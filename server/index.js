const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
require('dotenv').config();

/**
 * Build a deterministic SHA-256 fingerprint for a transaction row.
 * This is the SINGLE source of truth for deduplication —
 * it is stored in the DB as a UNIQUE column, so INSERT IGNORE
 * silently drops any row whose hash already exists.
 *
 * Fields used:  transaction_date | first-200-chars-of-remarks |
 *               withdrawal_amount | deposit_amount
 *
 * NOTE: balance is intentionally excluded because a corrected
 * statement might carry a different running balance for the same
 * underlying transaction.
 */
function buildTxHash(tx) {
  const raw = [
    tx.transaction_date || '',
    (tx.transaction_remarks || '').substring(0, 200).trim(),
    parseFloat(tx.withdrawal_amount || 0).toFixed(2),
    parseFloat(tx.deposit_amount   || 0).toFixed(2),
  ].join('|');
  return crypto.createHash('sha256').update(raw).digest('hex');
}

const db = require('./db');
const { parseExcel } = require('./parser');

const app = express();
const PORT = process.env.PORT || 3001;

// ─── Auto-migrate: add phone column if not yet present ─────────────────────
(async () => {
  try {
    await db.query(`ALTER TABLE transactions ADD COLUMN IF NOT EXISTS phone VARCHAR(15) NULL AFTER upi_id`);
    await db.query(`ALTER TABLE tenants     ADD COLUMN IF NOT EXISTS phone VARCHAR(15) NULL AFTER display_name`);
    // Backfill phone for existing rows that have a 10-digit UPI id
    await db.query(`
      UPDATE transactions
      SET phone = REGEXP_SUBSTR(upi_id, '[0-9]{10}')
      WHERE phone IS NULL AND upi_id REGEXP '[0-9]{10}'
    `);
    console.log('✅ DB migration OK (phone columns ready)');
  } catch (e) {
    // Older MySQL may not support IF NOT EXISTS on ALTER – ignore gracefully
    if (!e.message.includes('Duplicate column')) console.warn('Migration note:', e.message);
  }
})();

// Uploads directory
const UPLOAD_DIR = path.join(__dirname, '..', 'uploads');
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

// Multer setup
const storage = multer.memoryStorage();
const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    if (
      file.originalname.endsWith('.xls') ||
      file.originalname.endsWith('.xlsx') ||
      file.mimetype.includes('spreadsheet') ||
      file.mimetype.includes('excel') ||
      file.mimetype === 'application/vnd.ms-excel'
    ) {
      cb(null, true);
    } else {
      cb(new Error('Only .xls or .xlsx files are allowed'));
    }
  },
  limits: { fileSize: 20 * 1024 * 1024 },
});

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '..', 'public')));

// ─── Upload & Process Excel ────────────────────────────────────────────────
app.post('/api/upload', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

  const conn = await db.getConnection();
  try {
    const parsed = parseExcel(req.file.buffer);
    const { accountNumber, accountHolder, dateFrom, dateTo, transactions } = parsed;

    // ── Step 1: record this upload attempt ────────────────────────────────
    const [uploadResult] = await conn.query(
      `INSERT INTO uploads
         (filename, account_number, account_holder, date_from, date_to, total_rows)
       VALUES (?, ?, ?, ?, ?, ?)`,
      [req.file.originalname, accountNumber, accountHolder,
       dateFrom, dateTo, transactions.length]
    );
    const uploadId = uploadResult.insertId;

    // ── Step 2: bulk INSERT IGNORE – DB unique constraint on tx_hash
    //    is the final guard against duplicates.
    //    No SELECT-then-INSERT loop → no race conditions, no double counting.
    let newCount = 0;
    let dupCount = 0;

    for (const tx of transactions) {
      const hash = buildTxHash(tx);

      const [result] = await conn.query(
        `INSERT IGNORE INTO transactions
           (upload_id, sno, value_date, transaction_date, cheque_number,
            transaction_remarks, withdrawal_amount, deposit_amount, balance,
            sender_name, upi_id, phone, transaction_type, is_rent, month_year, tx_hash)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [
          uploadId,
          tx.sno,
          tx.value_date,
          tx.transaction_date,
          tx.cheque_number,
          tx.transaction_remarks,
          tx.withdrawal_amount,
          tx.deposit_amount,
          tx.balance,
          tx.sender_name,
          tx.upi_id,
          tx.phone || null,
          tx.transaction_type,
          tx.is_rent,
          tx.month_year,
          hash,
        ]
      );

      // affectedRows = 0 means the unique constraint fired → duplicate
      if (result.affectedRows === 0) {
        dupCount++;
      } else {
        newCount++;

        // ── Upsert tenant record for rent deposits ─────────────────────
        if (tx.is_rent && tx.sender_name) {
          await conn.query(
            `INSERT INTO tenants (name, display_name, phone, first_seen, last_seen)
             VALUES (?, ?, ?, ?, ?)
             ON DUPLICATE KEY UPDATE
               last_seen = GREATEST(COALESCE(last_seen, '1970-01-01'), VALUES(last_seen)),
               phone     = COALESCE(phone, VALUES(phone))`,
            [tx.sender_name, tx.sender_name, tx.phone || null,
             tx.transaction_date, tx.transaction_date]
          );
        }
      }
    }

    // ── Step 3: recalculate tenant totals from scratch (never accumulate) ─
    //    We always recompute from the raw transaction rows so the totals
    //    can never drift no matter how many times the same file is uploaded.
    await conn.query(`
      UPDATE tenants t
      JOIN (
        SELECT sender_name,
               SUM(deposit_amount)           AS total_paid,
               COUNT(DISTINCT month_year)    AS months_paid
        FROM   transactions
        WHERE  is_rent = 1
        GROUP  BY sender_name
      ) s ON t.name = s.sender_name
      SET t.total_paid  = s.total_paid,
          t.months_paid = s.months_paid,
          t.updated_at  = NOW()
    `);

    // ── Step 4: finalise upload record ────────────────────────────────────
    await conn.query(
      `UPDATE uploads
       SET new_transactions = ?, duplicate_transactions = ?
       WHERE id = ?`,
      [newCount, dupCount, uploadId]
    );

    res.json({
      success: true,
      uploadId,
      accountNumber,
      accountHolder,
      dateFrom,
      dateTo,
      total: transactions.length,
      newCount,
      dupCount,
      message: newCount > 0
        ? `${newCount} new transaction(s) added. ${dupCount} already existed and were skipped.`
        : `Nothing new — all ${dupCount} transaction(s) already exist in the database.`,
    });

  } catch (err) {
    console.error('Upload error:', err);
    res.status(500).json({ error: err.message });
  } finally {
    conn.release();
  }
});

// ─── Dashboard Summary ─────────────────────────────────────────────────────
app.get('/api/dashboard', async (req, res) => {
  try {
    const [[summary]] = await db.query(`
      SELECT
        COUNT(*) AS total_transactions,
        SUM(deposit_amount) AS total_deposits,
        SUM(withdrawal_amount) AS total_withdrawals,
        SUM(CASE WHEN is_rent=1 THEN deposit_amount ELSE 0 END) AS total_rent_received,
        COUNT(CASE WHEN is_rent=1 THEN 1 END) AS total_rent_txns
      FROM transactions
    `);

    const [tenants] = await db.query(`
      SELECT name, display_name, expected_rent, total_paid, months_paid,
             first_seen, last_seen, is_active
      FROM tenants
      WHERE is_active=1
      ORDER BY total_paid DESC
    `);

    const [monthlyRents] = await db.query(`
      SELECT month_year,
             SUM(deposit_amount) AS total,
             COUNT(*) AS count
      FROM transactions
      WHERE is_rent=1
      GROUP BY month_year
      ORDER BY month_year ASC
    `);

    const [recentUploads] = await db.query(`
      SELECT id, filename, account_holder, date_from, date_to,
             total_rows, new_transactions, duplicate_transactions, uploaded_at
      FROM uploads
      ORDER BY uploaded_at DESC
      LIMIT 5
    `);

    res.json({ summary, tenants, monthlyRents, recentUploads });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── Rent by Sender ────────────────────────────────────────────────────────
app.get('/api/rents/by-sender', async (req, res) => {
  try {
    const [rows] = await db.query(`
      SELECT
        sender_name,
        month_year,
        SUM(deposit_amount)                                          AS amount,
        COUNT(*)                                                     AS tx_count,
        MIN(transaction_date)                                        AS first_date,
        MAX(transaction_date)                                        AS last_date,
        -- pick the first non-null phone seen for this sender
        MAX(phone)                                                   AS phone,
        MAX(upi_id)                                                  AS upi_id,
        GROUP_CONCAT(transaction_remarks ORDER BY transaction_date SEPARATOR '||') AS remarks
      FROM transactions
      WHERE is_rent=1
      GROUP BY sender_name, month_year
      ORDER BY sender_name, month_year
    `);
    res.json(rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── Sender detail ─────────────────────────────────────────────────────────
app.get('/api/rents/sender/:name', async (req, res) => {
  try {
    const name = decodeURIComponent(req.params.name);
    const [rows] = await db.query(`
      SELECT *
      FROM transactions
      WHERE is_rent=1 AND sender_name = ?
      ORDER BY transaction_date ASC
    `, [name]);
    res.json(rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── Ledger (income only by sender) ───────────────────────────────────────
app.get('/api/ledger', async (req, res) => {
  try {
    const [rows] = await db.query(`
      SELECT
        sender_name,
        MAX(phone)                                                            AS phone,
        SUM(deposit_amount)                                                   AS total_received,
        COUNT(*)                                                              AS payment_count,
        MIN(transaction_date)                                                 AS first_payment,
        MAX(transaction_date)                                                 AS last_payment,
        COUNT(DISTINCT month_year)                                            AS months_paid,
        GROUP_CONCAT(DISTINCT month_year ORDER BY month_year SEPARATOR ',')  AS paid_months
      FROM transactions
      WHERE is_rent=1
      GROUP BY sender_name
      ORDER BY total_received DESC
    `);
    res.json(rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── Missed Rents ──────────────────────────────────────────────────────────
app.get('/api/missed-rents', async (req, res) => {
  try {
    // Get all month_years present in data
    const [allMonths] = await db.query(`
      SELECT DISTINCT month_year
      FROM transactions
      WHERE is_rent=1
      ORDER BY month_year
    `);

    // Get each sender's paid months
    const [senderMonths] = await db.query(`
      SELECT sender_name, GROUP_CONCAT(DISTINCT month_year ORDER BY month_year) AS paid_months,
             MIN(month_year) AS first_month, MAX(month_year) AS last_month
      FROM transactions
      WHERE is_rent=1
      GROUP BY sender_name
    `);

    const months = allMonths.map(r => r.month_year);
    const missed = [];

    for (const sender of senderMonths) {
      const paidSet = new Set((sender.paid_months || '').split(','));
      const firstMonth = sender.first_month;
      const lastMonth = sender.last_month;

      // Only check months between first and last payment
      const relevantMonths = months.filter(m => m >= firstMonth && m <= lastMonth);
      const missedMonths = relevantMonths.filter(m => !paidSet.has(m));

      if (missedMonths.length > 0) {
        missed.push({
          sender_name: sender.sender_name,
          paid_months: (sender.paid_months || '').split(','),
          missed_months: missedMonths,
          first_month: firstMonth,
          last_month: lastMonth,
        });
      }
    }

    res.json(missed);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── All transactions ──────────────────────────────────────────────────────
app.get('/api/transactions', async (req, res) => {
  try {
    const page = parseInt(req.query.page) || 1;
    const limit = parseInt(req.query.limit) || 50;
    const offset = (page - 1) * limit;
    const isRent = req.query.is_rent;
    const sender = req.query.sender;

    let where = '1=1';
    const params = [];
    if (isRent === '1') { where += ' AND is_rent=1'; }
    if (sender) { where += ' AND sender_name=?'; params.push(sender); }

    const [[{ total }]] = await db.query(`SELECT COUNT(*) as total FROM transactions WHERE ${where}`, params);
    const [rows] = await db.query(
      `SELECT * FROM transactions WHERE ${where} ORDER BY transaction_date DESC LIMIT ? OFFSET ?`,
      [...params, limit, offset]
    );
    res.json({ total, page, limit, rows });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── Update tenant expected rent ───────────────────────────────────────────
app.put('/api/tenants/:name', async (req, res) => {
  try {
    const name = decodeURIComponent(req.params.name);
    const { expected_rent, display_name, is_active } = req.body;
    await db.query(
      `UPDATE tenants SET
         expected_rent = COALESCE(?, expected_rent),
         display_name = COALESCE(?, display_name),
         is_active = COALESCE(?, is_active)
       WHERE name = ?`,
      [expected_rent ?? null, display_name ?? null, is_active ?? null, name]
    );
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── Toggle transaction rent flag ──────────────────────────────────────────
app.put('/api/transactions/:id/rent', async (req, res) => {
  try {
    const { is_rent } = req.body;
    await db.query(`UPDATE transactions SET is_rent=? WHERE id=?`, [is_rent ? 1 : 0, req.params.id]);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── Serve frontend ────────────────────────────────────────────────────────
app.get('/{*splat}', (req, res) => {
  res.sendFile(path.join(__dirname, '..', 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`✅ TrackMyRent server running at http://localhost:${PORT}`);
});
