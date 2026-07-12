/**
 * Parser for ICICI Bank detailed statement XLS files
 * Extracts sender name, UPI ID, transaction type from remarks column
 */

const XLSX = require('xlsx');

// Known non-rent senders to exclude
const EXCLUDE_KEYWORDS = [
  'netflix', 'amazon', 'airtel', 'autopay', 'phonepe', 'bbpsbp',
  'interest', 'int.pd', 'urban comp', 'urbancompany', 'mlpi',
  'dcardfee', 'kindle', 'refund', 'attestr', 'accountvalidati',
  'vanceaccountver', 'instantpay', 'apibanking', 'panacea',
  'ekoindiafi', 'ketla tech', 'credit car', 'gold', 'insurance',
  'bbps', 'bpay', 'dtax', 'idtx', 'electricity', 'ghmc',
  'greater hy', 'bill payment', 'axis/amazon', 'amazon pay',
  'payout', 'api bankin', 'impstxn'
];

// UPI name overrides — maps UPI sender token → clean display name
const UPI_NAME_MAP = {
  '8688032788@ybl':  'Pulapa Harish',
  '8688032788@axl':  'Pulapa Harish',
  '9010612889@axl':  'Chandra Shekar',
  'bandharamraghav': 'Sri Raghav Bandharam',
  'venkateshjoonla': 'Venkateshjoonla',
  'parameswaraormp': 'Parameswara Rao (Mangalagiri)',
  'yalamanchililsu': 'Yalamanchili',
  '9347474547@ibl':  'Mr Pattila Saibabu',
  'saibabu.pattela': 'Mr Pattila Saibabu',
  '9866648356@ybl':  'Addepalli',
  '9866648356@axl':  'Addepalli',
  '9866648356@ibl':  'Addepalli',
  '8125474245@axl':  'Lakshmana',
  '8125474245@ibl':  'Lakshmana',
  '8125474245@pts':  'Lakshmana',
  '8125474245-5@y':  'Lakshmana',
  '8125474245-5@a':  'Lakshmana',
  'krishna8888gvk':  'Guttikonda Krishna',
  'chandrashekarc':  'Y Chandra Shekar',
  '9010612889@axl':  'Y Chandra Shekar',
  'shilparenati1@':  'Renati Nagamani',
  'trivenichandu1':  'Yandamuri Triveni',
  'srilakshmistud':  'Kommareddy Srilakshmi',
  '7207723551@axl':  'Callapilli',
  '7207723551@ybl':  'Callapilli',
  '9398279718@ybl':  'M Hanumanta Rao',
  'yeruvasagar685':  'Yeruva Chandra Sagar',
  '8074740734@ybl':  'Sri Vighneswara',
  '9652518298-2@y':  'Muppidi Upendra',
  '9949794045-2@i':  'Gopireddy',
  '309160221230@y':  'Kowtharapu',
  'mranjithkumar0':  'Mallampalli Ranjith Kumar',
  '9948333884@ibl':  'Vijaya Sya',
  'javvadimanasa@':  'Javvadi Manasa',
};


/**
 * Extract a 10-digit Indian mobile number from a UPI ID or remarks string.
 * Handles patterns like:  9866648356@ybl  |  8125474245-5@axl  |  9652518298-2@ybl
 */
function extractPhone(upiId, remarks) {
  // Try UPI id first – the number before @ or before - suffix
  if (upiId) {
    const m = upiId.match(/^(\d{10})/);
    if (m) return m[1];
    // also handles 9652518298-2@ybl
    const m2 = upiId.match(/(\d{10})/);
    if (m2) return m2[1];
  }
  // Fallback – look for 10-digit run in the full remarks
  if (remarks) {
    const m = remarks.match(/\b(\d{10})\b/);
    if (m) return m[1];
  }
  return null;
}

/**
 * Extract sender name from transaction remarks
 * Handles UPI, IMPS, NEFT, BIL patterns
 */
function extractSenderInfo(remarks) {
  if (!remarks) return { senderName: null, upiId: null, phone: null, txType: 'OTHER' };

  const r = remarks.trim();

  // --- UPI pattern: UPI/SENDER NAME/upi@id/description/BANK/ref ---
  const upiMatch = r.match(/^UPI\/([^\/]+)\/([^\/\s]+)/i);
  if (upiMatch) {
    const rawName = upiMatch[1].trim();
    const upiId = upiMatch[2].trim().toLowerCase();

    // Check name map by UPI id first, then by raw name token
    let resolvedName = null;
    for (const [key, val] of Object.entries(UPI_NAME_MAP)) {
      if (upiId.includes(key.toLowerCase())) { resolvedName = val; break; }
    }
    if (!resolvedName) {
      for (const [key, val] of Object.entries(UPI_NAME_MAP)) {
        if (rawName.toLowerCase().includes(key.toLowerCase())) { resolvedName = val; break; }
      }
    }

    return {
      senderName: resolvedName || toTitleCase(rawName),
      upiId,
      phone: extractPhone(upiId, r),
      txType: 'UPI'
    };
  }

  // --- IMPS pattern: MMT/IMPS/refno/Name/bank ---
  const impsMatch = r.match(/MMT\/IMPS\/\d+\/([^\/]+)\//i);
  if (impsMatch) {
    const rawName = impsMatch[1].trim();
    // exclude system names
    if (['payout', 'na', 'imps transactio', 'gopi amount', 'bill payment'].includes(rawName.toLowerCase())) {
      return { senderName: null, upiId: null, phone: null, txType: 'IMPS' };
    }
    return {
      senderName: toTitleCase(rawName),
      upiId: null,
      phone: extractPhone(null, r),
      txType: 'IMPS'
    };
  }

  // --- NEFT pattern: NEFT-...-NAME-... ---
  const neftMatch = r.match(/NEFT[-\/][A-Z0-9]+-([A-Z ]+)-/i);
  if (neftMatch) {
    const rawName = neftMatch[1].trim();
    return {
      senderName: toTitleCase(rawName),
      upiId: null,
      phone: extractPhone(null, r),
      txType: 'NEFT'
    };
  }

  // --- BIL/NEFT pattern ---
  const bilNeftMatch = r.match(/BIL\/NEFT\/[A-Z0-9]+\/([^\/]+)\//i);
  if (bilNeftMatch) {
    const rawName = bilNeftMatch[1].trim();
    return {
      senderName: toTitleCase(rawName),
      upiId: null,
      phone: extractPhone(null, r),
      txType: 'NEFT'
    };
  }

  // --- RTGS ---
  const rtgsMatch = r.match(/RTGS\/[A-Z0-9]+\/([^\/]+)\/([^\/]+)/i);
  if (rtgsMatch) {
    const rawName = rtgsMatch[2].trim();
    return {
      senderName: toTitleCase(rawName),
      upiId: null,
      phone: extractPhone(null, r),
      txType: 'RTGS'
    };
  }

  // --- Interest credit ---
  if (r.includes('Int.Pd') || r.toLowerCase().includes('interest')) {
    return { senderName: 'Bank Interest', upiId: null, phone: null, txType: 'INTEREST' };
  }

  return { senderName: null, upiId: null, phone: null, txType: 'OTHER' };
}

function toTitleCase(str) {
  if (!str) return str;
  return str
    .toLowerCase()
    .replace(/\b\w/g, c => c.toUpperCase())
    .replace(/\s+/g, ' ')
    .trim();
}

function isExcluded(remarks, senderName) {
  const combined = ((remarks || '') + ' ' + (senderName || '')).toLowerCase();
  return EXCLUDE_KEYWORDS.some(kw => combined.includes(kw));
}

function parseDate(dateStr) {
  if (!dateStr) return null;
  // Format: DD/MM/YYYY
  const match = String(dateStr).match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (match) {
    return `${match[3]}-${match[2]}-${match[1]}`;
  }
  return null;
}

function parseAmount(val) {
  const num = parseFloat(String(val).replace(/,/g, '').trim());
  return isNaN(num) ? 0.00 : num;
}

function getMonthYear(dateStr) {
  // dateStr in YYYY-MM-DD
  if (!dateStr) return null;
  return dateStr.substring(0, 7); // YYYY-MM
}

/**
 * Parse the XLS buffer and return structured data
 */
function parseExcel(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  // Extract header info
  let accountNumber = '';
  let accountHolder = '';
  let dateFrom = null;
  let dateTo = null;

  const transactions = [];
  let headerRowFound = false;
  let pendingRemark = null; // for multi-line remarks

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const cells = row.map(c => String(c || '').trim());

    // Parse metadata rows
    if (cells[1] === 'Account Number' && cells[3]) {
      const accMatch = cells[3].match(/(\d+)\s*\(\s*INR\s*\)\s*-\s*(.+)/i);
      if (accMatch) {
        accountNumber = accMatch[1].trim();
        accountHolder = accMatch[2].trim();
      }
    }

    if (cells[1] === 'Transaction Date from' && cells[3]) {
      dateFrom = parseDate(cells[3]);
      dateTo = parseDate(cells[5]);
    }

    // Detect header row
    if (cells[1] === 'S No.' && cells[2] === 'Value Date') {
      headerRowFound = true;
      continue;
    }

    if (!headerRowFound) continue;

    // Stop at legends
    if (cells[1].startsWith('Legends') || cells[1].startsWith('1. INFT')) break;

    const sno = cells[1];
    const valueDate = cells[2];
    const txDate = cells[3];
    const cheque = cells[4];
    let remarks = cells[5];
    const withdrawal = cells[6];
    const deposit = cells[7];
    const balance = cells[8];

    // Handle continuation row (split remarks)
    if (sno === '' && remarks !== '') {
      // This is a continuation of the previous transaction's remarks
      if (transactions.length > 0) {
        transactions[transactions.length - 1].transaction_remarks += remarks;
      }
      continue;
    }

    // Skip empty rows
    if (!sno || !txDate) continue;

    const parsedDate = parseDate(txDate);
    const parsedValueDate = parseDate(valueDate);
    if (!parsedDate) continue;

    const withdrawalAmt = parseAmount(withdrawal);
    const depositAmt = parseAmount(deposit);

    const { senderName, upiId, phone, txType } = extractSenderInfo(remarks);

    const excluded = isExcluded(remarks, senderName);
    const isDeposit = depositAmt > 0;
    const isRent = isDeposit && !excluded && senderName !== null && senderName !== 'Bank Interest';

    transactions.push({
      sno: parseInt(sno) || 0,
      value_date: parsedValueDate,
      transaction_date: parsedDate,
      cheque_number: cheque,
      transaction_remarks: remarks,
      withdrawal_amount: withdrawalAmt,
      deposit_amount: depositAmt,
      balance: parseAmount(balance),
      sender_name: senderName,
      upi_id: upiId,
      phone: phone || null,
      transaction_type: txType,
      is_rent: isRent ? 1 : 0,
      month_year: getMonthYear(parsedDate),
    });
  }

  return {
    accountNumber,
    accountHolder,
    dateFrom,
    dateTo,
    transactions,
  };
}

module.exports = { parseExcel, extractSenderInfo };
