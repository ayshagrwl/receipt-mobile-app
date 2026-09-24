// ============================================================
//  OM MARKETING — Receipt Lookup API
//  Google Apps Script Web App (doGet + doPost)
//  Sheet: ALL INVOICE PARTY (OM MARKETING) - MARCH-SEPT
// ============================================================

const SHEET_ID   = '1aNzEsD0v92fVZtklSj79ZxDtXCkySQxw63HVlydg88g';
const SHEET_NAME = 'MARCH-SEPT';

// --------------- Column indices (0-based) -------------------
const COL = {
  STATUS_COL1 : 0,   // Column 1 (DELIVERIED / MISSING etc.)
  PRESENT     : 1,
  DATE        : 2,
  INVOICE     : 3,
  CUSTOMER    : 4,
  AMOUNT      : 5,
  OVERDUE     : 6,
  DISCOUNT    : 7,
  PAID_UP     : 8,
  STATUS      : 9,
  MODE        : 10,
  OUTSTANDING : 11,
  RECEIPT     : 12,
  REMARKS     : 13,
  BEAT        : 14,
  AGENT       : 15,
};

const MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

// Fast pure-JS date formatter (0.001ms vs 3ms for Utilities.formatDate)
function formatFastDate(d) {
  if (!d) return '';
  if (d instanceof Date) {
    return d.getDate() + ' ' + MONTH_NAMES[d.getMonth()] + ' ' + d.getFullYear();
  }
  return String(d).trim();
}

// ============================================================
//  CORS helper — wrap any response with proper headers
// ============================================================
function corsOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  doGet — handles fetchRecent, fetchAll, search, ping requests
//  ?action=fetchRecent&limit=1500 (super fast sync, ~1 second)
//  ?action=fetchAll (full sheet export, ~3 seconds)
//  ?action=search&q=QUERY (legacy query search)
//  ?action=ping (healthcheck)
// ============================================================
function doGet(e) {
  try {
    const action = (e.parameter.action || '').trim();
    const q      = (e.parameter.q      || '').trim();

    if (action === 'fetchRecent' || action === 'recent') {
      const limit = parseInt(e.parameter.limit || 1500);
      return corsOutput(fetchRecentRecords(limit));
    }

    if (action === 'fetchAll' || action === 'all') {
      return corsOutput(fetchAllRecords());
    }

    if (action === 'search') {
      return corsOutput(searchRecords(q));
    }

    if (action === 'ping') {
      return corsOutput({ status: 'ok', sheet: SHEET_NAME });
    }

    return corsOutput({ error: 'Unknown action. Use ?action=fetchRecent or ?action=fetchAll or ?action=search&q=QUERY' });
  } catch (err) {
    return corsOutput({ error: err.message });
  }
}

// ============================================================
//  doPost — handles payment submissions
//  Body JSON: { action:"pay", rowIndex:N, cash:500, bank:200, discount:50, receipt:"...", remarks:"..." }
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'pay') {
      return corsOutput(recordPayment(data));
    }
    return corsOutput({ error: 'Unknown POST action' });
  } catch (err) {
    return corsOutput({ error: err.message });
  }
}

// ============================================================
//  fetchRecentRecords — returns the latest N rows (~1 second!)
// ============================================================
function fetchRecentRecords(limit) {
  const ss      = SpreadsheetApp.openById(SHEET_ID);
  const sheet   = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { status: 'ok', total: 0, lastRow: 0, rows: [] };
  }

  const n = Math.min(parseInt(limit) || 1500, lastRow - 1);
  const startRow = Math.max(2, lastRow - n + 1);
  const numRows  = lastRow - startRow + 1;

  // Read only the 16 columns needed (columns A to P)
  const rows = sheet.getRange(startRow, 1, numRows, 16).getValues();
  const results = [];

  for (let i = 0; i < rows.length; i++) {
    const row      = rows[i];
    const invoice  = row[COL.INVOICE];
    const customer = row[COL.CUSTOMER];

    if (!invoice && !customer) continue;

    results.push([
      startRow + i,                                          // 0: rowIndex (1-based sheet row)
      formatFastDate(row[COL.DATE]),                         // 1: date
      invoice  ? String(invoice).trim()  : '',               // 2: invoice
      customer ? String(customer).trim() : '',               // 3: customer
      row[COL.AMOUNT]      ? String(row[COL.AMOUNT]).trim()      : '', // 4: amount
      row[COL.PAID_UP]     ? String(row[COL.PAID_UP]).trim()     : '', // 5: paidUp
      row[COL.STATUS]      ? String(row[COL.STATUS]).trim()      : '', // 6: status
      row[COL.MODE]        ? String(row[COL.MODE]).trim()        : '', // 7: mode
      row[COL.OUTSTANDING] ? String(row[COL.OUTSTANDING]).trim() : '', // 8: outstanding
      row[COL.RECEIPT]     ? String(row[COL.RECEIPT]).trim()     : '', // 9: receipt
      row[COL.REMARKS]     ? String(row[COL.REMARKS]).trim()     : '', // 10: remarks
      row[COL.BEAT]        ? String(row[COL.BEAT]).trim()        : '', // 11: beat
      row[COL.AGENT]       ? String(row[COL.AGENT]).trim()       : '', // 12: agent
      row[COL.DISCOUNT]    ? String(row[COL.DISCOUNT]).trim()    : ''  // 13: discount / CD
    ]);
  }

  return {
    status    : 'ok',
    mode      : 'recent',
    total     : results.length,
    lastRow   : lastRow,
    updatedAt : Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMM yyyy HH:mm'),
    rows      : results
  };
}

// ============================================================
//  fetchAllRecords — returns all invoice records (optimized)
// ============================================================
function fetchAllRecords() {
  const ss      = SpreadsheetApp.openById(SHEET_ID);
  const sheet   = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { status: 'ok', total: 0, rows: [] };
  }

  // Read only columns A to P (16 columns), skipping header directly
  const rows = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
  const results = [];

  for (let i = 0; i < rows.length; i++) {
    const row      = rows[i];
    const invoice  = row[COL.INVOICE];
    const customer = row[COL.CUSTOMER];

    if (!invoice && !customer) continue;

    results.push([
      i + 2,                                                 // 0: rowIndex (1-based sheet row)
      formatFastDate(row[COL.DATE]),                         // 1: date
      invoice  ? String(invoice).trim()  : '',               // 2: invoice
      customer ? String(customer).trim() : '',               // 3: customer
      row[COL.AMOUNT]      ? String(row[COL.AMOUNT]).trim()      : '', // 4: amount
      row[COL.PAID_UP]     ? String(row[COL.PAID_UP]).trim()     : '', // 5: paidUp
      row[COL.STATUS]      ? String(row[COL.STATUS]).trim()      : '', // 6: status
      row[COL.MODE]        ? String(row[COL.MODE]).trim()        : '', // 7: mode
      row[COL.OUTSTANDING] ? String(row[COL.OUTSTANDING]).trim() : '', // 8: outstanding
      row[COL.RECEIPT]     ? String(row[COL.RECEIPT]).trim()     : '', // 9: receipt
      row[COL.REMARKS]     ? String(row[COL.REMARKS]).trim()     : '', // 10: remarks
      row[COL.BEAT]        ? String(row[COL.BEAT]).trim()        : '', // 11: beat
      row[COL.AGENT]       ? String(row[COL.AGENT]).trim()       : '', // 12: agent
      row[COL.DISCOUNT]    ? String(row[COL.DISCOUNT]).trim()    : ''  // 13: discount / CD
    ]);
  }

  return {
    status    : 'ok',
    mode      : 'all',
    total     : results.length,
    lastRow   : lastRow,
    updatedAt : Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMM yyyy HH:mm'),
    rows      : results
  };
}

// ============================================================
//  searchRecords — searches invoice number OR customer name
// ============================================================
function searchRecords(q) {
  if (!q || q.length < 2) {
    return { error: 'Query too short. Minimum 2 characters.' };
  }

  const ss      = SpreadsheetApp.openById(SHEET_ID);
  const sheet   = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const rows    = sheet.getRange(2, 1, lastRow - 1, 16).getValues();

  const query = q.toLowerCase();
  const results = [];

  for (let i = 0; i < rows.length; i++) {
    const row      = rows[i];
    const invoice  = String(row[COL.INVOICE]  || '').toLowerCase();
    const customer = String(row[COL.CUSTOMER] || '').toLowerCase();

    if (!invoice && !customer) continue;

    if (invoice.includes(query) || customer.includes(query)) {
      results.push({
        rowIndex    : i + 2,
        date        : formatFastDate(row[COL.DATE]),
        invoice     : String(row[COL.INVOICE]     || ''),
        customer    : String(row[COL.CUSTOMER]    || ''),
        amount      : String(row[COL.AMOUNT]      || ''),
        discount    : String(row[COL.DISCOUNT]    || ''),
        paidUp      : String(row[COL.PAID_UP]     || ''),
        status      : String(row[COL.STATUS]      || ''),
        mode        : String(row[COL.MODE]        || ''),
        outstanding : String(row[COL.OUTSTANDING] || ''),
        receipt     : String(row[COL.RECEIPT]     || ''),
        remarks     : String(row[COL.REMARKS]     || ''),
        beat        : String(row[COL.BEAT]        || ''),
        agent       : String(row[COL.AGENT]       || ''),
      });
    }

    if (results.length >= 50) break;
  }

  return { query: q, count: results.length, results };
}

// ============================================================
//  recordPayment — updates invoice row + logs payment entry
//  Supports split Cash + Bank payments + Discount / CD
// ============================================================
function recordPayment(data) {
  const rowIndex       = parseInt(data.rowIndex);
  const cashAmount     = parseFloat(data.cash     || 0);
  const bankAmount     = parseFloat(data.bank     || 0);
  const discountAmount = parseFloat(data.discount || 0);
  const totalPayment   = cashAmount + bankAmount;
  const totalSettled   = totalPayment + discountAmount;
  const receiptNo      = String(data.receipt || '').trim();
  const remarks        = String(data.remarks || '').trim();

  if (!rowIndex || totalSettled <= 0) {
    return { error: 'Enter at least a Cash, Bank, or Discount amount.' };
  }

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  // Read current row (16 columns)
  const rowRange = sheet.getRange(rowIndex, 1, 1, 16);
  const row      = rowRange.getValues()[0];

  const currentPaid        = parseAmount(row[COL.PAID_UP]);
  const currentDiscount    = parseAmount(row[COL.DISCOUNT]);
  const currentOutstanding = parseAmount(row[COL.OUTSTANDING]);
  const existingReceipt    = String(row[COL.RECEIPT] || '').trim();
  const existingRemarks    = String(row[COL.REMARKS] || '').trim();

  const newPaid        = currentPaid + totalPayment;
  const newDiscount    = currentDiscount + discountAmount;
  const newOutstanding = currentOutstanding - totalSettled;

  // Determine MODE label
  let mode = '';
  if (cashAmount > 0 && bankAmount > 0) mode = 'Cash + Bank';
  else if (cashAmount > 0)              mode = 'Cash';
  else if (bankAmount > 0)              mode = 'Bank';
  else if (discountAmount > 0)          mode = 'Discount';

  // Append receipt number
  const newReceipt = existingReceipt
    ? (receiptNo ? existingReceipt + ' + ' + receiptNo : existingReceipt)
    : receiptNo;

  // Append remarks
  const newRemarks = existingRemarks
    ? (remarks ? existingRemarks + ' | ' + remarks : existingRemarks)
    : remarks;

  // ── Update invoice row ──────────────────────────────────────
  if (discountAmount > 0 || newDiscount > 0) {
    sheet.getRange(rowIndex, COL.DISCOUNT + 1).setValue(formatAmount(newDiscount));
  }
  sheet.getRange(rowIndex, COL.PAID_UP + 1).setValue(formatAmount(newPaid));
  sheet.getRange(rowIndex, COL.OUTSTANDING + 1).setValue(formatAmount(newOutstanding));
  if (mode) sheet.getRange(rowIndex, COL.MODE + 1).setValue(mode);
  if (receiptNo) sheet.getRange(rowIndex, COL.RECEIPT + 1).setValue(newReceipt);
  if (remarks)   sheet.getRange(rowIndex, COL.REMARKS + 1).setValue(newRemarks);

  // Auto-mark PAID if fully settled
  if (newOutstanding <= 0) {
    sheet.getRange(rowIndex, COL.STATUS + 1).setValue('PAID');
  }

  // ── Append to PAYMENT_LOG sheet ────────────────────────────
  logPayment(ss, {
    date        : Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMM yyyy HH:mm'),
    invoice     : String(row[COL.INVOICE]  || ''),
    customer    : String(row[COL.CUSTOMER] || ''),
    cash        : cashAmount > 0 ? formatAmount(cashAmount) : '',
    bank        : bankAmount > 0 ? formatAmount(bankAmount) : '',
    discount    : discountAmount > 0 ? formatAmount(discountAmount) : '',
    total       : formatAmount(totalPayment),
    mode        : mode,
    receipt     : receiptNo,
    remarks     : remarks,
    outstanding : formatAmount(newOutstanding),
  });

  return {
    success        : true,
    rowIndex       : rowIndex,
    newPaidUp      : formatAmount(newPaid),
    newDiscount    : formatAmount(newDiscount),
    newOutstanding : formatAmount(newOutstanding),
    mode           : mode,
    receipt        : newReceipt,
  };
}

// ============================================================
//  logPayment — appends a row to PAYMENT_LOG sheet
// ============================================================
function logPayment(ss, entry) {
  const LOG_SHEET = 'PAYMENT_LOG';
  let log = ss.getSheetByName(LOG_SHEET);

  if (!log) {
    log = ss.insertSheet(LOG_SHEET);
    log.appendRow([
      'Date & Time', 'Invoice Number', 'Customer',
      'Cash', 'Bank', 'Discount', 'Total', 'Mode',
      'Receipt No.', 'Remarks', 'Outstanding After'
    ]);
    log.getRange(1, 1, 1, 11).setFontWeight('bold');
  }

  log.appendRow([
    entry.date, entry.invoice, entry.customer,
    entry.cash, entry.bank, entry.discount || '', entry.total, entry.mode,
    entry.receipt, entry.remarks, entry.outstanding,
  ]);
}

// ============================================================
//  Helpers
// ============================================================
function parseAmount(val) {
  const cleaned = String(val || '0').replace(/[₹,\s]/g, '');
  const num     = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

function formatAmount(num) {
  const sign = num < 0 ? '-' : '';
  const abs  = Math.abs(num);
  return sign + '₹' + abs.toLocaleString('en-IN');
}
