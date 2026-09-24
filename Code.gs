// ============================================================
//  OM MARKETING — Receipt Entry & Lookup System
//  Google Apps Script Web App (doGet + doPost)
//  Sheet: ALL INVOICE PARTY (OM MARKETING) - MARCH-SEPT
//  Preserves exact logic from ReceiptForm.html & original Code.gs
//
//  Columns in Sales Voucher sheet:
//  D = Invoice Number (4)
//  E = Customer (5)
//  F = Amount (6)
//  G = Overdue (7)
//  H = Discount (8)     <-- Appended as formula: =old+val
//  I = PAID-U (9)       <-- Appended as formula: =old+val
//  J = Status (10)      <-- PAID if OUTSTAND <= 0 else PARTIAL
//  K = MODE (11)        <-- DD-MMM-YYYY CASH ₹... + BANK ₹...
//  L = OUTSTAND (12)    <-- NEVER TOUCHED (Formula preserved)
//  M = RECEIPT (13)     <-- Appended with ' + '
//  N = REMARKS (14)     <-- Appended with ' + '
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

// Pure-JS date formatter for sheet sync
function formatFastDate(d) {
  if (!d) return '';
  if (d instanceof Date) {
    return d.getDate() + ' ' + MONTH_NAMES[d.getMonth()] + ' ' + d.getFullYear();
  }
  return String(d).trim();
}

// Spreadsheet resolver: supports both container-bound and standalone execution
function getSpreadsheet() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && active.getId()) return active;
  } catch (e) {}
  return SpreadsheetApp.openById(SHEET_ID);
}

// ============================================================
//  Desktop Google Sheets UI Menu & Modal (ReceiptForm)
// ============================================================
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Receipt Entry')
      .addItem('Add Receipt', 'openReceiptForm')
      .addToUi();
  } catch (e) {}
}

function openReceiptForm() {
  const ss = getSpreadsheet();
  const sheet = ss.getActiveSheet ? ss.getActiveSheet() : ss.getSheetByName(SHEET_NAME);
  const row = sheet.getActiveCell().getRow();

  // Don't allow header row
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('Please select an invoice row first.');
    return;
  }

  // Read existing values
  const invoice     = sheet.getRange(row, 4).getDisplayValue();  // D
  const customer    = sheet.getRange(row, 5).getDisplayValue();  // E
  const amount      = sheet.getRange(row, 6).getDisplayValue();  // F
  const discount    = sheet.getRange(row, 8).getDisplayValue();  // H
  const paid        = sheet.getRange(row, 9).getDisplayValue();  // I
  const status      = sheet.getRange(row, 10).getDisplayValue(); // J
  const mode        = sheet.getRange(row, 11).getDisplayValue(); // K
  const outstanding = sheet.getRange(row, 12).getDisplayValue(); // L
  const receipt     = sheet.getRange(row, 13).getDisplayValue(); // M
  const remarks     = sheet.getRange(row, 14).getDisplayValue(); // N

  if (!invoice) {
    SpreadsheetApp.getUi().alert('The selected row does not contain an Invoice Number.');
    return;
  }

  const template = HtmlService.createTemplateFromFile('ReceiptForm');
  template.row         = row;
  template.invoice     = invoice;
  template.customer    = customer;
  template.amount      = amount;
  template.discount    = discount;
  template.paid        = paid;
  template.status      = status;
  template.mode        = mode;
  template.outstanding = outstanding;
  template.receipt     = receipt;
  template.remarks     = remarks;

  const html = template.evaluate().setWidth(470).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Record Receipt');
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
//  doGet — handles fetchRecent, fetchAll, search, ping, pay requests
// ============================================================
function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = (params.action || '').trim();
    const q      = (params.q      || '').trim();

    if (action === 'fetchRecent' || action === 'recent') {
      const limit = parseInt(params.limit || 1500);
      return corsOutput(fetchRecentRecords(limit));
    }

    if (action === 'fetchAll' || action === 'all') {
      return corsOutput(fetchAllRecords());
    }

    if (action === 'pay' || action === 'saveReceipt') {
      const payData = {
        row       : params.row || params.rowIndex,
        rowIndex  : params.row || params.rowIndex,
        sheetName : params.sheetName || SHEET_NAME,
        date      : params.date || '',
        cash      : params.cash || 0,
        bank      : params.bank || 0,
        discount  : params.discount || 0,
        receiptNo : params.receiptNo || params.receipt || '',
        receipt   : params.receiptNo || params.receipt || '',
        remarks   : params.remarks || ''
      };
      return corsOutput(saveReceipt(payData));
    }

    if (action === 'search') {
      return corsOutput(searchRecords(q));
    }

    if (action === 'ping') {
      return corsOutput({ status: 'ok', sheet: SHEET_NAME });
    }

    return corsOutput({ error: 'Unknown action. Use ?action=fetchRecent, ?action=fetchAll, or ?action=search&q=QUERY' });
  } catch (err) {
    return corsOutput({ error: err.message });
  }
}

// ============================================================
//  doPost — handles payment submissions from web app
// ============================================================
function doPost(e) {
  try {
    let data;
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    } else {
      return corsOutput({ error: 'No POST data received.' });
    }

    const action = String(data.action || '').trim();
    if (action === 'pay' || action === 'saveReceipt' || !action) {
      return corsOutput(saveReceipt(data));
    }

    return corsOutput({ error: 'Unknown POST action: ' + action });
  } catch (err) {
    return corsOutput({ error: err.message });
  }
}

// ============================================================
//  saveReceipt — exact implementation matching original Code.gs
//  and ReceiptForm.html
// ============================================================
function saveReceipt(data) {
  const ss = getSpreadsheet();
  const targetSheetName = data.sheetName || SHEET_NAME;
  const sheet = ss.getSheetByName(targetSheetName);

  if (!sheet) {
    throw new Error('Sales sheet not found: ' + targetSheetName);
  }

  const row = Number(data.row || data.rowIndex);
  if (!row || row < 2) {
    throw new Error('Invalid row number: ' + row);
  }

  const cash     = Number(data.cash)     || 0;
  const bank     = Number(data.bank)     || 0;
  const discount = Number(data.discount) || 0;

  // At least one payment/discount required
  if (cash === 0 && bank === 0 && discount === 0) {
    throw new Error('Please enter Cash, Bank, or Discount.');
  }

  /********************************************
   * H = DISCOUNT (Column 8)
   * Append numeric value while preserving formula
   ********************************************/
  if (discount !== 0) {
    const discountCell = sheet.getRange(row, 8);
    appendToExistingCell(discountCell, discount);
  }

  /********************************************
   * I = PAID-U (Column 9)
   * Existing formula is preserved:
   * =1000+600 becomes =1000+600+500
   ********************************************/
  const totalPayment = cash + bank;
  if (totalPayment !== 0) {
    const paidCell = sheet.getRange(row, 9);
    appendToExistingCell(paidCell, totalPayment);
  }

  /********************************************
   * J = STATUS (Column 10)
   *
   * We DO NOT touch OUTSTAND (Column 12).
   * It contains a formula (=F - H - I).
   * We flush and simply read its calculated result.
   ********************************************/
  SpreadsheetApp.flush();

  const outstandingCell = sheet.getRange(row, 12);
  const outstandingValue = Number(outstandingCell.getValue()) || 0;

  const statusCell = sheet.getRange(row, 10);
  if (outstandingValue <= 0) {
    statusCell.setValue('PAID');
  } else {
    statusCell.setValue('PARTIAL');
  }

  /********************************************
   * K = MODE (Column 11)
   * Append: DD-MMM-YYYY CASH ₹... + BANK ₹... + DISCOUNT ₹...
   ********************************************/
  const modeCell = sheet.getRange(row, 11);
  const dateText = formatReceiptDate(data.date);

  const paymentParts = [];
  if (cash > 0) {
    paymentParts.push('CASH ₹' + formatIndianNumber(cash));
  }
  if (bank > 0) {
    paymentParts.push('BANK ₹' + formatIndianNumber(bank));
  }
  if (discount > 0) {
    paymentParts.push('DISCOUNT ₹' + formatIndianNumber(discount));
  }

  const modeEntry = dateText + ' ' + paymentParts.join(' + ');
  appendText(modeCell, modeEntry);

  /********************************************
   * M = RECEIPT (Column 13)
   * Append receipt number
   ********************************************/
  const receiptNo = String(data.receiptNo || data.receipt || '').trim();
  if (receiptNo) {
    const receiptCell = sheet.getRange(row, 13);
    appendText(receiptCell, receiptNo);
  }

  /********************************************
   * N = REMARKS (Column 14)
   * Append new remarks
   ********************************************/
  const remarks = String(data.remarks || '').trim();
  if (remarks) {
    const remarksCell = sheet.getRange(row, 14);
    appendText(remarksCell, remarks);
  }

  /********************************************
   * CREATE / UPDATE RECEIPTS SHEET
   * Exact tab name: "Receipts" (11 columns)
   ********************************************/
  let receiptSheet = ss.getSheetByName('Receipts');
  if (!receiptSheet) {
    receiptSheet = ss.insertSheet('Receipts');
    receiptSheet.appendRow([
      'Receipt Date',
      'Invoice Number',
      'Customer',
      'Invoice Amount',
      'Cash',
      'Bank',
      'Discount',
      'Total Payment',
      'Receipt Number',
      'Remarks',
      'Sales Sheet Row'
    ]);
    receiptSheet.setFrozenRows(1);
    try {
      receiptSheet.getRange(1, 1, 1, 11).setFontWeight('bold');
    } catch (e) {}
  }

  const invoice       = sheet.getRange(row, 4).getDisplayValue();
  const customer      = sheet.getRange(row, 5).getDisplayValue();
  const invoiceAmount = sheet.getRange(row, 6).getValue();

  const receiptDate = data.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Kolkata', 'yyyy-MM-dd');

  receiptSheet.appendRow([
    receiptDate,
    invoice,
    customer,
    invoiceAmount,
    cash || '',
    bank || '',
    discount || '',
    totalPayment || '',
    receiptNo || '',
    remarks || '',
    row
  ]);

  SpreadsheetApp.flush();

  const finalOutstandingDisp = sheet.getRange(row, 12).getDisplayValue();
  const finalPaidDisp        = sheet.getRange(row, 9).getDisplayValue();
  const finalDiscountDisp    = sheet.getRange(row, 8).getDisplayValue();
  const finalModeDisp        = sheet.getRange(row, 11).getDisplayValue();
  const finalReceiptDisp     = sheet.getRange(row, 13).getDisplayValue();
  const finalRemarksDisp     = sheet.getRange(row, 14).getDisplayValue();
  const finalStatusDisp      = sheet.getRange(row, 10).getDisplayValue();

  return {
    success        : true,
    row            : row,
    rowIndex       : row,
    outstanding    : outstandingValue,
    newOutstanding : finalOutstandingDisp || ('₹' + formatIndianNumber(outstandingValue)),
    newPaidUp      : finalPaidDisp,
    newDiscount    : finalDiscountDisp,
    mode           : finalModeDisp,
    receipt        : finalReceiptDisp,
    remarks        : finalRemarksDisp,
    status         : finalStatusDisp
  };
}

// Alias for backwards compatibility
const recordPayment = saveReceipt;

// ============================================================
//  appendToExistingCell — appends numeric value while preserving formula
//  e.g. =1000+600 becomes =1000+600+500
// ============================================================
function appendToExistingCell(cell, value) {
  value = Number(value) || 0;
  if (value === 0) return;

  const formula = cell.getFormula();
  const currentValue = cell.getValue();

  // CASE 1: Cell already contains a formula
  if (formula) {
    cell.setFormula(formula + '+' + value);
    return;
  }

  // CASE 2: Cell is completely blank
  if (currentValue === '' || currentValue === null || currentValue === undefined) {
    cell.setFormula('=' + value);
    return;
  }

  // CASE 3: Cell contains a normal number or numeric string
  let numVal = Number(currentValue);
  if (isNaN(numVal) && typeof currentValue === 'string') {
    numVal = Number(currentValue.replace(/[₹,\s]/g, '')) || 0;
  }

  if (!numVal) {
    cell.setFormula('=' + value);
  } else {
    cell.setFormula('=' + numVal + '+' + value);
  }
}

// ============================================================
//  appendText — appends text with ' + ' separator
// ============================================================
function appendText(cell, newText) {
  if (!newText) return;

  const oldText = cell.getDisplayValue();
  if (!oldText) {
    cell.setValue(newText);
  } else {
    cell.setValue(oldText + ' + ' + newText);
  }
}

// ============================================================
//  formatReceiptDate — converts YYYY-MM-DD or date to DD-MMM-YYYY
// ============================================================
function formatReceiptDate(dateString) {
  if (!dateString) {
    return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Kolkata', 'dd-MMM-yyyy');
  }

  // If YYYY-MM-DD string, parse safely to avoid timezone day rollover
  if (typeof dateString === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(dateString.trim())) {
    const parts = dateString.trim().split('-');
    const year  = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const day   = parseInt(parts[2], 10);
    const dd    = String(day).padStart(2, '0');
    return dd + '-' + MONTH_NAMES[month] + '-' + year;
  }

  const date = new Date(dateString);
  if (isNaN(date.getTime())) return String(dateString);

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone() || 'Asia/Kolkata',
    'dd-MMM-yyyy'
  );
}

// ============================================================
//  formatIndianNumber — formats number with Indian comma grouping
// ============================================================
function formatIndianNumber(number) {
  return Number(number).toLocaleString('en-IN', {
    maximumFractionDigits: 2
  });
}

// ============================================================
//  fetchRecentRecords — returns the latest N rows (~1 second!)
// ============================================================
function fetchRecentRecords(limit) {
  const ss      = getSpreadsheet();
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
  const ss      = getSpreadsheet();
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

  const ss      = getSpreadsheet();
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
