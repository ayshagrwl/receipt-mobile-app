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

// ============================================================
//  CORS helper — wrap any response with proper headers
// ============================================================
function corsOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  doGet — handles search requests
//  ?action=search&q=QUERY
// ============================================================
function doGet(e) {
  try {
    const action = (e.parameter.action || '').trim();
    const q      = (e.parameter.q      || '').trim();

    if (action === 'search') {
      return corsOutput(searchRecords(q));
    }

    if (action === 'ping') {
      return corsOutput({ status: 'ok', sheet: SHEET_NAME });
    }

    return corsOutput({ error: 'Unknown action. Use ?action=search&q=QUERY' });
  } catch (err) {
    return corsOutput({ error: err.message });
  }
}

// ============================================================
//  doPost — handles payment submissions
//  Body JSON: { action:"pay", rowIndex:N, amount:500, mode:"Cash"|"Bank" }
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
//  searchRecords — searches invoice number OR customer name
// ============================================================
function searchRecords(q) {
  if (!q || q.length < 2) {
    return { error: 'Query too short. Minimum 2 characters.' };
  }

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const rows  = sheet.getDataRange().getValues();

  const query = q.toLowerCase();
  const results = [];

  // Start from row index 1 to skip header row
  for (let i = 1; i < rows.length; i++) {
    const row      = rows[i];
    const invoice  = String(row[COL.INVOICE]  || '').toLowerCase();
    const customer = String(row[COL.CUSTOMER] || '').toLowerCase();

    if (!invoice && !customer) continue; // skip completely empty rows

    if (invoice.includes(query) || customer.includes(query)) {
      results.push({
        rowIndex    : i + 1,                      // 1-based sheet row (header is row 1)
        date        : String(row[COL.DATE]        || ''),
        invoice     : String(row[COL.INVOICE]     || ''),
        customer    : String(row[COL.CUSTOMER]    || ''),
        amount      : String(row[COL.AMOUNT]      || ''),
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

    if (results.length >= 50) break; // cap at 50 results for speed
  }

  return { query: q, count: results.length, results };
}

// ============================================================
//  recordPayment — updates a row with payment info
// ============================================================
function recordPayment(data) {
  const rowIndex      = parseInt(data.rowIndex);
  const paymentAmount = parseFloat(data.amount);
  const mode          = String(data.mode || '').trim(); // "Cash" or "Bank"

  if (!rowIndex || isNaN(paymentAmount) || paymentAmount <= 0) {
    return { error: 'Invalid payment data.' };
  }
  if (mode !== 'Cash' && mode !== 'Bank') {
    return { error: 'Mode must be "Cash" or "Bank".' };
  }

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  // Read the current row
  const rowRange = sheet.getRange(rowIndex, 1, 1, 16);
  const row      = rowRange.getValues()[0];

  const currentPaid        = parseAmount(row[COL.PAID_UP]);
  const currentOutstanding = parseAmount(row[COL.OUTSTANDING]);

  const newPaid        = currentPaid + paymentAmount;
  const newOutstanding = currentOutstanding - paymentAmount;

  // Write updates
  sheet.getRange(rowIndex, COL.PAID_UP + 1)     .setValue(formatAmount(newPaid));
  sheet.getRange(rowIndex, COL.OUTSTANDING + 1)  .setValue(formatAmount(newOutstanding));
  sheet.getRange(rowIndex, COL.MODE + 1)         .setValue(mode);

  // Auto-mark PAID if fully settled
  if (newOutstanding <= 0) {
    sheet.getRange(rowIndex, COL.STATUS + 1).setValue('PAID');
  }

  return {
    success     : true,
    rowIndex    : rowIndex,
    newPaidUp   : formatAmount(newPaid),
    newOutstanding: formatAmount(newOutstanding),
    mode        : mode,
  };
}

// ============================================================
//  Helpers
// ============================================================
function parseAmount(val) {
  // Handles "₹7,100" / "-₹1,980" / "0" / ""
  const cleaned = String(val || '0').replace(/[₹,\s]/g, '');
  const num     = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

function formatAmount(num) {
  const sign = num < 0 ? '-' : '';
  const abs  = Math.abs(num);
  return sign + '₹' + abs.toLocaleString('en-IN');
}
