/**
 * Google Sheets Invoice Generator v1.0
 * 
 * Generates PDF invoices from a work log in Google Sheets.
 * Free, open-source, no third-party dependencies.
 * 
 * GitHub: [repo URL]
 * License: MIT
 */

// ============================================================
// CONSTANTS
// ============================================================

const SHEET_NAMES = {
  SETTINGS: 'Settings',
  CLIENTS: 'Clients',
  WORK_LOG: 'WorkLog',
  INVOICES: 'Invoices'
};

const WORK_LOG_COLS = {
  DATE: 1,        // A
  CLIENT_ID: 2,   // B
  DESCRIPTION: 3, // C
  QTY: 4,         // D
  RATE: 5,        // E
  TAX_PCT: 6,     // F
  STATUS: 7,      // G
  INVOICE_ID: 8   // H
};

const STATUS = {
  UNBILLED: 'Unbilled',
  INVOICED: 'Invoiced'
};

// ============================================================
// MENU
// ============================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('🧾 Invoice');

  // Check if setup has been completed (Settings sheet exists and has data)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  var isSetup = settingsSheet && settingsSheet.getLastRow() > 1;

  if (!isSetup) {
    menu.addItem('⚡ Set Up Template (run this first)', 'setupTemplate');
  }

  menu.addItem('Generate for Selected Client', 'generateForSelectedClient');
  menu.addItem('Generate for All Unbilled', 'generateForAllUnbilled');
  menu.addToUi();
}

// ============================================================
// TEMPLATE SETUP (one-click, idempotent)
// ============================================================

/**
 * Creates all four sheets with headers, formatting, data validation,
 * conditional formatting, and sample data. Safe to run multiple times —
 * skips any sheet that already exists and has data.
 */
function setupTemplate() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Guard: if Settings already has data, confirm before overwriting
  var existingSettings = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (existingSettings && existingSettings.getLastRow() > 1) {
    var confirm = ui.alert(
      '⚠️ Template Already Exists',
      'It looks like this spreadsheet has already been set up.\n\n' +
      'Running setup again will NOT overwrite existing sheets.\n' +
      'Only missing sheets will be created.\n\nContinue?',
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;
  }

  var created = [];
  var skipped = [];

  // --- Settings ---
  if (setupSettingsSheet_(ss)) {
    created.push('Settings');
  } else {
    skipped.push('Settings');
  }

  // --- Clients ---
  if (setupClientsSheet_(ss)) {
    created.push('Clients');
  } else {
    skipped.push('Clients');
  }

  // --- WorkLog ---
  if (setupWorkLogSheet_(ss)) {
    created.push('WorkLog');
  } else {
    skipped.push('WorkLog');
  }

  // --- Invoices ---
  if (setupInvoicesSheet_(ss)) {
    created.push('Invoices');
  } else {
    skipped.push('Invoices');
  }

  // Delete the default "Sheet1" if it's empty and our sheets exist
  cleanupDefaultSheet_(ss);

  // Report
  var msg = '';
  if (created.length > 0) msg += '✅ Created: ' + created.join(', ') + '\n';
  if (skipped.length > 0) msg += 'ℹ️ Already existed (skipped): ' + skipped.join(', ') + '\n';
  msg += '\nNext steps:\n';
  msg += '1. Fill in your business details in the Settings tab\n';
  msg += '2. Add your clients in the Clients tab\n';
  msg += '3. Log work in WorkLog\n';
  msg += '4. Use 🧾 Invoice menu to generate invoices!\n';
  msg += '\nRefresh the page to update the menu.';

  ui.alert('🧾 Setup Complete', msg, ui.ButtonSet.OK);
}

/**
 * Creates the Settings sheet. Returns true if created, false if skipped.
 */
function setupSettingsSheet_(ss) {
  if (sheetHasData_(ss, SHEET_NAMES.SETTINGS)) return false;

  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.SETTINGS);
  sheet.clear();

  // Headers
  var headers = [['Setting', 'Value']];
  sheet.getRange(1, 1, 1, 2).setValues(headers);

  // Default settings
  var settings = [
    ['BusinessName', 'Your Business Name'],
    ['BusinessAddress', '123 Main Street\nCity, ST 00000'],
    ['BusinessEmail', 'you@example.com'],
    ['LogoURL', ''],
    ['Currency', 'USD'],
    ['InvoicePrefix', 'INV-'],
    ['NextInvoiceNumber', 1],
    ['DefaultDueDays', 30],
    ['InvoiceNotes', 'Thank you for your business!']
  ];
  sheet.getRange(2, 1, settings.length, 2).setValues(settings);

  // Formatting
  var headerRange = sheet.getRange(1, 1, 1, 2);
  headerRange.setFontWeight('bold')
    .setBackground('#F3F3F3')
    .setFontSize(10);

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 400);
  sheet.setFrozenRows(1);

  // Protect setting keys (column A rows 2+) from accidental edits
  sheet.getRange(2, 1, settings.length, 1)
    .setFontWeight('bold')
    .setFontColor('#333333');

  // Highlight editable values
  sheet.getRange(2, 2, settings.length, 1)
    .setBackground('#FFFDE7');

  // Add note on row 1
  sheet.getRange(1, 2).setNote('Edit the values in this column. Do not change the setting names in column A.');

  return true;
}

/**
 * Creates the Clients sheet. Returns true if created, false if skipped.
 */
function setupClientsSheet_(ss) {
  if (sheetHasData_(ss, SHEET_NAMES.CLIENTS)) return false;

  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.CLIENTS);
  sheet.clear();

  // Headers
  var headers = [['ClientID', 'Name', 'Email', 'Address', 'PaymentTerms']];
  sheet.getRange(1, 1, 1, 5).setValues(headers);

  // Sample data
  var samples = [
    ['ACME', 'Acme Corporation', 'billing@acme.com', '456 Oak Avenue\nSan Francisco, CA 94102', 'Net 30'],
    ['GLOBEX', 'Globex Industries', 'ap@globex.com', '789 Pine Street\nLos Angeles, CA 90001', 'Net 15']
  ];
  sheet.getRange(2, 1, samples.length, 5).setValues(samples);

  // Formatting
  var headerRange = sheet.getRange(1, 1, 1, 5);
  headerRange.setFontWeight('bold')
    .setBackground('#F3F3F3')
    .setFontSize(10);

  sheet.setColumnWidth(1, 100);  // ClientID
  sheet.setColumnWidth(2, 200);  // Name
  sheet.setColumnWidth(3, 200);  // Email
  sheet.setColumnWidth(4, 250);  // Address
  sheet.setColumnWidth(5, 150);  // PaymentTerms
  sheet.setFrozenRows(1);

  // Add helper notes
  sheet.getRange(1, 1).setNote('Short ID you\'ll reference in WorkLog (e.g., ACME, GLOBEX). Keep it simple.');
  sheet.getRange(1, 5).setNote('Enter number of days (e.g., 30) or text (e.g., "Net 30"). If numeric, used to calculate due date. Otherwise falls back to DefaultDueDays from Settings.');

  return true;
}

/**
 * Creates the WorkLog sheet. Returns true if created, false if skipped.
 */
function setupWorkLogSheet_(ss) {
  if (sheetHasData_(ss, SHEET_NAMES.WORK_LOG)) return false;

  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.WORK_LOG);
  sheet.clear();

  // Headers
  var headers = [['Date', 'ClientID', 'Description', 'Qty', 'Rate', 'Tax %', 'Status', 'InvoiceID']];
  sheet.getRange(1, 1, 1, 8).setValues(headers);

  // Sample data
  var today = new Date();
  var yesterday = new Date(today.getTime() - 86400000);
  var twoDaysAgo = new Date(today.getTime() - 172800000);

  var samples = [
    [twoDaysAgo, 'ACME', 'Website redesign — homepage', 8, 75.00, 0, 'Unbilled', ''],
    [yesterday, 'ACME', 'Website redesign — about page', 4, 75.00, 0, 'Unbilled', ''],
    [today, 'ACME', 'Website redesign — contact form', 3, 75.00, 0, 'Unbilled', ''],
    [yesterday, 'GLOBEX', 'Logo design — 3 concepts', 1, 500.00, 8.5, 'Unbilled', ''],
    [today, 'GLOBEX', 'Brand guidelines document', 6, 85.00, 8.5, 'Unbilled', '']
  ];
  sheet.getRange(2, 1, samples.length, 8).setValues(samples);

  // Formatting
  var headerRange = sheet.getRange(1, 1, 1, 8);
  headerRange.setFontWeight('bold')
    .setBackground('#F3F3F3')
    .setFontSize(10);

  sheet.setColumnWidth(1, 110);  // Date
  sheet.setColumnWidth(2, 100);  // ClientID
  sheet.setColumnWidth(3, 300);  // Description
  sheet.setColumnWidth(4, 60);   // Qty
  sheet.setColumnWidth(5, 80);   // Rate
  sheet.setColumnWidth(6, 70);   // Tax %
  sheet.setColumnWidth(7, 90);   // Status
  sheet.setColumnWidth(8, 120);  // InvoiceID
  sheet.setFrozenRows(1);

  // Date format
  sheet.getRange(2, 1, 1000, 1).setNumberFormat('yyyy-mm-dd');

  // Number formats
  sheet.getRange(2, 4, 1000, 1).setNumberFormat('0.##');     // Qty
  sheet.getRange(2, 5, 1000, 1).setNumberFormat('#,##0.00'); // Rate
  sheet.getRange(2, 6, 1000, 1).setNumberFormat('0.#');      // Tax %

  // Right-align number columns
  sheet.getRange(2, 4, 1000, 3).setHorizontalAlignment('right');

  // Data validation: Status dropdown
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Unbilled', 'Invoiced'], true)
    .setAllowInvalid(false)
    .setHelpText('Select Unbilled or Invoiced')
    .build();
  sheet.getRange(2, 7, 1000, 1).setDataValidation(statusRule);

  // Conditional formatting: Unbilled = yellow, Invoiced = green
  var unbilledRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Unbilled')
    .setBackground('#FFF9C4')
    .setRanges([sheet.getRange(2, 7, 1000, 1)])
    .build();

  var invoicedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Invoiced')
    .setBackground('#C8E6C9')
    .setRanges([sheet.getRange(2, 7, 1000, 1)])
    .build();

  sheet.setConditionalFormatRules([unbilledRule, invoicedRule]);

  // Protect InvoiceID column with a warning
  sheet.getRange(1, 8).setNote('This column is filled automatically by the invoice generator. Do not edit manually.');

  return true;
}

/**
 * Creates the Invoices sheet. Returns true if created, false if skipped.
 */
function setupInvoicesSheet_(ss) {
  if (sheetHasData_(ss, SHEET_NAMES.INVOICES)) return false;

  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.INVOICES);
  sheet.clear();

  // Headers
  var headers = [['InvoiceID', 'ClientID', 'IssueDate', 'DueDate', 'Total', 'PDF_Link']];
  sheet.getRange(1, 1, 1, 6).setValues(headers);

  // Formatting
  var headerRange = sheet.getRange(1, 1, 1, 6);
  headerRange.setFontWeight('bold')
    .setBackground('#F3F3F3')
    .setFontSize(10);

  sheet.setColumnWidth(1, 120);  // InvoiceID
  sheet.setColumnWidth(2, 100);  // ClientID
  sheet.setColumnWidth(3, 110);  // IssueDate
  sheet.setColumnWidth(4, 110);  // DueDate
  sheet.setColumnWidth(5, 100);  // Total
  sheet.setColumnWidth(6, 350);  // PDF_Link
  sheet.setFrozenRows(1);

  // Date and currency formats
  sheet.getRange(2, 3, 1000, 2).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 5, 1000, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 5, 1000, 1).setHorizontalAlignment('right');

  // Note
  sheet.getRange(1, 1).setNote('This sheet is populated automatically. Each row is one generated invoice with a link to the PDF in Google Drive.');

  return true;
}

// ============================================================
// SETUP HELPERS
// ============================================================

/**
 * Returns true if a sheet exists and has data beyond the header row.
 */
function sheetHasData_(ss, name) {
  var sheet = ss.getSheetByName(name);
  return sheet && sheet.getLastRow() > 1;
}

/**
 * Gets a sheet by name or creates it. Returns the sheet.
 */
function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * Removes the default "Sheet1" if it exists, is empty, and all required sheets exist.
 */
function cleanupDefaultSheet_(ss) {
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (!defaultSheet) return;

  // Only delete if all four tabs exist (so we don't leave the spreadsheet empty)
  var allExist = [SHEET_NAMES.SETTINGS, SHEET_NAMES.CLIENTS, SHEET_NAMES.WORK_LOG, SHEET_NAMES.INVOICES]
    .every(function(name) { return ss.getSheetByName(name) !== null; });

  if (allExist && defaultSheet.getLastRow() <= 1 && defaultSheet.getLastColumn() <= 1) {
    ss.deleteSheet(defaultSheet);
  }
}

// ============================================================
// SETTINGS HELPERS
// ============================================================

/**
 * Reads the Settings sheet into a key-value object.
 * Settings sheet format: Column A = key, Column B = value.
 */
function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) throw new Error('Settings sheet not found. Please create a "Settings" tab.');

  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) { // skip header row
    const key = String(data[i][0]).trim();
    const val = data[i][1];
    if (key) settings[key] = val;
  }

  // Validate required settings
  const required = ['BusinessName', 'BusinessAddress', 'Currency', 'InvoicePrefix', 'NextInvoiceNumber', 'DefaultDueDays'];
  const missing = required.filter(k => !settings[k] && settings[k] !== 0);
  if (missing.length > 0) {
    throw new Error('Missing required settings: ' + missing.join(', '));
  }

  return settings;
}

/**
 * Increments the NextInvoiceNumber in the Settings sheet.
 */
function incrementInvoiceNumber(settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === 'NextInvoiceNumber') {
      sheet.getRange(i + 1, 2).setValue(Number(settings.NextInvoiceNumber) + 1);
      return;
    }
  }
}

// ============================================================
// CLIENT HELPERS
// ============================================================

/**
 * Returns a map of ClientID → client object from the Clients sheet.
 * Clients sheet columns: ClientID, Name, Email, Address, PaymentTerms
 */
function getClientsMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CLIENTS);
  if (!sheet) throw new Error('Clients sheet not found.');

  const data = sheet.getDataRange().getValues();
  const clients = {};

  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]).trim();
    if (id) {
      clients[id] = {
        id: id,
        name: String(data[i][1]).trim(),
        email: String(data[i][2]).trim(),
        address: String(data[i][3]).trim(),
        paymentTerms: String(data[i][4]).trim()
      };
    }
  }
  return clients;
}

// ============================================================
// WORK LOG HELPERS
// ============================================================

/**
 * Returns unbilled work log rows grouped by ClientID.
 * Each entry includes the row index (1-based) for later updating.
 */
function getUnbilledByClient(filterClientId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.WORK_LOG);
  if (!sheet) throw new Error('WorkLog sheet not found.');

  const data = sheet.getDataRange().getValues();
  const grouped = {};

  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][WORK_LOG_COLS.STATUS - 1]).trim();
    if (status !== STATUS.UNBILLED) continue;

    const clientId = String(data[i][WORK_LOG_COLS.CLIENT_ID - 1]).trim();
    if (!clientId) continue;
    if (filterClientId && clientId !== filterClientId) continue;

    if (!grouped[clientId]) grouped[clientId] = [];

    grouped[clientId].push({
      rowIndex: i + 1, // 1-based sheet row
      date: data[i][WORK_LOG_COLS.DATE - 1],
      clientId: clientId,
      description: String(data[i][WORK_LOG_COLS.DESCRIPTION - 1]),
      qty: Number(data[i][WORK_LOG_COLS.QTY - 1]) || 0,
      rate: Number(data[i][WORK_LOG_COLS.RATE - 1]) || 0,
      taxPct: Number(data[i][WORK_LOG_COLS.TAX_PCT - 1]) || 0
    });
  }

  return grouped;
}

/**
 * Marks work log rows as invoiced and writes the InvoiceID.
 */
function markRowsAsInvoiced(rowIndices, invoiceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.WORK_LOG);

  rowIndices.forEach(function(rowIndex) {
    sheet.getRange(rowIndex, WORK_LOG_COLS.STATUS).setValue(STATUS.INVOICED);
    sheet.getRange(rowIndex, WORK_LOG_COLS.INVOICE_ID).setValue(invoiceId);
  });
}

// ============================================================
// INVOICE RECORD
// ============================================================

/**
 * Appends a row to the Invoices sheet.
 */
function recordInvoice(invoiceId, clientId, issueDate, dueDate, total, pdfUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.INVOICES);
  if (!sheet) throw new Error('Invoices sheet not found.');

  sheet.appendRow([invoiceId, clientId, issueDate, dueDate, total, pdfUrl]);
}

// ============================================================
// INVOICE ID FORMATTING
// ============================================================

/**
 * Formats an invoice ID like "INV-000123".
 */
function formatInvoiceId(prefix, number) {
  const padded = String(number).padStart(6, '0');
  return prefix + padded;
}

// ============================================================
// PDF GENERATION
// ============================================================

/**
 * Builds an HTML invoice, converts to PDF via a temporary Google Doc,
 * saves to Drive, and returns the PDF file URL.
 */
function generateInvoicePdf(invoiceId, settings, client, lineItems, issueDate, dueDate, totals) {
  const html = buildInvoiceHtml(invoiceId, settings, client, lineItems, issueDate, dueDate, totals);

  // Create temporary Google Doc
  const doc = DocumentApp.create('_temp_invoice_' + invoiceId);
  const body = doc.getBody();

  // Clear default content and set up the document via HTML
  // Since DocumentApp doesn't support raw HTML, we'll build the doc programmatically
  buildInvoiceDoc(doc, invoiceId, settings, client, lineItems, issueDate, dueDate, totals);

  doc.saveAndClose();

  // Export as PDF
  const docId = doc.getId();
  const pdfBlob = DriveApp.getFileById(docId).getAs('application/pdf');

  // Build folder path: /Invoices/YYYY/MM/
  const folder = getOrCreateFolder(issueDate);

  // File name: INV-000123_ClientName_YYYY-MM-DD.pdf
  const safeName = client.name.replace(/[^a-zA-Z0-9_\-]/g, '_');
  const dateStr = Utilities.formatDate(issueDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const fileName = invoiceId + '_' + safeName + '_' + dateStr + '.pdf';

  pdfBlob.setName(fileName);
  const pdfFile = folder.createFile(pdfBlob);

  // Delete temporary doc
  DriveApp.getFileById(docId).setTrashed(true);

  return pdfFile.getUrl();
}

/**
 * Builds the invoice content inside a Google Doc programmatically.
 */
function buildInvoiceDoc(doc, invoiceId, settings, client, lineItems, issueDate, dueDate, totals) {
  const body = doc.getBody();
  body.clear();

  const tz = Session.getScriptTimeZone();
  const issueDateStr = Utilities.formatDate(issueDate, tz, 'MMMM d, yyyy');
  const dueDateStr = Utilities.formatDate(dueDate, tz, 'MMMM d, yyyy');

  // --- Header: Business info ---
  const header = body.appendParagraph(settings.BusinessName);
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

  if (settings.BusinessAddress) {
    const addrLines = String(settings.BusinessAddress).split('\n');
    addrLines.forEach(function(line) {
      const p = body.appendParagraph(line.trim());
      p.setFontSize(9);
      p.setForegroundColor('#666666');
    });
  }

  if (settings.BusinessEmail) {
    const emailP = body.appendParagraph(settings.BusinessEmail);
    emailP.setFontSize(9);
    emailP.setForegroundColor('#666666');
  }

  // Horizontal rule
  body.appendHorizontalRule();

  // --- Invoice title + metadata ---
  const titleP = body.appendParagraph('INVOICE');
  titleP.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  titleP.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  const metaLines = [
    'Invoice #: ' + invoiceId,
    'Date: ' + issueDateStr,
    'Due: ' + dueDateStr,
    'Currency: ' + settings.Currency
  ];
  metaLines.forEach(function(line) {
    const p = body.appendParagraph(line);
    p.setFontSize(10);
    p.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  });

  body.appendParagraph(''); // spacer

  // --- Bill To ---
  const billTo = body.appendParagraph('Bill To:');
  billTo.setBold(true);
  billTo.setFontSize(10);

  const clientNameP = body.appendParagraph(client.name);
  clientNameP.setFontSize(10);

  if (client.address) {
    const clientAddrLines = String(client.address).split('\n');
    clientAddrLines.forEach(function(line) {
      const p = body.appendParagraph(line.trim());
      p.setFontSize(10);
    });
  }

  if (client.email) {
    const clientEmailP = body.appendParagraph(client.email);
    clientEmailP.setFontSize(10);
    clientEmailP.setForegroundColor('#666666');
  }

  body.appendParagraph(''); // spacer

  // --- Line items table ---
  const tableData = [['Date', 'Description', 'Qty', 'Rate', 'Tax %', 'Amount']];

  lineItems.forEach(function(item) {
    const dateStr = (item.date instanceof Date)
      ? Utilities.formatDate(item.date, tz, 'yyyy-MM-dd')
      : String(item.date);
    const amount = item.qty * item.rate;
    const taxAmount = amount * (item.taxPct / 100);
    const lineTotal = amount + taxAmount;

    tableData.push([
      dateStr,
      item.description,
      String(item.qty),
      formatCurrency(item.rate, settings.Currency),
      item.taxPct > 0 ? String(item.taxPct) + '%' : '—',
      formatCurrency(lineTotal, settings.Currency)
    ]);
  });

  const table = body.appendTable(tableData);

  // Style header row
  const headerRow = table.getRow(0);
  for (let c = 0; c < headerRow.getNumCells(); c++) {
    const cell = headerRow.getCell(c);
    cell.setBackgroundColor('#333333');
    cell.editAsText().setForegroundColor('#FFFFFF').setBold(true).setFontSize(9);
  }

  // Style data rows
  for (let r = 1; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    const bgColor = (r % 2 === 0) ? '#F5F5F5' : '#FFFFFF';
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setBackgroundColor(bgColor);
      cell.editAsText().setFontSize(9);
    }
  }

  body.appendParagraph(''); // spacer

  // --- Totals ---
  const subtotalP = body.appendParagraph('Subtotal: ' + formatCurrency(totals.subtotal, settings.Currency));
  subtotalP.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  subtotalP.setFontSize(10);

  if (totals.totalTax > 0) {
    const taxP = body.appendParagraph('Tax: ' + formatCurrency(totals.totalTax, settings.Currency));
    taxP.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    taxP.setFontSize(10);
  }

  const totalP = body.appendParagraph('Total Due: ' + formatCurrency(totals.total, settings.Currency));
  totalP.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  totalP.setBold(true);
  totalP.setFontSize(13);

  body.appendParagraph(''); // spacer

  // --- Payment terms / notes ---
  if (client.paymentTerms) {
    body.appendHorizontalRule();
    const termsLabel = body.appendParagraph('Payment Terms:');
    termsLabel.setBold(true);
    termsLabel.setFontSize(9);
    const termsP = body.appendParagraph(client.paymentTerms);
    termsP.setFontSize(9);
    termsP.setForegroundColor('#666666');
  }

  if (settings.InvoiceNotes) {
    if (!client.paymentTerms) body.appendHorizontalRule();
    const notesLabel = body.appendParagraph('Notes:');
    notesLabel.setBold(true);
    notesLabel.setFontSize(9);
    const notesP = body.appendParagraph(String(settings.InvoiceNotes));
    notesP.setFontSize(9);
    notesP.setForegroundColor('#666666');
  }
}

/**
 * Gets or creates the folder path /Invoices/YYYY/MM/ in Google Drive.
 */
function getOrCreateFolder(date) {
  const tz = Session.getScriptTimeZone();
  const year = Utilities.formatDate(date, tz, 'yyyy');
  const month = Utilities.formatDate(date, tz, 'MM');

  const root = DriveApp.getRootFolder();

  let invoicesFolder = getSubfolder(root, 'Invoices');
  if (!invoicesFolder) invoicesFolder = root.createFolder('Invoices');

  let yearFolder = getSubfolder(invoicesFolder, year);
  if (!yearFolder) yearFolder = invoicesFolder.createFolder(year);

  let monthFolder = getSubfolder(yearFolder, month);
  if (!monthFolder) monthFolder = yearFolder.createFolder(month);

  return monthFolder;
}

/**
 * Finds a subfolder by name, or returns null.
 */
function getSubfolder(parent, name) {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return null;
}

/**
 * Formats a number as currency. Simple prefix-based formatting.
 */
function formatCurrency(amount, currency) {
  const symbols = {
    'USD': '$', 'EUR': '€', 'GBP': '£', 'CAD': 'CA$', 'AUD': 'A$',
    'JPY': '¥', 'CNY': '¥', 'INR': '₹', 'BRL': 'R$', 'MXN': 'MX$'
  };
  const symbol = symbols[currency] || currency + ' ';
  return symbol + Number(amount).toFixed(2);
}

// ============================================================
// LINE ITEM CALCULATIONS
// ============================================================

function calculateTotals(lineItems) {
  let subtotal = 0;
  let totalTax = 0;

  lineItems.forEach(function(item) {
    const lineAmount = item.qty * item.rate;
    const lineTax = lineAmount * (item.taxPct / 100);
    subtotal += lineAmount;
    totalTax += lineTax;
  });

  return {
    subtotal: subtotal,
    totalTax: totalTax,
    total: subtotal + totalTax
  };
}

// ============================================================
// MAIN GENERATION LOGIC
// ============================================================

/**
 * Generates an invoice for a single client's unbilled work.
 * Returns the invoice ID, or null if no unbilled work found.
 */
function generateInvoiceForClient(clientId, settings, clients, unbilledGroups) {
  const lineItems = unbilledGroups[clientId];
  if (!lineItems || lineItems.length === 0) return null;

  const client = clients[clientId];
  if (!client) throw new Error('Client ID "' + clientId + '" not found in Clients sheet.');

  // Generate invoice ID
  const invoiceId = formatInvoiceId(settings.InvoicePrefix, settings.NextInvoiceNumber);

  // Calculate dates
  const issueDate = new Date();
  const dueDays = Number(client.paymentTerms) || Number(settings.DefaultDueDays) || 30;
  const dueDate = new Date(issueDate.getTime() + dueDays * 24 * 60 * 60 * 1000);

  // Calculate totals
  const totals = calculateTotals(lineItems);

  // Generate PDF
  const pdfUrl = generateInvoicePdf(invoiceId, settings, client, lineItems, issueDate, dueDate, totals);

  // Mark rows as invoiced (idempotency: only marks UNBILLED rows)
  const rowIndices = lineItems.map(function(item) { return item.rowIndex; });
  markRowsAsInvoiced(rowIndices, invoiceId);

  // Record in Invoices sheet
  recordInvoice(invoiceId, clientId, issueDate, dueDate, totals.total, pdfUrl);

  // Increment invoice number
  incrementInvoiceNumber(settings);
  // Reload settings for next call since number changed
  settings.NextInvoiceNumber = Number(settings.NextInvoiceNumber) + 1;

  return {
    invoiceId: invoiceId,
    clientName: client.name,
    total: totals.total,
    pdfUrl: pdfUrl
  };
}

// ============================================================
// MENU ACTIONS
// ============================================================

/**
 * Generates an invoice for the client ID found in the currently selected row.
 */
function generateForSelectedClient() {
  const ui = SpreadsheetApp.getUi();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();

    // Must be on WorkLog sheet
    if (activeSheet.getName() !== SHEET_NAMES.WORK_LOG) {
      ui.alert('⚠️ Navigate to the WorkLog sheet first, then select a row for the client you want to invoice.');
      return;
    }

    const row = ss.getActiveRange().getRow();
    if (row <= 1) {
      ui.alert('⚠️ Select a data row (not the header).');
      return;
    }

    const clientId = String(activeSheet.getRange(row, WORK_LOG_COLS.CLIENT_ID).getValue()).trim();
    if (!clientId) {
      ui.alert('⚠️ No Client ID found in the selected row.');
      return;
    }

    const settings = getSettings();
    const clients = getClientsMap();
    const unbilled = getUnbilledByClient(clientId);

    if (!unbilled[clientId] || unbilled[clientId].length === 0) {
      ui.alert('ℹ️ No unbilled work found for client "' + clientId + '".');
      return;
    }

    const result = generateInvoiceForClient(clientId, settings, clients, unbilled);

    if (result) {
      ui.alert(
        '✅ Invoice Generated',
        'Invoice: ' + result.invoiceId + '\n' +
        'Client: ' + result.clientName + '\n' +
        'Total: ' + formatCurrency(result.total, settings.Currency) + '\n\n' +
        'PDF saved to Google Drive.\n' +
        'Link: ' + result.pdfUrl,
        ui.ButtonSet.OK
      );
    }

  } catch (e) {
    ui.alert('❌ Error', e.message, ui.ButtonSet.OK);
    Logger.log('Invoice generation error: ' + e.message + '\n' + e.stack);
  }
}

/**
 * Generates invoices for ALL clients with unbilled work.
 */
function generateForAllUnbilled() {
  const ui = SpreadsheetApp.getUi();

  try {
    const settings = getSettings();
    const clients = getClientsMap();
    const unbilled = getUnbilledByClient(null); // all clients

    const clientIds = Object.keys(unbilled);
    if (clientIds.length === 0) {
      ui.alert('ℹ️ No unbilled work found for any client.');
      return;
    }

    // Confirm
    const confirm = ui.alert(
      '📋 Generate All Invoices',
      'Found unbilled work for ' + clientIds.length + ' client(s):\n\n' +
      clientIds.map(function(id) {
        const name = clients[id] ? clients[id].name : id;
        return '• ' + name + ' (' + unbilled[id].length + ' items)';
      }).join('\n') + '\n\nProceed?',
      ui.ButtonSet.YES_NO
    );

    if (confirm !== ui.Button.YES) return;

    const results = [];
    clientIds.forEach(function(clientId) {
      const result = generateInvoiceForClient(clientId, settings, clients, unbilled);
      if (result) results.push(result);
    });

    if (results.length > 0) {
      ui.alert(
        '✅ ' + results.length + ' Invoice(s) Generated',
        results.map(function(r) {
          return r.invoiceId + ' — ' + r.clientName + ' — ' + formatCurrency(r.total, settings.Currency);
        }).join('\n') + '\n\nAll PDFs saved to Google Drive /Invoices/ folder.',
        ui.ButtonSet.OK
      );
    }

  } catch (e) {
    ui.alert('❌ Error', e.message, ui.ButtonSet.OK);
    Logger.log('Batch invoice generation error: ' + e.message + '\n' + e.stack);
  }
}

// ============================================================
// UTILITY: Build HTML invoice (kept for potential future use)
// ============================================================

function buildInvoiceHtml(invoiceId, settings, client, lineItems, issueDate, dueDate, totals) {
  const tz = Session.getScriptTimeZone();
  const issueDateStr = Utilities.formatDate(issueDate, tz, 'MMMM d, yyyy');
  const dueDateStr = Utilities.formatDate(dueDate, tz, 'MMMM d, yyyy');

  let rows = '';
  lineItems.forEach(function(item) {
    const dateStr = (item.date instanceof Date)
      ? Utilities.formatDate(item.date, tz, 'yyyy-MM-dd')
      : String(item.date);
    const amount = item.qty * item.rate;
    const taxAmt = amount * (item.taxPct / 100);
    const lineTotal = amount + taxAmt;

    rows += '<tr>' +
      '<td>' + dateStr + '</td>' +
      '<td>' + item.description + '</td>' +
      '<td style="text-align:right">' + item.qty + '</td>' +
      '<td style="text-align:right">' + formatCurrency(item.rate, settings.Currency) + '</td>' +
      '<td style="text-align:right">' + (item.taxPct > 0 ? item.taxPct + '%' : '—') + '</td>' +
      '<td style="text-align:right">' + formatCurrency(lineTotal, settings.Currency) + '</td>' +
      '</tr>';
  });

  return '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Arial,sans-serif;margin:40px;color:#333}' +
    'h1{margin:0;font-size:24px}' +
    '.meta{color:#666;font-size:12px;margin-top:4px}' +
    'hr{border:none;border-top:1px solid #ddd;margin:20px 0}' +
    '.invoice-title{text-align:right;font-size:28px;color:#333;letter-spacing:2px}' +
    '.invoice-meta{text-align:right;font-size:12px;color:#666}' +
    'table{width:100%;border-collapse:collapse;margin-top:20px}' +
    'th{background:#333;color:#fff;padding:8px 12px;text-align:left;font-size:11px}' +
    'td{padding:8px 12px;border-bottom:1px solid #eee;font-size:11px}' +
    'tr:nth-child(even){background:#f9f9f9}' +
    '.totals{text-align:right;margin-top:20px;font-size:13px}' +
    '.total-line{font-size:18px;font-weight:bold;margin-top:8px}' +
    '.footer{margin-top:30px;padding-top:15px;border-top:1px solid #ddd;font-size:10px;color:#999}' +
    '</style></head><body>' +
    '<h1>' + settings.BusinessName + '</h1>' +
    '<div class="meta">' + String(settings.BusinessAddress).replace(/\n/g, '<br>') + '</div>' +
    (settings.BusinessEmail ? '<div class="meta">' + settings.BusinessEmail + '</div>' : '') +
    '<hr>' +
    '<div class="invoice-title">INVOICE</div>' +
    '<div class="invoice-meta">' +
    '#' + invoiceId + '<br>' +
    'Date: ' + issueDateStr + '<br>' +
    'Due: ' + dueDateStr + '<br>' +
    settings.Currency +
    '</div>' +
    '<div style="margin-top:20px"><strong>Bill To:</strong><br>' +
    client.name + '<br>' +
    (client.address ? String(client.address).replace(/\n/g, '<br>') + '<br>' : '') +
    (client.email ? '<span style="color:#666">' + client.email + '</span>' : '') +
    '</div>' +
    '<table><tr><th>Date</th><th>Description</th><th style="text-align:right">Qty</th>' +
    '<th style="text-align:right">Rate</th><th style="text-align:right">Tax</th>' +
    '<th style="text-align:right">Amount</th></tr>' + rows + '</table>' +
    '<div class="totals">' +
    'Subtotal: ' + formatCurrency(totals.subtotal, settings.Currency) + '<br>' +
    (totals.totalTax > 0 ? 'Tax: ' + formatCurrency(totals.totalTax, settings.Currency) + '<br>' : '') +
    '<div class="total-line">Total Due: ' + formatCurrency(totals.total, settings.Currency) + '</div>' +
    '</div>' +
    (client.paymentTerms || settings.InvoiceNotes ?
      '<div class="footer">' +
      (client.paymentTerms ? '<strong>Payment Terms:</strong> ' + client.paymentTerms + '<br>' : '') +
      (settings.InvoiceNotes ? '<strong>Notes:</strong> ' + settings.InvoiceNotes : '') +
      '</div>' : '') +
    '</body></html>';
}
