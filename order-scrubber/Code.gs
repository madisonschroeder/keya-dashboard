/**
 * Keya's Snacks — Order Tracking Email Scrubber
 * ================================================
 * Reads order emails out of the orders inbox (Faire, Airgoods, the online
 * store, wholesale/distributor POs, and direct customer emails), asks Claude
 * to extract structured order data, and appends rows to the "Orders" tab of:
 * https://docs.google.com/spreadsheets/d/173Y576EPf013k1XCsQF0r2IgKM3uzUUHBBKUl8zUC_s
 *
 * SETUP (one time)
 * -----------------
 * 1. Go to script.google.com > New project. Paste this whole file in as Code.gs.
 * 2. Project Settings (gear icon) > Script Properties > add a property:
 *      CLAUDE_API_KEY = <your Anthropic API key>
 * 3. In Gmail, set up a filter that applies a label called "orders" to
 *    incoming order emails (from Faire, Airgoods, the store, distributors,
 *    customers, etc). This script only ever looks at mail labeled "orders" —
 *    it does not scan the whole inbox.
 * 4. Run `setup()` once from the Apps Script editor (select it in the
 *    function dropdown, click Run). It will:
 *      - create the Gmail sub-labels this script uses to track state
 *      - create the "Needs Review" tab on the sheet if it doesn't exist
 *      - install a time-driven trigger to run processOrderEmails() every 15 min
 *    The first run will prompt you to authorize Gmail + Sheets access.
 * 5. Done. New order emails will show up as rows within ~15 minutes.
 *
 * HOW IT DECIDES WHAT TO WRITE
 * -----------------------------
 * - Every message gets one Claude call. Claude returns a JSON object saying
 *   whether it's actually an order, plus the order fields and a list of
 *   line items (one row per SKU).
 * - Rule: REQUIRED FIELDS. If Claude can't find a customer name, a PO/order
 *   number, or at least one line item, the row is NOT written to the Orders
 *   tab — it goes to "Needs Review" instead so nobody has to dig through
 *   Gmail to find out why an order didn't show up.
 * - Rule: DEDUPE BY PO NUMBER. Before writing, the script reads every PO
 *   Number already in the Orders tab. If a parsed PO number already exists,
 *   the email is skipped (logged to Needs Review as a duplicate) so
 *   re-processing or forwarded copies never create double rows.
 * - Distributor POs sent as PDF attachments are read too (first PDF
 *   attachment is sent to Claude alongside the email text).
 */

// ─────────────────────────────────────────────────────────────────────────
// CONFIG — the only section you should need to touch
// ─────────────────────────────────────────────────────────────────────────
var CONFIG = {
  SPREADSHEET_ID: '173Y576EPf013k1XCsQF0r2IgKM3uzUUHBBKUl8zUC_s',
  SHEET_NAME: 'Orders',
  NEEDS_REVIEW_SHEET_NAME: 'Needs Review',

  // Gmail label applied (via a Gmail filter you set up) to incoming order
  // emails. The script only reads mail with this label.
  GMAIL_LABEL: 'orders',

  // Sub-labels the script manages itself to avoid reprocessing a message.
  LABEL_PROCESSED: 'orders/processed',
  LABEL_DUPLICATE: 'orders/duplicate',
  LABEL_NOT_ORDER: 'orders/not-an-order',
  LABEL_NEEDS_REVIEW: 'orders/needs-review',
  LABEL_ERROR: 'orders/error',

  CLAUDE_MODEL: 'claude-haiku-4-5', // swap to 'claude-opus-4-8' if extraction quality needs it
  MAX_THREADS_PER_RUN: 20,
};

var ORDERS_HEADER = [
  'Status', 'Customer', 'PO Number', 'PO Date', 'PO Entered',
  'PO to 3PL Date', 'Requested Delivery Date', 'Actual Delivery Date',
  'SKU Name', 'QTY', 'Method of Delivery', 'Notes',
];

var NEEDS_REVIEW_HEADER = [
  'Date Flagged', 'Reason', 'From', 'Subject', 'Gmail Link', 'Raw Extraction (JSON)',
];

// ─────────────────────────────────────────────────────────────────────────
// One-time setup
// ─────────────────────────────────────────────────────────────────────────
function setup() {
  ensureLabelsExist_();
  getOrCreateSheet_(CONFIG.NEEDS_REVIEW_SHEET_NAME, NEEDS_REVIEW_HEADER);
  ensureOrdersHeader_();
  installTrigger_();
  Logger.log('Setup complete. processOrderEmails() will run every 15 minutes.');
}

function installTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'processOrderEmails') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('processOrderEmails').timeBased().everyMinutes(15).create();
}

// ─────────────────────────────────────────────────────────────────────────
// Main entry point (also runnable manually / from the trigger)
// ─────────────────────────────────────────────────────────────────────────
function processOrderEmails() {
  ensureLabelsExist_();
  ensureOrdersHeader_();

  var query = 'label:' + CONFIG.GMAIL_LABEL +
    ' -label:' + CONFIG.LABEL_PROCESSED +
    ' -label:' + CONFIG.LABEL_DUPLICATE +
    ' -label:' + CONFIG.LABEL_NOT_ORDER +
    ' -label:' + CONFIG.LABEL_NEEDS_REVIEW +
    ' -label:' + CONFIG.LABEL_ERROR;

  var threads = GmailApp.search(query, 0, CONFIG.MAX_THREADS_PER_RUN);
  if (threads.length === 0) return;

  var existingPoNumbers = getExistingPoNumbers_();

  threads.forEach(function (thread) {
    thread.getMessages().forEach(function (message) {
      try {
        processMessage_(message, existingPoNumbers);
      } catch (err) {
        Logger.log('Error processing message ' + message.getId() + ': ' + err);
        addLabel_(thread, CONFIG.LABEL_ERROR);
      }
    });
  });
}

function processMessage_(message, existingPoNumbers) {
  var extraction = callClaudeApi_(message);
  var permalink = 'https://mail.google.com/mail/u/0/#inbox/' + message.getId();
  var thread = message.getThread();

  if (!extraction.is_order) {
    logNeedsReview_('Not an order (' + (extraction.notes || 'no reason given') + ')', message, permalink, extraction);
    addLabel_(thread, CONFIG.LABEL_NOT_ORDER);
    return;
  }

  var missing = [];
  if (!extraction.customer) missing.push('customer');
  if (!extraction.po_number) missing.push('po_number');
  if (!extraction.line_items || extraction.line_items.length === 0) missing.push('line_items');

  if (missing.length > 0) {
    logNeedsReview_('Missing required field(s): ' + missing.join(', '), message, permalink, extraction);
    addLabel_(thread, CONFIG.LABEL_NEEDS_REVIEW);
    return;
  }

  var poNumber = String(extraction.po_number).trim();
  if (existingPoNumbers.has(poNumber)) {
    logNeedsReview_('Duplicate PO Number "' + poNumber + '" — already in Orders tab', message, permalink, extraction);
    addLabel_(thread, CONFIG.LABEL_DUPLICATE);
    return;
  }
  existingPoNumbers.add(poNumber);

  appendOrderRows_(extraction, permalink);
  addLabel_(thread, CONFIG.LABEL_PROCESSED);
}

// ─────────────────────────────────────────────────────────────────────────
// Claude API call
// ─────────────────────────────────────────────────────────────────────────
var SYSTEM_PROMPT = [
  'You extract structured order data from a single email for Keya\'s Snacks,',
  'a potato chip company. The orders inbox receives several different kinds',
  'of email: Faire order notifications, Airgoods order notifications, the',
  'company\'s own online store notifications, wholesale/distributor purchase',
  'orders (sometimes as a PDF attachment, sometimes plain text), and direct,',
  'freeform emails from customers or distributors placing an order.',
  '',
  'Read the email (and any attached PDF) and decide: is this actually a new',
  'order or purchase order? If it is a shipping notification, marketing',
  'email, newsletter, receipt for something unrelated, or any other email',
  'that is not itself placing/confirming a new order, set is_order to false',
  'and explain briefly in notes why.',
  '',
  'If it is an order, extract:',
  '- customer: the customer, buyer, retailer, or distributor name (not Keya\'s',
  '  Snacks itself).',
  '- po_number: the purchase order number, order number, or confirmation',
  '  number. If the source is a platform like Faire that only shows an order',
  '  ID, use that ID.',
  '- po_date: the date the order was placed, as YYYY-MM-DD. Null if not stated.',
  '- requested_delivery_date: a requested/needed-by delivery date if the',
  '  customer specified one, as YYYY-MM-DD. Null if not stated.',
  '- delivery_method: shipping/delivery method or carrier if mentioned',
  '  (e.g. "UPS Ground", "LTL freight", "Faire fulfillment"). Null if not stated.',
  '- line_items: one entry per distinct product/SKU ordered, with the product',
  '  name as written (sku_name) and the quantity ordered (quantity, a number).',
  '- notes: anything else useful for the fulfillment team (shipping address,',
  '  special instructions, ambiguities in the source data). Keep it short.',
  '',
  'Always respond with the JSON object described by the schema. Do not',
  'invent data that is not in the email — use null for anything not present.',
].join('\n');

var ORDER_SCHEMA = {
  type: 'object',
  properties: {
    is_order: { type: 'boolean' },
    customer: { type: ['string', 'null'] },
    po_number: { type: ['string', 'null'] },
    po_date: { type: ['string', 'null'] },
    requested_delivery_date: { type: ['string', 'null'] },
    delivery_method: { type: ['string', 'null'] },
    notes: { type: 'string' },
    line_items: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          sku_name: { type: 'string' },
          quantity: { type: ['number', 'null'] },
        },
        required: ['sku_name', 'quantity'],
        additionalProperties: false,
      },
    },
  },
  required: ['is_order', 'customer', 'po_number', 'po_date', 'requested_delivery_date', 'delivery_method', 'notes', 'line_items'],
  additionalProperties: false,
};

function callClaudeApi_(message) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) throw new Error('CLAUDE_API_KEY script property is not set. See setup instructions at the top of this file.');

  var content = [];
  var pdf = getFirstPdfAttachment_(message);
  if (pdf) {
    content.push({
      type: 'document',
      source: {
        type: 'base64',
        media_type: 'application/pdf',
        data: Utilities.base64Encode(pdf.getBytes()),
      },
    });
  }
  content.push({ type: 'text', text: buildEmailText_(message) });

  var payload = {
    model: CONFIG.CLAUDE_MODEL,
    max_tokens: 2000,
    system: SYSTEM_PROMPT,
    messages: [{ role: 'user', content: content }],
    output_config: { format: { type: 'json_schema', schema: ORDER_SCHEMA } },
  };

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code !== 200) {
    throw new Error('Claude API error ' + code + ': ' + body);
  }

  var json = JSON.parse(body);
  var textBlock = null;
  for (var i = 0; i < json.content.length; i++) {
    if (json.content[i].type === 'text') { textBlock = json.content[i]; break; }
  }
  if (!textBlock) throw new Error('Claude response had no text block: ' + body);

  return JSON.parse(textBlock.text);
}

function buildEmailText_(message) {
  return [
    'From: ' + message.getFrom(),
    'Subject: ' + message.getSubject(),
    'Date: ' + message.getDate(),
    '',
    message.getPlainBody(),
  ].join('\n');
}

function getFirstPdfAttachment_(message) {
  var attachments = message.getAttachments();
  for (var i = 0; i < attachments.length; i++) {
    if (attachments[i].getContentType() === 'application/pdf') return attachments[i];
  }
  return null;
}

// ─────────────────────────────────────────────────────────────────────────
// Sheet writes
// ─────────────────────────────────────────────────────────────────────────
function appendOrderRows_(extraction, permalink) {
  var sheet = getOrCreateSheet_(CONFIG.SHEET_NAME, ORDERS_HEADER);
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var notes = extraction.notes ? extraction.notes + ' (source: ' + permalink + ')' : '(source: ' + permalink + ')';

  extraction.line_items.forEach(function (item) {
    sheet.appendRow([
      'New',                              // Status
      extraction.customer,                // Customer
      extraction.po_number,                // PO Number
      extraction.po_date || '',           // PO Date
      today,                               // PO Entered
      '',                                  // PO to 3PL Date (filled manually)
      extraction.requested_delivery_date || '', // Requested Delivery Date
      '',                                  // Actual Delivery Date (filled manually)
      item.sku_name,                       // SKU Name
      item.quantity,                       // QTY
      extraction.delivery_method || '',   // Method of Delivery
      notes,                               // Notes
    ]);
  });
}

function logNeedsReview_(reason, message, permalink, extraction) {
  var sheet = getOrCreateSheet_(CONFIG.NEEDS_REVIEW_SHEET_NAME, NEEDS_REVIEW_HEADER);
  sheet.appendRow([
    new Date(),
    reason,
    message.getFrom(),
    message.getSubject(),
    permalink,
    JSON.stringify(extraction),
  ]);
}

function getExistingPoNumbers_() {
  var sheet = getOrCreateSheet_(CONFIG.SHEET_NAME, ORDERS_HEADER);
  var lastRow = sheet.getLastRow();
  var set = new Set();
  if (lastRow < 2) return set;

  var poColIndex = ORDERS_HEADER.indexOf('PO Number') + 1;
  var values = sheet.getRange(2, poColIndex, lastRow - 1, 1).getValues();
  values.forEach(function (row) {
    var v = String(row[0] || '').trim();
    if (v) set.add(v);
  });
  return set;
}

function ensureOrdersHeader_() {
  getOrCreateSheet_(CONFIG.SHEET_NAME, ORDERS_HEADER);
}

function getOrCreateSheet_(name, header) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(header);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(header);
  }
  return sheet;
}

// ─────────────────────────────────────────────────────────────────────────
// Gmail labels
// ─────────────────────────────────────────────────────────────────────────
function ensureLabelsExist_() {
  [
    CONFIG.LABEL_PROCESSED,
    CONFIG.LABEL_DUPLICATE,
    CONFIG.LABEL_NOT_ORDER,
    CONFIG.LABEL_NEEDS_REVIEW,
    CONFIG.LABEL_ERROR,
  ].forEach(function (name) {
    if (!GmailApp.getUserLabelByName(name)) GmailApp.createLabel(name);
  });
}

function addLabel_(thread, labelName) {
  var label = GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);
  thread.addLabel(label);
}
