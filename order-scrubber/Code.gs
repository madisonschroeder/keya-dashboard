/**
 * Keya's Snacks — Order Tracking Email Scrubber + Faire API Pull
 * ================================================
 * Reads order emails out of the orders inbox (Airgoods, the online store,
 * Cureate, Hungryroot, Rainforest Distribution, KeHE/SPS Commerce, and
 * direct customer emails), asks Claude to extract structured order data,
 * and appends rows to the "Orders" tab of:
 * https://docs.google.com/spreadsheets/d/173Y576EPf013k1XCsQF0r2IgKM3uzUUHBBKUl8zUC_s
 *
 * Faire orders are pulled directly from Faire's own API instead (see the
 * "Faire API pull" section near the bottom of this file) since Faire's
 * order-notification emails weren't arriving reliably. Both pipelines
 * write to the same Orders tab and share the same dedupe/customer/SKU
 * rules, so an order can't get double-entered no matter which path it
 * came through.
 *
 * Business rules (SKU list, customer name aliases, hybrid manual-entry
 * sources, avocado oil flag, status vocabulary, dedupe logic) are ported
 * from the working Tasklet-based system already used for this client.
 *
 * SETUP (one time)
 * -----------------
 * 1. Log into the orders inbox's Google account directly (script triggers
 *    run as whoever authorizes the script, so it must be created from
 *    inside that account, not your own).
 * 2. Go to script.google.com > New project. Paste this whole file in as Code.gs.
 * 3. Project Settings (gear icon) > Script Properties > add:
 *      CLAUDE_API_KEY = <your Anthropic API key>
 *      FAIRE_APPLICATION_ID = <from Faire's Brand Portal / Developer Portal>
 *      FAIRE_APPLICATION_SECRET = <from Faire's Brand Portal / Developer Portal>
 *      FAIRE_ACCESS_TOKEN = <the API key generated via Settings > Integrations
 *        > "Have an unpublished integration?" in your Faire Brand Portal>
 *    (If Faire credentials aren't set yet, pullFaireOrders() just logs a
 *    warning and does nothing — the email pipeline works independently.)
 * 4. Run `setup()` once from the Apps Script editor (select it in the
 *    function dropdown, click Run). It will:
 *      - create the Gmail sub-labels this script uses to track state
 *      - create the "Needs Review" tab on the sheet if it doesn't exist
 *      - install time-driven triggers to run processOrderEmails() AND
 *        pullFaireOrders() at the hours listed in CONFIG.RUN_HOURS below
 *        (currently 8am, 1pm, 5pm, 8pm, in the script's time zone —
 *        Project Settings > time zone)
 *    The first run will prompt you to authorize Gmail + Sheets access.
 * 5. Done. New orders will show up as rows after the next scheduled run.
 *
 * This scans the whole inbox (no Gmail label/filter setup required) since
 * this inbox is mostly order emails already — Claude's is_order check does
 * the filtering instead of a label. If that stops being true and the inbox
 * picks up a lot of unrelated mail, set CONFIG.GMAIL_LABEL below to scope
 * it to a label instead. Only email dated on/after CONFIG.MIN_EMAIL_DATE is
 * looked at — otherwise the scrubber works through the entire historical
 * inbox backlog over time, resurfacing genuinely old (but valid) orders as
 * if they were new. Matches the same 7/1 floor used on the Faire side.
 *
 * HOW IT DECIDES WHAT TO WRITE
 * -----------------------------
 * - Every message gets one Claude call. Claude returns a JSON object saying
 *   whether it's actually an order, plus the order fields and a list of
 *   line items (one row per SKU).
 * - Rule: REQUIRED FIELDS. If Claude can't find a customer name or at least
 *   one line item, the row is NOT written to the Orders tab — it goes to
 *   "Needs Review" instead. A missing PO Number is NOT a blocker on its own
 *   — direct/local customer orders legitimately don't have one.
 * - Rule: HYBRID MANUAL ENTRY. Rainforest Distribution POs (PDF attachment)
 *   and KeHE/SPS Commerce notifications (details behind a portal login)
 *   still get a row — with SKU Name "TBD" and "MANUAL ENTRY NEEDED" in
 *   Notes — rather than being dropped or sent to Needs Review, matching how
 *   the team already works these.
 * - Rule: DEDUPE. Before writing, the script builds a key per existing row
 *   of Customer + PO Number + SKU Name when a PO Number is available
 *   (PO Number is more reliable than date — a BOL/pickup follow-up email
 *   about the same PO often carries a different or missing date than the
 *   original order email), falling back to Customer + PO Date + SKU Name
 *   when there's no PO Number at all (informal direct/local orders). A new
 *   line item matching an existing key is skipped as a duplicate. This
 *   allows multiple SKUs on the same PO to all get written, while still
 *   catching genuine re-processing of the same email/PDF.
 * - Customer names are normalized to canonical names (e.g. "Heritage
 *   Foods" -> "Virginia Heritage") and avocado oil SKUs get an "AVOCADO
 *   OIL" flag prepended to Notes — both editable in the CONFIG section
 *   below as the client's product/customer list evolves.
 * - Distributor POs sent as PDF attachments are read too (first PDF
 *   attachment is sent to Claude alongside the email text) — Claude reads
 *   the PDF directly, falling back to the manual-entry pattern only if the
 *   PDF text isn't legible.
 */

// ─────────────────────────────────────────────────────────────────────────
// CONFIG — the sections you'll actually need to touch over time
// ─────────────────────────────────────────────────────────────────────────
var CONFIG = {
  SPREADSHEET_ID: '173Y576EPf013k1XCsQF0r2IgKM3uzUUHBBKUl8zUC_s',
  SHEET_NAME: 'Orders',
  NEEDS_REVIEW_SHEET_NAME: 'Needs Review',

  // Optional: set this to a Gmail label (e.g. 'orders') to scope processing
  // to only mail with that label. Leave '' to scan the whole inbox.
  GMAIL_LABEL: '',

  // Hard floor, same as FAIRE_MIN_CREATED_AT below — never look at email
  // older than this date. Format must be YYYY/MM/DD (Gmail search syntax).
  // Without this, the scrubber works through the entire historical inbox
  // backlog over time, surfacing genuinely old (but valid) orders as if
  // they were new simply because it had never seen that email before.
  MIN_EMAIL_DATE: '2026/07/01',

  // Sub-labels the script manages itself to avoid reprocessing a message.
  LABEL_PROCESSED: 'orders/processed',
  LABEL_DUPLICATE: 'orders/duplicate',
  LABEL_NOT_ORDER: 'orders/not-an-order',
  LABEL_NEEDS_REVIEW: 'orders/needs-review',
  LABEL_ERROR: 'orders/error',

  CLAUDE_MODEL: 'claude-haiku-4-5', // swap to 'claude-opus-4-8' if extraction quality needs it
  MAX_THREADS_PER_RUN: 40,

  // Hours of the day (24-hour, script's time zone) to run processOrderEmails.
  // One trigger per hour listed here.
  RUN_HOURS: [8, 13, 17, 20],
};

var ORDERS_HEADER = [
  'Status', 'Customer', 'PO Number', 'PO Date', 'PO Entered',
  'PO to 3PL Date', 'Requested Delivery Date', 'Actual Delivery Date',
  'SKU Name', 'QTY', 'Method of Delivery', 'Notes',
];

var NEEDS_REVIEW_HEADER = [
  'Date Flagged', 'Reason', 'From', 'Subject', 'Gmail Link', 'Raw Extraction (JSON)',
];

// Canonical customer names. "Add new mappings as you discover them" — same
// rule as the original Tasklet system. Match is case-insensitive substring.
var CUSTOMER_ALIASES = [
  { match: ['heritage foods', 'heritage'], canonical: 'Virginia Heritage' },
  { match: ['hungry root', 'hungryroot'], canonical: 'Hungryroot' },
  { match: ['gravity brewing'], canonical: 'Final Gravity Brewing' },
];

// SKUs that trigger the "AVOCADO OIL" notes flag.
var AVOCADO_OIL_SKUS = ['bombay spice avo oil', 'black salt avo oil', 'avocado oil'];

// ─────────────────────────────────────────────────────────────────────────
// One-time setup
// ─────────────────────────────────────────────────────────────────────────
function setup() {
  ensureLabelsExist_();
  getOrCreateSheet_(CONFIG.NEEDS_REVIEW_SHEET_NAME, NEEDS_REVIEW_HEADER);
  ensureOrdersHeader_();
  installTrigger_();
  Logger.log('Setup complete. processOrderEmails() and pullFaireOrders() will run at hours: ' + CONFIG.RUN_HOURS.join(', ') + ' (script time zone).');
}

function installTrigger_() {
  ['processOrderEmails', 'pullFaireOrders'].forEach(function (fn) {
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t);
    });
    CONFIG.RUN_HOURS.forEach(function (hour) {
      ScriptApp.newTrigger(fn).timeBased().atHour(hour).everyDays(1).create();
    });
  });
}

// ─────────────────────────────────────────────────────────────────────────
// Main entry point (also runnable manually / from the trigger)
// ─────────────────────────────────────────────────────────────────────────
function processOrderEmails() {
  ensureLabelsExist_();
  ensureOrdersHeader_();

  var query = (CONFIG.GMAIL_LABEL ? 'label:' + CONFIG.GMAIL_LABEL : 'in:inbox') +
    ' after:' + CONFIG.MIN_EMAIL_DATE +
    ' -label:' + CONFIG.LABEL_PROCESSED +
    ' -label:' + CONFIG.LABEL_DUPLICATE +
    ' -label:' + CONFIG.LABEL_NOT_ORDER +
    ' -label:' + CONFIG.LABEL_NEEDS_REVIEW +
    ' -label:' + CONFIG.LABEL_ERROR;

  var threads = GmailApp.search(query, 0, CONFIG.MAX_THREADS_PER_RUN);
  if (threads.length === 0) return;

  var existingOrderKeys = getExistingOrderKeys_();

  threads.forEach(function (thread) {
    thread.getMessages().forEach(function (message) {
      try {
        processMessage_(message, existingOrderKeys);
      } catch (err) {
        Logger.log('Error processing message ' + message.getId() + ': ' + err);
        addLabel_(thread, CONFIG.LABEL_ERROR);
      }
    });
  });
}

function processMessage_(message, existingOrderKeys) {
  var extraction = callClaudeApi_(message);
  var permalink = 'https://mail.google.com/mail/u/0/#inbox/' + message.getId();
  var thread = message.getThread();

  if (!extraction.is_order) {
    logNeedsReview_('Not an order (' + (extraction.notes || 'no reason given') + ')', message, permalink, extraction);
    addLabel_(thread, CONFIG.LABEL_NOT_ORDER);
    return;
  }

  extraction.customer = normalizeCustomerName_(extraction.customer);

  // po_number is deliberately NOT required here: direct/local customer
  // orders (e.g. an informal email placing an order) legitimately have no
  // PO number, and should still be written with a blank PO Number rather
  // than stalling in Needs Review.
  var missing = [];
  if (!extraction.customer) missing.push('customer');
  if (!extraction.line_items || extraction.line_items.length === 0) missing.push('line_items');

  if (missing.length > 0) {
    logNeedsReview_('Missing required field(s): ' + missing.join(', '), message, permalink, extraction);
    addLabel_(thread, CONFIG.LABEL_NEEDS_REVIEW);
    return;
  }

  // Dedupe per line item: Customer + PO Date + (SKU Name, or PO Number when
  // the SKU is a manual-entry placeholder). This lets multiple genuine SKUs
  // on one PO all get written, while still catching re-processed emails.
  // Keys are NOT added to existingOrderKeys until after a successful write
  // below — otherwise a failed write would still "reserve" the key, making
  // every later mention of that same order silently look like a duplicate
  // instead of surfacing the real failure.
  var newItems = extraction.line_items.filter(function (item) {
    var key = orderKey_(extraction.customer, extraction.po_date, extraction.po_number, item.sku_name);
    return !existingOrderKeys.has(key);
  });

  if (newItems.length === 0) {
    logNeedsReview_('Duplicate — all line items already in Orders tab (PO ' + extraction.po_number + ')', message, permalink, extraction);
    addLabel_(thread, CONFIG.LABEL_DUPLICATE);
    return;
  }

  appendOrderRows_(extraction, newItems, permalink);

  newItems.forEach(function (item) {
    existingOrderKeys.add(orderKey_(extraction.customer, extraction.po_date, extraction.po_number, item.sku_name));
  });

  addLabel_(thread, CONFIG.LABEL_PROCESSED);
}

function orderKey_(customer, poDate, poNumber, skuName) {
  var isPlaceholder = !skuName || skuName.toUpperCase() === 'TBD';

  // PO Number is a much more reliable identifier than date when it's
  // available — a follow-up email about the same PO (a BOL/pickup
  // coordination message, a status check-in) will often carry a different
  // or missing date than the original order email, but the PO Number stays
  // the same. Match on PO Number + SKU in that case, ignoring date.
  if (poNumber) {
    var thirdPart = isPlaceholder ? 'manual-entry' : ('sku:' + skuName);
    return [customer || '', 'po:' + poNumber, thirdPart].join('|').toLowerCase();
  }

  // No PO Number (e.g. an informal direct/local order) — fall back to
  // Customer + PO Date + SKU as the best available proxy for a distinct order.
  var thirdPart = isPlaceholder ? 'manual-entry' : ('sku:' + skuName);
  return [customer || '', poDate || '', thirdPart].join('|').toLowerCase();
}

function normalizeCustomerName_(rawName) {
  if (!rawName) return rawName;
  var lower = rawName.toLowerCase();
  for (var i = 0; i < CUSTOMER_ALIASES.length; i++) {
    var alias = CUSTOMER_ALIASES[i];
    for (var j = 0; j < alias.match.length; j++) {
      if (lower.indexOf(alias.match[j]) !== -1) return alias.canonical;
    }
  }
  return rawName;
}

function isAvocadoOilSku_(skuName) {
  if (!skuName) return false;
  var lower = skuName.toLowerCase();
  return AVOCADO_OIL_SKUS.some(function (needle) { return lower.indexOf(needle) !== -1; });
}

// ─────────────────────────────────────────────────────────────────────────
// Claude API call
// ─────────────────────────────────────────────────────────────────────────
var SYSTEM_PROMPT = [
  'You extract structured order data from a single email for Keya\'s Snacks,',
  'a potato chip company. The orders inbox receives several different kinds',
  'of email:',
  '- Direct customer/distributor emails placing an order',
  '- Airgoods order notifications (retailer name is the customer, NOT',
  '  "Airgoods"). Faire orders are NOT handled here — they come in through',
  '  a separate direct API pull, so a Faire email in this inbox is almost',
  '  always marketing, not an order.',
  '- The company\'s own online store (Squarespace) order notifications',
  '- Cureate and Hungryroot order notifications',
  '- Rainforest Distribution purchase orders, usually as a PDF attachment',
  '- KeHE Distributors purchase orders, arriving as an SPS Commerce',
  '  notification email that does NOT contain the actual order details —',
  '  those live behind a portal login you do not have access to',
  '',
  'STEP 1 — Decide if this is actually an order.',
  'IS an order: a customer placing an order, a forwarded purchase order,',
  'orders from Cureate/Hungryroot/Faire/Airgoods, reorders from existing',
  'customers, sample requests with specific quantities.',
  'Is NOT an order: shipping/tracking updates only, payment confirmations,',
  'warehouse communications without a new order, general inquiries with no',
  'order placed, shipping quote requests, marketing/newsletters.',
  'Bills of Lading (BOLs), pickup ("PU") scheduling, freight/dimension',
  'requests, and any other logistics coordination for an ALREADY-PLACED',
  'order are NEVER an order — this is true even when the email restates the',
  'original PO number, SKU names, or quantities as context for scheduling',
  'the pickup/shipment. Restating old order details for logistics purposes',
  'is not the same as placing a new order. If the email is coordinating',
  'how/when something already ordered will ship, set is_order to false.',
  'If it is not an order, set is_order to false and briefly explain why in notes.',
  '',
  'STEP 2 — If it is an order, extract:',
  '- customer: the customer, buyer, retailer, or distributor name (not Keya\'s',
  '  Snacks itself). Use the name as written; do not guess a canonical form.',
  '- po_number: the purchase order number, order number, or confirmation ID.',
  '- po_date: the date the order was placed, as YYYY-MM-DD. Null if not stated.',
  '- requested_delivery_date: a requested/needed-by delivery date if stated,',
  '  as YYYY-MM-DD. Null if not stated.',
  '- delivery_method: prefer one of "LTL", "Drop Ship", or "Local" when you',
  '  can tell; "Local" usually means Richmond, VA area. Airgoods orders',
  '  default to "Drop Ship" unless the email says otherwise. Use "TBD" if',
  '  genuinely unclear.',
  '- line_items: one entry per distinct product/SKU ordered. Products are',
  '  potato chips in 1.5oz and 6oz sizes (flavors: Bombay Spice, Black Salt,',
  '  Golden Ranch) plus 5.5oz avocado oil variants (Bombay Spice Avo Oil,',
  '  Black Salt Avo Oil). Normalize sku_name to one of: "Bombay Spice 6oz",',
  '  "Black Salt 6oz", "Golden Ranch 6oz", "Bombay Spice Avo Oil", "Black',
  '  Salt Avo Oil", "Bombay Spice 1.5oz", "Black Salt 1.5oz", "Golden Ranch',
  '  1.5oz" whenever the email is describing one of these products, even if',
  '  worded differently. Rainforest Distribution emails sometimes reference',
  '  internal item codes instead of names — map them: 180400 = Bombay Spice',
  '  6oz, 180401 = Black Salt 6oz, 180402 = Bombay Spice Avo Oil, 180410 =',
  '  Bombay Spice 1.5oz, 180411 = Black Salt 1.5oz. "Case" is the standard',
  '  unit — quantity is the number of cases.',
  '- notes: PO numbers already captured elsewhere don\'t need repeating;',
  '  use this for special instructions, shipping address, or ambiguities.',
  '',
  'STEP 3 — Hybrid manual-entry sources (Rainforest PDF, KeHE/SPS Commerce).',
  'For a Rainforest Distribution PO, first try to read the attached PDF and',
  'extract real line items exactly like any other order. Only if the PDF',
  'text is not legible (garbled, scanned image you cannot read, etc.) fall',
  'back to: one line item with sku_name "TBD" and quantity null, and explain',
  'in notes: "MANUAL ENTRY NEEDED - Rainforest PDF PO. SKU/QTY in attached',
  'PDF." Still capture customer ("Rainforest Distribution" if no more',
  'specific name given), po_number, and po_date if visible.',
  'For KeHE / SPS Commerce notifications, the order details are never in',
  'the email itself — always use: customer "KeHE Distributors", one line',
  'item with sku_name "TBD" and quantity null, delivery_method "LTL", and',
  'notes: "MANUAL ENTRY NEEDED - KeHE via SPS Commerce. Login to portal for',
  'details." Still capture po_number and po_date if visible in the notification.',
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
    status: { type: 'string', enum: ['New Order', 'Cancelled'] },
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
  required: ['is_order', 'customer', 'po_number', 'po_date', 'requested_delivery_date', 'delivery_method', 'status', 'notes', 'line_items'],
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
function appendOrderRows_(extraction, lineItems, permalink) {
  var sheet = getOrCreateSheet_(CONFIG.SHEET_NAME, ORDERS_HEADER);
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var baseNotes = extraction.notes ? extraction.notes + ' (source: ' + permalink + ')' : '(source: ' + permalink + ')';

  lineItems.forEach(function (item) {
    var notes = isAvocadoOilSku_(item.sku_name) ? 'AVOCADO OIL | ' + baseNotes : baseNotes;
    sheet.appendRow([
      extraction.status || 'New Order',   // Status
      extraction.customer,                 // Customer
      extraction.po_number,                 // PO Number
      extraction.po_date || '',            // PO Date
      today,                                // PO Entered
      '',                                   // PO to 3PL Date (filled manually)
      extraction.requested_delivery_date || '', // Requested Delivery Date
      '',                                   // Actual Delivery Date (filled manually)
      item.sku_name,                        // SKU Name
      (item.quantity === null || item.quantity === undefined) ? 'TBD' : item.quantity, // QTY
      extraction.delivery_method || 'TBD', // Method of Delivery
      notes,                                // Notes
    ]);
  });
}

function logNeedsReview_(reason, message, permalink, extraction) {
  logNeedsReviewGeneric_(reason, message.getFrom(), message.getSubject(), permalink, extraction);
}

function logNeedsReviewGeneric_(reason, fromLabel, subjectLabel, link, extraction) {
  var sheet = getOrCreateSheet_(CONFIG.NEEDS_REVIEW_SHEET_NAME, NEEDS_REVIEW_HEADER);
  sheet.appendRow([
    new Date(),
    reason,
    fromLabel,
    subjectLabel,
    link,
    JSON.stringify(extraction),
  ]);
}

function getExistingOrderKeys_() {
  var sheet = getOrCreateSheet_(CONFIG.SHEET_NAME, ORDERS_HEADER);
  var lastRow = sheet.getLastRow();
  var set = new Set();
  if (lastRow < 2) return set;

  var customerCol = ORDERS_HEADER.indexOf('Customer') + 1;
  var poNumberCol = ORDERS_HEADER.indexOf('PO Number') + 1;
  var poDateCol = ORDERS_HEADER.indexOf('PO Date') + 1;
  var skuCol = ORDERS_HEADER.indexOf('SKU Name') + 1;

  var values = sheet.getRange(2, 1, lastRow - 1, ORDERS_HEADER.length).getValues();
  values.forEach(function (row) {
    var customer = row[customerCol - 1];
    var poNumber = row[poNumberCol - 1];
    var poDate = row[poDateCol - 1];
    var sku = row[skuCol - 1];
    set.add(orderKey_(customer, poDate, poNumber, sku));
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

// ─────────────────────────────────────────────────────────────────────────
// Faire API pull — bypasses Faire's order-notification emails entirely
// (Faire wasn't sending them) by pulling orders directly from Faire's
// External API. Requires three Script Properties, generated via Faire's
// Brand Portal (Settings > Integrations > "Have an unpublished
// integration?"):
//   FAIRE_APPLICATION_ID
//   FAIRE_APPLICATION_SECRET
//   FAIRE_ACCESS_TOKEN
// Reuses the same Orders-tab dedupe, customer aliasing, and avocado oil
// flag as the email pipeline, so a Faire order and an email mentioning the
// same PO number/SKU can never both get written.
// ─────────────────────────────────────────────────────────────────────────
var FAIRE_ORDERS_URL = 'https://www.faire.com/external-api/v2/orders';

// Hard floor: orders originally CREATED before this date are never pulled
// in, no matter what. This is separate from the incremental "what's new"
// cursor (FAIRE_UPDATED_AT_MIN) because Faire's updated_at_min filter
// tracks when an order was last TOUCHED (status change, payout, etc.), not
// when it was placed — an old order that gets its payout processed later
// shows up as "updated today" even though it's months old. Update this
// value if you ever intentionally want to move the floor.
var FAIRE_MIN_CREATED_AT = '2026-07-01';

function pullFaireOrders() {
  var headers = faireAuthHeaders_();
  if (!headers) {
    Logger.log('Faire API not configured — set FAIRE_APPLICATION_ID, FAIRE_APPLICATION_SECRET, and FAIRE_ACCESS_TOKEN script properties to enable.');
    return;
  }

  var props = PropertiesService.getScriptProperties();
  var floorIso = new Date(FAIRE_MIN_CREATED_AT + 'T00:00:00.000Z').toISOString();
  var updatedAtMin = props.getProperty('FAIRE_UPDATED_AT_MIN') || floorIso;
  if (updatedAtMin < floorIso) updatedAtMin = floorIso; // never let a stale cursor go below the floor either
  var existingOrderKeys = getExistingOrderKeys_();
  var maxUpdatedAtSeen = updatedAtMin;

  var cursor = null;
  do {
    var page = fetchFaireOrdersPage_(headers, updatedAtMin, cursor);
    var orders = page.orders || [];

    orders.forEach(function (order) {
      try {
        if (order.created_at && order.created_at < floorIso) return; // hard floor — always skip, regardless of why it showed up
        processFaireOrder_(order, existingOrderKeys);
        if (order.updated_at && order.updated_at > maxUpdatedAtSeen) maxUpdatedAtSeen = order.updated_at;
      } catch (err) {
        Logger.log('Error processing Faire order ' + order.id + ': ' + err);
      }
    });

    cursor = orders.length > 0 ? page.cursor : null;
  } while (cursor);

  props.setProperty('FAIRE_UPDATED_AT_MIN', maxUpdatedAtSeen);
}

function faireAuthHeaders_() {
  var props = PropertiesService.getScriptProperties();
  // .trim() guards against a stray trailing newline/space from copy-pasting
  // out of a browser, which silently breaks Base64 credentials and 401s
  // with no useful error message.
  var applicationId = (props.getProperty('FAIRE_APPLICATION_ID') || '').trim();
  var applicationSecret = (props.getProperty('FAIRE_APPLICATION_SECRET') || '').trim();
  var accessToken = (props.getProperty('FAIRE_ACCESS_TOKEN') || '').trim();
  if (!applicationId || !applicationSecret || !accessToken) return null;

  var credentials = Utilities.base64Encode(applicationId + ':' + applicationSecret);
  return {
    // Sent alongside the documented custom headers below, in case the 401
    // is coming from a gateway/container-level auth check that expects
    // standard HTTP Basic Auth rather than (or in addition to) the custom
    // X-FAIRE-APP-CREDENTIALS header — same underlying value either way.
    'Authorization': 'Basic ' + credentials,
    'X-FAIRE-APP-CREDENTIALS': credentials,
    'X-FAIRE-OAUTH-ACCESS-TOKEN': accessToken,
  };
}

// Diagnostic only — logs metadata about the three Faire script properties
// (length, whether they contain whitespace/newlines) WITHOUT ever logging
// the actual secret values, so the output is safe to paste into chat.
// No trailing underscore on purpose — Apps Script hides "_"-suffixed
// functions from the run dropdown, and this one needs to be run directly.
function debugFaireCredentials() {
  var props = PropertiesService.getScriptProperties();
  ['FAIRE_APPLICATION_ID', 'FAIRE_APPLICATION_SECRET', 'FAIRE_ACCESS_TOKEN'].forEach(function (key) {
    var raw = props.getProperty(key);
    if (raw === null) {
      Logger.log(key + ': NOT SET');
      return;
    }
    var trimmed = raw.trim();
    Logger.log(
      key + ': length=' + raw.length +
      ', trimmedLength=' + trimmed.length +
      ', hasWhitespaceOrNewline=' + (raw !== trimmed) +
      ', startsWith=' + raw.slice(0, 3) + '...'
    );
  });
}

function fetchFaireOrdersPage_(headers, updatedAtMin, cursor) {
  var url = FAIRE_ORDERS_URL + '?limit=50&updated_at_min=' + encodeURIComponent(updatedAtMin) +
    (cursor ? '&cursor=' + encodeURIComponent(cursor) : '');

  var response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: headers,
    muteHttpExceptions: true,
  });

  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code !== 200) {
    throw new Error('Faire API error ' + code + ': ' + body);
  }

  return JSON.parse(body);
}

function processFaireOrder_(order, existingOrderKeys) {
  var customer = normalizeCustomerName_(faireCustomerName_(order));
  var lineItems = (order.items || []).map(function (item) {
    return {
      sku_name: faireSkuToCanonical_(item.product_name, item.variant_name),
      quantity: item.quantity === undefined ? null : item.quantity,
    };
  });

  var extraction = {
    is_order: true,
    customer: customer,
    po_number: order.purchase_order_number || order.display_id,
    po_date: order.created_at ? isoToDate_(order.created_at) : null,
    requested_delivery_date: order.requested_ship_date ? isoToDate_(order.requested_ship_date) :
      (order.ship_after ? isoToDate_(order.ship_after) : null),
    delivery_method: 'Drop Ship',
    status: order.state === 'CANCELED' ? 'Cancelled' : 'New Order',
    notes: order.notes || '',
    line_items: lineItems,
  };

  var subjectLabel = 'Faire order ' + (order.display_id || order.id);
  var link = 'Faire order ' + (order.display_id || order.id) + ' (internal id: ' + order.id + ')';

  var missing = [];
  if (!extraction.customer) missing.push('customer');
  if (extraction.line_items.length === 0) missing.push('line_items');
  if (missing.length > 0) {
    logNeedsReviewGeneric_('Missing required field(s): ' + missing.join(', '), 'Faire API', subjectLabel, link, extraction);
    return;
  }

  var newItems = extraction.line_items.filter(function (item) {
    var key = orderKey_(extraction.customer, extraction.po_date, extraction.po_number, item.sku_name);
    return !existingOrderKeys.has(key);
  });

  if (newItems.length === 0) {
    logNeedsReviewGeneric_('Duplicate — all line items already in Orders tab (PO ' + extraction.po_number + ')', 'Faire API', subjectLabel, link, extraction);
    return;
  }

  appendOrderRows_(extraction, newItems, link);

  newItems.forEach(function (item) {
    existingOrderKeys.add(orderKey_(extraction.customer, extraction.po_date, extraction.po_number, item.sku_name));
  });
}

function faireCustomerName_(order) {
  if (order.address && order.address.company_name) return order.address.company_name;
  if (order.customer && (order.customer.first_name || order.customer.last_name)) {
    return [order.customer.first_name, order.customer.last_name].filter(Boolean).join(' ');
  }
  return null;
}

// Faire returns product_name + variant_name as separate free-text fields
// (e.g. "Bombay Spice Chips" + "6oz") rather than a single SKU-shaped
// string, so this maps them onto the same canonical names the email
// pipeline uses, via simple flavor + size keyword matching.
function faireSkuToCanonical_(productName, variantName) {
  var combined = ((productName || '') + ' ' + (variantName || '')).toLowerCase();

  var flavor = null;
  if (combined.indexOf('bombay') !== -1) flavor = 'Bombay Spice';
  else if (combined.indexOf('black salt') !== -1) flavor = 'Black Salt';
  else if (combined.indexOf('golden ranch') !== -1) flavor = 'Golden Ranch';

  var isAvocadoOil = combined.indexOf('avo oil') !== -1 || combined.indexOf('avocado oil') !== -1;
  if (isAvocadoOil && (flavor === 'Bombay Spice' || flavor === 'Black Salt')) {
    return flavor + ' Avo Oil';
  }

  var size = null;
  if (combined.indexOf('1.5') !== -1) size = '1.5oz';
  else if (combined.indexOf('6oz') !== -1 || combined.indexOf('6 oz') !== -1) size = '6oz';

  if (flavor && size) return flavor + ' ' + size;

  // No confident match — return the raw text so it's still visible on the
  // sheet row rather than silently dropped; worth adding a new keyword
  // rule above if this starts showing up often.
  return (productName || '') + (variantName ? ' - ' + variantName : '') || 'Unknown product';
}

function isoToDate_(isoTimestamp) {
  return isoTimestamp.slice(0, 10); // 'YYYY-MM-DDTHH:mm:ss.sssZ' -> 'YYYY-MM-DD'
}
