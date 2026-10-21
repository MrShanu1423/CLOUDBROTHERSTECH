/**
 * CLOUDBROTHERS.TECH — CloudMitra backend
 * ----------------------------------------
 * Two sheets:
 *
 *   1. "Customers" — ONE ROW PER CUSTOMER. Every column (lead details AND
 *      their latest star rating/review) lives in the same row, matched by
 *      phone number. So:
 *        - Someone who fills the callback form gets one row with their details.
 *        - If they (or anyone) later leaves a star rating/review with the SAME
 *          phone number, it's merged into that same row instead of a new one.
 *        - A visitor who leaves a review WITHOUT ever filling the lead form
 *          (no phone, or a phone we haven't seen) gets their own new row with
 *          just the review columns filled in.
 *      This is your CRM view — one line per customer, easy to scan/filter.
 *
 *   2. "Reviews" — an append-only log. EVERY review anyone ever submits gets
 *      its own new row here, even if the same phone number reviews more than
 *      once. This is what the website's testimonials carousel reads from
 *      (?action=reviews), so no review is ever lost or silently overwritten,
 *      and the app + this sheet always stay in sync.
 *
 * SETUP
 * 1. Create (or open) a Google Sheet that will hold your data.
 * 2. Extensions > Apps Script. Delete any starter code and paste this file in.
 * 3. Run `setup` once from the Apps Script editor (select it in the
 *    function dropdown, click ▶ Run). This creates both the "Customers" and
 *    "Reviews" tabs with headers, and asks you to authorize the script —
 *    accept the permissions. (If you're upgrading from an older version of
 *    this script that only had "Customers", just run `setup` again — it's
 *    safe to re-run, it only creates whatever tab is missing.)
 * 4. Deploy > New deployment > select type "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Click Deploy, copy the Web App URL it gives you.
 * 5. In cloud.html, find SHEET_WEBHOOK_URL inside the CloudMitra <script>
 *    block and paste that URL in.
 * 6. Re-deploy (Deploy > Manage deployments > Edit > New version) any
 *    time you change this script, otherwise the live URL keeps running
 *    the old code.
 */

const CUSTOMERS_SHEET_NAME = 'Customers';
const REVIEWS_SHEET_NAME = 'Reviews';

const CUSTOMERS_HEADERS = [
  'Timestamp', 'Name', 'Phone', 'Email', 'City',
  'Service', 'Budget', 'Preferred Time', 'Message',
  'Rating', 'Review Text', 'Review Timestamp'
];

// 1-based column numbers matching CUSTOMERS_HEADERS above, for readability.
const COL = {
  TIMESTAMP: 1, NAME: 2, PHONE: 3, EMAIL: 4, CITY: 5,
  SERVICE: 6, BUDGET: 7, TIME: 8, MESSAGE: 9,
  RATING: 10, REVIEW: 11, REVIEW_TIMESTAMP: 12
};

const REVIEWS_HEADERS = [
  'Timestamp', 'Name', 'Phone', 'City', 'Service', 'Rating', 'Review Text'
];

// 1-based column numbers matching REVIEWS_HEADERS above.
const RCOL = {
  TIMESTAMP: 1, NAME: 2, PHONE: 3, CITY: 4, SERVICE: 5, RATING: 6, REVIEW: 7
};

/** Run this once manually to create/prepare both sheet tabs. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, CUSTOMERS_SHEET_NAME, CUSTOMERS_HEADERS);
  ensureSheet_(ss, REVIEWS_SHEET_NAME, REVIEWS_HEADERS);
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** Finds the row number (1-based, including the header row) of an existing
 *  customer by phone number. Returns -1 if not found or phone is blank. */
function findRowByPhone_(sheet, phone) {
  if (!phone) return -1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const phones = sheet.getRange(2, COL.PHONE, lastRow - 1, 1).getValues();
  for (let i = 0; i < phones.length; i++) {
    if (String(phones[i][0]).trim() !== '' &&
        String(phones[i][0]).trim() === String(phone).trim()) {
      return i + 2; // +2: data starts at row 2, i is 0-based
    }
  }
  return -1;
}

function fillIfBlank_(sheet, row, col, value) {
  if (!value) return;
  const cell = sheet.getRange(row, col);
  if (!cell.getValue()) cell.setValue(value);
}

/**
 * Entry point the website calls. The frontend POSTs JSON like:
 *   { type: 'customer', name, phone, email, city, service, budget, time, message, timestamp }
 *   { type: 'review',   name, phone, city, service, rating, review, timestamp }
 *
 * Customer (lead) submissions: matched/merged into the "Customers" sheet by
 * phone number — one row per customer, columns filled in as data arrives.
 *
 * Review submissions: ALWAYS get their own new row appended to the "Reviews"
 * sheet (full history, nothing overwritten), AND are also merged into that
 * customer's row in "Customers" (so the CRM view shows their latest rating
 * at a glance) if a matching phone number exists there.
 */
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);
    const now = data.timestamp || new Date().toISOString();

    if (data.type === 'review') {
      // 1) Append-only log — every review, every time, its own row.
      const reviewsSheet = ensureSheet_(ss, REVIEWS_SHEET_NAME, REVIEWS_HEADERS);
      reviewsSheet.appendRow([
        now,
        data.name || '',
        data.phone || '',
        data.city || '',
        data.service || '',
        data.rating || '',
        data.review || ''
      ]);

      // 2) Also reflect the latest rating/review on that customer's row in
      //    the CRM sheet, if we recognise their phone number.
      const customersSheet = ensureSheet_(ss, CUSTOMERS_SHEET_NAME, CUSTOMERS_HEADERS);
      const existingRow = findRowByPhone_(customersSheet, data.phone);
      if (existingRow > 0) {
        customersSheet.getRange(existingRow, COL.RATING).setValue(data.rating || '');
        customersSheet.getRange(existingRow, COL.REVIEW).setValue(data.review || '');
        customersSheet.getRange(existingRow, COL.REVIEW_TIMESTAMP).setValue(now);
        // Backfill any identity fields we now have but the row was missing.
        fillIfBlank_(customersSheet, existingRow, COL.NAME, data.name);
        fillIfBlank_(customersSheet, existingRow, COL.CITY, data.city);
        fillIfBlank_(customersSheet, existingRow, COL.SERVICE, data.service);
      } else if (data.phone) {
        // Reviewer gave a phone number we haven't seen before — still worth
        // a row in the CRM sheet, not just the Reviews log.
        const row = new Array(CUSTOMERS_HEADERS.length).fill('');
        row[COL.TIMESTAMP - 1]        = now;
        row[COL.NAME - 1]             = data.name || '';
        row[COL.PHONE - 1]            = data.phone || '';
        row[COL.CITY - 1]             = data.city || '';
        row[COL.SERVICE - 1]          = data.service || '';
        row[COL.RATING - 1]           = data.rating || '';
        row[COL.REVIEW - 1]           = data.review || '';
        row[COL.REVIEW_TIMESTAMP - 1] = now;
        customersSheet.appendRow(row);
      }
      // No phone at all (anonymous review) — it's still safely logged in
      // the Reviews sheet above; nothing to merge into the CRM sheet.

    } else {
      // Lead / customer (callback) submission — CRM sheet only.
      const customersSheet = ensureSheet_(ss, CUSTOMERS_SHEET_NAME, CUSTOMERS_HEADERS);
      const existingRow = findRowByPhone_(customersSheet, data.phone);
      if (existingRow > 0) {
        customersSheet.getRange(existingRow, COL.NAME).setValue(data.name || '');
        customersSheet.getRange(existingRow, COL.EMAIL).setValue(data.email || '');
        customersSheet.getRange(existingRow, COL.CITY).setValue(data.city || '');
        customersSheet.getRange(existingRow, COL.SERVICE).setValue(data.service || '');
        customersSheet.getRange(existingRow, COL.BUDGET).setValue(data.budget || '');
        customersSheet.getRange(existingRow, COL.TIME).setValue(data.time || '');
        customersSheet.getRange(existingRow, COL.MESSAGE).setValue(data.message || '');
      } else {
        const row = new Array(CUSTOMERS_HEADERS.length).fill('');
        row[COL.TIMESTAMP - 1] = now;
        row[COL.NAME - 1]      = data.name || '';
        row[COL.PHONE - 1]     = data.phone || '';
        row[COL.EMAIL - 1]     = data.email || '';
        row[COL.CITY - 1]      = data.city || '';
        row[COL.SERVICE - 1]   = data.service || '';
        row[COL.BUDGET - 1]    = data.budget || '';
        row[COL.TIME - 1]      = data.time || '';
        row[COL.MESSAGE - 1]   = data.message || '';
        customersSheet.appendRow(row);
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * GET endpoint. Two uses:
 *   1. Open the URL directly in a browser -> health check.
 *   2. Called by the website with ?action=reviews -> returns the most recent
 *      4★ and 5★ rows from the "Reviews" sheet (see MIN_RATING_TO_SHOW) as
 *      JSON, so the site can render them in the testimonials carousel. Since
 *      every review is its own row in "Reviews", nothing a customer submits
 *      is ever skipped, no matter how many times that phone number reviews.
 */
const MIN_RATING_TO_SHOW = 4;

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'reviews') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ensureSheet_(ss, REVIEWS_SHEET_NAME, REVIEWS_HEADERS);
    const rows = sheet.getDataRange().getValues();
    rows.shift(); // drop header row

    const reviews = rows
      .map(r => ({
        timestamp: r[RCOL.TIMESTAMP - 1],
        name:      r[RCOL.NAME - 1],
        city:      r[RCOL.CITY - 1],
        service:   r[RCOL.SERVICE - 1],
        rating:    Number(r[RCOL.RATING - 1]),
        review:    r[RCOL.REVIEW - 1]
      }))
      .filter(r => r.name && r.rating >= MIN_RATING_TO_SHOW)
      .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))
      .slice(0, 20); // most recent 20

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', reviews: reviews }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'CloudMitra backend is live' }))
    .setMimeType(ContentService.MimeType.JSON);
}
