/**
 * DSR Orders backend — one shared Google Sheet + this one Apps Script Web App
 * serves BOTH Prabha's and Arun's order-slip apps (and any future DSR app),
 * scoped per request by a "dsrName" field/param. No claude.ai login required —
 * the DSR apps call this with plain fetch() from a phone browser.
 *
 * SETUP: see "DSR Orders API - Setup Instructions.txt" delivered alongside this
 * file. Short version: create a Google Sheet, Extensions > Apps Script, paste
 * this whole file in as Code.gs, Deploy > New deployment > Web app,
 * "Execute as: Me", "Who has access: Anyone", copy the /exec URL it gives you,
 * and send that URL back so it can be set as ORDERS_API_URL in the DSR apps.
 *
 * Sheet layout (auto-created on first run, in a sheet/tab named "Orders"):
 *   OrderID | DSR Name | Dealer | Delivery Date | Date Key | Month Key |
 *   Total Ltr | Items JSON | Order Text | Created At | Updated At | Deleted |
 *   Created At (IST) | Updated At (IST)
 */

var SHEET_NAME = "Orders";
var HEADERS = [
  "OrderID", "DSR Name", "Dealer", "Delivery Date", "Date Key", "Month Key",
  "Total Ltr", "Items JSON", "Order Text", "Created At", "Updated At", "Deleted",
  "Created At (IST)", "Updated At (IST)"
];
// Fixed 1-based column number of "Deleted" - kept as its own constant (rather
// than computed from HEADERS.length) because two more columns were appended
// AFTER it below (Created At (IST) / Updated At (IST)); if this were still
// derived from HEADERS.length, the delete action further down would start
// flipping the wrong column the moment those got added.
var DELETED_COL = 12;
// Human-readable timestamps are always India time regardless of whatever
// timezone this Apps Script project itself is configured with, so "Created
// At (IST)"/"Updated At (IST)" read correctly no matter what.
var DISPLAY_TIMEZONE = "Asia/Kolkata";

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < HEADERS.length) {
    // A newer version of this script added more columns (e.g. the "(IST)"
    // display columns below) after this Sheet already had rows in it - label
    // just the missing header cells so they're not left blank, without
    // touching any existing data.
    var startCol = sheet.getLastColumn() + 1;
    var missing = HEADERS.slice(sheet.getLastColumn());
    sheet.getRange(1, startCol, 1, missing.length).setValues([missing]);
  }
  return sheet;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Formats a JS Date as an India-time, human-readable string for the
// "Created At (IST)"/"Updated At (IST)" columns - e.g. "02 Sep 2026, 11:46:13
// AM". Purely for a human glancing at the Sheet directly; "Created At"/
// "Updated At" (plain ISO UTC) stay exactly as before and remain the columns
// everything else (sorting in doGet, the app's own order-list display) reads
// - changing THEIR format would silently break both, since ISO strings sort
// correctly as plain text and parse reliably everywhere, which a custom
// human-readable string does not. (Added 2 Sep 2026.)
function formatIst_(date) {
  return Utilities.formatDate(date, DISPLAY_TIMEZONE, "dd MMM yyyy, hh:mm:ss a");
}

// Google Sheets silently auto-converts a plain "2026-09-02"/"2026-09" string
// typed or written into a cell into a REAL date value if it looks like one -
// even though doPost only ever sends plain text. When that happens, reading
// the cell back gives a JS Date object, not the original string, so comparing
// it against the plain-text date/month the app asks for (in doGet, below)
// silently never matches - Today/This month/todaysOrders come back empty
// forever, even though the row is sitting right there in the Sheet. These two
// helpers convert the cell's value back to the expected "yyyy-MM-dd"/"yyyy-MM"
// text whether Sheets stored it as a real date or left it as text, so the
// comparison in doGet always works - for every order already in the Sheet
// too, not just new ones. (Bug found and fixed 2 Sep 2026.)
function normalizeDateKey_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(val || "").trim();
}

function normalizeMonthKey_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM");
  }
  return String(val || "").trim();
}

function readAllRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    rows.push({
      rowIndex: i + 2,
      orderId: String(r[0] || ""),
      dsrName: String(r[1] || ""),
      dealer: String(r[2] || ""),
      deliveryDate: r[3],
      dateKey: normalizeDateKey_(r[4]),
      monthKey: normalizeMonthKey_(r[5]),
      totalLtr: Number(r[6]) || 0,
      itemsJson: String(r[7] || ""),
      orderText: String(r[8] || ""),
      createdAt: r[9],
      updatedAt: r[10],
      deleted: r[11] === true || r[11] === "TRUE" || r[11] === "true",
      createdAtIst: String(r[12] || ""),
      updatedAtIst: String(r[13] || "")
    });
  }
  return rows;
}

/**
 * GET ?dsrName=...&date=YYYY-MM-DD&month=YYYY-MM
 * Returns today's total, this month's total, and the list of today's orders
 * (full detail, for the edit/delete list) for that one DSR only.
 */
function doGet(e) {
  try {
    var params = (e && e.parameter) || {};
    var dsrName = String(params.dsrName || "").trim();
    var dateKey = String(params.date || "").trim();
    var monthKey = String(params.month || "").trim();
    if (!dsrName) return jsonOut_({ ok: false, error: "dsrName required" });

    var rows = readAllRows_(getSheet_());
    var todayLtr = 0;
    var monthLtr = 0;
    var todaysOrders = [];

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (row.deleted) continue;
      if (row.dsrName !== dsrName) continue;
      if (monthKey && row.monthKey === monthKey) {
        monthLtr += row.totalLtr;
      }
      if (dateKey && row.dateKey === dateKey) {
        todayLtr += row.totalLtr;
        var items = [];
        try { items = JSON.parse(row.itemsJson || "[]"); } catch (err) { items = []; }
        todaysOrders.push({
          id: row.orderId,
          dealer: row.dealer,
          deliveryDate: row.deliveryDate,
          dateKey: row.dateKey,
          monthKey: row.monthKey,
          totalLtr: row.totalLtr,
          items: items,
          orderText: row.orderText,
          createdAt: row.createdAt,
          updatedAt: row.updatedAt
        });
      }
    }

    todaysOrders.sort(function (a, b) {
      return String(b.createdAt || "").localeCompare(String(a.createdAt || ""));
    });

    return jsonOut_({ ok: true, todayLtr: todayLtr, monthLtr: monthLtr, todaysOrders: todaysOrders });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

/**
 * POST body (JSON, sent as text/plain to skip CORS preflight — parsed as JSON
 * here regardless of the declared content-type):
 *   { action: "add",    dsrName, dealer, deliveryDate, dateKey, monthKey, totalLtr, items, orderText }
 *   { action: "update", id, dsrName, dealer, deliveryDate, dateKey, monthKey, totalLtr, items, orderText }
 *   { action: "delete", id }
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return jsonOut_({ ok: false, error: "Server busy, try again" });
  }
  try {
    var body;
    try {
      body = JSON.parse(e.postData.contents);
    } catch (err) {
      return jsonOut_({ ok: false, error: "Bad request body" });
    }
    var action = body.action;
    var sheet = getSheet_();

    if (action === "add") {
      var id = Utilities.getUuid();
      var nowDate = new Date();
      var now = nowDate.toISOString();
      var nowIst = formatIst_(nowDate);
      sheet.appendRow([
        id,
        body.dsrName || "",
        body.dealer || "",
        body.deliveryDate || "",
        body.dateKey || "",
        body.monthKey || "",
        Number(body.totalLtr) || 0,
        JSON.stringify(body.items || []),
        body.orderText || "",
        now,
        now,
        false,
        nowIst,
        nowIst
      ]);
      return jsonOut_({ ok: true, id: id, createdAt: now, updatedAt: now });
    }

    if (action === "update") {
      var orderId = body.id;
      if (!orderId) return jsonOut_({ ok: false, error: "id required" });
      var rows = readAllRows_(sheet);
      var target = null;
      for (var i = 0; i < rows.length; i++) {
        if (rows[i].orderId === orderId) { target = rows[i]; break; }
      }
      if (!target) return jsonOut_({ ok: false, error: "Order not found" });
      var now2Date = new Date();
      var now2 = now2Date.toISOString();
      var now2Ist = formatIst_(now2Date);
      sheet.getRange(target.rowIndex, 1, 1, HEADERS.length).setValues([[
        orderId,
        body.dsrName || target.dsrName,
        body.dealer || "",
        body.deliveryDate || "",
        body.dateKey || target.dateKey,
        body.monthKey || target.monthKey,
        Number(body.totalLtr) || 0,
        JSON.stringify(body.items || []),
        body.orderText || "",
        target.createdAt,
        now2,
        false,
        target.createdAtIst,
        now2Ist
      ]]);
      return jsonOut_({ ok: true, id: orderId, updatedAt: now2 });
    }

    if (action === "delete") {
      var delId = body.id;
      if (!delId) return jsonOut_({ ok: false, error: "id required" });
      var rows2 = readAllRows_(sheet);
      var target2 = null;
      for (var j = 0; j < rows2.length; j++) {
        if (rows2[j].orderId === delId) { target2 = rows2[j]; break; }
      }
      if (!target2) return jsonOut_({ ok: false, error: "Order not found" });
      // Soft delete only — flips the "Deleted" column rather than removing the row,
      // so nothing else's row index shifts underneath a concurrent request.
      sheet.getRange(target2.rowIndex, DELETED_COL).setValue(true);
      return jsonOut_({ ok: true, id: delId });
    }

    return jsonOut_({ ok: false, error: "Unknown action" });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
