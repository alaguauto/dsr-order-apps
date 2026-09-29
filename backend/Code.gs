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

// ===== Live Tally Sync (added 7 Sep 2026) =====
// Lets a local watcher script (tally_live_watcher.py, running on the office
// Tally PC only during business hours) push near-real-time stock/rate/mrp
// and dealer-outstanding updates in here, which every DSR app then polls via
// GET ?type=live (see fetchLiveOverlay_() in each app file). Purely additive
// - everything above this line (the existing order-saving doGet/doPost
// logic) is completely untouched.
//
// Two new Sheet tabs (auto-created on first write, same pattern as
// getSheet_() above): LiveStock and LiveOutstanding. Each holds ONE current
// row per item/dealer name (upserted, never appended-and-left), so a GET
// always returns the latest known state for everything, not just whatever
// changed in the most recent watcher poll.
//
// LIVE_SYNC_KEY is a shared secret so an anonymous POST can't be used to
// stuff arbitrary stock/outstanding data in here - it must exactly match the
// SYNC_KEY baked into tally_live_watcher.py. The GET side (?type=live) is
// deliberately NOT key-gated, matching the existing doGet's own pattern
// (reads need no auth) - it exposes nothing the DSR apps don't already show
// on-screen once they poll it.
var LIVE_SYNC_KEY = "c1a2ecd5610741281838e57e687f47a3a57a7e0d459553e2";
var LIVE_STOCK_SHEET = "LiveStock";
var LIVE_STOCK_HEADERS = ["Item Name", "Stock", "Rate", "MRP", "Updated At"];
var LIVE_OUTSTANDING_SHEET = "LiveOutstanding";
// "Balance" column added 8 Sep 2026 for the "Manage dealers" live-balance
// feature - see handleLiveOutstandingUpdate_()/handleLiveGet_() below.
// "Pending Amount" column added 17 Sep 2026 - BILLCL-based sum of a
// dealer's currently-open bills only (from Tally's own Bills Receivable
// report, the same one that already powers "Outstanding Since"), kept
// separate from "Balance" (the dealer's whole ledger CLOSINGBALANCE) -
// see tally_live_watcher.py's fetch_outstanding() docstring for the full
// reasoning. Trial run on Arun's app only for now.
// "Aged Pending (45d+)" column added 19 Sep 2026 for the payment-reminder
// feature (Arun-only trial) - the portion of Pending Amount whose OWN
// bill is more than 45 days old, computed per-bill in
// tally_live_watcher.py's fetch_outstanding() (see its agedPendingAmount
// docstring section) - never a dealer-wide guess, since a dealer can have
// both old and fresh open bills at once.
// "Band 30-59"/"Band 60-89"/"Band 90+" added 22 Sep 2026, fixing a real
// bug the owner flagged with a real example (BS Enterprises: a ₹33,336
// bill raised TODAY, a ₹13,955 bill 59 days old, a ₹3,000 bill 73 days
// old) - every consumer of "Outstanding Since"/"Pending Amount" until now
// (Needs Attention's stat tiles, the DSR-wise age chart, Dealer Lookup's
// "Outstanding" tile) attributed a dealer's WHOLE Pending Amount to a
// single age bucket based only on their OLDEST bill's date, so BS
// Enterprises showed as "₹50,291 (73d)" everywhere - implying the full
// amount was 73 days overdue, when only ₹3,000 of it actually was. These
// three columns hold the TRUE per-bill split (each bill's own age decides
// its own band, computed in tally_live_watcher.py's fetch_outstanding() -
// see its band3059/band6089/band90plus docstring section), so any display
// that sums them is now age-accurate no matter how many open bills of
// different ages a dealer has.
// "Band 0-29" added 23 Sep 2026 for the dashboard's DSR-wise aging TABLE
// (replacing the DSR-wise aging chart) - a genuinely complete 0-29/30-59/
// 60-89/90+ picture, not just the 30+ "needs attention" slice. Same
// per-bill, debit-only convention as the other three bands.
var LIVE_OUTSTANDING_HEADERS = ["Dealer Name", "Outstanding Since", "Balance", "Pending Amount", "Aged Pending (45d+)", "Band 0-29", "Band 30-59", "Band 60-89", "Band 90+", "Updated At"];
// Added 8 Sep 2026 - the new-dealer auto-add feature. "Key" is
// "<DSR Name>||<Dealer Name>" so the same dealer name can never collide
// across two different DSRs' rows; see handleLiveDealerRosterUpdate_()/
// handleLiveGet_() below and tally_live_watcher.py's compute_eligible_roster().
var LIVE_DEALER_ROSTER_SHEET = "LiveDealerRoster";
var LIVE_DEALER_ROSTER_HEADERS = ["Key", "DSR Name", "Dealer Name", "Updated At"];
// Added 19 Sep 2026 (Arun-only trial) for the "New dealers" feature -
// deliberately NOT derived from the Orders sheet (see doGet()'s comment
// above for why that approach was tried and reverted same day: the Orders
// sheet only goes back to when the app launched, so it can't tell a
// genuinely new dealer from a long-time customer whose app-recorded
// history just happens to start this month). Instead, this sheet only
// ever gets a row written the day tally_live_watcher.py's daily roster
// job (see compute_eligible_roster()/maybe_push_dealer_roster()) sees a
// dealer name that was NEVER in its eligible roster before - a genuine
// "Tally has never called this dealer active or in-debt until today"
// signal. "First Seen Date" is permanent once written (never re-dated on
// a later poll), so the app can filter to "this calendar month" itself.
var LIVE_NEW_DEALERS_SHEET = "LiveNewDealers";
var LIVE_NEW_DEALERS_HEADERS = ["Key", "DSR Name", "Dealer Name", "First Seen Date", "Updated At"];
// Added 12 Sep 2026 - "Collection" (Receipt-voucher total per DSR, see
// tally_live_watcher.py's fetch_collection()). One row per DSR, upserted
// (never appended-and-left) just like the sheets above, so a GET always
// returns the latest known totals, not just whatever changed in the most
// recent watcher poll. No manual input anywhere on this - Target lives as
// a baked-in constant in each app file (COLLECTION_TARGET_RS); this sheet
// only ever carries the live Collected figures from Tally.
// "Today Collected" added 14 Sep 2026, alongside "Month Collected"
// (previously the only figure, under a plain "Collected" header) - the
// combined Sales & Collection WhatsApp report needs a same-day figure,
// not just the month-to-date total the Collection panel itself shows.
var LIVE_COLLECTION_SHEET = "LiveCollection";
// "MB" (dealers who paid this month) added 21 Sep 2026, for the DSR apps'
// own Collection panel (Arun first) - see tally_live_watcher.py's
// fetch_collection() docstring for how it's counted. "Last Month Same
// Date" added the same day - same-date-basis Tally figure (e.g. 1-21
// Aug when today is the 21st), for that panel's own "Last month" chip
// (owner: "do same for collection also", after the Sales-side version).
// "Last Month Full" added the same day, right after - the WHOLE previous
// calendar month's Tally total, for the "Last month" chip's bracket
// figure (owner: "show in bracket the last month total amount").
// "Avg 6mo Collected Rs" (index 6) added 26 Sep 2026 for the "Full
// Petronas This Month Collection" area table's new "6M AVG" column -
// see tally_live_watcher.py's fetch_collection()/dsr_sixmo_totals
// comment. Inserted before "Updated At" (now index 7).
var LIVE_COLLECTION_HEADERS = ["DSR Name", "Today Collected", "Month Collected", "MB", "Last Month Same Date", "Last Month Full", "Avg 6mo Collected Rs", "Updated At"];
// Added 17 Sep 2026 for the dashboard's "Top customers - Collection"
// panel - per-DEALER breakdown of the same Receipt-voucher totals above,
// see handleLiveCollectionByDealerUpdate_()/handleLiveGet_() below and
// tally_live_watcher.py's fetch_collection() docstring.
var LIVE_COLLECTION_BY_DEALER_SHEET = "LiveCollectionByDealer";
// "Avg 6mo Collected Rs" added 23 Sep 2026 for the "this month vs 6-month
// avg" dealer review feature - see tally_live_watcher.py's
// fetch_collection() docstring (avg6moCollectedRs section).
// "Last Month Same Date"/"Last Month Full" added 25 Sep 2026, also for
// the Dealer Review popup (this month / last month / 6-mo avg) - see
// tally_live_watcher.py's fetch_collection() dealer_rec comment. Column
// indices (0-indexed, matching this array's order exactly):
//   0 Dealer Name, 1 Today Collected, 2 Month Collected,
//   3 Last Payment Date, 4 Avg 6mo Collected Rs, 5 Last Month Same Date,
//   6 Last Month Full, 7 Updated At.
// See the collectionByDealer read loop below (~handleLiveGet_) for the
// matching read-side indices.
// "Last Payment Amount" (index 7) added 26 Sep 2026 for the dealer-lookup
// card (owner: "IN LAST PAYMENT (AMOUNT)") - see tally_live_watcher.py's
// fetch_collection()/last_payment_amt_by_dealer comment. Inserted before
// "Updated At" (now index 8) - every read/write site below updated to match.
var LIVE_COLLECTION_BY_DEALER_HEADERS = ["Dealer Name", "Today Collected", "Month Collected", "Last Payment Date", "Avg 6mo Collected Rs", "Last Month Same Date", "Last Month Full", "Last Payment Amount", "Updated At"];

// Added 28 Sep 2026 - Collection Group Summary (Opening/Debit/Credit/
// Closing per Petronas sales-area group, all voucher types - dashboard's
// new expandable "+" panel with a per-dealer drill-down and full-data
// Excel export). Owner's original request: "need collection data as
// downloadable in excel as we did in daybook, but that has to be
// displayed with+ button as expandable and download to in the format of
// attached for all the group shown there, with the total of group should
// be displayed as show in the screenshot, data should be need from this
// month start, and download current month only, data should be saved in
// the same tab for all months" - also "this should be not shown to area
// manager" and "while export in excel full data to be exported".
//
// UNLIKE Day Book (one sheet tab PER CALENDAR MONTH, see
// daybookSheetName_() above), this is deliberately ONE sheet for every
// month, per the owner's own explicit "same tab for all months" wording -
// each row's own "Month" column (yyyy-MM) is what keeps different
// months' figures apart, upserted by a composite Group+Month key ("Group"
// key composite added below), rather than the sheet ever being
// month-scoped itself. See tally_live_watcher.py's fetch_group_summary()/
// _group_summary_key() for the matching watcher-side design and
// dump_group_summary()'s docstring/history for how the underlying
// Opening/Debit/Credit/Closing computation (including the post-dated-
// voucher correction) was validated against the owner's real Tally data
// across 7 diagnostic rounds before this was wired in.
var LIVE_GROUP_SUMMARY_SHEET = "LiveCollectionGroupSummary";
var LIVE_GROUP_SUMMARY_HEADERS = ["Group", "Month", "Opening", "Debit", "Credit", "Closing", "Updated At"];
// Per-dealer drill-down (the "+" expand's own rows) - composite
// Group+Month+Dealer key, same reasoning as LIVE_GROUP_SUMMARY_SHEET
// above, one level more specific.
var LIVE_GROUP_SUMMARY_BY_DEALER_SHEET = "LiveCollectionGroupSummaryByDealer";
var LIVE_GROUP_SUMMARY_BY_DEALER_HEADERS = ["Group", "Month", "Dealer", "Opening", "Debit", "Credit", "Closing", "Updated At"];

// Added 23 Sep 2026 for the fixed, bill-level "60+ days overdue,
// collected this month" stat (replacing the old daily-recomputed
// version) - see tally_live_watcher.py's snapshot_or_get_sixtyplus_baseline()/
// compute_sixtyplus_collected_by_dealer() docstrings for the full design.
// One row per dealer that has at least one bill in this month's frozen
// 60+ baseline; "Baseline Amount" is that frozen total, "Collected" is
// how much of THOSE SPECIFIC bills has since closed/reduced, "Bill Count"
// is how many individual bills make up the baseline (shown so the
// dashboard/DSR-card list can say e.g. "3 bills" alongside the amounts).
var LIVE_SIXTYPLUS_BY_DEALER_SHEET = "LiveSixtyPlusByDealer";
var LIVE_SIXTYPLUS_BY_DEALER_HEADERS = ["Dealer Name", "Baseline Amount", "Collected", "Bill Count", "Updated At"];

// Added 18 Sep 2026 - dashboard-only "actually billed" Sales, from
// Tally's own Sales vouchers (see tally_live_watcher.py's fetch_sales()
// docstring for the full verification history/scoping). Same
// today/month split as Collection, same per-DSR + per-dealer split, plus
// a third per-ITEM breakdown for the "Top 5 items sold" panel - none of
// this ever reaches any DSR app, dashboard-only per the owner's explicit
// scope request 17 Sep 2026 ("only in the dashboard not in dsr app"). A
// Tally-based "MB" column was briefly added here 21 Sep 2026 for the DSR
// apps' own Sales panel, then reverted the same day at the owner's own
// correction - that panel is entirely app-punched, so its MB is now a
// pure Orders-sheet dealer count instead (see doGet()'s
// monthDealerCount below), not this sheet. "Last Month Same Date" added
// 21 Sep 2026 - same-date-basis Tally figure (e.g. 1-21 Aug when today
// is the 21st), for the DSR apps' Sales "Last month" chip (owner
// approved "Same date last month" over the old full-previous-month
// app-punched comparison). "Last Month Full" added the same day, right
// after - the WHOLE previous calendar month's Tally total, for that same
// chip's bracket figure (owner: "show in bracket the last month total
// amount") - a Tally-based twin of the old app-punched
// ordersSummary.lastMonthLtr, used instead of it so the bracket isn't
// stuck at 0 for a month too young to have real app-punched history.
// "Today Sold Rs"/"Month Sold Rs" added 24 Sep 2026 - fixes the DSR apps'
// Collection Health panel, which reads data.sales[dsrName].monthSoldRs
// (added to that panel 23 Sep 2026) but never actually received it - this
// column plus handleLiveSalesUpdate_()/readSalesDsrMap_() below were the
// missing piece. See tally_live_watcher.py's fetch_sales()/diff_sales().
var LIVE_SALES_SHEET = "LiveSales";
// "Avg 6mo Sold Ltr"/"Avg 6mo Sold Rs" (indices 7/8) added 26 Sep 2026 for
// the "Full Petronas This Month Sales" area table's new "6M AVG" column -
// see tally_live_watcher.py's fetch_sales()/dsr_sixmo_totals comment.
// Inserted before "Updated At" (now index 9).
var LIVE_SALES_HEADERS = ["DSR Name", "Today Sold Ltr", "Month Sold Ltr", "Last Month Same Date", "Last Month Full", "Today Sold Rs", "Month Sold Rs", "Avg 6mo Sold Ltr", "Avg 6mo Sold Rs", "Updated At"];
var LIVE_SALES_BY_DEALER_SHEET = "LiveSalesByDealer";
// "Avg 6mo Sold Ltr" added 23 Sep 2026 for the "this month vs 6-month
// avg" dealer review feature - see tally_live_watcher.py's fetch_sales()
// docstring (avg6moLtr section).
// "Last Sale Before Month" added 23 Sep 2026 for "Dealers Reactivated
// This Month" - the dealer's last sale date STRICTLY BEFORE this
// calendar month (see tally_live_watcher.py's fetch_sales(),
// last_sale_before_month_by_dealer) - kept separate from "Last Sale
// Date" (which includes this month) so the dashboard can tell "were they
// dormant coming into this month" apart from "have they sold recently".
// Six more columns added 25 Sep 2026 for the Dealer Review popup (this
// month / last month / 6-mo avg, Ltr AND Rs) - see
// tally_live_watcher.py's fetch_sales() dealer_rec comment/
// diff_sales_by_dealer(). Column indices (0-indexed, matching this
// array's order exactly):
//   0 Dealer Name, 1 Today Sold Ltr, 2 Month Sold Ltr, 3 Last Sale Date,
//   4 Avg 6mo Sold Ltr, 5 Last Sale Before Month,
//   6 Last Month Same Date Ltr, 7 Last Month Full Ltr, 8 Today Sold Rs,
//   9 Month Sold Rs, 10 Last Month Same Date Rs, 11 Last Month Full Rs,
//   12 Avg 6mo Sold Rs, 13 Updated At.
// See the salesByDealer read loop below (~handleLiveGet_) for the
// matching read-side indices.
// "Last Sale Ltr" (index 13) added 26 Sep 2026 for the dealer-lookup card
// (owner: "IN LAST ORDER (HOW MANY LTR GIVEN)") - see
// tally_live_watcher.py's fetch_sales()/last_sale_ltr_by_dealer comment.
// Inserted before "Updated At" (now index 14) - every read/write site
// below updated to match.
var LIVE_SALES_BY_DEALER_HEADERS = ["Dealer Name", "Today Sold Ltr", "Month Sold Ltr", "Last Sale Date", "Avg 6mo Sold Ltr", "Last Sale Before Month", "Last Month Same Date Ltr", "Last Month Full Ltr", "Today Sold Rs", "Month Sold Rs", "Last Month Same Date Rs", "Last Month Full Rs", "Avg 6mo Sold Rs", "Last Sale Ltr", "Updated At"];
var LIVE_SALES_BY_ITEM_SHEET = "LiveSalesByItem";
var LIVE_SALES_BY_ITEM_HEADERS = ["Item Name", "Today Sold Ltr", "Month Sold Ltr", "Updated At"];

// Added 22 Sep 2026 for the dashboard's "Top selling SKU by DSR" panels
// and "part number covered" per-DSR stat (owner's request). Neither
// LIVE_SALES_SHEET (per-DSR, no item split) nor LIVE_SALES_BY_ITEM_SHEET
// (per-item, company-wide, no DSR split) can answer "what did THIS DSR
// actually sell" - this is the missing DSR x item combination, THIS
// MONTH only (no "today" split - nothing reads a live today-only figure
// here). Keyed by the combined "DSR Name||Item Name" string (see
// tally_live_watcher.py's dsr_item_totals) - bulkUpsertSheet_ just needs
// a unique key string, the two parts don't need their own key column.
var LIVE_SALES_BY_DSR_ITEM_SHEET = "LiveSalesByDsrItem";
var LIVE_SALES_BY_DSR_ITEM_HEADERS = ["DSR Name", "Item Name", "Month Sold Ltr", "Updated At"];

// Added 20 Sep 2026 for the dashboard's Sales+Collection trend chart
// rebuild - the owner asked for the chart on real Tally data, not the
// Orders-sheet/app-logged figures it used before ("i dont want log data
// chart"). Keyed by plain "yyyy-MM-dd" text (Column A, the upsert key -
// see bulkUpsertSheet_()), one row per day, upserted by
// tally_live_watcher.py's fetch_sales()/fetch_collection() every time
// that day's running total changes. Only ever holds this month + last
// month's dates (see _prev_month_range() in tally_live_watcher.py), so
// these sheets stay small - no pruning needed.
var LIVE_SALES_TREND_SHEET = "LiveSalesTrend";
var LIVE_SALES_TREND_HEADERS = ["Date", "Sold Ltr", "Updated At"];
var LIVE_COLLECTION_TREND_SHEET = "LiveCollectionTrend";
var LIVE_COLLECTION_TREND_HEADERS = ["Date", "Collected Rs", "Updated At"];

// Added 20 Sep 2026 for the Collection-to-Sales Ratio feature (owner's
// request: warn when cash collected lags sales BILLED in the same week) -
// a Rupee-value twin of LIVE_SALES_TREND_SHEET above, same shape/upsert
// pattern, own sheet so it's independently diffed from the Ltr-based one.
var LIVE_SALES_RS_TREND_SHEET = "LiveSalesRsTrend";
var LIVE_SALES_RS_TREND_HEADERS = ["Date", "Sold Rs", "Updated At"];

// Added 24 Sep 2026 for the "Ask" smart-question box (owner's request: "CAN
// WE RECORD THE QUE ASKED BY DSR FOR UNDERSTANDING") - every question typed
// into Ask, on any of the 8 DSR apps or the Dashboard, is appended here
// (fire-and-forget, best-effort - a failed log never blocks or breaks the
// answer the DSR/owner actually sees). Plain append-only log, one row per
// question - NOT an upsert like the Live* sheets above, so getLiveSheet_()
// is reused only for its get-or-create-with-self-healing-header behaviour.
// "Understood" is Yes/No depending on whether runAskQuery_() matched a
// known question shape or fell through to askDontUnderstand_() - filtering
// this sheet to "No" rows is the fastest way to see which real phrasings
// Ask still needs to be taught.
var ASK_LOG_SHEET = "AskLog";
var ASK_LOG_HEADERS = ["Timestamp", "Timestamp (IST)", "Source", "DSR Name", "Question", "Understood"];

// Day Book (Sales-PETRONAS/Sales-CBE, ALL dealers - no DSR-group
// filtering, unlike every other Live* sheet above) - added 26 Sep 2026 for
// the dashboard's Excel export button. See tally_live_watcher.py's
// fetch_daybook()/maybe_push_daybook() for the full sync design (GST-
// inclusive Amount, Tally-native Alt Qty text parse, 1 Sep backfill, 3-day
// lookback for late-entered vouchers) and handleLiveDaybookUpdate_()/
// handleDaybookMonthGet_() below for how it's stored and read back.
//
// UNLIKE every sheet above, this is not ONE sheet - it's one sheet PER
// CALENDAR MONTH, auto-created the first time a row for that month
// arrives (daybookSheetName_() below), e.g. "DayBook_2026_09",
// "DayBook_2026_10". This was the owner's own explicit choice (asked via
// AskUserQuestion: "a new sheet tab auto-created each month") - keeps
// each month's Day Book a manageable, independently-scrollable tab rather
// than one ever-growing sheet spanning the whole distributorship's
// history. "Row Key" (voucher#+inventory-line-index, e.g.
// "P-2891/26-27#1") is the upsert key bulkUpsertSheet_() uses - see
// tally_live_watcher.py's _daybook_row_key() docstring for why a voucher
// number alone isn't unique enough.
var DAYBOOK_HEADERS = ["Row Key", "Date", "Voucher Number", "Voucher Type", "Party", "Party Group",
  "Item Name", "Item Group", "Qty", "Alt Qty (Ltr)", "Rate", "Amount (incl GST)", "Updated At"];

// "2026-09" -> "DayBook_2026_09" - Sheets tab names sort/display fine with
// underscores and this matches the existing Live* naming style; a literal
// hyphenated "2026-09" would work too, but every other constant sheet name
// in this file avoids hyphens, so this keeps that convention.
function daybookSheetName_(monthKey) {
  return "DayBook_" + String(monthKey || "").replace(/-/g, "_");
}

function getLiveSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    return sheet;
  }
  // Self-healing header row - added 8 Sep 2026 alongside the new "Balance"
  // column, so a schema change like this one doesn't leave row 1's text out
  // of sync with what the columns actually hold on a sheet that already
  // existed before the change. Only touches row 1, never any data row.
  var currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var headersMatch = headers.every(function (h, i) { return currentHeaders[i] === h; });
  if (!headersMatch) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

// Bulk upsert - added 8 Sep 2026, replacing an earlier per-row version.
// The original upsertLiveRow_() re-read+re-scanned the WHOLE key column
// from scratch for EVERY item in a batch (one getRange().getValues() call
// per item, then a linear scan) - fine for a handful of changed rows each
// poll, but the very first live sync sends a full snapshot (8,809 stock
// items; thousands of dealers), which made this effectively O(n^2) and
// blew straight through the watcher's HTTP timeout every time (confirmed:
// "Read timed out (read timeout=30)" on both of the office's first two
// live poll attempts, 8 Sep 2026). Fixed by reading the sheet's data ONCE,
// merging every update into an in-memory array, and writing the whole
// thing back in ONE setValues() call - O(n) total regardless of batch
// size, and Apps Script's bulk range read/write is fast even for
// thousands of rows (unlike doing it one row-range at a time).
// existingKeyFn - optional, added 24 Sep 2026. Every call site's upsert
// key is column A verbatim EXCEPT LiveSalesByDsrItem, whose real key is
// the composite "DSR Name||Item Name" (columns A+B) - LiveDealerRoster
// and LiveNewDealers face the same composite-key need but solve it by
// writing the composite string INTO column A itself (a dedicated "Key"
// column), so they still match the column-A-only default below.
// LiveSalesByDsrItem never got that "Key" column, and inserting one now
// would reshuffle a sheet that already has months of live data in it -
// so instead this function takes an optional existingKeyFn(existingRow)
// to compute the same composite key from the raw existing row when the
// column-A-only default isn't the sheet's real key. Omit it and behavior
// for every other sheet is unchanged. Without this fix, a keyFn that
// looks for a composite key that never appears in column A means every
// poll's rows fail the "already exists" check and get appended as brand
// new rows forever - see handleLiveSalesByDsrItemUpdate_ below, the
// confirmed root cause of the "same item shown multiple times with
// different Ltr values" bug reported 24 Sep 2026.
function bulkUpsertSheet_(sheet, numCols, updates, keyFn, rowFn, existingKeyFn) {
  var lastRow = sheet.getLastRow();
  var existing = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, numCols).getValues() : [];
  var indexByKey = {};
  for (var i = 0; i < existing.length; i++) {
    var k = existingKeyFn ? String(existingKeyFn(existing[i]) || "").trim() : String(existing[i][0] || "").trim();
    if (k) indexByKey[k] = i;
  }
  var now = new Date().toISOString();
  var written = 0;
  for (var j = 0; j < updates.length; j++) {
    var item = updates[j];
    var key = keyFn(item);
    if (!key) continue;
    var row = rowFn(item, now);
    if (Object.prototype.hasOwnProperty.call(indexByKey, key)) {
      existing[indexByKey[key]] = row;
    } else {
      indexByKey[key] = existing.length;
      existing.push(row);
    }
    written++;
  }
  if (existing.length > 0) {
    sheet.getRange(2, 1, existing.length, numCols).setValues(existing);
  }
  return written;
}

function liveSheetRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
}

// POST { action: "liveStockUpdate", key, items: [{name, stock, rate, mrp}, ...] }
// Only fields actually present (non-null/non-undefined) on each item are
// written - the watcher only ever sends items whose value it detected as
// changed since its last poll (after the very first full-snapshot sync),
// so this only touches those rows in steady state.
function handleLiveStockUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_STOCK_SHEET, LIVE_STOCK_HEADERS);
  var items = body.items || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_STOCK_HEADERS.length, items,
    function (it) { return it && it.name ? String(it.name).trim() : null; },
    function (it, now) {
      return [
        it.name,
        (it.stock === null || it.stock === undefined) ? "" : it.stock,
        (it.rate === null || it.rate === undefined) ? "" : it.rate,
        (it.mrp === null || it.mrp === undefined) ? "" : it.mrp,
        now
      ];
    });
  return jsonOut_({ ok: true, updated: updated });
}

// POST { action: "liveOutstandingUpdate", key, dealers: [{name, outstandingSince, balance}, ...] }
// outstandingSince === "" or null/undefined explicitly means "this dealer is
// no longer outstanding (paid off)" - stored as an empty cell, NOT skipped,
// so a subsequent GET always reports the cleared state instead of silently
// keeping whatever stale date was there before. This is the specific gap the
// office flagged 7 Sep 2026 (Challenger Motors showing an old badge after
// paying) - see tally_live_watcher.py's fetch_outstanding() for how "paid
// off" gets detected from Tally's own bill-wise data.
// "balance" added 8 Sep 2026 for "Manage dealers" - the ledger's current
// CLOSINGBALANCE from Tally (negative = Dr/owed, positive = Cr/credit),
// stored the same "explicit clear beats silent skip" way: 0/null/undefined
// still writes an empty cell rather than leaving a stale figure behind.
function handleLiveOutstandingUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_OUTSTANDING_SHEET, LIVE_OUTSTANDING_HEADERS);
  var dealers = body.dealers || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_OUTSTANDING_HEADERS.length, dealers,
    function (d) { return d && d.name ? String(d.name).trim() : null; },
    function (d, now) {
      return [
        d.name,
        d.outstandingSince || "",
        (d.balance === null || d.balance === undefined) ? "" : d.balance,
        (d.pendingAmount === null || d.pendingAmount === undefined) ? "" : d.pendingAmount,
        (d.agedPendingAmount === null || d.agedPendingAmount === undefined) ? "" : d.agedPendingAmount,
        // "Band 0-29"/"Band 30-59"/"Band 60-89"/"Band 90+" - Band 0-29
        // added 23 Sep 2026, the other three 22 Sep 2026 - see
        // LIVE_OUTSTANDING_HEADERS above for why.
        (d.band0_29 === null || d.band0_29 === undefined) ? "" : d.band0_29,
        (d.band3059 === null || d.band3059 === undefined) ? "" : d.band3059,
        (d.band6089 === null || d.band6089 === undefined) ? "" : d.band6089,
        (d.band90plus === null || d.band90plus === undefined) ? "" : d.band90plus,
        now
      ];
    });
  return jsonOut_({ ok: true, updated: updated });
}

// GET ?type=live - returns the full current LiveStock/LiveOutstanding
// snapshot (not just a delta), so a DSR app polling for the first time (or
// after being closed/reopened) always gets complete current state in one
// call, regardless of how many small incremental POSTs built up to it.
// POST { action: "liveDealerRosterUpdate", key, roster: [{dsrName, name}, ...] }
// Added 8 Sep 2026 - the watcher computes this on a much slower cadence
// than the 10-min stock/outstanding poll (its own voucher-history fetch is
// heavy - see ROSTER_INTERVAL_SECONDS in tally_live_watcher.py), and sends
// the FULL current list of eligible dealers for each of the 8 known DSR
// groups every time it runs, not a diff. Rows are only ever upserted here,
// never deleted - a dealer that briefly stops qualifying just stops being
// re-sent, its row (and the client's copy of it) simply stays as-is rather
// than disappearing, matching "dealers can't be added or removed here"
// already stated in every app's Manage Dealers panel.
function handleLiveDealerRosterUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_DEALER_ROSTER_SHEET, LIVE_DEALER_ROSTER_HEADERS);
  var rows = body.roster || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_DEALER_ROSTER_HEADERS.length, rows,
    function (r) {
      if (!r || !r.dsrName || !r.name) return null;
      return String(r.dsrName).trim() + "||" + String(r.name).trim();
    },
    function (r, now) {
      var key = String(r.dsrName).trim() + "||" + String(r.name).trim();
      return [key, r.dsrName, r.name, now];
    });
  return jsonOut_({ ok: true, updated: updated });
}

// POST { action: "liveNewDealerUpdate", key, newDealers: [{dsrName, name, firstSeenDate}, ...] }
// Added 19 Sep 2026 - see LIVE_NEW_DEALERS_HEADERS above for why this
// exists as its own sheet instead of reusing LiveDealerRoster (that sheet
// re-sends EVERY eligible dealer every day, old and new alike, so it has
// no "when did we first see this one" signal on its own - the watcher does
// that diffing itself and only ever POSTs here the names that are
// genuinely new since the last time it checked). Upserted, never
// re-dated - once a dealer has a First Seen Date here, it stays fixed.
function handleLiveNewDealerUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_NEW_DEALERS_SHEET, LIVE_NEW_DEALERS_HEADERS);
  var rows = body.newDealers || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_NEW_DEALERS_HEADERS.length, rows,
    function (r) {
      if (!r || !r.dsrName || !r.name) return null;
      return String(r.dsrName).trim() + "||" + String(r.name).trim();
    },
    function (r, now) {
      var key = String(r.dsrName).trim() + "||" + String(r.name).trim();
      return [key, r.dsrName, r.name, r.firstSeenDate || "", now];
    });
  return jsonOut_({ ok: true, updated: updated });
}

// POST { action: "liveCollectionUpdate", key, collection: [{dsrName, todayCollected, monthCollected}, ...] }
// Added 12 Sep 2026, split into today/month 14 Sep 2026. One row per DSR,
// upserted (never appended-and-left) - see LIVE_COLLECTION_SHEET above.
// Both figures are Receipt-voucher totals computed by
// tally_live_watcher.py's fetch_collection() - confirmed against real
// Tally data (Arun, 12 Sep 2026) before this was wired in: "Receipt" is
// the exact voucher-type label on this Tally setup, and a dealer's own
// ledger entry on a Receipt voucher comes back POSITIVE, matching the
// negative=Dr/positive=Cr convention already used for CLOSINGBALANCE
// throughout the watcher script.
function handleLiveCollectionUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_COLLECTION_SHEET, LIVE_COLLECTION_HEADERS);
  var rows = body.collection || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_COLLECTION_HEADERS.length, rows,
    function (r) { return r && r.dsrName ? String(r.dsrName).trim() : null; },
    function (r, now) {
      return [
        r.dsrName,
        (r.todayCollected === null || r.todayCollected === undefined) ? "" : r.todayCollected,
        (r.monthCollected === null || r.monthCollected === undefined) ? "" : r.monthCollected,
        // "MB" added 21 Sep 2026 - see LIVE_COLLECTION_HEADERS above.
        (r.mb === null || r.mb === undefined) ? "" : r.mb,
        // "Last Month Same Date" added 21 Sep 2026 - see
        // LIVE_COLLECTION_HEADERS above.
        (r.lastMonthSameDate === null || r.lastMonthSameDate === undefined) ? "" : r.lastMonthSameDate,
        // "Last Month Full" added 21 Sep 2026 - see LIVE_COLLECTION_HEADERS above.
        (r.lastMonthFull === null || r.lastMonthFull === undefined) ? "" : r.lastMonthFull,
        // "Avg 6mo Collected Rs" added 26 Sep 2026 - see LIVE_COLLECTION_HEADERS above.
        (r.avg6moCollectedRs === null || r.avg6moCollectedRs === undefined) ? "" : r.avg6moCollectedRs,
        now
      ];
    });
  return jsonOut_({ ok: true, updated: updated });
}

// POST { action: "liveCollectionByDealerUpdate", key, collection: [{name, todayCollected, monthCollected}, ...] }
// Added 17 Sep 2026 - identical shape/pattern to handleLiveCollectionUpdate_()
// above, just keyed by dealer name (`name`) instead of `dsrName`, upserted
// into its own separate sheet. Feeds the dashboard's "Top customers -
// Collection" panel via handleLiveGet_()'s new collectionByDealer field
// below - the dashboard computes the actual top-5 sort client-side, the
// same way it already does for the DSR ranking panels.
function handleLiveCollectionByDealerUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_COLLECTION_BY_DEALER_SHEET, LIVE_COLLECTION_BY_DEALER_HEADERS);
  var rows = body.collection || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_COLLECTION_BY_DEALER_HEADERS.length, rows,
    function (r) { return r && r.name ? String(r.name).trim() : null; },
    function (r, now) {
      return [
        r.name,
        (r.todayCollected === null || r.todayCollected === undefined) ? "" : r.todayCollected,
        (r.monthCollected === null || r.monthCollected === undefined) ? "" : r.monthCollected,
        // "Last Payment Date" added 20 Sep 2026 for the dealer-lookup
        // card's "Last payment" - a yyyymmdd string from
        // tally_live_watcher.py's fetch_collection(), same plain-text
        // convention as LIVE_SALES_BY_DEALER_HEADERS' "Last Sale Date"
        // right above (read back as plain text too, no Date object).
        (r.lastPaymentDate === null || r.lastPaymentDate === undefined) ? "" : r.lastPaymentDate,
        // "Avg 6mo Collected Rs" added 23 Sep 2026 - see
        // LIVE_COLLECTION_BY_DEALER_HEADERS above.
        (r.avg6moCollectedRs === null || r.avg6moCollectedRs === undefined) ? "" : r.avg6moCollectedRs,
        // "Last Month Same Date"/"Last Month Full" added 25 Sep 2026 for
        // the Dealer Review popup - see LIVE_COLLECTION_BY_DEALER_HEADERS
        // above and tally_live_watcher.py's diff_collection_by_dealer().
        (r.lastMonthSameDate === null || r.lastMonthSameDate === undefined) ? "" : r.lastMonthSameDate,
        (r.lastMonthFull === null || r.lastMonthFull === undefined) ? "" : r.lastMonthFull,
        // "Last Payment Amount" added 26 Sep 2026 - see
        // LIVE_COLLECTION_BY_DEALER_HEADERS above.
        (r.lastPaymentAmount === null || r.lastPaymentAmount === undefined) ? "" : r.lastPaymentAmount,
        now
      ];
    });
  return jsonOut_({ ok: true, updated: updated });
}

// POST { action: "liveGroupSummaryUpdate", key, groups: [{group, monthKey,
// opening, debit, credit, closing}, ...] }
// Added 28 Sep 2026 - see LIVE_GROUP_SUMMARY_SHEET above and
// tally_live_watcher.py's fetch_group_summary()/maybe_push_group_summary()
// for the full design. Keyed by "Group||Month" (composite, columns A+B -
// same existingKeyFn pattern LIVE_SALES_BY_DSR_ITEM_SHEET already uses
// above), so a later month's row for the same group is a brand-new row,
// never overwriting an earlier month's - exactly how "same tab for all
// months" was asked for.
function handleLiveGroupSummaryUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_GROUP_SUMMARY_SHEET, LIVE_GROUP_SUMMARY_HEADERS);
  var rows = body.groups || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_GROUP_SUMMARY_HEADERS.length, rows,
    function (r) { return (r && r.group && r.monthKey) ? (String(r.group).trim() + "||" + String(r.monthKey).trim()) : null; },
    function (r, now) {
      return [
        r.group,
        r.monthKey,
        (r.opening === null || r.opening === undefined) ? "" : r.opening,
        (r.debit === null || r.debit === undefined) ? "" : r.debit,
        (r.credit === null || r.credit === undefined) ? "" : r.credit,
        (r.closing === null || r.closing === undefined) ? "" : r.closing,
        now
      ];
    },
    // existingRow is [Group, Month, Opening, Debit, Credit, Closing,
    // Updated At] - rebuild the same composite key from columns A+B.
    // Fixed 28 Sep 2026 (found the same night as the read-side monthKey
    // fix, via the owner's "some dealer name are in twice" report) -
    // this was a bare String(existingRow[1]).trim(), which breaks the
    // SECOND time a group's row is ever upserted: after the FIRST write,
    // Google Sheets silently auto-converts that Month cell into a real
    // Date value on its own (same bug class as normalizeMonthKey_()'s own
    // comment above), so reading it back with String() no longer
    // reproduces the plain "2026-09" the incoming update's own keyFn
    // computes - the composite key stops matching, and bulkUpsertSheet_
    // appends a brand-new row instead of overwriting the existing one
    // every single push after the first. normalizeMonthKey_() fixes the
    // readback the same way it already fixes handleLiveGet_() above.
    function (existingRow) { return String(existingRow[0] || "").trim() + "||" + normalizeMonthKey_(existingRow[1]); });
  return jsonOut_({ ok: true, updated: updated });
}

// POST { action: "liveGroupSummaryByDealerUpdate", key, dealers: [{group,
// monthKey, dealer, opening, debit, credit, closing}, ...] }
// Added 28 Sep 2026 - same shape/reasoning as handleLiveGroupSummaryUpdate_()
// above, one level more specific (per-dealer rows, feeding the dashboard
// panel's "+" expand and the full-data Excel export). Keyed by
// "Group||Month||Dealer".
function handleLiveGroupSummaryByDealerUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_GROUP_SUMMARY_BY_DEALER_SHEET, LIVE_GROUP_SUMMARY_BY_DEALER_HEADERS);
  var rows = body.dealers || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_GROUP_SUMMARY_BY_DEALER_HEADERS.length, rows,
    function (r) { return (r && r.group && r.monthKey && r.dealer) ? (String(r.group).trim() + "||" + String(r.monthKey).trim() + "||" + String(r.dealer).trim()) : null; },
    function (r, now) {
      return [
        r.group,
        r.monthKey,
        r.dealer,
        (r.opening === null || r.opening === undefined) ? "" : r.opening,
        (r.debit === null || r.debit === undefined) ? "" : r.debit,
        (r.credit === null || r.credit === undefined) ? "" : r.credit,
        (r.closing === null || r.closing === undefined) ? "" : r.closing,
        now
      ];
    },
    // existingRow is [Group, Month, Dealer, Opening, Debit, Credit,
    // Closing, Updated At] - rebuild the same composite key from columns
    // A+B+C. Fixed 28 Sep 2026 - identical root cause and fix as
    // handleLiveGroupSummaryUpdate_()'s own existingKeyFn just above
    // (normalizeMonthKey_() instead of a bare String() on the Month
    // column read back from the Sheet, which Sheets auto-converts to a
    // Date after the first write). THIS is what actually produced the
    // "some dealer name are in twice" symptom the owner reported: every
    // push after a dealer's first one failed to match the existing row
    // (different Group||Month||Dealer key each time, because Month kept
    // round-tripping through String(Date) instead of the plain "yyyy-MM"
    // the new push was keyed on), so a SECOND row was appended instead of
    // the first one being updated - two rows, same dealer, different
    // Credit/Closing values (whichever the watcher's push happened to
    // compute at each of those two sync times).
    function (existingRow) {
      return String(existingRow[0] || "").trim() + "||" + normalizeMonthKey_(existingRow[1]) + "||" + String(existingRow[2] || "").trim();
    });
  return jsonOut_({ ok: true, updated: updated });
}

// POST { action: "liveDaybookUpdate", key, rows: [{monthKey, rowKey, date,
// voucherNumber, voucherType, party, partyGroup, itemName, itemGroup, qty,
// altQtyLtr, rate, amount}, ...] }
// Added 26 Sep 2026 - see DAYBOOK_HEADERS/daybookSheetName_() above and
// tally_live_watcher.py's fetch_daybook()/maybe_push_daybook() for the
// full design. Rows can span more than one month's tab in a single POST
// (the one-time 1 Sep backfill does, by design, and any poll that happens
// to straddle a month boundary would too) - grouped here by monthKey so
// each group only ever touches its own tab, auto-created on first use.
function handleLiveDaybookUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var rows = body.rows || [];
  var byMonth = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var mk = r && r.monthKey;
    if (!mk) continue;
    if (!byMonth[mk]) byMonth[mk] = [];
    byMonth[mk].push(r);
  }
  var totalUpdated = 0;
  var months = Object.keys(byMonth);
  for (var m = 0; m < months.length; m++) {
    var monthKey = months[m];
    var sheet = getLiveSheet_(daybookSheetName_(monthKey), DAYBOOK_HEADERS);
    var updated = bulkUpsertSheet_(sheet, DAYBOOK_HEADERS.length, byMonth[monthKey],
      function (r) { return r && r.rowKey ? String(r.rowKey).trim() : null; },
      function (r, now) {
        return [
          r.rowKey,
          r.date || "",
          r.voucherNumber || "",
          r.voucherType || "",
          r.party || "",
          r.partyGroup || "",
          r.itemName || "",
          r.itemGroup || "",
          (r.qty === null || r.qty === undefined) ? "" : r.qty,
          (r.altQtyLtr === null || r.altQtyLtr === undefined) ? "" : r.altQtyLtr,
          (r.rate === null || r.rate === undefined) ? "" : r.rate,
          (r.amount === null || r.amount === undefined) ? "" : r.amount,
          now
        ];
      });
    totalUpdated += updated;
  }
  return jsonOut_({ ok: true, updated: totalUpdated, months: months });
}

// ---------------------------------------------------------------------
// Monthly card archive (PDF -> Google Drive), added 28 Sep 2026 at the
// owner's request: "need all the cards including all DSR cards to be
// download in pdf format automatically and to save in google drive by
// each month last date by 10 pm."
//
// This endpoint is deliberately dumb - it just takes a FINISHED PDF
// (already rendered from the real, live dashboard by a headless-browser
// script running on the office PC, see monthly_archive.py) and files it
// into Drive under a per-month folder. Doing it this way - capture the
// actual rendered page, not a second server-side copy of the card-
// rendering logic - means this archive can never drift out of sync with
// whatever the dashboard actually looks like; there's nothing here to
// keep updated every time a card's layout changes.
//
// Folder layout, in the DEPLOYING ACCOUNT's own My Drive (this Apps
// Script project already has whatever Drive access its own deploying
// Google account has - no new credentials needed, just a one-time
// permission-review click the first time DriveApp is used - see the
// setup notes sent alongside this file):
//   AAA Monthly Reports/
//     2026-09/
//       DSR - Arun.pdf
//       DSR - Prabha.pdf
//       ... one per DSR ...
//       Company - Full Petronas Sales.pdf
//       Company - Full Petronas Collection.pdf
//       Company - Collection Group Summary.pdf
//       Company - Day Book and Sales by Item Group.pdf
//       Company - Stocks.pdf
//       Combined - All Cards 2026-09.pdf
//
// Re-uploading the same file name for the same month OVERWRITES (deletes
// the old Drive file, writes the fresh one) rather than piling up
// duplicates, so the office-PC script is safe to re-run (e.g. with
// --force, to retry after a partial failure) without littering Drive.
function archiveRootFolder_() {
  return getOrCreateFolder_(DriveApp.getRootFolder(), "AAA Monthly Reports");
}
function getOrCreateFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}
function handleArchiveMonthlyPdf_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var monthKey = String(body.monthKey || "").trim();
  var fileName = String(body.fileName || "").trim();
  var pdfBase64 = body.pdfBase64;
  if (!/^\d{4}-\d{2}$/.test(monthKey)) return jsonOut_({ ok: false, error: "Bad monthKey, expected yyyy-MM" });
  if (!fileName) return jsonOut_({ ok: false, error: "fileName required" });
  if (!pdfBase64) return jsonOut_({ ok: false, error: "pdfBase64 required" });
  if (!/\.pdf$/i.test(fileName)) fileName = fileName + ".pdf";

  var monthFolder = getOrCreateFolder_(archiveRootFolder_(), monthKey);

  // Overwrite semantics - remove any existing file with this exact name
  // in this month's folder first, so a re-run replaces rather than
  // duplicates (bulkUpsertSheet_() elsewhere in this file follows the
  // same "re-sync should overwrite, never pile up" convention).
  var existing = monthFolder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);

  var bytes = Utilities.base64Decode(pdfBase64);
  var blob = Utilities.newBlob(bytes, "application/pdf", fileName);
  var file = monthFolder.createFile(blob);

  return jsonOut_({ ok: true, fileId: file.getId(), url: file.getUrl(), folderUrl: monthFolder.getUrl() });
}

// One-time cleanup, added 26 Sep 2026 at the owner's explicit request
// ("yes cleanout old data and place droping th esales") after the Day
// Book's voucher-type scope was narrowed to Sales-PETRONAS/Sales-CBE only
// (see tally_live_watcher.py's DAYBOOK_VOUCHER_TYPE_NAMES, 26 Sep 2026) -
// that fix only stops NEW plain-"Sales" rows (the unrelated auto-spares
// business) from being synced going forward; it doesn't touch rows
// already written into the Sheet under the old, wider scope before the
// fix landed. GET ?type=daybookCleanupOldSales&key=<LIVE_SYNC_KEY> -
// gated on the same shared secret as every POST action even though this
// rides on doGet, since it deletes real sheet rows; a GET (not POST) so
// it can be run with one visit to a URL in a browser instead of needing
// the Apps Script editor. Scans every "DayBook_*" tab and deletes any row
// whose Voucher Type column is exactly "Sales" (case-insensitive,
// trimmed - never touches "Sales - PETRONAS"/"Sales - CBE" rows, which
// don't match that exact string), bottom-to-top so row indices stay
// valid mid-scan, and reports what it removed per tab so the result can
// be checked without having to open the Sheet.
function handleDaybookCleanupOldSales_(key) {
  if (key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets().filter(function (s) { return /^DayBook_/.test(s.getName()); });
  var VOUCHER_TYPE_COL = DAYBOOK_HEADERS.indexOf("Voucher Type") + 1; // 1-indexed sheet column
  var report = [];
  var totalRemoved = 0;
  sheets.forEach(function (sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) { report.push({ sheet: sheet.getName(), removed: 0 }); return; }
    var values = sheet.getRange(2, VOUCHER_TYPE_COL, lastRow - 1, 1).getValues();
    var removed = 0;
    for (var i = values.length - 1; i >= 0; i--) {
      var v = String(values[i][0] || "").trim().toLowerCase();
      if (v === "sales") {
        sheet.deleteRow(i + 2); // +2: 1-indexed row + header row offset
        removed++;
      }
    }
    totalRemoved += removed;
    report.push({ sheet: sheet.getName(), removed: removed });
  });
  return jsonOut_({ ok: true, totalRemoved: totalRemoved, sheets: report });
}

// GET ?type=daybookMonth&month=2026-09 - returns every row currently
// stored in that month's Day Book tab, for the dashboard's Excel export
// button (which always requests the CURRENT calendar month, per the
// owner's own scope - "current month only"). Read-only, no LIVE_SYNC_KEY
// needed, same as handleLiveGet_()/handleTopSales_() - this is a plain
// GET the dashboard calls directly from a manager's browser.
function handleDaybookMonthGet_(monthKey) {
  if (!monthKey) return jsonOut_({ ok: false, error: "month required" });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(daybookSheetName_(monthKey));
  if (!sheet) return jsonOut_({ ok: true, month: monthKey, rows: [] });
  var raw = liveSheetRows_(sheet);
  var rows = [];
  for (var i = 0; i < raw.length; i++) {
    var r = raw[i];
    if (!r[0]) continue; // blank Row Key = blank trailing row, skip
    // Read-time safety net, added 26 Sep 2026 ("in round 6 daybook shows
    // all the vouchers incl sales") - the write side already stops NEW
    // plain-"Sales" rows (tally_live_watcher.py's DAYBOOK_VOUCHER_TYPE_
    // NAMES, narrowed the same day) and a one-time cleanup endpoint
    // (handleDaybookCleanupOldSales_ above) exists to remove any that
    // were already synced under the old wider scope - but both of those
    // depend on the owner having redeployed the narrowed watcher and/or
    // visited the cleanup URL, and the 1 Sep backfill that ran before
    // either fix landed would have written plain-"Sales" (unrelated
    // auto-spares business) rows straight into the sheet. Rather than
    // keep relying on that sequencing, this read path now ALSO drops any
    // row whose Voucher Type is exactly "Sales" (same case/whitespace-
    // insensitive exact-match rule as the cleanup endpoint - "Sales -
    // PETRONAS"/"Sales - CBE" never match) before it ever reaches the
    // dashboard's Day Book export - so the export is correct immediately,
    // regardless of whether the sheet has been cleaned or the watcher
    // redeployed yet.
    if (String(r[3] || "").trim().toLowerCase() === "sales") continue;
    rows.push({
      date: normalizeDateKey_(r[1]),
      voucherNumber: r[2],
      voucherType: r[3],
      party: r[4],
      partyGroup: r[5],
      itemName: r[6],
      itemGroup: r[7],
      qty: r[8],
      altQtyLtr: r[9],
      rate: r[10],
      amount: r[11]
    });
  }
  // Chronological + stable voucher-number order for the exported sheet -
  // the raw tab's own row order is upsert order, not calendar order, once
  // a late-entered voucher lands (see DAYBOOK_LOOKBACK_DAYS on the Python
  // side) - same "read the Sheet, sort once here" pattern as
  // todaysOrders/historyOrders in doGet() above.
  rows.sort(function (a, b) {
    var byDate = String(a.date || "").localeCompare(String(b.date || ""));
    if (byDate !== 0) return byDate;
    return String(a.voucherNumber || "").localeCompare(String(b.voucherNumber || ""));
  });
  return jsonOut_({ ok: true, month: monthKey, rows: rows });
}

// POST { action: "liveSixtyPlusBaselineUpdate", key, dealers: [{name, baseline, collected, billCount}, ...] }
// Added 23 Sep 2026 - see LIVE_SIXTYPLUS_BY_DEALER_HEADERS above and
// tally_live_watcher.py's snapshot_or_get_sixtyplus_baseline()/
// compute_sixtyplus_collected_by_dealer() docstrings for the full design
// this replaces (the old daily-recomputed sixtyPlusCollectedByDsr_ stat).
function handleLiveSixtyPlusBaselineUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_SIXTYPLUS_BY_DEALER_SHEET, LIVE_SIXTYPLUS_BY_DEALER_HEADERS);
  var rows = body.dealers || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_SIXTYPLUS_BY_DEALER_HEADERS.length, rows,
    function (r) { return r && r.name ? String(r.name).trim() : null; },
    function (r, now) {
      return [
        r.name,
        (r.baseline === null || r.baseline === undefined) ? "" : r.baseline,
        (r.collected === null || r.collected === undefined) ? "" : r.collected,
        (r.billCount === null || r.billCount === undefined) ? "" : r.billCount,
        now
      ];
    });
  return jsonOut_({ ok: true, updated: updated });
}

// POST { action: "liveSalesUpdate", key, sales: [{dsrName, todaySoldLtr, monthSoldLtr}, ...] }
// POST { action: "liveSalesByDealerUpdate", key, sales: [{name, todaySoldLtr, monthSoldLtr}, ...] }
// POST { action: "liveSalesByItemUpdate", key, sales: [{name, todaySoldLtr, monthSoldLtr}, ...] }
// Added 18 Sep 2026 - three handlers, identical bulk-upsert pattern to
// the Collection ones above, for the dashboard-only "actually billed"
// Sales feature (see LIVE_SALES_SHEET/etc. above and
// tally_live_watcher.py's fetch_sales() for the full history).
function handleLiveSalesUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_SALES_SHEET, LIVE_SALES_HEADERS);
  var rows = body.sales || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_SALES_HEADERS.length, rows,
    function (r) { return r && r.dsrName ? String(r.dsrName).trim() : null; },
    function (r, now) {
      return [
        r.dsrName,
        (r.todaySoldLtr === null || r.todaySoldLtr === undefined) ? "" : r.todaySoldLtr,
        (r.monthSoldLtr === null || r.monthSoldLtr === undefined) ? "" : r.monthSoldLtr,
        // "Last Month Same Date" added 21 Sep 2026 - see LIVE_SALES_HEADERS above.
        (r.lastMonthSameDateLtr === null || r.lastMonthSameDateLtr === undefined) ? "" : r.lastMonthSameDateLtr,
        // "Last Month Full" added 21 Sep 2026 - see LIVE_SALES_HEADERS above.
        (r.lastMonthFullLtr === null || r.lastMonthFullLtr === undefined) ? "" : r.lastMonthFullLtr,
        // "Today Sold Rs"/"Month Sold Rs" added 24 Sep 2026 for Collection
        // Health - see LIVE_SALES_HEADERS above.
        (r.todaySoldRs === null || r.todaySoldRs === undefined) ? "" : r.todaySoldRs,
        (r.monthSoldRs === null || r.monthSoldRs === undefined) ? "" : r.monthSoldRs,
        // "Avg 6mo Sold Ltr"/"Avg 6mo Sold Rs" added 26 Sep 2026 - see
        // LIVE_SALES_HEADERS above.
        (r.avg6moLtr === null || r.avg6moLtr === undefined) ? "" : r.avg6moLtr,
        (r.avg6moRs === null || r.avg6moRs === undefined) ? "" : r.avg6moRs,
        now
      ];
    });
  return jsonOut_({ ok: true, updated: updated });
}

function handleLiveSalesByDealerUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_SALES_BY_DEALER_SHEET, LIVE_SALES_BY_DEALER_HEADERS);
  var rows = body.sales || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_SALES_BY_DEALER_HEADERS.length, rows,
    function (r) { return r && r.name ? String(r.name).trim() : null; },
    function (r, now) {
      return [
        r.name,
        (r.todaySoldLtr === null || r.todaySoldLtr === undefined) ? "" : r.todaySoldLtr,
        (r.monthSoldLtr === null || r.monthSoldLtr === undefined) ? "" : r.monthSoldLtr,
        // "Last Sale Date" added 18 Sep 2026 for the "Dormant dealers"
        // panel - a yyyymmdd string from tally_live_watcher.py's
        // fetch_sales(), or blank if no real sale in its 120-day
        // lookback window at all. Kept as plain text (not parsed as a
        // date here) so it round-trips exactly, same reasoning as every
        // other date-shaped field in this project that hit the
        // Sheets-auto-converts-dates bug before (see dsr-app-registry.md's
        // "Seventh bug") - readSalesByDealerMap_ below reads it back as
        // plain text too, no Date object ever touches this column.
        (r.lastSaleDate === null || r.lastSaleDate === undefined) ? "" : r.lastSaleDate,
        // "Avg 6mo Sold Ltr" added 23 Sep 2026 - see
        // LIVE_SALES_BY_DEALER_HEADERS above.
        (r.avg6moLtr === null || r.avg6moLtr === undefined) ? "" : r.avg6moLtr,
        // "Last Sale Before Month" added 23 Sep 2026 - see
        // LIVE_SALES_BY_DEALER_HEADERS above. Same plain-text-date
        // handling as "Last Sale Date" just above (no Date object ever
        // touches this column).
        (r.lastSaleBeforeMonth === null || r.lastSaleBeforeMonth === undefined) ? "" : r.lastSaleBeforeMonth,
        // Six new fields added 25 Sep 2026 for the Dealer Review popup
        // (this month / last month / 6-mo avg, Ltr AND Rs) - see
        // LIVE_SALES_BY_DEALER_HEADERS above for the exact column order
        // these must match, and tally_live_watcher.py's
        // diff_sales_by_dealer() for where they come from.
        (r.lastMonthSameDateLtr === null || r.lastMonthSameDateLtr === undefined) ? "" : r.lastMonthSameDateLtr,
        (r.lastMonthFullLtr === null || r.lastMonthFullLtr === undefined) ? "" : r.lastMonthFullLtr,
        (r.todaySoldRs === null || r.todaySoldRs === undefined) ? "" : r.todaySoldRs,
        (r.monthSoldRs === null || r.monthSoldRs === undefined) ? "" : r.monthSoldRs,
        (r.lastMonthSameDateRs === null || r.lastMonthSameDateRs === undefined) ? "" : r.lastMonthSameDateRs,
        (r.lastMonthFullRs === null || r.lastMonthFullRs === undefined) ? "" : r.lastMonthFullRs,
        (r.avg6moRs === null || r.avg6moRs === undefined) ? "" : r.avg6moRs,
        // "Last Sale Ltr" added 26 Sep 2026 - see LIVE_SALES_BY_DEALER_HEADERS above.
        (r.lastSaleLtr === null || r.lastSaleLtr === undefined) ? "" : r.lastSaleLtr,
        now
      ];
    });
  return jsonOut_({ ok: true, updated: updated });
}

function handleLiveSalesByItemUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_SALES_BY_ITEM_SHEET, LIVE_SALES_BY_ITEM_HEADERS);
  var rows = body.sales || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_SALES_BY_ITEM_HEADERS.length, rows,
    function (r) { return r && r.name ? String(r.name).trim() : null; },
    function (r, now) {
      return [
        r.name,
        (r.todaySoldLtr === null || r.todaySoldLtr === undefined) ? "" : r.todaySoldLtr,
        (r.monthSoldLtr === null || r.monthSoldLtr === undefined) ? "" : r.monthSoldLtr,
        now
      ];
    });
  return jsonOut_({ ok: true, updated: updated });
}

// POST { action: "liveSalesByDsrItemUpdate", key, sales: [{dsrName, itemName, monthSoldLtr}, ...] }
// Added 22 Sep 2026 - see LIVE_SALES_BY_DSR_ITEM_SHEET above and
// tally_live_watcher.py's dsr_item_totals for the full story.
function handleLiveSalesByDsrItemUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_SALES_BY_DSR_ITEM_SHEET, LIVE_SALES_BY_DSR_ITEM_HEADERS);
  var rows = body.sales || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_SALES_BY_DSR_ITEM_HEADERS.length, rows,
    function (r) { return (r && r.dsrName && r.itemName) ? (String(r.dsrName).trim() + "||" + String(r.itemName).trim()) : null; },
    function (r, now) {
      return [
        r.dsrName,
        r.itemName,
        (r.monthSoldLtr === null || r.monthSoldLtr === undefined) ? "" : r.monthSoldLtr,
        now
      ];
    },
    // existingRow is [DSR Name, Item Name, Month Sold Ltr, Updated At] -
    // rebuild the same composite key from columns A+B so existing rows
    // actually match new updates instead of piling up as duplicates.
    function (existingRow) { return String(existingRow[0] || "").trim() + "||" + String(existingRow[1] || "").trim(); });
  return jsonOut_({ ok: true, updated: updated });
}

// POST { action: "liveSalesTrendUpdate", key, trend: [{date: "yyyy-MM-dd", ltr}, ...] }
// POST { action: "liveCollectionTrendUpdate", key, trend: [{date: "yyyy-MM-dd", rs}, ...] }
// Added 20 Sep 2026 - see LIVE_SALES_TREND_SHEET/LIVE_COLLECTION_TREND_SHEET
// above and tally_live_watcher.py's _prev_month_range()/diff_daily_trend()
// for the full story. Same bulk-upsert pattern as every other Live* handler,
// keyed by the date string itself (r.date) instead of a dealer/DSR/item name.
function handleLiveSalesTrendUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_SALES_TREND_SHEET, LIVE_SALES_TREND_HEADERS);
  var rows = body.trend || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_SALES_TREND_HEADERS.length, rows,
    function (r) { return r && r.date ? String(r.date).trim() : null; },
    function (r, now) {
      return [
        r.date,
        (r.ltr === null || r.ltr === undefined) ? "" : r.ltr,
        now
      ];
    });
  return jsonOut_({ ok: true, updated: updated });
}

// POST { action: "liveSalesRsTrendUpdate", key, trend: [{date: "yyyy-MM-dd", rs}, ...] }
// Added 20 Sep 2026 for the Collection-to-Sales Ratio feature - see
// LIVE_SALES_RS_TREND_SHEET above. Same pattern as the two handlers below.
function handleLiveSalesRsTrendUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_SALES_RS_TREND_SHEET, LIVE_SALES_RS_TREND_HEADERS);
  var rows = body.trend || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_SALES_RS_TREND_HEADERS.length, rows,
    function (r) { return r && r.date ? String(r.date).trim() : null; },
    function (r, now) {
      return [
        r.date,
        (r.rs === null || r.rs === undefined) ? "" : r.rs,
        now
      ];
    });
  return jsonOut_({ ok: true, updated: updated });
}

function handleLiveCollectionTrendUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_COLLECTION_TREND_SHEET, LIVE_COLLECTION_TREND_HEADERS);
  var rows = body.trend || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_COLLECTION_TREND_HEADERS.length, rows,
    function (r) { return r && r.date ? String(r.date).trim() : null; },
    function (r, now) {
      return [
        r.date,
        (r.rs === null || r.rs === undefined) ? "" : r.rs,
        now
      ];
    });
  return jsonOut_({ ok: true, updated: updated });
}

// POST { action: "logAsk", source, dsrName, question, understood }
// source: "DSR App" or "Dashboard". dsrName: the DSR's full Tally name for
// a DSR app call, blank for the Dashboard (company-wide, not tied to one
// DSR). No LIVE_SYNC_KEY needed - this comes straight from the app/
// dashboard itself, the same trust level as an order add. Not behind the
// order-saving LockService lock (see doPost below) since it's a simple
// append with nothing else to keep consistent with, same as the Live*
// update handlers above.
function handleLogAsk_(body) {
  var sheet = getLiveSheet_(ASK_LOG_SHEET, ASK_LOG_HEADERS);
  var now = new Date();
  sheet.appendRow([
    now.toISOString(),
    formatIst_(now),
    body.source || "DSR App",
    body.dsrName || "",
    body.question || "",
    body.understood === false ? "No" : "Yes"
  ]);
  return jsonOut_({ ok: true });
}

// ===== Dashboard/DSR-app login + usage-time tracking, added 26 Sep 2026 =====
// Owner's ask: password-protect the dashboard (4 named managers) and each
// DSR app (one password per DSR), and record how long each person actually
// uses the app so usage can be reviewed later.
//
// Two new sheets:
//   Users     - Name | Role (Owner/Manager/DSR) | PasswordHash | Salt | UpdatedAt
//               One row per person. Passwords are never stored in plain
//               text - only a SHA-256 hash of "password|salt", so reading
//               the Sheet directly never reveals anyone's actual password.
//   LoginLog  - SessionId | Name | Role | App | LoginAt | LastSeenAt | Heartbeats
//               One row per LOGIN (not per heartbeat - a session's own row
//               is updated in place on every heartbeat, so this sheet never
//               grows unbounded the way a full click/ping log would).
//               "How long did they use it" = LastSeenAt - LoginAt for a
//               session; the client only sends a heartbeat while its tab is
//               actually visible/foregrounded, so a backgrounded or closed
//               tab naturally stops advancing LastSeenAt rather than
//               needing a reliable "logout" event (which mobile browsers
//               don't fire consistently).
var USERS_SHEET = "Users";
// AllowedAreas added 27 Sep 2026 - a comma-separated list of DSRS[] `name`
// values (e.g. "Prabha,Arun,Chellamani") this user is restricted to seeing
// on the dashboard; blank/empty means unrestricted (sees every DSR, same
// as before this column existed). Only meaningful for role "Manager" -
// the Owner and every DSR login are unaffected by it either way (Owner
// always sees everything; a DSR app's own login never reads this field).
//
// HiddenUsageFor added 28 Sep 2026 ("bala also dont want to see anbu
// login but anbu wants bala login to see") - a comma-separated list of
// NAMES whose login/usage rows should be hidden from THIS user
// specifically in the usage report, independent of role. This is
// per-viewer and NOT symmetric by itself - Bala's own row gets
// HiddenUsageFor="Anbu" (Bala never sees Anbu's logins), while Anbu's row
// stays blank (Anbu still sees Bala's logins, and everyone else's) - each
// person's own preference lives on their own row. Blank = see everyone
// they'd otherwise be allowed to see (no change from before this column
// existed). This stacks with, but is a completely separate mechanism
// from, the existing "a restricted Area Manager never sees ANY manager's
// login" role-based rule (enforced client-side in the dashboard's own
// usageRowVisibleToViewer_(), same as this field is) - that one is
// automatic and role-driven, this one is an explicit, one-off,
// per-person exclusion list the Owner sets by hand.
// handleUsageSummary_() itself is unchanged - it still just hands back
// every raw row; both filters are applied client-side, same "dumb
// aggregation" convention as everything else on this dashboard.
var USERS_HEADERS = ["Name", "Role", "PasswordHash", "Salt", "UpdatedAt", "AllowedAreas", "HiddenUsageFor"];
var LOGINLOG_SHEET = "LoginLog";
var LOGINLOG_HEADERS = ["SessionId", "Name", "Role", "App", "LoginAt", "LastSeenAt", "Heartbeats"];

function hashPassword_(password, salt) {
  var digestBytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(password) + "|" + String(salt)
  );
  return digestBytes.map(function (b) {
    return ("0" + (b & 0xFF).toString(16)).slice(-2);
  }).join("");
}

// Row number (1-based, sheet-relative) of a user by name, case/whitespace-
// insensitive, or -1 if not found. Small linear scan - the Users sheet will
// only ever hold a few dozen rows at most (managers + DSRs), so this never
// needs the bulk-upsert machinery the big Live* sheets use.
function findUserRow_(usersSheet, name) {
  var lastRow = usersSheet.getLastRow();
  if (lastRow < 2) return -1;
  var names = usersSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var target = String(name || "").trim().toLowerCase();
  for (var i = 0; i < names.length; i++) {
    if (String(names[i][0] || "").trim().toLowerCase() === target) return i + 2;
  }
  return -1;
}

// allowedAreas/hiddenUsageFor: pass undefined to leave an existing user's
// value untouched (e.g. a plain password reset that doesn't mention it at
// all); pass a string (including "") to set/clear it - "" explicitly
// means "back to the default" (unrestricted / hides nobody), so it must
// NOT be treated the same as undefined.
function setUserPassword_(usersSheet, name, role, newPassword, allowedAreas, hiddenUsageFor) {
  var row = findUserRow_(usersSheet, name);
  var salt = Utilities.getUuid();
  var hash = hashPassword_(newPassword, salt);
  var now = new Date().toISOString();
  if (row === -1) {
    usersSheet.appendRow([name, role || "DSR", hash, salt, now, allowedAreas || "", hiddenUsageFor || ""]);
  } else {
    var existingRole = usersSheet.getRange(row, 2).getValue();
    var existingAreas = usersSheet.getRange(row, 6).getValue();
    var existingHidden = usersSheet.getRange(row, 7).getValue();
    var areasToWrite = (allowedAreas === undefined) ? existingAreas : allowedAreas;
    var hiddenToWrite = (hiddenUsageFor === undefined) ? existingHidden : hiddenUsageFor;
    usersSheet.getRange(row, 2, 1, 6).setValues([[role || existingRole, hash, salt, now, areasToWrite, hiddenToWrite]]);
  }
}

// POST {action:"login", name, password, app}
function handleLogin_(body) {
  var name = String(body.name || "").trim();
  var password = String(body.password || "");
  var app = String(body.app || "");
  if (!name || !password) return jsonOut_({ ok: false, error: "Name and password required" });
  var usersSheet = getLiveSheet_(USERS_SHEET, USERS_HEADERS);
  var row = findUserRow_(usersSheet, name);
  if (row === -1) return jsonOut_({ ok: false, error: "Invalid name or password" });
  var data = usersSheet.getRange(row, 1, 1, USERS_HEADERS.length).getValues()[0];
  var role = data[1], storedHash = data[2], salt = data[3];
  var allowedAreas = String(data[5] || "").trim();
  var hiddenUsageFor = String(data[6] || "").trim();
  if (hashPassword_(password, salt) !== storedHash) {
    return jsonOut_({ ok: false, error: "Invalid name or password" });
  }
  var sessionId = Utilities.getUuid();
  var now = new Date().toISOString();
  // 27 Sep 2026, owner's own request: "do not track me or show myname to
  // any of the managers" - the Owner's own logins/usage are never written
  // to LoginLog at all (not just hidden from the report), so there is
  // nothing to show anyone, himself included. Login itself still succeeds
  // normally - only the usage-tracking side effect is skipped.
  if (role !== "Owner") {
    var logSheet = getLiveSheet_(LOGINLOG_SHEET, LOGINLOG_HEADERS);
    logSheet.appendRow([sessionId, data[0], role, app, now, now, 1]);
  }
  // allowedAreas ("" when unrestricted) tells the dashboard which DSR
  // cards/totals a restricted Area Manager (e.g. Vimal) should be limited
  // to - see the "AllowedAreas" note on USERS_HEADERS above. hiddenUsageFor
  // ("" when nobody is hidden) is this same user's OWN per-viewer usage-
  // report exclusion list (28 Sep 2026) - see the "HiddenUsageFor" note on
  // USERS_HEADERS above.
  return jsonOut_({ ok: true, sessionId: sessionId, name: data[0], role: role, allowedAreas: allowedAreas, hiddenUsageFor: hiddenUsageFor });
}

// POST {action:"heartbeat", sessionId, name, role, app} - "still here" ping,
// sent every few minutes only while the tab is visible. Updates that
// session's own LastSeenAt/Heartbeats in place rather than appending a new
// row each time.
function handleHeartbeat_(body) {
  var sessionId = String(body.sessionId || "");
  if (!sessionId) return jsonOut_({ ok: false, error: "sessionId required" });
  // Same Owner exclusion as handleLogin_() above - an Owner session was
  // never given a LoginLog row to begin with, so every heartbeat for one
  // is a deliberate no-op rather than "recreate a minimal row" (which
  // would otherwise silently start tracking the Owner again from here).
  if (String(body.role || "") === "Owner") return jsonOut_({ ok: true });
  var logSheet = getLiveSheet_(LOGINLOG_SHEET, LOGINLOG_HEADERS);
  var lastRow = logSheet.getLastRow();
  var now = new Date().toISOString();
  if (lastRow >= 2) {
    var ids = logSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || "") === sessionId) {
        var r = i + 2;
        var heartbeats = Number(logSheet.getRange(r, 7).getValue()) || 0;
        logSheet.getRange(r, 6, 1, 2).setValues([[now, heartbeats + 1]]);
        return jsonOut_({ ok: true });
      }
    }
  }
  // Session row not found (e.g. the Sheet was cleared) - recreate a minimal
  // row from what the client resent, rather than silently dropping the ping.
  logSheet.appendRow([sessionId, String(body.name || ""), String(body.role || ""), String(body.app || ""), now, now, 1]);
  return jsonOut_({ ok: true });
}

// GET ?type=usersList - names + roles only, no password data, for the
// dashboard's admin panel "pick a user" dropdown.
function handleUsersList_() {
  var sheet = getLiveSheet_(USERS_SHEET, USERS_HEADERS);
  var lastRow = sheet.getLastRow();
  var out = [];
  if (lastRow >= 2) {
    // Reads all the way to column 7 now (HiddenUsageFor) - still no
    // password/hash data, same privacy guarantee as before.
    var data = sheet.getRange(2, 1, lastRow - 1, USERS_HEADERS.length).getValues();
    for (var i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      out.push({ name: data[i][0], role: data[i][1], allowedAreas: String(data[i][5] || ""), hiddenUsageFor: String(data[i][6] || "") });
    }
  }
  return jsonOut_({ ok: true, users: out });
}

// GET ?type=usageSummary - every login session on record. Left as an open
// read (no key), same as ?type=live - the dashboard itself is now
// password-gated at the UI level, and all 4 managers are meant to be able
// to see this. The client groups sessions by person/day and sums active
// minutes; kept dumb here on purpose, same pattern as every other
// aggregation on this dashboard.
function handleUsageSummary_() {
  var sheet = getLiveSheet_(LOGINLOG_SHEET, LOGINLOG_HEADERS);
  var lastRow = sheet.getLastRow();
  var out = [];
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, LOGINLOG_HEADERS.length).getValues();
    for (var i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      out.push({
        name: data[i][1], role: data[i][2], app: data[i][3],
        loginAt: data[i][4], lastSeenAt: data[i][5], heartbeats: data[i][6]
      });
    }
  }
  return jsonOut_({ ok: true, sessions: out });
}

// POST {action:"setPassword", targetName, targetRole, newPassword, auth:{name,password}}
// or {action:"setPassword", targetName, targetRole, newPassword, key:LIVE_SYNC_KEY}
// Two ways to authorize a password change, never a client-side "I'm the
// admin" flag: the shared LIVE_SYNC_KEY (master override, used for
// first-time setup before any Owner login exists) or the CURRENT Owner's
// own name+password re-entered fresh (normal day-to-day resets from the
// dashboard's admin panel). Creates the target user if they don't exist yet.
function handleSetPassword_(body) {
  var targetName = String(body.targetName || "").trim();
  var targetRole = String(body.targetRole || "").trim();
  var newPassword = String(body.newPassword || "");
  if (!targetName || !newPassword) return jsonOut_({ ok: false, error: "targetName and newPassword required" });
  if (newPassword.length < 4) return jsonOut_({ ok: false, error: "Password must be at least 4 characters" });

  var usersSheet = getLiveSheet_(USERS_SHEET, USERS_HEADERS);

  if (body.key) {
    if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  } else {
    var adminName = String((body.auth && body.auth.name) || "").trim();
    var adminPassword = String((body.auth && body.auth.password) || "");
    var adminRow = findUserRow_(usersSheet, adminName);
    if (adminRow === -1) return jsonOut_({ ok: false, error: "Not authorized" });
    var adminData = usersSheet.getRange(adminRow, 1, 1, USERS_HEADERS.length).getValues()[0];
    if (adminData[1] !== "Owner") return jsonOut_({ ok: false, error: "Not authorized" });
    if (hashPassword_(adminPassword, adminData[3]) !== adminData[2]) return jsonOut_({ ok: false, error: "Not authorized" });
  }

  // allowedAreas/hiddenUsageFor are both optional - omit either entirely
  // (undefined) to leave that existing value untouched; send "" to clear
  // it back to the default, or a comma-separated name list to set it.
  var allowedAreas = (body.allowedAreas === undefined) ? undefined : String(body.allowedAreas || "").trim();
  var hiddenUsageFor = (body.hiddenUsageFor === undefined) ? undefined : String(body.hiddenUsageFor || "").trim();
  setUserPassword_(usersSheet, targetName, targetRole, newPassword, allowedAreas, hiddenUsageFor);
  return jsonOut_({ ok: true });
}

// One-time bootstrap - GET ?type=provisionUsers&key=<LIVE_SYNC_KEY>. Creates
// every default account with a simple starting password ONLY if that name
// doesn't already exist in the Users sheet - safe to visit more than once,
// it will never overwrite a password that's since been changed. Visit this
// URL once after deploying this update, then change any of these from the
// dashboard's admin panel (Owner login required).
function handleProvisionUsers_(key) {
  if (key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var usersSheet = getLiveSheet_(USERS_SHEET, USERS_HEADERS);
  // 27 Sep 2026: Anbu and Bala are full Managers (blank AllowedAreas, see
  // everything - unchanged). Vimal is an "Area Manager" - its own distinct
  // Role value as of the 28 Sep 2026 follow-up (the owner's own request:
  // "can u create vimal as Area Manager under role so that we can define
  // a new role to new one added" - previously Area Manager was only a
  // Manager role + a non-blank AllowedAreas, with no way to pick it
  // directly when creating someone new) - restricted to only the
  // Prabha/Arun/Chellamani dashboard cards (his own request: "vimal is
  // looking after only arun,prabha,chellamani and saravanan area" -
  // Saravanan has no separate dashboard card of his own, his real Tally
  // sales/collection are already folded into Chellamani's card, so
  // "Chellamani" alone covers both areas here). Role is just a free string
  // everywhere in this file (only "Owner" is ever specially checked), so
  // "Area Manager" needed no schema/validation change - see
  // handleSetPassword_()'s own note on this. Bala's own HiddenUsageFor is
  // "Anbu" - his own request ("bala also dont want to see anbu login but
  // anbu wants bala login to see"), so this is deliberately one-directional:
  // Anbu's own row stays blank, so Anbu still sees Bala's logins fine.
  var defaults = [
    ["Alagu", "Owner", "Alagu123", "", ""],
    ["Anbu", "Manager", "Anbu123", "", ""],
    ["Bala", "Manager", "Bala123", "", "Anbu"],
    ["Vimal", "Area Manager", "Vimal123", "Prabha,Arun,Chellamani", ""],
    ["Prabha", "DSR", "Prabha123", "", ""],
    ["Arun", "DSR", "Arun123", "", ""],
    ["Nagaraj", "DSR", "Nagaraj123", "", ""],
    ["Chellamani", "DSR", "Chellamani123", "", ""],
    ["Anandh", "DSR", "Anandh123", "", ""],
    ["Madhu", "DSR", "Madhu123", "", ""],
    ["Saravanan", "DSR", "Saravanan123", "", ""],
    ["AAA Direct", "DSR", "Aaadirect123", "", ""]
  ];
  var created = [];
  for (var i = 0; i < defaults.length; i++) {
    var name = defaults[i][0], role = defaults[i][1], pw = defaults[i][2], areas = defaults[i][3], hidden = defaults[i][4];
    if (findUserRow_(usersSheet, name) === -1) {
      setUserPassword_(usersSheet, name, role, pw, areas, hidden);
      created.push(name);
    }
  }
  return jsonOut_({ ok: true, created: created });
}

function handleLiveGet_() {
  var stockRows = liveSheetRows_(getLiveSheet_(LIVE_STOCK_SHEET, LIVE_STOCK_HEADERS));
  var outstandingRows = liveSheetRows_(getLiveSheet_(LIVE_OUTSTANDING_SHEET, LIVE_OUTSTANDING_HEADERS));
  var collectionRows = liveSheetRows_(getLiveSheet_(LIVE_COLLECTION_SHEET, LIVE_COLLECTION_HEADERS));
  var collectionByDealerRows = liveSheetRows_(getLiveSheet_(LIVE_COLLECTION_BY_DEALER_SHEET, LIVE_COLLECTION_BY_DEALER_HEADERS));
  var groupSummaryRows = liveSheetRows_(getLiveSheet_(LIVE_GROUP_SUMMARY_SHEET, LIVE_GROUP_SUMMARY_HEADERS));
  var groupSummaryByDealerRows = liveSheetRows_(getLiveSheet_(LIVE_GROUP_SUMMARY_BY_DEALER_SHEET, LIVE_GROUP_SUMMARY_BY_DEALER_HEADERS));
  var salesRows = liveSheetRows_(getLiveSheet_(LIVE_SALES_SHEET, LIVE_SALES_HEADERS));
  var salesByDealerRows = liveSheetRows_(getLiveSheet_(LIVE_SALES_BY_DEALER_SHEET, LIVE_SALES_BY_DEALER_HEADERS));
  var salesByItemRows = liveSheetRows_(getLiveSheet_(LIVE_SALES_BY_ITEM_SHEET, LIVE_SALES_BY_ITEM_HEADERS));
  var salesByDsrItemRows = liveSheetRows_(getLiveSheet_(LIVE_SALES_BY_DSR_ITEM_SHEET, LIVE_SALES_BY_DSR_ITEM_HEADERS));
  var salesTrendRows = liveSheetRows_(getLiveSheet_(LIVE_SALES_TREND_SHEET, LIVE_SALES_TREND_HEADERS));
  var salesRsTrendRows = liveSheetRows_(getLiveSheet_(LIVE_SALES_RS_TREND_SHEET, LIVE_SALES_RS_TREND_HEADERS));
  var collectionTrendRows = liveSheetRows_(getLiveSheet_(LIVE_COLLECTION_TREND_SHEET, LIVE_COLLECTION_TREND_HEADERS));

  var stock = {};
  for (var i = 0; i < stockRows.length; i++) {
    var r = stockRows[i];
    var name = String(r[0] || "").trim();
    if (!name) continue;
    stock[name] = {
      stock: r[1] === "" ? null : Number(r[1]),
      rate: r[2] === "" ? null : Number(r[2]),
      mrp: r[3] === "" ? null : Number(r[3])
    };
  }

  var outstanding = {};
  for (var j = 0; j < outstandingRows.length; j++) {
    var r2 = outstandingRows[j];
    var dname = String(r2[0] || "").trim();
    if (!dname) continue;
    // Bug found and fixed 8 Sep 2026 (badges disappearing a few seconds
    // after page load): the "Outstanding Since" cell holds a plain
    // "yyyy-MM-dd" string when handleLiveOutstandingUpdate_() writes it,
    // but Google Sheets silently auto-converts that into a REAL Date value
    // on a freshly-created sheet (the exact same class of bug as the
    // "Seventh bug" in dsr-app-registry.md, which hit the Orders sheet's
    // Date Key/Month Key columns). Reading it back with getValues() then
    // returns a JS Date object, not the original string; JSON.stringify()-ing
    // a Date calls toISOString(), which converts to UTC - for a date built
    // at local midnight (IST, UTC+5:30) that shifts the date back a day AND
    // appends a time component, e.g. "2026-03-08T18:30:00.000Z" instead of
    // "2026-03-09". The client's daysBetweenISO() naively splits on "-" and
    // silently produces NaN/Invalid Date on that shape, which outstandingDays()
    // then reports as null - so outstandingBadgeHtml() renders nothing at
    // all. Net effect: any dealer with a genuinely-pending live outstanding
    // date had their badge vanish within seconds of the live-overlay fetch
    // resolving, even though the underlying data was correct. Reusing the
    // existing normalizeDateKey_() helper (built for the Seventh bug) fixes
    // this the same way: format a real Date back to plain "yyyy-MM-dd" text;
    // leave an already-plain string untouched. "" (genuinely cleared/paid
    // off) still normalizes to "" -> null below, unchanged.
    var sinceNorm = normalizeDateKey_(r2[1]);
    // Changed 8 Sep 2026: each dealer's entry is now {since, balance}
    // instead of a bare date value, so the live current balance (for
    // "Manage dealers") travels in the same response - see the matching
    // client-side applyLiveOverlay_() change in every DSR app file.
    // "pendingAmount" added 17 Sep 2026 (BILLCL-based, see
    // LIVE_OUTSTANDING_HEADERS above) - trial run, Arun's app only reads
    // this field for now, but it's sent to every app the same way since
    // the underlying Tally fetch is already company-wide.
    // "agedPendingAmount" added 19 Sep 2026 for the payment-reminder
    // feature, same trial scope (Arun-only reads it for now).
    // "band3059"/"band6089"/"band90plus" added 22 Sep 2026, "band0_29"
    // added 23 Sep 2026 - see LIVE_OUTSTANDING_HEADERS above for why
    // (age-accurate breakdown, replacing the old single-band-per-dealer
    // approach every display used until now).
    outstanding[dname] = {
      since: sinceNorm || null,
      balance: r2[2] === "" ? null : Number(r2[2]),
      pendingAmount: r2[3] === "" ? null : Number(r2[3]),
      agedPendingAmount: r2[4] === "" ? null : Number(r2[4]),
      band0_29: r2[5] === "" ? null : Number(r2[5]),
      band3059: r2[6] === "" ? null : Number(r2[6]),
      band6089: r2[7] === "" ? null : Number(r2[7]),
      band90plus: r2[8] === "" ? null : Number(r2[8])
    };
  }

  // Added 8 Sep 2026 - grouped {dsrName: [dealerName, ...]} for the
  // new-dealer auto-add feature. Every DSR app only ever looks up its own
  // dsrName's entry (see applyLiveOverlay_() in each app file), so sending
  // the full roster for all 8 DSRs to every app is simplest and matches how
  // stock/outstanding already work (whole snapshot, filtered client-side).
  var rosterRows = liveSheetRows_(getLiveSheet_(LIVE_DEALER_ROSTER_SHEET, LIVE_DEALER_ROSTER_HEADERS));
  var dealerRoster = {};
  for (var k = 0; k < rosterRows.length; k++) {
    var r3 = rosterRows[k];
    var rDsr = String(r3[1] || "").trim();
    var rName = String(r3[2] || "").trim();
    if (!rDsr || !rName) continue;
    if (!dealerRoster[rDsr]) dealerRoster[rDsr] = [];
    dealerRoster[rDsr].push(rName);
  }

  // Added 19 Sep 2026 (Arun-only trial) - grouped {dsrName: [{name,
  // firstSeenDate}, ...]}, same shape/scope pattern as dealerRoster just
  // above. See LIVE_NEW_DEALERS_HEADERS for why this is its own sheet
  // rather than derived from Orders-sheet history.
  var newDealerRows = liveSheetRows_(getLiveSheet_(LIVE_NEW_DEALERS_SHEET, LIVE_NEW_DEALERS_HEADERS));
  var newDealers = {};
  for (var nd = 0; nd < newDealerRows.length; nd++) {
    var r5 = newDealerRows[nd];
    var ndDsr = String(r5[1] || "").trim();
    var ndName = String(r5[2] || "").trim();
    if (!ndDsr || !ndName) continue;
    if (!newDealers[ndDsr]) newDealers[ndDsr] = [];
    newDealers[ndDsr].push({ name: ndName, firstSeenDate: normalizeDateKey_(r5[3]) || null });
  }

  // Added 12 Sep 2026 - {dsrName: amount}, one entry per DSR that has a
  // Collection figure. Every DSR app only ever looks up its own dsrName's
  // entry (see applyLiveOverlay_()), same pattern as stock/outstanding/
  // dealerRoster above.
  var collection = {};
  for (var m = 0; m < collectionRows.length; m++) {
    var r4 = collectionRows[m];
    var cDsr = String(r4[0] || "").trim();
    if (!cDsr) continue;
    collection[cDsr] = {
      today: r4[1] === "" ? null : Number(r4[1]),
      month: r4[2] === "" ? null : Number(r4[2]),
      // "MB" added 21 Sep 2026 - see LIVE_COLLECTION_HEADERS above.
      mb: r4[3] === "" || r4[3] === undefined ? null : Number(r4[3]),
      // "Last Month Same Date" added 21 Sep 2026 - see LIVE_COLLECTION_HEADERS above.
      lastMonthSameDate: r4[4] === "" || r4[4] === undefined ? null : Number(r4[4]),
      // "Last Month Full" added 21 Sep 2026 - see LIVE_COLLECTION_HEADERS above.
      lastMonthFull: r4[5] === "" || r4[5] === undefined ? null : Number(r4[5]),
      // "Avg 6mo Collected Rs" added 26 Sep 2026 for the "Full Petronas
      // This Month Collection" area table's new "6M AVG" column - see
      // LIVE_COLLECTION_HEADERS above.
      avg6moCollectedRs: r4[6] === "" || r4[6] === undefined ? null : Number(r4[6])
    };
  }

  // Added 17 Sep 2026 - {dealerName: {today, month}}, same shape as
  // `collection` above but per dealer instead of per DSR - feeds the
  // dashboard's "Top customers - Collection" panel only; no DSR app reads
  // this field (each DSR app only ever looks at its own dsrName's entry
  // in `collection` above).
  var collectionByDealer = {};
  for (var n = 0; n < collectionByDealerRows.length; n++) {
    var r5 = collectionByDealerRows[n];
    var cDealer = String(r5[0] || "").trim();
    if (!cDealer) continue;
    collectionByDealer[cDealer] = {
      today: r5[1] === "" ? null : Number(r5[1]),
      month: r5[2] === "" ? null : Number(r5[2]),
      // "Last payment" added 20 Sep 2026 for the dealer-lookup card,
      // same plain-text yyyymmdd-or-null convention as salesByDealer's
      // lastSaleDate just below.
      lastPaymentDate: r5[3] === "" || r5[3] === undefined ? null : String(r5[3]),
      // "Avg 6mo Collected Rs" added 23 Sep 2026 for the "this month vs
      // 6-month avg" dealer review feature.
      avg6moCollectedRs: r5[4] === "" || r5[4] === undefined ? null : Number(r5[4]),
      // "Last Month Same Date"/"Last Month Full" added 25 Sep 2026 for
      // the Dealer Review popup - column indices 5/6, matching
      // LIVE_COLLECTION_BY_DEALER_HEADERS' order exactly (0 Dealer Name,
      // 1 Today Collected, 2 Month Collected, 3 Last Payment Date,
      // 4 Avg 6mo Collected Rs, 5 Last Month Same Date, 6 Last Month
      // Full, 7 Last Payment Amount, 8 Updated At).
      lastMonthSameDate: r5[5] === "" || r5[5] === undefined ? null : Number(r5[5]),
      lastMonthFull: r5[6] === "" || r5[6] === undefined ? null : Number(r5[6]),
      // "Last Payment Amount" added 26 Sep 2026 for the dealer-lookup
      // card's "Last payment" tile - see LIVE_COLLECTION_BY_DEALER_HEADERS
      // above.
      lastPaymentAmount: r5[7] === "" || r5[7] === undefined ? null : Number(r5[7])
    };
  }

  // Added 18 Sep 2026 - dashboard-only "actually billed" Sales, from
  // Tally's own Sales vouchers (see LIVE_SALES_SHEET above and
  // tally_live_watcher.py's fetch_sales() for the full history/scoping).
  // Same {key: {today, month}} shape as collection/collectionByDealer -
  // sales is per-DSR (mirrors `collection`), salesByDealer/salesByItem
  // feed the dashboard's Top-5 panels once those switch over to this
  // Tally figure. salesByDealer/salesByItem stay dashboard-only; `sales`
  // (per-DSR) is ALSO read by the DSR apps as of 21 Sep 2026, but ONLY for
  // its lastMonthSameDateLtr field (Tally-based "Last month" comparison) -
  // each app's own Sales month/today figures stay app-punched, per the
  // owner's explicit correction earlier that day.
  function readSalesMap_(rows) {
    var out = {};
    for (var s = 0; s < rows.length; s++) {
      var r = rows[s];
      var key = String(r[0] || "").trim();
      if (!key) continue;
      out[key] = {
        today: r[1] === "" ? null : Number(r[1]),
        month: r[2] === "" ? null : Number(r[2])
      };
    }
    return out;
  }
  // `sales` (per-DSR) gets its own reader, not readSalesMap_ - added 21 Sep
  // 2026 for "Last Month Same Date" (LIVE_SALES_HEADERS' 4th column, see
  // above), a column salesByItem's sheet doesn't have.
  function readSalesDsrMap_(rows) {
    var out = {};
    for (var s = 0; s < rows.length; s++) {
      var r = rows[s];
      var key = String(r[0] || "").trim();
      if (!key) continue;
      out[key] = {
        today: r[1] === "" ? null : Number(r[1]),
        month: r[2] === "" ? null : Number(r[2]),
        lastMonthSameDateLtr: r[3] === "" || r[3] === undefined ? null : Number(r[3]),
        // "Last Month Full" added 21 Sep 2026 - see LIVE_SALES_HEADERS above.
        lastMonthFullLtr: r[4] === "" || r[4] === undefined ? null : Number(r[4]),
        // "Today Sold Rs"/"Month Sold Rs" added 24 Sep 2026 for Collection
        // Health - see LIVE_SALES_HEADERS above.
        todaySoldRs: r[5] === "" || r[5] === undefined ? null : Number(r[5]),
        monthSoldRs: r[6] === "" || r[6] === undefined ? null : Number(r[6]),
        // "Avg 6mo Sold Ltr"/"Avg 6mo Sold Rs" added 26 Sep 2026 for the
        // "Full Petronas This Month Sales" area table's new "6M AVG"
        // column - see LIVE_SALES_HEADERS above.
        avg6moLtr: r[7] === "" || r[7] === undefined ? null : Number(r[7]),
        avg6moRs: r[8] === "" || r[8] === undefined ? null : Number(r[8])
      };
    }
    return out;
  }
  var sales = readSalesDsrMap_(salesRows);
  var salesByItem = readSalesMap_(salesByItemRows);
  // salesByDealer gets its own reader (not readSalesMap_) - added 18 Sep
  // 2026 for the "Dormant dealers" panel, which needs each dealer's
  // "Last Sale Date" (column index 3, a plain yyyymmdd text string - see
  // handleLiveSalesByDealerUpdate_) alongside today/month, a column
  // sales/salesByItem don't have.
  var salesByDealer = {};
  for (var sd = 0; sd < salesByDealerRows.length; sd++) {
    var sdr = salesByDealerRows[sd];
    var sdKey = String(sdr[0] || "").trim();
    if (!sdKey) continue;
    salesByDealer[sdKey] = {
      today: sdr[1] === "" ? null : Number(sdr[1]),
      month: sdr[2] === "" ? null : Number(sdr[2]),
      lastSaleDate: sdr[3] === "" || sdr[3] === undefined ? null : String(sdr[3]),
      // "Avg 6mo Sold Ltr" added 23 Sep 2026 for the "this month vs
      // 6-month avg" dealer review feature.
      avg6moLtr: sdr[4] === "" || sdr[4] === undefined ? null : Number(sdr[4]),
      // "Last Sale Before Month" added 23 Sep 2026 for "Dealers
      // Reactivated This Month" - see LIVE_SALES_BY_DEALER_HEADERS above.
      lastSaleBeforeMonth: sdr[5] === "" || sdr[5] === undefined ? null : String(sdr[5]),
      // Six new fields added 25 Sep 2026 for the Dealer Review popup
      // (this month / last month / 6-mo avg, Ltr AND Rs) - column
      // indices 6-12, matching LIVE_SALES_BY_DEALER_HEADERS' order
      // exactly (0 Dealer Name, 1 Today Sold Ltr, 2 Month Sold Ltr,
      // 3 Last Sale Date, 4 Avg 6mo Sold Ltr, 5 Last Sale Before Month,
      // 6 Last Month Same Date Ltr, 7 Last Month Full Ltr, 8 Today Sold
      // Rs, 9 Month Sold Rs, 10 Last Month Same Date Rs, 11 Last Month
      // Full Rs, 12 Avg 6mo Sold Rs, 13 Last Sale Ltr, 14 Updated At).
      lastMonthSameDateLtr: sdr[6] === "" || sdr[6] === undefined ? null : Number(sdr[6]),
      lastMonthFullLtr: sdr[7] === "" || sdr[7] === undefined ? null : Number(sdr[7]),
      todaySoldRs: sdr[8] === "" || sdr[8] === undefined ? null : Number(sdr[8]),
      monthSoldRs: sdr[9] === "" || sdr[9] === undefined ? null : Number(sdr[9]),
      lastMonthSameDateRs: sdr[10] === "" || sdr[10] === undefined ? null : Number(sdr[10]),
      lastMonthFullRs: sdr[11] === "" || sdr[11] === undefined ? null : Number(sdr[11]),
      avg6moRs: sdr[12] === "" || sdr[12] === undefined ? null : Number(sdr[12]),
      // "Last Sale Ltr" added 26 Sep 2026 for the dealer-lookup card's
      // "Last order" tile - see LIVE_SALES_BY_DEALER_HEADERS above.
      lastSaleLtr: sdr[13] === "" || sdr[13] === undefined ? null : Number(sdr[13])
    };
  }

  // Added 23 Sep 2026 for the fixed, bill-level 60+ day overdue stat -
  // {dealerName: {baseline, collected, billCount}}, one entry per dealer
  // that has at least one bill in this month's frozen baseline. See
  // LIVE_SIXTYPLUS_BY_DEALER_HEADERS/handleLiveSixtyPlusBaselineUpdate_
  // above and tally_live_watcher.py's snapshot_or_get_sixtyplus_baseline()
  // docstring for the full design.
  var sixtyPlusByDealerRows = liveSheetRows_(getLiveSheet_(LIVE_SIXTYPLUS_BY_DEALER_SHEET, LIVE_SIXTYPLUS_BY_DEALER_HEADERS));
  var sixtyPlusByDealer = {};
  for (var sp = 0; sp < sixtyPlusByDealerRows.length; sp++) {
    var spr = sixtyPlusByDealerRows[sp];
    var spKey = String(spr[0] || "").trim();
    if (!spKey) continue;
    sixtyPlusByDealer[spKey] = {
      baseline: spr[1] === "" || spr[1] === undefined ? null : Number(spr[1]),
      collected: spr[2] === "" || spr[2] === undefined ? null : Number(spr[2]),
      billCount: spr[3] === "" || spr[3] === undefined ? null : Number(spr[3])
    };
  }

  // Added 20 Sep 2026 - {date: value} maps for the dashboard's Tally-based
  // trend chart (see LIVE_SALES_TREND_SHEET/LIVE_COLLECTION_TREND_SHEET
  // above). Sent as a flat map, not pre-split into this-month/last-month -
  // the dashboard already knows today's date and can bucket client-side
  // the same way it already does for every other date-keyed field here.
  function readTrendMap_(rows, valueCol) {
    var out = {};
    for (var t = 0; t < rows.length; t++) {
      var r = rows[t];
      var d = normalizeDateKey_(r[0]);
      if (!d) continue;
      out[d] = r[valueCol] === "" ? null : Number(r[valueCol]);
    }
    return out;
  }
  var salesTrend = readTrendMap_(salesTrendRows, 1);
  var salesRsTrend = readTrendMap_(salesRsTrendRows, 1);
  var collectionTrend = readTrendMap_(collectionTrendRows, 1);

  // Added 22 Sep 2026 for the dashboard's "Top selling SKU by DSR" panels
  // and "part number covered per DSR" stat - see LIVE_SALES_BY_DSR_ITEM_SHEET
  // above and tally_live_watcher.py's dsr_item_totals for the full story.
  // Sent as a flat array (not a nested {dsrName: {itemName: ltr}} object) -
  // the dashboard already groups/sorts flat per-dealer/per-item data
  // client-side everywhere else on this page, same pattern here.
  var salesByDsrItem = [];
  for (var di = 0; di < salesByDsrItemRows.length; di++) {
    var dir = salesByDsrItemRows[di];
    var diDsr = String(dir[0] || "").trim();
    var diItem = String(dir[1] || "").trim();
    if (!diDsr || !diItem) continue;
    var diMonth = dir[2] === "" || dir[2] === undefined ? null : Number(dir[2]);
    if (diMonth === null || diMonth <= 0) continue;
    salesByDsrItem.push({ dsrName: diDsr, itemName: diItem, monthLtr: diMonth });
  }

  // Added 28 Sep 2026 for the Collection Group Summary panel - flat arrays
  // (not nested maps), same "dashboard groups/sorts client-side" pattern
  // salesByDsrItem above already uses. Every past month's rows are sent
  // too (not just the current month) since LIVE_GROUP_SUMMARY_SHEET keeps
  // all of them in the one sheet by design (see its own comment above) -
  // the dashboard itself decides what to show expanded/collapsed and
  // which month the Excel export covers (current month only, per the
  // owner's "download current month only" - see the dashboard JS).
  var groupSummary = [];
  for (var gs = 0; gs < groupSummaryRows.length; gs++) {
    var gsr = groupSummaryRows[gs];
    var gsGroup = String(gsr[0] || "").trim();
    // Fixed 28 Sep 2026 (same night as the feature shipped) - "Month" is a
    // plain "yyyy-MM" string on write, but Google Sheets silently
    // auto-converts a cell that LOOKS like a date into a real Date value
    // (the exact same class of bug as normalizeDateKey_()'s own comment
    // above, and the "Seventh bug" in dsr-app-registry.md) - so a naive
    // String(gsr[1]) here was returning the Date's full toString(), e.g.
    // "Tue Sep 01 2026 00:00:00 GMT+0530 (India Standard Time)", instead of
    // "2026-09". The dashboard's renderGroupSummaryTable_() matches
    // r.monthKey against currentMonthKey_() with strict === , so that
    // never matched anything - anyData stayed false forever and the panel
    // was permanently stuck on "Waiting for Collection Group Summary to
    // sync..." even though the watcher's pushes were succeeding and the
    // Sheet had the right data all along (confirmed directly: opened the
    // Sheet and the dashboard's own Network tab, both showed correct
    // Group/Opening/Debit/Credit/Closing - only Month was corrupted on
    // readback). normalizeMonthKey_() (defined near the top of this file)
    // already exists for exactly this - reusing it instead of a bare
    // String() fixes every row already in the Sheet too, no data rewrite
    // needed.
    var gsMonth = normalizeMonthKey_(gsr[1]);
    if (!gsGroup || !gsMonth) continue;
    groupSummary.push({
      group: gsGroup,
      monthKey: gsMonth,
      opening: gsr[2] === "" || gsr[2] === undefined ? null : Number(gsr[2]),
      debit: gsr[3] === "" || gsr[3] === undefined ? null : Number(gsr[3]),
      credit: gsr[4] === "" || gsr[4] === undefined ? null : Number(gsr[4]),
      closing: gsr[5] === "" || gsr[5] === undefined ? null : Number(gsr[5]),
      // updatedAt added 28 Sep 2026 alongside the existingKeyFn fix above -
      // lets the dashboard tell which of two rows for the same key is the
      // freshest, as a defensive front-end safety net (see
      // renderGroupSummaryTable_()'s own dedupe comment).
      updatedAt: gsr[6] || ""
    });
  }
  var groupSummaryByDealer = [];
  for (var gsd = 0; gsd < groupSummaryByDealerRows.length; gsd++) {
    var gsdr = groupSummaryByDealerRows[gsd];
    var gsdGroup = String(gsdr[0] || "").trim();
    // Same Sheets-auto-converts-dates fix as groupSummary's own gsMonth
    // just above - see that comment for the full root-cause explanation.
    var gsdMonth = normalizeMonthKey_(gsdr[1]);
    var gsdDealer = String(gsdr[2] || "").trim();
    if (!gsdGroup || !gsdMonth || !gsdDealer) continue;
    groupSummaryByDealer.push({
      group: gsdGroup,
      monthKey: gsdMonth,
      dealer: gsdDealer,
      opening: gsdr[3] === "" || gsdr[3] === undefined ? null : Number(gsdr[3]),
      debit: gsdr[4] === "" || gsdr[4] === undefined ? null : Number(gsdr[4]),
      credit: gsdr[5] === "" || gsdr[5] === undefined ? null : Number(gsdr[5]),
      closing: gsdr[6] === "" || gsdr[6] === undefined ? null : Number(gsdr[6]),
      // updatedAt - see groupSummary's own identical field just above.
      updatedAt: gsdr[7] || ""
    });
  }

  return jsonOut_({
    ok: true,
    stock: stock,
    outstanding: outstanding,
    dealerRoster: dealerRoster,
    newDealers: newDealers,
    collection: collection,
    collectionByDealer: collectionByDealer,
    sales: sales,
    salesByDealer: salesByDealer,
    salesByItem: salesByItem,
    salesByDsrItem: salesByDsrItem,
    salesTrend: salesTrend,
    salesRsTrend: salesRsTrend,
    collectionTrend: collectionTrend,
    sixtyPlusByDealer: sixtyPlusByDealer,
    groupSummary: groupSummary,
    groupSummaryByDealer: groupSummaryByDealer
  });
}

/**
 * GET ?dsrName=...&date=YYYY-MM-DD&month=YYYY-MM&lastMonthKey=YYYY-MM
 * Returns today's total, this month's total, and the list of today's orders
 * (full detail, for the edit/delete list) for that one DSR only.
 *
 * `lastMonthKey` (optional, added 19 Sep 2026 for the Target-Balance
 * panel's "Last month" comparison chip) - the previous calendar month's
 * own "yyyy-MM" key, computed client-side and passed straight through;
 * returns `lastMonthLtr` alongside the existing `monthLtr`. An older
 * cached front-end that never sends it just gets `lastMonthLtr: 0` back.
 *
 * New dealers used to be computed here too (added, then REVERTED same day,
 * 19 Sep 2026) - "first order in the Orders sheet this month" turned out
 * to mean nothing on a system this young (the sheet only goes back to
 * ~3 Sep 2026, so ANY dealer's first-ever app-recorded order looks "new,"
 * including real long-time customers with existing Tally outstanding -
 * the owner caught this immediately on real data). New dealers now comes
 * from `?type=live`'s own `newDealers` field instead (see handleLiveGet_()
 * below) - sourced from the daily Tally roster job's own new-vs-known
 * diff, not order history, since only Tally's roster reflects a dealer's
 * real business history.
 *
 * GET ?dsrName=...&historyFrom=YYYY-MM-DD&historyTo=YYYY-MM-DD&dealerSearch=...
 * Added 12 Sep 2026 (the Past Orders date-range search's missing backend
 * half - this endpoint previously had NO handling at all for historyFrom/
 * historyTo/historyDate/dealerSearch, which is why every Past Orders search
 * always came back "No orders found" regardless of the date range picked -
 * the front end asked a question this backend never answered). Returns
 * `historyOrders`: every non-deleted order for this dsrName whose Date Key
 * falls within [historyFrom, historyTo] (either bound optional - an empty
 * historyFrom means "no lower bound", an empty historyTo means "no upper
 * bound") AND whose dealer name contains dealerSearch (case-insensitive
 * substring, optional). `historyDate` alone (no historyFrom/historyTo, for
 * any older cached front-end that hasn't picked up the range-picker update
 * yet) is treated as an exact single-day lookup - the same behavior this
 * search always claimed to have. historyOrders is sorted most-recent-date
 * first (then by createdAt within a date), same convention as todaysOrders.
 * This is purely additive - the date/month/todaysOrders logic above is
 * completely untouched.
 *
 * GET ?type=live
 * Returns the current live-overlay snapshot (see "Live Tally Sync" above) -
 * a completely separate, unauthenticated read path from the dsrName one
 * above; needs no dsrName/date/month and never touches the Orders sheet.
 */
// Only these 7 DSRs count toward this endpoint's figures - same scope as
// the dashboard's 7 cards (see sales-dashboard.md's mapping table);
// Saravanan's orders, if any, aren't included separately - they're logged
// under Chellamani's own dsrName since her app covers his merged
// territory, same as everywhere else on this dashboard. AAA Direct added
// 20 Sep 2026 for the Sales trend chart (this endpoint's `dailyTrend`
// field, the only part of this endpoint still in use - the Top 5 panels
// it originally served switched to Tally-live data 18 Sep 2026) so the
// trend covers all 7 dashboard cards, not just the original 6.
var TOP_SALES_DSR_NAMES_ = [
  "PETRONAS THENI AREA (PRABHA)",
  "PETRONAS DINDIGUL AREA (ARUN)",
  "PETRONAS VIRUDHUNAGAR AREA (NAGARAJ)",
  "PETRONAS MDU AREA (CHELLAMANI)",
  "PETRONAS TVL1 AREA (ANANDH)",
  "PETRONAS NKL AREA (MADHU)",
  "PETRONAS AAA DIRECT AREA SALES"
];

/**
 * GET ?type=topSales&month=YYYY-MM
 * Added 17 Sep 2026 for the dashboard's "Top 5 items sold" / "Top
 * customers - Sales" panels (see claude/sales-dashboard.md). Scans every
 * non-deleted order row for the 6 dashboard DSRs (TOP_SALES_DSR_NAMES_
 * above) whose Month Key matches, and sums: (a) each item's own `ltr`
 * field (see each app's collectOrderItemsForSave() - {id, name, qty,
 * unit, ltr} per line, already the Ltr-normalized figure used everywhere
 * else on this dashboard) grouped by item name; (b) each order's
 * `totalLtr` grouped by dealer name. Returns the top 5 of each, sorted
 * descending. A completely separate read path from the dsrName-scoped
 * summary below - no dsrName required, never touches Live* sheets.
 *
 * Extended 17 Sep 2026, same day, with two more fields for the "last
 * activity per DSR" and "Sales trend" dashboard additions - both reuse
 * this same single sheet scan rather than adding new endpoints/requests:
 * - `lastOrderAt`: {dsrName: isoTimestamp_or_null} - the most recent
 *   `createdAt` among ANY of that DSR's non-deleted orders, deliberately
 *   NOT restricted to `monthKey` (unlike the item/dealer/trend
 *   aggregates below) - the point is "how long since this DSR last
 *   logged anything at all," which should still answer correctly on the
 *   1st of a new month before anyone's placed an order yet.
 * - `dailyTrend`: {dsrName: [{date, ltr}, ...]} - each DSR's own
 *   `totalLtr` summed by `dateKey`, for the requested `monthKey` only,
 *   sorted ascending by date (oldest first, matching how a trend line
 *   reads left-to-right).
 */
function handleTopSales_(monthKey) {
  var rows = readAllRows_(getSheet_());
  var itemTotals = {};
  var dealerTotals = {};
  var lastOrderAt = {};
  var dailyTrend = {};
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (row.deleted) continue;
    if (TOP_SALES_DSR_NAMES_.indexOf(row.dsrName) === -1) continue;

    // lastOrderAt - every non-deleted row for this DSR counts, regardless
    // of month, so this stays correct across a month boundary.
    var createdAtStr = row.createdAt ? new Date(row.createdAt).toISOString() : null;
    if (createdAtStr && (!lastOrderAt[row.dsrName] || createdAtStr > lastOrderAt[row.dsrName])) {
      lastOrderAt[row.dsrName] = createdAtStr;
    }

    if (monthKey && row.monthKey !== monthKey) continue;

    if (row.dealer) {
      dealerTotals[row.dealer] = (dealerTotals[row.dealer] || 0) + row.totalLtr;
    }
    if (row.dateKey) {
      if (!dailyTrend[row.dsrName]) dailyTrend[row.dsrName] = {};
      dailyTrend[row.dsrName][row.dateKey] = (dailyTrend[row.dsrName][row.dateKey] || 0) + row.totalLtr;
    }
    var items = [];
    try { items = JSON.parse(row.itemsJson || "[]"); } catch (err) { items = []; }
    for (var k = 0; k < items.length; k++) {
      var it = items[k];
      var name = it && it.name ? String(it.name) : "";
      if (!name) continue;
      var ltr = Number(it.ltr) || 0;
      itemTotals[name] = (itemTotals[name] || 0) + ltr;
    }
  }

  function topN_(obj, n) {
    return Object.keys(obj)
      .map(function (k) { return { name: k, ltr: Math.round(obj[k] * 100) / 100 }; })
      .sort(function (a, b) { return b.ltr - a.ltr; })
      .slice(0, n);
  }

  var dailyTrendOut = {};
  Object.keys(dailyTrend).forEach(function (dsr) {
    var days = dailyTrend[dsr];
    dailyTrendOut[dsr] = Object.keys(days).sort().map(function (dateKey) {
      return { date: dateKey, ltr: Math.round(days[dateKey] * 100) / 100 };
    });
  });

  return jsonOut_({
    ok: true,
    topItems: topN_(itemTotals, 5),
    topCustomers: topN_(dealerTotals, 5),
    lastOrderAt: lastOrderAt,
    dailyTrend: dailyTrendOut
  });
}

function doGet(e) {
  try {
    var params = (e && e.parameter) || {};

    if (String(params.type || "").trim() === "live") {
      return handleLiveGet_();
    }
    if (String(params.type || "").trim() === "topSales") {
      return handleTopSales_(String(params.month || "").trim());
    }
    if (String(params.type || "").trim() === "daybookMonth") {
      return handleDaybookMonthGet_(String(params.month || "").trim());
    }
    if (String(params.type || "").trim() === "daybookCleanupOldSales") {
      return handleDaybookCleanupOldSales_(String(params.key || "").trim());
    }
    // Login/usage-tracking reads, added 26 Sep 2026 - see handleUsersList_(),
    // handleUsageSummary_() and handleProvisionUsers_() above.
    if (String(params.type || "").trim() === "usersList") {
      return handleUsersList_();
    }
    if (String(params.type || "").trim() === "usageSummary") {
      return handleUsageSummary_();
    }
    if (String(params.type || "").trim() === "provisionUsers") {
      return handleProvisionUsers_(String(params.key || "").trim());
    }

    var dsrName = String(params.dsrName || "").trim();
    var dateKey = String(params.date || "").trim();
    var monthKey = String(params.month || "").trim();
    // lastMonthKey added 19 Sep 2026 for the Target-Balance panel's
    // "Last month" comparison chip - the client computes and passes the
    // previous calendar month's own "yyyy-MM" key; optional, so an older
    // cached front-end that never sends it just gets lastMonthLtr: 0 back
    // (same "purely additive" pattern as every other param here).
    var lastMonthKey = String(params.lastMonthKey || "").trim();
    if (!dsrName) return jsonOut_({ ok: false, error: "dsrName required" });

    var rows = readAllRows_(getSheet_());
    var todayLtr = 0;
    var monthLtr = 0;
    var lastMonthLtr = 0;
    var todaysOrders = [];
    // "MB" (dealers billed this month), added 21 Sep 2026 - the owner's
    // own correction after seeing the Tally-based version: this panel's
    // Sales figures are ALL app-punched (todayLtr/monthLtr above), so MB
    // should be too, not a Tally number mixed into an otherwise all-app
    // panel. A plain {dealerName: true} set of every distinct dealer with
    // an order this month, straight from the same Orders-sheet rows
    // monthLtr is already summing - no Tally/watcher involvement at all,
    // unlike the dashboard's own MB (which stays Tally-based there,
    // unchanged) or this DSR app's Collection MB (Tally-live, unchanged -
    // Collection itself has always been Tally-live here, never punched).
    var monthDealerSet = {};
    // Added 22 Sep 2026 for the Sales App's own "part number covered" and
    // "top selling SKU" (owner's request, same wording as the dashboard's
    // versions of these - but THIS is the app-punched twin, matching the
    // app-punched convention its Sales panel already uses everywhere else,
    // not a Tally pull like the dashboard's DSR-item breakdown). A plain
    // {itemName: ltr} total across every order this DSR punched this
    // month, built from the SAME Items JSON column every order already
    // stores (see collectOrderItemsForSave() client-side, {id, name, qty,
    // unit, ltr} per line) - no new column, no new request shape.
    var monthItemTotals = {};

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (row.deleted) continue;
      if (row.dsrName !== dsrName) continue;
      if (monthKey && row.monthKey === monthKey) {
        monthLtr += row.totalLtr;
        if (row.dealer) monthDealerSet[row.dealer] = true;
        var miItems = [];
        try { miItems = JSON.parse(row.itemsJson || "[]"); } catch (miErr) { miItems = []; }
        for (var mi = 0; mi < miItems.length; mi++) {
          var miName = miItems[mi] && miItems[mi].name;
          var miLtr = miItems[mi] && Number(miItems[mi].ltr);
          if (!miName || isNaN(miLtr)) continue;
          monthItemTotals[miName] = (monthItemTotals[miName] || 0) + miLtr;
        }
      }
      if (lastMonthKey && row.monthKey === lastMonthKey) {
        lastMonthLtr += row.totalLtr;
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

    // --- Past Orders date-range search (added 12 Sep 2026) ---
    var historyFrom = String(params.historyFrom || "").trim();
    var historyTo = String(params.historyTo || "").trim();
    var historyDateParam = String(params.historyDate || "").trim();
    var dealerSearch = String(params.dealerSearch || "").trim().toLowerCase();

    // Back-compat: an older cached front-end that only ever sends
    // historyDate (no historyFrom/historyTo) gets the exact single-day
    // lookup it always claimed to do - both bounds pinned to that one date.
    if (!historyFrom && !historyTo && historyDateParam) {
      historyFrom = historyDateParam;
      historyTo = historyDateParam;
    }

    var historyOrders = [];
    if (historyFrom || historyTo || dealerSearch) {
      for (var hi = 0; hi < rows.length; hi++) {
        var hrow = rows[hi];
        if (hrow.deleted) continue;
        if (hrow.dsrName !== dsrName) continue;
        // Date Key is "yyyy-MM-dd" text, which sorts/compares correctly as
        // a plain string - no need to parse it into a real Date here.
        if (historyFrom && hrow.dateKey < historyFrom) continue;
        if (historyTo && hrow.dateKey > historyTo) continue;
        if (dealerSearch && hrow.dealer.toLowerCase().indexOf(dealerSearch) === -1) continue;

        var hItems = [];
        try { hItems = JSON.parse(hrow.itemsJson || "[]"); } catch (hErr) { hItems = []; }
        historyOrders.push({
          id: hrow.orderId,
          dealer: hrow.dealer,
          deliveryDate: hrow.deliveryDate,
          dateKey: hrow.dateKey,
          monthKey: hrow.monthKey,
          totalLtr: hrow.totalLtr,
          items: hItems,
          orderText: hrow.orderText,
          createdAt: hrow.createdAt,
          updatedAt: hrow.updatedAt
        });
      }
      historyOrders.sort(function (a, b) {
        var byDate = String(b.dateKey || "").localeCompare(String(a.dateKey || ""));
        if (byDate !== 0) return byDate;
        return String(b.createdAt || "").localeCompare(String(a.createdAt || ""));
      });
    }

    return jsonOut_({
      ok: true,
      todayLtr: todayLtr,
      monthLtr: monthLtr,
      lastMonthLtr: lastMonthLtr,
      monthDealerCount: Object.keys(monthDealerSet).length,
      monthItemTotals: monthItemTotals,
      todaysOrders: todaysOrders,
      historyOrders: historyOrders
    });
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
 *   { action: "liveStockUpdate",        key, items: [...] }   — see "Live Tally Sync" above
 *   { action: "liveOutstandingUpdate",  key, dealers: [...] } — see "Live Tally Sync" above
 *   { action: "liveDealerRosterUpdate", key, roster: [...] }  — see "Live Tally Sync" above
 *   { action: "liveCollectionUpdate",    key, collection: [...] } — see "Live Tally Sync" above
 *   { action: "liveCollectionByDealerUpdate", key, collection: [...] } — see "Live Tally Sync" above
 *   { action: "liveDaybookUpdate", key, rows: [...] } — see handleLiveDaybookUpdate_() above;
 *       read back via GET ?type=daybookMonth&month=yyyy-MM — see handleDaybookMonthGet_() above
 *   { action: "logAsk", source, dsrName, question, understood }  — see handleLogAsk_() above
 */
function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ ok: false, error: "Bad request body" });
  }

  // Live-sync actions are handled before the order-saving lock below - they
  // touch entirely separate sheets (LiveStock/LiveOutstanding/
  // LiveDealerRoster, never "Orders"), so there's no reason to make a Tally
  // poll wait behind (or block) a DSR's order save, or vice versa.
  if (body.action === "liveStockUpdate") return handleLiveStockUpdate_(body);
  if (body.action === "liveOutstandingUpdate") return handleLiveOutstandingUpdate_(body);
  if (body.action === "liveDealerRosterUpdate") return handleLiveDealerRosterUpdate_(body);
  if (body.action === "liveNewDealerUpdate") return handleLiveNewDealerUpdate_(body);
  if (body.action === "liveCollectionUpdate") return handleLiveCollectionUpdate_(body);
  if (body.action === "liveCollectionByDealerUpdate") return handleLiveCollectionByDealerUpdate_(body);
  if (body.action === "liveGroupSummaryUpdate") return handleLiveGroupSummaryUpdate_(body);
  if (body.action === "liveGroupSummaryByDealerUpdate") return handleLiveGroupSummaryByDealerUpdate_(body);
  if (body.action === "liveSixtyPlusBaselineUpdate") return handleLiveSixtyPlusBaselineUpdate_(body);
  if (body.action === "liveSalesUpdate") return handleLiveSalesUpdate_(body);
  if (body.action === "liveSalesByDealerUpdate") return handleLiveSalesByDealerUpdate_(body);
  if (body.action === "liveSalesByItemUpdate") return handleLiveSalesByItemUpdate_(body);
  if (body.action === "liveSalesByDsrItemUpdate") return handleLiveSalesByDsrItemUpdate_(body);
  if (body.action === "liveSalesTrendUpdate") return handleLiveSalesTrendUpdate_(body);
  if (body.action === "liveSalesRsTrendUpdate") return handleLiveSalesRsTrendUpdate_(body);
  if (body.action === "liveCollectionTrendUpdate") return handleLiveCollectionTrendUpdate_(body);
  if (body.action === "liveDaybookUpdate") return handleLiveDaybookUpdate_(body);
  if (body.action === "logAsk") return handleLogAsk_(body);
  // Monthly card archive (PDF -> Drive), added 28 Sep 2026 - touches only
  // Drive, never "Orders", so it belongs with the other pre-lock actions
  // above for the same reason they're all up here.
  if (body.action === "archiveMonthlyPdf") return handleArchiveMonthlyPdf_(body);
  // Login/heartbeat/admin-password actions, added 26 Sep 2026 - these touch
  // the Users/LoginLog sheets only, never "Orders", so (like the live-sync
  // actions above) there's no reason for them to wait behind the
  // order-saving lock below.
  if (body.action === "login") return handleLogin_(body);
  if (body.action === "heartbeat") return handleHeartbeat_(body);
  if (body.action === "setPassword") return handleSetPassword_(body);

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return jsonOut_({ ok: false, error: "Server busy, try again" });
  }
  try {
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
