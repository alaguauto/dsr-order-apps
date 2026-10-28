/**
 * ⚠ FULL-FILE REPLACEMENT — based on a 15 Oct 2026 Code.gs snapshot plus
 * every Bank Deposit change made through 23 Oct 2026, INCLUDING the date-
 * display fix (dates were showing as "Fri Oct 23 2026 00:00:00 GMT+0530
 * (India Standard Time)" instead of plain dates - now fixed). Safe to
 * paste in and overwrite your WHOLE Code.gs ONLY if you have not hand-
 * edited Code.gs in the Apps Script editor since 15 Oct 2026 — if you
 * have, use one of the paste-in snippets instead so you don't lose that
 * other work.
 *
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
// "Dismissed" - added 6 Oct 2026, alongside the "Not new?" button in every
// DSR app. Some dealers the roster job flags "new" turn out to be
// long-standing dealers with only an old carried-forward balance and no
// real sales voucher on file - the lastSaleBeforeMonth cross-check in
// handleLiveGet_() below can't disprove those (confirmed 6 Oct 2026: 9
// dealers from 3 Oct stuck showing as "new" in every app despite that
// fix). "Y" here hides a row from every app/dashboard permanently (see
// the newDealers-build loop in handleLiveGet_()), set via the new
// dismissNewDealer POST action below. Never cleared automatically - only
// a fresh genuine re-flag from the roster job (handleLiveNewDealerUpdate_
// rewriting that row) resets it.
var LIVE_NEW_DEALERS_HEADERS = ["Key", "DSR Name", "Dealer Name", "First Seen Date", "Updated At", "Dismissed"];
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

// Added 1 Oct 2026 - Petronas Business Summary, LIVE, current month only.
// Jan-Aug 2026 stay a static, owner-confirmed one-time snapshot baked into
// the dashboard (see dashboard_index.html's PETRONAS_SUMMARY_DATA_); the
// owner's explicit choice this round ("FROM SEP THERE IS LIVE DATA, U HAVE
// TO PUT THAT, FROM SEP ONWARD LIVE DATA SHOULD COME") was that the
// current month onward should compute and push itself instead of waiting
// on another manual screenshot/bake-in round. One row PER MONTH, upserted
// by Month Key (not just one row overwritten in place) - once a month
// rolls over, that month's row simply stops being updated and becomes its
// own small piece of history, same "upsert by key" pattern as every other
// live sheet here. See tally_live_watcher.py's fetch_petronas_summary()
// for the exact computation (a current-month-only version of the already-
// confirmed report_petronas_summary()) and
// maybe_push_petronas_summary()/PETRONAS_SUMMARY_POLL_INTERVAL_SECONDS for
// the push cadence (60 min, same heavy-whole-company-pull reasoning as
// Collection/Sales/Group Summary).
var LIVE_PETRONAS_SUMMARY_SHEET = "LivePetronasSummary";
// "Claims CN" added 15 Oct 2026 - see handleLivePetronasSummaryUpdate_()'s
// own comment for the full story. Appended AFTER "Updated At" (not
// alongside Payment, where it conceptually belongs) specifically so every
// row already sitting in this sheet from before 15 Oct 2026 stays
// correctly aligned - inserting it earlier would have shifted "Updated At"
// out of its existing column for every historical row.
var LIVE_PETRONAS_SUMMARY_HEADERS = ["Month Key", "Month Label", "Purchase Ltr", "Sales Ltr", "Sales Value", "Collected", "Purchase Value", "Payment", "Updated At", "Claims CN"];

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

// Added 11 Oct 2026 for the new per-DSR Sales/Collection weekday heatmap
// ("the sales heatmap above Ranking ... ONLY SALES & COLLECTION ... ONLY
// DSR BUT LIVE FROM TALLY"). Per-DSR twins of LIVE_SALES_TREND_SHEET/
// LIVE_COLLECTION_TREND_SHEET above - same upsert pattern, just keyed by
// "DSR||Date" (Column A) instead of plain Date, since a (DSR, date) pair
// is the actual upsert key here, not the date alone. See
// tally_live_watcher.py's daily_totals_by_dsr for the watcher side.
var LIVE_SALES_TREND_BY_DSR_SHEET = "LiveSalesTrendByDsr";
var LIVE_SALES_TREND_BY_DSR_HEADERS = ["DSR", "Date", "Sold Ltr", "Updated At"];
var LIVE_COLLECTION_TREND_BY_DSR_SHEET = "LiveCollectionTrendByDsr";
var LIVE_COLLECTION_TREND_BY_DSR_HEADERS = ["DSR", "Date", "Collected Rs", "Updated At"];

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

// Added 30 Sep 2026 - dealer credit-limit / credit-lock status, read live
// from Tally's own CREDITLIMIT ledger field (the owner's own Lock D
// formula, pushed into Tally natively via Alter Ledger / XML Import Data -
// this sheet reads it, never computes it). CONFIRMED against real data:
// N.S.K. AUTOMOBILES, THENI (OIL) = 587222, RAMYA AUTO STORES, CBM (OIL)
// = 60880, both matching the owner's own independently-known real
// figures exactly (see tally_live_watcher.py's fetch_creditlimit()
// docstring). Own isolated sheet, same "new/unverified field never
// touches an already-trusted sheet" convention as every other Live*
// sheet here - never folded into LiveOutstanding.
var LIVE_CREDITLIMIT_SHEET = "LiveCreditLimit";
var LIVE_CREDITLIMIT_HEADERS = ["Dealer Name", "Credit Limit", "Credit Period", "Updated At"];

// Transient-error retry wrapper - added 6 Oct 2026 after a real failure:
// "Exception: Service Spreadsheets timed out while accessing document
// with id ..." thrown from getLiveSheet_() while running
// seedMasterFileFromRaw() by hand. This is a known Apps Script/Google
// Sheets infrastructure hiccup, NOT a bug in the calling code - it shows
// up more often on a spreadsheet this size, with this much live-sync
// traffic hitting it every few minutes from tally_live_watcher.py plus
// every dashboard/DSR-app request. It is almost always transient: the
// exact same call usually succeeds a few seconds later. Wrapped around
// the two hottest/heaviest Spreadsheet-service call sites (getLiveSheet_
// and bulkUpsertSheet_'s bulk read/write) so a one-off hiccup during a
// live sync, a dashboard load, or a manual run like seedMasterFileFromRaw
// retries itself automatically instead of failing the whole request.
function isTransientSheetsError_(err) {
  var msg = String((err && err.message) || err || "");
  return msg.indexOf("timed out") !== -1 ||
    msg.indexOf("Service Spreadsheets") !== -1 ||
    msg.indexOf("Internal error") !== -1 ||
    msg.indexOf("try again") !== -1;
}
function withSheetsRetry_(fn) {
  var lastErr;
  for (var attempt = 1; attempt <= 3; attempt++) {
    try {
      return fn();
    } catch (err) {
      lastErr = err;
      if (!isTransientSheetsError_(err) || attempt === 3) throw err;
      Utilities.sleep(1500 * attempt); // 1.5s, then 3s
    }
  }
  throw lastErr;
}

function getLiveSheet_(name, headers) {
  return withSheetsRetry_(function () {
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
  });
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
  // Both Spreadsheet-service calls below (the bulk read and the bulk
  // write) are wrapped in withSheetsRetry_ - same transient-timeout
  // retry added 6 Oct 2026 as getLiveSheet_ above; this is the other
  // call site that touches the most rows most often (every live sync
  // from tally_live_watcher.py goes through here).
  var lastRow = sheet.getLastRow();
  var existing = lastRow >= 2
    ? withSheetsRetry_(function () { return sheet.getRange(2, 1, lastRow - 1, numCols).getValues(); })
    : [];
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
    withSheetsRetry_(function () { sheet.getRange(2, 1, existing.length, numCols).setValues(existing); });
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

// POST { action: "liveCreditLimitUpdate", key, dealers: [{name, creditLimit, creditPeriod}, ...] }
// Added 30 Sep 2026 - see LIVE_CREDITLIMIT_HEADERS above. Same
// "explicit clear beats silent skip" convention as
// handleLiveOutstandingUpdate_() above: 0/null/undefined still writes an
// empty cell rather than leaving a stale figure behind, so a ledger whose
// CREDITLIMIT is cleared in Tally correctly shows as cleared here too.
function handleLiveCreditLimitUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_CREDITLIMIT_SHEET, LIVE_CREDITLIMIT_HEADERS);
  var dealers = body.dealers || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_CREDITLIMIT_HEADERS.length, dealers,
    function (d) { return d && d.name ? String(d.name).trim() : null; },
    function (d, now) {
      return [
        d.name,
        (d.creditLimit === null || d.creditLimit === undefined) ? "" : d.creditLimit,
        (d.creditPeriod === null || d.creditPeriod === undefined) ? "" : d.creditPeriod,
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
  // Master File auto-add, added 6 Oct 2026 - a dealer new to Tally reaches
  // this same roster push (its normal once-daily cadence - see the big
  // comment above), so this is also the moment a brand-new dealer can be
  // added to DealerMeta automatically, instead of being invisible in
  // Master File/Dealer Targets until the owner's next monthly Excel
  // re-import. See autoAddMissingDealerMeta_() below.
  autoAddMissingDealerMeta_(rows);
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
      // Dismissed always comes back blank here - this is the watcher's own
      // daily push, not a dismiss action, so this only matters for a
      // dealer whose roster-diff genuinely re-fires after being dismissed
      // (Tally considering them newly-eligible again, a fresh signal worth
      // re-showing, not a bug - see the "Dismissed" header comment above).
      return [key, r.dsrName, r.name, r.firstSeenDate || "", now, ""];
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

// "202609" -> "Sep 2026". Used by handleLiveGet_() to build the
// petronasSummary response's monthLabel field straight from the reliable
// Month Key column, instead of trusting the sheet's own "Month Label" cell
// - see the big comment at that call site for why (Sheets silently
// date-converts a plain "Sep 2026" text value, corrupting it).
var MONTH_LABEL_NAMES_ = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
function monthLabelFromKey_(monthKey6) {
  var s = String(monthKey6 || "").trim();
  if (s.length !== 6) return s;
  var y = s.substring(0, 4);
  var mIdx = parseInt(s.substring(4, 6), 10) - 1;
  if (isNaN(mIdx) || mIdx < 0 || mIdx > 11) return s;
  return MONTH_LABEL_NAMES_[mIdx] + " " + y;
}

// POST { action: "livePetronasSummaryUpdate", key, petronasSummary:
//   {month, monthLabel, purchaseLtr, salesLtr, salesValue, collected,
//    purchaseValue, payment, claimsCn} }
// Added 1 Oct 2026 - see LIVE_PETRONAS_SUMMARY_SHEET above and
// tally_live_watcher.py's fetch_petronas_summary()/
// maybe_push_petronas_summary() for the full story. A single object, not
// an array (one aggregate figure for the whole business, not per-DSR/
// per-dealer) - wrapped in a 1-element array for bulkUpsertSheet_, upserted
// by Month Key so each month gets (and keeps) its own row.
//
// claimsCn added 15 Oct 2026, per the owner's own correction: "FROM SEP DO
// NOT CALCULATE SALES VALUE IF BILLED FOR PETRONAS LUBRICANTS AS IT IS FOR
// CLAIM PURPOSE AFTER CN COMES ONLY WE HAVE TO ADD THAT IN PAYMENT" -
// salesValue (above) now already excludes claim-billed vouchers (fixed in
// tally_live_watcher.py's fetch_petronas_summary()); claimsCn is the real
// Credit Notes Petronas has issued back against those claims, tracked as
// its OWN column on the dashboard per the owner's explicit choice ("Add as
// a separate line") rather than folded into payment above.
function handleLivePetronasSummaryUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_PETRONAS_SUMMARY_SHEET, LIVE_PETRONAS_SUMMARY_HEADERS);
  // Force column B (Month Label) to plain text so Sheets' automatic format
  // stops silently date-converting values like "Sep 2026"/"Oct 2026" - the
  // corruption that showed up as raw ISO timestamps on the dashboard (see
  // the monthLabelFromKey_() comment in handleLiveGet_() for the read-side
  // fix, which no longer depends on this cell anyway, but keeping the raw
  // sheet readable is worth the one extra call). Applied to the whole
  // column, once per push - cheap, and self-healing for any already-
  // corrupted existing rows the next time this runs.
  sheet.getRange(1, 2, Math.max(sheet.getMaxRows(), 2), 1).setNumberFormat("@");
  var r = body.petronasSummary;
  var rows = (r && r.month) ? [r] : [];
  var updated = bulkUpsertSheet_(sheet, LIVE_PETRONAS_SUMMARY_HEADERS.length, rows,
    function (rr) { return rr && rr.month ? String(rr.month).trim() : null; },
    function (rr, now) {
      return [
        rr.month,
        rr.monthLabel || "",
        (rr.purchaseLtr === null || rr.purchaseLtr === undefined) ? "" : rr.purchaseLtr,
        (rr.salesLtr === null || rr.salesLtr === undefined) ? "" : rr.salesLtr,
        (rr.salesValue === null || rr.salesValue === undefined) ? "" : rr.salesValue,
        (rr.collected === null || rr.collected === undefined) ? "" : rr.collected,
        (rr.purchaseValue === null || rr.purchaseValue === undefined) ? "" : rr.purchaseValue,
        (rr.payment === null || rr.payment === undefined) ? "" : rr.payment,
        now,
        (rr.claimsCn === null || rr.claimsCn === undefined) ? "" : rr.claimsCn
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
  // Dual auth, added 30 Sep 2026 at the owner's request ("that reportbot
  // not working so please keep a 'run full archive' download in only my
  // owner id"): monthly_archive.py (the office-PC script) keeps using its
  // body.key === LIVE_SYNC_KEY path unchanged. The new "Run Full Archive"
  // button inside dashboard/index.html cannot send LIVE_SYNC_KEY - that
  // file is served from the public repo, so embedding the key there would
  // leak it to anyone viewing page source. Instead it sends body.auth =
  // {name, password}, re-verified here against the Users sheet exactly
  // like handleSetPassword_() already does for its own admin-only actions,
  // and is only accepted when that user's Role is "Owner" - matching the
  // owner's explicit instruction that this stay Owner-only.
  var usersSheet = getLiveSheet_(USERS_SHEET, USERS_HEADERS);
  if (body.key) {
    if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  } else {
    var archiveAdminName = String((body.auth && body.auth.name) || "").trim();
    var archiveAdminPassword = String((body.auth && body.auth.password) || "");
    var archiveAdminRow = findUserRow_(usersSheet, archiveAdminName);
    if (archiveAdminRow === -1) return jsonOut_({ ok: false, error: "Not authorized" });
    var archiveAdminData = usersSheet.getRange(archiveAdminRow, 1, 1, USERS_HEADERS.length).getValues()[0];
    if (archiveAdminData[1] !== "Owner") return jsonOut_({ ok: false, error: "Not authorized" });
    if (hashPassword_(archiveAdminPassword, archiveAdminData[3]) !== archiveAdminData[2]) return jsonOut_({ ok: false, error: "Not authorized" });
  }
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

// GET ?type=itemGroupSixMoAvg - trailing-6-CLOSED-months average Alt Qty
// (Ltr) per raw Item Group, aggregated from Day Book's own per-month
// sheets (DayBook_yyyy_MM - the one source that actually keeps full
// month-by-month history; LiveSalesByItem, by contrast, only ever holds
// the CURRENT month, overwritten on every sync - see its own header
// comment). Added 6 Oct 2026 for the "Sales by Item Group (Alt Qty)"
// panel's new 6M AVG column (owner: "CAN ADD 6M AVG AFTER ALT QTY").
// Walks backward from the month before the current one, collecting up
// to 6 months that actually have a Day Book sheet (stops early if fewer
// than 6 exist yet - same "average over however many months have real
// data" convention as every other avg6mo* figure on this dashboard,
// rather than hardcoding a start-month constant that would need to stay
// in sync with the dashboard's own DAYBOOK_FIRST_MONTH_). Excludes plain
// "Sales" voucher-type rows, same as handleDaybookMonthGet_() above (the
// unrelated auto-spares business sharing these voucher types). Returns
// RAW item-group sums, not Petronas-filtered/main-group-merged - the
// dashboard's own isPetronasItemGroup_()/daybookMainGroup_() already do
// that exact filtering for "this month"'s Alt Qty, so this keeps that
// one business rule in one place (the dashboard) instead of duplicating
// it here and risking drift.
function handleItemGroupSixMoAvgGet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var todayMonthKey = Utilities.formatDate(new Date(), DISPLAY_TIMEZONE, "yyyy-MM");
  var closedMonths = [];
  var mk = addMonthsToMonthKey_(todayMonthKey, -1);
  // Walk back at most 24 months looking for up to 6 that actually have a
  // Day Book sheet - the 24-month cap is just a sane backstop against an
  // infinite loop, never expected to matter in practice.
  for (var guard = 0; guard < 24 && closedMonths.length < 6; guard++) {
    if (ss.getSheetByName(daybookSheetName_(mk))) closedMonths.push(mk);
    mk = addMonthsToMonthKey_(mk, -1);
  }
  closedMonths.sort(); // oldest first, just for a tidier response

  var sums = {}; // { itemGroupRaw: { sumAltQtyLtr: N, monthsWithData: N } }
  for (var i = 0; i < closedMonths.length; i++) {
    var sheet = ss.getSheetByName(daybookSheetName_(closedMonths[i]));
    if (!sheet) continue;
    var rows = liveSheetRows_(sheet);
    var seenGroupsThisMonth = {};
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      if (!row[0]) continue; // blank Row Key = blank trailing row
      if (String(row[3] || "").trim().toLowerCase() === "sales") continue; // same exclusion as handleDaybookMonthGet_()
      var itemGroup = String(row[7] || "").trim();
      if (!itemGroup) continue;
      var altQty = Number(row[9] || 0);
      if (!sums[itemGroup]) sums[itemGroup] = { sumAltQtyLtr: 0, monthsWithData: 0 };
      sums[itemGroup].sumAltQtyLtr += altQty;
      seenGroupsThisMonth[itemGroup] = true;
    }
    for (var g in seenGroupsThisMonth) sums[g].monthsWithData++;
  }
  var out = {};
  for (var g2 in sums) {
    out[g2] = sums[g2].monthsWithData > 0 ? Math.round(sums[g2].sumAltQtyLtr / sums[g2].monthsWithData) : null;
  }
  return jsonOut_({ ok: true, closedMonths: closedMonths, avgAltQtyLtrByItemGroup: out });
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

// POST { action: "liveSalesTrendByDsrUpdate", key, trend: [{dsrName, date: "yyyy-MM-dd", ltr}, ...] }
// POST { action: "liveCollectionTrendByDsrUpdate", key, trend: [{dsrName, date: "yyyy-MM-dd", rs}, ...] }
// Added 11 Oct 2026 for the new per-DSR Sales/Collection weekday heatmap -
// see LIVE_SALES_TREND_BY_DSR_SHEET/LIVE_COLLECTION_TREND_BY_DSR_SHEET
// above. Same bulk-upsert pattern as the two plain (company-wide) trend
// handlers just above, keyed by "dsrName||date" (same composite-key
// convention as handleLiveSalesByDsrItemUpdate_'s "dsrName||itemName").
//
// BUG FIXED 15 Oct 2026 ("Collection in Tally is 293285 but showing here
// wrongly, sales data also showing wrongly" - the owner's own Tally
// screenshot for 13-Oct-26 vs. the heatmap cell for the same day): this
// comment always SAID "composite-key convention", but the bulkUpsertSheet_
// call below never actually passed the existingKeyFn argument that makes
// that true - so every lookup fell back to bulkUpsertSheet_'s column-A-only
// default (DSR name alone), which never matches a "dsrName||date" key.
// Every poll where today's running total had changed (i.e. nearly every
// poll, for the current day) therefore found no existing row to overwrite
// and appended a brand-new one instead - exactly the "same item shown
// multiple times" bug already diagnosed and fixed for
// handleLiveSalesByDsrItemUpdate_ on 24 Sep 2026, just never applied here
// when this feature was built three weeks later. Each extra duplicate row
// held a real snapshot value (never fabricated data), but
// renderSalesCollectionHeatmap_()'s byDate aggregation in dashboard_index.html
// sums every row for a date with no de-duplication, so a day with several
// duplicate rows was shown inflated by exactly that many extra additions -
// worst for "today" (updated many times a day while this bug was live) and
// invisible for days from before 11 Oct 2026 (seeded once, no later changes
// to re-trigger the append). Fixed the same way
// handleLiveSalesByDsrItemUpdate_ already was: pass existingKeyFn so
// existing rows are matched by their own columns A (DSR) + B (Date), not by
// column A alone. See also readTrendByDsrRows_() below, which now also
// de-dupes defensively on read so any duplicate rows already sitting in the
// sheet from before this fix stop inflating the heatmap immediately,
// without needing the sheets to be manually cleaned first (see
// cleanupDuplicateTrendByDsrRows_() further down for that optional cleanup).
function handleLiveSalesTrendByDsrUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_SALES_TREND_BY_DSR_SHEET, LIVE_SALES_TREND_BY_DSR_HEADERS);
  var rows = body.trend || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_SALES_TREND_BY_DSR_HEADERS.length, rows,
    function (r) { return (r && r.dsrName && r.date) ? (String(r.dsrName).trim() + "||" + String(r.date).trim()) : null; },
    function (r, now) {
      return [
        r.dsrName,
        r.date,
        (r.ltr === null || r.ltr === undefined) ? "" : r.ltr,
        now
      ];
    },
    // existingRow is [DSR Name, Date, Sold Ltr, Updated At] - rebuild the
    // same "dsrName||date" composite key from columns A+B so an existing
    // row for today actually gets overwritten instead of duplicated.
    function (existingRow) { return String(existingRow[0] || "").trim() + "||" + String(existingRow[1] || "").trim(); });
  return jsonOut_({ ok: true, updated: updated });
}

function handleLiveCollectionTrendByDsrUpdate_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var sheet = getLiveSheet_(LIVE_COLLECTION_TREND_BY_DSR_SHEET, LIVE_COLLECTION_TREND_BY_DSR_HEADERS);
  var rows = body.trend || [];
  var updated = bulkUpsertSheet_(sheet, LIVE_COLLECTION_TREND_BY_DSR_HEADERS.length, rows,
    function (r) { return (r && r.dsrName && r.date) ? (String(r.dsrName).trim() + "||" + String(r.date).trim()) : null; },
    function (r, now) {
      return [
        r.dsrName,
        r.date,
        (r.rs === null || r.rs === undefined) ? "" : r.rs,
        now
      ];
    },
    // Same existingRow fix as handleLiveSalesTrendByDsrUpdate_ just above.
    function (existingRow) { return String(existingRow[0] || "").trim() + "||" + String(existingRow[1] || "").trim(); });
  return jsonOut_({ ok: true, updated: updated });
}

// One-time cleanup, added 15 Oct 2026 alongside the existingKeyFn fix above -
// physically removes the duplicate rows the bug already wrote into
// LiveSalesTrendByDsr/LiveCollectionTrendByDsr between 11 and 15 Oct 2026.
// NOT required for correctness (readTrendByDsrRows_() below already de-dupes
// on every read, so the dashboard is already showing the right numbers as
// soon as this file is redeployed) - this is purely sheet hygiene, so the
// two sheets don't keep carrying thousands of dead rows forever. Run this
// ONCE, by hand, from the Apps Script editor: open Code.gs there, pick
// "cleanupDuplicateTrendByDsrRows_" from the function dropdown next to the
// Run button, and click Run. Safe to run more than once - it's a no-op once
// each sheet is already clean (keeps the newest "Updated At" row per
// dsrName+date, deletes the rest).
function cleanupDuplicateTrendByDsrRows_() {
  [
    [LIVE_SALES_TREND_BY_DSR_SHEET, LIVE_SALES_TREND_BY_DSR_HEADERS],
    [LIVE_COLLECTION_TREND_BY_DSR_SHEET, LIVE_COLLECTION_TREND_BY_DSR_HEADERS]
  ].forEach(function (pair) {
    var sheet = getLiveSheet_(pair[0], pair[1]);
    var rows = liveSheetRows_(sheet);
    if (!rows.length) return;
    var bestByKey = {};
    for (var i = 0; i < rows.length; i++) {
      var key = String(rows[i][0] || "").trim() + "||" + String(rows[i][1] || "").trim();
      if (!key) continue;
      var prev = bestByKey[key];
      if (!prev || String(rows[i][3] || "") > String(prev[3] || "")) bestByKey[key] = rows[i];
    }
    var deduped = Object.keys(bestByKey).map(function (k) { return bestByKey[k]; });
    sheet.getRange(2, 1, sheet.getLastRow() - 1, pair[1].length).clearContent();
    if (deduped.length > 0) {
      sheet.getRange(2, 1, deduped.length, pair[1].length).setValues(deduped);
    }
    Logger.log(pair[0] + ": " + rows.length + " rows -> " + deduped.length + " rows (" + (rows.length - deduped.length) + " duplicates removed)");
  });
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
//
// RememberToken/RememberTokenExpiresAt added 1 Oct 2026 - DSRs complained
// the "ask every time, no remember me" rule (26 Sep 2026 design decision,
// still unchanged for the manager dashboard) made the app feel slow to get
// into, since they open it many times a day between dealer visits. A DSR
// app is a personal, one-person device, so a long-lived "stay logged in"
// session is a reasonable trade for that one app type specifically - see
// handleLogin_()/handleLoginWithToken_() below. RememberToken is a random,
// single-use-until-rotated opaque token (never the password itself) stored
// on this row; RememberTokenExpiresAt is an ISO timestamp, refreshed (slid
// forward) every time the token is used, so an actively-used device never
// gets logged out but one left idle for 30+ days asks again. Both are
// cleared on every password change (setUserPassword_() below) and by the
// explicit "Log out" action (handleClearRememberToken_()), so a stolen or
// copied token can't outlive either of those. Through 9 Oct 2026, only
// ever issued for role "DSR" (see handleLogin_()) - the Owner/Manager/Area
// Manager dashboard logins deliberately kept asking every time. Extended
// 10 Oct 2026 (owner's own request: "for dashboad login also can me make
// as dont ask password for everytime") to any role - the dashboard's own
// login form now has its own "Stay signed in on this device" checkbox
// (unchecked by default, since a dashboard account can see company-wide
// figures and the Admin panel, unlike a DSR's single-purpose phone) that
// sends remember:true only when the person ticks it.
var USERS_HEADERS = ["Name", "Role", "PasswordHash", "Salt", "UpdatedAt", "AllowedAreas", "HiddenUsageFor", "RememberToken", "RememberTokenExpiresAt"];
var LOGINLOG_SHEET = "LoginLog";
var LOGINLOG_HEADERS = ["SessionId", "Name", "Role", "App", "LoginAt", "LastSeenAt", "Heartbeats"];

// Real-owner login privacy, added 26 Sep 2026, NARROWED 11 Oct 2026 - this
// used to be "every Role=Owner login is invisible, by role" (see the old
// comments on createLoginSession_/handleHeartbeat_ below), which quietly
// also hid Rani the moment she was created as Role "Owner" for full
// dashboard access (10 Oct 2026) - her logins got ZERO LoginLog rows,
// invisible to literally everyone, Alagu included. Owner's actual ask
// ("alagu owner login is hidden...make rani login show only for me"):
// Alagu's own logins stay untracked as always, but Rani's should now be
// tracked and visible - to Alagu only. So this check is by NAME now, not
// by Role - Rani keeps full Role="Owner" admin/task permissions
// everywhere else in this file (those checks are all still role-based,
// unchanged), only this one specific "never write a LoginLog row" privacy
// carve-out is narrowed to the real owner. The "visible to Alagu only"
// half of the ask is done the same way every other one-off login
// visibility request on this project has been (see HiddenUsageFor's own
// note above "bala also dont want to see anbu login") - add "Rani" to
// every OTHER account's own HiddenUsageFor via the Admin panel (Anbu,
// Bala, Vimal today), so Rani's now-real LoginLog rows show to Alagu
// (whose own HiddenUsageFor stays blank) but nobody else.
var OWNER_PRIVACY_NAMES_ = ["Alagu"];

// "Section visibility (by role)" - added 1 Oct 2026 at the owner's request
// ("i have a idea in admin page write the heading and ask which one to be
// available for whom"), after a restricted Area Manager (Vimal) was still
// seeing Day Book/Collection Group Summary/the Report month picker even
// once his AllowedAreas was confirmed set correctly - rather than keep
// chasing that one account's data, this gives a direct, explicit,
// Owner-set per-ROLE switch for 5 company-wide dashboard sections, fully
// independent of any person's AllowedAreas. One row per role; DSR is
// deliberately excluded (DSRs don't use this dashboard at all).
var SECTION_VIS_SHEET = "DashboardSectionVisibility";
var SECTION_VIS_HEADERS = ["Role", "HiddenSections"];
var SECTION_VIS_ROLES_ = ["Owner", "Manager", "Area Manager"];
// The only section keys this feature knows about - matches the dashboard's
// own SECTION_VIS_KEYS array exactly (dashboard_index.html). Kept to just
// these 5 for now per the owner's own scope ("those 5" when asked whether
// to include more) - adding a 6th later is a matter of adding it to both
// this list and the dashboard's SECTION_VIS_KEYS/sectionHidden_() call
// sites, nothing structural.
var SECTION_VIS_KNOWN_KEYS_ = ["reportMonth", "dayBook", "groupSummary", "stocks", "petronasSummary"];
// Default hidden-sections list per role, used ONLY to seed the sheet the
// very first time it's read (see getSectionVisibilityMap_() below) and as
// a fallback for any role whose row is somehow missing afterwards. The
// Area Manager default was originally set (1 Oct 2026) to exactly that
// day's existing hardcoded hide-list, so turning this feature on was a
// pure no-op until the Owner actively changed a role's list from the
// Admin panel.
//
// 'groupSummary' removed from Area Manager's default 2 Oct 2026, at the
// owner's request ("Collection Group Summary to show his area only") -
// it's no longer a plain on/off switch for this role, it's now ALWAYS
// shown but scoped to the Area Manager's own dealers (see
// restrictedGroupDealers_()/restrictedGroupTotals_() in
// dashboard_index.html, used by both renderGroupSummaryTable_() and
// runGroupSummaryExport_()) - a data-availability scope tied to
// AllowedAreas, same as the Trend charts and Last Month report's own
// company-wide totals, not a configurable Admin-panel permission. The
// 'groupSummary' key itself is left in SECTION_VIS_KNOWN_KEYS_/the Admin
// panel's own checkbox, in case the Owner ever wants to hide it outright
// for some future role instead.
//
// IMPORTANT for an already-running spreadsheet: this default only seeds
// a BRAND NEW DashboardSectionVisibility sheet. If Area Manager's row was
// already seeded under the old default (any time since 1 Oct 2026), its
// stored HiddenSections value still has 'groupSummary' in it and this
// code change alone won't remove it - the Owner needs to open Admin ->
// Section visibility (by role), pick Area Manager, uncheck "Collection
// Group Summary", and save once.
var SECTION_VIS_DEFAULTS_ = {
  "Owner": "",
  "Manager": "",
  "Area Manager": "reportMonth,dayBook,stocks,petronasSummary"
};

// Reads the DashboardSectionVisibility sheet into {role: "comma,separated,
// keys"}. Seeds one row per SECTION_VIS_ROLES_ (using SECTION_VIS_DEFAULTS_)
// the first time this sheet has no data rows yet - getLiveSheet_() only
// creates the header row, never data, so seeding happens here, once, the
// first time anyone (a dashboard login or the Admin panel) actually asks.
// Falls back to the default for any role whose row is missing even after
// seeding (e.g. a future new role added to SECTION_VIS_ROLES_ later).
function getSectionVisibilityMap_() {
  var sheet = getLiveSheet_(SECTION_VIS_SHEET, SECTION_VIS_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SECTION_VIS_ROLES_.forEach(function (role) {
      sheet.appendRow([role, SECTION_VIS_DEFAULTS_[role] || ""]);
    });
    lastRow = sheet.getLastRow();
  }
  var map = {};
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, SECTION_VIS_HEADERS.length).getValues();
    for (var i = 0; i < data.length; i++) {
      var role = String(data[i][0] || "").trim();
      if (!role) continue;
      map[role] = String(data[i][1] || "").trim();
    }
  }
  SECTION_VIS_ROLES_.forEach(function (role) {
    if (!(role in map)) map[role] = SECTION_VIS_DEFAULTS_[role] || "";
  });
  return map;
}

// GET ?type=sectionVisibility - current per-role hidden-sections map, for
// the Admin panel's "Section visibility (by role)" block to load and show.
// Open read, no auth, same as ?type=usersList - it reveals only role-level
// config (which sections are on/off), never any password/user data.
function handleSectionVisibilityGet_() {
  return jsonOut_({ ok: true, visibility: getSectionVisibilityMap_() });
}

// POST {action:"setSectionVisibility", role, hiddenSections, auth:{name,password}}
// or {action:"setSectionVisibility", role, hiddenSections, key:LIVE_SYNC_KEY}
// Same two-path auth as handleSetPassword_() above (master key OR the
// current Owner's own name+password re-entered fresh) - never a trusted
// client-side flag. hiddenSections is a comma-separated subset of
// SECTION_VIS_KNOWN_KEYS_; unknown keys are dropped rather than rejected,
// so a slightly stale client can't wedge a role's row.
function handleSetSectionVisibility_(body) {
  var role = String(body.role || "").trim();
  if (SECTION_VIS_ROLES_.indexOf(role) === -1) {
    return jsonOut_({ ok: false, error: "Unknown role" });
  }

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

  var requested = String(body.hiddenSections || "").split(",").map(function (s) { return s.trim(); }).filter(function (s) { return s; });
  var cleaned = requested.filter(function (key) { return SECTION_VIS_KNOWN_KEYS_.indexOf(key) !== -1; });
  var hiddenSections = cleaned.join(",");

  var sheet = getLiveSheet_(SECTION_VIS_SHEET, SECTION_VIS_HEADERS);
  var lastRow = sheet.getLastRow();
  var targetRow = -1;
  if (lastRow >= 2) {
    var roles = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < roles.length; i++) {
      if (String(roles[i][0] || "").trim() === role) { targetRow = i + 2; break; }
    }
  }
  if (targetRow === -1) {
    sheet.appendRow([role, hiddenSections]);
  } else {
    sheet.getRange(targetRow, 2).setValue(hiddenSections);
  }
  return jsonOut_({ ok: true });
}

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
//
// newPassword may now be blank for an EXISTING user - added 11 Oct 2026
// alongside OWNER_PRIVACY_NAMES_ (see its comment above): making Rani's
// logins visible "to Alagu only" means adding "Rani" to Anbu/Bala/Vimal's
// own HiddenUsageFor via this same Admin panel form, and forcing a full
// password reset on three accounts just to tweak that one field was pure
// friction with no security benefit. A blank newPassword on an EXISTING
// user now leaves PasswordHash/Salt AND any remembered-device token
// completely untouched - only Role/AllowedAreas/HiddenUsageFor change. A
// brand NEW user still always needs a real password (handleSetPassword_
// enforces this before calling here) - there's no existing password to
// leave alone.
function setUserPassword_(usersSheet, name, role, newPassword, allowedAreas, hiddenUsageFor) {
  var row = findUserRow_(usersSheet, name);
  var now = new Date().toISOString();
  if (row === -1) {
    var newSalt = Utilities.getUuid();
    var newHash = hashPassword_(newPassword, newSalt);
    usersSheet.appendRow([name, role || "DSR", newHash, newSalt, now, allowedAreas || "", hiddenUsageFor || "", "", ""]);
    return;
  }
  var existingRole = usersSheet.getRange(row, 2).getValue();
  var existingHash = usersSheet.getRange(row, 3).getValue();
  var existingSalt = usersSheet.getRange(row, 4).getValue();
  var existingUpdatedAt = usersSheet.getRange(row, 5).getValue();
  var existingAreas = usersSheet.getRange(row, 6).getValue();
  var existingHidden = usersSheet.getRange(row, 7).getValue();
  var existingToken = usersSheet.getRange(row, 8).getValue();
  var existingTokenExpiry = usersSheet.getRange(row, 9).getValue();
  var areasToWrite = (allowedAreas === undefined) ? existingAreas : allowedAreas;
  var hiddenToWrite = (hiddenUsageFor === undefined) ? existingHidden : hiddenUsageFor;
  if (newPassword) {
    // An actual password change clears RememberToken/RememberTokenExpiresAt
    // (1 Oct 2026) - a device "stayed logged in" under the OLD password
    // must not keep working after the Owner resets it; that device will
    // simply fall back to asking for the new password next time it opens.
    var salt = Utilities.getUuid();
    var hash = hashPassword_(newPassword, salt);
    usersSheet.getRange(row, 2, 1, 8).setValues([[role || existingRole, hash, salt, now, areasToWrite, hiddenToWrite, "", ""]]);
  } else {
    usersSheet.getRange(row, 2, 1, 8).setValues([[role || existingRole, existingHash, existingSalt, existingUpdatedAt, areasToWrite, hiddenToWrite, existingToken, existingTokenExpiry]]);
  }
}

// Shared by handleLogin_() and handleLoginWithToken_() below - both a real
// password login and a silent remembered-device login create exactly the
// same kind of LoginLog session, so usage tracking/time-in-app is identical
// either way (a DSR that stays logged in is still fully tracked - "stay
// logged in" only skips re-typing the password, never the tracking the
// owner asked for on 26 Sep 2026). Returns the new sessionId; writes
// nothing for the real owner (OWNER_PRIVACY_NAMES_ - see its own comment
// above), same privacy rule as always, but anyone else with Role "Owner"
// (e.g. Rani) now DOES get a normal row, same as a Manager.
function createLoginSession_(name, role, app) {
  var sessionId = Utilities.getUuid();
  var now = new Date().toISOString();
  if (OWNER_PRIVACY_NAMES_.indexOf(name) === -1) {
    var logSheet = getLiveSheet_(LOGINLOG_SHEET, LOGINLOG_HEADERS);
    logSheet.appendRow([sessionId, name, role, app, now, now, 1]);
  }
  return sessionId;
}

// How long a remembered DSR-app device stays logged in without being used -
// slides forward on every silent re-login (handleLoginWithToken_()), so an
// actively-used phone never gets logged out; one left idle this long asks
// again. 1 Oct 2026, owner-confirmed ("stay logged in for ~30 days").
var REMEMBER_TOKEN_DAYS_ = 30;

function newRememberTokenExpiry_() {
  return new Date(Date.now() + REMEMBER_TOKEN_DAYS_ * 24 * 60 * 60 * 1000).toISOString();
}

// POST {action:"login", name, password, app, remember:true}
// `remember` works for any role as of 10 Oct 2026 (see USERS_HEADERS' own
// note on RememberToken above) - a manager/Area Manager/Owner login gets a
// token too whenever the client explicitly asks for one. Through 9 Oct
// 2026 this was hard-restricted to role "DSR" only; lifted because the
// dashboard's own login now offers the same "stay signed in" choice,
// opt-in via its own checkbox rather than always-on like the DSR apps.
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
  var sessionId = createLoginSession_(data[0], role, app);
  var out = { ok: true, sessionId: sessionId, name: data[0], role: role, allowedAreas: allowedAreas, hiddenUsageFor: hiddenUsageFor };
  if (body.remember === true) {
    var token = Utilities.getUuid();
    usersSheet.getRange(row, 8, 1, 2).setValues([[token, newRememberTokenExpiry_()]]);
    out.rememberToken = token;
  }
  // hiddenSections, added 1 Oct 2026 - this role's own current list from
  // DashboardSectionVisibility (see getSectionVisibilityMap_()/
  // SECTION_VIS_DEFAULTS_ above), looked up fresh on every login so an
  // Owner's Admin-panel change takes effect the next time each person logs
  // in, no redeploy needed. "" for a role (e.g. DSR, which never logs into
  // this dashboard) just means nothing is hidden for it.
  out.hiddenSections = getSectionVisibilityMap_()[role] || "";
  // allowedAreas ("" when unrestricted) tells the dashboard which DSR
  // cards/totals a restricted Area Manager (e.g. Vimal) should be limited
  // to - see the "AllowedAreas" note on USERS_HEADERS above. hiddenUsageFor
  // ("" when nobody is hidden) is this same user's OWN per-viewer usage-
  // report exclusion list (28 Sep 2026) - see the "HiddenUsageFor" note on
  // USERS_HEADERS above.
  return jsonOut_(out);
}

// POST {action:"loginWithToken", token, app} - silent re-login for any app
// (a DSR app, or the dashboard as of 10 Oct 2026) that has a remembered
// device token in localStorage (1 Oct 2026, see the RememberToken note on
// USERS_HEADERS above). Rotates the token and slides its expiry forward on
// every successful use, and creates a normal LoginLog session exactly like
// handleLogin_() - a remembered device is just as fully tracked as a
// freshly-typed-password one (Owner excepted, same as always).
function handleLoginWithToken_(body) {
  var token = String(body.token || "").trim();
  var app = String(body.app || "");
  if (!token) return jsonOut_({ ok: false, error: "invalid_token" });
  var usersSheet = getLiveSheet_(USERS_SHEET, USERS_HEADERS);
  var lastRow = usersSheet.getLastRow();
  if (lastRow < 2) return jsonOut_({ ok: false, error: "invalid_token" });
  var data = usersSheet.getRange(2, 1, lastRow - 1, USERS_HEADERS.length).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][7] || "") !== token) continue;
    var row = i + 2;
    var name = data[i][0], role = data[i][1];
    var allowedAreas = String(data[i][5] || "").trim();
    var hiddenUsageFor = String(data[i][6] || "").trim();
    var expiresAt = data[i][8];
    var expired = !expiresAt || new Date(expiresAt).getTime() < Date.now();
    if (expired) {
      usersSheet.getRange(row, 8, 1, 2).setValues([["", ""]]);
      return jsonOut_({ ok: false, error: "expired" });
    }
    var newToken = Utilities.getUuid();
    usersSheet.getRange(row, 8, 1, 2).setValues([[newToken, newRememberTokenExpiry_()]]);
    var sessionId = createLoginSession_(name, role, app);
    return jsonOut_({
      ok: true, sessionId: sessionId, name: name, role: role,
      allowedAreas: allowedAreas, hiddenUsageFor: hiddenUsageFor, rememberToken: newToken
    });
  }
  return jsonOut_({ ok: false, error: "invalid_token" });
}

// POST {action:"clearRememberToken", token} - any app's own "Log out"
// control (1 Oct 2026, dashboard included as of 10 Oct 2026). Invalidates
// the token server-side too (not just a
// localStorage.removeItem on the client) so a copied/leaked token can't be
// replayed after an explicit logout. Idempotent and silent either way - an
// already-invalid/unknown token is not an error from the client's point of
// view, it just means there's nothing left to clear.
function handleClearRememberToken_(body) {
  var token = String(body.token || "").trim();
  if (!token) return jsonOut_({ ok: true });
  var usersSheet = getLiveSheet_(USERS_SHEET, USERS_HEADERS);
  var lastRow = usersSheet.getLastRow();
  if (lastRow >= 2) {
    var tokens = usersSheet.getRange(2, 8, lastRow - 1, 1).getValues();
    for (var i = 0; i < tokens.length; i++) {
      if (String(tokens[i][0] || "") === token) {
        usersSheet.getRange(i + 2, 8, 1, 2).setValues([["", ""]]);
        break;
      }
    }
  }
  return jsonOut_({ ok: true });
}

// POST {action:"heartbeat", sessionId, name, role, app} - "still here" ping,
// sent every few minutes only while the tab is visible. Updates that
// session's own LastSeenAt/Heartbeats in place rather than appending a new
// row each time.
function handleHeartbeat_(body) {
  var sessionId = String(body.sessionId || "");
  if (!sessionId) return jsonOut_({ ok: false, error: "sessionId required" });
  // Same real-owner exclusion as createLoginSession_() above (by NAME, not
  // Role, since 11 Oct 2026 - see OWNER_PRIVACY_NAMES_'s own comment) - a
  // privacy-exempt session was never given a LoginLog row to begin with,
  // so every heartbeat for one is a deliberate no-op rather than "recreate
  // a minimal row" (which would otherwise silently start tracking them
  // again from here). Anyone else with Role "Owner" (e.g. Rani) heartbeats
  // normally, same as a Manager.
  if (OWNER_PRIVACY_NAMES_.indexOf(String(body.name || "").trim()) !== -1) return jsonOut_({ ok: true });
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
  if (!targetName) return jsonOut_({ ok: false, error: "targetName required" });
  if (newPassword && newPassword.length < 4) return jsonOut_({ ok: false, error: "Password must be at least 4 characters" });

  var usersSheet = getLiveSheet_(USERS_SHEET, USERS_HEADERS);
  // newPassword is only mandatory for a brand-new account - there's no
  // existing password to leave alone. Added 11 Oct 2026 alongside
  // setUserPassword_'s own metadata-only-save support (see its comment) -
  // an EXISTING user's Role/AllowedAreas/HiddenUsageFor can now be changed
  // here without also resetting their password.
  if (findUserRow_(usersSheet, targetName) === -1 && !newPassword) {
    return jsonOut_({ ok: false, error: "newPassword required for a new user" });
  }

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

// ============================================================
// Monthly Freeze - "all cards available for Last Month" (1 Oct 2026)
// ============================================================
// Owner's request, verbatim: "why cant all the month end card/tile will
// make available for the last month selection as it is... easy for
// review dsr meeting" - then, on being asked, confirmed EVERY card
// (not a smaller priority subset) and an AUTOMATIC monthly freeze, same
// schedule as the existing monthly PDF archive.
//
// Several cards already have real month-by-month history for free and
// need no new work: Day Book (its own per-month sheet), Collection
// Group Summary, Petronas Business Summary (both keep one row PER
// MONTH, never overwritten), and the Sales/Collection trend sheets
// (plain daily time series, never pruned) - this is exactly why those
// were the only things the OLD, narrower Last Month view could ever
// show. Everything else - Stock, dealer Outstanding/aging, Sales &
// Collection by DSR/Dealer/Item, Dealer Roster, New Dealers, and the
// 60+ days overdue baseline - lives in "Live*" sheets the Tally watcher
// OVERWRITES every sync (one row per item/dealer/DSR, no month
// dimension at all) - there has never been anywhere a past month's true
// numbers for these could be read back from.
//
// FREEZE_LIVE_CATEGORIES_ lists exactly those overwrite-style sheets.
// Once a month, handleFreezeMonthSnapshot_() below copies each one's
// CURRENT rows into a twin "Freeze_<SheetName>" archive sheet tagged
// with that Month Key (never touching a previous month's own frozen
// rows) - called by monthly_archive.py on the last day of the month,
// the same trigger/schedule as the existing PDF archive, right before
// it captures screenshots, so the frozen numbers and that month's PDF
// snapshot agree. GET ?type=monthlyFreeze&month=yyyy-MM
// (handleMonthlyFreezeGet_() below) then rebuilds the EXACT SAME
// payload shape as ?type=live, but sourced from those frozen archive
// sheets instead of the live ones - see buildLivePayload_()'s own
// `rowSource` parameter just below. The dashboard swaps this payload
// into the same globals fetchLive() already fills (see
// applyLivePayloadToState_() in dashboard_index.html), so every
// existing card-rendering function works completely unchanged against
// a frozen month - Stocks/Ranking/Needs attention/the per-DSR
// cards/etc. needed zero changes of their own.
var FREEZE_LIVE_CATEGORIES_ = [
  { sheet: LIVE_STOCK_SHEET, headers: LIVE_STOCK_HEADERS },
  { sheet: LIVE_OUTSTANDING_SHEET, headers: LIVE_OUTSTANDING_HEADERS },
  { sheet: LIVE_CREDITLIMIT_SHEET, headers: LIVE_CREDITLIMIT_HEADERS },
  { sheet: LIVE_DEALER_ROSTER_SHEET, headers: LIVE_DEALER_ROSTER_HEADERS },
  { sheet: LIVE_NEW_DEALERS_SHEET, headers: LIVE_NEW_DEALERS_HEADERS },
  { sheet: LIVE_COLLECTION_SHEET, headers: LIVE_COLLECTION_HEADERS },
  { sheet: LIVE_COLLECTION_BY_DEALER_SHEET, headers: LIVE_COLLECTION_BY_DEALER_HEADERS },
  { sheet: LIVE_SALES_SHEET, headers: LIVE_SALES_HEADERS },
  { sheet: LIVE_SALES_BY_DEALER_SHEET, headers: LIVE_SALES_BY_DEALER_HEADERS },
  { sheet: LIVE_SALES_BY_ITEM_SHEET, headers: LIVE_SALES_BY_ITEM_HEADERS },
  { sheet: LIVE_SALES_BY_DSR_ITEM_SHEET, headers: LIVE_SALES_BY_DSR_ITEM_HEADERS },
  { sheet: LIVE_SIXTYPLUS_BY_DEALER_SHEET, headers: LIVE_SIXTYPLUS_BY_DEALER_HEADERS }
];
var MONTHLY_FREEZE_PREFIX_ = "Freeze_";
var MONTHLY_FREEZE_LOG_SHEET = "MonthlyFreezeLog";
var MONTHLY_FREEZE_LOG_HEADERS = ["Month Key", "FrozenAt"];

function monthlyFreezeSheetName_(liveSheetName) {
  return MONTHLY_FREEZE_PREFIX_ + liveSheetName;
}

// Copies every CURRENT row of `liveSheetName` into its twin Freeze_*
// archive sheet, tagged with `monthKey`. Rewrites the whole archive
// sheet in one bulk read + one bulk write (never a per-row
// append/delete) - safe to re-run for the SAME month (a `--force`
// re-run from monthly_archive.py) since any existing rows already
// tagged with that exact month are replaced, not duplicated; every
// OTHER month's already-frozen rows are left completely untouched.
function freezeLiveSheetForMonth_(liveSheetName, liveHeaders, monthKey) {
  var liveRows = liveSheetRows_(getLiveSheet_(liveSheetName, liveHeaders));
  var archiveHeaders = ["Month Key"].concat(liveHeaders);
  var archiveSheet = getLiveSheet_(monthlyFreezeSheetName_(liveSheetName), archiveHeaders);
  var lastRow = archiveSheet.getLastRow();
  var kept = [];
  if (lastRow >= 2) {
    var existing = archiveSheet.getRange(2, 1, lastRow - 1, archiveHeaders.length).getValues();
    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i][0] || "").trim() !== monthKey) kept.push(existing[i]);
    }
  }
  var fresh = liveRows.map(function (r) { return [monthKey].concat(r); });
  var all = kept.concat(fresh);
  if (lastRow >= 2) archiveSheet.getRange(2, 1, lastRow - 1, archiveHeaders.length).clearContent();
  if (all.length) archiveSheet.getRange(2, 1, all.length, archiveHeaders.length).setValues(all);
  return fresh.length;
}

// Reads back `monthKey`'s frozen rows for `liveSheetName`, in the SAME
// shape liveSheetRows_() returns for a live sheet (Month Key column
// stripped) - so frozenRowSource_() below is a drop-in replacement for
// liveRowSource_() from buildLivePayload_()'s point of view.
function readFrozenSheetForMonth_(liveSheetName, liveHeaders, monthKey) {
  var archiveHeaders = ["Month Key"].concat(liveHeaders);
  var archiveSheet = getLiveSheet_(monthlyFreezeSheetName_(liveSheetName), archiveHeaders);
  var lastRow = archiveSheet.getLastRow();
  if (lastRow < 2) return [];
  var data = archiveSheet.getRange(2, 1, lastRow - 1, archiveHeaders.length).getValues();
  var out = [];
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0] || "").trim() !== monthKey) continue;
    out.push(data[i].slice(1));
  }
  return out;
}

function frozenRowSource_(monthKey) {
  return function (sheetName, headers) { return readFrozenSheetForMonth_(sheetName, headers, monthKey); };
}

// POST {action:"freezeMonthSnapshot", key:LIVE_SYNC_KEY, monthKey?, category?}
// Key-only auth (same as liveStockUpdate etc. above) - this is called by
// monthly_archive.py on the office PC, not a logged-in dashboard user,
// exactly like every other liveXxxUpdate_ action. `monthKey` is
// optional (defaults to the real current month) so a future manual
// backfill/re-freeze could target a specific month if ever needed - the
// normal automatic month-end call never sends it.
//
// `category` added 3 Oct 2026, fixing a real production failure: at real
// data scale (~19,454 rows total across these 12 sheets - LiveStock alone
// is 8884), freezing all 12 categories in ONE call means up to 48
// sequential Sheets API round-trips (read live + read archive + clear +
// write, per category) inside a single HTTP request - monthly_archive.py's
// own run showed this taking 40s-120s+ and never finishing within its
// 120s client timeout, coming back as either a bare 500 (likely the Apps
// Script web app gateway's own shorter response-time budget, well under
// its 6-minute execution budget) or a client-side read-timeout once the
// execution ran past 120s unanswered. Compare handleFreezeDaySnapshot_()
// just below, which stays fast because it only READS these same 12 sheets
// and does exactly ONE write (a single Drive JSON blob) - nothing close
// to this many Sheets API round-trips.
//
// When `category` is one of FREEZE_LIVE_CATEGORIES_'s own sheet names,
// only that one category is frozen (one call's worth of Sheets API work
// instead of twelve) - monthly_archive.py now calls this once per
// category instead of once for everything, so each individual request
// finishes in a few seconds instead of minutes, and one category being
// slow/retried never forces redoing the other eleven. Omitting `category`
// keeps the original all-at-once behavior for backward compatibility (in
// case anything else ever calls this action without it).
function handleFreezeMonthSnapshot_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var monthKey = String(body.monthKey || "").trim() || Utilities.formatDate(new Date(), DISPLAY_TIMEZONE, "yyyy-MM");
  var categoryFilter = String(body.category || "").trim();
  var cats = FREEZE_LIVE_CATEGORIES_;
  if (categoryFilter) {
    cats = FREEZE_LIVE_CATEGORIES_.filter(function (cat) { return cat.sheet === categoryFilter; });
    if (!cats.length) return jsonOut_({ ok: false, error: "Unknown category: " + categoryFilter });
  }
  var counts = {};
  cats.forEach(function (cat) {
    counts[cat.sheet] = freezeLiveSheetForMonth_(cat.sheet, cat.headers, monthKey);
  });
  var frozenAt = new Date().toISOString();
  var logSheet = getLiveSheet_(MONTHLY_FREEZE_LOG_SHEET, MONTHLY_FREEZE_LOG_HEADERS);
  var logLastRow = logSheet.getLastRow();
  var logRow = -1;
  if (logLastRow >= 2) {
    var logKeys = logSheet.getRange(2, 1, logLastRow - 1, 1).getValues();
    for (var i = 0; i < logKeys.length; i++) {
      if (String(logKeys[i][0] || "").trim() === monthKey) { logRow = i + 2; break; }
    }
  }
  if (logRow === -1) {
    logSheet.appendRow([monthKey, frozenAt]);
  } else {
    logSheet.getRange(logRow, 2).setValue(frozenAt);
  }
  return jsonOut_({ ok: true, monthKey: monthKey, frozenAt: frozenAt, counts: counts });
}

// GET ?type=monthlyFreeze&month=yyyy-MM - rebuilds that month's frozen
// payload (same shape as ?type=live) from the Freeze_* archive sheets,
// for the dashboard's Report month / Last Month view. {ok:true,
// found:false} (no payload fields) if that month was never frozen - a
// month before this feature existed, or the current month (which won't
// freeze until it ends) - and the dashboard falls back to its older,
// narrower Last Month view in that case, same as before this feature
// existed, so nothing regresses before the first automatic freeze runs.
// Pulled out of handleMonthlyFreezeGet_() (2 Oct 2026) so
// handlePeriodSummaryGet_() below can ask the same question ("was this
// month ever frozen, and when") without duplicating the MonthlyFreezeLog
// lookup.
function monthlyFreezeFrozenAt_(monthKey) {
  var logSheet = getLiveSheet_(MONTHLY_FREEZE_LOG_SHEET, MONTHLY_FREEZE_LOG_HEADERS);
  var logLastRow = logSheet.getLastRow();
  if (logLastRow < 2) return null;
  var logData = logSheet.getRange(2, 1, logLastRow - 1, 2).getValues();
  for (var i = 0; i < logData.length; i++) {
    if (String(logData[i][0] || "").trim() === monthKey) return logData[i][1];
  }
  return null;
}

function handleMonthlyFreezeGet_(monthKey) {
  if (!monthKey) return jsonOut_({ ok: false, error: "month required" });
  var frozenAt = monthlyFreezeFrozenAt_(monthKey);
  if (!frozenAt) return jsonOut_({ ok: true, found: false });
  var payload = buildLivePayload_(frozenRowSource_(monthKey), monthKey);
  payload.ok = true;
  payload.found = true;
  payload.frozenAt = frozenAt;
  payload.monthKey = monthKey;
  return jsonOut_(payload);
}

// "Auto-archive Master File History" - added 10 Oct 2026, owner's own
// follow-up right after the Sep 2026 Master File re-import: "after each
// month close that month will be automatically saved in the sheet?" -
// now it is. Before this, MasterFileHistory (the permanent Sales Ltr /
// Collection Rs record behind the Master File table and Review-by-DSR)
// only ever got a live-era month's real numbers (Sep 2026 onward - see
// MASTER_FILE_LIVE_START_MONTH_ above) if the owner manually re-exported
// Tally into a fresh MasterFileRaw Excel and re-ran
// seedMasterFileFromRaw() - otherwise it stayed blank forever for that
// month, even though the dashboard itself already shows the real figures
// live from LiveSalesByDealer/LiveCollectionByDealer (or their Monthly
// Freeze archive once the month closes).
//
// Closes that gap automatically: called by monthly_archive.py right
// after freezeMonthSnapshot (above) finishes freezing that month's
// LiveSalesByDealer/LiveCollectionByDealer rows into their Freeze_*
// archive. This reads those SAME just-frozen rows back
// (readFrozenSheetForMonth_ - guarantees it sees exactly the numbers
// Monthly Freeze captured, never a second, possibly-drifted live read)
// and upserts one MasterFileHistory row per dealer for that month -
// "Month Sold Ltr" / "Month Collected" (column index 2 of each sheet -
// same fields masterFileActualsForMonth_() above already reads for
// display) is that dealer's running month-to-date total at freeze time,
// which at a real month-end IS that month's final total.
//
// Safe to call again for the same month (a `--force` re-run, or the
// 31 Oct safety re-run already recommended after the 10 Oct
// partial-month mishap) - bulkUpsertSheet_'s composite dealer||monthKey
// key REPLACES that month's rows rather than duplicating them, exactly
// like seedMasterFileFromRaw()'s own MasterFileHistory upsert does. Never
// touches a year-total row ("2023", "2022", ... monthKeys, Is Year
// Total="Y") since those keys never collide with a real "yyyy-MM" one.
//
// POST {action:"archiveMasterFileHistoryMonth", key:LIVE_SYNC_KEY, monthKey}
// Key-only auth, same as freezeMonthSnapshot/freezeDaySnapshot above -
// called by monthly_archive.py, not a logged-in dashboard user.
function handleArchiveMasterFileHistoryMonth_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var monthKey = String(body.monthKey || "").trim();
  if (!monthKey) return jsonOut_({ ok: false, error: "monthKey required" });

  var salesRows = readFrozenSheetForMonth_(LIVE_SALES_BY_DEALER_SHEET, LIVE_SALES_BY_DEALER_HEADERS, monthKey) || [];
  var collRows = readFrozenSheetForMonth_(LIVE_COLLECTION_BY_DEALER_SHEET, LIVE_COLLECTION_BY_DEALER_HEADERS, monthKey) || [];
  if (!salesRows.length && !collRows.length) {
    return jsonOut_({ ok: false, error: "No frozen LiveSalesByDealer/LiveCollectionByDealer data for " + monthKey + " yet - run freezeMonthSnapshot for this month first." });
  }

  // Same dealer -> {salesLtr, collectionRs} merge masterFileActualsForMonth_
  // above already does for DISPLAY.
  var byDealer = {};
  for (var i = 0; i < salesRows.length; i++) {
    var r = salesRows[i];
    var name = String(r[0] || "").trim();
    if (!name) continue;
    if (!byDealer[name]) byDealer[name] = { salesLtr: null, collectionRs: null };
    byDealer[name].salesLtr = r[2] === "" || r[2] === undefined || r[2] === null ? null : Number(r[2]); // Month Sold Ltr
  }
  for (var j = 0; j < collRows.length; j++) {
    var r2 = collRows[j];
    var name2 = String(r2[0] || "").trim();
    if (!name2) continue;
    if (!byDealer[name2]) byDealer[name2] = { salesLtr: null, collectionRs: null };
    byDealer[name2].collectionRs = r2[2] === "" || r2[2] === undefined || r2[2] === null ? null : Number(r2[2]); // Month Collected
  }

  var updates = [];
  for (var dealerName in byDealer) {
    updates.push({ dealer: dealerName, monthKey: monthKey, salesLtr: byDealer[dealerName].salesLtr, collectionRs: byDealer[dealerName].collectionRs });
  }

  var histSheet = getLiveSheet_(MASTER_FILE_HISTORY_SHEET, MASTER_FILE_HISTORY_HEADERS);
  var written = bulkUpsertSheet_(histSheet, MASTER_FILE_HISTORY_HEADERS.length, updates,
    function (item) { return item.dealer + "||" + item.monthKey; },
    function (item) {
      return [
        item.dealer, item.monthKey,
        item.salesLtr === null ? "" : item.salesLtr,
        item.collectionRs === null ? "" : item.collectionRs,
        "" // Is Year Total - always blank for a real closed month, never "Y"
      ];
    },
    function (existingRow) { return String(existingRow[0] || "").trim() + "||" + String(existingRow[1] || "").trim(); }
  );
  return jsonOut_({ ok: true, monthKey: monthKey, dealersWritten: written });
}

// =======================================================================
// "Daily Freeze" (2 Oct 2026)
// =======================================================================
// Owner's follow-up to Monthly Freeze, after seeing another app's
// Today/Yesterday/This Week/Custom Date picker: can the same
// "last month frozen as it was" idea work for single PAST DAYS, not just
// month-end? Scoped (via AskUserQuestion) to: 90 days of rolling
// retention, single-day views only for this first build (Today stays
// live; Yesterday and any past Custom Date read a frozen snapshot; This
// Week/This Quarter rollups are a later phase, not built here).
//
// STORAGE: deliberately Google Drive JSON files, NOT a Sheets tab like
// Monthly Freeze uses. The per-dealer sheets alone (Outstanding, Credit
// Limit, Sales/Collection by Dealer, 60+ Overdue - thousands of dealers
// combined across all DSRs) make one day's freeze of all 12 categories
// roughly 100,000+ cells; at 90 days of DAILY retention that's close to
// Google Sheets' hard 10-million-cell-per-spreadsheet ceiling, on top of
// everything else already living in this same spreadsheet (the live
// sheets, Monthly Freeze's own archive sheets, Day Book's per-month
// tabs). Monthly Freeze is safe at that storage because it's one
// snapshot per MONTH; Daily Freeze is 30x the write volume, so it gets
// its own one-small-JSON-file-per-day folder in Drive instead - no cell
// limit at any retention length, and pruning is just deleting old files
// rather than rewriting a sheet. Apps Script already has Drive access
// (added for the PDF monthly archive, 28 Sep 2026) - no new permissions
// needed.
//
// Reuses the exact same FREEZE_LIVE_CATEGORIES_ list Monthly Freeze
// already covers (the "overwrite-style" live sheets) and the exact same
// buildLivePayload_(rowSource, refMonthKey) the dashboard's live and
// monthly-frozen views both already go through - a day's frozen payload
// is just a third rowSource implementation, so every existing
// card-rendering function on the dashboard works against it unchanged,
// same reasoning as Monthly Freeze's own design.
var DAILY_FREEZE_RETENTION_DAYS_ = 90;
var DAILY_FREEZE_FOLDER_NAME_ = "AAA Daily Freeze";

function dailyFreezeRootFolder_() {
  return getOrCreateFolder_(DriveApp.getRootFolder(), DAILY_FREEZE_FOLDER_NAME_);
}
function dailyFreezeFileName_(dateKey) {
  return dateKey + ".json";
}
// Overwrite semantics (removes any existing file for this exact date
// first) - same "re-run should replace, never duplicate" convention as
// archiveMonthlyPdf's own PDF uploads and Monthly Freeze's row-replace
// logic, so a re-run (e.g. the office script firing twice in one day)
// is always safe.
function writeDailyFreezeFile_(dateKey, fileObj) {
  var folder = dailyFreezeRootFolder_();
  var fileName = dailyFreezeFileName_(dateKey);
  var existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);
  var blob = Utilities.newBlob(JSON.stringify(fileObj), "application/json", fileName);
  folder.createFile(blob);
}
function readDailyFreezeFile_(dateKey) {
  var folder = dailyFreezeRootFolder_();
  var it = folder.getFilesByName(dailyFreezeFileName_(dateKey));
  if (!it.hasNext()) return null;
  try {
    return JSON.parse(it.next().getBlob().getDataAsString());
  } catch (e) {
    return null;
  }
}
// Pure "yyyy-MM-dd" + N days arithmetic via Date.UTC/getUTC*, not a
// plain `new Date(y, m, d)` - avoids any dependence on the Apps Script
// project's own default timezone (which need not match
// DISPLAY_TIMEZONE) for what is otherwise just whole-day offset math.
function addDaysToDateKey_(dateKey, deltaDays) {
  var p = dateKey.split("-");
  var d = new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2])) + deltaDays * 86400000);
  return d.getUTCFullYear() + "-" + String(d.getUTCMonth() + 1).padStart(2, "0") + "-" + String(d.getUTCDate()).padStart(2, "0");
}
// Deletes any Daily Freeze file older than the retention window. Run
// once per freeze call - this folder only ever holds about
// DAILY_FREEZE_RETENTION_DAYS_ files at a time, so listing all of them
// is cheap. This is Daily Freeze's entire storage-growth control; unlike
// a Sheets tab there is no other natural place to prune.
function pruneDailyFreezeFiles_(todayDateKey) {
  var cutoff = addDaysToDateKey_(todayDateKey, -DAILY_FREEZE_RETENTION_DAYS_);
  var files = dailyFreezeRootFolder_().getFiles();
  while (files.hasNext()) {
    var f = files.next();
    var m = /^(\d{4}-\d{2}-\d{2})\.json$/.exec(f.getName());
    if (m && m[1] < cutoff) f.setTrashed(true);
  }
}
// rowSource built from an already-read-and-parsed Daily Freeze file
// object (not from dateKey directly) so handleDailyFreezeGet_() below
// never reads the same Drive file twice.
function frozenRowSourceFromFile_(file) {
  return function (sheetName, headers) {
    return (file && file.categories && file.categories[sheetName]) || [];
  };
}

// POST {action:"freezeDaySnapshot", key:LIVE_SYNC_KEY, dateKey?}
// Key-only auth (same pattern as freezeMonthSnapshot) - called once a
// day by monthly_archive.py on the office PC, right alongside (but NOT
// gated behind) its existing month-end checks - see that script's own
// freeze_day_snapshot() comment. `dateKey` is optional (defaults to the
// real current day) so a future manual backfill could target a specific
// date if ever needed - the normal daily call never sends it.
function handleFreezeDaySnapshot_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var dateKey = String(body.dateKey || "").trim() || Utilities.formatDate(new Date(), DISPLAY_TIMEZONE, "yyyy-MM-dd");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) return jsonOut_({ ok: false, error: "Bad dateKey, expected yyyy-MM-dd" });
  var categories = {};
  var counts = {};
  FREEZE_LIVE_CATEGORIES_.forEach(function (cat) {
    var rows = liveSheetRows_(getLiveSheet_(cat.sheet, cat.headers));
    categories[cat.sheet] = rows;
    counts[cat.sheet] = rows.length;
  });
  var frozenAt = new Date().toISOString();
  writeDailyFreezeFile_(dateKey, { dateKey: dateKey, frozenAt: frozenAt, categories: categories });
  pruneDailyFreezeFiles_(dateKey);
  return jsonOut_({ ok: true, dateKey: dateKey, frozenAt: frozenAt, counts: counts });
}

// "Tally-based historical recovery" - added 26 Oct 2026. Daily Freeze
// never ran for 7/14/21 Oct 2026 (monthly_archive.py's own CONFIG had
// reverted to placeholder ORDERS_API_URL/REPORT_LOGIN_PASSWORD values -
// see claude/task-scheduler-setup-2026-10-03.md) so the "Report period"
// Week view's 3 October boundary dates have no Daily Freeze file, and can
// never be produced by the normal freezeDaySnapshot path above (that path
// always freezes TODAY's live sheet state under whatever dateKey is
// passed - it has no way to freeze a genuinely past day). This action
// instead accepts PRE-COMPUTED historical rows (built by
// backfill_week_boundaries.py on the office PC, from Tally's own retained
// voucher history) for exactly the 5 categories weekBoundaryData_() reads,
// and writes them straight to that date's Daily Freeze JSON file via
// writeDailyFreezeFile_ - WITHOUT ever calling getLiveSheet_ or touching
// the live Sheets, so today's real production data is never read or
// overwritten by this action.
//
// POST {action:"backfillDailyFreezeCategories", key:LIVE_SYNC_KEY,
//       dateKey:"yyyy-MM-dd", categories:{<sheetName>: [[...row...], ...]}}
// categories must be keyed by exactly the FREEZE_LIVE_CATEGORIES_ sheet
// names (LiveDealerRoster/LiveCollection/LiveCollectionByDealer/
// LiveSales/LiveSalesByDealer) - any other key is rejected. Each row's
// length is checked against that category's own header length so a
// malformed backfill payload fails loudly instead of silently writing a
// half-shaped row the dashboard would later misread. A category omitted
// from this call is simply left out of the written file (never filled in
// from the live sheet) - weekBoundaryData_'s own `cat[sheet] || []`
// fallback already treats a missing category as empty, same as any
// pre-Daily-Freeze date, so a partial backfill degrades safely instead of
// silently mixing in today's unrelated live data for whatever was left
// out.
function handleBackfillDailyFreezeCategories_(body) {
  if (body.key !== LIVE_SYNC_KEY) return jsonOut_({ ok: false, error: "Bad key" });
  var dateKey = String(body.dateKey || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) return jsonOut_({ ok: false, error: "Bad dateKey, expected yyyy-MM-dd" });
  var incoming = body.categories || {};
  var byName = {};
  FREEZE_LIVE_CATEGORIES_.forEach(function (cat) { byName[cat.sheet] = cat; });
  var categories = {};
  var counts = {};
  var providedNames = Object.keys(incoming);
  for (var i = 0; i < providedNames.length; i++) {
    var sheetName = providedNames[i];
    var cat = byName[sheetName];
    if (!cat) return jsonOut_({ ok: false, error: "Unknown category: " + sheetName });
    var rows = incoming[sheetName];
    if (!Array.isArray(rows)) return jsonOut_({ ok: false, error: "categories." + sheetName + " must be an array of rows" });
    for (var r = 0; r < rows.length; r++) {
      if (!Array.isArray(rows[r]) || rows[r].length !== cat.headers.length) {
        return jsonOut_({ ok: false, error: sheetName + " row " + r + " has " + (rows[r] && rows[r].length) + " column(s), expected " + cat.headers.length });
      }
    }
    categories[sheetName] = rows;
    counts[sheetName] = rows.length;
  }
  var frozenAt = new Date().toISOString();
  writeDailyFreezeFile_(dateKey, { dateKey: dateKey, frozenAt: frozenAt, categories: categories, backfilled: true });
  return jsonOut_({ ok: true, dateKey: dateKey, frozenAt: frozenAt, counts: counts });
}

// GET ?type=dailyFreeze&date=yyyy-MM-dd - rebuilds that single day's
// frozen payload (same shape as ?type=live) from its Drive JSON file,
// for the dashboard's Report date picker. {ok:true, found:false} (no
// payload fields) if that day was never frozen - before this feature
// existed, pruned past the 90-day retention window, or today (which has
// no end-of-day freeze yet, same as the current month under Monthly
// Freeze) - the dashboard's caller falls back to the live view / its
// existing "no data" handling in that case, same no-regression shape as
// handleMonthlyFreezeGet_() above.
function handleDailyFreezeGet_(dateKey) {
  if (!dateKey || !/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) return jsonOut_({ ok: false, error: "date required, expected yyyy-MM-dd" });
  var file = readDailyFreezeFile_(dateKey);
  if (!file) return jsonOut_({ ok: true, found: false });
  // refMonthKey is the MONTH `dateKey` falls in, not the day itself -
  // Group Summary/Petronas Summary/the trend sheets are always read live
  // inside buildLivePayload_() (see its own comment), filtered by month,
  // same as Monthly Freeze's handleMonthlyFreezeGet_() already does.
  var monthKey = dateKey.slice(0, 7);
  var payload = buildLivePayload_(frozenRowSourceFromFile_(file), monthKey);
  payload.ok = true;
  payload.found = true;
  payload.dateKey = dateKey;
  payload.frozenAt = file.frozenAt;
  return jsonOut_(payload);
}

// =======================================================================
// "Report period" - Week / Quarter / Year summaries (2 Oct 2026)
// =======================================================================
// Owner's correction after Daily Freeze shipped: a single-day view wasn't
// what was wanted at all - "i dont want daily view date , i want last
// month data... this month , 1st week, 2nd like that... qtr 1
// (jan,feb,mar 2026) , qtr2... this year , last year". Clarified via
// AskUserQuestion: Week = a FIXED 7-day block within one calendar month
// (1st-7th, 8th-14th, 15th-21st, 22nd-28th, 29th-end), never crossing a
// month boundary ("week should be frozen as month calendar"); Quarter/
// Year are calendar (Jan-Dec), not financial. Both Week and Quarter/Year
// show the SAME reduced view as the existing Last Month report: Combined
// Sales + Combined Collection (company-wide) + a By-DSR table (tap a name
// to expand to its dealers) - not the full every-card treatment.
//
// The COMBINED (company-wide) totals need no new endpoint at all: they
// come straight from the daily salesTrend/salesRsTrend/collectionTrend
// maps buildLivePayload_() already sends on every live/frozen fetch
// (added 20 Sep 2026 for the Trend chart) by summing over the period's
// date range client-side - exactly how renderLastMonthReport_() already
// computes Last Month's own combined totals (see dashboard_index.html's
// sumMapInRange_()). Those trend sheets keep one row per calendar day
// going back to when this system started (no 90-day window, unlike Daily
// Freeze), so this is both simpler AND more robust than anything built
// here.
//
// What DOES need new backend work is the BY-DSR/BY-DEALER breakdown,
// since Sales/SalesByDealer/Collection/CollectionByDealer are
// "Month Sold Ltr"/"Month Collected" running MONTH-TO-DATE totals, not
// daily deltas - there is no sheet anywhere that already holds "just
// this week" or "just this quarter" per DSR/dealer figures:
//
//  - WEEK (handleWeekSummaryGet_ below): a week's total = (MTD at the
//    week's last day) MINUS (MTD at the day right before the week
//    started) - except Week 1, where "the day before" is the previous
//    month (MTD resets to 0 there), so Week 1's total is just the plain
//    MTD-at-day-7 value, no subtraction. The two MTD snapshots come from
//    Daily Freeze's existing Drive JSON files (handleFreezeDaySnapshot_
//    above) - exactly the raw data that feature was built to capture,
//    now repurposed instead of powering its own (removed) single-day
//    picker. "Today" itself is never frozen yet, so a boundary that falls
//    on today reads the LIVE sheets directly instead (weekBoundaryData_()
//    below) - same reasoning Monthly Freeze already uses for the current
//    in-progress month.
//
//  - QUARTER/YEAR (handlePeriodSummaryGet_ below): unlike a week, each
//    MONTH's own frozen "Month Sold Ltr" etc. (from Monthly Freeze,
//    already built 1 Oct 2026) is already a complete, independent total
//    for that month - it resets at the next month's start, so a
//    quarter/year total is simply the SUM of each of its completed
//    months' already-frozen totals, no subtraction needed. The current
//    in-progress month (if the period includes it) is summed in from the
//    LIVE sheets instead of a frozen archive, same current-month handling
//    as everywhere else in this file.
//
// Both endpoints return {ok:true, found:false, reason:...} rather than a
// partial/guessed number whenever a boundary they need was never frozen
// (too long ago, before these features existed, or - for Daily Freeze -
// pruned past its 90-day window) - same "never show a number that might
// be wrong" discipline as Monthly/Daily Freeze's own found:false case.
// `partial:true` on a found:true response means the period hasn't fully
// finished yet (an in-progress week/quarter/year, clamped to today) - the
// figures are real, just not the FULL period's total yet.

// Generic "yyyy-MM" + N months arithmetic, the month-key twin of
// addDaysToDateKey_() above.
function addMonthsToMonthKey_(monthKey, delta) {
  var parts = monthKey.split("-");
  var y = Number(parts[0]);
  var m = Number(parts[1]) - 1 + delta;
  y += Math.floor(m / 12);
  m = ((m % 12) + 12) % 12;
  return y + "-" + (m + 1 < 10 ? "0" : "") + (m + 1);
}
// Inclusive list of "yyyy-MM" keys from startKey to endKey. Capped at 60
// (5 years) purely as a runaway guard, same spirit as the dashboard's own
// 60-month cap on the Report month dropdown - a real Quarter (3 months)
// or Year (12 months) request never comes close.
function monthKeyRange_(startKey, endKey) {
  var out = [];
  var cursor = startKey;
  var guard = 0;
  while (cursor <= endKey && guard < 60) {
    out.push(cursor);
    cursor = addMonthsToMonthKey_(cursor, 1);
    guard++;
  }
  return out;
}
function daysInMonth_(year, month1to12) {
  var leap = (year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0));
  var days = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  return days[month1to12 - 1];
}

// The four "Month Sold Ltr"/"Month Collected"-style rowsets Week/Quarter/
// Year both need, parsed into plain {name: {field: number}} maps - a
// SMALLER, purpose-built twin of buildLivePayload_()'s own row parsing
// (that function builds the full dashboard payload shape with every
// field for every card; these four parse only the running MTD total(s)
// this feature actually needs, from whichever rowset - live sheet,
// Monthly Freeze archive, or Daily Freeze JSON file - the caller already
// read). Column indices match LIVE_SALES_HEADERS/LIVE_SALES_BY_DEALER_HEADERS/
// LIVE_COLLECTION_HEADERS/LIVE_COLLECTION_BY_DEALER_HEADERS exactly.
function parseSalesRows_(rows) {
  var out = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var name = String(r[0] || "").trim();
    if (!name) continue;
    out[name] = { salesLtr: Number(r[2]) || 0, salesRs: Number(r[6]) || 0 };
  }
  return out;
}
function parseSalesByDealerRows_(rows) {
  var out = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var name = String(r[0] || "").trim();
    if (!name) continue;
    out[name] = { salesLtr: Number(r[2]) || 0, salesRs: Number(r[9]) || 0 };
  }
  return out;
}
function parseCollectionRows_(rows) {
  var out = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var name = String(r[0] || "").trim();
    if (!name) continue;
    out[name] = { collectionRs: Number(r[2]) || 0 };
  }
  return out;
}
function parseCollectionByDealerRows_(rows) {
  var out = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var name = String(r[0] || "").trim();
    if (!name) continue;
    out[name] = { collectionRs: Number(r[2]) || 0 };
  }
  return out;
}
// {dsrName: [dealerName, ...]} - same shape as buildLivePayload_()'s own
// dealerRoster, parsed from whichever rowset (live/frozen-month/
// frozen-day) the caller already read.
function parseRosterRows_(rows) {
  var out = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var dsr = String(r[1] || "").trim();
    var dealer = String(r[2] || "").trim();
    if (!dsr || !dealer) continue;
    if (!out[dsr]) out[dsr] = [];
    out[dsr].push(dealer);
  }
  return out;
}

// Sums the SAME field(s) across several parsed maps (one per completed
// month, for Quarter/Year) - union of every name seen in any map, missing
// entries treated as 0.
function sumParsedMaps_(mapsArray, fields) {
  var out = {};
  mapsArray.forEach(function (m) {
    Object.keys(m || {}).forEach(function (name) {
      if (!out[name]) {
        out[name] = {};
        fields.forEach(function (f) { out[name][f] = 0; });
      }
      fields.forEach(function (f) { out[name][f] += (m[name][f] || 0); });
    });
  });
  return out;
}
// endMap minus startMap (for Week) - every name in endMap, startMap
// entries missing/absent treated as 0 (so a dealer that only appears from
// partway through the month still gets its full MTD-so-far credited to
// the week it actually happened in, rather than being dropped). A
// negative result (a Tally correction/return inside the subtracted
// window) is left as-is rather than clamped to 0 - the existing cell
// renderers already show any value <=0 as "-", so nothing further is
// needed here.
function diffParsedMaps_(endMap, startMap, fields) {
  var out = {};
  Object.keys(endMap || {}).forEach(function (name) {
    out[name] = {};
    fields.forEach(function (f) {
      var endV = (endMap[name] && endMap[name][f]) || 0;
      var startV = (startMap && startMap[name] && startMap[name][f]) || 0;
      out[name][f] = endV - startV;
    });
  });
  return out;
}
function mergeRosterMaps_(mapsArray) {
  var out = {};
  mapsArray.forEach(function (m) {
    Object.keys(m || {}).forEach(function (dsr) {
      if (!out[dsr]) out[dsr] = [];
      (m[dsr] || []).forEach(function (dealerName) {
        if (out[dsr].indexOf(dealerName) === -1) out[dsr].push(dealerName);
      });
    });
  });
  return out;
}

// Reads one calendar day's Sales/SalesByDealer/Collection/
// CollectionByDealer/DealerRoster MTD snapshot, already parsed into the
// maps above - `dateKey === todayKey` reads the LIVE sheets (today has no
// end-of-day freeze yet); any other date reads Daily Freeze's Drive JSON
// file for it, returning found:false if that day was never frozen (too
// long ago, before Daily Freeze existed, or pruned past its 90-day
// window).
function weekBoundaryData_(dateKey, todayKey) {
  if (dateKey === todayKey) {
    return {
      found: true, live: true, frozenAt: null,
      sales: parseSalesRows_(liveSheetRows_(getLiveSheet_(LIVE_SALES_SHEET, LIVE_SALES_HEADERS))),
      salesByDealer: parseSalesByDealerRows_(liveSheetRows_(getLiveSheet_(LIVE_SALES_BY_DEALER_SHEET, LIVE_SALES_BY_DEALER_HEADERS))),
      collection: parseCollectionRows_(liveSheetRows_(getLiveSheet_(LIVE_COLLECTION_SHEET, LIVE_COLLECTION_HEADERS))),
      collectionByDealer: parseCollectionByDealerRows_(liveSheetRows_(getLiveSheet_(LIVE_COLLECTION_BY_DEALER_SHEET, LIVE_COLLECTION_BY_DEALER_HEADERS))),
      roster: parseRosterRows_(liveSheetRows_(getLiveSheet_(LIVE_DEALER_ROSTER_SHEET, LIVE_DEALER_ROSTER_HEADERS)))
    };
  }
  var file = readDailyFreezeFile_(dateKey);
  if (!file) return { found: false };
  var cat = file.categories || {};
  return {
    found: true, live: false, frozenAt: file.frozenAt,
    sales: parseSalesRows_(cat[LIVE_SALES_SHEET] || []),
    salesByDealer: parseSalesByDealerRows_(cat[LIVE_SALES_BY_DEALER_SHEET] || []),
    collection: parseCollectionRows_(cat[LIVE_COLLECTION_SHEET] || []),
    collectionByDealer: parseCollectionByDealerRows_(cat[LIVE_COLLECTION_BY_DEALER_SHEET] || []),
    roster: parseRosterRows_(cat[LIVE_DEALER_ROSTER_SHEET] || [])
  };
}

// GET ?type=weekSummary&month=yyyy-MM&week=1..5 - a fixed 7-day block
// within ONE calendar month (never crossing into the next), by-DSR/
// by-dealer Sales+Collection totals only (the company-wide combined
// totals are computed client-side from the trend maps - see the long
// comment above). {ok:true, found:false, reason:...} when:
//  - "no_such_week": that week number doesn't exist in this month (e.g.
//    Week 5 of a 28-day February).
//  - "future": the whole week hasn't started yet.
//  - "end_not_frozen" / "start_not_frozen": a boundary day this
//    computation needs was never frozen (`dateKey` names which one) -
//    this is the single-day gap case Daily Freeze's 90-day retention (or
//    simply not having run yet before 2 Oct 2026) can produce.
function handleWeekSummaryGet_(monthKey, weekNumRaw) {
  if (!/^\d{4}-\d{2}$/.test(monthKey || "")) return jsonOut_({ ok: false, error: "month required, expected yyyy-MM" });
  var weekNum = Number(weekNumRaw);
  if (!(weekNum >= 1 && weekNum <= 5)) return jsonOut_({ ok: false, error: "week required, expected 1-5" });
  var parts = monthKey.split("-");
  var year = Number(parts[0]), month = Number(parts[1]);
  var lastDay = daysInMonth_(year, month);
  var startDay = (weekNum - 1) * 7 + 1;
  if (startDay > lastDay) return jsonOut_({ ok: true, found: false, reason: "no_such_week" });
  var endDay = Math.min(startDay + 6, lastDay);
  var pad2 = function (n) { return (n < 10 ? "0" : "") + n; };
  var weekStartKey = monthKey + "-" + pad2(startDay);
  var weekEndKey = monthKey + "-" + pad2(endDay);
  var todayKey = Utilities.formatDate(new Date(), DISPLAY_TIMEZONE, "yyyy-MM-dd");
  if (weekStartKey > todayKey) return jsonOut_({ ok: true, found: false, reason: "future" });
  var effectiveEndKey = weekEndKey > todayKey ? todayKey : weekEndKey;
  var endData = weekBoundaryData_(effectiveEndKey, todayKey);
  if (!endData.found) return jsonOut_({ ok: true, found: false, reason: "end_not_frozen", dateKey: effectiveEndKey });
  // partial:true means the END boundary came from today's still-live,
  // not-yet-frozen data - covers both "the week's own last day hasn't
  // happened yet" (clamped to today) AND "today IS the week's last day,
  // but its end-of-day freeze hasn't run" - either way the figures are
  // real, just not the full week's final total yet.
  var partial = endData.live;
  var startData = null;
  if (weekNum > 1) {
    var startMinus1Key = addDaysToDateKey_(weekStartKey, -1);
    startData = weekBoundaryData_(startMinus1Key, todayKey);
    if (!startData.found) return jsonOut_({ ok: true, found: false, reason: "start_not_frozen", dateKey: startMinus1Key });
  }
  var salesFields = ["salesLtr", "salesRs"];
  var collFields = ["collectionRs"];
  return jsonOut_({
    ok: true, found: true,
    monthKey: monthKey, week: weekNum,
    weekStartKey: weekStartKey, weekEndKey: weekEndKey,
    partial: partial,
    frozenAt: endData.frozenAt,
    sales: diffParsedMaps_(endData.sales, startData ? startData.sales : null, salesFields),
    salesByDealer: diffParsedMaps_(endData.salesByDealer, startData ? startData.salesByDealer : null, salesFields),
    collection: diffParsedMaps_(endData.collection, startData ? startData.collection : null, collFields),
    collectionByDealer: diffParsedMaps_(endData.collectionByDealer, startData ? startData.collectionByDealer : null, collFields),
    dealerRoster: endData.roster
  });
}

// GET ?type=periodSummary&start=yyyy-MM&end=yyyy-MM - an inclusive
// calendar-month range (a Quarter is 3 months, a Year is 12 - Jan-Dec,
// confirmed NOT financial-year by the owner), by-DSR/by-dealer Sales+
// Collection totals only (company-wide combined totals again come from
// the client-side trend-map sum). Each completed month in the range is
// summed from its own already-frozen Monthly Freeze total (no
// subtraction needed - unlike a week, a month's own MTD total is already
// a complete, independent figure); the current in-progress month (if the
// range reaches it) is summed in from the LIVE sheets instead.
// {ok:true, found:false, reason:...} when:
//  - "future": the whole range hasn't started yet.
//  - "month_not_frozen": some completed month in the range (`missingMonth`)
//    was never frozen - before Monthly Freeze existed, or its automatic
//    month-end freeze simply hasn't run yet for some other reason. Partial
//    figures are never shown in this case; the owner's own "never fake a
//    number" rule from Monthly/Daily Freeze applies here too.
function handlePeriodSummaryGet_(startMonth, endMonth) {
  if (!/^\d{4}-\d{2}$/.test(startMonth || "") || !/^\d{4}-\d{2}$/.test(endMonth || "")) {
    return jsonOut_({ ok: false, error: "start and end required, expected yyyy-MM" });
  }
  if (endMonth < startMonth) return jsonOut_({ ok: false, error: "end before start" });
  var currentMonthKey = Utilities.formatDate(new Date(), DISPLAY_TIMEZONE, "yyyy-MM");
  if (startMonth > currentMonthKey) return jsonOut_({ ok: true, found: false, reason: "future" });
  var effectiveEnd = endMonth > currentMonthKey ? currentMonthKey : endMonth;
  // partial:true whenever the still-in-progress current month falls
  // inside the requested range - not just when the range reaches PAST it
  // (startMonth <= currentMonthKey is already guaranteed by the "future"
  // check above, so this is really "currentMonthKey <= endMonth"). The
  // current month's own figures are always live/incomplete until it
  // ends, even when it's exactly the range's last month and so needs no
  // clamping at all.
  var partial = (currentMonthKey <= endMonth);
  var months = monthKeyRange_(startMonth, effectiveEnd);
  var salesMaps = [], salesByDealerMaps = [], collectionMaps = [], collectionByDealerMaps = [], rosterMaps = [];
  var missingMonth = null;
  for (var i = 0; i < months.length; i++) {
    var mk = months[i];
    if (mk === currentMonthKey) {
      salesMaps.push(parseSalesRows_(liveSheetRows_(getLiveSheet_(LIVE_SALES_SHEET, LIVE_SALES_HEADERS))));
      salesByDealerMaps.push(parseSalesByDealerRows_(liveSheetRows_(getLiveSheet_(LIVE_SALES_BY_DEALER_SHEET, LIVE_SALES_BY_DEALER_HEADERS))));
      collectionMaps.push(parseCollectionRows_(liveSheetRows_(getLiveSheet_(LIVE_COLLECTION_SHEET, LIVE_COLLECTION_HEADERS))));
      collectionByDealerMaps.push(parseCollectionByDealerRows_(liveSheetRows_(getLiveSheet_(LIVE_COLLECTION_BY_DEALER_SHEET, LIVE_COLLECTION_BY_DEALER_HEADERS))));
      rosterMaps.push(parseRosterRows_(liveSheetRows_(getLiveSheet_(LIVE_DEALER_ROSTER_SHEET, LIVE_DEALER_ROSTER_HEADERS))));
      continue;
    }
    var frozenAt = monthlyFreezeFrozenAt_(mk);
    if (!frozenAt) { missingMonth = mk; break; }
    salesMaps.push(parseSalesRows_(readFrozenSheetForMonth_(LIVE_SALES_SHEET, LIVE_SALES_HEADERS, mk)));
    salesByDealerMaps.push(parseSalesByDealerRows_(readFrozenSheetForMonth_(LIVE_SALES_BY_DEALER_SHEET, LIVE_SALES_BY_DEALER_HEADERS, mk)));
    collectionMaps.push(parseCollectionRows_(readFrozenSheetForMonth_(LIVE_COLLECTION_SHEET, LIVE_COLLECTION_HEADERS, mk)));
    collectionByDealerMaps.push(parseCollectionByDealerRows_(readFrozenSheetForMonth_(LIVE_COLLECTION_BY_DEALER_SHEET, LIVE_COLLECTION_BY_DEALER_HEADERS, mk)));
    rosterMaps.push(parseRosterRows_(readFrozenSheetForMonth_(LIVE_DEALER_ROSTER_SHEET, LIVE_DEALER_ROSTER_HEADERS, mk)));
  }
  if (missingMonth) return jsonOut_({ ok: true, found: false, reason: "month_not_frozen", missingMonth: missingMonth });
  var salesFields = ["salesLtr", "salesRs"];
  var collFields = ["collectionRs"];
  return jsonOut_({
    ok: true, found: true,
    startMonth: startMonth, endMonth: endMonth, partial: partial,
    sales: sumParsedMaps_(salesMaps, salesFields),
    salesByDealer: sumParsedMaps_(salesByDealerMaps, salesFields),
    collection: sumParsedMaps_(collectionMaps, collFields),
    collectionByDealer: sumParsedMaps_(collectionByDealerMaps, collFields),
    dealerRoster: mergeRosterMaps_(rosterMaps)
  });
}

// GET ?type=todaySummary - by-DSR/by-dealer Sales+Collection totals for
// JUST today (company-wide combined totals again come from the
// client-side trend-map sum - see the "Report period" comment in
// dashboard_index.html). Added 3 Oct 2026 at the owner's request for a
// one-tap "today only" view, alongside the existing Month/Week/Quarter/
// Year options ("add only 1 option at last as today - to see only
// todays transaction"). This is really a week-of-one-day: today's own
// live MTD-so-far figures minus YESTERDAY's frozen end-of-day MTD
// figures leaves exactly today's own transactions, using the exact same
// weekBoundaryData_()/diffParsedMaps_() primitives handleWeekSummaryGet_()
// above already uses for a week's own last-day boundary - no new
// machinery needed.
// {ok:true, found:false, reason:"yesterday_not_frozen", dateKey:...} if
// yesterday was never frozen - Daily Freeze hasn't run yet, or this is
// the very first day the feature exists. Self-resolving: every day's
// 7:45 PM Daily Freeze run (monthly_archive.py) covers the NEXT day's
// Today view automatically, so this can only happen once.
function handleTodaySummaryGet_() {
  var todayKey = Utilities.formatDate(new Date(), DISPLAY_TIMEZONE, "yyyy-MM-dd");
  var yesterdayKey = addDaysToDateKey_(todayKey, -1);
  var endData = weekBoundaryData_(todayKey, todayKey);
  var startData = weekBoundaryData_(yesterdayKey, todayKey);
  if (!startData.found) return jsonOut_({ ok: true, found: false, reason: "yesterday_not_frozen", dateKey: yesterdayKey });
  var salesFields = ["salesLtr", "salesRs"];
  var collFields = ["collectionRs"];
  return jsonOut_({
    ok: true, found: true,
    dateKey: todayKey,
    sales: diffParsedMaps_(endData.sales, startData.sales, salesFields),
    salesByDealer: diffParsedMaps_(endData.salesByDealer, startData.salesByDealer, salesFields),
    collection: diffParsedMaps_(endData.collection, startData.collection, collFields),
    collectionByDealer: diffParsedMaps_(endData.collectionByDealer, startData.collectionByDealer, collFields),
    dealerRoster: endData.roster
  });
}

// Refactored 1 Oct 2026 out of the old handleLiveGet_() into
// buildLivePayload_(rowSource, refMonthKey) - see the long "Monthly
// Freeze" comment above for why. `rowSource(sheetName, headers)`
// abstracts "where do this category's rows come from": liveRowSource_()
// (the live Sheets, same as always) for a normal ?type=live call, or
// frozenRowSource_(monthKey) (the Freeze_* archive sheets) when
// rebuilding a past, frozen month's payload for ?type=monthlyFreeze.
// Petronas Summary/Group Summary/Group Summary By Dealer/the trend
// sheets are deliberately NOT read through rowSource - they already
// keep their own full month-by-month history and are never frozen (see
// FREEZE_LIVE_CATEGORIES_'s own comment), so they always read the live
// sheet directly regardless of which month this payload is for.
// `refMonthKey` is "what month is this payload as of", used only by the
// staleness checks below (isStaleLiveRow_) - the REAL current month for
// a live call, or the frozen month itself for a historical one, so a
// frozen Sep row's own Sep "Updated At" is never wrongly treated as
// stale once the calendar has moved on to October.
function buildLivePayload_(rowSource, refMonthKey) {
  var stockRows = rowSource(LIVE_STOCK_SHEET, LIVE_STOCK_HEADERS);
  var outstandingRows = rowSource(LIVE_OUTSTANDING_SHEET, LIVE_OUTSTANDING_HEADERS);
  var creditLimitRows = rowSource(LIVE_CREDITLIMIT_SHEET, LIVE_CREDITLIMIT_HEADERS);
  var collectionRows = rowSource(LIVE_COLLECTION_SHEET, LIVE_COLLECTION_HEADERS);
  var collectionByDealerRows = rowSource(LIVE_COLLECTION_BY_DEALER_SHEET, LIVE_COLLECTION_BY_DEALER_HEADERS);
  var petronasSummaryRows = liveSheetRows_(getLiveSheet_(LIVE_PETRONAS_SUMMARY_SHEET, LIVE_PETRONAS_SUMMARY_HEADERS));
  var groupSummaryRows = liveSheetRows_(getLiveSheet_(LIVE_GROUP_SUMMARY_SHEET, LIVE_GROUP_SUMMARY_HEADERS));
  var groupSummaryByDealerRows = liveSheetRows_(getLiveSheet_(LIVE_GROUP_SUMMARY_BY_DEALER_SHEET, LIVE_GROUP_SUMMARY_BY_DEALER_HEADERS));
  var salesRows = rowSource(LIVE_SALES_SHEET, LIVE_SALES_HEADERS);
  var salesByDealerRows = rowSource(LIVE_SALES_BY_DEALER_SHEET, LIVE_SALES_BY_DEALER_HEADERS);
  var salesByItemRows = rowSource(LIVE_SALES_BY_ITEM_SHEET, LIVE_SALES_BY_ITEM_HEADERS);
  var salesByDsrItemRows = rowSource(LIVE_SALES_BY_DSR_ITEM_SHEET, LIVE_SALES_BY_DSR_ITEM_HEADERS);
  var salesTrendRows = liveSheetRows_(getLiveSheet_(LIVE_SALES_TREND_SHEET, LIVE_SALES_TREND_HEADERS));
  var salesRsTrendRows = liveSheetRows_(getLiveSheet_(LIVE_SALES_RS_TREND_SHEET, LIVE_SALES_RS_TREND_HEADERS));
  var collectionTrendRows = liveSheetRows_(getLiveSheet_(LIVE_COLLECTION_TREND_SHEET, LIVE_COLLECTION_TREND_HEADERS));
  // Added 11 Oct 2026 for the per-DSR Sales/Collection heatmap - same
  // "always live, never frozen" treatment as the 3 trend sheets just
  // above (liveSheetRows_, not rowSource - these are rolling this-month+
  // last-month windows, not a monthly snapshot).
  var salesTrendByDsrRows = liveSheetRows_(getLiveSheet_(LIVE_SALES_TREND_BY_DSR_SHEET, LIVE_SALES_TREND_BY_DSR_HEADERS));
  var collectionTrendByDsrRows = liveSheetRows_(getLiveSheet_(LIVE_COLLECTION_TREND_BY_DSR_SHEET, LIVE_COLLECTION_TREND_BY_DSR_HEADERS));

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

  // Added 30 Sep 2026 - {dealerName: {creditLimit, creditPeriod}}, see
  // LIVE_CREDITLIMIT_HEADERS/handleLiveCreditLimitUpdate_() above. Same
  // whole-company scope as `outstanding` just above - not filtered to
  // known DSR dealers, so any ledger with a real CREDITLIMIT resolves
  // here regardless of which group it belongs to.
  var creditLimit = {};
  for (var cl = 0; cl < creditLimitRows.length; cl++) {
    var r5 = creditLimitRows[cl];
    var clName = String(r5[0] || "").trim();
    if (!clName) continue;
    creditLimit[clName] = {
      creditLimit: r5[1] === "" || r5[1] === undefined ? null : Number(r5[1]),
      creditPeriod: r5[2] === "" || r5[2] === undefined ? null : Number(r5[2])
    };
  }

  // Added 8 Sep 2026 - grouped {dsrName: [dealerName, ...]} for the
  // new-dealer auto-add feature. Every DSR app only ever looks up its own
  // dsrName's entry (see applyLiveOverlay_() in each app file), so sending
  // the full roster for all 8 DSRs to every app is simplest and matches how
  // stock/outstanding already work (whole snapshot, filtered client-side).
  var rosterRows = rowSource(LIVE_DEALER_ROSTER_SHEET, LIVE_DEALER_ROSTER_HEADERS);
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
  var newDealerRows = rowSource(LIVE_NEW_DEALERS_SHEET, LIVE_NEW_DEALERS_HEADERS);
  var newDealers = {};
  for (var nd = 0; nd < newDealerRows.length; nd++) {
    var r5 = newDealerRows[nd];
    var ndDsr = String(r5[1] || "").trim();
    var ndName = String(r5[2] || "").trim();
    if (!ndDsr || !ndName) continue;
    if (String(r5[5] || "").trim() === "Y") continue; // dismissed by hand - see "Not new?" / dismissNewDealer
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

  // Added 1 Oct 2026 - a defensive, backend-side safety net on top of the
  // same-day diff_sales_by_item()/diff_sales_by_dsr_item() watcher fix (see
  // tally_live_watcher.py). That fix only self-corrects a stale item when
  // the WATCHER's own local tally_live_state.json already remembers the
  // item from a previous poll - if that local cache was ever empty/reset
  // (office PC watcher restarted, state file moved/deleted, etc.) while an
  // item simply hasn't sold again since the month rolled over, the watcher
  // never revisits that item at all and the stale row just sits in
  // LiveSalesByItem/LiveSalesByDsrItem forever with its real last month's
  // number still in it. Confirmed happening for real on 1 Oct 2026 - the
  // owner found "20W40 55 LIT SPRINTA F300" still showing 2,035 Ltr/660 Ltr
  // in both sheets with "Updated At" timestamps from 29/30 Sep, despite the
  // watcher fix already being deployed and the sheet columns themselves
  // being correct. Rather than depend on the watcher's local state always
  // staying in sync with the sheet (which this incident shows it cannot be
  // relied on to do), this endpoint now ALSO refuses to serve any
  // salesByItem/salesByDsrItem row whose own "Updated At" isn't from the
  // current calendar month (Asia/Kolkata, matching DISPLAY_TIMEZONE/every
  // other "today"/"this month" calculation in this file) - a stale row is
  // treated exactly like the row doesn't exist, same as a genuine zero
  // already was. This is a read-time filter only; it never touches the
  // Sheet itself, so the watcher's own diff-and-correct mechanism (and
  // whatever becomes of that still-open investigation) keeps working
  // underneath it unchanged.
  function isStaleLiveRow_(updatedAtVal, currentMonthKey) {
    if (!updatedAtVal) return true;
    var d = (updatedAtVal instanceof Date) ? updatedAtVal : new Date(updatedAtVal);
    if (isNaN(d.getTime())) return true;
    return Utilities.formatDate(d, DISPLAY_TIMEZONE, "yyyy-MM") !== currentMonthKey;
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
    var nowKey = refMonthKey;
    for (var s = 0; s < rows.length; s++) {
      var r = rows[s];
      var key = String(r[0] || "").trim();
      if (!key) continue;
      // 1 Oct 2026 - see isStaleLiveRow_() above. Column index 3 is
      // "Updated At" per LIVE_SALES_BY_ITEM_HEADERS.
      if (isStaleLiveRow_(r[3], nowKey)) continue;
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

  // Added 5 Oct 2026 - cross-check "New dealers" against each dealer's own
  // Tally sales history before trusting it, instead of passing through
  // whatever the Tally watcher's LiveNewDealers sheet says unconditionally.
  //
  // Why this is needed: `newDealers` (built above from LIVE_NEW_DEALERS_SHEET)
  // is produced entirely by tally_live_watcher.py's once-a-day dealer-roster
  // diff, comparing today's full eligible-dealer list against a baseline it
  // keeps in a LOCAL file on the office PC (known_roster_names in
  // state.json) - Code.gs has never validated this against anything else,
  // it just relays whatever that one local file's diff decided. That local
  // baseline can legitimately go stale or get reset independent of whether
  // a dealer is actually new (the office PC's watcher not running on a
  // given evening, being restarted/moved to a new working directory - e.g.
  // for the Task Scheduler auto-start setup done 3 Oct 2026 - or any other
  // gap). When that happens, a long-standing dealer the watcher simply
  // didn't have in its local baseline gets flagged "new" exactly like a
  // genuinely brand-new one, with no way to tell the two apart downstream.
  //
  // Real-world trigger confirmed 5 Oct 2026: the owner reported 9 dealers
  // showing as "new" (all dated 3 Oct 2026, the same day the Tally watcher's
  // Task Scheduler entry was set up) despite no real new activity - a
  // single-day batch of "new" flags like that is the signature of a reset/
  // stale local baseline, not 9 coincidentally new customers landing on one
  // day.
  //
  // The fix: `salesByDealer` (built just above) already carries each
  // dealer's own `lastSaleBeforeMonth` - their last REAL Tally sale date
  // strictly before the current calendar month, added 23 Sep 2026 for
  // "Dealers Reactivated This Month". If a dealer flagged "new" already has
  // a `lastSaleBeforeMonth` on file, that's direct proof from a completely
  // different data path (LiveSalesByDealer, driven by fetch_sales()'s own
  // monthly voucher scan, not the roster-diff/local-baseline mechanism at
  // all) that they were already buying before this month - so they cannot
  // be genuinely new, whatever the roster diff concluded, and are silently
  // dropped here before being sent to any app. A dealer with no
  // `lastSaleBeforeMonth` on file (the normal case for an actually-new
  // dealer) is completely unaffected - this only ever removes a false
  // positive, never a real one.
  for (var ndDsrKey in newDealers) {
    newDealers[ndDsrKey] = newDealers[ndDsrKey].filter(function (nd) {
      var priorSale = salesByDealer[nd.name] && salesByDealer[nd.name].lastSaleBeforeMonth;
      return !priorSale;
    });
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

  // Added 11 Oct 2026 for the per-DSR Sales/Collection heatmap - sent as a
  // flat array of {dsrName, date, ltr|rs}, same convention as
  // salesByDsrItem just below (the dashboard already groups/sorts flat
  // per-DSR data client-side everywhere else on this page).
  //
  // De-duped defensively 15 Oct 2026 - see the long comment on
  // handleLiveSalesTrendByDsrUpdate_/handleLiveCollectionTrendByDsrUpdate_
  // above for the bug that let duplicate (dsrName, date) rows pile up in
  // the sheet between 11-15 Oct 2026. That write-side bug is fixed now, but
  // this read-side de-dupe is kept permanently (not just as a one-time
  // patch) as a safety net: it's cheap, and it means a future bug of the
  // same shape can never again inflate a number shown on the dashboard, it
  // would only ever show stale-but-correct-looking data for one row. Keeps
  // whichever duplicate row has the latest "Updated At" (column index 3) -
  // that's always the most recently known-correct snapshot for that
  // dsrName+date, since every pushed row holds that day's running total AS
  // OF the push, never an increment to add to the others.
  function readTrendByDsrRows_(rows, valueCol, valueKey) {
    var bestByKey = {};
    for (var t = 0; t < rows.length; t++) {
      var r = rows[t];
      var dsrName = String(r[0] || "").trim();
      var d = normalizeDateKey_(r[1]);
      if (!dsrName || !d) continue;
      var key = dsrName + "||" + d;
      var updatedAt = String(r[3] || "");
      var prev = bestByKey[key];
      if (!prev || updatedAt > prev._updatedAt) {
        bestByKey[key] = { dsrName: dsrName, date: d, _updatedAt: updatedAt, _raw: r };
      }
    }
    var out = [];
    Object.keys(bestByKey).forEach(function (key) {
      var picked = bestByKey[key];
      var rec = { dsrName: picked.dsrName, date: picked.date };
      rec[valueKey] = picked._raw[valueCol] === "" ? null : Number(picked._raw[valueCol]);
      out.push(rec);
    });
    return out;
  }
  var salesTrendByDsr = readTrendByDsrRows_(salesTrendByDsrRows, 2, "ltr");
  var collectionTrendByDsr = readTrendByDsrRows_(collectionTrendByDsrRows, 2, "rs");

  // Added 22 Sep 2026 for the dashboard's "Top selling SKU by DSR" panels
  // and "part number covered per DSR" stat - see LIVE_SALES_BY_DSR_ITEM_SHEET
  // above and tally_live_watcher.py's dsr_item_totals for the full story.
  // Sent as a flat array (not a nested {dsrName: {itemName: ltr}} object) -
  // the dashboard already groups/sorts flat per-dealer/per-item data
  // client-side everywhere else on this page, same pattern here.
  var salesByDsrItem = [];
  var salesByDsrItemNowKey_ = refMonthKey;
  for (var di = 0; di < salesByDsrItemRows.length; di++) {
    var dir = salesByDsrItemRows[di];
    var diDsr = String(dir[0] || "").trim();
    var diItem = String(dir[1] || "").trim();
    if (!diDsr || !diItem) continue;
    var diMonth = dir[2] === "" || dir[2] === undefined ? null : Number(dir[2]);
    if (diMonth === null || diMonth <= 0) continue;
    // 1 Oct 2026 - see isStaleLiveRow_() above. Column index 3 is
    // "Updated At" per LIVE_SALES_BY_DSR_ITEM_HEADERS. Without this, a row
    // that stopped getting fresh pushes (watcher restart with an empty
    // local state cache, same root cause as the salesByItem fix above)
    // keeps showing its real LAST month's Ltr forever under "Top selling
    // SKU - by DSR".
    if (isStaleLiveRow_(dir[3], salesByDsrItemNowKey_)) continue;
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

  // Petronas Business Summary - LIVE, one row PER MONTH - added 1 Oct 2026
  // per the owner's explicit choice ("FROM SEP ONWARD LIVE DATA SHOULD
  // COME"). Jan-Aug 2026 are a static, owner-confirmed snapshot baked
  // directly into the dashboard (PETRONAS_SUMMARY_DATA_); this sheet only
  // ever holds Sep 2026 onward, one upserted row per month key, so it's
  // returned here in full (no current-month-only filter like
  // salesByItem/salesByDsrItem above - a closed month's row is still real
  // history the dashboard should keep showing, not stale data to drop).
  // Sorted ascending by Month Key so the dashboard can just append these
  // after August without re-sorting.
  var petronasSummaryLive = [];
  for (var psi = 0; psi < petronasSummaryRows.length; psi++) {
    var psr = petronasSummaryRows[psi];
    var psMonth = String(psr[0] || "").trim();
    if (!psMonth) continue;
    petronasSummaryLive.push({
      month: psMonth,
      // Derived from the Month Key (psMonth, e.g. "202609"), never from the
      // sheet's own "Month Label" cell (psr[1]) - found 1 Oct 2026 that
      // Google Sheets auto-detects a plain "Sep 2026"/"Oct 2026" text value
      // as a real calendar date and silently converts that cell to a date
      // serial. SpreadsheetApp then reads it back as a JS Date, and
      // JSON.stringify() in jsonOut_() serializes that Date to an ISO-8601
      // UTC string (e.g. "2026-08-31T18:30:00.000Z" for "Sep 2026" in IST) -
      // exactly the raw timestamp the owner saw on the dashboard instead of
      // a month name. monthLabelFromKey_() below sidesteps the corrupted
      // cell entirely by computing the label from the Month Key column,
      // which is a plain digit string Sheets never date-converts (confirmed
      // by the live rows always sorting into the correct chronological
      // position).
      monthLabel: monthLabelFromKey_(psMonth),
      purchaseLtr: psr[2] === "" || psr[2] === undefined ? null : Number(psr[2]),
      salesLtr: psr[3] === "" || psr[3] === undefined ? null : Number(psr[3]),
      // x1.18 (GST) REMOVED here 15 Oct 2026 - this was a double-GST bug,
      // found while fixing the Petronas-claim exclusion the owner asked
      // for on this same Sales Value field. History: 6 Oct 2026 this line
      // started multiplying by 1.18 here, on the stated assumption that
      // "tally_live_watcher.py's own pushed figure (psr[4]) stays the
      // plain pre-GST taxable value". That assumption stopped being true
      // the DAY BEFORE (5 Oct 2026), when fetch_petronas_summary() in
      // tally_live_watcher.py was itself changed to push an
      // ALREADY-GST-inclusive salesValue (the exact same
      // "round(abs(amt) * 1.18, 2)" conversion, applied once there). The
      // two changes landed a day apart and nobody noticed the earlier one
      // had already made the later one redundant - net effect, every live
      // Sep/Oct 2026 Sales (Value) figure on the dashboard has been
      // showing 1.18 x 1.18 = 1.3924x the real pre-GST amount (about 18%
      // too high on top of GST, not just GST once) since 6 Oct 2026. Fixed
      // by removing the multiplication here and trusting psr[4] (already
      // GST-inclusive from Python) as-is - matches the static Jan-Aug
      // array's own convention (GST applied once, in the number itself,
      // nowhere else).
      salesValue: psr[4] === "" || psr[4] === undefined ? null : Number(psr[4]),
      collected: psr[5] === "" || psr[5] === undefined ? null : Number(psr[5]),
      purchaseValue: psr[6] === "" || psr[6] === undefined ? null : Number(psr[6]),
      payment: psr[7] === "" || psr[7] === undefined ? null : Number(psr[7]),
      updatedAt: psr[8] || "",
      // Added 15 Oct 2026 - see handleLivePetronasSummaryUpdate_()'s own
      // comment on claimsCn. null for any row pushed before this column
      // existed (reads back as "" from a sheet cell that was never
      // written) - shown as a dash on the dashboard, not a fabricated 0.
      claimsCn: psr[9] === "" || psr[9] === undefined ? null : Number(psr[9])
    });
  }
  petronasSummaryLive.sort(function (a, b) { return String(a.month).localeCompare(String(b.month)); });

  // Added 26 Oct 2026 - see dealerVisitCheckLive_()'s own long comment
  // (near dealerMetaMap_() above) for the full design/calibration story.
  // Always computed live (two Sheets reads, not Tally data), never
  // monthly-frozen - "today's" check has no meaning for a past month.
  var dealerVisitCheck = dealerVisitCheckLive_();

  return {
    stock: stock,
    outstanding: outstanding,
    creditLimit: creditLimit,
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
    salesTrendByDsr: salesTrendByDsr,
    collectionTrendByDsr: collectionTrendByDsr,
    sixtyPlusByDealer: sixtyPlusByDealer,
    groupSummary: groupSummary,
    groupSummaryByDealer: groupSummaryByDealer,
    petronasSummary: petronasSummaryLive,
    dealerVisitCheck: dealerVisitCheck
  };
}

// Thin wrapper - the real ?type=live handler, behavior unchanged from the
// caller's point of view. Added 1 Oct 2026 alongside buildLivePayload_()
// above (see its own comment, and the "Monthly Freeze" block further up,
// for the refactor rationale).
function liveRowSource_() {
  return function (sheetName, headers) { return liveSheetRows_(getLiveSheet_(sheetName, headers)); };
}

// --- Live Cache (added 5 Oct 2026) ---------------------------------------
// Added because the owner reported login/dashboard load feeling "too
// slow" and sent a screen recording confirming ~20-25 seconds of dead
// "Refreshing..." time before real figures appeared. Root cause:
// buildLivePayload_() above reads ~15+ Google Sheets from scratch
// (Stock alone is ~8,800+ rows) on EVERY SINGLE ?type=live call, with no
// caching anywhere in this file - every login/refresh pays the full
// rebuild cost even though the underlying Tally-synced data only
// actually changes every few minutes.
//
// Fix: reuse the exact same "precompute once, write to a Drive JSON
// file, read it back fast" pattern Daily Freeze already proved out
// (see writeDailyFreezeFile_()/readDailyFreezeFile_() above) instead of
// inventing a new mechanism. refreshLiveCache() is meant to be
// installed as a time-driven trigger (every 5 minutes - see the
// delivery notes for the one-time manual setup step in the Apps Script
// editor, which cannot be done via code alone) that does the slow
// rebuild in the background; handleLiveGet_() then just reads that
// file, which is fast regardless of how big Stock/Outstanding/etc. get.
//
// This is additive and self-healing, not a hard dependency: if the
// trigger is never set up, fails, or stops firing (the exact failure
// mode the Tally Live Watcher's Task Scheduler entry hit on 5 Oct -
// "Ready" in the scheduler but not actually running), readLiveCachePayload_()
// below simply returns null for a missing/corrupt/too-old file and
// handleLiveGet_() transparently falls back to the original always-
// correct live rebuild. Nothing can break by adding this; at worst it
// does nothing and today's slow path keeps working exactly as before.
var LIVE_CACHE_FOLDER_NAME_ = "AAA Live Cache";
var LIVE_CACHE_FILE_NAME_ = "live-cache.json";
// How old a cache file is allowed to be before handleLiveGet_() stops
// trusting it and rebuilds live instead - well above the trigger's own
// 5-minute interval so a couple of missed/slow runs still serve the
// fast path, but low enough that a trigger which silently stopped
// firing days ago can never quietly serve ancient data forever.
var LIVE_CACHE_MAX_AGE_MS_ = 15 * 60 * 1000;

function liveCacheRootFolder_() {
  return getOrCreateFolder_(DriveApp.getRootFolder(), LIVE_CACHE_FOLDER_NAME_);
}
// Overwrite semantics (trash any existing cache file first) - same
// re-run-safe convention as writeDailyFreezeFile_() above. There is
// always at most one live cache file, since there's only ever one
// "right now" to cache.
function writeLiveCachePayload_(payload) {
  var folder = liveCacheRootFolder_();
  var existing = folder.getFilesByName(LIVE_CACHE_FILE_NAME_);
  while (existing.hasNext()) existing.next().setTrashed(true);
  var blob = Utilities.newBlob(JSON.stringify(payload), "application/json", LIVE_CACHE_FILE_NAME_);
  folder.createFile(blob);
}
// Never throws - returns null for anything short of a healthy, fresh
// cache file (missing, corrupt JSON, or older than
// LIVE_CACHE_MAX_AGE_MS_) so handleLiveGet_() below can fall back to
// the live rebuild path with a single, unconditional check.
function readLiveCachePayload_() {
  var folder = liveCacheRootFolder_();
  var it = folder.getFilesByName(LIVE_CACHE_FILE_NAME_);
  if (!it.hasNext()) return null;
  var file;
  try {
    file = JSON.parse(it.next().getBlob().getDataAsString());
  } catch (e) {
    return null;
  }
  if (!file || !file._cachedAt) return null;
  var age = Date.now() - new Date(file._cachedAt).getTime();
  if (isNaN(age) || age > LIVE_CACHE_MAX_AGE_MS_) return null;
  return file;
}

// The function to install as a time-driven trigger (every 5 minutes) -
// see the delivery notes for the one-time manual setup steps. Does the
// full slow rebuild (same call handleLiveGet_() used to make directly
// on every request) and writes the result to Drive for handleLiveGet_()
// to read back quickly. `_cachedAt` lets the dashboard show truthfully
// when this data was actually built, instead of the client's own fetch
// time, when it's serving a cached response - see applyLivePayloadToState_()
// in dashboard/index.html.
// NOTE: deliberately no trailing underscore on this name, unlike every
// other helper in this file - Apps Script hides underscore-suffixed
// ("private") functions from the Run dropdown and the Add Trigger
// function picker, so a name like refreshLiveCache_() can never actually
// be selected there. This one has to be selectable as a trigger target.
function refreshLiveCache() {
  var nowKey = Utilities.formatDate(new Date(), DISPLAY_TIMEZONE, "yyyy-MM");
  var payload = buildLivePayload_(liveRowSource_(), nowKey);
  payload.ok = true;
  payload._cachedAt = new Date().toISOString();
  writeLiveCachePayload_(payload);
}

function handleLiveGet_() {
  var cached = readLiveCachePayload_();
  if (cached) return jsonOut_(cached);
  // Cache missing/stale/not set up yet - fall back to the original,
  // always-correct direct rebuild so the dashboard never breaks because
  // of this feature. No `_cachedAt` stamped here: this payload really
  // is being built this instant, so the client's own "Updated <now>"
  // text is already accurate as-is.
  var nowKey = Utilities.formatDate(new Date(), DISPLAY_TIMEZONE, "yyyy-MM");
  var payload = buildLivePayload_(liveRowSource_(), nowKey);
  payload.ok = true;
  return jsonOut_(payload);
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

// ===================================================================
// MASTER FILE / DEALER TARGETS / REMARKS / TASK MANAGER
// Added 6 Oct 2026 - the owner's "master file" system: per-dealer
// monthly Sales Target (Ltr) + Collection Plan (Rs), fillable by each
// DSR in their own app or edited by Owner/Manager from the dashboard;
// a Suggested Target computed from each dealer's own real Tally sales
// pattern; "Collectable" (45+ days overdue) - this REUSES the existing
// LiveOutstanding "Aged Pending (45d+)" column (agedPendingAmount),
// already pushed live by tally_live_watcher.py for the payment-
// reminder feature - no watcher change needed for this; "Overall
// Outstanding" as a MONTH-START SNAPSHOT, not a live-ticking figure -
// this reuses the existing Monthly Freeze archive of LiveOutstanding
// (freezeLiveSheetForMonth_() already freezes it every month-end, see
// FREEZE_LIVE_CATEGORIES_ above) one month back, since a freeze taken
// on month M's last day IS month M+1's opening balance; a short
// Remarks log per dealer, shown in that dealer's own DSR's app; and a
// Task Manager modeled on the owner's reference screenshots, visible
// to Owner/Manager always, to an Area Manager scoped to their own
// AllowedAreas, and to a DSR only for tasks assigned to or about their
// own area.
//
// New sheets (auto-created on first write via getLiveSheet_(), same
// convention as every Live* sheet above):
//   DealerTargets     - one row per (Month Key, Dealer Name): that
//                        month's Sales Target Ltr / Collection Plan
//                        Rs, who last edited it and when.
//   DealerRemarks     - append-only log, one row per remark.
//   DealerMeta        - static reference data per dealer (DSR, Area,
//                        Town, District, Segment, Beat Day) - seeded
//                        from the owner's own monthly Master File
//                        Excel via seedMasterFileFromRaw() below,
//                        never written by the DSR apps/dashboard.
//   MasterFileHistory - one row per (Dealer Name, Month Key): the
//                        FIXED historical Sales Ltr/Collection Rs for
//                        every month before live Tally tracking
//                        started, plus a once-a-year total for
//                        2018-2023 - also seeded from the Excel.
//   Tasks / TaskComments - the Task Manager.
// ===================================================================

var DEALER_TARGETS_SHEET = "DealerTargets";
var DEALER_TARGETS_HEADERS = ["Month Key", "Dealer Name", "DSR Name", "Sales Target Ltr", "Collection Plan Rs", "Edited By", "Edited At"];

var DEALER_REMARKS_SHEET = "DealerRemarks";
var DEALER_REMARKS_HEADERS = ["ID", "Dealer Name", "Remark", "By", "At"];

var DEALER_META_SHEET = "DealerMeta";
var DEALER_META_HEADERS = ["Dealer Name", "DSR Name", "Area", "Town", "District", "Segment", "Beat Day", "Reference Outstanding"];

// Bank Deposit - added 23 Oct 2026 ("want to add a column as bank deposit
// date, in that want to ask input from dsr paid amount, date, paid thro
// bank (cash deposit/ATM deposit/transfer), cash direct to office, cash
// thro, paid receipt number from/to date" - built into each DSR's own app
// next to the Sales/Collection weekday heatmap, and into the dashboard as
// its own panel with a DSR picker). A DSR logs how/when they paid in cash
// they've collected - NOT derived from Tally like everything else in
// ?type=live, this is purely a DSR-entered reconciliation log, same
// "self-reported, office reviews it" shape as DealerRemarks/Tasks above.
// "Paid Through" stores the plain display label directly (same convention
// as Task Type/Status), not a code, so the raw Sheet reads cleanly.
// "Thro Name" only applies to the "Cash — Thro" mode (who physically
// carried the cash); "Receipt No From/To" is a number RANGE with ONE date
// (the Date column) since a DSR may deposit several pay-in slips at once.
// "Remarks" added 23 Oct 2026 ("add remarks in bank transfer ... thats
// also must") - only applies to (and is required for) "Bank — Transfer",
// same single-mode-only convention as Thro Name, since an electronic
// transfer has no paper receipt number to log instead. Appended as the
// LAST column on purpose, so getLiveSheet_()'s self-healing header-row
// rewrite only ever touches row 1 (never a data row) on a BankDeposits
// sheet that already existed before this column was added - any row
// saved before today simply reads back with an empty Remarks.
var BANK_DEPOSIT_SHEET = "BankDeposits";
// "Receipt From Date", "Receipt To Date" and "Deleted At" appended 23
// Oct 2026 ("NEED A DELETE OR EDIT OPTION ... NEED ASK RECPT FROM AND
// TO DATE NEAR THAT RECP NUMBER" - then corrected same day: "ABOVE DATE
// IS PAYMENT DEPOSIT DATE, BELOW I WANT 2 DATE ... EG RCPT NUMBER 101 DT
// 15.10.2026 TO RCPT NO 105 DT 20.10.2026 PAYMNET DEPOSTIED 20000 TODAY
// SO FOR THIS I NEED 2 DATES NEAR THE RECPT FROM AND TO" - a receipt
// range can span several days of actual receipts, separate from the
// single deposit Date column, so From and To each get their own date).
// All three added as new LAST columns, safe on a sheet that already has
// data (existing rows just read back with these empty). Deleted At is a
// soft-delete stamp, same pattern as the existing Orders sheet's own
// Deleted column.
var BANK_DEPOSIT_HEADERS = ["ID", "DSR Name", "Date", "Amount", "Paid Through", "Thro Name", "Receipt No From", "Receipt No To", "Created At", "Remarks", "Receipt From Date", "Receipt To Date", "Deleted At"];

// Maps LiveDealerRoster's own "DSR Name" (the long Tally GROUP name, e.g.
// "PETRONAS DINDIGUL AREA (ARUN)" - see KNOWN_DSR_GROUP_NAMES in
// tally_live_watcher.py) to the short DSR name DealerMeta/DealerTargets/
// each DSR app actually use (matching DSR_SHORT_NAME in every DSR app's
// own HTML, e.g. "Arun"). Added 6 Oct 2026 for autoAddMissingDealerMeta_()
// below. Saravanan's Tally group maps to "Chellamani", not "Saravanan" -
// his dealers were merged into Chellamani's order-slip APP on 1 Sep 2026
// even though they're still their own separate Group in Tally itself (see
// _merge_saravanan_into_chellamani() in tally_live_watcher.py) - a dealer
// from his group needs to show up in Chellamani's Dealer Targets screen,
// not a "Saravanan" app that no longer takes orders.
var DEALER_ROSTER_GROUP_TO_SHORT_DSR_ = {
  "PETRONAS THENI AREA (PRABHA)": "Prabha",
  "PETRONAS DINDIGUL AREA (ARUN)": "Arun",
  "PETRONAS VIRUDHUNAGAR AREA (NAGARAJ)": "Nagaraj",
  "PETRONAS MDU AREA (CHELLAMANI)": "Chellamani",
  "PETRONAS TVL1 AREA (ANANDH)": "Anandh",
  "PETRONAS NKL AREA (MADHU)": "Madhu",
  "PETRONAS MDU AREA (SARAVANAN)": "Chellamani",
  "PETRONAS AAA DIRECT AREA SALES": "AAA Direct"
};

var MASTER_FILE_HISTORY_SHEET = "MasterFileHistory";
var MASTER_FILE_HISTORY_HEADERS = ["Dealer Name", "Month Key", "Sales Ltr", "Collection Rs", "Is Year Total"];

var TASKS_SHEET = "Tasks";
// Remarks/Reminder Time added 6 Oct 2026 at the end of the row (not
// inserted in the middle) specifically so every existing column-number
// reference elsewhere (handleTaskUpdateStatus_'s Status/Updated At,
// handleTaskAddComment_'s Updated At) keeps working unchanged - only
// code that reads/writes these two new columns needed to change.
// "Deleted At" added 6 Oct 2026, same end-of-row pattern, per the owner's
// explicit request: "after delete also the written should be on sheet as
// deleted" - handleTaskDelete_() below used to permanently deleteRow()
// the task and wipe its whole comment thread, leaving no trace on the
// Sheet at all. Now a delete is a SOFT delete: this column gets an ISO
// timestamp, the row and every one of its comments stay exactly where
// they are, and handleTasksGet_() below just filters deleted tasks out
// of the normal list - same disappears-from-the-UI behavior as before,
// but the complete record (task, status it had, every comment) now
// stays permanently on the Sheet instead of being erased.
// "Dealer Name" added 8 Oct 2026 (the owner: "show dealer name as per
// selected dsr name") - a 15th column at the END, same append-only
// pattern as Remarks/Reminder Time/Deleted At above, so no existing
// column-index reference in this file needed to change. This is PURELY
// a display field (which dealer the task is actually about, e.g.
// "Maharaja payment" -> dealer "Maharaja") - it is never read by
// taskRowVisible_(); visibility still runs entirely on Assigned To/Area/
// Also Visible To, unchanged.
var TASKS_HEADERS = ["ID", "Task Name", "Task Type", "Status", "Due Date", "Created By", "Assigned To", "Also Visible To", "Area", "Created At", "Updated At", "Remarks", "Reminder Time", "Deleted At", "Dealer Name"];

var TASK_COMMENTS_SHEET = "TaskComments";
var TASK_COMMENTS_HEADERS = ["ID", "Task ID", "Author", "Text", "At"];

// Fixed column map (0-based, matching sheet.getRange().getValues() row
// indices) for the owner's own Master File Excel layout, confirmed
// directly against "MASTER_FILE_SALES & COLL - SEP 2026.xlsx" (read 6
// Oct 2026). Columns 0-7 are the dealer-meta block; everything from 8
// onward is either a (monthKey, sales|coll) pair or a plain yearly
// total ("year"). seedMasterFileFromRaw() below is the only place
// this is used - the owner imports a fresh copy of their Excel as a
// sheet tab named exactly "MasterFileRaw" each time they want to
// refresh the Master File's fixed/historical data (their own existing
// monthly habit - "what I usually do every month... fill the last
// month data in it"), then runs seedMasterFileFromRaw once from the
// Apps Script editor. Re-running it is always safe - it fully
// replaces DealerMeta and upserts MasterFileHistory by (dealer,
// monthKey), never touches DealerTargets/DealerRemarks/Tasks.
var MASTER_FILE_RAW_MONTH_COLS_ = [
  { col: 8, monthKey: "2026-09", field: "sales" },
  { col: 10, monthKey: "2026-09", field: "coll" },
  { col: 11, monthKey: "2026-08", field: "sales" },
  { col: 12, monthKey: "2026-08", field: "coll" },
  { col: 13, monthKey: "2026-07", field: "coll" },
  { col: 14, monthKey: "2026-07", field: "sales" },
  { col: 15, monthKey: "2026-06", field: "sales" },
  { col: 16, monthKey: "2026-06", field: "coll" },
  { col: 17, monthKey: "2026-05", field: "sales" },
  { col: 18, monthKey: "2026-05", field: "coll" },
  { col: 19, monthKey: "2026-04", field: "sales" },
  { col: 20, monthKey: "2026-04", field: "coll" },
  { col: 21, monthKey: "2026-03", field: "sales" },
  { col: 22, monthKey: "2026-03", field: "coll" },
  { col: 23, monthKey: "2026-02", field: "sales" },
  { col: 24, monthKey: "2026-02", field: "coll" },
  { col: 25, monthKey: "2026-01", field: "sales" },
  { col: 26, monthKey: "2026-01", field: "coll" },
  { col: 27, monthKey: "2025-12", field: "sales" },
  { col: 28, monthKey: "2025-12", field: "coll" },
  { col: 29, monthKey: "2025-11", field: "sales" },
  { col: 30, monthKey: "2025-11", field: "coll" },
  { col: 31, monthKey: "2025-10", field: "sales" },
  { col: 32, monthKey: "2025-10", field: "coll" },
  { col: 33, monthKey: "2025-09", field: "sales" },
  { col: 34, monthKey: "2025-09", field: "coll" },
  { col: 35, monthKey: "2025-08", field: "sales" },
  { col: 36, monthKey: "2025-08", field: "coll" },
  { col: 37, monthKey: "2025-07", field: "sales" },
  { col: 38, monthKey: "2025-07", field: "coll" },
  { col: 39, monthKey: "2025-06", field: "sales" },
  { col: 40, monthKey: "2025-06", field: "coll" },
  { col: 41, monthKey: "2025-05", field: "sales" },
  { col: 42, monthKey: "2025-05", field: "coll" },
  { col: 43, monthKey: "2025-04", field: "sales" },
  { col: 44, monthKey: "2025-04", field: "coll" },
  { col: 45, monthKey: "2025-03", field: "sales" },
  { col: 46, monthKey: "2025-03", field: "coll" },
  { col: 47, monthKey: "2025-02", field: "sales" },
  { col: 48, monthKey: "2025-02", field: "coll" },
  { col: 49, monthKey: "2025-01", field: "sales" },
  { col: 50, monthKey: "2025-01", field: "coll" },
  { col: 51, monthKey: "2024-12", field: "sales" },
  { col: 52, monthKey: "2024-12", field: "coll" },
  { col: 53, monthKey: "2024-11", field: "sales" },
  { col: 54, monthKey: "2024-11", field: "coll" },
  { col: 55, monthKey: "2024-10", field: "sales" },
  { col: 56, monthKey: "2024-10", field: "coll" },
  { col: 57, monthKey: "2024-09", field: "sales" },
  { col: 58, monthKey: "2024-09", field: "coll" },
  { col: 59, monthKey: "2024-08", field: "sales" },
  { col: 60, monthKey: "2024-08", field: "coll" },
  { col: 61, monthKey: "2024-07", field: "sales" },
  { col: 62, monthKey: "2024-07", field: "coll" },
  { col: 63, monthKey: "2024-06", field: "sales" },
  { col: 64, monthKey: "2024-06", field: "coll" },
  { col: 65, monthKey: "2024-05", field: "sales" },
  { col: 66, monthKey: "2024-05", field: "coll" },
  { col: 67, monthKey: "2024-04", field: "sales" },
  { col: 68, monthKey: "2024-04", field: "coll" },
  { col: 69, monthKey: "2024-03", field: "sales" },
  { col: 70, monthKey: "2024-02", field: "sales" },
  { col: 71, monthKey: "2024-01", field: "sales" },
  { col: 72, monthKey: "2023", field: "year" },
  { col: 73, monthKey: "2022", field: "year" },
  { col: 74, monthKey: "2021", field: "year" },
  { col: 75, monthKey: "2020", field: "year" },
  { col: 76, monthKey: "2019", field: "year" },
  { col: 77, monthKey: "2018", field: "year" }
];

// The first monthKey (inclusive) that Master File treats as "live" -
// filled from Tally via the existing LiveSalesByDealer/
// LiveCollectionByDealer sheets (and, for a past live month, their
// Monthly Freeze archive) instead of the fixed MasterFileHistory
// import. Everything strictly before this stays fixed/historical.
// Moved from "2026-10" to "2026-09" on 10 Oct 2026, per the owner's
// explicit instruction on this same Master File re-import ("FROM SEP
// THAT SHOWLD BE LIVE FROM TALLY"). Sep 2026's Sales/Collection figures
// in this round's MasterFileRaw Excel are still imported into
// MasterFileHistory same as every other month (so there's a real
// number to fall back on / cross-check against), but
// masterFileMonthIsLive_() below now means Sep 2026 onward is always
// read from the live LiveSalesByDealer/LiveCollectionByDealer sheets
// (or their Monthly Freeze archive once Sep closes), never from this
// imported figure.
var MASTER_FILE_LIVE_START_MONTH_ = "2026-09";

function masterFileMonthIsLive_(monthKey) {
  return String(monthKey || "") >= MASTER_FILE_LIVE_START_MONTH_;
}

// Reads DealerMeta into {dealerName: {dsrName, area, town, district,
// seg, beatDay, referenceOutstanding}}.
function dealerMetaMap_() {
  var sheet = getLiveSheet_(DEALER_META_SHEET, DEALER_META_HEADERS);
  var rows = liveSheetRows_(sheet);
  var map = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var name = String(r[0] || "").trim();
    if (!name) continue;
    map[name] = {
      dsrName: String(r[1] || "").trim(),
      area: String(r[2] || "").trim(),
      town: String(r[3] || "").trim(),
      district: String(r[4] || "").trim(),
      seg: String(r[5] || "").trim(),
      beatDay: String(r[6] || "").trim(),
      referenceOutstanding: r[7] === "" || r[7] === undefined ? null : Number(r[7])
    };
  }
  return map;
}

// --- Today's Dealer Visit Check (added 26 Oct 2026) -----------------------
// Owner's request: "by eve report if he did not sync with that dealer can u
// inform me in dashboard with matching of todays dealer vs visited = missed
// dealer." Clarified: "visited" means cross-checked against the daily Shop
// Visit Remarks mail/sheet (NOT whether an order was saved in a DSR app),
// checked once that day's visit data has landed (the separate "Daily DSR
// remarks into Google Sheet" scheduled task runs ~9:15 PM and fills the
// sheet this reads), for all 7 DealerMeta DSRs.
//
// A real diagnostic was run before building this (430 planned-dealer-day
// rows across 5 real days, 20-24 Oct 2026): a naive EXACT name match came
// back ~0%, because DealerMeta's own dealer names always carry a trailing
// town/area suffix ("RANJITH MOTORS, DGL (OIL)") that the visit-log's own
// "Shop" column never has ("RANJIT MOTORS") - fixed below by comparing only
// the text before the first comma, parenthetical tags stripped. Even with
// that fixed, the REAL finding wasn't naming - it was coverage: Chellamani/
// Nagaraj/Prabha showed ZERO logged visits on most of the 5 sampled days,
// while Arun/Anandh/Madhu logged more consistently. The owner confirmed
// this isn't a data bug: Chellamani has quit (role vacant, a replacement
// will be hired later), and Nagaraj/Prabha both have broken phones and will
// resume once they get new ones. So "nothing logged today" for a DSR is a
// real, expected, non-accusatory state this feature must show plainly as
// its own thing - NEVER as "missed every dealer," which would be false and
// unfair to a DSR who simply isn't able to log anything right now.
//
// Matching uses a plain bigram Dice-coefficient string-similarity score (no
// external library available in Apps Script) on the normalized shop name
// (lowercased, parenthetical tag stripped, only the text before the first
// comma kept, punctuation stripped). Calibrated against 5 real name pairs
// found during the diagnostic:
//   "RANJITH MOTORS, DGL (OIL)"            vs "RANJIT MOTORS"        -> 0.880 (same shop)
//   "SHANTHI AUTOMOBILES, SURANDAI (OIL)"  vs "SHANTHI AUTO MOBILES" -> 0.919 (same shop)
//   "KRISHNA AUTOMOBILES,SURANDAI (OIL)"   vs "KRISHNA AUTO MOBILES" -> 0.919 (same shop)
//   "GANESH AUTO WORKS, VEDASANDUR (OIL)"  vs "GANAPATHY AUTO WORKS" -> 0.686 (DIFFERENT shop - false-positive risk)
//   "KANNAN AUTO STORE, PLN (OIL)"         vs "KANCHI AUTO STORE"    -> 0.750 (DIFFERENT shop - false-positive risk)
// DEALER_VISIT_MATCH_HIGH_ (0.85) sits above both false-positive-risk
// scores and below all 3 genuine matches - genuine spelling drift clears
// it, two different shops sharing common words ("AUTO WORKS"/"AUTO STORE")
// do not. DEALER_VISIT_MATCH_LOW_ (0.5) is the floor below which a planned
// dealer is called a real Missed rather than a Possible match worth a human
// glance (shown side by side so a person decides, never silently resolved
// either way).
var DEALER_VISIT_REMARKS_SPREADSHEET_ID_ = "1NCWstQV2BOM6n4hiYb0IJuBQM6zhjqVhoDMv4sILRcM";
var DEALER_VISIT_REMARKS_SHEET_NAME_ = "All Visits";
var DEALER_VISIT_MATCH_HIGH_ = 0.85;
var DEALER_VISIT_MATCH_LOW_ = 0.5;

// Dealer Visit Check Yes/No learning - added 27 Oct 2026 (owner: "for
// possible match ask me a yes or no there if i clicked yes then match
// the dealer with that and next time possible match cannot happen to
// the same and if no then leave and try to match with some another
// next time"). One row per (Dealer, Normalized Shop) pair the owner has
// actually decided on, in THIS SAME spreadsheet (not the separate Visit
// Remarks one) since it's AAA's own decision record, not Tally/visit
// data. "Normalized Shop" is dealerVisitNormalizeName_() applied to the
// logged shop text at decision time - this is deliberately the SAME
// normalization already used for the Dice-score matching, so a Yes/No
// decision keys on the exact spelling variant the DSR actually typed,
// not the real dealer name. "Key" (col 1) is Dealer + "||" + Normalized
// Shop, upserted via bulkUpsertSheet_() same as every other live sheet -
// deciding the same pair again (e.g. owner flips an old No to a Yes)
// overwrites in place rather than piling up duplicate rows.
// See dealerVisitDecisionsMap_()/handleDealerVisitMatchDecision_() below
// and the CONFIRMED/REJECTED handling inside dealerVisitCheckLive_().
var DEALER_VISIT_DECISIONS_SHEET = "DealerVisitMatchDecisions";
var DEALER_VISIT_DECISIONS_HEADERS = ["Key", "Dealer", "Normalized Shop", "Raw Shop Name", "Decision", "Decided By", "Decided At"];

function dealerVisitNormalizeName_(s) {
  s = String(s || "").trim().toLowerCase();
  s = s.replace(/\([^)]*\)/g, "");
  s = s.split(",")[0];
  s = s.replace(/[^a-z0-9 ]/g, " ");
  s = s.replace(/\s+/g, " ").trim();
  return s;
}

function dealerVisitBigrams_(s) {
  var out = [];
  for (var i = 0; i < s.length - 1; i++) out.push(s.substr(i, 2));
  return out;
}

// The daily trigger appends the Date column as "YYYY-MM-DD" text, which
// Google Sheets (USER_ENTERED) auto-converts into a real Date value -
// displayed as dd-mm-yyyy only because of the column's own number format.
// getDataRange().getValues() then hands back a genuine JS Date object for
// that cell, not a string - so a plain String(row[0]) never matches a
// "dd-MM-yyyy" comparison string (it produces something like "Thu Oct 26
// 2026 00:00:00 GMT+0000..."). This silently broke the date match for
// EVERY row, which is why the panel always showed "No visit log" for
// every DSR on every date, however many real rows existed for that date.
// Found and fixed 27 Oct 2026 - see claude/dsr-app-registry.md.
function dealerVisitRowDateStr_(val) {
  if (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, DISPLAY_TIMEZONE, "dd-MM-yyyy");
  }
  var s = String(val || "").trim();
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return iso[3] + "-" + iso[2] + "-" + iso[1];
  return s;
}

function dealerVisitDiceScore_(a, b) {
  if (!a || !b) return 0;
  if (a === b) return 1;
  var bgA = dealerVisitBigrams_(a), bgB = dealerVisitBigrams_(b);
  if (bgA.length === 0 || bgB.length === 0) return 0;
  var map = {};
  for (var i = 0; i < bgA.length; i++) map[bgA[i]] = (map[bgA[i]] || 0) + 1;
  var matches = 0;
  for (var j = 0; j < bgB.length; j++) {
    if (map[bgB[j]] > 0) { matches++; map[bgB[j]]--; }
  }
  return (2 * matches) / (bgA.length + bgB.length);
}

// Reads every Yes/No decision the owner has made so far into two lookup
// maps - confirmed[dealer][normShop] and rejected[dealer][normShop] -
// added 27 Oct 2026 alongside the Dealer Visit Check Yes/No feature (see
// the long comment above DEALER_VISIT_DECISIONS_SHEET). Wrapped in its
// own try/catch for the same reason dealerVisitCheckLive_() wraps the
// Visit Remarks read - a problem here must never break the rest of
// ?type=live, it should just mean no decisions are known yet this call
// (falls back to pure score-based matching, same as before this feature
// existed).
function dealerVisitDecisionsMap_() {
  var confirmed = {}, rejected = {};
  try {
    var sheet = getLiveSheet_(DEALER_VISIT_DECISIONS_SHEET, DEALER_VISIT_DECISIONS_HEADERS);
    var rows = liveSheetRows_(sheet);
    for (var i = 0; i < rows.length; i++) {
      var dealer = String(rows[i][1] || "").trim();
      var normShop = String(rows[i][2] || "").trim();
      var decision = String(rows[i][4] || "").trim().toLowerCase();
      if (!dealer || !normShop) continue;
      if (decision === "confirmed") {
        if (!confirmed[dealer]) confirmed[dealer] = {};
        confirmed[dealer][normShop] = true;
      } else if (decision === "rejected") {
        if (!rejected[dealer]) rejected[dealer] = {};
        rejected[dealer][normShop] = true;
      }
    }
  } catch (e) {
    // Leave both maps empty - matching just behaves as it did before
    // this feature existed.
  }
  return { confirmed: confirmed, rejected: rejected };
}

// POST { action: "dealerVisitMatchDecision", dealer, matchedShop,
// decision: "yes"|"no", requestedBy } - the dashboard's new Yes/No
// buttons on each "Possible match" row (added 27 Oct 2026). "yes" files
// a CONFIRMED decision for this exact (dealer, normalized shop) pair -
// dealerVisitCheckLive_() then treats any future day's visit log entry
// that normalizes to the same text as an automatic Visited for this
// dealer, no more glance needed. "no" files a REJECTED decision - that
// exact normalized shop text is never offered as this dealer's closest
// match again (the matcher falls through to the next-best candidate, or
// Missed if none clears the floor). Deciding the same pair again (owner
// changes their mind) simply overwrites the previous decision - see
// bulkUpsertSheet_()'s upsert-by-key behavior. No role check server-side,
// same lighter convention as handleDealerMetaSet_() above - this panel
// is already Owner-only in the dashboard UI (dealerVisitCheckBlock.hidden
// in afterLogin_()).
function handleDealerVisitMatchDecision_(body) {
  var dealer = String(body.dealer || "").trim();
  var rawShop = String(body.matchedShop || "").trim();
  var decisionIn = String(body.decision || "").trim().toLowerCase();
  if (!dealer || !rawShop) return jsonOut_({ ok: false, error: "dealer and matchedShop required" });
  if (decisionIn !== "yes" && decisionIn !== "no") return jsonOut_({ ok: false, error: "decision must be \"yes\" or \"no\"" });
  var normShop = dealerVisitNormalizeName_(rawShop);
  if (!normShop) return jsonOut_({ ok: false, error: "matchedShop did not normalize to anything usable" });
  var decisionOut = decisionIn === "yes" ? "confirmed" : "rejected";
  var key = dealer + "||" + normShop;
  var sheet = getLiveSheet_(DEALER_VISIT_DECISIONS_SHEET, DEALER_VISIT_DECISIONS_HEADERS);
  bulkUpsertSheet_(sheet, DEALER_VISIT_DECISIONS_HEADERS.length, [{ key: key, dealer: dealer, normShop: normShop, rawShop: rawShop, decision: decisionOut, by: String(body.requestedBy || "").trim() }],
    function (it) { return it.key; },
    function (it, now) {
      return [it.key, it.dealer, it.normShop, it.rawShop, it.decision, it.by, now];
    });
  return jsonOut_({ ok: true, decision: decisionOut });
}

// Sums today's (or whatever targetDate is) Day Book "order" amount per
// dealer, for the Dealer Visit Check panel's new "show me that also"
// cross-check (owner: "if possible from tally live if visited dealer
// gives order or payment show me that also") - added 27 Oct 2026.
// ORDER side only: reads the SAME-spreadsheet Day Book tab for
// targetDate's month (daybookSheetName_() above - Day Book already lives
// in this bound spreadsheet, unlike the separate Visit Remarks one, so
// this needs no cross-spreadsheet open) and sums "Amount (incl GST)" per
// Party for rows on targetDate, using the exact same "Voucher Type is
// literally 'Sales' (case/whitespace-insensitive) = unrelated auto-spares
// business, exclude it" rule handleDaybookMonthGet_() already applies -
// "Sales - PETRONAS"/"Sales - CBE"/anything else synced is kept, so this
// never drifts out of sync with what the Day Book export itself shows.
// Party is Tally's own ledger name for the dealer, which is also where
// DealerMeta's own Dealer Name column originates (both ultimately come
// from the same Tally roster/Excel) - so this matches by EXACT string,
// no fuzzy scoring needed, unlike the shop-visit-remarks side.
// PAYMENT side is NOT covered here - Day Book currently only ever
// receives Sales-type vouchers from tally_live_watcher.py (confirmed
// against the real sheet: no "Receipt" voucher type has ever been
// synced), so there is no live per-dealer per-day payment feed to check
// against yet. Flagged to the owner rather than guessed at - extending
// the watcher to also push Receipt vouchers would need a change on the
// office PC's script plus a Code.gs change here to match it.
// Returns { dealerName: totalAmount, ... } - only for dealers that had
// at least one qualifying row; a dealer with no entry simply has no
// order recorded for that date. Wrapped in try/catch, same self-healing
// reasoning as dealerVisitDecisionsMap_() above - a problem reading Day
// Book must never break the rest of the Dealer Visit Check panel.
function dayBookOrdersForDealersOnDate_(dealerNames, targetDate) {
  var result = {};
  if (!dealerNames || !dealerNames.length) return result;
  try {
    var wanted = {};
    for (var w = 0; w < dealerNames.length; w++) wanted[dealerNames[w]] = true;
    var monthKey = Utilities.formatDate(targetDate, DISPLAY_TIMEZONE, "yyyy-MM");
    var targetDateStr = Utilities.formatDate(targetDate, DISPLAY_TIMEZONE, "yyyy-MM-dd");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(daybookSheetName_(monthKey));
    if (!sheet) return result;
    var rows = liveSheetRows_(sheet);
    var DATE_COL = DAYBOOK_HEADERS.indexOf("Date");
    var VOUCHER_TYPE_COL = DAYBOOK_HEADERS.indexOf("Voucher Type");
    var PARTY_COL = DAYBOOK_HEADERS.indexOf("Party");
    var ALT_QTY_COL = DAYBOOK_HEADERS.indexOf("Alt Qty (Ltr)");
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (normalizeDateKey_(r[DATE_COL]) !== targetDateStr) continue;
      var voucherType = String(r[VOUCHER_TYPE_COL] || "").trim().toLowerCase();
      if (voucherType === "sales") continue; // unrelated auto-spares business - same exclusion as handleDaybookMonthGet_()
      var party = String(r[PARTY_COL] || "").trim();
      if (!wanted[party]) continue;
      var qtyLtr = Number(r[ALT_QTY_COL]) || 0;
      result[party] = (result[party] || 0) + qtyLtr;
    }
  } catch (e) {
    // Leave result empty/partial - the panel just shows no order info
    // rather than failing the whole Dealer Visit Check.
  }
  return result;
}

// Returns the per-DSR breakdown for TODAY only (Asia/Kolkata) - see the big
// comment above. Wrapped so a problem reaching the separate Visit Remarks
// spreadsheet (permissions, the sheet renamed/moved, etc.) can never break
// the rest of ?type=live - same self-healing philosophy as the Live Cache
// fallback elsewhere in this file: on any error this returns an empty list
// plus a short `error` string, rather than throwing.
// dateOverrideIso added 26 Oct 2026 for the dashboard's new date picker
// ("can i get date changing option inside") - optional "YYYY-MM-DD"
// string; when omitted this still means "today" exactly as before, so
// the existing no-arg call from handleLiveGet_() (?type=live, always
// today) is untouched. DealerMeta's Beat Day has no history - this
// still uses TODAY'S beat-day assignment to decide which dealers were
// "planned" on the chosen date, which is an approximation for any date
// whose beat-day assignment has since changed; the dashboard shows a
// note about this next to the date picker.
function dealerVisitCheckLive_(dateOverrideIso) {
  try {
    var targetDate = new Date();
    if (dateOverrideIso) {
      var dp = String(dateOverrideIso).split("-");
      if (dp.length === 3) {
        var parsed = new Date(Number(dp[0]), Number(dp[1]) - 1, Number(dp[2]));
        if (!isNaN(parsed.getTime())) targetDate = parsed;
      }
    }
    var todayStr = Utilities.formatDate(targetDate, DISPLAY_TIMEZONE, "dd-MM-yyyy");
    var todayWeekday = Utilities.formatDate(targetDate, DISPLAY_TIMEZONE, "EEEE");
    // Yes/No learning, added 27 Oct 2026 - see the long comment above
    // DEALER_VISIT_DECISIONS_SHEET and dealerVisitDecisionsMap_().
    var decisions = dealerVisitDecisionsMap_();

    var meta = dealerMetaMap_();
    var plannedByDsr = {};
    for (var dealerName in meta) {
      var m = meta[dealerName];
      if (!m.dsrName || m.beatDay !== todayWeekday) continue;
      if (!plannedByDsr[m.dsrName]) plannedByDsr[m.dsrName] = [];
      plannedByDsr[m.dsrName].push(dealerName);
    }

    var visitRows = [];
    try {
      var visitSs = SpreadsheetApp.openById(DEALER_VISIT_REMARKS_SPREADSHEET_ID_);
      var visitSheet = visitSs.getSheetByName(DEALER_VISIT_REMARKS_SHEET_NAME_);
      if (visitSheet) visitRows = visitSheet.getDataRange().getValues();
    } catch (openErr) {
      return { error: "Could not read Visit Remarks sheet: " + openErr, byDsr: [] };
    }

    var shopsByDsr = {};
    for (var r = 1; r < visitRows.length; r++) {
      var row = visitRows[r];
      var date = dealerVisitRowDateStr_(row[0]);
      if (date !== todayStr) continue;
      var staff = String(row[1] || "").trim();
      var shop = String(row[2] || "").trim();
      if (!staff || !shop) continue;
      if (!shopsByDsr[staff]) shopsByDsr[staff] = [];
      shopsByDsr[staff].push(shop);
    }

    var out = [];
    for (var dsrName in plannedByDsr) {
      var planned = plannedByDsr[dsrName];
      var visitedShops = shopsByDsr[dsrName] || [];
      if (visitedShops.length === 0) {
        out.push({ dsr: dsrName, plannedCount: planned.length, loggedToday: false, visited: [], possible: [], missed: [] });
        continue;
      }
      var normShops = [];
      for (var vi = 0; vi < visitedShops.length; vi++) {
        normShops.push({ raw: visitedShops[vi], norm: dealerVisitNormalizeName_(visitedShops[vi]) });
      }
      var visited = [], possible = [], missed = [];
      for (var pi = 0; pi < planned.length; pi++) {
        var dealer = planned[pi];
        var nd = dealerVisitNormalizeName_(dealer);
        var confirmedForDealer = decisions.confirmed[dealer] || {};
        var rejectedForDealer = decisions.rejected[dealer] || {};
        // A previously CONFIRMED (dealer, normalized shop) pair counts as
        // Visited immediately, bypassing the score threshold entirely -
        // this is what makes a Yes decision "stick" for that exact
        // shop-name spelling on every later day it reappears in the log.
        var confirmedMatch = null;
        for (var ci = 0; ci < normShops.length; ci++) {
          if (confirmedForDealer[normShops[ci].norm]) { confirmedMatch = normShops[ci].raw; break; }
        }
        if (confirmedMatch) {
          visited.push({ dealer: dealer, matchedShop: confirmedMatch, confirmed: true });
          continue;
        }
        // Normal best-score pass, but a previously REJECTED (dealer,
        // normalized shop) pair is skipped as a candidate entirely - a
        // No permanently removes that exact shop-name spelling from ever
        // being offered again for this dealer, so the next-best
        // candidate (or Missed, if none left) is used instead.
        var bestScore = 0, bestShop = null;
        for (var si = 0; si < normShops.length; si++) {
          if (rejectedForDealer[normShops[si].norm]) continue;
          var score = dealerVisitDiceScore_(nd, normShops[si].norm);
          if (score > bestScore) { bestScore = score; bestShop = normShops[si].raw; }
        }
        if (bestScore >= DEALER_VISIT_MATCH_HIGH_) {
          visited.push({ dealer: dealer, matchedShop: bestShop, confirmed: false });
        } else if (bestScore >= DEALER_VISIT_MATCH_LOW_) {
          possible.push({ dealer: dealer, matchedShop: bestShop, score: Math.round(bestScore * 100) });
        } else {
          missed.push(dealer);
        }
      }
      out.push({ dsr: dsrName, plannedCount: planned.length, loggedToday: true, visited: visited, possible: possible, missed: missed });
    }
    // Today's Tally order cross-check, added 27 Oct 2026 ("if visited
    // dealer gives order... show me that also") - one batched Day Book
    // read across every visited dealer from every DSR (not one read per
    // DSR), then each visited entry gets its own orderAmount attached.
    // See dayBookOrdersForDealersOnDate_()'s own comment for why this is
    // ORDER-only (no live per-dealer per-day payment feed exists yet).
    var allVisitedDealers = [];
    for (var oi = 0; oi < out.length; oi++) {
      for (var ovi = 0; ovi < out[oi].visited.length; ovi++) allVisitedDealers.push(out[oi].visited[ovi].dealer);
    }
    var orderQtysLtr = dayBookOrdersForDealersOnDate_(allVisitedDealers, targetDate);
    for (var oj = 0; oj < out.length; oj++) {
      for (var ovj = 0; ovj < out[oj].visited.length; ovj++) {
        var vEntry = out[oj].visited[ovj];
        vEntry.orderQtyLtr = Object.prototype.hasOwnProperty.call(orderQtysLtr, vEntry.dealer) ? orderQtysLtr[vEntry.dealer] : null;
      }
    }
    // Stable order matching the dashboard's own 7-card DSR order.
    var order = ["Prabha", "Arun", "Nagaraj", "Chellamani", "Anandh", "Madhu", "AAA Direct"];
    out.sort(function (a, b) { return order.indexOf(a.dsr) - order.indexOf(b.dsr); });
    // dateIso echoes back the date actually used (today, or the resolved
    // dateOverrideIso) so the dashboard's date picker/label can never drift
    // out of sync with what the server actually computed - added 26 Oct 2026
    // alongside the date-picker feature.
    var dateIso = Utilities.formatDate(targetDate, DISPLAY_TIMEZONE, "yyyy-MM-dd");
    return { error: null, byDsr: out, dateIso: dateIso };
  } catch (e) {
    return { error: String(e), byDsr: [] };
  }
}

// Auto-adds any dealer in a live roster push that DealerMeta doesn't know
// about yet - added 6 Oct 2026 so a brand-new Tally dealer shows up in
// Master File/Dealer Targets the same day the watcher next pushes the
// roster (once daily - see handleLiveDealerRosterUpdate_() above), rather
// than staying invisible until the owner's next monthly Excel re-import.
// DSR Name comes from DEALER_ROSTER_GROUP_TO_SHORT_DSR_ above (falls back
// to the raw Tally group name if a future 9th group is ever added and
// this map hasn't been updated yet, rather than silently dropping it).
// Area/Town/District/Segment/Beat Day are left blank ("Not set") -
// nothing in Tally can supply a beat day, so the owner fills that in once
// from the dashboard's Master File table (see handleDealerMetaSet_()
// below). Never removes or edits an existing DealerMeta row - only ever
// adds ones that are missing.
function autoAddMissingDealerMeta_(rosterItems) {
  if (!rosterItems || !rosterItems.length) return 0;
  var metaSheet = getLiveSheet_(DEALER_META_SHEET, DEALER_META_HEADERS);
  var existing = dealerMetaMap_();
  var toAdd = [];
  var seenThisCall = {};
  for (var i = 0; i < rosterItems.length; i++) {
    var item = rosterItems[i];
    var dealerName = item && item.name ? String(item.name).trim() : "";
    if (!dealerName || existing[dealerName] || seenThisCall[dealerName]) continue;
    seenThisCall[dealerName] = true;
    var groupName = String((item && item.dsrName) || "").trim();
    var shortDsr = DEALER_ROSTER_GROUP_TO_SHORT_DSR_[groupName] || groupName;
    toAdd.push([dealerName, shortDsr, "", "", "", "", "", ""]);
  }
  if (toAdd.length) {
    metaSheet.getRange(metaSheet.getLastRow() + 1, 1, toAdd.length, DEALER_META_HEADERS.length).setValues(toAdd);
  }
  return toAdd.length;
}

// POST { action: "dealerMetaSet", dealer, area, district, beatDay, dsrName }
// - the dashboard's Master File table lets Owner/Manager fill in these
// directly for a dealer DealerMeta doesn't have them for yet (most often
// one that just arrived via autoAddMissingDealerMeta_() above) - a
// lighter-weight path than re-importing the whole Excel for one dealer.
// dsrName added 6 Oct 2026, same request as Area/Beat Day: pick a DSR
// directly from the dashboard for a dealer that has none yet (blank/
// "Vacant Karaikudi"), or reassign one to a different DSR later (a
// dealer moving territory) - previously this column only ever came from
// the Excel import or the roster auto-add, with no way to fix or change
// it in between imports. district added 8 Oct 2026, same reasoning, when
// the dashboard's Master File table switched its editable geography column
// from Area to District ("DT") - Area is still readable/stored (still
// comes from the Excel import), just no longer the one the dashboard table
// edits. Each field is only written when the caller actually sends it
// (undefined = leave alone), so an Area/District/Beat-Day-only save still
// works exactly as before. Town/Segment are untouched either way - those
// still only come from the Excel import.
function handleDealerMetaSet_(body) {
  var dealer = String(body.dealer || "").trim();
  if (!dealer) return jsonOut_({ ok: false, error: "dealer required" });
  var sheet = getLiveSheet_(DEALER_META_SHEET, DEALER_META_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonOut_({ ok: false, error: "Dealer not found" });
  var names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < names.length; i++) {
    if (String(names[i][0] || "").trim() === dealer) {
      if (body.dsrName !== undefined) sheet.getRange(i + 2, 2, 1, 1).setValues([[String(body.dsrName || "").trim()]]);
      if (body.area !== undefined) sheet.getRange(i + 2, 3, 1, 1).setValues([[String(body.area || "").trim()]]);
      if (body.district !== undefined) sheet.getRange(i + 2, 5, 1, 1).setValues([[String(body.district || "").trim()]]);
      if (body.beatDay !== undefined) sheet.getRange(i + 2, 7, 1, 1).setValues([[String(body.beatDay || "").trim()]]);
      return jsonOut_({ ok: true });
    }
  }
  return jsonOut_({ ok: false, error: "Dealer not found" });
}

// One-time import of Beat Day + District values from the "MASTER_FILE_SALES &
// COLL - SEP 2026.xlsx" file the owner uploaded (8 Oct 2026), straight into
// the live DealerMeta sheet - requested alongside the dashboard's Master
// File table switching its editable "Area" column to "DT" (District).
// Only UPDATES a Beat Day/District cell for a dealer ALREADY in DealerMeta
// (matched by exact Dealer Name) - never adds new dealer rows (that's
// autoAddMissingDealerMeta_'s job via the daily Tally roster push) and
// never touches any other DealerMeta column (DSR Name/Area/Town/Segment/
// Reference Outstanding are left exactly as they are). A blank cell in the
// source data (the Excel had no Beat Day, or no District, for that row) is
// skipped, not written - so this can never CLEAR a value that's already
// set, only fill one in or correct it. Safe to run more than once; a
// second run just re-applies the same values.
// HOW TO RUN: Apps Script editor - select importBeatDayDistrictFromSep2026MasterFile
// in the function dropdown, click Run, then check the execution log.
function importBeatDayDistrictFromSep2026MasterFile() {
  var IMPORT_DATA_ = [
  ['AARAV ENTERPRISES,CBE (CBE-OIL)','','COIMBATORE'],
  ['SREE VARI AUTOMOBILES,UPT (OIL)','MONDAY','TIRUPPUR'],
  ['SREE ANNAMALAI AGENCIES,MDU (OIL)','','MADURAI'],
  ['SIVAROOPAM AUTO SPARES, DGL (OIL)','MONDAY','DINDIGUL'],
  ['THE PIONEER AUTO AGENCIES,MDU (OIL)','','MADURAI'],
  ['SASI BEARING ,MDU (ARAPALAYAM)','','MADURAI'],
  ['SHANTHI AUTOMOBILES, SURANDAI (OIL)','WEDNESDAY','TENKASI'],
  ['CHELLAM AUTOMOBILES, TENKASI (OIL)','TUESDAY','TENKASI'],
  ['VINAYAGAMANI AUTO AGENCY,MDU (OIL)','','MADURAI'],
  ['TIRUMANI ATI AUTO STORES, THENI (OIL)','MONDAY','THENI'],
  ['SRI ARUL AUTO AGENCY, MELUR (OIL)','THURSDAY','MADURAI'],
  ['SRI MEENAKSHI AUTO STORES,APT (OIL)','TUESDAY','THENI'],
  ['KANNAN AUTO STORE, PLN (OIL)','THURSDAY','DINDIGUL'],
  ['HIGHWAY AUTO SPARES, SANKARANKOVIL (OIL)','FRIDAY','TENKASI'],
  ['HIGH WAY AUTO SPARE PARTS, PULIYANGUDI (OIL)','FRIDAY','TENKASI'],
  ['SRI EASWARI MOTORS,PKM (OIL)','FRIDAY','THENI'],
  ['VENKATESWARA AUTO STORES, PALAYAMKOTTAI (OIL)','MONDAY','TIRUNELVELI'],
  ['A.S.P.AUTO SPARE PARTS,KKDI (OIL)','TUESDAY','SIVAGANGAI'],
  ['SRI SRINIVASA AUTO STORES,SKL (OIL)','FRIDAY','TENKASI'],
  ['NARAYANAN AUTO STORES,CMR  (OIL)','WEDNESDAY','THENI'],
  ['GOLDEN AUTO SPARES,CBM (OIL)','THURSDAY','THENI'],
  ['RK MOTORS,VALLIOOR (OIL)','FRIDAY','TIRUNELVELI'],
  ['AYYANAR TWO WHEELER WORK SHOP, RJPM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['THIRUMALAI AUTO STORES, CBM','THURSDAY','THENI'],
  ['S.P.K AUTOMOBILES, CBM (OIL)','THURSDAY','THENI'],
  ['SWAMY AUTO SPARE PARTS,ODC','THURSDAY','DINDIGUL'],
  ['SURABI MOTORS, KARAIKUDI (OIL)','TUESDAY','SIVAGANGAI'],
  ['SRI VINAYAKA MOTORS,MDU (OIL)','','MADURAI'],
  ['SRI BALA GURUNATHAN AUTO SPARES, RJPM (OIL)','THURSDAY','VIRUDHUNAGAR'],
  ['VISAKAN MOTORS, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['MAGNUM MOTORS, DGL (OIL)','MONDAY','DINDIGUL'],
  ['DEVIS AUTO PARTS,KKDI (OIL)','','SIVAGANGAI'],
  ['SUN AUTO STORE, CUMBUM (OIL)','THURSDAY','THENI'],
  ['RAMM RADHA AUTO SPARES,DVPM (OIL)**','THURSDAY','VIRUDHUNAGAR'],
  ['BALA KIRUTHIKA AUTO SPARES , VADAMADURAI (OIL)','WEDNESDAY','DINDIGUL'],
  ['S.T.R.AUTO STORES, UDAIYARPATTI (OIL)**','SATURDAY','TIRUNELVELI'],
  ['SRI AMMAN AUTO STORES KKDI (OIL)','TUESDAY','SIVAGANGAI'],
  ['GUGAN AUTO WORKS, KRISHNAPURAM (OIL)','THURSDAY','VIRUDHUNAGAR'],
  ['RAMSAKTHI AUTO SPARES, AMBAI (OIL)','THURSDAY','TIRUNELVELI'],
  ['AMAR AUTO STORES, KOVILPATTI (OIL)','MONDAY','TUTICORIN'],
  ['SRI DHARSINI ENTERPRISES, SVK (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['SATHIYA OIL AND  SPARE PARTS, VIRUDHUNAGAR (OIL)','TUESDAY','VIRUDHUNAGAR'],
  ['M.M.MOTORS, SIVAKASI (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['SRI AYYAPPAN TWO WHEELER WORKS, RAJAPALAYAM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['BS ENTERPRISES, IRAVATHANALLUR (OIL)','TUESDAY','MADURAI'],
  ['SHRI GURU AGENCIES, SATTUR (OIL)','MONDAY','VIRUDHUNAGAR'],
  ['BAVANI SRI AUTO STORES, NILAKOTTAI (OIL)','TUESDAY','DINDIGUL'],
  ['SRII THIRUMALAI AUTO SPARE,SVK (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['SRI BALAJI BATTERY CENTRE, KAYATHAR (OIL)','SATURDAY','TUTICORIN'],
  ['SALAM AUTO SPARES, CUMBUM (OIL)','THURSDAY','THENI'],
  ['SRI ANNAI AUTOMOBILES, MADURAI (OIL)','TUESDAY','MADURAI'],
  ['SRI VIVEK AUTO STORES,MDU (OIL)','WEDNESDAY','MADURAI'],
  ['SRI ANNAI MUTHUMARIAMMAN AUTO STORE, APT (OIL)','TUESDAY','THENI'],
  ['STAR AUTO WORKS, NATHAM (OIL)','SATURDAY','DINDIGUL'],
  ['S.MANIMEGALAI WORK SHOP SERVICE,KOTTAMPATTI (OIL)','THURSDAY','MADURAI'],
  ['SRI VINAYAGA AUTO STORE, ODC (OIL)','THURSDAY','DINDIGUL'],
  ['M.J.AUTOMOBILES, NKL (OIL)','MONDAY','KANNIYAKUMARI'],
  ['ARASAN AUTO SPARES, SANKARANKOVIL (OIL)','FRIDAY','TENKASI'],
  ['SRI SIVASAKTHI AUTOMOBILES, BATLAGUNDU (OIL)','TUESDAY','DINDIGUL'],
  ['ALEX AUTO STORE, SVPTR (OIL)','FRIDAY','VIRUDHUNAGAR'],
  ['EMPIRE AUTOMOBILES AGENCY,TVL (OIL)','MONDAY','TIRUNELVELI'],
  ['SAKTHI AUTO WORKS, KADAYAM (OIL)','THURSDAY','TENKASI'],
  ['VARSHAN AUTO SPARES, AMBAI (OIL)','THURSDAY','TIRUNELVELI'],
  ['ANBU TYRES,MDU (OIL)','MONDAY','MADURAI'],
  ['SARAVANA AUTO STORES,WATRAP','FRIDAY','VIRUDHUNAGAR'],
  ['MANI EASWARI TWO WHEELER WORKS, MDU (OIL)','','MADURAI'],
  ['AL MADHINA AUTO STORE, MDU (OIL)','TUESDAY','MADURAI'],
  ['PRAKASH AUTOMOBILES,RAJAPALAYAM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['RAJA AUTO STORE. GOPALPATTY (OIL)','SATURDAY','DINDIGUL'],
  ['Karthick Auto Store,Mdu (Oil)','WEDNESDAY','MADURAI'],
  ['SRI SRINIVAASA AUTO STORES, MDU (OIL)','TUESDAY','MADURAI'],
  ['Selvalakshmi Automobiles And Elect, Anlr (Oil)','','MADURAI'],
  ['M.S.S. SON,  PMDU','','MADURAI'],
  ['SANKAR TRACTOR WORK SHOP, VNR (OIL)','TUESDAY','VIRUDHUNAGAR'],
  ['KUMAR AUTO WORKS, THACHANALLUR (OIL)','MONDAY','TIRUNELVELI'],
  ['PAVITHRA AUTO CARRIAGE, OZHUGINASERY (OIL)','THURSDAY','KANNIYAKUMARI'],
  ['RJ MOTORS, ERAVIPUTHOORKADAI (OIL)','TUESDAY','KANNIYAKUMARI'],
  ['K.K.MOTORS, THENKASI (OIL)','TUESDAY','TENKASI'],
  ['KHALITH AUTOMOBILES, KNLR','WEDNESDAY','TENKASI'],
  ['S.K.AUTOMOBILES, SANKARANKOVIL (OIL)','FRIDAY','TENKASI'],
  ['JAYA AUTOMOBILES, USILAMPATTI (OIL)','FRIDAY','MADURAI'],
  ['VELAVAN MILL STORE, THENI (OIL)','MONDAY','THENI'],
  ['RAMYA AUTO STORES, CBM (OIL)','THURSDAY','THENI'],
  ['MURUGESAN AUTO WORKS, VLR (OIL)','FRIDAY','TIRUNELVELI'],
  ['ARUL,KARAIKUDI (OIL)','TUESDAY','SIVAGANGAI'],
  ['CROWN AUTO,VEDASANDUR (OIL)','WEDNESDAY','DINDIGUL'],
  ['SIVAGURU AUTO CARAGE, ODC (OIL)','THURSDAY','DINDIGUL'],
  ['VIGNESH MOTORS, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['AKILA AUTO SPARES, VALLIOOR (OIL)','FRIDAY','TIRUNELVELI'],
  ['S.S.AUTOMOBILES, TVL (OIL)','MONDAY','TIRUNELVELI'],
  ['SRI CHENDHUR AUTO SPARES, TVL PALAY (OIL)','SATURDAY','TIRUNELVELI'],
  ['SRI VINAYAGAM AUTO STORES, SANKARANKOVIL (OIL)','FRIDAY','TENKASI'],
  ['VISHAKAN AUTOMOBILES, BODI (OIL)','MONDAY','THENI'],
  ['N.S.K.AUTOMOBILES, THENI (OIL)','MONDAY','THENI'],
  ['SAKTHI AUTOMOBILES, KARAIKUDI (OIL)','TUESDAY','SIVAGANGAI'],
  ['MARI, TIRUPPUVANAM (OIL)','FRIDAY','SIVAGANGAI'],
  ['SRI MADURA AUTO SPARES,NATHAM ©','SATURDAY','DINDIGUL'],
  ['R.V.AUTO STORES,NKT (OIL)','TUESDAY','DINDIGUL'],
  ['NEW CHENNAI DIESEL SPARES ,VALLIYUR (OIL)','FRIDAY','TIRUNELVELI'],
  ['RM.MOTORS, RAJARATHINAM NAGAR (OIL)','FRIDAY','TIRUNELVELI'],
  ['SAI CARS, VEDASANDUR (OIL)','WEDNESDAY','DINDIGUL'],
  ['S.M.S AUTO STORE, DGL (OIL)','MONDAY','DINDIGUL'],
  ['M.S.K SPARE PARTS, BGU (OIL)','TUESDAY','DINDIGUL'],
  ['SURIYA AUTO STORES, CBM (OIL)','THURSDAY','THENI'],
  ['GANESHA MOORTHY OIL STORE, KKD (OIL)','TUESDAY','SIVAGANGAI'],
  ['ANNAI AUTO PARTS, KKNL (C)','FRIDAY','DINDIGUL'],
  ['SUPER MILL STORE, MELUR (OIL)','THURSDAY','MADURAI'],
  ['ISWARYA AUTO STORES, SANKARANKOVIL (OIL)','FRIDAY','TENKASI'],
  ['TRACK AUTOMOTIVE, KKMRI (OIL)','WEDNESDAY','KANNIYAKUMARI'],
  ['SRI GANAPATHY TWO WHEELER WORKS, K.KOVIL (OIL)','FRIDAY','VIRUDHUNAGAR'],
  ['SATHISH TWO WHEELER WORKS, NAHALAPURAM (OIL)','MONDAY','TUTICORIN'],
  ['GOLDEN GARAGES, RAJAPALAYAM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['SUSI MOTORS, SVK (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['M.J.MANI TWO WHEELER WORKS, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['Renuga Battery Oil And Tyres,Mdu (Oil)','TUESDAY','MADURAI'],
  ['SIVA SAKTHI TWO WHEELER SPARES, THOPPUR (OIL)','WEDNESDAY','KANNIYAKUMARI'],
  ['MUTHUKUMAR TWO WHEELER WORKS, REDDIARPATTI (OIL)','SATURDAY','TIRUNELVELI'],
  ['SREE LAKSHMI AUTO SPARE PARTS, THENTHAMARAIKULAM (O','WEDNESDAY','KANNIYAKUMARI'],
  ['SRI VETRIVINAYAGAR AUTOMOBILES, PALAYAMKOTTAI (OIL','SATURDAY','TIRUNELVELI'],
  ['GANESH AUTO WORKS, VEDASANDUR (OIL)','WEDNESDAY','DINDIGUL'],
  ['ROYAL AUTO WORKS, SIVAKASI (OIL)','TUESDAY','VIRUDHUNAGAR'],
  ['PONNI AUTOMOBILES,MDU (R)','','MADURAI'],
  ['RAJA TRACTOR SPARE,USILAI  (OIL) **','FRIDAY','MADURAI'],
  ['SRI JOTHI TRADERS, SIVAKASI (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['OM SARAVANA AUTO STORE, CHINNAMANUR (OIL)','WEDNESDAY','THENI'],
  ['ASOKA MOTORS, VNR (OIL)','TUESDAY','VIRUDHUNAGAR'],
  ['AJI AUTOMOBILES,  SVK (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['BANUMATHI STORE, KARUNGALAKUDI (R)','','MADURAI'],
  ['MANIRAM MOTORS ,MDU (OIL)**','TUESDAY','MADURAI'],
  ['Sri Krishna Auto Spares And Works, Vallioor (Oil)','FRIDAY','TIRUNELVELI'],
  ['PRANAV SRI AUTOS, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['S.M.K AUTO STORE, SVK (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['KEERTHIGA AUTO STORES, CHINNAMANUR (OIL)','WEDNESDAY','THENI'],
  ['TAJ AUTO STORES,RJPM  (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['CHALLENGER MOTORS, NKL (OIL)','MONDAY','KANNIYAKUMARI'],
  ['SRI VINAYAGA AUTO STORES, SRIVILLIPUTHUR (OIL)','FRIDAY','VIRUDHUNAGAR'],
  ['SRI BALAJI AUTO SPARES, DEVATHANAPATTI (OIL)','FRIDAY','THENI'],
  ['SUKI AUTO SPARES,CNMR  (OIL)','WEDNESDAY','THENI'],
  ['ARUN TYRES,VNR (OIL)','TUESDAY','VIRUDHUNAGAR'],
  ['KPS TRADERS, CUMBUM (OIL)','THURSDAY','CUMBAM'],
  ['HINDUSTAN AGENCIES, NAGERCOIL (OIL)','MONDAY','KANNIYAKUMARI'],
  ['ARUN AUTO PARTS,MDU (R)','','MADURAI'],
  ['SUNDHAR MOBILES AND2 WHEELER SPARES,ATHIKARIPATT (O','SATURDAY','DINDIGUL'],
  ['MURUGAN FOUR WHEELER SPARE PARTS, KANYAKUMARI (OIL','WEDNESDAY','KANNIYAKUMARI'],
  ['GANAPATHY AUTO WORKS, VEDASANDUR (OIL)','WEDNESDAY','DINDIGUL'],
  ['KUMARAN TRACTOR SPARES,USILAI','FRIDAY','MADURAI'],
  ['SRI PARANJOTHI MOTORS, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['A.M. AUTO MOBILES, KUTTHUKKALVALASAI (OIL)','TUESDAY','TENKASI'],
  ['SPARK AUTO SPARES, NAGERCOIL (OIL)','MONDAY','KANNIYAKUMARI'],
  ['VS BIKES TWO-WHEELER SPARE,TVL (OIL)','SATURDAY','TIRUNELVELI'],
  ['VINCENT AUTO SPARES & WORKS, TVL (OIL)','MONDAY','TIRUNELVELI'],
  ['RANJITH MOTORS, DGL (OIL)','WEDNESDAY','DINDIGUL'],
  ['SRI VENKATESWARA AUTO STORES,KMLI (R)','MONDAY','TUTICORIN'],
  ['JS TYRES AND LUBRICANTS, SETTIYARPATTI (OIL)','THURSDAY','VIRUDHUNAGAR'],
  ['SPR MOTORS,MDU (OIL)','MONDAY','MADURAI'],
  ['ASHA AUTO WORKS, REETHAPURAM (OIL)','TUESDAY','KANNIYAKUMARI'],
  ['INIYAL AUTO SPARES, SVG (OIL)','FRIDAY','SIVAGANGAI'],
  ['CITY AUTO SPARES, COLACHEL (OIL)','TUESDAY','KANNIYAKUMARI'],
  ['SRI PON AYYANAR MOTORS, SIVAKASI (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['KUMAR AUTO WORKS, APK (OIL)','MONDAY','VIRUDHUNAGAR'],
  ['VENKAT TWO WHEELER SPARES, USILAI (OIL)','FRIDAY','MADURAI'],
  ['MOTO STAR, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['SURESH AUTO SPARE PARTS, SVG (OIL)','FRIDAY','SIVAGANGAI'],
  ['Sri Maruthi Auto Stores,Villapuram Mdu (Oil)','WEDNESDAY','MADURAI'],
  ['KANCHI AUTO SPARES,ODC (OIL)','THURSDAY','DINDIGUL'],
  ['M.SUBA AUTO SPARES,USILAMPATTI (OIL)','TUESDAY','MADURAI'],
  ['SMR MOTORS, THIRUPALAI MDU (OIL)','MONDAY','MADURAI'],
  ['SRI GANESH AUTO AGENCY, PERIYAKULAM (OIL)','FRIDAY','THENI'],
  ['MOTO SPARK,MDU (OIL)','WEDNESDAY','MADURAI'],
  ['Ali Auto And Two Wheeler Spare Parts, Tenkasi (Oil)','TUESDAY','TENKASI'],
  ['OM MURUGA CAR DECORS & FANCY STORE, DGL (OIL)','MONDAY','DINDIGUL'],
  ['PRAKASH AUTOS, SVK (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['SRI AYYANAR MOTORS, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['GOLDEN GREEN AGRO SPARES, CUMBUM','THURSDAY','THENI'],
  ['VINAYAGA AUTO STORE, RJPM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['ANBU AUTO SPARE PARTS, ODC (OIL)','THURSDAY','DINDIGUL'],
  ['Sastha Enterprises,Mdu (Oil)','WEDNESDAY','MADURAI'],
  ['SRI SASTHAA LUBRICANTS, MDU (OIL)','MONDAY','MADURAI'],
  ['ESTEEM ENGINEERS, MDU  (OIL)','TUESDAY','MADURAI'],
  ['SIVA SAKTHI AUTOMOBILES, THIRUTHANGAL (OIL)**','TUESDAY','VIRUDHUNAGAR'],
  ['SRI NEW MARUTI MOTORS, VADAMADURAI (OIL)','FRIDAY','DINDIGUL'],
  ['SRI PALANIVELAN AUTO SPARE PARTS, KARAIKUDI (OIL)','TUESDAY','SIVAGANGAI'],
  ['SRI SWAMY AUTO STORE,BODI (OIL)','MONDAY','THENI'],
  ['SRI SENTHIL MURUGAN AUTO AGENCY,KKDI (OIL)','TUESDAY','SIVAGANGAI'],
  ['SIGARAM AUTOMOBILES AND HARDWARES, ANDIPATTI (OIL)','TUESDAY','THENI'],
  ['RAJESH TWO WHEELER WORKS, MULAGUMODU (OIL)','TUESDAY','KANNIYAKUMARI'],
  ['POTHI MOTORS, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['A to Z AUTO SPARES, NAGERCOIL (OIL)','WEDNESDAY','KANNIYAKUMARI'],
  ['ARUN MOTORS, BODINAYAKANUR (OIL)','MONDAY','THENI'],
  ['SUDHARSAN AUTO STORES, APT (OIL)','TUESDAY','THENI'],
  ['LAKSHMI AUTOMOBILES, CUMBUM (OIL)','THURSDAY','THENI'],
  ['SRI RAM & CO.,RJPM (OIL) **','SATURDAY','VIRUDHUNAGAR'],
  ['EVEREST AUTO STORE,  CUMBUM (OIL)','THURSDAY','THENI'],
  ['SIGARAM AUTOMOBILES, THENI (OIL)','MONDAY','THENI'],
  ['SRI VIGNESH AUTOMOBILES, BODI (OIL)','MONDAY','THENI'],
  ['SUSII MOTORS ,KKDI (OIL)','TUESDAY','SIVAGANGAI'],
  ['S.S AUTO SPARES, MDU (OIL)','TUESDAY','MADURAI'],
  ['SHRI GANAPATHY TYRE AGENCIES, MADURAI (OIL)','WEDNESDAY','MADURAI'],
  ['C.K.R.TWO WHEELER SPARES, USILAMPATTI (OIL)','FRIDAY','MADURAI'],
  ['PALANIAPPA TRADERS, KGLKDI (R)','THURSDAY','MADURAI'],
  ['BALAJI AUTO SPARES, GANDHIGRAM (OIL)','TUESDAY','DINDIGUL'],
  ['SARGURU AUTO SPARES, SEMPATTI (OIL)','TUESDAY','DINDIGUL'],
  ['SRI KAMATCHIAMMAN SPARE PARTS,PALAMEDU (OIL)','','MADURAI'],
  ['MARUTHI AGENCY,DGL (OIL)','MONDAY','DINDIGUL'],
  ['S.P.S.Motors, Elumalai','FRIDAY','MADURAI'],
  ['SHANMUGA BIKE POINT, DINDIGUL (OIL)','MONDAY','DINDIGUL'],
  ['RAJALAKSHMI TRADING CO, THIRUMANGALAM (OIL)','SATURDAY','MADURAI'],
  ['K.K.AUTO WORKS, DINDIGUL (OIL)','MONDAY','DINDIGUL'],
  ['KRISHNA AUTOMOBILES ,THENI (OIL)','MONDAY','THENI'],
  ['SRI KRISHNA SPARE PARTS, KALUNEERKULAM (OIL)','WEDNESDAY','TENKASI'],
  ['KAVIN WORKSHOP, MANAMADURAI (OIL)','FRIDAY','SIVAGANGAI'],
  ['BALU AUTO STORE, NAHALAPAURAM (OIL)','MONDAY','TUTICORIN'],
  ['SRI RAJAKALIAMMAN TWO WHEELER WORKS MDU (OIL)','MONDAY','MADURAI'],
  ['ARS AUTO WORKS, NATHAM (OIL)','SATURDAY','DINDIGUL'],
  ['S.K.V.M. AUTO WORKS, R.R.NAGAR (OIL)','TUESDAY','VIRUDHUNAGAR'],
  ['EASWARI SPARE PARTS, SVK (OIL)','TUESDAY','VIRUDHUNAGAR'],
  ['KARTHIKA AUTOMOBILES, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['ROYAL AUTO SPARES & HARDWARD, T.KALLUPATTI (OIL)','SATURDAY','MADURAI'],
  ['SALEEM MOTORS,  MADURAI (OIL)','SATURDAY','MADURAI'],
  ['SHANA AUTO SPARES, MARTHANDAM (OIL)','TUESDAY','KANNIYAKUMARI'],
  ['OM VINAYAGA SPARE PARTS, KARAIKUDI (OIL)','TUESDAY','SIVAGANGAI'],
  ['SRI LAKSHMI MOTORS, NAGERCOIL (OIL)','MONDAY','KANNIYAKUMARI'],
  ['SRI KRISHNA BATTERY OIL TYRES, MDU (OIL)','MONDAY','MADURAI'],
  ['BMS SERVICE POINT, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['TYRES WORLD, CHECKANURANI (OIL)','FRIDAY','MADURAI'],
  ['V.R.AUTOMOTIVE, THIRUPPUVANAM (OIL)','FRIDAY','MADURAI'],
  ['SRI SANKAR AUTO SPARES, SRIVILLIPUTHUR OIL)','FRIDAY','VIRUDHUNAGAR'],
  ['SAI MEENAKSHI AUTO STORES, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['SRI EASWARI MOTORS,PKM','FRIDAY','THENI'],
  ['NEW MUTHULAKSHMI AUTO STORE,  SIVAKASI (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['Smile Bay Enterprises,Mdu (Oil)','SATURDAY','MADURAI'],
  ['RAJ MOTORS, NKT (OIL)','TUESDAY','DINDIGUL'],
  ['SENTHIL SPARE PARTS, MDU (OIL)','TUESDAY','MADURAI'],
  ['PETHANATCHI AUTO STORE, USILAMPATTI (OIL)','FRIDAY','MADURAI'],
  ['SRI MEENAKSHI MOTORS, SVK (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['JEGAN AUTO STORE, AUNDIPATTI (OIL)','TUESDAY','THENI'],
  ['K.P.M.AUTO STORE, ERICHANATHAM (OIL)','FRIDAY','VIRUDHUNAGAR'],
  ['SRI SAASTHA OIL STORE, NILAKKOTTAI (OIL)','TUESDAY','DINDIGUL'],
  ['SRI KRISHNA TRADERS,VNR','TUESDAY','VIRUDHUNAGAR'],
  ['SHANTHI AUTO SPARES, SURANDAI (OIL)','WEDNESDAY','TENKASI'],
  ['AR AUTO SPARES,SHENCOTTAI (OIL)','TUESDAY','TENKASI'],
  ['BRAGADHEESH AUTO STORES, TVL (OIL)','MONDAY','TIRUNELVELI'],
  ['GRACE POWER TECH, AMBASAMUDRAM (OIL)','THURSDAY','TIRUNELVELI'],
  ['HIGH WAY AUTOMOBILES, KADAYANALLUR (OIL)','WEDNESDAY','TENKASI'],
  ['HONEST BIKE CARE, KELA SURANDAI (OIL)','WEDNESDAY','TENKASI'],
  ['KALAI TWO WHEELER WORKS, SURANDAI (OIL) **','WEDNESDAY','TENKASI'],
  ['KANI AUTO WORKS & SPARES, VANNIKONENDAL (OIL) **','FRIDAY','TIRUNELVELI'],
  ['A.J.OIL STORE,  CHUNKANKADAI (OIL)','WEDNESDAY','KANNIYAKUMARI'],
  ['KOTHAI AUTO SPARES, TENKASI (OIL)','TUESDAY','TENKASI'],
  ['KULASAI AUTO WORKS, TIRUNELVELI (OIL)','MONDAY','TIRUNELVELI'],
  ['MAHALAKSHMI TWO WHEELER WORKS, TVL (OIL)','MONDAY','TIRUNELVELI'],
  ['MANI AUTOMOBILE, ALANGULAM (OIL)','WEDNESDAY','TENKASI'],
  ['PANDIAN AUTO SPARES, AMBASAMUDRAM (OIL)**','THURSDAY','TIRUNELVELI'],
  ['R.J AUTO STORES,TVL (OIL)','SATURDAY','TIRUNELVELI'],
  ['SAKTHI UMAYAL AUTOMOBILES, ALANGULAM (OIL)**','WEDNESDAY','TENKASI'],
  ['SAKTHIVEL ALL CAR SERVICE , MANNARKOVIL (OIL)','THURSDAY','TIRUNELVELI'],
  ['SANTHANAMARI AUTO WORKS, TIRUNELVELI (OIL)','MONDAY','TIRUNELVELI'],
  ['SARAN AUTO PARTS, TVL (OIL)','MONDAY','TIRUNELVELI'],
  ['ALDEN AUTOMOBILES, NAGERCOIL (OIL)','THURSDAY','KANNIYAKUMARI'],
  ['SARATHA AUTOMOBILES, ALANGULAM (OIL)','WEDNESDAY','TENKASI'],
  ['SMS CARE, TIRUNELVELI (OIL)','SATURDAY','TIRUNELVELI'],
  ['SANKARA GOMATHI AUTO SPARES, SANKARANKOVIL (OIL)','FRIDAY','TENKASI'],
  ['SRI VARALAKSHMI AUTO AGENCIES, AMBAI (OIL)','THURSDAY','TIRUNELVELI'],
  ['SRI VARALAKSHMI AUTO SPARES,AMBAI (R)','THURSDAY','TIRUNELCELI'],
  ['SRI VIGNESH AUTO MOBILES, PKT (OIL)','SATURDAY','TIRUNELVELI'],
  ['BAGAVATHI TWO WHEELER WORKS, EATHAMOZHY (OIL)','THURSDAY','KANNIYAKUMARI'],
  ['BALAJI SPARES, VALLIYOOR (OIL)','FRIDAY','TIRUNELVELI'],
  ['BRIGHT TWO WHEELER WORKS, VALLIOOR (OIL)','FRIDAY','TIRUNELVELI'],
  ['SRI KAMATCHI AMMAN MOTORS, SVK (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['THIRUPPATHY BATTERY AUTO SPARES, SANKARANKOVIL (OIL','FRIDAY','TENKASI'],
  ['VASANTHAM AUTOMOBILES, TENKASI  (OIL)','TUESDAY','TENKASI'],
  ['VELMURUGAN AUTO PARTS & FINANCE, SENCOTTAH (OIL)','TUESDAY','TENKASI'],
  ['VENKATESWARA AUTO SPARES,TVL (OIL)','SATURDAY','TIRUNELVELI'],
  ['DERISH AUTO SPARES, KULASEKARAM (OIL)','TUESDAY','KANNIYAKUMARI'],
  ['JK AUTOMOBILES,SVK (OIL)','TUESDAY','VIRUDHUNAGAR'],
  ['KARTHICK CAR MECHANICAL WORKS, VALLIOOR (OIL)','FRIDAY','TIRUNELVELI'],
  ['KRISHNA AUTOMOBILES, ANJUGRAMAM (OIL)','FRIDAY','KANNIYAKUMARI'],
  ['MOTO SHOP TWO WHEELER WORK SHOP, CKLM (OIL)','MONDAY','KANNIYAKUMARI'],
  ['MP MOTORS, VALLIOOR (OIL)','FRIDAY','TIRUNELVELI'],
  ['P.S.S. AUTO PARTS, MARTHANDAM (OIL) *****','TUESDAY','KANNIYAKUMARI'],
  ['PATHUVAI MOTOR WORKS, VLR (OIL)','FRIDAY','TIRUNELVELI'],
  ['SRI VELAVAN TWO WHEELER WORKS, ODDANCHATHRAM (OIL)','THURSDAY','DINDIGUL'],
  ['DREAM MOTORS, THADIKOMBU (OIL)','WEDNESDAY','DINDIGUL'],
  ['R.K.TWO WHEELER WORKS, SUCHINDRUM (OIL)','MONDAY','KANNIYAKUMARI'],
  ['MURUGALAKSHMI AUTO SPARES,SVK (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['SRI SAKTHIVEL AUTO STORE, KOVILPATTI (OIL)','MONDAY','TUTICORIN'],
  ['VINOTHINI INIYA TWO WHEELER WORKS. RJPM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['SRI MURUGAN AUTOMOBILES, SHOLAVANDHAN (OIL)','WEDNESDAY','MADURAI'],
  ['THANGAM AUTO STORE, SRIVILLIPUTTUR (OIL)','FRIDAY','VIRUDHUNAGAR'],
  ['S.R.M. MOTORS ,THUCKALAY (OIL)******','','KANNIYAKUMARI'],
  ['CLASSICAL CAR CARRIAGE, ARUPPUKOTTAI (OIL)','MONDAY','VIRUDHUNAGAR'],
  ['RAGUMAN AUTO STORE, R.R.NAGAR (OIL)','MONDAY','VIRUDHUNAGAR'],
  ['ANIKA LUBRICANTS, RAJAPALAYAM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['SAI RAM TRADERS, PARVATHIPURAM (OIL)','MONDAY','KANNIYAKUMARI'],
  ['DEVI AUTO SPARE, KOVILPATTI (OIL)','MONDAY','TUTICORIN'],
  ['G.K.MOTORS, SVK (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['GMC OILS AND SPARES, SETTIYARPATTI  (OIL)','THURSDAY','VIRUDHUNAGAR'],
  ['SRI RAM ENTERPRISES, APK (OIL)','MONDAY','VIRUDHUNAGAR'],
  ['HIGHWAY AUTO SPARE PARTS, SRIVILLIPUTTUR (OIL)','FRIDAY','VIRUDHUNAGAR'],
  ['JEYAM TWO WHEELER WORKS, VIRUDHUNAGAR (OIL)','TUESDAY','VIRUDHUNAGAR'],
  ['JAVED AUTO STORE, PALANI (OIL)','THURSDAY','DINDIGUL'],
  ['K.R.AUTOMOBILES, RJPM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['GARUDA MOTORS, UDUMALPET (OIL)','FRIDAY','TIRUPUR'],
  ['JANATH AUTO SPARES, GOPALPATTI (OIL)','SATURDAY','DINDIGUL'],
  ['KEERTHI AUTO SPARES, APK (OIL)','MONDAY','VIRUDHUNAGAR'],
  ['LAKSHMI AUTO WORKS, RAJAPALAYAM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['H T AUTO WORKS, DGL (OIL)','MONDAY','DINDIGUL'],
  ['NEW SAKTHI TWO WHEELER SPARE PARTS, TPVNM (OIL)','FRIDAY','SIVAGANGAI'],
  ['SRI BABY AUTO WORKS, VELLODU (OIL)','MONDAY','DINDIGUL'],
  ['AS AUTO WORKS,DGL (OIL)','MONDAY','DINDIGUL'],
  ['NIZAM AUTO WORKS, DGL (OIL)','MONDAY','DINDIGUL'],
  ['M S AUTO CARE,ARUPPUKOTTAI (OIL)','MONDAY','VIRUDHUNAGAR'],
  ['OM MURUGA AUTO STORE,  SRIVILLIPUTHUR (OIL)','FRIDAY','VIRUDHUNAGAR'],
  ['SRI ANANDHA MOTORS, DGL (OIL)','MONDAY','DINDIGUL'],
  ['PANDI MUNEESHWARAN AUTO WORKS, SRIVILLIPUTTUR (OIL)','FRIDAY','VIRUDHUNAGAR'],
  ['RUPAN AUTOMOBILES, SVK (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['SENTHILANDAVAR AUTOMOBILES, KOVILPATTI (OIL)','MONDAY','TUTICORIN'],
  ['SHREE BALAJI AUTO AGENCY, TUTICORIN (OIL)','WEDNESDAY','TUTICORIN'],
  ['SHREE RAM AUTOMOBILES, DHALAVAIPURAM (OIL)','THURSDAY','VIRUDHUNAGAR'],
  ['SRI  KANTHAPANDI AUTOMOBILES, DHALAVAIPURAM (OIL)','THURSDAY','VIRUDHUNAGAR'],
  ['SRI BALAJI AUTOMOBILES, VIRUDHUNAGAR (OIL)','TUESDAY','VIRUDHUNAGAR'],
  ['SRI CHAKRA MARUTHI , SVPTR (OIL)','FRIDAY','VIRUDHUNAGAR'],
  ['SRI GANAPATHY MOTORS,SIVAKASI (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['SRI GANESH LUBRICANTS, SIVAKASI (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['Sudheesan Auto  Store And Service, Mnm (Oil)','FRIDAY','SIVAGANGAI'],
  ['SRI IYYVARAJA AUTO STORE, DVPM (OIL)','THURSDAY','VIRUDHUNAGAR'],
  ['NEW STAR AUTO SPARES, SIVGANGAI (OIL)','FRIDAY','SIVAGANGAI'],
  ['SRI KUMAR AUTO WORKS, RAJAPALAYAM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['ABU AUTO SPARES, ILAYANKUDI (OIL)','FRIDAY','SIVAGANGAI'],
  ['SRI SAKTHI TRADERS, KRISHNAN KOVIL (OIL)','FRIDAY','VIRUDHUNAGAR'],
  ['SRI VALATHTHI AMMAN AUTOS, SIVAKASI (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['THIRU SENTHIL ANDAVAR AUTO SPARES, KOODANKULAM (OIL','WEDNESDAY','KANNIYAKUMARI'],
  ['TREND GREEN TRADERS, ARALVAIMOZHI (OIL)','THURSDAY','KANNIYAKUMARI'],
  ['VENNILA AUTO SPARE PARTS, KOVILPATTI (OIL)','MONDAY','TUTICORIN'],
  ['VAILANKANNI AUTO SPARES, VALLIOOR (OIL)','FRIDAY','TIRUNELVELI'],
  ['W.S.AUTOMOBILES, NEYYOR (OIL)','TUESDAY','KANNIYAKUMARI'],
  ['SHANMUGAM MOTOR WORKS, THENI (OIL)','MONDAY','THENI'],
  ['RMS AUTO STORE,  MADURAI (OIL)','WEDNESDAY','MADURAI'],
  ['SIDDHESH CARS, DGL (OIL)','MONDAY','DINDIGUL'],
  ['Murugan Automobiles, Mdu (Oil)','SATURDAY','MADURAI'],
  ['A.H.TWO WHEELER SPARES, ATTOOR (OIL)','TUESDAY','KANNIYAKUMARI'],
  ['VMAS ENTERPRISES,TMLM (OIL)','SATURDAY','MADURAI'],
  ['GNANAM AUTO TWO WHEELER SPARES, MDU (','TUESDAY','MADURAI'],
  ['SIVAKASI AUTO PARTS, THRUTHANGAL (OIL)','TUESDAY','VIRUDHUNAGAR'],
  ['SRI JOTHI LAKSHMI MOTOR WORKS, SVPR (OIL) **','FRIDAY','VIRUDHUNAGAR'],
  ['FATHIMA AUTO WORKS, RJPM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['AYNGARAN ENTERPRISES, NKL (OIL)','WEDNESDAY','KANNIYAKUMARI'],
  ['BALAJI AUTO STORE, CHEKKANURANI (OIL)','FRIDAY','MADURAI'],
  ['P.SUBBURAJ AUTOMOBILES, MELUR (OIL)','THURSDAY','MADURAI'],
  ['THEERAN SRI AUTOMOBILE, SUNDHARNACHIYARPURAM (OIL)','THURSDAY','VIRUDHUNAGAR'],
  ['PALANI ANDAVAR WORK SHOP, MDU (OIL)','SATURDAY','MADURAI'],
  ['DURAIPANDI  AUTO STORE, WATRAB (OIL)','FRIDAY','VIRUDHUNAGAR'],
  ['CLASSIC DRIVE,  KKRI (OIL)','WEDNESDAY','KANYAKUMARI'],
  ['MANJAMMAL MOTORS, MDU (OIL)','THURSDAY','MADURAI'],
  ['G.R.MOTORS, NAGERCOIL (OIL)','MONDAY','KANNIYAKUMARI'],
  ['SRI AYYANAR AUTO STORE, SIVAKASI (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['HARSHIK TWO WHEELER WORK SHOP, VLR (OIL)','FRIDAY','TIRUNELVELI'],
  ['R.K.AUTO WORKS, MDU (OIL)','MONDAY','MADURAI'],
  ['V.S.K AUTOMOBILE, SIVAKASI (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['KABILESH AUTO STORES, RJPM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['POORNADEVI  LUBRICANTS, MDU (OIL)','MONDAY','MADURAI'],
  ['S S MANI AUTO WORKS,DGL (OIL)','MONDAY','DINDIGUL'],
  ['KUMAR AUTO WORKS, VALLIOOR (OIL)','FRIDAY','TIRUNELVELI'],
  ['VELAVAN TWO WHEELER AUTO WORKS, SURANDAI (OIL)','WEDNESDAY','TENKASI'],
  ['Sasi Autos, Dgl (Oil)','MONDAY','DINDIGUL'],
  ['MATHA MARINES, KKRI (OIL)','WEDNESDAY','KANYAKUMARI'],
  ['ASM MOTORS, TIRUNELVELI (OIL)','MONDAY','TIRUNELVELI'],
  ['MOHAN MARUTI CAR CARE, NKL (OIL)','MONDAY','KANNIYAKUMARI'],
  ['SRI BALAJI LUBRICANTS ,MLR (OIL)','THURSDAY','MADURAI'],
  ['DHURUVAN AUTO STORE, TIRUPPATHUR (OIL)','TUESDAY','SIVAGANGAI'],
  ['Arul Auto Stores ,Kkdi (Oil)','TUESDAY','SIVAGANGAI'],
  ['UDHAYA AUTO STORE, MDU (OIL)','MONDAY','MADURAI'],
  ['SRI KARPAGA VINAYAGAR TRADERS,KKOIL','FRIDAY','VIRUDHUNAGAR'],
  ['DINESH AUTO STORES, TIRUPPACHETHI (OIL)','FRIDAY','SIVAGANGAI'],
  ['KAVITHA AUTO STORE, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['MOORTHY AUTO GARAGE, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['KAVIN AUTOMOBILES, DVKT (OIL)','TUESDAY','SIVAGANGAI'],
  ['MAHASAYEE AUTO PARTS, MDU (OIL)','FRIDAY','MADURAI'],
  ['Bikers Point, Mdu (Oil)','WEDNESDAY','MADURAI'],
  ['VALARMATHI AUTO SPARES, TMLM (OIL)','SATURDAY','MADURAI'],
  ['SRI NACHIAR MOTORS, RAJAPALAYAM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['RAJAPALAYAM AUTO STORE, RAJAPALAYAM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['SRI SIVASAKTHI TWO WHEELER SPARE, MNM (OIL)','FRIDAY','SIVAGANGAI'],
  ['MOTO BIKE POINT, MADURAI (OIL)','TUESDAY','MADURAI'],
  ['MARAN AUTO SPARE PARTS, MNM (OIL)','FRIDAY','SIVAGANGAI'],
  ['MOTORS WORLD, MDU (OIL)','MONDAY','MADURAI'],
  ['SRI SARGURU AUTO STORE, SVG (OIL)','FRIDAY','SIVAGANGAI'],
  ['SRI MEENAKSHI AUTO STORES,APt (C)','TUESDAY','THENI'],
  ['HARIPRASATH AUTO STORE ,CNMNR (OIL)','WEDNESDAY','THENI'],
  ['KARTHIKEYAN AUTO STORE,MDU (OIL)','MONDAY','MADURAI'],
  ['DINAKARAN AUTO STORES,KALINGA PATTI (OIL)','MONDAY','TENKASI'],
  ['PONNI AUTOS, MDU  (C)','MONDAY','MADURAI'],
  ['SARAVANA AUTOMOBILES, MDU (OIL)','TUESDAY','MADURAI'],
  ['THIRUMURUGA AUTOMOBILES, THIRUMONCHOLAI (OIL)','FRIDAY','SIVAGANGAI'],
  ['SRI MAHA AUTO WORKS, RAJAPALAYAM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['MIDHUL AUTOMOBILE, SATTUR (OIL)','TUESDAY','VIRUDHUNAGAR'],
  ['LAXMI AUTO PARTS, MDU (OIL)','MONDAY','MADURAI'],
  ['SRI RAJA AUTO SPARES,SVK (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['RKS AUTO STORES, SIVAKASI (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['RH MOTORS, MDM (OIL)','TUESDAY','KANNIYAKUMARI'],
  ['SRI SASTHA AUTO SPARES, TIRUPPUVANAM (OIL)','FRIDAY','SIVAGANGAI'],
  ['SRI GANAPATHY AUTOMOBILES, USILAI (OIL)','FRIDAY','MADURAI'],
  ['DINESH KANNA AUTO PARTS, MDU (OIL)','MONDAY','MADURAI'],
  ['SRI AADHI TRADERS, MADURAI (OIL)','MONDAY','MADURAI'],
  ['MUKILA AUTO STORES, MDU (OIL)','MONDAY','MADURAI'],
  ['J.K.VISHNU AUTO STORE, APT (OIL)','TUESDAY','THENI'],
  ['ESSARR AUTO SPARES, PUDUVAYAL (OIL)','TUESDAY','MADURAI'],
  ['ADS AUTO STORE,SVG (OIL)','FRIDAY','SIVAGANGAI'],
  ['SAI AUTO STORE,MELUR (OIL)','THURSDAY','MADURAI'],
  ['MUTHAIAH APE AUTO WORKS, USILAI (OIL)','FRIDAY','MADURAI'],
  ['AL-SAFA , MADURAI (OIL)','TUESDAY','MADURAI'],
  ['MAHARAJ AGENCIES,MDU (OIL)','MONDAY','MADURAI'],
  ['KALAI TWO WHEELER WORKS, URANGANPATTI (OIL)','THURSDAY','MADURAI'],
  ['SRI SENTHUR AUTOMOBILES, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['TKT ENTERPRISES, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['EASAN MOTORS,  SRIVILLIPUTTUR (OIL)','FRIDAY','VIRUDHUNAGAR'],
  ['SRI VINAYAGA AUTO STORES, THIRUTHANGAL (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['MINI MOTORS, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['SREE SENTHIL AUTO GARAGE, MDU (OIL)','TUESDAY','MADURAI'],
  ['KANI AUTO SPARES, SIVAKASI (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['SRI SUBBURAMANIYAN TWO WHEELER WORKS, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['ALPHA ENTERPRISES,RJPM (OIL)','','VIRUDHUNAGAR'],
  ['ARUL AUTO STORE, MELUR (OIL)','THURSDAY','MADURAI'],
  ['M.P. TRACTOR PARTS,SVK (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['ARUN MOTORS, BODI (OIL)','MONDAY','THENI'],
  ['KRISHNA AUTO WORKS RAJAPALAYAM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['ARUNACHAL MOTORS PVT LTD, MADURAI (OIL)','','MADURAI'],
  ['B2B BAZAR, MDU','','MADURAI'],
  ['A.M.S.AUTO STORES, SETTIYARPATTI (OIL)','THURSDAY','VIRUDHUNAGAR'],
  ['SRI RAM AUTO SPARES, SKT(OIL)','TUESDAY','TENKASI'],
  ['S.R.M.AUTO SPARE, ANJUKANNUKALUNGU (OIL)','TUESDAY','KANNIYAKUMARI'],
  ['JAGUAR AUTOTRONICS, MDU (OIL)','THURSDAY','MADURAI'],
  ['SAI CAR CARE, NAGERCOIL (OIL)','SATURDAY','KANNIYAKUMARI'],
  ['AMMAN AUTO SPARE, VEERACHOLAN (OIL)','','VIRUDHUNAGAR'],
  ['M.P.K.AUTO GARAIGE, MELUR (OIL)','THURSDAY','MADURAI'],
  ['EAGLE MOTORS, POLLACHI (CBE-OIL)','','COIMBATORE'],
  ['CHELLAM AUTO SPARES, PAVOORCHATHRAM (OIL)','TUESDAY','TENKASI'],
  ['S K AUTOMOBILES, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['A.M.AUTO SPARES, MDU (OIL)','TUESDAY','MADURAI'],
  ['GHAYATHRI MOTORS, SIVAKASI (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['EVEREST MOTORS, PYR (OIL)','SATURDAY','MADURAI'],
  ['BALU AUTO AGENCY, APK (OIL)','MONDAY','VIRUDHUNAGAR'],
  ['BHARATH AUTO PARTS,SVK','WEDNESDAY','VIRUDHUNAGAR'],
  ['DHANALAKSHMI AUTO ELECTRICALS,VNR','TUESDAY','VIRUDHUNAGAR'],
  ['SIVASAKTHI TRADERS,SANKARANKOVIL (OIL)','FRIDAY','TENKASI'],
  ['GUGAN AUTO AGENCIES, MADURAI (OIL)','','MADURAI'],
  ['DINAKARAN AUTO STORES,KOVILPATTI (C)','MONDAY','TUTICORIN'],
  ['GUSTO FLOW TECH ENGINEERING, THENGAPATTANAM (OIL)','','KANNIYAKUMARI'],
  ['GUSTO FLOW TECK ENGINEERING,MARTHANDAM (OIL)','FRIDAY','KANNIYAKUMARI'],
  ['HANISHA AUTOMOBILES, PALAMDEU','','MADURAI'],
  ['THUPPAKKI MOTORS, MDU (OIL)','THURSDAY','MADURAI'],
  ['CRD AUTOMOTIVE, TIRUNELVELI (OIL)','MONDAY','TIRUNELVELI'],
  ['EZHIL AUTO STORE, VIRUDHUNAGAR  (OIL)','TUESDAY','VIRUDHUNAGAR'],
  ['TAJ TWO WHEELER WORKS, MDU (OIL)','TUESDAY','MADURAI'],
  ['INDIAN AUTO MOBILES,PMK (OIL)','','RAMNATHAPURAM'],
  ['S.R.VINAYAGA AUTO GARAGE, MELUR (OIL)','THURSDAY','MADURAI'],
  ['SRI RAGAVENDRA AUTO STORES, T.KALLUPATTI (OIL)','SATURDAY','MADURAI'],
  ['JESUS AUTO WORKS, NAZARETH (OIL)','','TUTICORIN'],
  ['JOTHI MOTOR SPARES,MDU (R)','','MADURAI'],
  ['ABINESH AUTOMOBILES, SATTUR (OIL)','MONDAY','VIRUDHUNAGAR'],
  ['SRI MANIKUMAR AUTO SPARES, MDU (OIL)','SATURDAY','MADURAI'],
  ['TECH MAN. MDU (OIL)','WEDNESDAY','MADURAI'],
  ['PRAKASH OIL STORES, MDU (OIL)','MONDAY','MADURAI'],
  ['KATHIRAVAN AUTO STORES, MDU (OIL)','TUESDAY','MADURAI'],
  ['SRI AMBAL AUTO STORES, MDU (OIL)','MONDAY','MADURAI'],
  ['JEYA AUTO MOBILES, SKOIL','FRIDAY','TENKASI'],
  ['SELVAKUMAR AUTO WORKS,  KAVALKINARU(OIL)','FRIDAY','TIRUNELVELI'],
  ['ESAKKI MATHESH AGENCY, TNLR (OIL)','MONDAY','TIRUNELVELI'],
  ['GENERAL AUTOMOBILES, UDAIYARPATTI (OIL)','SATURDAY','TIRUNELVELI'],
  ['SRI KARPAGAM AUTOMOBILES,MDU (OIL)','TUESDAY','MADURAI'],
  ['GOPAL AUTO PARTS ,SVPR (OIL)','FRIDAY','VIRUDHUNAGAR'],
  ['J2E ENTERPRISES, MDU (OIL)','TUESDAY','MADURAI'],
  ['LOGESH AUTO SPARES, ELUMALAI (OIL)','FRIDAY','MADURAI'],
  ['KUMAR AUTOMOBILES, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['JAYA AUTO WORKS, PERAIYUR (OIL)','SATURDAY','MADURAI'],
  ['SHREE SAI TYRE & OIL, VAZHUKKAMPARAI (OIL)','THURSDAY','KANNIYAKUMARI'],
  ['SK MOTORS, MDU (OIL)','TUESDAY','MADURAI'],
  ['JAWAHAR AUTO WORKS, MDU (OIL)','THURSDAY','MADURAI'],
  ['J.J ALL CAR SERVICE, AGASTHIYARPATTI (OIL)','THURSDAY','TIRUNELVELI'],
  ['SRI ANGALAMMAN AUTO SPARE PARTS, ODC (OIL)','THURSDAY','DINDIGUL'],
  ['SRI VELMURUGAN TRACTOR WORKS, DPM (OIL)','TUESDAY','DINDIGUL'],
  ['RR AUTO SPARES, VADAMADURI (OIL)','FRIDAY','DINDIGUL'],
  ['THIRUMURUGAN TWO WHEELER WORKSHOP, MDU (OIL)','MONDAY','MADURAI'],
  ['SRI  AANUR AMMAN TRADERS, GOVINDHAPURAM (CBE-OIL)','FRIDAY','THIRUPUR'],
  ['MARIYA AUTO SPARE PARTS, BATLAGUNDU  (OIL)','TUESDAY','DINDIGUL'],
  ['PASUMAI AGRO SOLUTIONS,MDU','','MADURAI'],
  ['PONMANI AUTO CENTRE, MDU (OIL)','','MADURAI'],
  ['SURESH MOTORS, THENI (OIL)','MONDAY','THENI'],
  ['RAJ MOTORS, KELA SURANDAI (OIL)','WEDNESDAY','TIRUNELVLEI'],
  ['NILESHKUMAR AUTO SPARES, USILAMPATTI (OIL)','TUESDAY','MADURAI'],
  ['SRI EASWARI MOTORS, TNI(OIL)','MONDAY','THENI'],
  ['SRI MARUTI MOTORS,MDU (OIL)','TUESDAY','MADURAI'],
  ['RAMACHANDRAN, VILATHIKULAM (OIL)','','TUTICORIN'],
  ['J.V.S TRAILORS, KNLR (OIL)','WEDNESDAY','TENKASI'],
  ['SAADHANA MOTORS,. PALANI (OIL)','THURSDAY','DINDIGUL'],
  ['JESI CAR ZONE, TVL (OIL)','SATURDAY','TIRUNELVELI'],
  ['RS AUTOMOBILES,T.KALLUPATTI','SATURDAY','MADURAI'],
  ['SS M HARDWARE & BUILDING MATERIALS, THENI','','THENI'],
  ['JP AUTOMOBILES SPARES & SERVICE ,UDUMALAI (OIL)','FRIDAY','TIRUPPUR'],
  ['S.K.LUBRICANTS, BODI (OIL)','','THENI'],
  ['PANDIAN TIRUMANGALAM (OIL)','SATURDAY','MADURAI'],
  ['SREE GANAPATHY MOTORS, NAGERCOIL (OIL)','SATURDAY','KANNIYAKUMARI'],
  ['DEVA AUTO PARTS, MDU (OIL)','','MADURAI'],
  ['SRI KRISHNA AUTO SPARES, MANAMADURAI (OIL)','FRIDAY','SIVAGANGAI'],
  ['SREE TIRUPATHI MOTORS,  NAGERCOIL (OIL)','THURSDAY','KANNIYAKUMARI'],
  ['R S CAR CARE, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['K.R.SENTHIL AUTOMOBILES, KADAYANALLUR (OIL)','WEDNESDAY','TENKASI'],
  ['KRISHNA AUTOMOBILES,SURANDAI (OIL)','WEDNESDAY','TENKASI'],
  ['HAKEEM TRADERS, UPLYM (OIL)','WEDNESDAY','THENI'],
  ['MAGHI SPARES AND BATTERY, UTHUMALAI (OIL)','WEDNESDAY','TIRUNELVELI'],
  ['SARASWATHI AUTO STORES, TCN','','TUTICORIN'],
  ['SARASWATHI AUTO STORES, TCN (OIL)','','TUTICORIN'],
  ['MAHESH AUTO SPARES, AMBAI (OIL)','THURSDAY','TIRUNELVELI'],
  ['BANUMATHI TRADERS, KGLKDI','THURSDAY','MADURAI'],
  ['MARI AUTO WORKS & SPARE PARTS, VANNIKONENDAL (OIL)','FRIDAY','TIRUNELVELI'],
  ['MARIYAMMAL AUTOMOBILES, AGLM (OIL)','WEDNESDAY','TENKASI'],
  ['APR AUTO SPARES,GOPALPATTI (OIL)','SATURDAY','DINDIGUL'],
  ['SHANMUGA JOTHI AUTO STORES,PUDUR (OIL)','','TUTICORIN'],
  ['THENDRAL AUTO WORKS, VEDASNDUR (OIL)','WEDNESDAY','DINDIGUL'],
  ['MURUGAN AUTOMOBILES, KADAYAM (OIL)','THURSDAY','TENKASI'],
  ['PANDIYAN  AUTO CENTRE,  AVUDAYANOOR (OIL)','THURSDAY','TENKASI'],
  ['NAGA TWO WHEELER WORK\'S, RAJAPALAYAM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['DEIVAM TRACTORS TWO WHEELERS, UPLM (OIL)','WEDNESDAY','THENI'],
  ['N.S.S.AUTO AGENCY, DINDIGUL (OIL)','MONDAY','DINDIGUL'],
  ['SIVASAKTHI LUBRICANTS,MDU','','MADURAI'],
  ['RAJ WORKSHOP,VIKRAMASINGAPURAM (OIL)','THURSDAY','TIRUNELVELI'],
  ['SOUNDAR AUTO AGENCIES,MDU (R)','','MADURAI'],
  ['SREE ABIRAMI AUTO STORES,MDU','','MADURAI'],
  ['SREE ANNAMALAI ENTERPRISES,MDU (OIL)','','MADURAI'],
  ['IBRAHIM TRADERS, UTHAMAPALAYAM (OIL)','WEDNESDAY','THENI'],
  ['SHRI HARINI AUTO STORE, MADURAI (OIL)','MONDAY','MADURAI'],
  ['SREE SAMY AUTOMOBILES, NARIKUDI (OIL)','','VIRUDHUNAGAR'],
  ['RATHNESH AUTO WORKS, AGASTHIYARPATTI (OIL)','THURSDAY','TIRUNELVELI'],
  ['RAJARAM AUTOMOBILES,  RAJAPALAYAM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['SRI  MEENAKSHI AUTO STORES, PYR (OIL)','SATURDAY','MADURAI'],
  ['AMBAL AUTO STORE,MDU (OIL)','MONDAY','MADURAI'],
  ['SRI ALAGENDRA AUTOMOBILES P.LTD., MDU','','MADURAI'],
  ['SRI AMMAN AUTO MOBILES,MDU','','MADURAI'],
  ['ARAVIND AGRO ENTERPRISES, USPT (R)','FRIDAY','MADURAI'],
  ['FRIENDS AUTO STORES, MDU (OIL)','MONDAY','MADURAI'],
  ['A.V.M. AUTO PARTS, NATHAM','SATURDAY','DINDIGUL'],
  ['SRI BALAA AGENCIES, MADURAI. (R)','','MADURAI'],
  ['GNANAM AUTO SPARES, MADURAI (OIL)','TUESDAY','MADURAI'],
  ['J.K.TRADERS, APT (R)','TUESDAY','THENI'],
  ['SAKTHI VEL TWO WHEELER WORK SHOP, SURANDAI (OIL)','WEDNESDAY','TENKASI'],
  ['JOSH MOTORS,MDU','MONDAY','MADURAI'],
  ['OM SARAVANA AUTO SPARES,CUMBUM (OIL)','THURSDAY','THENI'],
  ['SAKTHI SPARES & OIL STORE, AUNDIPATTY (OIL)','TUESDAY','THENI'],
  ['KUMAR AUTO WORKS,MDU (OIL)','SATURDAY','MADURAI'],
  ['SIVANANTHA AUTO SPARES, VNR (OIL)','TUESDAY','VIRUDHUNAGAR'],
  ['SRI ABIRAMI MOTOR CYCLE COMPANY,KPT (OIL)','MONDAY','TUTICORIN'],
  ['SARAVANA TRACTORS, TNI (OIL)','MONDAY','THENI'],
  ['SR BIKE POINT, DGL (OIL)','MONDAY','DINDIGUL'],
  ['ANNAI BEARING CENTRE, VLR (OIL)','FRIDAY','TIRUNELVELI'],
  ['SRI AMMAN AUTO STORES, APT (OIL)','TUESDAY','THENI'],
  ['SRI BALAJI AUTO PARTS, THENI (OIL)','MONDAY','THENI'],
  ['LAKSHMI AUTOMOBILE, ELUMALAI','FRIDAY','MADURAI'],
  ['SRI BALAJI MOTOR WORKS, CUMBUM (OIL)','THURSDAY','THENI'],
  ['A.M.S.&SONS, GANDHIGRAM (OIL)','TUESDAY','DINDIGUL'],
  ['SRI NANDEES TRADERS,CNMR','','THENI'],
  ['MUKUNDA MOTORS, MDU','','MADURAI'],
  ['LEELAVATHI AUTO SPARES, MDU (OIL)','TUESDAY','MADURAI'],
  ['MARUTHAM AUTO PARTS, MADURAI (OIL)','TUESDAY','MADURAI'],
  ['SRI SAKTHI AUTO PARTS,MDU (R)','','MADURAI'],
  ['SRI LAKSHMI TRADING CO, SVPR (OIL)','FRIDAY','VIRUDHUNAGAR'],
  ['SRI DHANALAKSHMI AUTO STORES, CUMBUM (OIL)','THURSDAY','THENI'],
  ['MALA AUTO SPARE PARTS, DVK (C)','TUESDAY','SIVAGANGAI'],
  ['MUTHURAJA TRACTOR WORK SHOP,THIRUPPACHETHI','FRIDAY','SIVAGANGAI'],
  ['SRI SASWANTH AUTOMOBILES,MDU','','MADURAI'],
  ['NEW ASIAN MOTORS, MDU (OIL)','TUESDAY','MADURAI'],
  ['PRABHA TWO WHEELER WORKS, MADURAI (OIL)','WEDNESDAY','MADURAI'],
  ['NACHIYAR AUTO STORE','MONDAY','MADURAI'],
  ['SAKTHI AUTO SPARE PARTS, MDU (OIL)','TUESDAY','MADURAI'],
  ['PANDIAN AUTO SPARES, KKDI (OIL)','TUESDAY','SIVAGANGAI'],
  ['SRI NATARAJAN AUTOMOBILES, KOVILPATTI (OIL)','MONDAY','TUTICORIN'],
  ['MAHA AUTOMOBILES, TVL (OIL)','SATURDAY','TIRUNELVELI'],
  ['P.S.S.AUTO AGENCY, PCHM (R)','TUESDAY','TENKASI'],
  ['SRI EASWARI MOTORS, TNI (R)','MONDAY','THENI'],
  ['SRI VINAYAGA TWO WHEELER WORKS, PANTHALKUDI (OIL)','MONDAY','VIRUDHUNAGAR'],
  ['SHALINI AUTO WORKS, TIRUNELVELI (OIL)','MONDAY','TIRUNELVELI'],
  ['SRI VIVEK AUTO STORES,MDU','','MADURAI'],
  ['SRI YOGALAKSHMI AUTO STORE,MDU (C)','','MADURAI'],
  ['SRI KUMARAN GAS CAR MOTORS, CUMBUM (OIL)','THURSDAY','THENI'],
  ['SRI PARASAKTHI AUTO STORE, RAJAPALAYAM  (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['RAJ AUTO SPARE','MONDAY','MADURAI'],
  ['SRI KUMARAN MOTORS, UTHAMAPALAYAM (OIL)','THURSDAY','THENI'],
  ['SRI KRISHNA AUTO STORES, MDU (OIL)','WEDNESDAY','MADURAI'],
  ['SRI ALAGARSAMY AUTO STORES,TPNM (R)','FRIDAY','MADURAI'],
  ['TAMILNADU AUTO STORES,MDU','','MADURAI'],
  ['SRI RAGAVENDRA AUTO STORE, PERAIYUR (OIL)','SATURDAY','MADURAI'],
  ['SRI RAGA AUTOMOBILES,NAGERCOIL (OIL)','THURSDAY','KANNIYAKUMARI'],
  ['SRI SANNASI AUTO SPARE PARTS, MDU (OIL)','MONDAY','MADURAI'],
  ['SRI VINAYAGA AUTO WORKS & SPARES, MDU (OIL)','MONDAY','MADURAI'],
  ['THE TRUE SAI WORKS, KARUR  (TATA) - 3008834','','KARUR'],
  ['TESLA ELECTRIC SOLUTION, VADAKKANKULAM (OIL)','FRIDAY','TIRUNELVELI'],
  ['SUBASHRE AUTO SPARES, SEDAPATTI','SATURDAY','MADURAI'],
  ['SREE SAKTHIVEL PETRO PRODUCTS,DGL (OIL)','MONDAY','DINDIGUL'],
  ['TIMING AUTO SPARES, MDU','','MADURAI'],
  ['SURYA MOTORS, MDU (OIL)','FRIDAY','MADURAI'],
  ['THANESH AUTO SPARE PARTS, MARTHANDAM (OIL)','TUESDAY','KANNIYAKUMARI'],
  ['TRN AUTOMOBILES, KALAKAD (OIL)','FRIDAY','TIRUNELVELI'],
  ['TAMILNADU AUTO ENTERPRISES, MDU (OIL)','MONDAY','MADURAI'],
  ['THANGAM TWO WHEELER SPARES, USILAI (OIL)**','FRIDAY','MADURAI'],
  ['SREE GOMATHI AUTO MOBILES, CHENKULAM (OIL)','THURSDAY','TIRUNELVELI'],
  ['SREE SANKARAGOMATHY TRADERS, SKL (OIL)','','TENKASI'],
  ['THAVASINATHAN AUTOMOBILES, CHELLAMPATTI (OIL)','FRIDAY','MADURAI'],
  ['SHAJI AUTOMOBILES, TENKASI (OIL)','TUESDAY','TENKASI'],
  ['SRI KUMARAN TRACTOR WORKS, THENI (OIL)','MONDAY','THENI'],
  ['SRI ANAND AUTOMOBILES, AGASTHIYARPATTI (OIL)','THURSDAY','TIRUNELVELI'],
  ['SRI VENKATESWARA WORKS, TIRUNELVELI (OIL)','MONDAY','TIRUNELVELI'],
  ['SUN AUTO WORKS, RAJAPALAYAM (OIL)','SATURDAY','VIRUDHUNAGAR'],
  ['S.L.K AUTO WORKS, SVK (OIL)','WEDNESDAY','VIRUDHUNAGAR'],
  ['VASANTH & COMPANY, MDU (OIL)','MONDAY','MADURAI'],
  ['SRI MOTORS, THENI (OIL)','MONDAY','THENI'],
  ['VAAGAI PATHINATHAN AUTO WORKS, AMBAI','THURSDAY','TIRUNELVELI'],
  ['VELLAIYANGIRI NATHAN AUTO WORKS, KADAYAM (OIL)','THURSDAY','TENKASI'],
  ['SENTHILMURUGAN AUTO STORES,ODC (R)','TUESDAY','SIVAGANGAI'],
  ['THIRUPATHI AUTO&ELECT SPARES, KLGM (OIL)**','MONDAY','TUTICORIN'],
  ['CHANDRA AGENCY,MDU','','MADURAI'],
  ['ROYAL FORD,PARAMAKUDI (OIL)','','RAMANATHAPURAM'],
  ['THAAI AUTO STORE, TIRUPPUVANAM','FRIDAY','SIVAGANGAI'],
  ['SRI SANJAY AUTO CARRIAGE, THENI (OIL)','MONDAY','THENI'],
  ['THIRUMALAI AUTO STORE','THURSDAY','THENI'],
  ['VINOTH AUTOMOBILES,CNMR','WEDNESDAY','THENI'],
  ['BALAJI MOTORS,HOSUR (OIL)','','HOSUR'],
  ['SRI MADURAI MEENAKSHI AUTO STORES, MDU (OIL)','','MADURAI'],
  ['MANI AUTO WORKS, KOVILPATTI (OIL)','MONDAY','TUTICORIN'],
  ['THULASI AUTOMOBILES','MONDAY','DINDIGUL'],
  ['ANEES AUTOMAK, POLLACHI (CBE-OIL)','','COIMBATORE'],
  ['EVEREST AUTOMOBILES,KLGM','FRIDAY','VIRUDHUNAGAR'],
  ['SRI GANABATHI AUTOMOBILES,MDU (R)','','MADURAI'],
  ['F.K.AUTO, ILAIYANKUDI','','SIVAGANGAI'],
  ['OM SAIRAM AUTO TRADING, MDU (OIL)','','MADURAI'],
  ['SELVAM AUTO STORE, NILAKOTTAI (OIL)','TUESDAY','DINDIGUL'],
  ['SRI SENTHIL MURUGAN AUTO SPARES,PMK','','RAMANATHAPURAM'],
  ['S.K.AUTO STORES ,SVPR','FRIDAY','VIRUDHUNAGAR'],
  ['S.YASMIN OIL STORE,DEVATHANAPATTI','FRIDAY','THENI'],
  ['SHIVAN CAR SERVICE, THENI (OIL)','SATURDAY','THENI'],
  ['SRI IYAPPAN CAR CARE, BODI (OIL)','MONDAY','THENI'],
  ['A1 AUTO STORES, KALLIDAIKURICHI (OIL)','THURSDAY','TIRUNELVELI'],
  ['SRI JEYAM TRADERS, VEDASANDUR (OIL)','WEDNESDAY','DINDIGUL'],
  ['GOKULAM TRADERS ,KVPT (OIL)','','TUTICORIN'],
  ['IRAIVAN AUTO STORE, VEERAPANDI (OIL)','MONDAY','THENI'],
  ['LAYA SRI LUBRICANTS, DGL (OIL)','MONDAY','DINDIGUL'],
  ['PITSPOT AUTO CARE, SWAMIYARMADAM (OIL)','TUESDAY','KANYAKUMARI'],
  ['SIDDESH AUTO SPARES, DGL (OIL)','MONDAY','DINDIGUL'],
  ['SIVA AUTO AGENCIES PVT LTD, SKL (OIL)','FRIDAY','TENKASI'],
  ['LEO AUTO LUBRICANTS, VALLIOOR (OIL)','FRIDAY','TIRUNELVELI'],
  ['SRI MADURA AUTO SPARES,NATHAM (C)','SATURDAY','DINDIGUL'],
  ['St. MICHAEL AUTOMOBILES, VLR (OIL)','FRIDAY','TIRUNELVELI'],
  ['FRIENDS TWO WHEELER WORKS, DINDIGUL (OIL)','MONDAY','DINDIGUL'],
  ['SRI RAM MOTOR & TRACTORS, SVGR','','TENKASI'],
  ['VINAYAGA TRADERS, VNR (OIL)','TUESDAY','VIRUDHUNAGAR'],
  ['J K AUTOMOBILES, ARALVAIMOZHI (OIL)','FRIDAY','KANYAKUMARI'],
  ['PANDIAN TRACTOR WORK SHOP,APK','','VIRUDHUNAGAR']
];

  var sheet = getLiveSheet_(DEALER_META_SHEET, DEALER_META_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("DealerMeta has no data rows - nothing to update.");
    return;
  }
  var numCols = DEALER_META_HEADERS.length;
  var values = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var rowIndexByDealer = {};
  for (var i = 0; i < values.length; i++) {
    var name = String(values[i][0] || "").trim();
    if (name) rowIndexByDealer[name] = i;
  }

  var beatDayUpdated = 0, districtUpdated = 0, notFound = [];
  for (var j = 0; j < IMPORT_DATA_.length; j++) {
    var entry = IMPORT_DATA_[j];
    var dealer = entry[0], beatDay = entry[1], district = entry[2];
    var idx = rowIndexByDealer[dealer];
    if (idx === undefined) { notFound.push(dealer); continue; }
    if (beatDay) { values[idx][6] = beatDay; beatDayUpdated++; }
    if (district) { values[idx][4] = district; districtUpdated++; }
  }

  sheet.getRange(2, 1, values.length, numCols).setValues(values);

  Logger.log("Beat Day updated on " + beatDayUpdated + " dealers, District updated on " + districtUpdated + " dealers.");
  if (notFound.length > 0) {
    Logger.log(notFound.length + " dealer name(s) from the Excel did not match any existing DealerMeta row (likely a naming mismatch, or a dealer not yet on the live roster) - first 20: " + notFound.slice(0, 20).join(" | "));
  }
}

// POST { action: "dismissNewDealer", dsrName, name } - added 6 Oct 2026,
// see the "Dismissed" header comment on LIVE_NEW_DEALERS_HEADERS above for
// why this exists. dsrName must be the dealer's REAL source Tally group
// (e.g. "PETRONAS MDU AREA (SARAVANAN)" even when viewed from Chellamani's
// merged app screen), matching the Key column written by
// handleLiveNewDealerUpdate_ - the app side is responsible for sending the
// right one. No LIVE_SYNC_KEY check - this is a normal app/dashboard
// action, not a watcher push.
function handleDismissNewDealer_(body) {
  var dsrName = String(body.dsrName || "").trim();
  var name = String(body.name || "").trim();
  if (!dsrName || !name) return jsonOut_({ ok: false, error: "dsrName and name required" });
  var key = dsrName + "||" + name;
  var sheet = getLiveSheet_(LIVE_NEW_DEALERS_SHEET, LIVE_NEW_DEALERS_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonOut_({ ok: false, error: "Not found" });
  var keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0] || "").trim() === key) {
      sheet.getRange(i + 2, 6, 1, 1).setValues([["Y"]]);
      return jsonOut_({ ok: true });
    }
  }
  return jsonOut_({ ok: false, error: "Not found" });
}

// Reads MasterFileHistory for one monthKey into {dealerName: {salesLtr,
// collectionRs}}. isYearTotal rows (field "year") are only ever read
// by handleMasterFileReviewGet_'s own yearly-total display, not here.
function masterFileHistoryMonthMap_(monthKey) {
  var sheet = getLiveSheet_(MASTER_FILE_HISTORY_SHEET, MASTER_FILE_HISTORY_HEADERS);
  var rows = liveSheetRows_(sheet);
  var map = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    // Same normalizeMonthKey_() fix as handleDealerTargetsGet_() above (7
    // Oct 2026) - this sheet's "Month Key" column is written the same
    // plain "yyyy-MM" way via bulkUpsertSheet_() and is just as exposed to
    // Sheets' silent text-to-Date auto-conversion, which would otherwise
    // break every month's lookup here (feeds the "Suggested Target"
    // 6-month average shown on the same Dealer Targets screen).
    if (normalizeMonthKey_(r[1]) !== monthKey) continue;
    var dealer = String(r[0] || "").trim();
    if (!dealer) continue;
    map[dealer] = {
      salesLtr: r[2] === "" || r[2] === undefined ? null : Number(r[2]),
      collectionRs: r[3] === "" || r[3] === undefined ? null : Number(r[3])
    };
  }
  return map;
}

// "Collectable" - the portion of a dealer's outstanding overdue by
// MORE THAN 45 DAYS. Reuses LiveOutstanding's existing
// "Aged Pending (45d+)" column (agedPendingAmount) verbatim - see the
// big comment block above. {dealerName: amountRs}
function collectableByDealerMap_() {
  var rows = liveSheetRows_(getLiveSheet_(LIVE_OUTSTANDING_SHEET, LIVE_OUTSTANDING_HEADERS));
  var map = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var name = String(r[0] || "").trim();
    if (!name) continue;
    // Math.abs() added 6 Oct 2026 - LiveOutstanding's Balance/Aged Pending
    // columns come straight from Tally's ledger convention, where an
    // amount a dealer OWES can be stored as negative; the dashboard's own
    // "Needs Attention" table (totalOutstandingByDsr in dashboard/index.html)
    // already takes Math.abs() of the same kind of figure for exactly this
    // reason - this brings the Master File table in line with that, so
    // "Collectable" shows as a plain positive rupee amount owed, not a
    // confusing negative one.
    var val = r[4] === "" || r[4] === undefined ? null : Number(r[4]);
    map[name] = val === null ? null : Math.abs(val);
  }
  return map;
}

// "Overall Outstanding" as a MONTH-START SNAPSHOT for monthKey - the
// frozen LiveOutstanding archive one month back (a freeze taken on the
// last day of month M is month M+1's opening balance - see
// freezeLiveSheetForMonth_()/FREEZE_LIVE_CATEGORIES_ above). Falls
// back to the LIVE current balance (isSnapshot:false) when no freeze
// exists yet for that prior month (e.g. the very first month this
// feature runs, before any monthly freeze has ever been taken) -
// same "purely additive, graceful fallback" convention as everywhere
// else in this file.
function overallOutstandingSnapshot_(monthKey) {
  var prevMonthKey = addMonthsToMonthKey_(monthKey, -1);
  var frozenRows = readFrozenSheetForMonth_(LIVE_OUTSTANDING_SHEET, LIVE_OUTSTANDING_HEADERS, prevMonthKey);
  var map = {};
  var isSnapshot = false;
  // Math.abs() added 6 Oct 2026, same reasoning/same fix as
  // collectableByDealerMap_() just above - Tally's ledger Balance can be
  // stored negative for an amount owed, and the dashboard's own "Needs
  // Attention" table already Math.abs()'s this exact figure elsewhere.
  if (frozenRows && frozenRows.length > 0) {
    isSnapshot = true;
    for (var i = 0; i < frozenRows.length; i++) {
      var r = frozenRows[i];
      var name = String(r[0] || "").trim();
      if (!name) continue;
      var val1 = r[2] === "" || r[2] === undefined ? null : Number(r[2]);
      map[name] = val1 === null ? null : Math.abs(val1);
    }
  } else {
    var liveRows = liveSheetRows_(getLiveSheet_(LIVE_OUTSTANDING_SHEET, LIVE_OUTSTANDING_HEADERS));
    for (var j = 0; j < liveRows.length; j++) {
      var r2 = liveRows[j];
      var name2 = String(r2[0] || "").trim();
      if (!name2) continue;
      var val2 = r2[2] === "" || r2[2] === undefined ? null : Number(r2[2]);
      map[name2] = val2 === null ? null : Math.abs(val2);
    }
  }
  return { map: map, isSnapshot: isSnapshot, asOfMonthKey: prevMonthKey };
}

// Suggested Target - each dealer's own recent purchase pattern, not a
// generic guess (owner: "A dealer will buy 210 someone 55 so have to
// guess when they will buy next as per the last purchase"). Prefers
// REAL Tally sales (LiveSalesByDealer's Avg 6mo Sold Ltr + Last Sale
// Date - already live for the Dormant Dealers feature, so this gets
// the owner's "sharper once Tally live is wired" refinement for free)
// and only falls back to this DSR app's own punched-Orders 3-month
// average for a dealer that has no live Tally sales row yet.
// {dealerName: {suggestedLtr, basis: "tally6mo"|"ordersAvg3mo"|null,
//   lastSaleDate}}
function suggestedTargetMap_() {
  var map = {};
  var liveRows = liveSheetRows_(getLiveSheet_(LIVE_SALES_BY_DEALER_SHEET, LIVE_SALES_BY_DEALER_HEADERS));
  for (var i = 0; i < liveRows.length; i++) {
    var r = liveRows[i];
    var name = String(r[0] || "").trim();
    if (!name) continue;
    var avg6 = r[4] === "" || r[4] === undefined ? null : Number(r[4]);
    if (avg6 === null) continue;
    map[name] = {
      suggestedLtr: Math.round(avg6),
      basis: "tally6mo",
      lastSaleDate: normalizeDateKey_(r[3]) || null
    };
  }

  // Fallback for any dealer with no live Tally sales row yet - this
  // DSR app's own punched Orders, averaged over its last 3 distinct
  // months with at least one order (not necessarily calendar-
  // consecutive), same convention as every other "average" figure in
  // this file.
  var orderRows = readAllRows_(getSheet_());
  var byDealer = {};
  for (var j = 0; j < orderRows.length; j++) {
    var row = orderRows[j];
    if (row.deleted || !row.dealer || !row.monthKey) continue;
    if (!byDealer[row.dealer]) byDealer[row.dealer] = {};
    byDealer[row.dealer][row.monthKey] = (byDealer[row.dealer][row.monthKey] || 0) + Number(row.totalLtr || 0);
  }
  for (var dealer in byDealer) {
    if (map[dealer]) continue; // already has the sharper Tally-based figure
    var months = Object.keys(byDealer[dealer]).sort().reverse().slice(0, 3);
    if (!months.length) continue;
    var sum = 0;
    for (var m = 0; m < months.length; m++) sum += byDealer[dealer][months[m]];
    map[dealer] = { suggestedLtr: Math.round(sum / months.length), basis: "ordersAvg3mo", lastSaleDate: null };
  }
  return map;
}

// Latest remark per dealer (for the DSR app's own targets screen,
// which only ever needs to show the most recent one) - {dealerName:
// {remark, by, at}}
function latestRemarkByDealerMap_() {
  var rows = liveSheetRows_(getLiveSheet_(DEALER_REMARKS_SHEET, DEALER_REMARKS_HEADERS));
  var map = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var dealer = String(r[1] || "").trim();
    if (!dealer) continue;
    var at = String(r[4] || "");
    if (!map[dealer] || at > map[dealer].at) {
      map[dealer] = { remark: String(r[2] || ""), by: String(r[3] || ""), at: at };
    }
  }
  return map;
}

// GET ?type=dealerTargets&month=yyyy-MM[&dsrName=X] - the combined
// Dealer Targets row set: target/plan (editable), suggested target,
// collectable (45d+), overall outstanding (month-start snapshot),
// latest remark, beat day/area. Powers both the DSR app's own Dealer
// Targets screen (dsrName required there) and the dashboard's Master
// File table (dsrName omitted for the full roster, or passed to
// filter by DSR there too).
function handleDealerTargetsGet_(monthKey, dsrNameFilter) {
  if (!monthKey) return jsonOut_({ ok: false, error: "month required" });
  var meta = dealerMetaMap_();
  var targetsSheet = getLiveSheet_(DEALER_TARGETS_SHEET, DEALER_TARGETS_HEADERS);
  var targetRows = liveSheetRows_(targetsSheet);
  var targets = {};
  for (var i = 0; i < targetRows.length; i++) {
    var r = targetRows[i];
    // normalizeMonthKey_() - not a raw String(r[0]) compare - same fix as
    // readAllRows_()'s Month Key handling (2 Sep 2026): Sheets silently
    // auto-converts a plain "yyyy-MM" text cell like "2026-10" into a real
    // Date object on write, so a raw string compare here never matches and
    // every saved Dealer Target would read back as unset - exactly the "set
    // target disappears after Save" bug reported 7 Oct 2026. (Bug found and
    // fixed 7 Oct 2026.)
    if (normalizeMonthKey_(r[0]) !== monthKey) continue;
    var dealer = String(r[1] || "").trim();
    if (!dealer) continue;
    targets[dealer] = {
      salesTargetLtr: r[3] === "" || r[3] === undefined ? null : Number(r[3]),
      collectionPlanRs: r[4] === "" || r[4] === undefined ? null : Number(r[4]),
      editedBy: String(r[5] || ""),
      editedAt: String(r[6] || "")
    };
  }
  var suggested = suggestedTargetMap_();
  var collectable = collectableByDealerMap_();
  var outstandingSnap = overallOutstandingSnapshot_(monthKey);
  var remarks = latestRemarkByDealerMap_();
  // "Sales so far this month" per dealer, added 7 Oct 2026 for the DSR
  // app's new "Dealer Targets - Balance To Do" list (owner: "FOR WHOM
  // TARGET GIVEN THAT DEALER HAS TO BE SHOWN ... AS TARGET - SALES =
  // BALANCE TO DO", "IF YOU CAN PUT LIVE OTHERWISE APP ORDER IS
  // ENOUGH"). Live Tally per-dealer month sales already exists (the
  // same LiveSalesByDealer "Month Sold Ltr" column suggestedTargetMap_
  // above already reads for its own 6-month-average figure) - no new
  // Tally diagnostic needed. Falls back to this app's own Orders sheet,
  // summed for the current month, for any dealer with no live Tally row
  // yet - same fallback convention/field (row.totalLtr) already used by
  // suggestedTargetMap_'s own "ordersAvg3mo" fallback above.
  var liveSalesByDealer = parseSalesByDealerRows_(liveSheetRows_(getLiveSheet_(LIVE_SALES_BY_DEALER_SHEET, LIVE_SALES_BY_DEALER_HEADERS)));
  var appOrdersThisMonth = {};
  var orderRowsForMonth = readAllRows_(getSheet_());
  for (var oi = 0; oi < orderRowsForMonth.length; oi++) {
    var orow = orderRowsForMonth[oi];
    if (orow.deleted || !orow.dealer || orow.monthKey !== monthKey) continue;
    appOrdersThisMonth[orow.dealer] = (appOrdersThisMonth[orow.dealer] || 0) + Number(orow.totalLtr || 0);
  }
  // "Collected so far this month" per dealer, added 7 Oct 2026 - same
  // Balance To Do treatment as Sales above, just for the Collection Plan
  // (Rs) side (owner: "IS COLLECTION TARGET-DONE=BALANCE TO DO ALSO
  // INCLUDED?", "YES THATS ONLY I SAID ALREADY"). Live per-dealer
  // Collection already exists (LiveCollectionByDealer "Month Collected"
  // - the same sheet the Dealer Lookup card's own Collection tile
  // already reads) - no new Tally diagnostic needed here either. No app-
  // side fallback exists for this one (DSRs don't record collections in
  // this app, only Tally does, via Receipt vouchers) - a dealer with a
  // Collection Plan set but no live Collection row yet just shows null/0
  // Collected until the next Tally sync.
  var liveCollectionByDealer = parseCollectionByDealerRows_(liveSheetRows_(getLiveSheet_(LIVE_COLLECTION_BY_DEALER_SHEET, LIVE_COLLECTION_BY_DEALER_HEADERS)));

  var out = [];
  for (var dealerName in meta) {
    var m = meta[dealerName];
    if (dsrNameFilter && m.dsrName !== dsrNameFilter) continue;
    var t = targets[dealerName] || {};
    var s = suggested[dealerName] || {};
    var salesLtr = null, salesBasis = null;
    if (liveSalesByDealer[dealerName]) {
      salesLtr = liveSalesByDealer[dealerName].salesLtr;
      salesBasis = "tally";
    } else if (Object.prototype.hasOwnProperty.call(appOrdersThisMonth, dealerName)) {
      salesLtr = appOrdersThisMonth[dealerName];
      salesBasis = "appOrders";
    }
    var collectedRs = liveCollectionByDealer[dealerName] ? liveCollectionByDealer[dealerName].collectionRs : null;
    out.push({
      dealer: dealerName,
      dsrName: m.dsrName,
      area: m.area,
      district: m.district,
      beatDay: m.beatDay,
      salesTargetLtr: (t.salesTargetLtr === undefined) ? null : t.salesTargetLtr,
      collectionPlanRs: (t.collectionPlanRs === undefined) ? null : t.collectionPlanRs,
      editedBy: t.editedBy || "",
      editedAt: t.editedAt || "",
      suggestedLtr: (s.suggestedLtr === undefined) ? null : s.suggestedLtr,
      suggestedBasis: s.basis || null,
      lastSaleDate: s.lastSaleDate || null,
      salesLtr: salesLtr,
      salesBasis: salesBasis,
      collectedRs: collectedRs,
      collectableRs: (collectable[dealerName] === undefined) ? null : collectable[dealerName],
      overallOutstandingRs: (outstandingSnap.map[dealerName] === undefined) ? null : outstandingSnap.map[dealerName],
      remark: remarks[dealerName] ? remarks[dealerName].remark : "",
      remarkBy: remarks[dealerName] ? remarks[dealerName].by : "",
      remarkAt: remarks[dealerName] ? remarks[dealerName].at : ""
    });
  }
  return jsonOut_({
    ok: true,
    month: monthKey,
    outstandingIsSnapshot: outstandingSnap.isSnapshot,
    outstandingAsOfMonthKey: outstandingSnap.asOfMonthKey,
    dealers: out
  });
}

// POST { action: "dealerTargetSet", month, dealer, salesTargetLtr,
//   collectionPlanRs, editedBy } - a DSR filling their own dealers'
// targets, or Owner/Manager editing any dealer's from the dashboard
// (owner: "Dsr to fill if any changes need we can edit"). editedBy is
// the editor's own display NAME as returned by handleLogin_ (e.g.
// "Alagu"), never a role label - see the project's naming preference.
// Upserts by (Month Key, Dealer Name); either numeric field may be
// omitted/null to leave it unset without touching the other.
function handleDealerTargetSet_(body) {
  var monthKey = String(body.month || "").trim();
  var dealer = String(body.dealer || "").trim();
  if (!monthKey || !dealer) return jsonOut_({ ok: false, error: "month and dealer required" });
  var meta = dealerMetaMap_();
  var dsrName = (meta[dealer] && meta[dealer].dsrName) || String(body.dsrName || "").trim();
  var sheet = getLiveSheet_(DEALER_TARGETS_SHEET, DEALER_TARGETS_HEADERS);
  var key = monthKey + "||" + dealer;
  var updated = bulkUpsertSheet_(sheet, DEALER_TARGETS_HEADERS.length, [{
    monthKey: monthKey, dealer: dealer, dsrName: dsrName,
    salesTargetLtr: body.salesTargetLtr, collectionPlanRs: body.collectionPlanRs,
    editedBy: String(body.editedBy || "").trim()
  }],
    function () { return key; },
    function (item, now) {
      return [
        item.monthKey, item.dealer, item.dsrName,
        (item.salesTargetLtr === null || item.salesTargetLtr === undefined || item.salesTargetLtr === "") ? "" : Number(item.salesTargetLtr),
        (item.collectionPlanRs === null || item.collectionPlanRs === undefined || item.collectionPlanRs === "") ? "" : Number(item.collectionPlanRs),
        item.editedBy, now
      ];
    },
    // normalizeMonthKey_() - not a raw String(existingRow[0]) compare - same
    // fix as handleDealerTargetsGet_'s Month Key handling just below. Sheets
    // silently auto-converts a plain "yyyy-MM" text cell like "2026-10" into
    // a real Date object on write, so a raw string compare here NEVER
    // matched an existing row once that row's Month Key had been
    // auto-converted - every re-save of an already-saved dealer+month
    // silently appended a brand new row instead of updating the existing
    // one in place. That is the exact cause of the duplicate DealerTargets
    // rows found 8 Oct 2026 (138 rows for only 39 real month+dealer
    // combinations - see cleanupDuplicateDealerTargets below for the
    // one-time fix for rows already written before this change).
    function (existingRow) { return normalizeMonthKey_(existingRow[0]) + "||" + String(existingRow[1] || "").trim(); }
  );
  return jsonOut_({ ok: true, updated: updated });
}

// One-time cleanup for the duplicate DealerTargets rows written before the
// normalizeMonthKey_ fix above (8 Oct 2026). Collapses every group of rows
// that share the same (normalized Month Key, Dealer Name) down to a single
// row, keeping the most recently edited values (highest "Edited At") so no
// DSR-entered data is lost - just the accidental duplicates are removed.
// HOW TO RUN: open this project in the Apps Script editor, select
// cleanupDuplicateDealerTargets in the function dropdown, click Run, then
// check the execution log for the summary. Safe to run more than once (a
// second run will find zero duplicates left and do nothing).
function cleanupDuplicateDealerTargets() {
  var sheet = getLiveSheet_(DEALER_TARGETS_SHEET, DEALER_TARGETS_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No data rows - nothing to clean up.");
    return;
  }
  var numCols = DEALER_TARGETS_HEADERS.length;
  var values = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var bestByKey = {}; // key -> { row, editedAt }
  var order = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var key = normalizeMonthKey_(row[0]) + "||" + String(row[1] || "").trim();
    if (!key || key === "||") continue;
    var editedAt = String(row[6] || "");
    var existing = bestByKey[key];
    if (!existing) {
      bestByKey[key] = { row: row, editedAt: editedAt };
      order.push(key);
    } else if (editedAt > existing.editedAt) {
      // Later "Edited At" wins; ties (e.g. blank on both) keep the first
      // (earliest-row) one, same effective behavior bulkUpsertSheet_ always
      // had going forward.
      bestByKey[key] = { row: row, editedAt: editedAt };
    }
  }
  var cleaned = order.map(function (k) { return bestByKey[k].row; });
  // Write the Month Key cell back out as plain "yyyy-MM" text - Sheets may
  // well auto-convert it straight back to a Date the instant it's written
  // (the same behavior that caused the duplicates in the first place), and
  // that is fine: normalizeMonthKey_() is what both the read path
  // (handleDealerTargetsGet_) and the write path (handleDealerTargetSet_'s
  // existingKeyFn, fixed above) use everywhere now, so it no longer
  // matters whether a given cell ends up stored as text or as a Date.
  for (var j = 0; j < cleaned.length; j++) {
    cleaned[j][0] = normalizeMonthKey_(cleaned[j][0]);
  }
  var before = values.length;
  var after = cleaned.length;
  // Clear the old range first (it may have been longer than the cleaned
  // data) then write the deduped rows back starting at row 2.
  sheet.getRange(2, 1, before, numCols).clearContent();
  if (after > 0) {
    sheet.getRange(2, 1, after, numCols).setValues(cleaned);
  }
  Logger.log("DealerTargets cleanup: " + before + " rows -> " + after + " rows (" + (before - after) + " duplicates removed).");
}

// GET ?type=dealerRemarks&dealer=X - full remark history for one
// dealer (dashboard's Master File "Remarks" column reads this for its
// history popup; the DSR app itself only needs the latest one, already
// included inline in ?type=dealerTargets above).
function handleDealerRemarksGet_(dealer) {
  if (!dealer) return jsonOut_({ ok: false, error: "dealer required" });
  var rows = liveSheetRows_(getLiveSheet_(DEALER_REMARKS_SHEET, DEALER_REMARKS_HEADERS));
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[1] || "").trim() !== dealer) continue;
    out.push({ id: String(r[0] || ""), dealer: dealer, remark: String(r[2] || ""), by: String(r[3] || ""), at: String(r[4] || "") });
  }
  out.sort(function (a, b) { return String(b.at).localeCompare(String(a.at)); });
  return jsonOut_({ ok: true, remarks: out });
}

// POST { action: "dealerRemarkAdd", dealer, remark, by } - Owner/
// Manager leaving a note against a dealer (e.g. from a Review-by-DSR
// session), shown in that dealer's own DSR's app (owner: "that
// remarks should be shown in dsr app"). "by" is the author's own
// display name (e.g. "Alagu"), not a role label.
function handleDealerRemarkAdd_(body) {
  var dealer = String(body.dealer || "").trim();
  var remark = String(body.remark || "").trim();
  if (!dealer || !remark) return jsonOut_({ ok: false, error: "dealer and remark required" });
  var sheet = getLiveSheet_(DEALER_REMARKS_SHEET, DEALER_REMARKS_HEADERS);
  var id = Utilities.getUuid();
  var now = new Date().toISOString();
  sheet.appendRow([id, dealer, remark, String(body.by || "").trim(), now]);
  return jsonOut_({ ok: true, id: id, at: now });
}

// ----- Bank Deposit -----

// Bank Deposit's Date/Receipt From Date/Receipt To Date columns get
// written as plain "yyyy-MM-dd" strings, but Google Sheets auto-detects
// those as real dates and silently stores them as a Date value anyway -
// so getValues() can hand back either a Date object or a string
// depending on how that particular cell happened to get written/edited.
// Added 23 Oct 2026 after the owner reported dates showing as "Fri Oct
// 23 2026 00:00:00 GMT+0530 (India Standard Time)" in the app - that was
// a Date object going straight through a naive String(), and the exact
// same mismatch was silently breaking the "is this today?" check inside
// handleBankDepositUpdate_/handleBankDepositDelete_ below (a Date object
// stringified that way never equals a plain "yyyy-MM-dd" today string,
// so edit/delete could be rejecting entries that actually were today's).
// This formats either shape back to one consistent "yyyy-MM-dd".
function bdDateStr_(v) {
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, DISPLAY_TIMEZONE, "yyyy-MM-dd");
  }
  return String(v || "").trim();
}

// GET ?type=bankDeposits&dsrName=X - last 5 bank-deposit entries for one
// DSR, newest first (by Date, then by Created At for same-day entries).
// Each DSR's own app always passes its own dsrName. The dashboard's Bank
// Deposits panel (DSR-picker driven, same convention as the heatmap's own
// "Total (all DSRs)" default) passes dsrName when a specific DSR is
// picked, or omits it for its default "All DSRs" view - an empty/missing
// dsrName returns the last 5 across every DSR instead of requiring one.
//
// fromDate/toDate added 27 Oct 2026 ("need date selecting options in Bank
// Deposits for checking previous deposits") - optional "yyyy-MM-dd"
// strings. When either is given, this returns EVERY matching entry in
// that range (not capped at 5) instead of the last-5 default - a
// separate "search previous deposits" mode layered on the same endpoint
// rather than a new route, since the filtering (dsrName, soft-delete,
// sort) is otherwise identical. Date comparison is plain string
// comparison, safe because bdDateStr_() always normalizes to "yyyy-MM-dd"
// first, which sorts the same as real date order.
// latestOnly (added 27 Oct 2026, "compulsory show all dsr name and
// each dsr only 1 last entry in the screen"): the dashboard's default
// "All DSRs"/single-DSR view now asks for this explicitly so its
// default list is capped at ONE (the most recent) entry PER DSR -
// never 5. The 8 DSR apps' own "last 5" lists never pass this flag, so
// their unrelated default behavior (still their own last 5) is
// completely unchanged - this only ever narrows an already-dsrName-
// filtered or unfiltered result further, never widens one.
function handleBankDepositsGet_(dsrName, fromDate, toDate, latestOnly) {
  dsrName = String(dsrName || "").trim();
  fromDate = String(fromDate || "").trim();
  toDate = String(toDate || "").trim();
  var rows = liveSheetRows_(getLiveSheet_(BANK_DEPOSIT_SHEET, BANK_DEPOSIT_HEADERS));
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[12] || "").trim()) continue; // soft-deleted - see handleBankDepositDelete_()
    var rowDsrName = String(r[1] || "").trim();
    if (dsrName && rowDsrName !== dsrName) continue;
    var rowDate = bdDateStr_(r[2]);
    if (fromDate && rowDate < fromDate) continue;
    if (toDate && rowDate > toDate) continue;
    out.push({
      id: String(r[0] || ""),
      dsrName: rowDsrName,
      date: rowDate,
      amount: Number(r[3]) || 0,
      paidThrough: String(r[4] || ""),
      throName: String(r[5] || ""),
      receiptFrom: String(r[6] || ""),
      receiptTo: String(r[7] || ""),
      at: String(r[8] || ""),
      remarks: String(r[9] || ""),
      receiptFromDate: bdDateStr_(r[10]),
      receiptToDate: bdDateStr_(r[11])
    });
  }
  out.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)) || String(b.at).localeCompare(String(a.at)); });
  var hasRange = !!(fromDate || toDate);
  if (hasRange) {
    return jsonOut_({ ok: true, deposits: out, range: true });
  }
  if (latestOnly) {
    // out is already sorted newest-first, so the first row seen per
    // dsrName is that DSR's own latest - works the same whether
    // dsrName was already filtered to one DSR above (yields exactly 1
    // row) or left blank for "all DSRs" (yields each DSR's own latest).
    var seenDsr = {};
    var latestEach = [];
    for (var j = 0; j < out.length; j++) {
      var dn = out[j].dsrName;
      if (seenDsr[dn]) continue;
      seenDsr[dn] = true;
      latestEach.push(out[j]);
    }
    return jsonOut_({ ok: true, deposits: latestEach, range: false });
  }
  return jsonOut_({ ok: true, deposits: out.slice(0, 5), range: false });
}

// POST { action: "bankDepositAdd", dsrName, date, amount, paidThrough,
// throName, receiptFrom, receiptFromDate, receiptTo, receiptToDate,
// remarks } - a DSR logging a cash deposit/handover. dsrName/date/
// amount/paidThrough are always required. "remarks" is required
// specifically when paidThrough is "Bank — Transfer" (added 23 Oct
// 2026, "add remarks in bank transfer ... thats also must") - a
// transfer has no paper receipt number, so Remarks is this mode's own
// required identifying field instead. For every OTHER paidThrough,
// receiptFrom/receiptFromDate/receiptTo/receiptToDate are now all
// required together (added same day, "remove the optional note from
// recpt and make it as must * ... need ask recpt from and to date near
// that recp number", then corrected to two SEPARATE dates - "I NEED 2
// DATES NEAR THE RECPT FROM AND TO" - since a receipt range can span
// several days of actual receipts) - validated here too (not just
// client-side) so a direct/scripted POST can't skip either rule.
function handleBankDepositAdd_(body) {
  var dsrName = String(body.dsrName || "").trim();
  var date = String(body.date || "").trim();
  var amount = Number(body.amount) || 0;
  var paidThrough = String(body.paidThrough || "").trim();
  var remarks = String(body.remarks || "").trim();
  var receiptFrom = String(body.receiptFrom || "").trim();
  var receiptFromDate = String(body.receiptFromDate || "").trim();
  var receiptTo = String(body.receiptTo || "").trim();
  var receiptToDate = String(body.receiptToDate || "").trim();
  if (!dsrName || !date || !amount || !paidThrough) {
    return jsonOut_({ ok: false, error: "dsrName, date, amount and paidThrough are required" });
  }
  if (paidThrough === "Bank — Transfer") {
    if (!remarks) return jsonOut_({ ok: false, error: "remarks is required for Bank — Transfer" });
  } else {
    if (!receiptFrom || !receiptFromDate || !receiptTo || !receiptToDate) {
      return jsonOut_({ ok: false, error: "receiptFrom, receiptFromDate, receiptTo and receiptToDate are required" });
    }
  }
  var sheet = getLiveSheet_(BANK_DEPOSIT_SHEET, BANK_DEPOSIT_HEADERS);
  var id = Utilities.getUuid();
  var now = new Date().toISOString();
  sheet.appendRow([
    id, dsrName, date, amount, paidThrough,
    String(body.throName || "").trim(),
    paidThrough === "Bank — Transfer" ? "" : receiptFrom,
    paidThrough === "Bank — Transfer" ? "" : receiptTo,
    now,
    // Remarks is stored for EVERY paidThrough now, not just Bank —
    // Transfer - changed 27 Oct 2026 ("need remarks coloumn for all type
    // of payments"). It's still required only for Bank — Transfer (the
    // validation above is unchanged); for every other mode it's just an
    // optional note the DSR may or may not fill in.
    remarks,
    paidThrough === "Bank — Transfer" ? "" : receiptFromDate,
    paidThrough === "Bank — Transfer" ? "" : receiptToDate,
    "" // Deleted At
  ]);
  return jsonOut_({ ok: true, id: id, at: now });
}

// POST { action: "bankDepositUpdate", id, dsrName, date, amount,
// paidThrough, throName, receiptFrom, receiptFromDate, receiptTo,
// receiptToDate, remarks } - added 23 Oct 2026 ("NEED A DELETE OR EDIT
// OPTION IN ALL DSR APP, THAT EDIT SHOULD BE DONE ONLY FOR TODAY DATE,
// NOT FOR ANY PRVIOUS DATE ENTRY"). Only the DSR who owns the entry can
// edit it (dsrName must match the row's own DSR Name), and only while
// the row's OWN Date - its value on the server right now, not whatever
// the edit form is submitting - is still today in DISPLAY_TIMEZONE; the
// submitted date is also rejected if it tries to move the entry away
// from today. Both checks are enforced here as a server-side backstop
// even though the app's own UI already only shows an Edit button on
// today's entries (the same way the existing "update" order action
// above relies purely on its UI gate) - the owner specifically called
// out the today-only restriction for this feature, so it's enforced on
// both sides.
function handleBankDepositUpdate_(body) {
  var id = String(body.id || "").trim();
  var dsrName = String(body.dsrName || "").trim();
  if (!id || !dsrName) return jsonOut_({ ok: false, error: "id and dsrName are required" });
  var sheet = getLiveSheet_(BANK_DEPOSIT_SHEET, BANK_DEPOSIT_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonOut_({ ok: false, error: "Entry not found" });
  var rows = sheet.getRange(2, 1, lastRow - 1, BANK_DEPOSIT_HEADERS.length).getValues();
  var today = Utilities.formatDate(new Date(), DISPLAY_TIMEZONE, "yyyy-MM-dd");
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0] || "").trim() !== id) continue;
    if (String(rows[i][12] || "").trim()) return jsonOut_({ ok: false, error: "Entry not found" });
    if (String(rows[i][1] || "").trim() !== dsrName) {
      return jsonOut_({ ok: false, error: "You can only edit your own entries." });
    }
    if (bdDateStr_(rows[i][2]) !== today) {
      return jsonOut_({ ok: false, error: "Only today's entries can be edited." });
    }
    var date = String(body.date || "").trim();
    var amount = Number(body.amount) || 0;
    var paidThrough = String(body.paidThrough || "").trim();
    var remarks = String(body.remarks || "").trim();
    var receiptFrom = String(body.receiptFrom || "").trim();
    var receiptFromDate = String(body.receiptFromDate || "").trim();
    var receiptTo = String(body.receiptTo || "").trim();
    var receiptToDate = String(body.receiptToDate || "").trim();
    if (!date || !amount || !paidThrough) {
      return jsonOut_({ ok: false, error: "date, amount and paidThrough are required" });
    }
    if (date !== today) {
      return jsonOut_({ ok: false, error: "The date can't be changed away from today." });
    }
    if (paidThrough === "Bank — Transfer") {
      if (!remarks) return jsonOut_({ ok: false, error: "remarks is required for Bank — Transfer" });
    } else {
      if (!receiptFrom || !receiptFromDate || !receiptTo || !receiptToDate) {
        return jsonOut_({ ok: false, error: "receiptFrom, receiptFromDate, receiptTo and receiptToDate are required" });
      }
    }
    sheet.getRange(i + 2, 1, 1, 12).setValues([[
      id, dsrName, date, amount, paidThrough,
      String(body.throName || "").trim(),
      paidThrough === "Bank — Transfer" ? "" : receiptFrom,
      paidThrough === "Bank — Transfer" ? "" : receiptTo,
      rows[i][8], // Created At - unchanged
      // Remarks stored for every paidThrough now - see the matching
      // comment in handleBankDepositAdd_() above, 27 Oct 2026.
      remarks,
      paidThrough === "Bank — Transfer" ? "" : receiptFromDate,
      paidThrough === "Bank — Transfer" ? "" : receiptToDate
    ]]);
    return jsonOut_({ ok: true, id: id });
  }
  return jsonOut_({ ok: false, error: "Entry not found" });
}

// POST { action: "bankDepositDelete", id, dsrName } - added 23 Oct 2026,
// same request as Update above and the same ownership/today-only checks.
// Soft delete only (stamps Deleted At, column 13) - mirrors the existing
// order "delete" action's soft-delete pattern (see the "delete" branch
// in doPost below) rather than removing the row, so a concurrent
// request's row index never shifts underneath it and the record
// survives permanently on the Sheet.
function handleBankDepositDelete_(body) {
  var id = String(body.id || "").trim();
  var dsrName = String(body.dsrName || "").trim();
  if (!id || !dsrName) return jsonOut_({ ok: false, error: "id and dsrName are required" });
  var sheet = getLiveSheet_(BANK_DEPOSIT_SHEET, BANK_DEPOSIT_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonOut_({ ok: false, error: "Entry not found" });
  var rows = sheet.getRange(2, 1, lastRow - 1, BANK_DEPOSIT_HEADERS.length).getValues();
  var today = Utilities.formatDate(new Date(), DISPLAY_TIMEZONE, "yyyy-MM-dd");
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0] || "").trim() !== id) continue;
    if (String(rows[i][12] || "").trim()) return jsonOut_({ ok: false, error: "Entry not found" });
    if (String(rows[i][1] || "").trim() !== dsrName) {
      return jsonOut_({ ok: false, error: "You can only delete your own entries." });
    }
    if (bdDateStr_(rows[i][2]) !== today) {
      return jsonOut_({ ok: false, error: "Only today's entries can be deleted." });
    }
    sheet.getRange(i + 2, 13, 1, 1).setValue(new Date().toISOString());
    return jsonOut_({ ok: true, id: id });
  }
  return jsonOut_({ ok: false, error: "Entry not found" });
}

// ----- Task Manager -----

// True when viewerRole/viewerName/allowedAreasCsv should see this raw
// Tasks row. Owner/Manager see everything. An Area Manager sees a task
// whose Area is one of their own AllowedAreas, or where they're
// explicitly named in Also Visible To. A DSR sees a task Assigned To
// them, or where they're named in Also Visible To.
function taskRowVisible_(row, viewerRole, viewerName, allowedAreasCsv) {
  if (viewerRole === "Owner" || viewerRole === "Manager") return true;
  var area = String(row[8] || "").trim();
  var assignedTo = String(row[6] || "").trim();
  var alsoVisible = String(row[7] || "").split(",").map(function (s) { return s.trim(); });
  if (viewerRole === "Area Manager") {
    var allowed = String(allowedAreasCsv || "").split(",").map(function (s) { return s.trim(); }).filter(Boolean);
    if (allowed.length === 0) return true; // unrestricted Area Manager
    if (area && allowed.indexOf(area) !== -1) return true;
    return alsoVisible.indexOf(viewerName) !== -1;
  }
  // DSR (or any other viewer): assigned-to or explicitly also-visible only.
  if (assignedTo && assignedTo === viewerName) return true;
  return alsoVisible.indexOf(viewerName) !== -1;
}

// Fixes a real display bug found 8 Oct 2026 from the owner's own
// screenshot ("Due Thu Oct 08 2026 00:00:00 GMT+0530 (India Standard
// Time) ... Reminder Sat Dec 30 1899 11:00:00 GMT+0521"). Root cause:
// Google Sheets auto-detects a cell's type from what's WRITTEN into it,
// regardless of how handleTaskCreate_ sends the value - a Due Date cell
// gets "2026-10-08" (looks like a date to Sheets) and a Reminder Time
// cell gets "11:00 AM" (looks like a time to Sheets), so Sheets silently
// converts both columns to Date/Time-typed cells. Apps Script then reads
// them back as real JS Date objects, not the plain strings that were
// sent - and a Date value run through String(...) (as both fields used
// to be here) calls the Date's own toString(), producing that whole
// ugly "Thu Oct 08 2026 00:00:00 GMT+0530 (India Standard Time)" line.
// For a genuinely blank Reminder Time, Sheets' own null-date epoch (Dec
// 30, 1899) is what showed through - nothing was ever actually wrong
// with the stored reminder, the cell was just blank AND Date-typed.
// formatTaskDate_/formatTaskReminderTime_ below defensively re-format a
// Date-typed cell back into the clean string the rest of this file
// always expected (a plain "yyyy-MM-dd" for Due Date, so every existing
// string-compare against todayIso elsewhere keeps working unchanged; the
// original "11:00 AM"/"3:00 PM"/"4:30 PM" slot label for Reminder Time,
// recovered from the cell's time-of-day and ignoring its meaningless
// Dec-30-1899 date part). A cell that was never auto-converted (still a
// plain string) passes through untouched either way.
function formatTaskDate_(v) {
  if (!v) return "";
  if (Object.prototype.toString.call(v) === "[object Date]") {
    if (isNaN(v.getTime())) return "";
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(v);
}
function formatTaskReminderTime_(v) {
  if (!v) return "";
  if (Object.prototype.toString.call(v) === "[object Date]") {
    if (isNaN(v.getTime())) return "";
    // Sheets stores a pure time-of-day value with its date part pinned to
    // its own null-date epoch (30 Dec 1899) - only the h:mm matters here.
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "h:mm a");
  }
  return String(v);
}
function taskRowToObj_(row, comments) {
  return {
    id: String(row[0] || ""),
    name: String(row[1] || ""),
    type: String(row[2] || ""),
    status: String(row[3] || ""),
    dueDate: formatTaskDate_(row[4]),
    createdBy: String(row[5] || ""),
    assignedTo: String(row[6] || ""),
    alsoVisibleTo: String(row[7] || "").split(",").map(function (s) { return s.trim(); }).filter(Boolean),
    area: String(row[8] || ""),
    createdAt: String(row[9] || ""),
    updatedAt: String(row[10] || ""),
    remarks: String(row[11] || ""),
    reminderTime: formatTaskReminderTime_(row[12]),
    deletedAt: String(row[13] || ""),
    // Dealer Name - added 8 Oct 2026, see TASKS_HEADERS comment above.
    dealerName: String(row[14] || ""),
    comments: comments || []
  };
}

// GET ?type=tasks&viewerName=X&viewerRole=Y[&allowedAreas=Z] - every
// task this viewer may see, newest-updated first, each with its own
// comment thread attached.
function handleTasksGet_(viewerName, viewerRole, allowedAreasCsv) {
  var taskRows = liveSheetRows_(getLiveSheet_(TASKS_SHEET, TASKS_HEADERS));
  var commentRows = liveSheetRows_(getLiveSheet_(TASK_COMMENTS_SHEET, TASK_COMMENTS_HEADERS));
  var commentsByTask = {};
  for (var c = 0; c < commentRows.length; c++) {
    var cr = commentRows[c];
    var tid = String(cr[1] || "").trim();
    if (!tid) continue;
    if (!commentsByTask[tid]) commentsByTask[tid] = [];
    commentsByTask[tid].push({ id: String(cr[0] || ""), taskId: tid, author: String(cr[2] || ""), text: String(cr[3] || ""), at: String(cr[4] || "") });
  }
  var out = [];
  for (var i = 0; i < taskRows.length; i++) {
    var row = taskRows[i];
    if (!String(row[0] || "").trim()) continue;
    // Soft-deleted (Deleted At column 14 set) - stays on the Sheet
    // forever as a record, per the owner's request, but is excluded from
    // the normal visible list here, same as a hard delete used to look
    // from every app's point of view.
    if (String(row[13] || "").trim()) continue;
    if (!taskRowVisible_(row, viewerRole, viewerName, allowedAreasCsv)) continue;
    var tid2 = String(row[0] || "").trim();
    var cmts = (commentsByTask[tid2] || []).sort(function (a, b) { return String(a.at).localeCompare(String(b.at)); });
    out.push(taskRowToObj_(row, cmts));
  }
  out.sort(function (a, b) { return String(b.updatedAt).localeCompare(String(a.updatedAt)); });
  return jsonOut_({ ok: true, tasks: out });
}

// POST { action: "taskCreate", name, type, status, dueDate, createdBy,
//   assignedTo, alsoVisibleTo (comma string or array), area, remarks,
//   reminderTime } - Owner/Manager creating a task (the dashboard's "+"
// button) or a DSR creating their own (see the aaadirect comment above
// myTasksDetails); modeled on the owner's reference screenshots.
// createdBy is the creator's own display name (e.g. "Alagu", "AAA
// Direct"), never a role label - handleTaskDelete_() below trusts this
// column as the record of who may delete the task. remarks (added 6 Oct
// 2026, "remarks need in create of task") is free text, optional.
// reminderTime (added same day, "time should ask while creating task
// only 3 time to show") is one of exactly three fixed slots the DSR/
// dashboard app UI offers ("11:00 AM", "3:00 PM", "4:30 PM") or blank for
// no reminder - not validated against that list here (both apps' own
// <select> only ever offers those three), so a wider UI language later
// doesn't need a backend change too.
function handleTaskCreate_(body) {
  var name = String(body.name || "").trim();
  if (!name) return jsonOut_({ ok: false, error: "name required" });
  // Due Date now required - added 8 Oct 2026 ("do not accept any task"
  // without a due date). Both apps' own form already validate this
  // client-side before the request is even sent; this is the server-side
  // backstop so a task can never end up with no due date no matter what
  // calls this endpoint.
  var dueDate = String(body.dueDate || "").trim();
  if (!dueDate) return jsonOut_({ ok: false, error: "Due date is required." });
  // Task Type now required - added 27 Oct 2026 ("compulsory" category
  // picker, no default/fallback value). Both apps' own Type <select> no
  // longer offers a pre-selected option and already block the request
  // client-side if nothing is picked; this is the server-side backstop,
  // same pattern as Due Date above. The old "General" fallback is gone -
  // there's no longer a category by that name in either app's dropdown.
  var type = String(body.type || "").trim();
  if (!type) return jsonOut_({ ok: false, error: "Task type is required." });
  var sheet = getLiveSheet_(TASKS_SHEET, TASKS_HEADERS);
  var id = Utilities.getUuid();
  var now = new Date().toISOString();
  var alsoVisibleTo = body.alsoVisibleTo;
  if (Array.isArray(alsoVisibleTo)) alsoVisibleTo = alsoVisibleTo.join(",");
  sheet.appendRow([
    id, name, type, String(body.status || "Open").trim(),
    dueDate, String(body.createdBy || "").trim(),
    String(body.assignedTo || "").trim(), String(alsoVisibleTo || "").trim(),
    String(body.area || "").trim(), now, now,
    String(body.remarks || "").trim(), String(body.reminderTime || "").trim(),
    "", // Deleted At - blank until handleTaskDelete_ stamps it
    String(body.dealerName || "").trim() // Dealer Name - added 8 Oct 2026
  ]);
  return jsonOut_({ ok: true, id: id });
}

// ---- One-time data cleanup - 27 Oct 2026 ("update the existing task
// also") ----
// The Task Type dropdown was narrowed/renamed this round: DSR apps now
// only ever offer Sales / Payment / Collection Follow up, and the
// dashboard's list dropped "General" and renamed "Stock" to
// "Stock / Inventory" (full new list: Sales, Payment, Collection Follow
// up, Dealer Visit, Scheme / Promotion, Stock / Inventory, Others).
// Existing tasks already saved with the old "Stock"/"General" values
// would otherwise sit there forever showing a name that's no longer in
// either dropdown. Run this ONCE from the Apps Script editor's Run
// dropdown (select migrateTaskTypes27Oct2026, click Run) after
// deploying this file - it rewrites every existing row in place:
//   "Stock"   -> "Stock / Inventory"
//   "General" -> "Others"
// Every other existing value (Sales, Payment, Others, blank, or anything
// already matching a current category) is left untouched. Safe to
// re-run - a second run finds nothing left to rename and changes 0 rows.
// NOTE: deliberately named with NO trailing underscore - Apps Script's
// editor hides any function ending in "_" from this exact Run dropdown
// (the same mistake this project already hit and fixed once before, see
// claude/master-file-task-manager-real-build-2026-10-06.md's "Apps
// Script run-dropdown fix", third delivery, 6 Oct 2026 - repeated here
// by accident on the first cut of this function and corrected the same
// day).
function migrateTaskTypes27Oct2026() {
  var RENAME_MAP_ = { "Stock": "Stock / Inventory", "General": "Others" };
  var sheet = getLiveSheet_(TASKS_SHEET, TASKS_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log("No task rows found - nothing to migrate."); return; }
  var typeCol = TASKS_HEADERS.indexOf("Task Type") + 1; // 1-based column index
  var range = sheet.getRange(2, typeCol, lastRow - 1, 1);
  var values = range.getValues();
  var changed = 0;
  for (var i = 0; i < values.length; i++) {
    var current = String(values[i][0] || "").trim();
    if (RENAME_MAP_.hasOwnProperty(current)) {
      values[i][0] = RENAME_MAP_[current];
      changed++;
    }
  }
  if (changed > 0) range.setValues(values);
  Logger.log("Task Type migration: " + changed + " row(s) updated out of " + values.length + " checked.");
}

// POST { action: "taskUpdateStatus", id, status, requestedByRole } - any
// viewer who can see a task can move it between the 4 statuses, EXCEPT:
// moving a task TO "In Progress" or "Completed" is now Owner/Manager
// only (added 8 Oct 2026, the owner's own words: "Only the complete in
// prograss option should be set my me or manager not area manager or
// dsr"). An Area Manager or DSR can still set Open/On Hold (e.g. to put
// something on hold or reopen it), just not mark it in-progress or done
// themselves - that stays a call for the owner or a Manager. Both apps'
// own status <select> already only offers the options this role is
// allowed to pick (see renderTaskManager_()/renderMyTasks_()); this is
// the server-side backstop so the restriction holds even if a request is
// ever sent some other way.
var TASK_STATUS_OWNER_MANAGER_ONLY_ = { "In Progress": true, "Completed": true };
function handleTaskUpdateStatus_(body) {
  var id = String(body.id || "").trim();
  var status = String(body.status || "").trim();
  var requestedByRole = String(body.requestedByRole || "").trim();
  if (!id || !status) return jsonOut_({ ok: false, error: "id and status required" });
  if (TASK_STATUS_OWNER_MANAGER_ONLY_[status] && requestedByRole !== "Owner" && requestedByRole !== "Manager") {
    return jsonOut_({ ok: false, error: "Only the Owner or a Manager can set this status." });
  }
  var sheet = getLiveSheet_(TASKS_SHEET, TASKS_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonOut_({ ok: false, error: "Task not found" });
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || "").trim() === id) {
      var now = new Date().toISOString();
      sheet.getRange(i + 2, 4, 1, 1).setValues([[status]]);
      sheet.getRange(i + 2, 11, 1, 1).setValues([[now]]);
      return jsonOut_({ ok: true });
    }
  }
  return jsonOut_({ ok: false, error: "Task not found" });
}

// POST { action: "taskAddComment", taskId, author, text }
function handleTaskAddComment_(body) {
  var taskId = String(body.taskId || "").trim();
  var text = String(body.text || "").trim();
  if (!taskId || !text) return jsonOut_({ ok: false, error: "taskId and text required" });
  var sheet = getLiveSheet_(TASK_COMMENTS_SHEET, TASK_COMMENTS_HEADERS);
  var id = Utilities.getUuid();
  var now = new Date().toISOString();
  sheet.appendRow([id, taskId, String(body.author || "").trim(), text, now]);
  var tasksSheet = getLiveSheet_(TASKS_SHEET, TASKS_HEADERS);
  var lastRow = tasksSheet.getLastRow();
  if (lastRow >= 2) {
    var ids = tasksSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || "").trim() === taskId) {
        tasksSheet.getRange(i + 2, 11, 1, 1).setValues([[now]]);
        break;
      }
    }
  }
  return jsonOut_({ ok: true, id: id, at: now });
}

// POST { action: "taskDelete", id, requestedBy, requestedByRole } - added
// 6 Oct 2026 ("add a delete button near a task... that delete should
// work only who created the task"). requestedBy must exactly match the
// task's own Created By value (column 6) - both apps always send their
// own fixed identity (CURRENT_USER_NAME on the dashboard, DSR_SHORT_NAME
// in a DSR app) as requestedBy, the same value they'd have sent as
// createdBy when the task was made, so this is a simple, reliable "only
// the creator" check with no separate permissions table needed.
//
// CHANGED 8 Oct 2026 - the owner's own words: "For my id put delete
// option who ever created i can also delete." This revisits the original
// "no bypass, not even for Owner/Manager" decision below, which the
// owner has now explicitly overridden for their own login specifically:
// requestedByRole === "Owner" skips the Created-By check entirely. A
// Manager/Area Manager/DSR still gets no bypass - only Owner.
//
// CHANGED 6 Oct 2026, same day - per the owner's explicit correction
// ("after delete also the written should be on sheet as deleted"): this
// used to hard-delete the row (sheet.deleteRow) and wipe the whole
// comment thread out of TaskComments, leaving nothing on the Sheet at
// all once a task was removed. Now a delete is a SOFT delete - it stamps
// the Deleted At column (14) with an ISO timestamp and leaves the task
// row, its Status (whatever it was at the moment of deletion), and every
// one of its comments exactly as they were. handleTasksGet_() filters
// deleted tasks out of the normal list, so every app's behavior looks
// identical to before (the task disappears) - the difference is purely
// that the complete record now survives permanently on the Sheet instead
// of being erased, which is exactly what was asked for.
function handleTaskDelete_(body) {
  var id = String(body.id || "").trim();
  var requestedBy = String(body.requestedBy || "").trim();
  var requestedByRole = String(body.requestedByRole || "").trim();
  if (!id || !requestedBy) return jsonOut_({ ok: false, error: "id and requestedBy required" });
  var sheet = getLiveSheet_(TASKS_SHEET, TASKS_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonOut_({ ok: false, error: "Task not found" });
  var rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues(); // just need ID (col 1) and Created By (col 6)
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0] || "").trim() === id) {
      var createdBy = String(rows[i][5] || "").trim();
      if (createdBy !== requestedBy && requestedByRole !== "Owner") {
        return jsonOut_({ ok: false, error: "Only the person who created this task can delete it." });
      }
      var now = new Date().toISOString();
      sheet.getRange(i + 2, 14, 1, 1).setValues([[now]]); // Deleted At - row and comments left untouched
      return jsonOut_({ ok: true });
    }
  }
  return jsonOut_({ ok: false, error: "Task not found" });
}

// ----- Master File (dashboard section) -----

// Per-dealer actual Sales Ltr/Collection Rs for monthKey, from
// whichever source is correct for that month: the fixed
// MasterFileHistory import for anything before
// MASTER_FILE_LIVE_START_MONTH_, the live LiveSalesByDealer/
// LiveCollectionByDealer sheets for the CURRENT month, or their
// Monthly Freeze archive for a past month that's already live-tracked
// (see FREEZE_LIVE_CATEGORIES_ above - both sheets are already frozen
// every month-end, so no new freeze plumbing is needed here).
function masterFileActualsForMonth_(monthKey) {
  if (!masterFileMonthIsLive_(monthKey)) {
    var h = masterFileHistoryMonthMap_(monthKey);
    var out = {};
    for (var d in h) out[d] = { salesLtr: h[d].salesLtr, collectionRs: h[d].collectionRs, isLive: false };
    return out;
  }
  var todayMonthKey = Utilities.formatDate(new Date(), DISPLAY_TIMEZONE, "yyyy-MM");
  var salesRows, collRows;
  if (monthKey === todayMonthKey) {
    salesRows = liveSheetRows_(getLiveSheet_(LIVE_SALES_BY_DEALER_SHEET, LIVE_SALES_BY_DEALER_HEADERS));
    collRows = liveSheetRows_(getLiveSheet_(LIVE_COLLECTION_BY_DEALER_SHEET, LIVE_COLLECTION_BY_DEALER_HEADERS));
  } else {
    salesRows = readFrozenSheetForMonth_(LIVE_SALES_BY_DEALER_SHEET, LIVE_SALES_BY_DEALER_HEADERS, monthKey) || [];
    collRows = readFrozenSheetForMonth_(LIVE_COLLECTION_BY_DEALER_SHEET, LIVE_COLLECTION_BY_DEALER_HEADERS, monthKey) || [];
  }
  var map = {};
  for (var i = 0; i < salesRows.length; i++) {
    var r = salesRows[i];
    var name = String(r[0] || "").trim();
    if (!name) continue;
    if (!map[name]) map[name] = { salesLtr: null, collectionRs: null, isLive: true };
    map[name].salesLtr = r[2] === "" || r[2] === undefined ? null : Number(r[2]); // Month Sold Ltr
  }
  for (var j = 0; j < collRows.length; j++) {
    var r2 = collRows[j];
    var name2 = String(r2[0] || "").trim();
    if (!name2) continue;
    if (!map[name2]) map[name2] = { salesLtr: null, collectionRs: null, isLive: true };
    map[name2].collectionRs = r2[2] === "" || r2[2] === undefined ? null : Number(r2[2]); // Month figure
  }
  return map;
}

// GET ?type=masterFile&month=yyyy-MM[&dsrName=X] - the dashboard's
// Master File table: the same per-dealer row shape as
// handleDealerTargetsGet_() above, PLUS that month's actual Sales Ltr
// / Collection Rs (masterFileActualsForMonth_()) and whether this
// month counts as a fixed/historical month or a live one.
function handleMasterFileGet_(monthKey, dsrNameFilter) {
  if (!monthKey) return jsonOut_({ ok: false, error: "month required" });
  var base = JSON.parse(handleDealerTargetsGet_(monthKey, dsrNameFilter).getContent());
  if (!base.ok) return jsonOut_(base);
  var actuals = masterFileActualsForMonth_(monthKey);
  for (var i = 0; i < base.dealers.length; i++) {
    var d = base.dealers[i];
    var a = actuals[d.dealer];
    d.salesLtr = a ? a.salesLtr : null;
    d.collectionRs = a ? a.collectionRs : null;
  }
  base.isLiveMonth = masterFileMonthIsLive_(monthKey);
  return jsonOut_(base);
}

// GET ?type=masterFileReview&dsrName=X - "Review by DSR": every month
// from Jan 2026 to the current month for every dealer under this DSR,
// plus a 6M AVG column (the trailing 6 CLOSED months - i.e. not
// counting the current, still-open month, the SAME "6M AVG" convention
// used throughout this dashboard already, e.g. LIVE_SALES_BY_DEALER's
// own Avg 6mo Sold Ltr column above).
function handleMasterFileReviewGet_(dsrName) {
  if (!dsrName) return jsonOut_({ ok: false, error: "dsrName required" });
  var meta = dealerMetaMap_();
  var dealers = [];
  for (var name in meta) if (meta[name].dsrName === dsrName) dealers.push(name);

  var todayMonthKey = Utilities.formatDate(new Date(), DISPLAY_TIMEZONE, "yyyy-MM");
  var months = [];
  var mk = "2026-01";
  while (mk <= todayMonthKey) { months.push(mk); mk = addMonthsToMonthKey_(mk, 1); }

  var byMonth = {};
  for (var mi = 0; mi < months.length; mi++) byMonth[months[mi]] = masterFileActualsForMonth_(months[mi]);

  // Trailing 6 CLOSED months (excludes the current, still-open month).
  var closedMonths = months.slice(0, months.length - 1).slice(-6);

  var out = [];
  for (var di = 0; di < dealers.length; di++) {
    var dealer = dealers[di];
    var monthVals = {};
    var sumSalesLtr = 0, sumCollRs = 0, cnt = 0;
    for (var m = 0; m < months.length; m++) {
      var rec = byMonth[months[m]][dealer] || { salesLtr: null, collectionRs: null };
      monthVals[months[m]] = { salesLtr: rec.salesLtr, collectionRs: rec.collectionRs };
    }
    for (var cm = 0; cm < closedMonths.length; cm++) {
      var crec = byMonth[closedMonths[cm]][dealer];
      if (crec && crec.salesLtr !== null && crec.salesLtr !== undefined) { sumSalesLtr += crec.salesLtr; cnt++; }
    }
    var avg6moLtr = cnt > 0 ? Math.round(sumSalesLtr / cnt) : null;
    out.push({ dealer: dealer, months: monthVals, avg6moLtr: avg6moLtr, avg6moFromMonths: closedMonths });
  }
  return jsonOut_({ ok: true, dsrName: dsrName, months: months, closedMonths: closedMonths, dealers: out });
}

// GET ?type=masterFileReviewAll - "Review DSR" (Owner-side dashboard,
// company-wide): one row per INDIVIDUAL DEALER across every DSR, same
// month-on-month Sales Ltr / Collection Rs columns from Jan 2026 to the
// current month, plus the same trailing-6-CLOSED-months AVG convention
// as handleMasterFileReviewGet_() above - just not filtered down to one
// DSR, with each dealer's own DSR Name/Area/Beat Day carried in the row
// so the dashboard can show/sort which DSR a dealer belongs to. Added
// 6 Oct 2026 for the owner's "I want file to review DSR... month on
// month Sales and Collection data like the Master File has" request -
// scoped to one row per dealer (not a DSR-level rollup) and Dashboard-
// only (Owner-side), per the owner's own answers when asked.
function handleMasterFileReviewAllGet_() {
  var meta = dealerMetaMap_();
  var dealers = [];
  for (var name in meta) dealers.push(name);
  dealers.sort(); // stable, predictable order for the dashboard table

  var todayMonthKey = Utilities.formatDate(new Date(), DISPLAY_TIMEZONE, "yyyy-MM");
  var months = [];
  var mk = "2026-01";
  while (mk <= todayMonthKey) { months.push(mk); mk = addMonthsToMonthKey_(mk, 1); }

  var byMonth = {};
  for (var mi = 0; mi < months.length; mi++) byMonth[months[mi]] = masterFileActualsForMonth_(months[mi]);

  // Trailing 6 CLOSED months (excludes the current, still-open month) -
  // same convention as handleMasterFileReviewGet_() and
  // LIVE_SALES_BY_DEALER's own Avg 6mo Sold Ltr column.
  var closedMonths = months.slice(0, months.length - 1).slice(-6);

  var out = [];
  for (var di = 0; di < dealers.length; di++) {
    var dealer = dealers[di];
    var m = meta[dealer];
    var monthVals = {};
    for (var mm = 0; mm < months.length; mm++) {
      var rec = byMonth[months[mm]][dealer] || { salesLtr: null, collectionRs: null };
      monthVals[months[mm]] = { salesLtr: rec.salesLtr, collectionRs: rec.collectionRs };
    }
    var sumSalesLtr = 0, cnt = 0;
    for (var cm = 0; cm < closedMonths.length; cm++) {
      var crec = byMonth[closedMonths[cm]][dealer];
      if (crec && crec.salesLtr !== null && crec.salesLtr !== undefined) { sumSalesLtr += crec.salesLtr; cnt++; }
    }
    var avg6moLtr = cnt > 0 ? Math.round(sumSalesLtr / cnt) : null;
    out.push({
      dealer: dealer,
      dsrName: m.dsrName,
      area: m.area,
      beatDay: m.beatDay,
      months: monthVals,
      avg6moLtr: avg6moLtr
    });
  }
  return jsonOut_({ ok: true, months: months, closedMonths: closedMonths, dealers: out });
}

// GET ?type=masterFileExport&month=yyyy-MM - exports the current
// Master File roster as a real .xlsx, in the SAME column layout as the
// owner's own upload (DEALER NAME, SALES REP, AREA, TOWN, DISTRICT,
// SEG, BEAT DAY, Overall Outstanding, then that one month's Sales/
// Collection, Target/Plan, Collectable - owner: "the whole master file
// has to be downloaded with a option and it has to be as per my
// upload format"). Builds a throwaway Google Sheet, exports it to
// .xlsx via the standard Sheets export-URL trick, saves the file to
// Drive (same archive folder convention as
// handleArchiveMonthlyPdf_() above) and returns its download URL - the
// dashboard opens that URL in a new tab.
function handleMasterFileExportGet_(monthKey) {
  if (!monthKey) return jsonOut_({ ok: false, error: "month required" });
  var full = JSON.parse(handleMasterFileGet_(monthKey, "").getContent());
  if (!full.ok) return jsonOut_(full);

  var ss = SpreadsheetApp.create("MasterFile_Export_" + monthKey + "_" + new Date().getTime());
  var sheet = ss.getSheets()[0];
  var headerRow = [
    "DEALER NAME", "SALES REP", "AREA", "BEAT DAY", "Overall Outstanding",
    "Sales Ltr", "Collection Rs", "Sales Target Ltr", "Collection Plan Rs",
    "Collectable (45d+)", "Suggested Target Ltr", "Remark"
  ];
  var dataRows = full.dealers.map(function (d) {
    return [
      d.dealer, d.dsrName, d.area, d.beatDay,
      d.overallOutstandingRs === null ? "" : d.overallOutstandingRs,
      d.salesLtr === null ? "" : d.salesLtr,
      d.collectionRs === null ? "" : d.collectionRs,
      d.salesTargetLtr === null ? "" : d.salesTargetLtr,
      d.collectionPlanRs === null ? "" : d.collectionPlanRs,
      d.collectableRs === null ? "" : d.collectableRs,
      d.suggestedLtr === null ? "" : d.suggestedLtr,
      d.remark || ""
    ];
  });
  sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  if (dataRows.length > 0) sheet.getRange(2, 1, dataRows.length, headerRow.length).setValues(dataRows);
  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();

  var fileId = ss.getId();
  var exportUrl = "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";
  var token = ScriptApp.getOAuthToken();
  var resp = UrlFetchApp.fetch(exportUrl, { headers: { Authorization: "Bearer " + token } });
  var folder = getOrCreateFolder_(archiveRootFolder_(), "MasterFileExports");
  var xlsxFile = folder.createFile(resp.getBlob().setName("MasterFile_" + monthKey + ".xlsx"));
  xlsxFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  DriveApp.getFileById(fileId).setTrashed(true); // drop the throwaway Sheet, keep only the .xlsx

  return jsonOut_({ ok: true, downloadUrl: "https://drive.google.com/uc?export=download&id=" + xlsxFile.getId() });
}

// Admin/setup: imports the owner's monthly Master File Excel (already
// uploaded as a sheet TAB named exactly "MasterFileRaw" - see the
// delivery notes for the exact import steps) into DealerMeta (fully
// replaced every run - this sheet is pure reference data, always safe
// to overwrite) and MasterFileHistory (upserted by dealer+monthKey, so
// re-running this after the owner adds a new month's columns to their
// Excel just adds that one new month without disturbing anything
// already live-tracked). Safe to re-run every time the owner refreshes
// their Excel - this is meant to become their normal monthly habit.
// Run manually from the Apps Script editor (select seedMasterFileFromRaw
// in the function dropdown, then Run) - never called by any DSR
// app/dashboard request.
// Normalizes a raw "SALES REP" cell to the exact DSR_SHORT_NAME each app
// actually filters on - added 6 Oct 2026 after a real bug: the owner's
// real Master File Excel uses "AAA" (89 rows) for his own direct-control
// dealers, not "AAA Direct" as the upload instructions asked for, and
// every other name in the real file comes in ALL CAPS ("CHELLAMANI",
// "ARUN", ...) - seedMasterFileFromRaw() used to store that raw text
// verbatim into DealerMeta's DSR Name column, so handleDealerTargetsGet_'s
// exact-match filter (dsrNameFilter, e.g. "AAA Direct") never matched any
// seeded row, and AAA Direct's Dealer Targets panel showed "No dealers
// found for this app yet." even though 684 dealers were seeded. Case-
// insensitive lookup here, falling back to the raw trimmed value
// unchanged for anything not recognized (e.g. "Vacant Karaikudi", or a
// blank SALES REP cell) - those dealers simply won't match any app's
// DSR_SHORT_NAME, same as already-documented behavior for a genuine
// spelling mistake, nothing is ever silently dropped or reassigned.
var MASTER_FILE_SALES_REP_ALIASES_ = {
  "aaa": "AAA Direct",
  "aaa direct": "AAA Direct",
  "prabha": "Prabha",
  "arun": "Arun",
  "nagaraj": "Nagaraj",
  "chellamani": "Chellamani",
  "anandh": "Anandh",
  "madhu": "Madhu"
};
function normalizeMasterFileSalesRep_(raw) {
  var trimmed = String(raw || "").trim();
  if (!trimmed) return "";
  var hit = MASTER_FILE_SALES_REP_ALIASES_[trimmed.toLowerCase()];
  return hit || trimmed;
}

// Added 10 Oct 2026 (District update re-import) - the owner's Master
// File Excel has District/Segment typed inconsistently (e.g. "Dindigul"
// vs "DINDIGUL" vs "Dindigul " with a trailing space, "Tenkasi" vs
// "TENKASI"), which would otherwise create silently-duplicate-looking
// District/Segment values in the dashboard's own District/Area filters.
// Trims, collapses repeated internal whitespace, and Title-Cases each
// word ("COIMBATORE" / "coimbatore" / "Coimbatore " -> "Coimbatore").
// A bare "0" (seen once in the real Segment column - clearly not a real
// segment name) is treated as blank, same as an empty cell, so it falls
// through to the carry-forward/blank handling below rather than being
// stored as the literal text "0".
function normalizeDistrictOrSegment_(raw) {
  var trimmed = String(raw || "").trim().replace(/\s+/g, " ");
  if (!trimmed || trimmed === "0") return "";
  return trimmed.toLowerCase().replace(/\b\w/g, function (c) { return c.toUpperCase(); });
}

// normalizeBeatDay_(raw) - added 10 Oct 2026, fixing a real bug reported
// by the owner ("beat day is not showing in Master File") even though the
// real Master File Excel actually HAS a Beat Day for the great majority
// of dealers (589 of 678 in this round's upload - MONDAY 122, TUESDAY 110,
// WEDNESDAY 107, FRIDAY 106, THURSDAY 79, SATURDAY 65). Root cause:
// seedMasterFileFromRaw() used to store this column completely
// unnormalized (plain trim only), but the Excel writes it in ALL CAPS
// ("MONDAY") while the dashboard's Beat Day <select> only offers
// Title-Case options ("Monday"). A stored "MONDAY" never exactly matches
// <option value="Monday">, so the <select> silently fell back to its
// first/blank option and showed "Not set" for EVERY dealer - including
// the 589 whose Beat Day was actually sitting there correctly in
// DealerMeta the whole time. Fixed the same way District/Segment already
// are: recognizes the full day name and the common 3-letter abbreviation
// in any casing, and treats a stray "0" (55 dealers this round - same
// empty-cell placeholder pattern already seen in Segment) as blank. A
// genuinely unrecognized value (a typo) is kept as-is rather than
// silently dropped - dashboard/index.html's own Beat Day <select> was
// also hardened (10 Oct 2026) to show an unmatched value as an extra
// option instead of defaulting to "Not set", so this exact failure mode
// can't silently recur even if some future value doesn't normalize here.
var BEAT_DAY_ALIASES_ = {
  "monday": "Monday", "mon": "Monday",
  "tuesday": "Tuesday", "tue": "Tuesday", "tues": "Tuesday",
  "wednesday": "Wednesday", "wed": "Wednesday",
  "thursday": "Thursday", "thu": "Thursday", "thur": "Thursday", "thurs": "Thursday",
  "friday": "Friday", "fri": "Friday",
  "saturday": "Saturday", "sat": "Saturday",
  "sunday": "Sunday", "sun": "Sunday"
};
function normalizeBeatDay_(raw) {
  var trimmed = String(raw || "").trim();
  if (!trimmed || trimmed === "0") return "";
  var key = trimmed.toLowerCase();
  return BEAT_DAY_ALIASES_[key] || trimmed;
}

function seedMasterFileFromRaw() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var raw = ss.getSheetByName("MasterFileRaw");
  if (!raw) throw new Error('Import your Master File Excel as a tab named exactly "MasterFileRaw" first (File > Import > Insert new sheet), then run this again.');
  var lastRow = raw.getLastRow();
  var lastCol = raw.getLastColumn();
  if (lastRow < 2) throw new Error("MasterFileRaw has no data rows.");
  var values = raw.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var metaSheet = getLiveSheet_(DEALER_META_SHEET, DEALER_META_HEADERS);
  // Snapshot any already-known DSR assignment per dealer BEFORE clearing -
  // added 6 Oct 2026 so a DSR picked from the dashboard's DSR dropdown
  // (dealerMetaSet) for a dealer the Excel still leaves blank ("give
  // blank dealers as no DSR, assign later from the dashboard") survives
  // the next MasterFileRaw re-import instead of being wiped back to
  // blank just because this round's Excel cell still says nothing. Only
  // a RECOGNIZED DSR name (one of the 7 app names) is carried forward -
  // never a leftover "Vacant Karaikudi" or similar - and only used when
  // this round's Excel cell does NOT resolve to a known DSR; a real,
  // recognized value in the new Excel always wins and overwrites it, same
  // as before.
  var existingKnownDsrByDealer_ = {};
  // Beat Day carry-forward added 6 Oct 2026, same day as the issue was
  // reported ("beat day is empty for all") - the owner's Master File
  // Excel simply has no Beat Day column filled in, so Beat Day can only
  // ever come from being picked per-dealer on the dashboard's Master
  // File table (the mf-beat-input <select>, via dealerMetaSet). Without
  // this, the very next MasterFileRaw re-import would wipe every Beat
  // Day the owner had just assigned, straight back to "Not set" - the
  // exact same trap already fixed for DSR Name above. Area gets the same
  // treatment for consistency, in case a future Excel round also leaves
  // some dealer's Area blank after it was filled in from the dashboard.
  var existingBeatDayByDealer_ = {};
  var existingAreaByDealer_ = {};
  // District/Segment carry-forward added 10 Oct 2026 (same reasoning as
  // Area/Beat Day above) - this round's Excel has 31 dealers with a
  // blank District cell; without this, a dealer whose District was ever
  // filled in (this round or a future dashboard edit) would get wiped
  // back to blank the next time a re-import happens to leave it blank
  // again. A real, non-blank value in the Excel always still wins.
  var existingDistrictByDealer_ = {};
  var existingSegmentByDealer_ = {};
  if (metaSheet.getLastRow() > 1) {
    var existingRows_ = metaSheet.getRange(2, 1, metaSheet.getLastRow() - 1, DEALER_META_HEADERS.length).getValues();
    for (var e = 0; e < existingRows_.length; e++) {
      var exDealer_ = String(existingRows_[e][0] || "").trim();
      if (!exDealer_) continue;
      var exDsr_ = String(existingRows_[e][1] || "").trim();
      if (exDsr_ && MASTER_FILE_SALES_REP_ALIASES_[exDsr_.toLowerCase()]) {
        existingKnownDsrByDealer_[exDealer_] = exDsr_;
      }
      var exArea_ = String(existingRows_[e][2] || "").trim();
      if (exArea_) existingAreaByDealer_[exDealer_] = exArea_;
      var exDistrict_ = String(existingRows_[e][4] || "").trim();
      if (exDistrict_) existingDistrictByDealer_[exDealer_] = exDistrict_;
      var exSegment_ = String(existingRows_[e][5] || "").trim();
      if (exSegment_) existingSegmentByDealer_[exDealer_] = exSegment_;
      var exBeatDay_ = String(existingRows_[e][6] || "").trim();
      if (exBeatDay_) existingBeatDayByDealer_[exDealer_] = exBeatDay_;
    }
  }
  // Fully replace DealerMeta - pure reference data, no history to lose
  // (DSR Name is the one field with the carry-forward exception above).
  if (metaSheet.getLastRow() > 1) metaSheet.getRange(2, 1, metaSheet.getLastRow() - 1, DEALER_META_HEADERS.length).clearContent();
  var metaOut = [];
  var historyUpdates = [];
  // unmatchedSalesRepCounts_ - every raw Excel value that didn't resolve
  // to a known DSR this round, blank cells included (labeled "(blank
  // SALES REP cell)" since they'd otherwise be invisible in this log).
  // carriedForwardDsrCount_ / stillUnassignedCount_ added alongside the
  // snapshot above - split those same rows into "kept an existing
  // dashboard/previous-import assignment" vs. "still has no DSR at all,
  // ready to be picked from the dashboard's Master File table".
  var unmatchedSalesRepCounts_ = {};
  var carriedForwardDsrCount_ = 0;
  var stillUnassignedCount_ = 0;
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var dealer = String(row[0] || "").trim();
    if (!dealer) continue;
    var rawSalesRep_ = String(row[1] || "").trim();
    var dsrNameNormalized = normalizeMasterFileSalesRep_(row[1]);
    var dsrNameIsKnown_ = !!(dsrNameNormalized && MASTER_FILE_SALES_REP_ALIASES_[dsrNameNormalized.toLowerCase()]);
    if (!dsrNameIsKnown_) {
      var unmatchedLabel_ = rawSalesRep_ ? dsrNameNormalized : "(blank SALES REP cell)";
      unmatchedSalesRepCounts_[unmatchedLabel_] = (unmatchedSalesRepCounts_[unmatchedLabel_] || 0) + 1;
      if (existingKnownDsrByDealer_[dealer]) {
        dsrNameNormalized = existingKnownDsrByDealer_[dealer];
        carriedForwardDsrCount_++;
      } else {
        dsrNameNormalized = "";
        stillUnassignedCount_++;
      }
    }
    var areaFromExcel_ = String(row[2] || "").trim();
    var areaOut_ = areaFromExcel_ || existingAreaByDealer_[dealer] || "";
    // District/Segment: normalize casing/whitespace first (see
    // normalizeDistrictOrSegment_ above - also turns a stray "0" cell
    // into blank), THEN carry forward a previous value if this round's
    // cell comes out blank either way.
    var districtFromExcel_ = normalizeDistrictOrSegment_(row[4]);
    var districtOut_ = districtFromExcel_ || existingDistrictByDealer_[dealer] || "";
    var segmentFromExcel_ = normalizeDistrictOrSegment_(row[5]);
    var segmentOut_ = segmentFromExcel_ || existingSegmentByDealer_[dealer] || "";
    var beatDayFromExcel_ = normalizeBeatDay_(row[6]);
    var beatDayOut_ = beatDayFromExcel_ || existingBeatDayByDealer_[dealer] || "";
    metaOut.push([
      dealer, dsrNameNormalized, areaOut_,
      String(row[3] || "").trim(), districtOut_, segmentOut_,
      beatDayOut_,
      row[7] === "" || row[7] === undefined || row[7] === null ? "" : Number(row[7])
    ]);
    // Month-wise Sales/Collection (field "sales"/"coll") AND the once-
    // a-year 2018-2023 totals (field "year") both land in
    // MasterFileHistory. Fixed 10 Oct 2026 - the "year" entries used to
    // be skipped here with a comment claiming they were "seeded
    // separately below", but no such code ever existed, so every
    // MasterFileRaw import silently dropped 2018-2023 entirely. Now
    // pushed through the same historyUpdates list as every other
    // month, tagged isYearTotal so the merge step below can store them
    // with Collection Rs deliberately left blank (yearly rows are
    // Sales-only, per the owner's own "only total sales data no
    // collection for that" instruction).
    for (var c = 0; c < MASTER_FILE_RAW_MONTH_COLS_.length; c++) {
      var spec = MASTER_FILE_RAW_MONTH_COLS_[c];
      var val = row[spec.col];
      if (val === "" || val === undefined || val === null) continue;
      historyUpdates.push({
        dealer: dealer, monthKey: spec.monthKey, field: spec.field, value: Number(val),
        isYearTotal: spec.field === "year"
      });
    }
  }
  if (metaOut.length > 0) metaSheet.getRange(2, 1, metaOut.length, DEALER_META_HEADERS.length).setValues(metaOut);

  // Merge sales+coll pairs for the same (dealer, monthKey) into one
  // MasterFileHistory row each, then upsert by that composite key.
  var merged = {};
  for (var u = 0; u < historyUpdates.length; u++) {
    var up = historyUpdates[u];
    var key = up.dealer + "||" + up.monthKey;
    if (!merged[key]) merged[key] = { dealer: up.dealer, monthKey: up.monthKey, salesLtr: null, collectionRs: null, isYearTotal: false };
    if (up.field === "sales") merged[key].salesLtr = up.value;
    if (up.field === "coll") merged[key].collectionRs = up.value;
    // A "year" entry is Sales-only by design (2018-2023 - the owner's
    // own "only total sales data no collection for that") - it always
    // sets salesLtr and marks the row, and since the "year" monthKeys
    // ("2023", "2022", ...) never collide with a real "yyyy-MM" month
    // key, this merge key is never shared with any real month's row.
    if (up.isYearTotal) { merged[key].salesLtr = up.value; merged[key].isYearTotal = true; }
  }
  var mergedList = [];
  for (var k in merged) mergedList.push(merged[k]);

  var histSheet = getLiveSheet_(MASTER_FILE_HISTORY_SHEET, MASTER_FILE_HISTORY_HEADERS);
  bulkUpsertSheet_(histSheet, MASTER_FILE_HISTORY_HEADERS.length, mergedList,
    function (item) { return item.dealer + "||" + item.monthKey; },
    function (item) {
      return [
        item.dealer, item.monthKey,
        item.salesLtr === null ? "" : item.salesLtr,
        item.isYearTotal ? "" : (item.collectionRs === null ? "" : item.collectionRs),
        item.isYearTotal ? "Y" : ""
      ];
    },
    function (existingRow) { return String(existingRow[0] || "").trim() + "||" + String(existingRow[1] || "").trim(); }
  );

  var unmatchedSalesRepSummary_ = [];
  for (var unmatchedName in unmatchedSalesRepCounts_) {
    unmatchedSalesRepSummary_.push(unmatchedName + " (" + unmatchedSalesRepCounts_[unmatchedName] + ")");
  }
  Logger.log("DealerMeta rows: " + metaOut.length + " | MasterFileHistory rows upserted: " + mergedList.length);
  if (unmatchedSalesRepSummary_.length > 0) {
    Logger.log("SALES REP values this round's Excel didn't resolve to a known DSR: " + unmatchedSalesRepSummary_.join(", ") + ". Of these, " + carriedForwardDsrCount_ + " kept the DSR already assigned to them (dashboard or a previous import), and " + stillUnassignedCount_ + " still have no DSR at all - pick one for those from the dashboard's Master File table.");
  }
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
    if (String(params.type || "").trim() === "itemGroupSixMoAvg") {
      return handleItemGroupSixMoAvgGet_();
    }
    // "Monthly Freeze" - added 1 Oct 2026, see the long comment above
    // handleMonthlyFreezeGet_() for the full feature.
    if (String(params.type || "").trim() === "monthlyFreeze") {
      return handleMonthlyFreezeGet_(String(params.month || "").trim());
    }
    // "Daily Freeze" - added 2 Oct 2026, see the long comment above
    // handleDailyFreezeGet_() for the full feature.
    if (String(params.type || "").trim() === "dailyFreeze") {
      return handleDailyFreezeGet_(String(params.date || "").trim());
    }
    // "Report period" Week/Quarter/Year - added 2 Oct 2026, see the long
    // comment above handleWeekSummaryGet_()/handlePeriodSummaryGet_() for
    // the full feature.
    if (String(params.type || "").trim() === "weekSummary") {
      return handleWeekSummaryGet_(String(params.month || "").trim(), String(params.week || "").trim());
    }
    if (String(params.type || "").trim() === "periodSummary") {
      return handlePeriodSummaryGet_(String(params.start || "").trim(), String(params.end || "").trim());
    }
    // "Today" - added 3 Oct 2026, see handleTodaySummaryGet_() above.
    if (String(params.type || "").trim() === "todaySummary") {
      return handleTodaySummaryGet_();
    }
    // Master File / Dealer Targets / Remarks / Task Manager reads, added
    // 6 Oct 2026 - see the big comment block above
    // handleDealerTargetsGet_() for the full feature.
    if (String(params.type || "").trim() === "dealerTargets") {
      return handleDealerTargetsGet_(String(params.month || "").trim(), String(params.dsrName || "").trim());
    }
    if (String(params.type || "").trim() === "dealerRemarks") {
      return handleDealerRemarksGet_(String(params.dealer || "").trim());
    }
    // Bank Deposit, added 23 Oct 2026 - see handleBankDepositsGet_() above.
    if (String(params.type || "").trim() === "bankDeposits") {
      return handleBankDepositsGet_(String(params.dsrName || "").trim(), String(params.from || "").trim(), String(params.to || "").trim(), String(params.latestOnly || "").trim() === "1");
    }
    // "Today's Dealer Visit Check" date picker, added 26 Oct 2026
    // ("can i get date changing option inside") - ?type=live always
    // computes TODAY's check (see handleLiveGet_() -> dealerVisitCheckLive_()
    // with no argument); this route lets the dashboard's date picker ask
    // for any other date on demand, via the same function with a
    // "YYYY-MM-DD" override. See dealerVisitCheckLive_()'s own comment for
    // the beat-day-has-no-history caveat this applies to a past date too.
    if (String(params.type || "").trim() === "dealerVisitCheck") {
      var dvcResult = dealerVisitCheckLive_(String(params.date || "").trim() || null);
      return jsonOut_({ ok: !dvcResult.error, error: dvcResult.error || null, dealerVisitCheck: dvcResult });
    }
    if (String(params.type || "").trim() === "tasks") {
      return handleTasksGet_(String(params.viewerName || "").trim(), String(params.viewerRole || "").trim(), String(params.allowedAreas || "").trim());
    }
    if (String(params.type || "").trim() === "masterFile") {
      return handleMasterFileGet_(String(params.month || "").trim(), String(params.dsrName || "").trim());
    }
    if (String(params.type || "").trim() === "masterFileReview") {
      return handleMasterFileReviewGet_(String(params.dsrName || "").trim());
    }
    if (String(params.type || "").trim() === "masterFileReviewAll") {
      return handleMasterFileReviewAllGet_();
    }
    if (String(params.type || "").trim() === "masterFileExport") {
      return handleMasterFileExportGet_(String(params.month || "").trim());
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
    // "Section visibility (by role)", added 1 Oct 2026 - see
    // handleSectionVisibilityGet_() above.
    if (String(params.type || "").trim() === "sectionVisibility") {
      return handleSectionVisibilityGet_();
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
 *   { action: "liveCreditLimitUpdate",  key, dealers: [...] } — see "Live Tally Sync" above
 *   { action: "liveDealerRosterUpdate", key, roster: [...] }  — see "Live Tally Sync" above
 *   { action: "liveCollectionUpdate",    key, collection: [...] } — see "Live Tally Sync" above
 *   { action: "liveCollectionByDealerUpdate", key, collection: [...] } — see "Live Tally Sync" above
 *   { action: "livePetronasSummaryUpdate", key, petronasSummary: {month, monthLabel, purchaseLtr,
 *       salesLtr, salesValue, collected, purchaseValue, payment} } — current-month-only live push
 *       from tally_live_watcher.py's maybe_push_petronas_summary(), upserted by Month Key into the
 *       LivePetronasSummary sheet; read back inside handleLiveGet_() as the petronasSummary field
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
  //
  // Wrapped in try/catch added 3 Oct 2026: a real production run of
  // freezeMonthSnapshot threw an uncaught exception in here (confirmed - this
  // whole pre-lock block had no exception handling of its own, unlike doGet()
  // above which wraps its entire dispatch). An uncaught exception in an Apps
  // Script web app comes back as a generic Google-produced 500 "Internal
  // Server Error" page with no error text at all, which is why
  // monthly_archive.py's log just showed "500 Server Error" three times with
  // no clue what actually broke. This catch turns that into a readable
  // {ok:false, error:...} JSON response instead, matching the pattern doGet()
  // already uses - it doesn't fix the underlying bug by itself, but the next
  // failure (after this is deployed) will say what the bug actually is.
  try {
    if (body.action === "liveStockUpdate") return handleLiveStockUpdate_(body);
    if (body.action === "liveOutstandingUpdate") return handleLiveOutstandingUpdate_(body);
    if (body.action === "liveCreditLimitUpdate") return handleLiveCreditLimitUpdate_(body);
    if (body.action === "liveDealerRosterUpdate") return handleLiveDealerRosterUpdate_(body);
    if (body.action === "liveNewDealerUpdate") return handleLiveNewDealerUpdate_(body);
    if (body.action === "liveCollectionUpdate") return handleLiveCollectionUpdate_(body);
    if (body.action === "liveCollectionByDealerUpdate") return handleLiveCollectionByDealerUpdate_(body);
    if (body.action === "livePetronasSummaryUpdate") return handleLivePetronasSummaryUpdate_(body);
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
    if (body.action === "liveSalesTrendByDsrUpdate") return handleLiveSalesTrendByDsrUpdate_(body);
    if (body.action === "liveCollectionTrendByDsrUpdate") return handleLiveCollectionTrendByDsrUpdate_(body);
    if (body.action === "liveDaybookUpdate") return handleLiveDaybookUpdate_(body);
    if (body.action === "logAsk") return handleLogAsk_(body);
    // Master File / Dealer Targets / Remarks / Task Manager writes, added
    // 6 Oct 2026 - see the big comment block above
    // handleDealerTargetsGet_() for the full feature.
    if (body.action === "dealerTargetSet") return handleDealerTargetSet_(body);
    if (body.action === "dealerRemarkAdd") return handleDealerRemarkAdd_(body);
    // Bank Deposit, added 23 Oct 2026 - see handleBankDepositAdd_()/
    // handleBankDepositUpdate_()/handleBankDepositDelete_() above. Touches
    // BankDeposits only, never "Orders", so belongs up here with the
    // other pre-lock actions.
    if (body.action === "bankDepositAdd") return handleBankDepositAdd_(body);
    if (body.action === "bankDepositUpdate") return handleBankDepositUpdate_(body);
    if (body.action === "bankDepositDelete") return handleBankDepositDelete_(body);
    if (body.action === "dealerMetaSet") return handleDealerMetaSet_(body);
    // Dealer Visit Check Yes/No buttons, added 27 Oct 2026 - see
    // handleDealerVisitMatchDecision_() above.
    if (body.action === "dealerVisitMatchDecision") return handleDealerVisitMatchDecision_(body);
    if (body.action === "taskCreate") return handleTaskCreate_(body);
    if (body.action === "taskUpdateStatus") return handleTaskUpdateStatus_(body);
    if (body.action === "taskAddComment") return handleTaskAddComment_(body);
    if (body.action === "taskDelete") return handleTaskDelete_(body);
    // "Not new?" dismiss, added 6 Oct 2026 - see the "Dismissed" header
    // comment on LIVE_NEW_DEALERS_HEADERS above.
    if (body.action === "dismissNewDealer") return handleDismissNewDealer_(body);
    // "Monthly Freeze" - added 1 Oct 2026, called by monthly_archive.py on
    // the last day of the month, right before its PDF capture - see the
    // long comment above handleFreezeMonthSnapshot_() above.
    if (body.action === "freezeMonthSnapshot") return handleFreezeMonthSnapshot_(body);
    // "Auto-archive Master File History" - added 10 Oct 2026, called by
    // monthly_archive.py right after freezeMonthSnapshot finishes for that
    // month - see the long comment above handleArchiveMasterFileHistoryMonth_()
    // for the full feature.
    if (body.action === "archiveMasterFileHistoryMonth") return handleArchiveMasterFileHistoryMonth_(body);
    // "Daily Freeze" - added 2 Oct 2026, called once a day by
    // monthly_archive.py on the office PC (not gated by the month-end
    // check) - see the long comment above handleFreezeDaySnapshot_() above.
    if (body.action === "freezeDaySnapshot") return handleFreezeDaySnapshot_(body);
    // "Tally-based historical recovery" - added 26 Oct 2026, called once
    // by the one-time backfill_week_boundaries.py script on the office
    // PC, never by the normal daily poll - see the long comment above
    // handleBackfillDailyFreezeCategories_() for the full feature.
    if (body.action === "backfillDailyFreezeCategories") return handleBackfillDailyFreezeCategories_(body);
    // Monthly card archive (PDF -> Drive), added 28 Sep 2026 - touches only
    // Drive, never "Orders", so it belongs with the other pre-lock actions
    // above for the same reason they're all up here.
    if (body.action === "archiveMonthlyPdf") return handleArchiveMonthlyPdf_(body);
    // Login/heartbeat/admin-password actions, added 26 Sep 2026 - these touch
    // the Users/LoginLog sheets only, never "Orders", so (like the live-sync
    // actions above) there's no reason for them to wait behind the
    // order-saving lock below.
    if (body.action === "login") return handleLogin_(body);
    if (body.action === "loginWithToken") return handleLoginWithToken_(body);
    if (body.action === "clearRememberToken") return handleClearRememberToken_(body);
    if (body.action === "heartbeat") return handleHeartbeat_(body);
    if (body.action === "setPassword") return handleSetPassword_(body);
    // "Section visibility (by role)", added 1 Oct 2026 - see
    // handleSetSectionVisibility_() above. Touches DashboardSectionVisibility
    // only, never "Orders", so it belongs up here with setPassword rather
    // than behind the order-saving lock below.
    if (body.action === "setSectionVisibility") return handleSetSectionVisibility_(body);
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }

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
