#!/usr/bin/env python3
"""
tally_live_watcher.py — polls TallyPrime's local XML data interface and
pushes changed stock/rate and dealer-outstanding data to the DSR apps'
shared Apps Script backend, so every DSR's phone sees near-real-time
numbers instead of waiting for the next manual Drive-sync cycle.

Run this ONLY on the office PC that has TallyPrime open, during business
hours (Tally has to actually be running and have the company loaded for
this to get anything). Stop it (Ctrl+C) when the office closes — nothing
bad happens if it's not running, the DSR apps just fall back to whatever
they last had (their normal DEFAULT_ITEMS/DEFAULT_DEALERS baked in from
the last Drive sync).

SETUP — do this once:
  1. In TallyPrime: Gateway of Tally → F1 (Help) → Settings → Connectivity
     → make sure "TallyPrime acting as" is set to allow this local XML
     access, and note the port shown there (confirmed 9008 on this office's
     TallyPrime, not the more common default of 9000 — check yours on the
     same screen before running this).
  2. Install Python 3 if not already there, then:  pip install requests
  3. Edit TALLY_PORT below if your Connectivity screen shows a different
     port than 9008.
  4. BACKEND_URL and SYNC_KEY below are already filled in to match the
     Code.gs deployed for these apps — leave them as-is unless a future
     session tells you Code.gs was redeployed with a new key/URL.
  5. Test first, without writing anything live:
       python3 tally_live_watcher.py --test
     This connects to Tally once, prints what it WOULD send for both stock
     and dealer-outstanding, and exits — nothing gets POSTed. Check the
     "outstanding" section especially: pick 2-3 dealers you know the real
     status of (one you know is still owing, one you know just paid off)
     and confirm the printout agrees before ever running it for real.
  6. Once that looks right, run it for real:
       python3 tally_live_watcher.py
     Leave this running (in its own terminal window) during business
     hours. It polls every 10 minutes and only sends what actually
     changed since its last poll.

If dealer-outstanding output looks wrong (a paid-off dealer still shows a
date, or a genuinely-owing dealer shows nothing), STOP and flag it back —
see "Still open" in claude/live-tally-sync-notes.md; the report/collection
name Tally uses for bill-wise outstanding can vary by version/configuration,
and this first cut may need its parsing adjusted against your real data
before it's trusted to run unattended.
"""

import csv
import json
import os
import re
import sys
import time
import traceback
import xml.etree.ElementTree as ET
from datetime import datetime, date, timedelta

import requests

# ── Configuration ────────────────────────────────────────────────────────
TALLY_HOST = "localhost"
TALLY_PORT = 9008  # confirmed 7 Sep 2026 via the office's own Connectivity screen — NOT Tally's more common 9000 default
TALLY_URL = "http://{}:{}".format(TALLY_HOST, TALLY_PORT)

BACKEND_URL = "https://script.google.com/macros/s/AKfycbxppCHNsSlDGVY5gWKC-N2MKJ4wrLOd97FsKC19-SXjg9NPh4myip5Ht4h5R1dDlSl4GQ/exec"
SYNC_KEY = "374e613bdcdf217cb4948e0560f3773ad07fa07244cad32e3ca5cb20e648e634"  # must exactly match LIVE_SYNC_KEY in Code.gs (rotated 30 Oct 2026, security fix)

POLL_SECONDS = 600  # 10 minutes — matches LIVE_OVERLAY_POLL_MS in the DSR app files

STATE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tally_live_state.json")

# ── Tally XML requests ───────────────────────────────────────────────────
# Tally's local XML interface accepts a plain HTTP POST of an XML request
# envelope to its root URL and returns an XML response. The "Collection"
# request type below is Tally's standard programmatic-export mechanism
# (Tally Developer's TDL COLLECTION object) — more reliable for integration
# than screen-scraping a formatted report.

# FETCH list updated 11 Sep 2026: added STANDARDPRICE alongside the existing
# fields (CLOSINGRATE/CLOSINGBALANCE/CLOSINGVALUE kept for the diagnostic
# comparison in dump_price_check() below, not removed). STANDARDPRICE is the
# real selling-rate field, confirmed by reading the office's own TDL source
# for the "AAA MDU Dealer PL" report they already trust
# (Petronas_Madurai_DealerPriceList 6.tdl) - its NDLP column is computed as
# `$StandardPrice * 1.18`, i.e. exactly the pre-GST selling rate this app
# needs. This is NOT the same field as `STANDARDSELLINGPRICE`, which was
# tried first (via --dump-mrp, see fetch_stock()'s docstring below) and came
# back completely absent from Tally's response - a different, non-existent
# field name, not just an empty one. `StandardPrice` in TDL syntax maps to
# the `STANDARDPRICE` XML tag below, following the same TDL-field-name-to-
# XML-tag convention already confirmed for CLOSINGBALANCE/CLOSINGRATE/
# CLOSINGVALUE - not yet independently re-confirmed for this specific tag
# via a raw XML dump, hence dump_price_check() below before this goes live.
STOCK_COLLECTION_XML = """<ENVELOPE>
 <HEADER>
  <VERSION>1</VERSION>
  <TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE>
  <ID>StockItemLiveCollection</ID>
 </HEADER>
 <BODY>
  <DESC>
   <STATICVARIABLES>
    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
   </STATICVARIABLES>
   <TDL>
    <TDLMESSAGE>
     <COLLECTION NAME="StockItemLiveCollection" ISMODIFY="No">
      <TYPE>StockItem</TYPE>
      <FETCH>NAME, CLOSINGBALANCE, CLOSINGRATE, CLOSINGVALUE, STANDARDPRICE</FETCH>
      <!-- MRPDETAILS.LIST added 11 Sep 2026 - confirmed via --dump-mrp-live's
           raw dump against 2 known items (Sprinta F700 697, Urania 800 5018,
           both matched exactly using the LAST MRPDETAILS.LIST entry's nested
           MRPRATEDETAILS.LIST/MRPRATE) - see fetch_stock()'s docstring and
           _find_mrp_rate() below for how this is parsed. -->
      <FETCH>MRPDETAILS.LIST</FETCH>
     </COLLECTION>
    </TDLMESSAGE>
   </TDL>
  </DESC>
 </BODY>
</ENVELOPE>"""

# Bill-wise outstanding via each Ledger's own BillAllocations - this is the
# standard way to get "which bills are still actually open" rather than just
# a running balance, so a dealer whose balance nets to zero (or whose open
# bills are all settled) correctly comes back with nothing outstanding, even
# if their ledger has plenty of historical voucher activity.
OUTSTANDING_COLLECTION_XML = """<ENVELOPE>
 <HEADER>
  <VERSION>1</VERSION>
  <TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE>
  <ID>LedgerOutstandingCollection</ID>
 </HEADER>
 <BODY>
  <DESC>
   <STATICVARIABLES>
    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
   </STATICVARIABLES>
   <TDL>
    <TDLMESSAGE>
     <COLLECTION NAME="LedgerOutstandingCollection" ISMODIFY="No">
      <TYPE>Ledger</TYPE>
      <FETCH>NAME, CLOSINGBALANCE</FETCH>
      <FETCH>BILLALLOCATIONS.LIST</FETCH>
     </COLLECTION>
    </TDLMESSAGE>
   </TDL>
  </DESC>
 </BODY>
</ENVELOPE>"""

# ── Diagnostic-only, added 8 Sep 2026 ────────────────────────────────────
# The Collection-based approach above returns EVERY bill allocation ever
# recorded against a ledger, with no reliable field distinguishing an
# already-fully-settled old bill from a genuinely still-pending one - this
# was confirmed the hard way (Akila Auto Spares' Collection data included a
# fully-paid Feb-2024 bill with no way to tell it was closed). Tally's own
# "Bills Receivable" report (Gateway of Tally -> Display More Reports ->
# Statement of Accounts -> Outstandings -> Receivables - the same
# underlying data as the "Ledger Voucher Outstanding" screen the office
# checked manually) already does exactly this filtering server-side, so
# exporting THAT native report directly - rather than trying to
# reconstruct settlement status from raw bill records - should sidestep
# the problem entirely. This is a NATIVE REPORT export (TALLYREQUEST =
# "Export Data" + REPORTNAME), a different request shape than the
# Collection/TDL requests above - not yet confirmed to work against this
# Tally version, hence --dump-report below to check its shape BEFORE
# wiring it into fetch_outstanding().
BILLS_RECEIVABLE_REPORT_XML = """<ENVELOPE>
 <HEADER>
  <TALLYREQUEST>Export Data</TALLYREQUEST>
 </HEADER>
 <BODY>
  <EXPORTDATA>
   <REQUESTDESC>
    <REPORTNAME>Bills Receivable</REPORTNAME>
    <STATICVARIABLES>
     <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
    </STATICVARIABLES>
   </REQUESTDESC>
  </EXPORTDATA>
 </BODY>
</ENVELOPE>"""

# Payment-reminder feature, added 19 Sep 2026 (Arun-only trial) - a bill
# only counts toward a dealer's "reminder-worthy" aged amount once it's
# been open this many days. See fetch_outstanding()'s agedPendingAmount.
OUTSTANDING_REMINDER_AGE_DAYS = 45


# Tally sometimes encodes stray control characters that show up inside a
# ledger/item name (typically pasted in from somewhere else) as literal
# numeric XML character references - e.g. "&#5;" - in its exported XML.
# Most of those code points (anything below 0x20 except tab/newline/CR) are
# not valid XML 1.0 characters at all, so Python's XML parser refuses the
# WHOLE response with "reference to invalid character number" the moment it
# hits even one, no matter how large or otherwise well-formed the rest of
# the document is. Found 8 Sep 2026 via --dump-groups, which is the first
# request that pulls literally every ledger name in the company (4,140 of
# them) in one go - one of those names apparently contains such a
# character. Stripping just the invalid numeric references (not touching
# anything else) fixes this without needing to know which ledger it is.
_INVALID_NUMERIC_ENTITY_RE = re.compile(r"&#(x?[0-9a-fA-F]+);")


def _is_valid_xml_codepoint(cp):
    return (
        cp in (0x9, 0xA, 0xD)
        or 0x20 <= cp <= 0xD7FF
        or 0xE000 <= cp <= 0xFFFD
        or 0x10000 <= cp <= 0x10FFFF
    )


# Tally sometimes emits a tag whose own name contains a colon - most often
# a custom/UDF field, e.g. "<UDF:SomeField>" - which is syntactically a
# valid XML namespace-prefixed name, but since no xmlns declaration for
# that prefix exists anywhere in Tally's export, Python's strict parser
# rejects the WHOLE response with "unbound prefix" the instant it hits one.
# Found 8 Sep 2026 via --dump-eligible's full voucher-history pull - the
# first request broad enough to include a voucher carrying one of these.
# None of the tags this file actually reads (VOUCHER, DATE, PARTYLEDGERNAME,
# ALLLEDGERENTRIES.LIST, LEDGERNAME, STOCKITEM, LEDGER, GROUP, BILLFIXED,
# ...) ever contain a colon themselves, so it's safe to fold any
# "prefix:name" tag into "prefix_name" - the field's own content is never
# read, only the document around it needs to parse.
_NAMESPACED_TAG_RE = re.compile(r"(</?)([A-Za-z_][\w.\-]*):([A-Za-z_][\w.\-]*)")


def _sanitize_xml_text(text):
    def repl(m):
        raw = m.group(1)
        try:
            cp = int(raw[1:], 16) if raw[:1] in ("x", "X") else int(raw)
        except ValueError:
            return ""  # malformed reference entirely - drop it, don't guess
        return m.group(0) if _is_valid_xml_codepoint(cp) else ""
    text = _INVALID_NUMERIC_ENTITY_RE.sub(repl, text)
    text = _NAMESPACED_TAG_RE.sub(lambda m: m.group(1) + m.group(2) + "_" + m.group(3), text)
    return text


# Default bumped 30s -> 60s on 16 Sep 2026. This is the timeout every regular
# (non-bulk) poll uses - stock and outstanding, every ~10 minutes - talking to
# Tally's own local XML export port, not any network/Google call. The owner's
# real watcher log showed 4 "Read timed out (read timeout=30)" errors in
# about 1.5 hours on 16 Sep (vs. 1 the day before), all during what's likely
# a busy billing stretch - Tally was just slow to answer within 30s, not
# actually failing. Nothing was lost either way (a timed-out poll always
# retries and catches up on the very next cycle, unmodified), so this is
# purely to cut down how often a poll gives up prematurely - matches the 60s
# already used elsewhere in this file (post_backend()) for the same reason.
def tally_request(xml_body, timeout=60):
    resp = requests.post(TALLY_URL, data=xml_body.encode("utf-8"),
                          headers={"Content-Type": "text/xml"}, timeout=timeout)
    resp.raise_for_status()
    return _sanitize_xml_text(resp.text)


def _text(el, tag, default=None):
    child = el.find(tag)
    return child.text if child is not None and child.text is not None else default


def _num(val):
    if val is None:
        return None
    s = str(val).strip().replace(",", "")
    if s == "":
        return None
    # Tally often suffixes quantities with a unit, e.g. "218 Nos" - keep only
    # the leading numeric part.
    num = ""
    for ch in s:
        if ch.isdigit() or ch in ".-":
            num += ch
        elif num:
            break
    try:
        return float(num) if num not in ("", "-", ".") else None
    except ValueError:
        return None


def fetch_stock():
    """Returns {item_name: {"stock": float|None, "rate": float|None, "mrp": float|None}}.

    RATE HISTORY (all 11 Sep 2026):
    1. Originally returned CLOSINGRATE as "rate". CONFIRMED WRONG - Tally's
       CLOSINGRATE is a stock VALUATION figure (closing stock value /
       closing quantity, a cost/accounting basis), not a selling or billing
       rate, but the app treats "rate" as the pre-GST SELLING rate
       (displays Math.round(rate * 1.18) as "Rate incl GST"). Pushing
       CLOSINGRATE in meant every item's displayed price was quietly wrong
       (too low, by roughly the item's margin) - the owner caught this by
       cross-checking the live app against Daily Stock Update.xlsx/Item
       Catalog (both correct) for multiple items at once, all wrong the
       same way.
    2. Emergency fix: rate forced to None here (stop sending a bad figure)
       and the app-side line consuming `rate` from the live overlay was
       commented out in all 8 apps (see dsr-app-registry.md) - live price
       fully off, apps fall back to their last Drive-synced rate.
    3. Field hunt: first guess `STANDARDSELLINGPRICE` (via --dump-mrp) came
       back completely ABSENT from Tally's response - not blank, the field
       doesn't exist under that name on this Tally setup. The owner then
       supplied the actual TDL source for the office's own trusted
       "AAA MDU Dealer PL" report (Petronas_Madurai_DealerPriceList 6.tdl),
       whose NDLP column computes as `$StandardPrice * 1.18` - i.e. the
       real TDL field is `StandardPrice`, mapping to the `STANDARDPRICE`
       XML tag.
    4. CONFIRMED, 11 Sep 2026: `--dump-price` (STOCK_COLLECTION_XML, which
       now fetches STANDARDPRICE) matched all 4 known-correct items exactly
       - Sprinta F700 438.14, Urania 800 2817.8, Syntium 500 210L 57694.07,
       Syntium 500 CI4+ Diesel 338.98 - and STANDARDPRICE is populated for
       8594 of 8797 total stock items (97.7% coverage). Same verification
       discipline fetch_outstanding() went through against Akila/Challenger/
       A2Z before being trusted. Now wired into the real `rate` output below.
       The app-side consuming line must be re-enabled to match (see
       live-tally-sync-notes.md / dsr-app-registry.md for that half of the
       fix) - both sides were disabled together and must come back together.
       For the ~3% of items with no STANDARDPRICE, `rate` comes back None
       here exactly like before, so those items simply keep their last
       Drive-synced rate (same graceful-fallback behaviour as always) -
       nothing crashes or shows blank.

    MRP HISTORY (11 Sep 2026, later same day): the owner asked whether the
    same TDL file also gives a live MRP field, explicitly wanting it kept
    ISOLATED from the rate/price work above. Tally models MRP as a
    two-level nested list (an item can have several MRPs over time, each
    with its own nested rate-details list) - `$MRPDetails[Last].
    MRPRateDetails[Last].MRPRate` in TDL syntax. An isolated diagnostic
    (`--dump-mrp-live`, its own separate Collection request, untouched by
    anything above) fetched `MRPDETAILS.LIST` and confirmed the structure
    against 2 known-correct items straight from the raw XML: Sprinta F700's
    LAST `MRPDETAILS.LIST` entry (by document order, which is chronological
    - FROMDATE 20260716) held `MRPRATEDETAILS.LIST/MRPRATE` = 697.00,
    matching the catalog's known-correct 697 exactly; Urania 800's last
    entry (FROMDATE 20260909) held 5018.00, matching 5018 exactly. Now
    wired into the real `mrp` output below via `_find_mrp_rate()` (defined
    further down, alongside the diagnostic). No app-side change was needed
    for this one - unlike `rate`, the app's `upd.mrp` consumption line in
    `applyLiveOverlay_()` was never disabled, so all 8 apps already pick up
    a live `mrp` the moment this function starts sending one.
    """
    xml_text = tally_request(STOCK_COLLECTION_XML)
    root = ET.fromstring(xml_text)
    out = {}
    for si in root.iter("STOCKITEM"):
        name = si.get("NAME") or _text(si, "NAME")
        if not name:
            continue
        mrp_val, _mrp_why = _find_mrp_rate(si)
        out[name.strip()] = {
            "stock": _num(_text(si, "CLOSINGBALANCE")),
            "rate": _num(_text(si, "STANDARDPRICE")),  # RE-ENABLED 11 Sep 2026 - confirmed correct field, see docstring above
            "mrp": mrp_val,  # RE-ENABLED 11 Sep 2026 - confirmed via MRPDETAILS.LIST, see docstring above
        }
    return out


def _bill_date_to_iso(tally_date):
    """Tally bill dates typically come back as YYYYMMDD (e.g. 20260829)."""
    if not tally_date:
        return None
    s = str(tally_date).strip()
    if len(s) == 8 and s.isdigit():
        return "{}-{}-{}".format(s[0:4], s[4:6], s[6:8])
    return None


def _billdate_ddmonyy_to_iso(tally_date):
    """Bills Receivable report dates come back as 'DD-Mon-YY' e.g. '29-Jun-26'."""
    if not tally_date:
        return None
    try:
        return datetime.strptime(str(tally_date).strip(), "%d-%b-%y").strftime("%Y-%m-%d")
    except ValueError:
        return None


def fetch_all_ledgers_with_balance():
    """
    Fetches every Ledger's name + current CLOSINGBALANCE in one pass
    (OUTSTANDING_COLLECTION_XML already fetches CLOSINGBALANCE for the
    bill-wise request below - previously only the name was kept from it).
    Returns {name: balance_float_or_None}, negative = Dr (owed), positive =
    Cr (credit/advance) - the same sign convention used everywhere else in
    this file.

    Serves two purposes: (1) the full universe of dealer names, so a dealer
    with NOTHING pending (never appears in the Bills Receivable report at
    all) can still be reported as explicitly "cleared" rather than silently
    omitted - see fetch_outstanding(); (2) added 8 Sep 2026, each dealer's
    live current balance, pushed to "Manage dealers" in the app so it stays
    accurate the moment a bill or receipt posts in Tally - a Dr balance
    shows as "<amount> due", a Cr balance as "<amount> credit".
    """
    xml_text = tally_request(OUTSTANDING_COLLECTION_XML)
    root = ET.fromstring(xml_text)
    out = {}
    for ledger in root.iter("LEDGER"):
        name = ledger.get("NAME") or _text(ledger, "NAME")
        if not name:
            continue
        out[name.strip()] = _num(_text(ledger, "CLOSINGBALANCE"))
    return out


def fetch_all_ledger_names():
    """Thin wrapper kept for anything that only needs the name set."""
    return set(fetch_all_ledgers_with_balance().keys())


# ── Dealer credit-limit diagnostic, added 30 Sep 2026 ───────────────────
# New feature: the Dealer Lookup card on the dashboard (and, later, each
# DSR app) should show a dealer's current credit-lock/credit-limit status
# alongside their Outstanding, so a DSR sees it before promising an order.
# The owner already worked out the actual credit-limit FORMULA separately
# (a payment-ratio-tiered rule, "Lock D", tracked in a billing-lock
# workbook - see /areas/dealer-billing-lock.md) and, per his own explicit
# decision, has been pushing that computed number into Tally's OWN native
# per-ledger "Credit Limit" field (Alter Ledger / native XML Import Data),
# rather than this project building any lock logic of its own. His
# instruction for this feature: "Take credit limit from tally" - read it
# live, straight from Tally, the exact same way Outstanding/Stock/MRP
# already are here, NOT from the separate billing-lock Google Sheet.
#
# CREDITLIMIT is the standard Tally Ledger-master field for this (TDL
# field `$CreditLimit`, following the same TDL-name-to-XML-tag convention
# already confirmed for CLOSINGBALANCE/CLOSINGRATE/STANDARDPRICE) - but,
# same as STANDARDSELLINGPRICE turning out not to exist under that guessed
# name (the real field was STANDARDPRICE), this has NEVER been checked
# against this office's real Tally data, so nothing here is wired into any
# live push yet. CREDITPERIOD (the credit period in days, a different,
# related ledger setting) is fetched alongside it purely for context/in
# case it's useful later - not used by this feature at all right now.
#
# Isolated in its own dedicated Collection request (own COLLECTION NAME,
# not reusing LEDGER_GROUP_COLLECTION_XML/OUTSTANDING_COLLECTION_XML),
# same "a new/unverified field never risks an already-trusted request"
# convention as STOCK_MRP_COLLECTION_XML being kept separate from
# STOCK_COLLECTION_XML.
CREDIT_LIMIT_DUMP_XML = """<ENVELOPE>
 <HEADER>
  <VERSION>1</VERSION>
  <TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE>
  <ID>LedgerCreditLimitDumpCollection</ID>
 </HEADER>
 <BODY>
  <DESC>
   <STATICVARIABLES>
    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
   </STATICVARIABLES>
   <TDL>
    <TDLMESSAGE>
     <COLLECTION NAME="LedgerCreditLimitDumpCollection" ISMODIFY="No">
      <TYPE>Ledger</TYPE>
      <FETCH>NAME, PARENT, CLOSINGBALANCE, CREDITLIMIT, CREDITPERIOD</FETCH>
     </COLLECTION>
    </TDLMESSAGE>
   </TDL>
  </DESC>
 </BODY>
</ENVELOPE>"""


def dump_creditlimit(dealer_fragment=None):
    """
    --dump-creditlimit [dealer or DSR fragment]

    Diagnostic only - pushes nothing anywhere, no fetch_creditlimit()
    exists yet. Fetches NAME/PARENT/CLOSINGBALANCE/CREDITLIMIT/CREDITPERIOD
    for every ledger in the company, saves the raw response, and reports:

    1. Coverage - how many of the total ledgers have a non-blank
       CREDITLIMIT at all (mirrors dump_price_check()'s STANDARDPRICE
       coverage check and dump_mrp_live()'s MRP coverage check - if this
       comes back 0 of however-many, that's the same "wrong field name"
       signature every earlier guessed field in this project has shown at
       least once, e.g. STANDARDSELLINGPRICE).
    2. One row per dealer in SAMPLE_DEALERS_BY_DSR (or, if
       dealer_fragment is given, every ledger whose name contains that
       fragment instead) - name, resolved DSR (for the sample list),
       parsed CREDITLIMIT, parsed CREDITPERIOD, and current
       CLOSINGBALANCE for context - so the owner can directly compare
       against a dealer whose real Tally credit limit he already knows
       (e.g. one he entered Lock D's number into himself) before this
       goes anywhere near production.
    3. The full raw <LEDGER>...</LEDGER> block for the first matched
       dealer, unparsed - so if CREDITLIMIT turns out to be the wrong tag
       name (blank/missing for a dealer known to have one set), the real
       tag name is visible directly in Tally's own response rather than
       guessed at again.

    Nothing here assumes Lock D has actually been pushed into Tally for
    every one of the 8 dealer areas yet - the coverage count in point 1
    is exactly how that gets confirmed one way or the other, per the
    owner's own open question about how far that push has gotten.
    """
    out_dir = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(out_dir, "tally_creditlimit_dump.xml")

    print("Fetching NAME, PARENT, CLOSINGBALANCE, CREDITLIMIT, CREDITPERIOD for every ledger...")
    xml_text = tally_request(CREDIT_LIMIT_DUMP_XML)
    with open(path, "w", encoding="utf-8") as f:
        f.write(xml_text)
    print("Saved full response to: {}  ({} bytes)".format(path, len(xml_text)))

    root = ET.fromstring(xml_text)
    by_name = {}
    have_creditlimit = 0
    total = 0
    for led in root.iter("LEDGER"):
        name = led.get("NAME") or _text(led, "NAME")
        if not name:
            continue
        name = name.strip()
        total += 1
        cl = _num(_text(led, "CREDITLIMIT"))
        cp = _num(_text(led, "CREDITPERIOD"))
        bal = _num(_text(led, "CLOSINGBALANCE"))
        parent = (_text(led, "PARENT") or "").strip()
        if cl is not None:
            have_creditlimit += 1
        by_name[name] = {"creditLimit": cl, "creditPeriod": cp, "balance": bal, "parent": parent}

    print("\n" + "=" * 70)
    print("Coverage: {} of {} total ledgers have a non-blank CREDITLIMIT.".format(
        have_creditlimit, total))
    print("=" * 70)
    if have_creditlimit == 0 and total > 0:
        print("ZERO ledgers have a non-blank CREDITLIMIT - this is the same signature every")
        print("wrong field-name guess in this project has shown before (e.g.")
        print("STANDARDSELLINGPRICE). Either CREDITLIMIT isn't the real tag name on this")
        print("Tally setup, or Lock D genuinely hasn't been pushed into Tally's Credit Limit")
        print("field for any dealer yet - the raw excerpt below should help tell which.")

    print("\n" + "=" * 70)
    if dealer_fragment:
        frag_upper = dealer_fragment.strip().upper()
        rows = [(n, r, "") for n, r in sorted(by_name.items()) if frag_upper in n.upper()]
        print("Ledgers matching '{}':".format(dealer_fragment))
    else:
        dsr_by_name = dict((n, d) for d, n in SAMPLE_DEALERS_BY_DSR)
        rows = []
        for dsr, target in SAMPLE_DEALERS_BY_DSR:
            rec = by_name.get(target)
            if rec is not None:
                rows.append((target, rec, dsr))
                continue
            # fallback: substring match, same convention as dump_groups()
            found = None
            for name in by_name:
                if target.upper() in name.upper():
                    found = name
                    break
            if found:
                rows.append(("{} (matched '{}')".format(target, found), by_name[found], dsr))
            else:
                rows.append((target, None, dsr))
        print("One known dealer per DSR (SAMPLE_DEALERS_BY_DSR):")
    print("=" * 70)
    if not rows:
        print("  No dealer matched.")
    for name, rec, dsr in rows:
        label = "  {}".format(name)
        if dsr:
            label += "   [DSR: {}]".format(dsr)
        print(label)
        if rec is None:
            print("      NOT FOUND in Tally's ledger list")
            continue
        print("      CREDITLIMIT={}   CREDITPERIOD={}   CLOSINGBALANCE={}   Group(PARENT)={}".format(
            rec["creditLimit"], rec["creditPeriod"], rec["balance"], rec["parent"] or "(blank)"))

    # Raw excerpt of the first matched dealer's whole <LEDGER> block, so
    # the real tag names are visible directly - same fallback used
    # throughout this file (dump_xml's Challenger Motors block, dump_mrp's
    # raw STOCKITEM blocks) whenever a guessed field name might be wrong.
    first_real_name = None
    for name, rec, _dsr in rows:
        if rec is not None:
            first_real_name = name.split(" (matched ")[0] if " (matched " not in name else name.split("'")[1]
            break
    if first_real_name:
        idx = xml_text.find('NAME="{}"'.format(first_real_name))
        if idx == -1:
            idx = xml_text.upper().find(first_real_name.upper())
        if idx != -1:
            block_start = xml_text.rfind("<LEDGER", 0, idx)
            block_end = xml_text.find("</LEDGER>", idx)
            block_end = block_end + len("</LEDGER>") if block_end != -1 else idx + 2000
            if block_start == -1:
                block_start = max(0, idx - 200)
            print("\n" + "=" * 70)
            print("EASY-COPY RAW EXCERPT - full <LEDGER> block for '{}':".format(first_real_name))
            print("=" * 70)
            print(xml_text[block_start:block_end])

    print("\n" + "=" * 70)
    print("Send this whole output back, plus 2-3 real credit-limit figures you already")
    print("know for specific dealers (e.g. ones you personally pushed Lock D's number into")
    print("for), so CREDITLIMIT can be confirmed correct - or, if it's wrong/blank, the raw")
    print("excerpt above should show whatever the real field is called instead. Nothing")
    print("from this diagnostic is wired into the live sync or the dashboard yet.")


def fetch_creditlimit():
    """
    Real (non-diagnostic) dealer credit-limit fetch. CONFIRMED against real
    Tally data 30 Sep 2026: --dump-creditlimit's raw output matched the
    owner's own independently-known real Lock D figures exactly for two
    dealers (N.S.K. AUTOMOBILES, THENI (OIL) = 587222.0; RAMYA AUTO STORES,
    CBM (OIL) = 60880.0), with the raw <LEDGER> excerpt confirming
    CREDITLIMIT is the correct native Tally field (not a guess, unlike
    STANDARDSELLINGPRICE earlier in this project). Real coverage at
    confirmation time: 559 of 4152 total ledgers had a non-blank
    CREDITLIMIT (Lock D not yet pushed to every dealer - expected, not a
    bug; see dump_creditlimit()'s own docstring above).

    Reuses CREDIT_LIMIT_DUMP_XML - the same isolated Collection request the
    diagnostic above uses. Never folded into LEDGER_GROUP_COLLECTION_XML or
    OUTSTANDING_COLLECTION_XML, same "a new field never risks an
    already-trusted request" convention used throughout this file.

    Returns {ledger_name: {"creditLimit": float_or_None, "creditPeriod":
    float_or_None}} for EVERY ledger in the company - same whole-company
    scope as fetch_all_ledgers_with_balance()/fetch_outstanding() (not
    filtered to known DSR dealers), so a ledger whose CREDITLIMIT is
    cleared in Tally is still reported as an explicit None rather than
    silently dropping out of the diff - same "explicit clear beats silent
    skip" convention diff_outstanding() already follows.
    """
    xml_text = tally_request(CREDIT_LIMIT_DUMP_XML)
    root = ET.fromstring(xml_text)
    out = {}
    for led in root.iter("LEDGER"):
        name = led.get("NAME") or _text(led, "NAME")
        if not name:
            continue
        name = name.strip()
        out[name] = {
            "creditLimit": _num(_text(led, "CREDITLIMIT")),
            "creditPeriod": _num(_text(led, "CREDITPERIOD")),
        }
    return out


def diff_creditlimit(prev, current):
    """
    Same shape/convention as diff_outstanding() above - a ledger whose
    creditLimit OR creditPeriod changed since the last poll (including one
    clearing to None) counts as changed and gets pushed.
    """
    changed = []
    for name, rec in current.items():
        old = prev.get(name, "__unset__")
        if old != rec:
            changed.append({
                "name": name,
                "creditLimit": rec["creditLimit"],
                "creditPeriod": rec["creditPeriod"],
            })
    return changed


def fetch_outstanding(all_ledger_balances=None, capture_bills_into=None):
    """
    Returns {dealer_name: {"since": iso_date_or_None, "pendingAmount": float_or_None}}.
    (Changed 17 Sep 2026 from a bare date value to this dict - see the
    BILLCL addition in the docstring below.)

    Takes an optional pre-fetched {name: balance} dict (from
    fetch_all_ledgers_with_balance()) to use as the full dealer-name
    universe, so a caller that already fetched balances this poll (see
    run_once(), added 8 Sep 2026 for the "Manage dealers" balance display)
    doesn't trigger a second, redundant Tally request just to get the same
    name list again. Fetches its own if not given one.

    THIRD REWRITE, 8 Sep 2026 - the previous two attempts both tried to
    reconstruct "is this bill still genuinely pending" from raw
    BILLALLOCATIONS.LIST records pulled via a Ledger Collection request,
    and both got real dealers wrong (round 1: picked up an old credit note
    as if it were a debt; round 2: picked up an old bill that had since
    been fully paid off, because OPENINGBALANCE only reflects a bill's
    amount when raised, never whether it's since been settled - Tally
    simply never returned a "still owes this much on this specific bill"
    field via that request, no matter what was tried).

    This version stops reconstructing settlement status entirely and
    instead asks Tally for its own **"Bills Receivable"** report
    (`BILLS_RECEIVABLE_REPORT_XML` above - the same underlying data as the
    "Ledger Voucher Outstanding" screen the office checks manually), which
    already does this filtering server-side - it only ever lists bills
    that are CURRENTLY still pending, full stop. Confirmed against real
    output 8 Sep 2026: Akila Auto Spares' section correctly contained only
    its genuinely-open 29-Jun-26 bill - no Jul-2025 credit note, no
    Feb-2024 settled-but-still-Dr bill - matching the owner's own
    confirmed correct answer exactly.

    The report's XML is a FLAT, repeating sequence with no common wrapper
    per row - each bill is a <BILLFIXED> element (containing <BILLDATE>,
    <BILLREF>, <BILLPARTY>) immediately followed by sibling <BILLCL>
    (closing/pending amount - confirmed matches the "Pending Amount"
    column exactly, e.g. -8685.00 for Akila's 29-Jun-26 bill),
    <BILLDUE>, and <BILLOVERDUE> elements - NOT nested inside <BILLFIXED>.
    Parsed below by walking the whole tree in document order and grouping
    each <BILLFIXED> with whatever <BILLCL>/<BILLDUE>/<BILLOVERDUE>
    immediately follow it, up to the next <BILLFIXED>.

    Because this report only lists dealers with something still pending,
    "Outstanding Since" per dealer = the earliest <BILLDATE> among their
    entries here. A dealer with NO entries at all is confidently NOT
    outstanding - but since they never appear in this report either way,
    fetch_all_ledger_names() is used separately to get the full name
    universe, so "never appeared" can still be reported as an explicit
    clear rather than silently doing nothing for them.

    BILLCL added 17 Sep 2026 (trial run, Arun's app only for now - see
    dsr-app-registry.md's standing "new/unproven feature -> one DSR first"
    scope rule): this is each bill's own CURRENT pending amount, as opposed
    to a bill's OPENINGBALANCE (its amount when first raised, which rounds
    1/2 of this whole rewrite got wrong precisely because it never reflects
    later part-payment). Confirmed against real data 8 Sep 2026 building
    this same report: -8685.00 for Akila's 29-Jun-26 bill, matching the
    owner's own "Pending Amount" screenshot column exactly - same
    negative=Dr/positive=Cr sign convention used everywhere else in this
    file. Summed per dealer across all their entries in this report (which,
    by construction, is already restricted to currently-open bills only) to
    give a rupee-precise "amount actually pending right now" per dealer -
    genuinely different from fetch_all_ledgers_with_balance()'s
    CLOSINGBALANCE, which is the dealer's WHOLE ledger balance (could in
    principle include non-bill-tracked entries the Bills Receivable report
    wouldn't catch). The two are expected to usually agree; the owner
    comparing them side by side on Arun's real dealers is this trial's way
    of confirming that for real before deciding whether/how to build on top
    of it, same verification discipline as every other Tally field in this
    project.

    agedPendingAmount added 19 Sep 2026, for the payment-reminder feature
    (Arun-only trial, same scope rule): the portion of pendingAmount whose
    OWN bill is more than OUTSTANDING_REMINDER_AGE_DAYS (45) days old,
    computed per-bill from each bill's own BILLDATE against today - not a
    single dealer-wide "since" date, since a dealer can have several open
    bills of different ages and only the genuinely old ones should count
    toward a reminder. Each BILLFIXED is now paired with its OWN following
    BILLCL directly (one bill dict per bill), instead of growing two
    separate per-dealer lists (dates, amounts) and trusting them to stay in
    the same order - a dealer with a BILLFIXED that's missing its BILLCL
    for any reason could otherwise have silently misaligned every bill
    after it. Only bills whose amount is negative (Dr/owed, same guard used
    everywhere else in this project, e.g. Needs Attention's `r.raw < 0`)
    count toward agedPendingAmount, in case the Bills Receivable report
    ever includes a stray Cr entry.

    band3059/band6089/band90plus added 22 Sep 2026, fixing a real bug the
    owner flagged with a real example: BS Enterprises had a ₹33,336 bill
    raised TODAY, a ₹13,955 bill 59 days old, and a ₹3,000 bill 73 days
    old - but every existing consumer of this function's old {since,
    pendingAmount} pair (the dashboard's Needs Attention stat tiles, its
    DSR-wise age chart, and Dealer Lookup's "Outstanding" tile) attributed
    the dealer's WHOLE pendingAmount (₹50,291) to a single age bucket based
    only on `since` (the OLDEST bill's date), so every one of those showed
    "₹50,291 (73d)" - implying the entire amount was 73 days overdue, when
    only ₹3,000 of it actually was. These three fields split pendingAmount
    the same way agedPendingAmount already does - per BILL, by that bill's
    own age - into the same 30-59/60-89/90+ non-overlapping bands the
    dashboard's age chart already uses client-side, so any consumer that
    sums them gets a genuinely age-accurate figure no matter how many open
    bills of different ages a dealer has. Same debit-only guard as
    agedPendingAmount (a stray Cr entry never counts toward any band), and
    same "None when nothing fell in this band" convention as
    agedPendingAmount (not 0.0), so the sheet/JSON round-trip's existing
    null-handling stays uniform across all three "amount in a sub-bucket"
    fields.

    band0_29 added 23 Sep 2026, for the dashboard's DSR-wise aging TABLE
    (replacing the DSR-wise aging chart) - the owner wanted a 0-29 day
    band shown alongside 30-59/60-89/90+ so the table has a genuinely
    complete picture, not just the 30+ "needs attention" slice. Same
    per-bill, debit-only, "None not 0.0 when empty" convention as the
    other three bands.

    capture_bills_into added 23 Sep 2026, for the 60+ day overdue
    "collected this month" dashboard stat - fixed from a daily-recomputed
    figure (which drifted because which bills count as "60+" changes
    every day) to a frozen, BILL-LEVEL monthly baseline (the owner's own
    explicit choice over a simpler dealer-cash-level measure, to avoid a
    dealer's new-purchase payment being misread as old-debt collection).
    When given a dict, this function fills it (side effect) with
    {dealer_name: [{"billref": ..., "date": ..., "amount": ...}, ...]}
    for every dealer's every currently-open bill (not just the aged
    ones - the whole bills_by_party structure this function already
    builds for itself), so the caller can freeze/compare specific bills
    by BILLREF+date+amount. BILLREF newly parsed from BILLFIXED here
    (previously unused) - confirmed present on every BILLFIXED element
    alongside BILLPARTY/BILLDATE, same report, no new Tally request.
    """
    xml_text = tally_request(BILLS_RECEIVABLE_REPORT_XML)
    root = ET.fromstring(xml_text)

    bills_by_party = {}
    current_party = None
    current_date_iso = None
    current_billref = None
    for el in root.iter():
        if el.tag == "BILLFIXED":
            party = _text(el, "BILLPARTY")
            current_party = party.strip() if party else None
            current_date_iso = _billdate_ddmonyy_to_iso(_text(el, "BILLDATE"))
            billref_raw = _text(el, "BILLREF")
            current_billref = billref_raw.strip() if billref_raw else None
        elif el.tag == "BILLCL":
            amt = _num(el.text)
            if current_party and current_date_iso and amt is not None:
                bills_by_party.setdefault(current_party, []).append(
                    {"date": current_date_iso, "amount": amt, "billref": current_billref}
                )
        elif el.tag in ("BILLDUE", "BILLOVERDUE"):
            pass  # not currently used - Tally's own due-date/overdue-days, kept in mind for later

    all_names = all_ledger_balances.keys() if all_ledger_balances is not None else fetch_all_ledger_names()

    today_iso = datetime.now().strftime("%Y-%m-%d")

    def _record(name):
        bills = bills_by_party.get(name)
        if not bills:
            return {"since": None, "pendingAmount": None, "agedPendingAmount": None,
                     "band0_29": None, "band3059": None, "band6089": None, "band90plus": None}
        dates = [b["date"] for b in bills]
        amounts = [b["amount"] for b in bills]
        aged_sum = 0.0
        has_aged = False
        band0_29_sum = 0.0
        band3059_sum = 0.0
        band6089_sum = 0.0
        band90plus_sum = 0.0
        has_band0_29 = False
        has_band3059 = False
        has_band6089 = False
        has_band90plus = False
        for b in bills:
            if b["amount"] >= 0:
                continue  # not a genuine debit - never counts toward the aged/reminder figure
            age_days = (datetime.strptime(today_iso, "%Y-%m-%d") - datetime.strptime(b["date"], "%Y-%m-%d")).days
            if age_days > OUTSTANDING_REMINDER_AGE_DAYS:
                aged_sum += b["amount"]
                has_aged = True
            # Age bands - see this function's docstring ("band3059/..."
            # section) for why these exist. Boundaries match the
            # dashboard's own existing 30-59/60-89/90+ bands exactly
            # (non-overlapping, a bill under 30 days old falls in none of
            # them - it's still part of pendingAmount, just not "overdue"
            # by these bands' own definition).
            if age_days >= 90:
                band90plus_sum += b["amount"]
                has_band90plus = True
            elif age_days >= 60:
                band6089_sum += b["amount"]
                has_band6089 = True
            elif age_days >= 30:
                band3059_sum += b["amount"]
                has_band3059 = True
            else:
                band0_29_sum += b["amount"]
                has_band0_29 = True
        return {
            "since": min(dates) if dates else None,
            "pendingAmount": sum(amounts) if amounts else None,
            "agedPendingAmount": aged_sum if has_aged else None,
            "band0_29": band0_29_sum if has_band0_29 else None,
            "band3059": band3059_sum if has_band3059 else None,
            "band6089": band6089_sum if has_band6089 else None,
            "band90plus": band90plus_sum if has_band90plus else None,
        }

    out = {}
    for name in all_names:
        out[name] = _record(name)
    # Also include any BILLPARTY name from the report that, for whatever
    # reason, didn't show up in the plain Ledger name fetch (safer to
    # over-report than to silently drop a genuinely-pending dealer).
    for name in bills_by_party.keys():
        if name not in out:
            out[name] = _record(name)
    if capture_bills_into is not None:
        capture_bills_into.update(bills_by_party)
    return out


# ── State diffing (only send what changed since the last poll) ─────────
def load_state():
    state = None
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                state = json.load(f)
        except Exception:
            pass
    if state is None:
        state = {"stock": {}, "outstanding": {}, "collection": {}, "collectionByDealer": {},
                  "sales": {}, "salesByDealer": {}, "salesByItem": {}}
    # known_roster_names/roster_baseline_seeded added 19 Sep 2026 for the
    # "New dealers" feature - setdefault (not part of the dict literal
    # above) so an EXISTING state file from before this feature, which
    # won't have these keys at all, still gets them added rather than
    # KeyError-ing the first time maybe_push_dealer_roster() runs.
    state.setdefault("known_roster_names", {})
    state.setdefault("roster_baseline_seeded", False)
    # salesByDsrItem added 22 Sep 2026 for the dashboard's "Top selling SKU
    # by DSR" panels and "part number covered" per-DSR stat (owner's
    # request) - see fetch_sales()'s own comment near dsr_item_totals.
    state.setdefault("salesByDsrItem", {})
    # salesTrendDaily/collectionTrendDaily added 20 Sep 2026 for the
    # dashboard's Tally-based Sales+Collection trend chart (replacing the
    # old Orders-sheet/app-logged one) - {"YYYY-MM-DD": total} for this
    # month + last month, diffed each push same as every other live map
    # in this file. setdefault so an existing state file from before this
    # feature doesn't KeyError the first time it runs.
    state.setdefault("salesTrendDaily", {})
    state.setdefault("collectionTrendDaily", {})
    # salesRsTrendDaily added 20 Sep 2026 for the Collection-to-Sales
    # Ratio feature - Rs-value twin of salesTrendDaily (Ltr).
    state.setdefault("salesRsTrendDaily", {})
    # salesTrendDailyByDsr/collectionTrendDailyByDsr added 11 Oct 2026 for
    # the new per-DSR Sales/Collection weekday heatmap - dsr -> {"YYYY-MM-
    # DD": total}, one level deeper than salesTrendDaily/collectionTrendDaily
    # above. See fetch_sales()/fetch_collection()'s own daily_totals_by_dsr
    # comments.
    state.setdefault("salesTrendDailyByDsr", {})
    state.setdefault("collectionTrendDailyByDsr", {})
    # sixtyPlusBaseline/sixtyPlusByDealer added 23 Sep 2026 for the fixed,
    # bill-level "60+ days overdue, collected this month" dashboard stat -
    # see snapshot_or_get_sixtyplus_baseline()'s own docstring below for
    # why this replaced the old daily-recomputed version.
    state.setdefault("sixtyPlusBaseline", {})
    state.setdefault("sixtyPlusByDealer", {})
    # daybook_backfill_done added 26 Sep 2026 for the real (non-diagnostic)
    # Day Book sync - see maybe_push_daybook()'s own comment below for why
    # this flag (rather than a synced-rows dict, unlike every other
    # live_* map above) is all the local state Day Book needs: rows
    # themselves live in the Google Sheet's monthly tabs, upserted
    # idempotently by voucher#+line key, so there's nothing to diff
    # locally - just "have we ever done the one-time 1 Sep backfill yet".
    state.setdefault("daybook_backfill_done", False)
    # creditlimit added 30 Sep 2026 - see fetch_creditlimit()/diff_creditlimit()
    # above and the LiveCreditLimit push in run_once() below. setdefault so
    # an existing state file from before this feature doesn't KeyError the
    # first time it runs - that first run will then see every ledger with a
    # non-blank CREDITLIMIT as "changed" (nothing was ever recorded before),
    # a one-time larger push, same as any other feature's first rollout.
    state.setdefault("creditlimit", {})
    return state


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=None)


def diff_stock(prev, current):
    changed = []
    for name, vals in current.items():
        old = prev.get(name)
        if old != vals:
            changed.append(dict(name=name, **vals))
    return changed


def diff_outstanding(prev, current):
    """
    current: {name: {"outstandingSince": iso_or_None, "balance": float_or_None,
    "pendingAmount": float_or_None, "agedPendingAmount": float_or_None,
    "band0_29": float_or_None, "band3059": float_or_None, "band6089": float_or_None,
    "band90plus": float_or_None}}
    (changed 8 Sep 2026 from a bare date value to a small dict so the live
    current balance travels alongside the outstanding-since date - see the
    "Manage dealers" balance feature; the "pendingAmount" key was added
    17 Sep 2026, see fetch_outstanding()'s BILLCL docstring section;
    "agedPendingAmount" added 19 Sep 2026 for the payment-reminder feature;
    "band3059"/"band6089"/"band90plus" added 22 Sep 2026, see
    fetch_outstanding()'s docstring for the age-accuracy bug they fix;
    "band0_29" added 23 Sep 2026 for the dashboard's DSR-wise aging table.)
    A dealer whose balance/pendingAmount/agedPendingAmount/any band moved
    but whose outstandingSince didn't (or vice versa) still counts as
    changed, since any part updating is worth pushing.
    """
    changed = []
    for name, rec in current.items():
        old = prev.get(name, "__unset__")
        if old != rec:
            changed.append({
                "name": name,
                "outstandingSince": rec["outstandingSince"],
                "balance": rec["balance"],
                "pendingAmount": rec["pendingAmount"],
                "agedPendingAmount": rec["agedPendingAmount"],
                "band0_29": rec["band0_29"],
                "band3059": rec["band3059"],
                "band6089": rec["band6089"],
                "band90plus": rec["band90plus"],
            })
    return changed


# ── Backend POST ─────────────────────────────────────────────────────────
# POST_CHUNK_SIZE / the 60s timeout were added 8 Sep 2026 after the very
# first live sync (8,809 changed stock items + thousands of dealers, since
# the watcher starts with no prior state to diff against) timed out
# repeatedly trying to send everything in one request ("Read timed out
# (read timeout=30)", twice in a row on the office's first live run).
# Paired with a bulk-upsert rewrite on the Code.gs side (see
# bulkUpsertSheet_ there, which turned an O(n^2) per-item sheet scan into
# one read + one write) - that fix alone should make even a single large
# POST fast, but chunking here stays as a safety net regardless of
# backend speed, and keeps each individual request's JSON payload a
# reasonable size. Steady-state polls (after the first one) only ever
# send a small diff, so this only really matters on day one / after a
# long gap.
POST_CHUNK_SIZE = 300

# Retry-with-backoff added 25 Sep 2026, after a real production run showed
# a batch mid-sync fail with "404 Client Error: Not Found for url:
# https://script.googleusercontent.com/macros/echo?user_content_key=...
# &lib=...". That URL is Google's OWN internal redirect target, not
# BACKEND_URL itself - every Apps Script Web App response is served this
# way: the initial POST to BACKEND_URL gets a 302 to a short-lived,
# one-time script.googleusercontent.com/macros/echo?... link, which
# `requests` follows automatically as part of one post_backend() call.
# A 404 on THAT link means the Apps Script execution itself glitched or
# ran unusually slowly server-side (the failing batch that night took over
# 2 minutes to fail, versus well under a second for the ones before it,
# on the exact same code/data shape) - not a bug in this file's own logic,
# and not something BACKEND_URL/SYNC_KEY being wrong would produce (a bad
# URL/key fails immediately and cleanly, not after a multi-minute hang).
# Before this fix, ANY single chunk failing (out of up to 30 for a big
# stock sync) aborted the WHOLE poll via run_once()'s exception - and
# since save_state() only runs after run_once() returns successfully (see
# main()'s while-loop), even the chunks that DID succeed and write to the
# Sheet before the failure never got reflected in local state, so the
# NEXT poll would recompute and resend the exact same large diff all over
# again - wasteful, and means a single transient Google-side hiccup can
# keep a big sync from ever fully landing. **Fixed**: each chunk now gets
# up to 3 attempts with a short backoff (3s, then 8s) before giving up -
# a fresh POST to BACKEND_URL always gets a brand-new redirect token, so
# retrying is very likely to simply succeed rather than hit the same
# stale link again. Only re-raises (letting the whole poll fail, same as
# before) after all 3 attempts are exhausted, so a genuinely dead
# backend/network still surfaces as an error rather than retrying forever.
POST_BACKEND_MAX_ATTEMPTS = 3
POST_BACKEND_RETRY_DELAYS = (3, 8)  # seconds, between attempts 1->2 and 2->3


def post_backend(action, payload_key, payload_value):
    body = {"action": action, "key": SYNC_KEY}
    body[payload_key] = payload_value
    last_err = None
    for attempt in range(1, POST_BACKEND_MAX_ATTEMPTS + 1):
        try:
            resp = requests.post(BACKEND_URL, data=json.dumps(body),
                                  headers={"Content-Type": "text/plain"}, timeout=60)
            resp.raise_for_status()
            return resp.json()
        except requests.exceptions.ConnectionError:
            raise  # Tally/network being down isn't transient in the same way - let the
            # existing "Tally not reachable" handling in main()'s loop catch this as before.
        except requests.exceptions.RequestException as e:
            last_err = e
            if attempt < POST_BACKEND_MAX_ATTEMPTS:
                delay = POST_BACKEND_RETRY_DELAYS[attempt - 1]
                print("    {} attempt {}/{} failed ({}) - retrying in {}s...".format(
                    action, attempt, POST_BACKEND_MAX_ATTEMPTS, e, delay))
                time.sleep(delay)
    raise last_err


def post_backend_chunked(action, payload_key, items, ts):
    """Splits a large update into POST_CHUNK_SIZE-sized POSTs, printing
    progress as it goes so a big first sync doesn't look stuck."""
    total = len(items)
    sent = 0
    chunks = [items[i:i + POST_CHUNK_SIZE] for i in range(0, total, POST_CHUNK_SIZE)]
    for idx, chunk in enumerate(chunks, 1):
        result = post_backend(action, payload_key, chunk)
        sent += result.get("updated", len(chunk))
        if len(chunks) > 1:
            print("[{}]   {} batch {}/{}: {} row(s) -> {}".format(
                ts, action, idx, len(chunks), len(chunk), result))
    return sent


# ── 60+ day overdue, bill-level monthly baseline ────────────────────────
# Added 23 Sep 2026, replacing the old daily-recomputed "60+ days
# outstanding, collected this month" DSR-card stat (built 22 Sep 2026).
# The owner flagged a real flaw: recomputing "which bills are 60+ days
# old" fresh on every poll means the 60+ POOL ITSELF drifts day to day
# (a bill that was 58 days old yesterday is 60+ today, and yesterday's
# 60+ bill that just got paid off simply vanishes from the pool) - so
# "collected this month against 60+ debt" was really measuring a moving
# target, not a stable one. The owner's own explicit choice (asked via
# AskUserQuestion): freeze the SPECIFIC BILLS (by BILLREF+date+amount,
# not just a dealer-level cash total) that are 60+ days old as of the
# 1st of each calendar month, and track collection against only THOSE
# frozen bills for the rest of the month - so a dealer's brand-new,
# unrelated purchase this month (and its payment) never gets misread as
# old-debt collection, since it simply isn't one of the frozen bills.
def snapshot_or_get_sixtyplus_baseline(state, bills_by_party, today_iso):
    """
    Returns this month's frozen 60+ day baseline, computing and storing a
    new one (into state["sixtyPlusBaseline"], mutated in place - caller
    still needs to save_state()) the first time this runs in a new
    calendar month, otherwise just returns the one already stored.

    bills_by_party is fetch_outstanding()'s capture_bills_into output -
    EVERY dealer's EVERY currently-open bill (not pre-filtered to 60+),
    since this function does its own age filtering at snapshot time.

    Baseline shape: {"month": "YYYY-MM", "asOfDate": "YYYY-MM-DD",
    "dealers": {dealerName: [{"billref", "date", "amount"}, ...]}} - only
    dealers with at least one 60+ day bill at snapshot time appear.

    Caveat (documented for the owner, not a bug): the very first time
    this feature ever runs, "month start" is necessarily "as of whatever
    day this update actually goes live" - the past can't be retroactively
    captured. From the next calendar month onward, every baseline is a
    true 1st-of-month snapshot, taken automatically the first time
    run_once() executes on or after that 1st.
    """
    current_month = today_iso[:7]  # "YYYY-MM"
    baseline = state.get("sixtyPlusBaseline") or {}
    if baseline.get("month") == current_month:
        return baseline
    today_dt = datetime.strptime(today_iso, "%Y-%m-%d")
    dealers = {}
    for name, bills in bills_by_party.items():
        frozen = []
        for b in bills:
            if b["amount"] >= 0:
                continue  # not a genuine debit - same guard used everywhere else in this file
            age_days = (today_dt - datetime.strptime(b["date"], "%Y-%m-%d")).days
            if age_days >= 60:
                frozen.append({"billref": b.get("billref"), "date": b["date"], "amount": b["amount"]})
        if frozen:
            dealers[name] = frozen
    baseline = {"month": current_month, "asOfDate": today_iso, "dealers": dealers}
    state["sixtyPlusBaseline"] = baseline
    return baseline


def _sixtyplus_bill_key(b):
    """A bill's own BILLREF uniquely identifies it across polls; a small
    number of older Tally bills can come back with no BILLREF at all, so
    those fall back to a (date, amount) composite key instead - not as
    precise (two different bills of the exact same age/amount for the
    same dealer would collide), but far better than dropping them from
    tracking entirely, and BILLREF is present on the large majority of
    real bills seen in this office's data."""
    return b.get("billref") or ("__noref__", b["date"], b["amount"])


def compute_sixtyplus_collected_by_dealer(baseline, bills_by_party):
    """
    For each dealer in the frozen baseline, checks ONLY their frozen
    bills (by the same key _sixtyplus_bill_key() uses) against the
    CURRENT bills_by_party snapshot: a frozen bill that no longer appears
    at all has been fully paid off (Tally's Bills Receivable report only
    ever lists still-open bills, so "gone" = "closed"); one that still
    appears with a smaller (less negative) amount has been part-paid;
    baseline_total minus what's still open across those specific bills =
    genuine old-debt collected this month. A dealer's new, unrelated bill
    this month is never one of the frozen bills, so it and its payment
    never enter this calculation at all - this is what makes the
    bill-level measure safe against the cash-level miscount the owner
    was worried about.
    """
    out = {}
    for name, frozen_bills in baseline.get("dealers", {}).items():
        current_by_key = {}
        for b in bills_by_party.get(name, []):
            current_by_key[_sixtyplus_bill_key(b)] = b["amount"]
        baseline_total = 0.0
        current_total = 0.0
        for fb in frozen_bills:
            baseline_total += abs(fb["amount"])
            still_open_amt = current_by_key.get(_sixtyplus_bill_key(fb))
            if still_open_amt is not None:
                current_total += abs(still_open_amt)
            # else: this specific frozen bill is gone from the open-bills
            # report entirely -> fully closed -> contributes 0 (i.e. its
            # whole frozen amount counts as collected).
        out[name] = {
            "baseline": round(baseline_total, 2),
            "collected": round(baseline_total - current_total, 2),
            "billCount": len(frozen_bills),
        }
    return out


def diff_sixtyplus_by_dealer(prev, current):
    changed = []
    for name, rec in current.items():
        old = prev.get(name, "__unset__")
        if old != rec:
            changed.append({
                "name": name,
                "baseline": rec["baseline"],
                "collected": rec["collected"],
                "billCount": rec["billCount"],
            })
    return changed


# ── 60+ day baseline diagnostic ─────────────────────────────────────────
# Added 25 Sep 2026, after the owner reported "60+ days overdue - collected
# this month" looking wrong on the dashboard, with no specifics given yet.
# Nothing here was guessable from the code alone - compute_sixtyplus_
# collected_by_dealer() itself is a straightforward "sum what's missing
# from the frozen bill list" calculation and has no obvious sign/logic bug
# on inspection - so before changing anything, this prints the RAW inputs
# to that calculation for one dealer (or every dealer matching a fragment),
# so a wrong-looking number can be traced back to its actual cause rather
# than guessed at. Same "show the real data before touching code" discipline
# as every other live-Tally diagnostic in this file (--check-collection,
# --dump-sales, etc).
#
# One real, plausible failure mode this is specifically built to catch:
# Tally's bill-by-bill tracking can, depending on this office's own
# "Maintain Bill-wise Details" configuration, issue a NEW BILLREF for the
# remaining balance when a bill is PART-paid (rather than keeping the same
# BILLREF with a smaller BILLCL) - if that happens here, a frozen 60+ bill
# would look "fully closed" (its old BILLREF vanishes from Bills
# Receivable) the moment ANY partial payment lands on it, so the WHOLE
# frozen amount gets counted as "collected" even though only part of it
# actually was - this would show up as this month's 60+ collected figure
# jumping to (at or near) 100% the instant a partial payment posts, which
# is exactly the kind of "not working" a manager would notice at a glance.
# This diagnostic surfaces that directly: for each dealer, "gone" frozen
# bills (counted as fully collected) are printed next to that dealer's
# CURRENT full bill list, so a same-dealer bill of a similar/smaller
# amount dated close to today, under a different BILLREF, is visible side
# by side rather than silently assumed.
def dump_sixtyplus(dealer_fragment=None):
    state = load_state()
    baseline = state.get("sixtyPlusBaseline") or {}
    if not baseline.get("dealers"):
        print("No 60+ day baseline stored yet in {} - the watcher hasn't run "
              "since this feature was added, or it's not yet reached its first "
              "poll of this calendar month. Run the watcher normally at least "
              "once first (not this diagnostic), then re-run this.".format(STATE_FILE))
        return
    print("Baseline snapshot: month={}, taken as of {}, {} dealer(s) in the frozen 60+ pool.\n".format(
        baseline.get("month"), baseline.get("asOfDate"), len(baseline.get("dealers", {}))))

    bills_raw = {}
    fetch_outstanding(capture_bills_into=bills_raw)
    collected_now = compute_sixtyplus_collected_by_dealer(baseline, bills_raw)

    names = sorted(baseline.get("dealers", {}).keys())
    if dealer_fragment:
        frag = dealer_fragment.strip().lower()
        names = [n for n in names if frag in n.lower()]
        if not names:
            print("No dealer in the frozen 60+ baseline matches \"{}\". Try a shorter "
                  "fragment, or omit it to list every dealer in the baseline.".format(dealer_fragment))
            return

    for name in names:
        frozen_bills = baseline["dealers"][name]
        current_bills = bills_raw.get(name, [])
        current_by_key = dict((_sixtyplus_bill_key(b), b) for b in current_bills)
        result = collected_now.get(name, {"baseline": 0, "collected": 0, "billCount": 0})
        print("=" * 70)
        print("{}  ->  collected {} / baseline {} ({} bill(s) frozen)".format(
            name, result["collected"], result["baseline"], result["billCount"]))
        print("  Frozen 60+ bills (as of {}):".format(baseline.get("asOfDate")))
        for fb in frozen_bills:
            key = _sixtyplus_bill_key(fb)
            still = current_by_key.get(key)
            if still is None:
                status = "GONE from current Bills Receivable -> counted as FULLY collected (Rs {})".format(abs(fb["amount"]))
            elif still["amount"] == fb["amount"]:
                status = "unchanged, still fully open -> Rs 0 collected on this bill"
            else:
                status = "still open but reduced to Rs {} -> Rs {} collected on this bill".format(
                    abs(still["amount"]), round(abs(fb["amount"]) - abs(still["amount"]), 2))
            print("    billref={!r} date={} amount={}  [{}]".format(
                fb.get("billref"), fb.get("date"), fb.get("amount"), status))
        print("  ALL of this dealer's CURRENT open bills (for comparison - look here for a "
              "same-amount/near-date bill under a different billref if a frozen bill shows "
              "as GONE above):")
        if current_bills:
            for cb in sorted(current_bills, key=lambda b: b.get("date") or ""):
                print("    billref={!r} date={} amount={}".format(cb.get("billref"), cb.get("date"), cb.get("amount")))
        else:
            print("    (none - dealer has no currently-open bills at all)")
        print()


# ── Main loop ─────────────────────────────────────────────────────────────
def run_once(state, dry_run=False):
    stock = fetch_stock()
    # Fetched once here and passed into fetch_outstanding() below so that
    # function doesn't re-fetch the same ledger list a second time just to
    # get the name universe - one Tally request now covers both the balance
    # figure and the dealer-name list fetch_outstanding() needs.
    all_ledger_balances = fetch_all_ledgers_with_balance()
    # bills_raw captured as a side effect of fetch_outstanding() below (no
    # extra Tally request) - every dealer's every currently-open bill, fed
    # to the 60+ day baseline block further down.
    bills_raw = {}
    outstanding_data = fetch_outstanding(all_ledger_balances, capture_bills_into=bills_raw)
    outstanding = dict(
        (name, {
            "outstandingSince": rec["since"],
            "balance": all_ledger_balances.get(name),
            "pendingAmount": rec["pendingAmount"],
            "agedPendingAmount": rec["agedPendingAmount"],
            "band0_29": rec["band0_29"],
            "band3059": rec["band3059"],
            "band6089": rec["band6089"],
            "band90plus": rec["band90plus"],
        })
        for name, rec in outstanding_data.items()
    )

    # Dealer credit-limit/credit-lock status - confirmed against real Tally
    # data 30 Sep 2026 (see fetch_creditlimit()'s own docstring). Same
    # weight class as fetch_all_ledgers_with_balance() above (one
    # whole-company Ledger Collection request, not a per-voucher pull like
    # Collection/Sales/Day Book/Group Summary below), so it runs on the same
    # 10-minute cadence as stock/outstanding rather than a slower one.
    creditlimit = fetch_creditlimit()

    stock_changed = diff_stock(state["stock"], stock)
    outstanding_changed = diff_outstanding(state["outstanding"], outstanding)
    creditlimit_changed = diff_creditlimit(state.get("creditlimit", {}), creditlimit)

    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if dry_run:
        print("[{}] --test mode - nothing sent, just showing what would happen".format(ts))
        print("Stock items read from Tally: {}".format(len(stock)))
        print("Dealers read from Tally: {}".format(len(outstanding)))
        print()
        print("Would send {} changed stock item(s):".format(len(stock_changed)))
        for it in stock_changed[:15]:
            print("  {}  stock={} rate={}".format(it["name"], it.get("stock"), it.get("rate")))
        if len(stock_changed) > 15:
            print("  ... and {} more".format(len(stock_changed) - 15))
        print()
        print("Would send {} changed dealer-outstanding row(s):".format(len(outstanding_changed)))
        for d in outstanding_changed[:30]:
            status = d["outstandingSince"] or "CLEARED (no longer outstanding)"
            print("  {}  ->  {}   (balance: {}, pending: {})".format(d["name"], status, d["balance"], d["pendingAmount"]))
        if len(outstanding_changed) > 30:
            print("  ... and {} more".format(len(outstanding_changed) - 30))
        print()
        print("Full outstanding snapshot (first 20 dealers, for spot-checking against real Tally):")
        for name in list(outstanding.keys())[:20]:
            rec = outstanding[name]
            print("  {}  ->  {}   (balance: {}, pending: {})".format(
                name, rec["outstandingSince"] or "not outstanding", rec["balance"], rec["pendingAmount"]))
        print()
        print("Would send {} changed credit-limit row(s):".format(len(creditlimit_changed)))
        for c in creditlimit_changed[:15]:
            print("  {}  ->  creditLimit={} creditPeriod={}".format(c["name"], c["creditLimit"], c["creditPeriod"]))
        if len(creditlimit_changed) > 15:
            print("  ... and {} more".format(len(creditlimit_changed) - 15))
        return state  # unchanged - --test never updates the state file

    if stock_changed:
        sent = post_backend_chunked("liveStockUpdate", "items", stock_changed, ts)
        print("[{}] stock: sent {} changed item(s) total".format(ts, sent))
    if outstanding_changed:
        sent = post_backend_chunked("liveOutstandingUpdate", "dealers", outstanding_changed, ts)
        print("[{}] outstanding: sent {} changed dealer(s) total".format(ts, sent))
    if creditlimit_changed:
        sent = post_backend_chunked("liveCreditLimitUpdate", "dealers", creditlimit_changed, ts)
        print("[{}] credit limit: sent {} changed ledger(s) total".format(ts, sent))
    if not stock_changed and not outstanding_changed and not creditlimit_changed:
        print("[{}] no changes".format(ts))

    state["stock"] = stock
    state["outstanding"] = outstanding
    state["creditlimit"] = creditlimit

    # 60+ day overdue, bill-level monthly baseline - see
    # snapshot_or_get_sixtyplus_baseline()'s docstring above for the full
    # design. Cheap (reuses bills_raw already captured above, no extra
    # Tally request), so this runs every poll, not on a slower cadence
    # like the roster/collection/sales blocks below.
    if not dry_run:
        sixtyplus_baseline = snapshot_or_get_sixtyplus_baseline(
            state, bills_raw, datetime.now().strftime("%Y-%m-%d"))
        sixtyplus_by_dealer = compute_sixtyplus_collected_by_dealer(sixtyplus_baseline, bills_raw)
        sixtyplus_changed = diff_sixtyplus_by_dealer(state.get("sixtyPlusByDealer", {}), sixtyplus_by_dealer)
        if sixtyplus_changed:
            sent = post_backend_chunked("liveSixtyPlusBaselineUpdate", "dealers", sixtyplus_changed, ts)
            print("[{}] 60+ baseline: sent {} changed dealer(s) total".format(ts, sent))
        state["sixtyPlusByDealer"] = sixtyplus_by_dealer

    # New-dealer roster - separate, much slower cadence, see
    # maybe_push_dealer_roster()'s own docstring for why. Runs after the
    # fast stock/outstanding work above is already sent, so it never delays
    # that regardless of how long the roster computation takes.
    state = maybe_push_dealer_roster(state, dry_run, ts)

    # DSR Collection (Receipt vouchers) - added 12 Sep 2026, its own
    # 30-min cadence, see maybe_push_collection()/COLLECTION_POLL_INTERVAL_SECONDS
    # above for why. Runs last, after everything faster is already sent.
    state = maybe_push_collection(state, dry_run, ts)

    # DSR Sales (Sales vouchers, dashboard only) - added 18 Sep 2026, same
    # heavy-whole-company-pull reasoning as Collection above, own 60-min
    # cadence - see maybe_push_sales()/SALES_POLL_INTERVAL_SECONDS.
    state = maybe_push_sales(state, dry_run, ts)

    # Day Book (Sales/Sales-PETRONAS/Sales-CBE, ALL dealers, no DSR
    # filtering - dashboard's Excel export) - added 26 Sep 2026, own 60-min
    # cadence, runs last since it's the newest/least time-critical of the
    # heavy whole-company pulls above. See maybe_push_daybook()/
    # DAYBOOK_POLL_INTERVAL_SECONDS above.
    state = maybe_push_daybook(state, dry_run, ts)

    # Collection Group Summary (Opening/Debit/Credit/Closing per Petronas
    # sales-area group, all voucher types, expandable "+" dashboard panel
    # with per-dealer drill-down and a full-data Excel export) - added
    # 28 Sep 2026, method fully validated against real Tally data first
    # (see dump_group_summary()'s docstring/history above - 7 diagnostic
    # rounds, ending with the post-dated-voucher correction). Own slow
    # cadence, confirmed with the owner via AskUserQuestion ("a few times
    # a day" / 1-2 hours, not every ~10-min poll) since this is the
    # heaviest single Tally pull in this whole file - see
    # maybe_push_group_summary()/GROUP_SUMMARY_POLL_INTERVAL_SECONDS.
    # Runs last, after every lighter/more time-critical push above.
    state = maybe_push_group_summary(state, dry_run, ts)

    # Petronas Business Summary - LIVE, current month only - added 1 Oct
    # 2026 per the owner's explicit choice ("FROM SEP ONWARD LIVE DATA
    # SHOULD COME"). Jan-Aug 2026 stay a static, owner-confirmed snapshot
    # baked into the dashboard; this computes and pushes the CURRENT
    # month's figures going forward. Same heavy-pull reasoning/cadence as
    # Collection/Sales/Group Summary above - runs last.
    state = maybe_push_petronas_summary(state, dry_run, ts)

    # S.REP - Sales & Collection Plan panel - added 4 Nov 2026, same
    # heavy-pull reasoning/cadence as Petronas Business Summary just
    # above - runs last.
    state = maybe_push_srep_summary(state, dry_run, ts)

    # Anbu/TN Head scope (Coimbatore + Maruti + TATA combined) - added 17
    # Nov 2026, same heavy-pull reasoning as Petronas Business Summary/
    # S.REP just above (a full-year unfiltered inventory scan every
    # cycle) - runs last of all.
    state = maybe_push_anbu_tn_scope(state, dry_run, ts)

    return state


def dump_xml():
    """
    Diagnostic mode: fetches the raw XML Tally sends back for both requests
    and saves it to files, without trying to parse/interpret any of it -
    for when --test's output looks wrong (e.g. every dealer coming back
    "not outstanding") and the fix needs to see Tally's actual field names/
    structure instead of guessing again. Also prints a couple of small,
    easy-to-copy-paste excerpts so a full multi-megabyte file doesn't have
    to be sent back and forth.
    """
    out_dir = os.path.dirname(os.path.abspath(__file__))
    stock_path = os.path.join(out_dir, "tally_stock_dump.xml")
    outstanding_path = os.path.join(out_dir, "tally_outstanding_dump.xml")

    print("Fetching raw stock XML from Tally...")
    stock_xml = tally_request(STOCK_COLLECTION_XML)
    with open(stock_path, "w", encoding="utf-8") as f:
        f.write(stock_xml)
    print("Saved full response to: {}  ({} bytes)".format(stock_path, len(stock_xml)))

    print("\nFetching raw outstanding/ledger XML from Tally...")
    outstanding_xml = tally_request(OUTSTANDING_COLLECTION_XML)
    with open(outstanding_path, "w", encoding="utf-8") as f:
        f.write(outstanding_xml)
    print("Saved full response to: {}  ({} bytes)".format(outstanding_path, len(outstanding_xml)))

    print("\n" + "=" * 70)
    print("EASY-COPY EXCERPTS - paste these back if asked:")
    print("=" * 70)

    # First full <STOCKITEM ...>...</STOCKITEM> block, raw, unparsed.
    si_start = stock_xml.find("<STOCKITEM")
    if si_start != -1:
        si_end = stock_xml.find("</STOCKITEM>", si_start)
        si_end = si_end + len("</STOCKITEM>") if si_end != -1 else si_start + 2000
        print("\n--- First <STOCKITEM> block (raw) ---")
        print(stock_xml[si_start:si_end])
    else:
        print("\n--- No <STOCKITEM tag found in the stock response at all - first 1000 chars instead: ---")
        print(stock_xml[:1000])

    # The full <LEDGER ...>...</LEDGER> block for Challenger Motors specifically,
    # since we know its real-world status (paid off) to check the fix against.
    idx = outstanding_xml.upper().find("CHALLENGER")
    if idx != -1:
        block_start = outstanding_xml.rfind("<LEDGER", 0, idx)
        block_end = outstanding_xml.find("</LEDGER>", idx)
        block_end = block_end + len("</LEDGER>") if block_end != -1 else idx + 3000
        if block_start == -1:
            block_start = max(0, idx - 200)
        print("\n--- Full <LEDGER> block for Challenger Motors (raw) ---")
        print(outstanding_xml[block_start:block_end])
    else:
        print("\n--- 'CHALLENGER' not found anywhere in the outstanding response. "
              "First <LEDGER> block instead: ---")
        l_start = outstanding_xml.find("<LEDGER")
        if l_start != -1:
            l_end = outstanding_xml.find("</LEDGER>", l_start)
            l_end = l_end + len("</LEDGER>") if l_end != -1 else l_start + 3000
            print(outstanding_xml[l_start:l_end])
        else:
            print("No <LEDGER tag found in the outstanding response at all - first 1000 chars instead:")
            print(outstanding_xml[:1000])

    print("\n" + "=" * 70)
    print("Send the excerpts above back (copy-paste this terminal output is fine).")
    print("If they're too long to paste, the two saved .xml files can be attached instead.")


def dump_price_check():
    """
    Diagnostic mode (added 11 Sep 2026): fetches the stock collection (now
    including STANDARDPRICE alongside the old CLOSINGRATE) and prints both
    side by side for a handful of items whose CORRECT pre-GST selling rate
    is already independently known (confirmed against the office's own
    price-list uploads and cross-checked with the owner directly) - so this
    can be verified the same way fetch_outstanding() was, before
    STANDARDPRICE is ever wired into the live rate push in fetch_stock().

    Also reports how many stock items in the whole company have a non-empty
    STANDARDPRICE at all, in case it's only set for some items (e.g. only
    ones actively sold through this counter) rather than universally - a
    live-rate feature that only works for a fraction of the catalog would
    need to leave everything else on the Drive-sync rate, same as items
    with no live data today already do.
    """
    KNOWN_ITEMS = {
        "10W30 1 LIT SPRINTA F700 (20)": 438.14,
        "CF4 15W40 10 LIT URANIA 800": 2817.8,
        "5W30 SYNTIUM 500 SN/CF 210 LIT": 57694.07,
        "15W40 1 LIT SYNTIUM 500 CI4+ DIESEL (20)": 338.98,
    }
    print("Fetching stock collection (NAME, CLOSINGRATE, STANDARDPRICE, CLOSINGBALANCE)...")
    xml_text = tally_request(STOCK_COLLECTION_XML)
    root = ET.fromstring(xml_text)
    by_name = {}
    have_standard_price = 0
    total = 0
    for si in root.iter("STOCKITEM"):
        name = si.get("NAME") or _text(si, "NAME")
        if not name:
            continue
        name = name.strip()
        total += 1
        sp = _num(_text(si, "STANDARDPRICE"))
        cr = _num(_text(si, "CLOSINGRATE"))
        if sp is not None:
            have_standard_price += 1
        by_name[name] = {"standardPrice": sp, "closingRate": cr}

    print("\n" + "=" * 70)
    print("Known-answer items (expected = the pre-GST rate already confirmed correct):")
    print("=" * 70)
    for name, expected in KNOWN_ITEMS.items():
        rec = by_name.get(name)
        if rec is None:
            print("  {}  -> NOT FOUND in Tally's stock list under this exact name".format(name))
            continue
        sp = rec["standardPrice"]
        cr = rec["closingRate"]
        sp_match = "MATCH" if (sp is not None and abs(sp - expected) < 1) else "MISMATCH"
        print("  {}".format(name))
        print("    expected (known-correct): {}".format(expected))
        print("    STANDARDPRICE from Tally: {}   [{}]".format(sp, sp_match if sp is not None else "MISSING"))
        print("    CLOSINGRATE from Tally (for reference, should NOT match - cost, not price): {}".format(cr))

    print("\n" + "=" * 70)
    print("Coverage: {} of {} total stock items have a non-empty STANDARDPRICE.".format(
        have_standard_price, total))
    print("=" * 70)
    print("\nSend this whole output back before live rate is wired in for real.")


# ── MRP live-fetch investigation, 11 Sep 2026 ────────────────────────────
# The owner asked whether the same TDL file that revealed StandardPrice can
# also confirm a live MRP field, and specifically wants this kept as an
# ISOLATED check that does not touch price/rate at all (rate is freshly
# re-enabled and confirmed working - no reason to risk it while poking at
# something unrelated). So this uses its OWN Collection request
# (StockMRPCollection below), completely separate from STOCK_COLLECTION_XML/
# fetch_stock()/dump_price_check() - it cannot affect the rate/price path in
# any way, by construction, not just by convention.
#
# What the TDL confirms: `[Field: AAA MDD Item MRP] Set as :
# $MRPDetails[Last].MRPRateDetails[Last].MRPRate` and
# `[Field: AAA MDD Item MRPDate] Set as : $MRPDetails[Last].FromDate`
# (the .tdl file even has a comment: "VERIFIED working: FromDate (NOT
# ApplicableFrom)" - from whoever originally wrote the report). So Tally
# models MRP as a nested, dated LIST (an item can have had several MRPs
# over time, same idea as a price-list history) - `MRPDetails` is a list of
# entries, each entry itself has its own nested `MRPRateDetails` list, and
# `[Last]` picks the most recent one on each level. That is a TWO-level
# nested list, unlike StandardPrice (a plain flat field) or even
# BATCHALLOCATIONS.LIST (one level) - so unlike StandardPrice, this genuinely
# cannot be fetched as a single flat FETCH field name. The Collection's own
# `Fetch` line in the TDL just says `MRPDetails` (Tally expands that
# automatically when rendering *inside a report*), but a raw XML Collection
# export needs the actual list tag name, which is conventionally
# `<X.LIST>` for a repeating field in Tally's XML (confirmed already for
# BATCHALLOCATIONS.LIST and BILLALLOCATIONS.LIST) - so `MRPDETAILS.LIST` is
# the best-informed guess, not yet confirmed. This diagnostic fetches it,
# prints the RAW block for known items (so the actual tag names can be read
# directly if the guess is wrong, same fallback used for --dump-mrp and
# --dump-report), and also attempts a best-effort parse + known-answer
# check using 3 items whose correct MRP is already in the catalog:
# `10W30 1 LIT SPRINTA F700 (20)` (697), `CF4 15W40 10 LIT URANIA 800`
# (5018), `15W40 1 LIT SYNTIUM 500 CI4+ DIESEL (20)` (630) - deliberately
# NOT the barrel item used for the rate checks (`5W30 SYNTIUM 500 SN/CF 210
# LIT`), since barrel-size items don't carry an MRP in this catalog at all.
#
# fetch_stock()'s "mrp" output is NOT touched by this - stays None, exactly
# as it has been all along - until this is confirmed the same way
# STANDARDPRICE was.
STOCK_MRP_COLLECTION_XML = """<ENVELOPE>
 <HEADER>
  <VERSION>1</VERSION>
  <TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE>
  <ID>StockMRPOnlyCollection</ID>
 </HEADER>
 <BODY>
  <DESC>
   <STATICVARIABLES>
    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
   </STATICVARIABLES>
   <TDL>
    <TDLMESSAGE>
     <COLLECTION NAME="StockMRPOnlyCollection" ISMODIFY="No">
      <TYPE>StockItem</TYPE>
      <FETCH>NAME, CLOSINGBALANCE</FETCH>
      <FETCH>MRPDETAILS.LIST</FETCH>
     </COLLECTION>
    </TDLMESSAGE>
   </TDL>
  </DESC>
 </BODY>
</ENVELOPE>"""


def _find_mrp_rate(si_elem):
    """
    Best-effort walk of one <STOCKITEM> element looking for the live MRP,
    following the TDL's $MRPDetails[Last].MRPRateDetails[Last].MRPRate path:
    take the LAST <MRPDETAILS.LIST> child, then within it the LAST nested
    list child (guessed name MRPRATEDETAILS.LIST), then a numeric field on
    it whose tag name contains "RATE". Returns (value, path_description) or
    (None, reason) - the reason string is printed so a MISS can be diagnosed
    from the summary alone, without necessarily needing the raw block.
    """
    mrp_lists = si_elem.findall("MRPDETAILS.LIST")
    if not mrp_lists:
        # fall back to a loose search, in case the real tag differs slightly
        mrp_lists = [e for e in si_elem if "MRPDETAIL" in e.tag.upper()]
    if not mrp_lists:
        return None, "no MRPDETAILS.LIST-like element found on this item"

    last_mrp = mrp_lists[-1]
    rate_lists = last_mrp.findall("MRPRATEDETAILS.LIST")
    if not rate_lists:
        rate_lists = [e for e in last_mrp if "MRPRATEDETAIL" in e.tag.upper()]

    search_root = rate_lists[-1] if rate_lists else last_mrp
    rate_val = None
    rate_tag = None
    for e in search_root.iter():
        if e is search_root:
            continue
        if "RATE" in e.tag.upper() and e.text and e.text.strip():
            v = _num(e.text)
            if v is not None:
                rate_val = v
                rate_tag = e.tag
                break
    if rate_val is None:
        return None, "found MRPDETAILS.LIST but no *RATE* field inside it"
    return rate_val, "via <{}>".format(rate_tag)


def dump_mrp_live():
    """
    Diagnostic mode (added 11 Sep 2026): fetches ONLY MRPDETAILS.LIST (its
    own isolated Collection request, StockMRPOnlyCollection - see the module
    comment above) and prints raw + best-effort-parsed results for a couple
    of known items, plus a known-answer comparison. Does not touch rate or
    price in any way.
    """
    KNOWN_MRP = {
        "10W30 1 LIT SPRINTA F700 (20)": 697,
        "CF4 15W40 10 LIT URANIA 800": 5018,
        "15W40 1 LIT SYNTIUM 500 CI4+ DIESEL (20)": 630,
    }
    out_dir = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(out_dir, "tally_mrp_live_dump.xml")

    print("Fetching MRPDETAILS.LIST only (isolated from rate/price) from Tally...")
    xml_text = tally_request(STOCK_MRP_COLLECTION_XML)
    with open(path, "w", encoding="utf-8") as f:
        f.write(xml_text)
    print("Saved full response to: {}  ({} bytes)".format(path, len(xml_text)))

    root = ET.fromstring(xml_text)
    by_name = {}
    have_mrp = 0
    total = 0
    for si in root.iter("STOCKITEM"):
        name = si.get("NAME") or _text(si, "NAME")
        if not name:
            continue
        name = name.strip()
        total += 1
        val, why = _find_mrp_rate(si)
        if val is not None:
            have_mrp += 1
        by_name[name] = (val, why)

    print("\n" + "=" * 70)
    print("Known-answer items (expected = the MRP already confirmed correct in the catalog):")
    print("=" * 70)
    for name, expected in KNOWN_MRP.items():
        rec = by_name.get(name)
        if rec is None:
            print("  {}  -> NOT FOUND in Tally's stock list under this exact name".format(name))
            continue
        val, why = rec
        if val is None:
            print("  {}".format(name))
            print("    expected (known-correct): {}".format(expected))
            print("    parsed MRP: MISSING ({})".format(why))
        else:
            match = "MATCH" if abs(val - expected) < 1 else "MISMATCH"
            print("  {}".format(name))
            print("    expected (known-correct): {}".format(expected))
            print("    parsed MRP from Tally: {}   [{}]  ({})".format(val, match, why))

    print("\n" + "=" * 70)
    print("Coverage: {} of {} total stock items have a parseable MRP.".format(have_mrp, total))
    print("=" * 70)

    print("\n" + "=" * 70)
    print("EASY-COPY RAW EXCERPTS (send these back too, especially if any item above")
    print("came back MISSING or MISMATCH - the actual tag names can be read from this):")
    print("=" * 70)
    for label, needle in [("SPRINTA F700", "SPRINTA F700"), ("URANIA 800", "URANIA 800")]:
        idx = xml_text.upper().find(needle)
        if idx != -1:
            block_start = xml_text.rfind("<STOCKITEM", 0, idx)
            block_end = xml_text.find("</STOCKITEM>", idx)
            block_end = block_end + len("</STOCKITEM>") if block_end != -1 else idx + 4000
            if block_start == -1:
                block_start = max(0, idx - 200)
            print("\n--- First <STOCKITEM> block containing '{}' (raw) ---".format(label))
            print(xml_text[block_start:block_end])
        else:
            print("\n--- '{}' not found in this response ---".format(needle))

    print("\n" + "=" * 70)
    print("Send this whole output back. rate/price are completely untouched by this -")
    print("fetch_stock()'s mrp output stays None until this is confirmed the same way")
    print("STANDARDPRICE was.")


def dump_report():
    """
    Diagnostic mode (added 8 Sep 2026, after the Collection-based bill data
    turned out to include already-settled old bills with no way to detect
    that): fetches Tally's own native "Bills Receivable" report - the same
    underlying data as the "Ledger Voucher Outstanding" screen - instead of
    reconstructing pending/settled status ourselves from raw bill records.
    Saves the raw response and prints an excerpt for Akila Auto Spares
    specifically (known real numbers to check the shape against: should
    show ONLY its genuinely still-pending bills, starting 2026-06-29, with
    NO Feb-2024 or Jul-2025 entries in it - those are the ones that should
    be gone if this report is doing the filtering correctly).
    """
    out_dir = os.path.dirname(os.path.abspath(__file__))
    report_path = os.path.join(out_dir, "tally_bills_receivable_dump.xml")

    print("Fetching Tally's native 'Bills Receivable' report...")
    report_xml = tally_request(BILLS_RECEIVABLE_REPORT_XML)
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report_xml)
    print("Saved full response to: {}  ({} bytes)".format(report_path, len(report_xml)))

    print("\n" + "=" * 70)
    print("EASY-COPY EXCERPT - paste this back:")
    print("=" * 70)

    idx = report_xml.upper().find("AKILA")
    if idx != -1:
        start = max(0, idx - 300)
        end = min(len(report_xml), idx + 2500)
        print("\n--- Section around 'AKILA' (raw) ---")
        print(report_xml[start:end])
    else:
        print("\n--- 'AKILA' not found anywhere in this report's response. "
              "First 2000 chars instead (so the overall shape can still be checked): ---")
        print(report_xml[:2000])
        print("\n(If the report came back empty or with an error/ID-not-found message, "
              "that's useful too - paste whatever printed above.)")


# ── Diagnostic-only, added 8 Sep 2026: MRP + dealer-group dumps ─────────
# Two follow-up questions from the office: (1) can MRP go live too, same as
# stock/rate — fetch_stock() deliberately leaves "mrp" as None today, since
# a plain STOCKITEM collection (NAME/CLOSINGBALANCE/CLOSINGRATE/
# CLOSINGVALUE) never carried it; MRP is commonly stored per-BATCH in Tally
# (relevant when batch-wise tracking is on, common for FMCG/lubricant
# distribution), not as a flat field on the item itself, so this widens the
# FETCH list and also requests BATCHALLOCATIONS.LIST (the same nested-LIST
# fetch trick that already worked for BILLALLOCATIONS.LIST on the
# outstanding side) to see whether MRP shows up there instead. (2) if a
# brand-new item/dealer in Tally should be auto-added straight into the
# right DSR's app, this needs to know which DSR/area a given ledger
# belongs to — dumps each named ledger's Tally group (PARENT) to see
# whether Tally's own grouping already encodes that, before any auto-add
# logic gets built around a guess.
STOCK_MRP_DUMP_XML = """<ENVELOPE>
 <HEADER>
  <VERSION>1</VERSION>
  <TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE>
  <ID>StockItemMRPDumpCollection</ID>
 </HEADER>
 <BODY>
  <DESC>
   <STATICVARIABLES>
    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
   </STATICVARIABLES>
   <TDL>
    <TDLMESSAGE>
     <COLLECTION NAME="StockItemMRPDumpCollection" ISMODIFY="No">
      <TYPE>StockItem</TYPE>
      <FETCH>NAME, CLOSINGBALANCE, CLOSINGRATE, CLOSINGVALUE, BASEUNITS, STANDARDCOSTPRICE, STANDARDSELLINGPRICE</FETCH>
      <FETCH>BATCHALLOCATIONS.LIST</FETCH>
     </COLLECTION>
    </TDLMESSAGE>
   </TDL>
  </DESC>
 </BODY>
</ENVELOPE>"""

LEDGER_GROUP_COLLECTION_XML = """<ENVELOPE>
 <HEADER>
  <VERSION>1</VERSION>
  <TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE>
  <ID>LedgerGroupDumpCollection</ID>
 </HEADER>
 <BODY>
  <DESC>
   <STATICVARIABLES>
    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
   </STATICVARIABLES>
   <TDL>
    <TDLMESSAGE>
     <COLLECTION NAME="LedgerGroupDumpCollection" ISMODIFY="No">
      <TYPE>Ledger</TYPE>
      <FETCH>NAME, PARENT, CLOSINGBALANCE</FETCH>
     </COLLECTION>
    </TDLMESSAGE>
   </TDL>
  </DESC>
 </BODY>
</ENVELOPE>"""

# Owner's eligibility rule (8 Sep 2026): a ledger under a known DSR group
# only counts as a real, currently-active dealer worth auto-adding if it
# has EITHER a voucher/transaction dated 1-Apr-2025 or later, OR a current
# debit (Dr) balance right now - not just group membership, since Tally's
# ledger master turned out to carry hundreds of old/inactive accounts per
# DSR that were never deleted (see --dump-all-groups: every DSR's resolved
# count came back roughly 1.5-2x its app's current dealer count). Fetches
# every voucher from that date onward and pulls both PARTYLEDGERNAME (the
# main party on Sales/Receipt/Payment-type vouchers) and every line inside
# ALLLEDGERENTRIES.LIST (catches ledgers only touched via a journal-style
# voucher that has no single "party") - the union of both is the "has
# recent activity" set. SVTODATE is filled in at request time with today's
# date.
ACTIVE_SINCE_DATE = "20250401"  # 1 Apr 2025, per the owner's rule
VOUCHER_ACTIVITY_COLLECTION_TEMPLATE = """<ENVELOPE>
 <HEADER>
  <VERSION>1</VERSION>
  <TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE>
  <ID>VoucherActivityDumpCollection</ID>
 </HEADER>
 <BODY>
  <DESC>
   <STATICVARIABLES>
    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
    <SVFROMDATE>{from_date}</SVFROMDATE>
    <SVTODATE>{to_date}</SVTODATE>
   </STATICVARIABLES>
   <TDL>
    <TDLMESSAGE>
     <COLLECTION NAME="VoucherActivityDumpCollection" ISMODIFY="No">
      <TYPE>Voucher</TYPE>
      <FETCH>DATE, PARTYLEDGERNAME</FETCH>
      <FETCH>ALLLEDGERENTRIES.LIST</FETCH>
     </COLLECTION>
    </TDLMESSAGE>
   </TDL>
  </DESC>
 </BODY>
</ENVELOPE>"""

# One known dealer per DSR app, so a single --dump-groups run (no args)
# covers every DSR's grouping in one go rather than needing 8 separate
# --check-style commands.
SAMPLE_DEALERS_BY_DSR = [
    ("Prabha",     "ARUN MOTORS, BODINAYAKANUR (OIL)"),
    ("Arun",       "GOLDEN AUTO SPARES, DHARAPURAM (CBE-OIL)"),
    ("Nagaraj",    "2 WHEELER SPARES, SVK (OIL)"),
    ("Ajith",      "ADHI PARASAKTHI AUTO GARAGE, MDU (OIL)"),  # was "Chellamani" - renamed 14 Nov 2026
    ("Anandh",     "AKSHAYA MOTORS, TENKASI (OIL)"),
    ("Madhu",      "ADLIN AUTOMOBILES, COLACHEL (OIL)"),
    ("Saravanan",  "ADS AUTO STORE,SVG (OIL)"),
    ("AAA Direct", "ADM ENTERPRISES,PONDY (OIL)"),
]


def dump_mrp():
    """
    Diagnostic mode: fetches stock items with a wider field list (including
    BATCHALLOCATIONS.LIST) and saves/prints raw excerpts so we can see
    whether MRP shows up anywhere in this Tally's data before wiring
    anything into fetch_stock(). Prints the raw block for a couple of
    well-known catalog items (searched by name fragment) plus the very
    first item in the response, in case the known names don't match
    Tally's own item naming exactly.
    """
    out_dir = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(out_dir, "tally_mrp_dump.xml")

    print("Fetching stock items with an expanded field list (incl. batch data) from Tally...")
    xml_text = tally_request(STOCK_MRP_DUMP_XML)
    with open(path, "w", encoding="utf-8") as f:
        f.write(xml_text)
    print("Saved full response to: {}  ({} bytes)".format(path, len(xml_text)))

    print("\n" + "=" * 70)
    print("EASY-COPY EXCERPTS - paste these back:")
    print("=" * 70)

    for label, needle in [("URANIA", "URANIA"), ("SPRINTA", "SPRINTA")]:
        idx = xml_text.upper().find(needle)
        if idx != -1:
            block_start = xml_text.rfind("<STOCKITEM", 0, idx)
            block_end = xml_text.find("</STOCKITEM>", idx)
            block_end = block_end + len("</STOCKITEM>") if block_end != -1 else idx + 3000
            if block_start == -1:
                block_start = max(0, idx - 200)
            print("\n--- First <STOCKITEM> block containing '{}' (raw) ---".format(label))
            print(xml_text[block_start:block_end])
        else:
            print("\n--- '{}' not found in this response ---".format(needle))

    si_start = xml_text.find("<STOCKITEM")
    if si_start != -1:
        si_end = xml_text.find("</STOCKITEM>", si_start)
        si_end = si_end + len("</STOCKITEM>") if si_end != -1 else si_start + 3000
        print("\n--- Very first <STOCKITEM> block overall (raw, for comparison) ---")
        print(xml_text[si_start:si_end])
    else:
        print("\n--- No <STOCKITEM tag found at all - first 1500 chars instead: ---")
        print(xml_text[:1500])

    print("\n" + "=" * 70)
    print("Send the excerpts above back. If none of them show anything MRP-shaped")
    print("(look for tags like <MRP>, <RATE>, or an <MRPRATE>/<PRICE> inside a")
    print("<BATCHALLOCATIONS.LIST> block), that likely means this Tally setup")
    print("doesn't track MRP as structured data at all, and it'll have to stay")
    print("on the Drive-sync cadence rather than going live.")


def dump_stock_unit():
    """
    Diagnostic mode: --dump-stock-unit - checks what UNIT fetch_stock()'s
    CLOSINGBALANCE is actually expressed in for live stock. This has NEVER
    been independently confirmed the way rate (STANDARDPRICE, confirmed
    11 Sep 2026) and mrp (MRPDETAILS.LIST, confirmed 11 Sep 2026) both
    were - it was only ever an inherited ASSUMPTION written into the
    ITEM_LTR_CATALOG/_catalog_ltr_guess() comment above ("the 'Case'/
    order-unit CLOSINGBALANCE already behaves for live stock"), stated
    while diagnosing a DIFFERENT field (a Sales voucher's own quantity).
    That separate field turned out to be in individual Nos, not Case,
    once real data was actually checked (18 Sep 2026 update above) - so
    the same "Case" assumption for CLOSINGBALANCE needs its own real-data
    check before being trusted, same discipline as every other Tally
    field in this project.

    Why this matters now: the owner flagged the dashboard's new Stocks-
    section Ltr totals as "CALCULATION IS WRONG" on 25 Sep 2026.
    stockCatalogItemLtr_() on the dashboard currently computes
    stock * nosPerUnit * sizeLtr (i.e. assumes CLOSINGBALANCE is in
    Case/order-units) - if it's actually in individual Nos like the
    Sales-voucher quantity was, every case-packed item's shown Ltr total
    is overstated by exactly its own nosPerUnit (e.g. 20x too high for a
    20-per-case item), which matches the kind of implausibly large total
    the owner's own screenshot showed.

    Prints CLOSINGBALANCE and BASEUNITS RAW (unparsed - keeping any unit
    suffix Tally attaches, e.g. "218.00 Nos" vs "20.00 Case") for a
    sample of real catalog items with different case sizes, reusing the
    same STOCK_MRP_DUMP_XML request dump_mrp() already fetches
    successfully (no new/unverified request shape here) - plus, for each
    item, BOTH candidate Ltr totals (today's Case-based formula vs the
    Nos-based alternative) so the owner can directly compare against a
    real known total (Daily Stock Update.xlsx or a physical count) and
    say which one is actually right.
    """
    out_dir = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(out_dir, "tally_stock_unit_dump.xml")

    print("Fetching stock items (incl. BASEUNITS) from Tally...")
    xml_text = tally_request(STOCK_MRP_DUMP_XML)
    with open(path, "w", encoding="utf-8") as f:
        f.write(xml_text)
    print("Saved full response to: {}  ({} bytes)".format(path, len(xml_text)))

    root = ET.fromstring(xml_text)
    sample_names = [
        "10W30 1 LIT SPRINTA F700 (20)",
        "CF4 15W40 1 LIT URANIA 800 (20)",
        "0W20 4 LIT SYNTIUM 7000 SN HYBRID FULLSYN (4)",
        "3 WHEELER LIFE ENGINE OIL 2.75 LIT PETRONAS (6)",
        "10W30 210 LIT SPRINTA F300 BARREL",
    ]

    print("\n" + "=" * 70)
    print("EASY-COPY EXCERPTS - paste these back:")
    print("=" * 70)

    found_any = False
    for si in root.iter("STOCKITEM"):
        name = (si.get("NAME") or _text(si, "NAME") or "").strip()
        if name not in sample_names:
            continue
        found_any = True
        raw_bal = _text(si, "CLOSINGBALANCE")
        raw_base = _text(si, "BASEUNITS")
        entry = ITEM_LTR_CATALOG.get(name)
        print("\n--- {} ---".format(name))
        print("  Raw CLOSINGBALANCE: {!r}".format(raw_bal))
        print("  Raw BASEUNITS:      {!r}".format(raw_base))
        if entry:
            bal_num = _num(raw_bal)
            if bal_num is not None:
                nos_per_unit = entry["nosPerUnit"] or 1
                case_guess = round(bal_num * nos_per_unit * entry["sizeLtr"], 2)
                nos_guess = round(bal_num * entry["sizeLtr"], 2)
                print("  Catalog: {} Ltr/unit, {} units per {}".format(entry["sizeLtr"], nos_per_unit, entry["unit"]))
                print("  If CLOSINGBALANCE is in {} (today's dashboard formula): {} Ltr".format(entry["unit"], case_guess))
                print("  If CLOSINGBALANCE is actually in individual Nos:        {} Ltr".format(nos_guess))

    if not found_any:
        print("\nNone of the sample item names matched exactly - printing the first")
        print("5 <STOCKITEM> blocks raw instead so the real names can be read off:")
        count = 0
        for si in root.iter("STOCKITEM"):
            if count >= 5:
                break
            name = (si.get("NAME") or _text(si, "NAME") or "").strip()
            print("\n--- {} ---".format(name))
            print("  Raw CLOSINGBALANCE: {!r}".format(_text(si, "CLOSINGBALANCE")))
            print("  Raw BASEUNITS:      {!r}".format(_text(si, "BASEUNITS")))
            count += 1

    print("\n" + "=" * 70)
    print("Send the excerpts above back. Two things settle this:")
    print("1. If CLOSINGBALANCE's own text carries a unit suffix (e.g. '218.00")
    print("   Nos' vs '20.00 Case'), that says it directly, same as the Sales-")
    print("   voucher quantity check did on 18 Sep 2026.")
    print("2. Whichever of the two Ltr totals above (Case-based vs Nos-based)")
    print("   matches your own known correct total (Daily Stock Update.xlsx or")
    print("   a physical count) for these items tells us which formula is")
    print("   right - the dashboard's Stocks-section total will be fixed to")
    print("   match once you confirm.")


# Group masters have their own PARENT (Tally groups nest: a sub-area group
# can itself sit under a broader DSR-level group) - fetched separately from
# ledgers so the full chain can be walked, not just one hop.
GROUP_HIERARCHY_COLLECTION_XML = """<ENVELOPE>
 <HEADER>
  <VERSION>1</VERSION>
  <TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE>
  <ID>GroupHierarchyDumpCollection</ID>
 </HEADER>
 <BODY>
  <DESC>
   <STATICVARIABLES>
    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
   </STATICVARIABLES>
   <TDL>
    <TDLMESSAGE>
     <COLLECTION NAME="GroupHierarchyDumpCollection" ISMODIFY="No">
      <TYPE>Group</TYPE>
      <FETCH>NAME, PARENT</FETCH>
     </COLLECTION>
    </TDLMESSAGE>
   </TDL>
  </DESC>
 </BODY>
</ENVELOPE>"""

# The 8 DSR-level group names the app/Drive-sync side actually uses (must
# exactly match the "DSR Name" column in Dealer Outstanding.xlsx / each
# app's dsrName) - a ledger only counts as belonging to a DSR once walking
# up its group chain lands on one of these, however many sub-area groups
# deep that takes.
KNOWN_DSR_GROUP_NAMES = [
    "PETRONAS THENI AREA (PRABHA)",
    "PETRONAS DINDIGUL AREA (ARUN)",
    "PETRONAS VIRUDHUNAGAR AREA (NAGARAJ)",
    "PETRONAS MDU AREA (AJITH)",  # was "(CHELLAMANI)" - owner renamed the real Tally group itself, 14 Nov 2026
    "PETRONAS TVL1 AREA (ANANDH)",
    "PETRONAS NKL AREA (MADHU)",
    "PETRONAS MDU AREA (SARAVANAN)",
    "PETRONAS AAA DIRECT AREA SALES",
]

# Added 18 Sep 2026 - Saravanan's dealers are their own Tally Group
# (Saravanan's role is on hold; his dealers were merged into Chellamani's
# order-slip APP on 1 Sep 2026, but that merge never touched Tally - his
# dealers still sit under his own separate Group there). Owner confirmed
# today (dashboard reporting Chellamani's Tally-actual Sales looking
# lower than her app-logged total) that both DSR-keyed live-data fetches
# (fetch_sales(), fetch_collection()) should fold Saravanan's Tally
# totals into Chellamani's before returning, so the dashboard's
# Tally-actual figure matches the same merged territory her app total
# already reflects - matching real dealers to a real DSR group, not
# double-counting or losing anything, just attributing his group's
# numbers to her card the same way the app already does.
_SARAVANAN_GROUP_NAME = "PETRONAS MDU AREA (SARAVANAN)"
_AJITH_GROUP_NAME = "PETRONAS MDU AREA (AJITH)"  # was "PETRONAS MDU AREA (CHELLAMANI)" - owner renamed the real Tally group itself, 14 Nov 2026


def _merge_saravanan_into_chellamani(dsr_totals):
    """Mutates dsr_totals in place: adds Saravanan's figures into
    Chellamani's entry (if both are present) and removes the standalone
    Saravanan entry, so callers/consumers never see him as a separate
    row. Safe to call even if one or both keys are missing.

    Generalized 21 Sep 2026 (was: hardcoded to just "today"/"month") to
    merge whatever numeric keys Saravanan's own record actually has -
    added for "lastMonthSameDate" (the new Tally-based "Last month" chip)
    without needing a matching one-off change here every time a new
    per-DSR field is added elsewhere. Every value in a dsr_totals/totals
    record is always a plain number, so summing every key is always safe.
    """
    sarv = dsr_totals.get(_SARAVANAN_GROUP_NAME)
    if not sarv:
        return
    chell = dsr_totals.setdefault(_AJITH_GROUP_NAME, {"today": 0.0, "month": 0.0})
    for key, val in sarv.items():
        chell[key] = (chell.get(key) or 0.0) + (val or 0.0)
    del dsr_totals[_SARAVANAN_GROUP_NAME]


# Added 11 Oct 2026 for daily_totals_by_dsr (the new per-DSR Sales/
# Collection heatmap feature) - same flat "dsr -> {key: number}" merge as
# _merge_saravanan_into_chellamani() above, but WITHOUT that function's
# {"today": 0.0, "month": 0.0} default seed for Chellamani's entry, which
# is wrong for a dict keyed by date instead of by today/month (would have
# silently injected two fake "today"/"month" = 0 entries into Chellamani's
# date-keyed dict on any day she herself had no entry yet before the
# merge). Safe if either/both keys are missing.
def _merge_saravanan_into_chellamani_flat(dsr_key_dict):
    sarv = dsr_key_dict.get(_SARAVANAN_GROUP_NAME)
    if not sarv:
        return
    chell = dsr_key_dict.setdefault(_AJITH_GROUP_NAME, {})
    for key, val in sarv.items():
        chell[key] = (chell.get(key) or 0.0) + (val or 0.0)
    del dsr_key_dict[_SARAVANAN_GROUP_NAME]


# Added 21 Sep 2026 for "MB" (dealers billed/dealers who paid this month) -
# same Saravanan-into-Chellamani fold as _merge_saravanan_into_chellamani()
# above, just for a plain {dsr: count} dict instead of a {dsr: {today,
# month}} dict, since that function's today/month math doesn't apply here.
def _merge_saravanan_count_into_chellamani(counts):
    sarv = counts.get(_SARAVANAN_GROUP_NAME)
    if not sarv:
        return
    counts[_AJITH_GROUP_NAME] = counts.get(_AJITH_GROUP_NAME, 0) + sarv
    del counts[_SARAVANAN_GROUP_NAME]


# Added 1 Oct 2026 for the historical-summary diagnostic's STEP C (monthly
# closing outstanding, by age-bucket, per DSR) - same Saravanan-into-
# Chellamani fold as _merge_saravanan_into_chellamani() above, but for a
# structure that's TWO levels deep: dsr_month_dict[dsr][month_key] is
# itself a dict of numeric fields (pendingAmount/band3059/band6089/
# band90plus), not one flat {field: number} dict per DSR. Calling the
# existing flat-case helper on this structure crashes ("unsupported
# operand type(s) for +: 'dict' and 'dict'" - exactly the traceback the
# owner's own first real run hit, since it tried to add a whole per-month
# band-dict to another band-dict as if it were a single number).
# Deliberately a separate function, not a change to
# _merge_saravanan_into_chellamani() itself (used at 5 other call sites
# for the flat case) - isolated here so this still-experimental STEP C
# diagnostic can never risk changing behaviour for an already-trusted
# call site.
def _merge_saravanan_into_chellamani_nested_dict(dsr_month_dict):
    sarv = dsr_month_dict.get(_SARAVANAN_GROUP_NAME)
    if not sarv:
        return
    chell = dsr_month_dict.setdefault(_AJITH_GROUP_NAME, {})
    for month_key, sarv_rec in sarv.items():
        chell_rec = chell.setdefault(month_key, {})
        for field, val in sarv_rec.items():
            chell_rec[field] = (chell_rec.get(field) or 0.0) + (val or 0.0)
    del dsr_month_dict[_SARAVANAN_GROUP_NAME]


def dump_all_groups():
    """
    Diagnostic mode: --dump-groups only checked one sample dealer per DSR
    and found at least one (Arun's) sitting under a sub-area group
    ("PETRONAS DHARAPURAM AREA (ARUN)") rather than the exact DSR-level
    group name the app itself uses - meaning a simple "ledger's immediate
    PARENT == DSR name" match would silently miss dealers under any other
    sub-area group. This fetches BOTH the full ledger list (NAME+PARENT)
    AND the full Group master's own hierarchy (NAME+PARENT for groups),
    then walks each ledger's group chain upward (ledger -> its group ->
    that group's parent -> ... ) until it lands on one of the 8 known
    DSR-level group names in KNOWN_DSR_GROUP_NAMES, however many hops that
    takes. Prints, for every group name touched by this walk, how many
    ledgers resolved through it and which top-level DSR they landed on -
    so the real breadth of the grouping can be seen before any auto-add
    logic gets built on top of it.
    """
    print("Fetching full ledger list (name + group) from Tally...")
    ledger_xml = tally_request(LEDGER_GROUP_COLLECTION_XML)
    ledger_root = ET.fromstring(ledger_xml)
    ledger_parent = {}
    for led in ledger_root.iter("LEDGER"):
        name = led.get("NAME") or _text(led, "NAME")
        parent = _text(led, "PARENT")
        if name:
            ledger_parent[name.strip()] = (parent or "").strip()
    print("  {} ledgers read.".format(len(ledger_parent)))

    print("Fetching full Group master hierarchy from Tally...")
    group_xml = tally_request(GROUP_HIERARCHY_COLLECTION_XML)
    group_root = ET.fromstring(group_xml)
    group_parent = {}
    for grp in group_root.iter("GROUP"):
        name = grp.get("NAME") or _text(grp, "NAME")
        parent = _text(grp, "PARENT")
        if name:
            group_parent[name.strip()] = (parent or "").strip()
    print("  {} groups read.".format(len(group_parent)))

    known_set = set(KNOWN_DSR_GROUP_NAMES)

    def resolve_dsr(immediate_group):
        """Walk up the group chain from a ledger's immediate group until a
        known DSR-level group name is hit, or the chain runs out (root, a
        loop, or an unrecognised top-level group). Returns
        (resolved_dsr_or_None, chain_list)."""
        chain = []
        current = immediate_group
        seen = set()
        while current and current not in seen:
            chain.append(current)
            seen.add(current)
            if current in known_set:
                return current, chain
            current = group_parent.get(current, "")
        return None, chain

    # Only look at ledgers whose chain actually touches something
    # PETRONAS/AREA-shaped, so this doesn't dump the whole 4,140-ledger
    # company chart of accounts (bank accounts, tax ledgers, etc.).
    per_dsr_count = {}
    unresolved_but_relevant = []
    all_touched_group_names = set()
    for lname, lparent in ledger_parent.items():
        if not lparent:
            continue
        dsr, chain = resolve_dsr(lparent)
        touched_relevant = any(("PETRONAS" in g.upper() or "AREA" in g.upper()) for g in chain)
        if dsr:
            per_dsr_count[dsr] = per_dsr_count.get(dsr, 0) + 1
            all_touched_group_names.update(chain)
        elif touched_relevant:
            unresolved_but_relevant.append((lname, chain))

    print("\n" + "=" * 70)
    print("Ledgers resolved to a known DSR group (however many hops up):")
    print("=" * 70)
    for dsr in KNOWN_DSR_GROUP_NAMES:
        print("  {}  ->  {} ledgers".format(dsr, per_dsr_count.get(dsr, 0)))

    print("\n" + "=" * 70)
    print("Every distinct group name that appears in a resolved chain")
    print("(these are the sub-area groups actually in use, if any):")
    print("=" * 70)
    for g in sorted(all_touched_group_names):
        print("  {}".format(g))

    if unresolved_but_relevant:
        print("\n" + "=" * 70)
        print("PETRONAS/AREA-ish ledgers that did NOT resolve to any of the 8 known")
        print("DSR groups (worth a look - possibly a typo'd group, a genuinely new")
        print("DSR/area, or a group this list doesn't know about yet):")
        print("=" * 70)
        for lname, chain in unresolved_but_relevant[:40]:
            print("  {}  ->  {}".format(lname, " -> ".join(chain) if chain else "(no group)"))
        if len(unresolved_but_relevant) > 40:
            print("  ... and {} more".format(len(unresolved_but_relevant) - 40))

    print("\n" + "=" * 70)
    print("Send this whole output back.")


def fetch_active_ledger_names(from_date=ACTIVE_SINCE_DATE):
    """
    Returns the set of ledger names with at least one voucher dated
    from_date (YYYYMMDD) or later - the union of each voucher's
    PARTYLEDGERNAME and every ledger name inside its ALLLEDGERENTRIES.LIST,
    so both party-style vouchers (Sales/Receipt/Payment) and journal-style
    ones (no single "party") are covered.
    """
    to_date = datetime.now().strftime("%Y%m%d")
    xml_body = VOUCHER_ACTIVITY_COLLECTION_TEMPLATE.format(from_date=from_date, to_date=to_date)
    # Timeout raised 120s -> 600s on 9 Sep 2026: with the roster cadence now
    # a once-daily 7:30 PM check (see ROSTER_DAILY_HOUR above) rather than
    # every ~4h, this request only runs once a day and is no longer on the
    # critical path for the fast 10-min stock/outstanding polls - so it's
    # safe to let it take longer. This was necessary, not just generous:
    # the very first evening under the new daily cadence (9 Sep 2026) hit
    # the old 120s timeout twice in a row at 19:32:17 and 19:44:42, both
    # with the identical "Read timed out (read timeout=120)" error - this
    # office's full voucher history since ACTIVE_SINCE_DATE (currently
    # 1 Apr 2025, 1.5+ years) genuinely takes Tally longer than 2 minutes
    # to export. If 600s still isn't enough on some future evening, either
    # raise this further or - a better long-term fix if that becomes
    # routine - narrow ACTIVE_SINCE_DATE's window (e.g. to the last 6-12
    # months) so this pull covers less history each time; a shorter window
    # only affects how far back a still-open OLD debt can make a ledger
    # "eligible", not whether a genuinely new/recent dealer is caught.
    xml_text = tally_request(xml_body, timeout=600)
    root = ET.fromstring(xml_text)
    active = set()
    for v in root.iter("VOUCHER"):
        party = _text(v, "PARTYLEDGERNAME")
        if party:
            active.add(party.strip())
        for entry in v.iter("ALLLEDGERENTRIES.LIST"):
            lname = _text(entry, "LEDGERNAME")
            if lname:
                active.add(lname.strip())
    return active


# When the new-dealer roster gets (re)computed and pushed - separate from
# POLL_SECONDS (the 10-min stock/outstanding cadence), since this one also
# fetches the whole company's voucher history since ACTIVE_SINCE_DATE (see
# fetch_active_ledger_names()), which is far heavier and slower than
# anything else this script does.
#
# CHANGED 9 Sep 2026, per the owner's request: was a rolling "every ~4h
# since the last attempt" interval (ROSTER_INTERVAL_SECONDS = 4 * 3600),
# now a fixed once-a-day check at ROSTER_DAILY_HOUR:ROSTER_DAILY_MINUTE
# (device local time). Two reasons: (1) the owner wants dealers added in
# Tally picked up and pushed to their DSR's app once daily by early
# evening, not on a rolling multi-hour clock; (2) the old interval logic
# only advanced its timestamp on a SUCCESSFUL push, so once the 4h
# threshold had passed, a FAILING push (e.g. Tally's voucher-history query
# timing out) retried on every single 10-min poll instead of waiting
# another 4h - this actually happened 9 Sep 2026 (8 straight timeouts,
# ~12 min apart, over 90+ minutes) and is what prompted this change.
# Tracked by calendar date (state["last_roster_push_date"], a "YYYY-MM-DD"
# string) rather than elapsed seconds, so a push that fails still retries
# on the next poll - same evening, until it succeeds or the day rolls
# over - but a SUCCESSFUL push only needs to happen once per day.
ROSTER_DAILY_HOUR = 19
ROSTER_DAILY_MINUTE = 30


def compute_eligible_roster():
    """
    Returns {dsr_group_name: [dealer_name, ...]} - every ledger that
    resolves (via its full group chain, however many hops) to one of the 8
    KNOWN_DSR_GROUP_NAMES, AND passes the owner's eligibility rule: a
    voucher dated ACTIVE_SINCE_DATE or later, OR a current debit balance.
    Same logic as --dump-eligible (kept as a separate, deliberately
    untouched function so the proven diagnostic never risks drifting out of
    sync with what actually gets pushed) - this is the version that feeds
    the real liveDealerRosterUpdate push in maybe_push_dealer_roster()
    below, returning full name lists instead of just printing counts.
    """
    ledger_xml = tally_request(LEDGER_GROUP_COLLECTION_XML)
    ledger_root = ET.fromstring(ledger_xml)
    ledger_parent = {}
    ledger_balance = {}
    for led in ledger_root.iter("LEDGER"):
        name = led.get("NAME") or _text(led, "NAME")
        if not name:
            continue
        name = name.strip()
        ledger_parent[name] = (_text(led, "PARENT") or "").strip()
        ledger_balance[name] = _num(_text(led, "CLOSINGBALANCE"))

    group_xml = tally_request(GROUP_HIERARCHY_COLLECTION_XML)
    group_root = ET.fromstring(group_xml)
    group_parent = {}
    for grp in group_root.iter("GROUP"):
        gname = grp.get("NAME") or _text(grp, "NAME")
        if gname:
            group_parent[gname.strip()] = (_text(grp, "PARENT") or "").strip()

    active_names = fetch_active_ledger_names()
    known_set = set(KNOWN_DSR_GROUP_NAMES)

    def resolve_dsr(immediate_group):
        current = immediate_group
        seen = set()
        while current and current not in seen:
            seen.add(current)
            if current in known_set:
                return current
            current = group_parent.get(current, "")
        return None

    roster = dict((dsr, []) for dsr in KNOWN_DSR_GROUP_NAMES)
    for lname, lparent in ledger_parent.items():
        if not lparent:
            continue
        dsr = resolve_dsr(lparent)
        if not dsr:
            continue
        bal = ledger_balance.get(lname)
        is_dr = bal is not None and bal < -0.5
        has_activity = lname in active_names
        if is_dr or has_activity:
            roster[dsr].append(lname)
    return roster


def maybe_push_dealer_roster(state, dry_run, ts):
    """
    Runs compute_eligible_roster() and pushes it, but only once per
    calendar day, at or after ROSTER_DAILY_HOUR:ROSTER_DAILY_MINUTE
    (device local time) - see the note above those constants for why this
    changed 9 Sep 2026 from a rolling ~4h interval. Called from run_once()
    after the fast stock/outstanding work is already done and sent, so a
    slow roster computation never delays those. Never runs in --test mode
    (--dump-eligible already exists as the dedicated way to preview this).

    state["last_roster_push_date"] ("YYYY-MM-DD") tracks the last day a
    push actually SUCCEEDED. On any given day: before
    ROSTER_DAILY_HOUR:ROSTER_DAILY_MINUTE, this is a no-op every poll;
    from that time onward, it runs on the next poll and, if it fails,
    keeps retrying on every subsequent poll (still that same evening)
    until it succeeds - a failure never sets last_roster_push_date, so it
    is not treated as "done for today." Once it succeeds, it will not run
    again until a new calendar date's ROSTER_DAILY_HOUR:ROSTER_DAILY_MINUTE
    arrives.

    "New dealers" diff, added 19 Sep 2026 (Arun-only trial): this function
    already recomputes the FULL eligible roster every day (old and new
    dealers mixed together, no "since when" signal of its own) - state
    ["known_roster_names"] now remembers every name this function has EVER
    seen in that roster, per DSR, so a name showing up here that was NEVER
    in known_roster_names before is a genuine "Tally has never called this
    dealer active or in-debt until today" signal - unlike the Orders
    sheet's order history, which only goes back to when the app itself
    launched and so can't tell a real new dealer from a long-time customer
    who simply hadn't ordered through the app before (see doGet()'s own
    comment in Code.gs for the real bug this replaces).

    First run ever with this feature (state["roster_baseline_seeded"] is
    False): today's FULL roster becomes the known baseline and NOTHING is
    reported as new - this is a one-time seed, not a real "0 new dealers
    today," so every dealer already in Tally today is correctly treated as
    already-known rather than flooding the app with false "new" dealers on
    rollout day. From the next successful run onward, only names that
    genuinely weren't in the baseline get pushed to LiveNewDealers, and are
    then folded into known_roster_names so they're never re-flagged.
    """
    if dry_run:
        return state
    now = datetime.now()
    today_str = now.strftime("%Y-%m-%d")
    if state.get("last_roster_push_date") == today_str:
        return state  # already pushed successfully today
    if (now.hour, now.minute) < (ROSTER_DAILY_HOUR, ROSTER_DAILY_MINUTE):
        return state  # not time yet today
    print("[{}] computing new-dealer roster (runs once daily at {:02d}:{:02d}, not every "
          "poll - this step can take a while, it's a full voucher-history pull)...".format(
              ts, ROSTER_DAILY_HOUR, ROSTER_DAILY_MINUTE))
    try:
        roster = compute_eligible_roster()
    except Exception as e:
        print("[{}] dealer-roster computation failed this round (will retry at the "
              "next poll, still today, until it succeeds): {}".format(ts, e))
        return state
    total_sent = 0
    for dsr in KNOWN_DSR_GROUP_NAMES:
        names = roster.get(dsr, [])
        if not names:
            continue
        payload = [{"dsrName": dsr, "name": n} for n in names]
        sent = post_backend_chunked("liveDealerRosterUpdate", "roster", payload, ts)
        total_sent += sent
    print("[{}] dealer roster: sent {} eligible dealer row(s) total across {} DSR(s)".format(
        ts, total_sent, len(KNOWN_DSR_GROUP_NAMES)))

    if not state.get("roster_baseline_seeded"):
        state["known_roster_names"] = dict((dsr, list(roster.get(dsr, []))) for dsr in KNOWN_DSR_GROUP_NAMES)
        state["roster_baseline_seeded"] = True
        total_baseline = sum(len(v) for v in state["known_roster_names"].values())
        print("[{}] new-dealers: first run with this feature - treating today's {} eligible "
              "dealer(s) as the existing baseline, not new. Genuinely new dealers will start "
              "showing up from tomorrow's roster run.".format(ts, total_baseline))
    else:
        new_by_dsr = {}
        for dsr in KNOWN_DSR_GROUP_NAMES:
            known = set(state["known_roster_names"].get(dsr, []))
            new_names = [n for n in roster.get(dsr, []) if n not in known]
            if new_names:
                new_by_dsr[dsr] = new_names
        total_new = 0
        for dsr, names in new_by_dsr.items():
            payload = [{"dsrName": dsr, "name": n, "firstSeenDate": today_str} for n in names]
            total_new += post_backend_chunked("liveNewDealerUpdate", "newDealers", payload, ts)
        print("[{}] new-dealers: {} genuinely new dealer(s) today across {} DSR(s)".format(
            ts, total_new, len(new_by_dsr)))
        # Fold today's full roster into known_roster_names regardless of
        # whether it had anything new, so a dealer who briefly stopped
        # qualifying and comes back later is never re-flagged as new.
        for dsr in KNOWN_DSR_GROUP_NAMES:
            merged = set(state["known_roster_names"].get(dsr, [])) | set(roster.get(dsr, []))
            state["known_roster_names"][dsr] = list(merged)

    state["last_roster_push_date"] = today_str
    return state


# --push-roster-now - a one-shot command to force the dealer roster
# (LiveDealerRoster push + new-dealer diffing) right away, without
# waiting for the once-daily ROSTER_DAILY_HOUR:ROSTER_DAILY_MINUTE
# window, and without needing a new calendar date even if today's
# regular window already ran - same pattern and same reason as
# --push-srep-summary-now/--push-petronas-summary-now above. REAL push
# (post_backend_chunked() writes to LiveDealerRoster), not a diagnostic
# - --dump-eligible already exists for a no-push preview. Added 14 Nov
# 2026 so the owner could verify a live Tally-side rename (his old
# "PETRONAS MDU AREA (CHELLAMANI)" group renamed to "...(AJITH)") landed
# correctly the same day, instead of waiting for that day's 19:30 run
# (already passed) or the next day's.
#
# Deliberately reuses the exact same new-dealer diffing logic as
# maybe_push_dealer_roster() (against known_roster_names, with the same
# first-run baseline-seeding) rather than a stripped-down "just push,
# skip the diffing" version - running this command must behave
# identically to a normal scheduled run (same dedup, same state
# updates), never produce a second, different "new today" list just
# because it ran through this door instead of the regular one. The ONLY
# difference from maybe_push_dealer_roster() is that this ignores
# state["last_roster_push_date"] entirely (so it runs regardless of the
# time of day or whether today's regular run already happened) -
# everything else, including setting last_roster_push_date to today's
# date at the end, is identical, so the regular 19:30 check later today
# correctly sees "already done today" and skips a redundant second run.
def push_everything_now():
    """--push-everything-now - one-shot: clears every maybe_push_*()
    feature's own "last pushed at/on" throttle in state, then calls
    run_once() - the EXACT same code the main polling loop calls every
    600 seconds - so the very next run genuinely re-pushes everything
    (stock, outstanding, sales, collection, daybook, dealer roster,
    Petronas summary, S.REP summary, Collection Group Summary)
    regardless of how recently each last ran, instead of waiting out
    each feature's own interval (1-4 hours, or once daily for the
    roster) one at a time.

    Deliberately reuses run_once() itself rather than re-deriving each
    feature's diff/push logic by hand here - this file's standing "run
    the real function, don't reimplement" rule (same reasoning as
    compute_eligible_roster() vs. dump_eligible() above). Stock and
    outstanding have no throttle of their own (they already push every
    poll) so nothing needs clearing for those two.

    Added 14 Nov 2026 so the owner could force a same-day full resync
    after a live Tally-side change (a DSR group rename) instead of
    waiting out each feature's own gate separately. Does NOT touch
    known_roster_names/roster_baseline_seeded (new-dealer detection) or
    any other state besides the timers listed below, so running this
    can never wrongly flag an existing dealer as "new" just because it
    ran through this door.
    """
    state = load_state()
    for key in ("last_sales_push_ts", "last_collection_push_ts",
                "last_daybook_push_ts", "last_petronas_summary_push_ts",
                "last_srep_summary_push_ts", "last_group_summary_push_ts",
                "last_roster_push_date"):
        state.pop(key, None)
    print("Forcing a full refresh now (stock, outstanding, sales, collection, "
          "daybook, dealer roster, Petronas summary, S.REP summary, Collection "
          "Group Summary) - this can take a few minutes, it's everything at "
          "once, including a full voucher-history pull for the roster...")
    state = run_once(state, dry_run=False)
    save_state(state)
    print("\nDone. Refresh the dashboard (or wait a few seconds for its own "
          "auto-refresh) to see everything updated.")


def push_dealer_roster_now():
    state = load_state()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    today_str = datetime.now().strftime("%Y-%m-%d")
    print("[{}] computing eligible dealer roster now (ignoring the usual {:02d}:{:02d} "
          "daily window) - this can take a while, it's a full voucher-history pull...".format(
              ts, ROSTER_DAILY_HOUR, ROSTER_DAILY_MINUTE))
    roster = compute_eligible_roster()
    total_sent = 0
    for dsr in KNOWN_DSR_GROUP_NAMES:
        names = roster.get(dsr, [])
        if not names:
            continue
        payload = [{"dsrName": dsr, "name": n} for n in names]
        sent = post_backend_chunked("liveDealerRosterUpdate", "roster", payload, ts)
        total_sent += sent
    print("[{}] dealer roster: sent {} eligible dealer row(s) total across {} DSR(s)".format(
        ts, total_sent, len(KNOWN_DSR_GROUP_NAMES)))

    if not state.get("roster_baseline_seeded"):
        state["known_roster_names"] = dict((dsr, list(roster.get(dsr, []))) for dsr in KNOWN_DSR_GROUP_NAMES)
        state["roster_baseline_seeded"] = True
        total_baseline = sum(len(v) for v in state["known_roster_names"].values())
        print("[{}] new-dealers: first run with this feature - treating today's {} eligible "
              "dealer(s) as the existing baseline, not new.".format(ts, total_baseline))
    else:
        new_by_dsr = {}
        for dsr in KNOWN_DSR_GROUP_NAMES:
            known = set(state["known_roster_names"].get(dsr, []))
            new_names = [n for n in roster.get(dsr, []) if n not in known]
            if new_names:
                new_by_dsr[dsr] = new_names
        total_new = 0
        for dsr, names in new_by_dsr.items():
            payload = [{"dsrName": dsr, "name": n, "firstSeenDate": today_str} for n in names]
            total_new += post_backend_chunked("liveNewDealerUpdate", "newDealers", payload, ts)
        print("[{}] new-dealers: {} genuinely new dealer(s) today across {} DSR(s)".format(
            ts, total_new, len(new_by_dsr)))
        for dsr in KNOWN_DSR_GROUP_NAMES:
            merged = set(state["known_roster_names"].get(dsr, [])) | set(roster.get(dsr, []))
            state["known_roster_names"][dsr] = list(merged)

    state["last_roster_push_date"] = today_str
    save_state(state)
    print("\nDone. Refresh the dashboard (or wait a few seconds for its own auto-refresh) "
          "to see the updated roster.")


def dump_eligible():
    """
    Diagnostic mode (added 8 Sep 2026, after --dump-all-groups showed every
    DSR's Tally-resolved dealer count running ~1.5-2x its app's current
    count): applies the owner's actual eligibility rule - a ledger under a
    known DSR group counts as a real, currently-active dealer only if it
    has a voucher dated 1-Apr-2025 or later, OR a current debit (Dr)
    balance - and reports how many ledgers per DSR pass, so the counts can
    be sanity-checked against reality before this rule gets wired into the
    live auto-add pipeline for real.
    """
    print("Fetching full ledger list (name + group + balance) from Tally...")
    ledger_xml = tally_request(LEDGER_GROUP_COLLECTION_XML)
    ledger_root = ET.fromstring(ledger_xml)
    ledger_parent = {}
    ledger_balance = {}
    for led in ledger_root.iter("LEDGER"):
        name = led.get("NAME") or _text(led, "NAME")
        if not name:
            continue
        name = name.strip()
        ledger_parent[name] = (_text(led, "PARENT") or "").strip()
        ledger_balance[name] = _num(_text(led, "CLOSINGBALANCE"))
    print("  {} ledgers read.".format(len(ledger_parent)))

    print("Fetching full Group master hierarchy from Tally...")
    group_xml = tally_request(GROUP_HIERARCHY_COLLECTION_XML)
    group_root = ET.fromstring(group_xml)
    group_parent = {}
    for grp in group_root.iter("GROUP"):
        gname = grp.get("NAME") or _text(grp, "NAME")
        if gname:
            group_parent[gname.strip()] = (_text(grp, "PARENT") or "").strip()
    print("  {} groups read.".format(len(group_parent)))

    print("Fetching every voucher from {} onward (this can take a while - "
          "it's the whole company's transaction history since then)...".format(ACTIVE_SINCE_DATE))
    active_names = fetch_active_ledger_names()
    print("  {} distinct ledger names have at least one voucher in range.".format(len(active_names)))

    known_set = set(KNOWN_DSR_GROUP_NAMES)

    def resolve_dsr(immediate_group):
        current = immediate_group
        seen = set()
        while current and current not in seen:
            seen.add(current)
            if current in known_set:
                return current
            current = group_parent.get(current, "")
        return None

    resolved_count = {}
    eligible_count = {}
    eligible_samples = {}
    for lname, lparent in ledger_parent.items():
        if not lparent:
            continue
        dsr = resolve_dsr(lparent)
        if not dsr:
            continue
        resolved_count[dsr] = resolved_count.get(dsr, 0) + 1
        bal = ledger_balance.get(lname)
        is_dr = bal is not None and bal < -0.5
        has_activity = lname in active_names
        if is_dr or has_activity:
            eligible_count[dsr] = eligible_count.get(dsr, 0) + 1
            eligible_samples.setdefault(dsr, []).append(
                "{}  (Dr balance: {}, recent activity: {})".format(lname, is_dr, has_activity))

    print("\n" + "=" * 70)
    print("Per DSR: total resolved in Tally  vs.  eligible under the rule")
    print("(has a voucher since {} OR a current debit balance)".format(ACTIVE_SINCE_DATE))
    print("=" * 70)
    for dsr in KNOWN_DSR_GROUP_NAMES:
        print("  {}  ->  {} resolved, {} eligible".format(
            dsr, resolved_count.get(dsr, 0), eligible_count.get(dsr, 0)))

    print("\n" + "=" * 70)
    print("Sample of up to 10 eligible ledgers per DSR (for a quick sanity check):")
    print("=" * 70)
    for dsr in KNOWN_DSR_GROUP_NAMES:
        samples = eligible_samples.get(dsr, [])
        print("\n{}:".format(dsr))
        if not samples:
            print("  (none)")
        for s in samples[:10]:
            print("  " + s)
        if len(samples) > 10:
            print("  ... and {} more".format(len(samples) - 10))

    print("\n" + "=" * 70)
    print("Send this whole output back.")


def dump_groups(names=None):
    """
    Diagnostic mode: fetches every ledger's own Tally group (PARENT) and
    prints it for a set of known dealers (one per DSR by default, or
    whatever names are passed on the command line) - so we can see whether
    Tally's own grouping already tells us which DSR/area a ledger belongs
    to, before building any auto-add-new-dealer logic around a guess.
    """
    pairs = names if names else [n for _, n in SAMPLE_DEALERS_BY_DSR]
    dsr_by_name = dict((n, d) for d, n in SAMPLE_DEALERS_BY_DSR)

    print("Fetching ledger name + group (PARENT) data from Tally...")
    xml_text = tally_request(LEDGER_GROUP_COLLECTION_XML)
    root = ET.fromstring(xml_text)
    parent_by_name = {}
    for led in root.iter("LEDGER"):
        name = led.get("NAME") or _text(led, "NAME")
        parent = _text(led, "PARENT")
        if name:
            parent_by_name[name.strip()] = (parent or "").strip()

    print("\n" + "=" * 70)
    print("Dealer  ->  Tally Group (PARENT)   [known DSR, if we have one]")
    print("=" * 70)
    for target in pairs:
        target_upper = target.upper()
        exact = parent_by_name.get(target)
        if exact is not None:
            known_dsr = dsr_by_name.get(target, "")
            print("  {}  ->  {}{}".format(target, exact or "(blank)",
                  "   [known DSR: {}]".format(known_dsr) if known_dsr else ""))
            continue
        # fallback: substring match, same as --check
        found = False
        for name, parent in parent_by_name.items():
            if target_upper in name.upper():
                known_dsr = dsr_by_name.get(target, "")
                print("  {} (matched '{}')  ->  {}{}".format(
                    target, name, parent or "(blank)",
                    "   [known DSR: {}]".format(known_dsr) if known_dsr else ""))
                found = True
                break
        if not found:
            print("  {}  ->  NOT FOUND in Tally's ledger list".format(target))

    print("\n" + "=" * 70)
    print("Send this output back. If every dealer's group cleanly matches its")
    print("known DSR/area (e.g. all of Madhu's dealers share one group, all of")
    print("Arun's share a different one), we can auto-detect a new ledger's DSR")
    print("from its Tally group. If the groups look unrelated to DSR/area (e.g.")
    print("everything is just 'Sundry Debtors'), we'll need a different approach.")


def check_dealers(search_terms):
    """
    --check "name fragment" [more fragments...] - prints only the dealers
    whose name contains any of the given fragments (case-insensitive),
    against a fresh live computation. Added 8 Sep 2026 so specific
    already-verified dealers (Akila, Challenger Motors, A2Z, ...) can be
    re-checked precisely after a code change, instead of scrolling through
    a --test run's thousands of lines / random sample hoping to spot them.
    """
    print("Fetching live outstanding data from Tally...")
    all_ledger_balances = fetch_all_ledgers_with_balance()
    outstanding = fetch_outstanding(all_ledger_balances)
    terms = [t.upper() for t in search_terms]
    matched = False
    for name in sorted(outstanding.keys()):
        if any(t in name.upper() for t in terms):
            matched = True
            rec = outstanding[name]
            status = rec["since"] or "CLEARED (not outstanding)"
            bal = all_ledger_balances.get(name)
            print("  {}  ->  {}   (balance: {}, pending on open bills: {}, "
                  "aged >{}d: {})".format(
                name, status, bal, rec["pendingAmount"],
                OUTSTANDING_REMINDER_AGE_DAYS, rec["agedPendingAmount"]))
    if not matched:
        print("No dealer name matched any of: {}".format(", ".join(search_terms)))


# ── DSR Collection (Receipt vouchers), diagnostic added 12 Sep 2026 ──────
# The owner wants a new "Collection" figure per DSR - the total of all
# Receipt vouchers posted against that DSR's dealers this month, no manual
# entry anywhere, drilled down in the app as Target (owner-supplied) vs
# Collection (this) vs Balance. His own words: "COLLECTION MEANS ALL RCPT
# VOUCHERS" - a narrower definition than the office's own
# AAA_Group_Monthly_Collection.tdl report, which sums ANY credit-side
# ledger entry across a whole month (would also catch credit notes/journal
# credits), not Receipt vouchers specifically. That TDL is still useful
# context though: it confirms Tally already groups each DSR's dealers under
# their own named Group (the same KNOWN_DSR_GROUP_NAMES already proven for
# the outstanding-balance/roster features above), so a per-DSR Collection
# query is just "sum Receipt-voucher entries for ledgers under DSR X's
# group" - a normal Tally query, not a new kind of thing.
#
# This diagnostic (--check-collection) fetches every voucher for the
# current month (or an explicit range), resolves each ledger entry's dealer
# to a DSR the same way compute_eligible_roster() does (walk the ledger's
# Group chain up to a KNOWN_DSR_GROUP_NAMES entry), and for ONE requested
# DSR prints every distinct VOUCHERTYPENAME seen among its dealers' entries
# this period with its own credit/debit subtotal - so the real Receipt-type
# voucher name on THIS Tally setup, and which side (credit/debit) the money
# actually collected shows up on, can be confirmed from Tally's own output
# rather than assumed. Nothing is pushed anywhere by this - diagnostic
# only, same discipline as every other new Tally field in this file
# (rate/MRP/outstanding all went through this exact kind of check first).
#
# PERFORMANCE FIX, 14 Sep 2026 - the owner reported Tally itself running
# slow after this feature went live. Root cause: this request originally
# pulled EVERY voucher of EVERY type across the WHOLE company for the
# month-to-date, then filtered down to "Receipt" only in Python - the
# 1-12 Sep test alone came back with 50,101 total vouchers company-wide,
# of which only 39 actually touched Arun's dealers (and of those, only
# 22 were even Receipt type) - so Tally was doing a huge amount of
# needless export work every 30 minutes, all day, competing for the same
# engine a DSR/office staffer is actively using. **Fixed** by adding a
# `<FILTER>` referencing a `<SYSTEM TYPE="Formulae">` formula
# (`ReceiptVouchersOnly`, same ad hoc filter-formula pattern the office's
# own AAA_Group_Monthly_Collection.tdl uses for `NOTIsOptionalVoucher`,
# just written for a raw XML request instead of a persistent .tdl file)
# so Tally itself only exports Receipt-type vouchers, not the whole
# company's transaction history - this should cut the payload by roughly
# two orders of magnitude based on the 1-12 Sep sample. The Python-side
# `vtype.lower() != "receipt"` check in fetch_collection()/
# check_collection() below is left in place regardless, as a safety net
# in case this filter doesn't take effect exactly as expected on this
# Tally version - correctness never depended on it working, only speed
# does. **Not yet independently verified that the filter actually reduces
# the returned voucher count on this Tally setup** - fetch_collection()
# now prints how many vouchers came back each time it runs for real, so
# this can be confirmed from the watcher's own console output on its next
# run, and check_collection() now reports the same count too. Also
# widened COLLECTION_POLL_INTERVAL_SECONDS from 30 to 60 minutes as an
# extra, zero-risk safety margin on top of the filter (see that constant
# below) - Collection doesn't need to be as fresh as stock/outstanding.
COLLECTION_VOUCHERS_TEMPLATE = """<ENVELOPE>
 <HEADER>
  <VERSION>1</VERSION>
  <TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE>
  <ID>CollectionVouchersDumpCollection</ID>
 </HEADER>
 <BODY>
  <DESC>
   <STATICVARIABLES>
    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
    <SVFROMDATE>{from_date}</SVFROMDATE>
    <SVTODATE>{to_date}</SVTODATE>
   </STATICVARIABLES>
   <TDL>
    <TDLMESSAGE>
     <SYSTEM TYPE="Formulae" NAME="ReceiptVouchersOnly">$VoucherTypeName = "Receipt" OR $VoucherTypeName = "Cheque"</SYSTEM>
     <COLLECTION NAME="CollectionVouchersDumpCollection" ISMODIFY="No">
      <TYPE>Voucher</TYPE>
      <FETCH>DATE, VOUCHERTYPENAME, PARTYLEDGERNAME</FETCH>
      <FETCH>ALLLEDGERENTRIES.LIST</FETCH>
      <FILTER>ReceiptVouchersOnly</FILTER>
     </COLLECTION>
    </TDLMESSAGE>
   </TDL>
  </DESC>
 </BODY>
</ENVELOPE>"""

# Widened 18 Sep 2026 to also include "Cheque" vouchers, per the owner's
# explicit request ("in collection part i think data from cheque vouchers
# has not added") - Collection was originally scoped to "Receipt" only,
# matching his own 12 Sep 2026 words ("COLLECTION MEANS ALL RCPT
# VOUCHERS"), but he confirmed today that Cheque should count too. The
# real 12 Sep --check-collection sample already showed a real, non-zero
# "Cheque" voucher type for Arun's dealers (2 entries, Rs 32,500 over 12
# days) landing in the same positive-amount bucket as Receipt - structurally
# trusted on that basis rather than a fresh dedicated verification round.
# `COLLECTION_VOUCHERS_TEMPLATE`'s server-side `<FILTER>` formula and the
# client-side type check both widened to match either type
# (`COLLECTION_VOUCHER_TYPE_NAMES = ("receipt", "cheque")`).
COLLECTION_VOUCHER_TYPE_NAMES = ("receipt", "cheque")


def _build_group_resolver():
    """
    Fetches the ledger->group and group->parent maps once and returns
    (ledger_parent, resolve) where resolve(immediate_group) walks up the
    chain to a KNOWN_DSR_GROUP_NAMES entry, same logic already proven in
    compute_eligible_roster()/dump_eligible()/dump_all_groups() above.
    Deliberately duplicated here rather than shared with those, so this
    new, not-yet-verified Collection diagnostic can never risk changing
    behaviour for the already-trusted roster/outstanding code paths.
    """
    ledger_xml = tally_request(LEDGER_GROUP_COLLECTION_XML)
    ledger_root = ET.fromstring(ledger_xml)
    ledger_parent = {}
    for led in ledger_root.iter("LEDGER"):
        name = led.get("NAME") or _text(led, "NAME")
        if name:
            ledger_parent[name.strip()] = (_text(led, "PARENT") or "").strip()

    group_xml = tally_request(GROUP_HIERARCHY_COLLECTION_XML)
    group_root = ET.fromstring(group_xml)
    group_parent = {}
    for grp in group_root.iter("GROUP"):
        gname = grp.get("NAME") or _text(grp, "NAME")
        if gname:
            group_parent[gname.strip()] = (_text(grp, "PARENT") or "").strip()

    known_set = set(KNOWN_DSR_GROUP_NAMES)

    def resolve(immediate_group):
        current = immediate_group
        seen = set()
        while current and current not in seen:
            seen.add(current)
            if current in known_set:
                return current
            current = group_parent.get(current, "")
        return None

    return ledger_parent, resolve


def _resolve_dsr_group_name(fragment):
    """Matches a DSR name/fragment (case-insensitive, either the exact
    KNOWN_DSR_GROUP_NAMES string or a substring of it, e.g. "arun") to the
    one real group name it should use. Returns None if nothing matched."""
    frag_upper = fragment.upper()
    for g in KNOWN_DSR_GROUP_NAMES:
        if frag_upper == g.upper():
            return g
    for g in KNOWN_DSR_GROUP_NAMES:
        if frag_upper in g.upper():
            return g
    return None
def _prev_month_range(now):
    prev_month_end = now.replace(day=1) - timedelta(days=1)
    return prev_month_end.strftime("%Y%m01"), prev_month_end.strftime("%Y%m%d")


# Added 21 Sep 2026 for the DSR apps' "Last month" chip going Tally-real
# (owner's request, after confirming the app-punched version was showing
# 0 for August - the apps only started being used in September, so there
# was never any real app-logged August data to show). A fair mid-month
# pace comparison needs the SAME number of days on both sides - e.g. "Sep
# 1-21" needs "Aug 1-21", not all of August - so this returns the
# previous month's start plus a same-day-of-month upper bound, clamped to
# that month's own last day (so e.g. today=31 in a 30-day month still
# returns a valid date instead of overflowing into the month after).
def _prev_month_samedate_range(now):
    prev_month_end = now.replace(day=1) - timedelta(days=1)
    prev_month_start = prev_month_end.replace(day=1)
    day = min(now.day, prev_month_end.day)
    prev_samedate = prev_month_start.replace(day=day)
    return prev_month_start.strftime("%Y%m01"), prev_samedate.strftime("%Y%m%d")


def _sixmo_window_range(now):
    """
    Added 23 Sep 2026 for the "this month vs 6-month average" dealer
    review feature (owner: "this month sales vs 6 month avg sales dealer
    wise and collections also same that"). Returns (window_start,
    window_end) as "YYYYMMDD" strings covering the 6 FULL calendar
    months immediately before the current one - e.g. on 23 Sep 2026,
    Mar 1 through Aug 31 2026. The current in-progress month is
    deliberately excluded from the average (it's shown separately as
    "this month"), same convention as every other month-to-date figure
    in this file.

    Handles the year-wrap case (Jan/Feb/etc. looking back into the
    previous year) via plain integer month arithmetic rather than
    relying on a library that may not be available - relativedelta isn't
    imported anywhere else in this file, so this stays dependency-free.
    """
    y, m = now.year, now.month
    # previous calendar month (the window's END month)
    end_total = y * 12 + (m - 1) - 1
    end_y, end_m = end_total // 12, end_total % 12 + 1
    # 6 months back from the CURRENT month (the window's START month)
    start_total = y * 12 + (m - 1) - 6
    start_y, start_m = start_total // 12, start_total % 12 + 1
    window_start = "{:04d}{:02d}01".format(start_y, start_m)
    next_total = end_y * 12 + (end_m - 1) + 1
    next_y, next_m = next_total // 12, next_total % 12 + 1
    last_day = (datetime(next_y, next_m, 1) - timedelta(days=1)).day
    window_end = "{:04d}{:02d}{:02d}".format(end_y, end_m, last_day)
    return window_start, window_end


def fetch_collection():
    """
    Returns (dsr_totals, dealer_totals):
      dsr_totals = {dsr_group_name: {"today": rs, "month": rs}} for every
      one of the 8 KNOWN_DSR_GROUP_NAMES - the sum of POSITIVE amounts on
      each DSR's own dealers' ledger entries within "Receipt" vouchers.
      dealer_totals = {dealer_ledger_name: {"today": rs, "month": rs}},
      added 17 Sep 2026 for the dashboard's "Top customers - Collection"
      panel (see claude/sales-dashboard.md) - the SAME per-entry amounts
      already being summed into dsr_totals, just also kept keyed by the
      individual dealer ledger name rather than only rolled up to its
      DSR. This needed no new Tally request or field: each
      ALLLEDGERENTRIES.LIST entry's own LEDGERNAME *is* the dealer's
      ledger name (that's exactly what ledger_to_dsr.get(lname) already
      resolves to a DSR below), so a dealer-level breakdown was already
      sitting right there in the same loop - a genuinely small addition,
      not a new build.
    "month" is the CURRENT CALENDAR MONTH so far (month start through
    today); "today" is just today's own vouchers, a subset of "month" -
    both computed in one pass over the same fetched vouchers (added 14 Sep
    2026, for the combined Sales & Collection WhatsApp report's "Today:
    ... Collection = ..." line, which needs a same-day figure, not just
    the month-to-date total the Collection panel itself already showed).

    CONFIRMED against real Tally data for Arun, 12 Sep 2026, via
    --check-collection before this function was written: "Receipt" is the
    exact voucher-type label on this Tally setup (not e.g. "Receipt
    Voucher" or something custom), and a dealer's own ledger entry on a
    Receipt voucher comes back POSITIVE - a real sample voucher showed
    +10,000.00 on the dealer's own ledger entry and -10,000.00 on the bank
    ledger entry in the same voucher - matching the negative=Dr/
    positive=Cr convention already used for CLOSINGBALANCE everywhere else
    in this file. No <ISDEBIT> flag comes back on these entries at all on
    this Tally setup, so sign (not a debit/credit tag) is the only signal,
    and it was confirmed reliable across a real 12-day sample (39 vouchers,
    4 distinct voucher types, Receipt's 22 entries were ALL positive with
    zero negative - a clean signal, not a mixed one).

    ASSUMED PERIOD: calendar month-to-date, matching the existing liters
    Target-Balance panel's "This month" convention. This has NOT been
    explicitly confirmed with the owner as the period he wants for
    Collection specifically - worth double-checking if a DSR's displayed
    Collection ever looks off by roughly a month's worth.

    Fetches the whole company's vouchers ONCE for the month - Tally's
    SVFROMDATE/SVTODATE do NOT scope a raw TYPE="Voucher" Collection
    export (confirmed 12 Sep 2026: a request for 12 days came back with a
    voucher dated over a year earlier), so each voucher's own <DATE> is
    filtered client-side here instead, same fix already proven in
    check_collection() above. Every ledger entry is then bucketed by
    resolving its ledger to a DSR via the same Group-chain walk
    compute_eligible_roster() uses (built fresh here via
    _build_group_resolver(), deliberately not shared/cached across calls,
    so this stays simple and correct even though it re-fetches the ledger/
    group hierarchy each time it's called).
    """
    now = datetime.now()
    from_d = now.strftime("%Y%m01")
    to_d = now.strftime("%Y%m%d")
    today_str = now.strftime("%Y%m%d")
    prev_from_d, prev_to_d = _prev_month_range(now)
    # Added 20 Sep 2026 for the dashboard's dealer-lookup card ("Last
    # payment", alongside the existing "Last order"/lastSaleDate from
    # fetch_sales()) - same 120-day lookback as fetch_sales() uses for its
    # own last-sale-date tracking, a safety margin comfortably wider than
    # the 100-day Inactive-dealer threshold so a dealer sitting right at
    # the edge is never missed by a few days' slack in date math.
    LAST_PAYMENT_LOOKBACK_DAYS = 120
    lookback_d = (now - timedelta(days=LAST_PAYMENT_LOOKBACK_DAYS)).strftime("%Y%m%d")
    # Added 21 Sep 2026 for the DSR apps' Collection panel's own "Last
    # month" chip going same-date (owner: "do same for collection also") -
    # see _prev_month_samedate_range()'s own comment. Already comfortably
    # inside the 120-day lookback above, no widening needed.
    prev_samedate_from_d, prev_samedate_to_d = _prev_month_samedate_range(now)
    # Added 23 Sep 2026 for the "this month vs 6-month avg" dealer review
    # feature - see _sixmo_window_range()'s own docstring. widest_lookback_d
    # replaces lookback_d as the loop's own skip cutoff below so vouchers
    # in the 6-month window (which starts well before the 120-day
    # last-payment lookback most months) aren't skipped before they ever
    # reach the per-dealer accumulation.
    sixmo_from_d, sixmo_to_d = _sixmo_window_range(now)
    widest_lookback_d = min(lookback_d, sixmo_from_d)

    ledger_parent, resolve = _build_group_resolver()
    ledger_to_dsr = {}
    for lname, lparent in ledger_parent.items():
        if not lparent:
            continue
        dsr = resolve(lparent)
        if dsr:
            ledger_to_dsr[lname] = dsr

    xml_body = COLLECTION_VOUCHERS_TEMPLATE.format(from_date=from_d, to_date=to_d)
    xml_text = tally_request(xml_body, timeout=600)
    root = ET.fromstring(xml_text)

    totals = dict((dsr, {"today": 0.0, "month": 0.0, "lastMonthSameDate": 0.0, "lastMonthFull": 0.0}) for dsr in KNOWN_DSR_GROUP_NAMES)
    dealer_totals = {}
    dealer_sixmo_totals = {}  # dealer -> sum of Rs across the 6-month window (not yet averaged)
    # DSR-level twin of dealer_sixmo_totals above, added 26 Sep 2026 for
    # the "Full Petronas This Month Collection" area table's new "6M AVG"
    # column - see fetch_sales()'s dsr_sixmo_totals comment for the same
    # pattern applied there.
    dsr_sixmo_totals = {}
    daily_totals = {}  # "YYYY-MM-DD" -> Rs, this month + last month only - see _prev_month_range()
    # Per-DSR daily trend, added 11 Oct 2026 for the dashboard's new Sales/
    # Collection weekday heatmap ("the sales heatmap above Ranking ...
    # ONLY SALES & COLLECTION ... ONLY DSR BUT LIVE FROM TALLY"). Same
    # dict-of-dicts shape (dsr -> {"YYYY-MM-DD": Rs}) as daily_totals above,
    # just one level deeper - no new Tally request or field: `dsr` and
    # `iso_date` are both already computed in the loop below for
    # dsr_totals/daily_totals respectively, this just also keys the same
    # amount by both at once. Merged Saravanan-into-Chellamani below via
    # _merge_saravanan_into_chellamani_nested_dict(), same as every other
    # per-DSR structure here.
    daily_totals_by_dsr = {}
    last_payment_date_by_dealer = {}
    # dsr -> set of dealer ledger names who paid THIS CALENDAR MONTH -
    # added 10 Oct 2026, the real fix for "Collection MB its seems to
    # wrong in all app". Kept SEPARATE from dealer_totals on purpose: the
    # 25 Sep 2026 widening below (see the dealer_rec comment just under
    # this) made dealer_totals.setdefault() run for ANY voucher anywhere
    # in widest_lookback_d..to_d (up to the 120-day lookback or the
    # 6-month average window, whichever is wider) - not just this month -
    # so dealer_totals.keys() has included every dealer who paid any time
    # in the last several months since that change, not just this month.
    # mb_by_dsr below used to count straight off dealer_totals.keys()
    # (added 21 Sep 2026, BEFORE the 25 Sep widening existed) on the
    # mistaken assumption - true when it was written, no longer true
    # after 25 Sep - that dealer_totals only ever held this-month dealers
    # at this point. The result: "MB" (meant to be "dealers who paid THIS
    # month") was actually counting dealers who paid at any point in
    # roughly the last 120 days to 6 months, on both the dashboard and
    # every DSR app (same shared fetch_collection() feeds both). This new
    # set is populated only when is_current_month is true, immediately
    # below, so MB goes back to meaning exactly what its own label says.
    mb_dealers_by_dsr = {}
    voucher_count = 0
    for v in root.iter("VOUCHER"):
        voucher_count += 1
        vdate = (_text(v, "DATE") or "").strip()
        # Widened 20 Sep 2026 twice: first (was: skip anything outside the
        # current month) so last month's vouchers reach daily_totals below,
        # then again to lookback_d (120 days) so last_payment_date_by_dealer
        # below can see a payment from further back than last calendar
        # month too - is_current_month (added below) still gates totals/
        # dealer_totals exactly as before, so the existing month-to-date
        # figures are completely unaffected by either widening. Widened a
        # third time, 23 Sep 2026, to widest_lookback_d (whichever of the
        # 120-day lookback or the 6-month average window starts earlier)
        # for the same reason - is_sixmo_window (below) is its own gate,
        # unaffected by how wide the loop's outer skip is.
        if vdate and (vdate < widest_lookback_d or vdate > to_d):
            continue
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip()
        if vtype.lower() not in COLLECTION_VOUCHER_TYPE_NAMES:
            continue
        is_current_month = from_d <= vdate <= to_d
        is_prev_month = prev_from_d <= vdate <= prev_to_d
        is_prev_month_samedate = prev_samedate_from_d <= vdate <= prev_samedate_to_d
        is_sixmo_window = sixmo_from_d <= vdate <= sixmo_to_d
        is_today = vdate == today_str
        for entry in v.iter("ALLLEDGERENTRIES.LIST"):
            lname = _text(entry, "LEDGERNAME")
            if not lname:
                continue
            lname = lname.strip()
            dsr = ledger_to_dsr.get(lname)
            if not dsr:
                continue
            amt = _num(_text(entry, "AMOUNT"))
            if amt is None or amt <= 0:
                continue
            # Per-dealer record, moved here (before the
            # is_prev_month_samedate/is_prev_month checks below) and given
            # its lastMonthSameDate/lastMonthFull keys, 25 Sep 2026, for
            # the Dealer Review popup's new "6-mo avg / last month /
            # current month" Collection figures - mirrors totals[dsr]'s
            # own lastMonthSameDate/lastMonthFull just below, at dealer
            # granularity. Was previously only set up under
            # "if not is_current_month: continue" further down, so a
            # dealer with only a prior-month payment (no line this month)
            # never got a dealer_totals record at all until the
            # lastPaymentDate merge after the loop - now it does, with
            # both new keys already defaulted so the += below never
            # KeyErrors on any code path.
            dealer_rec = dealer_totals.setdefault(lname, {"today": 0.0, "month": 0.0, "lastMonthSameDate": 0.0, "lastMonthFull": 0.0})
            # See mb_dealers_by_dsr's own comment above - this is the only
            # correct place to decide "did this dealer pay THIS month",
            # since is_current_month is already computed right here and
            # amt > 0/dsr-resolved were already checked just above.
            if is_current_month:
                mb_dealers_by_dsr.setdefault(dsr, set()).add(lname)
            if is_prev_month_samedate:
                totals[dsr]["lastMonthSameDate"] += amt
                dealer_rec["lastMonthSameDate"] += amt
            # "Last Month Full" (the WHOLE previous calendar month's Rs),
            # added 21 Sep 2026 - same reasoning/pattern as fetch_sales()'s
            # own lastMonthFull, for the Collection panel's "Last month"
            # chip bracket figure ("show in bracket the last month total
            # amount"). is_prev_month was already computed above for
            # daily_totals' trend chart - this just also feeds it into a
            # per-DSR total.
            if is_prev_month:
                totals[dsr]["lastMonthFull"] += amt
                dealer_rec["lastMonthFull"] += amt
            # 6-month average window - added 23 Sep 2026, see
            # _sixmo_window_range()'s docstring. Summed here per dealer;
            # divided by 6 once, after the loop, to get the average.
            if is_sixmo_window:
                dealer_sixmo_totals[lname] = dealer_sixmo_totals.get(lname, 0.0) + amt
                # DSR-level accumulation, added 26 Sep 2026 - see
                # dsr_sixmo_totals' own comment above.
                dsr_sixmo_totals[dsr] = dsr_sixmo_totals.get(dsr, 0.0) + amt
            # Added 20 Sep 2026 for the dealer-lookup card's "Last payment" -
            # tracked for every dealer within the 120-day lookback,
            # independent of is_current_month, same "widen the loop, add a
            # side-tracker" pattern fetch_sales() already uses for
            # lastSaleDate.
            prev_payment = last_payment_date_by_dealer.get(lname)
            if not prev_payment or vdate > prev_payment:
                last_payment_date_by_dealer[lname] = vdate
            if is_current_month or is_prev_month:
                iso_date = vdate[0:4] + "-" + vdate[4:6] + "-" + vdate[6:8]
                daily_totals[iso_date] = daily_totals.get(iso_date, 0.0) + amt
                _dsr_day = daily_totals_by_dsr.setdefault(dsr, {})
                _dsr_day[iso_date] = _dsr_day.get(iso_date, 0.0) + amt
            if not is_current_month:
                continue
            totals[dsr]["month"] += amt
            dealer_rec["month"] += amt
            if is_today:
                totals[dsr]["today"] += amt
                dealer_rec["today"] += amt
    # Printed so the effect of the 14 Sep 2026 Tally-side FILTER (see the
    # module comment above COLLECTION_VOUCHERS_TEMPLATE) is visible from
    # the watcher's own console output on its very next real run, without
    # needing a separate diagnostic round - before this fix, this number
    # was regularly in the tens of thousands (company-wide, every voucher
    # type); if the filter is working, it should now be close to just the
    # DSRs' own Receipt/Cheque-voucher count for the period.
    print("  (collection: Tally returned {} voucher(s) for this fetch)".format(voucher_count))

    # "Last payment (amount)" - added 26 Sep 2026 for the dealer-lookup
    # card (owner: "IN LAST PAYMENT (AMOUNT)"). last_payment_date_by_dealer
    # above only records WHEN a dealer's most recent Receipt was, not how
    # much it was for - this second pass (over the SAME already-fetched
    # `root`, no new Tally request) sums the amount of every ledger entry
    # on every voucher dated exactly that dealer's own
    # last_payment_date_by_dealer date (more than one receipt on the same
    # calendar day folds into one "last payment" figure, same reasoning as
    # fetch_sales()'s own last_sale_ltr_by_dealer just above). Two passes
    # for the same reason as that one: Tally doesn't guarantee voucher
    # order, so the true max date per dealer isn't final until the main
    # loop above has seen every voucher.
    last_payment_amt_by_dealer = {}
    for v in root.iter("VOUCHER"):
        vdate = (_text(v, "DATE") or "").strip()
        if not vdate:
            continue
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip()
        if vtype.lower() not in COLLECTION_VOUCHER_TYPE_NAMES:
            continue
        for entry in v.iter("ALLLEDGERENTRIES.LIST"):
            lname = _text(entry, "LEDGERNAME")
            if not lname:
                continue
            lname = lname.strip()
            if not ledger_to_dsr.get(lname):
                continue
            amt = _num(_text(entry, "AMOUNT"))
            if amt is None or amt <= 0:
                continue
            target_date = last_payment_date_by_dealer.get(lname)
            if not target_date or vdate != target_date:
                continue
            last_payment_amt_by_dealer[lname] = last_payment_amt_by_dealer.get(lname, 0.0) + amt

    # avg6moCollectedRs at DSR level - added 26 Sep 2026, see
    # dsr_sixmo_totals' own comment above. Folded in BEFORE
    # _merge_saravanan_into_chellamani() below for the same reason as
    # fetch_sales()'s own avg6moLtr/avg6moRs fold-in (safe to sum two
    # already-divided averages here since both share the same divisor, 6).
    for _dsr, _sixmo_total in dsr_sixmo_totals.items():
        totals[_dsr]["avg6moCollectedRs"] = round(_sixmo_total / 6.0, 2)

    _merge_saravanan_into_chellamani(totals)

    # "MB" (dealers who paid this month), added 21 Sep 2026 for the DSR
    # apps' Collection panel. FIXED 10 Oct 2026 ("check collection MB its
    # seems to wrong in all app"): this used to count from
    # dealer_totals.keys() directly, on the assumption (true on 21 Sep
    # 2026, no longer true after the 25 Sep 2026 widening above) that
    # dealer_totals only held dealers with a CURRENT-MONTH voucher at this
    # point - it had actually been holding every dealer with a voucher
    # anywhere in the last 120 days-to-6-months for over two weeks, so MB
    # was silently counting far more dealers than "paid this month" ever
    # should. Now counts from mb_dealers_by_dsr (see its own comment
    # above), which is built with an explicit is_current_month check in
    # the same loop that computes everything else here - MB is correct
    # again, and every other field in `totals` is completely unaffected
    # (dealer_totals itself is untouched, still needed exactly as before
    # for lastMonthSameDate/lastMonthFull/the Dealer Review popup).
    mb_by_dsr = {}
    for dsr, dealer_set in mb_dealers_by_dsr.items():
        mb_by_dsr[dsr] = len(dealer_set)
    _merge_saravanan_count_into_chellamani(mb_by_dsr)
    for dsr in totals:
        totals[dsr]["mb"] = mb_by_dsr.get(dsr, 0)

    # Fold lastPaymentDate into dealer_totals for every dealer that has one,
    # even a dealer with nothing collected THIS month (dealer_totals
    # otherwise only ever gets an entry for a dealer with a current-month
    # receipt) - same merge fetch_sales() does for lastSaleDate, so a
    # dealer who paid last month but not yet this month still shows their
    # real last-payment date instead of nothing.
    # Full default (matching dealer_rec's own setdefault inside the loop
    # above, updated 25 Sep 2026 for lastMonthSameDate/lastMonthFull) kept
    # here too, defensively - in practice dealer_rec above now setdefaults
    # unconditionally for every entry with amt > 0 regardless of date
    # window, so these later setdefaults should always find an existing
    # record already.
    _DEALER_REC_DEFAULT = lambda: {"today": 0.0, "month": 0.0, "lastMonthSameDate": 0.0, "lastMonthFull": 0.0}
    all_dealer_names = set(dealer_totals.keys()) | set(last_payment_date_by_dealer.keys())
    for name in all_dealer_names:
        rec = dealer_totals.setdefault(name, _DEALER_REC_DEFAULT())
        rec["lastPaymentDate"] = last_payment_date_by_dealer.get(name)
        # "Last payment (amount)" - see last_payment_amt_by_dealer's own
        # comment above. None (not 0.0) for a dealer with no
        # lastPaymentDate at all, same "None means genuinely absent"
        # convention as avg6moCollectedRs below.
        rec["lastPaymentAmount"] = last_payment_amt_by_dealer.get(name)

    # avg6moCollectedRs - added 23 Sep 2026, see _sixmo_window_range()'s
    # docstring. Only set for a dealer with at least one voucher somewhere
    # in the 6-month window (None, not 0.0, for a dealer with zero history
    # there at all - same "None means genuinely absent" convention used
    # for the age bands elsewhere in this file, so a dealer who simply
    # never had activity in that window isn't misread as "averaging Rs 0").
    for name, sixmo_total in dealer_sixmo_totals.items():
        rec = dealer_totals.setdefault(name, _DEALER_REC_DEFAULT())
        rec["avg6moCollectedRs"] = round(sixmo_total / 6.0, 2)

    # daily_totals_by_dsr is dsr -> {date: Rs} - flat, but keyed by date
    # (not today/month), so the DEDICATED flat merge is used here, not
    # _merge_saravanan_into_chellamani() (which seeds a fresh Chellamani
    # entry with {"today":0.0,"month":0.0} - wrong defaults for a date-
    # keyed dict) and not _merge_saravanan_into_chellamani_nested_dict()
    # (expects a 3rd level down, dsr -> month -> {field: number} - wrong
    # shape entirely, throws on a plain float). See
    # _merge_saravanan_into_chellamani_flat()'s own comment.
    _merge_saravanan_into_chellamani_flat(daily_totals_by_dsr)

    return totals, dealer_totals, daily_totals, daily_totals_by_dsr


def diff_collection(prev, current):
    changed = []
    for dsr, rec in current.items():
        old = prev.get(dsr)
        if old != rec:
            # "mb" added 21 Sep 2026 - see fetch_collection()'s own MB comment.
            # "lastMonthSameDate" added 21 Sep 2026 - see fetch_collection()'s
            # own comment near prev_samedate_from_d/prev_samedate_to_d.
            # "lastMonthFull" added the same day (right after) - see
            # fetch_collection()'s own comment above.
            changed.append({
                "dsrName": dsr,
                "todayCollected": rec["today"],
                "monthCollected": rec["month"],
                "mb": rec.get("mb"),
                "lastMonthSameDate": rec.get("lastMonthSameDate"),
                "lastMonthFull": rec.get("lastMonthFull"),
                # "avg6moCollectedRs" added 26 Sep 2026 for the "Full
                # Petronas This Month Collection" area table's new
                # "6M AVG" column - see fetch_collection()'s
                # dsr_sixmo_totals comment.
                "avg6moCollectedRs": rec.get("avg6moCollectedRs"),
            })
    return changed


# Added 17 Sep 2026 alongside dealer_totals in fetch_collection() - same
# shape as diff_collection() above, just keyed by dealer name instead of
# DSR name and posted to a separate sheet/action (see
# handleLiveCollectionByDealerUpdate_() in Code.gs) so the existing
# per-DSR LiveCollection sheet/POST is completely untouched.
def diff_collection_by_dealer(prev, current):
    changed = []
    for dealer, rec in current.items():
        old = prev.get(dealer)
        if old != rec:
            changed.append({
                "name": dealer,
                "todayCollected": rec["today"],
                "monthCollected": rec["month"],
                # Added 20 Sep 2026 for the dealer-lookup card's "Last
                # payment" - see fetch_collection()'s docstring/lastPaymentDate
                # fold-in above.
                "lastPaymentDate": rec.get("lastPaymentDate"),
                # Added 26 Sep 2026 for the dealer-lookup card's "Last
                # payment (amount)" - see fetch_collection()'s
                # last_payment_amt_by_dealer comment.
                "lastPaymentAmount": rec.get("lastPaymentAmount"),
                # Added 23 Sep 2026 for the "this month vs 6-month avg"
                # dealer review feature - see fetch_collection()'s
                # avg6moCollectedRs comment.
                "avg6moCollectedRs": rec.get("avg6moCollectedRs"),
                # Added 25 Sep 2026 for the Dealer Review popup's new
                # "last month" Collection figure - Rs, no suffix needed
                # (Collection is Rs-only), matching the DSR-level
                # totals[dsr] convention in diff_collection() above which
                # also uses bare "lastMonthSameDate"/"lastMonthFull".
                "lastMonthSameDate": rec.get("lastMonthSameDate"),
                "lastMonthFull": rec.get("lastMonthFull"),
            })
    return changed


# Collection runs on its own, lighter cadence (30 min) rather than every
# 10-min main poll, since fetch_collection() pulls the WHOLE company's
# vouchers for the month each time (heavier than the small incremental
# diffs stock/outstanding send) - same reasoning that gave the new-dealer
# roster its own slower cadence. Simpler gate than the roster's, though:
# a plain elapsed-seconds check (not a fixed daily time) is safe here
# because the OUTER poll loop is already 10 minutes apart, so a failed
# attempt naturally waits for the next main poll before retrying - no
# retry-storm risk like the roster's original ~4h-interval bug (see that
# bug's writeup above for why THAT one needed a stricter daily-time gate
# instead of a plain interval).
# Widened 30 -> 60 min, 14 Sep 2026, after the owner reported Tally running
# slow - a zero-risk extra safety margin stacked on top of the FILTER fix
# above (see the module comment near COLLECTION_VOUCHERS_TEMPLATE);
# Collection doesn't need anywhere near stock/outstanding's freshness.
COLLECTION_POLL_INTERVAL_SECONDS = 60 * 60


def maybe_push_collection(state, dry_run, ts):
    if dry_run:
        return state
    last = state.get("last_collection_push_ts", 0)
    if time.time() - last < COLLECTION_POLL_INTERVAL_SECONDS:
        return state
    print("[{}] computing DSR collection (Receipt vouchers, this month so far - "
          "runs every {} min, not every poll)...".format(ts, COLLECTION_POLL_INTERVAL_SECONDS // 60))
    try:
        collection, collection_by_dealer, daily_collection, daily_collection_by_dsr = fetch_collection()
    except Exception as e:
        print("[{}] collection computation failed this round (will retry at the next "
              "poll): {}".format(ts, e))
        return state
    changed = diff_collection(state.get("collection", {}), collection)
    if changed:
        sent = post_backend_chunked("liveCollectionUpdate", "collection", changed, ts)
        print("[{}] collection: sent {} changed DSR row(s) total".format(ts, sent))
    else:
        print("[{}] collection: no changes".format(ts))
    # Per-dealer breakdown, added 17 Sep 2026 for the dashboard's "Top
    # customers - Collection" panel - own diff/push, own sheet/action, so
    # the existing per-DSR push above is completely unaffected either way.
    dealer_changed = diff_collection_by_dealer(state.get("collectionByDealer", {}), collection_by_dealer)
    if dealer_changed:
        dealer_sent = post_backend_chunked("liveCollectionByDealerUpdate", "collection", dealer_changed, ts)
        print("[{}] collection (by dealer): sent {} changed dealer row(s) total".format(ts, dealer_sent))
    else:
        print("[{}] collection (by dealer): no changes".format(ts))

    # Daily trend, added 20 Sep 2026 - same pattern as sales' own trend
    # push in maybe_push_sales() above, see diff_daily_trend()'s comment.
    trend_changed = diff_daily_trend(state.get("collectionTrendDaily", {}), daily_collection, "rs")
    if trend_changed:
        trend_sent = post_backend_chunked("liveCollectionTrendUpdate", "trend", trend_changed, ts)
        print("[{}] collection (daily trend): sent {} changed day(s) total".format(ts, trend_sent))
    else:
        print("[{}] collection (daily trend): no changes".format(ts))

    # Per-DSR daily trend, added 11 Oct 2026 for the new Sales/Collection
    # heatmap - own diff/push/state key, same pattern as the company-wide
    # trend just above.
    dsr_trend_changed = diff_daily_trend_by_dsr(state.get("collectionTrendDailyByDsr", {}), daily_collection_by_dsr, "rs")
    if dsr_trend_changed:
        dsr_trend_sent = post_backend_chunked("liveCollectionTrendByDsrUpdate", "trend", dsr_trend_changed, ts)
        print("[{}] collection (daily trend by DSR): sent {} changed row(s) total".format(ts, dsr_trend_sent))
    else:
        print("[{}] collection (daily trend by DSR): no changes".format(ts))

    state["collection"] = collection
    state["collectionByDealer"] = collection_by_dealer
    state["collectionTrendDaily"] = daily_collection
    state["collectionTrendDailyByDsr"] = daily_collection_by_dsr
    state["last_collection_push_ts"] = time.time()
    return state


def check_collection(dsr_fragment, from_date=None, to_date=None):
    """
    --check-collection "arun"   (or the exact group name, e.g.
    "PETRONAS DINDIGUL AREA (ARUN)" - matched case-insensitively, either
    exact or as a substring against KNOWN_DSR_GROUP_NAMES)

    Diagnostic only - pushes nothing anywhere. Fetches this calendar
    month's vouchers by default (from_date/to_date can override, both
    YYYYMMDD, e.g. to check a different period), resolves each ledger
    entry to a DSR via its Tally group chain (same walk as
    compute_eligible_roster()), and for the requested DSR prints every
    distinct voucher type seen among its dealers' entries this period with
    its own credit/debit subtotal and entry count - so the real Receipt
    voucher type name/amount/sign on this Tally setup can be read directly
    off Tally's own output, rather than guessed, before fetch_collection()
    (not yet written) gets wired to sum exactly the right thing.
    """
    target = _resolve_dsr_group_name(dsr_fragment)
    if not target:
        print("'{}' didn't match any of the known DSR groups:".format(dsr_fragment))
        for g in KNOWN_DSR_GROUP_NAMES:
            print("  " + g)
        return

    now = datetime.now()
    from_d = from_date or now.strftime("%Y%m01")
    to_d = to_date or now.strftime("%Y%m%d")

    print("Resolving Tally's ledger/group hierarchy...")
    ledger_parent, resolve = _build_group_resolver()

    dealer_names_for_dsr = set(
        lname for lname, lparent in ledger_parent.items()
        if lparent and resolve(lparent) == target
    )
    print("  {} ledger(s) resolve to {}.".format(len(dealer_names_for_dsr), target))
    if not dealer_names_for_dsr:
        print("  No dealers resolved to this group at all - stopping here; "
              "something upstream of the voucher fetch would need a look first.")
        return

    # Timeout raised 120s -> 600s (12 Sep 2026): even a short 12-day range
    # timed out at 120s on the office's Tally - this request pulls EVERY
    # voucher in the WHOLE company for the period (not pre-filtered to one
    # DSR's dealers, since Tally's Collection XML export has no simple way
    # to filter Vouchers by ledger-group membership), so it can be slow
    # regardless of how short the date window is if the company has a lot
    # of overall transaction volume. Same fix already proven for
    # fetch_active_ledger_names()'s much larger 1.5-year pull (raised to
    # 600s on 9 Sep 2026) - this is a diagnostic run only, so a slower
    # response is a fine tradeoff for not timing out.
    print("Fetching vouchers from {} to {} - this can take a few minutes on a busy "
          "company, please wait...".format(from_d, to_d))
    xml_body = COLLECTION_VOUCHERS_TEMPLATE.format(from_date=from_d, to_date=to_d)
    xml_text = tally_request(xml_body, timeout=600)
    root = ET.fromstring(xml_text)

    # Bucket key -> {"positive": sum of positive AMOUNTs, "negative": sum of
    # negative AMOUNTs (kept negative, not abs'd), "count": n}. Split by
    # actual sign rather than an <ISDEBIT> flag - the raw sample from the
    # first real run showed entries with NO <ISDEBIT> tag at all on this
    # Tally setup (only <AMOUNT>), so sign is the only signal available.
    # Positive/negative are kept as their own numbers (not abs'd together)
    # so the two directions of money movement stay visible and distinct.
    by_voucher_type = {}
    sample_voucher_xml = None       # first voucher touching this DSR, any type
    sample_receipt_voucher_xml = None  # first voucher whose type contains "receipt"
    voucher_count_touched = 0
    total_vouchers_in_response = 0
    out_of_range_count = 0

    # Date filtering is done HERE, client-side, on each voucher's own <DATE>
    # (YYYYMMDD text - compares correctly as a plain string, same convention
    # used for dateKey everywhere else in this project) rather than trusting
    # SVFROMDATE/SVTODATE to have scoped the request server-side. Added
    # 12 Sep 2026 after the first real run of this diagnostic came back with
    # a sample voucher dated 20250401 despite requesting only 20260901 to
    # 20260912 - proving Tally's STATICVARIABLES period does NOT constrain a
    # raw TYPE="Voucher" Collection export the way it does for the
    # report-style requests elsewhere in this file (Bills Receivable, the
    # native report export, respects it; this raw Collection export
    # apparently doesn't). This makes the result correct regardless of what
    # Tally actually sends back.
    for v in root.iter("VOUCHER"):
        total_vouchers_in_response += 1
        vdate = (_text(v, "DATE") or "").strip()
        if vdate and (vdate < from_d or vdate > to_d):
            out_of_range_count += 1
            continue
        vtype = (_text(v, "VOUCHERTYPENAME") or "(blank)").strip()
        touched_this_voucher = False
        for entry in v.iter("ALLLEDGERENTRIES.LIST"):
            lname = _text(entry, "LEDGERNAME")
            if not lname or lname.strip() not in dealer_names_for_dsr:
                continue
            amt = _num(_text(entry, "AMOUNT"))
            if amt is None:
                continue
            touched_this_voucher = True
            bucket = by_voucher_type.setdefault(
                vtype, {"positive": 0.0, "negative": 0.0, "count": 0})
            bucket["count"] += 1
            if amt >= 0:
                bucket["positive"] += amt
            else:
                bucket["negative"] += amt  # stays negative
        if touched_this_voucher:
            voucher_count_touched += 1
            if sample_voucher_xml is None:
                sample_voucher_xml = ET.tostring(v, encoding="unicode")
            if sample_receipt_voucher_xml is None and "receipt" in vtype.lower():
                sample_receipt_voucher_xml = ET.tostring(v, encoding="unicode")

    print("\n" + "=" * 70)
    print("{}  -  vouchers touching its {} dealer(s), {} to {}".format(
        target, len(dealer_names_for_dsr), from_d, to_d))
    print("=" * 70)
    print("Tally returned {} voucher(s) total in this response; {} were outside "
          "{}-{} and excluded here (client-side date filter - see the code comment "
          "if this number looks large).".format(
              total_vouchers_in_response, out_of_range_count, from_d, to_d))
    print("{} voucher(s) touched one of this DSR's dealers this period.\n".format(voucher_count_touched))
    if not by_voucher_type:
        print("No matching voucher entries found for this DSR in this date range.")
    for vtype in sorted(by_voucher_type.keys()):
        b = by_voucher_type[vtype]
        print("  Voucher type: {}".format(vtype))
        print("    entries: {}   positive-amount total: {}   negative-amount total: {}".format(
            b["count"], b["positive"], b["negative"]))

    print("\n" + "=" * 70)
    print("No <ISDEBIT> flag came back on any entry in this response, so the split above")
    print("is by the AMOUNT's own sign (+/-), not a debit/credit label. Everywhere else in")
    print("this file, a NEGATIVE amount on a dealer's own ledger = Dr = they owe more (e.g.")
    print("a Sales invoice), and POSITIVE = Cr = they owe less (e.g. a Receipt) - same")
    print("convention as CLOSINGBALANCE throughout this script. If that holds here too,")
    print("Collection for this DSR = the 'Receipt' row's POSITIVE total. Tell Claude if")
    print("that number looks right, or if it should be a different label/sign.")
    if sample_receipt_voucher_xml:
        print("\n--- One matching voucher whose type contains 'receipt', raw (check its AMOUNT sign) ---")
        print(sample_receipt_voucher_xml[:3000])
    elif sample_voucher_xml:
        print("\n--- One matching voucher, raw (for a sanity read - not a Receipt-labelled one) ---")
        print(sample_voucher_xml[:3000])
    print("\nSend this whole output back.")



# -- Dashboard-only live Sales (from Tally billing), diagnostic added
# 17 Sep 2026 ------------------------------------------------------------
# The owner's concern: the dashboard's Sales numbers only ever reflect
# what a DSR typed into the app as an order - a DSR who forgets to log
# one, or an order billed straight from the office without ever touching
# the app, would never show up. His ask: "like collection live can u pull
# sales also live only for the dashboard" - pull the real, actually-BILLED
# Ltr straight from Tally's own Sales vouchers, the same way Collection
# already pulls real Receipt vouchers, for the dashboard ONLY (never the
# DSR apps themselves - their own order-taking flow/target panel/WhatsApp
# report all stay exactly as they are, untouched).
#
# This mirrors fetch_collection()'s whole approach (a Sales-only FILTER
# formula pushed server-side from day one, learning from the "Tally
# running slow" mistake that shipped before Collection got the same
# treatment; client-side date filtering, since SVFROMDATE/SVTODATE do not
# scope a raw TYPE="Voucher" Collection export on this Tally setup -
# confirmed for Collection, assumed to hold here too until proven
# otherwise) - but with one extra unknown Collection never had: a Receipt
# voucher's amount is already in rupees, nothing to convert; a Sales
# voucher's inventory line reports a QUANTITY, and turning that into the
# same Ltr figure every DSR app already computes for its own orders
# needs two things resolved first, neither yet confirmed against real
# data:
#   1. Which field carries billed quantity - ACTUALQTY vs BILLEDQTY (they
#      can differ, e.g. free/scheme quantity) - and whether either is
#      even present the way expected on this Tally setup, the same kind
#      of "does this field even exist here" question STANDARDSELLINGPRICE
#      failed and StandardPrice answered for rate.
#   2. What UNIT that quantity is actually expressed in for each stock
#      item. This project already tracks each item's "Case size" (how
#      many individual bottles/cans make one order-unit, and how many
#      Ltr each individual bottle/can holds) in claude/item-catalog.json -
#      embedded below as ITEM_LTR_CATALOG - but Tally's own SALES-voucher
#      quantity could plausibly be expressed either in that same
#      "Case"/order-unit (matching how CLOSINGBALANCE already behaves for
#      live stock - see fetch_stock()) or in individual "Nos" (bottles),
#      and guessing wrong would silently scale every item's Ltr figure by
#      its own case size - a systematic, hard-to-notice error across the
#      WHOLE dashboard (Combined Sales, Ranking, Top 5 Items), not a
#      single-field bug like every previous wrong Tally guess in this
#      project. Tally's own raw quantity text conventionally carries its
#      unit suffix (e.g. "20.00 Case" vs "218.00 Nos") - printed verbatim,
#      unparsed, below specifically so this can be read directly off real
#      output rather than guessed, same discipline as every prior new
#      Tally field here (Outstanding's BILLTYPE, Rate's CLOSINGRATE vs
#      StandardPrice, MRP's nested list structure).
#
# fetch_sales() (the real, non-diagnostic version) does NOT exist yet -
# deliberately not written until --dump-sales's real output has been seen
# and both unknowns above are settled, same sequence Collection followed
# (check_collection() first, fetch_collection() only after "Receipt"/sign
# were confirmed against real vouchers).
#
# UPDATE 18 Sep 2026, same day - both unknowns settled against a real
# voucher's raw XML (owner's own --dump-sales output, a real Sales
# invoice from 1 Apr 2026 - the type filter itself is correct, the type
# just genuinely wasn't present on TODAY's own vouchers yet at the time
# this ran, same SVFROMDATE/SVTODATE-doesn't-scope quirk Collection
# already had):
#   1. ACTUALQTY/BILLEDQTY are both present and populated - confirmed.
#   2. The unit is individual NOS (pieces), NOT the "Case"/order-unit
#      CLOSINGBALANCE uses for live stock - confirmed from the raw XML
#      itself: `<ACTUALQTY TYPE="Quantity"> 5 NOS</ACTUALQTY>`.
#      _catalog_ltr_guess() fixed same day to multiply by sizeLtr alone
#      (no nosPerUnit) - the original guess would have overstated every
#      case-packed item's Ltr by its own case size.
# Still open before fetch_sales() gets written for real: the sample
# voucher seen so far billed a non-lubricant spare part (not in
# ITEM_LTR_CATALOG), so the fixed formula hasn't yet been checked against
# a REAL known lubricant item's Ltr-guess total for a day the owner can
# independently verify - waiting on a --dump-sales run against a wider/
# more recent date range (today's vouchers specifically hadn't posted yet
# when the diagnostic was last run) to confirm that end-to-end.

ITEM_LTR_CATALOG = json.loads('''{"10W30 1 LIT SPRINTA F700 (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"0W20 4 LIT SYNTIUM 7000 SN HYBRID FULLSYN (4)":{"sizeLtr":4,"nosPerUnit":4,"unit":"Case"},"0W30 MGDO 5.3 LIT SYNTIUM 3000SZ (2)":{"sizeLtr":5.3,"nosPerUnit":2,"unit":"Case"},"0W40  4 LIT SYNTIUM 7000 SN (COOLTECH) (5)":{"sizeLtr":4,"nosPerUnit":5,"unit":"Case"},"10W30 210 LIT SPRINTA F300 BARREL":{"sizeLtr":210,"nosPerUnit":null,"unit":"Barrel"},"10W30 800 ML SPRINTA A300 (COMBO PACK - WITH GEAR OIL) (10)":{"sizeLtr":0.8,"nosPerUnit":10,"unit":"Case"},"10W30 800 ML SPRINTA A300 (SPOUT PACK) (20)":{"sizeLtr":0.8,"nosPerUnit":20,"unit":"Case"},"10W30 900 ML SPRINTA F300 (20)":{"sizeLtr":0.9,"nosPerUnit":20,"unit":"Case"},"10W30 F100 SPRINTA 55 LIT":{"sizeLtr":55,"nosPerUnit":null,"unit":"Drum"},"10W40 1 LIT SN SPRINTA RACING (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"10W40 1 LIT SPRINTA F900 (MOTO GP PACK) (16)":{"sizeLtr":1,"nosPerUnit":16,"unit":"Case"},"10W40 1 LIT SYNTIUM 800 (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"10W40 3 LIT SYNTIUM 800 (6)":{"sizeLtr":3,"nosPerUnit":6,"unit":"Case"},"10W40 3.5 LIT SYNTIUM 800 SN (4)":{"sizeLtr":3.5,"nosPerUnit":4,"unit":"Case"},"10W50 1 LIT SN SPRINTA RACING (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"15W40 1 LIT SYNTIUM 500 CH-4 DIESEL *****":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"15W40 1 LIT SYNTIUM 500 CI4+ DIESEL (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"15W40 5 LIT SYNTIUM 500 CI4+ DIESEL":{"sizeLtr":5,"nosPerUnit":4,"unit":"Case"},"15W40 7 LIT SYNTIUM 500 CI4+ DIESEL":{"sizeLtr":7,"nosPerUnit":null,"unit":"Bucket"},"15W50 1 LIT SPRINTA F700 (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"15W50 2.5 LIT SPRINTA F700 (6)":{"sizeLtr":2.5,"nosPerUnit":6,"unit":"Case"},"15W50 SPRINTA F700 50 LIT":{"sizeLtr":50,"nosPerUnit":null,"unit":"Drum"},"20W40 1 LIT (CF4) MOTOLUBE":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"20W40 1 LIT SPRINTA F300 (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"20W40 210 LIT SPRINTA F300":{"sizeLtr":210,"nosPerUnit":null,"unit":"Barrel"},"20W40 55 LIT SPRINTA F300":{"sizeLtr":55,"nosPerUnit":null,"unit":"Drum"},"20W40 900 ML SPRINTA F300 (20)":{"sizeLtr":0.9,"nosPerUnit":20,"unit":"Case"},"20W50 1 LIT SPRINTA F300 (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"20W50 1 LIT SYNTIUM 300 SL (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"20W50 1 LTR CNG LUBE (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"20W50 210 LTR CNG LUBE BARREL":{"sizeLtr":210,"nosPerUnit":null,"unit":"Barrel"},"20W50 3 LIT SYNTIUM 300 SL (6)":{"sizeLtr":3,"nosPerUnit":6,"unit":"Case"},"20W50 3.5 LIT SYNTIUM 300 SL (4)":{"sizeLtr":3.5,"nosPerUnit":4,"unit":"Case"},"20W50 3.5 LTR CNG LUBE (4)":{"sizeLtr":3.5,"nosPerUnit":4,"unit":"Case"},"20W50 500 ML CNG LUBE (20)":{"sizeLtr":0.5,"nosPerUnit":20,"unit":"Case"},"3 WHEELER LIFE ENGINE OIL 2.75 LIT PETRONAS (6)":{"sizeLtr":2.75,"nosPerUnit":6,"unit":"Case"},"3 WHEELER LIFE ENGINE OIL 3 LIT PETRONAS (6)":{"sizeLtr":3,"nosPerUnit":6,"unit":"Case"},"3 WHEELER LIFE ENGINE OIL 500 ML PETRONAS (20)":{"sizeLtr":0.5,"nosPerUnit":20,"unit":"Case"},"5W30  1 LIT SYNTIUM 500 SN/CF (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"5W30  3 LIT SYNTIUM 500 SN/CF (6)":{"sizeLtr":3,"nosPerUnit":6,"unit":"Case"},"5W30  3.5 LIT SYNTIUM 500 SN/CF (4)":{"sizeLtr":3.5,"nosPerUnit":4,"unit":"Case"},"5W30 1 LIT SYNTIUM 3000 SN PLUS FULLY SYN (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"5W30 3.5 LIT SYNTIUM 3000 (SP) FULLY SYN (4)":{"sizeLtr":3.5,"nosPerUnit":4,"unit":"Case"},"5W30 3.5 LIT SYNTIUM 800 SN SEMI SYNT (4)":{"sizeLtr":3.5,"nosPerUnit":4,"unit":"Case"},"5W30 4 LIT TATA MOTORS GENUINE OIL (4)":{"sizeLtr":4,"nosPerUnit":4,"unit":"Case"},"5W30 5 LIT SYNTIUM 500 SN/CF (4)":{"sizeLtr":5,"nosPerUnit":4,"unit":"Case"},"5W30 50 LIT SYNTIUM 500 SN/CF":{"sizeLtr":50,"nosPerUnit":null,"unit":"Drum"},"5W30 600 ML SPRINTA A700 (SPOUT - SCOOTER PACK) (20)":{"sizeLtr":0.6,"nosPerUnit":20,"unit":"Case"},"5W30 SYNTIUM 500 SN/CF 210 LIT":{"sizeLtr":210,"nosPerUnit":null,"unit":"Barrel"},"5W40 4 LIT SYNTIUM 3000 SN (4)":{"sizeLtr":4,"nosPerUnit":4,"unit":"Case"},"5W40 MGDO 3.5 LIT SYNTIUM 3000SZ (4)":{"sizeLtr":3.5,"nosPerUnit":4,"unit":"Case"},"75W90 1 LTR TUTELA TRANSMISSION OIL (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"75W90 2.5 LTR TUTELA TRANSMISSION OIL (6)":{"sizeLtr":2.5,"nosPerUnit":6,"unit":"Case"},"80W90 1 LTR TUTELA TRANSMISSION OIL":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"80W90 20 LTR TUTELA TRANSMISSION OIL":{"sizeLtr":20,"nosPerUnit":null,"unit":"Bucket"},"80W90 TUTELA 300 MTF TRANSMISSION OIL 7 LIT":{"sizeLtr":7,"nosPerUnit":null,"unit":"Bucket"},"80W90 TUTELA TRANSMISSION OIL 2.5 LIT (6)":{"sizeLtr":2.5,"nosPerUnit":6,"unit":"Case"},"80W90 TUTELA TRANSMISSION OIL 5 LIT":{"sizeLtr":5,"nosPerUnit":4,"unit":"Case"},"85W140 1 LIT TUTELA TRANSMISSION OIL (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"85W140 210 LTR TUTELA TRANSMISSION OIL":{"sizeLtr":210,"nosPerUnit":null,"unit":"Barrel"},"85W140 5LIT TUTELA TRANSMISSION OIL":{"sizeLtr":5,"nosPerUnit":4,"unit":"Case"},"AGRI GOLD TRACTOR ENGINE OIL 7.5 LIT":{"sizeLtr":7.5,"nosPerUnit":null,"unit":"Bucket"},"AGRI GOLD TRACTOR ENGINE OIL 8.5 LIT":{"sizeLtr":8.5,"nosPerUnit":null,"unit":"Bucket"},"AGRI GOLD TRACTOR TRANSMISSION OIL 26 LIT":{"sizeLtr":26,"nosPerUnit":null,"unit":"Bucket"},"AKROS 20W30 20 LIT MULTI UTTO":{"sizeLtr":20,"nosPerUnit":null,"unit":"Bucket"},"AMBRA MASTER GOLD HSP 7.7 LIT":{"sizeLtr":7.7,"nosPerUnit":null,"unit":"Bucket"},"AMBRA SUPER GOLD 10 LIT":{"sizeLtr":10,"nosPerUnit":null,"unit":"Bucket"},"AMBRA SUPERTRAN 5 LIT TRANSMISSION OIL (4)":{"sizeLtr":5,"nosPerUnit":4,"unit":"Case"},"ATUL AUTO GENUINE OIL 1 LIT (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"ATUL AUTO GENUINE OIL 2.5 LIT BSVI (6)":{"sizeLtr":2.5,"nosPerUnit":6,"unit":"Case"},"ATUL AUTO GENUINE OIL 3 LIT (6)":{"sizeLtr":3,"nosPerUnit":6,"unit":"Case"},"ATUL AUTO GENUINE OIL 500 ML (20)":{"sizeLtr":0.5,"nosPerUnit":20,"unit":"Case"},"CF4 15W40 1 LIT URANIA 800 (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"CF4 15W40 10 LIT URANIA 800":{"sizeLtr":10,"nosPerUnit":null,"unit":"Bucket"},"CF4 15W40 3 LIT URANIA 800(6)":{"sizeLtr":3,"nosPerUnit":6,"unit":"Case"},"CF4 15W40 6 LIT URANIA 800":{"sizeLtr":6,"nosPerUnit":null,"unit":"Bucket"},"CF4 15W40 7.5 LIT URANIA 800":{"sizeLtr":7.5,"nosPerUnit":null,"unit":"Bucket"},"CF4 20W40 1 LIT URANIA 800":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"CF4 20W40 10 LIT URANIA 800":{"sizeLtr":10,"nosPerUnit":null,"unit":"Bucket"},"CF4 20W40 210 LIT URANIA 800":{"sizeLtr":210,"nosPerUnit":null,"unit":"Barrel"},"CF4 20W40 6 LIT URANIA 800":{"sizeLtr":6,"nosPerUnit":null,"unit":"Bucket"},"CF4 20W40 7.5 LIT URANIA 800":{"sizeLtr":7.5,"nosPerUnit":null,"unit":"Bucket"},"CF4 20W40 8.5 LIT URANIA 800":{"sizeLtr":8.5,"nosPerUnit":null,"unit":"Bucket"},"CH4 15W40 1 LIT URANIA 1000 (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"CH4 15W40 10 LIT URANIA 1000":{"sizeLtr":10,"nosPerUnit":null,"unit":"Bucket"},"CH4 15W40 210 LIT URANIA 1000":{"sizeLtr":210,"nosPerUnit":null,"unit":"Barrel"},"CH4 15W40 3 LIT URANIA 1000(6)":{"sizeLtr":3,"nosPerUnit":6,"unit":"Case"},"CH4 15W40 55 LIT URANIA 1000":{"sizeLtr":55,"nosPerUnit":null,"unit":"Drum"},"CH4 15W40 7.5 LIT URANIA 1000":{"sizeLtr":7.5,"nosPerUnit":null,"unit":"Bucket"},"CI4+ 15W 40 210 LIT URANIA 3000":{"sizeLtr":210,"nosPerUnit":null,"unit":"Barrel"},"CI4+ 15W40 1 LIT URANIA 3000":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"CI4+ 15W40 15 LIT URANIA 3000":{"sizeLtr":15,"nosPerUnit":null,"unit":"Bucket"},"CI4+ 15W40 2.5 LIT URANIA 3000 (6)":{"sizeLtr":2.5,"nosPerUnit":6,"unit":"Case"},"CI4+ 15W40 3.5 LIT SYNTIUM 500 (4)":{"sizeLtr":3.5,"nosPerUnit":4,"unit":"Case"},"CI4+ 15W40 55 LIT URANIA 3000":{"sizeLtr":55,"nosPerUnit":null,"unit":"Drum"},"CI4+ 15W40 6 LIT URANIA 3000":{"sizeLtr":6,"nosPerUnit":null,"unit":"Bucket"},"CI4+ 15W40 7.5 LIT URANIA 3000":{"sizeLtr":7.5,"nosPerUnit":null,"unit":"Bucket"},"CI4+ 15W40 8.5 LIT URANIA 3000":{"sizeLtr":8.5,"nosPerUnit":null,"unit":"Bucket"},"CK4 15W40 1 LIT URANIA 5000 IN (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"CK4 15W40 10 LIT URANIA 5000 IN":{"sizeLtr":10,"nosPerUnit":null,"unit":"Bucket"},"CK4 15W40 15 LIT URANIA 5000 IN":{"sizeLtr":15,"nosPerUnit":null,"unit":"Bucket"},"EPYX 140 1 LIT GEAR OIL TUTELA (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"EPYX 140 20 LIT GEAR OIL TUTELA":{"sizeLtr":20,"nosPerUnit":null,"unit":"Bucket"},"EPYX 140 5 LIT GEAR OIL TUTELA  (4)":{"sizeLtr":5,"nosPerUnit":4,"unit":"Case"},"EPYX 90 1 LIT GEAR OIL TUTELA (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"EPYX 90 20 LIT GEAR OIL TUTELA":{"sizeLtr":20,"nosPerUnit":null,"unit":"Bucket"},"EPYX 90 5 LIT GEAR OIL TUTELA (4)":{"sizeLtr":5,"nosPerUnit":4,"unit":"Case"},"HYDROCER 46 HYD OIL 26 LIT":{"sizeLtr":26,"nosPerUnit":null,"unit":"Bucket"},"MACH5 20W50 3.5 LIT SG (4)":{"sizeLtr":3.5,"nosPerUnit":4,"unit":"Case"},"MACH5 CH4 15W40 5 LIT DIESEL (4)":{"sizeLtr":5,"nosPerUnit":4,"unit":"Case"},"MOTOLUB 15W40 10 LTR (CF4)":{"sizeLtr":10,"nosPerUnit":4,"unit":"Bucket"},"MOTOLUB 15W40 7.5 LTR (CF4)":{"sizeLtr":7.5,"nosPerUnit":null,"unit":"Bucket"},"MOTOLUB 20W40 10 LIT (CF4)":{"sizeLtr":10,"nosPerUnit":null,"unit":"Bucket"},"MOTOLUB 20W40 7.5 (CF4)":{"sizeLtr":7.5,"nosPerUnit":null,"unit":"Bucket"},"MOTOLUB CF4 15W40 0.5 LIT (20)":{"sizeLtr":0.5,"nosPerUnit":20,"unit":"Case"},"MOTOLUB CF4 15W40 1 LIT (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"MOTOLUB CF4 15W40 3 LIT (6)":{"sizeLtr":3,"nosPerUnit":6,"unit":"Case"},"MOTOLUB CF4 15W40 6 LIT":{"sizeLtr":6,"nosPerUnit":null,"unit":"Bucket"},"NEXPRO TRANSMISSION OIL 20 LIT":{"sizeLtr":20,"nosPerUnit":null,"unit":"Bucket"},"PROFAR ANTIFREEZE COOLANT OIL 1 LIT (10)":{"sizeLtr":1,"nosPerUnit":10,"unit":"Case"},"PUMPSET OIL 3.5 LIT CF40 PETRONAS (4)":{"sizeLtr":3.5,"nosPerUnit":4,"unit":"Case"},"SPRINTA 20W40 F100 55 LIT SJ DRUM":{"sizeLtr":55,"nosPerUnit":null,"unit":"Drum"},"SPRINTA 20W50 55 LIT F300":{"sizeLtr":55,"nosPerUnit":null,"unit":"Drum"},"SPRINTA F100 15W50 1.2 LIT- SL (10)":{"sizeLtr":1.2,"nosPerUnit":10,"unit":"Case"},"SPRINTA F100 15W50 2.5 LIT- SL (6)":{"sizeLtr":2.5,"nosPerUnit":6,"unit":"Case"},"SPRINTA F100 20W40 1 LIT- SJ (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"SPRINTA F100 20W40 900 ML - SJ (20)":{"sizeLtr":0.9,"nosPerUnit":20,"unit":"Case"},"SPRINTA F300 20W50 1.2 L - SL (10)":{"sizeLtr":1.2,"nosPerUnit":10,"unit":"Case"},"SPRINTA F900 15W50 2.5 LIT (6)":{"sizeLtr":2.5,"nosPerUnit":6,"unit":"Case"},"TATA MOTORS CH4 15W40 GENUINE OIL 4 LIT (4)":{"sizeLtr":4,"nosPerUnit":4,"unit":"Case"},"TATA MOTORS GENUINE OIL 1 LIT (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"TATA MOTORS GENUINE OIL 5 LIT (4)":{"sizeLtr":5,"nosPerUnit":4,"unit":"Case"},"TMGO 0W20 5 LIT FULLY SYNT":{"sizeLtr":5,"nosPerUnit":4,"unit":"Case"},"TMGO 5W30 5 LIT FULLY SYNT (HARRIER)":{"sizeLtr":5,"nosPerUnit":4,"unit":"Case"},"TMGO CI4 + 15W40 5 LIT (4)":{"sizeLtr":5,"nosPerUnit":4,"unit":"Case"},"TMOO T-PRO 1000 0W20 3.5 LIT (4)":{"sizeLtr":3.5,"nosPerUnit":4,"unit":"Case"},"TMOO T-PRO 1000 0W20 50 LIT":{"sizeLtr":50,"nosPerUnit":null,"unit":"Drum"},"TUTELA AKROS 80W 20 LIT UTTO F100":{"sizeLtr":20,"nosPerUnit":null,"unit":"Bucket"},"TUTELA AKROS 80W 5 LIT UTTO F100 (4)":{"sizeLtr":5,"nosPerUnit":4,"unit":"Case"},"TUTELA GREASE 1 KG (MR3) (12)":{"sizeLtr":1,"nosPerUnit":12,"unit":"Case"},"TUTELA GREASE 2 KG (MR3) (6)":{"sizeLtr":2,"nosPerUnit":6,"unit":"Case"},"TUTELA GREASE 3 KG  (4) (MR3)":{"sizeLtr":3,"nosPerUnit":4,"unit":"Case"},"TUTELA GREASE 500GMS (MR3)(24)":{"sizeLtr":0.5,"nosPerUnit":24,"unit":"Case"},"TUTELA LONG LIFE GREASE 1 KG (RED - NLGI 3) (12)":{"sizeLtr":1,"nosPerUnit":12,"unit":"Case"},"TUTELA LONG LIFE GREASE 20 KG (RED - NLGI 3)":{"sizeLtr":20,"nosPerUnit":null,"unit":"Case"},"TUTELA LONG LIFE GREASE 3 KG (RED - NLGI 3) (4)":{"sizeLtr":3,"nosPerUnit":4,"unit":"Case"},"TUTELA LONG LIFE GREASE 500 GMS (RED - NLGI 3) (24)":{"sizeLtr":0.5,"nosPerUnit":24,"unit":"Case"},"TUTELA LONG LIFE GREASE 7 KG (RED - NLGI 3)":{"sizeLtr":7,"nosPerUnit":null,"unit":"Case"},"TUTELA MULTI UTTO 5 LIT 20W30 (AKROS) (4)":{"sizeLtr":5,"nosPerUnit":4,"unit":"Case"},"TUTELA SUPREME LL GREASE 1 KG (GREEN-NLGI 3) (12)":{"sizeLtr":1,"nosPerUnit":12,"unit":"Case"},"TUTELA SUPREME LL GREASE 5 KG (GREEN-NLGI 3) (2)":{"sizeLtr":5,"nosPerUnit":2,"unit":"Case"},"TUTELA SUPREME LL GREASE 500 GM (GREEN-NLGI 3) (24)":{"sizeLtr":0.5,"nosPerUnit":24,"unit":"Case"},"TVS TRU4 RACE PRO 15W50 1.7 L (OE PACK) (9)":{"sizeLtr":1.7,"nosPerUnit":9,"unit":"Case"},"TVS TRU4 RACE PRO 15W50 FULLY SYN 1.2 LIT (10)":{"sizeLtr":1.2,"nosPerUnit":10,"unit":"Case"},"TVS TRU4 RACE PRO 15W50 FULLY SYN 2.5 LIT (6)":{"sizeLtr":2.5,"nosPerUnit":6,"unit":"Case"},"10W30 55 LIT SPRINTA F300 SL":{"sizeLtr":55,"nosPerUnit":null,"unit":"Drum"},"5W30 7 LIT SYNTIUM 500 SN/CF":{"sizeLtr":7,"nosPerUnit":null,"unit":"Bucket"},"5W40 1 LIT SYNTIUM 3000 SM (FULL SYNT) COOLTECH (20":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"5W40 MGDO SYNTIUM 3000SZ 210 LIT BARREL":{"sizeLtr":210,"nosPerUnit":null,"unit":"Barrel"},"80W90 210 LTR TUTELA TRANSMISSION OIL":{"sizeLtr":210,"nosPerUnit":null,"unit":"Barrel"},"85W140 12 LIT TUTELA 300 EP TRANSMISSION OIL":{"sizeLtr":12,"nosPerUnit":null,"unit":"Bucket"},"AGRI GOLD TRACTOR ENGINE OIL 10 LIT":{"sizeLtr":10,"nosPerUnit":null,"unit":"Bucket"},"AGRI GOLD UTTO TRANSMISSION OIL 26LIT (AKROS)":{"sizeLtr":26,"nosPerUnit":null,"unit":"Bucket"},"AGRI GOLD UTTO TRANSMISSION OIL 5 LIT (AKROS)":{"sizeLtr":5,"nosPerUnit":4,"unit":"Case"},"CF4 15W40 210 LIT URANIA 800":{"sizeLtr":210,"nosPerUnit":null,"unit":"Barrel"},"CF4 15W40 55 LIT URANIA 800":{"sizeLtr":55,"nosPerUnit":null,"unit":"Drum"},"CF4 20W40 55 LIT URANIA 800":{"sizeLtr":55,"nosPerUnit":null,"unit":"Drum"},"CI4+ 15W40 10 LIT URANIA 3000":{"sizeLtr":10,"nosPerUnit":null,"unit":"Bucket"},"CI4+ 15W40 11 LIT URANIA 3000":{"sizeLtr":11,"nosPerUnit":null,"unit":"Bucket"},"CI4+ 15W40 18 LIT URANIA 3000":{"sizeLtr":18,"nosPerUnit":null,"unit":"Bucket"},"EPYX 140 TUTELA 210 LIT TRANSMISSION OIL":{"sizeLtr":210,"nosPerUnit":null,"unit":"Barrel"},"EPYX 90 TUTELA 210 LIT TRANSMISSION OIL":{"sizeLtr":210,"nosPerUnit":null,"unit":"Barrel"},"HYDROCER 68 HYD OIL 20 LIT":{"sizeLtr":20,"nosPerUnit":null,"unit":"Bucket"},"HYDROCER 68 HYD OIL 210 LIT":{"sizeLtr":210,"nosPerUnit":null,"unit":"Barrel"},"HYDROCER 68 HYD OIL 26 LIT":{"sizeLtr":26,"nosPerUnit":null,"unit":"Bucket"},"POWER STEERING OIL ATF D3 1 LIT (20)":{"sizeLtr":1,"nosPerUnit":20,"unit":"Case"},"TMOO T-PRO 1000 0W20 210 LIT":{"sizeLtr":210,"nosPerUnit":null,"unit":"Barrel"},"TUTELA GREASE 20 KG (MR3)":{"sizeLtr":20,"nosPerUnit":null,"unit":"Bucket"},"TUTELA LONG LIFE GREASE 5 KG (RED - NLGI 3) (2)":{"sizeLtr":5,"nosPerUnit":2,"unit":"Case"},"5W30  4.5 LIT SYNTIUM 500 SN/CF (4)":{"sizeLtr":4.5,"nosPerUnit":4,"unit":"Case"}}''')


def _catalog_ltr_guess(stock_item_name, qty):
    """
    Ltr for one Sales-voucher inventory line, given the item's own name
    (matched against ITEM_LTR_CATALOG by exact string, same convention as
    every other live-overlay item match in this project) and Tally's raw
    quantity NUMBER (already stripped of its unit suffix by _num() before
    this is called).

    CONFIRMED 18 Sep 2026 against a real Sales voucher's raw XML (owner's
    own --dump-sales output): Tally's ACTUALQTY/BILLEDQTY for a Sales
    voucher inventory line is in individual NOS (pieces) - e.g.
    "<ACTUALQTY TYPE=\"Quantity\"> 5 NOS</ACTUALQTY>" - NOT in the same
    "Case"/order-unit CLOSINGBALANCE uses for live stock, as originally
    guessed (and flagged as unverified) when this function was first
    written. So the correct conversion is simply qty * sizeLtr - each
    unit sold already IS one individual bottle/can/drum, no nosPerUnit
    multiplication needed (that would overstate every case-packed item's
    Ltr by its own case size, e.g. 20x too high for a 20-per-case item).
    `nosPerUnit` is kept in ITEM_LTR_CATALOG/the entry lookup for
    reference and any future use, just no longer multiplied in here.
    Returns None if the item name isn't in the catalog at all, or qty is
    missing.
    """
    if qty is None:
        return None
    entry = ITEM_LTR_CATALOG.get(stock_item_name)
    if not entry:
        return None
    return round(qty * entry["sizeLtr"], 2)


SALES_VOUCHERS_TEMPLATE = """<ENVELOPE>
 <HEADER>
  <VERSION>1</VERSION>
  <TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE>
  <ID>SalesVouchersDumpCollection</ID>
 </HEADER>
 <BODY>
  <DESC>
   <STATICVARIABLES>
    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
    <SVFROMDATE>{from_date}</SVFROMDATE>
    <SVTODATE>{to_date}</SVTODATE>
   </STATICVARIABLES>
   <TDL>
    <TDLMESSAGE>
     <SYSTEM TYPE="Formulae" NAME="SalesVouchersOnly">$VoucherTypeName = "Sales" OR $VoucherTypeName = "Sales - PETRONAS"</SYSTEM>
     <COLLECTION NAME="SalesVouchersDumpCollection" ISMODIFY="No">
      <TYPE>Voucher</TYPE>
      <FETCH>DATE, VOUCHERTYPENAME, PARTYLEDGERNAME</FETCH>
      <FETCH>ALLINVENTORYENTRIES.LIST</FETCH>
      <FILTER>SalesVouchersOnly</FILTER>
     </COLLECTION>
    </TDLMESSAGE>
   </TDL>
  </DESC>
 </BODY>
</ENVELOPE>"""

# Widened 18 Sep 2026, same day the real fetch_sales() first ran live -
# the owner's real watcher reported "0 lubricant-item line(s)" for every
# DSR, which was suspicious: the 12 Sep 2026 Collection investigation
# (check_collection(), see dsr-app-registry.md) already found a SEPARATE
# voucher type on this exact Tally setup, "Sales - PETRONAS", carrying
# far more rupee value than plain "Sales" for the same dealers (Arun's
# 12-day sample: Sales Rs 21,525 vs Sales - PETRONAS Rs 2,37,716) - but
# every --dump-sales run up to this point only ever asked Tally for
# VoucherTypeName = "Sales", server-side, so "Sales - PETRONAS" vouchers
# were never even included in the response to filter client-side - the
# zero wasn't a client-side matching bug, Tally was never sending them.
# The <FILTER> formula above and both type-checks below (dump_sales() and
# fetch_sales()) now match either name. Not yet independently re-verified
# against a real --dump-sales run with this widened filter (that
# verification happens live, via the owner's own real dashboard numbers
# against what he knows was actually billed - same final-check pattern
# used throughout this project) - if "Sales - PETRONAS" vouchers turn out
# to have a different inventory-line shape than plain "Sales" (untested,
# though very likely identical since both are standard Sales-class
# vouchers), the raw ACTUALQTY/BILLEDQTY text printed by --dump-sales
# will show it immediately, same as every other field in this project.
SALES_VOUCHER_TYPE_NAMES = ("sales", "sales - petronas")


def dump_sales(dsr_fragment=None, from_date=None, to_date=None, max_lines=40):
    """
    --dump-sales ["arun"] [YYYYMMDD-from] [YYYYMMDD-to]

    Diagnostic only - pushes nothing anywhere, no fetch_sales() exists
    yet. Defaults to TODAY only if no dates given (deliberately narrow -
    a Sales voucher dump can be large company-wide; widen with explicit
    from/to once today's shape looks right). dsr_fragment optionally
    narrows printed vouchers to one DSR's dealers (matched the same way
    --check-collection does); omit it to see every dealer's Sales
    vouchers for the period.

    Prints, per matching voucher: DATE, dealer (PARTYLEDGERNAME), the
    resolved DSR (or "(no DSR match)"), then EVERY inventory line's
    STOCKITEMNAME, the RAW unparsed ACTUALQTY/BILLEDQTY text (unit suffix
    and all - this is the actual answer to the open unit question, read
    it directly), RATE, AMOUNT, and - only if the item name matches
    ITEM_LTR_CATALOG - a computed Ltr guess (see _catalog_ltr_guess()'s
    docstring for the assumption it makes). Capped at max_lines total
    inventory lines printed (not vouchers) so a busy day doesn't flood
    the terminal; prints how many more exist beyond the cap.
    """
    now = datetime.now()
    from_d = from_date or now.strftime("%Y%m%d")
    to_d = to_date or now.strftime("%Y%m%d")

    target = None
    dealer_names_for_dsr = None
    print("Resolving Tally's ledger/group hierarchy...")
    ledger_parent, resolve = _build_group_resolver()

    if dsr_fragment:
        target = _resolve_dsr_group_name(dsr_fragment)
        if not target:
            print("'{}' didn't match any of the known DSR groups:".format(dsr_fragment))
            for g in KNOWN_DSR_GROUP_NAMES:
                print("  " + g)
            return
        dealer_names_for_dsr = set(
            lname for lname, lparent in ledger_parent.items()
            if lparent and resolve(lparent) == target
        )
        print("  {} ledger(s) resolve to {}.".format(len(dealer_names_for_dsr), target))

    def resolve_dealer_to_dsr(dealer_name):
        lparent = ledger_parent.get(dealer_name)
        if not lparent:
            return None
        return resolve(lparent)

    print("Fetching Sales vouchers from {} to {} - this can take a while on a busy "
          "company, please wait...".format(from_d, to_d))
    xml_body = SALES_VOUCHERS_TEMPLATE.format(from_date=from_d, to_date=to_d)
    xml_text = tally_request(xml_body, timeout=600)
    try:
        with open("tally_sales_dump.xml", "w", encoding="utf-8") as f:
            f.write(xml_text)
        print("(raw response saved to tally_sales_dump.xml)")
    except Exception as e:
        print("(could not save raw response: {})".format(e))
    root = ET.fromstring(xml_text)

    voucher_count = 0
    leaf_voucher_count = 0
    in_range_count = 0
    matched_voucher_count = 0
    irrelevant_voucher_count = 0
    lines_printed = 0
    lines_skipped = 0
    dsr_ltr_guess_totals = {}
    distinct_voucher_types_seen = set()
    sample_voucher_xml = None

    for v in root.iter("VOUCHER"):
        voucher_count += 1
        if len(list(v)) == 0:
            leaf_voucher_count += 1
            continue
        if sample_voucher_xml is None:
            sample_voucher_xml = ET.tostring(v, encoding="unicode")
        vdate = (_text(v, "DATE") or "").strip()
        if vdate and (vdate < from_d or vdate > to_d):
            continue
        in_range_count += 1
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip()
        distinct_voucher_types_seen.add(vtype)
        if vtype.lower() not in SALES_VOUCHER_TYPE_NAMES:
            continue
        party = (_text(v, "PARTYLEDGERNAME") or "").strip()
        if dealer_names_for_dsr is not None and party not in dealer_names_for_dsr:
            continue
        matched_voucher_count += 1
        voucher_dsr = resolve_dealer_to_dsr(party) or "(no DSR match)"

        entries = list(v.iter("ALLINVENTORYENTRIES.LIST"))
        line_results = []
        voucher_has_catalog_item = False
        for entry in entries:
            item_name = (_text(entry, "STOCKITEMNAME") or "").strip()
            raw_actual = _text(entry, "ACTUALQTY")
            raw_billed = _text(entry, "BILLEDQTY")
            rate = _text(entry, "RATE")
            amount = _num(_text(entry, "AMOUNT"))
            qty_num = _num(raw_billed) if raw_billed is not None else _num(raw_actual)
            ltr_guess = _catalog_ltr_guess(item_name, qty_num)
            if ltr_guess is not None:
                voucher_has_catalog_item = True
            if voucher_dsr != "(no DSR match)" and ltr_guess is not None:
                dsr_ltr_guess_totals[voucher_dsr] = dsr_ltr_guess_totals.get(voucher_dsr, 0.0) + ltr_guess
            line_results.append((item_name, raw_actual, raw_billed, rate, amount, ltr_guess))

        if not voucher_has_catalog_item:
            irrelevant_voucher_count += 1
            continue

        if lines_printed < max_lines:
            print("\n[{}] {}  ->  {}".format(vdate, party, voucher_dsr))
        for item_name, raw_actual, raw_billed, rate, amount, ltr_guess in line_results:
            if lines_printed < max_lines:
                print("    {}  |  ACTUALQTY(raw)={!r}  BILLEDQTY(raw)={!r}  RATE={}  AMOUNT={}  "
                      "Ltr-guess={}".format(item_name, raw_actual, raw_billed, rate, amount, ltr_guess))
                lines_printed += 1
            else:
                lines_skipped += 1

    print("\n" + "=" * 70)
    print("Total elements tagged <VOUCHER>: {}   of which {} were real records (had child "
          "elements) and {} were stray/leaf (e.g. '<VOUCHER>0</VOUCHER>', not counted below "
          "at all).".format(voucher_count, voucher_count - leaf_voucher_count, leaf_voucher_count))
    print("Real voucher records in requested date range ({} to {}): {}".format(
        from_d, to_d, in_range_count))
    print("Distinct voucher types seen among in-range real vouchers: {}".format(
        sorted(distinct_voucher_types_seen)))
    print("Sales vouchers matched (period{}): {}   of which {} had NO known lubricant item "
          "(not printed above - this company also sells auto/tractor spare parts through the "
          "same Tally company) and {} had at least one, printed above.".format(
              " + " + target if target else "", matched_voucher_count, irrelevant_voucher_count,
              matched_voucher_count - irrelevant_voucher_count))
    if (voucher_count - leaf_voucher_count) > in_range_count:
        print("({} real vouchers in the response were OUTSIDE the requested date range - same "
              "SVFROMDATE/SVTODATE-doesn't-scope-a-raw-Voucher-Collection quirk already "
              "found for Collection, not a new bug.)".format((voucher_count - leaf_voucher_count) - in_range_count))
    if lines_skipped:
        print("({} more inventory line(s) not printed - narrow the date range or DSR "
              "to see them.)".format(lines_skipped))
    print("\nLtr-guess totals by DSR this period (ASSUMES qty is in Case/order-units - "
          "confirm against the raw ACTUALQTY/BILLEDQTY unit text above before trusting "
          "these numbers at all):")
    if dsr_ltr_guess_totals:
        for dsr, total in sorted(dsr_ltr_guess_totals.items()):
            print("  {}  ->  {} Ltr (guess)".format(dsr, round(total, 2)))
    else:
        print("  (none matched a known item name - nothing to total)")
    real_voucher_count = voucher_count - leaf_voucher_count
    if matched_voucher_count == 0 and real_voucher_count > 0:
        print("\n" + "=" * 70)
        print("ZERO vouchers matched 'sales' by type, out of {} real voucher record(s) found -".format(real_voucher_count))
        print("this is the same 'near-100% unmatched' signature that meant a wrong field/")
        print("value guess in every earlier round of this project (Outstanding's BILLTYPE,")
        print("Rate's CLOSINGRATE vs StandardPrice) - most likely this Tally company's real")
        print("Sales voucher type name isn't literally 'Sales' (or the <FILTER> isn't being")
        print("honored the way it was for Receipt - same kind of raw-Voucher-Collection quirk")
        print("already found once for SVFROMDATE/SVTODATE). Raw sample voucher below, exactly")
        print("as Tally sent it - the real VOUCHERTYPENAME value (or its absence) is in here:")
        if sample_voucher_xml:
            print("\n--- One raw voucher record from the response, unfiltered by type ---")
            print(sample_voucher_xml[:3000])
        else:
            print("(no voucher at all came back in the response - see tally_sales_dump.xml)")
    elif real_voucher_count == 0:
        print("\n" + "=" * 70)
        print("No real voucher records found at all (every <VOUCHER>-tagged element in the")
        print("response was a stray leaf, not an actual transaction) - open tally_sales_dump.xml")
        print("directly (any text editor) and search it for a dealer name you know was billed")
        print("today, then send back ~40-50 lines around that match so the real record shape")
        print("can be read directly - this diagnostic's own tag-based search isn't finding it.")
    print("\nSend this whole output back - specifically: (1) does the raw ACTUALQTY/")
    print("BILLEDQTY text say 'Case' (or similar order-unit) or 'Nos' (individual bottles)")
    print("for a few known items, and (2) does the Ltr-guess total for a DSR you can check")
    print("against a known real day's billing look right, too high, or too low.")


# -- Day Book export diagnostic, added 26 Sep 2026 -----------------------
# Owner's request (Round 5, with a real "Petronas Day Book" Tally
# screenshot + his own TDL script attached): a table-format Excel export
# matching that Day Book report exactly - Voucher Number/Date/Voucher
# Type/Party/Party Group/Item/Item Group/Qty/Alt Qty/Amount (incl 18%
# GST, confirmed by the owner directly - not the TDL's own raw
# pre-GST AmtBeforeGST field), one row per voucher INVENTORY LINE (not
# one row per bill). Per the owner's own "ANY DOUBT ASK AND CONFIRM ME
# BEFORE PROCEED" instruction and this project's standing diagnostic-
# first discipline (every other Tally-derived field in this file -
# Outstanding, Rate, MRP, Sales qty unit, Stock qty unit - needed a real
# dump before being trusted), this ships as a DIAGNOSTIC ONLY - nothing
# pushed to the Sheet, no backfill, no monthly-tab storage, no Excel
# export button yet. Those all wait on the owner confirming this
# diagnostic's output against his real Sep 1-2 Day Book screenshot.
STOCK_ITEM_GROUP_COLLECTION_XML = """<ENVELOPE>
 <HEADER>
  <VERSION>1</VERSION>
  <TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE>
  <ID>StockItemGroupDumpCollection</ID>
 </HEADER>
 <BODY>
  <DESC>
   <STATICVARIABLES>
    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
   </STATICVARIABLES>
   <TDL>
    <TDLMESSAGE>
     <COLLECTION NAME="StockItemGroupDumpCollection" ISMODIFY="No">
      <TYPE>StockItem</TYPE>
      <FETCH>NAME, PARENT</FETCH>
     </COLLECTION>
    </TDLMESSAGE>
   </TDL>
  </DESC>
 </BODY>
</ENVELOPE>"""


# -- Dealer Purchase History diagnostic, added 11 Oct 2026 ---------------
# Owner's request ("want to see a dealer history of purchase by selection
# of a dealer name"), confirmed after a mockup: full Tally-backed,
# item-by-item, on both the dashboard (any dealer, any DSR) and each DSR
# app (their own dealers). fetch_sales() above already proves the raw
# voucher shape (date/dealer/items/qty/rate/amount) parses fine - this is
# NOT re-diagnosing that. What's new here and needs confirming before the
# real feature gets built: (1) how far back a SINGLE dealer's own
# purchase history actually goes in this Tally company (fetch_sales()
# only ever looks 120 days back - a full history view needs much more,
# and nobody's confirmed yet how many years of real data exist), and (2)
# whether a dealer's history is cleanly under ONE ledger name the whole
# time, or split across more than one near-identical name (a known
# recurring issue on this project - see --dump-groups' own notes) which
# would silently truncate the "full" history if not accounted for.
def dump_purchase_history(dealer_fragment, days_back=None):
    """
    --dump-purchase-history <dealer fragment> [days-back]

    Diagnostic only - pushes nothing anywhere, no fetch_purchase_history()
    exists yet. Unlike fetch_sales(), which only keeps running TOTALS per
    dealer (today/month/6-mo avg), this prints every individual Sales
    voucher for ONE matched dealer, OLDEST FIRST: date, every item line
    (name, qty, rate, amount), and the voucher's own total - exactly the
    order-by-order view the real feature needs to show.

    dealer_fragment is matched the same substring-on-ledger-name way as
    --dump-creditlimit/--dump-groups (case-insensitive, anywhere in the
    name). If more than one ledger matches, ALL of them are listed (so a
    narrower fragment can be used on a follow-up run) but only the FIRST
    match's vouchers are printed, and every other matching name is noted
    at the end in case the dealer's real history turns out to be split
    across more than one ledger.

    days_back (default 365) caps how far back the request asks Tally to
    look - deliberately not unbounded, since this project's Tally company
    may hold many years of vouchers and a request with no lower bound at
    all risks being very slow for no benefit if real history turns out to
    only go back a year or two anyway. Pass 0 for "everything Tally has"
    (uses a fixed 1 Jan 2015 floor instead of no bound at all, since Tally
    still needs SOME from-date).

    Reports at the end: how many vouchers were found and the earliest/
    latest date among them (so the owner can see how far back real data
    actually goes for this dealer - compare against what he knows from
    Tally directly), total Rs and total Ltr-guess across everything
    printed, and the "more than one matching ledger name" warning above
    if relevant.
    """
    out_dir = os.path.dirname(os.path.abspath(__file__))
    now = datetime.now()
    if days_back is None:
        days_back = 365
    days_back = int(days_back)
    from_d = "20150101" if days_back <= 0 else (now - timedelta(days=days_back)).strftime("%Y%m%d")
    to_d = now.strftime("%Y%m%d")

    print("Resolving Tally's ledger/group hierarchy...")
    ledger_parent, resolve = _build_group_resolver()
    frag_upper = dealer_fragment.strip().upper()
    all_matches = sorted(set(
        lname for lname in ledger_parent.keys() if frag_upper in lname.upper()
    ))
    if not all_matches:
        print("'{}' didn't match any ledger name at all. Try a shorter fragment.".format(dealer_fragment))
        return
    target_dealer = all_matches[0]
    print("Matched {} ledger(s) containing '{}': {}".format(len(all_matches), dealer_fragment, all_matches))
    print("Using the first match: '{}'\n".format(target_dealer))

    print("Fetching Sales vouchers from {} to {} (days_back={}) - this can take a while on a "
          "busy company, please wait...".format(from_d, to_d, days_back))
    xml_body = SALES_VOUCHERS_TEMPLATE.format(from_date=from_d, to_date=to_d)
    xml_text = tally_request(xml_body, timeout=600)
    path = os.path.join(out_dir, "tally_purchase_history_dump.xml")
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write(xml_text)
        print("(raw response saved to {})".format(path))
    except Exception as e:
        print("(could not save raw response: {})".format(e))
    root = ET.fromstring(xml_text)

    vouchers_out = []
    for v in root.iter("VOUCHER"):
        if len(list(v)) == 0:
            continue
        vdate = (_text(v, "DATE") or "").strip()
        if not vdate or vdate < from_d or vdate > to_d:
            continue
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip()
        if vtype.lower() not in SALES_VOUCHER_TYPE_NAMES:
            continue
        party = (_text(v, "PARTYLEDGERNAME") or "").strip()
        if party != target_dealer:
            continue
        line_results = []
        voucher_rs_total = 0.0
        voucher_ltr_total = 0.0
        for entry in v.iter("ALLINVENTORYENTRIES.LIST"):
            item_name = (_text(entry, "STOCKITEMNAME") or "").strip()
            if not item_name:
                continue
            raw_billed = _text(entry, "BILLEDQTY")
            raw_actual = _text(entry, "ACTUALQTY")
            qty_num = _num(raw_billed) if raw_billed is not None else _num(raw_actual)
            amount = _num(_text(entry, "AMOUNT"))
            rate = _text(entry, "RATE")
            ltr_guess = _catalog_ltr_guess(item_name, qty_num)
            if amount is not None:
                voucher_rs_total += abs(amount)
            if ltr_guess is not None:
                voucher_ltr_total += ltr_guess
            line_results.append((item_name, raw_actual, raw_billed, rate, amount, ltr_guess))
        if not line_results:
            continue
        vouchers_out.append((vdate, line_results, voucher_rs_total, voucher_ltr_total))

    vouchers_out.sort(key=lambda row: row[0])  # oldest first

    print("\n" + "=" * 70)
    if not vouchers_out:
        print("No Sales vouchers found for '{}' in {} to {}.".format(target_dealer, from_d, to_d))
        print("Either this dealer has no Sales in that window, or the ledger name match above")
        print("isn't the right one - try a different fragment, or re-run with a larger")
        print("days_back (or 0 for everything back to 2015).")
    else:
        for vdate, line_results, voucher_rs_total, voucher_ltr_total in vouchers_out:
            print("\n[{}]  Total Rs {}   Ltr-guess {}".format(
                vdate, round(voucher_rs_total, 2), round(voucher_ltr_total, 2)))
            for item_name, raw_actual, raw_billed, rate, amount, ltr_guess in line_results:
                print("    {}  |  qty(raw)={!r}/{!r}  RATE={}  AMOUNT={}  Ltr-guess={}".format(
                    item_name, raw_billed, raw_actual, rate, amount, ltr_guess))
        total_rs = sum(row[2] for row in vouchers_out)
        total_ltr = sum(row[3] for row in vouchers_out)
        print("\n" + "=" * 70)
        print("{} voucher(s) for '{}', {} to {}.".format(
            len(vouchers_out), target_dealer, vouchers_out[0][0], vouchers_out[-1][0]))
        print("Total across all of them: Rs {}   Ltr-guess {}".format(
            round(total_rs, 2), round(total_ltr, 2)))
    if len(all_matches) > 1:
        print("\nNOTE: {} other ledger name(s) also matched '{}' and were NOT included above - "
              "if '{}'s real history is split across more than one name, re-run this once per "
              "name and compare: {}".format(
                  len(all_matches) - 1, dealer_fragment, target_dealer,
                  [n for n in all_matches if n != target_dealer]))
    print("\nSend this whole output back - specifically: (1) does the earliest date above match")
    print("what you'd expect from Tally directly for this dealer (or does real history actually")
    print("go back further than days_back covered - re-run with a bigger number or 0 if so), and")
    print("(2) do the Rs/Ltr-guess totals for a voucher you can check against a known real order")
    print("look right.")


# -- Invoice stats diagnostic, added 11 Oct 2026 --------------------------
# Owner's request, scoped down after the DSR year-month table mockup
# ("CUSTOMER COUNT IS THE MB COUNT IN SALES WE ATLREADY CALCULATED IN DSR
# APP, SKU COUNT IS THE ITEMS SALES COUNT WE ALREADY DID, INVOICES AND AVG
# INVOICE ONLY WANTS TO TAKE DATA"): Customer count (= the existing Sales
# "MB" - dealers billed this month, see countMB_() in dashboard_index.html
# and the "MB" comment in fetch_sales() above) and Sku count (= the
# existing distinct-items-sold-per-DSR count, from dsr_item_totals/
# salesByDsrItem) are BOTH already computed elsewhere on this project and
# need no new diagnostic. Invoices (count of distinct Sales vouchers/bills)
# and Avg Qty/Avg Val per invoice have never been computed anywhere on this
# project before - per the standing "never guess a new Tally structure"
# rule, this ships as a diagnostic first, same as dump_sales()/
# dump_purchase_history() above: pushes nothing anywhere, no real
# fetch_invoice_stats() exists yet.
#
# One <VOUCHER> element (that matches a known Sales voucher type and has
# at least one recognized lubricant item line, same filter dump_sales()
# already uses) = one invoice/bill, same assumption the rest of this file
# already makes everywhere a "voucher" is treated as one sale. Per-DSR
# totals resolve each voucher's dealer the same way dump_sales()/
# check-collection do (via the Tally ledger/group hierarchy walk).
def dump_invoice_stats(dsr_fragment=None, from_date=None, to_date=None):
    """
    --dump-invoice-stats ["arun"] [YYYYMM01-from] [YYYYMMDD-to]

    Diagnostic only. Defaults to the current calendar month if no dates
    given (the DSR year-month table needs one month at a time - re-run
    once per month, e.g. for Apr'26 through Oct'26, to build up a full
    financial year). dsr_fragment optionally narrows to one DSR (matched
    the same way --check-collection/--dump-sales do); omit it to see
    every DSR for the period.

    Prints, per matching invoice (voucher): DATE, dealer, resolved DSR,
    the voucher's own total Ltr-guess and total Rs. Then, per DSR: invoice
    count, total Ltr, total Rs, Avg Qty/invoice (Ltr), Avg Val/invoice
    (Rs) - exactly the 3 figures the DSR year-month table mockup's
    "Invoices"/"Avg Qty/invoice"/"Avg Val/invoice" rows need, one DSR-
    month at a time.
    """
    now = datetime.now()
    from_d = from_date or now.strftime("%Y%m01")
    to_d = to_date or now.strftime("%Y%m%d")

    target = None
    dealer_names_for_dsr = None
    print("Resolving Tally's ledger/group hierarchy...")
    ledger_parent, resolve = _build_group_resolver()

    if dsr_fragment:
        target = _resolve_dsr_group_name(dsr_fragment)
        if not target:
            print("'{}' didn't match any of the known DSR groups:".format(dsr_fragment))
            for g in KNOWN_DSR_GROUP_NAMES:
                print("  " + g)
            return
        dealer_names_for_dsr = set(
            lname for lname, lparent in ledger_parent.items()
            if lparent and resolve(lparent) == target
        )
        print("  {} ledger(s) resolve to {}.".format(len(dealer_names_for_dsr), target))

    def resolve_dealer_to_dsr(dealer_name):
        lparent = ledger_parent.get(dealer_name)
        if not lparent:
            return None
        return resolve(lparent)

    print("Fetching Sales vouchers from {} to {} - this can take a while on a busy "
          "company, please wait...".format(from_d, to_d))
    xml_body = SALES_VOUCHERS_TEMPLATE.format(from_date=from_d, to_date=to_d)
    xml_text = tally_request(xml_body, timeout=600)
    try:
        with open("tally_invoice_stats_dump.xml", "w", encoding="utf-8") as f:
            f.write(xml_text)
        print("(raw response saved to tally_invoice_stats_dump.xml)")
    except Exception as e:
        print("(could not save raw response: {})".format(e))
    root = ET.fromstring(xml_text)

    by_dsr = {}  # dsr -> {"invoices": int, "ltr": float, "rs": float}
    printed = 0
    real_voucher_count = 0
    matched_voucher_count = 0

    for v in root.iter("VOUCHER"):
        if len(list(v)) == 0:
            continue
        real_voucher_count += 1
        vdate = (_text(v, "DATE") or "").strip()
        if vdate and (vdate < from_d or vdate > to_d):
            continue
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip()
        if vtype.lower() not in SALES_VOUCHER_TYPE_NAMES:
            continue
        party = (_text(v, "PARTYLEDGERNAME") or "").strip()
        if dealer_names_for_dsr is not None and party not in dealer_names_for_dsr:
            continue
        voucher_dsr = resolve_dealer_to_dsr(party)
        if not voucher_dsr:
            continue

        voucher_ltr = 0.0
        voucher_rs = 0.0
        has_catalog_item = False
        for entry in v.iter("ALLINVENTORYENTRIES.LIST"):
            item_name = (_text(entry, "STOCKITEMNAME") or "").strip()
            raw_billed = _text(entry, "BILLEDQTY")
            raw_actual = _text(entry, "ACTUALQTY")
            qty_num = _num(raw_billed) if raw_billed is not None else _num(raw_actual)
            ltr_guess = _catalog_ltr_guess(item_name, qty_num)
            amount = _num(_text(entry, "AMOUNT"))
            if ltr_guess is not None:
                has_catalog_item = True
                voucher_ltr += ltr_guess
            if amount is not None:
                voucher_rs += abs(amount)
        if not has_catalog_item:
            continue

        matched_voucher_count += 1
        rec = by_dsr.setdefault(voucher_dsr, {"invoices": 0, "ltr": 0.0, "rs": 0.0})
        rec["invoices"] += 1
        rec["ltr"] += voucher_ltr
        rec["rs"] += voucher_rs
        if printed < 40:
            print("[{}] {}  ->  {}   Ltr-guess={}  Rs={}".format(
                vdate, party, voucher_dsr, round(voucher_ltr, 2), round(voucher_rs, 2)))
            printed += 1

    print("\n" + "=" * 70)
    print("Real voucher records in range: {}   Sales-type with a known lubricant item: {}".format(
        real_voucher_count, matched_voucher_count))
    print("\nPer-DSR invoice stats, {} to {}{}:".format(
        from_d, to_d, " (" + target + " only)" if target else ""))
    if not by_dsr:
        print("  (no matching invoices for this period/DSR)")
    for dsr in sorted(by_dsr.keys()):
        rec = by_dsr[dsr]
        avg_ltr = round(rec["ltr"] / rec["invoices"], 2) if rec["invoices"] else 0
        avg_rs = round(rec["rs"] / rec["invoices"], 2) if rec["invoices"] else 0
        print("  {}".format(dsr))
        print("    Invoices: {}   Total Ltr-guess: {}   Total Rs: {}".format(
            rec["invoices"], round(rec["ltr"], 2), round(rec["rs"], 2)))
        print("    Avg Qty/invoice (Ltr): {}   Avg Val/invoice (Rs): {}".format(avg_ltr, avg_rs))
    print("\nSend this whole output back - specifically: (1) does the per-DSR invoice count for a")
    print("month you can check look right against what you'd expect, and (2) do the Avg Qty/Avg")
    print("Val per invoice figures look reasonable (not wildly off from a typical order you know).")
    print("Re-run once per month (pass that month's from/to dates) to build up the full financial")
    print("year the DSR year-month table needs.")


def _build_item_group_resolver():
    """Returns {stock_item_name: immediate_parent_stock_group_name}, for
    the Day Book's "Item Group" column."""
    xml_text = tally_request(STOCK_ITEM_GROUP_COLLECTION_XML)
    root = ET.fromstring(xml_text)
    item_group = {}
    for it in root.iter("STOCKITEM"):
        name = it.get("NAME") or _text(it, "NAME")
        if name:
            item_group[name.strip()] = (_text(it, "PARENT") or "").strip()
    return item_group


# 1 Oct 2026 - the owner's real Tally group master has several stock groups
# whose names start with "petronas" but that are NOT the Petronas lubricants
# business this dashboard tracks. Confirmed via a real --dump-petronas-item-
# groups run (Sep 2026): PETRONAS TATA BRAKE FLUID / PETRONAS TATA CBE GROUP
# / PETRONAS TATA OE GROUP are the owner's separate TATA OE lubricants line
# (AAA distributes both) - together they were inflating Outward Ltr to ~2.6x
# the owner's own Tally "Stock Group Analysis" report for the same period
# (42,281.7 Ltr computed vs 25,934.9 Ltr actual), an excess that matched
# those three groups' own totals almost exactly. Owner confirmed by direct
# instruction ("THIS 3 GROUPS HAS TO BE OMITTED"), then separately added
# PETRONAS MARUTI GROUP to the same list ("THIS GRP ALSO TO BE OMITTED"), no
# further reason given - taken at face value, same as the other three. Any
# future owner-confirmed exclusion should be added here, not spot-patched at
# an individual call site - every live and diagnostic Purchase/Sales Ltr
# computation in this file goes through _is_petronas_item_group() below.
PETRONAS_ITEM_GROUP_EXCLUSIONS_ = set([
    "petronas tata brake fluid",
    "petronas tata cbe group",
    "petronas tata oe group",
    "petronas maruti group",
])


def _is_petronas_item_group(item_grp):
    """True if this stock item's immediate-parent Tally group should count
    toward Petronas Purchase/Sales Ltr: its name starts with "petronas" AND
    it is not one of the explicitly owner-excluded groups above that share
    that naming prefix without actually being Petronas-brand stock."""
    g = (item_grp or "").strip().lower()
    if not g.startswith("petronas"):
        return False
    if g in PETRONAS_ITEM_GROUP_EXCLUSIONS_:
        return False
    return True


# "POS BILLING"/"Sales - CASH" added 10 Oct 2026 - found by comparing a
# real Tally Day Book export against the dashboard's own for the same
# dates (the owner's own request: "compare with ur daybook, tell whats
# dealer missed here"). Neither was ever a missing-DEALER problem - every
# party involved already shows up correctly elsewhere under a recognized
# voucher type - these two voucher types are genuine Petronas counter/
# walk-in sales (booked to party "CASH") that were simply never fetched
# from Tally at all (33 lines / ~390 Ltr in that one sample alone).
# Owner confirmed via AskUserQuestion: add these two, but leave
# "Sales - Suspense" out (a pending/not-yet-reclassified voucher type in
# Tally - it will naturally become a real voucher type once the office
# reclassifies it, same as any other pending entry) - so Suspense is
# deliberately NOT added here.
DAYBOOK_VOUCHER_TYPE_NAMES = ("sales", "sales - petronas", "sales - cbe", "sales - cash", "pos billing")

DAYBOOK_EXCLUDED_PARTY = "tata motors passenger vehicles limited,pune"
DAYBOOK_EXCLUDED_ITEM = "business support services"

DAYBOOK_VOUCHERS_TEMPLATE = """<ENVELOPE>
 <HEADER>
  <VERSION>1</VERSION>
  <TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE>
  <ID>DayBookVouchersDumpCollection</ID>
 </HEADER>
 <BODY>
  <DESC>
   <STATICVARIABLES>
    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
    <SVFROMDATE>{from_date}</SVFROMDATE>
    <SVTODATE>{to_date}</SVTODATE>
   </STATICVARIABLES>
   <TDL>
    <TDLMESSAGE>
     <SYSTEM TYPE="Formulae" NAME="DayBookVouchersOnly">$VoucherTypeName = "Sales" OR $VoucherTypeName = "Sales - PETRONAS" OR $VoucherTypeName = "Sales - CBE" OR $VoucherTypeName = "Sales - CASH" OR $VoucherTypeName = "POS BILLING"</SYSTEM>
     <COLLECTION NAME="DayBookVouchersDumpCollection" ISMODIFY="No">
      <TYPE>Voucher</TYPE>
      <FETCH>VOUCHERNUMBER, DATE, VOUCHERTYPENAME, PARTYLEDGERNAME</FETCH>
      <FETCH>ALLINVENTORYENTRIES.LIST</FETCH>
      <FILTER>DayBookVouchersOnly</FILTER>
     </COLLECTION>
    </TDLMESSAGE>
   </TDL>
  </DESC>
 </BODY>
</ENVELOPE>"""


def dump_daybook(from_date=None, to_date=None, max_lines=60):
    """--dump-daybook [YYYYMMDD-from] [YYYYMMDD-to] - diagnostic only."""
    now = datetime.now()
    from_d = from_date or now.strftime("%Y%m%d")
    to_d = to_date or now.strftime("%Y%m%d")

    print("Resolving Tally's ledger group hierarchy (for Party Group)...")
    ledger_parent, _resolve_dsr = _build_group_resolver()
    print("Resolving Tally's stock item group hierarchy (for Item Group)...")
    item_group = _build_item_group_resolver()
    print("  {} stock item(s) resolved to a group.".format(len(item_group)))

    print("Fetching Sales/Sales-PETRONAS/Sales-CBE vouchers from {} to {} - this can "
          "take a while on a busy company, please wait...".format(from_d, to_d))
    xml_body = DAYBOOK_VOUCHERS_TEMPLATE.format(from_date=from_d, to_date=to_d)
    xml_text = tally_request(xml_body, timeout=600)
    try:
        with open("tally_daybook_dump.xml", "w", encoding="utf-8") as f:
            f.write(xml_text)
        print("(raw response saved to tally_daybook_dump.xml)")
    except Exception as e:
        print("(could not save raw response: {})".format(e))
    root = ET.fromstring(xml_text)

    voucher_count = 0
    in_range_count = 0
    matched_voucher_count = 0
    excluded_party_count = 0
    lines_printed = 0
    lines_skipped = 0
    lines_excluded_item = 0

    for v in root.iter("VOUCHER"):
        if len(list(v)) == 0:
            continue
        voucher_count += 1
        vdate = (_text(v, "DATE") or "").strip()
        if vdate and (vdate < from_d or vdate > to_d):
            continue
        in_range_count += 1
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip()
        if vtype.lower() not in DAYBOOK_VOUCHER_TYPE_NAMES:
            continue
        party = (_text(v, "PARTYLEDGERNAME") or "").strip()
        if party.lower() == DAYBOOK_EXCLUDED_PARTY:
            excluded_party_count += 1
            continue
        matched_voucher_count += 1
        vnum = (_text(v, "VOUCHERNUMBER") or "").strip()
        party_group = ledger_parent.get(party, "")

        entries = list(v.iter("ALLINVENTORYENTRIES.LIST"))
        printed_header = False
        for entry in entries:
            item_name = (_text(entry, "STOCKITEMNAME") or "").strip()
            if not item_name:
                continue
            if item_name.lower() == DAYBOOK_EXCLUDED_ITEM:
                lines_excluded_item += 1
                continue
            item_grp = item_group.get(item_name, "(no group found)")
            raw_actual = _text(entry, "ACTUALQTY")
            raw_billed = _text(entry, "BILLEDQTY")
            rate = _text(entry, "RATE")
            raw_amount = _text(entry, "AMOUNT")
            amt = _num(raw_amount)
            amt_abs = abs(amt) if amt is not None else None
            amt_incl_gst = round(amt_abs * 1.18, 2) if amt_abs is not None else None
            qty_num = _num(raw_billed) if raw_billed is not None else _num(raw_actual)
            ltr_guess = _catalog_ltr_guess(item_name, qty_num)
            alt_qty_native = (_text(entry, "ALTQTY") or _text(entry, "SECONDARYQTY"))

            if lines_printed < max_lines:
                if not printed_header:
                    print("\n[{}] Voucher# {}  {}  ->  Party Group: {}".format(
                        vdate, vnum or "(blank)", party, party_group or "(no group found)"))
                    printed_header = True
                print("    {}  |  Item Group: {}".format(item_name, item_grp))
                print("        ACTUALQTY(raw)={!r}  BILLEDQTY(raw)={!r}  RATE={}".format(
                    raw_actual, raw_billed, rate))
                print("        AMOUNT(raw)={}  AMOUNT*1.18(candidate incl-GST)={}".format(
                    amt_abs, amt_incl_gst))
                print("        Alt Qty - Tally-native attempt={}  |  catalog-based guess={}".format(
                    alt_qty_native if alt_qty_native is not None else "n/a (field didn't resolve)",
                    ltr_guess if ltr_guess is not None else "n/a (not a known catalog item)"))
                lines_printed += 1
            else:
                lines_skipped += 1

    print("\n" + "=" * 70)
    print("Real voucher records in requested date range ({} to {}): {}".format(from_d, to_d, in_range_count))
    print("Matched Sales/Sales-PETRONAS/Sales-CBE vouchers: {}  ({} excluded as {!r})".format(
        matched_voucher_count, excluded_party_count, DAYBOOK_EXCLUDED_PARTY))
    if lines_excluded_item:
        print("{} inventory line(s) excluded as item {!r}.".format(lines_excluded_item, DAYBOOK_EXCLUDED_ITEM))
    if lines_skipped:
        print("({} more inventory line(s) not printed - narrow the date range to see them.)".format(lines_skipped))
    print("\nSend this whole output back, plus your real Sep 1-2 Day Book screenshot side by")
    print("side, and confirm: (1) does 'Alt Qty - Tally-native attempt' show real numbers, or")
    print("'n/a' every time (if n/a, the catalog-based guess is what the real feature will use")
    print("instead); (2) does AMOUNT(raw) or AMOUNT*1.18 match your screenshot's own Amount")
    print("column (tells us whether raw AMOUNT already includes GST or not); (3) do the Party")
    print("Group / Item Group values look right for a few rows you recognize.")


# Matches the LTR figure in either order a raw ACTUALQTY/BILLEDQTY string
# can show it: '<nos> NOS = <ltr> LTR' (most items - Ltr is an alternate
# unit of a Nos-based item) OR '<ltr> LTR = <nos> NOS' (some items - e.g.
# the TATA CBE/OE group items - have Ltr AS their primary unit, with Nos
# as the alternate/compound reading instead). Originally only matched the
# first order (anchored on '=\s*' immediately before the number); FIXED
# 30 Sep 2026 after a real --dump-petronas-purchase run showed the second
# order is common for a whole item group ("Purchase - CBE" native-Ltr
# coverage was 6/31 under the old anchored regex, almost all TATA CBE/OE
# items in the reversed order) - now just matches whichever number sits
# immediately before the literal "LTR", regardless of which side of "="
# it's on, since that's the only distinguishing anchor that holds in both
# orders. A raw string with no Ltr reading at all (plain '10 Nos') still
# correctly returns None either way.
_NATIVE_LTR_RE = re.compile(r"([\d,]+\.?\d*)\s*LTR", re.IGNORECASE)


def _parse_native_ltr(raw_text):
    """Pulls the LTR figure Tally itself already embeds in a lubricant
    item's raw ACTUALQTY/BILLEDQTY text, in whichever order it appears
    (see _NATIVE_LTR_RE's comment above). Returns None (not 0.0) when no
    LTR reading is present at all, so that's distinguishable from a
    genuine zero."""
    if not raw_text:
        return None
    m = _NATIVE_LTR_RE.search(raw_text)
    if not m:
        return None
    try:
        return float(m.group(1).replace(",", ""))
    except ValueError:
        return None


def _daybook_row_key(voucher_number, line_seq):
    """voucher_number + this line's 0-based position among that
    voucher's inventory lines, so two lines for the same item on one
    voucher get distinct, stable keys across re-syncs."""
    return "{}#{}".format(voucher_number or "(blank)", line_seq)


def fetch_daybook(from_date, to_date):
    """Real (non-diagnostic) Day Book fetch for [from_date, to_date].
    Returns a flat list of row dicts, one per matched inventory line:
      {monthKey, rowKey, date, voucherNumber, voucherType, party,
       partyGroup, itemName, itemGroup, qty, altQtyLtr, rate, amount}
    Amount is always GST-inclusive (owner confirmed); altQtyLtr is
    parsed from Tally's own embedded "= X.XXX LTR" text, not the
    catalog guess (see the module comment above dump_daybook())."""
    ledger_parent, _resolve_dsr = _build_group_resolver()
    item_group = _build_item_group_resolver()

    xml_body = DAYBOOK_VOUCHERS_TEMPLATE.format(from_date=from_date, to_date=to_date)
    xml_text = tally_request(xml_body, timeout=600)
    root = ET.fromstring(xml_text)

    rows = []
    for v in root.iter("VOUCHER"):
        if len(list(v)) == 0:
            continue
        vdate = (_text(v, "DATE") or "").strip()
        if vdate and (vdate < from_date or vdate > to_date):
            continue
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip()
        if vtype.lower() not in DAYBOOK_VOUCHER_TYPE_NAMES:
            continue
        party = (_text(v, "PARTYLEDGERNAME") or "").strip()
        if party.lower() == DAYBOOK_EXCLUDED_PARTY:
            continue
        vnum = (_text(v, "VOUCHERNUMBER") or "").strip()
        party_group = ledger_parent.get(party, "")
        vdate_iso = "{}-{}-{}".format(vdate[0:4], vdate[4:6], vdate[6:8]) if len(vdate) == 8 else vdate
        month_key = vdate_iso[:7] if len(vdate_iso) >= 7 else ""

        entries = list(v.iter("ALLINVENTORYENTRIES.LIST"))
        for line_seq, entry in enumerate(entries):
            item_name = (_text(entry, "STOCKITEMNAME") or "").strip()
            if not item_name:
                continue
            if item_name.lower() == DAYBOOK_EXCLUDED_ITEM:
                continue
            item_grp = item_group.get(item_name, "")
            raw_actual = _text(entry, "ACTUALQTY")
            raw_billed = _text(entry, "BILLEDQTY")
            rate = _num(_text(entry, "RATE"))
            amt = _num(_text(entry, "AMOUNT"))
            amt_abs = abs(amt) if amt is not None else None
            amount_incl_gst = round(amt_abs * 1.18, 2) if amt_abs is not None else None
            qty_num = _num(raw_billed) if raw_billed is not None else _num(raw_actual)
            alt_qty_ltr = _parse_native_ltr(raw_billed)
            if alt_qty_ltr is None:
                alt_qty_ltr = _parse_native_ltr(raw_actual)

            if not month_key:
                continue
            rows.append({
                "monthKey": month_key,
                "rowKey": _daybook_row_key(vnum, line_seq),
                "date": vdate_iso,
                "voucherNumber": vnum,
                "voucherType": vtype,
                "party": party,
                "partyGroup": party_group,
                "itemName": item_name,
                "itemGroup": item_grp,
                "qty": qty_num,
                "altQtyLtr": alt_qty_ltr,
                "rate": rate,
                "amount": amount_incl_gst,
            })
    return rows


DAYBOOK_BACKFILL_START_DATE = "20260901"
DAYBOOK_LOOKBACK_DAYS = 3
DAYBOOK_POLL_INTERVAL_SECONDS = 60 * 60


def maybe_push_daybook(state, dry_run, ts):
    if dry_run:
        return state
    last = state.get("last_daybook_push_ts", 0)
    if time.time() - last < DAYBOOK_POLL_INTERVAL_SECONDS:
        return state

    today = datetime.now()
    if not state.get("daybook_backfill_done"):
        from_date = DAYBOOK_BACKFILL_START_DATE
        print("[{}] Day Book: first run - backfilling {} through today ({}) - this can "
              "take a while on a busy company, please wait...".format(
                  ts, from_date, today.strftime("%Y%m%d")))
    else:
        from_date = (today - timedelta(days=DAYBOOK_LOOKBACK_DAYS)).strftime("%Y%m%d")
        print("[{}] computing Day Book ({} to today, {}-day lookback for late entries - "
              "runs every {} min, not every poll)...".format(
                  ts, from_date, DAYBOOK_LOOKBACK_DAYS, DAYBOOK_POLL_INTERVAL_SECONDS // 60))
    to_date = today.strftime("%Y%m%d")

    try:
        rows = fetch_daybook(from_date, to_date)
    except Exception as e:
        print("[{}] Day Book computation failed this round (will retry at the next "
              "poll): {}".format(ts, e))
        return state

    if rows:
        sent = post_backend_chunked("liveDaybookUpdate", "rows", rows, ts)
        print("[{}] Day Book: sent {} row(s) total ({} to {})".format(ts, sent, from_date, to_date))
    else:
        print("[{}] Day Book: no matching rows in range ({} to {})".format(ts, from_date, to_date))

    state["daybook_backfill_done"] = True
    state["last_daybook_push_ts"] = time.time()
    return state


# -- Real, non-diagnostic live Sales (dashboard only), 18 Sep 2026 -------
# VOUCHERTYPENAME "Sales"/"Sales - PETRONAS"; quantity in individual NOS;
# Tally's own ACTUALQTY/BILLEDQTY text carries a native "= X.XXX LTR"
# conversion matching _catalog_ltr_guess()'s formula. This Tally company
# also runs an unrelated auto/tractor-spares business through the same
# company file, so fetch_sales() scopes dealer/item totals to only
# entries that match a known lubricant item AND resolve to a known DSR
# group.
def fetch_sales():
    """Returns (dsr_totals, dealer_totals, item_totals, daily_totals,
    daily_rs_totals, dsr_item_totals, daily_totals_by_dsr) - "today"/"month"
    Ltr for each, mirroring fetch_collection()'s shape (which now also
    returns a 4th daily_totals_by_dsr - added 11 Oct 2026 for the per-DSR
    Sales/Collection heatmap, see its own comment near fetch_collection()'s
    daily_totals_by_dsr init). dealer_totals also carries
    lastSaleDate (yyyymmdd or None), lastSaleLtr, avg6moLtr/avg6moRs,
    lastMonthSameDate/lastMonthFull (Ltr+Rs), lastSaleBeforeMonth."""
    now = datetime.now()
    from_d = now.strftime("%Y%m01")
    to_d = now.strftime("%Y%m%d")
    today_str = now.strftime("%Y%m%d")
    LOOKBACK_DAYS = 120
    lookback_d = (now - timedelta(days=LOOKBACK_DAYS)).strftime("%Y%m%d")
    prev_from_d, prev_to_d = _prev_month_range(now)
    prev_samedate_from_d, prev_samedate_to_d = _prev_month_samedate_range(now)
    sixmo_from_d, sixmo_to_d = _sixmo_window_range(now)
    widest_lookback_d = min(lookback_d, sixmo_from_d)

    ledger_parent, resolve = _build_group_resolver()
    ledger_to_dsr = {}
    for lname, lparent in ledger_parent.items():
        if not lparent:
            continue
        dsr = resolve(lparent)
        if dsr:
            ledger_to_dsr[lname] = dsr

    xml_body = SALES_VOUCHERS_TEMPLATE.format(from_date=from_d, to_date=to_d)
    xml_text = tally_request(xml_body, timeout=600)
    root = ET.fromstring(xml_text)

    dsr_totals = dict((dsr, {"today": 0.0, "month": 0.0, "lastMonthSameDate": 0.0, "lastMonthFull": 0.0, "todayRs": 0.0, "monthRs": 0.0}) for dsr in KNOWN_DSR_GROUP_NAMES)
    dealer_totals = {}
    item_totals = {}
    dsr_item_totals = {}
    dealer_sixmo_totals = {}
    dealer_sixmo_rs_totals = {}
    dsr_sixmo_totals = {}
    dsr_sixmo_rs_totals = {}
    last_sale_date_by_dealer = {}
    last_sale_before_month_by_dealer = {}
    daily_totals = {}
    daily_rs_totals = {}
    # Per-DSR daily trend (Ltr), added 11 Oct 2026 - Rs-value twin of this
    # same idea lives in fetch_collection()'s daily_totals_by_dsr, see its
    # own comment for the full reasoning (new heatmap feature, no new
    # Tally request needed - `dsr`/`iso_date`/`ltr` below are already
    # computed for dsr_totals/daily_totals).
    daily_totals_by_dsr = {}
    voucher_count = 0

    for v in root.iter("VOUCHER"):
        if len(list(v)) == 0:
            continue
        vdate = (_text(v, "DATE") or "").strip()
        if not vdate or vdate < widest_lookback_d:
            continue
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip()
        if vtype.lower() not in SALES_VOUCHER_TYPE_NAMES:
            continue
        party = (_text(v, "PARTYLEDGERNAME") or "").strip()
        dsr = ledger_to_dsr.get(party)
        if not dsr:
            continue
        is_current_month = from_d <= vdate <= to_d
        is_prev_month = prev_from_d <= vdate <= prev_to_d
        is_prev_month_samedate = prev_samedate_from_d <= vdate <= prev_samedate_to_d
        is_sixmo_window = sixmo_from_d <= vdate <= sixmo_to_d
        is_today = vdate == today_str

        voucher_has_match = False
        for entry in v.iter("ALLINVENTORYENTRIES.LIST"):
            item_name = (_text(entry, "STOCKITEMNAME") or "").strip()
            if not item_name:
                continue
            raw_billed = _text(entry, "BILLEDQTY")
            raw_actual = _text(entry, "ACTUALQTY")
            qty_num = _num(raw_billed) if raw_billed is not None else _num(raw_actual)
            ltr = _catalog_ltr_guess(item_name, qty_num)
            if ltr is None:
                continue
            voucher_has_match = True
            raw_amount = _text(entry, "AMOUNT")
            amt = _num(raw_amount) if raw_amount is not None else None
            dealer_rec = dealer_totals.setdefault(party, {
                "today": 0.0, "month": 0.0, "todayRs": 0.0, "monthRs": 0.0,
                "lastMonthSameDate": 0.0, "lastMonthFull": 0.0,
                "lastMonthSameDateRs": 0.0, "lastMonthFullRs": 0.0,
            })
            if is_sixmo_window:
                dealer_sixmo_totals[party] = dealer_sixmo_totals.get(party, 0.0) + ltr
                if amt is not None:
                    dealer_sixmo_rs_totals[party] = dealer_sixmo_rs_totals.get(party, 0.0) + abs(amt)
                dsr_sixmo_totals[dsr] = dsr_sixmo_totals.get(dsr, 0.0) + ltr
                if amt is not None:
                    dsr_sixmo_rs_totals[dsr] = dsr_sixmo_rs_totals.get(dsr, 0.0) + abs(amt)
            if is_current_month or is_prev_month:
                iso_date = vdate[0:4] + "-" + vdate[4:6] + "-" + vdate[6:8]
                daily_totals[iso_date] = daily_totals.get(iso_date, 0.0) + ltr
                if amt is not None:
                    daily_rs_totals[iso_date] = daily_rs_totals.get(iso_date, 0.0) + abs(amt)
                _dsr_day = daily_totals_by_dsr.setdefault(dsr, {})
                _dsr_day[iso_date] = _dsr_day.get(iso_date, 0.0) + ltr
            if is_prev_month_samedate:
                dsr_totals[dsr]["lastMonthSameDate"] += ltr
                dealer_rec["lastMonthSameDate"] += ltr
                if amt is not None:
                    dealer_rec["lastMonthSameDateRs"] += abs(amt)
            if is_prev_month:
                dsr_totals[dsr]["lastMonthFull"] += ltr
                dealer_rec["lastMonthFull"] += ltr
                if amt is not None:
                    dealer_rec["lastMonthFullRs"] += abs(amt)
            if is_current_month:
                voucher_count += 1
                dsr_totals[dsr]["month"] += ltr
                if amt is not None:
                    dsr_totals[dsr]["monthRs"] += abs(amt)
                dealer_rec["month"] += ltr
                if amt is not None:
                    dealer_rec["monthRs"] += abs(amt)
                item_rec = item_totals.setdefault(item_name, {"today": 0.0, "month": 0.0})
                item_rec["month"] += ltr
                dsr_item_key = dsr + "||" + item_name
                dsr_item_totals[dsr_item_key] = dsr_item_totals.get(dsr_item_key, 0.0) + ltr
                if is_today:
                    dsr_totals[dsr]["today"] += ltr
                    if amt is not None:
                        dsr_totals[dsr]["todayRs"] += abs(amt)
                    dealer_rec["today"] += ltr
                    if amt is not None:
                        dealer_rec["todayRs"] += abs(amt)
                    item_rec["today"] += ltr

        if voucher_has_match:
            prev = last_sale_date_by_dealer.get(party)
            if not prev or vdate > prev:
                last_sale_date_by_dealer[party] = vdate
            if not is_current_month:
                prevb = last_sale_before_month_by_dealer.get(party)
                if not prevb or vdate > prevb:
                    last_sale_before_month_by_dealer[party] = vdate

    print("  (sales: {} lubricant-item line(s) counted toward DSR/dealer/item totals "
          "this fetch, last-sale-date tracked for {} dealer(s) within {} days)".format(
              voucher_count, len(last_sale_date_by_dealer), LOOKBACK_DAYS))

    last_sale_ltr_by_dealer = {}
    for v in root.iter("VOUCHER"):
        if len(list(v)) == 0:
            continue
        vdate = (_text(v, "DATE") or "").strip()
        if not vdate:
            continue
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip()
        if vtype.lower() not in SALES_VOUCHER_TYPE_NAMES:
            continue
        party = (_text(v, "PARTYLEDGERNAME") or "").strip()
        if not party:
            continue
        target_date = last_sale_date_by_dealer.get(party)
        if not target_date or vdate != target_date:
            continue
        for entry in v.iter("ALLINVENTORYENTRIES.LIST"):
            item_name = (_text(entry, "STOCKITEMNAME") or "").strip()
            if not item_name:
                continue
            raw_billed = _text(entry, "BILLEDQTY")
            raw_actual = _text(entry, "ACTUALQTY")
            qty_num = _num(raw_billed) if raw_billed is not None else _num(raw_actual)
            ltr = _catalog_ltr_guess(item_name, qty_num)
            if ltr is None:
                continue
            last_sale_ltr_by_dealer[party] = last_sale_ltr_by_dealer.get(party, 0.0) + ltr

    for _dsr, _sixmo_total in dsr_sixmo_totals.items():
        dsr_totals[_dsr]["avg6moLtr"] = round(_sixmo_total / 6.0, 2)
    for _dsr, _sixmo_rs_total in dsr_sixmo_rs_totals.items():
        dsr_totals[_dsr]["avg6moRs"] = round(_sixmo_rs_total / 6.0, 2)

    _merge_saravanan_into_chellamani(dsr_totals)

    _sarv_item_prefix = _SARAVANAN_GROUP_NAME + "||"
    _chell_item_prefix = _AJITH_GROUP_NAME + "||"
    for _key in list(dsr_item_totals.keys()):
        if _key.startswith(_sarv_item_prefix):
            _item_name = _key[len(_sarv_item_prefix):]
            _chell_key = _chell_item_prefix + _item_name
            dsr_item_totals[_chell_key] = dsr_item_totals.get(_chell_key, 0.0) + dsr_item_totals.pop(_key)

    _DEALER_REC_DEFAULT = lambda: {
        "today": 0.0, "month": 0.0, "todayRs": 0.0, "monthRs": 0.0,
        "lastMonthSameDate": 0.0, "lastMonthFull": 0.0,
        "lastMonthSameDateRs": 0.0, "lastMonthFullRs": 0.0,
    }
    all_dealer_names = set(dealer_totals.keys()) | set(last_sale_date_by_dealer.keys()) | set(last_sale_before_month_by_dealer.keys())
    for name in all_dealer_names:
        rec = dealer_totals.setdefault(name, _DEALER_REC_DEFAULT())
        rec["lastSaleDate"] = last_sale_date_by_dealer.get(name)
        rec["lastSaleBeforeMonth"] = last_sale_before_month_by_dealer.get(name)
        rec["lastSaleLtr"] = last_sale_ltr_by_dealer.get(name)

    for name, sixmo_total in dealer_sixmo_totals.items():
        rec = dealer_totals.setdefault(name, _DEALER_REC_DEFAULT())
        rec["avg6moLtr"] = round(sixmo_total / 6.0, 2)

    for name, sixmo_rs_total in dealer_sixmo_rs_totals.items():
        rec = dealer_totals.setdefault(name, _DEALER_REC_DEFAULT())
        rec["avg6moRs"] = round(sixmo_rs_total / 6.0, 2)

    # daily_totals_by_dsr is dsr -> {date: Ltr} - flat but date-keyed, same
    # situation as fetch_collection()'s own daily_totals_by_dsr, so the
    # dedicated flat merge is used here too (see
    # _merge_saravanan_into_chellamani_flat()'s own comment for why the
    # other two merge helpers are both wrong for this shape).
    _merge_saravanan_into_chellamani_flat(daily_totals_by_dsr)

    return dsr_totals, dealer_totals, item_totals, daily_totals, daily_rs_totals, dsr_item_totals, daily_totals_by_dsr


def diff_daily_trend(prev, current, value_key):
    changed = []
    for date, val in current.items():
        if prev.get(date) != val:
            changed.append({"date": date, value_key: val})
    return changed


# Added 11 Oct 2026 for the per-DSR Sales/Collection heatmap - same idea as
# diff_daily_trend() above, one level deeper (prev/current are dsr ->
# {date: val} dicts, see fetch_sales()/fetch_collection()'s own
# daily_totals_by_dsr comments). Only emits rows for a (dsr, date)
# combination whose value actually changed, same minimal-diff convention
# every other push in this file uses.
def diff_daily_trend_by_dsr(prev, current, value_key):
    changed = []
    for dsr, day_map in current.items():
        prev_day_map = prev.get(dsr, {})
        for date, val in day_map.items():
            if prev_day_map.get(date) != val:
                changed.append({"dsrName": dsr, "date": date, value_key: val})
    return changed


def diff_sales(prev, current):
    changed = []
    for dsr, rec in current.items():
        old = prev.get(dsr)
        if old != rec:
            changed.append({
                "dsrName": dsr,
                "todaySoldLtr": rec["today"],
                "monthSoldLtr": rec["month"],
                "lastMonthSameDateLtr": rec.get("lastMonthSameDate"),
                "lastMonthFullLtr": rec.get("lastMonthFull"),
                "todaySoldRs": rec.get("todayRs"),
                "monthSoldRs": rec.get("monthRs"),
                "avg6moLtr": rec.get("avg6moLtr"),
                "avg6moRs": rec.get("avg6moRs"),
            })
    return changed


def diff_sales_by_dealer(prev, current):
    changed = []
    for dealer, rec in current.items():
        old = prev.get(dealer)
        if old != rec:
            changed.append({
                "name": dealer,
                "todaySoldLtr": rec["today"],
                "monthSoldLtr": rec["month"],
                "lastSaleDate": rec.get("lastSaleDate"),
                "lastSaleLtr": rec.get("lastSaleLtr"),
                "avg6moLtr": rec.get("avg6moLtr"),
                "lastSaleBeforeMonth": rec.get("lastSaleBeforeMonth"),
                "lastMonthSameDateLtr": rec.get("lastMonthSameDate"),
                "lastMonthFullLtr": rec.get("lastMonthFull"),
                "todaySoldRs": rec.get("todayRs"),
                "monthSoldRs": rec.get("monthRs"),
                "lastMonthSameDateRs": rec.get("lastMonthSameDateRs"),
                "lastMonthFullRs": rec.get("lastMonthFullRs"),
                "avg6moRs": rec.get("avg6moRs"),
            })
    return changed


def diff_sales_by_dsr_item(prev, current):
    # Fixed 1 Oct 2026 - was `for key, month_ltr in current.items()`, which
    # only ever visits items that have SOME current-month Ltr. fetch_sales()
    # builds dsr_item_totals fresh from scratch every poll (only items with
    # an actual current-month voucher line get an entry at all - there's no
    # equivalent of dsr_totals' pre-seeded-with-every-known-DSR backstop, or
    # dealer_totals' all_dealer_names backstop, for this dict), so at a
    # calendar-month rollover, an item that sold last month but hasn't sold
    # yet this month simply has NO key in `current` at all. The old code
    # never visited that key, so no "reset to 0" update was ever sent - the
    # Sheet (and the dashboard's "Items Covered"/"Top selling SKU by DSR"
    # cards, which read it straight back) kept showing last month's stale
    # Ltr figure indefinitely, for every item not yet sold again this month.
    # Iterating the UNION of prev's and current's keys (defaulting a missing
    # current entry to 0.0) means a dropped-to-zero item gets pushed exactly
    # once, the first poll after it drops - matching how dsr_totals/
    # dealer_totals already self-heal across a month rollover.
    changed = []
    for key in set(prev.keys()) | set(current.keys()):
        month_ltr = current.get(key, 0.0)
        if prev.get(key) != month_ltr:
            dsr_name, item_name = key.split("||", 1)
            changed.append({"dsrName": dsr_name, "itemName": item_name, "monthSoldLtr": month_ltr})
    return changed


def diff_sales_by_item(prev, current):
    # Same fix, same date, same reason as diff_sales_by_dsr_item() above -
    # item_totals has the identical "only items with a current-month
    # voucher line get a dict entry at all" shape, so this had the exact
    # same stale-after-month-rollover bug for the plain item-level map
    # (feeds the homescreen "Items sold" panel and petronasCoverage_()'s
    # own company-wide "Items Covered" count on the dashboard).
    changed = []
    _default_rec = {"today": 0.0, "month": 0.0}
    for item in set(prev.keys()) | set(current.keys()):
        rec = current.get(item, _default_rec)
        old = prev.get(item)
        if old != rec:
            changed.append({"name": item, "todaySoldLtr": rec["today"], "monthSoldLtr": rec["month"]})
    return changed


SALES_POLL_INTERVAL_SECONDS = 60 * 60


def maybe_push_sales(state, dry_run, ts):
    if dry_run:
        return state
    last = state.get("last_sales_push_ts", 0)
    if time.time() - last < SALES_POLL_INTERVAL_SECONDS:
        return state
    print("[{}] computing DSR sales (Sales vouchers, this month so far - "
          "runs every {} min, not every poll)...".format(ts, SALES_POLL_INTERVAL_SECONDS // 60))
    try:
        dsr_sales, dealer_sales, item_sales, daily_sales, daily_rs_sales, dsr_item_sales, daily_sales_by_dsr = fetch_sales()
    except Exception as e:
        print("[{}] sales computation failed this round (will retry at the next "
              "poll): {}".format(ts, e))
        return state

    changed = diff_sales(state.get("sales", {}), dsr_sales)
    if changed:
        sent = post_backend_chunked("liveSalesUpdate", "sales", changed, ts)
        print("[{}] sales: sent {} changed DSR row(s) total".format(ts, sent))
    else:
        print("[{}] sales: no changes".format(ts))

    dealer_changed = diff_sales_by_dealer(state.get("salesByDealer", {}), dealer_sales)
    if dealer_changed:
        dealer_sent = post_backend_chunked("liveSalesByDealerUpdate", "sales", dealer_changed, ts)
        print("[{}] sales (by dealer): sent {} changed dealer row(s) total".format(ts, dealer_sent))
    else:
        print("[{}] sales (by dealer): no changes".format(ts))

    item_changed = diff_sales_by_item(state.get("salesByItem", {}), item_sales)
    if item_changed:
        item_sent = post_backend_chunked("liveSalesByItemUpdate", "sales", item_changed, ts)
        print("[{}] sales (by item): sent {} changed item row(s) total".format(ts, item_sent))
    else:
        print("[{}] sales (by item): no changes".format(ts))

    dsr_item_changed = diff_sales_by_dsr_item(state.get("salesByDsrItem", {}), dsr_item_sales)
    if dsr_item_changed:
        dsr_item_sent = post_backend_chunked("liveSalesByDsrItemUpdate", "sales", dsr_item_changed, ts)
        print("[{}] sales (by DSR+item): sent {} changed row(s) total".format(ts, dsr_item_sent))
    else:
        print("[{}] sales (by DSR+item): no changes".format(ts))

    trend_changed = diff_daily_trend(state.get("salesTrendDaily", {}), daily_sales, "ltr")
    if trend_changed:
        trend_sent = post_backend_chunked("liveSalesTrendUpdate", "trend", trend_changed, ts)
        print("[{}] sales (daily trend): sent {} changed day(s) total".format(ts, trend_sent))
    else:
        print("[{}] sales (daily trend): no changes".format(ts))

    rs_trend_changed = diff_daily_trend(state.get("salesRsTrendDaily", {}), daily_rs_sales, "rs")
    if rs_trend_changed:
        rs_trend_sent = post_backend_chunked("liveSalesRsTrendUpdate", "trend", rs_trend_changed, ts)
        print("[{}] sales (daily Rs trend): sent {} changed day(s) total".format(ts, rs_trend_sent))
    else:
        print("[{}] sales (daily Rs trend): no changes".format(ts))

    # Per-DSR daily trend (Ltr), added 11 Oct 2026 for the new Sales/
    # Collection heatmap - own diff/push/state key, same pattern as the
    # company-wide trend just above and as collection's own DSR trend
    # push in maybe_push_collection().
    dsr_trend_changed = diff_daily_trend_by_dsr(state.get("salesTrendDailyByDsr", {}), daily_sales_by_dsr, "ltr")
    if dsr_trend_changed:
        dsr_trend_sent = post_backend_chunked("liveSalesTrendByDsrUpdate", "trend", dsr_trend_changed, ts)
        print("[{}] sales (daily trend by DSR): sent {} changed row(s) total".format(ts, dsr_trend_sent))
    else:
        print("[{}] sales (daily trend by DSR): no changes".format(ts))

    state["sales"] = dsr_sales
    state["salesByDealer"] = dealer_sales
    state["salesByItem"] = item_sales
    state["salesByDsrItem"] = dsr_item_sales
    state["salesTrendDaily"] = daily_sales
    state["salesRsTrendDaily"] = daily_rs_sales
    state["salesTrendDailyByDsr"] = daily_sales_by_dsr
    state["last_sales_push_ts"] = time.time()
    return state


# -- Collection Group Summary diagnostic, added 28 Sep 2026 --------------
# New feature: a dashboard panel mirroring the owner's Tally "Group
# Summary - COLLECTION - My View" screenshot (ALAGU AUTO AGENCY, PETRONAS
# TOTAL AREA SALES, 1-Sep-26 to 30-Sep-26) - 8 raw Tally party GROUPS
# (Opening Balance / Debit / Credit / Closing Balance, with a Grand Total
# row), each expandable to its individual dealers, downloadable to Excel
# for the current month, with every month's data kept in the sheet.
#
# These 8 "sales area" groups are a DIFFERENT branch of Tally's Group
# hierarchy from KNOWN_DSR_GROUP_NAMES above (e.g. "PETRONAS MADURAI
# TOTAL AREA" here vs. "PETRONAS MDU AREA (CHELLAMANI)" there) - kept as
# their own separate constant and their own separate resolver below, per
# this file's standing "never risk an already-trusted code path" rule
# (_build_group_resolver() stays untouched, used only for DSR
# resolution).
#
# Genuinely unconfirmed and the reason this ships diagnostic-only first:
# a stock Tally "Group Summary" report may show only Closing Balance by
# default - whether Opening/Debit/Credit/Closing all come back the way
# the owner's screenshot shows (his screenshot says "My View", which may
# be a customised view rather than the report's own default shape) is
# unknown until real output is seen. This diagnostic (a) confirms the 8
# group names actually exist in the Group master and prints each one's
# immediate PARENT, (b) enumerates every ledger/dealer that resolves to
# each of the 8 groups, and (c) attempts a native report export
# (TALLYREQUEST="Export Data" + REPORTNAME="Group Summary", mirroring
# BILLS_RECEIVABLE_REPORT_XML's already-proven shape, with SVFROMDATE/
# SVTODATE set) and saves/prints the raw XML so the real column shape can
# be read off real Tally output - nothing here is guessed into a live
# fetch_group_summary() yet.
# 28 Sep 2026: owner confirmed real Tally name is "PETRONAS  LEFT OUT
# DEALERS" (two spaces) but explicitly does NOT want this group included
# in the Collection Group Summary feature at all ("this is left out
# dealers dont want this group") - so it's deliberately left OUT of this
# list. Only the 7 real sales-area groups below are in scope.
KNOWN_COLLECTION_GROUP_NAMES = [
    "PETRONAS AAA DIRECT AREA SALES",
    "PETRONAS MADURAI TOTAL AREA",
    "PETRONAS OE AM WS SALES",
    "PETRONAS OTHER AREA",
    "PETRONAS RAMNAD AREA (OTHER)",
    "PETRONAS RAMNAD AREA (PIONEER)",
    "PETRONAS TIRUNELVELI TOTAL AREA",
]

# Same "Export Data" + REPORTNAME native-report shape already proven
# working for BILLS_RECEIVABLE_REPORT_XML - REPORTNAME guessed as "Group
# Summary" (Tally's own standard report name for this screen, Display
# More Reports -> Statement of Accounts -> Group Summary). Not yet
# confirmed this report name is accepted by this Tally version/setup, or
# that it can be scoped to just these 8 groups via STATICVARIABLES alone
# (a plain Group Summary export may come back for EVERY group in the
# company, in which case the diagnostic below filters client-side to the
# 8 names, same "ask Tally for everything, filter here" pattern already
# used for Collection/Sales vouchers elsewhere in this file).
GROUP_SUMMARY_REPORT_XML = """<ENVELOPE>
 <HEADER>
  <TALLYREQUEST>Export Data</TALLYREQUEST>
 </HEADER>
 <BODY>
  <EXPORTDATA>
   <REQUESTDESC>
    <REPORTNAME>Group Summary</REPORTNAME>
    <STATICVARIABLES>
     <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
     <SVFROMDATE>{from_date}</SVFROMDATE>
     <SVTODATE>{to_date}</SVTODATE>
    </STATICVARIABLES>
   </REQUESTDESC>
  </EXPORTDATA>
 </BODY>
</ENVELOPE>"""

# Added for STEP 4 below, after the real 28 Sep 2026 diagnostic run showed
# GROUP_SUMMARY_REPORT_XML above does NOT give per-custom-group Opening/
# Debit/Credit/Closing (only the company's standard primary groups, each
# with a single Closing Dr/Cr figure). This is the same "ask Tally for
# every raw voucher, compute the rest client-side" approach already
# proven and trusted for fetch_collection()/fetch_sales() - deliberately
# UNFILTERED (no <FILTER> tag, unlike COLLECTION_VOUCHERS_TEMPLATE's
# ReceiptVouchersOnly formula) because a real Group Summary's Debit/
# Credit movement is driven by EVERY voucher type touching a dealer's
# ledger (Sales, Receipt, Credit Note, Journal, etc.), not just Receipts.
# Flagged as the heaviest request in this file so far - whole company,
# every voucher type, unfiltered - kept diagnostic-only until its output
# is confirmed against the owner's real screenshot numbers.
ALL_VOUCHERS_FOR_GROUP_SUMMARY_XML = """<ENVELOPE>
 <HEADER>
  <VERSION>1</VERSION>
  <TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE>
  <ID>GroupSummaryAllVouchersCollection</ID>
 </HEADER>
 <BODY>
  <DESC>
   <STATICVARIABLES>
    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
    <SVFROMDATE>{from_date}</SVFROMDATE>
    <SVTODATE>{to_date}</SVTODATE>
   </STATICVARIABLES>
   <TDL>
    <TDLMESSAGE>
     <COLLECTION NAME="GroupSummaryAllVouchersCollection" ISMODIFY="No">
      <TYPE>Voucher</TYPE>
      <FETCH>DATE, VOUCHERTYPENAME, PARTYLEDGERNAME</FETCH>
      <FETCH>ALLLEDGERENTRIES.LIST</FETCH>
     </COLLECTION>
    </TDLMESSAGE>
   </TDL>
  </DESC>
 </BODY>
</ENVELOPE>"""


def _build_collection_group_resolver():
    """Same ledger->group->parent chain-walk as _build_group_resolver(),
    but resolving against KNOWN_COLLECTION_GROUP_NAMES instead of
    KNOWN_DSR_GROUP_NAMES - its own isolated fetch, deliberately not
    shared with the DSR resolver, same reasoning given throughout this
    file for every new, not-yet-verified resolver."""
    ledger_xml = tally_request(LEDGER_GROUP_COLLECTION_XML)
    ledger_root = ET.fromstring(ledger_xml)
    ledger_parent = {}
    for led in ledger_root.iter("LEDGER"):
        name = led.get("NAME") or _text(led, "NAME")
        if name:
            ledger_parent[name.strip()] = (_text(led, "PARENT") or "").strip()

    group_xml = tally_request(GROUP_HIERARCHY_COLLECTION_XML)
    group_root = ET.fromstring(group_xml)
    group_parent = {}
    for grp in group_root.iter("GROUP"):
        gname = grp.get("NAME") or _text(grp, "NAME")
        if gname:
            group_parent[gname.strip()] = (_text(grp, "PARENT") or "").strip()

    known_set = set(KNOWN_COLLECTION_GROUP_NAMES)

    def resolve(immediate_group):
        current = immediate_group
        seen = set()
        while current and current not in seen:
            seen.add(current)
            if current in known_set:
                return current
            current = group_parent.get(current, "")
        return None

    return ledger_parent, group_parent, resolve


def dump_group_summary(from_date=None, to_date=None):
    """
    --dump-group-summary [YYYYMMDD-from] [YYYYMMDD-to]

    Diagnostic only - pushes nothing anywhere, no fetch_group_summary()
    exists yet. Defaults to the current calendar month-to-date if no
    dates given, matching the owner's own screenshot's whole-month range.

    Step 1: confirms each of the 8 KNOWN_COLLECTION_GROUP_NAMES actually
    exists in this Tally's Group master and prints its immediate PARENT
    (so a typo'd/renamed group is caught immediately, before anything
    else is even attempted).

    Step 2: enumerates every ledger that resolves to each of the 8
    groups via the group-chain walk (same pattern already proven for DSR
    resolution), and prints a per-group dealer count plus up to 10 sample
    names - so the group membership itself can be sanity-checked against
    what the owner expects before any Rupee figures are trusted.

    Step 3: attempts the native "Group Summary" report export
    (GROUP_SUMMARY_REPORT_XML above) and saves the raw response to
    tally_group_summary_dump.xml, printing an excerpt - this is the
    genuinely open question (does it return Opening/Debit/Credit/Closing,
    or just Closing Balance) - read directly from real Tally output here,
    not guessed.

    CONFIRMED 28 Sep 2026 from a real run of Steps 1-3 above: 7 of the 8
    groups FOUND (all under parent "PETRONAS TOTAL AREA SALES"), only
    "PETRONAS LEFT OUT DEALERS" MISSING; and Step 3's native report export
    does NOT return the 8 custom groups at all - it returns the company's
    standard primary chart-of-accounts groups, each with only a single
    Closing Dr/Cr figure (DSPCLDRAMTA/DSPCLCRAMTA), no Opening Balance, no
    period Debit/Credit movement. So two more steps were added in response:

    Step 1b: lists every OTHER immediate child of "PETRONAS TOTAL AREA
    SALES" not already in KNOWN_COLLECTION_GROUP_NAMES, to try to resolve
    what "PETRONAS LEFT OUT DEALERS" is really called (or confirm it isn't
    a real Group at all).

    Step 4: since the native report path is a dead end, computes Opening/
    Debit/Credit/Closing per group directly from ledger entries instead -
    the same "ask Tally for every raw voucher, compute client-side"
    approach already proven for fetch_collection()/fetch_sales(), but
    UNFILTERED (every voucher type, not just Receipt) since a real Group
    Summary's movement includes Sales, Credit Notes, Journals, etc. Prints
    a table to compare against the owner's real screenshot before any live
    fetch_group_summary() gets written.
    """
    now = datetime.now()
    from_d = from_date or now.strftime("%Y%m01")
    to_d = to_date or now.strftime("%Y%m%d")

    print("Fetching Tally's Group master hierarchy...")
    group_xml = tally_request(GROUP_HIERARCHY_COLLECTION_XML)
    group_root = ET.fromstring(group_xml)
    group_parent_all = {}
    for grp in group_root.iter("GROUP"):
        gname = grp.get("NAME") or _text(grp, "NAME")
        if gname:
            group_parent_all[gname.strip()] = (_text(grp, "PARENT") or "").strip()

    print("\n" + "=" * 70)
    print("STEP 1: do the 7 target group names exist in Tally's Group master?")
    print("=" * 70)
    all_found = True
    for gname in KNOWN_COLLECTION_GROUP_NAMES:
        if gname in group_parent_all:
            print("  FOUND   {}  ->  parent: {}".format(gname, group_parent_all[gname] or "(top-level)"))
        else:
            all_found = False
            print("  MISSING {}  -- not found in the Group master at all (check spelling/"
                  "case against your real Tally, or send the exact name as Tally shows it)".format(gname))
    if not all_found:
        print("\n  At least one group name didn't match exactly - group membership below")
        print("  will be incomplete for any group marked MISSING above.")

    print("\n" + "=" * 70)
    print("STEP 1b: any OTHER immediate children of 'PETRONAS TOTAL AREA SALES'")
    print("         not already in the known group names? (sanity check - the")
    print("         owner confirmed 'PETRONAS  LEFT OUT DEALERS' exists but is")
    print("         intentionally excluded from this feature, so it's expected")
    print("         to show up here every time)")
    print("=" * 70)
    known_set_display = set(KNOWN_COLLECTION_GROUP_NAMES)
    siblings = sorted(
        gname for gname, gparent in group_parent_all.items()
        if gparent == "PETRONAS TOTAL AREA SALES" and gname not in known_set_display
    )
    if siblings:
        print("  {} other group(s) under 'PETRONAS TOTAL AREA SALES' (excluded on purpose):".format(len(siblings)))
        for s in siblings:
            print("    " + s)
        print("\n  If a NEW name shows up here that isn't the Left Out Dealers group,")
        print("  send it back - that would mean a new sales area was added in Tally.")
    else:
        print("  None found - only the 7 known groups exist under 'PETRONAS TOTAL")
        print("  AREA SALES' right now.")

    print("\n" + "=" * 70)
    print("STEP 2: which dealers resolve to each group (via the group chain walk)?")
    print("=" * 70)
    ledger_parent, _group_parent, resolve = _build_collection_group_resolver()
    dealers_by_group = dict((g, []) for g in KNOWN_COLLECTION_GROUP_NAMES)
    for lname, lparent in ledger_parent.items():
        if not lparent:
            continue
        g = resolve(lparent)
        if g:
            dealers_by_group[g].append(lname)
    for gname in KNOWN_COLLECTION_GROUP_NAMES:
        names = sorted(dealers_by_group[gname])
        print("\n  {}  ->  {} dealer(s)".format(gname, len(names)))
        for n in names[:10]:
            print("    " + n)
        if len(names) > 10:
            print("    ... and {} more".format(len(names) - 10))

    print("\n" + "=" * 70)
    print("STEP 3: attempting Tally's native 'Group Summary' report export")
    print("        ({} to {})".format(from_d, to_d))
    print("=" * 70)
    xml_body = GROUP_SUMMARY_REPORT_XML.format(from_date=from_d, to_date=to_d)
    try:
        report_xml = tally_request(xml_body, timeout=300)
    except Exception as e:
        print("  Request failed: {}".format(e))
        print("  This likely means 'Group Summary' isn't the right REPORTNAME for this")
        print("  Tally version/setup - send this error back so a different report name")
        print("  can be tried.")
        return
    out_dir = os.path.dirname(os.path.abspath(__file__))
    dump_path = os.path.join(out_dir, "tally_group_summary_dump.xml")
    with open(dump_path, "w", encoding="utf-8") as f:
        f.write(report_xml)
    print("  Saved full raw response to: {}  ({} bytes)".format(dump_path, len(report_xml)))

    # Excerpt around the first target group name found in the raw text,
    # so the real column/tag shape can be read directly without opening
    # the (possibly large) saved file.
    excerpt_shown = False
    for gname in KNOWN_COLLECTION_GROUP_NAMES:
        idx = report_xml.upper().find(gname.upper())
        if idx != -1:
            start = max(0, idx - 200)
            end = min(len(report_xml), idx + 2000)
            print("\n--- Raw excerpt around '{}' ---".format(gname))
            print(report_xml[start:end])
            excerpt_shown = True
            break
    if not excerpt_shown:
        print("\n  None of the 7 group names were found as literal text anywhere in the")
        print("  raw response - first 2000 characters instead, so the overall shape can")
        print("  still be read (an error/'ID not found'-style message would show here):")
        print(report_xml[:2000])

    print("\n" + "=" * 70)
    print("STEP 4: computing Opening/Debit/Credit/Closing per group directly from")
    print("        ledger entries (ALL voucher types, {} to {})".format(from_d, to_d))
    print("        This re-fetches EVERY voucher company-wide for the period -")
    print("        heavier than anything else this file asks Tally for. If Tally")
    print("        feels slow while this runs, that's expected for a one-off")
    print("        diagnostic; the real live feature would need its own lighter,")
    print("        incremental approach before shipping.")
    print("=" * 70)

    print("  Fetching each ledger's current CLOSINGBALANCE ...")
    ledger_xml2 = tally_request(LEDGER_GROUP_COLLECTION_XML)
    ledger_root2 = ET.fromstring(ledger_xml2)
    ledger_closing = {}
    # Raw (unparsed) CLOSINGBALANCE text kept alongside the parsed number,
    # for STEP 5 below - added after the 28 Sep run showed 3 of the 7
    # groups (the ones with dealer names containing a trailing "*****" in
    # STEP 2's listing) off by a fixed Rs amount on Opening/Closing only
    # (Debit/Credit matched Tally's screenshot exactly), so a specific
    # dealer's CLOSINGBALANCE not being captured the way expected is the
    # leading suspect - seeing the raw tag text (missing entirely vs.
    # present-but-blank vs. a real number) is needed to tell which.
    ledger_closing_raw = {}
    for led in ledger_root2.iter("LEDGER"):
        lname2 = led.get("NAME") or _text(led, "NAME")
        if lname2:
            lname2 = lname2.strip()
            raw_val = _text(led, "CLOSINGBALANCE")
            ledger_closing_raw[lname2] = raw_val
            # _num() returns None (not 0.0) when CLOSINGBALANCE is missing/
            # blank - a real, common case for a dealer ledger sitting at a
            # genuine zero balance. Coalesced to 0.0 here so the sum() below
            # over dealers_by_group never hits float + None (confirmed bug
            # from the 28 Sep run: "TypeError: unsupported operand type(s)
            # for +: 'float' and 'NoneType'" on this exact line's old form).
            ledger_closing[lname2] = _num(raw_val) or 0.0

    print("  Fetching ALL vouchers company-wide for the period (unfiltered) ...")
    xml_body4 = ALL_VOUCHERS_FOR_GROUP_SUMMARY_XML.format(from_date=from_d, to_date=to_d)
    try:
        vouchers_xml = tally_request(xml_body4, timeout=600)
    except Exception as e:
        print("  Request failed: {}".format(e))
        print("  Skipping STEP 4 - send this error back.")
        vouchers_xml = None

    if vouchers_xml is not None:
        # Printed unconditionally, before any parsing, so a suspiciously
        # fast/small response (Tally returning an error or empty result
        # instead of real voucher data) is visible immediately rather than
        # silently producing an empty-looking table below.
        print("  Response received: {} bytes".format(len(vouchers_xml)))
        try:
            root4 = ET.fromstring(vouchers_xml)
            group_debit = dict((g, 0.0) for g in KNOWN_COLLECTION_GROUP_NAMES)
            group_credit = dict((g, 0.0) for g in KNOWN_COLLECTION_GROUP_NAMES)
            group_entry_count = dict((g, 0) for g in KNOWN_COLLECTION_GROUP_NAMES)
            # Per-DEALER (not just per-group) Debit/Credit, added for STEP 4b
            # below - after the 28 Sep run showed group totals close but not
            # exact (Debit/Credit matched the real screenshot exactly, only
            # Opening/Closing off by a fixed Rs amount per group), the next
            # question is WHICH dealer(s) - this makes that answerable by a
            # row-by-row CSV compare against the owner's own Tally drill-down
            # exports, instead of guessing from group-level totals alone.
            dealer_debit = {}
            dealer_credit = {}
            # POST-DATED correction, added 28 Sep 2026 after the owner
            # identified the real cause of every remaining mismatch: SRI
            # VINAYAKA MOTORS,MDU (OIL) had a post-dated cheque dated
            # 01.10.2026 already entered in Tally. ledger_closing (the plain
            # LEDGER CLOSINGBALANCE fetch, no date scoping) reflects EVERY
            # voucher ever entered for that ledger, INCLUDING ones dated
            # after this period's TO date - so a post-dated cheque already
            # nets out of it even though, as of the period's own end date,
            # it hasn't actually cleared and the dealer still legitimately
            # owes the money. ALL_VOUCHERS_FOR_GROUP_SUMMARY_XML already
            # comes back unfiltered by date (confirmed elsewhere in this
            # file: SVFROMDATE/SVTODATE don't scope a raw TYPE="Voucher"
            # export), so every post-dated voucher is sitting right there in
            # the same response - it was just being skipped by the old
            # "vdate outside from_d..to_d -> continue" filter. Now tracked
            # separately per dealer and SUBTRACTED back out of
            # ledger_closing below (_corrected_closing()), so Closing
            # reflects the ledger's true balance as of the period's own TO
            # date, matching what Tally's own Group Summary report shows -
            # not the ledger's real-time "as of right now" balance.
            dealer_future_debit = {}
            dealer_future_credit = {}
            voucher_count4 = 0
            entries_matched = 0
            future_voucher_count = 0
            for v in root4.iter("VOUCHER"):
                vdate = _text(v, "DATE")
                if not vdate:
                    continue
                is_future = vdate > to_d
                if not is_future and vdate < from_d:
                    continue
                if is_future:
                    future_voucher_count += 1
                else:
                    voucher_count4 += 1
                for entry in v.iter("ALLLEDGERENTRIES.LIST"):
                    lname3 = (_text(entry, "LEDGERNAME") or "").strip()
                    if not lname3:
                        continue
                    lparent3 = ledger_parent.get(lname3)
                    if not lparent3:
                        continue
                    g = resolve(lparent3)
                    if not g:
                        continue
                    amt = _num(_text(entry, "AMOUNT"))
                    if is_future:
                        if amt >= 0:
                            dealer_future_credit[lname3] = dealer_future_credit.get(lname3, 0.0) + amt
                        else:
                            dealer_future_debit[lname3] = dealer_future_debit.get(lname3, 0.0) + (-amt)
                        continue
                    if amt >= 0:
                        group_credit[g] += amt
                        dealer_credit[lname3] = dealer_credit.get(lname3, 0.0) + amt
                    else:
                        group_debit[g] += -amt
                        dealer_debit[lname3] = dealer_debit.get(lname3, 0.0) + (-amt)
                    group_entry_count[g] += 1
                    entries_matched += 1

            def _corrected_closing(d):
                """Raw current CLOSINGBALANCE, minus any post-dated (after
                to_d) movement already baked into it - see the note above.
                Falls back to the raw value for a dealer with no post-dated
                vouchers at all (the common case)."""
                raw = ledger_closing.get(d, 0.0)
                fut_net = dealer_future_credit.get(d, 0.0) - dealer_future_debit.get(d, 0.0)
                return raw - fut_net

            print("\n  {} vouchers in range, {} ledger entries matched to one of the 7 groups".format(
                voucher_count4, entries_matched))
            if future_voucher_count:
                print("  ({} post-dated voucher(s) found dated after {} - excluded from Debit/".format(
                    future_voucher_count, to_d))
                print("  Credit above, and backed out of Closing below, so Closing reflects the")
                print("  balance as of {} rather than today's real-time ledger balance.)".format(to_d))
            if voucher_count4 == 0:
                print("  ZERO vouchers matched the date range {} to {} - the raw response's".format(from_d, to_d))
                print("  first 1000 characters are printed below so the actual shape/error can")
                print("  be read directly (e.g. Tally may have returned an <ID not found> style")
                print("  message instead of voucher data for this unfiltered request):")
                print(vouchers_xml[:1000])
            print("\n  {:<38} {:>14} {:>14} {:>14} {:>14}".format(
                "Group", "Opening", "Debit", "Credit", "Closing"))
            for gname in KNOWN_COLLECTION_GROUP_NAMES:
                closing = sum(_corrected_closing(d) for d in dealers_by_group[gname])
                debit = group_debit[gname]
                credit = group_credit[gname]
                net_movement = credit - debit
                opening = closing - net_movement
                print("  {:<38} {:>14.2f} {:>14.2f} {:>14.2f} {:>14.2f}".format(
                    gname, opening, debit, credit, closing))
            print("\n  Compare this table against your real screenshot's September numbers")
            print("  for each group. If they match (even loosely, allowing for rounding),")
            print("  this ledger-entry approach is the one the real feature will use. If")
            print("  any group is off, send back which one and by roughly how much.")

            # STEP 4b: per-dealer CSV, one row per dealer in every group -
            # written so it can be diffed directly against the owner's own
            # Tally drill-down exports (already have two of these: Madurai
            # and Tirunelveli) to find the EXACT dealer(s) behind a group's
            # mismatch, rather than guessing from group totals. Now includes
            # the post-dated correction above, plus its own "PostDated"
            # column so a dealer with a post-dated cheque is visible
            # directly in the CSV instead of looking like an unexplained gap.
            csv_out_dir = os.path.dirname(os.path.abspath(__file__))
            csv_path = os.path.join(csv_out_dir, "tally_group_summary_dealer_breakdown.csv")
            with open(csv_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["Group", "Dealer", "Opening", "Debit", "Credit", "Closing", "PostDated"])
                for gname in KNOWN_COLLECTION_GROUP_NAMES:
                    for d in sorted(dealers_by_group[gname]):
                        d_closing = _corrected_closing(d)
                        d_debit = dealer_debit.get(d, 0.0)
                        d_credit = dealer_credit.get(d, 0.0)
                        d_opening = d_closing - (d_credit - d_debit)
                        d_postdated = dealer_future_credit.get(d, 0.0) - dealer_future_debit.get(d, 0.0)
                        writer.writerow([gname, d, "{:.2f}".format(d_opening),
                                          "{:.2f}".format(d_debit), "{:.2f}".format(d_credit),
                                          "{:.2f}".format(d_closing), "{:.2f}".format(d_postdated)])
            print("\n  STEP 4b: wrote a per-dealer breakdown (every dealer in every group,")
            print("  same 4 figures as above, plus a PostDated column) to:")
            print("  " + csv_path)
            print("  Please attach/send this CSV file back along with the console output -")
            print("  I already have your Madurai and Tirunelveli drill-down Excel exports,")
            print("  so I can diff this against them row-by-row to find the exact dealer(s)")
            print("  causing each group's Opening/Closing mismatch, without more guessing.")
        except Exception:
            print("  STEP 4 processing failed - full error below (send this back as-is):")
            print("  " + "-" * 66)
            traceback.print_exc()
            print("  " + "-" * 66)
            print("  Raw response's first 1000 characters, for context:")
            print(vouchers_xml[:1000])

    print("\n" + "=" * 70)
    print("STEP 5: dealers whose ledger name contains a '*' (as seen in STEP 2's")
    print("        listings, e.g. 'ARUNACHAL MOTORS PVT LTD, MADURAI *****') -")
    print("        checking whether these line up with STEP 4's Opening/Closing")
    print("        mismatches. On the 28 Sep run, 3 groups (AAA Direct, Madurai,")
    print("        Tirunelveli) were off by a fixed Rs amount on Opening AND")
    print("        Closing only - Debit/Credit matched the real screenshot")
    print("        exactly - which points at specific dealer(s)' CLOSINGBALANCE")
    print("        not being captured the way expected, not a formula bug.")
    print("=" * 70)
    any_starred = False
    for gname in KNOWN_COLLECTION_GROUP_NAMES:
        starred = sorted(d for d in dealers_by_group[gname] if "*" in d)
        if not starred:
            continue
        any_starred = True
        subtotal = 0.0
        print("\n  {} - {} starred dealer(s):".format(gname, len(starred)))
        for d in starred:
            raw = ledger_closing_raw.get(d, None)
            parsed = ledger_closing.get(d, 0.0)
            print("    {:<55} raw CLOSINGBALANCE={!r:<12} parsed={:.2f}".format(d, raw, parsed))
            subtotal += parsed
        print("    subtotal of these starred dealers' CLOSINGBALANCE: {:.2f}".format(subtotal))
    if not any_starred:
        print("  No dealer names contain '*' in any of the 7 groups.")
    print("\n  Send this section back too - if a starred dealer's raw CLOSINGBALANCE")
    print("  above is None or blank despite them being a real, active dealer, that's")
    print("  very likely the cause of STEP 4's Opening/Closing mismatch for their")
    print("  group, and I'll need to fetch that dealer's balance a different way.")

    print("\n" + "=" * 70)
    print("Send this whole output back, plus tally_group_summary_dump.xml AND")
    print("tally_group_summary_dealer_breakdown.csv (both saved next to this script).")
    print("Specifically: (1) did all 7 groups show FOUND in STEP 1; (2) do the dealer")
    print("counts/names in STEP 2 look right; (3) STEP 3 is expected to NOT match your")
    print("screenshot (already known); (4) does STEP 4's table now match your real")
    print("screenshot's Opening/Debit/Credit/Closing numbers per group, now that")
    print("post-dated vouchers are excluded/backed out - this SHOULD close the gap")
    print("on AAA/Madurai/Tirunelveli if post-dated cheques were the whole story;")
    print("(5) the STEP 4b CSV's new PostDated column - if any group is STILL off,")
    print("that column shows exactly which dealer(s) still don't reconcile.")


# -- Petronas Business Summary diagnostic, added 30 Sep 2026 -------------
# Owner's request (verbatim, all caps): a month-wise (Jan 2026 onward)
# company-wide table with Purchase (Ltr), Sales (Ltr) [already have],
# Sales (Value) [already have], Collected Amount [already have],
# Purchase (Value) and Payment Made to Petronas - the last two both come
# from one ledger, "PETRONAS Total Purchase A/C" (confirmed by the owner
# 30 Sep 2026 to sit under Sundry Creditors): its CREDIT movement is
# Purchase (Value) (a purchase booking increases what AAA owes), its
# DEBIT movement is Payment Made (a payment reduces what AAA owes) - the
# same amt>=0-is-Credit/amt<0-is-Debit convention already proven in
# dump_group_summary()'s STEP 4 above, just on a creditor ledger instead
# of 200+ debtor ledgers. Purchase (Ltr) is the inward quantity for the
# same 17 "starts with PETRONAS" stock item groups already proven for
# Sales (isPetronasItemGroup_ in script3_v2.js / the daybook dump above)
# - but which Tally VOUCHER TYPE actually carries inward/purchase entries
# in this company (plain "Purchase"? "Purchase - PETRONAS"? something
# else?) has never been confirmed, and Purchase-side qty might not be in
# the same NOS-per-bottle unit Sales turned out to be (dump_stock_unit
# above exists for exactly this reason on the stock side). So, per the
# owner's own "Diagnostic first" choice and this project's standing
# "never guess a new Tally structure" rule: DIAGNOSTIC ONLY, nothing
# pushed anywhere, no fetch_petronas_purchase() exists yet. Deliberately
# UNFILTERED on voucher type for the inventory scan (unlike
# DAYBOOK_VOUCHERS_TEMPLATE's Sales-only filter) specifically so every
# voucher type touching a Petronas item group shows up, including ones
# that turn out to be irrelevant (e.g. Stock Journal transfers between
# godowns) - the point is to SEE the real shape before deciding what to
# filter on.
#
# FIRST REAL RUN, 30 Sep 2026 (owner's own office PC, Jan-Sep 2026):
# STEP 1 came back MISSING for "PETRONAS Total Purchase A/C" exactly as
# spelled/spaced here - the owner's real ledger name differs somehow
# (case/spacing/wording) and still needs to be sent back exact. STEP 3
# was very informative even without the ledger fixed: real voucher type
# names in this company are "Purchase" (558 lines) and "Purchase - CBE"
# (31 lines) - separate from "Sales - PETRONAS"/"Sales - CBE"/"Sales"/
# "Sales - CASH"/plus "Branch Transfer"/"Credit Note"/"Debit Note"/
# "POS BILLING" which also touch Petronas item groups but are presumably
# not real inward Purchase. Also confirmed: some raw ACTUALQTY/BILLEDQTY
# readings come back as a compound '<qty> NOS = <ltr> LTR' string (see
# _parse_native_ltr(), further up this file, already in use by
# dump_daybook()'s "Alt Qty - Tally-native attempt" column) rather than a
# bare qty - if this also holds for the real "Purchase"/"Purchase - CBE"
# lines (not yet confirmed in this first run, which only sampled the
# first 15 lines overall, nearly all of which happened to be Sales -
# PETRONAS), that's a far more trustworthy Ltr figure than
# ITEM_LTR_CATALOG's hardcoded guess.
#
# SECOND REAL RUN, 30 Sep 2026, same day - STEP 1 still MISSING (still
# waiting on the owner's exact ledger name), but STEP 3's per-voucher-
# type sample fix paid off immediately: "Purchase" showed 451/558 lines
# with a native Ltr reading, "Purchase - CBE" only 6/31 - and the 6
# CBE ones that DID show a reading were all in the FLIPPED order
# ('1260.000 LTR =  6 NOS', Ltr before '=', not after) that
# _NATIVE_LTR_RE didn't match yet - a real, pre-existing gap in
# _parse_native_ltr() (shared with dump_daybook(), so this affects that
# diagnostic's "Alt Qty" column too, not just this one), not a Purchase-
# specific problem. FIXED the same day - _NATIVE_LTR_RE now matches the
# LTR figure in either order (see its own comment for detail). Nearly
# every TATA CBE/OE group item (item groups "PETRONAS TATA CBE GROUP"/
# "PETRONAS TATA OE GROUP") uses the flipped order, since Ltr is their
# PRIMARY unit (Nos is the alternate/compound one) rather than the other
# way around like most other Petronas items.
PETRONAS_PURCHASE_LEDGER_NAME = "PETRONAS Total Purchase A/C"  # STILL UNCONFIRMED as of the second run - STEP 1 keeps coming back MISSING, waiting on the owner's exact name (candidate ledgers are now auto-suggested on a miss, see STEP 1 below)

PETRONAS_INVENTORY_SCAN_TEMPLATE = """<ENVELOPE>
 <HEADER>
  <VERSION>1</VERSION>
  <TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE>
  <ID>PetronasInventoryScanCollection</ID>
 </HEADER>
 <BODY>
  <DESC>
   <STATICVARIABLES>
    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
    <SVFROMDATE>{from_date}</SVFROMDATE>
    <SVTODATE>{to_date}</SVTODATE>
   </STATICVARIABLES>
   <TDL>
    <TDLMESSAGE>
     <COLLECTION NAME="PetronasInventoryScanCollection" ISMODIFY="No">
      <TYPE>Voucher</TYPE>
      <FETCH>VOUCHERNUMBER, DATE, VOUCHERTYPENAME, PARTYLEDGERNAME</FETCH>
      <FETCH>ALLINVENTORYENTRIES.LIST</FETCH>
     </COLLECTION>
    </TDLMESSAGE>
   </TDL>
  </DESC>
 </BODY>
</ENVELOPE>"""


def _month_key(vdate8):
    """'20260315' -> '2026-03'."""
    return vdate8[:6] if vdate8 and len(vdate8) >= 6 else "unknown"


def _month_label(month_key6):
    """'202603' -> 'Mar 2026'; 'unknown' passed through as-is."""
    if month_key6 == "unknown" or len(month_key6) != 6:
        return month_key6
    try:
        return datetime.strptime(month_key6, "%Y%m").strftime("%b %Y")
    except ValueError:
        return month_key6


def dump_petronas_purchase(from_date=None, to_date=None):
    """
    --dump-petronas-purchase [YYYYMMDD-from] [YYYYMMDD-to]

    Diagnostic only - see the big comment above. Defaults to 1 Jan 2026
    through today, matching the owner's own "from Jan 2026 month wise"
    request, so a plain run with no args already produces the real
    month-wise shape he asked for (pending his confirmation below).

    STEP 1: confirms PETRONAS_PURCHASE_LEDGER_NAME exists in the Ledger
    master and prints its PARENT (should be Sundry Creditors, or a group
    that resolves to it) plus its current CLOSINGBALANCE.

    STEP 2: fetches every voucher company-wide (unfiltered, reusing
    ALL_VOUCHERS_FOR_GROUP_SUMMARY_XML - same heaviest-request pattern as
    dump_group_summary()'s STEP 4), pulls every ALLLEDGERENTRIES.LIST
    line matching that one ledger, and prints Debit/Credit per calendar
    month AND per voucher type (so a stray Journal/Debit Note/Credit Note
    entry - which might need different handling than plain Purchase/
    Payment - is visible rather than silently lumped in).

    STEP 3: fetches every voucher company-wide with inventory entries
    (PETRONAS_INVENTORY_SCAN_TEMPLATE, unfiltered on voucher type on
    purpose), resolves each inventory line's stock item to its immediate
    parent group via _build_item_group_resolver(), keeps only lines whose
    group starts with "petronas" (case-insensitive - same isPetronasItem
    Group_ rule already proven for Sales), and prints inward qty per
    calendar month AND per voucher type - both the RAW qty text (so the
    unit - Nos/Case/whatever Tally shows - can be read directly, since
    Sales turned out to be Nos-of-bottles and Purchase has never been
    checked) and a catalog-based Ltr guess (labelled as unverified for
    Purchase until confirmed).
    """
    now = datetime.now()
    from_d = from_date or "20260101"
    to_d = to_date or now.strftime("%Y%m%d")

    print("=" * 70)
    print("PETRONAS BUSINESS SUMMARY - DIAGNOSTIC ONLY - {} to {}".format(from_d, to_d))
    print("Nothing in this run is pushed anywhere or saved to the Sheet.")
    print("=" * 70)

    # STEP 1 ---------------------------------------------------------
    print("\n" + "=" * 70)
    print("STEP 1: does '{}' exist in the Ledger master?".format(PETRONAS_PURCHASE_LEDGER_NAME))
    print("=" * 70)
    ledger_xml = tally_request(LEDGER_GROUP_COLLECTION_XML)
    ledger_root = ET.fromstring(ledger_xml)
    target_lower = PETRONAS_PURCHASE_LEDGER_NAME.strip().lower()
    found_name = None
    found_parent = None
    found_closing_raw = None
    # All ledgers, kept for the candidate-search fallback below - built
    # regardless of whether the exact match succeeds, since it's cheap
    # (already-fetched XML, just iterated once) and STEP 2 needs the
    # confirmed name either way.
    all_ledgers = []  # (name, parent, closing_raw)
    for led in ledger_root.iter("LEDGER"):
        lname = led.get("NAME") or _text(led, "NAME")
        if not lname:
            continue
        lname = lname.strip()
        lparent = (_text(led, "PARENT") or "").strip()
        lclosing = _text(led, "CLOSINGBALANCE")
        all_ledgers.append((lname, lparent, lclosing))
        if lname.lower() == target_lower:
            found_name = lname
            found_parent = lparent
            found_closing_raw = lclosing

    resolved_ledger_name = None  # what STEP 2 will actually search for
    if found_name:
        print("  FOUND   {}".format(found_name))
        print("  PARENT: {}{}".format(
            found_parent or "(top-level)",
            "" if found_parent.strip().lower() == "sundry creditors" else
            "  <- NOT literally 'Sundry Creditors', check this is still the right group"))
        print("  Current CLOSINGBALANCE (raw, real-time, not date-scoped): {!r}".format(found_closing_raw))
        resolved_ledger_name = found_name
    else:
        print("  MISSING - '{}' was not found (exact match, case-insensitive) in the Ledger".format(
            PETRONAS_PURCHASE_LEDGER_NAME))
        print("  master at all.")
        # Candidate search, added 30 Sep 2026 after the exact name missed
        # twice in a row - rather than making the owner hunt through
        # Tally's UI a third time, surface every ledger whose name
        # contains both 'petronas' and 'purchase', OR sits directly under
        # Sundry Creditors with 'petronas' in its name, so the real name
        # is very likely just sitting in this list already.
        candidates = []
        seen = set()
        for lname, lparent, lclosing in all_ledgers:
            ll = lname.lower()
            is_candidate = (
                ("petronas" in ll and "purchase" in ll) or
                ("petronas" in ll and lparent.strip().lower() == "sundry creditors")
            )
            if is_candidate and lname not in seen:
                seen.add(lname)
                candidates.append((lname, lparent, lclosing))
        if candidates:
            print("  {} possible candidate(s) found (name contains 'petronas'+'purchase', or".format(
                len(candidates)))
            print("  'petronas' with PARENT exactly 'Sundry Creditors') - is one of these it?")
            for lname, lparent, lclosing in candidates[:20]:
                print("    {!r}  (PARENT: {}, CLOSINGBALANCE: {!r})".format(lname, lparent or "(top-level)", lclosing))
            if len(candidates) == 1:
                # Unambiguous - use it for STEP 2 below rather than
                # showing another all-zero table, but still flagged
                # clearly as a guess, not a confirmed match.
                resolved_ledger_name = candidates[0][0]
                print("\n  Exactly one candidate - using it for STEP 2 below, but CONFIRM this is")
                print("  really the right ledger before trusting the numbers.")
            else:
                print("\n  More than one candidate - STEP 2 below will show ZERO matches until you")
                print("  send back which exact name is the real one.")
        else:
            print("  No candidates found either (nothing with 'petronas'+'purchase' in the name,")
            print("  nothing under Sundry Creditors with 'petronas' in the name). Send back the")
            print("  EXACT name as Tally shows it (Gateway of Tally -> Chart of Accounts ->")
            print("  Ledgers, or Display -> Statements of Accounts -> Outstandings -> Ledger,")
            print("  under Sundry Creditors) - spelling/spacing/case must match exactly for")
            print("  STEP 2 below to find anything.")

    # STEP 2 ---------------------------------------------------------
    out_dir = os.path.dirname(os.path.abspath(__file__))
    print("\n" + "=" * 70)
    if resolved_ledger_name is None:
        print("STEP 2: SKIPPED - no ledger name to search for (STEP 1 found zero")
        print("        candidates). Send back the exact ledger name and re-run.")
        print("=" * 70)
    else:
        print("STEP 2: month-wise Debit (Payment Made)/Credit (Purchase Value) for '{}',".format(
            resolved_ledger_name))
        print("        {} to {} (fetching ALL vouchers company-wide, unfiltered - this is".format(
            from_d, to_d))
        print("        the heaviest request in this file, please be patient)")
        print("=" * 70)
    resolved_lower = resolved_ledger_name.strip().lower() if resolved_ledger_name else None

    month_debit = {}
    month_credit = {}
    vtype_debit = {}
    vtype_credit = {}
    vtype_count = {}
    ledger_voucher_count = 0
    vouchers_xml = ""
    if resolved_ledger_name is not None:
        vouchers_xml = tally_request(
            ALL_VOUCHERS_FOR_GROUP_SUMMARY_XML.format(from_date=from_d, to_date=to_d), timeout=600)
        ledger_dump_path = os.path.join(out_dir, "tally_petronas_purchase_ledger_dump.xml")
        with open(ledger_dump_path, "w", encoding="utf-8") as f:
            f.write(vouchers_xml)
        print("  (raw response saved to {}, {} bytes)".format(ledger_dump_path, len(vouchers_xml)))
        root2 = ET.fromstring(vouchers_xml)

        for v in root2.iter("VOUCHER"):
            vdate = (_text(v, "DATE") or "").strip()
            if not vdate or vdate < from_d or vdate > to_d:
                continue
            vtype = (_text(v, "VOUCHERTYPENAME") or "(blank)").strip()
            matched_this_voucher = False
            for entry in v.iter("ALLLEDGERENTRIES.LIST"):
                lname = (_text(entry, "LEDGERNAME") or "").strip()
                if lname.lower() != resolved_lower:
                    continue
                matched_this_voucher = True
                amt = _num(_text(entry, "AMOUNT")) or 0.0
                mk = _month_key(vdate)
                if amt >= 0:
                    month_credit[mk] = month_credit.get(mk, 0.0) + amt
                    vtype_credit[vtype] = vtype_credit.get(vtype, 0.0) + amt
                else:
                    month_debit[mk] = month_debit.get(mk, 0.0) + (-amt)
                    vtype_debit[vtype] = vtype_debit.get(vtype, 0.0) + (-amt)
                vtype_count[vtype] = vtype_count.get(vtype, 0) + 1
            if matched_this_voucher:
                ledger_voucher_count += 1

    print("\n  {} voucher(s) touched this ledger in range.".format(ledger_voucher_count))
    if ledger_voucher_count == 0:
        print("  ZERO matches - either the ledger name didn't match (see STEP 1) or it")
        print("  genuinely has no activity in this range. First 1000 chars of the raw")
        print("  response, for a sanity check:")
        print("  " + vouchers_xml[:1000])
    else:
        all_months = sorted(set(list(month_debit.keys()) + list(month_credit.keys())))
        print("\n  {:<12} {:>18} {:>18}".format("Month", "Debit (Payment)", "Credit (Purchase)"))
        tot_d = tot_c = 0.0
        for mk in all_months:
            d = month_debit.get(mk, 0.0)
            c = month_credit.get(mk, 0.0)
            tot_d += d
            tot_c += c
            print("  {:<12} {:>18.2f} {:>18.2f}".format(_month_label(mk), d, c))
        print("  {:<12} {:>18.2f} {:>18.2f}".format("Total", tot_d, tot_c))

        print("\n  Per voucher type (so an unexpected Journal/Debit Note/Credit Note is")
        print("  visible, not silently lumped into Payment/Purchase):")
        print("  {:<28} {:>10} {:>16} {:>16}".format("Voucher Type", "Count", "Debit", "Credit"))
        for vtype in sorted(vtype_count.keys()):
            print("  {:<28} {:>10} {:>16.2f} {:>16.2f}".format(
                vtype, vtype_count[vtype], vtype_debit.get(vtype, 0.0), vtype_credit.get(vtype, 0.0)))

    # STEP 3 ---------------------------------------------------------
    print("\n" + "=" * 70)
    print("STEP 3: month-wise inward qty for the 17 Petronas item groups, {} to {}".format(from_d, to_d))
    print("        (fetching ALL vouchers company-wide with inventory entries,")
    print("        unfiltered on voucher type on purpose)")
    print("=" * 70)
    print("Resolving Tally's stock item group hierarchy (for Item Group)...")
    item_group = _build_item_group_resolver()
    print("  {} stock item(s) resolved to a group.".format(len(item_group)))

    inv_xml = tally_request(PETRONAS_INVENTORY_SCAN_TEMPLATE.format(from_date=from_d, to_date=to_d), timeout=600)
    inv_dump_path = os.path.join(out_dir, "tally_petronas_purchase_inventory_dump.xml")
    with open(inv_dump_path, "w", encoding="utf-8") as f:
        f.write(inv_xml)
    print("  (raw response saved to {}, {} bytes)".format(inv_dump_path, len(inv_xml)))
    root3 = ET.fromstring(inv_xml)

    month_qty_raw = {}       # month_key -> summed _num(BILLEDQTY or ACTUALQTY), unit-agnostic
    month_qty_ltr = {}       # month_key -> summed catalog Ltr guess (only for items in catalog)
    month_qty_native_ltr = {}  # month_key -> summed Tally-native '= X LTR' parse (see _parse_native_ltr)
    vtype_inv_count = {}
    vtype_inv_qty_raw = {}
    vtype_native_ltr_count = {}  # vtype -> how many of its lines had a native '= X LTR' reading
    sample_lines_by_vtype = {}   # vtype -> up to 3 (item, item_grp, raw_actual, raw_billed, native_ltr) tuples,
                                  # PER voucher type (not just the first 15 overall) - added after the first
                                  # real run (30 Sep 2026) showed all 15 flat samples happened to come from one
                                  # voucher type (Sales - PETRONAS, by far the most common), leaving Purchase's
                                  # own raw format completely unseen.
    matched_lines = 0

    for v in root3.iter("VOUCHER"):
        vdate = (_text(v, "DATE") or "").strip()
        if not vdate or vdate < from_d or vdate > to_d:
            continue
        vtype = (_text(v, "VOUCHERTYPENAME") or "(blank)").strip()
        mk = _month_key(vdate)
        for entry in v.iter("ALLINVENTORYENTRIES.LIST"):
            item_name = (_text(entry, "STOCKITEMNAME") or "").strip()
            if not item_name:
                continue
            item_grp = item_group.get(item_name, "")
            if not _is_petronas_item_group(item_grp):
                continue
            raw_actual = _text(entry, "ACTUALQTY")
            raw_billed = _text(entry, "BILLEDQTY")
            qty_num = _num(raw_billed) if raw_billed is not None else _num(raw_actual)
            if qty_num is None:
                continue
            matched_lines += 1
            month_qty_raw[mk] = month_qty_raw.get(mk, 0.0) + qty_num
            vtype_inv_count[vtype] = vtype_inv_count.get(vtype, 0) + 1
            vtype_inv_qty_raw[vtype] = vtype_inv_qty_raw.get(vtype, 0.0) + qty_num
            ltr_guess = _catalog_ltr_guess(item_name, qty_num)
            if ltr_guess is not None:
                month_qty_ltr[mk] = month_qty_ltr.get(mk, 0.0) + ltr_guess
            native_ltr = _parse_native_ltr(raw_billed) or _parse_native_ltr(raw_actual)
            if native_ltr is not None:
                month_qty_native_ltr[mk] = month_qty_native_ltr.get(mk, 0.0) + native_ltr
                vtype_native_ltr_count[vtype] = vtype_native_ltr_count.get(vtype, 0) + 1
            bucket = sample_lines_by_vtype.setdefault(vtype, [])
            if len(bucket) < 3:
                bucket.append((item_name, item_grp, raw_actual, raw_billed, native_ltr))

    print("\n  {} matching inventory line(s) (Petronas item groups, any voucher type).".format(matched_lines))
    if matched_lines == 0:
        print("  ZERO matches - first 1000 chars of the raw response, for a sanity check:")
        print("  " + inv_xml[:1000])
    else:
        all_months3 = sorted(month_qty_raw.keys())
        print("\n  {:<12} {:>16} {:>20} {:>20}".format(
            "Month", "Raw qty (summed)", "Tally-native Ltr", "Catalog Ltr guess"))
        tot_raw = tot_ltr = tot_native = 0.0
        for mk in all_months3:
            r = month_qty_raw.get(mk, 0.0)
            n = month_qty_native_ltr.get(mk, 0.0)
            l = month_qty_ltr.get(mk, 0.0)
            tot_raw += r
            tot_native += n
            tot_ltr += l
            print("  {:<12} {:>16.2f} {:>20.2f} {:>20.2f}".format(_month_label(mk), r, n, l))
        print("  {:<12} {:>16.2f} {:>20.2f} {:>20.2f}".format("Total", tot_raw, tot_native, tot_ltr))
        print("\n  NOTE: 'Raw qty (summed)' sums whatever number Tally returned regardless of")
        print("  unit (Nos/Case/etc mixed together if a voucher type's unit differs from Sales)")
        print("  - do NOT trust it as Ltr. 'Tally-native Ltr' is parsed directly from a 'LTR'")
        print("  reading Tally itself includes on some raw ACTUALQTY/BILLEDQTY text, in whichever")
        print("  order it appears ('44 NOS = 154 LTR' or '1260 LTR = 6 NOS', both handled) - the")
        print("  most trustworthy Ltr figure available where present, no catalog guess needed.")
        print("  'Catalog Ltr guess' is the older, less trustworthy fallback (hardcoded per-item")
        print("  Python table) - only relevant for the rare line with no native Ltr reading at all.")

        print("\n  Per voucher type (tells you which voucher type(s) are real inward/Purchase")
        print("  entries vs. something else, e.g. Stock Journal godown transfers), plus how many")
        print("  of that type's lines carried a native '= X LTR' reading:")
        print("  {:<28} {:>10} {:>16} {:>14}".format("Voucher Type", "Lines", "Raw qty (sum)", "Has native Ltr"))
        for vtype in sorted(vtype_inv_count.keys()):
            print("  {:<28} {:>10} {:>16.2f} {:>14}".format(
                vtype, vtype_inv_count[vtype], vtype_inv_qty_raw.get(vtype, 0.0),
                "{}/{}".format(vtype_native_ltr_count.get(vtype, 0), vtype_inv_count[vtype])))

        print("\n  Sample raw lines (up to 3 PER voucher type, so Purchase's own raw format is")
        print("  actually visible, not just whichever type happened to sort first):")
        for vtype in sorted(sample_lines_by_vtype.keys()):
            print("\n    -- {} --".format(vtype))
            for item_name, item_grp, raw_actual, raw_billed, native_ltr in sample_lines_by_vtype[vtype]:
                print("    {}  (Item Group: {})".format(item_name, item_grp))
                print("        ACTUALQTY(raw)={!r}  BILLEDQTY(raw)={!r}  native Ltr parsed={}".format(
                    raw_actual, raw_billed, native_ltr if native_ltr is not None else "n/a"))

    print("\n" + "=" * 70)
    print("Send this whole console output back (the two saved XML dump files are huge -")
    print("no need to send those unless specifically asked).")
    print("Specifically: (1) STEP 1 - if MISSING, send the EXACT ledger name as Tally")
    print("shows it (Gateway of Tally -> Chart of Accounts -> Ledgers -> search")
    print("'Petronas'/'Purchase' under Sundry Creditors) so STEP 2 can be re-run against")
    print("the right name; (2) STEP 2 - once the name is right, does the month-wise")
    print("Debit/Credit table match your real Tally Ledger Vouchers view for a month or")
    print("two you can check by eye, and does the per-voucher-type breakdown look right")
    print("(Purchase voucher = mostly Credit, Payment voucher = mostly Debit, nothing")
    print("unexpected in Journal/Debit Note/Credit Note rows); (3) STEP 3 - which voucher")
    print("type(s) in the per-voucher-type table are the REAL inward/Purchase entries")
    print("(send the exact name(s) back - 'Purchase' and 'Purchase - CBE' are the obvious")
    print("candidates if this company books Coimbatore separately), and for THOSE types")
    print("specifically: does 'Has native Ltr' show a fraction close to the line count (a")
    print("native Ltr reading is present), or close to 0 (no native reading, would need")
    print("the older catalog-guess fallback instead)?")


# -- Petronas Business Summary - THE REAL one-time report, 30 Sep 2026 --
# All three open unknowns from --dump-petronas-purchase are now confirmed
# by the owner against his real Tally screen:
#
# 1. "PETRONAS TOTAL PURCHASE A/C" is a GROUP (a real Tally "Group
#    Alteration" screenshot, under Sundry Creditors / Current
#    Liabilities, with "Group behaves like a sub-ledger: Yes") - NOT a
#    plain Ledger. That's exactly why the diagnostic's straight
#    LEDGER-type lookup kept coming back MISSING twice - it was looking
#    for the wrong kind of Tally object. A "behaves like a sub-ledger"
#    Group consolidates every individual ledger filed under it into one
#    reporting line in Tally's own UI, so the real fix is: resolve every
#    LEDGER whose PARENT chain reaches this Group (same chain-walk
#    pattern already proven in _build_collection_group_resolver()), then
#    sum Debit/Credit across ALL of them combined - not a single-ledger-
#    name lookup at all.
# 2. "NO NEED OF Purchase - CBE TO ADD" - the owner's own words, 30 Sep
#    2026 - only the plain "Purchase" voucher type counted as real inward/
#    Purchase for this report, "Purchase - CBE" deliberately excluded.
#    SUPERSEDED 6 Oct 2026 - owner asked to add it back in ("please add
#    that also, only want to omit is tata oem purchase") - see
#    PETRONAS_PURCHASE_VOUCHER_TYPE_NAMES below for the current rule.
# 3. "YES PURCHASE ALSO CARRIES SAME conversion" - the owner's own
#    words, confirming "Purchase" voucher lines carry the same native
#    '<qty> NOS = <ltr> LTR' reading already proven for Sales, so
#    _parse_native_ltr() is trustworthy for Purchase too (falling back
#    to the older ITEM_LTR_CATALOG guess only for the rare line with no
#    native reading).
#
# So this is no longer a diagnostic - it's the real one-time report the
# owner asked for, per his own "One-time report" choice (a snapshot
# pulled once and manually baked into the dashboard, not a live poll).
# Sales (Ltr/Value) and Collected Amount are recomputed here too, even
# though the owner said "ALREADY U HAVE DATA" for them, because that
# phrase meant the underlying mechanism (the item-group Sales filter,
# the Receipt-voucher Collection filter) is already proven elsewhere in
# this project, not that historical Jan-onward MONTH-WISE figures
# already exist anywhere - today's live dashboard only ever shows
# current-month totals. Sales reuses the exact same voucher-type scope
# already proven for the live "Sales by Item Group" feature (Sales,
# Sales - PETRONAS, Sales - CBE - CBE is NOT excluded here, only for
# Purchase, since the owner's CBE exclusion was stated for Purchase
# specifically). Collected Amount reuses the exact same "Receipt
# vouchers against the 7 known collection groups" logic already proven
# in fetch_collection()/dump_group_summary(), just company-wide and
# month-wise instead of current-month-only.
PETRONAS_PURCHASE_GROUP_NAME = "PETRONAS TOTAL PURCHASE A/C"  # CONFIRMED 30 Sep 2026 - a Group, not a Ledger
# Widened 6 Oct 2026, per the owner's explicit request: "please add that
# also [Purchase - CBE], only want to omit is tata oem purchase." Purchase
# (Ltr) was excluding the whole "Purchase - CBE" voucher type (same as the
# already-included "Sales - CBE" on the sales side, DAYBOOK_VOUCHER_TYPE_
# NAMES below) - now both count. TATA OEM stays excluded exactly as before,
# but that exclusion was NEVER voucher-type-based to begin with - it's the
# separate, independent item-group filter every one of these 4 call sites
# already applies right after this check (_is_petronas_item_group(item_grp),
# which excludes PETRONAS_ITEM_GROUP_EXCLUSIONS_ = "petronas tata brake
# fluid"/"petronas tata cbe group"/"petronas tata oe group"). So a TATA OEM
# item purchased on a "Purchase - CBE" voucher is still correctly left out,
# by item group, regardless of this voucher-type widening.
PETRONAS_PURCHASE_VOUCHER_TYPE_NAMES = ("purchase", "purchase - cbe")

# Added 6 Oct 2026, per the owner's explicit report (with a real Tally
# "Group Summary - COLLECTION - My View" screenshot, Sep 2026, Credit
# column Grand Total Rs 87,64,145): "FOR COLLECTED AMOUNT ALSO I NEED TO
# INCLUDE ALL VOUCHERS OF Cheque,Commercial Credit Note,Credit Note,
# Receipt,SCHEME CREDIT NOTE". Business Summary's live "Collected" (the
# fetch_petronas_summary() function below, which drives the dashboard's
# Business Summary panel Sep 2026 onward) was still filtering to "receipt"
# only - narrower even than COLLECTION_VOUCHER_TYPE_NAMES above (which
# already widened the separate DSR Collection panel to receipt+cheque back
# on 18 Sep 2026). This is the Business-Summary-specific fix: every one of
# the 5 voucher types the owner named, scoped to the same 7 known
# collection groups as before (coll_resolve), same amt >= 0 = Cr = money
# in convention already proven for Receipt.
PETRONAS_SUMMARY_COLLECTION_VOUCHER_TYPE_NAMES = (
    "receipt", "cheque", "commercial credit note", "credit note", "scheme credit note"
)

# Added 15 Oct 2026, per the owner's direct correction on the rendered
# Petronas Business Summary table: "FROM SEP DO NOT CALCULATE SALES VALUE
# IF BILLED FOR PETRONAS LUBRICANTS AS IT IS FOR CLAIM PURPOSE AFTER CN
# COMES ONLY WE HAVE TO ADD THAT IN PAYMENT" - clarified in chat: these are
# Sales-type vouchers billed to the party "BILL TO PETRONAS LUBRICANTS
# (INDIA) Pvt. Ltd. CHENNAI" itself (not a real dealer) - a claim document
# styled as a sales invoice, not real dealer revenue, so it should never
# count toward Sales (Value). When Petronas later confirms the claim with
# a real Credit Note, that CN amount is tracked separately below
# (claimsCn) and shown as its OWN column on the dashboard ("Add as a
# separate line" - the owner's own choice over folding it into Payment
# Made), not added into Sales (Value) or Payment Made to Petronas.
#
# _is_petronas_claim_party_() is tolerant of minor punctuation/whitespace
# differences from what actually sits in Tally (periods, extra spaces,
# "Pvt Ltd" vs "Pvt. Ltd.") since the owner typed the name once in chat,
# not copy-pasted from Tally itself - matches on the distinctive substring
# "petronas lubricants" + "chennai" after normalizing, rather than an exact
# string compare that a single stray character would silently break.
PETRONAS_CLAIM_PARTY_NAME_ = "PETRONAS LUBRICANTS (INDIA) PVT. LTD. CHENNAI"


def _normalize_party_name_(name):
    import re as _re
    return _re.sub(r"[^a-z0-9]+", " ", (name or "").lower()).strip()


def _is_petronas_claim_party_(party_name):
    norm = _normalize_party_name_(party_name)
    return "petronas lubricants" in norm and "chennai" in norm


# Same confirmed Credit Note voucher-type family as
# PETRONAS_SUMMARY_COLLECTION_VOUCHER_TYPE_NAMES above (owner-confirmed
# real Tally voucher-type names on this setup) - used to find the Credit
# Notes Petronas itself raises against AAA's claim (party = the claim
# party name above), as opposed to a dealer-side credit note.
PETRONAS_CLAIM_CN_VOUCHER_TYPE_NAMES = ("commercial credit note", "credit note", "scheme credit note")


def _resolve_ledgers_under_group(target_group_name):
    """Every LEDGER whose PARENT chain eventually reaches
    target_group_name (exact name match at any point in the chain,
    case-insensitive) - same chain-walk pattern as
    _build_collection_group_resolver(), just resolving against one
    single target Group instead of a known list of several. Also
    includes the target name itself if it happens to exist as its own
    LEDGER (defensive - a 'behaves like a sub-ledger' Group can
    reportedly still carry its own direct Ledger in some Tally setups,
    not just child ledgers). Returns a set of ledger names."""
    ledger_xml = tally_request(LEDGER_GROUP_COLLECTION_XML)
    ledger_root = ET.fromstring(ledger_xml)
    ledger_parent = {}
    for led in ledger_root.iter("LEDGER"):
        name = led.get("NAME") or _text(led, "NAME")
        if name:
            ledger_parent[name.strip()] = (_text(led, "PARENT") or "").strip()

    group_xml = tally_request(GROUP_HIERARCHY_COLLECTION_XML)
    group_root = ET.fromstring(group_xml)
    group_parent = {}
    for grp in group_root.iter("GROUP"):
        gname = grp.get("NAME") or _text(grp, "NAME")
        if gname:
            group_parent[gname.strip()] = (_text(grp, "PARENT") or "").strip()

    target_lower = target_group_name.strip().lower()

    def resolves_to_target(immediate_group):
        current = immediate_group
        seen = set()
        while current and current not in seen:
            if current.strip().lower() == target_lower:
                return True
            seen.add(current)
            current = group_parent.get(current, "")
        return False

    matching = set()
    for lname, lparent in ledger_parent.items():
        if lparent and resolves_to_target(lparent):
            matching.add(lname)
    for lname in ledger_parent:
        if lname.strip().lower() == target_lower:
            matching.add(lname)
    return matching


def report_petronas_summary(from_date=None, to_date=None):
    """
    --petronas-summary [YYYYMMDD-from] [YYYYMMDD-to]

    THE REAL one-time report (see the big comment above for what got
    confirmed and why this is no longer diagnostic-only). Defaults to 1
    Jan 2026 through today, matching "from Jan 2026 month wise". Prints
    the final month-wise table in the exact column order the owner
    approved on the mockup: Purchase (Ltr), Sales (Ltr), Stock
    (Purchase - Sales), Sales (Value), Collected Amount, Purchase
    (Value), Payment Made to Petronas. Still pushes nothing live and
    writes nothing to the Sheet - these numbers get manually reviewed
    against what the owner independently knows, then baked into the
    dashboard as a static section, per his own "one-time report" choice.
    """
    now = datetime.now()
    from_d = from_date or "20260101"
    to_d = to_date or now.strftime("%Y%m%d")
    out_dir = os.path.dirname(os.path.abspath(__file__))

    print("=" * 70)
    print("PETRONAS BUSINESS SUMMARY - REAL ONE-TIME REPORT - {} to {}".format(from_d, to_d))
    print("Nothing pushed anywhere or saved to the Sheet - review these numbers, then")
    print("they get manually baked into the dashboard.")
    print("=" * 70)

    print("\nResolving every ledger under the '{}' group...".format(PETRONAS_PURCHASE_GROUP_NAME))
    purchase_ac_ledgers = _resolve_ledgers_under_group(PETRONAS_PURCHASE_GROUP_NAME)
    print("  {} ledger(s) resolved:".format(len(purchase_ac_ledgers)))
    for lname in sorted(purchase_ac_ledgers)[:20]:
        print("    " + lname)
    if not purchase_ac_ledgers:
        print("  ZERO ledgers resolved - the group name may not match either. Stopping here -")
        print("  send this output back before anything else is computed.")
        return

    print("\nResolving the {} known collection (Sales-area) groups' dealer ledgers...".format(
        len(KNOWN_COLLECTION_GROUP_NAMES)))
    coll_ledger_parent, _coll_group_parent, coll_resolve = _build_collection_group_resolver()
    print("  {} ledger(s) total in the company's Ledger master.".format(len(coll_ledger_parent)))

    print("\nFetching ALL vouchers company-wide, {} to {} (unfiltered - heaviest request".format(
        from_d, to_d))
    print("in this file, please be patient)...")
    vouchers_xml = tally_request(
        ALL_VOUCHERS_FOR_GROUP_SUMMARY_XML.format(from_date=from_d, to_date=to_d), timeout=600)
    print("  Response received: {} bytes.".format(len(vouchers_xml)))
    vroot = ET.fromstring(vouchers_xml)

    month_purchase_value = {}
    month_payment_made = {}
    month_collected = {}
    # Kept in sync with fetch_petronas_summary()'s live twin, added 15 Oct
    # 2026 - see PETRONAS_CLAIM_PARTY_NAME_'s own comment for the story.
    month_claims_cn = {}
    purchase_ac_lower = set(n.lower() for n in purchase_ac_ledgers)
    # Added 5 Oct 2026, per the owner's report: "in Payment Made to Petronas i
    # need to show here with omit of voucher made in Sales - PETRONAS". Root
    # cause - this loop below never filtered by voucher type at all: ANY
    # voucher (Sales, Journal, Credit/Debit Note, whatever) that happens to
    # touch a ledger resolving to the "PETRONAS Total Purchase A/C" group got
    # summed into Purchase Value / Payment Made, purely because of which
    # LEDGER it touched, with no check on what KIND of voucher it was. A
    # Sales voucher is never a real purchase or a real payment to Petronas,
    # whatever ledger it happens to post against - so every Sales-type
    # voucher (not just "Sales - PETRONAS" by name - "Sales"/"Sales - CBE"
    # carry the identical risk, same voucher-type set already proven as
    # DAYBOOK_VOUCHER_TYPE_NAMES/sales_vtypes further below in this same
    # function) is now excluded from this ledger-entry loop entirely.
    sales_vtypes_for_purchase_ac = set(DAYBOOK_VOUCHER_TYPE_NAMES)  # ("sales", "sales - petronas", "sales - cbe", "sales - cash", "pos billing") - 10 Oct 2026

    for v in vroot.iter("VOUCHER"):
        vdate = (_text(v, "DATE") or "").strip()
        if not vdate or vdate < from_d or vdate > to_d:
            continue
        mk = _month_key(vdate)
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip().lower()
        is_sales_voucher = vtype in sales_vtypes_for_purchase_ac
        v_party = (_text(v, "PARTYLEDGERNAME") or "").strip()
        is_claim_cn_voucher = vtype in PETRONAS_CLAIM_CN_VOUCHER_TYPE_NAMES and _is_petronas_claim_party_(v_party)
        for entry in v.iter("ALLLEDGERENTRIES.LIST"):
            lname = (_text(entry, "LEDGERNAME") or "").strip()
            if not lname:
                continue
            amt = _num(_text(entry, "AMOUNT")) or 0.0
            if lname.lower() in purchase_ac_lower and not is_sales_voucher:
                if amt >= 0:
                    month_purchase_value[mk] = month_purchase_value.get(mk, 0.0) + amt
                else:
                    month_payment_made[mk] = month_payment_made.get(mk, 0.0) + (-amt)
            # Collection: widened 6 Oct 2026 to all 5 owner-named voucher types
            # (see PETRONAS_SUMMARY_COLLECTION_VOUCHER_TYPE_NAMES above) against
            # a dealer resolving to one of the 7 known collection groups. Kept in
            # sync with fetch_petronas_summary()'s live twin below even though
            # this one-time report only ever fed the already-baked Jan-Aug
            # static dashboard snapshot - so a future manual rerun matches.
            if vtype in PETRONAS_SUMMARY_COLLECTION_VOUCHER_TYPE_NAMES:
                lparent = coll_ledger_parent.get(lname)
                if lparent and coll_resolve(lparent) and amt >= 0:
                    month_collected[mk] = month_collected.get(mk, 0.0) + amt
            # Petronas claim Credit Note - see fetch_petronas_summary()'s own
            # claims_cn comment, added 15 Oct 2026.
            if is_claim_cn_voucher and _is_petronas_claim_party_(lname):
                month_claims_cn[mk] = month_claims_cn.get(mk, 0.0) + abs(amt)

    print("\nFetching stock item group hierarchy...")
    item_group = _build_item_group_resolver()
    print("  {} stock item(s) resolved to a group.".format(len(item_group)))

    print("\nFetching ALL vouchers company-wide with inventory entries, {} to {} (unfiltered -".format(
        from_d, to_d))
    print("on voucher type, second-heaviest request, please be patient)...")
    inv_xml = tally_request(PETRONAS_INVENTORY_SCAN_TEMPLATE.format(from_date=from_d, to_date=to_d), timeout=600)
    print("  Response received: {} bytes.".format(len(inv_xml)))
    iroot = ET.fromstring(inv_xml)

    month_purchase_ltr = {}
    month_sales_ltr = {}
    month_sales_value = {}
    sales_vtypes = set(DAYBOOK_VOUCHER_TYPE_NAMES)  # ("sales", "sales - petronas", "sales - cbe", "sales - cash", "pos billing") - 10 Oct 2026

    for v in iroot.iter("VOUCHER"):
        vdate = (_text(v, "DATE") or "").strip()
        if not vdate or vdate < from_d or vdate > to_d:
            continue
        mk = _month_key(vdate)
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip().lower()
        is_purchase = (vtype in PETRONAS_PURCHASE_VOUCHER_TYPE_NAMES)
        is_sales = (vtype in sales_vtypes)
        if not is_purchase and not is_sales:
            continue
        # Added 15 Oct 2026 - kept in sync with fetch_petronas_summary()'s
        # own is_claim_party_voucher, see its comment for the full story.
        is_claim_party_voucher = is_sales and _is_petronas_claim_party_(_text(v, "PARTYLEDGERNAME"))
        for entry in v.iter("ALLINVENTORYENTRIES.LIST"):
            item_name = (_text(entry, "STOCKITEMNAME") or "").strip()
            if not item_name:
                continue
            item_grp = item_group.get(item_name, "")
            if not _is_petronas_item_group(item_grp):
                continue
            raw_actual = _text(entry, "ACTUALQTY")
            raw_billed = _text(entry, "BILLEDQTY")
            qty_num = _num(raw_billed) if raw_billed is not None else _num(raw_actual)
            ltr = _parse_native_ltr(raw_billed) or _parse_native_ltr(raw_actual)
            if ltr is None and qty_num is not None:
                ltr = _catalog_ltr_guess(item_name, qty_num)
            if ltr is None:
                continue
            if is_purchase:
                month_purchase_ltr[mk] = month_purchase_ltr.get(mk, 0.0) + ltr
            else:
                month_sales_ltr[mk] = month_sales_ltr.get(mk, 0.0) + ltr
                if is_claim_party_voucher:
                    continue
                amt = _num(_text(entry, "AMOUNT"))
                if amt is not None:
                    # Added 5 Oct 2026, per the owner's report: "i think Sales
                    # (Value) in dashboard Petronas Business Summary is
                    # without 18% - so i need that amount including gst".
                    # Confirmed: this AMOUNT comes straight from each stock
                    # item's own ALLINVENTORYENTRIES.LIST line, which in this
                    # company's Tally setup is the pre-GST (basic/taxable)
                    # value - GST is posted as separate CGST/SGST/IGST ledger
                    # entries elsewhere in the voucher, never folded into the
                    # item line's own amount. This is the exact same
                    # "amt_incl_gst = round(amt_abs * 1.18, 2)" conversion
                    # already proven and shipped for the Day Book export's own
                    # "Value incl GST" column (dump_daybook(), earlier in this
                    # file) - reused here rather than introducing a second,
                    # possibly-drifting GST formula, and matching the flat 18%
                    # GST rate already hardcoded dashboard-/DSR-app-wide
                    # (GST_RATE_PCT in every DSR app's own index.html).
                    month_sales_value[mk] = month_sales_value.get(mk, 0.0) + round(abs(amt) * 1.18, 2)

    all_months = sorted(set(
        list(month_purchase_ltr.keys()) + list(month_sales_ltr.keys()) +
        list(month_sales_value.keys()) + list(month_collected.keys()) +
        list(month_purchase_value.keys()) + list(month_payment_made.keys()) +
        list(month_claims_cn.keys())))

    print("\n" + "=" * 70)
    print("RESULT - Petronas Business Summary, {} to {}".format(from_d, to_d))
    print("=" * 70)
    hdr = "{:<10} {:>12} {:>12} {:>12} {:>14} {:>14} {:>14} {:>14} {:>14}".format(
        "Month", "Purch(Ltr)", "Sales(Ltr)", "Stock", "Sales(Rs)", "Collected(Rs)", "Purch(Rs)", "Payment(Rs)", "ClaimsCN(Rs)")
    print(hdr)
    totals = dict(pl=0.0, sl=0.0, stock=0.0, sv=0.0, coll=0.0, pv=0.0, pay=0.0, cn=0.0)
    for mk in all_months:
        pl = month_purchase_ltr.get(mk, 0.0)
        sl = month_sales_ltr.get(mk, 0.0)
        stock = pl - sl
        sv = month_sales_value.get(mk, 0.0)
        coll = month_collected.get(mk, 0.0)
        pv = month_purchase_value.get(mk, 0.0)
        pay = month_payment_made.get(mk, 0.0)
        cn = month_claims_cn.get(mk, 0.0)
        totals["pl"] += pl; totals["sl"] += sl; totals["stock"] += stock
        totals["sv"] += sv; totals["coll"] += coll; totals["pv"] += pv; totals["pay"] += pay; totals["cn"] += cn
        print("{:<10} {:>12.1f} {:>12.1f} {:>12.1f} {:>14.2f} {:>14.2f} {:>14.2f} {:>14.2f} {:>14.2f}".format(
            _month_label(mk), pl, sl, stock, sv, coll, pv, pay, cn))
    print("{:<10} {:>12.1f} {:>12.1f} {:>12.1f} {:>14.2f} {:>14.2f} {:>14.2f} {:>14.2f} {:>14.2f}".format(
        "Total", totals["pl"], totals["sl"], totals["stock"], totals["sv"], totals["coll"],
        totals["pv"], totals["pay"], totals["cn"]))

    print("\n" + "=" * 70)
    print("Send this whole output back. Check a month or two you know well against your")
    print("own real Tally screens: does Purchase(Rs)/Payment(Rs) match your Ledger")
    print("Vouchers view for '{}'; does Collected(Rs) match".format(PETRONAS_PURCHASE_GROUP_NAME))
    print("the dashboard's own Combined Collection figure for that month; does Sales(Rs)")
    print("look right against what you know was billed - it should now EXCLUDE anything")
    print("billed to 'PETRONAS LUBRICANTS (INDIA) Pvt. Ltd. CHENNAI' as a claim; and does")
    print("ClaimsCN(Rs) match the real Credit Notes you've received from Petronas for")
    print("confirmed claims, if any landed in this date range. Once confirmed, these")
    print("numbers get baked into the dashboard as a static section, matching the")
    print("approved mockup.")


# ── Petronas item-group coverage diagnostic, 1 Oct 2026 ────────────────
#
# Owner sent a real Tally "Stock Group Analysis" screenshot for Sep 2026
# showing TWO rows under Petronas: "PETRONAS TOTAL GROUPS" (Inward 52 NOS
# /100.000 LTR, Outward 10,446 NOS/25,663.900 LTR) and "PETRONAS TOTAL
# GROUPS (OEM)" (Outward only, 77 NOS/271.000 LTR - no Inward row shown),
# and asked directly: "did u adding TATA OEM also? recheck." Every live
# Purchase/Sales Ltr computation in this file (fetch_petronas_summary(),
# report_petronas_summary(), fetch_sales(), dump_daybook(), etc.) scopes
# items to their stock group via _build_item_group_resolver() + an
# `item_grp.strip().lower().startswith("petronas")` check on that item's
# OWN IMMEDIATE PARENT group only (not the full ancestor chain the way
# _resolve_ledgers_under_group() walks for ledgers). Per the 30 Sep 2026
# native-Ltr diagnostic, OEM/TATA items are known to live under groups
# literally named "PETRONAS TATA CBE GROUP"/"PETRONAS TATA OE GROUP" -
# both start with "petronas", so on paper they SHOULD already be
# included. But that was never directly verified against the real group
# master, and the immediate-parent-only check is fragile if Tally's real
# hierarchy nests OEM items one level deeper than expected (e.g. under a
# group that does NOT itself start with "petronas", several levels below
# an ancestor that does) - this diagnostic answers that directly instead
# of guessing. Never pushes anything; prints only.
def dump_petronas_item_group_breakdown(from_date=None, to_date=None):
    """--dump-petronas-item-groups [YYYYMMDD-from] [YYYYMMDD-to] - diagnostic
    only. Defaults to the 1st of this month through today if no dates given.
    Buckets every Purchase/Sales(-PETRONAS/-CBE) voucher's inventory lines by
    the item's OWN IMMEDIATE Tally stock group (no "starts with petronas"
    filter applied here - every group touched is shown), so the real group
    names - and whether each one currently passes fetch_petronas_summary()'s
    own filter - can be checked directly against a real Tally screenshot."""
    now = datetime.now()
    from_d = from_date or now.strftime("%Y%m01")
    to_d = to_date or now.strftime("%Y%m%d")

    print("Fetching stock item group hierarchy (immediate parent only, same")
    print("lookup every live Sales/Purchase feature in this file uses)...")
    item_group = _build_item_group_resolver()
    print("  {} stock item(s) resolved to a group.".format(len(item_group)))

    print("\nFetching ALL vouchers with inventory entries, {} to {} (unfiltered on".format(from_d, to_d))
    print("voucher type - heavy request, please be patient)...")
    inv_xml = tally_request(PETRONAS_INVENTORY_SCAN_TEMPLATE.format(from_date=from_d, to_date=to_d), timeout=600)
    print("  Response received: {} bytes.".format(len(inv_xml)))
    iroot = ET.fromstring(inv_xml)

    sales_vtypes = set(DAYBOOK_VOUCHER_TYPE_NAMES)  # ("sales", "sales - petronas", "sales - cbe", "sales - cash", "pos billing") - 10 Oct 2026

    # {group_name: {"in_qty":.., "in_ltr":.., "in_value":.., "out_qty":.., "out_ltr":.., "out_value":..}}
    groups = {}
    unmatched_items = set()  # items with inventory entries but no group resolved at all

    for v in iroot.iter("VOUCHER"):
        vdate = (_text(v, "DATE") or "").strip()
        if not vdate or vdate < from_d or vdate > to_d:
            continue
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip().lower()
        is_purchase = (vtype in PETRONAS_PURCHASE_VOUCHER_TYPE_NAMES)
        is_sales = (vtype in sales_vtypes)
        if not is_purchase and not is_sales:
            continue
        for entry in v.iter("ALLINVENTORYENTRIES.LIST"):
            item_name = (_text(entry, "STOCKITEMNAME") or "").strip()
            if not item_name:
                continue
            item_grp = item_group.get(item_name, "")
            if not item_grp:
                unmatched_items.add(item_name)
                item_grp = "(no group resolved)"
            raw_actual = _text(entry, "ACTUALQTY")
            raw_billed = _text(entry, "BILLEDQTY")
            qty_num = _num(raw_billed) if raw_billed is not None else _num(raw_actual)
            ltr = _parse_native_ltr(raw_billed) or _parse_native_ltr(raw_actual)
            if ltr is None and qty_num is not None:
                ltr = _catalog_ltr_guess(item_name, qty_num)
            amt = _num(_text(entry, "AMOUNT"))
            row = groups.setdefault(item_grp, dict(in_qty=0.0, in_ltr=0.0, in_value=0.0,
                                                     out_qty=0.0, out_ltr=0.0, out_value=0.0))
            if is_purchase:
                if qty_num is not None:
                    row["in_qty"] += abs(qty_num)
                if ltr is not None:
                    row["in_ltr"] += abs(ltr)
                if amt is not None:
                    row["in_value"] += abs(amt)
            else:
                if qty_num is not None:
                    row["out_qty"] += abs(qty_num)
                if ltr is not None:
                    row["out_ltr"] += abs(ltr)
                if amt is not None:
                    row["out_value"] += abs(amt)

    print("\n" + "=" * 100)
    print("RESULT - every stock group touched by a Purchase/Sales voucher, {} to {}".format(from_d, to_d))
    print("(cross-check this against your real Tally 'Stock Group Analysis' screenshot)")
    print("=" * 100)
    hdr = "{:<42} {:>4} {:>10} {:>10} {:>12}  {:>10} {:>10} {:>12}".format(
        "Group (item's immediate parent)", "In?", "In Qty", "In Ltr", "In Value",
        "Out Qty", "Out Ltr", "Out Value")
    print(hdr)
    totals_included = dict(in_qty=0.0, in_ltr=0.0, in_value=0.0, out_qty=0.0, out_ltr=0.0, out_value=0.0)
    totals_excluded = dict(in_qty=0.0, in_ltr=0.0, in_value=0.0, out_qty=0.0, out_ltr=0.0, out_value=0.0)
    flagged_excluded = []
    for grp in sorted(groups.keys()):
        r = groups[grp]
        included = _is_petronas_item_group(grp)
        policy_excluded = grp.strip().lower() in PETRONAS_ITEM_GROUP_EXCLUSIONS_
        flag = "YES" if included else ("OMIT" if policy_excluded else "no")
        print("{:<42} {:>4} {:>10.1f} {:>10.1f} {:>12.2f}  {:>10.1f} {:>10.1f} {:>12.2f}".format(
            grp[:42], flag, r["in_qty"], r["in_ltr"], r["in_value"],
            r["out_qty"], r["out_ltr"], r["out_value"]))
        target = totals_included if included else totals_excluded
        for k in target:
            target[k] += r[k]
        if not included and not policy_excluded and (
                "petronas" in grp.lower() or "oem" in grp.lower() or "tata" in grp.lower()):
            flagged_excluded.append(grp)

    print("-" * 100)
    print("{:<42} {:>4} {:>10.1f} {:>10.1f} {:>12.2f}  {:>10.1f} {:>10.1f} {:>12.2f}".format(
        "TOTAL - currently INCLUDED", "", totals_included["in_qty"],
        totals_included["in_ltr"], totals_included["in_value"], totals_included["out_qty"],
        totals_included["out_ltr"], totals_included["out_value"]))
    print("{:<42} {:>4} {:>10.1f} {:>10.1f} {:>12.2f}  {:>10.1f} {:>10.1f} {:>12.2f}".format(
        "TOTAL - EXCLUDED (incl. OMIT below)", "", totals_excluded["in_qty"],
        totals_excluded["in_ltr"], totals_excluded["in_value"], totals_excluded["out_qty"],
        totals_excluded["out_ltr"], totals_excluded["out_value"]))
    print("\n'OMIT' = starts with \"petronas\" but is on the owner-confirmed exclusion list")
    print("(" + ", ".join(sorted(PETRONAS_ITEM_GROUP_EXCLUSIONS_)) + ") - a separate business line,")
    print("not counted in Petronas Purchase/Sales Ltr even though the group name starts with petronas.")

    if flagged_excluded:
        print("\n*** FLAG: these groups are currently EXCLUDED from Purchase/Sales Ltr but their")
        print("    name mentions petronas/oem/tata and they're NOT already on the OMIT list -")
        print("    check these specifically: ***")
        for g in flagged_excluded:
            print("    - " + g)
    else:
        print("\nNo other excluded group name mentions petronas/oem/tata - nothing else obviously missed.")

    if unmatched_items:
        print("\n{} item(s) had NO group resolved at all (not in the Stock Item master's group".format(
            len(unmatched_items)))
        print("hierarchy, or a name mismatch) - these are silently skipped by every live feature:")
        for n in sorted(unmatched_items)[:20]:
            print("    - " + n)
        if len(unmatched_items) > 20:
            print("    ... and {} more".format(len(unmatched_items) - 20))

    print("\n" + "=" * 100)
    print("Send this whole output back. Compare the 'TOTAL - currently INCLUDED' Outward")
    print("Qty/Ltr row above against your screenshot's Grand Total Outward row for the same")
    print("date range - if they don't match, the group(s) making up the gap will be sitting")
    print("in the EXCLUDED list (or the *** FLAG *** list) above, by name.")


# ── Anbu / TN Head scope diagnostic - 16 Nov 2026 ──────────────────────
#
# Owner split Anbu (K. Anbumani) out as TN Head, now overseeing the
# WHOLE business across TATA, Maruti, and Coimbatore, not just the
# regular South TN Petronas DSR territory - wants to give him visibility
# into Sales + Purchase for: "PETRONAS COIMBATORE TOTAL AREA SALES",
# "PETRONAS MARUTI GROUP", and the 3 TATA groups (Brake Fluid/CBE/OE)
# combined as one "TATA" bucket, per the owner's own "ALL THIS TATA
# UNDER TATA" instruction. "Petronas Group India" in the owner's list is
# the regular Petronas-branded business already covered by
# report_petronas_summary()/fetch_petronas_summary() above (owner
# confirmed via AskUserQuestion) - no new diagnostic needed for that
# piece, Anbu's eventual dashboard view just reuses that already-live
# data directly. 4 of these 5 groups (everything except Coimbatane) were
# already confirmed to exist in the real Group master back on 1 Oct 2026
# (see PETRONAS_ITEM_GROUP_EXCLUSIONS_ above, built from that exact
# finding) - this diagnostic re-confirms all 5 anyway, since Tally's
# master data can change and this is new, separate-business reporting
# going to someone outside the company, not routine internal use.
# Owner also confirmed (via AskUserQuestion) that TATA/Maruti PURCHASE
# comes through a DIFFERENT supplier ledger than PETRONAS_PURCHASE_
# GROUP_NAME - STEP 3 below self-serve-searches for those, same pattern
# as dump_petronas_purchase()'s STEP 1 candidate search. Diagnostic
# only - nothing pushed, nothing saved to the Sheet, same discipline as
# every other new Tally data type in this project.
ANBU_TN_SCOPE_TARGETS_ = [
    # (bucket label, kind, exact name as given by the owner (lowercased),
    #  candidate search keywords for if it's MISSING)
    #
    # kind="item_group": a PRODUCT-LINE split (Maruti, TATA x3) - resolved
    #   the same way dump_petronas_item_group_breakdown()/
    #   _is_petronas_item_group() already resolve the main Petronas
    #   Sales/Purchase split: by each STOCK ITEM's own immediate-parent
    #   Stock Group (via _build_item_group_resolver(), TYPE=StockItem).
    #   Confirmed real precedent, 1 Oct 2026: "PETRONAS TATA BRAKE FLUID" /
    #   "PETRONAS TATA CBE GROUP" / "PETRONAS TATA OE GROUP" / "PETRONAS
    #   MARUTI GROUP" are all real STOCK ITEM group names in this exact
    #   spelling (see PETRONAS_ITEM_GROUP_EXCLUSIONS_ above) - this is a
    #   DIFFERENT Tally master than the accounting Group master below.
    #
    # kind="ledger_group": an AREA split (Coimbatore) - resolved by WHICH
    #   DEALER (debtor) LEDGERS fall under a named accounting Group (via
    #   GROUP_HIERARCHY_COLLECTION_XML, TYPE=Group + _resolve_ledgers_
    #   under_group()) - the exact same mechanism KNOWN_DSR_GROUP_NAMES
    #   already uses to resolve each DSR's own dealers. "...TOTAL AREA
    #   SALES" names a customer/dealer grouping, not a product line, so it
    #   has no stock-item-group counterpart at all.
    #
    # First real diagnostic run, 16 Nov 2026, confirmed this split is real
    # and necessary, not a guess: STEP 1 originally checked ALL 5 targets
    # against the accounting Group master only. Coimbatore matched there
    # exactly (PARENT: Sundry Debtors - a genuine dealer/area grouping,
    # same shape as every KNOWN_DSR_GROUP_NAMES entry). Maruti/TATA did
    # NOT match there - instead turning up a different, unrelated pair of
    # real accounting groups per line (a "...SALES" debtor group and a
    # "...PURCHASE"/"...PURCHASE A/C" creditor group) - which is exactly
    # what you'd expect if Maruti/TATA's real split lives in the STOCK
    # ITEM group master, not the accounting Group master. Because STEP 2
    # then tried to match stock items' own parent STOCK group against
    # those wrong accounting-Group names, it silently computed zero for
    # every single bucket ("no resolved group... zero matching vouchers")
    # even though real data existed - a genuine diagnostic bug, not a data
    # gap. Fixed below by routing each target through the right master for
    # its own kind.
    ("Coimbatore", "ledger_group", "petronas coimbatore total area sales", ["coimbatore"]),
    ("Maruti", "item_group", "petronas maruti group", ["maruti"]),
    ("TATA", "item_group", "petronas tata brake fluid", ["tata", "brake"]),
    ("TATA", "item_group", "petronas tata cbe group", ["tata", "cbe"]),
    ("TATA", "item_group", "petronas tata oe group", ["tata", "oe"]),
]

# Real accounting-Group candidates surfaced by the 16 Nov 2026 run while
# searching for Maruti/TATA by name in the WRONG (accounting Group) master
# - kept here as named constants because they look like strong leads for
# STEP 3's PURCHASE-ledger search below, not because they're the item-group
# answer. "PETRONAS MARUTI PURCHASE A/C" and "PETRONAS TATA OE PURCHASE" are
# both accounting Groups (the same "behaves like a sub-ledger" shape already
# confirmed for PETRONAS_PURCHASE_GROUP_NAME on 30 Sep 2026), so if either is
# real, _resolve_ledgers_under_group() against it should return the actual
# supplier ledger(s) directly, the same way PETRONAS_PURCHASE_GROUP_NAME
# already does for the main business. UNCONFIRMED - STEP 3 tests these
# live and prints what it finds rather than assuming they're correct.
ANBU_PURCHASE_GROUP_HYPOTHESES_ = {
    "Maruti": "PETRONAS MARUTI PURCHASE A/C",
    "TATA": "PETRONAS TATA OE PURCHASE",
}


def dump_anbu_tn_scope(from_date=None, to_date=None):
    """
    --dump-anbu-tn-scope [YYYYMMDD-from] [YYYYMMDD-to]

    Diagnostic only - see the big comment above ANBU_TN_SCOPE_TARGETS_.
    Defaults to 1 Jan 2026 through today, same convention as
    --dump-petronas-purchase / --petronas-summary.

    Fixed 16 Nov 2026 after the first real run came back with every bucket
    empty - see the comment above ANBU_TN_SCOPE_TARGETS_ for the root cause
    (Maruti/TATA are a STOCK ITEM group split; Coimbatore is an ACCOUNTING
    GROUP / dealer-area split; the original version checked all 5 against
    only the accounting Group master).

    STEP 1a: confirms Maruti/TATA's 3 target names exist as real STOCK
    ITEM groups (each stock item's own immediate-parent group, same method
    dump_petronas_item_group_breakdown() uses) - the right master for a
    product-line split. Any MISSING name gets a keyword candidate search
    over the real distinct stock-group names instead of a blank table.

    STEP 1b: confirms Coimbatore's target name exists as a real ACCOUNTING
    group (GROUP_HIERARCHY_COLLECTION_XML) - the right master for an
    area/dealer split - and prints its PARENT (worth checking whether it
    resolves up to a group literally named "Petronas Group India").

    STEP 2: for Maruti/TATA (item_group), buckets Purchase Ltr / Sales Ltr
    / Sales Value by month using the resolved real stock-group name(s),
    same item-group-matching pattern as report_petronas_summary() above.
    For Coimbatore (ledger_group), resolves every dealer ledger under the
    resolved accounting group (_resolve_ledgers_under_group(), the same
    mechanism KNOWN_DSR_GROUP_NAMES already uses per-DSR) and sums Sales
    for regular Petronas-catalog items (_is_petronas_item_group()) sold to
    those dealers specifically - NOT by item-group, since an area split
    has no stock-item-group counterpart. Purchase in Rupees is deliberately
    NOT computed here for any bucket - that needs a confirmed ledger name
    first, which is what STEP 3 is for. Coimbatore's own Purchase Ltr/Value
    is left blank on principle: a PURCHASE voucher's party is the SUPPLIER,
    not a dealer, so there is no dealer-area way to attribute a purchase to
    "Coimbatore" at all - see the STEP 3 note on this.

    STEP 3: self-serve candidate search for TATA and Maruti PURCHASE/
    supplier ledgers, since the owner confirmed these are NOT
    PETRONAS_PURCHASE_GROUP_NAME. First tests the two accounting-Group
    hypotheses surfaced incidentally by STEP 1b's own candidate search
    (ANBU_PURCHASE_GROUP_HYPOTHESES_) by resolving real ledgers under each,
    same mechanism as the confirmed PETRONAS_PURCHASE_GROUP_NAME Group.
    Then falls back to the original raw ledger-name keyword search for
    anything those two hypotheses miss (in particular: TATA Brake Fluid /
    CBE have no distinct accounting Group of their own at all, so they may
    route through PETRONAS TATA OE PURCHASE combined, or through its own
    parent "PETRONAS WS-AM PURCHASE" - flagged, not assumed). Coimbatore is
    still assumed to share the regular Petronas purchase ledger (it's a
    sales-AREA split of the same branded stock, going by its own name -
    "...TOTAL AREA SALES" - not a separate supplier) - flagged for the
    owner to confirm or correct, not assumed silently.
    """
    now = datetime.now()
    from_d = from_date or "20260101"
    to_d = to_date or now.strftime("%Y%m%d")

    print("=" * 70)
    print("ANBU / TN HEAD SCOPE - DIAGNOSTIC ONLY - {} to {}".format(from_d, to_d))
    print("Nothing in this run is pushed anywhere or saved to the Sheet.")
    print("=" * 70)

    item_targets = [(b, t, kw) for b, k, t, kw in ANBU_TN_SCOPE_TARGETS_ if k == "item_group"]
    ledger_targets = [(b, t, kw) for b, k, t, kw in ANBU_TN_SCOPE_TARGETS_ if k == "ledger_group"]

    # STEP 1a (item_group targets: Maruti, TATA) ---------------------
    print("\n" + "=" * 70)
    print("STEP 1a: do Maruti/TATA exist as real STOCK ITEM groups? (product-line split)")
    print("=" * 70)
    item_group = _build_item_group_resolver()
    distinct_stock_groups = sorted(set(v for v in item_group.values() if v))
    distinct_stock_groups_lower = {g.lower(): g for g in distinct_stock_groups}

    resolved_item_group_names = {}  # bucket -> set of real stock-group names
    for bucket, target_lower, search_keywords in item_targets:
        real = distinct_stock_groups_lower.get(target_lower)
        if real:
            print("  FOUND   [{}] {!r}  (real Stock Item group)".format(bucket, real))
            resolved_item_group_names.setdefault(bucket, set()).add(real)
        else:
            print("  MISSING [{}] {!r} - not found among {} distinct Stock Item groups".format(
                bucket, target_lower, len(distinct_stock_groups)))
            candidates = [g for g in distinct_stock_groups
                          if all(kw in g.lower() for kw in search_keywords)]
            if candidates:
                print("    {} possible candidate(s) (name contains every one of {}):".format(
                    len(candidates), search_keywords))
                for g in candidates[:10]:
                    print("      {!r}".format(g))
                if len(candidates) == 1:
                    resolved_item_group_names.setdefault(bucket, set()).add(candidates[0])
                    print("    Exactly one candidate - using it for STEP 2 below, but CONFIRM this is")
                    print("    really the right group before trusting the numbers.")
            else:
                print("    No candidates found containing all of {} either - send back the exact".format(
                    search_keywords))
                print("    name as Tally's Stock Group master (item-level) shows it.")

    # STEP 1b (ledger_group target: Coimbatore) -----------------------
    print("\n" + "=" * 70)
    print("STEP 1b: does Coimbatore exist as a real ACCOUNTING group? (dealer/area split)")
    print("=" * 70)
    group_xml = tally_request(GROUP_HIERARCHY_COLLECTION_XML)
    group_root = ET.fromstring(group_xml)
    all_groups = []  # (name, parent)
    for grp in group_root.iter("GROUP"):
        gname = grp.get("NAME") or _text(grp, "NAME")
        if gname:
            all_groups.append((gname.strip(), (_text(grp, "PARENT") or "").strip()))

    resolved_ledger_group_names = {}  # bucket -> set of real accounting-group names
    for bucket, target_lower, search_keywords in ledger_targets:
        found = None
        for n, p in all_groups:
            if n.lower() == target_lower:
                found = (n, p)
                break
        if found:
            real_name, parent = found
            flag_note = ""
            if parent.strip().lower() == "petronas group india":
                flag_note = "  <- PARENT IS LITERALLY 'Petronas Group India' - confirms that's a real parent group"
            print("  FOUND   [{}] {!r}  (PARENT: {}){}".format(
                bucket, real_name, parent or "(top-level)", flag_note))
            resolved_ledger_group_names.setdefault(bucket, set()).add(real_name)
        else:
            print("  MISSING [{}] {!r} - not found (exact match, case-insensitive)".format(bucket, target_lower))
            candidates = [(n, p) for n, p in all_groups
                          if all(kw in n.lower() for kw in search_keywords)]
            if candidates:
                print("    {} possible candidate(s) (name contains every one of {}):".format(
                    len(candidates), search_keywords))
                for n, p in candidates[:10]:
                    print("      {!r}  (PARENT: {})".format(n, p or "(top-level)"))
                if len(candidates) == 1:
                    real_name, parent = candidates[0]
                    resolved_ledger_group_names.setdefault(bucket, set()).add(real_name)
                    print("    Exactly one candidate - using it for STEP 2 below, but CONFIRM this is")
                    print("    really the right group before trusting the numbers.")
            else:
                print("    No candidates found containing all of {} either - send back the exact".format(
                    search_keywords))
                print("    name as Tally's accounting Group master shows it.")

    # Also surface the two accounting-Group PURCHASE hypotheses right here,
    # since they came from this same master - STEP 3 below tests them for real.
    maruti_tata_purchase_groups_seen = {}
    for bucket_label, hyp_name in ANBU_PURCHASE_GROUP_HYPOTHESES_.items():
        for n, p in all_groups:
            if n.strip().lower() == hyp_name.strip().lower():
                maruti_tata_purchase_groups_seen[bucket_label] = (n, p)
                print("  (incidental) [{}] accounting Group {!r} exists (PARENT: {}) - a candidate".format(
                    bucket_label, n, p or "(top-level)"))
                print("  PURCHASE-side group for STEP 3 to test below, same shape as the confirmed")
                print("  {!r} Group.".format(PETRONAS_PURCHASE_GROUP_NAME))
                break

    # 17 Nov 2026 - a real run's raw ledger-keyword search turned up TWO
    # "coimbatore"+"purchase" accounting Groups, not one: "COIMBATORE AREA
    # PURCHASE" (33 ledgers - all agro/farm-implement companies, AAA's
    # OTHER, unrelated business line in that area) and "COIMBATORE PETRONAS
    # LUBRICANTS PURCHASE" (5 ledgers, all real regional "PETRONAS
    # LUBRICANTS (I/INDIA)..." billing entities - the exact same shape as
    # every other confirmed Petronas purchase group in this file). Narrowed
    # the auto-resolve to candidates whose name ALSO contains "petronas",
    # so the real one is used automatically and the irrelevant one is
    # still shown (for transparency) but never silently trusted.
    coimbatore_purchase_candidates_all = [(n, p) for n, p in all_groups
                                           if "coimbatore" in n.lower() and "purchase" in n.lower()]
    coimbatore_purchase_group_name = None
    coimbatore_purchase_ledgers = set()
    if coimbatore_purchase_candidates_all:
        petronas_candidates = [(n, p) for n, p in coimbatore_purchase_candidates_all
                                if "petronas" in n.lower()]
        print("\n  (incidental) [Coimbatore] found {} accounting Group(s) with 'coimbatore' AND".format(
            len(coimbatore_purchase_candidates_all)))
        print("  'purchase' in the name:")
        for n, p in coimbatore_purchase_candidates_all:
            tag = "  <- also contains 'petronas', auto-using this one" if "petronas" in n.lower() else ""
            print("    {!r}  (PARENT: {}){}".format(n, p or "(top-level)", tag))
        if len(petronas_candidates) == 1:
            coimbatore_purchase_group_name, cpg_parent = petronas_candidates[0]
            coimbatore_purchase_ledgers = _resolve_ledgers_under_group(coimbatore_purchase_group_name)
            print("  Resolved {} ledger(s) under it - using this for Coimbatore's Purchase in STEP 2,".format(
                len(coimbatore_purchase_ledgers)))
            print("  overriding the 'shares the main ledger' default assumption. CONFIRM this is right.")
            for n in sorted(coimbatore_purchase_ledgers)[:20]:
                print("    {!r}".format(n))
        elif len(petronas_candidates) > 1:
            print("  More than one candidate also contains 'petronas' - not auto-resolving, picking")
            print("  blindly here would be a guess. CONFIRM which one (if any) is right.")
        else:
            print("  None of these also contain 'petronas' - not auto-resolving. Coimbatore's Purchase")
            print("  stays on the default 'shares the main ledger' assumption below unless corrected.")

    if not resolved_item_group_names and not resolved_ledger_group_names:
        print("\nZero of the 5 groups resolved - stopping here, send this output back first.")
        return

    # STEP 2 ---------------------------------------------------------
    print("\n" + "=" * 70)
    print("STEP 2: month-wise Purchase Ltr / Sales Ltr & Value for each resolved bucket")
    print("        (fetching ALL vouchers company-wide with inventory entries,")
    print("        unfiltered on voucher type - heavy request, please be patient)")
    print("=" * 70)

    item_to_bucket = {}
    for bucket, names in resolved_item_group_names.items():
        for n in names:
            item_to_bucket[n.lower()] = bucket

    coimbatore_dealers = set()
    if resolved_ledger_group_names.get("Coimbatore"):
        for real_name in resolved_ledger_group_names["Coimbatore"]:
            print("  Resolving dealer ledgers under {!r}...".format(real_name))
            coimbatore_dealers |= _resolve_ledgers_under_group(real_name)
        print("  {} dealer ledger(s) resolve to Coimbatore.".format(len(coimbatore_dealers)))

    inv_xml = tally_request(PETRONAS_INVENTORY_SCAN_TEMPLATE.format(from_date=from_d, to_date=to_d), timeout=600)
    print("  Response received: {} bytes.".format(len(inv_xml)))
    # 16 Nov 2026 - a real run came back with only 1532 bytes here (every
    # other confirmed run of this exact request, unfiltered company-wide,
    # has been tens to hundreds of MB) - genuinely too small to be real
    # data, most likely a transient Tally-side hiccup (this request is the
    # heaviest one in the whole script and runs after 3-4 lighter requests
    # already fired in STEP 1a/1b/the incidental purchase-group check).
    # Retried once automatically rather than failing the whole run on what
    # is probably a one-off; if the retry is ALSO small, the raw text is
    # printed so the real Tally-side error (if any) is visible instead of
    # guessing blind a second time.
    if len(inv_xml) < 50000:
        print("  That's a suspiciously small response for an unfiltered company-wide fetch")
        print("  (every confirmed run of this exact request has been tens-to-hundreds of MB) -")
        print("  retrying once automatically...")
        inv_xml = tally_request(PETRONAS_INVENTORY_SCAN_TEMPLATE.format(from_date=from_d, to_date=to_d), timeout=600)
        print("  Retry response: {} bytes.".format(len(inv_xml)))
        if len(inv_xml) < 50000:
            print("  Still small after a retry. Raw response below so the real Tally-side")
            print("  message (if any) is visible, instead of guessing why STEP 2 came back empty:")
            print("  " + "-" * 66)
            print("  " + inv_xml[:2000])
            print("  " + "-" * 66)
    iroot = ET.fromstring(inv_xml)

    sales_vtypes = set(DAYBOOK_VOUCHER_TYPE_NAMES)
    # {bucket: {month_key: {"purchase_ltr":.., "sales_ltr":.., "sales_value":..}}}
    bucket_month = {}
    # 17 Nov 2026 - raw (unfiltered-by-item-group) diagnostic counters for
    # Coimbatore's resolved Purchase ledgers, so a 0.0 Purchase total below
    # can be told apart from "these ledgers have zero voucher activity at
    # all in this range" vs "they have activity but it's being excluded by
    # the _is_petronas_item_group() filter" - the real 17 Nov run showed
    # 0.0 despite 5 resolved ledgers, which needs this to diagnose.
    coimb_purch_raw_voucher_count = 0
    coimb_purch_raw_item_groups_seen = {}

    for v in iroot.iter("VOUCHER"):
        vdate = (_text(v, "DATE") or "").strip()
        if not vdate or vdate < from_d or vdate > to_d:
            continue
        mk = _month_key(vdate)
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip().lower()
        party = (_text(v, "PARTYLEDGERNAME") or "").strip()
        is_purchase = (vtype in PETRONAS_PURCHASE_VOUCHER_TYPE_NAMES)
        is_sales = (vtype in sales_vtypes)
        if not is_purchase and not is_sales:
            continue
        is_coimbatore_sale = (is_sales and party in coimbatore_dealers)
        is_coimbatore_purchase = (is_purchase and party in coimbatore_purchase_ledgers)
        if is_coimbatore_purchase:
            coimb_purch_raw_voucher_count += 1
        if not is_coimbatore_sale and not is_coimbatore_purchase and not item_to_bucket:
            continue
        for entry in v.iter("ALLINVENTORYENTRIES.LIST"):
            item_name = (_text(entry, "STOCKITEMNAME") or "").strip()
            if not item_name:
                continue
            item_grp = item_group.get(item_name, "")
            raw_actual = _text(entry, "ACTUALQTY")
            raw_billed = _text(entry, "BILLEDQTY")
            qty_num = _num(raw_billed) if raw_billed is not None else _num(raw_actual)
            ltr = _parse_native_ltr(raw_billed) or _parse_native_ltr(raw_actual)
            if ltr is None and qty_num is not None:
                ltr = _catalog_ltr_guess(item_name, qty_num)
            amt = _num(_text(entry, "AMOUNT"))

            if is_coimbatore_purchase:
                grp_key = item_grp or "(no group resolved)"
                coimb_purch_raw_item_groups_seen[grp_key] = coimb_purch_raw_item_groups_seen.get(grp_key, 0) + 1

            # Coimbatore: a dealer-area bucket, so match on WHO bought it
            # (party in the resolved dealer set) + regular Petronas-catalog
            # scope (_is_petronas_item_group()), not on the item's own
            # parent group matching "Coimbatore" by name (it never will -
            # Coimbatore has no stock-item-group counterpart at all).
            if is_coimbatore_sale and _is_petronas_item_group(item_grp):
                row = bucket_month.setdefault("Coimbatore", {}).setdefault(
                    mk, dict(purchase_ltr=0.0, sales_ltr=0.0, sales_value=0.0))
                if ltr is not None:
                    row["sales_ltr"] += abs(ltr)
                if amt is not None:
                    row["sales_value"] += round(abs(amt) * 1.18, 2)

            # Coimbatore Purchase - only computed when STEP 1b's incidental
            # search found exactly one 'coimbatore'+'purchase'+'petronas'
            # accounting Group (coimbatore_purchase_ledgers non-empty);
            # otherwise this never fires and Purchase stays 0/blank, same
            # as before this was added.
            if is_coimbatore_purchase and _is_petronas_item_group(item_grp):
                row = bucket_month.setdefault("Coimbatore", {}).setdefault(
                    mk, dict(purchase_ltr=0.0, sales_ltr=0.0, sales_value=0.0))
                if ltr is not None:
                    row["purchase_ltr"] += abs(ltr)

            # Maruti/TATA: a product-line bucket, so match on the item's
            # own parent stock group, same as the main Petronas split.
            item_bucket = item_to_bucket.get(item_grp.lower())
            if item_bucket:
                row = bucket_month.setdefault(item_bucket, {}).setdefault(
                    mk, dict(purchase_ltr=0.0, sales_ltr=0.0, sales_value=0.0))
                if is_purchase:
                    if ltr is not None:
                        row["purchase_ltr"] += abs(ltr)
                elif is_sales:
                    if ltr is not None:
                        row["sales_ltr"] += abs(ltr)
                    if amt is not None:
                        row["sales_value"] += round(abs(amt) * 1.18, 2)

    for bucket in ("Coimbatore", "Maruti", "TATA"):
        if bucket not in bucket_month:
            print("\n  [{}] - no resolved group for this bucket, or zero matching vouchers in range.".format(bucket))
            continue
        print("\n  [{}]".format(bucket))
        if bucket == "Coimbatore":
            # 17 Nov 2026 - the real run showed 'COIMBATORE PETRONAS LUBRICANTS
            # PURCHASE' (5 real ledgers) has ZERO Purchase vouchers at all in
            # range - genuinely dormant, not just filtered out. Confirmed
            # dormant -> fall back to the ORIGINAL default assumption (shares
            # the main company-wide ledger, not separately splittable) rather
            # than printing a falsely-confident "computed from N ledgers" note
            # for a number that is 0 only because the source is unused.
            if coimbatore_purchase_group_name and coimb_purch_raw_voucher_count == 0:
                print("  (Purchase Ltr 0/blank here - the resolved candidate {!r}".format(
                    coimbatore_purchase_group_name))
                print("  ({} ledger(s)) has ZERO Purchase vouchers at all in this date range, confirmed".format(
                    len(coimbatore_purchase_ledgers)))
                print("  via the raw diagnostic below - it's genuinely dormant, not a real Purchase".format())
                print("  source this period. Falling back to the default assumption: Coimbatore shares")
                print("  the main {!r} ledger like the rest of the business, so its".format(
                    PETRONAS_PURCHASE_GROUP_NAME))
                print("  Purchase isn't separately splittable from the company-wide total.)")
            elif coimbatore_purchase_group_name:
                print("  (Purchase Ltr computed from {} resolved ledger(s) under {!r} -".format(
                    len(coimbatore_purchase_ledgers), coimbatore_purchase_group_name))
                print("  CONFIRM this is really the right Purchase source for Coimbatore.)")
            else:
                print("  (Purchase Ltr intentionally 0/blank here - no single confirmed Coimbatore")
                print("  Purchase ledger/Group yet; a Purchase voucher's party is the SUPPLIER, not")
                print("  a dealer, so there's no dealer-area way to infer one without a name match.)")
        print("  {:<10} {:>14} {:>14} {:>18}".format("Month", "Purch(Ltr)", "Sales(Ltr)", "Sales(Rs incl GST)"))
        months = sorted(bucket_month[bucket].keys())
        tot = dict(purchase_ltr=0.0, sales_ltr=0.0, sales_value=0.0)
        for mk in months:
            r = bucket_month[bucket][mk]
            for k in tot:
                tot[k] += r[k]
            print("  {:<10} {:>14.1f} {:>14.1f} {:>18.2f}".format(
                _month_label(mk), r["purchase_ltr"], r["sales_ltr"], r["sales_value"]))
        print("  {:<10} {:>14.1f} {:>14.1f} {:>18.2f}".format(
            "Total", tot["purchase_ltr"], tot["sales_ltr"], tot["sales_value"]))
        if bucket == "Coimbatore" and coimbatore_purchase_group_name and tot["purchase_ltr"] == 0.0:
            print("\n  (diagnostic) Coimbatore's resolved Purchase ledgers show 0.0 Ltr above - raw check")
            print("  (ignoring the Petronas-item-group filter) of vouchers against those same ledgers:")
            if coimb_purch_raw_voucher_count == 0:
                print("    0 Purchase vouchers at all found against these {} ledger(s) in this date".format(
                    len(coimbatore_purchase_ledgers)))
                print("    range - they look genuinely unused/dormant for this period. Coimbatore's real")
                print("    Purchase may still be booked through the shared main ledger after all (the")
                print("    original default assumption) - CONFIRM with the owner before trusting 0.0.")
            else:
                print("    {} Purchase voucher entries found against these ledgers, but their item's".format(
                    coimb_purch_raw_voucher_count))
                print("    resolved Stock Item group(s) didn't start with 'petronas' so they were excluded:")
                for grp, cnt in sorted(coimb_purch_raw_item_groups_seen.items(), key=lambda kv: -kv[1])[:10]:
                    print("      {!r}: {} entrie(s)".format(grp, cnt))
                print("    If one of these SHOULD count as Petronas stock, that's a real gap in")
                print("    _is_petronas_item_group()/the item master, not a 0 Purchase month.")

    # STEP 3 ---------------------------------------------------------
    print("\n" + "=" * 70)
    print("STEP 3: self-serve candidate search for TATA / Maruti PURCHASE (supplier) ledgers")
    print("        (owner confirmed these are a DIFFERENT supplier than {!r})".format(
        PETRONAS_PURCHASE_GROUP_NAME))
    print("=" * 70)

    for bucket_label, hyp_name in ANBU_PURCHASE_GROUP_HYPOTHESES_.items():
        if bucket_label in maruti_tata_purchase_groups_seen:
            real_name, parent = maruti_tata_purchase_groups_seen[bucket_label]
            resolved = _resolve_ledgers_under_group(real_name)
            print("\n  [{}] testing accounting Group {!r} (PARENT: {}):".format(
                bucket_label, real_name, parent or "(top-level)"))
            if resolved:
                print("    {} ledger(s) resolve under it:".format(len(resolved)))
                for n in sorted(resolved)[:20]:
                    print("      {!r}".format(n))
            else:
                print("    0 ledgers resolve under it directly - may still be right if it carries")
                print("    its own balance as a 'behaves like a sub-ledger' Group with no children.")
        else:
            print("\n  [{}] hypothesis group {!r} was not found in STEP 1b's accounting-Group".format(
                bucket_label, hyp_name))
            print("    scan this run - ignore, or check the raw keyword list below.")

    ledger_xml = tally_request(LEDGER_GROUP_COLLECTION_XML)
    ledger_root = ET.fromstring(ledger_xml)
    all_ledgers = []
    for led in ledger_root.iter("LEDGER"):
        lname = led.get("NAME") or _text(led, "NAME")
        if not lname:
            continue
        lname = lname.strip()
        lparent = (_text(led, "PARENT") or "").strip()
        lclosing = _text(led, "CLOSINGBALANCE")
        all_ledgers.append((lname, lparent, lclosing))

    for keyword in ("tata", "maruti"):
        print("\n  Raw ledger-name search, ledgers containing {!r}:".format(keyword))
        candidates = [(n, p, c) for n, p, c in all_ledgers if keyword in n.lower()]
        if candidates:
            for n, p, c in candidates[:20]:
                flag = ""
                if "creditor" in p.lower() or "supplier" in p.lower() or "purchase" in p.lower():
                    flag = "  <- PARENT mentions creditor/supplier/purchase, likely candidate"
                print("    {!r}  (PARENT: {}, CLOSINGBALANCE: {!r}){}".format(
                    n, p or "(top-level)", c, flag))
        else:
            print("    None found containing {!r}.".format(keyword))

    print("\n  TATA Brake Fluid / CBE have no distinct accounting Group of their own in")
    print("  this master (confirmed above) - if their real Purchase doesn't fall under")
    print("  'PETRONAS TATA OE PURCHASE' combined, it may sit under that group's own")
    print("  PARENT ('PETRONAS WS-AM PURCHASE' per STEP 1b, if seen) instead - worth")
    print("  checking that one too if the OE-only figure looks short.")

    # 17 Nov 2026 - Coimbatore's own candidate search/resolution now happens
    # up in STEP 1b (where the 'coimbatore'+'purchase' Groups were found),
    # and STEP 2 above already used whatever it resolved - this is just a
    # plain-language recap, not a second search/request.
    if coimbatore_purchase_group_name and coimb_purch_raw_voucher_count == 0:
        print("\n  Coimbatore: STEP 1b resolved {!r} ({} ledger(s)),".format(
            coimbatore_purchase_group_name, len(coimbatore_purchase_ledgers)))
        print("  but STEP 2's raw diagnostic above found ZERO Purchase vouchers against it in")
        print("  range - confirmed dormant, not a real Purchase source this period. Fell back to")
        print("  the default assumption below (shares the main company-wide ledger).")
    elif coimbatore_purchase_group_name:
        print("\n  Coimbatore: STEP 1b already resolved and used {!r}".format(
            coimbatore_purchase_group_name))
        print("  ({} ledger(s)) for Coimbatore's Purchase in STEP 2 above - CONFIRM this is".format(
            len(coimbatore_purchase_ledgers)))
        print("  really the right Purchase source, since it overrides the default assumption below.")
    elif coimbatore_purchase_candidates_all:
        print("\n  Coimbatore: STEP 1b found {} 'coimbatore'+'purchase' accounting Group(s) but".format(
            len(coimbatore_purchase_candidates_all)))
        print("  could not auto-resolve one (see STEP 1b above) - Purchase stays on the default")
        print("  assumption below unless you pick the right one from that list.")
    else:
        print("\n  Coimbatore: no distinct 'coimbatore'+'purchase' accounting Group found -")
        print("  no contradiction of the default assumption below turned up this run.")

    print("\n  Coimbatore (default assumption, only used if nothing above overrode it): it's a")
    print("  SALES-area split (the name itself says 'TOTAL AREA SALES'), assumed to draw from")
    print("  the SAME {!r} purchase ledger as the rest of the regular".format(PETRONAS_PURCHASE_GROUP_NAME))
    print("  Petronas business, not a separate supplier.")

    print("\n" + "=" * 70)
    print("Send this whole output back. Check STEP 1a/1b's resolved names, STEP 2's")
    print("month-wise figures against what you know, and pick the right ledger/Group from")
    print("STEP 3's candidates (or confirm none of them are right, or that Coimbatore's")
    print("assumption above is wrong) before anything gets built live for Anbu's login.")


# ── Anbu / TN Head scope - LIVE, full month-wise history - added 17 Nov
# 2026 ──
#
# The real thing, finally - after 3 rounds of --dump-anbu-tn-scope against
# real Tally data (16-17 Nov 2026, see the big comment above
# ANBU_TN_SCOPE_TARGETS_ and dump_anbu_tn_scope()'s own docstring for the
# full diagnostic history), these 3 buckets are CONFIRMED:
#   - Maruti: a real STOCK ITEM group, "PETRONAS MARUTI GROUP".
#   - TATA (combined): 3 real STOCK ITEM groups - "PETRONAS TATA BRAKE
#     FLUID", "PETRONAS TATA CBE GROUP", "PETRONAS TATA OE GROUP".
#   - Coimbatore: a real ACCOUNTING group (dealer-area split),
#     "PETRONAS COIMBATORE TOTAL AREA SALES" - SALES only. Purchase is
#     deliberately NOT computed for Coimbatore: the one real-looking
#     purchase-group candidate ("COIMBATORE PETRONAS LUBRICANTS
#     PURCHASE", 5 real regional ledgers) was confirmed DORMANT on the
#     17 Nov 2026 real run (zero Purchase vouchers against it across the
#     whole Jan-Nov 2026 range) - Coimbatore shares the same main
#     PETRONAS_PURCHASE_GROUP_NAME ledger as the rest of the business,
#     with no area tag to split a Coimbatore-only figure back out of it.
#
# Owner's own choice (17 Nov 2026, asked via AskUserQuestion before this
# was built): show the FULL month-wise history (from
# ANBU_TN_SCOPE_LIVE_START_DATE_ through today), not just the current
# month like fetch_petronas_summary() - so this function re-scans the
# WHOLE range every push, not just a 1-2 month window. Pushed to its own
# LiveAnbuTnScope sheet (one row per Bucket+Month, upserted - see
# Code.gs's handleLiveAnbuTnScopeUpdate_()). On the dashboard this is
# shown ONLY to Anbu and Alagu (Owner), gated by NAME not role - Bala
# shares Anbu's "Manager" role but was never asked for this panel (the
# owner's own second AskUserQuestion answer, 17 Nov 2026).
ANBU_TN_SCOPE_LIVE_START_DATE_ = "20260101"

# The now-CONFIRMED real names, used directly - no more self-serve
# searching here (STEP 1a/1b/3 in dump_anbu_tn_scope() already did that
# job; this is the live path, same "confirm via diagnostic, then hardcode
# it" convention as PETRONAS_PURCHASE_GROUP_NAME/PETRONAS_ITEM_GROUP_
# EXCLUSIONS_ elsewhere in this file).
ANBU_TN_SCOPE_CONFIRMED_ = {
    "Coimbatore": {"kind": "ledger_group", "names": ["PETRONAS COIMBATORE TOTAL AREA SALES"]},
    "Maruti": {"kind": "item_group", "names": ["PETRONAS MARUTI GROUP"]},
    "TATA": {"kind": "item_group", "names": [
        "PETRONAS TATA BRAKE FLUID", "PETRONAS TATA CBE GROUP", "PETRONAS TATA OE GROUP"]},
}


def fetch_anbu_tn_scope():
    """Computes the full month-wise Purchase Ltr / Sales Ltr / Sales
    Value history (ANBU_TN_SCOPE_LIVE_START_DATE_ through today) for the
    3 confirmed buckets - Coimbatore (Sales only - see the big comment
    above for why Purchase is never computed for it), Maruti, TATA
    (combined across its 3 item groups). Returns a flat list of row
    dicts: [{"bucket", "month", "monthLabel", "purchaseLtr" (None for
    every Coimbatore row), "salesLtr", "salesValue"}, ...], one per
    bucket per month that has any matching voucher - mirrors
    dump_anbu_tn_scope()'s own STEP 2 computation exactly (same
    resolution mechanism, same _is_petronas_item_group() filter, same
    retry-once-on-small-response safeguard), just against the confirmed
    names directly and returning data instead of printing it."""
    from_d = ANBU_TN_SCOPE_LIVE_START_DATE_
    to_d = datetime.now().strftime("%Y%m%d")

    item_group = _build_item_group_resolver()
    item_to_bucket = {}
    for name in ANBU_TN_SCOPE_CONFIRMED_["Maruti"]["names"]:
        item_to_bucket[name.lower()] = "Maruti"
    for name in ANBU_TN_SCOPE_CONFIRMED_["TATA"]["names"]:
        item_to_bucket[name.lower()] = "TATA"

    coimbatore_dealers = _resolve_ledgers_under_group(
        ANBU_TN_SCOPE_CONFIRMED_["Coimbatore"]["names"][0])

    inv_xml = tally_request(PETRONAS_INVENTORY_SCAN_TEMPLATE.format(from_date=from_d, to_date=to_d), timeout=600)
    if len(inv_xml) < 50000:
        # Same transient-small-response safeguard as dump_anbu_tn_scope()'s
        # STEP 2 - see that function's own 16 Nov 2026 comment for why a
        # retry (rather than failing/guessing) is the right response here.
        inv_xml = tally_request(PETRONAS_INVENTORY_SCAN_TEMPLATE.format(from_date=from_d, to_date=to_d), timeout=600)
    iroot = ET.fromstring(inv_xml)

    sales_vtypes = set(DAYBOOK_VOUCHER_TYPE_NAMES)
    bucket_month = {}

    for v in iroot.iter("VOUCHER"):
        vdate = (_text(v, "DATE") or "").strip()
        if not vdate or vdate < from_d or vdate > to_d:
            continue
        mk = _month_key(vdate)
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip().lower()
        party = (_text(v, "PARTYLEDGERNAME") or "").strip()
        is_purchase = (vtype in PETRONAS_PURCHASE_VOUCHER_TYPE_NAMES)
        is_sales = (vtype in sales_vtypes)
        if not is_purchase and not is_sales:
            continue
        is_coimbatore_sale = (is_sales and party in coimbatore_dealers)
        if not is_coimbatore_sale and not item_to_bucket:
            continue
        for entry in v.iter("ALLINVENTORYENTRIES.LIST"):
            item_name = (_text(entry, "STOCKITEMNAME") or "").strip()
            if not item_name:
                continue
            item_grp = item_group.get(item_name, "")
            raw_actual = _text(entry, "ACTUALQTY")
            raw_billed = _text(entry, "BILLEDQTY")
            qty_num = _num(raw_billed) if raw_billed is not None else _num(raw_actual)
            ltr = _parse_native_ltr(raw_billed) or _parse_native_ltr(raw_actual)
            if ltr is None and qty_num is not None:
                ltr = _catalog_ltr_guess(item_name, qty_num)
            amt = _num(_text(entry, "AMOUNT"))

            # Coimbatore - Sales only, matched by dealer (area split), same
            # as dump_anbu_tn_scope()'s STEP 2. Purchase is never touched
            # here - see the big comment above this function for why.
            if is_coimbatore_sale and _is_petronas_item_group(item_grp):
                row = bucket_month.setdefault("Coimbatore", {}).setdefault(
                    mk, dict(purchase_ltr=None, sales_ltr=0.0, sales_value=0.0))
                if ltr is not None:
                    row["sales_ltr"] += abs(ltr)
                if amt is not None:
                    row["sales_value"] += round(abs(amt) * 1.18, 2)

            # Maruti/TATA - Purchase + Sales, matched by the item's own
            # stock-item group (product-line split).
            item_bucket = item_to_bucket.get(item_grp.lower())
            if item_bucket:
                row = bucket_month.setdefault(item_bucket, {}).setdefault(
                    mk, dict(purchase_ltr=0.0, sales_ltr=0.0, sales_value=0.0))
                if is_purchase:
                    if ltr is not None:
                        row["purchase_ltr"] += abs(ltr)
                elif is_sales:
                    if ltr is not None:
                        row["sales_ltr"] += abs(ltr)
                    if amt is not None:
                        row["sales_value"] += round(abs(amt) * 1.18, 2)

    rows = []
    for bucket in ("Coimbatore", "Maruti", "TATA"):
        for mk in sorted(bucket_month.get(bucket, {}).keys()):
            r = bucket_month[bucket][mk]
            rows.append({
                "bucket": bucket,
                "month": mk,
                "monthLabel": _month_label(mk),
                "purchaseLtr": None if r["purchase_ltr"] is None else round(r["purchase_ltr"], 1),
                "salesLtr": round(r["sales_ltr"], 1),
                "salesValue": round(r["sales_value"], 2),
            })
    return rows


def dump_anbu_tn_scope_live_check():
    """--dump-anbu-tn-scope-live-check - diagnostic only, pushes nothing.
    Runs the new LIVE fetch_anbu_tn_scope() (the confirmed-names version,
    no self-serve searching) and prints it in the same per-bucket table
    shape as dump_anbu_tn_scope()'s STEP 2, so the owner can confirm
    switching to the hardcoded confirmed names didn't change any figure,
    before trusting this to actually push to the dashboard."""
    rows = fetch_anbu_tn_scope()
    by_bucket = {}
    for r in rows:
        by_bucket.setdefault(r["bucket"], []).append(r)
    for bucket in ("Coimbatore", "Maruti", "TATA"):
        brows = by_bucket.get(bucket, [])
        print("\n  [{}]".format(bucket))
        print("  {:<10} {:>14} {:>14} {:>18}".format("Month", "Purch(Ltr)", "Sales(Ltr)", "Sales(Rs incl GST)"))
        tot_p, tot_s, tot_v, any_p = 0.0, 0.0, 0.0, False
        for r in brows:
            p = r["purchaseLtr"]
            print("  {:<10} {:>14} {:>14.1f} {:>18.2f}".format(
                r["monthLabel"], ("—" if p is None else "{:.1f}".format(p)), r["salesLtr"], r["salesValue"]))
            if p is not None:
                tot_p += p
                any_p = True
            tot_s += r["salesLtr"]
            tot_v += r["salesValue"]
        print("  {:<10} {:>14} {:>14.1f} {:>18.2f}".format(
            "Total", ("—" if not any_p else "{:.1f}".format(tot_p)), tot_s, tot_v))


# --push-anbu-tn-scope-now / the watcher's own cadence - added 17 Nov 2026.
# A full-year unfiltered inventory scan every cycle (same heavy single
# request dump_anbu_tn_scope()'s STEP 2 already does) - same order of
# heaviness as Group Summary/Petronas Business Summary, so it gets a slow,
# few-times-a-day cadence rather than anything more frequent. 12h (2x/day)
# chosen as the starting cadence since this is a reference panel for
# Anbu/Owner, not one of the main time-critical business figures - change
# ANBU_TN_SCOPE_POLL_INTERVAL_SECONDS if the owner wants it fresher.
ANBU_TN_SCOPE_POLL_INTERVAL_SECONDS = 12 * 60 * 60


def maybe_push_anbu_tn_scope(state, dry_run, ts):
    if dry_run:
        return state
    last = state.get("last_anbu_tn_scope_push_ts", 0)
    if time.time() - last < ANBU_TN_SCOPE_POLL_INTERVAL_SECONDS:
        return state
    print("[{}] computing live Anbu/TN-scope summary (full month-wise history - runs every "
          "{} hours, not every poll)...".format(ts, ANBU_TN_SCOPE_POLL_INTERVAL_SECONDS // 3600))
    try:
        rows = fetch_anbu_tn_scope()
    except Exception as e:
        print("[{}] Anbu/TN-scope computation failed this round (will retry at the next "
              "poll): {}".format(ts, e))
        return state
    post_backend("liveAnbuTnScopeUpdate", "rows", rows)
    print("[{}] Anbu/TN-scope: pushed {} row(s) across 3 buckets.".format(ts, len(rows)))
    state["last_anbu_tn_scope_push_ts"] = time.time()
    return state


# --push-anbu-tn-scope-now - one-shot, same reasoning as
# --push-petronas-summary-now above: get the figures onto the dashboard
# right away without waiting for the watcher's normal 12h cadence or
# restarting the whole continuous loop. REAL push (post_backend() +
# writes to the LiveAnbuTnScope sheet) - not a diagnostic.
def push_anbu_tn_scope_now():
    state = load_state()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("[{}] computing Anbu/TN-scope full month-wise history...".format(ts))
    rows = fetch_anbu_tn_scope()
    post_backend("liveAnbuTnScopeUpdate", "rows", rows)
    state["last_anbu_tn_scope_push_ts"] = time.time()
    save_state(state)
    print("[{}] pushed {} row(s) across 3 buckets.".format(ts, len(rows)))
    print("\nDone. Refresh the dashboard (or wait a few seconds for its own auto-refresh) "
          "to see the updated figures.")


# ── Petronas Business Summary - LIVE, current month only (1 Oct 2026) ──
#
# Owner's explicit choice, 1 Oct 2026, after sending the real Jan-Aug 2026
# figures (screenshot, baked into the dashboard as a static one-time
# snapshot - the "one-time report" decision from 30 Sep 2026 stands for
# those closed months): "FROM SEP THERE IS LIVE DATA, U HAVE TO PUT THAT,
# FROM SEP ONWARD LIVE DATA SHOULD COME" - i.e. the CURRENT calendar month
# onward should no longer wait for a manual screenshot/bake-in cycle; it
# should compute and push itself, same as every other live panel.
#
# fetch_petronas_summary() below is a live, non-diagnostic, CURRENT-MONTH-
# ONLY version of report_petronas_summary() above - it reuses the exact
# same ledger-resolution (_resolve_ledgers_under_group against the
# confirmed PETRONAS_PURCHASE_GROUP_NAME Group), the exact same Collection
# rule (Receipt vouchers against a dealer resolving to one of the known
# collection groups), and the exact same inventory-scan Purchase/Sales Ltr
# logic (Purchase voucher type only, Sales/Sales-PETRONAS/Sales-CBE,
# native-Ltr-first-then-catalog-guess) already proven correct against real
# Tally data in report_petronas_summary() - the only difference is the
# date window (1st of THIS month through today, not an arbitrary range)
# and that it returns one dict instead of printing a table. Nothing new
# to diagnose here: the underlying computation already went through this
# file's own diagnostic-first rounds for report_petronas_summary() back in
# 30 Sep 2026 and was confirmed against the owner's real Tally screens.
def _month_end_date_str(month_key6):
    # "202609" -> "20260930". Used only by fetch_petronas_summary()'s
    # explicit-month path (the September backfill below) - the normal
    # current-month path never needs a month's last day since "today" is
    # always inside the current month.
    y = int(month_key6[:4])
    m = int(month_key6[4:6])
    if m == 12:
        ny, nm = y + 1, 1
    else:
        ny, nm = y, m + 1
    last_day = date(ny, nm, 1) - timedelta(days=1)
    return last_day.strftime("%Y%m%d")


def fetch_petronas_summary(month_key=None):
    # month_key=None (the normal poll path): current month, month-to-date -
    # 1st of this month through today. month_key="YYYYMM" (used once, for
    # the September 2026 backfill below): that specific month in full, or
    # up to today if it's still the current month - lets the same function
    # serve both "live, still running" and "closed, compute once" months.
    if month_key:
        from_d = month_key + "01"
        today_d = datetime.now().strftime("%Y%m%d")
        to_d = min(_month_end_date_str(month_key), today_d)
        mk = month_key
    else:
        now = datetime.now()
        from_d = now.strftime("%Y%m01")
        to_d = now.strftime("%Y%m%d")
        mk = _month_key(to_d)

    purchase_ac_ledgers = _resolve_ledgers_under_group(PETRONAS_PURCHASE_GROUP_NAME)
    purchase_ac_lower = set(n.lower() for n in purchase_ac_ledgers)
    coll_ledger_parent, _coll_group_parent, coll_resolve = _build_collection_group_resolver()

    purchase_value = 0.0
    payment_made = 0.0
    collected = 0.0
    # Petronas scheme/claim Credit Notes received, added 15 Oct 2026 - see
    # PETRONAS_CLAIM_PARTY_NAME_'s own comment above for the full story.
    # Tracked SEPARATELY from payment_made on purpose (the owner's own
    # choice: "Add as a separate line", not folded into Payment Made to
    # Petronas) - shown as its own "Claims (CN)" column on the dashboard,
    # Sep 2026 onward only.
    claims_cn = 0.0

    # Added 5 Oct 2026 - same fix and same reasoning as report_petronas_
    # summary()'s own sales_vtypes_for_purchase_ac (see that function's
    # comment): this is the LIVE twin of that one-time report, actually
    # driving Sep 2026 onward on the real dashboard every 12h, so it carried
    # the identical bug - a Sales-type voucher touching a ledger under the
    # PETRONAS Purchase group was being counted as a real Purchase/Payment,
    # purely because of which ledger it happened to touch, not what kind of
    # voucher it was. Owner's report: "in Payment Made to Petronas i need to
    # show here with omit of voucher made in Sales - PETRONAS" - scoped to
    # the whole sales-voucher-type set (DAYBOOK_VOUCHER_TYPE_NAMES), not just
    # "Sales - PETRONAS" by name, since "Sales"/"Sales - CBE" carry the exact
    # same risk and this function already has that set defined just below as
    # sales_vtypes for the inventory-scan loop - reused here, not redefined.
    sales_vtypes_for_purchase_ac = set(DAYBOOK_VOUCHER_TYPE_NAMES)

    if purchase_ac_ledgers:
        vouchers_xml = tally_request(
            ALL_VOUCHERS_FOR_GROUP_SUMMARY_XML.format(from_date=from_d, to_date=to_d), timeout=600)
        vroot = ET.fromstring(vouchers_xml)
        for v in vroot.iter("VOUCHER"):
            vdate = (_text(v, "DATE") or "").strip()
            if not vdate or vdate < from_d or vdate > to_d:
                continue
            vtype = (_text(v, "VOUCHERTYPENAME") or "").strip().lower()
            is_sales_voucher = vtype in sales_vtypes_for_purchase_ac
            # Added 15 Oct 2026 for claims_cn below - ALL_VOUCHERS_FOR_GROUP_
            # SUMMARY_XML already fetches PARTYLEDGERNAME (see its own FETCH
            # list), just never read in this function before now.
            v_party = (_text(v, "PARTYLEDGERNAME") or "").strip()
            is_claim_cn_voucher = vtype in PETRONAS_CLAIM_CN_VOUCHER_TYPE_NAMES and _is_petronas_claim_party_(v_party)
            for entry in v.iter("ALLLEDGERENTRIES.LIST"):
                lname = (_text(entry, "LEDGERNAME") or "").strip()
                if not lname:
                    continue
                amt = _num(_text(entry, "AMOUNT")) or 0.0
                if lname.lower() in purchase_ac_lower and not is_sales_voucher:
                    if amt >= 0:
                        purchase_value += amt
                    else:
                        payment_made += (-amt)
                # Widened 6 Oct 2026 from "receipt" only to all 5 owner-named
                # voucher types - see PETRONAS_SUMMARY_COLLECTION_VOUCHER_TYPE_NAMES
                # above for the full reasoning and the owner's exact request.
                if vtype in PETRONAS_SUMMARY_COLLECTION_VOUCHER_TYPE_NAMES:
                    lparent = coll_ledger_parent.get(lname)
                    if lparent and coll_resolve(lparent) and amt >= 0:
                        collected += amt
                # Petronas claim Credit Note, added 15 Oct 2026 - the specific
                # ledger line (within an already-confirmed Credit Note-family
                # voucher billed to the claim party) that carries the CN's own
                # Rs amount. UNCONFIRMED against a real claim CN yet - owner
                # should check the first live figure this produces against an
                # actual Credit Note he knows, same "diagnostic-first, then
                # trust it" pattern every other live figure in this file went
                # through before being treated as final.
                if is_claim_cn_voucher and _is_petronas_claim_party_(lname):
                    claims_cn += abs(amt)

    item_group = _build_item_group_resolver()
    purchase_ltr = 0.0
    sales_ltr = 0.0
    sales_value = 0.0
    sales_vtypes = set(DAYBOOK_VOUCHER_TYPE_NAMES)

    inv_xml = tally_request(PETRONAS_INVENTORY_SCAN_TEMPLATE.format(from_date=from_d, to_date=to_d), timeout=600)
    iroot = ET.fromstring(inv_xml)
    for v in iroot.iter("VOUCHER"):
        vdate = (_text(v, "DATE") or "").strip()
        if not vdate or vdate < from_d or vdate > to_d:
            continue
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip().lower()
        is_purchase = (vtype in PETRONAS_PURCHASE_VOUCHER_TYPE_NAMES)
        is_sales = (vtype in sales_vtypes)
        if not is_purchase and not is_sales:
            continue
        # Added 15 Oct 2026 - see PETRONAS_CLAIM_PARTY_NAME_'s own comment
        # above. A Sales-type voucher billed to Petronas itself (a claim
        # document, not a real dealer sale) still moves real stock - so Ltr
        # stays counted below exactly as before - but its Rs value is not
        # real dealer revenue, so it's excluded from sales_value only.
        is_claim_party_voucher = is_sales and _is_petronas_claim_party_(_text(v, "PARTYLEDGERNAME"))
        for entry in v.iter("ALLINVENTORYENTRIES.LIST"):
            item_name = (_text(entry, "STOCKITEMNAME") or "").strip()
            if not item_name:
                continue
            item_grp = item_group.get(item_name, "")
            if not _is_petronas_item_group(item_grp):
                continue
            raw_actual = _text(entry, "ACTUALQTY")
            raw_billed = _text(entry, "BILLEDQTY")
            qty_num = _num(raw_billed) if raw_billed is not None else _num(raw_actual)
            ltr = _parse_native_ltr(raw_billed) or _parse_native_ltr(raw_actual)
            if ltr is None and qty_num is not None:
                ltr = _catalog_ltr_guess(item_name, qty_num)
            if ltr is None:
                continue
            if is_purchase:
                purchase_ltr += ltr
            else:
                sales_ltr += ltr
                if is_claim_party_voucher:
                    continue
                amt = _num(_text(entry, "AMOUNT"))
                if amt is not None:
                    # Added 5 Oct 2026 - same fix as report_petronas_summary()'s
                    # own month_sales_value (see that function's comment for the
                    # full reasoning): this AMOUNT is the pre-GST/basic value,
                    # confirmed by the same already-shipped "* 1.18" conversion
                    # proven for the Day Book export's "Value incl GST" column.
                    # This is the LIVE path (Sep 2026 onward on the real
                    # dashboard), so this is the fix that actually matters going
                    # forward - report_petronas_summary() only ever covers the
                    # Jan-Aug static backfill.
                    sales_value += round(abs(amt) * 1.18, 2)

    return {
        "month": mk,                    # "YYYYMM", e.g. "202610" - same _month_key() convention as elsewhere
        "monthLabel": _month_label(mk),  # "Oct 2026" - precomputed so the dashboard doesn't need to reparse mk
        "purchaseLtr": round(purchase_ltr, 1),
        "salesLtr": round(sales_ltr, 1),
        "salesValue": round(sales_value, 2),
        "collected": round(collected, 2),
        "purchaseValue": round(purchase_value, 2),
        "payment": round(payment_made, 2),
        "claimsCn": round(claims_cn, 2),
    }

# ============================================================================
# S.REP - Sales & Collection Plan one-page summary panel, added 4 Nov 2026.
# Owner's request (after sharing an Excel screenshot of his own "OCT '26
# Month" S.REP tracking sheet): a new dashboard panel, directly above the
# existing "DSR / Dealer / Beat - Sales & Collection Plan" section, showing
# Sales(Ltr) and Collection(Rs) - each with Plan/Actual/Pending - for 10
# rows, one TOTAL row, a month dropdown, NO per-row expand (owner explicitly
# retracted his own original "+ to expand" framing: "no need for expand just
# total need").
#
# Design fully nailed down via two AskUserQuestion rounds + a mockup the
# owner approved ("ALL OK") with two corrections ("REMOVE SIVAGANGAI NO NEED
# HERE" and "BY DSR SALES APP PLAN U FIRST FILL HERE..." - see Code.gs's own
# SrepMonthlyPlan sheet comment for how that second correction is handled;
# this file only computes the live Sales/Collection actuals, never the
# owner-entered Plan numbers).
#
# Row -> data-source mapping, exactly as the owner confirmed:
#   AAA, Anandh, Arun, Chellamani, Madhu, Nagaraj, Prabha - the same 7
#     KNOWN_DSR_GROUP_NAMES groups already used everywhere else in this
#     file (Saravanan's separate group folds into Chellamani, same
#     convention as every other per-DSR structure here).
#   Pioneer - the existing "PETRONAS RAMNAD AREA (PIONEER)" group, already
#     used by the live Collection Group Summary feature
#     (_build_collection_group_resolver() / KNOWN_COLLECTION_GROUP_NAMES).
#   VARI - single dealer, exact Tally ledger name "SREE VARI AUTOMOBILES,UPT
#     (OIL)".
#   Aarav - single dealer, exact Tally ledger name "AARAV ENTERPRISES,CBE
#     (CBE-OIL)" (a Coimbatore dealer; owner: "i will share his target month
#     on month" - Aarav's Plan is owner-supplied, never computed here).
#
# Deliberately built as its OWN self-contained voucher scan (Sales(Ltr) from
# fetch_daybook() - the same arbitrary-date-range Sales source already used
# for Day Book/Petronas Business Summary - and its own ALL_VOUCHERS_FOR_
# GROUP_SUMMARY_XML-based Collection(Rs) scan), rather than retrofitting
# month-parameterization into fetch_collection() (which is hardcoded to
# current-month-only and already heavily tuned for its own DSR/dealer/daily/
# 6-month structures). This mirrors fetch_petronas_summary()'s own "collected"
# figure, which likewise does its own fresh voucher scan rather than calling
# fetch_collection() - the established pattern in this file for a new
# month-parameterized live figure. The DSR-group and Collection-group
# resolvers are each built fresh here too (_build_group_resolver(),
# _build_collection_group_resolver()), per this file's own "never share a
# resolver across not-yet-verified features" convention.
#
# UNCONFIRMED against real Tally data yet - this is the diagnostic-first
# step every other live-data feature in this file went through
# (--dump-srep-summary below) before being wired into any sheet push or
# dashboard UI. Nothing is pushed anywhere by this function; it only
# computes and returns.

SREP_ROW_LABELS = ["AAA", "Pioneer", "VARI", "Anandh", "Arun", "Ajith", "Madhu", "Nagaraj", "Prabha", "Aarav"]
# Renamed from "Chellamani" to "Ajith", 14 Nov 2026 - she quit and he took
# over her area (plus part of Saravanan's, see the dsr_group_to_row
# override just below fetch_srep_summary()). Must exactly match Code.gs's
# own SREP_ROW_LABELS_, which was renamed the same day - a mismatch here
# is why this row would otherwise show 0 live Sales/Collection on the
# S.REP panel even though the dashboard's main DSR card renders correctly.

# Exact Tally ledger names, as given by the owner - matched case-
# insensitively against Day Book's "party" / a voucher's PARTYLEDGERNAME.
SREP_DEALER_ROWS = {
    "VARI": "SREE VARI AUTOMOBILES,UPT (OIL)",
    "Aarav": "AARAV ENTERPRISES,CBE (CBE-OIL)",
}

# DSR row label -> its KNOWN_DSR_GROUP_NAMES entry.
SREP_DSR_ROW_TO_GROUP = {
    "AAA": "PETRONAS AAA DIRECT AREA SALES",
    "Anandh": "PETRONAS TVL1 AREA (ANANDH)",
    "Arun": "PETRONAS DINDIGUL AREA (ARUN)",
    "Ajith": "PETRONAS MDU AREA (AJITH)",  # Tally group itself renamed by the owner, 14 Nov 2026 (was "(CHELLAMANI)")
    "Madhu": "PETRONAS NKL AREA (MADHU)",
    "Nagaraj": "PETRONAS VIRUDHUNAGAR AREA (NAGARAJ)",
    "Prabha": "PETRONAS THENI AREA (PRABHA)",
}


def fetch_srep_summary(month_key=None):
    """month_key=None (normal poll path): current month, month-to-date -
    same convention as fetch_petronas_summary(). month_key="YYYYMM": that
    specific month in full, or up to today if it's the current month - the
    same shared signature, for the same reason: lets one function serve
    both the live poll and the month dropdown's historical-month requests.

    Returns {"month": mk, "monthLabel": ..., "rows": [{"label", "salesLtr",
    "collectionRs"}, ...]} - 10 rows, SREP_ROW_LABELS order. No Plan/
    Pending here - those are owner-entered/derived, computed in Code.gs
    alongside this live data, not in this file.
    """
    if month_key:
        from_d = month_key + "01"
        today_d = datetime.now().strftime("%Y%m%d")
        to_d = min(_month_end_date_str(month_key), today_d)
        mk = month_key
    else:
        now = datetime.now()
        from_d = now.strftime("%Y%m01")
        to_d = now.strftime("%Y%m%d")
        mk = _month_key(to_d)

    sales_ltr = dict((r, 0.0) for r in SREP_ROW_LABELS)
    collection_rs = dict((r, 0.0) for r in SREP_ROW_LABELS)

    dsr_group_to_row = dict((g, r) for r, g in SREP_DSR_ROW_TO_GROUP.items())
    # Saravanan's separate Tally group folds into Chellamani's row - same
    # convention as every other per-DSR structure in this file (see the
    # _merge_saravanan_into_chellamani_* family above for the original
    # reasoning: his dealers moved into Chellamani's order-slip app on 1
    # Sep 2026, but that merge never touched Tally, so his dealers still
    # sit under his own separate Group there).
    dsr_group_to_row["PETRONAS MDU AREA (SARAVANAN)"] = "Ajith"  # renamed 14 Nov 2026

    dealer_row_lower = dict((v.lower(), k) for k, v in SREP_DEALER_ROWS.items())

    ledger_parent_dsr, resolve_dsr = _build_group_resolver()
    coll_ledger_parent, _coll_group_parent, resolve_coll = _build_collection_group_resolver()

    # ---- Sales (Ltr), from Day Book, same [from_d, to_d] window ----
    for r in fetch_daybook(from_d, to_d):
        party = r.get("party") or ""
        ltr = r.get("altQtyLtr") or 0.0
        if not ltr:
            continue
        party_lower = party.lower()
        row_label = dealer_row_lower.get(party_lower)
        if not row_label:
            parent = ledger_parent_dsr.get(party, "")
            dsr_group = resolve_dsr(parent)
            if dsr_group:
                row_label = dsr_group_to_row.get(dsr_group)
            if not row_label:
                coll_group = resolve_coll(parent)
                if coll_group == "PETRONAS RAMNAD AREA (PIONEER)":
                    row_label = "Pioneer"
        if row_label:
            sales_ltr[row_label] += ltr

    # ---- Collection (Rs), own whole-company voucher scan, same window -
    # same ALL_VOUCHERS_FOR_GROUP_SUMMARY_XML request and positive-amount-
    # on-the-dealer's-own-ledger-entry convention already proven for
    # fetch_collection()/fetch_petronas_summary()'s own "collected" figure.
    vouchers_xml = tally_request(
        ALL_VOUCHERS_FOR_GROUP_SUMMARY_XML.format(from_date=from_d, to_date=to_d), timeout=600)
    vroot = ET.fromstring(vouchers_xml)
    for v in vroot.iter("VOUCHER"):
        vdate = (_text(v, "DATE") or "").strip()
        if not vdate or vdate < from_d or vdate > to_d:
            continue
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip().lower()
        if vtype not in COLLECTION_VOUCHER_TYPE_NAMES:
            continue
        for entry in v.iter("ALLLEDGERENTRIES.LIST"):
            lname = (_text(entry, "LEDGERNAME") or "").strip()
            if not lname:
                continue
            amt = _num(_text(entry, "AMOUNT")) or 0.0
            if amt <= 0:
                continue
            lname_lower = lname.lower()
            row_label = dealer_row_lower.get(lname_lower)
            if not row_label:
                parent = ledger_parent_dsr.get(lname, "")
                dsr_group = resolve_dsr(parent)
                if dsr_group:
                    row_label = dsr_group_to_row.get(dsr_group)
                if not row_label:
                    coll_parent = coll_ledger_parent.get(lname, "")
                    coll_group = resolve_coll(coll_parent)
                    if coll_group == "PETRONAS RAMNAD AREA (PIONEER)":
                        row_label = "Pioneer"
            if row_label:
                collection_rs[row_label] += amt

    rows = []
    for label in SREP_ROW_LABELS:
        rows.append({
            "label": label,
            "salesLtr": round(sales_ltr[label], 1),
            "collectionRs": round(collection_rs[label], 2),
        })

    return {
        "month": mk,
        "monthLabel": _month_label(mk),
        "rows": rows,
    }


def dump_srep_summary(month_key=None):
    """--dump-srep-summary [YYYYMM] - diagnostic only, pushes nothing.
    Prints the 10 rows' live Sales(Ltr)/Collection(Rs) so the owner can
    check them against his own Excel sheet / real Tally screens before
    this is wired into any sheet push or dashboard panel."""
    data = fetch_srep_summary(month_key)
    print("S.REP summary - {} ({})".format(data["monthLabel"], data["month"]))
    print("{:<12} {:>12} {:>16}".format("Row", "Sales(Ltr)", "Collection(Rs)"))
    total_ltr = 0.0
    total_rs = 0.0
    for row in data["rows"]:
        print("{:<12} {:>12.1f} {:>16.2f}".format(row["label"], row["salesLtr"], row["collectionRs"]))
        total_ltr += row["salesLtr"]
        total_rs += row["collectionRs"]
    print("{:<12} {:>12.1f} {:>16.2f}".format("TOTAL", total_ltr, total_rs))


# Added 4 Nov 2026 - the owner flagged that his own Oct '26 Excel sheet's
# per-DSR Collection figures (which he says were ALSO Tally, Receipt+Cheque,
# resolved to DSR group - the same method as fetch_srep_summary() above)
# don't match this live recomputation: AAA/Pioneer/Anandh/Chellamani all
# ~40-46% HIGHER here, Arun/Madhu/Nagaraj/Prabha 4-27% LOWER - a mixed
# direction, not a simple double-count. VARI (a plain exact-ledger-name
# match, no group-chain walk at all) came back correct and confirmed by the
# owner, which points the suspicion specifically at _build_group_resolver()'s
# ledger->group->parent walk: if even one or two dealer ledgers' GROUP
# ASSIGNMENT in Tally has changed since the owner's sheet was made (a
# dealer moved between DSR area groups), re-resolving October's vouchers
# against TODAY's hierarchy would misattribute those dealers' October
# receipts to a different row than the owner's sheet originally counted
# them under - exactly the "some rows up, some down" pattern seen. This
# diagnostic prints the actual DEALER-level contributors behind each DSR
# row's Collection total, so the owner can spot a receipt sitting under a
# row it doesn't belong to, by name, against his own records - same
# "dealer-level breakdown so the owner can eyeball it" pattern already used
# elsewhere in this file's diagnostics (e.g. dump_group_summary()).
def dump_srep_collection_detail(month_key=None, top_n=12):
    """--dump-srep-collection-detail [YYYYMM] [top_n] - diagnostic only,
    pushes nothing. For each of the 10 S.REP rows, prints its Collection
    total PLUS the individual dealer ledgers that made it up (highest
    amount first, capped at top_n per row - pass a bigger top_n to see
    more). Re-fetches everything fresh (own Tally requests, not shared
    with fetch_srep_summary()) so this stays a simple, standalone
    diagnostic that can't accidentally change that function's behaviour."""
    if month_key:
        from_d = month_key + "01"
        today_d = datetime.now().strftime("%Y%m%d")
        to_d = min(_month_end_date_str(month_key), today_d)
        mk = month_key
    else:
        now = datetime.now()
        from_d = now.strftime("%Y%m01")
        to_d = now.strftime("%Y%m%d")
        mk = _month_key(to_d)

    dsr_group_to_row = dict((g, r) for r, g in SREP_DSR_ROW_TO_GROUP.items())
    dsr_group_to_row["PETRONAS MDU AREA (SARAVANAN)"] = "Ajith"  # renamed 14 Nov 2026
    dealer_row_lower = dict((v.lower(), k) for k, v in SREP_DEALER_ROWS.items())

    ledger_parent_dsr, resolve_dsr = _build_group_resolver()
    coll_ledger_parent, _coll_group_parent, resolve_coll = _build_collection_group_resolver()

    # row -> {dealer_ledger_name: amount}
    by_row = dict((r, {}) for r in SREP_ROW_LABELS)

    vouchers_xml = tally_request(
        ALL_VOUCHERS_FOR_GROUP_SUMMARY_XML.format(from_date=from_d, to_date=to_d), timeout=600)
    vroot = ET.fromstring(vouchers_xml)
    for v in vroot.iter("VOUCHER"):
        vdate = (_text(v, "DATE") or "").strip()
        if not vdate or vdate < from_d or vdate > to_d:
            continue
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip().lower()
        if vtype not in COLLECTION_VOUCHER_TYPE_NAMES:
            continue
        for entry in v.iter("ALLLEDGERENTRIES.LIST"):
            lname = (_text(entry, "LEDGERNAME") or "").strip()
            if not lname:
                continue
            amt = _num(_text(entry, "AMOUNT")) or 0.0
            if amt <= 0:
                continue
            lname_lower = lname.lower()
            row_label = dealer_row_lower.get(lname_lower)
            if not row_label:
                parent = ledger_parent_dsr.get(lname, "")
                dsr_group = resolve_dsr(parent)
                if dsr_group:
                    row_label = dsr_group_to_row.get(dsr_group)
                if not row_label:
                    coll_parent = coll_ledger_parent.get(lname, "")
                    coll_group = resolve_coll(coll_parent)
                    if coll_group == "PETRONAS RAMNAD AREA (PIONEER)":
                        row_label = "Pioneer"
            if row_label:
                by_row[row_label][lname] = by_row[row_label].get(lname, 0.0) + amt

    print("S.REP Collection detail - {} ({})".format(_month_label(mk), mk))
    for row_label in SREP_ROW_LABELS:
        dealers = by_row[row_label]
        row_total = sum(dealers.values())
        print("")
        print("{} - total Rs {:.2f} - {} dealer(s)".format(row_label, row_total, len(dealers)))
        for dealer_name, amt in sorted(dealers.items(), key=lambda kv: -kv[1])[:top_n]:
            print("    {:<45} {:>14.2f}".format(dealer_name[:45], amt))
        if len(dealers) > top_n:
            print("    ... and {} more dealer(s) not shown".format(len(dealers) - top_n))


# Same heavy-whole-company-pull reasoning as Collection/Sales/Day Book
# above (this scans every voucher for the month, twice - once for the
# ledger-based Purchase Value/Payment/Collection figures, once for the
# inventory-based Purchase/Sales Ltr figures, and it does this for TWO
# months every cycle now - see below) - own cadence, not every ~10-min
# poll. Changed from hourly to twice a day (12h) 1 Oct 2026 at the
# owner's own request ("not to check for every hour make it to do for 2
# times a day") - this data doesn't move fast enough to need hourly
# polling, and halves the heavy whole-company voucher pulls this feature
# costs Tally per day.
# Changed from 12h to 4h (6x/day) 7 Oct 2026 at the owner's own request
# (first asked for 6h, then corrected in the same breath to "not 6 make
# it 4") - owner wants fresher Business Summary figures through the day;
# triples the heavy whole-company pulls this feature costs Tally per day
# versus the 1 Oct 12h cadence - still well short of the original hourly
# cadence this was changed from on 1 Oct.
PETRONAS_SUMMARY_POLL_INTERVAL_SECONDS = 4 * 60 * 60


def _previous_month_key():
    # "current month is 202610" -> "202609". Rolls forward automatically
    # as real time passes - no hardcoded month, unlike the one-time
    # September backfill this replaced (see the comment on
    # maybe_push_petronas_summary() below for why a fixed one-time
    # backfill wasn't the right fix).
    now = datetime.now()
    y, m = now.year, now.month
    if m == 1:
        py, pm = y - 1, 12
    else:
        py, pm = y, m - 1
    return "{:04d}{:02d}".format(py, pm)


def maybe_push_petronas_summary(state, dry_run, ts):
    if dry_run:
        return state
    last = state.get("last_petronas_summary_push_ts", 0)
    if time.time() - last < PETRONAS_SUMMARY_POLL_INTERVAL_SECONDS:
        return state
    print("[{}] computing live Petronas Business Summary (current month so far, PLUS "
          "last month - runs every {} hours, not every poll)...".format(
              ts, PETRONAS_SUMMARY_POLL_INTERVAL_SECONDS // 3600))
    # Recomputes BOTH the current month (month-to-date) AND the previous
    # calendar month every cycle - added 1 Oct 2026 at the owner's own
    # explanation: "month end purchase will be recvd at our in firstweek
    # then only we do purchase so always last month data also to be
    # updated normally" - Purchase invoices for a month routinely arrive
    # and get entered into Tally during the FIRST WEEK of the following
    # month, dated back to the month they're actually for. A month that
    # simply froze the moment it stopped being "current" (the original
    # 30 Sep/1 Oct design, and still how Collection/Sales/Group Summary
    # behave) would miss every one of those late entries. Recomputing
    # "last month" on every cycle, indefinitely, means a late Tally entry
    # for last month is picked up automatically within one cycle (up to
    # 4 hours, since the 7 Oct 2026 cadence change - see
    # PETRONAS_SUMMARY_POLL_INTERVAL_SECONDS above) rather than needing a
    # manual re-sync request. This also
    # supersedes the old one-time September-2026 backfill (removed) -
    # September was "last month" on 1 Oct 2026, so it's covered by this
    # same mechanism without a special case; a month older than "last
    # month" still freezes at whatever it last computed, same as before,
    # since purchase invoices arriving 2+ months late is not the
    # scenario the owner described.
    for month_key in (None, _previous_month_key()):
        try:
            summary = fetch_petronas_summary(month_key)
        except Exception as e:
            print("[{}] Petronas summary computation failed this round for {} (will "
                  "retry at the next poll): {}".format(
                      ts, month_key or "current month", e))
            continue
        post_backend("livePetronasSummaryUpdate", "petronasSummary", summary)
        print("[{}] Petronas summary: pushed {} (Purchase {} Ltr, Sales {} Ltr, "
              "Sales Rs {}, Collected Rs {}, Purchase Rs {}, Payment Rs {})".format(
                  ts, summary["month"], summary["purchaseLtr"], summary["salesLtr"],
                  summary["salesValue"], summary["collected"], summary["purchaseValue"],
                  summary["payment"]))
    state["last_petronas_summary_push_ts"] = time.time()
    return state


# --push-petronas-summary-now - added 1 Oct 2026, same day as the TATA/
# MARUTI group exclusion fix above, at the owner's own request ("CAN U MAKE
# WATCHER TO START THE UPDATE THE DATA NOW FRESH") - a one-shot command to
# get the corrected figures onto the dashboard right away, without waiting
# for the watcher's normal cadence (maybe_push_petronas_summary() above,
# 4-hourly since 7 Oct 2026 - see PETRONAS_SUMMARY_POLL_INTERVAL_SECONDS)
# or restarting the whole continuous watcher loop. REAL push - this
# calls post_backend() and writes to the LivePetronasSummary sheet, exactly
# like a normal cycle; it is NOT a diagnostic. Computes and pushes the exact
# same two months maybe_push_petronas_summary() would (current month-to-
# date, PLUS last month), then resets last_petronas_summary_push_ts so the
# next normal cycle counts from this push, not from whenever the
# watcher last happened to run one.
def push_petronas_summary_now():
    state = load_state()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    pushed_any = False
    for month_key in (None, _previous_month_key()):
        print("[{}] computing {} ...".format(ts, month_key or "current month, month-to-date"))
        summary = fetch_petronas_summary(month_key)
        post_backend("livePetronasSummaryUpdate", "petronasSummary", summary)
        pushed_any = True
        print("[{}] pushed {} ({}) - Purchase {} Ltr, Sales {} Ltr, Sales Rs {}, "
              "Collected Rs {}, Purchase Rs {}, Payment Rs {}".format(
                  ts, summary["month"], summary["monthLabel"], summary["purchaseLtr"],
                  summary["salesLtr"], summary["salesValue"], summary["collected"],
                  summary["purchaseValue"], summary["payment"]))
    if pushed_any:
        state["last_petronas_summary_push_ts"] = time.time()
        save_state(state)
    print("\nDone. Refresh the dashboard (or wait a few seconds for its own auto-refresh) "
          "to see the updated figures.")


# S.REP - Sales & Collection Plan panel push, added 4 Nov 2026 - see
# fetch_srep_summary() above for the full row-source story, and Code.gs's
# handleLiveSrepSummaryUpdate_()/SREP_MONTHLY_PLAN_SHEET for the receiving
# side. Pushes BOTH the current month (month-to-date) AND the previous
# calendar month every cycle, exactly like maybe_push_petronas_summary()
# above and for the identical reason: a Receipt/Cheque voucher for a month
# can get entered into Tally after that month has closed, so refreshing
# "last month" indefinitely (not just once at month-end) picks up a late
# entry automatically within one cycle instead of needing a manual
# --push-srep-summary-now. Own cadence (SREP_SUMMARY_POLL_INTERVAL_SECONDS
# below), not every ~10-min poll - this does TWO whole-company voucher
# scans (Day Book + the Collection scan) for EACH of the 2 months, the
# same order of heaviness as the Petronas Business Summary push, so it
# gets the same few-times-a-day cadence rather than anything more
# frequent, until the owner says otherwise.
SREP_SUMMARY_POLL_INTERVAL_SECONDS = 4 * 60 * 60


def maybe_push_srep_summary(state, dry_run, ts):
    if dry_run:
        return state
    last = state.get("last_srep_summary_push_ts", 0)
    if time.time() - last < SREP_SUMMARY_POLL_INTERVAL_SECONDS:
        return state
    print("[{}] computing S.REP Sales & Collection summary (current month so far, PLUS "
          "last month - runs every {} hours, not every poll)...".format(
              ts, SREP_SUMMARY_POLL_INTERVAL_SECONDS // 3600))
    for month_key in (None, _previous_month_key()):
        try:
            summary = fetch_srep_summary(month_key)
        except Exception as e:
            print("[{}] S.REP summary computation failed this round for {} (will "
                  "retry at the next poll): {}".format(
                      ts, month_key or "current month", e))
            continue
        post_backend("liveSrepSummaryUpdate", "srepSummary", summary)
        total_ltr = sum(r["salesLtr"] for r in summary["rows"])
        total_rs = sum(r["collectionRs"] for r in summary["rows"])
        print("[{}] S.REP summary: pushed {} ({} rows, total {} Ltr / Rs {})".format(
            ts, summary["month"], len(summary["rows"]), total_ltr, total_rs))
    state["last_srep_summary_push_ts"] = time.time()
    return state


# --push-srep-summary-now - a one-shot command to get the S.REP panel's
# figures onto the dashboard right away, without waiting for the normal
# cadence or restarting the whole watcher loop - same pattern and same
# reason as --push-petronas-summary-now above. REAL push (post_backend() +
# writes to LiveSrepSummary), not a diagnostic. Pushes the same two months
# maybe_push_srep_summary() would, then resets last_srep_summary_push_ts.
def push_srep_summary_now():
    state = load_state()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    pushed_any = False
    for month_key in (None, _previous_month_key()):
        print("[{}] computing {} ...".format(ts, month_key or "current month, month-to-date"))
        summary = fetch_srep_summary(month_key)
        post_backend("liveSrepSummaryUpdate", "srepSummary", summary)
        pushed_any = True
        total_ltr = sum(r["salesLtr"] for r in summary["rows"])
        total_rs = sum(r["collectionRs"] for r in summary["rows"])
        print("[{}] pushed {} ({}) - {} rows, total {} Ltr / Rs {}".format(
            ts, summary["month"], summary["monthLabel"], len(summary["rows"]), total_ltr, total_rs))
    if pushed_any:
        state["last_srep_summary_push_ts"] = time.time()
        save_state(state)
    print("\nDone. Refresh the dashboard (or wait a few seconds for its own auto-refresh) "
          "to see the updated figures.")


# ── Historical per-DSR summary (Jan 2026 onward) - diagnostic only ─────
#
# Owner's request, 1 Oct 2026: "pickup data from tally from jan 2026 , not
# full req only sales collection in ltr and value , purchase ltr and value
# monthly closing outstanding for all dealer and purchase from petronas and
# what else main data u need." Clarified via 3 questions before building
# anything (per this file's own standing "ask first" discipline): (1) scope
# - Petronas only, NOT also TATA OE (owner's own choice - AAA also
# distributes TATA OE lubricants per the business profile, but this pull
# stays scoped exactly like the already-confirmed Petronas Business Summary
# above); (2) per-DSR breakdown - owner chose YES, not company-wide totals
# only, so this can later feed a historical Team Business Summary/per-DSR
# trend, not just the company-wide table already built; (3) outstanding
# detail - owner chose an AGE-BUCKET summary (30-59/60-89/90+, matching the
# live dashboard's own existing bands) per DSR per month, NOT a full
# per-dealer monthly closing balance (hundreds of dealers x 9 months).
#
# Sales/Purchase Ltr+Value and Collection/Payment here are NOT a new build -
# they're report_petronas_summary()'s own already-confirmed logic, just with
# each voucher's dealer/party resolved to a DSR via the same
# _build_group_resolver()/ledger_to_dsr pattern fetch_sales()/
# fetch_collection() already use, so the real company-wide totals this
# produces should match report_petronas_summary()'s own Apr-Sep 2026 figures
# exactly when summed back across all DSRs - that cross-check is printed
# automatically below.
#
# The monthly closing-outstanding age-bucket piece is genuinely NEW and
# UNVERIFIED: fetch_outstanding() only ever asks Tally for a live, as-of-
# TODAY "Bills Receivable" snapshot (BILLS_RECEIVABLE_REPORT_XML, no date
# variables at all) - nothing in this project has ever asked Tally for that
# same canned report as of a PAST date. This diagnostic adds an SVTODATE to
# that request and hopes Tally's own "Bills Receivable" report honours it as
# an "as on" date, same as the office's own manual Alt+P month-end habit
# would expect - but this is a DIFFERENT request mechanism (EXPORTDATA on a
# named report) than the raw TYPE="Voucher" Collection requests whose
# SVFROMDATE/SVTODATE were already found to be silently ignored elsewhere in
# this file (Collection, Sales, the Day Book dump), so it might work fine
# here or might not. Rather than guess either way, this prints the 30-59/
# 60-89/90+ totals for 3 different past month-ends and flags it loudly if
# they come out suspiciously identical - the same symptom that first
# exposed the SVFROMDATE/SVTODATE quirk on the voucher-list requests.
#
# --company lets this ONE diagnostic target a different Tally company than
# whatever's currently open - added specifically for the Jan-Mar 2026 gap
# found 30 Sep 2026 while building the Petronas Business Summary (every
# metric came back zero for those 3 months, most likely because TallyPrime
# keeps each financial year, Apr-Mar in India, as its own separate company
# data file, and only FY 2026-27 was open that day). No other function in
# this file can target a named company - this is deliberately isolated to
# this new diagnostic so it can never change behaviour for an already-
# trusted live path.

HISTORICAL_LEDGER_ENTRIES_TEMPLATE = """<ENVELOPE>
 <HEADER>
  <VERSION>1</VERSION>
  <TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE>
  <ID>HistoricalLedgerEntriesCollection</ID>
 </HEADER>
 <BODY>
  <DESC>
   <STATICVARIABLES>
    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
    <SVFROMDATE>{from_date}</SVFROMDATE>
    <SVTODATE>{to_date}</SVTODATE>{company_tag}
   </STATICVARIABLES>
   <TDL>
    <TDLMESSAGE>
     <COLLECTION NAME="HistoricalLedgerEntriesCollection" ISMODIFY="No">
      <TYPE>Voucher</TYPE>
      <FETCH>DATE, VOUCHERTYPENAME, PARTYLEDGERNAME</FETCH>
      <FETCH>ALLLEDGERENTRIES.LIST</FETCH>
     </COLLECTION>
    </TDLMESSAGE>
   </TDL>
  </DESC>
 </BODY>
</ENVELOPE>"""

HISTORICAL_INVENTORY_SCAN_TEMPLATE = """<ENVELOPE>
 <HEADER>
  <VERSION>1</VERSION>
  <TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE>
  <ID>HistoricalInventoryScanCollection</ID>
 </HEADER>
 <BODY>
  <DESC>
   <STATICVARIABLES>
    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
    <SVFROMDATE>{from_date}</SVFROMDATE>
    <SVTODATE>{to_date}</SVTODATE>{company_tag}
   </STATICVARIABLES>
   <TDL>
    <TDLMESSAGE>
     <COLLECTION NAME="HistoricalInventoryScanCollection" ISMODIFY="No">
      <TYPE>Voucher</TYPE>
      <FETCH>VOUCHERNUMBER, DATE, VOUCHERTYPENAME, PARTYLEDGERNAME</FETCH>
      <FETCH>ALLINVENTORYENTRIES.LIST</FETCH>
     </COLLECTION>
    </TDLMESSAGE>
   </TDL>
  </DESC>
 </BODY>
</ENVELOPE>"""

HISTORICAL_BILLS_RECEIVABLE_TEMPLATE = """<ENVELOPE>
 <HEADER>
  <TALLYREQUEST>Export Data</TALLYREQUEST>
 </HEADER>
 <BODY>
  <EXPORTDATA>
   <REQUESTDESC>
    <REPORTNAME>Bills Receivable</REPORTNAME>
    <STATICVARIABLES>
     <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
     <SVTODATE>{as_of_date}</SVTODATE>{company_tag}
    </STATICVARIABLES>
   </REQUESTDESC>
  </EXPORTDATA>
 </BODY>
</ENVELOPE>"""


def _month_end_dates_in_range(from_d, to_d):
    """Returns [(month_key6, as_of_date8, is_partial), ...] - the last
    calendar day of every month from from_d's month through to_d's month,
    except the final entry if to_d itself falls before that month's own
    last day, in which case to_d is used instead and is_partial=True (the
    current, in-progress month - e.g. today, not yet month-end)."""
    start = datetime.strptime(from_d, "%Y%m%d").replace(day=1)
    end_full = datetime.strptime(to_d, "%Y%m%d")
    out = []
    cur = start
    while cur <= end_full:
        if cur.month == 12:
            next_month = cur.replace(year=cur.year + 1, month=1, day=1)
        else:
            next_month = cur.replace(month=cur.month + 1, day=1)
        month_last_day = next_month - timedelta(days=1)
        if month_last_day > end_full:
            out.append((cur.strftime("%Y%m"), end_full.strftime("%Y%m%d"), True))
            break
        out.append((cur.strftime("%Y%m"), month_last_day.strftime("%Y%m%d"), False))
        cur = next_month
    return out


def dump_historical_summary(from_date=None, to_date=None, company=None):
    """
    --dump-historical-summary [YYYYMMDD-from] [YYYYMMDD-to] [--company "<exact Tally company name>"]

    DIAGNOSTIC ONLY - nothing pushed anywhere, nothing saved to the Sheet.
    See the big comment block right above this function for the full
    background (owner's 1 Oct 2026 request, the 3 clarifying answers, and
    why the outstanding piece is genuinely new/unverified). Defaults to
    1 Jan 2026 through today.
    """
    now = datetime.now()
    from_d = from_date or "20260101"
    to_d = to_date or now.strftime("%Y%m%d")
    company_tag = ""
    if company:
        company_tag = "\n    <SVCURRENTCOMPANY>{}</SVCURRENTCOMPANY>".format(_sanitize_xml_text(company))

    print("=" * 70)
    print("HISTORICAL PER-DSR SUMMARY - DIAGNOSTIC ONLY - {} to {}".format(from_d, to_d))
    if company:
        print("Targeting company: {!r} (via SVCURRENTCOMPANY - UNCONFIRMED whether Tally".format(company))
        print("honours this for a company other than the one currently open)")
    print("Scope: Petronas only (not TATA OE), per-DSR breakdown, outstanding as age-buckets.")
    print("Nothing in this run is pushed anywhere or saved to the Sheet.")
    print("=" * 70)

    print("\nBuilding DSR ledger resolver (same chain-walk as fetch_sales()/fetch_collection())...")
    ledger_parent, resolve = _build_group_resolver()
    ledger_to_dsr = {}
    for lname, lparent in ledger_parent.items():
        if not lparent:
            continue
        dsr = resolve(lparent)
        if dsr:
            ledger_to_dsr[lname] = dsr
    print("  {} dealer ledger(s) resolved to one of the {} known DSR groups.".format(
        len(ledger_to_dsr), len(KNOWN_DSR_GROUP_NAMES)))

    print("\nResolving every ledger under '{}' (Purchase A/C, company-wide only - this".format(
        PETRONAS_PURCHASE_GROUP_NAME))
    print("liability sits with AAA as a whole, not any one DSR, same as report_petronas_summary())...")
    purchase_ac_ledgers = _resolve_ledgers_under_group(PETRONAS_PURCHASE_GROUP_NAME)
    purchase_ac_lower = set(n.lower() for n in purchase_ac_ledgers)
    print("  {} ledger(s) resolved.".format(len(purchase_ac_ledgers)))

    # ---- STEP A: Collection (per DSR per month) + Purchase Value/Payment (company-wide) ----
    print("\n" + "=" * 70)
    print("STEP A: Collection per DSR per month; Purchase(Rs)/Payment(Rs) company-wide")
    print("        (fetching ALL vouchers company-wide, unfiltered - please be patient)")
    print("=" * 70)
    ledger_xml_body = HISTORICAL_LEDGER_ENTRIES_TEMPLATE.format(
        from_date=from_d, to_date=to_d, company_tag=company_tag)
    vouchers_xml = tally_request(ledger_xml_body, timeout=600)
    print("  Response received: {} bytes.".format(len(vouchers_xml)))
    vroot = ET.fromstring(vouchers_xml)

    # dsr_month_collected[dsr][month_key] = Rs
    dsr_month_collected = {}
    month_purchase_value = {}
    month_payment_made = {}
    unresolved_collection_rs = {}  # month_key -> Rs from a Receipt entry whose dealer didn't resolve to any known DSR

    for v in vroot.iter("VOUCHER"):
        vdate = (_text(v, "DATE") or "").strip()
        if not vdate or vdate < from_d or vdate > to_d:
            continue
        mk = _month_key(vdate)
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip().lower()
        for entry in v.iter("ALLLEDGERENTRIES.LIST"):
            lname = (_text(entry, "LEDGERNAME") or "").strip()
            if not lname:
                continue
            amt = _num(_text(entry, "AMOUNT")) or 0.0
            if lname.lower() in purchase_ac_lower:
                if amt >= 0:
                    month_purchase_value[mk] = month_purchase_value.get(mk, 0.0) + amt
                else:
                    month_payment_made[mk] = month_payment_made.get(mk, 0.0) + (-amt)
            if vtype == "receipt" and amt >= 0:
                dsr = ledger_to_dsr.get(lname)
                if dsr:
                    dsr_month_collected.setdefault(dsr, {})
                    dsr_month_collected[dsr][mk] = dsr_month_collected[dsr].get(mk, 0.0) + amt
                else:
                    unresolved_collection_rs[mk] = unresolved_collection_rs.get(mk, 0.0) + amt

    _merge_saravanan_into_chellamani(dsr_month_collected)

    # ---- STEP B: Sales + Purchase Ltr/Value, per DSR per month (inventory scan) ----
    print("\n" + "=" * 70)
    print("STEP B: Sales/Purchase Ltr+Value per DSR per month (inventory scan, Petronas")
    print("        item groups only - please be patient)")
    print("=" * 70)
    item_group = _build_item_group_resolver()
    inv_xml_body = HISTORICAL_INVENTORY_SCAN_TEMPLATE.format(
        from_date=from_d, to_date=to_d, company_tag=company_tag)
    inv_xml = tally_request(inv_xml_body, timeout=600)
    print("  Response received: {} bytes.".format(len(inv_xml)))
    iroot = ET.fromstring(inv_xml)

    # dsr_month_sales_ltr[dsr][month_key] = Ltr, etc.
    dsr_month_sales_ltr = {}
    dsr_month_sales_value = {}
    dsr_month_purchase_ltr = {}
    month_purchase_ltr_total = {}  # company-wide cross-check against report_petronas_summary()
    month_sales_ltr_total = {}
    unresolved_sales_ltr = {}

    sales_vtypes = set(DAYBOOK_VOUCHER_TYPE_NAMES)  # ("sales", "sales - petronas", "sales - cbe", "sales - cash", "pos billing") - 10 Oct 2026

    for v in iroot.iter("VOUCHER"):
        vdate = (_text(v, "DATE") or "").strip()
        if not vdate or vdate < from_d or vdate > to_d:
            continue
        mk = _month_key(vdate)
        vtype = (_text(v, "VOUCHERTYPENAME") or "").strip().lower()
        is_purchase = (vtype in PETRONAS_PURCHASE_VOUCHER_TYPE_NAMES)
        is_sales = (vtype in sales_vtypes)
        if not is_purchase and not is_sales:
            continue
        party = (_text(v, "PARTYLEDGERNAME") or "").strip()
        dsr = ledger_to_dsr.get(party)
        for entry in v.iter("ALLINVENTORYENTRIES.LIST"):
            item_name = (_text(entry, "STOCKITEMNAME") or "").strip()
            if not item_name:
                continue
            item_grp = item_group.get(item_name, "")
            if not _is_petronas_item_group(item_grp):
                continue
            raw_actual = _text(entry, "ACTUALQTY")
            raw_billed = _text(entry, "BILLEDQTY")
            qty_num = _num(raw_billed) if raw_billed is not None else _num(raw_actual)
            ltr = _parse_native_ltr(raw_billed) or _parse_native_ltr(raw_actual)
            if ltr is None and qty_num is not None:
                ltr = _catalog_ltr_guess(item_name, qty_num)
            if ltr is None:
                continue
            if is_purchase:
                month_purchase_ltr_total[mk] = month_purchase_ltr_total.get(mk, 0.0) + ltr
                if dsr:
                    dsr_month_purchase_ltr.setdefault(dsr, {})
                    dsr_month_purchase_ltr[dsr][mk] = dsr_month_purchase_ltr[dsr].get(mk, 0.0) + ltr
            else:
                month_sales_ltr_total[mk] = month_sales_ltr_total.get(mk, 0.0) + ltr
                amt = _num(_text(entry, "AMOUNT"))
                if dsr:
                    dsr_month_sales_ltr.setdefault(dsr, {})
                    dsr_month_sales_ltr[dsr][mk] = dsr_month_sales_ltr[dsr].get(mk, 0.0) + ltr
                    if amt is not None:
                        dsr_month_sales_value.setdefault(dsr, {})
                        dsr_month_sales_value[dsr][mk] = dsr_month_sales_value[dsr].get(mk, 0.0) + abs(amt)
                else:
                    unresolved_sales_ltr[mk] = unresolved_sales_ltr.get(mk, 0.0) + ltr

    _merge_saravanan_into_chellamani(dsr_month_sales_ltr)
    _merge_saravanan_into_chellamani(dsr_month_sales_value)
    _merge_saravanan_into_chellamani(dsr_month_purchase_ltr)

    all_months = sorted(set(
        list(month_purchase_ltr_total.keys()) + list(month_sales_ltr_total.keys()) +
        list(month_purchase_value.keys()) + list(month_payment_made.keys())))
    dsrs_for_print = [d for d in KNOWN_DSR_GROUP_NAMES if d != "PETRONAS MDU AREA (SARAVANAN)"]

    print("\n" + "=" * 70)
    print("RESULT A - per-DSR Sales(Ltr)/Sales(Rs)/Collection(Rs)/Purchase(Ltr), by month")
    print("=" * 70)
    for mk in all_months:
        print("\n  {} —".format(_month_label(mk)))
        print("  {:<38} {:>10} {:>14} {:>14} {:>12}".format(
            "DSR", "Sales(Ltr)", "Sales(Rs)", "Collect(Rs)", "Purch(Ltr)"))
        for dsr in dsrs_for_print:
            sl = dsr_month_sales_ltr.get(dsr, {}).get(mk, 0.0)
            sv = dsr_month_sales_value.get(dsr, {}).get(mk, 0.0)
            co = dsr_month_collected.get(dsr, {}).get(mk, 0.0)
            pl = dsr_month_purchase_ltr.get(dsr, {}).get(mk, 0.0)
            if sl == 0.0 and sv == 0.0 and co == 0.0 and pl == 0.0:
                continue
            print("  {:<38} {:>10.1f} {:>14.2f} {:>14.2f} {:>12.1f}".format(
                dsr, sl, sv, co, pl))
        if unresolved_sales_ltr.get(mk):
            print("  (unresolved dealer, no known DSR) Sales(Ltr): {:.1f}".format(
                unresolved_sales_ltr.get(mk, 0.0)))
        if unresolved_collection_rs.get(mk):
            print("  (unresolved dealer, no known DSR) Collect(Rs): {:.2f}".format(
                unresolved_collection_rs.get(mk, 0.0)))

    print("\n" + "=" * 70)
    print("CROSS-CHECK against report_petronas_summary()'s own company-wide totals -")
    print("summing every DSR above (+ unresolved) per month should match these exactly:")
    print("=" * 70)
    print("  {:<10} {:>12} {:>12} {:>14} {:>14} {:>14} {:>14}".format(
        "Month", "Purch(Ltr)", "Sales(Ltr)", "Sales(Rs)", "Purch(Rs)", "Payment(Rs)", "(sum check)"))
    for mk in all_months:
        pl_t = month_purchase_ltr_total.get(mk, 0.0)
        sl_t = month_sales_ltr_total.get(mk, 0.0)
        pv = month_purchase_value.get(mk, 0.0)
        pay = month_payment_made.get(mk, 0.0)
        sum_sl = sum(dsr_month_sales_ltr.get(d, {}).get(mk, 0.0) for d in dsrs_for_print) + unresolved_sales_ltr.get(mk, 0.0)
        ok = "OK" if abs(sum_sl - sl_t) < 0.5 else "MISMATCH {:.1f} vs {:.1f}".format(sum_sl, sl_t)
        print("  {:<10} {:>12.1f} {:>12.1f} {:>14.2f} {:>14.2f} {:>14.2f} {:>14}".format(
            _month_label(mk), pl_t, sl_t, 0.0, pv, pay, ok))
    print("  (Sales(Rs) column left blank here - see RESULT A above for the real per-DSR figures;")
    print("  this row only exists to carry the per-DSR-sum-vs-total Sales(Ltr) check in the last column.)")
    print("\n  If these look right against report_petronas_summary()'s own already-reviewed")
    print("  Apr-Sep 2026 output, the per-DSR breakdown above can be trusted the same way.")

    # ---- STEP C: monthly closing outstanding, by age-bucket, per DSR (EXPERIMENTAL) ----
    print("\n" + "=" * 70)
    print("STEP C: monthly closing outstanding, by age-bucket, per DSR - EXPERIMENTAL,")
    print("        see this function's own big comment above for why this is unverified.")
    print("        One Tally request PER month-end below - please be patient.")
    print("=" * 70)
    month_ends = _month_end_dates_in_range(from_d, to_d)
    print("  Will check these {} as-of date(s): {}".format(
        len(month_ends), ", ".join("{}{}".format(d, " (in progress)" if p else "") for _, d, p in month_ends)))

    # dsr_month_outstanding[dsr][month_key] = {"pendingAmount":.., "band3059":.., "band6089":.., "band90plus":..}
    dsr_month_outstanding = {}
    unresolved_outstanding = {}
    grand_total_by_asof = {}  # as_of_date -> grand total pendingAmount, for the sameness check below

    for mk, as_of, is_partial in month_ends:
        bills_xml_body = HISTORICAL_BILLS_RECEIVABLE_TEMPLATE.format(as_of_date=as_of, company_tag=company_tag)
        bills_xml = tally_request(bills_xml_body)
        broot = ET.fromstring(bills_xml)

        bills_by_party = {}
        current_party = None
        current_date_iso = None
        for el in broot.iter():
            if el.tag == "BILLFIXED":
                party = _text(el, "BILLPARTY")
                current_party = party.strip() if party else None
                current_date_iso = _billdate_ddmonyy_to_iso(_text(el, "BILLDATE"))
            elif el.tag == "BILLCL":
                amt = _num(el.text)
                if current_party and current_date_iso and amt is not None:
                    bills_by_party.setdefault(current_party, []).append(
                        {"date": current_date_iso, "amount": amt})

        grand_total = 0.0
        for party, bills in bills_by_party.items():
            dsr = ledger_to_dsr.get(party)
            target = dsr_month_outstanding.setdefault(dsr, {}) if dsr else unresolved_outstanding
            rec = target.setdefault(mk, {"pendingAmount": 0.0, "band3059": 0.0, "band6089": 0.0, "band90plus": 0.0})
            for b in bills:
                if b["amount"] >= 0:
                    continue  # not a genuine debit
                rec["pendingAmount"] += b["amount"]
                grand_total += b["amount"]
                try:
                    age_days = (datetime.strptime(as_of, "%Y%m%d") - datetime.strptime(b["date"], "%Y-%m-%d")).days
                except ValueError:
                    continue
                if age_days >= 90:
                    rec["band90plus"] += b["amount"]
                elif age_days >= 60:
                    rec["band6089"] += b["amount"]
                elif age_days >= 30:
                    rec["band3059"] += b["amount"]
        grand_total_by_asof[as_of] = grand_total

    _merge_saravanan_into_chellamani_nested_dict(dsr_month_outstanding)

    print("\n" + "=" * 70)
    print("RESULT C - closing outstanding as of each month-end, by DSR (Rs, owed = positive)")
    print("=" * 70)
    for mk, as_of, is_partial in month_ends:
        print("\n  As of {}{} —".format(as_of, " (in progress, not a real month-end)" if is_partial else ""))
        print("  {:<38} {:>14} {:>12} {:>12} {:>12}".format(
            "DSR", "Total owed", "30-59d", "60-89d", "90+d"))
        for dsr in dsrs_for_print:
            rec = dsr_month_outstanding.get(dsr, {}).get(mk)
            if not rec or rec["pendingAmount"] == 0.0:
                continue
            print("  {:<38} {:>14.2f} {:>12.2f} {:>12.2f} {:>12.2f}".format(
                dsr, (-rec["pendingAmount"]) or 0.0, (-rec["band3059"]) or 0.0,
                (-rec["band6089"]) or 0.0, (-rec["band90plus"]) or 0.0))
        u = unresolved_outstanding.get(mk)
        if u and u["pendingAmount"] != 0.0:
            print("  (unresolved dealer, no known DSR) Total owed: {:.2f}".format(-u["pendingAmount"]))

    print("\n" + "=" * 70)
    print("SANITY CHECK: do these as-of totals actually differ by date? (Grand total owed,")
    print("company-wide, per as-of date)")
    print("=" * 70)
    asof_list = sorted(grand_total_by_asof.keys())
    for as_of in asof_list:
        print("  {}: {:.2f}".format(as_of, (-grand_total_by_asof[as_of]) or 0.0))
    distinct_values = set(round(v, 2) for v in grand_total_by_asof.values())
    if len(asof_list) >= 2 and len(distinct_values) == 1:
        print("\n  *** WARNING: every as-of date above produced the EXACT SAME grand total. ***")
        print("  This is the same symptom that first exposed the SVFROMDATE/SVTODATE-doesn't-")
        print("  scope-a-raw-Voucher-Collection quirk elsewhere in this file - it strongly")
        print("  suggests Tally is NOT honouring SVTODATE on this canned 'Bills Receivable'")
        print("  report either, and is just returning today's live snapshot every time")
        print("  regardless of which as-of date was requested. If so, STEP C's numbers above")
        print("  are NOT real historical closing-outstanding figures - they're all just")
        print("  today's live snapshot repeated. Do not trust STEP C until this is resolved.")
    else:
        print("\n  Totals differ across as-of dates - at least a GOOD SIGN that SVTODATE is")
        print("  doing something here. Still worth confirming one of these numbers by hand")
        print("  against a real Tally 'Bills Receivable as on <date>' screen for a month you")
        print("  know well before trusting the rest.")

    print("\n" + "=" * 70)
    print("Send this whole output back. Specifically useful to check: does STEP A's per-DSR")
    print("Sales/Collection for a month you know well add up to what you expect; does the")
    print("CROSS-CHECK table say OK for every month; and most importantly, did the SANITY")
    print("CHECK above come back a warning or a good sign for STEP C. If Jan-Mar 2026 still")
    print("show nothing, send the exact company name for FY 2025-26 as Tally shows it, and")
    print("re-run with --company \"<that name>\".")


def _group_summary_key(group, month_key):
    """Composite key for a group-level Group Summary row - Group +
    calendar month, e.g. "PETRONAS MADURAI TOTAL AREA||2026-09". Distinct
    per month on purpose (never overwritten month to month), since the
    owner explicitly wants every month's figures kept in the SAME sheet/
    tab going forward ("data should be saved in the same tab for all
    months") rather than Day Book's separate-tab-per-month pattern - a
    plain dict keyed this way, never cleared, naturally accumulates that
    history across every calendar month this watcher runs through."""
    return "{}||{}".format(group, month_key)


def _group_summary_dealer_key(group, month_key, dealer):
    """Same as _group_summary_key() above, one level more specific - one
    row per dealer within a group within a month."""
    return "{}||{}||{}".format(group, month_key, dealer)


def fetch_group_summary():
    """
    Real (non-diagnostic) Collection Group Summary fetch - current
    calendar month-to-date, all 7 KNOWN_COLLECTION_GROUP_NAMES. This is a
    straight refactor of dump_group_summary()'s own STEP 4/4b logic
    (unfiltered whole-company voucher pull, per-dealer Debit/Credit,
    post-dated-voucher correction via _corrected_closing()) into a clean,
    reusable, non-printing function once that method was fully validated
    against the owner's real Tally data across 7 diagnostic rounds (see
    dump_group_summary()'s docstring/history above) - EXACT matches for
    AAA/Madurai, and Tirunelveli within an explained live-data-drift
    margin, on the final (28 Sep 2026) validation run.

    Returns (group_summary, dealer_summary):
      group_summary = {key: {"group", "monthKey", "opening", "debit",
                              "credit", "closing"}}, keyed by
                       _group_summary_key(group, monthKey) - one row per
                       group for the current month.
      dealer_summary = {key: {"group", "monthKey", "dealer", "opening",
                               "debit", "credit", "closing"}}, keyed by
                        _group_summary_dealer_key(group, monthKey,
                        dealer) - one row per dealer within every group
                        for the current month (even a dealer with zero
                        movement this month gets a row, from
                        dealers_by_group's full membership list, so the
                        dashboard's "+" expand always shows every dealer
                        that belongs to the group, not just ones with a
                        voucher this month).

    Both dicts are keyed so the SAME group/dealer in a later calendar
    month gets a brand-new key (monthKey changes) rather than overwriting
    this month's row - see _group_summary_key()'s own comment above for
    why that matters here specifically.
    """
    now = datetime.now()
    from_d = now.strftime("%Y%m01")
    to_d = now.strftime("%Y%m%d")
    month_key = now.strftime("%Y-%m")

    ledger_parent, _group_parent, resolve = _build_collection_group_resolver()
    dealers_by_group = dict((g, []) for g in KNOWN_COLLECTION_GROUP_NAMES)
    for lname, lparent in ledger_parent.items():
        if not lparent:
            continue
        g = resolve(lparent)
        if g:
            dealers_by_group[g].append(lname)

    ledger_xml = tally_request(LEDGER_GROUP_COLLECTION_XML)
    ledger_root = ET.fromstring(ledger_xml)
    ledger_closing = {}
    for led in ledger_root.iter("LEDGER"):
        lname = led.get("NAME") or _text(led, "NAME")
        if lname:
            lname = lname.strip()
            # _num() returns None (not 0.0) for a blank/missing
            # CLOSINGBALANCE - coalesced here, same fix as
            # dump_group_summary()'s own ledger_closing loop (confirmed
            # bug from a real 28 Sep run: "float + NoneType" TypeError).
            ledger_closing[lname] = _num(_text(led, "CLOSINGBALANCE")) or 0.0

    xml_body = ALL_VOUCHERS_FOR_GROUP_SUMMARY_XML.format(from_date=from_d, to_date=to_d)
    vouchers_xml = tally_request(xml_body, timeout=600)
    root = ET.fromstring(vouchers_xml)

    group_debit = dict((g, 0.0) for g in KNOWN_COLLECTION_GROUP_NAMES)
    group_credit = dict((g, 0.0) for g in KNOWN_COLLECTION_GROUP_NAMES)
    dealer_debit = {}
    dealer_credit = {}
    # Post-dated (dated after to_d) movement, tracked separately and
    # backed back out of each dealer's raw CLOSINGBALANCE below via
    # _corrected_closing() - see dump_group_summary()'s STEP 4 comment
    # (the owner's own diagnosis) for the full reasoning.
    dealer_future_debit = {}
    dealer_future_credit = {}

    for v in root.iter("VOUCHER"):
        vdate = _text(v, "DATE")
        if not vdate:
            continue
        is_future = vdate > to_d
        if not is_future and vdate < from_d:
            continue
        for entry in v.iter("ALLLEDGERENTRIES.LIST"):
            lname = (_text(entry, "LEDGERNAME") or "").strip()
            if not lname:
                continue
            lparent = ledger_parent.get(lname)
            if not lparent:
                continue
            g = resolve(lparent)
            if not g:
                continue
            amt = _num(_text(entry, "AMOUNT"))
            if amt is None:
                continue
            if is_future:
                if amt >= 0:
                    dealer_future_credit[lname] = dealer_future_credit.get(lname, 0.0) + amt
                else:
                    dealer_future_debit[lname] = dealer_future_debit.get(lname, 0.0) + (-amt)
                continue
            if amt >= 0:
                group_credit[g] += amt
                dealer_credit[lname] = dealer_credit.get(lname, 0.0) + amt
            else:
                group_debit[g] += -amt
                dealer_debit[lname] = dealer_debit.get(lname, 0.0) + (-amt)

    def _corrected_closing(d):
        raw = ledger_closing.get(d, 0.0)
        fut_net = dealer_future_credit.get(d, 0.0) - dealer_future_debit.get(d, 0.0)
        return raw - fut_net

    group_summary = {}
    dealer_summary = {}
    for gname in KNOWN_COLLECTION_GROUP_NAMES:
        debit = round(group_debit[gname], 2)
        credit = round(group_credit[gname], 2)
        closing = round(sum(_corrected_closing(d) for d in dealers_by_group[gname]), 2)
        opening = round(closing - (credit - debit), 2)
        group_summary[_group_summary_key(gname, month_key)] = {
            "group": gname,
            "monthKey": month_key,
            "opening": opening,
            "debit": debit,
            "credit": credit,
            "closing": closing,
        }
        for d in sorted(dealers_by_group[gname]):
            d_closing = round(_corrected_closing(d), 2)
            d_debit = round(dealer_debit.get(d, 0.0), 2)
            d_credit = round(dealer_credit.get(d, 0.0), 2)
            d_opening = round(d_closing - (d_credit - d_debit), 2)
            dealer_summary[_group_summary_dealer_key(gname, month_key, d)] = {
                "group": gname,
                "monthKey": month_key,
                "dealer": d,
                "opening": d_opening,
                "debit": d_debit,
                "credit": d_credit,
                "closing": d_closing,
            }
    return group_summary, dealer_summary


def diff_group_summary(prev, current):
    """prev/current are both {key: row} dicts from fetch_group_summary()'s
    first return value - a changed/new key's row is returned as-is (it
    already carries its own group/monthKey fields, unlike diff_collection()
    which has to re-attach the dict key as a field)."""
    changed = []
    for key, rec in current.items():
        if prev.get(key) != rec:
            changed.append(rec)
    return changed


def diff_group_summary_by_dealer(prev, current):
    """Same shape/reasoning as diff_group_summary() above, one level more
    specific (per-dealer rows instead of per-group)."""
    changed = []
    for key, rec in current.items():
        if prev.get(key) != rec:
            changed.append(rec)
    return changed


# Collection Group Summary runs on the SLOWEST cadence of anything in this
# file - confirmed with the owner via AskUserQuestion ("A few times a day"
# / 1-2 hours, not every ~10-min poll like the other panels), because
# fetch_group_summary() above pulls EVERY voucher type company-wide for
# the month (no Receipt-only FILTER like Collection, no Sales-only
# VOUCHERTYPENAME filter like Sales/Day Book) - confirmed the heaviest
# single Tally response in this whole file during the 7-round diagnostic
# validation (dump_group_summary()'s STEP 4, tens of MB per fetch on a
# mid-month run). 2 hours picked as the top of the owner's confirmed
# "1-2 hours" range, matching that this is the heaviest pull of all of
# them - Tally stays the most responsive with this one polled the least.
GROUP_SUMMARY_POLL_INTERVAL_SECONDS = 2 * 60 * 60


def maybe_push_group_summary(state, dry_run, ts):
    if dry_run:
        return state
    last = state.get("last_group_summary_push_ts", 0)
    if time.time() - last < GROUP_SUMMARY_POLL_INTERVAL_SECONDS:
        return state
    print("[{}] computing Collection Group Summary (all voucher types, this month so far - "
          "runs every {} min, not every poll - the heaviest single pull in this file)...".format(
              ts, GROUP_SUMMARY_POLL_INTERVAL_SECONDS // 60))
    try:
        group_summary, dealer_summary = fetch_group_summary()
    except Exception as e:
        print("[{}] group summary computation failed this round (will retry at the next "
              "eligible poll): {}".format(ts, e))
        return state

    changed = diff_group_summary(state.get("groupSummary", {}), group_summary)
    if changed:
        sent = post_backend_chunked("liveGroupSummaryUpdate", "groups", changed, ts)
        print("[{}] group summary: sent {} changed group row(s) total".format(ts, sent))
    else:
        print("[{}] group summary: no changes".format(ts))

    dealer_changed = diff_group_summary_by_dealer(state.get("groupSummaryByDealer", {}), dealer_summary)
    if dealer_changed:
        dealer_sent = post_backend_chunked("liveGroupSummaryByDealerUpdate", "dealers", dealer_changed, ts)
        print("[{}] group summary (by dealer): sent {} changed dealer row(s) total".format(ts, dealer_sent))
    else:
        print("[{}] group summary (by dealer): no changes".format(ts))

    # Both dicts kept in full (not just this month's slice) so every past
    # month's rows stay in state and are never re-diffed-as-"new" again -
    # see _group_summary_key()'s own comment for why a new month simply
    # adds new keys rather than replacing old ones.
    state["groupSummary"] = group_summary if not state.get("groupSummary") else dict(state["groupSummary"], **group_summary)
    state["groupSummaryByDealer"] = dealer_summary if not state.get("groupSummaryByDealer") else dict(state["groupSummaryByDealer"], **dealer_summary)
    state["last_group_summary_push_ts"] = time.time()
    return state


def main():
    dry_run = "--test" in sys.argv

    if "--check" in sys.argv:
        idx = sys.argv.index("--check")
        terms = [a for a in sys.argv[idx + 1:] if not a.startswith("--")]
        if not terms:
            print("Usage: python3 tally_live_watcher.py --check \"dealer name or fragment\" [more...]")
            sys.exit(1)
        try:
            check_dealers(terms)
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--check-collection" in sys.argv:
        idx = sys.argv.index("--check-collection")
        rest = [a for a in sys.argv[idx + 1:] if not a.startswith("--")]
        if not rest:
            print("Usage: python3 tally_live_watcher.py --check-collection \"arun\" "
                  "[YYYYMMDD-from] [YYYYMMDD-to]")
            sys.exit(1)
        dsr_fragment = rest[0]
        from_date = rest[1] if len(rest) > 1 else None
        to_date = rest[2] if len(rest) > 2 else None
        try:
            check_collection(dsr_fragment, from_date, to_date)
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-sixtyplus" in sys.argv:
        idx = sys.argv.index("--dump-sixtyplus")
        rest = [a for a in sys.argv[idx + 1:] if not a.startswith("--")]
        dealer_fragment = rest[0] if rest else None
        try:
            dump_sixtyplus(dealer_fragment)
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-sales" in sys.argv:
        idx = sys.argv.index("--dump-sales")
        rest = [a for a in sys.argv[idx + 1:] if not a.startswith("--")]
        def _looks_like_date(s):
            return len(s) == 8 and s.isdigit()
        dsr_fragment = None
        dates = rest
        if rest and not _looks_like_date(rest[0]):
            dsr_fragment = rest[0]
            dates = rest[1:]
        from_date = dates[0] if len(dates) > 0 else None
        to_date = dates[1] if len(dates) > 1 else None
        try:
            dump_sales(dsr_fragment, from_date, to_date)
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-purchase-history" in sys.argv:
        idx = sys.argv.index("--dump-purchase-history")
        rest = [a for a in sys.argv[idx + 1:] if not a.startswith("--")]
        if not rest:
            print("Usage: --dump-purchase-history <dealer fragment> [days-back, default 365, 0=everything]")
            sys.exit(1)
        dealer_fragment = rest[0]
        days_back = rest[1] if len(rest) > 1 else None
        try:
            dump_purchase_history(dealer_fragment, days_back)
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-invoice-stats" in sys.argv:
        idx = sys.argv.index("--dump-invoice-stats")
        rest = [a for a in sys.argv[idx + 1:] if not a.startswith("--")]
        def _looks_like_date(s):
            return len(s) == 8 and s.isdigit()
        dsr_fragment = None
        dates = rest
        if rest and not _looks_like_date(rest[0]):
            dsr_fragment = rest[0]
            dates = rest[1:]
        from_date = dates[0] if len(dates) > 0 else None
        to_date = dates[1] if len(dates) > 1 else None
        try:
            dump_invoice_stats(dsr_fragment, from_date, to_date)
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-group-summary" in sys.argv:
        idx = sys.argv.index("--dump-group-summary")
        rest = [a for a in sys.argv[idx + 1:] if not a.startswith("--")]
        from_date = rest[0] if len(rest) > 0 else None
        to_date = rest[1] if len(rest) > 1 else None
        try:
            dump_group_summary(from_date, to_date)
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-petronas-purchase" in sys.argv:
        idx = sys.argv.index("--dump-petronas-purchase")
        rest = [a for a in sys.argv[idx + 1:] if not a.startswith("--")]
        from_date = rest[0] if len(rest) > 0 else None
        to_date = rest[1] if len(rest) > 1 else None
        try:
            dump_petronas_purchase(from_date, to_date)
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-petronas-item-groups" in sys.argv:
        idx = sys.argv.index("--dump-petronas-item-groups")
        rest = [a for a in sys.argv[idx + 1:] if not a.startswith("--")]
        from_date = rest[0] if len(rest) > 0 else None
        to_date = rest[1] if len(rest) > 1 else None
        try:
            dump_petronas_item_group_breakdown(from_date, to_date)
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-anbu-tn-scope" in sys.argv:
        idx = sys.argv.index("--dump-anbu-tn-scope")
        rest = [a for a in sys.argv[idx + 1:] if not a.startswith("--")]
        from_date = rest[0] if len(rest) > 0 else None
        to_date = rest[1] if len(rest) > 1 else None
        try:
            dump_anbu_tn_scope(from_date, to_date)
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--push-petronas-summary-now" in sys.argv:
        try:
            push_petronas_summary_now()
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-anbu-tn-scope-live-check" in sys.argv:
        try:
            dump_anbu_tn_scope_live_check()
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--push-anbu-tn-scope-now" in sys.argv:
        try:
            push_anbu_tn_scope_now()
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--push-srep-summary-now" in sys.argv:
        try:
            push_srep_summary_now()
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--push-roster-now" in sys.argv:
        try:
            push_dealer_roster_now()
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--push-everything-now" in sys.argv:
        try:
            push_everything_now()
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--petronas-summary" in sys.argv:
        idx = sys.argv.index("--petronas-summary")
        rest = [a for a in sys.argv[idx + 1:] if not a.startswith("--")]
        from_date = rest[0] if len(rest) > 0 else None
        to_date = rest[1] if len(rest) > 1 else None
        try:
            report_petronas_summary(from_date, to_date)
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-historical-summary" in sys.argv:
        idx = sys.argv.index("--dump-historical-summary")
        rest = [a for a in sys.argv[idx + 1:] if not a.startswith("--")]
        from_date = rest[0] if len(rest) > 0 else None
        to_date = rest[1] if len(rest) > 1 else None
        company = None
        if "--company" in sys.argv:
            cidx = sys.argv.index("--company")
            if cidx + 1 < len(sys.argv):
                company = sys.argv[cidx + 1]
            else:
                print("Usage: python3 tally_live_watcher.py --dump-historical-summary "
                      "[YYYYMMDD-from] [YYYYMMDD-to] --company \"<exact Tally company name>\"")
                sys.exit(1)
        try:
            dump_historical_summary(from_date, to_date, company)
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-daybook" in sys.argv:
        idx = sys.argv.index("--dump-daybook")
        rest = [a for a in sys.argv[idx + 1:] if not a.startswith("--")]
        from_date = rest[0] if len(rest) > 0 else None
        to_date = rest[1] if len(rest) > 1 else None
        try:
            dump_daybook(from_date, to_date)
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-srep-summary" in sys.argv:
        idx = sys.argv.index("--dump-srep-summary")
        rest = [a for a in sys.argv[idx + 1:] if not a.startswith("--")]
        month_key = rest[0] if len(rest) > 0 else None
        try:
            dump_srep_summary(month_key)
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-srep-collection-detail" in sys.argv:
        idx = sys.argv.index("--dump-srep-collection-detail")
        rest = [a for a in sys.argv[idx + 1:] if not a.startswith("--")]
        month_key = rest[0] if len(rest) > 0 else None
        top_n = 12
        if len(rest) > 1:
            try:
                top_n = int(rest[1])
            except ValueError:
                pass
        try:
            dump_srep_collection_detail(month_key, top_n)
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-price" in sys.argv:
        try:
            dump_price_check()
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-mrp-live" in sys.argv:
        try:
            dump_mrp_live()
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-report" in sys.argv:
        try:
            dump_report()
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-xml" in sys.argv:
        try:
            dump_xml()
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-mrp" in sys.argv:
        try:
            dump_mrp()
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-stock-unit" in sys.argv:
        try:
            dump_stock_unit()
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-groups" in sys.argv:
        idx = sys.argv.index("--dump-groups")
        custom_names = [a for a in sys.argv[idx + 1:] if not a.startswith("--")]
        try:
            dump_groups(custom_names if custom_names else None)
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-all-groups" in sys.argv:
        try:
            dump_all_groups()
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-eligible" in sys.argv:
        try:
            dump_eligible()
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    if "--dump-creditlimit" in sys.argv:
        idx = sys.argv.index("--dump-creditlimit")
        rest = [a for a in sys.argv[idx + 1:] if not a.startswith("--")]
        dealer_fragment = rest[0] if rest else None
        try:
            dump_creditlimit(dealer_fragment)
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return


    if dry_run:
        state = load_state()
        try:
            run_once(state, dry_run=True)
        except requests.exceptions.ConnectionError:
            print("Could not connect to Tally at {} — is TallyPrime open with a company loaded, "
                  "and is Connectivity enabled on port {}? (Gateway of Tally -> F1 Help -> Settings "
                  "-> Connectivity)".format(TALLY_URL, TALLY_PORT))
            sys.exit(1)
        return

    print("Live Tally watcher starting. Polling every {} seconds. Ctrl+C to stop.".format(POLL_SECONDS))
    state = load_state()
    while True:
        try:
            state = run_once(state, dry_run=False)
            save_state(state)
        except requests.exceptions.ConnectionError:
            print("[{}] Tally not reachable right now (is it open?) - will retry next poll".format(
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        except Exception as e:
            print("[{}] error this poll (will retry next time): {}".format(
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"), e))
        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
