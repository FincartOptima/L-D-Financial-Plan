"""
Rebuilds both dashboards from the latest source workbooks in the project folder:

  - plan-approval-lead-status.html   full detail, client names + emails, LOCAL ONLY (git-ignored)
  - index.html                       de-identified, PUBLISHED (private repo only - see the
                                     "unmapped" note in main() for the one exception)

Both are rendered from the SAME template (templates/dashboard.template.html); only the payload
differs. That is deliberate - when the two pages had separate templates they drifted apart and
grew inconsistent bugs.

The client universe is the de-duplicated union of all four data sheets, keyed on client email:

  - Plan Approval Sheet Q1 / Q2   the approval base, and the only source of revenue and of the
                                  dates the turnaround measures need. Always counted.
  - FY 2026-2027 Q1 / Q2          every planning ticket. Each carries a Ticket Subject, and the
                                  dashboard's master filter decides which subjects get added on
                                  top of the approval base.

Source files are auto-detected by pattern + picked by most-recent modified time, so this
still works after the b2c export and the planning workbook get replaced with new filenames:

  - b2c / lead export:      FIN<digits>_*.xlsx      (sheet 'Data')
  - planning workbook:      Financial Planning Tickets Summary-Dashboard*.xlsx

Run with no arguments: `python tools/build_dashboard.py`
"""
import datetime
import json
import re
import sys
import unicodedata
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
TEMPLATES = Path(__file__).resolve().parent / "templates"

B2C_RE = re.compile(r"^FIN\d+_.*\.xlsx$", re.IGNORECASE)
PLAN_RE = re.compile(r"^Financial Planning Tickets Summary-Dashboard.*\.xlsx$", re.IGNORECASE)
EMP_RE = re.compile(r"^EMPLOYEE_REF.*\.xlsx$", re.IGNORECASE)

# The workbook's four data sheets, read into one de-duplicated client universe keyed on email.
#   kind "pa" - Plan Approval sheets: carry revenue and the dates the turnaround measures need.
#   kind "fy" - FY ticket sheets:     carry Ticket Subject, which drives the master filter.
# Column names differ between sheets (Email Id vs Client Mail ID) and even between quarters of the
# same kind, so every field is looked up by header name per sheet, never by position.
SHEETS = [
    {"name": "Plan Approval Sheet Q1", "kind": "pa", "q": "Q1", "email": "Email Id",     "month": "D",     "date": "Date"},
    {"name": "Plan Approval Sheet Q2", "kind": "pa", "q": "Q2", "email": "Email Id",     "month": "D",     "date": "Date"},
    {"name": "FY 2026-2027 Q1",        "kind": "fy", "q": "Q1", "email": "Client Mail ID", "month": "Month", "date": "Tkt Recd Date"},
    {"name": "FY 2026-2027 Q2",        "kind": "fy", "q": "Q2", "email": "Client Mail ID", "month": "Month", "date": "Tkt Recd Date"},
]

# Dropped wholesale: this month's FY rows are almost entirely unusable (274 of 275 carry no client
# email at all, so they can never join the lead data), and it was excluded by request.
EXCLUDE_MONTHS = {"april"}

# The financial year the workbook covers. Dates outside it are treated as data-entry slips.
FY_START = datetime.date(2026, 4, 1)
FY_END = datetime.date(2027, 3, 31)

# The Client Type column holds two unrelated taxonomies depending on which sheet it came from:
# the FY sheets grade clients Alpha/Beta/Gamma (tier), the Plan Approval sheets record New/Existing
# (tenure). They are kept as separate dimensions rather than merged, because a client can be both
# ("Alpha" and "New") and the two vocabularies are not comparable to each other.
TIER_VALUES = {"alpha": "Alpha", "beta": "Beta", "gamma": "Gamma", "gama": "Gamma"}
TENURE_VALUES = {"new": "New", "existing": "Existing"}
BLANK = "Blank"


def client_tier(raw):
    """FY-sheet Client Type -> canonical tier. Case and the recurring 'Gama' typo are folded in;
    anything else the sheet contains is kept verbatim so a new grade shows up rather than vanishing."""
    t = (raw or "").strip()
    if not t:
        return ""
    return TIER_VALUES.get(t.lower(), t)


def client_tenure(raw):
    """Plan-Approval Client Type -> canonical tenure (New / Existing)."""
    t = (raw or "").strip()
    if not t:
        return ""
    return TENURE_VALUES.get(t.lower(), t)

# Revenue/allocation columns on the plan approval tabs, in the order shown in the product filter.
# "All products" sums these. (Monthly Revised Surplus is a monthly cashflow figure rather than an
# allocation — it is included because it was requested, but summing it with the allocations mixes
# two different units, so read the "All products" total with that in mind.)
PRODUCTS = [
    ("unlisted", "Unlisted Allocation"),
    ("grip", "GRIP Allocation"),
    ("ulip", "ULIP Allocation"),
    ("surplus", "Monthly Revised Surplus"),
    ("equity", "Total Equity allocation"),
]

# Advisor / Lead Source are supposed to be short internal category labels, and are the only two
# per-row fields that flow into the public aggregate cube. They're manually typed in the source
# sheet, so guard against fat-finger mistakes (e.g. a client's email or phone pasted into the
# wrong column) leaking through into the public file. Match => treat the cell as blank.
LOOKS_LIKE_PII_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+|\d{7,}")

# The b2c export and the employee reference sheet occasionally spell the same person differently,
# which breaks the RM -> team join. Variants are listed explicitly rather than fuzzy-matched:
# an edit-distance match would eventually pull two genuinely different people together, and a
# wrong team assignment is worse than an unassigned one because it looks correct.
# Left = spelling as it appears in b2c (currentRmName); right = spelling in the employee sheet.
# Both sides are compared after name_key() normalisation (lowercased, whitespace collapsed).
RM_NAME_ALIASES = {
    "bhawna surana": "bhawana surana",
}

# Same idea for client emails. These are typos in the plan approval sheet that point at a real
# lead record - listed one by one rather than fuzzy-matched, because an edit-distance rule on
# email addresses would eventually bind a plan to the wrong person's lead history.
# Fix these at source in the planning sheet and the entry here becomes redundant.
# Left = as typed in the plan approval sheet; right = the real address in b2c.
EMAIL_ALIASES = {
    "inayakpatilvp111@gmail.com": "vinayakpatilvp111@gmail.com",   # dropped leading "v"
}

# Turnaround bands, in days. Cumulative: each counts plans taking strictly more than that many days.
TAT_BUCKETS = [1, 2, 3, 5, 7, 10, 15, 30]
TAT_FIELDS = [("approve", "tatApprove"), ("convert", "tatConvert"), ("inprocess", "tatInProcess")]

# Each measure is scoped to the clients who actually reached that stage, so the denominator is the
# population that could have produced the outcome. Without this, every non-converted client counts
# as "missing a date" against the conversion measure and drags its percentages down.
TAT_UNIVERSE = {
    "approve": None,              # every approved plan has a raised and an approved date
    "convert": "CONVERTED",
    "inprocess": "IN PROCESS",
}

# A conversion dated before the plan was raised is either a client who converted in an earlier
# cycle and came back for a plan, or a conversion logged around the time of drafting but ahead of
# it — the wrong order procedurally. The two are separated by how far back the conversion sits:
# in this data the long tail runs to years (median 214 days) while the suspect cases cluster
# within days of the plan.
#
# The lead's createdDate cannot make this distinction: a lead is always created before a plan is
# drafted for it, so that test classifies every negative as an old client and never fires.
TAT_OLD_CLIENT_DAYS = 30
TAT_SPLIT_NEGATIVES = {"convert"}
TAT_EARLY_LABEL = {
    "convert": "Client Converted Before Drafting Financial Plan",
}
# In-process legitimately precedes drafting — the RM works the lead, then a plan is written — so
# those negatives are reported as one neutral row rather than being judged.
TAT_NEUTRAL_NEG_LABEL = {
    "inprocess": "Reached in-process before the plan was raised",
    # An approval dated before the plan was raised cannot happen in reality — it is a data entry
    # slip in the sheet, so it is surfaced rather than quietly dropped.
    "approve": "Approval Date precedes Date — check the sheet",
}


def scrub(value):
    return "" if value and LOOKS_LIKE_PII_RE.search(value) else value

MONTHS = {
    "jan": (1, "Jan"), "january": (1, "Jan"), "feb": (2, "Feb"), "february": (2, "Feb"),
    "mar": (3, "Mar"), "march": (3, "Mar"), "apr": (4, "Apr"), "april": (4, "Apr"),
    "may": (5, "May"), "jun": (6, "Jun"), "june": (6, "Jun"),
    "jul": (7, "Jul"), "july": (7, "Jul"), "aug": (8, "Aug"), "august": (8, "Aug"),
    "sep": (9, "Sep"), "sept": (9, "Sep"), "september": (9, "Sep"),
    "oct": (10, "Oct"), "october": (10, "Oct"), "nov": (11, "Nov"), "november": (11, "Nov"),
    "dec": (12, "Dec"), "december": (12, "Dec"),
}


def find_latest(pattern):
    candidates = [
        p for p in ROOT.glob("*.xlsx")
        if not p.name.startswith("~$") and pattern.match(p.name)
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


def s(v):
    if v is None:
        return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def name_key(v):
    """Normalized person-name key for joining the b2c RM name to the employee reference sheet."""
    return re.sub(r"\s+", " ", str(v or "").strip()).lower()


def parse_date(v):
    """Excel dates arrive as datetimes; b2c exports them as strings ('2026-07-28 18:36:56.0'),
    and unset values show up as 'N/A' or blank. Returns a date or None."""
    if v is None:
        return None
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    t = str(v).strip().replace("T", " ")
    if not t or t.upper() in ("N/A", "NA", "-", "NULL", "NONE"):
        return None
    # %d.%m.%Y is how the FY sheets write "Tkt Recd Date" (01.04.2026).
    for fmt in ("%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y"):
        try:
            return datetime.datetime.strptime(t[:26], fmt).date()
        except ValueError:
            pass
    try:
        return datetime.date.fromisoformat(t[:10])
    except ValueError:
        return None


def day_gap(later, earlier):
    """Whole days between two dates, or None if either side is missing."""
    if later is None or earlier is None:
        return None
    return (later - earlier).days


def load_team_map(path):
    """Employee reference sheet -> {normalized RM name: team}. Used to roll b2c leads up to a team."""
    if path is None:
        return {}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        it = ws.iter_rows(values_only=True)
        header = [s(x) for x in next(it)]
        try:
            n_i, t_i = header.index("Name"), header.index("Team")
        except ValueError:
            return {}
        out = {}
        for r in it:
            n, t = name_key(r[n_i]), s(r[t_i])
            if n and t:
                out[n] = t
        # Register known b2c spellings against the employee-sheet entry they refer to. If the
        # canonical name is missing (say the sheet was re-exported and the person left), the alias
        # is skipped and the usual "no row in the employee reference sheet" warning still fires.
        for variant, canonical in RM_NAME_ALIASES.items():
            if canonical in out and variant not in out:
                out[variant] = out[canonical]
        return out
    finally:
        wb.close()


def norm_key(v):
    """Canonical join key for matching an email between the two sheets. Strips ALL whitespace
    (not just leading/trailing — a stray space anywhere breaks equality) plus invisible/control
    Unicode characters (zero-width space, BOM, etc.) that plain .strip() doesn't catch, then
    lowercases. Used only for matching; display values keep their original spelling via s()."""
    if v is None:
        return ""
    text = "".join(ch for ch in str(v) if not unicodedata.category(ch).startswith("C"))
    return re.sub(r"\s+", "", text).lower()


def load_lead_master(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb["Data"]
        it = ws.iter_rows(values_only=True)
        header = [s(x) for x in next(it)]
        idx = {name: header.index(name) for name in header if name}
        lead = {}
        for r in it:
            email = norm_key(r[idx["userId"]])
            if not email:
                continue
            lead[email] = {
                "leadStatus": s(r[idx["leadStatus"]]).upper() or "BLANK",
                "rm": s(r[idx["currentRmName"]]),
                "leadHead": s(r[idx["leadHead"]]),
                "created": s(r[idx["createdDate"]])[:10],
                "lastStatus": s(r[idx["lastStatusDate"]])[:10],
                "converted": s(r[idx["convertedDate"]])[:10],
                "landing": s(r[idx["landingPage"]]),
                "platform": s(r[idx["platformName"]]),
                "category": s(r[idx["categoryName"]]),
                "isClient": s(r[idx["isClient"]]),
                "convertedDt": parse_date(r[idx["convertedDate"]]),
                "inProcessDt": parse_date(r[idx["leadInProcessDate"]]),
                "createdDt": parse_date(r[idx["createdDate"]]),
                "clientCategory": s(r[idx["clientCategory"]]),
            }
        return lead
    finally:
        wb.close()


def load_all_rows(path, lead, team_map):
    """Reads all four data sheets into a de-duplicated client universe keyed on email.

    Returns (tickets, clients, diag):
      tickets - one record per source row that carries an email. Plan-Approval tickets additionally
                carry revenue and the turnaround measures; FY tickets carry a Ticket Subject.
      clients - one record per distinct email, unioned across every sheet.
      diag    - counts for the build log, so rows dropped here are reported rather than vanishing.

    A row with no email cannot be joined to the lead data or de-duplicated against the other
    sheets, so it is skipped and counted in diag instead of being silently folded in.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        tickets = []
        diag = {"no_email": 0, "excluded_month": 0, "bad_dates": 0,
                "per_sheet": {}, "missing_sheets": []}

        for cfg in SHEETS:
            if cfg["name"] not in wb.sheetnames:
                diag["missing_sheets"].append(cfg["name"])
                continue
            ws = wb[cfg["name"]]
            raw = list(ws.iter_rows(values_only=True))
            if not raw:
                continue
            hdr = [s(c) for c in raw[0]]
            col = {name: i for i, name in enumerate(hdr) if name}
            kept = 0

            def g(r, *names):
                """First non-empty value among the given header names — quarters of the same sheet
                spell some columns differently (e.g. 'Lead Status' vs 'Lead Statues')."""
                for name in names:
                    i = col.get(name)
                    if i is not None and i < len(r):
                        v = s(r[i])
                        if v:
                            return v
                return ""

            def gnum(r, name):
                i = col.get(name)
                if i is None or i >= len(r):
                    return 0.0
                v = r[i]
                if isinstance(v, (int, float)):
                    return float(v)
                try:
                    return float(str(v).replace(",", "").strip() or 0)
                except ValueError:
                    return 0.0

            def gdate(r, name):
                i = col.get(name)
                return parse_date(r[i]) if i is not None and i < len(r) else None

            for r in raw[1:]:
                month_raw = g(r, cfg["month"])
                if month_raw.strip().lower() in EXCLUDE_MONTHS:
                    diag["excluded_month"] += 1
                    continue

                email = g(r, cfg["email"])
                key = norm_key(email)
                key = EMAIL_ALIASES.get(key, key)
                if not key:
                    # No email: unjoinable and un-dedupable. Counted, not guessed at.
                    if g(r, "Client Name"):
                        diag["no_email"] += 1
                    continue

                # The typed date is not trustworthy on its own: the FY sheets contain finger-slips
                # like 11.06.2926 and 04.07.2027 whose Month column still reads correctly. So the
                # month name drives the bucket, and a date outside the financial year is discarded
                # rather than allowed to invent a month (or a nonsense turnaround).
                row_dt = gdate(r, cfg["date"])
                if row_dt and not (FY_START <= row_dt <= FY_END):
                    diag["bad_dates"] += 1
                    row_dt = None

                mnum, mlabel = MONTHS.get(month_raw.lower(), (0, ""))
                if not mnum and row_dt:
                    mnum, mlabel = row_dt.month, row_dt.strftime("%b")
                if not mnum:
                    mlabel = month_raw or "Unknown"
                # FY 2026-27: April-December falls in 2026, January-March in 2027.
                year = 2026 if mnum >= 4 else 2027
                lead_rec = lead.get(key)
                rm_name = lead_rec["rm"] if lead_rec else ""

                rec = {
                    "kind": cfg["kind"],
                    "q": cfg["q"],
                    "mk": "%04d-%02d" % (year, mnum) if mnum else "zzzz",
                    "m": "%s %d" % (mlabel, year) if mnum else (mlabel or "Unknown"),
                    "date": s(row_dt) if row_dt else "",
                    "key": key,
                    "email": email,
                    "name": g(r, "Client Name"),
                    "subject": g(r, "Ticket Subject") if cfg["kind"] == "fy" else "",
                    "tierRaw": client_tier(g(r, "Client Type")) if cfg["kind"] == "fy" else "",
                    "tenureRaw": client_tenure(g(r, "Client Type")) if cfg["kind"] == "pa" else "",
                    "ticket": g(r, "Ticket Id", "SR No"),
                    "advisor": scrub(g(r, "Advisor", "RM Name")),
                    "appr": g(r, "Approved By"),
                    "sheetSrc": scrub(g(r, "Lead Source")),
                    # joined from the lead export
                    "status": lead_rec["leadStatus"] if lead_rec else "NOT IN LEAD DATA",
                    "matched": bool(lead_rec),
                    "rm": scrub(rm_name),
                    "team": team_map.get(name_key(rm_name), "Unassigned" if rm_name else "No lead record"),
                    "platform": lead_rec["platform"] if lead_rec else "",
                    "src": scrub(lead_rec["landing"] if lead_rec and lead_rec["landing"] else g(r, "Lead Source")),
                    "b2cCategory": lead_rec["clientCategory"] if lead_rec else "",
                    "created": lead_rec["created"] if lead_rec else "",
                    "converted": lead_rec["converted"] if lead_rec else "",
                }

                if cfg["kind"] == "pa":
                    raised_dt = row_dt
                    approved_dt = gdate(r, "Approval Date")
                    converted_dt = lead_rec["convertedDt"] if lead_rec else None
                    rec.update({
                        "appdate": s(approved_dt) if approved_dt else "",
                        "verdict": g(r, "Approved/Rejected"),
                        "rev": {k: gnum(r, header) for k, header in PRODUCTS},
                        "convertPreDraftDays": ((raised_dt - converted_dt).days
                                                if converted_dt and raised_dt and converted_dt < raised_dt else None),
                        "tatApprove": day_gap(approved_dt, raised_dt),
                        "tatConvert": day_gap(converted_dt, approved_dt),
                        "tatInProcess": day_gap(lead_rec["inProcessDt"] if lead_rec else None, raised_dt),
                    })
                else:
                    rec.update({"appdate": "", "verdict": "", "rev": {k: 0.0 for k, _ in PRODUCTS},
                                "convertPreDraftDays": None,
                                "tatApprove": None, "tatConvert": None, "tatInProcess": None})

                tickets.append(rec)
                kept += 1

            diag["per_sheet"][cfg["name"]] = kept

        clients = build_clients(tickets)
        # Stamp each ticket with its client's resolved tier / tenure / subject-set so the
        # ticket-level cube (revenue, turnaround) can be sliced by exactly the same dimensions and
        # the same master filter as the client-level one.
        by_key = {c["key"]: c for c in clients}
        for t in tickets:
            c = by_key.get(t["key"])
            if c:
                t["tier"], t["tenure"] = c["tier"], c["tenure"]
                t["subjKey"], t["inPA"] = c["subjKey"], c["inPA"]
            else:
                t["tier"] = t["tenure"] = BLANK
                t["subjKey"], t["inPA"] = "", False
        return tickets, clients, diag
    finally:
        wb.close()


def build_clients(tickets):
    """Collapses the ticket records into one row per email — the de-duplicated union.

    A client's month is the EARLIEST they appear across every sheet (when they entered the
    pipeline), so each client lands in exactly one month and the month-wise table still totals to
    the client count. Revenue sums across all of that client's plan approvals.
    """
    by = {}
    for t in sorted(tickets, key=lambda r: (r["mk"], r["date"], r["ticket"])):
        c = by.get(t["key"])
        if c is None:
            c = by[t["key"]] = {
                "key": t["key"], "email": t["email"], "name": t["name"],
                "mk": t["mk"], "m": t["m"], "q": t["q"],
                "inPA": False, "subjects": set(), "tier": "", "tenure": "",
                "rev": {k: 0.0 for k, _ in PRODUCTS},
                "nTickets": 0, "nPA": 0, "nFY": 0,
                "advisor": "", "appr": "", "ticket": "",
                "status": t["status"], "matched": t["matched"], "rm": t["rm"], "team": t["team"],
                "platform": t["platform"], "src": t["src"], "b2cCategory": t["b2cCategory"],
                "created": t["created"], "converted": t["converted"],
                "date": t["date"],
            }
        c["nTickets"] += 1
        if t["kind"] == "pa":
            c["inPA"] = True
            c["nPA"] += 1
            for k, _ in PRODUCTS:
                c["rev"][k] += t["rev"].get(k, 0.0)
            # Plan-Approval fields win for the client-level record: they are the richer source.
            if t["advisor"]:
                c["advisor"] = t["advisor"]
            if t["appr"]:
                c["appr"] = t["appr"]
            if t["ticket"]:
                c["ticket"] = t["ticket"]
        else:
            c["nFY"] += 1
            if t["subject"]:
                c["subjects"].add(t["subject"])
        if t["tierRaw"] and not c["tier"]:
            c["tier"] = t["tierRaw"]
        if t["tenureRaw"] and not c["tenure"]:
            c["tenure"] = t["tenureRaw"]
        if not c["advisor"] and t["advisor"]:
            c["advisor"] = t["advisor"]
        if not c["name"] and t["name"]:
            c["name"] = t["name"]

    out = []
    for c in by.values():
        subs = sorted(c["subjects"])
        c["subjects"] = subs
        # One value per client, so the master filter can be applied to an aggregated cell without
        # a client ever being counted under two different subjects.
        c["subjKey"] = "|".join(subs)
        c["tier"] = c["tier"] or BLANK
        c["tenure"] = c["tenure"] or BLANK
        c["rev"] = {k: round(v, 2) for k, v in c["rev"].items()}
        out.append(c)
    out.sort(key=lambda r: (r["mk"], r["name"].lower()))
    return out


def collect_strings(node, out):
    """Every string value reachable in the payload, ignoring numbers. Measures (revenue sums,
    turnaround counts) are numeric by construction, and a revenue figure like 2000000 would
    otherwise trip the phone-number heuristic. Anything genuinely identifying — an email, or a
    phone pasted into a text column — arrives as a string and is still caught."""
    if isinstance(node, str):
        out.append(node)
    elif isinstance(node, dict):
        for v in node.values():
            collect_strings(v, out)
    elif isinstance(node, list):
        for v in node:
            collect_strings(v, out)


def embed(template_path, payload, pii_scan_keys=None):
    """pii_scan_keys: top-level payload keys sourced from row data (advisor/src/team/...) to check
    for email/phone-like leftovers. Deliberately excludes 'meta' — filenames legitimately contain
    long digit runs (report IDs, timestamps) that would false-positive as phone numbers."""
    text = template_path.read_text(encoding="utf-8")
    data = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    if "</script" in data.lower():
        raise ValueError("payload contains a literal </script — refusing to embed unsafely")
    if pii_scan_keys:
        strings = []
        for k in pii_scan_keys:
            collect_strings(payload.get(k), strings)
        for value in strings:
            hit = LOOKS_LIKE_PII_RE.search(value)
            if hit:
                raise ValueError(
                    f"Refusing to write the public dashboard: found an email/phone-like value "
                    f"({hit.group()!r}) inside {value!r} in {pii_scan_keys}. This should be "
                    f"impossible after scrub() — stopping instead of risking a public leak. Check "
                    f"Advisor/Lead Source data, or whether a new text field was added without scrubbing."
                )
    if "__DATA__" not in text:
        raise ValueError(f"{template_path.name} has no __DATA__ placeholder")
    return text.replace("__DATA__", data)


def main():
    b2c_path = find_latest(B2C_RE)
    plan_path = find_latest(PLAN_RE)

    if not b2c_path:
        print("ERROR: no b2c/lead export found in this folder (expected a file matching 'FIN<digits>_*.xlsx').")
        return 1
    if not plan_path:
        print("ERROR: no plan approval workbook found in this folder (expected 'Financial Planning Tickets Summary-Dashboard*.xlsx').")
        return 1

    emp_path = find_latest(EMP_RE)

    print(f"b2c/lead file   : {b2c_path.name}")
    print(f"planning file   : {plan_path.name}")
    print(f"employee ref    : {emp_path.name if emp_path else '(none found - team roll-up disabled)'}")

    team_map = load_team_map(emp_path)
    print(f"  {len(team_map)} RM -> team mappings")

    print("Reading lead master...")
    lead = load_lead_master(b2c_path)
    print(f"  {len(lead):,} leads loaded")

    print("Reading all four data sheets (Plan Approval Q1/Q2 + FY Q1/Q2)...")
    tickets, clients, diag = load_all_rows(plan_path, lead, team_map)
    pa_tickets = [t for t in tickets if t["kind"] == "pa"]
    for sheet, n in diag["per_sheet"].items():
        print(f"    {sheet:<26} {n:>5} usable rows")
    if diag["missing_sheets"]:
        print(f"  WARNING: sheet(s) not found in the workbook: {diag['missing_sheets']}")
    print(f"  excluded {diag['excluded_month']} rows from {'/'.join(sorted(EXCLUDE_MONTHS)).title()} (by request)")
    print(f"  skipped  {diag['no_email']} rows that carry a client name but no email "
          f"(cannot be joined to lead data or de-duplicated)")
    if diag["bad_dates"]:
        print(f"  {diag['bad_dates']} rows had a date outside FY 2026-27 (typo years such as "
              f"11.06.2926) - bucketed by their Month column instead")

    matched = sum(1 for c in clients if c["matched"])
    in_pa = sum(1 for c in clients if c["inPA"])
    fy_only = len(clients) - in_pa
    print(f"  {len(tickets)} rows -> {len(clients)} unique clients in the de-duplicated union")
    print(f"    {in_pa} appear on a Plan Approval sheet (always counted)")
    print(f"    {fy_only} come only from the FY sheets (counted subject to the Ticket Subject filter)")
    print(f"    {matched} matched to lead data, {len(clients) - matched} unmapped")

    unresolved = sorted({c["rm"] for c in clients if c["matched"] and c["team"] == "Unassigned"})
    print(f"  teams: {len(set(c['team'] for c in clients))} distinct; "
          f"{sum(1 for c in clients if c['team'] not in ('Unassigned', 'No lead record'))} clients mapped to a team")
    if unresolved:
        print(f"  WARNING: {len(unresolved)} RM name(s) in b2c have no row in the employee reference sheet:")
        for n in unresolved[:10]:
            print(f"    - {n}")
        if len(unresolved) > 10:
            print(f"    ... and {len(unresolved) - 10} more")

    # Turnaround is measured on plan approvals only - the FY sheets carry no approval date.
    for name, field in TAT_FIELDS:
        universe = TAT_UNIVERSE[name]
        scope = [r for r in pa_tickets if not universe or r["status"] == universe]
        have = [r for r in scope if r[field] is not None]
        usable = [r for r in have if r[field] >= 0]
        who = f"status={universe}" if universe else "all plans"
        line = (f"  TAT days-to-{name}: universe {len(scope)} ({who}); "
                f"{len(have)} have both dates; {len(usable)} usable (>=0)")
        neg = [r for r in have if r[field] < 0]
        if name in TAT_SPLIT_NEGATIVES:
            old = sum(1 for r in neg
                      if r["convertPreDraftDays"] is not None
                      and r["convertPreDraftDays"] > TAT_OLD_CLIENT_DAYS)
            line += (f"; {len(neg)} negative -> {old} old clients (>{TAT_OLD_CLIENT_DAYS}d before "
                     f"drafting), {len(neg) - old} logged before drafting")
        elif neg:
            line += f"; {len(neg)} before the plan was raised (reported as one neutral row)"
        print(line)

    tiers, tenures, subjects = {}, {}, {}
    for c in clients:
        tiers[c["tier"]] = tiers.get(c["tier"], 0) + 1
        tenures[c["tenure"]] = tenures.get(c["tenure"], 0) + 1
        for sub in c["subjects"]:
            subjects[sub] = subjects.get(sub, 0) + 1
    print(f"  Client Tier   (FY sheets):        {tiers}")
    print(f"  Client Tenure (Plan Approval):    {tenures}")
    print(f"  Ticket Subject values: {len(subjects)} distinct")
    for k, v in sorted(subjects.items(), key=lambda kv: -kv[1]):
        print(f"    {v:>5}  {k}")

    all_subjects = sorted({sub for c in clients for sub in c["subjects"]})

    meta = {
        "generated": datetime.datetime.now().strftime("%d %b %Y, %I:%M %p"),
        "dashFile": plan_path.name,
        "leadFile": b2c_path.name,
        "empFile": emp_path.name if emp_path else "",
        "leadRows": len(lead),
        "tabs": [c["name"] for c in SHEETS],
        "excludedMonths": sorted(m.title() for m in EXCLUDE_MONTHS),
        "noEmailRows": diag["no_email"],
        "products": [{"k": k, "label": label} for k, label in PRODUCTS],
        "subjects": all_subjects,
        "blankLabel": BLANK,
        "tatBuckets": TAT_BUCKETS,
        "tatUniverse": TAT_UNIVERSE,
        "tatSplit": sorted(TAT_SPLIT_NEGATIVES),
        "tatEarlyLabel": TAT_EARLY_LABEL,
        "tatNeutralNegLabel": TAT_NEUTRAL_NEG_LABEL,
        "tatOldClientDays": TAT_OLD_CLIENT_DAYS,
    }

    # ---- full detail (local only) ----
    # Both pages render from ONE template. The only difference is the payload: the local build
    # carries names/emails, the published one has them stripped. Keeping a single template is what
    # stops the two pages drifting apart - divergence between them has already caused real bugs.
    unmapped_list = [
        {"mk": c["mk"], "m": c["m"], "date": c["date"], "ticket": c["ticket"],
         "tier": c["tier"], "tenure": c["tenure"], "name": c["name"], "email": c["email"],
         "advisor": c["advisor"], "subjects": ", ".join(c["subjects"])}
        for c in clients if not c["matched"]
    ]
    full_payload = {"meta": dict(meta, hasPII=True), "clients": clients,
                    "tickets": pa_tickets, "unmapped": unmapped_list}
    full_html = embed(TEMPLATES / "dashboard.template.html", full_payload)
    (ROOT / "plan-approval-lead-status.html").write_text(full_html, encoding="utf-8")
    print("Wrote plan-approval-lead-status.html (full detail, local only)")

    # ---- published build ----
    # De-identified: no client name, email or ticket id on any analysable record. The unmapped
    # list below is the one deliberate exception - see the note there.
    STRIP = ("name", "email", "ticket", "key")

    def deid(rec):
        return {k: v for k, v in rec.items() if k not in STRIP}

    public_payload = {
        "meta": dict(meta, hasPII=False),
        "clients": [deid(c) for c in clients],
        "tickets": [deid(t) for t in pa_tickets],
        # DELIBERATE EXCEPTION: the only client-identifying data on the published page. It exists
        # so the unmapped list can actually be chased down, and it is only acceptable because the
        # repository is private. If the repo is ever made public again, remove this key - and note
        # that anything already pushed stays in git history.
        "unmapped": unmapped_list,
    }
    public_html = embed(TEMPLATES / "dashboard.template.html", public_payload,
                        pii_scan_keys=["clients", "tickets"])
    (ROOT / "index.html").write_text(public_html, encoding="utf-8")
    n_unmapped = len(public_payload["unmapped"])
    print(f"Wrote index.html (de-identified + {n_unmapped} unmapped clients WITH names/emails)")
    if n_unmapped:
        print(f"  NOTE: index.html carries {n_unmapped} client names and email addresses.")
        print("        Only publish it from a PRIVATE repository.")

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)
