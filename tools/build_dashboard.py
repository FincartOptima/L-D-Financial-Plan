"""
Rebuilds both dashboards from the latest source workbooks in the project folder:

  - plan-approval-lead-status.html   full detail, client names + emails, LOCAL ONLY (git-ignored)
  - index.html                       aggregate counts only, no client-level data, PUBLISHED to GitHub Pages

Source files are auto-detected by pattern + picked by most-recent modified time, so this
still works after the b2c export and the planning workbook get replaced with new filenames:

  - b2c / lead export:      FIN<digits>_*.xlsx      (sheet 'Data')
  - plan approval workbook: Financial Planning Tickets Summary-Dashboard*.xlsx
                             (tabs 'Plan Approval Sheet Q1' and 'Plan Approval Sheet Q2')

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
    for fmt in ("%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
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


def load_plan_rows(path, lead, team_map):
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        rows_out = []
        for sheet, quarter in [("Plan Approval Sheet Q1", "Q1"), ("Plan Approval Sheet Q2", "Q2")]:
            ws = wb[sheet]
            raw = list(ws.iter_rows(values_only=True))
            hdr = [s(c) for c in raw[0]]
            col = {name: i for i, name in enumerate(hdr) if name}
            # Columns are looked up by header name, never by position: the Q1 and Q2 tabs have
            # different layouts (Q2 adds Lead Status / Client Status), so positional reads would
            # silently pull the wrong column on one of them.

            def g(r, name):
                i = col.get(name)
                return s(r[i]) if i is not None and i < len(r) else ""

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
                ticket = g(r, "Ticket Id")
                if not ticket:
                    continue
                email = g(r, "Email Id")
                key = norm_key(email)
                key = EMAIL_ALIASES.get(key, key)
                mraw = g(r, "D").lower()
                mnum, mlabel = MONTHS.get(mraw, (0, g(r, "D") or "Unknown"))
                date = g(r, "Date")
                year = int(date[:4]) if len(date) >= 4 and date[:4].isdigit() else datetime.date.today().year
                lead_rec = lead.get(key)

                raised_dt = gdate(r, "Date")
                approved_dt = gdate(r, "Approval Date")
                # How many days before the plan was raised (col B) the client converted, or None if
                # the conversion is not before it. This is what separates a returning client from a
                # conversion logged just ahead of drafting.
                converted_dt = lead_rec["convertedDt"] if lead_rec else None
                pre_draft = ((raised_dt - converted_dt).days
                             if converted_dt and raised_dt and converted_dt < raised_dt else None)
                rev = {k: gnum(r, header) for k, header in PRODUCTS}
                rm_name = lead_rec["rm"] if lead_rec else ""
                rows_out.append({
                    "q": quarter,
                    "mk": "%04d-%02d" % (year, mnum) if mnum else "zzzz",
                    "m": "%s %d" % (mlabel, year) if mnum else (mlabel or "Unknown"),
                    "date": date,
                    "ticket": ticket,
                    "advisor": scrub(g(r, "Advisor")),
                    "name": g(r, "Client Name"),
                    "email": email,
                    "clientType": g(r, "Client Type"),
                    "key": key,
                    # Lead source comes from the b2c record the email mapped to (landingPage —
                    # the column whose vocabulary matches the planning sheet: Direct Registration,
                    # FinancialPlan_VG, Workshop...). Unmatched clients have no b2c record, so they
                    # fall back to whatever the planning sheet recorded rather than going blank.
                    "src": scrub(lead_rec["landing"] if lead_rec and lead_rec["landing"] else g(r, "Lead Source")),
                    "srcSheet": scrub(g(r, "Lead Source")),
                    "draft": g(r, "Drafted By"),
                    "appr": g(r, "Approved By"),
                    "appdate": g(r, "Approval Date"),
                    "verdict": g(r, "Approved/Rejected"),
                    "status": lead_rec["leadStatus"] if lead_rec else "NOT IN LEAD DATA",
                    "matched": bool(lead_rec),
                    # Scrubbed like advisor/src: rm now also flows into the public cube as a
                    # dimension, so a fat-fingered email/phone in this CRM field should blank out
                    # rather than block the whole build via the PII guard.
                    "rm": scrub(lead_rec["rm"]) if lead_rec else "",
                    "leadHead": lead_rec["leadHead"] if lead_rec else "",
                    "created": lead_rec["created"] if lead_rec else "",
                    "lastStatus": lead_rec["lastStatus"] if lead_rec else "",
                    "converted": lead_rec["converted"] if lead_rec else "",
                    "landing": lead_rec["landing"] if lead_rec else "",
                    "platform": lead_rec["platform"] if lead_rec else "",
                    "category": lead_rec["category"] if lead_rec else "",
                    "isClient": lead_rec["isClient"] if lead_rec else "",
                    "team": team_map.get(name_key(rm_name), "Unassigned" if rm_name else "No lead record"),
                    "rev": rev,
                    # TAT metrics, in whole days. null = the underlying date is missing, in which
                    # case the row is left out of that metric's denominator rather than counted as 0.
                    # Each is measured from a different baseline on purpose:
                    #   approve   = plan raised (col B)    -> plan approved (col M)
                    #   convert   = plan approved (col M)  -> converted (b2c convertedDate)
                    #   inprocess = plan raised (col B)    -> in-process (b2c leadInProcessDate)
                    "convertPreDraftDays": pre_draft,
                    # Client category as the lead system holds it, for the clients whose email
                    # mapped. Kept separate from the planning sheet's own "Client Type" column.
                    "b2cCategory": lead_rec["clientCategory"] if lead_rec else "",
                    "tatApprove": day_gap(approved_dt, raised_dt),
                    "tatConvert": day_gap(lead_rec["convertedDt"] if lead_rec else None, approved_dt),
                    "tatInProcess": day_gap(lead_rec["inProcessDt"] if lead_rec else None, raised_dt),
                })
        return rows_out
    finally:
        wb.close()


def dedupe_clients(rows):
    """One row per client email: the row with the latest (month, date, ticket) wins.
    Mirrors the identical tie-break used client-side in the full dashboard's JS."""
    by = {}
    for r in rows:
        prev = by.get(r["key"])
        if not prev or (r["mk"] + r["date"] + r["ticket"]) > (prev["mk"] + prev["date"] + prev["ticket"]):
            by[r["key"]] = r
    return list(by.values())


DIMS = ("mk", "m", "q", "advisor", "src", "team", "rm", "platform", "b2cCategory", "status")


def build_cube(rows, with_measures=False):
    """Groups rows down to (month, quarter, advisor, source, team, RM, platform, client type,
    status). No client names, emails, ticket ids or dates survive this step, so the public page can
    slice these dimensions but never reach a named client. RM is a staff name, not a client's — same
    sensitivity class as advisor, which the cube already carries.

    with_measures adds revenue sums and turnaround bucket counts, which are per-approval figures —
    only pass it for the ticket-level cube, never the de-duplicated client one."""
    cells = {}
    for r in rows:
        key = tuple(r[d] for d in DIMS)
        c = cells.get(key)
        if c is None:
            c = cells[key] = {"n": 0}
            if with_measures:
                c["rev"] = {k: 0.0 for k, _ in PRODUCTS}
                c["tat"] = {name: {"d": 0, "neg": 0, "negOld": 0, "negEarly": 0,
                                   "miss": 0, "b": [0] * len(TAT_BUCKETS)}
                            for name, _ in TAT_FIELDS}
        c["n"] += 1
        if with_measures:
            for k, _ in PRODUCTS:
                c["rev"][k] += r["rev"].get(k, 0.0)
            for name, field in TAT_FIELDS:
                universe = TAT_UNIVERSE[name]
                if universe and r["status"] != universe:
                    continue          # out of scope for this measure — not missing, just not asked
                v, t = r[field], c["tat"][name]
                if v is None:
                    t["miss"] += 1
                elif v < 0:
                    # Event predates its baseline, so it is held out of the denominator rather than
                    # counted as a zero-day turnaround. For convert/in-process it is also split by
                    # whether the lead pre-dated the plan (see TAT_SPLIT_NEGATIVES).
                    t["neg"] += 1
                    if name in TAT_SPLIT_NEGATIVES:
                        gap = r["convertPreDraftDays"]
                        if gap is None:
                            # converted after drafting but before approval — still ahead of the
                            # plan being signed off, so it belongs with the flagged group
                            t["negEarly"] += 1
                        elif gap > TAT_OLD_CLIENT_DAYS:
                            t["negOld"] += 1
                        else:
                            t["negEarly"] += 1
                else:
                    t["d"] += 1
                    for i, b in enumerate(TAT_BUCKETS):
                        if v > b:
                            t["b"][i] += 1

    out = []
    for key, c in cells.items():
        cell = dict(zip(DIMS, key))
        cell["n"] = c["n"]
        if with_measures:
            cell["rev"] = {k: round(v, 2) for k, v in c["rev"].items()}
            cell["tat"] = c["tat"]
        out.append(cell)
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

    print("Reading plan approval tabs (Q1 + Q2)...")
    rows = load_plan_rows(plan_path, lead, team_map)
    rows.sort(key=lambda r: (r["mk"], r["date"], r["ticket"]))
    clients = dedupe_clients(rows)
    matched = sum(1 for r in clients if r["matched"])
    print(f"  {len(rows)} plan-approval rows -> {len(clients)} unique clients ({matched} matched to lead data)")

    # Diagnostics for the newly-joined dimensions, so a silently-empty section is obvious here
    # rather than only being noticed as a blank card on the dashboard.
    unresolved = sorted({r["rm"] for r in clients if r["matched"] and r["team"] == "Unassigned"})
    print(f"  teams: {len(set(r['team'] for r in clients))} distinct; "
          f"{sum(1 for r in clients if r['team'] not in ('Unassigned', 'No lead record'))} clients mapped to a team")
    if unresolved:
        print(f"  WARNING: {len(unresolved)} RM name(s) in b2c have no row in the employee reference sheet:")
        for n in unresolved[:10]:
            print(f"    - {n}")
        if len(unresolved) > 10:
            print(f"    ... and {len(unresolved) - 10} more")
    for name, field in TAT_FIELDS:
        universe = TAT_UNIVERSE[name]
        scope = [r for r in rows if not universe or r["status"] == universe]
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
    ctypes = {}
    for r in clients:
        ctypes[r["clientType"] or "(blank)"] = ctypes.get(r["clientType"] or "(blank)", 0) + 1
    print(f"  Client Type (planning sheet): {ctypes}")
    b2ccat = {}
    for r in clients:
        b2ccat[r["b2cCategory"] or "(blank)"] = b2ccat.get(r["b2cCategory"] or "(blank)", 0) + 1
    print(f"  clientCategory (b2c):        {b2ccat}")

    meta = {
        "generated": datetime.datetime.now().strftime("%d %b %Y, %I:%M %p"),
        "dashFile": plan_path.name,
        "leadFile": b2c_path.name,
        "empFile": emp_path.name if emp_path else "",
        "leadRows": len(lead),
        "tabs": ["Plan Approval Sheet Q1", "Plan Approval Sheet Q2"],
    }

    # ---- full detail (local only) ----
    meta = dict(meta,
                products=[{"k": k, "label": label} for k, label in PRODUCTS],
                tatUniverse=TAT_UNIVERSE,
                tatSplit=sorted(TAT_SPLIT_NEGATIVES),
                tatEarlyLabel=TAT_EARLY_LABEL,
                tatNeutralNegLabel=TAT_NEUTRAL_NEG_LABEL,
                tatOldClientDays=TAT_OLD_CLIENT_DAYS)
    full_payload = {"meta": meta, "rows": rows}
    full_html = embed(TEMPLATES / "full_dashboard.template.html", full_payload)
    (ROOT / "plan-approval-lead-status.html").write_text(full_html, encoding="utf-8")
    print("Wrote plan-approval-lead-status.html (full detail, local only)")

    # ---- public aggregate (pushed to GitHub) ----
    months = sorted({(r["mk"], r["m"]) for r in rows})
    public_payload = {
        "meta": dict(meta, tatBuckets=TAT_BUCKETS),
        "months": [{"mk": mk, "m": m} for mk, m in months],
        "advisors": sorted({r["advisor"] for r in rows if r["advisor"]}),
        "sources": sorted({r["src"] for r in rows if r["src"]}),
        "teams": sorted({r["team"] for r in rows if r["team"]}),
        "platforms": sorted({r["platform"] for r in rows if r["platform"]}),
        "cube": {"tickets": build_cube(rows, with_measures=True), "clients": build_cube(clients)},
        # DELIBERATE EXCEPTION: this is the only client-identifying data on the published page.
        # It exists so the unmapped list can be chased down, and it is only acceptable because the
        # repository is private. If the repo is ever made public again, remove this key — and note
        # that anything already pushed stays in git history.
        "unmapped": [
            {"mk": r["mk"], "m": r["m"], "date": r["date"], "ticket": r["ticket"],
             "clientType": r["clientType"], "name": r["name"], "email": r["email"],
             "advisor": r["advisor"], "appr": r["appr"]}
            for r in sorted(clients, key=lambda x: (x["mk"], x["date"], x["ticket"]))
            if not r["matched"]
        ],
    }
    public_html = embed(TEMPLATES / "public_dashboard.template.html", public_payload,
                        pii_scan_keys=["cube", "advisors", "sources", "teams", "platforms"])
    (ROOT / "index.html").write_text(public_html, encoding="utf-8")
    n_unmapped = len(public_payload["unmapped"])
    print(f"Wrote index.html (aggregates + {n_unmapped} unmapped clients WITH names/emails)")
    if n_unmapped:
        print(f"  NOTE: index.html now carries {n_unmapped} client names and email addresses.")
        print("        Only publish it from a PRIVATE repository.")

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)
