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
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
TEMPLATES = Path(__file__).resolve().parent / "templates"

B2C_RE = re.compile(r"^FIN\d+_.*\.xlsx$", re.IGNORECASE)
PLAN_RE = re.compile(r"^Financial Planning Tickets Summary-Dashboard.*\.xlsx$", re.IGNORECASE)

# Advisor / Lead Source are supposed to be short internal category labels, and are the only two
# per-row fields that flow into the public aggregate cube. They're manually typed in the source
# sheet, so guard against fat-finger mistakes (e.g. a client's email or phone pasted into the
# wrong column) leaking through into the public file. Match => treat the cell as blank.
LOOKS_LIKE_PII_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+|\d{7,}")


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


def load_lead_master(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb["Data"]
        it = ws.iter_rows(values_only=True)
        header = [s(x) for x in next(it)]
        idx = {name: header.index(name) for name in header if name}
        lead = {}
        for r in it:
            email = s(r[idx["userId"]]).lower()
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
            }
        return lead
    finally:
        wb.close()


def load_plan_rows(path, lead):
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        rows_out = []
        for sheet, quarter in [("Plan Approval Sheet Q1", "Q1"), ("Plan Approval Sheet Q2", "Q2")]:
            ws = wb[sheet]
            raw = list(ws.iter_rows(values_only=True))
            hdr = [s(c) for c in raw[0]]
            col = {name: i for i, name in enumerate(hdr) if name}

            def g(r, name):
                i = col.get(name)
                return s(r[i]) if i is not None and i < len(r) else ""

            for r in raw[1:]:
                ticket = g(r, "Ticket Id")
                if not ticket:
                    continue
                email = g(r, "Email Id")
                key = email.lower()
                mraw = g(r, "D").lower()
                mnum, mlabel = MONTHS.get(mraw, (0, g(r, "D") or "Unknown"))
                date = g(r, "Date")
                year = int(date[:4]) if len(date) >= 4 and date[:4].isdigit() else datetime.date.today().year
                lead_rec = lead.get(key)
                rows_out.append({
                    "q": quarter,
                    "mk": "%04d-%02d" % (year, mnum) if mnum else "zzzz",
                    "m": "%s %d" % (mlabel, year) if mnum else (mlabel or "Unknown"),
                    "date": date,
                    "ticket": ticket,
                    "advisor": scrub(g(r, "Advisor")),
                    "name": g(r, "Client Name"),
                    "email": email,
                    "key": key,
                    "src": scrub(g(r, "Lead Source")),
                    "draft": g(r, "Drafted By"),
                    "appr": g(r, "Approved By"),
                    "appdate": g(r, "Approval Date"),
                    "verdict": g(r, "Approved/Rejected"),
                    "status": lead_rec["leadStatus"] if lead_rec else "NOT IN LEAD DATA",
                    "matched": bool(lead_rec),
                    "rm": lead_rec["rm"] if lead_rec else "",
                    "leadHead": lead_rec["leadHead"] if lead_rec else "",
                    "created": lead_rec["created"] if lead_rec else "",
                    "lastStatus": lead_rec["lastStatus"] if lead_rec else "",
                    "converted": lead_rec["converted"] if lead_rec else "",
                    "landing": lead_rec["landing"] if lead_rec else "",
                    "platform": lead_rec["platform"] if lead_rec else "",
                    "category": lead_rec["category"] if lead_rec else "",
                    "isClient": lead_rec["isClient"] if lead_rec else "",
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


def build_cube(rows):
    """Grouped counts only: (month, quarter, advisor, source, status) -> n. No names/emails survive this step."""
    counts = {}
    for r in rows:
        key = (r["mk"], r["m"], r["q"], r["advisor"], r["src"], r["status"])
        counts[key] = counts.get(key, 0) + 1
    return [
        {"mk": mk, "m": m, "q": q, "advisor": advisor, "src": src, "status": status, "n": n}
        for (mk, m, q, advisor, src, status), n in counts.items()
    ]


def embed(template_path, payload, pii_scan_keys=None):
    """pii_scan_keys: names of top-level payload keys sourced from row data (advisor/src/etc.) to
    check for email/phone-like leftovers. Deliberately excludes 'meta' — filenames legitimately
    contain long digit runs (report IDs, timestamps) that would false-positive as phone numbers."""
    text = template_path.read_text(encoding="utf-8")
    data = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    if "</script" in data.lower():
        raise ValueError("payload contains a literal </script — refusing to embed unsafely")
    if pii_scan_keys:
        scan_text = json.dumps({k: payload.get(k) for k in pii_scan_keys}, ensure_ascii=False)
        hit = LOOKS_LIKE_PII_RE.search(scan_text)
        if hit:
            raise ValueError(
                f"Refusing to write the public dashboard: found an email/phone-like value "
                f"({hit.group()!r}) in {pii_scan_keys}. This should be impossible after scrub() — "
                f"stopping instead of risking a public leak. Check Advisor/Lead Source data, or "
                f"whether a new field was added without scrubbing."
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

    print(f"b2c/lead file   : {b2c_path.name}")
    print(f"planning file   : {plan_path.name}")
    print("Reading lead master...")
    lead = load_lead_master(b2c_path)
    print(f"  {len(lead):,} leads loaded")

    print("Reading plan approval tabs (Q1 + Q2)...")
    rows = load_plan_rows(plan_path, lead)
    rows.sort(key=lambda r: (r["mk"], r["date"], r["ticket"]))
    clients = dedupe_clients(rows)
    matched = sum(1 for r in clients if r["matched"])
    print(f"  {len(rows)} plan-approval rows -> {len(clients)} unique clients ({matched} matched to lead data)")

    meta = {
        "generated": datetime.datetime.now().strftime("%d %b %Y, %I:%M %p"),
        "dashFile": plan_path.name,
        "leadFile": b2c_path.name,
        "leadRows": len(lead),
        "tabs": ["Plan Approval Sheet Q1", "Plan Approval Sheet Q2"],
    }

    # ---- full detail (local only) ----
    full_payload = {"meta": meta, "rows": rows}
    full_html = embed(TEMPLATES / "full_dashboard.template.html", full_payload)
    (ROOT / "plan-approval-lead-status.html").write_text(full_html, encoding="utf-8")
    print("Wrote plan-approval-lead-status.html (full detail, local only)")

    # ---- public aggregate (pushed to GitHub) ----
    months = sorted({(r["mk"], r["m"]) for r in rows})
    advisors = sorted({r["advisor"] for r in rows if r["advisor"]})
    sources = sorted({r["src"] for r in rows if r["src"]})
    public_payload = {
        "meta": meta,
        "months": [{"mk": mk, "m": m} for mk, m in months],
        "advisors": advisors,
        "sources": sources,
        "cube": {"tickets": build_cube(rows), "clients": build_cube(clients)},
    }
    public_html = embed(TEMPLATES / "public_dashboard.template.html", public_payload, pii_scan_keys=["cube", "advisors", "sources"])
    (ROOT / "index.html").write_text(public_html, encoding="utf-8")
    print("Wrote index.html (aggregate only, no client names/emails)")

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)
