#!/usr/bin/env python3
"""Refresh the data built into index.html from the Client Directory workbook.

Usage:
    pip install openpyxl
    python build_data.py "path/to/Client Directory.xlsx"

index.html is self-contained: the data, fonts and icon are built into the file.
This script replaces only the data block inside index.html (same folder).

Reads the main directory tab, keeps one record per CMX ID (one row per
database), applies the scope rules below and writes dashboard-data.js.
"""
import datetime
import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

import openpyxl

SHEET = "Main - Customer Directory"
OUT = Path(__file__).resolve().parent / "index.html"

# Columns are located by header text, so inserted or moved columns don't break the build.
HEADERS = {
    "account": "Account", "ecr": "ECR ID", "dbname": "Database Name", "cmx": "CMX ID",
    "value": "Contract Value", "phase": "Phase", "status": "Migration Status",
    "clh_status": "CLH Customer Account Status (Active or Canceled)",
    "client_type": "Client Type", "learners": "Number of Active Learners",
    "admins": "Number of Active Admins", "country": "Country",
    "hr": "Functionality: Automated User Upload (HR Download)",
    "scorm": "Functionality: SCORM Uploader (SCO12 file type)",
    "orgpicker": "Multi-database (Org Picker)",
    "linked_tests": "Linked Tests", "linkedcontent": "LinkedContent (Tests)",
    "selfreg": "Self-Registration Link on Login Page",
    "selfenroll": "Self-Enroll: eLearning/Skills Turned On",
    "ck_sub": "CK Nursing Subscription", "ck_link": "CK Nursing Link on Skills Instance",
    "change_type": "Change Type",
    "sso": "SSO Eligibility", "skills": "Functionality: Mosby Skills",
}

# Scope: databases still to migrate.
# If the main tab has Excel filters set, those filters define the base scope.
# Otherwise these fallback rules apply; they mirror the filters used previously
# (Phase excludes Cancel/Remove and other non-migration phases; Migration Status
# excludes Complete, NA, Net New, TTP Cutover and Handoff to AM).
EXCLUDE_PHASE = {"Cancel/Remove", "Cancelled", "Internal", "NA", "Data Moved to CLH", "DirectCourse", "REVIEW"}
EXCLUDE_STATUS = {"Complete", "NA", "Net New", "TTP Cutover", "Handoff to AM"}
REPORTED_STATUS = ["Complete", "Net New", "TTP Cutover"]
EXCLUDE_CLIENT_TYPE = {"Workforce Development", "Internal MCS Site",
                       "Connect Partner", "EPM Connect", "LMS Connect", "Clinical Skills Connect"}
ANZ_COUNTRIES = {"Australia", "New Zealand"}
CONNECT_TYPES = {"Connect Partner", "EPM Connect", "LMS Connect", "Clinical Skills Connect"}
ID_PLACEHOLDERS = {"", "NULL", "N/A", "NA", "NONE", "-", "NEW CUSTOMER", "TTP CUT OVER", "TTP ONLY"}


def yes(v):
    return isinstance(v, str) and v.strip().lower().startswith("yes")


def num(v):
    return int(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0


def clean_id(v):
    s = "" if v is None else str(v).strip()
    return "" if s.upper() in ID_PLACEHOLDERS else s


# Gates: key, label, target phase, test(row, col)
GATES = [
    ("orgpicker", "Multi-database (Org Picker)", "Not phased", lambda r, c: yes(r[c["orgpicker"]])),
    ("hr", "HR Download (automated user management)", "Phase 4", lambda r, c: yes(r[c["hr"]])),
    ("selfreg", "Self-registration / self-enrollment", "Not phased",
     lambda r, c: yes(r[c["selfreg"]]) or yes(r[c["selfenroll"]])),
    ("scorm", "SCORM Uploader (SCO12)", "Phase 5", lambda r, c: yes(r[c["scorm"]])),
    ("linked", "LinkedContent / Linked Tests", "Phase 5",
     lambda r, c: yes(r[c["linkedcontent"]]) or yes(r[c["linked_tests"]])),
    ("ck", "CK Nursing (subscription / Skills link)", "Not phased",
     lambda r, c: yes(r[c["ck_sub"]]) or yes(r[c["ck_link"]])),
]


def phase_group(p):
    if not p:
        return "Unassigned"
    p = str(p).strip()
    for k in ["Phase 3", "Phase 4", "Phase 5", "Phase 6", "Phase 7", "Phase 2"]:
        if p.startswith(k):
            return k
    return "Other / review"


def read_autofilter(path, sheet):
    """Return {column_index: (allowed_values, blanks_allowed)} from the sheet's Excel filter, or {}."""
    try:
        z = zipfile.ZipFile(path)
        wbx = z.read("xl/workbook.xml").decode("utf8")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8")
        rid = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="(rId\d+)"' % re.escape(sheet), wbx).group(1)
        tgt = re.search(r'<Relationship [^>]*Id="%s"[^>]*Target="([^"]+)"' % rid, rels)
        tgt = tgt or re.search(r'<Relationship [^>]*Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
        xml = z.read("xl/" + tgt.group(1).lstrip("/").replace("xl/", "", 1)).decode("utf8")
    except Exception:
        return {}
    af = re.search(r"<autoFilter.*?</autoFilter>", xml, re.S)
    if not af:
        return {}
    out = {}
    for m in re.finditer(r'<filterColumn colId="(\d+)">(.*?)</filterColumn>', af.group(0), re.S):
        vals = {v.strip() for v in re.findall(r'<filter val="([^"]*)"', m.group(2))}
        if vals:
            out[int(m.group(1))] = (vals, 'blank="1"' in m.group(2))
    return out


def passes_filter(r, flt):
    for col, (vals, blanks) in flt.items():
        v = r[col] if col < len(r) else None
        v = "" if v is None else str(v).strip()
        if v == "":
            if not blanks:
                return False
        elif v not in vals:
            return False
    return True


def read_rows(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[SHEET]
    rows, blank = [], 0
    for r in ws.iter_rows(max_col=120, values_only=True):
        if not any(v not in (None, "") for v in r):
            blank += 1
            if blank > 50:  # sheet has ~1M formatted empty rows; stop early
                break
            continue
        blank = 0
        rows.append(r)
    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    col = {}
    for key, name in HEADERS.items():
        if name not in header:
            sys.exit(f'Column "{name}" not found in {SHEET}')
        col[key] = header.index(name)
    return rows[1:], col


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    rows, c = read_rows(sys.argv[1])
    ex = {s: sum(1 for r in rows if str(r[c["status"]] or "").strip() == s) for s in REPORTED_STATUS}

    flt = read_autofilter(sys.argv[1], SHEET)
    if flt:
        scope = [r for r in rows if passes_filter(r, flt) and r[c["account"]]]
        ex["Excel filter"] = len(rows) - len(scope)
    else:
        scope = [r for r in rows
                 if str(r[c["phase"]] or "").strip() not in EXCLUDE_PHASE
                 and str(r[c["status"]] or "").strip() not in EXCLUDE_STATUS
                 and r[c["account"]]]
    ex["REVIEW"] = sum(1 for r in scope if str(r[c["phase"]] or "").strip() == "REVIEW")
    scope = [r for r in scope if str(r[c["phase"]] or "").strip() != "REVIEW"]
    ex["After filter"] = len(scope)

    def canceled(r):
        vals = [str(r[c[k]] or "").strip().upper() for k in ("clh_status", "change_type")]
        return any(v in {"CANCELED", "CANCELLED"} for v in vals)

    ex["Canceled"] = sum(1 for r in scope if canceled(r))
    scope = [r for r in scope if not canceled(r)]

    # ANZ: Phase "ANZ TBD" or country Australia / New Zealand. Excluded from the dashboard
    # counts and summarised separately in section 9.
    def is_anz(r):
        return (str(r[c["phase"]] or "").strip() == "ANZ TBD"
                or str(r[c["country"]] or "").strip() in ANZ_COUNTRIES)

    anz_rows = [r for r in scope if is_anz(r)]
    ex["ANZ"] = len(anz_rows)
    scope = [r for r in scope if not is_anz(r)]
    connect_rows = [r for r in scope if (r[c["client_type"]] or "").strip() in CONNECT_TYPES]
    ex["Client type"] = sum(1 for r in scope if (r[c["client_type"]] or "").strip() in EXCLUDE_CLIENT_TYPE)
    scope = [r for r in scope if (r[c["client_type"]] or "").strip() not in EXCLUDE_CLIENT_TYPE]
    ex["No CMX ID"] = sum(1 for r in scope if not clean_id(r[c["cmx"]]))

    # One record per CMX ID (first occurrence wins)
    dbs, seen, dup = [], set(), 0
    for r in scope:
        cmx = clean_id(r[c["cmx"]])
        if not cmx:
            continue
        if cmx in seen:
            dup += 1
            continue
        seen.add(cmx)
        dbs.append({
            "cmx": cmx,
            "ecr": clean_id(r[c["ecr"]]),
            "name": str(r[c["account"]]).strip(),
            "db": str(r[c["dbname"]] or "").strip(),
            "country": str(r[c["country"]] or "").strip() or "Unspecified",
            "ct": r[c["client_type"]] or "Unspecified",
            "ph": phase_group(r[c["phase"]]),
            "ms": r[c["status"]] or "",
            "value": round(num(r[c["value"]])),
            "l": num(r[c["learners"]]),
            "a": num(r[c["admins"]]),
            "needs": [1 if g[3](r, c) else 0 for g in GATES],
            "sso": str(r[c["sso"]] or "").strip(),
            "skills": 1 if yes(r[c["skills"]]) else 0,
        })
    ex["Duplicate CMX ID"] = dup

    def summarize(rs):
        """Reference totals for databases outside the dashboard scope (one per CMX ID)."""
        out, seen = [], set()
        for r in rs:
            k = clean_id(r[c["cmx"]])
            if not k or k in seen:
                continue
            seen.add(k)
            out.append(r)
        nd = [[1 if g[3](r, c) else 0 for g in GATES] for r in out]
        return {
            "n": len(out),
            "value": round(sum(num(r[c["value"]]) for r in out)),
            "value_recorded": sum(1 for r in out if num(r[c["value"]]) > 0),
            "l": sum(num(r[c["learners"]]) for r in out),
            "a": sum(num(r[c["admins"]]) for r in out),
            "blocked": sum(1 for x in nd if any(x)),
            "needs": [sum(x[i] for x in nd) for i in range(len(GATES))],
            "status": dict(Counter(str(r[c["status"]] or "(blank)").strip() for r in out).most_common()),
            "ct": dict(Counter(str(r[c["client_type"]] or "Unspecified").strip() for r in out).most_common()),
        }

    outside = {
        "connect": summarize(connect_rows),
        "anz": summarize(anz_rows),
        "directcourse": summarize([r for r in rows if str(r[c["phase"]] or "").strip() == "DirectCourse" and r[c["account"]]]),
    }

    data = {
        "F": [g[1] for g in GATES], "P": [g[2] for g in GATES], "K": [g[0] for g in GATES],
        "C": dbs, "rows": len(rows), "excl": ex, "filter": bool(flt), "outside": outside,
        "source": Path(sys.argv[1]).name,
        "generated": datetime.date.today().strftime("%B %d, %Y").replace(" 0", " "),
    }
    block = ("/* Generated by scripts/build_data.py. Do not edit by hand. */\n"
             "window.DASHBOARD_DATA = " + json.dumps(data, separators=(",", ":")).replace("</", "<\\/") + ";\n")
    html = OUT.read_text(encoding="utf-8")
    new, n = re.subn(r"<script>/\* Generated by scripts/build_data\.py.*?</script>",
                     lambda m: "<script>" + block + "</script>", html, count=1, flags=re.S)
    if not n:
        sys.exit("Data block not found in index.html")
    OUT.write_text(new, encoding="utf-8")
    print(f"Updated {OUT} ({len(dbs)} databases in scope of {len(rows)} rows); excluded: {ex}")


if __name__ == "__main__":
    main()
