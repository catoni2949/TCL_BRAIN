#!/usr/bin/env python3
import argparse, csv, os, re, time
from pathlib import Path
from difflib import SequenceMatcher

import pandas as pd

def now_ts():
    return time.strftime("%Y-%m-%d %H:%M:%S")

def norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s).lower().strip()
    s = s.replace("&", " and ")
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\b(inc|llc|ltd|co|company|corporation|corp|services|service|group|construction|contracting)\b", " ", s)
    s = re.sub(r"\s{2,}", " ", s).strip()
    return s

def best_fuzzy_match(name_norm: str, master_norm_list):
    best = ("", 0.0)
    for m in master_norm_list:
        if not m:
            continue
        score = SequenceMatcher(None, name_norm, m).ratio()
        if score > best[1]:
            best = (m, score)
    return best[0], int(round(best[1] * 100))

def load_master_vendors(master_xls: Path):
    # Reads all sheets; pulls any column that looks like vendor/company
    xls = pd.ExcelFile(master_xls)
    candidates = []
    for sh in xls.sheet_names:
        df = xls.parse(sh)
        if df is None or df.empty:
            continue
        cols = [c for c in df.columns]
        # pick likely columns
        likely = []
        for c in cols:
            cl = str(c).lower()
            if any(k in cl for k in ["vendor", "company", "sub", "subcontractor", "name"]):
                likely.append(c)
        if not likely:
            # fallback first column
            likely = [cols[0]]
        for c in likely[:2]:
            for v in df[c].dropna().astype(str).tolist():
                v2 = v.strip()
                if v2 and v2.lower() not in ("nan", "none"):
                    candidates.append(v2)
    # de-dupe preserving order
    seen = set()
    out = []
    for v in candidates:
        if v not in seen:
            seen.add(v)
            out.append(v)
    return out

def main():
    ap = argparse.ArgumentParser(description="Read-only: normalize ledger vendor names against master vendor list.")
    ap.add_argument("--ledger-csv", required=True)
    ap.add_argument("--master-xls", required=True)
    ap.add_argument("--out-dir", required=True)
    args = ap.parse_args()

    ledger_csv = Path(os.path.expanduser(args.ledger_csv)).resolve()
    master_xls = Path(os.path.expanduser(args.master_xls)).resolve()
    out_dir = Path(os.path.expanduser(args.out_dir)).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not ledger_csv.exists():
        raise SystemExit(f"FATAL: ledger csv not found: {ledger_csv}")
    if not master_xls.exists():
        raise SystemExit(f"FATAL: master xls not found: {master_xls}")

    master_raw = load_master_vendors(master_xls)
    master_norm = [norm(v) for v in master_raw]
    norm_to_raw = {}
    for raw, n in zip(master_raw, master_norm):
        if n and n not in norm_to_raw:
            norm_to_raw[n] = raw

    rows = []
    with open(ledger_csv, "r", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            rows.append(row)

    for row in rows:
        v = row.get("vendor_normalized") or row.get("vendor_raw") or ""
        vn = norm(v)

        # exact match on normalized
        if vn in norm_to_raw and vn:
            row["vendor_master_match"] = norm_to_raw[vn]
            row["match_type"] = "exact"
            row["match_score"] = "100"
            row["needs_review"] = "no"
            continue

        # fuzzy match
        if vn:
            best_norm, score = best_fuzzy_match(vn, master_norm)
            if score >= 92 and best_norm:
                row["vendor_master_match"] = norm_to_raw.get(best_norm, "")
                row["match_type"] = "fuzzy"
                row["match_score"] = str(score)
                row["needs_review"] = "no"
            elif score >= 80 and best_norm:
                row["vendor_master_match"] = norm_to_raw.get(best_norm, "")
                row["match_type"] = "fuzzy"
                row["match_score"] = str(score)
                row["needs_review"] = "yes"
            else:
                row["vendor_master_match"] = ""
                row["match_type"] = "none"
                row["match_score"] = str(score if vn else 0)
                row["needs_review"] = "yes"
        else:
            row["vendor_master_match"] = ""
            row["match_type"] = "none"
            row["match_score"] = "0"
            row["needs_review"] = "yes"

    ts = time.strftime("%Y%m%d_%H%M%S")
    out_csv = out_dir / f"{ledger_csv.stem}__vendor_normalized_{ts}.csv"

    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()) if rows else [])
        if rows:
            w.writeheader()
            w.writerows(rows)

    # quick summary
    exact = sum(1 for r in rows if r.get("match_type") == "exact")
    fuzzy_ok = sum(1 for r in rows if r.get("match_type") == "fuzzy" and r.get("needs_review") == "no")
    review = sum(1 for r in rows if r.get("needs_review") == "yes")
    print("OK")
    print(f"OUT: {out_csv}")
    print(f"Summary: exact={exact} fuzzy_ok={fuzzy_ok} needs_review={review} total={len(rows)}")

if __name__ == "__main__":
    main()
