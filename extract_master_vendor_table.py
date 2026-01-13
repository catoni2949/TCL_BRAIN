#!/usr/bin/env python3
import argparse, os, re, csv, time
from pathlib import Path

import pandas as pd

def is_section_header(s: str) -> bool:
    s = (s or "").strip()
    if not s:
        return False
    # treat as header if mostly uppercase letters/spaces/&/- and length reasonable
    letters = re.sub(r"[^A-Za-z]", "", s)
    if not letters:
        return False
    if s.upper() != s:
        return False
    # ignore obvious top-of-sheet junk
    if "TENANT IMPROVEMENT VENDOR LIST" in s.upper():
        return True
    if "PROJECT NAME" in s.upper():
        return True
    return True

def is_junk_line(s: str) -> bool:
    s = (s or "").strip()
    if not s:
        return True
    u = s.upper()
    if u in ("TENANT IMPROVEMENT VENDOR LIST", "PROJECT NAME:", "PROJECT NAME"):
        return True
    return False

def clean_vendor(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"\s{2,}", " ", s)
    return s

def main():
    ap = argparse.ArgumentParser(description="Extract clean vendor master table (company + trade section + contact) from TCL Bid List .xls")
    ap.add_argument("--master-xls", required=True)
    ap.add_argument("--sheet", default="Sheet2")
    ap.add_argument("--out-dir", required=True)
    args = ap.parse_args()

    master = Path(os.path.expanduser(args.master_xls)).resolve()
    out_dir = Path(os.path.expanduser(args.out_dir)).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not master.exists():
        raise SystemExit(f"FATAL: not found: {master}")

    df = pd.read_excel(master, sheet_name=args.sheet, header=None)
    # Expect columns 0..3 to be: company/section, contact, phone, email
    current_section = "UNSPECIFIED"
    out_rows = []

    for _, row in df.iterrows():
        col0 = "" if pd.isna(row.get(0)) else str(row.get(0)).strip()
        col1 = "" if pd.isna(row.get(1)) else str(row.get(1)).strip()
        col2 = "" if pd.isna(row.get(2)) else str(row.get(2)).strip()
        col3 = "" if pd.isna(row.get(3)) else str(row.get(3)).strip()

        if not col0 and not col1 and not col2 and not col3:
            continue

        if col0 and is_section_header(col0) and not col1 and not col2 and not col3:
            # pure section row
            sec = col0.strip()
            current_section = sec
            continue

        # Some sheets put header + data in same row; treat ALL CAPS in col0 as section if col1 is blank-ish
        if col0 and is_section_header(col0) and (not col1 or col1.lower() in ("nan","none")) and (not col3):
            current_section = col0.strip()
            continue

        # ignore junk
        if is_junk_line(col0):
            continue

        # treat as vendor row when col0 is present and not ALL CAPS section header
        vendor = clean_vendor(col0)
        if not vendor:
            continue

        # skip if vendor looks like a section header (ALL CAPS) AND there is no contact/email
        if vendor.upper() == vendor and not col3 and not col1:
            # likely another section header line
            current_section = vendor
            continue

        out_rows.append({
            "trade_section": current_section,
            "vendor_company": vendor,
            "contact": col1,
            "phone": col2,
            "email": col3
        })

    ts = time.strftime("%Y%m%d_%H%M%S")
    out_csv = out_dir / f"TCL_Bid_List_Master_Vendors_{ts}.csv"

    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["trade_section","vendor_company","contact","phone","email"])
        w.writeheader()
        w.writerows(out_rows)

    print("OK")
    print(f"SHEET: {args.sheet}")
    print(f"VENDORS: {len(out_rows)}")
    print(f"OUT: {out_csv}")

if __name__ == "__main__":
    main()
