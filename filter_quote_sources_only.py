#!/usr/bin/env python3
import argparse, csv, os, time
from pathlib import Path

QUOTE_EXT = {".pdf",".eml",".xls",".xlsx",".xlsm"}
FOLDER_TOKENS = ["sub bids", "subbids", "sub email", "bids"]

def is_quote_source(path_str: str, ext: str) -> bool:
    s = path_str.lower()
    if ext.lower() not in QUOTE_EXT:
        return False
    return any(tok in s for tok in FOLDER_TOKENS)

def main():
    ap = argparse.ArgumentParser(description="READ-ONLY: filter ledger to quote-source-only artifacts.")
    ap.add_argument("--in-csv", required=True)
    ap.add_argument("--out-dir", required=True)
    args = ap.parse_args()

    in_csv = Path(os.path.expanduser(args.in_csv)).resolve()
    out_dir = Path(os.path.expanduser(args.out_dir)).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not in_csv.exists():
        raise SystemExit(f"FATAL: not found: {in_csv}")

    rows = []
    with open(in_csv, "r", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            rows.append(row)

    keep = []
    drop = []

    for row in rows:
        p = row.get("source_path","")
        ext = (row.get("file_ext") or "").lower()
        if is_quote_source(p, ext):
            keep.append(row)
        else:
            drop.append(row)

    ts = time.strftime("%Y%m%d_%H%M%S")
    keep_csv = out_dir / f"{in_csv.stem}__QUOTE_ONLY_{ts}.csv"
    drop_csv = out_dir / f"{in_csv.stem}__EXCLUDED_{ts}.csv"

    if keep:
        with open(keep_csv, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=list(keep[0].keys()))
            w.writeheader()
            w.writerows(keep)
    else:
        with open(keep_csv, "w", encoding="utf-8") as f:
            f.write("")

    if drop:
        with open(drop_csv, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=list(drop[0].keys()))
            w.writeheader()
            w.writerows(drop)

    print("OK")
    print(f"IN: {in_csv}")
    print(f"QUOTE_ONLY: {keep_csv}  count={len(keep)}")
    print(f"EXCLUDED:  {drop_csv}  count={len(drop)}")

if __name__ == "__main__":
    main()
