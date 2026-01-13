#!/usr/bin/env python3
import argparse, csv, os, time
from pathlib import Path

def load_map(map_csv: Path):
    m = {}
    with open(map_csv, "r", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            k = (row.get("master_trade_section") or "").strip()
            v = (row.get("tcl_trade_suggested") or "").strip()
            if k and v:
                m[k] = v
    return m

def main():
    ap = argparse.ArgumentParser(description="READ-ONLY: apply master-section->TCL trade mapping and produce estimator review sheet.")
    ap.add_argument("--quote-ledger-csv", required=True)
    ap.add_argument("--trade-map-csv", required=True)
    ap.add_argument("--out-dir", required=True)
    args = ap.parse_args()

    qcsv = Path(os.path.expanduser(args.quote_ledger_csv)).resolve()
    mcsv = Path(os.path.expanduser(args.trade_map_csv)).resolve()
    out_dir = Path(os.path.expanduser(args.out_dir)).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not qcsv.exists():
        raise SystemExit(f"FATAL: quote ledger not found: {qcsv}")
    if not mcsv.exists():
        raise SystemExit(f"FATAL: trade map not found: {mcsv}")

    trade_map = load_map(mcsv)

    rows = []
    with open(qcsv, "r", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            rows.append(row)

    # Apply mapping (only if master_trade_section exists AND map exists)
    mapped = 0
    for row in rows:
        sec = (row.get("master_trade_section") or "").strip()
        tcl_trade = trade_map.get(sec, "")
        row["tcl_trade"] = tcl_trade
        if tcl_trade and (row.get("needs_review") or "").strip().lower() != "yes":
            mapped += 1

    # Review sheet: needs_review=yes OR tcl_trade blank
    review = []
    for row in rows:
        needs = (row.get("needs_review") or "").strip().lower() == "yes"
        trade_blank = not (row.get("tcl_trade") or "").strip()
        if needs or trade_blank:
            review.append({
                "source_path": row.get("source_path",""),
                "file_ext": row.get("file_ext",""),
                "vendor_master_match": row.get("vendor_master_match",""),
                "master_trade_section": row.get("master_trade_section",""),
                "match_type": row.get("match_type",""),
                "match_score": row.get("match_score",""),
                "tcl_trade": row.get("tcl_trade",""),
                "needs_review": row.get("needs_review","yes"),
                "review_action": "Confirm vendor + trade OR mark as non-quote / wrong folder",
                "notes": ""
            })

    ts = time.strftime("%Y%m%d_%H%M%S")
    out_ledger = out_dir / f"{qcsv.stem}__TRADE_ASSIGNED_{ts}.csv"
    out_review = out_dir / f"{qcsv.stem}__ESTIMATOR_REVIEW_{ts}.csv"

    # write updated ledger
    with open(out_ledger, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()) if rows else [])
        if rows:
            w.writeheader()
            w.writerows(rows)

    # write review sheet
    if review:
        with open(out_review, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=list(review[0].keys()))
            w.writeheader()
            w.writerows(review)
    else:
        with open(out_review, "w", encoding="utf-8") as f:
            f.write("")

    print("OK")
    print(f"OUT_LEDGER: {out_ledger}")
    print(f"OUT_REVIEW: {out_review}")
    print(f"Summary: quote_rows={len(rows)} review_rows={len(review)} trade_mapped_nonreview={mapped}")

if __name__ == "__main__":
    main()
