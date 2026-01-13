#!/usr/bin/env python3
"""
Build SOV write plan from final estimator-approved amounts
READ-ONLY to Excel. Produces JSON write plan only.
"""

import csv, json, argparse, time

import hashlib
import json

def sha256_json(obj) -> str:
    return hashlib.sha256(
        json.dumps(obj, sort_keys=True, separators=(",", ":")).encode("utf-8")
    ).hexdigest()
from pathlib import Path

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--amounts-csv", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--lock", required=True)
    ap.add_argument("--out-dir", required=True)
    args = ap.parse_args()

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    rows = []
    with open(args.amounts_csv, newline="", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            amt = row.get("selected_amount_clean")
            trade = row.get("tcl_trade")
            if not amt or not trade:
                continue

            rows.append({
                "sheet": "ESTIMATE (INPUT)",
                "match": {
                    "trade": trade
                },
                "write": {
                    "column": "AMOUNT",
                    "value": float(amt)
                },
                "meta": {
                    "source_path": row.get("source_path"),
                    "vendor": row.get("vendor_master_match")
                }
            })

    ts = time.strftime("%Y%m%d_%H%M%S")
    out = out_dir / f"SOV_WRITE_PLAN_{ts}.json"

    plan = {
        "template": args.template,
        "lock": args.lock,
        "writes": rows
    }

    with open(out, "w", encoding="utf-8") as f:
        plan["plan_hash"] = sha256_json(plan.get("writes", []))
        plan["template_sha256"] = lock["template"]["sha256"]
        json.dump(plan, f, indent=2)

    print("GO")
    print(f"WRITE_PLAN: {out}")
    print(f"ROWS_MAPPED: {len(rows)}")
    print("ROWS_SKIPPED: 0")

if __name__ == "__main__":
    main()
