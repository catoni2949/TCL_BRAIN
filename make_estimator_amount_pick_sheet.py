#!/usr/bin/env python3
import argparse, csv, json
from pathlib import Path

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in-csv", required=True)
    ap.add_argument("--out-dir", required=True)
    args = ap.parse_args()

    in_csv = Path(args.in_csv)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    rows_out = []

    with open(in_csv, newline="", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            cjson = (row.get("amount_candidates_json") or "").strip()
            cands = []
            if cjson:
                try:
                    cands = json.loads(cjson)
                except Exception:
                    cands = []

            top = cands[:5]
            def fmt(c):
                lab = (c.get("label") or "").strip()
                amt = (c.get("amount_clean") or "").strip()
                ctx = (c.get("context") or "").strip()
                return f"{lab}|{amt}|{ctx}" if lab else f"{amt}|{ctx}"

            rows_out.append({
                "source_path": row.get("source_path",""),
                "vendor_master_match": row.get("vendor_master_match",""),
                "vendor_proved_in_text": row.get("vendor_proved_in_text",""),
                "tcl_trade": row.get("tcl_trade",""),
                "option_applicability": row.get("option_applicability",""),
                "needs_review": row.get("needs_review",""),
                "review_reason": row.get("review_reason",""),
                "candidate_1": fmt(top[0]) if len(top) > 0 else "",
                "candidate_2": fmt(top[1]) if len(top) > 1 else "",
                "candidate_3": fmt(top[2]) if len(top) > 2 else "",
                "candidate_4": fmt(top[3]) if len(top) > 3 else "",
                "candidate_5": fmt(top[4]) if len(top) > 4 else "",
                "estimator_selected_amount_clean": "",
                "estimator_notes": ""
            })

    out_csv = out_dir / (in_csv.stem[:60] + "__ESTIMATOR_AMOUNT_PICK.csv")
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows_out[0].keys()) if rows_out else [])
        if rows_out:
            w.writeheader()
            w.writerows(rows_out)

    print("OK")
    print(f"IN:  {in_csv}")
    print(f"OUT: {out_csv}")
    print(f"ROWS: {len(rows_out)}")

if __name__ == "__main__":
    main()
