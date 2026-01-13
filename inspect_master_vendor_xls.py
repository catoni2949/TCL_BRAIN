#!/usr/bin/env python3
import argparse, os
from pathlib import Path
import pandas as pd

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--master-xls", required=True)
    ap.add_argument("--max-rows", type=int, default=8)
    args = ap.parse_args()

    master = Path(os.path.expanduser(args.master_xls)).resolve()
    if not master.exists():
        raise SystemExit(f"FATAL: not found: {master}")

    xls = pd.ExcelFile(master)
    print("OK")
    print(f"MASTER: {master}")
    print("SHEETS:")
    for sh in xls.sheet_names:
        print(f"- {sh}")

    print("\nDETAIL (top candidate columns per sheet):")
    for sh in xls.sheet_names:
        df = xls.parse(sh)
        if df is None or df.empty:
            print(f"\n[{sh}] EMPTY")
            continue

        # show columns + non-null counts
        cols = list(df.columns)
        nn = {c: int(df[c].notna().sum()) for c in cols}

        # pick columns likely to contain vendor names
        def score_col(c):
            cl = str(c).lower()
            score = 0
            if "vendor" in cl: score += 5
            if "company" in cl: score += 5
            if "sub" in cl: score += 4
            if "name" in cl: score += 3
            if "email" in cl or "phone" in cl: score += 1
            # prefer columns with lots of values
            score += min(nn.get(c,0), 200) / 200.0
            return score

        ranked = sorted(cols, key=score_col, reverse=True)
        top = ranked[:6]

        print(f"\n[{sh}] rows={len(df)} cols={len(cols)}")
        for c in top:
            print(f"  - col='{c}' nonnull={nn.get(c,0)}")
            # print sample values
            vals = [str(v).strip() for v in df[c].dropna().astype(str).head(args.max_rows).tolist()]
            vals = [v for v in vals if v and v.lower() not in ("nan","none")]
            if vals:
                for v in vals[:args.max_rows]:
                    print(f"      sample: {v}")

if __name__ == "__main__":
    main()
