#!/usr/bin/env python3
import argparse, csv, time
from pathlib import Path

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in-csv", required=True)
    ap.add_argument("--out-dir", required=True)
    args = ap.parse_args()

    in_csv = Path(args.in_csv)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    with open(in_csv, newline="", encoding="utf-8") as f:
        r = csv.DictReader(f)
        rows = list(r)

    ts = time.strftime("%Y%m%d_%H%M%S")
    base = in_csv.stem[:60]
    short_hash = hex(abs(hash(in_csv.stem)))[2:10]

    out_quote   = out_dir / f"{base}__QUOTE_APPROVED_{ts}_{short_hash}.csv"
    out_non     = out_dir / f"{base}__NON_QUOTE_{ts}_{short_hash}.csv"
    out_unknown = out_dir / f"{base}__UNKNOWN_{ts}_{short_hash}.csv"
    out_block   = out_dir / f"{base}__PIPELINE_BLOCK_{ts}_{short_hash}.txt"

    quote, nonq, unk, review = [], [], [], []
    for row in rows:
        role = (row.get("doc_role") or "").strip().upper()
        if role == "QUOTE":
            quote.append(row)
        elif role == "NON_QUOTE":
            nonq.append(row)
        elif role == "UNKNOWN" or role == "":
            unk.append(row)
        elif role == "REVIEW":
            review.append(row)
        else:
            # anything else is effectively unknown
            unk.append(row)

    def write_csv(path: Path, items):
        with open(path, "w", newline="", encoding="utf-8") as f:
            if not items:
                f.write("")
                return
            w = csv.DictWriter(f, fieldnames=list(items[0].keys()))
            w.writeheader()
            w.writerows(items)

    write_csv(out_quote, quote)
    write_csv(out_non, nonq)
    write_csv(out_unknown, unk)

    blocked = False
    reasons = []
    if len(review) > 0:
        blocked = True
        reasons.append(f"BLOCK: REVIEW rows exist ({len(review)}). Must be resolved before proceeding.")
    if len(quote) == 0:
        reasons.append("NOTE: zero QUOTE rows found. No priced quotes available from this input set.")
    if len(unk) > 0:
        reasons.append(f"NOTE: UNKNOWN rows exist ({len(unk)}). These are not eligible for extraction.")

    if blocked:
        with open(out_block, "w", encoding="utf-8") as f:
            f.write("PIPELINE STATUS: BLOCKED\n")
            for line in reasons:
                f.write(line + "\n")
        print("NO-GO")
        print(f"BLOCK_FILE: {out_block}")
    else:
        print("GO")

    print(f"IN: {in_csv}")
    print(f"QUOTE_APPROVED: {out_quote}  count={len(quote)}")
    print(f"NON_QUOTE:      {out_non}  count={len(nonq)}")
    print(f"UNKNOWN:        {out_unknown}  count={len(unk)}")
    print(f"REVIEW:         {len(review)}")

if __name__ == "__main__":
    main()
