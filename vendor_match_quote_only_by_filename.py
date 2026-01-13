#!/usr/bin/env python3
import argparse, csv, os, re, time
from pathlib import Path

def norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s).lower().strip()
    s = s.replace("&", " and ")
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\b(inc|llc|ltd|co|company|corporation|corp|services|service|group)\b", " ", s)
    s = re.sub(r"\s{2,}", " ", s).strip()
    return s

def tokenize(s: str):
    s = norm(s)
    toks = [t for t in s.split(" ") if len(t) >= 3]
    return toks

def load_master(master_csv: Path):
    master = []
    with open(master_csv, "r", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            vc = (row.get("vendor_company") or "").strip()
            if not vc:
                continue
            master.append({
                "vendor_company": vc,
                "trade_section": (row.get("trade_section") or "").strip(),
                "vendor_norm": norm(vc),
                "tokens": tokenize(vc),
            })
    # sort longest names first to avoid short-name collisions
    master.sort(key=lambda x: len(x["vendor_norm"]), reverse=True)
    return master

def filename_contains_vendor(filename_norm: str, vendor_norm: str, tokens) -> bool:
    # require either full normalized vendor substring OR at least two strong tokens present
    if vendor_norm and vendor_norm in filename_norm and len(vendor_norm) >= 6:
        return True
    hits = 0
    for t in tokens:
        if t in filename_norm:
            hits += 1
        if hits >= 2:
            return True
    return False

def main():
    ap = argparse.ArgumentParser(description="READ-ONLY: match QUOTE_ONLY ledger rows to master vendor list using filename containment (Sub Bids only).")
    ap.add_argument("--quote-ledger-csv", required=True)
    ap.add_argument("--master-csv", required=True)
    ap.add_argument("--out-dir", required=True)
    args = ap.parse_args()

    qcsv = Path(os.path.expanduser(args.quote_ledger_csv)).resolve()
    master_csv = Path(os.path.expanduser(args.master_csv)).resolve()
    out_dir = Path(os.path.expanduser(args.out_dir)).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not qcsv.exists():
        raise SystemExit(f"FATAL: quote ledger not found: {qcsv}")
    if not master_csv.exists():
        raise SystemExit(f"FATAL: master not found: {master_csv}")

    master = load_master(master_csv)

    rows = []
    with open(qcsv, "r", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            rows.append(row)

    matched = 0
    scanned = 0

    for row in rows:
        # preserve prior deterministic matches (email_exact, etc.)
        if (row.get("match_type") or "").strip() in ("email_exact",):
            continue

        spath = row.get("source_path","")
        ext = (row.get("file_ext") or "").lower()
        if ext not in (".pdf",".xls",".xlsx",".xlsm",".eml"):
            continue

        # only do filename match when clearly in Sub Bids folder
        if "sub bids" not in spath.lower() and "subbids" not in spath.lower():
            continue

        scanned += 1
        fname = Path(spath).name
        fname_norm = norm(fname)

        best = None
        for m in master:
            if filename_contains_vendor(fname_norm, m["vendor_norm"], m["tokens"]):
                best = m
                break

        if best:
            row["vendor_master_match"] = best["vendor_company"]
            row["master_trade_section"] = best["trade_section"]
            row["match_type"] = "filename_contains"
            row["match_score"] = "95"
            row["needs_review"] = "no"
            matched += 1
        else:
            # leave as review
            row["needs_review"] = "yes"

    ts = time.strftime("%Y%m%d_%H%M%S")
    out_csv = out_dir / f"{qcsv.stem}__vendor_filename_matched_{ts}.csv"

    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()) if rows else [])
        if rows:
            w.writeheader()
            w.writerows(rows)

    print("OK")
    print(f"OUT: {out_csv}")
    print(f"Summary: scanned_subbids={scanned} filename_matched={matched} total_quote_rows={len(rows)}")

if __name__ == "__main__":
    main()
