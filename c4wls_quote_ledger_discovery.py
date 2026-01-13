#!/usr/bin/env python3
import argparse, os, re, csv, json, time
from pathlib import Path

ALLOWED_EXT = {".pdf",".eml",".msg",".xlsx",".xls",".xlsm",".docx",".doc",".txt",".rtf",".csv",".png",".jpg",".jpeg",".heic"}

def now_ts():
    return time.strftime("%Y-%m-%d %H:%M:%S")

def infer_option(path_str: str) -> str:
    s = path_str.lower()
    # strong signals
    if "opt1" in s or "option 1" in s or "option_1" in s or "option-1" in s:
        return "1"
    if "opt2" in s or "option 2" in s or "option_2" in s or "option-2" in s:
        return "2"
    return "unknown"

def normalize_vendor_from_filename(name: str) -> str:
    # strip extension and common junk
    base = re.sub(r"\.[^.]+$", "", name).strip()
    base = re.sub(r"[_\-]+", " ", base).strip()
    base = re.sub(r"\s+\(\d+\)$", "", base).strip()  # "file (1)"
    base = re.sub(r"\bproposal\b|\bquote\b|\bbid\b|\bestimate\b", "", base, flags=re.I).strip()
    base = re.sub(r"\s{2,}", " ", base).strip()
    return base if base else "unknown"

def walk_files(root: Path):
    for p in root.rglob("*"):
        if p.is_file():
            ext = p.suffix.lower()
            if ext in ALLOWED_EXT:
                yield p

def main():
    ap = argparse.ArgumentParser(description="READ-ONLY: build quote ledger inventory from a project folder")
    ap.add_argument("--project", required=True)
    ap.add_argument("--root", required=True, help="Project folder root in Dropbox (absolute path)")
    ap.add_argument("--out-dir", required=True)
    args = ap.parse_args()

    project = args.project
    root = Path(os.path.expanduser(args.root)).resolve()
    out_dir = Path(os.path.expanduser(args.out_dir)).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not root.exists():
        raise SystemExit(f"FATAL: root not found: {root}")

    rows = []
    for p in walk_files(root):
        rel = str(p.relative_to(root))
        opt = infer_option(str(p))
        vendor = normalize_vendor_from_filename(p.name)

        rows.append({
            "project": project,
            "option": opt,
            "vendor_raw": vendor,
            "vendor_normalized": vendor,   # later we’ll map to master vendor list
            "trade": "unknown",
            "source_type": "bid_pdf" if p.suffix.lower()==".pdf" else ("email_quote" if p.suffix.lower()==".eml" else "other"),
            "source_path": str(p),
            "locator": "TBD",              # later: page/section for PDF or email date/subject for EML
            "file_ext": p.suffix.lower(),
            "file_size_bytes": p.stat().st_size,
            "status": "unparsed",
            "notes": rel
        })

    # sort for readability
    rows.sort(key=lambda r: (r["option"], r["source_type"], r["vendor_normalized"], r["source_path"]))

    ts = time.strftime("%Y%m%d_%H%M%S")
    csv_path = out_dir / f"C4WLS_quote_ledger_inventory_{ts}.csv"
    json_path = out_dir / f"C4WLS_quote_ledger_inventory_{ts}.json"

    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()) if rows else [])
        if rows:
            w.writeheader()
            w.writerows(rows)

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump({"created_ts": now_ts(), "root": str(root), "count": len(rows), "rows": rows}, f, indent=2)

    print("OK")
    print(f"ROOT: {root}")
    print(f"COUNT: {len(rows)}")
    print(f"CSV:  {csv_path}")
    print(f"JSON: {json_path}")

if __name__ == "__main__":
    main()
