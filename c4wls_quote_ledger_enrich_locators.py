#!/usr/bin/env python3
import argparse, csv, json, os, time
from pathlib import Path
from email import policy
from email.parser import BytesParser

import openpyxl

def now_ts():
    return time.strftime("%Y-%m-%d %H:%M:%S")

def parse_eml_meta(path: Path):
    try:
        msg = BytesParser(policy=policy.default).parsebytes(path.read_bytes())
        dt = (msg.get("date") or "").strip()
        frm = (msg.get("from") or "").strip()
        subj = (msg.get("subject") or "").strip()
        # keep it short but traceable
        locator = f"email date={dt} | from={frm} | subject={subj}"
        return locator[:4000]
    except Exception as e:
        return f"email parse failed: {e}"

def pdf_page_count(path: Path):
    # Use a very light approach: count '/Type /Page' tokens (fast, not perfect but good enough for locator)
    try:
        b = path.read_bytes()
        # crude count; avoids external deps
        n = b.count(b"/Type /Page")
        if n == 0:
            # some PDFs use /Pages; still return unknown
            return "pages=unknown"
        return f"pages={n}"
    except Exception as e:
        return f"pages=unknown ({e})"

def excel_sheet_list(path: Path):
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=False, keep_vba=True)
        return "sheets=" + ",".join(wb.sheetnames[:25])
    except Exception as e:
        return f"sheets=unknown ({e})"

def main():
    ap = argparse.ArgumentParser(description="READ-ONLY: enrich quote ledger inventory with locators (email meta, pdf pages, excel sheets).")
    ap.add_argument("--in-csv", required=True)
    ap.add_argument("--out-dir", required=True)
    args = ap.parse_args()

    in_csv = Path(os.path.expanduser(args.in_csv)).resolve()
    out_dir = Path(os.path.expanduser(args.out_dir)).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not in_csv.exists():
        raise SystemExit(f"FATAL: input csv not found: {in_csv}")

    rows = []
    with open(in_csv, "r", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            rows.append(row)

    enriched = 0
    missing = 0

    for row in rows:
        p = Path(row["source_path"])
        if not p.exists():
            row["locator"] = "MISSING FILE"
            row["status"] = "missing_file"
            missing += 1
            continue

        ext = (row.get("file_ext") or "").lower()
        if ext == ".eml":
            row["locator"] = parse_eml_meta(p)
            row["status"] = "indexed_email"
            enriched += 1
        elif ext == ".pdf":
            row["locator"] = pdf_page_count(p)
            row["status"] = "indexed_pdf"
            enriched += 1
        elif ext in (".xls", ".xlsx", ".xlsm"):
            row["locator"] = excel_sheet_list(p)
            row["status"] = "indexed_excel"
            enriched += 1
        else:
            # keep as-is
            if not row.get("locator") or row["locator"] == "TBD":
                row["locator"] = "unindexed"
            row["status"] = row.get("status") or "unindexed"

    ts = time.strftime("%Y%m%d_%H%M%S")
    out_csv = out_dir / f"{in_csv.stem}__enriched_{ts}.csv"
    out_json = out_dir / f"{in_csv.stem}__enriched_{ts}.json"

    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()) if rows else [])
        if rows:
            w.writeheader()
            w.writerows(rows)

    with open(out_json, "w", encoding="utf-8") as f:
        json.dump({
            "created_ts": now_ts(),
            "input_csv": str(in_csv),
            "count": len(rows),
            "enriched": enriched,
            "missing_files": missing
        }, f, indent=2)

    print("OK")
    print(f"IN:   {in_csv}")
    print(f"OUT:  {out_csv}")
    print(f"META: {out_json}")
    print(f"Enriched: {enriched}  Missing files: {missing}  Total: {len(rows)}")

if __name__ == "__main__":
    main()
