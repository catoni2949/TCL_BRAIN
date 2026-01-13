#!/usr/bin/env python3
import argparse, csv, os, re, time
from pathlib import Path
from email import policy
from email.parser import BytesParser

def extract_email_addr(s: str) -> str:
    if not s:
        return ""
    s = s.strip()
    m = re.search(r"([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})", s, flags=re.I)
    return (m.group(1) if m else "").lower().strip()

def parse_eml_from_email(path: Path) -> str:
    try:
        msg = BytesParser(policy=policy.default).parsebytes(path.read_bytes())
        frm = (msg.get("from") or "").strip()
        return extract_email_addr(frm)
    except Exception:
        return ""

def main():
    ap = argparse.ArgumentParser(description="READ-ONLY: match .eml ledger rows to master vendor list by sender email (exact).")
    ap.add_argument("--ledger-csv", required=True)
    ap.add_argument("--master-csv", required=True)
    ap.add_argument("--out-dir", required=True)
    args = ap.parse_args()

    ledger_csv = Path(os.path.expanduser(args.ledger_csv)).resolve()
    master_csv = Path(os.path.expanduser(args.master_csv)).resolve()
    out_dir = Path(os.path.expanduser(args.out_dir)).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not ledger_csv.exists():
        raise SystemExit(f"FATAL: ledger not found: {ledger_csv}")
    if not master_csv.exists():
        raise SystemExit(f"FATAL: master not found: {master_csv}")

    # master email index
    email_to_master = {}
    with open(master_csv, "r", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            em = extract_email_addr(row.get("email",""))
            if not em:
                continue
            # first one wins (stable)
            if em not in email_to_master:
                email_to_master[em] = {
                    "vendor_company": (row.get("vendor_company") or "").strip(),
                    "trade_section": (row.get("trade_section") or "").strip(),
                }

    rows = []
    with open(ledger_csv, "r", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            rows.append(row)

    matched = 0
    eml_rows = 0

    for row in rows:
        ext = (row.get("file_ext") or "").lower()
        if ext != ".eml":
            continue

        eml_rows += 1
        p = Path(row.get("source_path",""))
        if not p.exists():
            continue

        from_email = parse_eml_from_email(p)
        row["eml_from_email"] = from_email

        if from_email and from_email in email_to_master:
            m = email_to_master[from_email]
            row["vendor_master_match"] = m["vendor_company"]
            row["master_trade_section"] = m["trade_section"]
            row["match_type"] = "email_exact"
            row["match_score"] = "100"
            row["needs_review"] = "no"
            matched += 1
        else:
            # keep existing match fields; just mark email miss for clarity
            if not row.get("match_type"):
                row["match_type"] = "none"

    ts = time.strftime("%Y%m%d_%H%M%S")
    out_csv = out_dir / f"{ledger_csv.stem}__email_matched_{ts}.csv"

    # ensure new column exists in header
    if rows and "eml_from_email" not in rows[0].keys():
        # add empty for non-eml rows
        for r in rows:
            r.setdefault("eml_from_email", "")

    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()) if rows else [])
        if rows:
            w.writeheader()
            w.writerows(rows)

    print("OK")
    print(f"OUT: {out_csv}")
    print(f"Summary: eml_rows={eml_rows} email_exact_matched={matched} total={len(rows)}")

if __name__ == "__main__":
    main()
