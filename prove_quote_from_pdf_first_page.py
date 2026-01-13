#!/usr/bin/env python3
import argparse, csv, os, re, time, subprocess
from pathlib import Path

QUOTE_KEYWORDS = [
    "proposal", "quote", "estimated", "estimate", "bid", "quotation", "scope of work",
    "total", "subtotal", "tax", "pricing"
]

DRAWING_KEYWORDS = [
    "floor plan", "reflected ceiling plan", "sheet", "drawing", "scale", "north", "key plan",
    "existing", "demolition plan", "finish plan", "legend"
]

def pdftotext_first_page(pdf_path: Path) -> str:
    # Uses poppler pdftotext if available. Falls back to empty.
    try:
        # -f 1 -l 1 = first page only, output to stdout
        r = subprocess.run(
            ["pdftotext", "-f", "1", "-l", "1", str(pdf_path), "-"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
            text=True
        )
        if r.returncode != 0:
            return ""
        return (r.stdout or "")
    except FileNotFoundError:
        return ""

def norm(s: str) -> str:
    s = (s or "").lower()
    s = re.sub(r"\s{2,}", " ", s)
    return s

def has_any(text: str, words) -> bool:
    t = norm(text)
    return any(w in t for w in words)

def find_vendor_in_text(text: str, vendor_name: str) -> bool:
    if not vendor_name:
        return False
    t = norm(text)
    v = norm(vendor_name)
    # simple containment on normalized vendor name (no fancy fuzzy here)
    return v and v in t

def snippet(text: str, maxlen: int = 220) -> str:
    t = re.sub(r"\s+", " ", (text or "")).strip()
    return t[:maxlen]

def main():
    ap = argparse.ArgumentParser(description="READ-ONLY: classify QUOTE vs NON_QUOTE by PDF first-page proof.")
    ap.add_argument("--in-csv", required=True, help="Estimator review CSV or quote ledger CSV containing source_path, file_ext, vendor_master_match")
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

    proved_quote = 0
    proved_nonquote = 0
    no_text = 0

    for row in rows:
        spath = row.get("source_path","")
        ext = (row.get("file_ext") or "").lower()
        if ext != ".pdf":
            row["doc_role"] = "UNKNOWN"
            row["proof_note"] = "non-pdf"
            continue

        p = Path(spath)
        if not p.exists():
            row["doc_role"] = "UNKNOWN"
            row["proof_note"] = "missing file"
            continue

        text = pdftotext_first_page(p)

        # ---- FORCED RULES (DETERMINISTIC, AUDITABLE) ----
        # These project option PDFs are drawings, not subcontractor bids.
        bn = (p.name or '').lower()
        if bn.startswith('c4wls_option 1') or bn.startswith('c4wls_option 2'):
            row['doc_role'] = 'NON_QUOTE'
            row['proof_note'] = 'option drawing (forced)'
            row['proof_vendor_found'] = 'no'
            row['proof_quote_keywords'] = 'no'
            row['proof_drawing_keywords'] = 'yes'
            row['proof_page'] = '1'
            row['proof_snippet'] = snippet(text)
            proved_nonquote += 1
            continue
        if not text.strip():
            row["doc_role"] = "UNKNOWN"
            row["proof_note"] = "no_text_first_page (pdftotext missing or scanned pdf)"
            no_text += 1
            continue

        vendor = (row.get("vendor_master_match") or "").strip()
        vendor_found = find_vendor_in_text(text, vendor) if vendor else False
        quote_kw = has_any(text, QUOTE_KEYWORDS)
        drawing_kw = has_any(text, DRAWING_KEYWORDS)

        row["proof_vendor_found"] = "yes" if vendor_found else "no"
        row["proof_quote_keywords"] = "yes" if quote_kw else "no"
        row["proof_drawing_keywords"] = "yes" if drawing_kw else "no"
        row["proof_page"] = "1"
        row["proof_snippet"] = snippet(text)

        # Conservative classification:
        # QUOTE requires quote keywords AND (vendor found OR filename is in Sub Bids folder)
        in_subbids = ("sub bids" in spath.lower() or "subbids" in spath.lower())
        if quote_kw and (vendor_found or in_subbids):
            row["doc_role"] = "QUOTE"
            row["proof_note"] = "first-page quote proof"
            proved_quote += 1
        elif drawing_kw and not quote_kw:
            row["doc_role"] = "NON_QUOTE"
            row["proof_note"] = "looks like drawing/plan"
            proved_nonquote += 1
        else:
            row["doc_role"] = "REVIEW"
            row["proof_note"] = "ambiguous first page"

    ts = time.strftime("%Y%m%d_%H%M%S")
    # ---- SAFE OUTPUT NAME ----
    base = in_csv.stem
    safe_base = base[:60]  # cap stem
    short_hash = hex(abs(hash(base)))[2:10]
    out_csv = out_dir / f"{safe_base}__PDF_PROOF_{ts}_{short_hash}.csv"


    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()) if rows else [])
        if rows:
            w.writeheader()
            w.writerows(rows)

    print("OK")
    print(f"OUT: {out_csv}")
    print(f"Summary: quote={proved_quote} non_quote={proved_nonquote} review={len(rows)-proved_quote-proved_nonquote} no_text={no_text} total={len(rows)}")
    print("Note: requires 'pdftotext' installed (poppler). If no_text is high, PDFs may be scanned or poppler missing.")

if __name__ == "__main__":
    main()
