#!/usr/bin/env python3
import argparse, csv, re, subprocess, time, json
from pathlib import Path

# Money patterns (conservative)
MONEY_RE = re.compile(
    r"(?i)(total|proposal\s+total|bid\s+total|grand\s+total|sum|price|amount)\s*[:\-]?\s*\$?\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{2})?|[0-9]+(?:\.[0-9]{2})?)"
)
ALT_MONEY_RE = re.compile(r"\$\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{2})?)")

OPT1_RE = re.compile(r"(?i)\b(option\s*one|option\s*1|opt\s*1)\b")
OPT2_RE = re.compile(r"(?i)\b(option\s*two|option\s*2|opt\s*2)\b")

EXCL_RE = re.compile(r"(?i)\b(exclusion|excluded|not included)\b")
SCOPE_RE = re.compile(r"(?i)\b(scope|work includes|we will provide|furnish and install|labor and materials)\b")

# Extra “quote-ish” words to help classify (still not pricing)
QUOTE_WORDS = re.compile(r"(?i)\b(proposal|estimate|quotation|quote|bid summary|budget estimate)\b")

def pdftotext_pages(pdf_path: Path, pages: int = 2) -> str:
    try:
        out = subprocess.check_output(
            ["pdftotext", "-f", "1", "-l", str(pages), str(pdf_path), "-"],
            stderr=subprocess.DEVNULL
        )
        return out.decode("utf-8", errors="ignore")
    except Exception:
        return ""

def clean_money(s: str) -> str:
    return s.strip().replace(",", "")

def normalize_vendor_tokens(name: str):
    # Very light normalization; we want “proof”, not fuzzy magic
    n = re.sub(r"(?i)\b(inc|llc|co|company|corp|corporation|ltd)\b\.?", " ", name)
    n = re.sub(r"[^a-z0-9]+", " ", n.lower()).strip()
    toks = [t for t in n.split() if len(t) >= 3]
    return toks

def vendor_found_in_text(text: str, vendor: str) -> bool:
    toks = normalize_vendor_tokens(vendor)
    if not toks:
        return False
    tlow = re.sub(r"[^a-z0-9]+", " ", text.lower())
    hits = sum(1 for t in toks if t in tlow)
    # require at least 2 token hits if vendor has 2+ tokens, else 1 hit
    if len(toks) >= 2:
        return hits >= 2
    return hits >= 1

def pick_money_candidates(text: str):
    cands = []
    for m in MONEY_RE.finditer(text):
        label = (m.group(1) or "").strip()
        amt_raw = (m.group(2) or "").strip()
        amt = clean_money(amt_raw)
        ctx = " ".join(m.group(0).split())
        score = 0
        if re.search(r"(?i)\b(total|grand)\b", label): score += 3
        if re.search(r"(?i)\bproposal\b", label): score += 2
        if re.search(r"(?i)\bbid\b", label): score += 2
        if "$" in m.group(0): score += 1
        if len(ctx) > 20: score += 1
        cands.append({"label": label, "amount_clean": amt, "amount_raw": amt_raw, "context": ctx, "score": score})

    if not cands:
        for m2 in ALT_MONEY_RE.finditer(text):
            amt_raw = m2.group(1)
            cands.append({"label": "", "amount_clean": clean_money(amt_raw), "amount_raw": amt_raw, "context": m2.group(0).strip(), "score": 0})

    # sort by score desc, then amount length desc (weak tie-breaker)
    cands.sort(key=lambda d: (d["score"], len(d["context"])), reverse=True)
    return cands

def option_tag(text: str) -> str:
    has1 = bool(OPT1_RE.search(text))
    has2 = bool(OPT2_RE.search(text))
    if has1 and has2:
        return "BOTH (explicit)"
    if has1:
        return "OPTION 1 (explicit)"
    if has2:
        return "OPTION 2 (explicit)"
    return "UNKNOWN"

def snippet_near(text: str, pattern: re.Pattern, max_chars: int = 220) -> str:
    m = pattern.search(text)
    if not m:
        return ""
    start = max(0, m.start() - 80)
    end = min(len(text), m.end() + 140)
    s = " ".join(text[start:end].split())
    return s[:max_chars]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--quote-approved-csv", required=True)
    ap.add_argument("--out-dir", required=True)
    ap.add_argument("--pages", type=int, default=2)
    ap.add_argument("--emit-json", action="store_true", help="also emit a JSON alongside CSV (same basename)")
    args = ap.parse_args()

    in_csv = Path(args.quote_approved_csv)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    with open(in_csv, newline="", encoding="utf-8") as f:
        r = csv.DictReader(f)
        rows = list(r)

    ts = time.strftime("%Y%m%d_%H%M%S")
    base = in_csv.stem[:50]
    short_hash = hex(abs(hash(in_csv.stem)))[2:10]
    out_csv = out_dir / f"{base}__QUOTE_LEDGER_EXTRACT_{ts}_{short_hash}.csv"
    out_json = out_dir / f"{base}__QUOTE_LEDGER_EXTRACT_{ts}_{short_hash}.json"

    out_rows = []
    missing_text = 0

    for row in rows:
        spath = (row.get("source_path") or "").strip()
        pdf = Path(spath)
        vendor = (row.get("vendor_master_match") or "").strip()
        trade = (row.get("tcl_trade") or "").strip()

        out = {
            "source_path": spath,
            "vendor_master_match": vendor,
            "tcl_trade": trade,
            "doc_role": (row.get("doc_role") or "").strip(),
            "option_applicability": "",
            "vendor_proved_in_text": "",
            "amount_found": "no",
            "amount_clean": "",
            "amount_raw": "",
            "amount_label": "",
            "amount_context": "",
            "amount_candidates_count": "0",
            "amount_candidates_json": "",
            "scope_snippet": "",
            "exclusions_snippet": "",
            "extract_pages": str(args.pages),
            "extract_status": "",
            "needs_review": "",
            "review_reason": ""
        }

        if not pdf.exists():
            out["extract_status"] = "MISSING_FILE"
            out["needs_review"] = "yes"
            out["review_reason"] = "missing file on disk"
            out_rows.append(out)
            continue

        text = pdftotext_pages(pdf, pages=args.pages)
        if not text.strip():
            missing_text += 1
            out["extract_status"] = "NO_TEXT"
            out["needs_review"] = "yes"
            out["review_reason"] = "no text extracted (scanned? encrypted?)"
            out_rows.append(out)
            continue

        out["option_applicability"] = option_tag(text)
        out["scope_snippet"] = snippet_near(text, SCOPE_RE)
        out["exclusions_snippet"] = snippet_near(text, EXCL_RE)

        vfound = vendor_found_in_text(text, vendor) if vendor else False
        out["vendor_proved_in_text"] = "yes" if vfound else "no"

        cands = pick_money_candidates(text)
        out["amount_candidates_count"] = str(len(cands))
        out["amount_candidates_json"] = json.dumps(cands[:12], ensure_ascii=False)

        if cands:
            best = cands[0]
            out["amount_found"] = "yes"
            out["amount_clean"] = best.get("amount_clean","")
            out["amount_raw"] = best.get("amount_raw","")
            out["amount_label"] = best.get("label","")
            out["amount_context"] = best.get("context","")
            out["extract_status"] = "OK"
        else:
            out["extract_status"] = "NO_AMOUNT_FOUND"

        # Review logic (strict)
        reasons = []
        if out["extract_status"] != "OK":
            reasons.append(out["extract_status"])
        if vendor and out["vendor_proved_in_text"] == "no":
            reasons.append("vendor_mismatch_or_not_found_in_text")
        if len(cands) > 1:
            reasons.append("multi_amounts_found (possible alternates/options)")
        if out["option_applicability"] == "UNKNOWN":
            reasons.append("option_unknown (no explicit opt text)")

        out["needs_review"] = "yes" if reasons else "no"
        out["review_reason"] = "; ".join(reasons) if reasons else ""

        out_rows.append(out)

    fieldnames = list(out_rows[0].keys()) if out_rows else []
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(out_rows)

    if args.emit_json:
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(out_rows, f, ensure_ascii=False, indent=2)

    print("OK")
    print(f"IN:  {in_csv}")
    print(f"OUT: {out_csv}")
    if args.emit_json:
        print(f"JSON: {out_json}")
    print(f"Summary: rows={len(out_rows)} no_text={missing_text}")

if __name__ == "__main__":
    main()
