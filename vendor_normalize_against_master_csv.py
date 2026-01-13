#!/usr/bin/env python3
import argparse, csv, os, re, time
from pathlib import Path
from difflib import SequenceMatcher

def norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s).lower().strip()
    s = s.replace("&", " and ")
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\b(inc|llc|ltd|co|company|corporation|corp|services|service|group)\b", " ", s)
    s = re.sub(r"\s{2,}", " ", s).strip()
    return s

def best_fuzzy(name_norm: str, master_norm_list):
    best_norm = ""
    best_score = 0
    for mn in master_norm_list:
        if not mn:
            continue
        sc = int(round(SequenceMatcher(None, name_norm, mn).ratio() * 100))
        if sc > best_score:
            best_score = sc
            best_norm = mn
    return best_norm, best_score

def load_master_csv(path: Path):
    rows = []
    with open(path, "r", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            vc = (row.get("vendor_company") or "").strip()
            if not vc:
                continue
            rows.append({
                "trade_section": (row.get("trade_section") or "").strip(),
                "vendor_company": vc,
                "contact": (row.get("contact") or "").strip(),
                "phone": (row.get("phone") or "").strip(),
                "email": (row.get("email") or "").strip(),
                "vendor_norm": norm(vc)
            })
    return rows

def main():
    ap = argparse.ArgumentParser(description="Read-only: normalize ledger vendor names against extracted master vendor CSV.")
    ap.add_argument("--ledger-csv", required=True)
    ap.add_argument("--master-csv", required=True)
    ap.add_argument("--out-dir", required=True)
    args = ap.parse_args()

    ledger_csv = Path(os.path.expanduser(args.ledger_csv)).resolve()
    master_csv = Path(os.path.expanduser(args.master_csv)).resolve()
    out_dir = Path(os.path.expanduser(args.out_dir)).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not ledger_csv.exists():
        raise SystemExit(f"FATAL: ledger csv not found: {ledger_csv}")
    if not master_csv.exists():
        raise SystemExit(f"FATAL: master csv not found: {master_csv}")

    master = load_master_csv(master_csv)
    master_norms = [m["vendor_norm"] for m in master]
    norm_to_master = {}
    for m in master:
        if m["vendor_norm"] and m["vendor_norm"] not in norm_to_master:
            norm_to_master[m["vendor_norm"]] = m

    rows = []
    with open(ledger_csv, "r", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            rows.append(row)

    exact = fuzzy_ok = review = none = 0

    for row in rows:
        v = (row.get("vendor_normalized") or row.get("vendor_raw") or "").strip()
        vn = norm(v)

        row["vendor_master_match"] = ""
        row["master_trade_section"] = ""
        row["match_type"] = "none"
        row["match_score"] = "0"
        row["needs_review"] = "yes"

        if not vn:
            none += 1
            continue

        if vn in norm_to_master:
            m = norm_to_master[vn]
            row["vendor_master_match"] = m["vendor_company"]
            row["master_trade_section"] = m["trade_section"]
            row["match_type"] = "exact"
            row["match_score"] = "100"
            row["needs_review"] = "no"
            exact += 1
            continue

        best_norm, score = best_fuzzy(vn, master_norms)
        if best_norm:
            m = norm_to_master.get(best_norm)
            if m:
                row["vendor_master_match"] = m["vendor_company"]
                row["master_trade_section"] = m["trade_section"]
                row["match_type"] = "fuzzy"
                row["match_score"] = str(score)
                # thresholds
                if score >= 92:
                    row["needs_review"] = "no"
                    fuzzy_ok += 1
                else:
                    row["needs_review"] = "yes"
                    review += 1
            else:
                none += 1
        else:
            none += 1

    ts = time.strftime("%Y%m%d_%H%M%S")
    out_csv = out_dir / f"{ledger_csv.stem}__vendor_normalized_masterCSV_{ts}.csv"

    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()) if rows else [])
        if rows:
            w.writeheader()
            w.writerows(rows)

    print("OK")
    print(f"OUT: {out_csv}")
    print(f"Summary: exact={exact} fuzzy_ok={fuzzy_ok} needs_review={review} none_or_blank={none} total={len(rows)}")

if __name__ == "__main__":
    main()
