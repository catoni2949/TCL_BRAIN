#!/usr/bin/env python3
import argparse, csv, json, os, re, time
from pathlib import Path

def is_money(s: str) -> bool:
    if s is None:
        return False
    s = str(s).strip()
    if not s:
        return False
    # allow "12,345.67" or "12345" etc
    s = s.replace(",", "")
    return bool(re.match(r'^\d+(\.\d{1,2})?$', s))

def money_norm(s: str) -> str:
    s = str(s).strip().replace(",", "")
    # normalize to 2 decimals
    if "." in s:
        whole, frac = s.split(".", 1)
        frac = (frac + "00")[:2]
        return f"{int(whole)}.{frac}"
    return f"{int(s)}.00"

def safe_stem(stem: str, maxlen: int = 70) -> str:
    keep = []
    for ch in stem:
        if ch.isalnum() or ch in ("-", "_"):
            keep.append(ch)
        else:
            keep.append("_")
    out = "".join(keep)
    return out[:maxlen]

def main():
    ap = argparse.ArgumentParser(description="Apply estimator-selected amounts into a finalized quote ledger (read-only, no Excel writes).")
    ap.add_argument("--pick-csv", required=True, help="Estimator pick sheet CSV (with estimator_selected_amount_clean filled).")
    ap.add_argument("--out-dir", required=True, help="Directory for outputs.")
    args = ap.parse_args()

    pick_csv = Path(args.pick_csv).expanduser()
    out_dir = Path(args.out_dir).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not pick_csv.exists():
        raise SystemExit(f"FATAL: pick csv not found: {pick_csv}")

    with open(pick_csv, newline="", encoding="utf-8") as f:
        r = csv.DictReader(f)
        rows = list(r)

    required_cols = [
        "source_path",
        "vendor_master_match",
        "tcl_trade",
        "option_applicability",
        "candidate_1",
        "estimator_selected_amount_clean",
        "estimator_notes",
    ]
    missing = [c for c in required_cols if c not in (r.fieldnames or [])]
    if missing:
        raise SystemExit(f"FATAL: pick sheet missing columns: {missing}")

    ok = 0
    bad = 0
    out_rows = []
    issues = []

    for i, row in enumerate(rows, start=2):  # 2 = header + first data row
        spath = (row.get("source_path") or "").strip()
        sel = (row.get("estimator_selected_amount_clean") or "").strip()

        if not sel:
            bad += 1
            issues.append({"row": i, "source_path": spath, "issue": "missing estimator_selected_amount_clean"})
            continue
        if not is_money(sel):
            bad += 1
            issues.append({"row": i, "source_path": spath, "issue": f"invalid money format: {sel}"})
            continue

        sel_norm = money_norm(sel)

        out_rows.append({
            "source_path": spath,
            "vendor_master_match": (row.get("vendor_master_match") or "").strip(),
            "tcl_trade": (row.get("tcl_trade") or "").strip(),
            "option_applicability": (row.get("option_applicability") or "").strip(),
            "selected_amount_clean": sel_norm,
            "estimator_notes": (row.get("estimator_notes") or "").strip(),
            "picked_from_candidates": "yes",
            "pick_csv": str(pick_csv),
            "pick_row": str(i),
        })
        ok += 1

    ts = time.strftime("%Y%m%d_%H%M%S")
    base = safe_stem(pick_csv.stem)
    short_hash = hex(abs(hash(pick_csv.stem)))[2:10]

    out_csv = out_dir / f"{base}__AMOUNTS_FINAL_{ts}_{short_hash}.csv"
    out_json = out_dir / f"{base}__AMOUNTS_FINAL_{ts}_{short_hash}.json"

    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(out_rows[0].keys()) if out_rows else [
            "source_path","vendor_master_match","tcl_trade","option_applicability","selected_amount_clean",
            "estimator_notes","picked_from_candidates","pick_csv","pick_row"
        ])
        w.writeheader()
        for rrow in out_rows:
            w.writerow(rrow)

    meta = {
        "timestamp": ts,
        "pick_csv": str(pick_csv),
        "rows_total": len(rows),
        "rows_ok": ok,
        "rows_bad": bad,
        "issues": issues,
        "out_csv": str(out_csv),
    }
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2)

    if bad:
        print("NO-GO")
        print(f"OK_ROWS: {ok}  BAD_ROWS: {bad}")
        print(f"OUT_CSV: {out_csv}")
        print(f"OUT_JSON: {out_json}")
        print("Fix the pick sheet rows listed in OUT_JSON -> issues, then re-run.")
        raise SystemExit(2)

    print("GO")
    print(f"OK_ROWS: {ok}")
    print(f"OUT_CSV: {out_csv}")
    print(f"OUT_JSON: {out_json}")

if __name__ == "__main__":
    main()
