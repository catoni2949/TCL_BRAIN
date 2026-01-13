#!/usr/bin/env python3
import argparse, csv, os, re, time
from pathlib import Path

def norm(s: str) -> str:
    s = (s or "").strip().lower()
    s = re.sub(r"\s{2,}", " ", s)
    return s

# Conservative mapping: only map when obvious.
# Everything else stays UNMAPPED for human review.
DEFAULT_MAP = {
    "final clean": "Final Cleaning",
    "demo": "Demolition",
    "demolition": "Demolition",
    "drywall": "Drywall / ACT",
    "framing": "Drywall / ACT",
    "acoustical ceilings": "Drywall / ACT",
    "act": "Drywall / ACT",
    "flooring": "Flooring",
    "paint": "Painting",
    "doors": "Doors / Hardware",
    "door & hardware": "Doors / Hardware",
    "millwork": "Millwork",
    "casework": "Millwork",
    "electrical": "Electrical",
    "plumbing": "Plumbing",
    "hvac": "Mechanical",
    "mechanical": "Mechanical",
    "fire protection": "Fire Sprinkler",
    "sprinkler": "Fire Sprinkler",
    "fire alarm": "Fire Alarm",
    "glazing": "Glass / Glazing",
    "storefront": "Glass / Glazing",
    "concrete": "Concrete",
    "scaffolding": "Scaffolding",
    "dumpsters": "Dumpsters",
    "hauling": "Demo / Haul-Off",
    "abatement": "Hazmat / Abatement",
    "insulation": "Insulation",
    "roofing": "Roofing",
    "landscaping": "Landscaping",
    "signage": "Signage",
    "low voltage": "Low Voltage",
    "data": "Low Voltage",
    "controls": "Controls",
}

def main():
    ap = argparse.ArgumentParser(description="Build a GLOBAL mapping file from master trade_section -> TCL trade (conservative defaults).")
    ap.add_argument("--master-csv", required=True, help="TCL_Bid_List_Master_Vendors_*.csv")
    ap.add_argument("--out-dir", required=True)
    args = ap.parse_args()

    master_csv = Path(os.path.expanduser(args.master_csv)).resolve()
    out_dir = Path(os.path.expanduser(args.out_dir)).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    if not master_csv.exists():
        raise SystemExit(f"FATAL: master not found: {master_csv}")

    sections = set()
    with open(master_csv, "r", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            sec = (row.get("trade_section") or "").strip()
            if sec:
                sections.add(sec)

    rows = []
    for sec in sorted(sections, key=lambda x: x.lower()):
        sec_n = norm(sec)
        tcl = ""
        # match by normalized equality against DEFAULT_MAP keys
        for k, v in DEFAULT_MAP.items():
            if sec_n == norm(k):
                tcl = v
                break
        rows.append({
            "master_trade_section": sec,
            "tcl_trade_suggested": tcl,
            "status": "MAPPED" if tcl else "UNMAPPED",
            "notes": ""
        })

    ts = time.strftime("%Y%m%d_%H%M%S")
    out_csv = out_dir / f"MASTER_SECTION_TO_TCL_TRADE_MAP_{ts}.csv"

    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["master_trade_section","tcl_trade_suggested","status","notes"])
        w.writeheader()
        w.writerows(rows)

    print("OK")
    print(f"SECTIONS: {len(rows)}")
    print(f"OUT: {out_csv}")

if __name__ == "__main__":
    main()
