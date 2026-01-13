#!/usr/bin/env python3
"""
Apply Write Plan (ENFORCED)

- Loads plan JSON
- Loads template lock JSON
- Validates plan against lock (no Excel writes if NO-GO)
- Applies writes via LockedSOVWriter (enforces merged/formula protection too)
- Saves output workbook + audit JSONL

ZERO tolerance:
- missing source/meta -> fail
- illegal cell -> fail
- wrong sheet -> fail
"""

import argparse, json, os
from pathlib import Path

from sov_write_plan import validate_plan, load_lock
from sov_locked_writer import LockedSOVWriter, SourceRef

def main():
    ap = argparse.ArgumentParser(description="Apply a validated write plan to the SOV template using the locked writer.")
    ap.add_argument("--plan", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--lock", required=True)
    ap.add_argument("--out-xlsm", required=True)
    ap.add_argument("--audit-jsonl", required=True)
    args = ap.parse_args()

    plan_path = Path(os.path.expanduser(args.plan)).resolve()
    template_path = Path(os.path.expanduser(args.template)).resolve()
    lock_path = Path(os.path.expanduser(args.lock)).resolve()
    out_xlsm = Path(os.path.expanduser(args.out_xlsm)).resolve()
    audit = Path(os.path.expanduser(args.audit_jsonl)).resolve()

    if not plan_path.exists():
        raise SystemExit(f"FATAL: plan not found: {plan_path}")
    if not template_path.exists():
        raise SystemExit(f"FATAL: template not found: {template_path}")
    if not lock_path.exists():
        raise SystemExit(f"FATAL: lock not found: {lock_path}")

    plan = json.loads(plan_path.read_text())
    lock = load_lock(str(lock_path))

    errs = validate_plan(plan, lock)
    if errs:
        print("NO-GO (plan validation failed)")
        for e in errs[:200]:
            print("-", e)
        raise SystemExit(2)

    # clear any prior audit file
    audit.parent.mkdir(parents=True, exist_ok=True)
    if audit.exists():
        audit.unlink()

    w = LockedSOVWriter(
        template_path=str(template_path),
        lock_path=str(lock_path),
        audit_log_path=str(audit)
    )

    writes = plan.get("writes", [])
    for i, pw in enumerate(writes):
        srcd = pw["source"]
        src = SourceRef(
            source_type=srcd["source_type"],
            source_path=srcd["source_path"],
            locator=srcd["locator"],
            notes=srcd.get("notes", "")
        )
        w.write_cell(
            cell_addr=pw["cell"],
            new_value=((pw.get('value') if isinstance(pw, dict) else None) or ((pw.get('write') or {}).get('value') if isinstance(pw, dict) else None)),
            source=src,
            meta=pw["meta"]
        )

    out_xlsm.parent.mkdir(parents=True, exist_ok=True)
    w.save_as(str(out_xlsm))

    print("OK")
    print(f"Writes applied: {len(writes)}")
    print(f"OUT:   {out_xlsm}")
    print(f"AUDIT: {audit}")

if __name__ == "__main__":
    main()
