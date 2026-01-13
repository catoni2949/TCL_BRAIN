#!/usr/bin/env python3
"""
Apply Write Plan (HARDENED)

Hardening:
- Requires TCL_HARDENED_APPLY=1
- Requires plan_hash and enforces it against writes (tamper detection)
- Prefight:
  - No duplicate (sheet, cell)
  - Valid Excel cell refs only
  - Never writes to column T (known formula col)
- Validates plan against lock using validate_plan (same as enforced apply)
"""

import argparse
import collections
import hashlib
import json
import os
import re
import sys
from pathlib import Path

from sov_write_plan import validate_plan, load_lock
from sov_locked_writer import LockedSOVWriter, SourceRef

RX_CELL = re.compile(r"^[A-Z]{1,3}[0-9]{1,5}$")


def fatal(msg: str):
    print(f"❌ {msg}")
    raise SystemExit(2)


def sha256_json(obj) -> str:
    return hashlib.sha256(
        json.dumps(obj, sort_keys=True, separators=(",", ":")).encode("utf-8")
    ).hexdigest()


def _get_write_value(w: dict):
    # Accept either legacy shape or normalized shape
    if "value" in w and w["value"] is not None:
        return w["value"]
    wr = w.get("write")
    if isinstance(wr, dict) and wr.get("value") is not None:
        return wr.get("value")
    return None


def _normalize_source(srcd: dict) -> dict:
    # Accept either:
    #  {source_type, source_path, locator, notes}
    # or legacy-ish:
    #  {type, path, locator, notes}
    if not isinstance(srcd, dict):
        return {"source_type": "unknown", "source_path": "unknown", "locator": "unknown", "notes": ""}

    st = srcd.get("source_type") or srcd.get("type") or "unknown"
    sp = srcd.get("source_path") or srcd.get("path") or "unknown"
    loc = srcd.get("locator") or "unknown"
    notes = srcd.get("notes") or ""
    return {"source_type": st, "source_path": sp, "locator": loc, "notes": notes}


def main():
    if os.environ.get("TCL_HARDENED_APPLY") != "1":
        fatal("Direct apply disabled. Run with TCL_HARDENED_APPLY=1")

    ap = argparse.ArgumentParser(description="Apply a validated write plan to the SOV template (HARDENED).")
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
        fatal(f"plan not found: {plan_path}")
    if not template_path.exists():
        fatal(f"template not found: {template_path}")
    if not lock_path.exists():
        fatal(f"lock not found: {lock_path}")

    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    lock = load_lock(str(lock_path))
    lock_sha = lock["template"]["sha256"]

    # --- template sha must match ---
    if plan.get("template_sha256") != lock_sha:
        fatal(f"TEMPLATE SHA mismatch: plan={plan.get('template_sha256')} lock={lock_sha}")

    writes = plan.get("writes", [])
    if not isinstance(writes, list) or not writes:
        fatal("Plan has no writes")

    # --- HARDENING: plan fingerprint enforcement ---
    expected = plan.get("plan_hash")
    actual = sha256_json(writes)
    if not expected:
        fatal("plan_hash missing from plan")
    if expected != actual:
        fatal(f"PLAN TAMPER DETECTED: expected={expected} actual={actual}")

    # --- PREFLIGHT (HARDENED) ---
    cnt = collections.Counter((w.get("sheet"), w.get("cell")) for w in writes)
    for (sh, cell), n in cnt.items():
        if n > 1:
            fatal(f"DUPLICATE CELL: {sh}!{cell} x{n}")

    for i, w in enumerate(writes):
        cell = (w.get("cell") or "").strip()
        if not RX_CELL.match(cell):
            fatal(f"INVALID CELL at writes[{i}]: {cell}")
        if cell.startswith("T"):
            fatal(f"FORMULA COLUMN TARGET at writes[{i}]: {cell}")
        if not w.get("meta"):
            fatal(f"Missing meta at writes[{i}]")
        if not w.get("source"):
            fatal(f"Missing source at writes[{i}]")

    print(f"✅ PREFLIGHT OK: {len(writes)} writes")

    # Full lock validation (policy enforcement)
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

    writer = LockedSOVWriter(
        template_path=str(template_path),
        lock_path=str(lock_path),
        audit_log_path=str(audit),
    )

    for i, w in enumerate(writes):
        srcd = _normalize_source(w.get("source") or {})
        src = SourceRef(
            source_type=srcd["source_type"],
            source_path=srcd["source_path"],
            locator=srcd["locator"],
            notes=srcd.get("notes", ""),
        )

        val = _get_write_value(w)
        writer.write_cell(
            cell_addr=w["cell"],
            new_value=val,
            source=src,
            meta=w["meta"],
        )

    out_xlsm.parent.mkdir(parents=True, exist_ok=True)
    writer.save_as(str(out_xlsm))

    print("✅ OK")
    print(f"Writes applied: {len(writes)}")
    print(f"OUT:   {out_xlsm}")
    print(f"AUDIT: {audit}")


if __name__ == "__main__":
    main()
