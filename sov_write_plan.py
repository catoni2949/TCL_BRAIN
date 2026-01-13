#!/usr/bin/env python3
"""
Write Plan (DETERMINISTIC, AUDITABLE)

A write plan is a JSON file containing a list of cell writes.
No Excel file is modified unless:
- plan validates against template lock
- every write has a SourceRef
- every write has WriteMeta
"""

import json, os, time
from dataclasses import dataclass, asdict
from typing import Any, Dict, List, Optional

from openpyxl.utils.cell import coordinate_from_string
from openpyxl.utils.cell import column_index_from_string

@dataclass
class SourceRef:
    source_type: str      # "bid_pdf" | "email_quote" | "allowance"
    source_path: str
    locator: str
    notes: str = ""

@dataclass
class WriteMeta:
    project: str
    option: str           # "1" or "2"
    trade: str
    bucket_code: str
    line_id: str
    note: str = ""

@dataclass
class PlannedWrite:
    sheet: str
    cell: str
    value: Any
    source: SourceRef
    meta: WriteMeta

@dataclass
class WritePlan:
    plan_version: str
    created_ts: str
    template_lock_path: str
    template_sha256: str
    writes: List[PlannedWrite]

def now_ts() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")

def load_lock(lock_path: str) -> Dict:
    lock_path = os.path.abspath(os.path.expanduser(lock_path))
    lock = json.loads(open(lock_path, "r", encoding="utf-8").read())
    return lock

def validate_plan(plan: Dict, lock: Dict) -> List[str]:
    errors: List[str] = []

    gov_sheet = lock["governing"]["sheet"]
    allowed_cols = set(lock["write_policy"]["allowed_columns"])
    denied_cols = set(lock["write_policy"]["denied_columns"])
    span = lock["governing"]["code_rows"]["span"]
    row_min = int(span["first"])
    row_max = int(span["last"])
    lock_sha = lock["template"]["sha256"]

    if plan.get("template_sha256") != lock_sha:
        errors.append(f"TEMPLATE SHA mismatch: plan={plan.get('template_sha256')} lock={lock_sha}")

    writes = plan.get("writes", [])
    if not writes:
        errors.append("Plan has no writes.")
        return errors

    for i, w in enumerate(writes):
        path = f"writes[{i}]"

        sheet = w.get("sheet")
        cell = w.get("cell")
        value = w.get("value", None)
        source = w.get("source")
        meta = w.get("meta")

        if sheet != gov_sheet:
            errors.append(f"{path}: sheet '{sheet}' is not governing sheet '{gov_sheet}'")

        if not cell or not isinstance(cell, str):
            errors.append(f"{path}: missing/invalid cell")
            continue

        try:
            col_letter, row = coordinate_from_string(cell)
            row = int(row)
        except Exception:
            errors.append(f"{path}: invalid cell address '{cell}'")
            continue

        if row < row_min or row > row_max:
            # Allow header block writes above the code table (rows 1-12)
            if row > 12:
                errors.append(f"{path}: row {row} outside code row span {row_min}-{row_max}")

        if col_letter in denied_cols:
            errors.append(f"{path}: column {col_letter} is LOCKED")
        if col_letter not in allowed_cols:
            errors.append(f"{path}: column {col_letter} not in allowed write columns")

        # SourceRef required
        if not source or not isinstance(source, dict):
            errors.append(f"{path}: missing source")
        else:
            for k in ("source_type","source_path","locator"):
                if not source.get(k):
                    errors.append(f"{path}: source missing {k}")

        # WriteMeta required
        if not meta or not isinstance(meta, dict):
            errors.append(f"{path}: missing meta")
        else:
            for k in ("project","option","trade","bucket_code","line_id"):
                if not meta.get(k):
                    errors.append(f"{path}: meta missing {k}")

        # Value rule: we don't forbid strings/numbers here; writer wrapper will still block formulas/merges at apply time.
        _ = value

    return errors

def main():
    import argparse
    ap = argparse.ArgumentParser(description="Validate a write plan JSON against template lock (no Excel writes).")
    ap.add_argument("--plan", required=True)
    ap.add_argument("--lock", required=True)
    args = ap.parse_args()

    plan_path = os.path.abspath(os.path.expanduser(args.plan))
    lock_path = os.path.abspath(os.path.expanduser(args.lock))

    plan = json.loads(open(plan_path, "r", encoding="utf-8").read())
    lock = load_lock(lock_path)

    errs = validate_plan(plan, lock)
    if errs:
        print("NO-GO")
        for e in errs[:200]:
            print("-", e)
        raise SystemExit(2)

    print("GO")
    print(f"Validated writes: {len(plan.get('writes', []))}")

if __name__ == "__main__":
    main()
