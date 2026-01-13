#!/usr/bin/env python3
import json
from pathlib import Path

from sov_locked_writer import LockedSOVWriter, SourceRef

TEMPLATE = "/Volumes/TCL_DATA 1TB/TCL Partners Dropbox/Ryan Caton/TCL/TCL_SOV_TEMPLATE GBT.xlsm"
LOCK = "/Users/tclserver/TCL_BRAIN/schema/TCL_SOV_TEMPLATE_GBT.lock.v1.json"

OUT = "/Users/tclserver/TCL_BRAIN/reports/DRYRUN_TCL_SOV_TEMPLATE_GBT.xlsm"
AUDIT = "/Users/tclserver/TCL_BRAIN/reports/DRYRUN_TCL_SOV_TEMPLATE_GBT.audit.jsonl"

def pick_cells_from_lock(lock_path: str):
    lock = json.loads(Path(lock_path).read_text())
    row_first = int(lock["governing"]["code_rows"]["span"]["first"])
    allowed_cols = lock["write_policy"]["allowed_columns"]
    denied_cols = lock["write_policy"]["denied_columns"]
    if not allowed_cols or not denied_cols:
        raise SystemExit("Lock file missing allowed/denied columns.")
    # pick first allowed col for legal, first denied for illegal
    legal = f"{allowed_cols[0]}{row_first}"
    illegal = f"{denied_cols[0]}{row_first}"
    return legal, illegal

def main():
    # clean old audit if any
    Path(AUDIT).unlink(missing_ok=True)

    legal_cell, illegal_cell = pick_cells_from_lock(LOCK)

    w = LockedSOVWriter(
        template_path=TEMPLATE,
        lock_path=LOCK,
        audit_log_path=AUDIT
    )

    src = SourceRef(
        source_type="allowance",
        source_path="DRYRUN",
        locator="unit-test",
        notes="Test only; not a real estimate value"
    )

    print(f"LEGAL write target:  {legal_cell}")
    print(f"ILLEGAL write target: {illegal_cell}")

    # 1) Legal write (tiny harmless marker)
    w.write_cell(
        cell_addr=legal_cell,
        new_value="DRYRUN_OK",
        source=src,
        meta={"test": "legal_write"}
    )

    # 2) Illegal write (should hard fail)
    try:
        w.write_cell(
            cell_addr=illegal_cell,
            new_value="DRYRUN_BLOCK",
            source=src,
            meta={"test": "illegal_write_should_fail"}
        )
        raise SystemExit("ERROR: illegal write unexpectedly succeeded (lock failed).")
    except PermissionError as e:
        print(f"Expected BLOCK: {e}")

    # Save output workbook (artifact)
    w.save_as(OUT)
    print(f"Saved dryrun workbook: {OUT}")

    # Confirm audit log wrote exactly 1 line (legal write only)
    lines = Path(AUDIT).read_text().splitlines()
    print(f"Audit lines: {len(lines)}")
    if len(lines) != 1:
        raise SystemExit("ERROR: audit log line count != 1")
    print("DRYRUN PASS")

if __name__ == "__main__":
    main()
