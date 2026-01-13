"""
Single authorized write API for TCL SOV automation.

Any Excel write must go through:
- LockedSOVWriter
- Template lock enforcement
- Mandatory SourceRef
- Mandatory audit log

If someone tries to write without this, it's a bug by definition.
"""

from dataclasses import dataclass
from typing import Any, Dict, Optional

from sov_locked_writer import LockedSOVWriter, SourceRef

@dataclass
class WriteMeta:
    project: str
    option: str                # "1" or "2"
    trade: str
    bucket_code: str           # template bucket code like "1000"
    line_id: str               # stable ID for traceability (e.g., "UR_HVAC_BASE" or UUID)
    note: str = ""

def open_writer(template_path: str, lock_path: str, audit_log_path: str, plan_hash=None) -> LockedSOVWriter:
    return LockedSOVWriter(template_path=template_path, lock_path=lock_path, audit_log_path=audit_log_path, plan_hash=plan_hash)

def write_value(writer: LockedSOVWriter,
                cell_addr: str,
                new_value: Any,
                source: SourceRef,
                meta: WriteMeta,
                extra: Optional[Dict[str, Any]] = None):
    payload = {
        "project": meta.project,
        "option": meta.option,
        "trade": meta.trade,
        "bucket_code": meta.bucket_code,
        "line_id": meta.line_id,
        "note": meta.note,
    }
    if extra:
        payload.update(extra)
    writer.write_cell(cell_addr=cell_addr, new_value=new_value, source=source, meta=payload)
