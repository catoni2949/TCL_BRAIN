#!/usr/bin/env python3
from __future__ import annotations

import json
import re
import sys
import hashlib
from pathlib import Path
from collections import defaultdict
from typing import Any, Dict, List, Tuple, Optional


A1_RE = re.compile(r"^[A-Z]{1,3}[1-9][0-9]*$")

# Hard guard: block writing into formula column(s).
# From your failures: column T (e.g., T2, T26) is a formula column target.
FORMULA_COLS = {"T"}

# Required meta keys (from your validator output)
REQUIRED_META_KEYS = ["project", "option", "trade", "bucket_code", "line_id"]


def sha256_json(obj: Any) -> str:
    b = json.dumps(obj, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    return hashlib.sha256(b).hexdigest()


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def get_lock_sha(lock: Dict[str, Any]) -> Optional[str]:
    # Accept common keys
    for k in ("template_sha256", "template_sha", "sha256", "sha"):
        v = lock.get(k)
        if isinstance(v, str) and v:
            return v
    return None


def cell_to_col(cell: str) -> Optional[str]:
    # cell is validated as A1 already; return column letters
    m = re.match(r"^([A-Z]{1,3})[0-9]+$", cell)
    return m.group(1) if m else None


def write_has_required_meta(w: Dict[str, Any]) -> bool:
    meta = w.get("meta")
    if not isinstance(meta, dict):
        return False
    for k in REQUIRED_META_KEYS:
        if k not in meta or meta.get(k) in (None, ""):
            return False
    return True


def write_has_required_source(w: Dict[str, Any]) -> Tuple[bool, bool]:
    """
    Returns (has_source, has_complete_source)
    """
    src = w.get("source")
    if src is None:
        return (False, False)
    if not isinstance(src, dict):
        return (True, False)
    need = ["source_type", "source_path", "locator"]
    ok = all((k in src and src.get(k) not in (None, "")) for k in need)
    return (True, ok)


def plan_template_sha(plan: Dict[str, Any]) -> Optional[str]:
    v = plan.get("template_sha256") or plan.get("template_sha")
    return v if isinstance(v, str) and v else None


def is_valid_write_obj(w: Any) -> bool:
    return isinstance(w, dict)


def extract_target(w: Dict[str, Any]) -> Tuple[Optional[str], Optional[str]]:
    """
    Returns (sheet, cell) if present.
    """
    sheet = w.get("sheet")
    cell = w.get("cell")
    if isinstance(sheet, str) and sheet and isinstance(cell, str) and cell:
        return sheet, cell
    return None, None


def main() -> int:
    if len(sys.argv) != 4:
        print("usage: batch_select_plans.py <plans_dir> <lock.json> <out_txt>", file=sys.stderr)
        return 2

    plans_dir = Path(sys.argv[1])
    lock_path = Path(sys.argv[2])
    out_txt = Path(sys.argv[3])

    lock = load_json(lock_path)
    if not isinstance(lock, dict):
        print(f"lock is not a dict: {lock_path}", file=sys.stderr)
        return 2

    lock_sha = get_lock_sha(lock)
    print("LOCK_SHA:", lock_sha)

    skipped = defaultdict(int)
    ready: List[str] = []

    # Where we write fixed plans (plan_hash injected only)
    ready_dir = plans_dir / "_batch_ready"
    ready_dir.mkdir(parents=True, exist_ok=True)

    # Scan *.json under plans_dir (non-recursive + reports style; include nested if needed)
    candidates = sorted(set(plans_dir.rglob("*.json")))

    for p in candidates:
        # ignore things in signatures by default
        if "signatures" in p.parts:
            skipped["signatures"] += 1
            continue

        try:
            plan = load_json(p)
        except Exception:
            skipped["bad_json"] += 1
            continue

        if not isinstance(plan, dict):
            skipped["not_dict"] += 1
            continue

        writes = plan.get("writes")
        if not isinstance(writes, list) or not writes:
            skipped["no_writes"] += 1
            continue

        # Require template sha present (per your latest rule)
        plan_sha = plan_template_sha(plan)
        if not plan_sha:
            skipped["missing_template_sha"] += 1
            continue

        # If lock_sha is known, require match
        if lock_sha and plan_sha != lock_sha:
            skipped["sha_mismatch"] += 1
            continue

        # Per-plan checks
        bad = False
        seen_targets: Dict[Tuple[str, str], int] = {}

        for idx, w in enumerate(writes):
            if not is_valid_write_obj(w):
                skipped["bad_write_obj"] += 1
                bad = True
                break

            sheet, cell = extract_target(w)
            if not sheet or not cell:
                skipped["missing_cell_or_sheet"] += 1
                bad = True
                break

            if not A1_RE.match(cell):
                skipped["invalid_cell"] += 1
                bad = True
                break

            col = cell_to_col(cell)
            if col in FORMULA_COLS:
                skipped["formula_col_target"] += 1
                bad = True
                break

            key = (sheet, cell)
            seen_targets[key] = seen_targets.get(key, 0) + 1
            if seen_targets[key] > 1:
                skipped["dup_cell_in_plan"] += 1
                bad = True
                break

            has_src, src_complete = write_has_required_source(w)
            if not has_src:
                skipped["missing_source"] += 1
                bad = True
                break
            if not src_complete:
                skipped["incomplete_source"] += 1
                bad = True
                break

            if not write_has_required_meta(w):
                skipped["missing_meta"] += 1
                bad = True
                break

        if bad:
            continue

        # Ensure plan_hash exists (compute; doesn’t change writes)
        if not plan.get("plan_hash"):
            plan["plan_hash"] = sha256_json(writes)
            outp = ready_dir / p.name
            outp.write_text(json.dumps(plan, indent=2, sort_keys=True), encoding="utf-8")
            ready.append(str(outp))
            skipped["fixed_plan_hash"] += 1
        else:
            ready.append(str(p))

    out_txt.parent.mkdir(parents=True, exist_ok=True)
    out_txt.write_text("\n".join(ready) + ("\n" if ready else ""), encoding="utf-8")

    print("READY:", len(ready))
    print("WROTE:", out_txt)
    print("SKIPPED:", dict(skipped))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
