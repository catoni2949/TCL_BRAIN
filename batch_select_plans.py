#!/usr/bin/env python3
import json
import re, re, sys
from pathlib import Path
from collections import Counter

A1 = re.compile(r"^[A-Z]{1,3}[1-9][0-9]*$")

def cell_col(a1: str) -> str:
    m = re.match(r"^([A-Z]+)", (a1 or "").upper())
    return m.group(1) if m else ""

def get_forbidden_cols(lock: dict) -> set:
    # Prefer explicit lock config if present; otherwise default to Excel column T
    cols = set()
    for k in ("formula_columns", "forbidden_columns", "locked_formula_columns"):
        v = lock.get(k)
        if isinstance(v, list):
            cols |= {str(x).upper() for x in v}
    if not cols:
        cols = {"T"}
    return cols

def sha256_json(obj):
    import hashlib
    b = json.dumps(obj, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    return hashlib.sha256(b).hexdigest()

def load_lock_sha(lock_path: Path) -> str:
    lock = json.loads(lock_path.read_text(encoding="utf-8"))
    return (lock.get("template", {}) or {}).get("sha256") or ""

def main():
    if len(sys.argv) != 4:
        print("usage: batch_select_plans.py <reports_dir> <lock_json> <out_txt>", file=sys.stderr)
        return 2

    reports = Path(sys.argv[1])
    lock_path = Path(sys.argv[2])
    out_txt = Path(sys.argv[3])

    lock_sha = load_lock_sha(lock_path)
    skipped = Counter()
    ready = []

    for p in sorted(reports.glob("*.json")):
        try:
            plan = json.loads(p.read_text(encoding="utf-8"))
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

        plan_sha = plan.get("template_sha256") or plan.get("template_sha")
    if not plan_sha:
        skipped["missing_plan_sha"] += 1
        continue
plan_sha = plan.get("template_sha256") or plan.get("template_sha")
        if plan_sha and lock_sha and plan_sha != lock_sha:
            skipped["sha_mismatch"] += 1
            continue

        bad = False
        seen = set()

        for w in writes:
            if not isinstance(w, dict):
                skipped["bad_write_obj"] += 1
                bad = True; break

            cell = w.get("cell")
            sheet = w.get("sheet")
            if not sheet or not cell or not isinstance(cell, str):
                skipped["missing_cell_or_sheet"] += 1
                bad = True; break

            if not A1.match(cell):
                skipped["invalid_cell"] += 1
                bad = True; break

            key = (sheet, cell)
            if key in seen:
                skipped["dup_cell_in_plan"] += 1
                bad = True; break
            seen.add(key)

            src = w.get("source")
            if not isinstance(src, dict):
                skipped["missing_source"] += 1
                bad = True; break
            for k in ("source_type", "source_path", "locator"):
                if not src.get(k):
                    skipped["incomplete_source"] += 1
                    bad = True; break
            if bad:
                break

        if bad:
            continue

        if not plan.get("plan_hash"):
            plan["plan_hash"] = sha256_json(writes)
            p_fixed = reports / "_batch_ready"
            p_fixed.mkdir(parents=True, exist_ok=True)
            outp = p_fixed / p.name
            outp.write_text(json.dumps(plan, indent=2, sort_keys=True), encoding="utf-8")
            ready.append(str(outp))
            skipped["fixed_plan_hash"] += 1
        else:
            ready.append(str(p))

    out_txt.parent.mkdir(parents=True, exist_ok=True)
    out_txt.write_text("\n".join(ready) + ("\n" if ready else ""), encoding="utf-8")

    print("LOCK_SHA:", lock_sha)
    print("READY:", len(ready))
    print("WROTE:", out_txt)
    print("SKIPPED:", dict(skipped))
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
