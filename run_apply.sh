#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

: "${PLAN:?set PLAN=/path/to/plan.json}"
: "${TPL:?set TPL=/path/to/template.xlsm}"
: "${LOCK:?set LOCK=/path/to/lock.json}"

TS="$(date +%Y%m%d_%H%M%S)"
OUT_XLSM="reports/APPLIED_HARDENED_${TS}.xlsm"
OUT_AUDIT="reports/APPLIED_HARDENED_${TS}.audit.jsonl"

./_hardened_guard.sh
./_git_clean_guard.sh

TCL_HARDENED_APPLY=1 ./apply_write_plan.py \
  --plan "$PLAN" \
  --template "$TPL" \
  --lock "$LOCK" \
  --out-xlsm "$OUT_XLSM" \
  --audit-jsonl "$OUT_AUDIT"

echo "✅ DONE"
echo "OUT:   $OUT_XLSM"
echo "AUDIT: $OUT_AUDIT"
