#!/usr/bin/env bash
set -euo pipefail

log(){ echo "$@" >&2; }
die(){ log "$@"; exit 2; }

# Ensure we run from repo root so relative paths work
cd "$(dirname "$0")"

: "${PLAN:?set PLAN=/path/to/plan.json}"
: "${TPL:?set TPL=/path/to/template.xlsm}"
: "${LOCK:?set LOCK=/path/to/lock.json}"

OUT_TS="20 20 12 61 79 80 81 398 701 33 98 100 204 250 395 399 400date +%Y%m%d_%H%M%S_%N)_80693"
OUT_XLSM="reports/APPLIED_HARDENED_${OUT_TS}.xlsm"
OUT_AUDIT="reports/APPLIED_HARDENED_${OUT_TS}.audit.jsonl"

# Everything below MUST NOT leak to stdout if you use eval/source
{
  ./_hardened_guard.sh
  ./_git_clean_guard.sh

  TCL_HARDENED_APPLY=1 ./apply_write_plan.py \
    --plan "$PLAN" \
    --template "$TPL" \
    --lock "$LOCK" \
    --out-xlsm "$OUT_XLSM" \
    --audit-jsonl "$OUT_AUDIT"

  log "✅ DONE"
  log "OUT:   $OUT_XLSM"
  log "AUDIT: $OUT_AUDIT"
} >&2

# stdout ONLY (so eval/source won't choke)
printf 'export OUT_XLSM=%q\n' "$OUT_XLSM"
printf 'export OUT_AUDIT=%q\n' "$OUT_AUDIT"
