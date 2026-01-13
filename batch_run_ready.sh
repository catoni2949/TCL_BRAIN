#!/usr/bin/env bash
set -euo pipefail

LIST="${1:-reports/BATCH_READY.txt}"
: "${TPL:?set TPL=/path/to/template.xlsm}"
: "${LOCK:?set LOCK=schema/...lock.json}"

ts="$(date +%Y%m%d_%H%M%S)"
LOG="reports/BATCH_RUN_${ts}.log"
mkdir -p reports

echo "LOG: $LOG" | tee -a "$LOG"
echo "LIST: $LIST" | tee -a "$LOG"

# De-dupe list defensively (blank lines removed)
tmp="$(mktemp)"
awk 'NF' "$LIST" | sort -u > "$tmp"

i=0
while IFS= read -r PLAN; do
  [ -z "$PLAN" ] && continue
  i=$((i+1))
  {
    echo "-----"
    echo "PLAN: $PLAN"
    PLAN="$PLAN" TPL="$TPL" LOCK="$LOCK" ./run_apply.sh
  } 2>&1 | tee -a "$LOG"
done < "$tmp"

rm -f "$tmp"

echo "-----" | tee -a "$LOG"
echo "DONE plans=$i" | tee -a "$LOG"
echo "OK_count=$(grep -c 'Writes applied:' "$LOG" || true)" | tee -a "$LOG"
echo "FAIL_count=$(grep -cE 'NO-GO|abort|mismatch|FORMULA COLUMN TARGET' "$LOG" || true)" | tee -a "$LOG"
