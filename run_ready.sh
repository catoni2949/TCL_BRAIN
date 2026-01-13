#!/usr/bin/env bash
set -euo pipefail

# ---- config defaults (override via env) ----
: "${REPORTS_DIR:=reports}"
: "${READY_LIST:=$REPORTS_DIR/BATCH_READY.txt}"
: "${LOCK:=schema/TCL_SOV_TEMPLATE_GBT.lock.v1.json}"
: "${TPL:?set TPL=/path/to/template.xlsm}"

SELECTOR="./batch_select_plans.py"
RUNNER="./batch_run_ready.sh"

# ---- helpers ----
die() { echo "❌ $*" >&2; exit 1; }

# ---- Gate 1: Preconditions ----
git diff --quiet || die "DIRTY GIT STATE (unstaged) — aborting"
git diff --cached --quiet || die "DIRTY GIT STATE (staged) — aborting"
[ -x "$SELECTOR" ] || die "missing or not executable: $SELECTOR"
[ -x "$RUNNER" ]   || die "missing or not executable: $RUNNER"
[ -f "$LOCK" ]     || die "missing lock: $LOCK"
[ -f "$TPL" ]      || die "missing template: $TPL"
mkdir -p "$REPORTS_DIR"

echo "✅ PRECONDITIONS OK"
echo "TPL=$TPL"
echo "LOCK=$LOCK"

# ---- Gate 2: Select plans (fresh) ----
tmp_list="$(mktemp)"
trap 'rm -f "$tmp_list"' EXIT

# selector writes to READY_LIST; we copy it into tmp_list after dedupe
"$SELECTOR" "$REPORTS_DIR" "$LOCK" "$READY_LIST" >/dev/null

awk 'NF' "$READY_LIST" | sort -u > "$tmp_list"
count="$(wc -l < "$tmp_list" | tr -d ' ')"

if [ "$count" = "0" ]; then
  echo "✅ No ready plans found. Done."
  exit 0
fi

echo "✅ Selected plans: $count"

# ---- Gate 3: Apply (single batch run) ----
# batch_run_ready.sh already logs; it wants args: LIST
export TPL LOCK
"$RUNNER" "$tmp_list"

# ---- Promote GOLD/LATEST ----
gold_src="$(ls -1t "$REPORTS_DIR"/APPLIED_HARDENED_*.xlsm 2>/dev/null | head -1 || true)"
[ -n "$gold_src" ] || die "No APPLIED_HARDENED_*.xlsm found to promote"

cp -f "$gold_src" "$REPORTS_DIR/APPLIED_HARDENED_GOLD.xlsm"
ln -sf "$(basename "$gold_src")" "$REPORTS_DIR/APPLIED_HARDENED_LATEST.xlsm"

echo "✅ GOLD promoted:"
echo "GOLD_SRC=$gold_src"
echo "GOLD=$REPORTS_DIR/APPLIED_HARDENED_GOLD.xlsm"
echo "LATEST=$REPORTS_DIR/APPLIED_HARDENED_LATEST.xlsm"
