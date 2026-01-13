#!/usr/bin/env bash
set -euo pipefail

FILES=(
  "$HOME/TCL_BRAIN/apply_write_plan.py"
  "$HOME/TCL_BRAIN/build_sov_write_plan_from_amounts.py"
  "$HOME/TCL_BRAIN/auto_resolve_sov_cells.py"
)

for f in "${FILES[@]}"; do
  perms=$(stat -f "%Lp" "$f")
  if [[ "$perms" != "555" ]]; then
    echo "❌ HARDENING VIOLATION: $f permissions=$perms (expected 555)"
    exit 2
  fi
done

echo "✅ HARDENED GUARD PASSED"
