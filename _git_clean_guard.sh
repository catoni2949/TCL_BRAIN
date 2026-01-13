#!/usr/bin/env bash
set -euo pipefail

cd "$HOME/TCL_BRAIN"

if ! git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
  echo "❌ NOT A GIT REPO — aborting"
  exit 2
fi

if [[ -n "$(git status --porcelain)" ]]; then
  echo "❌ DIRTY GIT STATE — aborting"
  git status --short
  exit 2
fi

echo "✅ GIT CLEAN"
