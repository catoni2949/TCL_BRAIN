#!/usr/bin/env bash
set -euo pipefail

cmd="${1:-help}"
shift || true

case "$cmd" in
  help|-h|--help)
    cat <<'H'
TCL_BRAIN MVP CLI

Usage:
  ./tcl rfi <project_name>
  ./tcl schedule <project_name>
  ./tcl bid-level <project_name>

Outputs go into ./out/<project_name>/
H
    ;;
  rfi)
    project="${1:-demo}"
    out="out/$project"
    mkdir -p "$out"
    cat > "$out/RFI_DRAFT.md" <<'RFI'
# RFI Draft

**Project:** PLACEHOLDER
**Date:** PLACEHOLDER

## Question
(Write the question)

## Background / Reference
(Add drawings/spec references)

## Proposed Options
1.
2.

## Requested Response By
DATE
RFI
    echo "Wrote $out/RFI_DRAFT.md"
    ;;
  schedule)
    project="${1:-demo}"
    out="out/$project"
    mkdir -p "$out"
    cat > "$out/schedule_template.csv" <<'CSV'
activity_id,activity_name,duration_days,predecessors,notes
A100,Notice to Proceed,0,,
A110,Submittals,10,A100,
A120,Procurement,20,A110,
A130,Install,15,A120,
A140,Punch,5,A130,
CSV
    echo "Wrote $out/schedule_template.csv"
    ;;
  bid-level)
    project="${1:-demo}"
    out="out/$project"
    mkdir -p "$out"
    cat > "$out/bid_level_template.csv" <<'CSV'
trade,vendor,base_bid,alt_1,alt_2,exclusions,notes
HVAC,,,,,,
Plumbing,,,,,,
Electrical,,,,,,
CSV
    echo "Wrote $out/bid_level_template.csv"
    ;;
  *)
    echo "Unknown command: $cmd"
    echo "Run: ./tcl --help"
    exit 1
    ;;
esac
