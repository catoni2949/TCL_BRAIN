#!/bin/zsh

echo "TCL_BRAIN STATUS"
date -u "+UTC now: %Y-%m-%dT%H:%M:%SZ"
echo ""

echo "LAUNCHD JOBS:"
dw="$(launchctl list 2>/dev/null | awk '$3=="com.tcl.dropwatcher"{print "LOADED pid=" $1; found=1} END{if(!found) print "NOT LOADED"}')"
wr="$(launchctl list 2>/dev/null | awk '$3=="com.tcl.waitingreport"{print "LOADED pid=" $1; found=1} END{if(!found) print "NOT LOADED"}')"
echo " - dropwatcher: $dw"
echo " - waitingreport: $wr"
echo ""

echo "DROP WATCHER (last 5 log lines):"
if [ -f "$HOME/TCL_BRAIN/logs/drop_watcher.log" ]; then
  tail -n 5 "$HOME/TCL_BRAIN/logs/drop_watcher.log"
else
  echo "NO drop_watcher.log"
fi
echo ""

echo "WAITING REPORT (header):"
if [ -f "$HOME/TCL_BRAIN/reports/waiting_report.txt" ]; then
  head -n 5 "$HOME/TCL_BRAIN/reports/waiting_report.txt"
else
  echo "NO waiting_report.txt"
fi
