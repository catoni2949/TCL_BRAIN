#!/bin/zsh
set -euo pipefail

cd /Users/tclserver/TCL_BRAIN

# Load secrets for THIS run
source /Users/tclserver/TCL_BRAIN/secrets.env

# Run poller
/usr/bin/python3 /Users/tclserver/TCL_BRAIN/email_poller_graph.py
