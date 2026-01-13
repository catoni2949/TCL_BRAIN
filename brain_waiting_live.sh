#!/bin/zsh
cd "$(dirname "$0")"
/usr/bin/python3 brain_waiting.py --feed feeds/work_events_email.live.jsonl --days 30
