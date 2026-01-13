import json
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path

FEEDS = [
    "feeds/work_events_email.history.jsonl",
    "feeds/work_events_email.live.jsonl",
]

DAYS = 30
STATUSES = {"waiting_on_us", "waiting_on_them", "neutral"}

cutoff = datetime.now(timezone.utc) - timedelta(days=DAYS)

events = []
missing = []

for feed in FEEDS:
    p = Path(feed)
    if not p.exists():
        missing.append(feed)
        continue
    with p.open("r") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            evt = json.loads(line)
            status = evt.get("waiting_status")
            if status not in STATUSES:
                continue
            ts = datetime.fromisoformat(evt["ts"].replace("Z", "+00:00"))
            if ts >= cutoff:
                events.append(evt)

by_project = defaultdict(list)
for e in events:
    key = e["project"].strip() if e.get("project") else ""
    if not key:
        key = "(no project)"
    by_project[key].append(e)

print("Feeds:")
for f in FEEDS:
    print(f"- {f}")
if missing:
    print("Missing feeds:")
    for f in missing:
        print(f"- {f}")

print(f"\nWindow: last {DAYS} days (cutoff={cutoff.isoformat()})")
print(f"Statuses: {', '.join(sorted(STATUSES))}")
print(f"Events: {len(events)}")

if not by_project:
    print("No matching events in window.")
else:
    for project, items in sorted(by_project.items(), key=lambda x: len(x[1]), reverse=True):
        print(f"\n=== {project} ({len(items)}) ===")
        for e in sorted(items, key=lambda x: x["ts"], reverse=True)[:8]:
            print(f"- {e['ts']} | {e['sender']} | {e['subject']} | {e.get('waiting_status')}")
