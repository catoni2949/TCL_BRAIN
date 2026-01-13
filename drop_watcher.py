import json
import time
from pathlib import Path
from datetime import datetime, timezone

INBOX = Path.home() / "TCL_BRAIN" / "inbox_drop"
PROCESSED = Path.home() / "TCL_BRAIN" / "archive" / "drop_processed"
LIVE = Path.home() / "TCL_BRAIN" / "feeds" / "work_events_email.live.jsonl"
LOG = Path.home() / "TCL_BRAIN" / "logs" / "drop_watcher.log"

REQUIRED_KEYS = {"ts", "sender", "subject", "project", "waiting_status"}

def log(msg: str) -> None:
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    line = f"{ts} {msg}"
    print(line)
    with LOG.open("a") as f:
        f.write(line + "\n")

    try:
        print(line, flush=True)
    except Exception:
        print(msg, flush=True)
def normalize_events(text: str):
    text = text.strip()
    if not text:
        return []
    # If it's a JSON array, explode it
    if text.startswith("["):
        data = json.loads(text)
        if isinstance(data, list):
            return data
        return []
    # Otherwise treat as JSONL (one JSON object per line)
    out = []
    for ln in text.splitlines():
        ln = ln.strip()
        if not ln:
            continue
        out.append(json.loads(ln))
    return out

def validate_event(evt: dict) -> bool:
    if not isinstance(evt, dict):
        return False
    missing = REQUIRED_KEYS - set(evt.keys())
    if missing:
        return False
    # Minimal sanity: ts must be string, sender/subject strings
    return isinstance(evt["ts"], str) and isinstance(evt["sender"], str) and isinstance(evt["subject"], str)

def main():
    log(f"Watcher starting. INBOX={INBOX}")
    INBOX.mkdir(parents=True, exist_ok=True)
    PROCESSED.mkdir(parents=True, exist_ok=True)
    LIVE.parent.mkdir(parents=True, exist_ok=True)

    while True:
        files = sorted([p for p in INBOX.iterdir() if p.is_file() and not p.name.startswith(".")])
        if not files:
            time.sleep(1)
            continue

        for p in files:
            try:
                raw = p.read_text(errors="replace")
                events = normalize_events(raw)
                good = []
                bad = 0
                for e in events:
                    if validate_event(e):
                        good.append(e)
                    else:
                        bad += 1

                if good:
                    with LIVE.open("a") as f:
                        for e in good:
                            f.write(json.dumps(e, ensure_ascii=False) + "\n")

                stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
                dest = PROCESSED / f"{p.name}.{stamp}.done"
                p.rename(dest)

                log(f"Processed {p.name}: appended={len(good)} bad={bad} -> {dest.name}")

            except Exception as ex:
                stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
                dest = PROCESSED / f"{p.name}.{stamp}.ERROR"
                try:
                    p.rename(dest)
                except Exception:
                    pass
                log(f"ERROR processing {p.name}: {ex}")

if __name__ == "__main__":
    main()
