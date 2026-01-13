#!/usr/bin/env python3
import os, json, ssl, imaplib
from pathlib import Path
from datetime import datetime, timezone

BASE = Path.home() / "TCL_BRAIN"
STATE = BASE / "state" / "imap_state.json"
INBOX_DROP = BASE / "inbox_drop"

IMAP_HOST = os.environ.get("TCL_IMAP_HOST", "").strip()
IMAP_USER = os.environ.get("TCL_IMAP_USER", "").strip()
IMAP_PASS = os.environ.get("TCL_IMAP_PASS", "").strip()
IMAP_FOLDER = os.environ.get("TCL_IMAP_FOLDER", "INBOX").strip()

# State schema (per key):
# {
#   "<host>|<user>|<folder>": {
#       "baseline_uid": <int>,   # set ONCE on first successful run
#       "last_uid": <int>        # advancing watermark
#   },
#   ...
# }
#
# Migration: if existing STATE[key] is an int, treat it as last_uid and
# set baseline_uid = last_uid (preserves "already live" behavior).

def utc_now_z():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

def load_state():
    if STATE.exists():
        try:
            return json.loads(STATE.read_text())
        except Exception:
            return {}
    return {}

def save_state(st):
    STATE.parent.mkdir(parents=True, exist_ok=True)
    STATE.write_text(json.dumps(st, ensure_ascii=False, indent=2))

def parse_headers(raw: bytes):
    sender = ""
    subject = ""
    for line in raw.decode("utf-8", "ignore").splitlines():
        if line.lower().startswith("from:"):
            sender = line.split(":", 1)[1].strip()
        elif line.lower().startswith("subject:"):
            subject = line.split(":", 1)[1].strip()
        if sender and subject:
            break
    return sender, subject

def get_key():
    return f"{IMAP_HOST}|{IMAP_USER}|{IMAP_FOLDER}"

def coerce_entry(entry):
    # Migration: old format stored an int (last_uid)
    if isinstance(entry, int):
        return {"baseline_uid": entry, "last_uid": entry}
    if isinstance(entry, dict):
        b = int(entry.get("baseline_uid", 0) or 0)
        l = int(entry.get("last_uid", 0) or 0)
        # If baseline missing but last exists, set baseline to last (safe)
        if b == 0 and l > 0:
            b = l
        return {"baseline_uid": b, "last_uid": l}
    return {"baseline_uid": 0, "last_uid": 0}

def main():
    if not (IMAP_HOST and IMAP_USER and IMAP_PASS):
        raise SystemExit("Missing IMAP env vars")

    INBOX_DROP.mkdir(parents=True, exist_ok=True)

    state = load_state()
    key = get_key()
    entry = coerce_entry(state.get(key, {}))

    baseline_uid = int(entry.get("baseline_uid", 0) or 0)
    last_uid = int(entry.get("last_uid", 0) or 0)

    ctx = ssl.create_default_context()
    M = imaplib.IMAP4_SSL(IMAP_HOST, ssl_context=ctx)
    M.login(IMAP_USER, IMAP_PASS)
    M.select(IMAP_FOLDER)

    # Search from last_uid+1
    typ, data = M.uid("SEARCH", None, "UID %d:*" % (last_uid + 1))
    if typ != "OK":
        M.logout()
        print(f"{utc_now_z()} search failed")
        return

    uids = [int(x) for x in data[0].split()] if data and data[0] else []
    if not uids:
        M.logout()
        print(f"{utc_now_z()} no new mail")
        return

    max_uid = max(uids)

    # First successful run cutoff:
    # If baseline_uid is not set, we "start live" at the newest UID and do NOT emit history.
    if baseline_uid == 0:
        baseline_uid = max_uid
        last_uid = max_uid
        state[key] = {"baseline_uid": baseline_uid, "last_uid": last_uid}
        save_state(state)
        M.logout()
        print(f"{utc_now_z()} initialized baseline_uid={baseline_uid} (no backfill)")
        return

    events = []
    new_uids = [u for u in uids if u > last_uid]
    # Limit per run
    new_uids = new_uids[-200:]

    for uid in new_uids:
        typ, msgdata = M.uid("FETCH", str(uid), "(RFC822.HEADER)")
        if typ != "OK":
            continue

        raw = b"".join(p[1] for p in msgdata if isinstance(p, tuple))
        sender, subject = parse_headers(raw)

        events.append({
            "ts": utc_now_z(),
            "sender": sender,
            "subject": subject,
            "project": "",
            "waiting_status": None
        })

        if uid > last_uid:
            last_uid = uid

    M.logout()

    if events:
        stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
        out = INBOX_DROP / f"imap_poll_{stamp}.jsonl"
        with out.open("w", encoding="utf-8") as f:
            for e in events:
                f.write(json.dumps(e, ensure_ascii=False) + "\n")

        state[key] = {"baseline_uid": baseline_uid, "last_uid": last_uid}
        save_state(state)
        print(f"{utc_now_z()} wrote {len(events)} events -> {out.name}")
    else:
        # Still advance the watermark if we saw higher UID but didn’t parse
        if max_uid > last_uid:
            last_uid = max_uid
            state[key] = {"baseline_uid": baseline_uid, "last_uid": last_uid}
            save_state(state)
        print(f"{utc_now_z()} no parsable events")

if __name__ == "__main__":
    main()