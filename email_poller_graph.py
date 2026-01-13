#!/usr/bin/env python3
import os
import json
import time
import uuid
from datetime import datetime, timezone

import msal
import requests

GRAPH = "https://graph.microsoft.com/v1.0"

STATE_FILE = "state/email_poller_graph.state.json"
MSG_MAP_FILE = "state/email_poller_graph.msgid_to_uid.json"

DROP_DIR = "inbox_drop"
MAILBOX_UPN_ENV = "TEST_MAILBOX_UPN"

FOLDERS = ["Inbox", "SentItems"]

# Key in state: when we first started tracking SentItems (UTC Z string).
SENTITEMS_CUTOFF_KEY = "sentitems_cutoff_utc"


def utc_now_z() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def load_json(path: str, default):
    try:
        with open(path, "r") as f:
            return json.load(f)
    except FileNotFoundError:
        return default


def save_json(path: str, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w") as f:
        json.dump(obj, f, indent=2, sort_keys=True)
    os.replace(tmp, path)


def graph_get(url: str, headers: dict, params=None):
    r = requests.get(url, headers=headers, params=params, timeout=30)
    if r.status_code != 200:
        try:
            body = r.json()
        except Exception:
            body = r.text
        raise RuntimeError(f"Graph GET failed: {r.status_code} url={url} body={body}")
    return r.json()


def make_event(msg: dict, mailbox_upn: str, uid: int, folder: str) -> dict:
    frm = msg.get("from", {}).get("emailAddress", {}).get("address", "") or ""
    subj = msg.get("subject", "") or ""
    ts = msg.get("receivedDateTime") or utc_now_z()
    return {
        "ts": ts,
        "sender": frm,
        "subject": subj,
        "project": f"EMAIL/{mailbox_upn}",
        "waiting_status": "neutral",
        "uid": uid,
        "graph_message_id": msg.get("id"),
        "mailbox": mailbox_upn,
        "folder": folder,
    }


def main():
    tenant = os.environ["TCL_TENANT_ID"]
    cid = os.environ["TCL_CLIENT_ID"]
    sec = os.environ["TCL_CLIENT_SECRET"]
    mailbox_upn = os.environ[MAILBOX_UPN_ENV]

    app = msal.ConfidentialClientApplication(
        client_id=cid,
        authority=f"https://login.microsoftonline.com/{tenant}",
        client_credential=sec,
    )
    tok = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in tok:
        raise RuntimeError(f"Token failure: {tok}")

    headers = {"Authorization": "Bearer " + tok["access_token"]}

    st = load_json(STATE_FILE, {})
    msg_map = load_json(MSG_MAP_FILE, {"next_uid": 1, "seen": {}})

    # MIGRATION: old versions used st["delta_link"] for Inbox only
    if "delta_link" in st and "delta_links" not in st:
        st["delta_links"] = {"Inbox": st.get("delta_link")}
        st.pop("delta_link", None)

    if "delta_links" not in st or not isinstance(st["delta_links"], dict):
        st["delta_links"] = {}

    # SentItems cutoff stored once; if missing, initialize as None.
    if SENTITEMS_CUTOFF_KEY not in st:
        st[SENTITEMS_CUTOFF_KEY] = None

    all_new_events = []
    pages_total = 0

    for folder in FOLDERS:
        # Set SentItems cutoff the FIRST time we ever run with SentItems enabled.
        # This prevents history ingestion; only messages AFTER this moment will be emitted.
        if folder == "SentItems" and st.get(SENTITEMS_CUTOFF_KEY) is None:
            st[SENTITEMS_CUTOFF_KEY] = utc_now_z()

        delta_link = st["delta_links"].get(folder)

        if delta_link:
            url = delta_link
            params = None
        else:
            url = f"{GRAPH}/users/{mailbox_upn}/mailFolders('{folder}')/messages/delta"
            params = {"$select": "id,subject,receivedDateTime,from", "$top": "50"}

        while True:
            pages_total += 1
            data = graph_get(url, headers, params=params)

            for msg in data.get("value", []):
                mid = msg.get("id")
                if not mid:
                    continue

                # SentItems cutoff: suppress anything that existed before we started tracking SentItems
                if folder == "SentItems":
                    cutoff = st.get(SENTITEMS_CUTOFF_KEY)
                    received = msg.get("receivedDateTime")
                    if cutoff and received and received <= cutoff:
                        continue

                if mid in msg_map["seen"]:
                    continue

                uid = int(msg_map["next_uid"])
                msg_map["next_uid"] = uid + 1
                msg_map["seen"][mid] = uid
                all_new_events.append(make_event(msg, mailbox_upn, uid, folder))

            next_link = data.get("@odata.nextLink")
            if next_link:
                url = next_link
                params = None
                continue

            new_delta = data.get("@odata.deltaLink")
            if new_delta:
                st["delta_links"][folder] = new_delta
                st["updated_utc"] = utc_now_z()
            break

    save_json(STATE_FILE, st)
    save_json(MSG_MAP_FILE, msg_map)

    if not all_new_events:
        print(f"{utc_now_z()} no new messages (pages={pages_total})")
        return

    os.makedirs(DROP_DIR, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    safe_mailbox = mailbox_upn.replace("@", "_at_").replace("/", "_")
    fname = f"email_poll_{safe_mailbox}_{stamp}_{uuid.uuid4().hex[:8]}.jsonl"
    out_path = os.path.join(DROP_DIR, fname)

    with open(out_path, "w") as f:
        for ev in all_new_events:
            f.write(json.dumps(ev) + "\n")

    print(f"{utc_now_z()} wrote {len(all_new_events)} events -> {DROP_DIR}/{fname}")


if __name__ == "__main__":
    main()
