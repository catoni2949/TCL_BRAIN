#!/usr/bin/env python3
import json
import re
from pathlib import Path
from datetime import datetime, timezone, timedelta

BASE = Path.home() / "TCL_BRAIN"
FEED = BASE / "feeds" / "work_events_email.live.jsonl"
REPORTS_DIR = BASE / "reports"
REPORT_FILE = REPORTS_DIR / "waiting_report.txt"
ROUTES_FILE = BASE / "state" / "email_routes.json"

STATUSES = {"neutral", "waiting_on_them", "waiting_on_us"}


def utc_now():
    return datetime.now(timezone.utc)


def utc_now_z():
    return utc_now().strftime("%Y-%m-%dT%H:%M:%SZ")


def parse_ts(s):
    if not s or not isinstance(s, str):
        return None
    try:
        # Accept Z
        if s.endswith("Z"):
            return datetime.fromisoformat(s.replace("Z", "+00:00"))
        return datetime.fromisoformat(s)
    except Exception:
        return None


def norm(s):
    if s is None:
        return ""
    return str(s).strip()


def load_json(path, default):
    try:
        if Path(path).exists():
            return json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        pass
    return default


def load_routes():
    routes = load_json(ROUTES_FILE, {})
    if not isinstance(routes, dict):
        routes = {}

    routes.setdefault("version", 1)
    routes.setdefault("default_email_bucket", "EMAIL/Unrouted")

    # ignore: list of strings/regexes (optional)
    routes.setdefault("ignore", [])

    # rules: list of dicts
    # Supported keys:
    #   - bucket or project: destination group name
    #   - ignore: true/false
    #   - sender_contains / subject_contains
    #   - sender_regex / subject_regex
    #   - mailbox (optional exact match)
    routes.setdefault("rules", [])
    if not isinstance(routes["rules"], list):
        routes["rules"] = []
    if not isinstance(routes["ignore"], list):
        routes["ignore"] = []

    return routes


def match_rule(e, rule):
    sender = norm(e.get("sender", ""))
    subject = norm(e.get("subject", ""))
    mailbox = norm(e.get("mailbox", ""))

    want_mailbox = norm(rule.get("mailbox", ""))
    if want_mailbox and want_mailbox.lower() != mailbox.lower():
        return False

    sc = norm(rule.get("sender_contains", ""))
    if sc and sc.lower() not in sender.lower():
        return False

    subc = norm(rule.get("subject_contains", ""))
    if subc and subc.lower() not in subject.lower():
        return False

    sr = norm(rule.get("sender_regex", ""))
    if sr:
        try:
            if not re.search(sr, sender, re.IGNORECASE):
                return False
        except re.error:
            return False

    sur = norm(rule.get("subject_regex", ""))
    if sur:
        try:
            if not re.search(sur, subject, re.IGNORECASE):
                return False
        except re.error:
            return False

    # If rule provided none of the match keys, it's not a match rule.
    if not any([sc, subc, sr, sur, want_mailbox]):
        return False

    return True


def route_email_event(e, routes):
    """
    Returns: (bucket, ignored)
      - bucket: string bucket name (e.g. 'St Francis Medical – Floor 2 TI' or 'EMAIL/Unrouted')
      - ignored: True if the event should be skipped entirely
    """
    default_bucket = routes.get("default_email_bucket", "EMAIL/Unrouted")

    sender = norm(e.get("sender", ""))
    subject = norm(e.get("subject", ""))

    # Ignore rules (match against sender OR subject)
    for ig in routes.get("ignore", []) or []:
        ig_n = norm(ig)
        if ig_n and (ig_n in sender or ig_n in subject):
            return (default_bucket, True)

    # Routing rules
    for r in routes.get("rules", []) or []:
        if not isinstance(r, dict):
            continue

        bucket = norm(r.get("bucket")) or norm(r.get("project")) or default_bucket

        fc = norm(r.get("from_contains", ""))
        if fc and fc in sender:
            return (bucket, False)

        c = norm(r.get("contains", ""))
        if c and c in subject:
            return (bucket, False)

    return (default_bucket, False)

    return (norm(routes.get("default_email_bucket", "EMAIL/Unrouted")), False)


def effective_project(e, routes):
    """
    Decide final grouping bucket for this event.
    Priority:
      1) If event has a non-empty project, keep it (even EMAIL/...).
      2) Else if mailbox exists -> EMAIL/{mailbox}
      3) Else apply routes default bucket.
    Then, if it's an email bucket (blank or starts EMAIL/), apply routing rules (may map to real project).
    """
    proj = norm(e.get("project", ""))
    mailbox = norm(e.get("mailbox", ""))

    if not proj:
        if mailbox:
            proj = f"EMAIL/{mailbox}"
        else:
            proj = norm(routes.get("default_email_bucket", "EMAIL/Unrouted"))

    # Apply routing rules only for email/blank projects
    if (not proj) or proj.startswith("EMAIL/"):
        routed, ignored = route_email_event(e, routes)
        if ignored:
            return ("", True)
        if routed:
            proj = routed

    return (proj, False)


def read_events(path):
    if not path.exists():
        return []

    out = []
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                e = json.loads(line)
                if isinstance(e, dict):
                    out.append(e)
            except Exception:
                continue
    return out


def main():
    routes = load_routes()
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)

    events = read_events(FEED)

    now = utc_now()
    cutoff = now - timedelta(days=30)

    bucketed = []
    kept = 0

    for e in events:
        ts_raw = e.get("ts")
        dt = parse_ts(ts_raw)
        if not dt:
            continue
        if dt < cutoff:
            continue

        status = norm(e.get("waiting_status", "neutral")) or "neutral"
        if status not in STATUSES:
            continue

        proj, ignored = effective_project(e, routes)
        if ignored:
            continue

        group = proj or "(no project)"
        sender = norm(e.get("sender", ""))
        subject = norm(e.get("subject", ""))

        bucketed.append((group, dt, sender, subject, status))
        kept += 1

    # Sort newest first globally
    bucketed.sort(key=lambda x: x[1], reverse=True)

    # Group
    groups = {}
    for g, dt, sender, subject, status in bucketed:
        groups.setdefault(g, []).append((dt, sender, subject, status))

    # Order groups:
    #   1) non-email projects (not starting EMAIL/)
    #   2) EMAIL/* buckets
    #   3) (no project)
    def group_key(name):
        if name == "(no project)":
            return (2, name.lower())
        if name.startswith("EMAIL/"):
            return (1, name.lower())
        return (0, name.lower())

    ordered_group_names = sorted(groups.keys(), key=group_key)

    lines = []
    lines.append("TCL_BRAIN waiting report")
    lines.append(f"Generated (UTC): {utc_now_z()}")
    lines.append(f"Window: last 30 days (cutoff={cutoff.isoformat()})")
    lines.append("Statuses: neutral, waiting_on_them, waiting_on_us")
    lines.append(f"Events: {kept}")
    lines.append("")

    for name in ordered_group_names:
        rows = groups[name]
        lines.append(f"=== {name} ({len(rows)}) ===")
        for dt, sender, subject, status in rows:
            # show Z timestamps
            ts = dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            lines.append(f"- {ts} | {sender} | {subject} | {status}")
        lines.append("")

    REPORT_FILE.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")
    print(f"Wrote: {REPORT_FILE} ({len(lines)} lines)")


if __name__ == "__main__":
    main()
