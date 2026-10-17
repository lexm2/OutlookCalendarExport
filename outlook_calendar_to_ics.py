#!/usr/bin/env python3
"""
Export every Outlook calendar on your account to .ics files by reading the
events directly through Microsoft Graph. This skips Outlook's "share/publish"
feature entirely.

Works with personal (outlook.com / hotmail / live) and work/school (Microsoft 365)
accounts. Standard library only, so there's nothing to install.

Usage:
    python3 outlook_calendar_to_ics.py                      # all calendars, 2000 -> 2036
    python3 outlook_calendar_to_ics.py --start 2015-01-01 --end 2030-01-01
    python3 outlook_calendar_to_ics.py --out ~/cal_export

You'll get a code to type into https://microsoft.com/devicelogin. Sign in there
and the script continues by itself.

Output: one .ics file per calendar, plus all_calendars.ics with everything combined.

Note: recurring events are expanded into individual occurrences (via calendarView),
so every moved, edited or cancelled instance comes out exactly as Outlook shows it.
"""

import argparse
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone

# Public client ID of Microsoft's own "Microsoft Graph Command Line Tools" app.
# It supports device-code sign-in for personal and work accounts. If your work
# tenant blocks it, register your own app in Entra ID (public client, delegated
# Calendars.Read permission) and pass its ID with --client-id.
DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
SCOPES = "Calendars.Read Calendars.Read.Shared offline_access"
GRAPH = "https://graph.microsoft.com/v1.0"

SELECT_FIELDS = ",".join([
    "id", "iCalUId", "subject", "body", "start", "end", "isAllDay", "location",
    "locations", "organizer", "attendees", "categories", "showAs", "sensitivity",
    "isCancelled", "importance", "createdDateTime", "lastModifiedDateTime",
    "onlineMeeting", "onlineMeetingUrl", "webLink", "seriesMasterId", "type",
    "isReminderOn", "reminderMinutesBeforeStart",
])


# --------------------------------------------------------------------------- auth

def post_form(url, data):
    req = urllib.request.Request(
        url, data=urllib.parse.urlencode(data).encode(), method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    try:
        with urllib.request.urlopen(req) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        return json.load(e)


def get_token(client_id, tenant):
    base = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0"
    dc = post_form(f"{base}/devicecode", {"client_id": client_id, "scope": SCOPES})
    if "device_code" not in dc:
        sys.exit(f"Could not start sign-in: {dc.get('error_description', dc)}")

    print("\n" + dc["message"] + "\n", flush=True)

    interval = int(dc.get("interval", 5))
    deadline = time.time() + int(dc.get("expires_in", 900))
    while time.time() < deadline:
        time.sleep(interval)
        tok = post_form(f"{base}/token", {
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": client_id,
            "device_code": dc["device_code"],
        })
        if "access_token" in tok:
            print("Signed in.\n")
            return tok["access_token"]
        err = tok.get("error")
        if err == "authorization_pending":
            continue
        if err == "slow_down":
            interval += 5
            continue
        sys.exit(f"Sign-in failed: {tok.get('error_description', tok)}")
    sys.exit("Sign-in timed out.")


# -------------------------------------------------------------------------- graph

def graph_get(url, token):
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        # Times in UTC, bodies as plain text, big pages.
        "Prefer": 'outlook.timezone="UTC", outlook.body-content-type="text", odata.maxpagesize=500',
    }
    for attempt in range(8):
        req = urllib.request.Request(url, headers=headers)
        try:
            with urllib.request.urlopen(req) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            if e.code in (429, 503, 504):
                wait = int(e.headers.get("Retry-After", 2 ** attempt))
                print(f"  throttled, waiting {wait}s...", flush=True)
                time.sleep(wait)
                continue
            body = e.read().decode(errors="replace")
            raise RuntimeError(f"HTTP {e.code} for {url}\n{body}") from None
    raise RuntimeError(f"Gave up after repeated throttling: {url}")


def graph_all(url, token):
    while url:
        page = graph_get(url, token)
        yield from page.get("value", [])
        url = page.get("@odata.nextLink")


# ---------------------------------------------------------------------------- ics

def esc(text):
    if text is None:
        return ""
    return (str(text).replace("\\", "\\\\").replace(";", "\\;")
            .replace(",", "\\,").replace("\r\n", "\\n").replace("\n", "\\n")
            .replace("\r", ""))


def fold(line):
    """Fold to 75 octets per RFC 5545 without splitting UTF-8 characters."""
    out, cur, size = [], "", 0
    for ch in line:
        n = len(ch.encode("utf-8"))
        limit = 75 if not out else 74  # continuation lines start with a space
        if size + n > limit:
            out.append(cur)
            cur, size = "", 0
        cur += ch
        size += n
    out.append(cur)
    return "\r\n ".join(out)


def utc_stamp(graph_dt):
    """'2026-10-17T14:00:00.0000000' (UTC) -> '20261017T140000Z'"""
    return datetime.strptime(graph_dt[:19], "%Y-%m-%dT%H:%M:%S").strftime("%Y%m%dT%H%M%SZ")


def iso_stamp(iso):
    """'2026-10-17T14:00:00.123Z' -> '20261017T140000Z'"""
    return utc_stamp(iso.replace("Z", "")[:19])


def date_only(graph_dt):
    return graph_dt[:10].replace("-", "")


def param_quote(v):
    v = (v or "").replace('"', "'")
    return f'"{v}"' if re.search(r'[:;,]', v) else v


PARTSTAT = {
    "accepted": "ACCEPTED", "declined": "DECLINED",
    "tentativelyAccepted": "TENTATIVE", "none": "NEEDS-ACTION",
    "notResponded": "NEEDS-ACTION", "organizer": "ACCEPTED",
}
ROLE = {"required": "REQ-PARTICIPANT", "optional": "OPT-PARTICIPANT", "resource": "NON-PARTICIPANT"}
CLASS = {"normal": "PUBLIC", "personal": "PRIVATE", "private": "PRIVATE", "confidential": "CONFIDENTIAL"}
PRIORITY = {"high": "1", "normal": "5", "low": "9"}
BUSY = {"free": "FREE", "tentative": "TENTATIVE", "busy": "BUSY", "oof": "OOF",
        "workingElsewhere": "WORKINGELSEWHERE"}


def event_to_vevent(ev):
    L = ["BEGIN:VEVENT"]
    uid = ev.get("iCalUId") or ev["id"]
    L.append(f"UID:{esc(uid)}")
    L.append(f"DTSTAMP:{datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')}")

    if ev.get("isAllDay"):
        L.append(f"DTSTART;VALUE=DATE:{date_only(ev['start']['dateTime'])}")
        L.append(f"DTEND;VALUE=DATE:{date_only(ev['end']['dateTime'])}")
    else:
        L.append(f"DTSTART:{utc_stamp(ev['start']['dateTime'])}")
        L.append(f"DTEND:{utc_stamp(ev['end']['dateTime'])}")

    L.append(f"SUMMARY:{esc(ev.get('subject') or '(No title)')}")

    loc = (ev.get("location") or {}).get("displayName")
    if not loc and ev.get("locations"):
        loc = "; ".join(l.get("displayName", "") for l in ev["locations"] if l.get("displayName"))
    if loc:
        L.append(f"LOCATION:{esc(loc)}")

    desc = ((ev.get("body") or {}).get("content") or "").strip()
    join_url = (ev.get("onlineMeeting") or {}).get("joinUrl") or ev.get("onlineMeetingUrl")
    if join_url:
        L.append(f"URL:{esc(join_url)}")
        if join_url not in desc:
            desc = (desc + "\n\nJoin: " + join_url).strip()
    if desc:
        L.append(f"DESCRIPTION:{esc(desc)}")

    org = (ev.get("organizer") or {}).get("emailAddress") or {}
    if org.get("address"):
        L.append(f"ORGANIZER;CN={param_quote(org.get('name'))}:mailto:{org['address']}")

    for a in ev.get("attendees") or []:
        em = a.get("emailAddress") or {}
        if not em.get("address"):
            continue
        status = PARTSTAT.get(((a.get("status") or {}).get("response")) or "none", "NEEDS-ACTION")
        role = ROLE.get(a.get("type"), "REQ-PARTICIPANT")
        L.append(f"ATTENDEE;CN={param_quote(em.get('name'))};ROLE={role};PARTSTAT={status}"
                 f":mailto:{em['address']}")

    if ev.get("categories"):
        L.append("CATEGORIES:" + ",".join(esc(c) for c in ev["categories"]))

    show = ev.get("showAs") or "busy"
    L.append(f"TRANSP:{'TRANSPARENT' if show == 'free' else 'OPAQUE'}")
    L.append(f"X-MICROSOFT-CDO-BUSYSTATUS:{BUSY.get(show, 'BUSY')}")
    L.append(f"CLASS:{CLASS.get(ev.get('sensitivity'), 'PUBLIC')}")
    L.append(f"PRIORITY:{PRIORITY.get(ev.get('importance'), '5')}")
    L.append(f"STATUS:{'CANCELLED' if ev.get('isCancelled') else 'CONFIRMED'}")

    if ev.get("createdDateTime"):
        L.append(f"CREATED:{iso_stamp(ev['createdDateTime'])}")
    if ev.get("lastModifiedDateTime"):
        L.append(f"LAST-MODIFIED:{iso_stamp(ev['lastModifiedDateTime'])}")

    if ev.get("isReminderOn") and ev.get("reminderMinutesBeforeStart") is not None:
        L += ["BEGIN:VALARM", "ACTION:DISPLAY", "DESCRIPTION:Reminder",
              f"TRIGGER:-PT{int(ev['reminderMinutesBeforeStart'])}M", "END:VALARM"]

    L.append("END:VEVENT")
    return L


def write_ics(path, cal_name, events):
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0",
             "PRODID:-//outlook_calendar_to_ics//EN", "CALSCALE:GREGORIAN",
             "METHOD:PUBLISH", f"X-WR-CALNAME:{esc(cal_name)}"]
    for ev in events:
        lines += event_to_vevent(ev)
    lines.append("END:VCALENDAR")
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write("\r\n".join(fold(l) for l in lines) + "\r\n")


def safe_name(s):
    return re.sub(r'[^\w\-. ]+', "_", s).strip() or "calendar"


# --------------------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(description="Export all Outlook calendars to .ics")
    ap.add_argument("--start", default="2000-01-01", help="range start, YYYY-MM-DD (default 2000-01-01)")
    ap.add_argument("--end", default="2036-01-01", help="range end, YYYY-MM-DD (default 2036-01-01)")
    ap.add_argument("--out", default="outlook_export", help="output folder (default ./outlook_export)")
    ap.add_argument("--client-id", default=DEFAULT_CLIENT_ID)
    ap.add_argument("--tenant", default="common", help="'common', 'consumers', 'organizations' or a tenant ID")
    args = ap.parse_args()

    out = os.path.expanduser(args.out)
    os.makedirs(out, exist_ok=True)

    token = get_token(args.client_id, args.tenant)

    calendars = list(graph_all(f"{GRAPH}/me/calendars?$select=id,name,owner", token))
    print(f"Found {len(calendars)} calendar(s).")

    start = f"{args.start}T00:00:00Z"
    end = f"{args.end}T00:00:00Z"
    combined, seen = [], set()

    for cal in calendars:
        name = cal.get("name") or "calendar"
        print(f"- {name} ...", end=" ", flush=True)
        q = urllib.parse.urlencode({
            "startDateTime": start, "endDateTime": end,
            "$select": SELECT_FIELDS, "$orderby": "start/dateTime",
        })
        try:
            events = list(graph_all(f"{GRAPH}/me/calendars/{cal['id']}/calendarView?{q}", token))
        except RuntimeError as e:
            print(f"skipped ({str(e).splitlines()[0]})")
            continue

        path = os.path.join(out, safe_name(name) + ".ics")
        write_ics(path, name, events)
        print(f"{len(events)} events -> {path}")

        for ev in events:
            key = ev.get("iCalUId") or ev["id"]
            if key not in seen:
                seen.add(key)
                combined.append(ev)

    combined.sort(key=lambda e: e["start"]["dateTime"])
    all_path = os.path.join(out, "all_calendars.ics")
    write_ics(all_path, "All Outlook calendars", combined)
    print(f"\nCombined: {len(combined)} events -> {all_path}")


if __name__ == "__main__":
    main()
