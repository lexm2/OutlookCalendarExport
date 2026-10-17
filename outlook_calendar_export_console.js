// Outlook on the web -> .ics exporter (paste into the browser DevTools console)
//
// 1. Open Outlook on the web (outlook.live.com, outlook.office.com,
//    outlook.office365.com or outlook.cloud.microsoft), signed in as usual.
//    Any Outlook page works (Mail, Calendar, Settings...): the script loads the
//    calendar itself in a hidden frame.
// 2. Press F12 -> Console. Paste this whole file and press Enter.
//    (Chrome may ask you to type "allow pasting" first.)
//    If you're on some other site, it takes you to Outlook's calendar instead;
//    paste it again there (browsers don't let a script survive leaving a site).
// 3. Leave the tab open while it runs. Progress is logged here.
// 4. Downloads one .ics per calendar plus all_calendars.ics.
//
// Uses the same internal API Outlook's own page uses (owa/service.svc).
//
// outlook_calendar_export_console.min.js is a minified copy for pasting;
// rebuild it with ./build.sh after editing this file.

(async () => {
  // Date range to export. Recurring events are expanded within it.
  const START_YEAR = 2000;
  const END_YEAR = 2035; // inclusive

  const log = (...a) => console.log("%c[ICS export]", "color:#0a7;font-weight:bold", ...a);
  const warn = (...a) => console.warn("[ICS export]", ...a);
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  // Run `fn` over `items` with at most `n` in flight. Slow connections are
  // dominated by latency, so a few parallel requests make a big difference.
  async function pool(items, n, fn) {
    const out = new Array(items.length);
    let next = 0;
    await Promise.all(Array.from({ length: Math.min(n, items.length) }, async () => {
      while (next < items.length) { const k = next++; out[k] = await fn(items[k], k); }
    }));
    return out;
  }

  // ------------------------------------------------------------- right site?
  // The API only answers same-origin requests, so this must run on Outlook.
  if (!/^outlook\.(live\.com|office\.com|office365\.com|cloud\.microsoft)$/i.test(location.hostname)) {
    const personal = confirm(
      "This isn't Outlook on the web, so the export can't run here.\n\n" +
      "OK = personal account (outlook.com / hotmail / live)\nCancel = work or school account\n\n" +
      "You'll be taken to your Outlook calendar. Once it has loaded, paste the script into the console again."
    );
    location.href = personal ? "https://outlook.live.com/calendar/" : "https://outlook.office.com/calendar/";
    return;
  }

  // ------------------------------------------------------------ borrow Outlook's auth
  // Outlook signs each request with a short-lived token, and after start-up the
  // page itself rarely makes requests (a background worker does). So: load
  // Outlook in a hidden same-origin iframe, hook its fetch before it boots, and
  // copy the headers of the first authenticated service.svc request it makes.
  let auth = null; // { headers, base, at }
  const origFetch = window.fetch;

  async function refreshAuth() {
    const frame = document.createElement("iframe");
    frame.style.cssText = "position:fixed;left:-10000px;top:0;width:1200px;height:800px;border:0";
    let got = null;
    const hook = (win) => {
      if (!win || win.__icsHooked) return;
      win.__icsHooked = true;
      const f = win.fetch;
      win.fetch = function (input, init) {
        try {
          const url = String(typeof input === "string" ? input : input && input.url);
          const h = {};
          new win.Headers((init && init.headers) || (input instanceof win.Request ? input.headers : undefined))
            .forEach((v, k) => (h[k] = v));
          // Skip /owa/published/service.svc, which only serves anonymous lookups.
          if (!got && /\/owa\/service\.svc/i.test(url) && h.authorization) {
            got = { headers: h, base: new URL(url, location.href).href.split("?")[0], at: Date.now() };
          }
        } catch (e) {}
        return f.apply(this, arguments);
      };
    };
    document.body.appendChild(frame);
    // Always the calendar, whichever Outlook page we were started from.
    frame.src = location.origin + "/calendar/";
    // The iframe's window object survives its first same-origin navigation, so
    // hooking it as early as possible catches Outlook's start-up requests.
    const iv = setInterval(() => { try { hook(frame.contentWindow); } catch (e) {} }, 1);
    const t0 = Date.now();
    let signedOut = false, lastNote = t0;
    // Up to 5 minutes: on a slow connection Outlook can take a while to boot.
    while (!got && Date.now() - t0 < 5 * 60 * 1000) {
      await sleep(100);
      try { void frame.contentWindow.location.href; } catch (e) { signedOut = true; break; } // bounced to a login site
      if (Date.now() - lastNote >= 15000) {
        lastNote = Date.now();
        log(`  still waiting for Outlook to load (${Math.round((lastNote - t0) / 1000)}s)...`);
      }
    }
    clearInterval(iv);
    frame.remove();
    if (!got) {
      // No point carrying on with other calendars: every request needs this.
      const e = new Error(signedOut
        ? "Outlook sent the hidden frame to a sign-in page. Make sure you're signed in, reload, and try again."
        : "Couldn't capture Outlook's login (no authenticated request within 5 minutes). Reload the page and try again.");
      e.fatal = true;
      throw e;
    }
    auth = got;
  }

  // Share one refresh between parallel requests.
  let refreshing = null;
  const ensureAuth = () => refreshing || (refreshing = refreshAuth().finally(() => (refreshing = null)));

  const HEADER = {
    __type: "JsonRequestHeaders:#Exchange",
    RequestServerVersion: "V2018_01_08",
    TimeZoneContext: {
      __type: "TimeZoneContext:#Exchange",
      TimeZoneDefinition: { __type: "TimeZoneDefinitionType:#Exchange", Id: "UTC" },
    },
  };

  async function svc(action, body) {
    let lastErr = "";
    for (let attempt = 0; attempt < 6; attempt++) {
      if (!auth || Date.now() - auth.at > 3 * 60 * 1000) await ensureAuth();
      const used = auth;
      const headers = { ...used.headers, action, "content-type": "application/json; charset=utf-8" };
      delete headers["x-owa-urlpostdata"];
      delete headers["content-length"];
      // Give up on a request that stalls (dropped connection) and retry it.
      const ctl = new AbortController();
      const timer = setTimeout(() => ctl.abort(), 120 * 1000);
      let r, text;
      try {
        r = await origFetch(`${used.base}?action=${action}&app=Calendar`, {
          method: "POST",
          credentials: "include",
          headers,
          signal: ctl.signal,
          body: JSON.stringify(body ? { __type: `${action}JsonRequest:#Exchange`, Header: HEADER, Body: body } : {}),
        });
        text = await r.text();
      } catch (e) {
        lastErr = ctl.signal.aborted ? "timed out" : e.message;
        warn(`${action} ${lastErr}, retrying...`);
        await sleep(2000 * 2 ** attempt);
        continue;
      } finally {
        clearTimeout(timer);
      }
      if (r.status === 401 || r.status === 440) { // token expired, grab a new one
        if (auth === used) auth.at = 0;
        lastErr = `HTTP ${r.status}`;
        continue;
      }
      // 429/503: throttled. 500/502/504: Outlook has these transient failures
      // (e.g. too many requests in flight), so they're worth a retry too.
      if ([429, 500, 502, 503, 504].includes(r.status)) {
        const wait = Number(r.headers.get("retry-after")) || 2 ** attempt;
        warn(`${action} HTTP ${r.status}, retrying in ${wait}s`);
        lastErr = `HTTP ${r.status}`;
        await sleep(wait * 1000);
        continue;
      }
      if (!r.ok) throw new Error(`${action} HTTP ${r.status}: ${text.slice(0, 300)}`);
      try {
        return JSON.parse(text);
      } catch (e) {
        // An expired session can come back as a 200 HTML page; get a new login.
        if (auth === used) auth.at = 0;
        lastErr = "response wasn't JSON";
        continue;
      }
    }
    throw new Error(`${action}: gave up after repeated failures (${lastErr})`);
  }

  // --------------------------------------------------------------------- calendars
  log("Getting Outlook's login...");
  await ensureAuth();

  const folders = await svc("GetCalendarFolders", null);
  const calendars = [];
  for (const g of folders.CalendarGroups || []) {
    for (const c of g.Calendars || []) {
      if (c.CalendarFolderId && c.CalendarFolderId.Id) {
        calendars.push({ name: c.CalendarName, id: c.CalendarFolderId.Id, group: g.GroupName });
      } else {
        warn(`Skipping "${c.CalendarName}" (no folder id; probably a shared/linked calendar)`);
      }
    }
  }
  if (!calendars.length) throw new Error("No calendars found: " + JSON.stringify(folders).slice(0, 300));
  log(`Found ${calendars.length} calendar(s):`, calendars.map((c) => `${c.name} [${c.group}]`).join(", "));

  // ------------------------------------------------------------------------ events
  async function viewRange(folderId, start, end) {
    const j = await svc("GetCalendarView", {
      __type: "GetCalendarViewRequest:#Exchange",
      CalendarId: { __type: "TargetFolderId:#Exchange", BaseFolderId: { __type: "FolderId:#Exchange", Id: folderId } },
      RangeStart: start,
      RangeEnd: end,
    });
    if (!j.Body || j.Body.ResponseClass !== "Success") throw new Error((j.Body && j.Body.MessageText) || "GetCalendarView failed");
    return j.Body.Items || [];
  }

  // GetItem for a list of item ids, 25 per request. Returns a Map id -> item;
  // ids that fail are simply missing from it.
  let detailFailures = 0; // events exported without notes/attendees
  async function getItems(ids) {
    const batches = [];
    for (let i = 0; i < ids.length; i += 25) batches.push(ids.slice(i, i + 25));
    const found = new Map();
    let done = 0;
    await pool(batches, 6, async (batch) => {
      try {
        const j = await svc("GetItem", {
          __type: "GetItemRequest:#Exchange",
          ItemShape: { __type: "ItemResponseShape:#Exchange", BaseShape: "AllProperties", BodyType: "Text" },
          ItemIds: batch.map((Id) => ({ __type: "ItemId:#Exchange", Id })),
        });
        ((j.Body && j.Body.ResponseMessages && j.Body.ResponseMessages.Items) || []).forEach((m, k) => {
          if (m.ResponseClass === "Success" && m.Items && m.Items[0]) found.set(batch[k], m.Items[0]);
        });
      } catch (e) {
        if (e.fatal) throw e;
        warn("GetItem failed for a batch:", e.message);
      }
      done += batch.length;
      if (ids.length > 100 && done % 250 < 25) log(`  details ${done}/${ids.length}`);
    });
    return found;
  }

  // Full details (notes, attendees, SMTP addresses). Falls back to the view data.
  //
  // Plain recurring occurrences are identical to their series apart from
  // their own id, times and type, so fetch each series once and copy its
  // details onto every occurrence. Singles and exceptions (occurrences edited
  // on their own) are fetched individually.
  async function getFull(items) {
    const masterOf = (it) =>
      it.CalendarItemType === "Occurrence" && it.SeriesMasterItemId && it.SeriesMasterItemId.Id;
    const ids = new Set();
    for (const it of items) ids.add(masterOf(it) || it.ItemId.Id);
    let found = await getItems([...ids]);

    // If a series couldn't be fetched, fall back to fetching its occurrences.
    const retry = items.filter((it) => masterOf(it) && !found.has(masterOf(it))).map((it) => it.ItemId.Id);
    if (retry.length) {
      warn(`  ${retry.length} occurrences: series lookup failed, fetching them one by one`);
      found = new Map([...found, ...(await getItems(retry))]);
    }

    // GetItem leaves out some fields the view has (e.g. Location), so layer it
    // on top of the view item. Keep the view's Start/End: for occurrences
    // they're the occurrence's own times.
    return items.map((it) => {
      const own = found.get(it.ItemId.Id);
      if (own) return { ...it, ...own, Start: it.Start, End: it.End };
      const series = masterOf(it) && found.get(masterOf(it));
      if (series) {
        return { ...it, ...series, Start: it.Start, End: it.End, ItemId: it.ItemId, CalendarItemType: it.CalendarItemType };
      }
      detailFailures++;
      return it;
    });
  }

  // -------------------------------------------------------------------------- ics
  const esc = (t) => String(t == null ? "" : t).replace(/\\/g, "\\\\").replace(/;/g, "\\;")
    .replace(/,/g, "\\,").replace(/\r\n|\n/g, "\\n").replace(/\r/g, "");
  // UTF-8 length of one code point.
  const u8len = (cp) => (cp < 0x80 ? 1 : cp < 0x800 ? 2 : cp < 0x10000 ? 3 : 4);
  const fold = (line) => {
    if (line.length <= 75 && !/[^\x00-\x7f]/.test(line)) return line; // short ASCII: nothing to do
    const out = []; let cur = "", size = 0;
    for (const ch of line) {
      const n = u8len(ch.codePointAt(0)), lim = out.length ? 74 : 75;
      if (size + n > lim) { out.push(cur); cur = ""; size = 0; }
      cur += ch; size += n;
    }
    out.push(cur);
    return out.join("\r\n ");
  };
  const pad = (n) => String(n).padStart(2, "0");
  const utc = (s) => { const d = new Date(s); return `${d.getUTCFullYear()}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}T${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}${pad(d.getUTCSeconds())}Z`; };
  // Times are requested in UTC. All-day events are "floating": Outlook returns
  // them as midnight in whatever zone is requested, so the UTC date is the date.
  const ymd = (s) => { const d = new Date(s); return `${d.getUTCFullYear()}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}`; };
  // Parameter values (e.g. CN=) can't contain quotes or control characters.
  const pq = (v) => { v = String(v || "").replace(/"/g, "'").replace(/[\x00-\x1f\x7f]+/g, " "); return /[:;,]/.test(v) ? `"${v}"` : v; };
  const smtp = (mb) => mb && (mb.SmtpEmailAddress || (/@/.test(mb.EmailAddress || "") ? mb.EmailAddress : ""));

  const PARTSTAT = { Accept: "ACCEPTED", Decline: "DECLINED", Tentative: "TENTATIVE", Organizer: "ACCEPTED" };
  const CLASS = { Normal: "PUBLIC", Personal: "PRIVATE", Private: "PRIVATE", Confidential: "CONFIDENTIAL" };
  const PRIORITY = { High: "1", Normal: "5", Low: "9" };
  const BUSY = { Free: "FREE", Tentative: "TENTATIVE", Busy: "BUSY", OOF: "OOF", WorkingElsewhere: "WORKINGELSEWHERE" };
  const stamp = utc(new Date().toISOString());

  // Recurring occurrences share the series UID; make each one unique so
  // importers don't collapse them into a single event. Use the occurrence's
  // original slot (not where it was moved to) so the UID stays the same if
  // it's moved and exported again.
  const uidOf = (it) => {
    const base = it.UID || (it.ItemId && it.ItemId.Id);
    if (!/Occurrence|Exception/.test(it.CalendarItemType || "")) return base;
    return `${base}-${utc(it.OriginalStart || it.RecurrenceId || it.Start)}`;
  };

  function vevent(it) {
    const L = ["BEGIN:VEVENT", `UID:${esc(uidOf(it))}`, `DTSTAMP:${stamp}`];
    if (it.IsAllDayEvent) L.push(`DTSTART;VALUE=DATE:${ymd(it.Start)}`, `DTEND;VALUE=DATE:${ymd(it.End)}`);
    else L.push(`DTSTART:${utc(it.Start)}`, `DTEND:${utc(it.End)}`);
    L.push(`SUMMARY:${esc(it.Subject || "(No title)")}`);

    const loc = typeof it.Location === "string" ? it.Location : it.Location && it.Location.DisplayName;
    if (loc) L.push(`LOCATION:${esc(loc)}`);

    let desc = ((it.Body && it.Body.Value) || "").trim();
    const join = it.JoinOnlineMeetingUrl || it.OnlineMeetingJoinUrl ||
      (it.OnlineMeetingSettings && it.OnlineMeetingSettings.JoinUrl) ||
      (it.SkypeTeamsProperties && it.SkypeTeamsProperties.JoinUrl);
    if (join) {
      L.push(`URL:${String(join).replace(/[\r\n]+/g, "")}`); // a URI value, not TEXT: no backslash escaping
      if (!desc.includes(join)) desc = (desc + "\n\nJoin: " + join).trim();
    }
    if (desc) L.push(`DESCRIPTION:${esc(desc)}`);

    const org = it.Organizer && it.Organizer.Mailbox;
    if (smtp(org)) L.push(`ORGANIZER;CN=${pq(org.Name)}:mailto:${smtp(org)}`);
    for (const [list, role] of [["RequiredAttendees", "REQ-PARTICIPANT"], ["OptionalAttendees", "OPT-PARTICIPANT"], ["Resources", "NON-PARTICIPANT"]]) {
      for (const a of it[list] || []) {
        const addr = smtp(a.Mailbox);
        if (!addr) continue;
        L.push(`ATTENDEE;CN=${pq(a.Mailbox.Name)};ROLE=${role};PARTSTAT=${PARTSTAT[a.ResponseType] || "NEEDS-ACTION"}:mailto:${addr}`);
      }
    }
    if (it.Categories && it.Categories.length) L.push("CATEGORIES:" + it.Categories.map(esc).join(","));

    const fb = it.FreeBusyType || it.LegacyFreeBusyStatus || "Busy";
    L.push(`TRANSP:${fb === "Free" ? "TRANSPARENT" : "OPAQUE"}`,
           `X-MICROSOFT-CDO-BUSYSTATUS:${BUSY[fb] || "BUSY"}`,
           `CLASS:${CLASS[it.Sensitivity] || "PUBLIC"}`,
           `PRIORITY:${PRIORITY[it.Importance] || "5"}`,
           `STATUS:${it.IsCancelled ? "CANCELLED" : "CONFIRMED"}`);
    if (it.DateTimeCreated) L.push(`CREATED:${utc(it.DateTimeCreated)}`);
    if (it.LastModifiedTime) L.push(`LAST-MODIFIED:${utc(it.LastModifiedTime)}`);
    if (it.ReminderIsSet && it.ReminderMinutesBeforeStart != null) {
      L.push("BEGIN:VALARM", "ACTION:DISPLAY", "DESCRIPTION:Reminder",
             `TRIGGER:-PT${parseInt(it.ReminderMinutesBeforeStart, 10) || 0}M`, "END:VALARM");
    }
    L.push("END:VEVENT");
    return L;
  }

  // Each event is rendered once and reused for its own calendar's file and
  // for all_calendars.ics.
  const rendered = new WeakMap();
  const veventText = (it) => {
    if (!rendered.has(it)) rendered.set(it, vevent(it).map(fold).join("\r\n"));
    return rendered.get(it);
  };
  const buildIcs = (name, items) => [
    ...["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//outlook-console-export//EN",
        "CALSCALE:GREGORIAN", "METHOD:PUBLISH", `X-WR-CALNAME:${esc(name)}`].map(fold),
    ...items.map(veventText), "END:VCALENDAR",
  ].join("\r\n") + "\r\n";

  const download = (filename, text) => {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([text], { type: "text/calendar;charset=utf-8" }));
    a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 10000);
  };
  const safe = (s) => (String(s || "calendar").replace(/[^\p{L}\p{N}\-_. ]+/gu, "_").trim() || "calendar");

  // ------------------------------------------------------------------------- run
  const all = new Map();
  const summary = {};
  const skipped = [];
  try {
    for (const cal of calendars) {
      log(`Reading "${cal.name}"...`);
      // Two calendars can share a name; keep both in the summary.
      let key = cal.name;
      for (let n = 2; key in summary; n++) key = `${cal.name} (${n})`;
      const byId = new Map();
      try {
        // Outlook allows at most 1825 days per request, so ask for 4 years at a
        // time. A chunk that comes back very full is re-read year by year, in
        // case there's a per-response cap (none seen up to 2,600 events).
        const yearRange = (from, to) => viewRange(cal.id, `${from}-01-01T00:00:00.000`, `${to}-01-01T00:00:00.000`);
        const chunks = [];
        for (let y = START_YEAR; y <= END_YEAR; y += 4) chunks.push([y, Math.min(y + 4, END_YEAR + 1)]);
        const perChunk = await pool(chunks, 6, ([from, to]) => yearRange(from, to));
        const years = [];
        chunks.forEach(([from, to], k) => {
          if (perChunk[k].length >= 2500 && to - from > 1) {
            for (let y = from; y < to; y++) years.push(y);
            perChunk[k] = [];
          }
        });
        const perYear = await pool(years, 6, (y) => yearRange(y, y + 1));
        for (const it of [...perChunk, ...perYear].flat()) byId.set(it.ItemId.Id, it); // events spanning chunk edges show up twice
        log(`  ${byId.size} events found, fetching details...`);
      } catch (e) {
        if (e.fatal) throw e;
        warn(`  skipped "${cal.name}": ${e.message}`);
        skipped.push(`${cal.name} (${e.message})`);
        continue;
      }
      const items = (await getFull([...byId.values()])).sort((a, b) => new Date(a.Start) - new Date(b.Start));
      log(`  ${items.length} events`);
      summary[key] = items.length;
      if (items.length) download(`${safe(cal.name)}.ics`, buildIcs(cal.name, items));
      for (const it of items) all.set(uidOf(it), it);
    }
  } finally {
    auth = null;
  }

  const combined = [...all.values()].sort((a, b) => new Date(a.Start) - new Date(b.Start));
  if (!combined.length) {
    warn("No events found in any calendar between", START_YEAR, "and", END_YEAR, "- nothing downloaded.");
  } else {
    download("all_calendars.ics", buildIcs("All Outlook calendars", combined));
    log(`Done. ${combined.length} events total -> all_calendars.ics`);
  }
  // Make partial exports impossible to miss.
  if (skipped.length) console.error(`[ICS export] NOT exported: ${skipped.join("; ")}`);
  if (detailFailures) console.error(`[ICS export] ${detailFailures} events exported without notes/attendees (details lookup failed).`);
  return { calendars: summary, total: combined.length, skipped, detailFailures };
})().catch((e) => console.error("[ICS export] failed:", e));
