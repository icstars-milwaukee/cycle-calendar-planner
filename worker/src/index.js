/**
 * Realtime sync backend for the Cycle Calendar Planner.
 *
 * GitHub Pages can only serve static files, so the planner UI stays there and this
 * Worker holds the shared board. One Durable Object instance per room name gives us
 * a single authoritative copy plus a WebSocket fan-out to everyone viewing it.
 *
 * Anyone who knows a room name can read and write that room -- there is no login.
 * Treat the room name like a "anyone with the link can edit" share URL.
 */

const ALLOWED_ORIGINS = [
  "https://icstars-milwaukee.github.io",
  "http://localhost:8000",
  "http://127.0.0.1:8000",
];

// A single board's JSON is a few KB. Cap well above that so a bug or a paste-bomb
// cannot fill the object's storage.
const MAX_PLAN_BYTES = 512 * 1024;

/**
 * Shared-password gate. The password lives in the ROOM_PASSWORD secret, never in
 * this repository -- the repo is public, so a committed password would be no
 * password at all. If the secret is missing we deny everything rather than falling
 * open, so a misconfigured deploy cannot quietly expose the board.
 */
function timingSafeEqual(a, b) {
  const ab = new TextEncoder().encode(a);
  const bb = new TextEncoder().encode(b);
  // Compare every byte regardless of mismatch position so failure time does not
  // leak how much of the password was correct.
  let diff = ab.length ^ bb.length;
  const len = Math.max(ab.length, bb.length);
  for (let i = 0; i < len; i++) diff |= (ab[i] || 0) ^ (bb[i] || 0);
  return diff === 0;
}

function authorized(request, env) {
  if (!env.ROOM_PASSWORD) return false;
  const url = new URL(request.url);
  // Browsers cannot set headers on a WebSocket handshake, so the query string is
  // the only channel available. It travels inside the TLS tunnel.
  const supplied = url.searchParams.get("pw") || request.headers.get("X-Board-Password") || "";
  return timingSafeEqual(supplied, env.ROOM_PASSWORD);
}

/**
 * Projects a stored board into concrete calendar events.
 *
 * The board itself is weekday-relative ("Tuesday, 10:00, 60 minutes"), which is why
 * plan.weekOf exists: it pins the board to a real Monday so a date can be computed.
 * Without it there is no answer to "what date is this block on" and nothing can be
 * pushed to a calendar, so this returns an explicit error rather than guessing.
 *
 * Times are emitted as local wall-clock plus an IANA zone, matching Microsoft Graph's
 * dateTimeTimeZone shape so a bridge can forward them without conversion. Sending
 * local time avoids a whole class of daylight-saving bugs: 9am stays 9am across the
 * March and November transitions, which a fixed UTC offset would silently shift.
 */
const CAL_TZ = "America/Chicago";

// A cycle is always 14 weeks. This is a property of the programme, not a setting:
// the board describes one week's rhythm and that rhythm runs for the whole cycle.
const CYCLE_WEEKS = 14;

function pad2(n) { return (n < 10 ? "0" : "") + n; }

// ---------- holidays ----------
// Computed rather than listed so a cycle in any year works without maintenance.
// All arithmetic is on the UTC calendar so a server timezone can never shift a date.

function isoOf(d) {
  return d.getUTCFullYear() + "-" + pad2(d.getUTCMonth() + 1) + "-" + pad2(d.getUTCDate());
}
function utc(y, m, d) { return new Date(Date.UTC(y, m, d)); }

// n-th given weekday of a month, 1-based. weekday: 0 = Sunday.
function nthWeekday(year, month, weekday, n) {
  const first = utc(year, month, 1);
  const shift = (weekday - first.getUTCDay() + 7) % 7;
  return utc(year, month, 1 + shift + (n - 1) * 7);
}
function lastWeekday(year, month, weekday) {
  const last = utc(year, month + 1, 0);
  const shift = (last.getUTCDay() - weekday + 7) % 7;
  return utc(year, month, last.getUTCDate() - shift);
}

// Federal holidays on a fixed date move when they land on a weekend: Saturday is
// observed the Friday before, Sunday the Monday after. Without this, a cycle would
// keep scheduling through the day the office is actually shut.
function observed(d) {
  const day = d.getUTCDay();
  if (day === 6) return utc(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate() - 1);
  if (day === 0) return utc(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate() + 1);
  return d;
}

function federalHolidays(year) {
  return [
    { name: "New Year's Day", date: observed(utc(year, 0, 1)) },
    { name: "Martin Luther King Jr. Day", date: nthWeekday(year, 0, 1, 3) },
    { name: "Presidents' Day", date: nthWeekday(year, 1, 1, 3) },
    { name: "Memorial Day", date: lastWeekday(year, 4, 1) },
    { name: "Juneteenth", date: observed(utc(year, 5, 19)) },
    { name: "Independence Day", date: observed(utc(year, 6, 4)) },
    { name: "Labor Day", date: nthWeekday(year, 8, 1, 1) },
    { name: "Indigenous Peoples' Day", date: nthWeekday(year, 9, 1, 2) },
    { name: "Veterans Day", date: observed(utc(year, 10, 11)) },
    { name: "Thanksgiving Day", date: nthWeekday(year, 10, 4, 4) },
    { name: "Day after Thanksgiving", date: utc(year, 10, nthWeekday(year, 10, 4, 4).getUTCDate() + 1) },
    { name: "Christmas Day", date: observed(utc(year, 11, 25)) },
  ].map((h) => ({ name: h.name, date: isoOf(h.date) }));
}

// Every holiday between two ISO dates inclusive. A 14-week cycle can straddle a year
// boundary, so both years are considered.
function holidaysBetween(startISO, endISO) {
  const y0 = +startISO.slice(0, 4);
  const y1 = +endISO.slice(0, 4);
  const out = [];
  for (let y = y0; y <= y1; y++) {
    for (const h of federalHolidays(y)) {
      if (h.date >= startISO && h.date <= endISO) out.push(h);
    }
  }
  out.sort((a, b) => a.date.localeCompare(b.date));
  return out;
}

// The last weekday of the cycle: Friday of week 14.
function cycleEnd(weekOf) {
  return addDaysISO(weekOf, (CYCLE_WEEKS - 1) * 7 + 4);
}

// Days are added on the UTC calendar so a host timezone can never shift the date.
function addDaysISO(mondayISO, days) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(mondayISO || "");
  if (!m) return null;
  const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  d.setUTCDate(d.getUTCDate() + days);
  return d.getUTCFullYear() + "-" + pad2(d.getUTCMonth() + 1) + "-" + pad2(d.getUTCDate());
}

/**
 * The fixed weekly skeleton: the rituals that happen whether or not a workshop is
 * planned. These live in the planner as a drawn backdrop, never as placed blocks, so
 * nothing ever sent them to the calendar -- which left the published day missing All
 * Skate, High Tea and every break. Mirrored here so the calendar shows the real day.
 *
 * days omitted means all five. Kept in step with the FIXED array in index.html.
 */
const HOLDS = [
  { title: "All Skate",         start: 510,  end: 540 },
  { title: "Lunch",             start: 720,  end: 780 },
  { title: "Wellness Workshop", start: 780,  end: 840,  days: [2] },
  { title: "Out of cycle",      start: 840,  end: 1155, days: [2] },
  { title: "Tea Prep",          start: 930,  end: 960,  days: [0, 1, 3, 4] },
  { title: "High Tea",          start: 960,  end: 1020, days: [0, 1, 3, 4] },
  { title: "Tea Debrief",       start: 1020, end: 1050, days: [0, 1, 3, 4] },
  { title: "Break",             start: 1050, end: 1080, days: [0, 1, 3, 4] },
];

function holdsForDay(day) {
  return HOLDS.filter((h) => !h.days || h.days.indexOf(day) >= 0);
}

function slug(s) {
  return String(s).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
}

function hhmm(minutesOfDay) {
  const m = Math.max(0, Math.min(24 * 60 - 1, Math.round(minutesOfDay)));
  return pad2(Math.floor(m / 60)) + ":" + pad2(m % 60) + ":00";
}

function toEvents(plan, room) {
  if (!plan) return { ok: true, weekOf: null, events: [] };
  if (!plan.weekOf) {
    return {
      ok: false,
      error: "This board has no week set. Pick the Monday it starts in the planner before syncing.",
      weekOf: null,
      events: [],
    };
  }

  const cats = Array.isArray(plan.cats) ? plan.cats : [];
  const catName = (id) => (cats.find((c) => c.id === id) || {}).name || "";
  const prefix = "cycle-planner-" + (room || "board") + "-";

  const endISO = cycleEnd(plan.weekOf);
  const holidays = holidaysBetween(plan.weekOf, endISO);

  // A holiday blocks scheduling unless it has been explicitly unblocked. Defaulting to
  // blocked is the safe direction: forgetting to mark one leaves an empty day on the
  // calendar, while the opposite books sessions on a day the office is closed.
  const blocks = plan.holidayBlocks || {};
  const blocked = new Set(holidays.filter((h) => blocks[h.date] !== false).map((h) => h.date));

  const events = [];

  // Each of the fourteen weeks has its own board. Older boards stored a single week
  // meant to repeat, so fall back to that and treat every week as identical.
  const weeks = Array.isArray(plan.weeks) && plan.weeks.length === CYCLE_WEEKS
    ? plan.weeks
    : Array.from({ length: CYCLE_WEEKS }, () => plan.placed || []);

  for (let w = 0; w < CYCLE_WEEKS; w++) {
    for (const p of Array.isArray(weeks[w]) ? weeks[w] : []) {
      const date = addDaysISO(plan.weekOf, w * 7 + p.day);
      if (!date) continue;
      if (blocked.has(date)) continue;
      // Placed blocks carry `uid`; `id` is only used by shelf items and categories.
      // Reading the wrong field collapsed every block in a week onto one identifier,
      // so the calendar received a single event per week instead of the whole day.
      const blockId = p.uid || p.id;
      if (!blockId) continue;
      events.push({
        uid: prefix + blockId + "-w" + (w + 1),
        subject: p.title,
        category: catName(p.cat || (p.refId ? (plan.custom || []).find((c) => c.id === p.refId) || {} : {}).cat),
        start: { dateTime: date + "T" + hhmm(p.start), timeZone: CAL_TZ },
        end: { dateTime: date + "T" + hhmm(p.start + p.mins), timeZone: CAL_TZ },
        week: w + 1,
        day: p.day,
        minutes: p.mins,
      });
    }
  }

  // The fixed skeleton, unless someone has turned it off. Emitted per day so a holiday
  // that closes a day removes its rituals too, not just its workshops.
  if (plan.includeHolds !== false) {
    for (let w = 0; w < CYCLE_WEEKS; w++) {
      for (let d = 0; d < 5; d++) {
        const date = addDaysISO(plan.weekOf, w * 7 + d);
        if (!date || blocked.has(date)) continue;
        for (const h of holdsForDay(d)) {
          events.push({
            uid: prefix + "hold-" + slug(h.title) + "-w" + (w + 1) + "-d" + d,
            subject: h.title,
            category: "Fixed hold",
            start: { dateTime: date + "T" + hhmm(h.start), timeZone: CAL_TZ },
            end: { dateTime: date + "T" + hhmm(h.end), timeZone: CAL_TZ },
            week: w + 1,
            day: d,
            minutes: h.end - h.start,
            isHold: true,
          });
        }
      }
    }
  }

  for (const h of holidays) {
    events.push({
      uid: prefix + "holiday-" + h.date,
      subject: h.name,
      category: "Holiday",
      isAllDay: true,
      // Graph treats an all-day event as a half-open range, so it ends the next midnight.
      start: { dateTime: h.date + "T00:00:00", timeZone: CAL_TZ },
      end: { dateTime: addDaysISO(h.date, 1) + "T00:00:00", timeZone: CAL_TZ },
      blocksScheduling: blocked.has(h.date),
    });
  }

  // Deterministic order so a diff between two pushes is meaningful.
  events.sort((a, b) => a.start.dateTime.localeCompare(b.start.dateTime) || a.uid.localeCompare(b.uid));

  return {
    ok: true,
    weekOf: plan.weekOf,
    cycleWeeks: CYCLE_WEEKS,
    cycleEnd: endISO,
    timeZone: CAL_TZ,
    holidays: holidays.map((h) => ({ ...h, blocksScheduling: blocked.has(h.date) })),
    count: events.length,
    events,
  };
}

/**
 * How far the live board has drifted from what was last published.
 *
 * The board is edited continuously and collaboratively, so it is never a safe thing to
 * mirror directly -- a half-finished drag would reach people's calendars. The calendar
 * follows an explicitly published snapshot instead, and this reports what publishing
 * would change so the UI can say so before anyone commits.
 */
function pendingChanges(livePlan, publishedPlan, room) {
  const live = toEvents(livePlan, room);
  const pub = toEvents(publishedPlan, room);

  // An unset week is a blocker, not a change count: nothing can be published yet.
  if (!live.ok) return { ready: false, error: live.error, added: 0, changed: 0, removed: 0, total: 0 };

  const pubMap = new Map(pub.ok ? pub.events.map((e) => [e.uid, eventHash(e)]) : []);
  let added = 0, changed = 0;
  for (const ev of live.events) {
    const prior = pubMap.get(ev.uid);
    if (prior === undefined) added++;
    else if (prior !== eventHash(ev)) changed++;
  }
  const liveUids = new Set(live.events.map((e) => e.uid));
  let removed = 0;
  for (const uid of pubMap.keys()) if (!liveUids.has(uid)) removed++;

  return { ready: true, added, changed, removed, total: added + changed + removed };
}

/**
 * Reconciliation against a Microsoft 365 Group calendar.
 *
 * Microsoft Graph has no application permission for group calendars, so the push has
 * to run as a signed-in member -- in practice a Power Automate flow. Flows are a poor
 * place for diffing logic, so the Worker does it here and hands the flow three flat
 * lists to loop over.
 *
 * The Worker remembers which Graph event each block became, keyed by the block's
 * stable uid. That mapping is what makes a re-push an update instead of a duplicate.
 */

// Content fingerprint. Only the fields a calendar actually shows are included, so
// unrelated board edits do not churn the calendar with pointless updates.
function eventHash(ev) {
  const basis = [ev.subject, ev.start.dateTime, ev.end.dateTime, ev.start.timeZone,
                 ev.category, ev.isAllDay ? "allday" : ""].join("|");
  let h = 5381;
  for (let i = 0; i < basis.length; i++) h = ((h << 5) + h + basis.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36);
}

// Shaped so a flow can hand it to "Create group event" with no transformation.
function graphBody(ev) {
  return {
    subject: ev.subject,
    start: ev.start,
    end: ev.end,
    isAllDay: !!ev.isAllDay,
    // The planner category rides as a real Outlook category, not just body text.
    // Outlook maps the NAME to a color from the calendar's own category list, which
    // Graph does not expose for group mailboxes -- so someone defines each name's
    // color once in Outlook, and every event wearing that name follows it.
    categories: ev.category ? [ev.category] : [],
    // Graph uses transactionId to make event creation idempotent: re-POSTing the same
    // one does not create a second event. Without it, any run that creates events and
    // then dies before reporting back duplicates all of them on the next attempt --
    // which is not a rare edge case when 284 events meet a 100-call-per-minute limit.
    transactionId: ev.uid,
    body: {
      contentType: "text",
      content: "Cycle Calendar Planner" + (ev.category ? " - " + ev.category : "") +
        (ev.week ? "\nWeek " + ev.week + " of " + CYCLE_WEEKS : "") +
        "\nDo not edit here; edit the planner and it will be overwritten on the next sync." +
        "\nref: " + ev.uid,
    },
  };
}

function buildSyncPlan(publishedPlan, room, pushed) {
  // Deliberately the *published* snapshot, never the live board. Someone mid-drag must
  // not reach anyone's calendar.
  if (!publishedPlan) {
    return {
      ok: false,
      error: "Nothing has been published yet. Open the planner and choose Publish to calendar.",
      create: [], update: [], delete: [],
    };
  }

  const projection = toEvents(publishedPlan, room);

  // No week means no dates. Returning empty lists rather than an error with deletions
  // matters: an unset week must never be read as "the board is empty, remove everything".
  if (!projection.ok) {
    return { ok: false, error: projection.error, create: [], update: [], delete: [] };
  }

  const live = new Map(projection.events.map((e) => [e.uid, e]));
  const create = [];
  const update = [];
  const remove = [];

  for (const ev of projection.events) {
    const prior = pushed[ev.uid];
    const hash = eventHash(ev);
    if (!prior) {
      create.push({ uid: ev.uid, hash, event: graphBody(ev) });
    } else if (prior.hash !== hash) {
      update.push({ uid: ev.uid, hash, eventId: prior.graphId, event: graphBody(ev) });
    }
  }

  for (const uid of Object.keys(pushed)) {
    if (!live.has(uid)) remove.push({ uid, eventId: pushed[uid].graphId });
  }

  return {
    ok: true,
    weekOf: projection.weekOf,
    timeZone: projection.timeZone,
    create,
    update,
    delete: remove,
    // Lets a flow skip its loops entirely when the calendar is already correct.
    changes: create.length + update.length + remove.length,
  };
}

/**
 * Reconciles the published cycle against what is actually on the calendar.
 *
 * The older sync-plan trusted a stored map of "what we created". That map is wrong the
 * moment a run creates events and dies before reporting back, and a wrong map produces
 * duplicates on the next run -- which is exactly what happened. This takes the calendar
 * itself as the source of truth, so it is correct no matter how badly a previous run
 * failed, and needs no acknowledgement step at all.
 *
 * Every event we own carries its uid twice: in transactionId, and as "ref: <uid>" in the
 * body. Older events predate transactionId, so the body is the fallback.
 */
function uidOfExisting(ev, marker) {
  const tx = ev && ev.transactionId;
  if (typeof tx === "string" && tx.startsWith(marker)) return tx;
  const content = (ev && ev.body && ev.body.content) || ev.bodyPreview || "";
  const m = /ref:\s*(\S+)/.exec(String(content));
  if (m && m[1].startsWith(marker)) return m[1];
  return null;
}

function reconcile(publishedPlan, room, existingRaw) {
  if (!publishedPlan) {
    return { ok: false, error: "Nothing has been published yet.", create: [], update: [], delete: [] };
  }
  const marker = "cycle-planner-" + room + "-";

  // A retired room: everything it ever put on the calendar is deleted, holidays and
  // holds included, and nothing is created. Used to clean a room up before wiping it.
  if (publishedPlan.purgeAll === true) {
    const remove = [];
    let foreignCount = 0;
    for (const ev of Array.isArray(existingRaw) ? existingRaw : []) {
      const uid = uidOfExisting(ev, marker);
      if (uid) remove.push({ uid, eventId: ev.id, reason: "room retired" });
      else foreignCount++;
    }
    return {
      ok: true, purge: true, marker,
      inspected: (existingRaw || []).length, ours: remove.length, foreign: foreignCount,
      create: [], update: [], delete: remove, changes: remove.length,
    };
  }

  const projection = toEvents(publishedPlan, room);
  if (!projection.ok) {
    return { ok: false, error: projection.error, create: [], update: [], delete: [] };
  }
  const desired = new Map(projection.events.map((e) => [e.uid, e]));

  // Group what is on the calendar by uid. More than one entry for a uid means an earlier
  // run duplicated it; the extras are deleted rather than left for a human to find.
  const seen = new Map();
  let foreign = 0;
  for (const ev of Array.isArray(existingRaw) ? existingRaw : []) {
    const uid = uidOfExisting(ev, marker);
    if (!uid) { foreign++; continue; }   // not ours -- never touched
    if (!seen.has(uid)) seen.set(uid, []);
    seen.get(uid).push(ev);
  }

  const create = [];
  const update = [];
  const remove = [];
  // Every board event's calendar identity, refreshed each time the calendar is read.
  const known = {};

  for (const [uid, ev] of desired) {
    const matches = seen.get(uid);
    if (!matches || !matches.length) {
      create.push({ uid, event: graphBody(ev) });
      continue;
    }
    // Keep the first, drop any duplicates of it.
    const keep = matches[0];
    known[uid] = keep.id;
    for (let i = 1; i < matches.length; i++) remove.push({ uid, eventId: matches[i].id, reason: "duplicate" });

    // Only touch it if what is on the calendar actually differs.
    const sameSubject = (keep.subject || "") === ev.subject;
    const sameStart = String((keep.start && keep.start.dateTime) || "").slice(0, 19) === ev.start.dateTime;
    const sameEnd = String((keep.end && keep.end.dateTime) || "").slice(0, 19) === ev.end.dateTime;
    // Category drift counts too -- but only when the flow sent categories at all.
    // A flow that predates the category column would otherwise make every event look
    // uncategorized and rewrite the entire calendar on every run.
    const sameCat = keep.categories === undefined ||
      ((keep.categories || [])[0] || "") === (ev.category || "");
    if (!sameSubject || !sameStart || !sameEnd || !sameCat) {
      update.push({ uid, eventId: keep.id, event: graphBody(ev) });
    }
  }

  // Ours, in the window, but no longer part of the cycle.
  for (const [uid, matches] of seen) {
    if (desired.has(uid)) continue;
    for (const ev of matches) remove.push({ uid, eventId: ev.id, reason: "no longer scheduled" });
  }

  return {
    ok: true,
    weekOf: projection.weekOf,
    marker,
    inspected: (existingRaw || []).length,
    ours: seen.size,
    foreign,
    known,
    create,
    update,
    delete: remove,
    changes: create.length + update.length + remove.length,
  };
}

function deny(cors) {
  return new Response(JSON.stringify({ ok: false, error: "Wrong password." }), {
    status: 401,
    headers: { ...cors, "Content-Type": "application/json" },
  });
}

function corsHeaders(origin) {
  const allow = ALLOWED_ORIGINS.includes(origin) ? origin : ALLOWED_ORIGINS[0];
  return {
    "Access-Control-Allow-Origin": allow,
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Vary": "Origin",
  };
}

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    const origin = request.headers.get("Origin") || "";
    const cors = corsHeaders(origin);

    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: cors });

    // Unauthenticated: reports only that the service is up and whether a password
    // was ever configured. Never reveals the password or any board content.
    if (url.pathname === "/health") {
      return new Response(JSON.stringify({ ok: true, locked: !!env.ROOM_PASSWORD }), {
        headers: { ...cors, "Content-Type": "application/json" },
      });
    }

    // Lets the UI tell "wrong password" apart from "network is down" before it
    // opens a socket, since a failed WebSocket handshake surfaces no useful reason.
    if (url.pathname === "/auth") {
      if (!authorized(request, env)) return deny(cors);
      return new Response(JSON.stringify({ ok: true }), {
        headers: { ...cors, "Content-Type": "application/json" },
      });
    }

    // Asks the groups flow to refresh the list now. Separate from the calendar flow on
    // purpose: refreshing groups should not run a calendar sync, and either flow being
    // broken should not take the other down.
    if (url.pathname === "/refresh-groups") {
      if (!authorized(request, env)) return deny(cors);
      if (request.method !== "POST") {
        return new Response("POST required", { status: 405, headers: cors });
      }
      if (!env.GROUPS_WEBHOOK) {
        return new Response(JSON.stringify({ ok: false, error: "No groups flow is configured." }), {
          status: 503, headers: { ...cors, "Content-Type": "application/json" },
        });
      }
      ctx.waitUntil(fetch(env.GROUPS_WEBHOOK, { method: "POST" }).catch(() => {}));
      return new Response(JSON.stringify({ ok: true, asked: true }), {
        headers: { ...cors, "Content-Type": "application/json" },
      });
    }

    // The cycle registry. GET lists cycles, POST creates one.
    // /groups holds the Microsoft 365 groups the flow can see, so the planner can offer
    // a real choice instead of a group name compiled into the flow.
    if (url.pathname === "/cycles" || url.pathname === "/groups") {
      if (!authorized(request, env)) return deny(cors);
      return env.CYCLES.get(env.CYCLES.idFromName("directory")).fetch(request);
    }

    // /room/<name>        -> WebSocket upgrade for live editing
    // /room/<name>/plan   -> plain GET snapshot, handy for debugging and backups
    const match = url.pathname.match(
      /^\/room\/([A-Za-z0-9_-]{1,64})(\/plan|\/events|\/sync-plan|\/sync-ack|\/init|\/reset-tracking|\/purge-info|\/reconcile|\/check|\/wipe|\/ids)?$/);
    if (!match) return new Response("Not found", { status: 404, headers: cors });

    // Every route that touches board data is gated, including the upgrade itself,
    // so an unauthenticated socket is never accepted in the first place.
    if (!authorized(request, env)) return deny(cors);

    const room = match[1];
    const id = env.BOARD.idFromName(room);
    return env.BOARD.get(id).fetch(request);
  },
};

/**
 * The list of cycles.
 *
 * A cycle is a name plus a start Monday; its 14 weeks follow from that and are not
 * separately editable. Keeping the start here rather than only inside a board means
 * "Cycle 21 starts 31 August" is a fact about the cycle, not a field someone can nudge
 * while dragging workshops around.
 */
export class Directory {
  constructor(ctx, env) {
    this.ctx = ctx;
    this.env = env;
  }

  async list() {
    return (await this.ctx.storage.get("cycles")) || [];
  }

  async fetch(request) {
    const origin = request.headers.get("Origin") || "";
    const cors = corsHeaders(origin);
    const json = (body, status) => new Response(JSON.stringify(body), {
      status: status || 200, headers: { ...cors, "Content-Type": "application/json" },
    });

    // The planner cannot call Microsoft Graph -- it is a static page with no sign-in, and
    // getting one would need admin consent. The flow can, using the connection that
    // already exists, so it posts the list here and the planner reads it back.
    if (new URL(request.url).pathname === "/groups") {
      if (request.method === "GET") {
        return json({
          ok: true,
          groups: (await this.ctx.storage.get("groups")) || [],
          updatedAt: (await this.ctx.storage.get("groupsUpdatedAt")) || null,
        });
      }
      if (request.method !== "POST") return json({ ok: false, error: "Unsupported method." }, 405);

      let payload;
      try { payload = await request.json(); } catch (e) { return json({ ok: false, error: "Body was not JSON." }, 400); }
      const incoming = Array.isArray(payload.groups) ? payload.groups : (payload.value || []);
      const groups = incoming
        .map((g) => ({
          id: String(g.id || "").slice(0, 64),
          name: String(g.displayName || g.name || "").slice(0, 120),
          mail: String(g.mail || "").slice(0, 160),
        }))
        .filter((g) => g.id && g.name);
      groups.sort((a, b) => a.name.localeCompare(b.name));

      // An empty list almost certainly means the lookup failed rather than that the
      // account belongs to nothing; keeping the previous list avoids emptying the
      // planner's picker on a transient error.
      if (!groups.length) {
        // Deliberate clearing has to be explicit, so a failed lookup cannot empty the
        // picker while a genuine reset is still possible.
        if (new URL(request.url).searchParams.get("clear") === "1") {
          await this.ctx.storage.put({ groups: [], groupsUpdatedAt: new Date().toISOString() });
          return json({ ok: true, cleared: true });
        }
        return json({ ok: true, kept: ((await this.ctx.storage.get("groups")) || []).length, note: "empty list ignored" });
      }

      await this.ctx.storage.put({ groups, groupsUpdatedAt: new Date().toISOString() });
      return json({ ok: true, stored: groups.length });
    }

    const cycles = await this.list();

    if (request.method === "GET") {
      return json({
        ok: true,
        cycleWeeks: CYCLE_WEEKS,
        cycles: cycles.map((c) => ({ ...c, end: cycleEnd(c.start) })),
      });
    }

    // Retiring a cycle removes it from the list and erases its board. The calendar is
    // left alone -- purge it first if its events should go too.
    if (request.method === "DELETE") {
      const id = new URL(request.url).searchParams.get("id") || "";
      const idx = cycles.findIndex((c) => c.id === id);
      if (idx < 0) return json({ ok: false, error: "No such cycle." }, 404);
      cycles.splice(idx, 1);
      await this.ctx.storage.put("cycles", cycles);
      try {
        const board = this.env.BOARD.get(this.env.BOARD.idFromName(id));
        await board.fetch(new Request("https://do/room/" + id + "/wipe", { method: "POST" }));
      } catch (e) { /* the list entry is gone either way */ }
      return json({ ok: true, removed: id });
    }

    if (request.method !== "POST") return json({ ok: false, error: "Unsupported method." }, 405);

    let body;
    try { body = await request.json(); } catch (e) { return json({ ok: false, error: "Body was not JSON." }, 400); }

    const name = String(body.name || "").trim();
    const start = String(body.start || "").trim();

    if (!name) return json({ ok: false, error: "Give the cycle a name, for example \"Cycle 21\"." }, 400);
    if (name.length > 60) return json({ ok: false, error: "That name is too long." }, 400);

    if (!/^\d{4}-\d{2}-\d{2}$/.test(start)) {
      return json({ ok: false, error: "Give a start date as YYYY-MM-DD." }, 400);
    }
    const startDate = new Date(start + "T00:00:00Z");
    if (isNaN(startDate.getTime()) || isoOf(startDate) !== start) {
      return json({ ok: false, error: "That is not a real date." }, 400);
    }
    // Week 1 Monday anchors every date in the cycle; a non-Monday start would put
    // "Monday" blocks on a Tuesday for all 14 weeks.
    if (startDate.getUTCDay() !== 1) {
      return json({ ok: false, error: "A cycle has to start on a Monday." }, 400);
    }

    // Slug becomes the room name and lives in URLs, so keep it conservative.
    let id = name.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 48);
    if (!id) id = "cycle";
    if (cycles.some((c) => c.id === id)) {
      return json({ ok: false, error: "A cycle called \"" + name + "\" already exists.", id }, 409);
    }

    const cycle = {
      id,
      name,
      start,
      weeks: CYCLE_WEEKS,
      // Chosen at creation so publishing works the moment the board opens, instead of
      // depending on someone finding the setting afterwards.
      groupId: String(body.groupId || "").slice(0, 64) || null,
      groupName: String(body.groupName || "").slice(0, 120) || null,
      createdAt: new Date().toISOString(),
      createdBy: String(body.by || "").slice(0, 40) || null,
    };
    cycles.push(cycle);
    cycles.sort((a, b) => a.start.localeCompare(b.start));
    await this.ctx.storage.put("cycles", cycles);

    // Seed the board so its week matches the cycle from the moment it is opened,
    // rather than depending on whoever connects first to set it.
    try {
      const board = this.env.BOARD.get(this.env.BOARD.idFromName(id));
      await board.fetch(new Request("https://do/room/" + id + "/init", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ weekOf: start, groupId: cycle.groupId, groupName: cycle.groupName }),
      }));
    } catch (e) { /* the board will still pick the week up from the record on open */ }

    return json({ ok: true, cycle: { ...cycle, end: cycleEnd(start) } }, 201);
  }
}

export class Board {
  constructor(ctx, env) {
    this.ctx = ctx;
    this.env = env;
  }

  async fetch(request) {
    const url = new URL(request.url);
    const origin = request.headers.get("Origin") || "";
    const cors = corsHeaders(origin);

    // A Durable Object does not know the name it was addressed by, but uids and diffs
    // need it. Remember it, since WebSocket handlers get no URL to read it from.
    const room = (url.pathname.match(/^\/room\/([A-Za-z0-9_-]{1,64})/) || [])[1] || "board";
    if ((await this.ctx.storage.get("room")) !== room) await this.ctx.storage.put("room", room);

    // Called once when a cycle is created. Only ever sets the week on an empty board,
    // so it can never overwrite a board people are already working in.
    if (url.pathname.endsWith("/init")) {
      let body = {};
      try { body = await request.json(); } catch (e) { /* treated as empty */ }
      const existing = (await this.ctx.storage.get("plan")) || null;
      if (!existing && /^\d{4}-\d{2}-\d{2}$/.test(body.weekOf || "")) {
        await this.ctx.storage.put({
          plan: {
            placed: [], custom: [], seq: 1, cats: [], weekOf: body.weekOf,
            groupId: body.groupId || null, groupName: body.groupName || null,
          },
          version: 1,
          updated: new Date().toISOString(),
        });
      }
      return new Response(JSON.stringify({ ok: true }), {
        headers: { ...cors, "Content-Type": "application/json" },
      });
    }

    // The calendar identity of every board event, as of the last time the calendar was
    // read. uid -> Graph event id.
    if (url.pathname.endsWith("/ids")) {
      return new Response(JSON.stringify({
        ok: true,
        ids: (await this.ctx.storage.get("idMap")) || {},
      }), { headers: { ...cors, "Content-Type": "application/json" } });
    }

    // Asks the flow to look at the calendar now. Used by the planner's "Check calendar"
    // button so someone can confirm the schedule is live without waiting for a publish.
    if (url.pathname.endsWith("/check")) {
      if (request.method !== "POST") {
        return new Response("POST required", { status: 405, headers: cors });
      }
      this.ctx.waitUntil(this.notifyFlow(room, new Date().toISOString()));
      return new Response(JSON.stringify({ ok: true, asked: true }), {
        headers: { ...cors, "Content-Type": "application/json" },
      });
    }

    // Erases the room entirely: board, published snapshot, tracking, everything. The
    // calendar is untouched -- purge first if its events should go too.
    if (url.pathname.endsWith("/wipe")) {
      if (request.method !== "POST") {
        return new Response("POST required", { status: 405, headers: cors });
      }
      for (const ws of this.ctx.getWebSockets()) { try { ws.close(); } catch (e) {} }
      await this.ctx.storage.deleteAll();
      return new Response(JSON.stringify({ ok: true, wiped: room }), {
        headers: { ...cors, "Content-Type": "application/json" },
      });
    }

    // Forgets which Graph events belong to this board, without touching the calendar.
    // Used after a purge, so the next sync rebuilds from nothing rather than trying to
    // update events that were deleted underneath it.
    if (url.pathname.endsWith("/reset-tracking")) {
      if (request.method !== "POST") {
        return new Response("POST required", { status: 405, headers: cors });
      }
      const previous = Object.keys((await this.ctx.storage.get("pushed")) || {}).length;
      await this.ctx.storage.put({ pushed: {}, lastSync: null });
      return new Response(JSON.stringify({ ok: true, forgot: previous }), {
        headers: { ...cors, "Content-Type": "application/json" },
      });
    }

    // Everything the calendar should hold for this cycle, plus the window to search.
    // A purge needs the date range and the marker; it cannot rely on `pushed`, because
    // orphaned events are by definition the ones that were never recorded there.
    if (url.pathname.endsWith("/purge-info")) {
      const published = (await this.ctx.storage.get("published")) || null;
      const plan = published || (await this.ctx.storage.get("plan")) || null;
      if (!plan || !plan.weekOf) {
        return new Response(JSON.stringify({ ok: false, error: "No cycle to purge." }), {
          status: 400, headers: { ...cors, "Content-Type": "application/json" },
        });
      }
      return new Response(JSON.stringify({
        ok: true,
        marker: "cycle-planner-" + room + "-",
        // A day either side so timezone handling at the boundaries cannot hide an event.
        start: addDaysISO(plan.weekOf, -1) + "T00:00:00",
        end: addDaysISO(cycleEnd(plan.weekOf), 2) + "T00:00:00",
        // Which group this cycle publishes to. Null means nobody has chosen one yet, and
        // the flow should refresh the group list and stop rather than guess a calendar.
        groupId: plan.groupId || null,
        groupName: plan.groupName || null,
      }), { headers: { ...cors, "Content-Type": "application/json" } });
    }

    // The flow sends what is currently on the calendar; this answers with what to change.
    // Self-correcting: a half-finished previous run changes nothing about the answer.
    if (url.pathname.endsWith("/reconcile")) {
      if (request.method !== "POST") {
        return new Response("POST required", { status: 405, headers: cors });
      }
      let body;
      try { body = await request.json(); } catch (e) {
        return new Response(JSON.stringify({ ok: false, error: "Body was not JSON." }), {
          status: 400, headers: { ...cors, "Content-Type": "application/json" },
        });
      }
      const published = (await this.ctx.storage.get("published")) || null;
      const result = reconcile(published, room, body.events || body.value || []);

      // Remember what the calendar actually looked like. This is the only moment anything
      // observes it, so it is the only basis for telling someone whether their schedule
      // is really live rather than merely published.
      await this.ctx.storage.put("lastReconcile", {
        at: new Date().toISOString(),
        ok: result.ok !== false,
        error: result.error || null,
        inspected: result.inspected || 0,
        ours: result.ours || 0,
        foreign: result.foreign || 0,
        create: (result.create || []).length,
        update: (result.update || []).length,
        delete: (result.delete || []).length,
        changes: result.changes || 0,
        tracked: Object.keys(result.known || {}).length,
      });
      // The uid -> calendar-event-id map, refreshed on every look at the calendar.
      // This is what lets a future edit target its exact event instead of searching.
      if (result.known) await this.ctx.storage.put("idMap", result.known);
      this.ctx.waitUntil(this.broadcastPublishState(room));

      return new Response(JSON.stringify(result), {
        headers: { ...cors, "Content-Type": "application/json" },
      });
    }

    if (url.pathname.endsWith("/sync-plan")) {
      const published = (await this.ctx.storage.get("published")) || null;
      const pushed = (await this.ctx.storage.get("pushed")) || {};
      const plan = buildSyncPlan(published, room, pushed);
      const publishedAt = (await this.ctx.storage.get("publishedAt")) || null;
      const lastNotify = (await this.ctx.storage.get("lastNotify")) || null;
      return new Response(JSON.stringify({ ...plan, publishedAt, lastNotify }), {
        headers: { ...cors, "Content-Type": "application/json" },
      });
    }

    // The flow reports back what Graph actually accepted. Recording it only after the
    // fact means a failed or half-finished run just retries next time instead of
    // leaving the Worker believing in events that were never created.
    if (url.pathname.endsWith("/sync-ack")) {
      if (request.method !== "POST") {
        return new Response("POST required", { status: 405, headers: cors });
      }
      let body;
      try { body = await request.json(); } catch (e) {
        return new Response(JSON.stringify({ ok: false, error: "Body was not JSON." }), {
          status: 400, headers: { ...cors, "Content-Type": "application/json" },
        });
      }

      const pushed = (await this.ctx.storage.get("pushed")) || {};
      let created = 0, updated = 0, removed = 0;

      for (const row of Array.isArray(body.created) ? body.created : []) {
        if (!row || !row.uid || !row.eventId) continue;
        pushed[row.uid] = { graphId: String(row.eventId), hash: String(row.hash || "") };
        created++;
      }
      for (const row of Array.isArray(body.updated) ? body.updated : []) {
        if (!row || !row.uid || !pushed[row.uid]) continue;
        pushed[row.uid].hash = String(row.hash || "");
        if (row.eventId) pushed[row.uid].graphId = String(row.eventId);
        updated++;
      }
      for (const row of Array.isArray(body.deleted) ? body.deleted : []) {
        const uid = row && (row.uid || row);
        if (uid && pushed[uid]) { delete pushed[uid]; removed++; }
      }

      await this.ctx.storage.put({ pushed, lastSync: new Date().toISOString() });
      return new Response(JSON.stringify({ ok: true, created, updated, deleted: removed, tracked: Object.keys(pushed).length }), {
        headers: { ...cors, "Content-Type": "application/json" },
      });
    }

    // Preview of the *live* board. Useful for checking what publishing would send;
    // the calendar itself is driven by /sync-plan, which reads the published snapshot.
    if (url.pathname.endsWith("/events")) {
      const snapshot = await this.snapshot();
      return new Response(JSON.stringify({ ...toEvents(snapshot.plan, room), version: snapshot.version }), {
        headers: { ...cors, "Content-Type": "application/json" },
      });
    }

    if (url.pathname.endsWith("/plan")) {
      const snapshot = await this.snapshot();
      return new Response(JSON.stringify(snapshot), {
        headers: { ...cors, "Content-Type": "application/json" },
      });
    }

    if (request.headers.get("Upgrade") !== "websocket") {
      return new Response("Expected a WebSocket upgrade", { status: 426, headers: cors });
    }

    const pair = new WebSocketPair();
    const [client, server] = Object.values(pair);

    // Hibernation API: the Durable Object can be evicted between messages while the
    // sockets stay open, so idle rooms cost nothing against the free tier's
    // duration budget. Handlers below are called when it wakes.
    this.ctx.acceptWebSocket(server);

    const snapshot = await this.snapshot();
    const publish = await this.publishState(room);
    server.send(JSON.stringify({ type: "snapshot", ...snapshot, publish, peers: this.peerCount() }));
    this.broadcastPeers();

    return new Response(null, { status: 101, webSocket: client });
  }

  async snapshot() {
    const plan = (await this.ctx.storage.get("plan")) || null;
    const version = (await this.ctx.storage.get("version")) || 0;
    const updated = (await this.ctx.storage.get("updated")) || null;
    return { plan, version, updated };
  }

  // What the calendar currently reflects, plus how far the live board has moved past it.
  async publishState(room) {
    const plan = (await this.ctx.storage.get("plan")) || null;
    const published = (await this.ctx.storage.get("published")) || null;
    const publishedAt = (await this.ctx.storage.get("publishedAt")) || null;
    return {
      publishedAt,
      everPublished: !!published,
      pending: pendingChanges(plan, published, room || "board"),
      groupName: (plan && plan.groupName) || null,
      groupChosen: !!(plan && plan.groupId),
      // What the calendar itself last looked like, and when.
      calendar: (await this.ctx.storage.get("lastReconcile")) || null,
      lastNotify: (await this.ctx.storage.get("lastNotify")) || null,
    };
  }

  /**
   * Tells the Power Automate flow to run right now.
   *
   * PUBLISH_WEBHOOK is the flow's "When an HTTP request is received" trigger URL. That
   * URL carries its own signature and is a credential, so it lives in a secret rather
   * than the repo. Unset simply means no instant push -- the scheduled run still works.
   */
  async notifyFlow(room, publishedAt) {
    const hook = this.env.PUBLISH_WEBHOOK;
    if (!hook) {
      await this.ctx.storage.put("lastNotify", { at: publishedAt, status: "no webhook configured" });
      return;
    }
    try {
      const res = await fetch(hook, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ room, publishedAt }),
      });
      await this.ctx.storage.put("lastNotify", { at: publishedAt, status: res.status, ok: res.ok });
    } catch (e) {
      await this.ctx.storage.put("lastNotify", { at: publishedAt, status: "failed: " + (e && e.message) });
    }
  }

  async broadcastPublishState(room) {
    const state = await this.publishState(room);
    this.broadcast({ type: "publish-state", ...state });
    return state;
  }

  peerCount() {
    return this.ctx.getWebSockets().length;
  }

  broadcast(payload, except) {
    const text = JSON.stringify(payload);
    for (const ws of this.ctx.getWebSockets()) {
      if (ws === except) continue;
      try { ws.send(text); } catch (e) { /* socket already gone; close handler cleans up */ }
    }
  }

  broadcastPeers() {
    this.broadcast({ type: "peers", peers: this.peerCount() });
  }

  async webSocketMessage(ws, raw) {
    let msg;
    try { msg = JSON.parse(raw); } catch (e) {
      return ws.send(JSON.stringify({ type: "error", message: "Malformed message." }));
    }

    if (msg.type === "ping") return ws.send(JSON.stringify({ type: "pong" }));

    // Promote the live board to the snapshot the calendar follows. Explicit and
    // deliberate: nothing anyone drags reaches a calendar until someone does this.
    if (msg.type === "publish") {
      const room = (await this.ctx.storage.get("room")) || "board";
      const plan = (await this.ctx.storage.get("plan")) || null;

      if (!plan || !plan.weekOf) {
        return ws.send(JSON.stringify({
          type: "publish-result",
          ok: false,
          error: "Pick the Monday this week starts before publishing to the calendar.",
        }));
      }

      const publishedAt = new Date().toISOString();
      // Deep copy so later edits to the live board cannot mutate what was published.
      await this.ctx.storage.put({ published: JSON.parse(JSON.stringify(plan)), publishedAt });

      // Poke the flow so the calendar updates now instead of at the next scheduled run.
      // Deliberately not awaited: publishing must not hang or fail because Power
      // Automate is slow or down. If this never lands, the flow's scheduled run is the
      // backstop and the calendar simply catches up a few minutes later.
      this.ctx.waitUntil(this.notifyFlow(room, publishedAt));

      ws.send(JSON.stringify({ type: "publish-result", ok: true, publishedAt, by: msg.by || null }));
      const state = await this.broadcastPublishState(room);
      return state;
    }

    if (msg.type !== "update") return;

    if (!msg.plan || typeof msg.plan !== "object" || !Array.isArray(msg.plan.placed)) {
      return ws.send(JSON.stringify({ type: "error", message: "Update did not contain a board." }));
    }

    const encoded = JSON.stringify(msg.plan);
    if (encoded.length > MAX_PLAN_BYTES) {
      return ws.send(JSON.stringify({ type: "error", message: "That board is too large to sync." }));
    }

    // Last write wins. With a handful of planners this is fine, but a client that
    // was editing an older version needs to know its base was superseded, so the
    // version travels back with every sync and stale senders get told.
    const current = (await this.ctx.storage.get("version")) || 0;
    const version = current + 1;
    const updated = new Date().toISOString();

    await this.ctx.storage.put({ plan: msg.plan, version, updated });

    const stale = typeof msg.base === "number" && msg.base < current;
    ws.send(JSON.stringify({ type: "ack", version, updated, overwrote: stale }));
    this.broadcast({ type: "sync", plan: msg.plan, version, updated, by: msg.by || null }, ws);

    // Every edit moves the board further from what the calendar shows; keep the
    // pending count in front of everyone rather than letting drift go unnoticed.
    const room = (await this.ctx.storage.get("room")) || "board";
    await this.broadcastPublishState(room);
  }

  async webSocketClose(ws) {
    try { ws.close(); } catch (e) { /* already closed */ }
    this.broadcastPeers();
  }

  async webSocketError(ws) {
    this.broadcastPeers();
  }
}
