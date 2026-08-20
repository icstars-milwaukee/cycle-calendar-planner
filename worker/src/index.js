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
  // The planner's home since it joined the hub. The old GitHub Pages origin
  // stays allowed so anyone on a cached copy of the page keeps syncing.
  "https://icstars-tech-inventory.pages.dev",
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

// Wednesday and Friday run virtual (mirrors DAY_MODE in the planner). Workshops on
// these days become Teams online meetings so every session carries a Join button.
// From week 3 onward: week 1 is Team Week and week 2 was drafted before this rule.
const VIRTUAL_DAYS = [2, 4];
const TEAMS_FROM_WEEK = 3;

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
  { id: "h1", title: "All Skate",         start: 510,  end: 540 },
  { id: "h2", title: "Lunch",             start: 720,  end: 780 },
  { id: "h3", title: "Wellness Workshop", start: 780,  end: 840,  days: [2] },
  { id: "h4", title: "Out of cycle",      start: 840,  end: 1155, days: [2] },
  { id: "h5", title: "Tea Prep",          start: 930,  end: 960,  days: [0, 1, 3, 4] },
  { id: "h6", title: "High Tea",          start: 960,  end: 1020, days: [0, 1, 3, 4] },
  { id: "h7", title: "Tea Debrief",       start: 1020, end: 1050, days: [0, 1, 3, 4] },
  { id: "h8", title: "Break",             start: 1050, end: 1080, days: [0, 1, 3, 4] },
];

// Holds live on the board now, so each cycle can shape its own week; the constant
// above only covers boards that have never edited theirs.
function holdsForDay(plan, day) {
  const list = Array.isArray(plan.holds) && plan.holds.length ? plan.holds : HOLDS;
  return list.filter((h) => h && h.title && h.end > h.start &&
    (!h.days || h.days.indexOf(day) >= 0));
}

function slug(s) {
  return String(s).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
}

function hhmm(minutesOfDay) {
  const m = Math.max(0, Math.min(24 * 60 - 1, Math.round(minutesOfDay)));
  return pad2(Math.floor(m / 60)) + ":" + pad2(m % 60) + ":00";
}

// Facilitators for a block: its own assignment, or the shelf card's if it has none.
// Normalised to lowercase addresses, deduped, capped -- these become real invitations.
function assignedTo(p, plan) {
  let list = Array.isArray(p.who) ? p.who : null;
  if (!list || !list.length) {
    const card = p.refId ? (plan.custom || []).find((c) => c.id === p.refId) : null;
    list = card && Array.isArray(card.who) ? card.who : [];
  }
  const out = [];
  for (const raw of list) {
    const e = String(raw || "").trim().toLowerCase();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) continue;
    if (out.indexOf(e) < 0) out.push(e);
    if (out.length >= 10) break;
  }
  return out;
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

  // Weeks that opt out of the daily skeleton entirely -- a team week or off-site is
  // one banner, not All Skate and tea service. Also all in person: no Teams links.
  // 1-based week numbers.
  const skipHolds = new Set(Array.isArray(plan.holdSkipWeeks) ? plan.holdSkipWeeks : []);

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
        // Per-event choice wins; the day rule is only the default. undefined follows
        // the rule (virtual on Wed/Fri from week 3, never on a skeleton-free week);
        // true forces a Teams link anywhere; false forces in person anywhere.
        teams: p.virtual === true ? true : p.virtual === false ? false :
          (w + 1 >= TEAMS_FROM_WEEK && VIRTUAL_DAYS.indexOf(p.day) >= 0 && !skipHolds.has(w + 1)),
        notes: typeof p.notes === "string" ? p.notes.slice(0, 4000) : "",
        // Facilitators become real attendees, which is what puts the session on their
        // own calendar. Inherited from the shelf card when the block carries none.
        who: assignedTo(p, plan),
        // Curriculum code, if this workshop has been mapped to one.
        code: typeof p.code === "string" ? p.code.slice(0, 16) : "",
      });
    }
  }

  // Week-long banners: one all-day event spanning Monday to Friday of the given week.
  for (const we of Array.isArray(plan.weekEvents) ? plan.weekEvents : []) {
    const wk = we && we.week | 0;
    if (wk < 1 || wk > CYCLE_WEEKS || !we.title) continue;
    const mon = addDaysISO(plan.weekOf, (wk - 1) * 7);
    events.push({
      uid: prefix + "weekev-w" + wk + "-" + slug(we.title),
      subject: String(we.title),
      category: we.cat || "Team Week",
      isAllDay: true,
      start: { dateTime: mon + "T00:00:00", timeZone: CAL_TZ },
      // All-day ranges are half-open: ending Saturday midnight covers Mon-Fri.
      end: { dateTime: addDaysISO(mon, 5) + "T00:00:00", timeZone: CAL_TZ },
      week: wk,
    });
  }

  // The fixed skeleton, unless someone has turned it off. Emitted per day so a holiday
  // that closes a day removes its rituals too, not just its workshops.
  if (plan.includeHolds !== false) {
    // One-off exceptions: "holdId:week:day" keys name single occurrences someone has
    // skipped. The mapped hold stays; only that occurrence's event disappears.
    const holdSkips = new Set(Array.isArray(plan.holdSkips) ? plan.holdSkips : []);
    // Per-occurrence details, same keying: an override wins over the hold's own notes.
    const holdNotes = plan.holdNotes && typeof plan.holdNotes === "object" ? plan.holdNotes : {};
    // Per-occurrence titles, same keying again: "Lunch" can be "Pizza Friday" on one
    // day. The uid stays anchored to the hold itself, so a renamed occurrence updates
    // its existing calendar event in place instead of duplicating it.
    const holdTitles = plan.holdTitles && typeof plan.holdTitles === "object" ? plan.holdTitles : {};
    for (let w = 0; w < CYCLE_WEEKS; w++) {
      if (skipHolds.has(w + 1)) continue;
      for (let d = 0; d < 5; d++) {
        const date = addDaysISO(plan.weekOf, w * 7 + d);
        if (!date || blocked.has(date)) continue;
        for (const h of holdsForDay(plan, d)) {
          if (holdSkips.has((h.id || slug(h.title)) + ":" + (w + 1) + ":" + d)) continue;
          // Default holds (ids h1-h8) keep the slug-of-title anchor their existing
          // calendar events were created under, so this change churns nothing.
          // Custom holds anchor to their unique id, so retiming or renaming one
          // updates its events in place rather than duplicating.
          const anchor = (!h.id || /^h\d$/.test(h.id)) ? slug(h.title) : h.id;
          const titleOv = holdTitles[(h.id || slug(h.title)) + ":" + (w + 1) + ":" + d];
          events.push({
            uid: prefix + "hold-" + anchor + "-w" + (w + 1) + "-d" + d,
            subject: (typeof titleOv === "string" && titleOv.trim())
              ? titleOv.trim().slice(0, 120) : h.title,
            category: "Fixed hold",
            start: { dateTime: date + "T" + hhmm(h.start), timeZone: CAL_TZ },
            end: { dateTime: date + "T" + hhmm(h.end), timeZone: CAL_TZ },
            week: w + 1,
            day: d,
            minutes: h.end - h.start,
            isHold: true,
            notes: (() => {
              const ov = holdNotes[(h.id || slug(h.title)) + ":" + (w + 1) + ":" + d];
              const t = typeof ov === "string" ? ov : (typeof h.notes === "string" ? h.notes : "");
              return t.slice(0, 4000);
            })(),
          });
        }
      }
    }
  }

  // Every program day carries a banner: the week number and whether the day
  // is in person or virtual. All-day events pin to the top of the day in
  // Outlook, so the hybrid schedule is the first thing on every date - the
  // person who shows up on a virtual day at least did it against a labeled
  // calendar. Same rule the per-event Teams flag uses: Wed/Fri are virtual
  // from week 3, and skeleton-free (team) weeks are all in person.
  for (let w = 0; w < CYCLE_WEEKS; w++) {
    for (let d = 0; d < 5; d++) {
      const date = addDaysISO(plan.weekOf, w * 7 + d);
      if (!date || blocked.has(date)) continue;
      const virtual = (w + 1) >= TEAMS_FROM_WEEK && VIRTUAL_DAYS.indexOf(d) >= 0 && !skipHolds.has(w + 1);
      events.push({
        uid: prefix + "daymode-w" + (w + 1) + "-d" + d,
        subject: "Week " + (w + 1) + " · " + (virtual ? "Virtual Day" : "In-Person Day"),
        category: "Schedule",
        isAllDay: true,
        start: { dateTime: date + "T00:00:00", timeZone: CAL_TZ },
        end: { dateTime: addDaysISO(date, 1) + "T00:00:00", timeZone: CAL_TZ },
      });
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
    // Virtual-day sessions are Teams meetings: Graph generates the join link when an
    // event is created with these set. They only apply reliably at creation, which is
    // why converting existing events is done by re-creating them, not patching.
    isOnlineMeeting: !!ev.teams,
    onlineMeetingProvider: ev.teams ? "teamsForBusiness" : "unknown",
    // The planner category rides as a real Outlook category, not just body text.
    // Outlook maps the NAME to a color from the calendar's own category list, which
    // Graph does not expose for group mailboxes -- so someone defines each name's
    // color once in Outlook, and every event wearing that name follows it.
    categories: ev.category ? [ev.category] : [],
    // Assigned facilitators are invited, which is what puts the session on their own
    // calendar. Graph sends the invitation from the group mailbox on create, and on
    // update to anyone newly added.
    attendees: (ev.who || []).map((a) => ({
      emailAddress: { address: a, name: a.split("@")[0] },
      type: "required",
    })),
    // Graph uses transactionId to make event creation idempotent: re-POSTing the same
    // one does not create a second event. Without it, any run that creates events and
    // then dies before reporting back duplicates all of them on the next attempt --
    // which is not a rare edge case when 284 events meet a 100-call-per-minute limit.
    transactionId: ev.uid,
    body: {
      contentType: "text",
      // The team's own details lead; the metadata trails. The nh line is a fingerprint
      // of the details so the reconciler can tell when they changed on the board and
      // update the event body in place -- comparing full bodies is unreliable because
      // Outlook rewrites them.
      content: (ev.notes ? ev.notes + "\n\n----\n" : "") +
        "Cycle Calendar Planner" + (ev.category ? " - " + ev.category : "") +
        (ev.week ? "\nWeek " + ev.week + " of " + CYCLE_WEEKS : "") +
        (ev.code ? "\nCurriculum: " + ev.code : "") +
        ((ev.who || []).length ? "\nFacilitator: " + ev.who.join(", ") : "") +
        "\nDo not edit here; edit the planner and it will be overwritten on the next sync." +
        "\nref: " + ev.uid +
        "\nnh: " + notesHash(ev.notes || "") +
        // Fingerprint for the fields the flow's calendarView listing does not return
        // -- attendees (the same gap that hides category drift) and the curriculum
        // code. Reassigning or remapping is detected through the body, which it does.
        "\nas: " + extrasHash(ev),
    },
  };
}

function extrasHash(ev) {
  return notesHash((ev.who || []).join(",") + "|" + (ev.code || ""));
}

function notesHash(s) {
  let h = 5381;
  for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36);
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
  const foreignSamples = [];
  for (const ev of Array.isArray(existingRaw) ? existingRaw : []) {
    const uid = uidOfExisting(ev, marker);
    if (!uid) {
      foreign++;   // not ours -- never touched
      if (foreignSamples.length < 40) {
        foreignSamples.push({ subject: String(ev.subject || "").slice(0, 80),
          start: String((ev.start && ev.start.dateTime) || "").slice(0, 16) });
      }
      continue;
    }
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
    // Details drift: compare the fingerprint stamped into the body against what the
    // board wants now. An event with no fingerprint (created before notes existed)
    // only updates once someone actually writes details for it.
    const bodyText = String((keep.body && keep.body.content) || keep.bodyPreview || "");
    const foundNh = (/nh:\s*([a-z0-9]+)/.exec(bodyText) || [])[1] || notesHash("");
    const sameNotes = foundNh === notesHash(ev.notes || "");
    // Reassignment or a curriculum remap: same fingerprint trick. An event stamped
    // before either existed carries no as: line and reads as unassigned and unmapped,
    // so it only updates once someone actually assigns or maps it.
    const foundAs = (/as:\s*([a-z0-9]+)/.exec(bodyText) || [])[1] || extrasHash({});
    const sameWho = foundAs === extrasHash(ev);
    if (!sameSubject || !sameStart || !sameEnd || !sameCat || !sameNotes || !sameWho) {
      // Online-meeting fields are creation-only in Graph; patching them onto an
      // existing event can fail the whole update, so they stay out of PATCH bodies.
      const body = graphBody(ev);
      delete body.isOnlineMeeting;
      delete body.onlineMeetingProvider;
      update.push({ uid, eventId: keep.id, event: body });
    }
  }

  // Ours, in the window, but no longer part of the cycle.
  for (const [uid, matches] of seen) {
    if (desired.has(uid)) continue;
    for (const ev of matches) remove.push({ uid, eventId: ev.id, reason: "no longer scheduled" });
  }

  // ONE-TIME CLEANUP, off by default. Deletes foreign events that were created by this
  // system under a DIFFERENT room's marker -- strays left by test rooms in the era when
  // the flow synced any room to the real calendar. DANGER: if two live cycles ever
  // publish to the same group calendar, this flag would make one delete the other's
  // events. It must be set for a single run and then removed, never left on.
  if (publishedPlan.cleanPlannerStrays === true) {
    for (const ev of Array.isArray(existingRaw) ? existingRaw : []) {
      if (uidOfExisting(ev, marker)) continue;               // ours -- handled above
      const anyPlanner = uidOfExisting(ev, "cycle-planner-"); // made by this system, other room
      if (anyPlanner) remove.push({ uid: anyPlanner, eventId: ev.id, reason: "test stray" });
    }
  }

  return {
    ok: true,
    weekOf: projection.weekOf,
    marker,
    inspected: (existingRaw || []).length,
    ours: seen.size,
    foreign,
    known,
    foreignSamples,
    create,
    update,
    delete: remove,
    changes: create.length + update.length + remove.length,
  };
}

/**
 * Login identity. Everyone signs in with the team password AND their work email; the
 * email is validated against an allowlist of domains and bound server-side to the
 * connection, so every decision in the log carries who made it.
 *
 * Honest limit: this asserts identity, it does not prove mailbox ownership -- sending
 * verification codes needs a sending domain this Cloudflare account does not have.
 * The shape is ready for that upgrade without changing anything downstream.
 */
function emailOk(email, env) {
  const e = String(email || "").trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e) || e.length > 120) return null;
  const allowed = String(env.ALLOWED_EMAIL_DOMAINS || "icstars.org")
    .toLowerCase().split(",").map((d) => d.trim()).filter(Boolean);
  const domain = e.split("@")[1];
  return allowed.indexOf(domain) >= 0 ? e : null;
}

// Admins are the only people who can delete anything -- cycles, boards, blocks, holds,
// categories. Two tiers: founders named in the ADMIN_EMAILS secret, who cannot be
// removed from the app, and admins appointed by other admins, stored in the Directory.
// Same honesty caveat as login: identity is asserted, not proven.
function isFounder(email, env) {
  const e = String(email || "").trim().toLowerCase();
  if (!e) return false;
  return String(env.ADMIN_EMAILS || "").toLowerCase()
    .split(",").map((x) => x.trim()).filter(Boolean).indexOf(e) >= 0;
}

// Founders plus appointed admins. One subrequest to the Directory; callers on hot
// paths cache the result.
async function isAdminFull(email, env) {
  if (isFounder(email, env)) return true;
  const e = String(email || "").trim().toLowerCase();
  if (!e) return false;
  try {
    const dir = env.CYCLES.get(env.CYCLES.idFromName("directory"));
    const res = await dir.fetch(new Request("https://do/admins", { method: "GET" }));
    const j = await res.json();
    return (j.admins || []).some((a) => a.email === e);
  } catch (err) { return false; }
}

/**
 * Does the incoming board LOSE anything the current board has? Replacements do not
 * count: a block whose uid vanished but whose title/day/time survives in the same week
 * is a re-identification (format flips, tooling), not a deletion.
 */
function detectLoss(current, incoming) {
  if (!current) return null;
  const weeksOf = (p) => Array.isArray(p.weeks) && p.weeks.length
    ? p.weeks : [Array.isArray(p.placed) ? p.placed : []];
  const curW = weeksOf(current), incW = weeksOf(incoming);

  for (let i = 0; i < curW.length; i++) {
    const inc = Array.isArray(incW[i]) ? incW[i] : [];
    const incUids = new Set(inc.map((b) => b.uid));
    for (const b of curW[i] || []) {
      if (incUids.has(b.uid)) continue;
      const replaced = inc.some((x) => x.title === b.title && x.day === b.day &&
        x.start === b.start && x.mins === b.mins);
      if (!replaced) return "the workshop “" + (b.title || "?") + "” in week " + (i + 1);
    }
  }
  if (Array.isArray(current.holds)) {
    const incIds = new Set((incoming.holds || []).map((h) => h.id));
    for (const h of current.holds) {
      if (!incIds.has(h.id)) return "the hold “" + h.title + "”";
    }
  }
  // Skipping a hold occurrence removes one calendar event, so it is a deletion too.
  {
    const curSkips = new Set(Array.isArray(current.holdSkips) ? current.holdSkips : []);
    for (const k of Array.isArray(incoming.holdSkips) ? incoming.holdSkips : []) {
      if (!curSkips.has(k)) return "a hold occurrence (week " + (String(k).split(":")[1] || "?") + ")";
    }
  }
  const incCats = new Set((incoming.cats || []).map((c) => c.id));
  for (const c of current.cats || []) {
    if (!incCats.has(c.id)) return "the category “" + c.name + "”";
  }
  return null;
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
      // If an email was offered, it must be a valid one -- the UI treats this as the
      // login, so a typo'd or off-domain address is rejected here, before any socket.
      const offered = url.searchParams.get("user");
      if (offered !== null) {
        const user = emailOk(offered, env);
        if (!user) {
          return new Response(JSON.stringify({ ok: false, error: "Use your work email address." }), {
            status: 403, headers: { ...cors, "Content-Type": "application/json" },
          });
        }
        return new Response(JSON.stringify({ ok: true, user, admin: await isAdminFull(user, env) }), {
          headers: { ...cors, "Content-Type": "application/json" },
        });
      }
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
    if (url.pathname === "/cycles" || url.pathname === "/groups" || url.pathname === "/admins") {
      if (!authorized(request, env)) return deny(cors);
      return env.CYCLES.get(env.CYCLES.idFromName("directory")).fetch(request);
    }

    // /room/<name>        -> WebSocket upgrade for live editing
    // /room/<name>/plan   -> plain GET snapshot, handy for debugging and backups
    const match = url.pathname.match(
      /^\/room\/([A-Za-z0-9_-]{1,64})(\/plan|\/events|\/sync-plan|\/sync-ack|\/init|\/reset-tracking|\/purge-info|\/reconcile|\/check|\/wipe|\/ids|\/log|\/tamper-test|\/backups|\/restore-backup)?$/);
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

    // Admin management. Founders come from the ADMIN_EMAILS secret and cannot be
    // touched here; appointed admins are stored and managed by any current admin.
    // Internal calls (hostname "do") may read the list; external reads need an admin.
    if (new URL(request.url).pathname === "/admins") {
      const aurl = new URL(request.url);
      const stored = (await this.ctx.storage.get("admins")) || [];
      const internal = aurl.hostname === "do";
      const caller = String(aurl.searchParams.get("user") || "").trim().toLowerCase();
      const callerIsAdmin = isFounder(caller, this.env) || stored.indexOf(caller) >= 0;

      if (request.method === "GET") {
        if (!internal && !callerIsAdmin) return json({ ok: false, error: "Admins only." }, 403);
        const founders = String(this.env.ADMIN_EMAILS || "").toLowerCase()
          .split(",").map((x) => x.trim()).filter(Boolean);
        return json({
          ok: true,
          admins: founders.map((e) => ({ email: e, fixed: true }))
            .concat(stored.map((e) => ({ email: e, fixed: false }))),
        });
      }
      if (!callerIsAdmin) return json({ ok: false, error: "Only an admin can manage admins." }, 403);

      if (request.method === "POST") {
        let body;
        try { body = await request.json(); } catch (e) { return json({ ok: false, error: "Body was not JSON." }, 400); }
        const email = emailOk(body.email, this.env);
        if (!email) return json({ ok: false, error: "Use a work email address on an allowed domain." }, 400);
        if (isFounder(email, this.env) || stored.indexOf(email) >= 0) {
          return json({ ok: false, error: email + " is already an admin." }, 409);
        }
        stored.push(email);
        await this.ctx.storage.put("admins", stored);
        return json({ ok: true, added: email, by: caller });
      }
      if (request.method === "DELETE") {
        const email = String(aurl.searchParams.get("email") || "").trim().toLowerCase();
        if (isFounder(email, this.env)) {
          return json({ ok: false, error: "Founder admins can only be changed in the worker configuration." }, 400);
        }
        const i = stored.indexOf(email);
        if (i < 0) return json({ ok: false, error: "Not an appointed admin." }, 404);
        stored.splice(i, 1);
        await this.ctx.storage.put("admins", stored);
        return json({ ok: true, removed: email, by: caller });
      }
      return json({ ok: false, error: "Unsupported method." }, 405);
    }

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

    // Retiring a cycle removes it from the list and erases its board. Admin only.
    if (request.method === "DELETE") {
      const delCaller = String(new URL(request.url).searchParams.get("user") || "").trim().toLowerCase();
      const delStored = (await this.ctx.storage.get("admins")) || [];
      if (!isFounder(delCaller, this.env) && delStored.indexOf(delCaller) < 0) {
        return json({ ok: false, error: "Only an admin can remove a cycle." }, 403);
      }
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

    // A cycle is tied to its group calendar at creation -- setup is the one moment the
    // choice is cheap. Without this, boards drifted along publishing nowhere until
    // someone discovered the setting.
    if (!body.groupId || !body.groupName) {
      return json({ ok: false, error: "Choose the group calendar this cycle publishes to." }, 400);
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

    // Snapshot history for this room. Admin-only, like everything that can rewind.
    if (url.pathname.endsWith("/backups")) {
      if (!(await this.checkAdmin(url.searchParams.get("user")))) {
        return new Response(JSON.stringify({ ok: false, error: "Admins only." }), {
          status: 403, headers: { ...cors, "Content-Type": "application/json" },
        });
      }
      const listed = this.env.BACKUPS ? await this.env.BACKUPS.list({ prefix: "bk:" + room + ":" }) : { keys: [] };
      const backups = [];
      for (const k of listed.keys.slice(-25).reverse()) {
        try {
          const v = JSON.parse(await this.env.BACKUPS.get(k.name));
          backups.push({ key: k.name, at: v.at, reason: v.reason, blocks: v.blocks });
        } catch (e) { backups.push({ key: k.name }); }
      }
      return new Response(JSON.stringify({ ok: true, backups }), {
        headers: { ...cors, "Content-Type": "application/json" },
      });
    }

    // Puts a snapshot back as the LIVE board. Deliberately does not publish: the
    // restored board is reviewed on screen, then published like any other change.
    if (url.pathname.endsWith("/restore-backup")) {
      if (request.method !== "POST") {
        return new Response("POST required", { status: 405, headers: cors });
      }
      if (!(await this.checkAdmin(url.searchParams.get("user")))) {
        return new Response(JSON.stringify({ ok: false, error: "Only an admin can restore a backup." }), {
          status: 403, headers: { ...cors, "Content-Type": "application/json" },
        });
      }
      let body;
      try { body = await request.json(); } catch (e) {
        return new Response(JSON.stringify({ ok: false, error: "Body was not JSON." }), {
          status: 400, headers: { ...cors, "Content-Type": "application/json" },
        });
      }
      const key = String(body.key || "");
      if (!key.startsWith("bk:" + room + ":")) {
        return new Response(JSON.stringify({ ok: false, error: "That backup belongs to a different cycle." }), {
          status: 400, headers: { ...cors, "Content-Type": "application/json" },
        });
      }
      const raw = this.env.BACKUPS ? await this.env.BACKUPS.get(key) : null;
      if (!raw) {
        return new Response(JSON.stringify({ ok: false, error: "No such backup." }), {
          status: 404, headers: { ...cors, "Content-Type": "application/json" },
        });
      }
      // Snapshot the current board first, so a restore is itself reversible.
      await this.saveBackup(room, "pre-restore");
      const snap = JSON.parse(raw);
      const version = ((await this.ctx.storage.get("version")) || 0) + 1;
      await this.ctx.storage.put({ plan: snap.plan, version, updated: new Date().toISOString() });
      this.broadcast({ type: "sync", plan: snap.plan, version, by: url.searchParams.get("user") });
      await this.appendLog({ t: "restore", by: url.searchParams.get("user"), from: snap.at, blocks: snap.blocks });
      this.ctx.waitUntil(this.broadcastPublishState(room));
      return new Response(JSON.stringify({ ok: true, restored: snap.at, blocks: snap.blocks }), {
        headers: { ...cors, "Content-Type": "application/json" },
      });
    }

    // The room's activity history, newest first.
    if (url.pathname.endsWith("/log")) {
      return new Response(JSON.stringify({
        ok: true,
        log: (await this.ctx.storage.get("log")) || [],
      }), { headers: { ...cors, "Content-Type": "application/json" } });
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
      this.ctx.waitUntil(this.ensureGuard(room));
      return new Response(JSON.stringify({ ok: true, asked: true }), {
        headers: { ...cors, "Content-Type": "application/json" },
      });
    }

    // Erases the room entirely: board, published snapshot, tracking, everything. The
    // calendar is untouched -- purge first if its events should go too. Admin only,
    // except scratch rooms outside the registry, which tests create and destroy freely.
    if (url.pathname.endsWith("/wipe")) {
      if (request.method !== "POST") {
        return new Response("POST required", { status: 405, headers: cors });
      }
      const registered = await this.guardEligibleRegistry(room);
      if (registered && !(await this.checkAdmin(url.searchParams.get("user")))) {
        return new Response(JSON.stringify({ ok: false, error: "Only an admin can wipe a cycle." }), {
          status: 403, headers: { ...cors, "Content-Type": "application/json" },
        });
      }
      for (const ws of this.ctx.getWebSockets()) { try { ws.close(); } catch (e) {} }
      // The last thing that happens before a board dies is a snapshot of it. Wiping,
      // including via cycle deletion, is therefore always reversible from KV.
      await this.saveBackup(room, "pre-wipe");
      await this.ctx.storage.deleteAlarm();
      await this.ctx.storage.deleteAll();
      return new Response(JSON.stringify({ ok: true, wiped: room }), {
        headers: { ...cors, "Content-Type": "application/json" },
      });
    }

    // Guard drill: makes the next sync delete ONE event the board still wants, exactly
    // as if someone removed it in Outlook. The guard's next heartbeat must put it back.
    // Self-neutralizing by design -- the deletion cannot outlive one heartbeat.
    if (url.pathname.endsWith("/tamper-test")) {
      if (request.method !== "POST") {
        return new Response("POST required", { status: 405, headers: cors });
      }
      if (!(await this.checkAdmin(url.searchParams.get("user")))) {
        return new Response(JSON.stringify({ ok: false, error: "Only an admin can run the guard drill." }), {
          status: 403, headers: { ...cors, "Content-Type": "application/json" },
        });
      }
      const ids = (await this.ctx.storage.get("idMap")) || {};
      const uids = Object.keys(ids).filter((u) => !/holiday|hold-|weekev/.test(u));
      if (!uids.length) {
        return new Response(JSON.stringify({ ok: false, error: "No tracked workshop events to test with." }), {
          status: 400, headers: { ...cors, "Content-Type": "application/json" },
        });
      }
      const victim = uids[uids.length - 1];   // a late-cycle session, least disruptive
      await this.ctx.storage.put("tamperOnce", victim);
      this.ctx.waitUntil(this.notifyFlow(room, new Date().toISOString(), true));
      return new Response(JSON.stringify({ ok: true, victim }), {
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
      if (!(await this.checkAdmin(url.searchParams.get("user")))) {
        return new Response(JSON.stringify({ ok: false, error: "Only an admin can reset tracking." }), {
          status: 403, headers: { ...cors, "Content-Type": "application/json" },
        });
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

      // An armed guard drill: inject the deletion of one wanted event into this run's
      // plan, then disarm immediately so it can never fire twice.
      const tamper = await this.ctx.storage.get("tamperOnce");
      if (tamper && result.ok && result.known && result.known[tamper]) {
        result.delete.push({ uid: tamper, eventId: result.known[tamper], reason: "guard drill" });
        result.changes = (result.changes || 0) + 1;
        await this.ctx.storage.delete("tamperOnce");
      }

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
        foreignSamples: result.foreignSamples || [],
      });
      // The uid -> calendar-event-id map, refreshed on every look at the calendar.
      // This is what lets a future edit target its exact event instead of searching.
      if (result.known) await this.ctx.storage.put("idMap", result.known);

      // Log what this run set out to do, with a sample of the actual operations so the
      // history reads as changes, not just numbers. changes=0 is logged too: that is
      // the completion signal for whatever job came before it.
      const opLine = (mark) => (x) => ({
        op: mark,
        s: String((x.event && x.event.subject) || x.uid || "").slice(0, 60),
        d: String((x.event && x.event.start && x.event.start.dateTime) || "").slice(0, 16),
      });
      const ops = []
        .concat((result.create || []).slice(0, 8).map(opLine("+")))
        .concat((result.update || []).slice(0, 8).map(opLine("~")))
        .concat((result.delete || []).slice(0, 8).map((x) => ({ op: "−", s: String(x.uid || "").slice(0, 60), d: x.reason || "" })));
      await this.appendLog({
        t: "sync",
        ok: result.ok !== false,
        error: result.error || null,
        purge: result.purge === true,
        inspected: result.inspected || 0,
        create: (result.create || []).length,
        update: (result.update || []).length,
        delete: (result.delete || []).length,
        changes: result.changes || 0,
        ops,
      });
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

    // Identity rides the connection: offered at login, validated here, and attached to
    // the socket so it survives hibernation. Absent is tolerated (tests, old clients);
    // present-but-invalid is refused.
    const offeredUser = url.searchParams.get("user");
    let connUser = null;
    if (offeredUser !== null && offeredUser !== "") {
      connUser = emailOk(offeredUser, this.env);
      if (!connUser) {
        return new Response("Use your work email address.", { status: 403, headers: cors });
      }
    }

    const pair = new WebSocketPair();
    const [client, server] = Object.values(pair);

    // Hibernation API: the Durable Object can be evicted between messages while the
    // sockets stay open, so idle rooms cost nothing against the free tier's
    // duration budget. Handlers below are called when it wakes.
    this.ctx.acceptWebSocket(server);
    server.serializeAttachment({ u: connUser });

    const snapshot = await this.snapshot();
    const publish = await this.publishState(room);
    server.send(JSON.stringify({
      type: "snapshot", ...snapshot, publish, peers: this.peerCount(), who: this.peerList(),
      youAre: { email: connUser, admin: await this.checkAdmin(connUser) },
    }));
    this.broadcastPeers();
    // Revive the guard on any connection -- covers deploys and stalled alarms.
    this.ctx.waitUntil(this.ensureGuard(room));

    return new Response(null, { status: 101, webSocket: client });
  }

  /**
   * Snapshots. The whole board is written to KV -- a separate storage system from this
   * Durable Object -- on every publish (at most once per 10 minutes), once a day via
   * the guard heartbeat, and unconditionally before anything destroys a board. Sixty
   * snapshots are kept per room; restore is admin-only and never auto-publishes, so a
   * restored board is reviewed before it reaches the calendar.
   */
  async saveBackup(room, reason) {
    if (!this.env.BACKUPS) return;
    const plan = (await this.ctx.storage.get("plan")) || null;
    if (!plan) return;
    const ts = new Date().toISOString();
    const blocks = (Array.isArray(plan.weeks) ? plan.weeks : [plan.placed || []])
      .reduce((n, w) => n + ((w && w.length) || 0), 0);
    try {
      await this.env.BACKUPS.put("bk:" + room + ":" + ts,
        JSON.stringify({ at: ts, room, reason, blocks, plan }));
      await this.ctx.storage.put("lastBackupAt", ts);
      // Prune beyond the newest 60. List returns keys sorted lexicographically, and
      // ISO timestamps sort chronologically, so the head of the list is the oldest.
      const listed = await this.env.BACKUPS.list({ prefix: "bk:" + room + ":" });
      const extra = listed.keys.length - 60;
      for (let i = 0; i < extra; i++) await this.env.BACKUPS.delete(listed.keys[i].name);
    } catch (e) { /* a failed backup must never break the operation it rode on */ }
  }

  async maybeBackup(room, reason, minGapMs) {
    const last = (await this.ctx.storage.get("lastBackupAt")) || "";
    if (last && Date.now() - new Date(last).getTime() < minGapMs) return;
    await this.saveBackup(room, reason);
  }

  // The room's activity log: publishes, sync runs and what they did, pokes and their
  // fate. Bounded so a busy cycle cannot grow storage without limit.
  async appendLog(entry) {
    const log = (await this.ctx.storage.get("log")) || [];
    // The guard checks the calendar every minute; a healthy calendar would bury real
    // history under "nothing to do" lines. Consecutive clean checks collapse into one
    // rolling entry whose timestamp is the latest verification.
    if (entry.t === "sync" && entry.changes === 0 && entry.ok !== false &&
        log[0] && log[0].t === "sync" && log[0].changes === 0 && log[0].ok !== false) {
      log[0].at = new Date().toISOString();
      log[0].inspected = entry.inspected;
      log[0].checks = (log[0].checks || 1) + 1;
      await this.ctx.storage.put("log", log);
      return;
    }
    log.unshift({ at: new Date().toISOString(), ...entry });
    if (log.length > 200) log.length = 200;
    await this.ctx.storage.put("log", log);
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
      // Whether the every-minute self-heal heartbeat is armed for this room.
      guard: (await this.ctx.storage.getAlarm()) !== null,
    };
  }

  /**
   * Tells the Power Automate flow to run right now.
   *
   * PUBLISH_WEBHOOK is the flow's "When an HTTP request is received" trigger URL. That
   * URL carries its own signature and is a credential, so it lives in a secret rather
   * than the repo. Unset simply means no instant push -- the scheduled run still works.
   */
  async notifyFlow(room, publishedAt, quiet) {
    const hook = this.env.PUBLISH_WEBHOOK;
    if (!hook) {
      await this.ctx.storage.put("lastNotify", { at: publishedAt, status: "no webhook configured" });
      if (!quiet) await this.appendLog({ t: "poke", ok: false, status: "no webhook configured", job: publishedAt });
      return;
    }
    try {
      const res = await fetch(hook, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ room, publishedAt }),
      });
      await this.ctx.storage.put("lastNotify", { at: publishedAt, status: res.status, ok: res.ok });
      // Guard heartbeats are quiet: 1,440 successful pokes a day is noise, not history.
      // A failed poke is always worth a line, but only one line per streak of failures.
      if (!quiet) await this.appendLog({ t: "poke", ok: res.ok, status: res.status, job: publishedAt });
      else if (!res.ok) await this.logPokeFailureOnce(res.status);
    } catch (e) {
      await this.ctx.storage.put("lastNotify", { at: publishedAt, status: "failed: " + (e && e.message) });
      if (!quiet) await this.appendLog({ t: "poke", ok: false, status: "failed: " + (e && e.message), job: publishedAt });
      else await this.logPokeFailureOnce("failed: " + (e && e.message));
    }
  }

  async logPokeFailureOnce(status) {
    const log = (await this.ctx.storage.get("log")) || [];
    if (log[0] && log[0].t === "poke" && log[0].ok === false) return;
    await this.appendLog({ t: "poke", ok: false, status, guard: true });
  }

  /**
   * The guard: a permanent 60-second heartbeat (GUARD_SECONDS to tune) that pokes the
   * sync flow for this room. The reconciler puts the calendar back the way the board
   * says, so anything deleted or edited directly in Outlook is undone within about a
   * minute -- enforcement by futility rather than permissions, and it also covers the
   * connection owner, who no permission scheme could restrict.
   *
   * Only rooms that are registered cycles with a group tied and a published board
   * heartbeat; anything else (scratch rooms, tests) never starts one, and a room that
   * stops qualifying stops beating on its next alarm.
   */
  guardInterval() {
    const n = parseInt(this.env.GUARD_SECONDS || "60", 10);
    return Math.max(30, isNaN(n) ? 60 : n) * 1000;
  }

  // Admin lookup for the hot paths: founders answer instantly from the env; appointed
  // admins come from the Directory, cached for a minute so a drag burst costs one
  // subrequest, not one per message. The cost: appointment or removal can take up to
  // a minute to reach an already-open connection.
  async checkAdmin(email) {
    if (isFounder(email, this.env)) return true;
    const e = String(email || "").trim().toLowerCase();
    if (!e) return false;
    const now = Date.now();
    if (!this._adminCache || now - this._adminCache.at > 60000) {
      try {
        const dir = this.env.CYCLES.get(this.env.CYCLES.idFromName("directory"));
        const j = await (await dir.fetch(new Request("https://do/admins", { method: "GET" }))).json();
        this._adminCache = { at: now, set: new Set((j.admins || []).map((a) => a.email)) };
      } catch (err) {
        this._adminCache = { at: now, set: new Set() };
      }
    }
    return this._adminCache.set.has(e);
  }

  async guardEligibleRegistry(room) {
    try {
      const dir = this.env.CYCLES.get(this.env.CYCLES.idFromName("directory"));
      const res = await dir.fetch(new Request("https://do/cycles", { method: "GET" }));
      const j = await res.json();
      return (j.cycles || []).some((c) => c.id === room);
    } catch (e) { return false; }
  }

  async guardEligible(room) {
    const published = (await this.ctx.storage.get("published")) || null;
    if (!published || !published.groupId || published.guard === false) return false;
    return this.guardEligibleRegistry(room);
  }

  async ensureGuard(room) {
    if (!(await this.guardEligible(room))) return;
    if ((await this.ctx.storage.getAlarm()) === null) {
      await this.ctx.storage.setAlarm(Date.now() + this.guardInterval());
    }
  }

  async alarm() {
    const room = (await this.ctx.storage.get("room")) || "board";
    if (!(await this.guardEligible(room))) return;   // stop beating; ensureGuard revives it
    // Reschedule BEFORE working: a failure in the poke must never kill the heartbeat,
    // and the alarm being momentarily unset made the guard flag flicker false.
    await this.ctx.storage.setAlarm(Date.now() + this.guardInterval());
    await this.notifyFlow(room, new Date().toISOString(), true);
    // The daily safety net, riding the heartbeat: even a cycle nobody publishes gets
    // a snapshot a day while it is being planned.
    await this.maybeBackup(room, "daily", 24 * 60 * 60 * 1000);
  }

  async broadcastPublishState(room) {
    const state = await this.publishState(room);
    this.broadcast({ type: "publish-state", ...state });
    return state;
  }

  peerCount() {
    return this.ctx.getWebSockets().length;
  }

  // Who is in the room, from the identity bound to each socket at connect.
  peerList() {
    const who = [];
    for (const ws of this.ctx.getWebSockets()) {
      let u = null;
      try { u = (ws.deserializeAttachment() || {}).u || null; } catch (e) { /* gone */ }
      who.push(u || "guest");
    }
    return who;
  }

  broadcast(payload, except) {
    const text = JSON.stringify(payload);
    for (const ws of this.ctx.getWebSockets()) {
      if (ws === except) continue;
      try { ws.send(text); } catch (e) { /* socket already gone; close handler cleans up */ }
    }
  }

  broadcastPeers() {
    this.broadcast({ type: "peers", peers: this.peerCount(), who: this.peerList() });
  }

  async webSocketMessage(ws, raw) {
    let msg;
    try { msg = JSON.parse(raw); } catch (e) {
      return ws.send(JSON.stringify({ type: "error", message: "Malformed message." }));
    }

    // Server-stamped actor: the email bound at connect wins over anything the client
    // writes into a message, so the log records who was signed in, not who a payload
    // claims to be.
    let actor = null;
    try { actor = (ws.deserializeAttachment() || {}).u || null; } catch (e) {}
    if (actor) msg.by = actor; else if (msg.by) msg.by = String(msg.by).slice(0, 40);

    if (msg.type === "ping") return ws.send(JSON.stringify({ type: "pong" }));

    // Live presence: where someone's drag is hovering right now. Relayed to the other
    // sockets and never stored -- ephemeral by design, so it costs no versions, no log
    // entries, and vanishes with the drag.
    if (msg.type === "cursor") {
      return this.broadcast({
        type: "cursor", by: msg.by || "guest",
        day: msg.day | 0, min: msg.min | 0, mins: msg.mins | 0, wk: msg.wk | 0,
        title: String(msg.title || "").slice(0, 80),
        done: !!msg.done,
      }, ws);
    }

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
      // Mass-deletion flags are admin-only, wherever the plan came from.
      if ((plan.purgeAll === true || plan.cleanPlannerStrays === true) && !(await this.checkAdmin(msg.by))) {
        return ws.send(JSON.stringify({
          type: "publish-result", ok: false,
          error: "Only an admin can publish a purge.",
        }));
      }

      const publishedAt = new Date().toISOString();
      // Deep copy so later edits to the live board cannot mutate what was published.
      await this.ctx.storage.put({ published: JSON.parse(JSON.stringify(plan)), publishedAt });
      await this.appendLog({ t: "publish", by: msg.by || null, job: publishedAt });

      // Poke the flow so the calendar updates now instead of at the next scheduled run.
      // Deliberately not awaited: publishing must not hang or fail because Power
      // Automate is slow or down. If this never lands, the guard heartbeat repairs it
      // within about a minute.
      this.ctx.waitUntil(this.notifyFlow(room, publishedAt));
      this.ctx.waitUntil(this.ensureGuard(room));
      this.ctx.waitUntil(this.maybeBackup(room, "publish", 10 * 60 * 1000));

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

    // Only admins delete. Anyone can add, move, retime, recategorize, edit details --
    // but an update that LOSES a workshop, hold or category is refused unless the
    // signed-in email is an admin, and the client is handed the authoritative board
    // back so its optimistic local copy reverts.
    if (!(await this.checkAdmin(actor))) {
      const currentPlan = (await this.ctx.storage.get("plan")) || null;
      const lost = detectLoss(currentPlan, msg.plan);
      if (lost) {
        const version = (await this.ctx.storage.get("version")) || 0;
        ws.send(JSON.stringify({
          type: "error",
          message: "Only an admin can delete " + lost + ". Nothing was changed.",
        }));
        return ws.send(JSON.stringify({ type: "sync", plan: currentPlan, version, by: null }));
      }
    }

    // Last write wins. With a handful of planners this is fine, but a client that
    // was editing an older version needs to know its base was superseded, so the
    // version travels back with every sync and stale senders get told.
    const current = (await this.ctx.storage.get("version")) || 0;
    const version = current + 1;
    const updated = new Date().toISOString();

    const stale = typeof msg.base === "number" && msg.base < current;

    // A stale writer is about to overwrite a board it never saw. Whole-board
    // last-write-wins is acceptable for block moves, but it silently deletes any
    // category or shelf card someone else created in the gap -- which is how two
    // people end up looking at different category lists. Deliberate deletes always
    // act on a fresh base (the delete button operates on the live list), so on a
    // stale base anything missing is an accident: merge it back in.
    if (stale) {
      const cur = (await this.ctx.storage.get("plan")) || null;
      for (const field of ["cats", "custom"]) {
        if (!cur || !Array.isArray(cur[field])) continue;
        if (!Array.isArray(msg.plan[field])) msg.plan[field] = [];
        const have = new Set(msg.plan[field].map((x) => x && x.id));
        for (const x of cur[field]) if (x && !have.has(x.id)) msg.plan[field].push(x);
      }
    }

    await this.ctx.storage.put({ plan: msg.plan, version, updated });
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
