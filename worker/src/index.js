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

function pad2(n) { return (n < 10 ? "0" : "") + n; }

// Days are added on the UTC calendar so a host timezone can never shift the date.
function addDaysISO(mondayISO, days) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(mondayISO || "");
  if (!m) return null;
  const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  d.setUTCDate(d.getUTCDate() + days);
  return d.getUTCFullYear() + "-" + pad2(d.getUTCMonth() + 1) + "-" + pad2(d.getUTCDate());
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

  const events = (plan.placed || []).map((p) => {
    const date = addDaysISO(plan.weekOf, p.day);
    if (!date) return null;
    return {
      // Stable across pushes: the same block keeps the same id for its whole life,
      // so a bridge can update or delete rather than creating duplicates.
      uid: "cycle-planner-" + (room || "board") + "-" + p.id,
      subject: p.title,
      category: catName(p.cat || (p.refId ? (plan.custom || []).find((c) => c.id === p.refId) || {} : {}).cat),
      start: { dateTime: date + "T" + hhmm(p.start), timeZone: CAL_TZ },
      end: { dateTime: date + "T" + hhmm(p.start + p.mins), timeZone: CAL_TZ },
      day: p.day,
      minutes: p.mins,
    };
  }).filter(Boolean);

  // Deterministic order so a diff between two pushes is meaningful.
  events.sort((a, b) => (a.day - b.day) || a.start.dateTime.localeCompare(b.start.dateTime));

  return {
    ok: true,
    weekOf: plan.weekOf,
    timeZone: CAL_TZ,
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
  const basis = [ev.subject, ev.start.dateTime, ev.end.dateTime, ev.start.timeZone, ev.category].join("|");
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
    body: {
      contentType: "text",
      content: "Cycle Calendar Planner" + (ev.category ? " - " + ev.category : "") +
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
  async fetch(request, env) {
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

    // /room/<name>        -> WebSocket upgrade for live editing
    // /room/<name>/plan   -> plain GET snapshot, handy for debugging and backups
    const match = url.pathname.match(
      /^\/room\/([A-Za-z0-9_-]{1,64})(\/plan|\/events|\/sync-plan|\/sync-ack)?$/);
    if (!match) return new Response("Not found", { status: 404, headers: cors });

    // Every route that touches board data is gated, including the upgrade itself,
    // so an unauthenticated socket is never accepted in the first place.
    if (!authorized(request, env)) return deny(cors);

    const room = match[1];
    const id = env.BOARD.idFromName(room);
    return env.BOARD.get(id).fetch(request);
  },
};

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
