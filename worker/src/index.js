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
    const match = url.pathname.match(/^\/room\/([A-Za-z0-9_-]{1,64})(\/plan)?$/);
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
    server.send(JSON.stringify({ type: "snapshot", ...snapshot, peers: this.peerCount() }));
    this.broadcastPeers();

    return new Response(null, { status: 101, webSocket: client });
  }

  async snapshot() {
    const plan = (await this.ctx.storage.get("plan")) || null;
    const version = (await this.ctx.storage.get("version")) || 0;
    const updated = (await this.ctx.storage.get("updated")) || null;
    return { plan, version, updated };
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
  }

  async webSocketClose(ws) {
    try { ws.close(); } catch (e) { /* already closed */ }
    this.broadcastPeers();
  }

  async webSocketError(ws) {
    this.broadcastPeers();
  }
}
