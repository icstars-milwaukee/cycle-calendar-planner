# Cycle Calendar Planner

A single-page week planner for i.c.stars Milwaukee Tech Foundations cycles. Drag workshops
onto the week grid, group them by category, and adjust the schedule in the browser.

**Live site:** https://icstars-milwaukee.github.io/cycle-calendar-planner/

## How it works

The whole app is one self-contained `index.html` — no build step, no dependencies, no server.
It runs entirely in the browser.

## The shared board

Opening the page asks for a team password. Enter it and you join a live board: everyone
with the same password and room sees each other's edits within about a quarter of a second.
The status pill in the header shows the connection and how many other people are editing.

Choose **Work offline** instead and the planner behaves as it always did — private to your
browser, nothing shared.

Rooms come from the URL: `#room=cycle21` is the default, and `#room=anything-else` is a
separate board. Anyone with the password and the room name can read and edit that board, so
treat it like an "anyone with the link can edit" share.

**Simultaneous edits are last-write-wins.** Two people dragging different blocks is fine.
Two people dragging *the same* block at the same moment means one edit overwrites the other,
with no merge. Good for a few planners; not a general collaborative editor.

## Where your data lives

Two places, and it matters which:

- **Your browser** (`localStorage`) always holds a copy, so the board survives a reload and
  keeps working if the network drops.
- **The shared board** lives in a Cloudflare Durable Object, one per room, and is the source
  of truth whenever you are connected.

Practical consequences:

- Clearing site data erases the local copy; rejoining the room restores it from the server.
- Private/incognito windows discard the local copy when the window closes.
- If the Worker is unreachable, you keep planning locally and changes sync on reconnect.
- **Backup (.json)** writes the whole board to a file, and **Restore** reads it back. Unlike
  "Copy schedule" (a text summary) and `.ics` export (placed blocks only), this round-trips
  everything: custom workshops, categories, and colors.

## The password

The password lives only in the Worker's `ROOM_PASSWORD` secret and in each person's browser
after they type it. It is deliberately **not** in this repository, which is public — a
committed password would be no password at all.

To set or change it:

```
cd worker
npx wrangler secret put ROOM_PASSWORD
```

Everyone is prompted again the next time they connect. Note that the gate protects the
**board data**, not the page itself: the empty planner UI is on public GitHub Pages and
anyone can view its source. No schedule content is reachable without the password.

## Running it locally

Clone the repo and open `index.html` in a browser. That's it — no server needed.

```
git clone https://github.com/icstars-milwaukee/cycle-calendar-planner.git
```

## Deployment

The UI is served by GitHub Pages from the `main` branch root. A `.nojekyll` file is present
so Pages publishes the files as-is rather than running them through Jekyll.

The sync backend is a Cloudflare Worker in `worker/`:

```
cd worker
npx wrangler deploy
```

`SYNC_URL` near the bottom of `index.html` must point at the deployed Worker. Setting it to
an empty string disables sync and makes every board offline-only.

Durable Objects on the Cloudflare free plan must use the SQLite storage backend, which is
what `wrangler.jsonc` declares via `new_sqlite_classes`. WebSocket hibernation keeps idle
rooms from consuming the daily duration budget.

For local development, put the password in `worker/.dev.vars` (gitignored):

```
ROOM_PASSWORD=your-password-here
```

then `npx wrangler dev`.
