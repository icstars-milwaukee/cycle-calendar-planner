# Pushing the planner into a Microsoft 365 Group calendar

The planner is the source of truth. This flow mirrors it into a Group calendar on a
schedule: it creates events that are new, updates ones that moved or were renamed, and
deletes ones that were removed from the board.

## Why a flow and not a server job

Microsoft Graph has **no application permission for group calendars**. Only delegated
`Group.ReadWrite.All` works, meaning the write must happen as a signed-in member of the
group. A Cloudflare Worker holding a secret cannot do it. A Power Automate flow runs
under a member's connection, which satisfies that requirement without an app
registration or admin consent.

Consequence worth knowing up front: **the flow runs as whoever owns the connection.**
Every event will show that person as organizer, and the flow breaks if they leave or
lose the license. If this matters, build it under a service account rather than a
personal one.

## What the Worker does for you

Diffing is done server-side so the flow stays simple. `sync-plan` returns three flat
lists and the flow just loops them:

```
GET  /room/{room}/sync-plan?pw={password}
POST /room/{room}/sync-ack?pw={password}
```

`sync-plan` response:

```jsonc
{
  "ok": true,
  "weekOf": "2026-08-17",
  "changes": 2,               // 0 means the calendar is already correct; skip the loops
  "create": [ { "uid": "...", "hash": "...", "event": { /* Graph event body */ } } ],
  "update": [ { "uid": "...", "hash": "...", "eventId": "AAMk...", "event": { } } ],
  "delete": [ { "uid": "...", "eventId": "AAMk..." } ]
}
```

Each `event` is already shaped for Graph — `subject`, `start`/`end` as
`{dateTime, timeZone}` — so it can be passed through without transformation.

`sync-ack` is how the Worker learns which Graph event each block became. **Send it only
after Graph accepts each call**, and only for calls that succeeded. If the flow dies
halfway, the unacked work is simply offered again on the next run.

If the board has no week set, `sync-plan` returns `ok: false` with **empty lists** — it
will never interpret a missing week as "the board is empty, delete everything."

## Building the flow

Create a **Scheduled cloud flow**. Every 15 minutes is a reasonable starting point;
5 is fine too. There is no benefit to going faster than the planner is edited.

### 1. HTTP — get the plan

| Field | Value |
|---|---|
| Method | `GET` |
| URI | `https://cycle-calendar-sync.icstars-milwaukee.workers.dev/room/cycle21/sync-plan?pw=YOUR_PASSWORD` |

> The HTTP action is a **premium connector**. If your licence lacks it, see
> "If you don't have the HTTP connector" below.

### 2. Parse JSON

Content: the HTTP `Body`. Use this schema:

```json
{
  "type": "object",
  "properties": {
    "ok": { "type": "boolean" },
    "weekOf": { "type": "string" },
    "changes": { "type": "integer" },
    "create": {
      "type": "array",
      "items": {
        "type": "object",
        "properties": {
          "uid": { "type": "string" },
          "hash": { "type": "string" },
          "event": { "type": "object" }
        }
      }
    },
    "update": {
      "type": "array",
      "items": {
        "type": "object",
        "properties": {
          "uid": { "type": "string" },
          "hash": { "type": "string" },
          "eventId": { "type": "string" },
          "event": { "type": "object" }
        }
      }
    },
    "delete": {
      "type": "array",
      "items": {
        "type": "object",
        "properties": {
          "uid": { "type": "string" },
          "eventId": { "type": "string" }
        }
      }
    }
  }
}
```

### 3. Condition — stop early

Continue only when `ok` is `true` **and** `changes` is greater than `0`. This keeps the
run history readable: most runs will do nothing.

### 4. Initialise three array variables

`ackCreated`, `ackUpdated`, `ackDeleted` — all type Array. These collect what actually
succeeded.

### 5. Apply to each — `create`

Inside, use **Office 365 Groups → Create a group event (V2)** (or V4 if offered):

| Field | Value |
|---|---|
| Group Id | your group |
| Subject | `items('Apply_to_each')?['event']?['subject']` |
| Start time | `items('Apply_to_each')?['event']?['start']?['dateTime']` |
| End time | `items('Apply_to_each')?['event']?['end']?['dateTime']` |
| Time zone | `Central Time (US & Canada)` |

Then **Append to array variable** `ackCreated`:

```
{
  "uid": "@{items('Apply_to_each')?['uid']}",
  "hash": "@{items('Apply_to_each')?['hash']}",
  "eventId": "@{body('Create_a_group_event_(V2)')?['id']}"
}
```

> The connector takes a time zone by **name**, not the IANA `America/Chicago` the feed
> carries. Set it to `Central Time (US & Canada)` and pass the bare `dateTime` string.
> Passing a UTC timestamp here instead is the classic way to land every event an hour
> off half the year.

### 6. Apply to each — `update`

Use **Update a group event (V2)** with `Id` = `items('Apply_to_each_2')?['eventId']`,
and the same field mapping as above. Append to `ackUpdated`:

```
{
  "uid": "@{items('Apply_to_each_2')?['uid']}",
  "hash": "@{items('Apply_to_each_2')?['hash']}",
  "eventId": "@{items('Apply_to_each_2')?['eventId']}"
}
```

Set this action's **Configure run after** to include *has failed*, followed by a
condition that only appends on success — otherwise one deleted-by-hand event aborts the
whole run. Alternatively leave it strict at first and loosen once it is stable.

### 7. Apply to each — `delete`

Use **Delete a group event** with `Id` = `items('Apply_to_each_3')?['eventId']`.
Append `{ "uid": "@{items('Apply_to_each_3')?['uid']}" }` to `ackDeleted`.

A delete that fails because someone already removed the event by hand is harmless —
allow the run to continue.

### 8. HTTP — acknowledge

| Field | Value |
|---|---|
| Method | `POST` |
| URI | `.../room/cycle21/sync-ack?pw=YOUR_PASSWORD` |
| Header | `Content-Type: application/json` |
| Body | `{ "created": @{variables('ackCreated')}, "updated": @{variables('ackUpdated')}, "deleted": @{variables('ackDeleted')} }` |

## Verifying it works

1. Run the flow once with an empty board — it should do nothing.
2. Add one workshop in the planner, set the **Week of** date, run the flow. One event
   appears in the Group calendar.
3. Run the flow again without touching the board. **Nothing should change** — this is
   the duplicate check, and the whole design rests on it.
4. Drag the workshop to another day, run again. The existing event moves; no second
   event appears.
5. Delete it from the shelf, run again. The event disappears.

## Known limitations

- **The organizer is whoever owns the connection**, not the person who edited the board.
- **One-way.** Editing an event in Outlook is overwritten on the next sync. The event
  body says so, but people will still try.
- **Deletions are driven by the Worker's memory**, not by scanning the calendar. If
  someone deletes an event in Outlook, the Worker still believes it exists and will not
  recreate it until that block changes. Editing the block in the planner repairs it.
- **A flow that fails after creating an event but before acking leaves an orphan** — a
  real event the Worker does not track, which will be created again next run. Rare, but
  it means duplicates are possible under partial failure even though repeated healthy
  runs never duplicate. Deleting the stray by hand is the fix.
- **No attendees, location, or reminders** are pushed. The planner does not model them.

## If you don't have the HTTP connector

The HTTP action is premium. Without it, the same shape works using a low-code
alternative: the Worker's `sync-plan` output can be fetched with the **RSS** or
**Outlook Send an HTTP request** connectors in some tenants, but the cleaner fallback is
to run the reconciliation as a small scheduled script elsewhere that authenticates as a
group member. Ask before going down that path — it reintroduces the token-storage
problem the flow was chosen to avoid.
