# Admin consent request

Text to forward to whoever administers the icstars.org Microsoft 365 tenant. It is
written to be read by an admin who will reasonably want to know the blast radius, so it
states the awkward part rather than burying it.

---

**Subject: One-time admin consent for a Cycle Calendar Planner app registration**

We have a scheduling tool for Tech Foundations cycles. Staff plan a 14-week cycle in it,
press Publish, and the sessions should appear on the **Cycle21-MKE** Microsoft 365 Group
calendar so they show against everyone's other meetings.

To write to a Group calendar, Microsoft Graph requires the delegated permission
**`Group.ReadWrite.All`**. There is no narrower option — Microsoft provides no application
permission for group calendars at all, and no per-calendar scope. That permission
requires one-time admin consent, which is what I am asking for.

**What I would like:**

1. I register an application in Entra ID (tenant default already allows this).
2. You click **Grant admin consent** once on that app for delegated
   `Group.ReadWrite.All`.

Nothing ongoing after that.

**What this does and does not allow — the honest version:**

- It is a **delegated** permission, so the app can only ever act *as a signed-in user*,
  and only within what that user could already do by hand in Outlook. It cannot run
  unattended as itself and cannot reach anything the signed-in person cannot reach.
- However, `Group.ReadWrite.All` is **broader than the one calendar we need**. As
  Microsoft scopes it, it also permits creating groups and updating content for groups
  the signed-in user belongs to. I am not asking for that capability because we want it;
  it is bundled, and Microsoft offers nothing narrower for this task.
- Consent is revocable at any time from Entra ID → Enterprise applications, which
  immediately stops the integration.
- The app would sign in as a single account we nominate. If you would prefer that be a
  service account rather than a staff member, that suits us better anyway, since the
  integration then survives people changing roles.

**If you would rather not grant it:** there is a route that needs no consent at all. A
Power Automate flow using the Office 365 Groups connector can write the same events,
because that connector is a Microsoft first-party application already consented in every
tenant. It needs a Power Automate licence covering the HTTP action and is more fragile to
maintain, but it requires nothing from you. Tell me which you prefer and I will build
that instead.

---

## Notes for us, not for the admin

- If the answer is no, nothing is lost: the reconciliation logic lives in the Cloudflare
  Worker, not in the client, so both routes consume the same `sync-plan` / `sync-ack`
  endpoints. Only the last mile changes.
- If the answer is yes, ask for a **service account** to own the connection. Whichever
  identity signs in becomes the organiser on every event, and the integration breaks when
  that person leaves.
