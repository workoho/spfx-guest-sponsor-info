---
layout: doc
lang: en
title: Why This Web Part Exists
permalink: /en/why/
description: >-
  Understand the guest onboarding gap this web part closes — why Teams contact
  buttons break for new guests, and how the web part helps.
lead: >-
  When a guest accepts an invitation, they may exist in Entra but not yet in
  Teams. Their contact buttons appear — but silently fail. This web part makes
  that situation visible.
mermaid: false
---

## After the Invitation {#after-the-invitation}

When someone outside your organisation receives a Microsoft 365 guest invitation
and clicks "Accept", something real happens in Entra ID: a guest account is
created, an invitation is recorded, and technically the person is inside your
tenant.

What is not guaranteed to happen immediately: a presence in Microsoft Teams and
Microsoft Exchange for that guest in your tenant.

---

## Two Ways to Invite a Guest {#two-ways-to-invite-a-guest}

How a guest ends up in your tenant determines what they can do — and how quickly.

### Implicit invitation

An employee adds an external contact directly to a Teams team or a SharePoint
site. Microsoft generates an invitation email in the background. The moment the
guest accepts, they are added to the team — and Teams begins establishing their
presence immediately. **The guest is ready in Teams within minutes.**

### Managed lifecycle invitation

A provisioning process — a governance platform, a script, or a workflow in the
Entra admin center — creates the guest account formally. The guest account exists
in Entra. But no one has added this person to a Teams team yet. Teams has not
yet established a presence for them in your tenant.

**The guest exists in Entra. They do not yet exist in Teams.**

This is the gap. And it is entirely invisible unless something explicitly shows
it.

---

## The Teams Timing Problem {#the-teams-timing-problem}

When a guest account exists in Entra but has not yet been added to any Teams
team, their Teams presence in your tenant is not established.

On a SharePoint landing page, a sponsor contact card might show:

- The sponsor's name and profile photo — ✓ available via Entra
- An email address — ✓ available
- A chat button — rendered on screen, but **silently broken**
- A call button — rendered on screen, but **silently broken**

The guest sees a contact card that looks ready. The buttons do nothing, or
navigate somewhere unexpected. There is no error message. There is no
explanation. The guest has no way to know whether the button is broken, whether
they did something wrong, or whether the feature simply is not meant for them.

This is not a Teams bug. It is the expected behaviour when a Teams presence has
not been established for the guest yet. But without something that detects and
communicates this, the guest experience is silently broken.

---

## What This Web Part Does {#what-this-web-part-does}

The **Guest Sponsor Info** web part is designed for the guest landing page
scenario. It is placed on the SharePoint Entrance Area — the page guests land on
after accepting an invitation.

It does two things:

1. **Shows the guest's sponsors** — the internal employees assigned in Entra as
   responsible for the guest's access. Names, photos, titles, and contact
   options. No per-guest configuration. No manual updates as sponsors change.

2. **Detects Teams readiness** — if the guest does not yet have a Teams presence
   in your tenant, the web part detects this and adjusts the contact card
   accordingly: the chat and call buttons are greyed out, and a clear status
   message explains the situation. The guest is not left wondering. They see a
   face, a name, and an honest status.

The result: a guest whose access is still being provisioned can still reach their
sponsor — by email — and knows that Teams access is on its way. No confusion. No
silent failure.

---

## What You Need {#prerequisites}

1. A SharePoint site that serves as a guest landing page (the Entrance Area).
2. An Azure Function deployment (the Guest Sponsor API) — this is what the web
   part calls to retrieve sponsor data.
3. Guest accounts with sponsors assigned in Microsoft Entra ID.

See the [Setup Guide](/en/setup/) for step-by-step instructions.
