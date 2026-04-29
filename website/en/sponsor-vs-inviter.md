---
layout: doc
lang: en
title: Sponsor vs Inviter
permalink: /en/sponsor-vs-inviter/
description: >-
  Understand the difference between the sponsor and the inviter in Microsoft
  Entra B2B guest onboarding, and why a SharePoint guest landing page should
  show the sponsor relationship clearly.
lead: >-
  In Microsoft Entra B2B guest onboarding, the inviter and the sponsor can be
  the same person, but they often are not. That distinction matters whenever a
  guest needs a real support contact on the SharePoint landing page.
---

## Short Answer

The **inviter** is the person or process that triggered the invitation.

The **sponsor** is the person or group recorded in Microsoft Entra's
**Sponsors** field for that guest relationship.

Sometimes both roles point to the same employee. In structured guest
onboarding, lifecycle governance, or admin-led invitation flows, they often do
not.

## Sponsor or "Owner"?

Some products and internal processes use the word **owner** for that same
responsible contact or group. On this website we prefer **sponsor** because
that is the term Microsoft Entra uses in the **Sponsors** field.

Here, sponsor does **not** mean budget owner, line manager, or someone with
special authority over the guest. It simply means the responsible internal
contact or group recorded for that guest relationship.

## Why Guests Mix Them Up

Guests usually see the inviter first because the invitation email, approval
flow, or request history exposes that name. The sponsor relationship in
Microsoft Entra is much less visible to them.

That creates a practical problem. A guest may remember who sent the invitation,
but still have no idea who is currently responsible for their access, who the
backup sponsor is, or who they should contact when onboarding is stuck.

This is one of the reasons a SharePoint guest landing page is useful: it can
become the place where sponsor visibility finally exists.

## Sponsor vs Inviter at a Glance

| Question | Sponsor | Inviter |
|---|---|---|
| What is the role? | Person or group recorded in Microsoft Entra's Sponsors field for the guest relationship | Person or process that triggered the invitation |
| What does the guest usually see first? | Usually hidden by default | Usually visible in the invitation process |
| Is it the right contact for ongoing support? | Usually yes | Not necessarily |
| Can there be multiple responsible contacts? | Yes, Microsoft Entra supports up to five sponsors | Not really; it is tied to the invitation event |
| Should a guest landing page emphasize it? | Yes, because it improves sponsor visibility | Only as context, not as the main support path |

## Why the Difference Matters on a Guest Landing Page

If your SharePoint guest landing page only repeats the inviter name from the
email, the core onboarding problem is still unresolved.

The guest still does not know:

- who is responsible for their access now
- who to contact if the primary sponsor is unavailable
- whether Teams guest access is ready yet
- whether the person from the invitation mail is even the right person anymore

Sponsor visibility solves a more durable problem than simply reminding someone
who sent the original invitation.

## What Guest Sponsor Info Shows Instead

Guest Sponsor Info is designed for exactly this gap. On the SharePoint landing
page it can show:

- the guest's assigned sponsor contacts from Microsoft Entra
- backup sponsors when multiple responsible contacts exist
- manager context and contact details
- honest Teams onboarding status when chat or call actions are not ready yet

That gives the guest a stable support surface during guest onboarding and after
it.

## When Sponsor and Inviter Are the Same

There are simple invitation flows where the same employee both invites the
guest and remains the ongoing sponsor. In standard Entra invitation flows,
that is also the default unless another sponsor is specified. SharePoint
sharing invitations to brand-new external users are a documented exception:
sponsors may need to be added manually there.

Even then, showing the sponsor relationship explicitly is still useful. It
confirms that the visible contact is not just the historical sender of an email
but the current responsible contact inside the tenant.

<div class="doc-cta-box">
  <div>
    <p class="doc-cta-title">Use the sponsor relationship on purpose</p>
    <p class="doc-cta-sub">Start with the Why page or go straight to the setup guide.</p>
  </div>
  <div class="doc-cta-actions">
    <a href="{{ '/en/why/' | relative_url }}" class="btn btn-outline">Why this exists</a>
    <a href="{{ '/en/setup/' | relative_url }}" class="btn btn-teal">Setup Guide</a>
  </div>
</div>
