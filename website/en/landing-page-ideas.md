---
layout: doc
lang: en
title: Landing Page Ideas
permalink: /en/landing-page-ideas/
description: >-
  Optional inspiration for what else to place on a guest landing page —
  especially Quick Links areas and tenant-pinned Microsoft 365 deep links.
lead: >-
  A small optional add-on for admins who want the landing page itself to be
  more helpful. These ideas sit around the Guest Sponsor Info web part;
  they are not required for the product to work.
---

## Why This Page Exists

This page is intentionally small. The goal is not to prescribe a complete
landing-page blueprint, but to show a few proven ideas for what else a shared
guest entry page can contain besides the sponsor web part.

If you use the SharePoint **Quick Links** web part, you can already build most
of this without custom code. With audience targeting enabled, the same landing
page can show different links to employees and guests.

The examples below use placeholders such as `<tenant-id>`, `<tenant-name>`,
and `<tenant-domain>`. Replace them with your own values.

## A Good Default Pattern

For many tenants, two Quick Links web parts are already enough:

- One area for **Microsoft 365** entry points in the correct tenant.
- One area for **My Guest Account** self-service actions.
- Optionally, a small standalone Quick Links web part for **More Apps** so it
  looks secondary rather than like the primary destination.

This works especially well when the landing page itself uses audience targeting.
Employees can see internal employee resources, while guests see only the links
that actually help them in the resource tenant.

The screenshot below shows one possible composition: a shared entry page with
tenant-pinned Quick Links near the top and the sponsor area further down the
page.

<img src="{{ '/assets/images/entrance-landingpage-example.jpg' | relative_url }}" alt="Landing page example screenshot.">

## Make The Page Easy To Return To

Guests benefit from a page they can find again without friction.

- If the landing page eventually lives at the tenant root site (`/`), the URL
  is often easy enough to remember on its own.
- If that is not possible, consider a memorable short URL or shortlink that
  resolves to the landing page.
- Add a visible call to action near the top of the page asking guests to
  bookmark the page after their first successful sign-in.
- Keep mentioning the page in invitation and onboarding emails, but do not
  assume guests will keep those emails forever or want to search for them
  later.

Because browsers do not offer one universal "add bookmark" link that works
cleanly everywhere, this is usually best as a simple instruction or callout,
not as a special scripted button.

## Area 1 — Microsoft 365

This area gives guests stable entry points into the resource tenant. Where a
Microsoft URL supports `tenantId`, include it. For SharePoint, the tenant
hostname itself already fixes the tenant context.

### Microsoft Teams

Use a tenant-pinned Teams entry link when you want Teams to open in the correct
resource tenant instead of whichever tenant happened to be active before.

```text
https://teams.cloud.microsoft/?tenantId=<tenant-id>
```

This is useful because it does not assume that the guest already knows how to
switch tenant context manually. It also avoids sending the guest into a team-
specific deep link before you know that team membership already exists.

### Microsoft SharePoint

Link to a tenant-owned overview page, hub site, or site directory that helps
guests find shared workspaces and storage locations even without navigating
through Teams first.

```text
https://<tenant-name>.sharepoint.com/teams/overview
```

In some tenants this is a hub site. In others it is a manually curated overview
page. Either is fine. The important part is that the URL is already tenant-
fixed because it uses your SharePoint hostname. It can also help guests find
Team-connected storage areas or plain team sites that are not "teamified".

Making the landing page itself a **hub site** can be useful even before you
have any other sites to associate. You already gain a shared navigation layer,
hub branding options, and a clearer identity for the whole area.

That can also help with naming. For example, the underlying site can keep a
friendly site title such as `Welcome @ Contoso`, while the hub identity and hub
navigation present the broader area as something like `Entrance Area`.

If you later do associate more sites, the hub becomes even more useful. You can
add links to associated or non-associated sites in the hub navigation and use
audience targeting so employees and guests do not have to see the same hub
links.

### Viva Engage

If your tenant uses Viva Engage as a broader community layer, it can be a
useful parallel entry point beside Teams and SharePoint.

```text
https://engage.cloud.microsoft/main/org/<tenant-domain>/
```

This works best when the guest actually has access to relevant communities. If
not, keep it audience-targeted or omit it.

### More Apps

A tenant-pinned My Applications link is useful as a fallback, but on the
landing page it usually works better as a secondary action than as the primary
entry point.

```text
https://myapplications.microsoft.com/?tenantId=<tenant-id>
```

A nice pattern is to render this as its own small Quick Links web part without
a visible section title, so it looks like an extra utility link rather than the
main path.

## Area 2 — My Guest Account

This area focuses on self-service. It helps guests manage their account inside
the correct resource tenant without first figuring out tenant switching on
their own.

### Guest Account

Link directly to the guest's account view in the correct tenant.

```text
https://myaccount.microsoft.com/?tenantId=<tenant-id>
```

This is useful when the guest needs to review account context, organization
information, or account-related prompts in the resource tenant.

### Security Info

This link is especially helpful when the resource tenant requires MFA
registration for guests, when MFA trust is not in place, or when the invitation
was redeemed with an identity that did not already carry the needed
authentication methods.

```text
https://mysignins.microsoft.com/security-info?tenantId=<tenant-id>
```

If the guest has to register authentication methods in the resource tenant,
this is one of the most valuable links on the whole page.

### Terms of Use

Terms of Use acceptances are easy to lose in portal navigation. A direct link
makes them re-openable and reviewable.

```text
https://myaccount.microsoft.com/termsofuse/myacceptances?tenantId=<tenant-id>
```

This is helpful when Conditional Access has presented one or more Terms of Use
over time and the guest needs to revisit what they accepted.

### Delete Guest Access

Many guests do not know that they can leave an organization themselves. If you
want to support privacy and clean offboarding, make that option discoverable.

```text
https://myaccount.microsoft.com/organizations/leave/<tenant-id>?tenant=<tenant-id>
```

You may want to label this link more explicitly on the page, for example as
**Delete Guest Access**, **Remove my guest access**, or **Leave this
organization**.

## Deep-Link Rules Of Thumb

- Use `tenantId` wherever the target service supports it.
- For SharePoint, use a tenant-owned URL instead of a generic Microsoft 365
  home page.
- Prefer overview pages and navigation hubs over links that only work after
  team membership exists.
- Test every important link while signed into a home tenant and at least one
  additional guest tenant.
- Re-test these links periodically. The Teams patterns are documented, but some
  account-portal URLs are practical deep-link patterns that Microsoft can
  change over time.

## A Simple Audience-Targeting Model

On the same landing page, guests and employees do not need to see the same
surrounding content.

- Show guests sponsor help, tenant-pinned app links, security-info self-
  service, and optional guest policy links.
- Show employees internal navigation, internal IT support, HR resources, and
  internal-only collaboration destinations.
- Keep the Guest Sponsor Info web part where it adds value, but use Quick Links
  around it to make the whole page feel intentional.

## Related Microsoft Guidance

- [Use the Quick Links web part](https://support.microsoft.com/office/use-the-quick-links-web-part-e1df7561-209d-4362-96d4-469f85ab2a82)
- [Deep links in Microsoft Teams](https://learn.microsoft.com/microsoftteams/platform/concepts/build-and-test/deep-link-teams)
- [Planning your SharePoint hub sites](https://learn.microsoft.com/sharepoint/planning-hub-sites)
