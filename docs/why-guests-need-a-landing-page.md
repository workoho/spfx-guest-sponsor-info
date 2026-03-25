# Your Guests Accepted the Invitation. Now What?

**In this article**

- [The Moment After the Invitation](#the-moment-after-the-invitation)
- [Why Guests Keep Getting Lost](#why-guests-keep-getting-lost)
- [Why Deep Links with an Explicit Tenant ID Are the Reliable Fix](#why-deep-links-with-an-explicit-tenant-id-are-the-reliable-fix)
  - [Why Common Alternatives Fall Short](#why-common-alternatives-fall-short)
  - [The Landing Page as the Simpler Answer](#the-landing-page-as-the-simpler-answer)
- [What a Guest Landing Page Should Contain](#what-a-guest-landing-page-should-contain)
  - [Location, URL, and Audience Strategy](#location-url-and-audience-strategy)
  - [Branding Does Part of the Work](#branding-does-part-of-the-work)
  - [A Practical Page Structure](#a-practical-page-structure)
  - [One More Thing: A URL Worth Sharing](#one-more-thing-a-url-worth-sharing)
- [One Feature Worth Adding: Sponsor Visibility](#one-feature-worth-adding-sponsor-visibility)
  - [The Guest Sponsor Info Web Part](#the-guest-sponsor-info-web-part)
  - [The Sponsor Field: Useful but Static](#the-sponsor-field-useful-but-static)
  - [Lifecycle Management: Tooling Options](#lifecycle-management-tooling-options)
  - [What It Shows on the Page](#what-it-shows-on-the-page)
  - [Going Further: Guest Type Classification](#going-further-guest-type-classification)
- [This Is Also a Governance Question](#this-is-also-a-governance-question)
- [Working on This Kind of Challenge](#working-on-this-kind-of-challenge)

---

An employee adds an external contact to a Teams team. Or triggers a formal invitation
through a provisioning process, or simply via the Azure portal — however your organization
handles it. The guest receives an email, clicks "Accept", authenticates, and within seconds
they are technically inside your Microsoft 365 environment.

What happens in the next thirty seconds usually gets no design attention at all.

---

## The Moment After the Invitation

The B2B guest redemption flow itself is smooth. Microsoft has invested heavily in making
the invitation and authentication experience work reliably across organizations, identity
providers, and devices. The problem is the landing point.

After accepting a guest invitation, most users end up on the Microsoft MyApps portal, a
generic account switcher, or — depending on how their browser session is configured —
their own organization's home screen. There is no context. No "you are now in Acme Corp's
Microsoft 365 environment." No indication of where to go next.

The MyApps portal deserves a special mention here, because it makes things worse rather
than better: in most organizations it is simply not maintained. Guests land there and see a
list of applications — many of which they have no actual access to, or that are entirely
irrelevant to their collaboration scope. It creates the impression of a toolbox without
telling you which tool to pick up, or whether any of them even work for you. Managing what
guests see in MyApps — not just who can access what, but what is visible to whom — is its
own topic and is worth addressing separately. For now, the point is: MyApps is not a
substitute for a proper landing page. It is, however, worth adding one prominent link there
that points guests to the Entrance Area — so that even guests who land in MyApps have a
clear exit ramp to somewhere more useful.

For experienced Microsoft 365 users, this is a mild nuisance. For everyone else — vendors,
external consultants, partners who primarily work in different tools — it is a genuine moment
of confusion that often leads to abandoned sessions, repeat authentication attempts, or
continued reliance on email when the whole point of the invitation was to enable richer
collaboration.

---

## Why Guests Keep Getting Lost

The confusion is not random. It has a structural cause.

Microsoft 365 is a multi-tenant platform. When a guest user clicks a link — any link —
the browser and the Microsoft 365 client have to decide which tenant context to use. If the
guest has their own Microsoft 365 account, there is already an active session for their home
tenancy. Generic links often resolve there first.

So a guest clicks a link to a Teams channel and ends up in their own organization's Teams.
The channel they were supposed to join is not there. It is not missing — it exists in
*your* tenant, not theirs. But nothing in the experience told them to switch.

The result: a support ticket, or silence — or a chain of messages across organizations as the
confused guest reaches out to whoever seems reachable: the project manager who sent the
invitation, a colleague, a helpdesk that does not have the right context. And the person
asked often does not know the answer either. The cost of these small failures is real: time
spent by multiple people at more than one organization, for a problem that should never have
created friction in the first place.

---

## Why Deep Links with an Explicit Tenant ID Are the Reliable Fix

This is where a technical detail becomes genuinely important to understand —
without needing to dive into the mechanics.

Microsoft 365 supports deep links — URLs that point to specific resources — across Teams,
SharePoint, Viva Engage, and other workloads. When these links include the target tenant's
ID as a URL parameter, they carry enough context for the client to load the correct
organizational environment. The guest still authenticates as themselves, but the system knows
exactly which tenant to pull from.

A Teams deep link that includes your tenant ID will prompt the guest to switch to the correct
account if they are not already using the right one. A generic Teams link will not. That small
difference has a large practical impact.

The difference is literally one URL parameter. A generic link to a Teams channel:

```text
https://teams.cloud.microsoft/l/channel/{channelId}/{channelName}?groupId={teamId}
```

The same link, tenant-scoped:

```text
https://teams.cloud.microsoft/l/channel/{channelId}/{channelName}?groupId={teamId}&tenantId={tenantId}
```

Your tenant ID is a GUID you can find in the Microsoft Entra admin center. Teams generates
full channel deep links automatically — right-click any channel and choose "Get link to
channel" — and those links already include the `groupId`. Adding `&tenantId=` to that URL
is the only manual step required.

### Why Common Alternatives Fall Short

The alternatives are well-intentioned but unreliable:

- **Carefully worded email instructions** depend on the guest reading and following them
  precisely — a low-confidence bet, especially days or weeks after the original invitation.
- **A Teams deep link in the invitation email** is a natural instinct — the guest accepts
  and lands straight in Microsoft Teams. Except it only works once Teams has established a
  presence for the guest in your tenant, which only happens after they have been added to at
  least one team. Accepting a B2B invitation alone is not enough — Teams and SharePoint are
  separate services, and an accepted invitation does not automatically give a guest a Teams
  identity in your environment. Microsoft's own documentation is explicit on this point: until
  that Teams presence exists, your tenant will not appear in the guest's tenant switcher, and
  any deep link into Teams will not resolve correctly. The guest ends up not in their own
  tenant, not in yours, but in an undefined state with no indication of what went wrong.
  A SharePoint landing page avoids this entirely — it works on day one,
  before any team membership exists. And because SharePoint is a separate service from Teams,
  it can even be used to prepare the guest for the situation: a dedicated SharePoint web part
  ([covered below](#one-feature-worth-adding-sponsor-visibility)) can detect that Teams access is not yet ready and display
  that information clearly on the page, while simultaneously greying out the call and chat
  buttons on the sponsor's contact card. The guest is not left stranded — they see a face,
  a name, and an honest status: *Teams access is being set up. Reach out by email for now.*
- **Custom Azure AD B2C flows or redirect apps** can technically address this, but they add
  significant engineering overhead and ongoing maintenance cost.
- **Conditional Access policies** help with access control but do not solve the orientation
  problem — they do not tell a guest where to go, only whether they can go there.
  That said, requiring MFA for guests is a clear security best practice and worth doing
  regardless of how you solve the orientation problem. The good news: for most guests who
  already have an Entra account in their home organization, you do not need to force them
  to register new MFA credentials in your tenant. The recommended default is to trust MFA
  from the guest's home tenant — Microsoft detects the account type at invitation redemption
  and applies the trust accordingly. With the right Conditional Access configuration, guests
  who already use MFA at work will barely notice the requirement in yours. Getting that
  balance right — secure but frictionless — is a topic in its own right, and one we are
  happy to talk through.

### The Landing Page as the Simpler Answer

A well-structured landing page with properly formed, tenant-scoped deep links is the simpler,
more durable answer. It is low-cost to build, easy to maintain, and it works regardless of
whether the guest is a seasoned Microsoft 365 user or someone opening Teams for the first time.

One practical prerequisite: for a custom landing page URL to actually reach the guest, it
needs to be embedded in the invitation itself. Microsoft Entra supports this — the invitation
flow accepts a redirect URL that the guest is forwarded to after accepting. But this only
works when the invitation process is managed. When employees invite guests implicitly — by
adding them directly to a Teams team or a SharePoint site, for example — Microsoft generates
a standard invitation email with no custom redirect. That is one more reason to move away
from unmanaged, ad hoc guest provisioning: it removes the one moment in the flow where you
could point a new guest somewhere intentional.

---

## What a Guest Landing Page Should Contain

We call this page the *Entrance Area* — a deliberate name that reflects its purpose: not a
dashboard, not a portal, not an intranet. Just a clear, calm place where the guest arrives
and immediately knows where they are.

### Location, URL, and Audience Strategy

The URL matters more than it might seem. A path like `/sites/entrance` is clean and
consistent. But the most practical option is to configure the Entrance Area as the root site
of your SharePoint tenant. A root site is both easy to remember and, for a technically
curious guest, easy to anticipate — someone who knows your domain can reasonably guess that
`https://yourtenant.sharepoint.com` leads somewhere meaningful. That kind of predictability
has real value when a guest returns weeks later and has lost the original link.

The page can also be built as audience-aware: SharePoint's targeting capabilities let you
show different content to internal employees, to guest users, and to external staff who have
a company account in your tenant. One practical note worth knowing: audience targeting only
works with certain web parts — not all of them support it. The Quick Links web part is the
most versatile choice here, as it handles both navigation and audience-filtered content well.
For internal employees on managed devices, a browser policy or device configuration can
already redirect them to the intranet directly — which means the root site does not need to
double as an intranet gateway. That leaves it free to do one thing well: serve as an
unambiguous entrance for everyone who is not yet sure where they belong.

The goal is not a polished marketing experience. The goal is orientation — getting the guest
from "where am I?" to "I know exactly what to do" in under two minutes.

### Branding Does Part of the Work

One element that contributes to that orientation without any extra effort is branding itself.
SharePoint's global navigation and the Microsoft 365 app bar already display your
organization's name and logo — and for a guest, seeing those immediately answers the most
basic question: *whose environment am I in?* But this only works if you have actually
configured them. A SharePoint tenant with no custom logo, no navigation, and default colors
tells a guest nothing about where they have landed. If you have set up a Microsoft 365
organization profile and a custom SharePoint theme — one that reflects your company's colors
and visual identity — then the Entrance Area does not just orient guests through its content.
It does it through its appearance the moment the page loads. This matters more here than on
any internal intranet page, because this is where you present yourself not only to your own
employees, but to the outside world.

### A Practical Page Structure

Here is a practical structure that works well in both simple and more complex collaboration
scenarios:

**1. A short welcome and context statement**

Confirm they are in the right place, name the organization that invited them, and briefly
describe the nature of the collaboration. Two or three sentences is enough. The guest should
immediately feel like they have arrived somewhere intentional — not somewhere accidental. If
you want this welcome text to be audience-targeted — shown only to guests, for instance —
the Quick Links web part is a practical workaround: configure a link that points back to the
same page, set the display style to look like plain content rather than a navigation element,
and apply audience targeting to the web part itself. It is a small trick, but it gets the
job done without extra infrastructure.

**2. Where to go first**

Not a comprehensive index of every resource — just the most important first step. A Teams
team, a shared channel, a project site. Prioritize the one or two things they are most likely
to need in their first session.

**3. Tenant-scoped deep links to key tools**

- **Microsoft Teams**: A direct link to the relevant team or channel, with your tenant ID
  included. This ensures the Teams client loads in your tenant context, not the guest's.
  Use the channel link Teams generates (right-click → "Get link to channel") and append
  `&tenantId={yourTenantId}`. To link to a team overview rather than a specific channel,
  use `https://teams.cloud.microsoft/?tenantId={yourTenantId}` as a simple tenant-scoped
  entry point.

- **Viva Engage**: If your organization uses Viva Engage for company-wide or cross-functional
  community, a properly scoped community link looks like:
  `https://engage.cloud.microsoft/main/groups/{communityId}/all`
  The community ID is visible in the URL when you open the community in a browser.
  Viva Engage resolves tenant context from the active Microsoft 365 session — which is
  exactly why pointing guests to a tenant-scoped Teams or SharePoint link *first* matters:
  it establishes the correct session before they navigate further.

**4. Direct links to relevant resources**

SharePoint sites, document libraries, project wikis, or Microsoft Loop workspaces. The
specific ones they will need in the first week. Not a complete sitemap — a curated shortlist.
If your MyApps portal is actively maintained, link to it here as well — alongside the
Microsoft 365 apps you want to highlight explicitly. A curated shortlist of tools is far more
useful than a portal full of tiles the guest cannot actually open. For the same reason that
applies to Teams links, the MyApps link should include a tenant ID parameter so the portal
opens in the correct organizational context:
`https://myapps.microsoft.com/?tenantid={yourTenantId}`

**5. A news or announcements section**

For most organizations, email is the only channel that reliably reaches guest users. The
Entrance Area can change that. A simple news section — a SharePoint News web part scoped to
guest-relevant content — gives you a place to post updates that matter to external
collaborators: policy changes, new tools, scheduled maintenance, announcements that affect
the collaboration. Guests who know they can find current information there have a reason to
return. Which points to a broader design principle: the Entrance Area should contain enough
genuinely useful content that a guest *wants* to bookmark it — not just land there once and
move on. A page worth saving is a page that keeps doing its job long after the original
invitation email has been archived.

**6. Self-service and transparency links**

Several links that almost no guest knows exist — but that matter:

- **Review your guest account**: The Microsoft MyAccount portal is the central self-service
  hub for a guest's account in your tenant. The tenant-scoped entry point is:
  `https://myaccount.microsoft.com/?tenantId={yourTenantId}`
  From there, guests can review their profile, see which organizations they belong to, and
  manage their presence in your tenant. Without this link, navigating there independently
  is genuinely difficult. Surfacing it on the Entrance Area respects informational
  self-determination and is a meaningful gesture toward GDPR compliance.

- **Leave this organization**: Guests have the right to remove their own guest account at
  any time. MyAccount supports this, but the option is buried several clicks deep. A direct
  link takes them straight to the confirmation page for your specific tenant:
  `https://myaccount.microsoft.com/organizations/leave/{yourTenantId}?tenant={yourTenantId}`
  Replace both occurrences of `{yourTenantId}` with your tenant's GUID. Publishing this
  link openly is a genuine transparency signal — it tells guests that their presence in your
  environment is their choice, and that leaving is straightforward.

- **View accepted terms and conditions**: If your organization uses Conditional Access
  policies that require guests to accept terms of use before accessing resources, most guests
  have no idea they can review what they agreed to. The MyAccount portal surfaces this
  under a dedicated terms acceptance page. A direct link:
  `https://myaccount.microsoft.com/termsofuse/myacceptances?tenantId={yourTenantId}`
  Providing this link gives guests the transparency they are entitled to — and rarely receive.

- **Manage security information**: For guests whose multi-factor authentication is registered
  directly in your tenant — rather than being trusted from their home organization — the
  My Sign-Ins portal is where they manage their MFA methods, including the Authenticator
  app registration that gets created when MFA is set up in your tenant:
  `https://mysignins.microsoft.com/security-info?tenantId={yourTenantId}`
  Most guests will never need this link. But for those who do, knowing it exists prevents a
  support call. And it is also a reminder of the one scenario the Entrance Area cannot help
  with at all: if a guest's MFA methods registered in your tenant stop working — a lost
  phone, a reset Authenticator app — and they have no fallback method configured, they are
  locked out before they can reach any page you have built. A reset process then requires
  the guest to contact their internal point of contact (ideally their sponsor), who must
  contact your helpdesk on their behalf, who in turn must be able to verify that the caller
  is actually that guest's sponsor — because only then can someone credibly vouch that they
  know the guest and confirm their identity. The chain is long, the verification is difficult,
  and the contact details that would have made it easier are sitting behind the login screen
  nobody can reach. This is one of the stronger arguments for keeping MFA in the home tenant
  wherever possible: the problem simply does not exist if there is nothing to reset in yours.

**7. Contact and support information**

Who to call or message if something is not working. A name and a channel, not just a generic
helpdesk alias. This one is consistently overlooked and consistently appreciated.

---

### One More Thing: A URL Worth Sharing

Even a well-designed Entrance Area is only useful if guests can find it again. Alongside a
memorable SharePoint URL, consider publishing a short link through a corporate link
shortening service — something like `go.contoso.com/entrance` or a similar pattern. A short,
stable URL you can drop into a Teams message, print on a welcome slide, or mention verbally
in an onboarding call is far easier to communicate than a full SharePoint path. It also gives
you a single, consistent address to update in one place if the underlying site ever moves.

---

## One Feature Worth Adding: Sponsor Visibility

When working with organizations on exactly the challenges described in this article, we kept
running into the same cluster of gaps in the out-of-box SharePoint experience for guests.
One of the most persistent: by the time a guest actually needs help, they often have no
practical way to reach the right person.

The invitation email does name whoever sent it — but that email is usually weeks or months
old by the time something goes wrong. It may have been sent by an automated provisioning
system rather than the actual person responsible for the collaboration. It gives an email
address, not a Teams chat link or a phone number or any indication of whether that person is
still the right contact. And in many cases, they are not: the original inviter may have moved
to a different team, left the organization, or simply handed the project to someone else.
None of that is reflected in the invitation email, because the invitation email is a
snapshot — it captures who clicked "Invite" on a specific day, not who is currently
responsible for the guest's access.

### The Guest Sponsor Info Web Part

So we built something. The **Guest Sponsor Info** web part is a free, open-source SharePoint
web part developed by Workoho and published for the Microsoft 365 community. It was designed
specifically for the Entrance Area scenario: a single web part that closes several gaps at
once, without requiring per-guest configuration or manual maintenance.

In Microsoft Entra, every guest user has one or more assigned *sponsors* — internal employees
who requested or approved the guest's access. This information sits in the directory, but
guests never see it. And when a guest runs into trouble, or simply wants to know who to reach
out to — or who to go to when the primary contact is unavailable — this is exactly what they
need, and what nothing in the default SharePoint experience provides.

Microsoft supports up to five sponsors per guest account, so a second sponsor can formally
serve as a substitute rather than just an informal backup. In the absence of a dedicated
substitute, the manager is often the natural fallback — and the web part can surface the
manager's contact alongside the primary sponsor when that is useful. That said, more names do
not automatically mean better coverage. Once an account has four or five sponsors attached,
the practical effect is often that no one feels primarily responsible. Diffuse accountability
tends to behave like no accountability at all: when a lifecycle review comes up and someone
needs to respond, it becomes genuinely unclear whose job that is. We recommend keeping it to
two — one primary contact, one designated substitute. That is enough for real coverage
without enough ambiguity for everyone to assume someone else is handling it.

### The Sponsor Field: Useful but Static

There is a catch, though. Microsoft populates the sponsor field automatically at the moment
of invitation — whoever sent the invite becomes the sponsor. That is a sensible default, but
it is also where the story ends. Once set, the sponsor field has no built-in workflow around
it: there is no automated process to flag when a sponsor leaves the organization or changes
roles, no prompt to initiate a handover to someone else, and no lifecycle process that asks
anyone to review whether the guest's access is still appropriate. The information is there,
but it is essentially static — a snapshot of who clicked "Invite" at a given moment in time,
not an actively maintained responsibility.

### Lifecycle Management: Tooling Options

Microsoft does offer tooling for guest lifecycle management through **Microsoft Entra ID
Governance** — in particular, Access Reviews and Entitlement Management. These features can
address some of these gaps, but they come with P2 licensing requirements and a level of
complexity that makes them genuinely difficult to implement for most mid-sized organizations.
The cost alone is a barrier for many: Entra ID Governance is priced per user per month, and
for organizations with many guests, the numbers add up quickly — often to a point where the
case for implementation becomes hard to make.

**EasyLife 365** approaches the same problem from a different direction: practical, affordable
lifecycle management designed for organizations that want governance without building an
enterprise-grade IAM program around it. EasyLife 365 lets you actively manage the sponsor
relationship — including substitution and handover — and can trigger review and renewal
processes automatically. For many organizations, the cost savings from cleaning up stale guest
accounts and removing unnecessary access pay for the tool almost immediately. It is, for a
lot of customers, genuinely a no-brainer.

All of this matters for the Entrance Area because the value of the sponsor field — and of the
**Guest Sponsor Info** web part that surfaces it — depends directly on whether that
information is being maintained. A sponsor who is still shown as the contact six months after
they left the organization is not a safety net for the guest; it is a dead end. The web part
does its job best when the directory behind it is being actively managed.

### What It Shows on the Page

The web part reads the visiting guest's Entra profile, identifies their sponsor, and displays
that person's name, title, and contact details automatically — with no configuration required
per guest and no manual updates as sponsors change. It also handles the Teams readiness
scenario described above: if the guest does not yet have an active Teams presence in your
tenant, the web part detects this and adjusts the sponsor's contact card accordingly —
greying out the call and chat buttons and displaying a clear status message, so the guest
knows what to expect rather than wondering why nothing works.

It is a small addition to a landing page. But it reliably eliminates the "who do I even ask?"
moment that is common in the first days of cross-tenant work. The guest sees a face and a
name. The sponsor knows they have a responsibility. The connection becomes visible.

### Going Further: Guest Type Classification

Beyond the sponsor, EasyLife 365 also lets you classify guests into different types — vendor,
auditor, partner, contractor, and whatever else fits your model — and drive group memberships
automatically from that classification. With a small addition to the Entrance Area, you can
surface the guest's own type visibly on the page. For a guest, seeing "you are registered as
a Partner" is a small but meaningful signal that their access is intentional and structured,
not accidental. For the organization, it reinforces that the directory is governed — and that
guest categories translate into real access boundaries.

Explore the web part and its documentation here:
[Guest Sponsor Info on GitHub](https://github.com/workoho/spfx-guest-sponsor-info)

---

## This Is Also a Governance Question

A structured landing page for guests is not just about user experience. It is about directing
guests toward the right tools, channels, and workflows from the start — rather than leaving
them to find their way independently.

When guests orient themselves without guidance, they often end up in the wrong places:
personal OneDrive shares used as ad hoc collaboration spaces, informal Teams chats that bypass
structured channels, or external platforms that fall entirely outside your governance policies.
A landing page that actively guides them into the correct environment reduces this risk
meaningfully.

It also signals something about how your organization approaches external collaboration:
that it is structured, intentional, and governed — not an afterthought.

---

## Working on This Kind of Challenge

This is exactly the type of problem we at **Workoho** help organizations solve. As an
**EasyLife 365 Platinum Partner**, we work with IT teams who want to move beyond reactive
external sharing policies toward structured, auditable guest collaboration — where every
invited user knows where they are, what they can access, and who to talk to.

Think of us less as consultants who hand over a report, and less as implementers who simply
execute a spec. We work more like mountain guides: we help IT teams assess the terrain, plan
a realistic route, and then walk it together with them — technically, process-wise, and
strategically. We pick up tools when the situation calls for it, but our primary goal is to
guide and enable, so that your team comes out of the climb with the capability and confidence
to manage the path going forward. And if your organization works with an external IT service
provider for day-to-day operations, that is not a complication — we are equally happy to
work alongside them.

If guests are regularly asking "I got an invitation — what do I do now?", that is a signal
worth taking seriously before it becomes a governance issue.

Workoho is happy to have a practical conversation about what this could look like in your
environment. No predefined blueprint, no pressure — just a straightforward discussion about
your setup and what would genuinely help.
