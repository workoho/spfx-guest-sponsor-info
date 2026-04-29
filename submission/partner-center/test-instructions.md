# Test Instructions — Guest Sponsor Info (AppSource / Partner Center)

> **Status:** Draft — not yet uploaded to Partner Center.\
> **Version these instructions apply to:** 0.21.x\
> **Last updated:** 2026-03-26

---

## Overview

**Guest Sponsor Info** is a SharePoint Online web part that surfaces the
Microsoft Entra B2B sponsors of a guest user. On a shared guest landing page,
the web part shows every guest exactly who to call, chat with, or email —
live profile photo, contact details, and one-click Teams chat included.

For users who are **not** guests the web part renders nothing in view mode.
In **edit mode** it shows a static placeholder regardless of the visitor's
account type — no Microsoft Graph calls are made while editing.

---

## Test Environment (pre-provisioned)

The following has already been set up in the test tenant before submission.
No additional admin configuration is required.

| Item | Details |
|---|---|
| Tenant | `[TENANT].onmicrosoft.com` (to be filled in) |
| SharePoint site | `https://[TENANT].sharepoint.com/sites/guests` |
| Page with web part | `Guest Landing Page` (default page of the site) |
| `.sppkg` deployed | Site Collection App Catalog of the guests site |
| Web part added | Already on the page, pre-configured |

---

## Test Accounts

| Role | UPN | Password | Purpose |
|---|---|---|---|
| Internal editor | `reviewer-editor@[TENANT].onmicrosoft.com` | `[PASSWORD]` | Edit mode, property pane, Demo Mode |
| Guest user | `reviewer-guest_partner.com#EXT#@[TENANT].onmicrosoft.com` | `[PASSWORD]` | Real guest experience (Path B) |

The guest account is a B2B guest from a separate partner tenant. It has been
assigned two sponsors in Entra ID: one with a profile photo, one without
(to verify the initials-avatar fallback).

---

## Path A — Demo Mode (no guest account needed)

Use this path to verify all core UI features without signing in as a guest.
Demo Mode shows simulated sponsor cards to **any** signed-in user.

### Step 1 — Sign in as the internal editor

1. Open a private browser window.
2. Navigate to `https://[TENANT].sharepoint.com/sites/guests`.
3. Sign in with the **internal editor** account.

> **Expected:** You are taken to the guest landing page. The web part area is
> **empty** (no cards, no message) because Demo Mode is off by default when
> viewing as an internal user.

---

### Step 2 — Open the page in edit mode

1. Click **Edit** (pencil icon, top-right of the page).

> **Expected:** The web part zone shows a text placeholder indicating that the
> web part is in edit mode. Sponsor cards are not rendered in edit mode.

---

### Step 3 — Open the property pane

1. Click the web part to select it, then click the **pencil / edit** icon that
   appears in the web part toolbar.

> **Expected:** The property pane slides in from the right. It contains at least
> the following sections: **Settings**, **Guest notifications**, **Display**,
> **Advanced display**.

---

### Step 4 — Enable Demo Mode

1. In the property pane, scroll to the **Settings** section.
2. Toggle **Enable public demo mode for internal users** to **On**.

> **Expected:** A description explaining Demo Mode appears below the toggle.
> The web part canvas still shows the edit-mode placeholder
> ("Demo mode active — Switch to view mode to see mock sponsors.").

---

### Step 5 — Switch to view mode

1. Click **Publish** or use the page toolbar to switch to view mode.

> **Expected:** The web part now shows **2 mock sponsor cards** (the default
> count) with simulated names, job titles, and Workoho-coloured initials avatars.
> No real profile data is fetched.

---

### Step 6 — Inspect a sponsor card (hover / focus)

1. Hover over (or Tab to) one of the sponsor cards.

   **Expected:** A contact popover appears containing: display name and job
   title, email address with a copy button, business phone (if configured),
   Teams Chat and Teams Call buttons, office location / address fields (if
   configured), and a map link (external, since no Azure Maps key is set).

2. Click the **copy** button next to the email address.

   **Expected:** A checkmark appears for ~1.5 seconds, then reverts to the
   copy icon.

3. Click **Teams Chat**.

   **Expected:** The browser attempts to open the Teams client (or the Teams
   web app) in a new tab/window. The link uses the correct guest UPN format.

---

### Step 7 — Verify accessibility

1. With the web part in view, use **Tab** to navigate through the sponsor cards.
2. Press **Enter** or **Space** on a card.

> **Expected:** The contact popover opens and all interactive elements inside
> (copy buttons, Teams buttons, map link) are reachable by keyboard. The popover
> closes when focus moves away or **Escape** is pressed.

---

### Step 8 — Disable Demo Mode before Path B

1. Return to edit mode (click **Edit**).
2. Open the property pane and toggle **Demo Mode** back to **Off**.
3. Publish / save the page.

---

## Path B — Real Guest User

Use this path to validate the actual end-to-end guest experience.

### Step 1 — Sign in as the guest

1. Open a private browser window.
2. Navigate to `https://[TENANT].sharepoint.com/sites/guests`.
3. Sign in with the **guest user** account
   (`reviewer-guest_partner.com#EXT#@[TENANT].onmicrosoft.com`).

> **Expected:** The page loads. The web part automatically detects the `#EXT#`
> marker in the login name, calls Microsoft Graph (`/me/sponsors`), and renders
> the sponsor cards.
>
> - No "configure" dialog or sign-in prompt appears in the web part.
> - Two sponsor cards appear (matching the assigned sponsors in Entra ID).
> - One card shows a real profile photo; the other shows a coloured initials
>   avatar (because no photo is available for that sponsor).

---

### Step 2 — Sponsor cards in full layout

> **Expected card content:**
>
> - Sponsor display name (bold)
> - Job title (below name)
> - Profile photo or deterministic initials avatar with a unique colour derived
>   from the sponsor's name
>
> **Expected layout:** Full tile grid (default `Auto` mode at this column width).

---

### Step 3 — Contact popover

1. Hover over (or Tab to) the first sponsor card and wait for the popover.

   **Expected:** The popover shows real data pulled from Microsoft Graph:
   display name, email address with copy button, business phone / mobile (if
   in the sponsor's Entra profile), office location, Teams Chat and Teams Call
   buttons.

2. Click **Teams Chat**.

   **Expected:** Teams opens (or attempts to open) a chat with that sponsor,
   pre-addressed to the sponsor's UPN.

---

### Step 4 — Non-guest internal user sees nothing

1. In a separate private window, navigate to
   `https://[TENANT].sharepoint.com/sites/guests` and sign in with the
   **internal editor** account (ensure Demo Mode is still off from Step 8 above).

> **Expected:** With Demo Mode off, the web part renders **nothing** for an
> internal user in view mode — no cards, no error message, no empty space.

---

## Property Pane Feature Walkthrough

Sign in as the **internal editor** and open the page in edit mode.

### Settings group

| Control | What to test | Expected |
|---|---|---|
| **Visible sponsors** (slider, 1–5) | Drag to 1, then to 5 | Card count changes when Demo Mode is on and you save |
| **Enable public demo mode** (toggle) | Toggle on/off | Description text appears below the toggle |

### Guest notifications group

Toggles to show or suppress advisory banners for specific edge cases
(Teams access pending, version mismatch, sponsor unavailable, no sponsors).
These banners appear on the live page when the corresponding condition is
detected at runtime.

### Display group

| Control | What to test | Expected |
|---|---|---|
| **Card layout** (Auto / Full / Compact) | Switch between options | Card grid changes in Demo Mode view |
| **Show business phone** / **mobile** | Toggle off | Field disappears from the contact popover |
| **Show work location** / address fields | Toggle off | Location row disappears from the popover |
| **External map provider** (Bing / Google / Apple / OSM ) | Change selection | Map link updates accordingly |

### Advanced display group

Toggles for showing the sponsor's manager, presence status, and photos.

---

## API Permissions Verification

The web part requests **no Microsoft Graph permissions** of its own. No
`webApiPermissionRequests` are declared in the solution package.

To verify: open **SharePoint Admin Center → Advanced → API access**. There are
no pending or approved permissions originating from this solution — the queue
remains empty.

All Graph calls are made server-side by the companion Azure Function using its
Managed Identity. The web part authenticates exclusively against the Azure
Function's App Registration — there are no outbound requests to
`graph.microsoft.com` from the browser.

---

## Known Limitations (not defects)

- **Edit mode preview is static.** The web part intentionally shows only a
  text placeholder in edit mode. No Graph calls are made while editing.
  This is consistent with the SharePoint People web part behaviour.
- **Profile photos match what Microsoft 365 stores.** Low-resolution or
  missing photos reflect the user's Entra profile, not the web part.
- **Demo Mode uses mock data.** Names, job titles, and avatars in Demo Mode
  are simulated and unrelated to any actual users.

---

## Support Contact

Workoho GmbH · [support@workoho.com](mailto:support@workoho.com)\
GitHub: <https://github.com/workoho/spfx-guest-sponsor-info/issues>
