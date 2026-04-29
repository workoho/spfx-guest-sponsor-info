// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0
//
// Assigns the Microsoft Graph application permissions required by the
// Azure Function's system-assigned Managed Identity.
//
// Permissions assigned:
//   User.Read.All         (df021288-...) — required: read sponsor profiles and photos
//   Presence.Read.All     (a70e0c2d-...) — optional: requires Teams at runtime
//   MailboxSettings.Read  (40f97065-...) — optional: filter shared/room/equipment mailboxes
//   TeamMember.Read.All   (660b7406-...) — optional: Teams membership as provisioning signal
//
// All four permissions are available to assign regardless of tenant licensing.
// Licensing only affects data access at runtime — the function degrades gracefully
// when Teams or Exchange Online is not licensed.
//
// SECURITY NOTE: These are *application* permissions (not delegated). The
// Managed Identity acts as itself — there is no signed-in user context. Scope is
// limited to the four role IDs listed above; no tenant-wide admin consent beyond
// these is required. If your security policy prohibits assigning these during
// Bicep deployment, set skipRoleAssignments=true and run setup-graph-permissions.ps1
// after deployment instead.

extension microsoftGraphV1

@description('Object ID (principalId) of the Function App system-assigned Managed Identity.')
param managedIdentityObjectId string

@description('When true, all role assignments are skipped. Use when Graph role assignment requires a separate Privileged Role Administrator step.')
param skipRoleAssignments bool

// ── Microsoft Graph service principal (well-known, tenant-stable) ─────────────
// appId 00000003-0000-0000-c000-000000000000 is the global identifier for
// Microsoft Graph — it is the same in every Entra tenant worldwide.
// Referenced as resourceId in the appRoleAssignedTo resources below.
resource graphSp 'Microsoft.Graph/servicePrincipals@v1.0' existing = if (!skipRoleAssignments) {
  appId: '00000003-0000-0000-c000-000000000000'
}

// BCP318: safe — graphSp is always resolved when !skipRoleAssignments.
#disable-next-line BCP318
var graphSpObjectId = graphSp.id

// ── Graph application role assignments → Managed Identity ────────────────────
// Each appRoleAssignedTo resource creates one Graph application permission grant.
// The Bicep Microsoft Graph extension writes these directly to the Entra tenant
// via the Graph API — no separate portal step is needed.

resource graphRoleUserReadAll 'Microsoft.Graph/appRoleAssignedTo@v1.0' = if (!skipRoleAssignments) {
  // User.Read.All — required: read sponsors, their profiles, and profile photos.
  appRoleId: 'df021288-bdef-4463-88db-98f22de89214'
  principalId: managedIdentityObjectId
  resourceId: graphSpObjectId
}

resource graphRolePresenceReadAll 'Microsoft.Graph/appRoleAssignedTo@v1.0' = if (!skipRoleAssignments) {
  // Presence.Read.All — optional: show sponsor online/away/busy presence in the card.
  // Returns empty results at runtime when Teams is not licensed; no error is thrown.
  appRoleId: 'a70e0c2d-e793-494c-94c4-118fa0a67f42'
  principalId: managedIdentityObjectId
  resourceId: graphSpObjectId
}

resource graphRoleMailboxSettingsRead 'Microsoft.Graph/appRoleAssignedTo@v1.0' = if (!skipRoleAssignments) {
  // MailboxSettings.Read — optional: filter out shared/room/equipment mailboxes.
  // Degrades gracefully when Exchange Online is not available.
  appRoleId: '40f97065-369a-49f4-947c-6a255697ae91'
  principalId: managedIdentityObjectId
  resourceId: graphSpObjectId
}

resource graphRoleTeamMemberReadAll 'Microsoft.Graph/appRoleAssignedTo@v1.0' = if (!skipRoleAssignments) {
  // TeamMember.Read.All — optional: verify Teams membership as a provisioning signal.
  // Degrades gracefully when Teams is not licensed.
  appRoleId: '660b7406-55f1-41ca-a0ed-0b035e182f3e'
  principalId: managedIdentityObjectId
  resourceId: graphSpObjectId
}
