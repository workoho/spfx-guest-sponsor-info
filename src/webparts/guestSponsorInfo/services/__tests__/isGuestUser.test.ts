// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0

import { isGuestUser } from '../SponsorService';

describe('isGuestUser', () => {
  it('returns true for a standard guest UPN containing #EXT#', () => {
    expect(isGuestUser('alice_contoso.com#EXT#@fabrikam.onmicrosoft.com')).toBe(true);
  });

  it('returns true when #EXT# appears anywhere in the login name', () => {
    // Some tenants use longer UPN prefixes due to special characters in the home UPN.
    expect(isGuestUser('alice+tag_contoso.com#EXT#@fabrikam.onmicrosoft.com')).toBe(true);
  });

  it('returns false for a regular member UPN without #EXT#', () => {
    expect(isGuestUser('alice@fabrikam.onmicrosoft.com')).toBe(false);
  });

  it('returns false for a service account UPN', () => {
    expect(isGuestUser('svc-sharepoint@fabrikam.onmicrosoft.com')).toBe(false);
  });

  it('returns false for an empty string', () => {
    expect(isGuestUser('')).toBe(false);
  });

  it('is case-sensitive – Entra UPNs always use uppercase #EXT#', () => {
    expect(isGuestUser('alice_contoso.com#ext#@fabrikam.onmicrosoft.com')).toBe(false);
  });
});
