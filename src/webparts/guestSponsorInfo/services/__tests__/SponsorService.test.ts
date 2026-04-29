// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0

/**
 * Unit tests for SponsorService helper functions.
 *
 * All Graph calls have been moved to the Azure Function (proxy-only architecture).
 * The web part no longer calls Graph directly; all data arrives via AadHttpClient.
 * These tests cover the two client-side utilities that remain:
 *
 *   isGuestUser      — guest detection heuristic
 *   loadManagerPhotosViaProxy — lazy manager-photo fetching via the /api/getPhoto endpoint
 */
import { isGuestUser, loadManagerPhotosViaProxy, ISponsorsResult } from '../SponsorService';

// Re-export so the type is reachable from tests without a separate import cycle.
export type { ISponsorsResult };

// ─── Mock AadHttpClient ────────────────────────────────────────────────────────

interface MockResponse {
  ok: boolean;
  json: () => Promise<unknown>;
}

function buildAadClient(handler: (url: string) => Promise<MockResponse>): unknown {
  return {
    get: jest.fn(async (url: string) => handler(url)),
    configurations: { v1: {} },
  };
}

// ─── Fixtures ──────────────────────────────────────────────────────────────────

const SPONSOR_A_ID = 'aaaaaaaa-0000-0000-0000-000000000001';
const MANAGER_A_ID = 'cccccccc-0000-0000-0000-000000000003';

// ─── isGuestUser tests ─────────────────────────────────────────────────────────

describe('isGuestUser', () => {
  it('returns true when the loginName contains #EXT#', () => {
    expect(isGuestUser('alice_contoso.com#EXT#@fabrikam.onmicrosoft.com')).toBe(true);
  });

  it('returns false for an internal member UPN', () => {
    expect(isGuestUser('alice@contoso.com')).toBe(false);
  });

  it('returns false for an empty string', () => {
    expect(isGuestUser('')).toBe(false);
  });

  it('is case-sensitive: lowercase #ext# does not match', () => {
    // The Entra marker is always upper-case; the check mirrors that.
    expect(isGuestUser('alice_contoso.com#ext#@fabrikam.onmicrosoft.com')).toBe(false);
  });
});

// ─── loadManagerPhotosViaProxy tests ───────────────────────────────────────────

describe('loadManagerPhotosViaProxy', () => {
  it('calls onUpdate with the photoUrl when the proxy returns a photo', async () => {
    const photoDataUrl = 'data:image/jpeg;base64,/9j/abc123';
    const client = buildAadClient(async () => ({
      ok: true,
      json: async () => ({ photoUrl: photoDataUrl }),
    }));

    const onUpdate = jest.fn();
    const sponsors = [{ id: SPONSOR_A_ID, managerId: MANAGER_A_ID, displayName: 'Alice', businessPhones: [] }];

    loadManagerPhotosViaProxy('https://func.example.com/api/getPhoto', client as never, undefined, sponsors as never, onUpdate);
    await new Promise(resolve => setTimeout(resolve, 0));

    expect(onUpdate).toHaveBeenCalledTimes(1);
    expect(onUpdate.mock.calls[0][0]).toBe(SPONSOR_A_ID);
    expect(onUpdate.mock.calls[0][1]).toBe(photoDataUrl);
  });

  it('calls onUpdate with undefined when the proxy returns a non-ok response', async () => {
    const client = buildAadClient(async () => ({
      ok: false,
      json: async () => ({}),
    }));

    const onUpdate = jest.fn();
    const sponsors = [{ id: SPONSOR_A_ID, managerId: MANAGER_A_ID, displayName: 'Alice', businessPhones: [] }];

    loadManagerPhotosViaProxy('https://func.example.com/api/getPhoto', client as never, undefined, sponsors as never, onUpdate);
    await new Promise(resolve => setTimeout(resolve, 0));

    expect(onUpdate).toHaveBeenCalledTimes(1);
    expect(onUpdate.mock.calls[0][1]).toBeUndefined();
  });

  it('skips sponsors without a managerId', async () => {
    const client = buildAadClient(async () => ({
      ok: true,
      json: async () => ({ photoUrl: 'data:image/jpeg;base64,abc' }),
    }));

    const onUpdate = jest.fn();
    // No managerId on this sponsor.
    const sponsors = [{ id: SPONSOR_A_ID, displayName: 'Alice', businessPhones: [] }];

    loadManagerPhotosViaProxy('https://func.example.com/api/getPhoto', client as never, undefined, sponsors as never, onUpdate);
    await new Promise(resolve => setTimeout(resolve, 0));

    // No manager → nothing to fetch → onUpdate never called.
    expect(onUpdate).not.toHaveBeenCalled();
  });

  it('calls onUpdate with undefined when the network call throws', async () => {
    const client = buildAadClient(async () => { throw new Error('network error'); });

    const onUpdate = jest.fn();
    const sponsors = [{ id: SPONSOR_A_ID, managerId: MANAGER_A_ID, displayName: 'Alice', businessPhones: [] }];

    loadManagerPhotosViaProxy('https://func.example.com/api/getPhoto', client as never, undefined, sponsors as never, onUpdate);
    await new Promise(resolve => setTimeout(resolve, 0));

    expect(onUpdate).toHaveBeenCalledWith(SPONSOR_A_ID, undefined);
  });

  it('sends the X-Presence-Token header when presenceToken is provided', async () => {
    let capturedHeaders: Record<string, string> | undefined;
    const client = {
      get: jest.fn(async (_url: string, _config: unknown, options?: { headers?: Record<string, string> }) => {
        capturedHeaders = options?.headers;
        return { ok: true, json: async () => ({ photoUrl: 'data:image/jpeg;base64,abc' }) };
      }),
    };

    const onUpdate = jest.fn();
    const sponsors = [{ id: SPONSOR_A_ID, managerId: MANAGER_A_ID, displayName: 'Alice', businessPhones: [] }];

    loadManagerPhotosViaProxy('https://func.example.com/api/getPhoto', client as never, 'signed-token-value', sponsors as never, onUpdate);
    await new Promise(resolve => setTimeout(resolve, 0));

    expect(capturedHeaders?.['X-Presence-Token']).toBe('signed-token-value');
  });
});
