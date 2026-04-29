// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';
// Type-only imports are erased at compile time and generate no require() calls.
import type { IGuestSponsorInfoProps } from '../IGuestSponsorInfoProps';

// ─── Mock setup ────────────────────────────────────────────────────────────────
//
// This project runs jest against pre-compiled CJS output (lib-commonjs/) without
// applying babel-jest. That means jest.mock() is NOT automatically hoisted above
// require() calls. To guarantee jest.mock() runs before any module that imports
// SponsorService is loaded, we:
//   1. Keep only side-effect-free imports above (React, ReactDOM, act, import type).
//   2. Call jest.mock() here — in the compiled CJS it appears after the React requires
//      but BEFORE the inline require() calls below.
//   3. Load GuestSponsorInfo and SponsorService via inline require() so they pick up
//      the mock that was just registered.
//
// eslint-disable-next-line @rushstack/hoist-jest-mock
jest.mock('../../services/SponsorService');

// eslint-disable-next-line @typescript-eslint/no-var-requires
const GuestSponsorInfo = (require('../GuestSponsorInfo') as { default: React.ComponentType<IGuestSponsorInfoProps> }).default;

// eslint-disable-next-line @typescript-eslint/no-var-requires
const _service = require('../../services/SponsorService') as { isGuestUser: jest.Mock; getSponsorsViaProxy: jest.Mock };
const mockIsGuestUser = _service.isGuestUser;
const mockGetSponsors = _service.getSponsorsViaProxy;

// DisplayMode values that match the @microsoft/sp-core-library shim (Read=1, Edit=2).
// Imported via the shim; redefining here avoids importing the SPFx package in tests.
const DisplayMode = { Read: 1 as never, Edit: 2 as never };
const SESSION_CACHE_KEY_PREFIX = 'gsi:session-cache';
const SPONSOR_CACHE_CHANNEL_PREFIX = 'gsi:sponsor-cache';

class MockBroadcastChannel {
  private static readonly channels = new Map<string, Set<MockBroadcastChannel>>();

  public onmessage: ((event: MessageEvent) => void) | undefined;

  public constructor(public readonly name: string) {
    const listeners = MockBroadcastChannel.channels.get(name) ?? new Set<MockBroadcastChannel>();
    listeners.add(this);
    MockBroadcastChannel.channels.set(name, listeners);
  }

  public postMessage(message: unknown): void {
    const listeners = MockBroadcastChannel.channels.get(this.name);
    if (!listeners) return;

    listeners.forEach((peer) => {
      if (peer !== this) {
        peer.onmessage?.({ data: message } as MessageEvent);
      }
    });
  }

  public close(): void {
    const listeners = MockBroadcastChannel.channels.get(this.name);
    if (!listeners) return;

    listeners.delete(this);
    if (listeners.size === 0) {
      MockBroadcastChannel.channels.delete(this.name);
    }
  }

  public static reset(): void {
    MockBroadcastChannel.channels.clear();
  }
}

const originalBroadcastChannel = globalThis.BroadcastChannel;

beforeAll(() => {
  globalThis.BroadcastChannel = MockBroadcastChannel as unknown as typeof BroadcastChannel;
});

afterAll(() => {
  globalThis.BroadcastChannel = originalBroadcastChannel;
});

// ─── Fixture ───────────────────────────────────────────────────────────────────

const SPONSOR = {
  id: 'aaaaaaaa-1111-1111-1111-111111111111',
  displayName: 'Alice Smith',
  mail: 'alice@contoso.com',
  jobTitle: 'Project Manager',
  businessPhones: [],
  mobilePhone: undefined,
  photoUrl: undefined,
};

const SPONSOR_UNAVAILABLE = {
  id: 'bbbbbbbb-2222-2222-2222-222222222222',
  displayName: 'Bob Jones (unavailable)',
  businessPhones: [] as string[],
};

// ─── DOM helpers ───────────────────────────────────────────────────────────────

let container: HTMLDivElement;

beforeEach(() => {
  container = document.createElement('div');
  document.body.appendChild(container);
  jest.clearAllMocks();
  MockBroadcastChannel.reset();
  sessionStorage.clear();
  // Default: any login name containing #EXT# is treated as a guest.
  mockIsGuestUser.mockImplementation((name: string) => name.includes('#EXT#'));
});

afterEach(() => {
  act(() => { ReactDOM.unmountComponentAtNode(container); });
  container.remove();
});

function renderWebPart(overrides: Partial<IGuestSponsorInfoProps> = {}): void {
  const defaults: IGuestSponsorInfoProps = {
    loginName: 'guest_contoso.com#EXT#@fabrikam.onmicrosoft.com',
    isExternalGuestUser: true,
    displayMode: DisplayMode.Read,
    title: 'My Sponsors',
    mockMode: false,
    maxSponsorCount: 2,
    mockSponsorCount: 2,
    mockSimulatedHint: 'none',
    cardLayout: 'auto',
    cardLayoutAutoThreshold: 3,
    hostTenantId: 'aaaabbbb-0000-0000-0000-000000000001',
    functionUrl: 'https://func.example.com/api/getGuestSponsors',
    presenceUrl: undefined,
    pingUrl: undefined,
    photoUrl: undefined,
    webPartClientId: undefined,
    aadHttpClient: {} as never,
    showBusinessPhones: true,
    showMobilePhone: true,
    showWorkLocation: true,
    showCity: false,
    showCountry: false,
    showStreetAddress: false,
    showPostalCode: false,
    showState: false,
    sponsorFilter: 'teams',
    requireUserMailbox: true,
    sessionCacheTtlMinutes: 30,
    azureMapsSubscriptionKey: undefined,
    externalMapProvider: 'bing',
    showManager: true,
    showPresence: true,
    useInformalAddress: false,
    showSponsorJobTitle: true,
    showManagerJobTitle: true,
    showSponsorDepartment: false,
    showManagerDepartment: false,
    showSponsorPhoto: true,
    showManagerPhoto: true,
    showTeamsAccessPendingHint: true,
    showVersionMismatchHint: true,
    showSponsorUnavailableHint: true,
    showNoSponsorsHint: true,
    clientVersion: '0.0.1',
    welcomeSeen: false,
    onWelcomeComplete: () => undefined,
    onWelcomeSkip: () => undefined,
    onWelcomeFinish: () => undefined,
    fluentProviderId: 'gsi-test',
  };
  ReactDOM.render(<GuestSponsorInfo {...defaults} {...overrides} />, container);
}

/** Flushes all pending microtasks so async useEffect callbacks resolve. */
async function flushAsync(): Promise<void> {
  await act(async () => {
    await new Promise<void>(resolve => setTimeout(resolve, 50));
  });
}

// ─── Tests ─────────────────────────────────────────────────────────────────────

describe('GuestSponsorInfo', () => {

  // ── Edit mode ───────────────────────────────────────────────────────────────

  describe('edit mode', () => {
    it('shows mock sponsor cards for a guest author', () => {
      act(() => { renderWebPart({ displayMode: DisplayMode.Edit, welcomeSeen: true }); });
      expect(container.textContent).toContain('Anna Müller');
    });

    it('shows mock sponsor cards even for a non-guest author', () => {
      mockIsGuestUser.mockReturnValue(false);
      act(() => {
        renderWebPart({
          displayMode: DisplayMode.Edit,
          loginName: 'member@fabrikam.onmicrosoft.com',
          isExternalGuestUser: false,
          welcomeSeen: true,
        });
      });
      expect(container.textContent).toContain('Anna Müller');
    });

    it('never calls getSponsorsViaProxy in edit mode', () => {
      act(() => { renderWebPart({ displayMode: DisplayMode.Edit, welcomeSeen: true }); });
      expect(mockGetSponsors).not.toHaveBeenCalled();
    });
  });

  // ── View mode - non-guest visitor ────────────────────────────────────────────

  describe('view mode - non-guest visitor', () => {
    it('renders nothing so the web part is invisible to members', () => {
      mockIsGuestUser.mockReturnValue(false);
      act(() => {
        renderWebPart({ loginName: 'member@fabrikam.onmicrosoft.com', isExternalGuestUser: false });
      });
      expect(container.firstChild).toBeNull();
    });

    it('never calls getSponsorsViaProxy for a non-guest visitor', () => {
      mockIsGuestUser.mockReturnValue(false);
      act(() => { renderWebPart({ loginName: 'member@fabrikam.onmicrosoft.com', isExternalGuestUser: false }); });
      expect(mockGetSponsors).not.toHaveBeenCalled();
    });
  });

  // ── View mode - guest visitor ────────────────────────────────────────────────

  describe('view mode - guest visitor', () => {
    it('does not call getSponsorsViaProxy when neither functionUrl nor aadHttpClient is provided', () => {
      act(() => { renderWebPart({ functionUrl: undefined, aadHttpClient: undefined }); });
      expect(mockGetSponsors).not.toHaveBeenCalled();
    });

    it('shows the skeleton placeholder while getSponsorsViaProxy is still pending', () => {
      // Return a promise that never resolves so loading state persists throughout the test.
      mockGetSponsors.mockReturnValue(new Promise(() => { /* intentionally pending */ }));
      act(() => { renderWebPart({}); });
      // The loading text is replaced by a shimmer skeleton grid (aria-busy="true").
      expect(container.querySelector('[aria-busy="true"]')).not.toBeNull();
    });

    it('calls getSponsorsViaProxy exactly once on first render', async () => {
      mockGetSponsors.mockResolvedValue({ activeSponsors: [], unavailableCount: 0 });
      act(() => { renderWebPart({}); });
      await flushAsync();
      expect(mockGetSponsors).toHaveBeenCalledTimes(1);
    });

    it('reuses the sessionStorage sponsor cache and skips getSponsorsViaProxy', async () => {
      sessionStorage.setItem(
        `${SESSION_CACHE_KEY_PREFIX}:gsi:sponsors:guest_contoso.com#ext#@fabrikam.onmicrosoft.com:https://func.example.com/api/getguestsponsors:teams:mailbox`,
        JSON.stringify({
          ts: Date.now(),
          clientVersion: '0.0.1',
          activeSponsors: [SPONSOR],
          unavailableSponsors: [],
          sponsorOrder: [SPONSOR.id],
        })
      );

      act(() => { renderWebPart({}); });
      await flushAsync();

      expect(container.textContent).toContain('Alice Smith');
      expect(mockGetSponsors).not.toHaveBeenCalled();
    });

    it('hydrates sponsor cache from another tab via BroadcastChannel and skips getSponsorsViaProxy', async () => {
      const cacheKey = 'gsi:sponsors:guest_contoso.com#ext#@fabrikam.onmicrosoft.com:https://func.example.com/api/getguestsponsors:teams:mailbox';
      const peerChannel = new MockBroadcastChannel(`${SPONSOR_CACHE_CHANNEL_PREFIX}:${cacheKey}`);

      peerChannel.onmessage = (event) => {
        const message = event.data as { type: string; requestId?: string; cacheKey?: string };
        if (message.type !== 'request' || !message.requestId || !message.cacheKey) return;

        peerChannel.postMessage({
          type: 'response',
          cacheKey: message.cacheKey,
          requestId: message.requestId,
          payload: {
            ts: Date.now(),
            clientVersion: '0.0.1',
            activeSponsors: [SPONSOR],
            unavailableSponsors: [],
            sponsorOrder: [SPONSOR.id],
          },
        });
      };

      act(() => { renderWebPart({}); });
      await flushAsync();

      expect(container.textContent).toContain('Alice Smith');
      expect(mockGetSponsors).not.toHaveBeenCalled();

      peerChannel.close();
    });

    it('renders sponsor cards when getSponsorsViaProxy resolves with active sponsors', async () => {
      mockGetSponsors.mockResolvedValue({ activeSponsors: [SPONSOR], unavailableCount: 0 });
      act(() => { renderWebPart({}); });
      await flushAsync();
      expect(container.textContent).toContain('Alice Smith');
    });

    it('renders the title heading above the sponsor list', async () => {
      mockGetSponsors.mockResolvedValue({ activeSponsors: [SPONSOR], unavailableCount: 0 });
      act(() => { renderWebPart({ title: 'Your Sponsors' }); });
      await flushAsync();
      expect(container.querySelector('h2')!.textContent).toBe('Your Sponsors');
    });

    it('omits the title heading when title is empty', async () => {
      mockGetSponsors.mockResolvedValue({ activeSponsors: [SPONSOR], unavailableCount: 0 });
      act(() => { renderWebPart({ title: '' }); });
      await flushAsync();
      expect(container.querySelector('h2')).toBeNull();
    });

    it('shows the "no sponsors" message when the list is empty and none were unavailable', async () => {
      mockGetSponsors.mockResolvedValue({ activeSponsors: [], unavailableCount: 0 });
      act(() => { renderWebPart({}); });
      await flushAsync();
      expect(container.textContent).toContain('No sponsors found');
    });

    it('shows the "sponsor unavailable" message when all assigned sponsors are gone', async () => {
      // The guest has sponsors assigned in Entra, but every one of them was deleted.
      mockGetSponsors.mockResolvedValue({
        activeSponsors: [],
        unavailableCount: 1,
        unavailableSponsors: [SPONSOR_UNAVAILABLE],
        sponsorOrder: [SPONSOR_UNAVAILABLE.id],
      });
      act(() => { renderWebPart({}); });
      await flushAsync();
      expect(container.textContent).toContain('no longer available');
    });

    it('shows unavailable sponsor tiles alongside the "unavailable" banner', async () => {
      mockGetSponsors.mockResolvedValue({
        activeSponsors: [],
        unavailableCount: 1,
        unavailableSponsors: [SPONSOR_UNAVAILABLE],
        sponsorOrder: [SPONSOR_UNAVAILABLE.id],
      });
      act(() => { renderWebPart({}); });
      await flushAsync();
      expect(container.textContent).toContain('Bob Jones (unavailable)');
    });

    it('lets active sponsors nachrücken when a higher-priority sponsor is unavailable', async () => {
      // Order: SPONSOR_UNAVAILABLE first, then SPONSOR (active).
      // With maxSponsorCount=1, SPONSOR should fill the single visible slot.
      mockGetSponsors.mockResolvedValue({
        activeSponsors: [SPONSOR],
        unavailableCount: 1,
        unavailableSponsors: [SPONSOR_UNAVAILABLE],
        sponsorOrder: [SPONSOR_UNAVAILABLE.id, SPONSOR.id],
      });
      act(() => { renderWebPart({ maxSponsorCount: 1 }); });
      await flushAsync();
      // Active sponsor should be visible.
      expect(container.textContent).toContain('Alice Smith');
      // Unavailable sponsor should also appear (read-only tile) since it came first.
      expect(container.textContent).toContain('Bob Jones (unavailable)');
      // Banner should NOT appear because there is one active sponsor visible.
      expect(container.textContent).not.toContain('no longer available');
    });

    it('shows remaining active sponsors even when some are unavailable', async () => {
      // One sponsor is still active; the deleted one should not prevent rendering.
      mockGetSponsors.mockResolvedValue({
        activeSponsors: [SPONSOR],
        unavailableCount: 1,
        unavailableSponsors: [SPONSOR_UNAVAILABLE],
        sponsorOrder: [SPONSOR.id, SPONSOR_UNAVAILABLE.id],
      });
      act(() => { renderWebPart({}); });
      await flushAsync();
      expect(container.textContent).toContain('Alice Smith');
      // The "no longer available" banner must not appear when at least one card is shown.
      expect(container.textContent).not.toContain('no longer available');
    });

    it('shows the error message when getSponsorsViaProxy rejects with a 4xx status', async () => {
      // Transient (network) errors are retried silently; only 4xx errors surface immediately.
      const err = Object.assign(new Error('Unauthorized'), { statusCode: 401 });
      mockGetSponsors.mockRejectedValue(err);
      act(() => { renderWebPart({}); });
      await flushAsync();
      expect(container.textContent).toContain('Could not load sponsor information');
    });

    it('shows the temporary service issue message for GRAPH_AUTHORIZATION_FAILED', async () => {
      const err = Object.assign(new Error('Forbidden'), {
        statusCode: 403,
        reasonCode: 'GRAPH_AUTHORIZATION_FAILED',
      });
      mockGetSponsors.mockRejectedValue(err);
      act(() => { renderWebPart({}); });
      await flushAsync();
      expect(container.textContent).toContain('Temporary service issue');
      expect(container.textContent).toContain('internal authorization step');
    });
  });
});
