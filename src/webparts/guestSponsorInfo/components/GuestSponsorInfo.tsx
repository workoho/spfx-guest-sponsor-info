// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0

import * as React from 'react';
import { DisplayMode } from '@microsoft/sp-core-library';
import { FluentProvider, MessageBar, MessageBarBody, Skeleton, SkeletonItem, makeStyles, mergeClasses, tokens, typographyStyles, webLightTheme, webDarkTheme } from '@fluentui/react-components';
import type { Theme } from '@fluentui/react-components';
import { createV9Theme } from '@fluentui/react-migration-v8-v9';
import { RendererProvider } from '@griffel/react';
import * as strings from 'GuestSponsorInfoWebPartStrings';
import type { IGuestSponsorInfoProps } from './IGuestSponsorInfoProps';
import { ISponsor } from '../services/ISponsor';
import { isGuestUser, getSponsorsViaProxy, pingProxy, loadManagerPhotosViaProxy, getPresencesViaProxy } from '../services/SponsorService';
import { MOCK_SPONSORS } from '../services/MockSponsorService';
import SponsorCard from './SponsorCard';
import WelcomeDialog from './WelcomeDialog';
import { griffelRenderer } from '../griffelRenderer';

const DEFAULT_SESSION_CACHE_TTL_MINUTES = 30;
const MIN_SESSION_CACHE_TTL_MINUTES = 2;
const MAX_SESSION_CACHE_TTL_MINUTES = 480;
const SESSION_CACHE_KEY_PREFIX = 'gsi:session-cache';
const SPONSOR_CACHE_CHANNEL_PREFIX = 'gsi:sponsor-cache';
const SPONSOR_CACHE_PEER_TIMEOUT_MS = 25;

interface ICachedSponsorsPayload {
  ts: number;
  clientVersion: string;
  activeSponsors: ISponsor[];
  unavailableSponsors: ISponsor[];
  sponsorOrder: string[];
  guestHasTeamsAccess?: boolean;
  presenceToken?: string;
  functionVersion?: string;
}

type ISponsorCacheWritePayload = Omit<ICachedSponsorsPayload, 'ts'>;

type ISponsorCacheMessage =
  | { type: 'request'; cacheKey: string; requestId: string }
  | { type: 'response'; cacheKey: string; requestId: string; payload: ICachedSponsorsPayload }
  | { type: 'invalidate'; cacheKey: string };

function buildSponsorCacheKey(
  loginName: string,
  functionUrl: string,
  sponsorFilter: 'any' | 'exchange' | 'teams',
  requireUserMailbox: boolean
): string {
  const normalizedUrl = functionUrl.replace(/\/+$/, '').toLowerCase();
  const normalizedLogin = loginName.trim().toLowerCase();
  return `gsi:sponsors:${normalizedLogin}:${normalizedUrl}:${sponsorFilter}:${requireUserMailbox ? 'mailbox' : 'license'}`;
}

function getSessionCacheTtlMs(minutes: number | undefined): number {
  const fallback = minutes ?? DEFAULT_SESSION_CACHE_TTL_MINUTES;
  const normalized = Number.isFinite(fallback) ? Math.floor(fallback) : DEFAULT_SESSION_CACHE_TTL_MINUTES;
  const clamped = Math.min(MAX_SESSION_CACHE_TTL_MINUTES, Math.max(MIN_SESSION_CACHE_TTL_MINUTES, normalized));
  return clamped * 60 * 1000;
}

function getSponsorCacheStorageKey(cacheKey: string): string {
  return `${SESSION_CACHE_KEY_PREFIX}:${cacheKey}`;
}

function validateSponsorCachePayload(
  parsed: Partial<ICachedSponsorsPayload>,
  cacheTtlMs: number,
  clientVersion: string
): ICachedSponsorsPayload | undefined {
  if (
    typeof parsed.ts !== 'number' ||
    Date.now() - parsed.ts > cacheTtlMs ||
    parsed.clientVersion !== clientVersion ||
    (typeof parsed.functionVersion === 'string' && parsed.functionVersion !== clientVersion)
  ) {
    return undefined;
  }

  return {
    ts: parsed.ts,
    clientVersion,
    activeSponsors: Array.isArray(parsed.activeSponsors) ? parsed.activeSponsors as ISponsor[] : [],
    unavailableSponsors: Array.isArray(parsed.unavailableSponsors) ? parsed.unavailableSponsors as ISponsor[] : [],
    sponsorOrder: Array.isArray(parsed.sponsorOrder)
      ? parsed.sponsorOrder.filter((id): id is string => typeof id === 'string')
      : [],
    guestHasTeamsAccess: typeof parsed.guestHasTeamsAccess === 'boolean'
      ? parsed.guestHasTeamsAccess
      : undefined,
    presenceToken: typeof parsed.presenceToken === 'string' ? parsed.presenceToken : undefined,
    functionVersion: typeof parsed.functionVersion === 'string' ? parsed.functionVersion : undefined,
  };
}

function createSponsorCacheChannel(cacheKey: string): BroadcastChannel | undefined {
  if (typeof BroadcastChannel === 'undefined') return undefined;
  try {
    return new BroadcastChannel(`${SPONSOR_CACHE_CHANNEL_PREFIX}:${cacheKey}`);
  } catch {
    return undefined;
  }
}

function deleteSponsorCache(cacheKey: string): void {
  try {
    sessionStorage.removeItem(getSponsorCacheStorageKey(cacheKey));
  } catch {
    // Ignore storage failures.
  }
}

function readSponsorCache(cacheKey: string, cacheTtlMs: number, clientVersion: string): ICachedSponsorsPayload | undefined {
  try {
    const storageKey = getSponsorCacheStorageKey(cacheKey);
    const raw = sessionStorage.getItem(storageKey);
    if (!raw) return undefined;

    const validated = validateSponsorCachePayload(JSON.parse(raw) as Partial<ICachedSponsorsPayload>, cacheTtlMs, clientVersion);
    if (!validated) {
      sessionStorage.removeItem(storageKey);
      return undefined;
    }
    return validated;
  } catch {
    return undefined;
  }
}

function writeSponsorCache(cacheKey: string, payload: ISponsorCacheWritePayload): void {
  try {
    sessionStorage.setItem(getSponsorCacheStorageKey(cacheKey), JSON.stringify({ ts: Date.now(), ...payload }));
  } catch {
    // Ignore storage quota / browser privacy mode failures.
  }
}

const useWebPartStyles = makeStyles({
  webPart: {
    // No top/side padding — the SharePoint section provides uniform spacing on
    // all sides, so adding top padding here would make the gap above the title
    // larger than the natural left/right gaps (mismatching first-party web parts
    // like Quick Links). Bottom padding keeps breathing room below the last card.
    padding: `0 0 ${tokens.spacingVerticalS}`,
    overflow: 'visible',
    maxWidth: '100%',
    boxSizing: 'border-box',
    containerType: 'inline-size',
  },
  webPartContent: {
    minHeight: '200px',
  },
  title: {
    // Font size and weight are NOT set here — each size variant (titleH2, titleH3,
    // titleH4, titleNormal) spreads the appropriate Fluent v9 typographyStyles entry
    // so that font size, weight, and line-height all come from the design system.
    // spacingVerticalL (16px) — half of the 35px Quick Links spacing, visually
    // balanced between the title and the sponsor card grid.
    margin: `0 0 ${tokens.spacingVerticalL}`,
    color: tokens.colorNeutralForeground1,
  },
  // Extra styles applied to the title only in edit mode: signals editability
  // matching the pattern used by Microsoft first-party web parts (Quick Links).
  titleEditable: {
    cursor: 'text',
    outline: 'none',
    borderRadius: tokens.borderRadiusMedium,
    // A transparent 1px border is always present so the layout does not shift
    // when the border becomes visible on focus (border takes up the same space
    // whether visible or not).
    border: `1px solid transparent`,
    // No padding — the text sits flush against the border, matching the
    // behaviour of Microsoft Quick Links in edit mode.
    // Only the top border needs compensation so the heading does not jump
    // down by 1px when edit mode activates.
    marginTop: '-1px',
    // minHeight keeps the area visible and clickable when no title text exists.
    minHeight: '1.2em',
    // Placeholder via :empty — works reliably because the onInput handler
    // clears innerHTML (removing any browser-inserted <br>) whenever the text
    // content is empty, ensuring this selector always matches.
    '&:empty::before': {
      content: 'attr(data-placeholder)',
      // colorNeutralForeground3 maps to the SharePoint theme's neutralTertiary
      // (~#B1AFAD in the default Office theme) — matching Microsoft's own
      // placeholder style in first-party web parts like Quick Links.
      color: tokens.colorNeutralForeground3,
    },
    '&:focus': {
      border: `1px solid ${tokens.colorBrandStroke1}`,
    },
  },
  // Typography variant classes applied on top of the base `title` class.
  // Each class spreads the matching Fluent v9 typographyStyles entry so that
  // font-size, font-weight, and line-height all come from the design system.
  // Responsive font-size overrides shrink the text on narrow containers.
  titleH2: {
    // title2: fontSizeHero700 = 28px, semibold — matches Quick Links "Heading 2"
    ...typographyStyles.title2,
  },
  titleH3: {
    // title3: fontSizeBase600 = 24px, semibold — matches Quick Links "Heading 3"
    ...typographyStyles.title3,
  },
  titleH4: {
    // subtitle1: fontSizeBase500 = 20px, semibold — matches Quick Links "Heading 4"
    ...typographyStyles.subtitle1,
  },
  titleNormal: {
    // 18 px regular — between subtitle1 (20 px) and subtitle2 (16 px), matches
    // the Quick Links "Normal" size which is not a standard Fluent token.
    fontSize: '18px',
    fontWeight: tokens.fontWeightRegular,
  },
  sponsorGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(136px, 1fr))',
    gap: tokens.spacingHorizontalL,
    listStyle: 'none',
    margin: '0',
    padding: '0',
    '@container (max-width: 319px)': {
      gap: tokens.spacingHorizontalS,
    },
    '@container (min-width: 320px) and (max-width: 479px)': {
      gap: tokens.spacingHorizontalM,
    },
  },
  sponsorGridCompact: {
    display: 'grid',
    gridTemplateColumns: '1fr',
    gap: tokens.spacingHorizontalXS,
    listStyle: 'none',
    margin: '0',
    padding: '0',
  },
  sponsorItem: {
    display: 'block',
  },
  teamsAccessBanner: {
    marginTop: tokens.spacingVerticalL,
    '@container (max-width: 319px)': {
      marginTop: tokens.spacingVerticalMNudge,
    },
  },
  // Skeleton reuses the card tile layout so loading shimmer matches the real
  // card dimensions pixel-perfectly. Only layout properties are needed here;
  // interactive properties (cursor, focus-visible) are irrelevant because
  // skeletonItem disables pointer-events.
  skeletonCard: {
    position: 'relative' as const,
    width: '100%',
    boxSizing: 'border-box' as const,
    minHeight: '122px',
    display: 'flex',
    flexDirection: 'column' as const,
    alignItems: 'center' as const,
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalM,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: 'transparent',
  },
  skeletonCardCompact: {
    position: 'relative' as const,
    display: 'inline-flex',
    flexDirection: 'row' as const,
    alignItems: 'center' as const,
    gap: tokens.spacingHorizontalMNudge,
    padding: tokens.spacingVerticalSNudge,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: 'transparent',
    maxWidth: '100%',
  },
  skeletonItem: {
    pointerEvents: 'none' as const,
    cursor: 'default',
  },
});

/**
 * Renders the sponsor grid and owns the single shared "active card" state.
 * Only one popup is ever visible at a time; switching cards cancels any
 * pending hide-timeout from the previously active card so there is no
 * overlap between the outgoing and incoming popup.
 */
interface ISponsorListProps {
  sponsors: ISponsor[];
  hostTenantId: string;
  /** When true, render compact horizontal cards instead of full 136px tiles. */
  compact: boolean;
  showBusinessPhones: boolean;
  showMobilePhone: boolean;
  showWorkLocation: boolean;
  showCity: boolean;
  showCountry: boolean;
  showStreetAddress: boolean;
  showPostalCode: boolean;
  showState: boolean;
  azureMapsSubscriptionKey: string | undefined;
  externalMapProvider: 'bing' | 'google' | 'apple' | 'openstreetmap' | 'none';
  showManager: boolean;
  showPresence: boolean;
  showSponsorJobTitle: boolean;
  showManagerJobTitle: boolean;
  showSponsorDepartment: boolean;
  showManagerDepartment: boolean;
  showSponsorPhoto: boolean;
  showManagerPhoto: boolean;
  useInformalAddress: boolean;
  onActiveCardChange?: (hasActiveCard: boolean) => void;
  /** Propagated from ISponsorsResult — false shows disabled buttons in each card. */
  guestHasTeamsAccess?: boolean;
  /**
   * Set of sponsor IDs that should be rendered as read-only tiles (no popup,
   * no keyboard activation). Used for disabled/resource-account sponsors that
   * are shown for context but cannot be interacted with.
   */
  readOnlyIds?: ReadonlySet<string>;
  /** Fluent v9 theme to forward into SponsorCard portal FluentProviders. */
  v9Theme?: Theme;
}

const SponsorList: React.FC<ISponsorListProps> = ({ sponsors, hostTenantId, compact, showBusinessPhones, showMobilePhone, showWorkLocation, showCity, showCountry, showStreetAddress, showPostalCode, showState, azureMapsSubscriptionKey, externalMapProvider, showManager, showPresence, showSponsorJobTitle, showManagerJobTitle, showSponsorDepartment, showManagerDepartment, showSponsorPhoto, showManagerPhoto, useInformalAddress, onActiveCardChange, guestHasTeamsAccess, readOnlyIds, v9Theme }) => {
  const [activeId, setActiveId] = React.useState<string | null>(null);
  const showTimeout = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  const hideTimeout = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  // Synchronous mirror of activeId — readable inside callbacks without stale closure.
  const activeIdRef = React.useRef<string | null>(null);
  // True while the currently active card was opened by an explicit click or focus.
  // When pinned, mouse-leave and blur do NOT schedule a deactivation, so the user
  // can move the cursor freely without accidentally closing the card.
  const isPinned = React.useRef(false);
  const classes = useWebPartStyles();

  // Delayed activation — hover path (500 ms delay).
  // Bails out early when the same card is already showing, so a re-enter after a
  // brief mouse-leave on the Popover surface does not restart the animation.
  const activate = (id: string): void => {
    if (hideTimeout.current) { clearTimeout(hideTimeout.current); hideTimeout.current = null; }
    if (activeIdRef.current === id) return; // already showing — nothing to do
    if (showTimeout.current) return;         // timer already running
    showTimeout.current = setTimeout(() => {
      showTimeout.current = null;
      isPinned.current = false;
      activeIdRef.current = id;
      setActiveId(id);
      onActiveCardChange?.(true);
    }, 500);
  };

  // Immediate activation — click / focus path.
  // Cancels any pending hover timer and shows the card synchronously.
  // Clicking the same card twice is a no-op (no toggling).
  const activateNow = (id: string): void => {
    if (hideTimeout.current) { clearTimeout(hideTimeout.current); hideTimeout.current = null; }
    if (showTimeout.current) { clearTimeout(showTimeout.current); showTimeout.current = null; }
    if (activeIdRef.current === id) return; // same card — don't toggle
    isPinned.current = true;
    activeIdRef.current = id;
    setActiveId(id);
    onActiveCardChange?.(true);
  };

  // Schedule deactivation — mouse-leave / blur path.
  // Ignored while a card is pinned (click-activated), so mouse movement over
  // other elements after a click does not accidentally close the card.
  const scheduleDeactivate = (): void => {
    if (showTimeout.current) { clearTimeout(showTimeout.current); showTimeout.current = null; }
    if (isPinned.current) return; // card was pinned by click — don't auto-close
    hideTimeout.current = setTimeout(() => {
      isPinned.current = false;
      activeIdRef.current = null;
      setActiveId(null);
      onActiveCardChange?.(false);
    }, 150);
  };

  // Forceful deactivation — explicit dismiss path (outside-click, Escape key,
  // mobile drawer close button). Always closes regardless of pin state.
  const forceDeactivate = (): void => {
    if (showTimeout.current) { clearTimeout(showTimeout.current); showTimeout.current = null; }
    if (hideTimeout.current) { clearTimeout(hideTimeout.current); hideTimeout.current = null; }
    isPinned.current = false;
    activeIdRef.current = null;
    setActiveId(null);
    onActiveCardChange?.(false);
  };

  React.useEffect(() => {
    return () => {
      if (showTimeout.current) clearTimeout(showTimeout.current);
      if (hideTimeout.current) clearTimeout(hideTimeout.current);
      isPinned.current = false;
      activeIdRef.current = null;
      onActiveCardChange?.(false);
    };
  }, [onActiveCardChange]);

  return (
    <ul className={compact ? classes.sponsorGridCompact : classes.sponsorGrid}>
      {sponsors.map(sponsor => {
        const isReadOnly = readOnlyIds?.has(sponsor.id) ?? false;
        return (
        <li key={sponsor.id} className={classes.sponsorItem}>
          <SponsorCard
            sponsor={sponsor}
            hostTenantId={hostTenantId}
            compact={compact}
            isActive={!isReadOnly && activeId === sponsor.id}
            onActivate={isReadOnly ? () => undefined : () => activate(sponsor.id)}
            onActivateNow={isReadOnly ? () => undefined : () => activateNow(sponsor.id)}
            onScheduleDeactivate={isReadOnly ? () => undefined : scheduleDeactivate}
            onForceDeactivate={isReadOnly ? () => undefined : forceDeactivate}
            readOnly={isReadOnly}
            showBusinessPhones={showBusinessPhones}
            showMobilePhone={showMobilePhone}
            showWorkLocation={showWorkLocation}
            showCity={showCity}
            showCountry={showCountry}
            showStreetAddress={showStreetAddress}
            showPostalCode={showPostalCode}
            showState={showState}
            azureMapsSubscriptionKey={azureMapsSubscriptionKey}
            externalMapProvider={externalMapProvider}
            showManager={showManager}
            showPresence={showPresence}
            showSponsorJobTitle={showSponsorJobTitle}
            showManagerJobTitle={showManagerJobTitle}
            showSponsorDepartment={showSponsorDepartment}
            showManagerDepartment={showManagerDepartment}
            showSponsorPhoto={showSponsorPhoto}
            showManagerPhoto={showManagerPhoto}
            useInformalAddress={useInformalAddress}
            guestHasTeamsAccess={guestHasTeamsAccess}
            v9Theme={v9Theme}
          />
        </li>
        );
      })}
    </ul>
  );
};

/**
 * Loading skeleton using Fluent UI Skeleton + SkeletonItem components.
 * Reuses the real card CSS classes so spacing is pixel-perfect.
 * Fluent’s Skeleton handles the shimmer animation and automatically
 * respects prefers-reduced-motion without custom media-query code.
 */
const SponsorGridSkeleton: React.FC<{ compact: boolean }> = ({ compact }) => {
  const classes = useWebPartStyles();
  return (
  <ul className={compact ? classes.sponsorGridCompact : classes.sponsorGrid} aria-busy="true">
    {[0, 1].map(i => (
      <li key={i} className={classes.sponsorItem}>
        {compact ? (
          <Skeleton
            className={mergeClasses(classes.skeletonCardCompact, classes.skeletonItem)}
            aria-label={strings.LoadingMessage}
          >
            <SkeletonItem shape="circle" size={40} />
            <SkeletonItem
              shape="rectangle"
              size={16}
              style={{ width: 120, borderRadius: 'var(--borderRadiusLarge, 6px)' }}
            />
          </Skeleton>
        ) : (
          <Skeleton
            className={mergeClasses(classes.skeletonCard, classes.skeletonItem)}
            aria-label={strings.LoadingMessage}
          >
            <SkeletonItem shape="circle" size={72} />
            <SkeletonItem
              shape="rectangle"
              size={12}
              style={{ width: 88, borderRadius: 'var(--borderRadiusLarge, 6px)' }}
            />
            <SkeletonItem
              shape="rectangle"
              size={12}
              style={{ width: 64, borderRadius: 'var(--borderRadiusLarge, 6px)' }}
            />
          </Skeleton>
        )}
      </li>
    ))}
  </ul>
  );
};

/**
 * Builds the two visible sponsor sets, respecting both the configured cap and
 * the original Entra ordering. Only active accounts count toward the cap;
 * unavailable accounts (disabled / deleted) are shown alongside the active
 * ones in their original position so the guest can always see who their
 * sponsors are, while active accounts "nachrücken" to fill the visible slots.
 *
 * Falls back to independent slice() when sponsorOrder is empty (e.g. older
 * Function versions without ordering info).
 */
function buildVisibleSponsorSets(
  activeSponsors: ISponsor[],
  unavailableSponsors: ISponsor[],
  sponsorOrder: string[],
  maxSponsorCount: number
): { visibleActive: ISponsor[]; visibleUnavailable: ISponsor[] } {
  if (sponsorOrder.length === 0) {
    return {
      visibleActive: activeSponsors.slice(0, maxSponsorCount),
      visibleUnavailable: unavailableSponsors.slice(0, maxSponsorCount),
    };
  }
  const activeMap = new Map(activeSponsors.map(s => [s.id, s]));
  const unavailableMap = new Map(unavailableSponsors.map(s => [s.id, s]));
  const visibleActive: ISponsor[] = [];
  const visibleUnavailable: ISponsor[] = [];
  for (const id of sponsorOrder) {
    if (visibleActive.length >= maxSponsorCount) break;
    const active = activeMap.get(id);
    if (active) { visibleActive.push(active); continue; }
    const unavail = unavailableMap.get(id);
    if (unavail) { visibleUnavailable.push(unavail); }
  }
  return { visibleActive, visibleUnavailable };
}

/** Maximum number of transient-error retries before giving up and showing an error. */
const MAX_RETRIES = 3;

const GuestSponsorInfo: React.FC<IGuestSponsorInfoProps> = ({
  loginName,
  isExternalGuestUser,
  displayMode,
  title,
  showTitle,
  titleSize,
  mockMode,
  mockSponsorCount,
  maxSponsorCount,
  mockSimulatedHint,
  showTeamsAccessPendingHint,
  showVersionMismatchHint,
  showSponsorUnavailableHint,
  showNoSponsorsHint,
  cardLayout,
  cardLayoutAutoThreshold,
  hostTenantId,
  functionUrl,
  presenceUrl,
  pingUrl,
  photoUrl,
  aadHttpClient,
  showBusinessPhones,
  showMobilePhone,
  showWorkLocation,
  showCity,
  showCountry,
  showStreetAddress,
  showPostalCode,
  showState,
  azureMapsSubscriptionKey,
  externalMapProvider,
  showManager,
  showPresence,
  showSponsorJobTitle,
  showManagerJobTitle,
  showSponsorDepartment,
  showManagerDepartment,
  showSponsorPhoto,
  showManagerPhoto,
  useInformalAddress,
  sponsorFilter,
  requireUserMailbox,
  sessionCacheTtlMinutes,
  clientVersion,
  onProxyStatusChange,
  onVersionMismatch,
  onTitleChange,
  welcomeSeen,
  onWelcomeComplete,
  onWelcomeSkip,
  onWelcomeFinish,
  fluentProviderId,
  theme,
}) => {
  // Slice the full MOCK_SPONSORS pool to the count configured in the property pane.
  const mockSponsors = MOCK_SPONSORS.slice(0, mockSponsorCount);

  // Helper: pick the informal string variant when useInformalAddress is enabled and
  // the current locale provides one (languages with T-V distinction like de, fr, es, it, nl).
  const fstr = <K extends keyof typeof strings>(key: K): string => {
    if (useInformalAddress) {
      const informalKey = `${key}Informal` as keyof typeof strings;
      const informal = strings[informalKey];
      if (informal) return informal as string;
    }
    return strings[key] as string;
  };

  // Primary signal: pageContext.user.isExternalGuestUser (authoritative, set from Entra token).
  // Fallback: #EXT# in loginName (heuristic; may be absent when the SharePoint user profile
  // has not yet been created for the guest, causing SP.UserProfile to return HTTP 500).
  const isGuest = isExternalGuestUser || isGuestUser(loginName);
  const isEditMode = displayMode === DisplayMode.Edit;
  const classes = useWebPartStyles();
  const sponsorCacheTtlMs = getSessionCacheTtlMs(sessionCacheTtlMinutes);
  const sponsorCacheKey = !isEditMode && !mockMode && isGuest && functionUrl
    ? buildSponsorCacheKey(loginName, functionUrl, sponsorFilter, requireUserMailbox)
    : undefined;
  const initialSponsorCacheRef = React.useRef<ICachedSponsorsPayload | undefined>(
    sponsorCacheKey ? readSponsorCache(sponsorCacheKey, sponsorCacheTtlMs, clientVersion) : undefined
  );
  const initialSponsorCache = initialSponsorCacheRef.current;

  const [sponsors, setSponsors] = React.useState<ISponsor[]>(initialSponsorCache?.activeSponsors ?? []);
  const [sponsorOrder, setSponsorOrder] = React.useState<string[]>(initialSponsorCache?.sponsorOrder ?? []);
  const [unavailableSponsors, setUnavailableSponsors] = React.useState<ISponsor[]>(initialSponsorCache?.unavailableSponsors ?? []);
  // Start in loading state immediately for guests and demo mode so the shimmer
  // is visible on the very first render — before the first useEffect tick.
  // Without this, React paints a brief "no sponsors" flash before the effect
  // can call setLoading(true).
  const [loading, setLoading] = React.useState(!isEditMode && (mockMode || (isGuest && !initialSponsorCache)));
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [isPermissionError, setIsPermissionError] = React.useState(false);
  const [retryCount, setRetryCount] = React.useState(0);
  const [hasActiveCard, setHasActiveCard] = React.useState(false);
  const [guestHasTeamsAccess, setGuestHasTeamsAccess] = React.useState<boolean | undefined>(
    initialSponsorCache?.guestHasTeamsAccess
  );
  const [versionMismatch, setVersionMismatch] = React.useState(false);

  // First-run welcome wizard: shown in edit mode until the admin completes it.
  // Skipping does NOT set the flag — the wizard reappears each edit session until
  // the admin actually chooses API or Demo mode. The flag is stored in the web part
  // property bag (SharePoint) so it is shared across all users and devices.
  const [showWelcomeDialog, setShowWelcomeDialog] = React.useState<boolean>(
    isEditMode && !welcomeSeen
  );
  // handleWelcomeCommit: saves the wizard result to web part properties but does NOT
  // close the dialog — the wizard advances to the Done step first so the admin sees
  // a genuine post-save confirmation before the pane opens.
  const handleWelcomeCommit = (config: import('./IGuestSponsorInfoProps').IWelcomeSetupConfig): void => {
    onWelcomeComplete(config);
  };
  const handleWelcomeDismiss = (): void => {
    setShowWelcomeDialog(false);
    // Open the property pane now that the wizard is gone so the admin can
    // review and fine-tune all settings.
    onWelcomeFinish();
  };
  // handleWelcomeSkip: X button or "Not now" — closes the dialog without
  // committing and delegates to the host to open the property pane.
  const handleWelcomeSkip = (): void => {
    setShowWelcomeDialog(false);
    onWelcomeSkip();
  };
  // Signed token issued by getGuestSponsors; passed to getPresence so the function
  // can validate sponsor IDs without server-side state or extra Graph calls.
  const [presenceToken, setPresenceToken] = React.useState<string | undefined>(initialSponsorCache?.presenceToken);

  // Ref that always holds the IDs of currently displayed sponsors.
  // The presence refresh interval reads this without capturing sponsors in its closure.
  const sponsorIdsRef = React.useRef<string[]>(initialSponsorCache?.activeSponsors.map(s => s.id) ?? []);
  const sponsorCacheMetaRef = React.useRef<ISponsorCacheWritePayload | undefined>(
    initialSponsorCache
      ? {
          activeSponsors: initialSponsorCache.activeSponsors,
          clientVersion,
          unavailableSponsors: initialSponsorCache.unavailableSponsors,
          sponsorOrder: initialSponsorCache.sponsorOrder,
          guestHasTeamsAccess: initialSponsorCache.guestHasTeamsAccess,
          presenceToken: initialSponsorCache.presenceToken,
          functionVersion: initialSponsorCache.functionVersion,
        }
      : undefined
  );
  const sponsorCacheDisabledRef = React.useRef(false);
  const sponsorCacheChannelRef = React.useRef<BroadcastChannel | undefined>(undefined);
  const pendingSponsorCacheRequestsRef = React.useRef(new Map<string, {
    resolve: (payload: ICachedSponsorsPayload | undefined) => void;
    timeoutId: ReturnType<typeof setTimeout>;
  }>());
  const primePresenceRefreshRef = React.useRef(!!initialSponsorCache);

  const invalidateSponsorCache = (): void => {
    if (!sponsorCacheKey) return;
    sponsorCacheDisabledRef.current = true;
    sponsorCacheMetaRef.current = undefined;
    deleteSponsorCache(sponsorCacheKey);
    sponsorCacheChannelRef.current?.postMessage({ type: 'invalidate', cacheKey: sponsorCacheKey } satisfies ISponsorCacheMessage);
  };

  const requestSponsorCacheFromPeers = (): Promise<ICachedSponsorsPayload | undefined> => {
    const channel = sponsorCacheChannelRef.current;
    if (!channel || !sponsorCacheKey || sponsorCacheDisabledRef.current) {
      return Promise.resolve(undefined);
    }

    const requestId = typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
      ? crypto.randomUUID()
      : `${Date.now()}-${Math.random().toString(36).slice(2)}`;

    return new Promise(resolve => {
      const timeoutId = setTimeout(() => {
        pendingSponsorCacheRequestsRef.current.delete(requestId);
        resolve(undefined);
      }, SPONSOR_CACHE_PEER_TIMEOUT_MS);

      pendingSponsorCacheRequestsRef.current.set(requestId, { resolve, timeoutId });
      channel.postMessage({ type: 'request', cacheKey: sponsorCacheKey, requestId } satisfies ISponsorCacheMessage);
    });
  };

  const persistSponsorCache = (
    activeSponsors: ISponsor[],
    overrides: Partial<Omit<ISponsorCacheWritePayload, 'activeSponsors'>> = {}
  ): void => {
    if (!sponsorCacheKey || sponsorCacheDisabledRef.current) return;
    const payload: ISponsorCacheWritePayload = {
      activeSponsors,
      clientVersion,
      unavailableSponsors: 'unavailableSponsors' in overrides
        ? overrides.unavailableSponsors ?? []
        : sponsorCacheMetaRef.current?.unavailableSponsors ?? [],
      sponsorOrder: 'sponsorOrder' in overrides
        ? overrides.sponsorOrder ?? activeSponsors.map(s => s.id)
        : sponsorCacheMetaRef.current?.sponsorOrder ?? activeSponsors.map(s => s.id),
      guestHasTeamsAccess: 'guestHasTeamsAccess' in overrides
        ? overrides.guestHasTeamsAccess
        : sponsorCacheMetaRef.current?.guestHasTeamsAccess,
      presenceToken: 'presenceToken' in overrides
        ? overrides.presenceToken
        : sponsorCacheMetaRef.current?.presenceToken,
      functionVersion: 'functionVersion' in overrides
        ? overrides.functionVersion
        : sponsorCacheMetaRef.current?.functionVersion,
    };
    if (payload.functionVersion && payload.functionVersion !== clientVersion) {
      invalidateSponsorCache();
      return;
    }
    sponsorCacheMetaRef.current = payload;
    writeSponsorCache(sponsorCacheKey, payload);
  };

  React.useEffect(() => {
    if (!sponsorCacheKey) return;

    const channel = createSponsorCacheChannel(sponsorCacheKey);
    sponsorCacheChannelRef.current = channel;
    if (!channel) {
      return () => { sponsorCacheChannelRef.current = undefined; };
    }

    channel.onmessage = (event: MessageEvent<ISponsorCacheMessage>) => {
      const message = event.data;
      if (!message || message.cacheKey !== sponsorCacheKey) return;

      if (message.type === 'request') {
        if (sponsorCacheDisabledRef.current) return;
        const cached = readSponsorCache(sponsorCacheKey, sponsorCacheTtlMs, clientVersion);
        if (!cached) return;
        channel.postMessage({
          type: 'response',
          cacheKey: sponsorCacheKey,
          requestId: message.requestId,
          payload: cached,
        } satisfies ISponsorCacheMessage);
        return;
      }

      if (message.type === 'response') {
        const pending = pendingSponsorCacheRequestsRef.current.get(message.requestId);
        if (!pending) return;

        clearTimeout(pending.timeoutId);
        pendingSponsorCacheRequestsRef.current.delete(message.requestId);
        pending.resolve(validateSponsorCachePayload(message.payload, sponsorCacheTtlMs, clientVersion));
        return;
      }

      sponsorCacheMetaRef.current = undefined;
      deleteSponsorCache(sponsorCacheKey);
    };

    return () => {
      sponsorCacheChannelRef.current = undefined;
      pendingSponsorCacheRequestsRef.current.forEach(({ resolve, timeoutId }) => {
        clearTimeout(timeoutId);
        resolve(undefined);
      });
      pendingSponsorCacheRequestsRef.current.clear();
      channel.close();
    };
  }, [sponsorCacheKey, sponsorCacheTtlMs, clientVersion]);

  // Edit-mode proxy health check: verify the Azure Function is reachable while the
  // page author has the web part selected.  Uses the lightweight /api/ping endpoint
  // so the check works for any authenticated user — not just guests.
  // The ping response also carries `x-api-version`; if that differs from the web
  // part's own manifest version we surface the version-mismatch banner to the editor
  // so they are informed regardless of the guest-facing showVersionMismatchHint setting.
  React.useEffect(() => {
    if (!isEditMode || !pingUrl) return;
    if (!aadHttpClient) { onProxyStatusChange?.('error'); return; }
    let cancelled = false;
    onProxyStatusChange?.('checking');
    pingProxy(pingUrl, aadHttpClient)
      .then((functionVersion) => {
        if (!cancelled) {
          onProxyStatusChange?.('ok');
          const detected = !!(clientVersion && functionVersion && clientVersion !== functionVersion);
          setVersionMismatch(detected);
          onVersionMismatch?.(detected);
        }
      })
      .catch(() => { if (!cancelled) { onProxyStatusChange?.('error'); } });
    return () => { cancelled = true; };
  }, [isEditMode, pingUrl, aadHttpClient, clientVersion]);

  React.useEffect(() => {
    if (isEditMode) return;

    // Demo mode: use static mock data, no Graph calls needed.
    if (mockMode) {
      // "No sponsors found" — no sponsors are assigned, so no tiles appear.
      if (mockSimulatedHint === 'noSponsors') {
        setSponsors([]);
        setSponsorOrder([]);
        setUnavailableSponsors([]);
      } else if (mockSimulatedHint === 'sponsorUnavailable') {
        // All sponsors are unavailable — show their tiles read-only.
        setSponsors([]);
        setSponsorOrder(mockSponsors.map(s => s.id));
        setUnavailableSponsors(mockSponsors);
      } else {
        setSponsors(mockSponsors);
        setSponsorOrder(mockSponsors.map(s => s.id));
        setUnavailableSponsors([]);
      }
      setGuestHasTeamsAccess(mockSimulatedHint === 'teamsAccessPending' ? false : undefined);
      setLoading(false);
      return;
    }

    // Only fetch sponsors when in view mode, the user is a guest, and a data source is ready.
    if (!isGuest) return;

    // All Graph calls go through the Azure Function proxy.
    if (functionUrl === undefined || aadHttpClient === undefined) return;

    let cancelled = false;
    setError(undefined);
    setIsPermissionError(false);

    const applyCachedSponsors = (cachedSponsors: ICachedSponsorsPayload): void => {
      sponsorCacheDisabledRef.current = false;
      sponsorCacheMetaRef.current = {
        activeSponsors: cachedSponsors.activeSponsors,
        clientVersion,
        unavailableSponsors: cachedSponsors.unavailableSponsors,
        sponsorOrder: cachedSponsors.sponsorOrder,
        guestHasTeamsAccess: cachedSponsors.guestHasTeamsAccess,
        presenceToken: cachedSponsors.presenceToken,
        functionVersion: cachedSponsors.functionVersion,
      };
      sponsorIdsRef.current = cachedSponsors.activeSponsors.map(s => s.id);
      setSponsors(cachedSponsors.activeSponsors);
      setSponsorOrder(cachedSponsors.sponsorOrder);
      setUnavailableSponsors(cachedSponsors.unavailableSponsors);
      setGuestHasTeamsAccess(cachedSponsors.guestHasTeamsAccess);
      setPresenceToken(cachedSponsors.presenceToken);
      setVersionMismatch(false);
      setLoading(false);
      setRetryCount(0);
      primePresenceRefreshRef.current = true;

      if (photoUrl && cachedSponsors.activeSponsors.some(s => s.managerId && !s.managerPhotoUrl)) {
        loadManagerPhotosViaProxy(photoUrl, aadHttpClient, cachedSponsors.presenceToken, cachedSponsors.activeSponsors, (sponsorId, managerPhotoUrl) => {
          if (!cancelled) {
            setSponsors(prev => {
              const nextSponsors = prev.map(s => (
                s.id === sponsorId ? { ...s, managerPhotoUrl } : s
              ));
              persistSponsorCache(nextSponsors);
              return nextSponsors;
            });
          }
        });
      }
    };

    const cachedSponsors = sponsorCacheKey ? readSponsorCache(sponsorCacheKey, sponsorCacheTtlMs, clientVersion) : undefined;
    if (cachedSponsors) {
      applyCachedSponsors(cachedSponsors);
      return () => { cancelled = true; };
    }

    sponsorCacheDisabledRef.current = false;
    setLoading(true);
    setVersionMismatch(false);
    setGuestHasTeamsAccess(undefined);

    const loadSponsors = async (): Promise<void> => {
      const peerCachedSponsors = sponsorCacheKey ? await requestSponsorCacheFromPeers() : undefined;
      if (cancelled) return;

      if (peerCachedSponsors) {
        persistSponsorCache(peerCachedSponsors.activeSponsors, {
          unavailableSponsors: peerCachedSponsors.unavailableSponsors,
          sponsorOrder: peerCachedSponsors.sponsorOrder,
          guestHasTeamsAccess: peerCachedSponsors.guestHasTeamsAccess,
          presenceToken: peerCachedSponsors.presenceToken,
          functionVersion: peerCachedSponsors.functionVersion,
        });
        applyCachedSponsors(peerCachedSponsors);
        return;
      }

      try {
        const result = await getSponsorsViaProxy(functionUrl, aadHttpClient, clientVersion, sponsorFilter, requireUserMailbox);
        if (cancelled) return;

        // Sponsor photos are bundled in the proxy response; keep them as-is.
        const active = result.activeSponsors;
        const nextSponsorOrder = result.sponsorOrder ?? active.map(s => s.id);
        const nextUnavailableSponsors = result.unavailableSponsors ?? [];
        setSponsors(active);
        sponsorIdsRef.current = active.map(s => s.id);
        setSponsorOrder(nextSponsorOrder);
        // All unavailable = sponsors were assigned but every account is disabled/deleted.
        setUnavailableSponsors(nextUnavailableSponsors);
        setGuestHasTeamsAccess(result.guestHasTeamsAccess);
        setPresenceToken(result.presenceToken);
        setLoading(false);
        setRetryCount(0);
        persistSponsorCache(active, {
          unavailableSponsors: nextUnavailableSponsors,
          sponsorOrder: nextSponsorOrder,
          guestHasTeamsAccess: result.guestHasTeamsAccess,
          presenceToken: result.presenceToken,
          functionVersion: result.functionVersion,
        });

        // Log a UI notice when the web part and function versions diverge.
        if (clientVersion && result.functionVersion && clientVersion !== result.functionVersion) {
          setVersionMismatch(true);
        }

        // Phase 2: lazily fetch manager photos via the proxy photo endpoint.
        // Sponsor photos already arrived inline; only manager photos need a round-trip.
        if (photoUrl && active.length > 0) {
          loadManagerPhotosViaProxy(photoUrl, aadHttpClient, result.presenceToken, active, (sponsorId, managerPhotoUrl) => {
            if (!cancelled) {
              setSponsors(prev => {
                const nextSponsors = prev.map(s => (
                  s.id === sponsorId ? { ...s, managerPhotoUrl } : s
                ));
                persistSponsorCache(nextSponsors);
                return nextSponsors;
              });
            }
          });
        }
      } catch (err: unknown) {
        if (cancelled) return;

        const status = (err as { statusCode?: number }).statusCode;
        const reasonCode = (err as { reasonCode?: string }).reasonCode;
        const referenceId = (err as { referenceId?: string }).referenceId;
        const retryable = (err as { retryable?: boolean }).retryable;
        // Structured console log for first-level operations triage.
        console.error('[GuestSponsorInfo] getSponsorsViaProxy failed', {
          status,
          reasonCode,
          referenceId,
          retryable,
          error: err,
        });

        const is4xx = status !== undefined && status >= 400 && status < 500;
        const shouldRetry = retryable === true || (!is4xx && retryable !== false);
        if (!shouldRetry || retryCount >= MAX_RETRIES) {
          // Permanent error (e.g. 401, 403) or retry limit reached:
          // stop retrying and show the error message so the shimmer disappears.
          if (reasonCode === 'GRAPH_PERMISSION_DENIED') {
            setIsPermissionError(true);
            setError(strings.InsufficientPermissionsMessage);
          } else {
            const supportRef = referenceId ? ` (Ref: ${referenceId})` : '';
            setError(`${fstr('ErrorMessage')}${supportRef}`);
          }
          setLoading(false);
        } else {
          // Transient error: retry with exponential backoff capped at 30 seconds.
          const delay = Math.min(3000 * Math.pow(3, retryCount), 30 * 1000);
          setTimeout(() => { if (!cancelled) setRetryCount(n => n + 1); }, delay);
        }
      }
    };

    loadSponsors().catch((unexpectedError: unknown) => {
      console.error('[GuestSponsorInfo] unexpected sponsor load failure', unexpectedError);
    });

    return () => { cancelled = true; };
  }, [isGuest, isEditMode, mockMode, mockSponsorCount, functionUrl, aadHttpClient, clientVersion, sponsorFilter, requireUserMailbox, retryCount, sponsorCacheTtlMs]);

  // Presence refresh: poll faster while a card is actively open and the tab is visible,
  // but back off when the tab is hidden to reduce Graph traffic.
  //
  // When presenceUrl is configured (Azure Function proxy) the poll calls the
  // /api/getPresence endpoint using application permissions (Managed Identity),
  // which works reliably for guest callers.  Otherwise falls back to the delegated
  // Graph call, which may silently return empty results for guests on tenants with
  // restrictive guest-access policies.
  const PRESENCE_REFRESH_ACTIVE_MS = 30 * 1000;
  const PRESENCE_REFRESH_VISIBLE_MS = 2 * 60 * 1000;
  const PRESENCE_REFRESH_HIDDEN_MS = 5 * 60 * 1000;
  React.useEffect(() => {
    if (isEditMode || mockMode || !isGuest || loading || error) return;
    if (!presenceUrl || !aadHttpClient) return;

    const refreshPresence = (): void => {
      const ids = sponsorIdsRef.current;
      if (ids.length === 0) return;
      getPresencesViaProxy(presenceUrl, aadHttpClient, ids, presenceToken)
        .then(({ map: presenceMap, presenceToken: renewedToken }) => {
          if (renewedToken) setPresenceToken(renewedToken);
          setSponsors(prev => {
            const nextSponsors = prev.map(s => {
              const snapshot = presenceMap.get(s.id);
              return {
                ...s,
                presence: snapshot?.availability ?? s.presence,
                presenceActivity: snapshot?.activity ?? s.presenceActivity,
              };
            });
            if (renewedToken) {
              persistSponsorCache(nextSponsors, { presenceToken: renewedToken });
            }
            return nextSponsors;
          });
        })
        .catch((err: unknown) => {
          // Presence refresh failures are silent — the existing data stays on screen.
          console.warn('[GuestSponsorInfo] Presence refresh failed:', err);
        });
    };

    const refreshImmediatelyIfNeeded = (): void => {
      if (!hasActiveCard && !primePresenceRefreshRef.current) return;
      primePresenceRefreshRef.current = false;
      refreshPresence();
    };

    const getRefreshIntervalMs = (): number => {
      if (document.visibilityState !== 'visible') return PRESENCE_REFRESH_HIDDEN_MS;
      return hasActiveCard ? PRESENCE_REFRESH_ACTIVE_MS : PRESENCE_REFRESH_VISIBLE_MS;
    };

    let id = setInterval(refreshPresence, getRefreshIntervalMs());

    const restartInterval = (): void => {
      clearInterval(id);
      id = setInterval(refreshPresence, getRefreshIntervalMs());
    };

    const onVisibilityChange = (): void => {
      restartInterval();
      if (document.visibilityState === 'visible') {
        refreshImmediatelyIfNeeded();
      }
    };

    document.addEventListener('visibilitychange', onVisibilityChange);

    if (document.visibilityState === 'visible') {
      refreshImmediatelyIfNeeded();
    }

    return () => {
      clearInterval(id);
      document.removeEventListener('visibilitychange', onVisibilityChange);
    };
  }, [isEditMode, mockMode, isGuest, loading, error, hasActiveCard, presenceUrl, aadHttpClient, presenceToken]);

  // Guard: SPFx AMD locale bundles load asynchronously. On the very first synchronous
  // render (edit mode renders immediately, before AMD resolves), strings may still be
  // undefined. Return null here — after all hooks have been called unconditionally —
  // so React can re-render once any state update or framework re-render provides a
  // populated strings object.
  if (!(strings as unknown as object | undefined)) return null;

  // Derive a Fluent v9 theme from the SPFx host site theme supplied by the ThemeProvider
  // service. Falls back to webLightTheme when the host theme is not yet available
  // (e.g. first render in the workbench before ThemeProvider initialises, or tests).
  // Without an explicit fallback, FluentProvider sets no token variables — Avatar
  // colorful backgrounds disappear and PresenceBadge dots render as black.
  // The second argument selects the correct v9 base theme so that v9 tokens with
  // no direct v8 mapping (e.g. dark-mode semantic colours) default to the right
  // values instead of always falling back to webLightTheme internally.
  const v9Theme = theme
    ? createV9Theme(
        theme as unknown as Parameters<typeof createV9Theme>[0],
        theme.isInverted ? webDarkTheme : webLightTheme
      )
    : webLightTheme;

  // Edit mode: always show a live preview using mock sponsor cards so page authors
  // can see the real layout and adjust display settings before going live.
  if (isEditMode) {
    // First-run wizard takes priority. Render it inline, full-width, instead of
    // the sponsor preview — no portals, no z-index conflict with SP chrome.
    // The property pane opens automatically once the admin finishes the wizard.
    if (showWelcomeDialog) {
      return (
        <RendererProvider renderer={griffelRenderer}>
        <FluentProvider theme={v9Theme} id={`${fluentProviderId}-edit`}>
          <WelcomeDialog open onCommit={handleWelcomeCommit} onSkip={handleWelcomeSkip} onDismiss={handleWelcomeDismiss} semver={clientVersion?.split('.').slice(0, 3).join('.')} isDark={theme?.isInverted} />
        </FluentProvider>
        </RendererProvider>
      );
    }

    const visibleMockSponsors = mockSponsors.slice(0, maxSponsorCount);
    const mockCompact = cardLayout === 'compact' || (cardLayout === 'auto' && visibleMockSponsors.length >= cardLayoutAutoThreshold);
    // Hide the sponsor list only when simulating "No sponsors found" — in that
    // state no sponsors are assigned at all so no tiles would appear. For
    // "Sponsor not available" the tiles ARE shown (read-only) so the editor can
    // see how the layout looks when all sponsors are unavailable.
    const showMockCards = mockSimulatedHint !== 'noSponsors';
    // When simulating "no sponsors" and the notice toggle is off, nothing remains
    // beyond the title — hide the entire web part so the editor sees the same
    // empty result the guest would see. The simulation hint is always respected
    // regardless of whether mockMode is active.
    // Using an IIFE avoids TypeScript control-flow narrowing through the || chain.
    const hasEditContent = ((): boolean => {
      if (mockSimulatedHint !== 'noSponsors') return true;     // shows tiles or a banner
      return showNoSponsorsHint;                               // only the notice banner remains
    })();
    if (!hasEditContent) return null;
    return (
      <RendererProvider renderer={griffelRenderer}>
      <FluentProvider theme={v9Theme} id={`${fluentProviderId}-edit`}>
        <section className={classes.webPart}>
          {(showTitle ?? true) && (title || isEditMode) ? (
            <h2
              className={mergeClasses(
                classes.title,
                titleSize === 'h2' && classes.titleH2,
                titleSize === 'h3' && classes.titleH3,
                titleSize === 'h4' && classes.titleH4,
                titleSize === 'normal' && classes.titleNormal,
                isEditMode && classes.titleEditable,
              )}
              contentEditable={isEditMode ? 'true' : undefined}
              suppressContentEditableWarning={isEditMode || undefined}
              data-placeholder={strings.TitlePlaceholder}
              onInput={isEditMode ? (e) => {
                // Browsers insert a <br> when the last character is deleted,
                // which prevents :empty from matching and hides the placeholder.
                // Clearing innerHTML when the visible text is gone fixes this.
                if ((e.currentTarget.textContent ?? '').trim() === '') {
                  e.currentTarget.innerHTML = '';
                }
              } : undefined}
              onBlur={isEditMode ? (e) => onTitleChange?.(e.currentTarget.textContent?.trim() ?? '') : undefined}
            >
              {/* Render undefined (not empty string) when title is blank so :empty CSS matches */}
              {title || undefined}
            </h2>
          ) : null}
          {showMockCards && (
          <SponsorList
            sponsors={visibleMockSponsors}
            hostTenantId={hostTenantId}
            compact={mockCompact}
            showBusinessPhones={showBusinessPhones}
            showMobilePhone={showMobilePhone}
            showWorkLocation={showWorkLocation}
            showCity={showCity}
            showCountry={showCountry}
            showStreetAddress={showStreetAddress}
            showPostalCode={showPostalCode}
            showState={showState}
            azureMapsSubscriptionKey={azureMapsSubscriptionKey}
            externalMapProvider={externalMapProvider}
            showManager={showManager}
            showPresence={showPresence}
            showSponsorJobTitle={showSponsorJobTitle}
            showManagerJobTitle={showManagerJobTitle}
            showSponsorDepartment={showSponsorDepartment}
            showManagerDepartment={showManagerDepartment}
            showSponsorPhoto={showSponsorPhoto}
            showManagerPhoto={showManagerPhoto}
            useInformalAddress={useInformalAddress}
            onActiveCardChange={() => undefined}
            guestHasTeamsAccess={mockSimulatedHint === 'teamsAccessPending' ? false : undefined}
            readOnlyIds={mockSimulatedHint === 'sponsorUnavailable' ? new Set(visibleMockSponsors.map(s => s.id)) : undefined}
            v9Theme={v9Theme}
          />
          )}
          {mockSimulatedHint === 'teamsAccessPending' && (
            <MessageBar intent="warning" className={classes.teamsAccessBanner}>
              <MessageBarBody>
                <b>{strings.TeamsAccessPendingTitle}</b><br />
                {fstr('TeamsAccessPendingMessage')}
              </MessageBarBody>
            </MessageBar>
          )}
          {mockSimulatedHint === 'versionMismatch' && (
            <MessageBar intent="info" className={classes.teamsAccessBanner}>
              <MessageBarBody>
                <b>{strings.VersionMismatchTitle}</b><br />
                {strings.VersionMismatchMessage}
              </MessageBarBody>
            </MessageBar>
          )}
          {mockSimulatedHint === 'sponsorUnavailable' && showSponsorUnavailableHint && (
            <MessageBar intent="warning" className={classes.teamsAccessBanner}>
              <MessageBarBody>
                <b>{strings.SponsorUnavailableTitle}</b><br />
                {fstr('SponsorUnavailableMessage')}
              </MessageBarBody>
            </MessageBar>
          )}
          {mockSimulatedHint === 'noSponsors' && showNoSponsorsHint && (
            <MessageBar intent="info">
              <MessageBarBody>
                <b>{strings.NoSponsorsTitle}</b><br />
                {fstr('NoSponsorsMessage')}
              </MessageBarBody>
            </MessageBar>
          )}
        </section>
      </FluentProvider>
      </RendererProvider>
    );
  }

  // View mode: render nothing for non-guest visitors (unless demo mode is active).
  if (!isGuest && !mockMode) {
    return null;
  }

  // View mode + guest user: render the sponsor list.
  // When some sponsors are unavailable we still show their tiles (read-only, no popup)
  // so the guest can see who their sponsors are. Active accounts "nachrücken" to fill
  // the visible slots: only active accounts count toward the maxSponsorCount cap.
  const { visibleActive, visibleUnavailable } = buildVisibleSponsorSets(
    sponsors, unavailableSponsors, sponsorOrder, maxSponsorCount
  );
  const someUnavailable = visibleUnavailable.length > 0;
  const noActiveSponsor = visibleActive.length === 0;

  // When nothing meaningful remains to show beyond the title (no tiles, no banners),
  // hide the entire web part so the guest never sees a lone heading.
  const hasVisibleContent =
    loading ||
    !!error ||
    visibleActive.length > 0 ||
    someUnavailable ||
    (noActiveSponsor && someUnavailable && showSponsorUnavailableHint) ||
    (noActiveSponsor && !someUnavailable && showNoSponsorsHint) ||
    (guestHasTeamsAccess === false && showTeamsAccessPendingHint) ||
    (versionMismatch && showVersionMismatchHint);
  if (!hasVisibleContent) return null;
  const noResults = !loading && !error && visibleActive.length === 0 && !someUnavailable;
  const contentClassNames = (loading || error || noResults) ? mergeClasses(classes.webPart, classes.webPartContent) : classes.webPart;
  return (
    <RendererProvider renderer={griffelRenderer}>
    <FluentProvider theme={v9Theme} id={`${fluentProviderId}-view`}>
      <section className={contentClassNames}>
        {(showTitle ?? true) && title && (
          <h2 className={mergeClasses(
            classes.title,
            titleSize === 'h2' && classes.titleH2,
            titleSize === 'h3' && classes.titleH3,
            titleSize === 'h4' && classes.titleH4,
            titleSize === 'normal' && classes.titleNormal,
          )}>
            {title}
          </h2>
        )}
        {loading && <SponsorGridSkeleton compact={cardLayout === 'compact'} />}
        {!loading && error && !isPermissionError && (
          <MessageBar intent="error">
            <MessageBarBody>
              <b>{strings.ErrorMessageTitle}</b><br />
              {error}
            </MessageBarBody>
          </MessageBar>
        )}
        {!loading && isPermissionError && error && (
          <MessageBar intent="error">
            <MessageBarBody>
              <b>{strings.InsufficientPermissionsTitle}</b><br />
              {error}
            </MessageBarBody>
          </MessageBar>
        )}
        {!loading && !error && (visibleActive.length > 0 || someUnavailable) && (
          <SponsorList
            sponsors={[...visibleActive, ...visibleUnavailable]}
            hostTenantId={hostTenantId}
            compact={cardLayout === 'compact' || (cardLayout === 'auto' && (visibleActive.length + visibleUnavailable.length) >= cardLayoutAutoThreshold)}
            showBusinessPhones={showBusinessPhones}
            showMobilePhone={showMobilePhone}
            showWorkLocation={showWorkLocation}
            showCity={showCity}
            showCountry={showCountry}
            showStreetAddress={showStreetAddress}
            showPostalCode={showPostalCode}
            showState={showState}
            azureMapsSubscriptionKey={azureMapsSubscriptionKey}
            externalMapProvider={externalMapProvider}
            showManager={showManager}
            showPresence={showPresence}
            showSponsorJobTitle={showSponsorJobTitle}
            showManagerJobTitle={showManagerJobTitle}
            showSponsorDepartment={showSponsorDepartment}
            showManagerDepartment={showManagerDepartment}
            showSponsorPhoto={showSponsorPhoto}
            showManagerPhoto={showManagerPhoto}
            useInformalAddress={useInformalAddress}
            onActiveCardChange={setHasActiveCard}
            guestHasTeamsAccess={guestHasTeamsAccess}
            readOnlyIds={someUnavailable ? new Set(visibleUnavailable.map(s => s.id)) : undefined}
            v9Theme={v9Theme}
          />
        )}
        {/* "Sponsor not available" notice — rendered below the tiles (if any).
            Only shown when no active sponsor is visible in the current set. */}
        {!loading && !error && noActiveSponsor && someUnavailable && showSponsorUnavailableHint && (
          <MessageBar
            intent="warning"
            className={classes.teamsAccessBanner}
          >
            <MessageBarBody>
              <b>{strings.SponsorUnavailableTitle}</b><br />
              {fstr('SponsorUnavailableMessage')}
            </MessageBarBody>
          </MessageBar>
        )}
        {!loading && !error && noActiveSponsor && !someUnavailable && showNoSponsorsHint && (
          <MessageBar intent="info">
            <MessageBarBody>
              <b>{strings.NoSponsorsTitle}</b><br />
              {fstr('NoSponsorsMessage')}
            </MessageBarBody>
          </MessageBar>
        )}
        {!loading && !error && guestHasTeamsAccess === false && showTeamsAccessPendingHint && (
          <MessageBar intent="warning" className={classes.teamsAccessBanner}>
            <MessageBarBody>
              <b>{strings.TeamsAccessPendingTitle}</b><br />
              {fstr('TeamsAccessPendingMessage')}
            </MessageBarBody>
          </MessageBar>
        )}
        {versionMismatch && showVersionMismatchHint && (
          <MessageBar intent="info" className={classes.teamsAccessBanner}>
            <MessageBarBody>
              <b>{strings.VersionMismatchTitle}</b><br />
              {strings.VersionMismatchMessage}
            </MessageBarBody>
          </MessageBar>
        )}
      </section>
    </FluentProvider>
    </RendererProvider>
  );
};

export default GuestSponsorInfo;


