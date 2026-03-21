import * as React from 'react';
import { DisplayMode } from '@microsoft/sp-core-library';
import { Shimmer, ShimmerElementType, ShimmerElementsGroup, MessageBar, MessageBarType } from '@fluentui/react';
import * as strings from 'GuestSponsorInfoWebPartStrings';
import styles from './GuestSponsorInfo.module.scss';
import type { IGuestSponsorInfoProps } from './IGuestSponsorInfoProps';
import { ISponsor } from '../services/ISponsor';
import { isGuestUser, getSponsors, getSponsorsViaProxy, loadPhotosProgressively, fetchPresences } from '../services/SponsorService';
import { MOCK_SPONSORS } from '../services/MockSponsorService';
import SponsorCard from './SponsorCard';

/**
 * Renders the sponsor grid and owns the single shared "active card" state.
 * Only one popup is ever visible at a time; switching cards cancels any
 * pending hide-timeout from the previously active card so there is no
 * overlap between the outgoing and incoming popup.
 */
interface ISponsorListProps {
  sponsors: ISponsor[];
  hostTenantId: string;
  showBusinessPhones: boolean;
  showMobilePhone: boolean;
  showWorkLocation: boolean;
  showManager: boolean;
  useInformalAddress: boolean;
  onActiveCardChange?: (hasActiveCard: boolean) => void;
  /** Propagated from ISponsorsResult — false shows disabled buttons in each card. */
  guestHasTeamsAccess?: boolean;
}

const SponsorList: React.FC<ISponsorListProps> = ({ sponsors, hostTenantId, showBusinessPhones, showMobilePhone, showWorkLocation, showManager, useInformalAddress, onActiveCardChange, guestHasTeamsAccess }) => {
  const [activeId, setActiveId] = React.useState<string | null>(null);
  const hideTimeout = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  const activate = (id: string): void => {
    if (hideTimeout.current) { clearTimeout(hideTimeout.current); hideTimeout.current = null; }
    setActiveId(id);
    onActiveCardChange?.(true);
  };
  const scheduleDeactivate = (): void => {
    hideTimeout.current = setTimeout(() => {
      setActiveId(null);
      onActiveCardChange?.(false);
    }, 150);
  };

  React.useEffect(() => {
    return () => {
      if (hideTimeout.current) clearTimeout(hideTimeout.current);
      onActiveCardChange?.(false);
    };
  }, [onActiveCardChange]);

  return (
    <ul className={styles.sponsorGrid}>
      {sponsors.map(sponsor => (
        <li key={sponsor.id} className={styles.sponsorItem}>
          <SponsorCard
            sponsor={sponsor}
            hostTenantId={hostTenantId}
            isActive={activeId === sponsor.id}
            onActivate={() => activate(sponsor.id)}
            onScheduleDeactivate={scheduleDeactivate}
            showBusinessPhones={showBusinessPhones}
            showMobilePhone={showMobilePhone}
            showWorkLocation={showWorkLocation}
            showManager={showManager}
            useInformalAddress={useInformalAddress}
            guestHasTeamsAccess={guestHasTeamsAccess}
          />
        </li>
      ))}
    </ul>
  );
};

// Skeleton shimmer for a single sponsor card (136 × 138 px — matches .card exactly).
// Pixel values: card width 136px, avatar 72px centered, name 90px, job title 68px.
// Defined once outside the component so it is not recreated on every render.
const sponsorCardShimmer = (
  <ShimmerElementsGroup
    flexWrap
    width="136px"
    shimmerElements={[
      // top padding 12px
      { type: ShimmerElementType.gap, width: '100%', height: 12 },
      // avatar row: (136-72)/2 = 32px gap each side
      { type: ShimmerElementType.gap,    width: 32, height: 72 },
      { type: ShimmerElementType.circle,            height: 72 },
      { type: ShimmerElementType.gap,    width: 32, height: 72 },
      // gap between avatar and name
      { type: ShimmerElementType.gap, width: '100%', height: 8 },
      // name line: 1 line (18px) — matches the single-line truncation used when a job title is present
      { type: ShimmerElementType.gap,  width: 23, height: 18 },
      { type: ShimmerElementType.line, width: 90, height: 18 },
      { type: ShimmerElementType.gap,  width: 23, height: 18 },
      // gap between name and job title
      { type: ShimmerElementType.gap, width: '100%', height: 8 },
      // job title line (~68px centered): (136-68)/2 = 34px each side; height = 12px font × 1.3 line-height ≈ 16px
      { type: ShimmerElementType.gap,  width: 34, height: 16 },
      { type: ShimmerElementType.line, width: 68, height: 16 },
      { type: ShimmerElementType.gap,  width: 34, height: 16 },
      // bottom padding 12px
      { type: ShimmerElementType.gap, width: '100%', height: 12 },
    ]}
  />
);

/**
 * Renders a single shimmer placeholder card in the same grid as the real sponsor list.
 * Shown instead of a loading text while Graph data is being fetched.
 * Displaying only one card reduces layout shift on page load while still providing
 * visual feedback that data is being fetched.
 */
const SponsorGridSkeleton: React.FC = () => (
  <ul className={styles.sponsorGrid} aria-busy="true">
    <li className={styles.sponsorItem}>
      <Shimmer customElementsGroup={sponsorCardShimmer} width="136px" />
    </li>
  </ul>
);

/** Maximum number of transient-error retries before giving up and showing an error. */
const MAX_RETRIES = 3;

type ProxyStatus = 'checking' | 'ok' | 'error';

const GuestSponsorInfo: React.FC<IGuestSponsorInfoProps> = ({
  loginName,
  isExternalGuestUser,
  displayMode,
  graphClient,
  title,
  mockMode,
  hostTenantId,
  functionUrl,
  aadHttpClient,
  showBusinessPhones,
  showMobilePhone,
  showWorkLocation,
  showManager,
  useInformalAddress,
}) => {
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

  const [sponsors, setSponsors] = React.useState<ISponsor[]>([]);
  const [allUnavailable, setAllUnavailable] = React.useState(false);
  // Start in loading state immediately for guests and demo mode so the shimmer
  // is visible on the very first render — before the first useEffect tick.
  // Without this, React paints a brief "no sponsors" flash before the effect
  // can call setLoading(true).
  const [loading, setLoading] = React.useState(!isEditMode && (mockMode || isGuest));
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [proxyStatus, setProxyStatus] = React.useState<ProxyStatus>('checking');
  const [retryCount, setRetryCount] = React.useState(0);
  const [hasActiveCard, setHasActiveCard] = React.useState(false);
  const [guestHasTeamsAccess, setGuestHasTeamsAccess] = React.useState<boolean | undefined>(undefined);

  // Ref that always holds the IDs of currently displayed sponsors.
  // The presence refresh interval reads this without capturing sponsors in its closure.
  const sponsorIdsRef = React.useRef<string[]>([]);

  // Edit-mode proxy health check: verify the Azure Function is reachable while the
  // page author has the web part selected. Only fires when functionUrl is configured.
  React.useEffect(() => {
    if (!isEditMode || !functionUrl) return;
    if (!aadHttpClient) { setProxyStatus('error'); return; }
    let cancelled = false;
    setProxyStatus('checking');
    getSponsorsViaProxy(functionUrl, aadHttpClient)
      .then(() => { if (!cancelled) setProxyStatus('ok'); })
      .catch(() => { if (!cancelled) setProxyStatus('error'); });
    return () => { cancelled = true; };
  }, [isEditMode, functionUrl, aadHttpClient]);

  React.useEffect(() => {
    if (isEditMode) return;

    // Demo mode: use static mock data, no Graph calls needed.
    if (mockMode) {
      setSponsors(MOCK_SPONSORS);
      setAllUnavailable(false);
      setLoading(false);
      return;
    }

    // Only fetch sponsors when in view mode, the user is a guest, and a data source is ready.
    if (!isGuest) return;

    // Prefer the function proxy when configured; fall back to direct Graph.
    const useProxy = functionUrl !== undefined && aadHttpClient !== undefined;
    if (!useProxy && graphClient === undefined) return;

    let cancelled = false;
    setLoading(true);
    setError(undefined);
    setGuestHasTeamsAccess(undefined);
    const loadFn = useProxy
      ? () => getSponsorsViaProxy(functionUrl as string, aadHttpClient!)
      : () => getSponsors(graphClient!);

    loadFn()
      .then(result => {
        if (!cancelled) {
          // Photo rendering is always client-side from direct Graph photo endpoints.
          // Strip proxy-provided photo fields defensively to keep this contract explicit.
          const active = result.activeSponsors.map(({ photoUrl: _photoUrl, managerPhotoUrl: _managerPhotoUrl, ...s }) => s);
          setSponsors(active);
          sponsorIdsRef.current = active.map(s => s.id);
          // All unavailable = sponsors were assigned but every account is disabled/deleted.
          setAllUnavailable(active.length === 0 && result.unavailableCount > 0);
          setGuestHasTeamsAccess(result.guestHasTeamsAccess);
          setLoading(false);
          setRetryCount(0);

          // Phase 2: progressively fetch photos without blocking the initial render.
          // graphClient is always obtained in onInit regardless of proxy mode,
          // so it is safe to use here even on the proxy path.
          if (graphClient && active.length > 0) {
            loadPhotosProgressively(graphClient, active, (sponsorId, photoUrl, managerPhotoUrl) => {
              if (!cancelled) {
                setSponsors(prev => prev.map(s =>
                  s.id === sponsorId ? { ...s, photoUrl, managerPhotoUrl } : s
                ));
              }
            });
          }
        }
      })
      .catch((err: unknown) => {
        if (!cancelled) {
          // Log the raw Graph error to the browser console so admins/developers
          // can see the exact status code and error message without opening the
          // network tab (useful when the web part is embedded in a guest session).
          console.error('[GuestSponsorInfo] getSponsors failed:', err);
          const status = (err as { statusCode?: number }).statusCode;
          const is4xx = status !== undefined && status >= 400 && status < 500;
          if (is4xx || retryCount >= MAX_RETRIES) {
            // Permanent error (e.g. 401, 403) or retry limit reached:
            // stop retrying and show the error message so the shimmer disappears.
            setError(fstr('ErrorMessage'));
            setLoading(false);
          } else {
            // Transient error: retry with exponential backoff capped at 30 seconds.
            const delay = Math.min(3000 * Math.pow(3, retryCount), 30 * 1000);
            setTimeout(() => { if (!cancelled) setRetryCount(n => n + 1); }, delay);
          }
        }
      });

    return () => { cancelled = true; };
  }, [isGuest, isEditMode, graphClient, mockMode, functionUrl, aadHttpClient, retryCount]);

  // Presence refresh: poll faster while a card is actively open and the tab is visible,
  // but back off when the tab is hidden to reduce Graph traffic.
  // Uses graphClient directly for both proxy and direct paths — graphClient is
  // always acquired in onInit regardless of whether the proxy is configured.
  const PRESENCE_REFRESH_ACTIVE_MS = 30 * 1000;
  const PRESENCE_REFRESH_VISIBLE_MS = 2 * 60 * 1000;
  const PRESENCE_REFRESH_HIDDEN_MS = 5 * 60 * 1000;
  React.useEffect(() => {
    if (isEditMode || mockMode || !isGuest || loading || error) return;
    if (!graphClient) return;

    const refreshPresence = (): void => {
      const ids = sponsorIdsRef.current;
      if (ids.length === 0) return;
      fetchPresences(graphClient, ids)
        .then(presenceMap => {
          setSponsors(prev => prev.map(s => {
            const snapshot = presenceMap.get(s.id);
            return {
              ...s,
              presence: snapshot?.availability ?? s.presence,
              presenceActivity: snapshot?.activity ?? s.presenceActivity,
            };
          }));
        })
        .catch((err: unknown) => {
          // Presence refresh failures are silent — the existing data stays on screen.
          console.warn('[GuestSponsorInfo] Presence refresh failed:', err);
        });
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
        refreshPresence();
      }
    };

    document.addEventListener('visibilitychange', onVisibilityChange);

    if (document.visibilityState === 'visible' && hasActiveCard) {
      refreshPresence();
    }

    return () => {
      clearInterval(id);
      document.removeEventListener('visibilitychange', onVisibilityChange);
    };
  }, [isEditMode, mockMode, isGuest, loading, error, graphClient, hasActiveCard]);

  // Edit mode: always show a lightweight placeholder so page authors can position the web part.
  if (isEditMode) {
    let placeholderText: string;
    if (mockMode) {
      placeholderText = strings.MockModePlaceholder;
    } else if (isGuest) {
      placeholderText = strings.EditModePlaceholder;
    } else {
      placeholderText = strings.GuestOnlyPlaceholder;
    }
    const proxyStatusClass = proxyStatus === 'ok'
      ? styles.proxyStatusOk
      : proxyStatus === 'error'
        ? styles.proxyStatusError
        : styles.proxyStatusChecking;
    return (
      <section className={styles.webPart}>
        <div className={styles.editPlaceholder}>
          <span>{placeholderText}</span>
          {functionUrl && (
            <div className={`${styles.proxyStatus} ${proxyStatusClass}`}>
              <span className={styles.proxyStatusDot} aria-hidden="true" />
              <span>
                {proxyStatus === 'checking' ? strings.ProxyStatusChecking
                  : proxyStatus === 'ok' ? strings.ProxyStatusOk
                  : strings.ProxyStatusError}
              </span>
            </div>
          )}
        </div>
      </section>
    );
  }

  // View mode: render nothing for non-guest visitors (unless demo mode is active).
  if (!isGuest && !mockMode) {
    return null;
  }

  // View mode + guest user: render the sponsor list.
  const contentClassNames = (loading || error) ? `${styles.webPart} ${styles.webPartContent}` : styles.webPart;
  return (
    <section className={contentClassNames}>
      {title && <h2 className={styles.title}>{title}</h2>}
      {loading && <SponsorGridSkeleton />}
      {!loading && error && (
        <p className={styles.statusMessage}>{error}</p>
      )}
      {!loading && !error && sponsors.length === 0 && allUnavailable && (
        <p className={styles.statusMessage}>{fstr('SponsorUnavailableMessage')}</p>
      )}
      {!loading && !error && sponsors.length === 0 && !allUnavailable && (
        <p className={styles.statusMessage}>{fstr('NoSponsorsMessage')}</p>
      )}
      {!loading && !error && sponsors.length > 0 && (
        <SponsorList
          sponsors={sponsors}
          hostTenantId={hostTenantId}
          showBusinessPhones={showBusinessPhones}
          showMobilePhone={showMobilePhone}
          showWorkLocation={showWorkLocation}
          showManager={showManager}
          useInformalAddress={useInformalAddress}
          onActiveCardChange={setHasActiveCard}
          guestHasTeamsAccess={guestHasTeamsAccess}
        />
      )}
      {!loading && !error && guestHasTeamsAccess === false && (
        <MessageBar
          messageBarType={MessageBarType.warning}
          isMultiline
          className={styles.teamsAccessBanner}
        >
          {fstr('TeamsAccessPendingMessage')}
        </MessageBar>
      )}
    </section>
  );
};

export default GuestSponsorInfo;

