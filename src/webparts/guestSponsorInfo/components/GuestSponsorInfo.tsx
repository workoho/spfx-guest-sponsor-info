import * as React from 'react';
import { DisplayMode } from '@microsoft/sp-core-library';
import { MessageBar, MessageBarType } from '@fluentui/react';
import * as strings from 'GuestSponsorInfoWebPartStrings';
import styles from './GuestSponsorInfo.module.scss';
import type { IGuestSponsorInfoProps } from './IGuestSponsorInfoProps';
import { ISponsor } from '../services/ISponsor';
import { isGuestUser, getSponsors, getSponsorsViaProxy, pingProxy, loadPhotosProgressively, fetchPresences, getPresencesViaProxy } from '../services/SponsorService';
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
  externalMapProvider: 'bing' | 'google' | 'apple' | 'openstreetmap' | 'here' | 'none';
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
}

const SponsorList: React.FC<ISponsorListProps> = ({ sponsors, hostTenantId, compact, showBusinessPhones, showMobilePhone, showWorkLocation, showCity, showCountry, showStreetAddress, showPostalCode, showState, azureMapsSubscriptionKey, externalMapProvider, showManager, showPresence, showSponsorJobTitle, showManagerJobTitle, showSponsorDepartment, showManagerDepartment, showSponsorPhoto, showManagerPhoto, useInformalAddress, onActiveCardChange, guestHasTeamsAccess }) => {
  const [activeId, setActiveId] = React.useState<string | null>(null);
  const showTimeout = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  const hideTimeout = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  const activate = (id: string): void => {
    if (hideTimeout.current) { clearTimeout(hideTimeout.current); hideTimeout.current = null; }
    if (showTimeout.current) return; // already pending for this or another card
    showTimeout.current = setTimeout(() => {
      showTimeout.current = null;
      setActiveId(id);
      onActiveCardChange?.(true);
    }, 500);
  };
  const scheduleDeactivate = (): void => {
    if (showTimeout.current) { clearTimeout(showTimeout.current); showTimeout.current = null; }
    hideTimeout.current = setTimeout(() => {
      setActiveId(null);
      onActiveCardChange?.(false);
    }, 150);
  };

  React.useEffect(() => {
    return () => {
      if (showTimeout.current) clearTimeout(showTimeout.current);
      if (hideTimeout.current) clearTimeout(hideTimeout.current);
      onActiveCardChange?.(false);
    };
  }, [onActiveCardChange]);

  return (
    <ul className={compact ? styles.sponsorGridCompact : styles.sponsorGrid}>
      {sponsors.map(sponsor => (
        <li key={sponsor.id} className={styles.sponsorItem}>
          <SponsorCard
            sponsor={sponsor}
            hostTenantId={hostTenantId}
            compact={compact}
            isActive={activeId === sponsor.id}
            onActivate={() => activate(sponsor.id)}
            onScheduleDeactivate={scheduleDeactivate}
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
          />
        </li>
      ))}
    </ul>
  );
};

/**
 * Pure-CSS loading skeleton — uses the identical DOM structure and CSS classes
 * as the real sponsor cards so spacing is pixel-perfect and requires no manual
 * height arithmetic.  Two placeholder cards are rendered (matching the typical
 * mock-data count) to give a realistic sense of the list size.
 */
const SponsorGridSkeleton: React.FC<{ compact: boolean }> = ({ compact }) => (
  <ul className={compact ? styles.sponsorGridCompact : styles.sponsorGrid} aria-busy="true">
    {[0, 1].map(i => (
      <li key={i} className={styles.sponsorItem}>
        {compact ? (
          <div className={`${styles.cardCompact} ${styles.skeletonItem}`}>
            <div className={styles.skeletonCircleCompact} />
            <div className={styles.skeletonLine} style={{ width: 120 }} />
          </div>
        ) : (
          <div className={`${styles.card} ${styles.skeletonItem}`}>
            <div className={styles.skeletonCircle} />
            <div className={styles.skeletonLine} style={{ width: 88 }} />
            <div className={styles.skeletonLine} style={{ width: 64 }} />
          </div>
        )}
      </li>
    ))}
  </ul>
);

/** Maximum number of transient-error retries before giving up and showing an error. */
const MAX_RETRIES = 3;

const GuestSponsorInfo: React.FC<IGuestSponsorInfoProps> = ({
  loginName,
  isExternalGuestUser,
  displayMode,
  graphClient,
  title,
  mockMode,
  mockSimulatedHint,
  showTeamsAccessPendingHint,
  showVersionMismatchHint,
  showSponsorUnavailableHint,
  showNoSponsorsHint,
  cardLayout,
  hostTenantId,
  functionUrl,
  presenceUrl,
  pingUrl,
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
  clientVersion,
  onProxyStatusChange,
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
  const [isPermissionError, setIsPermissionError] = React.useState(false);
  const [retryCount, setRetryCount] = React.useState(0);
  const [hasActiveCard, setHasActiveCard] = React.useState(false);
  const [guestHasTeamsAccess, setGuestHasTeamsAccess] = React.useState<boolean | undefined>(undefined);
  const [versionMismatch, setVersionMismatch] = React.useState(false);
  // Signed token issued by getGuestSponsors; passed to getPresence so the function
  // can validate sponsor IDs without server-side state or extra Graph calls.
  const [presenceToken, setPresenceToken] = React.useState<string | undefined>(undefined);

  // Ref that always holds the IDs of currently displayed sponsors.
  // The presence refresh interval reads this without capturing sponsors in its closure.
  const sponsorIdsRef = React.useRef<string[]>([]);

  // Edit-mode proxy health check: verify the Azure Function is reachable while the
  // page author has the web part selected.  Uses the lightweight /api/ping endpoint
  // so the check works for any authenticated user — not just guests.
  React.useEffect(() => {
    if (!isEditMode || !pingUrl) return;
    if (!aadHttpClient) { onProxyStatusChange?.('error'); return; }
    let cancelled = false;
    onProxyStatusChange?.('checking');
    pingProxy(pingUrl, aadHttpClient)
      .then(() => { if (!cancelled) { onProxyStatusChange?.('ok'); } })
      .catch(() => { if (!cancelled) { onProxyStatusChange?.('error'); } });
    return () => { cancelled = true; };
  }, [isEditMode, pingUrl, aadHttpClient]);

  React.useEffect(() => {
    if (isEditMode) return;

    // Demo mode: use static mock data, no Graph calls needed.
    // Simulate a guest whose Teams access hasn't been provisioned yet so
    // the warning banner and disabled Chat/Call buttons are visible.
    if (mockMode) {
      setSponsors(MOCK_SPONSORS);
      setAllUnavailable(false);
      setGuestHasTeamsAccess(mockSimulatedHint === 'teamsAccessPending' ? false : undefined);
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
    setIsPermissionError(false);
    setVersionMismatch(false);
    setGuestHasTeamsAccess(undefined);
    const loadFn = useProxy
      ? () => getSponsorsViaProxy(functionUrl as string, aadHttpClient!, clientVersion)
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
          setPresenceToken(result.presenceToken);
          setLoading(false);
          setRetryCount(0);

          // Log a UI notice when the web part and function versions diverge.
          if (clientVersion && result.functionVersion && clientVersion !== result.functionVersion) {
            setVersionMismatch(true);
          }

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
          const status = (err as { statusCode?: number }).statusCode;
          const reasonCode = (err as { reasonCode?: string }).reasonCode;
          const referenceId = (err as { referenceId?: string }).referenceId;
          const retryable = (err as { retryable?: boolean }).retryable;
          // Structured console log for first-level operations triage.
          console.error('[GuestSponsorInfo] getSponsors failed', {
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
      });

    return () => { cancelled = true; };
  }, [isGuest, isEditMode, graphClient, mockMode, functionUrl, aadHttpClient, clientVersion, retryCount]);

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
    const useProxy = !!(presenceUrl && aadHttpClient);
    if (!useProxy && !graphClient) return;

    const refreshPresence = (): void => {
      const ids = sponsorIdsRef.current;
      if (ids.length === 0) return;
      const presencePromise = (presenceUrl && aadHttpClient)
        ? getPresencesViaProxy(presenceUrl, aadHttpClient, ids, presenceToken)
        : fetchPresences(graphClient!, ids).then(presenceMap => ({ map: presenceMap, presenceToken: undefined }));
      presencePromise
        .then(({ map: presenceMap, presenceToken: renewedToken }) => {
          if (renewedToken) setPresenceToken(renewedToken);
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
  }, [isEditMode, mockMode, isGuest, loading, error, graphClient, hasActiveCard, presenceUrl, aadHttpClient, presenceToken]);

  // Guard: SPFx AMD locale bundles load asynchronously. On the very first synchronous
  // render (edit mode renders immediately, before AMD resolves), strings may still be
  // undefined. Return null here — after all hooks have been called unconditionally —
  // so React can re-render once any state update or framework re-render provides a
  // populated strings object.
  if (!(strings as unknown as object | undefined)) return null;

  // Edit mode: always show a live preview using mock sponsor cards so page authors
  // can see the real layout and adjust display settings before going live.
  if (isEditMode) {
    const mockCompact = cardLayout === 'compact' || (cardLayout === 'auto' && MOCK_SPONSORS.length > 2);
    return (
      <section className={styles.webPart}>
        {title && <h2 className={styles.title}>{title}</h2>}
        <SponsorList
          sponsors={MOCK_SPONSORS}
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
          guestHasTeamsAccess={mockMode && mockSimulatedHint === 'teamsAccessPending' ? false : undefined}
        />
        {mockMode && mockSimulatedHint === 'teamsAccessPending' && (
          <MessageBar
            messageBarType={MessageBarType.warning}
            isMultiline
            delayedRender={false}
            className={styles.teamsAccessBanner}
          >
            <b>{strings.TeamsAccessPendingTitle}</b><br />
            {fstr('TeamsAccessPendingMessage')}
          </MessageBar>
        )}
        {mockMode && mockSimulatedHint === 'versionMismatch' && (
          <MessageBar
            messageBarType={MessageBarType.warning}
            isMultiline
            delayedRender={false}
            className={styles.teamsAccessBanner}
          >
            <b>{strings.VersionMismatchTitle}</b><br />
            {strings.VersionMismatchMessage}
          </MessageBar>
        )}
        {mockMode && mockSimulatedHint === 'sponsorUnavailable' && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline delayedRender={false}>
            <b>{strings.SponsorUnavailableTitle}</b><br />
            {fstr('SponsorUnavailableMessage')}
          </MessageBar>
        )}
        {mockMode && mockSimulatedHint === 'noSponsors' && (
          <MessageBar messageBarType={MessageBarType.info} isMultiline delayedRender={false}>
            <b>{strings.NoSponsorsTitle}</b><br />
            {fstr('NoSponsorsMessage')}
          </MessageBar>
        )}
      </section>
    );
  }

  // View mode: render nothing for non-guest visitors (unless demo mode is active).
  if (!isGuest && !mockMode) {
    return null;
  }

  // View mode + guest user: render the sponsor list.
  const noResults = !loading && !error && sponsors.length === 0;
  const contentClassNames = (loading || error || noResults) ? `${styles.webPart} ${styles.webPartContent}` : styles.webPart;
  return (
    <section className={contentClassNames}>
      {title && <h2 className={styles.title}>{title}</h2>}
      {loading && <SponsorGridSkeleton compact={cardLayout === 'compact'} />}
      {!loading && error && !isPermissionError && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline delayedRender={false}>
          <b>{strings.ErrorMessageTitle}</b><br />
          {error}
        </MessageBar>
      )}
      {!loading && isPermissionError && error && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline delayedRender={false}>
          <b>{strings.InsufficientPermissionsTitle}</b><br />
          {error}
        </MessageBar>
      )}
      {!loading && !error && sponsors.length === 0 && allUnavailable && showSponsorUnavailableHint && (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline delayedRender={false}>
          <b>{strings.SponsorUnavailableTitle}</b><br />
          {fstr('SponsorUnavailableMessage')}
        </MessageBar>
      )}
      {!loading && !error && sponsors.length === 0 && !allUnavailable && showNoSponsorsHint && (
        <MessageBar messageBarType={MessageBarType.info} isMultiline delayedRender={false}>
          <b>{strings.NoSponsorsTitle}</b><br />
          {fstr('NoSponsorsMessage')}
        </MessageBar>
      )}
      {!loading && !error && sponsors.length > 0 && (
        <SponsorList
          sponsors={sponsors}
          hostTenantId={hostTenantId}
          compact={cardLayout === 'compact' || (cardLayout === 'auto' && sponsors.length > 2)}
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
        />
      )}
      {!loading && !error && guestHasTeamsAccess === false && showTeamsAccessPendingHint && (
        <MessageBar
          messageBarType={MessageBarType.warning}
          isMultiline
          delayedRender={false}
          className={styles.teamsAccessBanner}
        >
          <b>{strings.TeamsAccessPendingTitle}</b><br />
          {fstr('TeamsAccessPendingMessage')}
        </MessageBar>
      )}
      {versionMismatch && showVersionMismatchHint && (
        <MessageBar
          messageBarType={MessageBarType.warning}
          isMultiline
          delayedRender={false}
          className={styles.teamsAccessBanner}
        >
          <b>{strings.VersionMismatchTitle}</b><br />
          {strings.VersionMismatchMessage}
        </MessageBar>
      )}
    </section>
  );
};

export default GuestSponsorInfo;

