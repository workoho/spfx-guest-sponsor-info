import * as React from 'react';
import { DisplayMode } from '@microsoft/sp-core-library';
import { FluentProvider, MessageBar, MessageBarBody, Skeleton, SkeletonItem, mergeClasses, webLightTheme, webDarkTheme } from '@fluentui/react-components';
import type { Theme } from '@fluentui/react-components';
import { createV9Theme } from '@fluentui/react-migration-v8-v9';
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
  /**
   * When true, sponsor tiles are shown as static visuals only — no hover popup,
   * no keyboard activation. Used when all sponsors are unavailable.
   */
  readOnly?: boolean;
  /** Fluent v9 theme to forward into SponsorCard portal FluentProviders. */
  v9Theme?: Theme;
}

const SponsorList: React.FC<ISponsorListProps> = ({ sponsors, hostTenantId, compact, showBusinessPhones, showMobilePhone, showWorkLocation, showCity, showCountry, showStreetAddress, showPostalCode, showState, azureMapsSubscriptionKey, externalMapProvider, showManager, showPresence, showSponsorJobTitle, showManagerJobTitle, showSponsorDepartment, showManagerDepartment, showSponsorPhoto, showManagerPhoto, useInformalAddress, onActiveCardChange, guestHasTeamsAccess, readOnly, v9Theme }) => {
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
            isActive={!readOnly && activeId === sponsor.id}
            onActivate={readOnly ? () => undefined : () => activate(sponsor.id)}
            onScheduleDeactivate={readOnly ? () => undefined : scheduleDeactivate}
            readOnly={readOnly}
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
      ))}
    </ul>
  );
};

/**
 * Loading skeleton using Fluent UI Skeleton + SkeletonItem components.
 * Reuses the real card CSS classes so spacing is pixel-perfect.
 * Fluent’s Skeleton handles the shimmer animation and automatically
 * respects prefers-reduced-motion without custom media-query code.
 */
const SponsorGridSkeleton: React.FC<{ compact: boolean }> = ({ compact }) => (
  <ul className={compact ? styles.sponsorGridCompact : styles.sponsorGrid} aria-busy="true">
    {[0, 1].map(i => (
      <li key={i} className={styles.sponsorItem}>
        {compact ? (
          <Skeleton
            className={mergeClasses(styles.cardCompact, styles.skeletonItem)}
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
            className={mergeClasses(styles.card, styles.skeletonItem)}
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

/**
 * Builds the two visible sponsor sets, respecting both the configured cap and
 * the original Entra ordering. Only active accounts count toward the cap;
 * unavailable accounts (disabled / deleted) are shown alongside the active
 * ones in their original position so the guest can always see who their
 * sponsors are, while active accounts "nachrücken" to fill the visible slots.
 *
 * Falls back to independent slice() when sponsorOrder is empty (e.g. older
 * Function versions or the direct Graph path without ordering info).
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
  graphClient,
  title,
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

  const [sponsors, setSponsors] = React.useState<ISponsor[]>([]);
  const [sponsorOrder, setSponsorOrder] = React.useState<string[]>([]);
  const [unavailableSponsors, setUnavailableSponsors] = React.useState<ISponsor[]>([]);
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
          setVersionMismatch(
            !!(clientVersion && functionVersion && clientVersion !== functionVersion)
          );
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
          setSponsorOrder(result.sponsorOrder ?? active.map(s => s.id));
          // All unavailable = sponsors were assigned but every account is disabled/deleted.
          setUnavailableSponsors(result.unavailableSponsors ?? []);
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
  }, [isGuest, isEditMode, graphClient, mockMode, mockSponsorCount, functionUrl, aadHttpClient, clientVersion, retryCount]);

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
      if (versionMismatch) return true;                        // real version-mismatch ping result
      return showNoSponsorsHint;                               // only the notice banner remains
    })();
    if (!hasEditContent) return null;
    return (
      <FluentProvider theme={v9Theme} id={`${fluentProviderId}-edit`}>
        <section className={styles.webPart}>
          {title && <h2 className={styles.title}>{title}</h2>}
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
            readOnly={mockSimulatedHint === 'sponsorUnavailable'}
            v9Theme={v9Theme}
          />
          )}
          {/* Real version mismatch detected via ping: always shown to the editor, independent
              of the showVersionMismatchHint guest-facing toggle. */}
          {versionMismatch && mockSimulatedHint !== 'versionMismatch' && (
            <MessageBar intent="warning" className={styles.teamsAccessBanner}>
              <MessageBarBody>
                <b>{strings.VersionMismatchTitle}</b><br />
                {strings.VersionMismatchMessage}
              </MessageBarBody>
            </MessageBar>
          )}
          {mockSimulatedHint === 'teamsAccessPending' && (
            <MessageBar intent="warning" className={styles.teamsAccessBanner}>
              <MessageBarBody>
                <b>{strings.TeamsAccessPendingTitle}</b><br />
                {fstr('TeamsAccessPendingMessage')}
              </MessageBarBody>
            </MessageBar>
          )}
          {mockSimulatedHint === 'versionMismatch' && (
            <MessageBar intent="warning" className={styles.teamsAccessBanner}>
              <MessageBarBody>
                <b>{strings.VersionMismatchTitle}</b><br />
                {strings.VersionMismatchMessage}
              </MessageBarBody>
            </MessageBar>
          )}
          {mockSimulatedHint === 'sponsorUnavailable' && showSponsorUnavailableHint && (
            <MessageBar intent="warning" className={styles.teamsAccessBanner}>
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
  const contentClassNames = (loading || error || noResults) ? `${styles.webPart} ${styles.webPartContent}` : styles.webPart;
  return (
    <FluentProvider theme={v9Theme} id={`${fluentProviderId}-view`}>
      <section className={contentClassNames}>
        {title && <h2 className={styles.title}>{title}</h2>}
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
        {!loading && !error && visibleActive.length > 0 && (
          <SponsorList
            sponsors={visibleActive}
            hostTenantId={hostTenantId}
            compact={cardLayout === 'compact' || (cardLayout === 'auto' && visibleActive.length >= cardLayoutAutoThreshold)}
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
            v9Theme={v9Theme}
          />
        )}
        {/* Unavailable sponsors: show tiles read-only (no hover popup) so the guest
            can still see who their sponsors are, even if the accounts are currently
            disabled or deleted. Shown whenever there are unavailable sponsors in the
            visible set — not only when all sponsors are unavailable. */}
        {!loading && !error && someUnavailable && (
          <SponsorList
            sponsors={visibleUnavailable}
            hostTenantId={hostTenantId}
            compact={cardLayout === 'compact' || (cardLayout === 'auto' && visibleUnavailable.length >= cardLayoutAutoThreshold)}
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
            readOnly
            v9Theme={v9Theme}
          />
        )}
        {/* "Sponsor not available" notice — rendered below the tiles (if any).
            Only shown when no active sponsor is visible in the current set. */}
        {!loading && !error && noActiveSponsor && someUnavailable && showSponsorUnavailableHint && (
          <MessageBar
            intent="warning"
            className={styles.teamsAccessBanner}
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
          <MessageBar intent="warning" className={styles.teamsAccessBanner}>
            <MessageBarBody>
              <b>{strings.TeamsAccessPendingTitle}</b><br />
              {fstr('TeamsAccessPendingMessage')}
            </MessageBarBody>
          </MessageBar>
        )}
        {versionMismatch && showVersionMismatchHint && (
          <MessageBar intent="warning" className={styles.teamsAccessBanner}>
            <MessageBarBody>
              <b>{strings.VersionMismatchTitle}</b><br />
              {strings.VersionMismatchMessage}
            </MessageBarBody>
          </MessageBar>
        )}
      </section>
    </FluentProvider>
  );
};

export default GuestSponsorInfo;


