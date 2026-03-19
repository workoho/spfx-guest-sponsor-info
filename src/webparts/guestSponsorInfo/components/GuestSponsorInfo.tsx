import * as React from 'react';
import { DisplayMode } from '@microsoft/sp-core-library';
import * as strings from 'GuestSponsorInfoWebPartStrings';
import styles from './GuestSponsorInfo.module.scss';
import type { IGuestSponsorInfoProps } from './IGuestSponsorInfoProps';
import { ISponsor } from '../services/ISponsor';
import { isGuestUser, getSponsors, getSponsorsViaProxy } from '../services/SponsorService';
import { MOCK_SPONSORS } from '../services/MockSponsorService';
import SponsorCard from './SponsorCard';

/**
 * Renders the sponsor grid and owns the single shared "active card" state.
 * Only one popup is ever visible at a time; switching cards cancels any
 * pending hide-timeout from the previously active card so there is no
 * overlap between the outgoing and incoming popup.
 */
const SponsorList: React.FC<{ sponsors: ISponsor[]; hostTenantId: string }> = ({ sponsors, hostTenantId }) => {
  const [activeId, setActiveId] = React.useState<string | null>(null);
  const hideTimeout = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  const activate = (id: string): void => {
    if (hideTimeout.current) { clearTimeout(hideTimeout.current); hideTimeout.current = null; }
    setActiveId(id);
  };
  const scheduleDeactivate = (): void => {
    hideTimeout.current = setTimeout(() => setActiveId(null), 150);
  };

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
          />
        </li>
      ))}
    </ul>
  );
};

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
}) => {
  const [sponsors, setSponsors] = React.useState<ISponsor[]>([]);
  const [allUnavailable, setAllUnavailable] = React.useState(false);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [proxyStatus, setProxyStatus] = React.useState<ProxyStatus>('checking');
  const [retryCount, setRetryCount] = React.useState(0);

  // Primary signal: pageContext.user.isExternalGuestUser (authoritative, set from Entra token).
  // Fallback: #EXT# in loginName (heuristic; may be absent when the SharePoint user profile
  // has not yet been created for the guest, causing SP.UserProfile to return HTTP 500).
  const isGuest = isExternalGuestUser || isGuestUser(loginName);
  const isEditMode = displayMode === DisplayMode.Edit;

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
    const loadFn = useProxy
      ? () => getSponsorsViaProxy(functionUrl as string, aadHttpClient!)
      : () => getSponsors(graphClient!);

    loadFn()
      .then(result => {
        if (!cancelled) {
          setSponsors(result.activeSponsors);
          // All unavailable = sponsors were assigned but every account is disabled/deleted.
          setAllUnavailable(result.activeSponsors.length === 0 && result.unavailableCount > 0);
          setLoading(false);
          setRetryCount(0);
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
          if (is4xx) {
            // Permanent error (e.g. 401, 403): stop retrying and show the message.
            setError(strings.ErrorMessage);
            setLoading(false);
          } else {
            // Transient error: retry with exponential backoff capped at 5 minutes.
            // Spinner stays visible — the user sees "Loading…" throughout.
            const delay = Math.min(3000 * Math.pow(3, retryCount), 5 * 60 * 1000);
            setTimeout(() => { if (!cancelled) setRetryCount(n => n + 1); }, delay);
          }
        }
      });

    return () => { cancelled = true; };
  }, [isGuest, isEditMode, graphClient, mockMode, functionUrl, aadHttpClient, retryCount]);

  // Presence refresh: silently re-fetch sponsor data every 5 minutes so that
  // presence indicators stay current without the user needing to reload the page.
  // Only active in view mode after at least one successful load (sponsors or empty result).
  const PRESENCE_REFRESH_MS = 5 * 60 * 1000;
  React.useEffect(() => {
    if (isEditMode || mockMode || !isGuest || loading || error) return;
    const useProxy = functionUrl !== undefined && aadHttpClient !== undefined;
    if (!useProxy && graphClient === undefined) return;

    const loadFn = useProxy
      ? () => getSponsorsViaProxy(functionUrl as string, aadHttpClient!)
      : () => getSponsors(graphClient!);

    const id = setInterval(() => {
      loadFn()
        .then(result => {
          setSponsors(result.activeSponsors);
          setAllUnavailable(result.activeSponsors.length === 0 && result.unavailableCount > 0);
        })
        .catch((err: unknown) => {
          // Presence refresh failures are silent — the existing data stays on screen.
          console.warn('[GuestSponsorInfo] Presence refresh failed:', err);
        });
    }, PRESENCE_REFRESH_MS);

    return () => clearInterval(id);
  }, [isEditMode, mockMode, isGuest, loading, error, graphClient, functionUrl, aadHttpClient]);

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
  return (
    <section className={styles.webPart}>
      {title && <h2 className={styles.title}>{title}</h2>}
      {loading && (
        <p className={styles.statusMessage}>{strings.LoadingMessage}</p>
      )}
      {!loading && error && (
        <p className={styles.statusMessage}>{error}</p>
      )}
      {!loading && !error && sponsors.length === 0 && allUnavailable && (
        <p className={styles.statusMessage}>{strings.SponsorUnavailableMessage}</p>
      )}
      {!loading && !error && sponsors.length === 0 && !allUnavailable && (
        <p className={styles.statusMessage}>{strings.NoSponsorsMessage}</p>
      )}
      {!loading && !error && sponsors.length > 0 && (
        <SponsorList sponsors={sponsors} hostTenantId={hostTenantId} />
      )}
    </section>
  );
};

export default GuestSponsorInfo;

