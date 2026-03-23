import * as React from 'react';
import { ActionButton, Callout, DirectionalHint, Icon, IconButton, Link, Panel, PanelType, Persona, PersonaPresence, PersonaSize, TooltipHost } from '@fluentui/react';
import type { IButtonStyles } from '@fluentui/react';
import * as strings from 'GuestSponsorInfoWebPartStrings';
import { ISponsor } from '../services/ISponsor';
import styles from './GuestSponsorInfo.module.scss';

/** Fluent UI persona colours used as avatar backgrounds when no photo is available. */
const PERSONA_COLORS = [
  '#D13438', '#CA5010', '#986F0B', '#498205',
  '#038387', '#004E8C', '#8764B8', '#69797E',
  '#C19C00', '#00B294', '#E3008C', '#0099BC',
];

/** Derives a consistent colour from a display name string. */
function getInitialsColor(name: string): string {
  let hash = 0;
  for (let i = 0; i < name.length; i++) {
    hash = (hash << 5) - hash + name.charCodeAt(i);
    hash |= 0; // Convert to 32-bit integer
  }
  return PERSONA_COLORS[Math.abs(hash) % PERSONA_COLORS.length];
}

/** Extracts up to two initials from a display name. */
function getInitials(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length >= 2) {
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }
  return name.substring(0, 2).toUpperCase();
}

/**
 * Returns "givenName surname" when either part is non-empty, otherwise falls
 * back to displayName. Mirrors how Microsoft renders names in Teams/Outlook.
 */
function resolvePersonName(
  givenName: string | undefined,
  surname: string | undefined,
  displayName: string | undefined
): string {
  const first = givenName?.trim() ?? '';
  const last = surname?.trim() ?? '';
  if (first || last) return [first, last].filter(Boolean).join(' ');
  return displayName?.trim() ?? '';
}

function buildExternalMapLink(
  provider: 'bing' | 'google' | 'apple' | 'openstreetmap' | 'here',
  address: string
): string {
  const query = encodeURIComponent(address);
  switch (provider) {
    case 'google':
      return `https://www.google.com/maps/search/?api=1&query=${query}`;
    case 'apple':
      return `https://maps.apple.com/?q=${query}`;
    case 'openstreetmap':
      return `https://www.openstreetmap.org/search?query=${query}`;
    case 'here':
      return `https://wego.here.com/search/${query}`;
    case 'bing':
    default:
      return `https://www.bing.com/maps?q=${query}`;
  }
}

/** Maps Graph presence availability (and activity) token → display colour. */
const PRESENCE_COLORS: Record<string, string> = {
  // Matches Fluent PersonaPresence defaults for v8 (Microsoft).
  Available:       '#6BB700',
  AvailableIdle:   '#6BB700',
  Away:            '#FFAA44',
  BeRightBack:     '#FFAA44',
  Busy:            '#C43148',
  BusyIdle:        '#C43148',
  DoNotDisturb:    '#C50F1F',
  Focusing:        '#6264A7',
  InACall:         '#C43148',
  InAMeeting:      '#C43148',
  OutOfOffice:     '#B4009E',
  Presenting:      '#C43148',
  Offline:         '#8A8886',
  PresenceUnknown: '#8A8886',
};

/**
 * Maps Graph presence availability and activity tokens → localised label.
 * Activity tokens (InAMeeting, InACall, …) take priority over the base availability token,
 * matching Microsoft's profile card display behaviour.
 */
/**
 * Returns a map of Graph presence tokens to localised labels.
 * Evaluated lazily (called at render time, not at module load time) so that
 * the SPFx AMD string bundle is guaranteed to be loaded before access.
 */
function getPresenceLabels(): Record<string, string> {
  return {
    // availability tokens
    Available:       strings.PresenceAvailable,
    AvailableIdle:   strings.PresenceAvailableIdle,
    Away:            strings.PresenceAway,
    BeRightBack:     strings.PresenceBeRightBack,
    Busy:            strings.PresenceBusy,
    BusyIdle:        strings.PresenceBusyIdle,
    DoNotDisturb:    strings.PresenceDoNotDisturb,
    Offline:         strings.PresenceOffline,
    PresenceUnknown: '',
    // activity-specific tokens (refine the base availability label)
    Focusing:        strings.PresenceFocusing,
    InACall:         strings.PresenceInACall,
    InAMeeting:      strings.PresenceInAMeeting,
    OffWork:         strings.PresenceOffline,
    OutOfOffice:     strings.PresenceOutOfOffice,
    Presenting:      strings.PresencePresenting,
  };
}

/**
 * Converts a Graph presence activity token (PascalCase, e.g. InAMeeting) into a
 * localised, human-readable label matching Microsoft's profile card behaviour.
 * All documented tokens are resolved via getPresenceLabels(); unknown tokens fall back
 * to a generic PascalCase word-splitter (English only).
 */
function formatPresenceActivity(activity: string): string {
  const normalized = activity.trim();
  if (!normalized || normalized === 'PresenceUnknown') return '';

  // Use the typed map for all documented tokens (also covers availability mirrors).
  const typed = getPresenceLabels()[normalized];
  if (typed !== undefined) return typed;

  // Fallback for undocumented activity tokens: split PascalCase into words.
  const words = normalized
    .replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2')
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .split(/\s+/)
    .filter(Boolean);

  if (words.length === 0) return '';

  const lowercaseJoiners = new Set([
    'a', 'an', 'and', 'as', 'at', 'by', 'for', 'from', 'in', 'of', 'on', 'or', 'the', 'to', 'with',
  ]);

  return words
    .map((w, index) => {
      const lower = w.toLowerCase();
      if (index > 0 && lowercaseJoiners.has(lower)) return lower;
      return index === 0 ? lower.charAt(0).toUpperCase() + lower.slice(1) : lower;
    })
    .join(' ');
}

/**
 * Maps Graph presence availability and activity tokens to Fluent UI v8 Persona
 * presence props. All standard states are handled natively by Persona; only
 * Focusing requires a custom presence span (no Fluent enum equivalent).
 */
function graphPresenceToPersonaPresence(
  availability: string | undefined,
  activity: string | undefined
): { presence: PersonaPresence; isOutOfOffice: boolean } {
  if (activity === 'OutOfOffice') {
    return { presence: PersonaPresence.away, isOutOfOffice: true };
  }
  if (activity === 'Focusing') {
    return { presence: PersonaPresence.none, isOutOfOffice: false };
  }
  switch (availability) {
    case 'Available':
    case 'AvailableIdle':
      return { presence: PersonaPresence.online, isOutOfOffice: false };
    case 'Away':
    case 'BeRightBack':
      return { presence: PersonaPresence.away, isOutOfOffice: false };
    case 'Busy':
    case 'BusyIdle':
      return { presence: PersonaPresence.busy, isOutOfOffice: false };
    case 'DoNotDisturb':
      return { presence: PersonaPresence.dnd, isOutOfOffice: false };
    case 'Offline':
      return { presence: PersonaPresence.offline, isOutOfOffice: false };
    default:
      return { presence: PersonaPresence.none, isOutOfOffice: false };
  }
}

/**
 * Small copy-to-clipboard button shown at the trailing edge of each contact row.
 * Switches to a checkmark for 1.5 s after a successful copy.
 */
const CopyButton: React.FC<{ value: string; ariaLabel: string }> = ({ value, ariaLabel }) => {
  const [copied, setCopied] = React.useState(false);

  const handleCopy = (e: React.MouseEvent<HTMLElement>): void => {
    e.stopPropagation();
    navigator.clipboard.writeText(value).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 1500);
    }).catch(() => { /* clipboard access denied – silently ignore */ });
  };

  return (
    <TooltipHost content={copied ? strings.CopiedFeedback : ariaLabel}>
      <IconButton
        iconProps={{ iconName: copied ? 'Accept' : 'Copy' }}
        ariaLabel={copied ? strings.CopiedFeedback : ariaLabel}
        onClick={handleCopy}
        className={`${styles.copyButton}${copied ? ` ${styles.copyButtonCopied}` : ''}`}
        styles={{
          root: { background: 'none', border: 'none', borderRadius: 4, color: 'inherit' },
          rootHovered: { background: 'var(--neutralLight, #edebe9)' },
          icon: { fontSize: 14, lineHeight: '1' },
        }}
      />
    </TooltipHost>
  );
};

/** Styles for the icon-only action buttons in the rich card header. */
const actionButtonStyles: IButtonStyles = {
  root: {
    padding: '8px',
    borderRadius: 4,
    minWidth: 40,
    width: 40,
    height: 40,
    border: 'none',
    background: 'none',
  },
  rootHovered: {
    background: 'var(--neutralLight, #edebe9)',
    textDecoration: 'none',
  },
  rootDisabled: {
    opacity: 0.4,
    background: 'none',
  },
  flexContainer: {
    justifyContent: 'center',
    alignItems: 'center',
  },
  icon: {
    fontSize: 15, // Tier 1: matches richInfoIcon below
    lineHeight: '1',
    // neutralSecondary matches the info-row icons below — dark grey, no theme blue.
    color: 'var(--neutralSecondary, #666666)',
    margin: 0,
    height: 'auto',
  },
  iconHovered: {
    color: 'var(--neutralSecondary, #666666)', // stays the same on hover
  },
  iconDisabled: {
    color: 'var(--neutralTertiary, #a19f9d)',
  },
};

/**
 * Returns true when the viewport is ≤ 480 px (phone-sized).
 * Starts as false so it is safe in SSR and test environments (jsdom has no matchMedia).
 * Updates reactively when the viewport resizes across the breakpoint.
 */
function useIsMobile(): boolean {
  const [isMobile, setIsMobile] = React.useState(false);
  React.useEffect(() => {
    if (!window.matchMedia) return;
    const mq = window.matchMedia('(max-width: 480px)');
    setIsMobile(mq.matches);
    const handler = (e: MediaQueryListEvent): void => setIsMobile(e.matches);
    mq.addEventListener('change', handler);
    return () => mq.removeEventListener('change', handler);
  }, []);
  return isMobile;
}

interface ISponsorCardProps {
  sponsor: ISponsor;
  /** Entra ID tenant ID of the host tenant — used to build Teams guest-context deep links. */
  hostTenantId: string;
  /** When true, render a compact horizontal row instead of a full 136px tile. */
  compact: boolean;
  /** Controlled by the parent SponsorList — true when this card's rich popup should be visible. */
  isActive: boolean;
  /** Called when this card wants to show its popup. Parent cancels any pending hide timer. */
  onActivate: () => void;
  /** Called when the mouse/focus leaves this card or its popup. Parent starts the hide timer. */
  onScheduleDeactivate: () => void;
  /** Show business phone numbers in the contact details section. */
  showBusinessPhones: boolean;
  /** Show the mobile phone number in the contact details section. */
  showMobilePhone: boolean;
  /** Show the work location row in the contact details section. */
  showWorkLocation: boolean;
  /** Show the sponsor's city. */
  showCity: boolean;
  /** Show the sponsor's country or region. */
  showCountry: boolean;
  /** Show the sponsor's street address. */
  showStreetAddress: boolean;
  /** Show the sponsor's postal code. */
  showPostalCode: boolean;
  /** Show the sponsor's state or province. */
  showState: boolean;
  /** Optional Azure Maps subscription key used for inline preview. */
  azureMapsSubscriptionKey: string | undefined;
  /** External map provider used for fallback links. 'none' disables the link. */
  externalMapProvider: 'bing' | 'google' | 'apple' | 'openstreetmap' | 'here' | 'none';
  /** Show the manager section below the contact details. */
  showManager: boolean;
  /** Show the presence status indicator (dot) and label. */
  showPresence: boolean;
  /** Show the sponsor's job title in the rich card header. */
  showSponsorJobTitle: boolean;
  /** Show the manager's job title in the manager row. */
  showManagerJobTitle: boolean;
  /** Show the sponsor's profile photo. When false, only initials are shown. */
  showSponsorPhoto: boolean;
  /** Show the manager's profile photo. When false, only initials are shown. */
  showManagerPhoto: boolean;
  /** Show the sponsor's department in the Organization section. */
  showSponsorDepartment: boolean;
  /** Show the manager's department in the manager row. */
  showManagerDepartment: boolean;
  /** Use informal address for user-facing tooltips. */
  useInformalAddress: boolean;
  /**
   * Whether the signed-in guest's Teams service account has been provisioned.
   * false = disable Teams Chat and Call buttons and show an explanatory tooltip.
   * undefined = unknown (fail-open — buttons remain active).
   */
  guestHasTeamsAccess?: boolean;
  /**
   * When true, the card is displayed as a visual tile only — no hover popup,
   * no keyboard activation. Used when the sponsor account is unavailable
   * (disabled or deleted) and only the name is shown for context.
   */
  readOnly?: boolean;
}

const SponsorCard: React.FC<ISponsorCardProps> = ({
  sponsor,
  hostTenantId,
  compact,
  isActive,
  onActivate,
  onScheduleDeactivate,
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
  showSponsorPhoto,
  showManagerPhoto,
  showSponsorDepartment,
  showManagerDepartment,
  useInformalAddress,
  guestHasTeamsAccess,
  readOnly,
}) => {
  const cardRef = React.useRef<HTMLDivElement>(null);

  const resolvedName = resolvePersonName(sponsor.givenName, sponsor.surname, sponsor.displayName);
  const resolvedManagerName = resolvePersonName(sponsor.managerGivenName, sponsor.managerSurname, sponsor.managerDisplayName);

  // Pick informal string variant when the property is enabled and the locale provides one.
  const fstr = <K extends keyof typeof strings>(key: K): string => {
    if (useInformalAddress) {
      const informalKey = `${key}Informal` as keyof typeof strings;
      const informal = strings[informalKey];
      if (informal) return informal as string;
    }
    return strings[key] as string;
  };

  const initials = getInitials(resolvedName);
  const bgColor = getInitialsColor(resolvedName);
  const isFocusing = sponsor.presenceActivity === 'Focusing';
  const isOof = sponsor.presenceActivity === 'OutOfOffice';
  const { presence: personaPresence, isOutOfOffice: personaOof } = graphPresenceToPersonaPresence(
    sponsor.presence, sponsor.presenceActivity
  );
  const showPresenceIndicator = showPresence && sponsor.hasTeams !== false && isActive;
  const isOffline = sponsor.presence === 'Offline' || sponsor.presence === 'PresenceUnknown';
  const isAvailable = sponsor.presence === 'Available';
  const useCustomPresenceDot = showPresenceIndicator && (isFocusing || isOffline || isAvailable);
  const customDotColor = isFocusing ? PRESENCE_COLORS.Focusing : isAvailable ? PRESENCE_COLORS.Available : PRESENCE_COLORS.Offline;
  const effectivePresence = showPresenceIndicator && !useCustomPresenceDot ? personaPresence : PersonaPresence.none;
  const effectiveOof = showPresenceIndicator && !useCustomPresenceDot ? personaOof : false;
  const presenceColor = isOof
    ? PRESENCE_COLORS.OutOfOffice
    : sponsor.presence ? (PRESENCE_COLORS[sponsor.presence] ?? '#8A8886') : undefined;
  const presenceLabel = React.useMemo(() => {
    const availability = sponsor.presence;
    const activity = sponsor.presenceActivity;
    if (!availability && !activity) return undefined;
    if (isOof) {
      // OutOfOffice is a suffix modifier: "Available, out of office"
      // When availability mirrors a generic state, prepend it.
      const base = availability ? (getPresenceLabels()[availability] ?? '') : '';
      const suffix = strings.PresenceOutOfOfficeSuffix || ', out of office';
      return base ? `${base}${suffix}` : (strings.PresenceOutOfOffice || 'Out of office');
    }
    if (activity) return formatPresenceActivity(activity);
    return availability ? (getPresenceLabels()[availability] ?? '') : undefined;
  }, [sponsor.presence, sponsor.presenceActivity, isOof]);
  const managerInitials = resolvedManagerName ? getInitials(resolvedManagerName) : '';
  const managerBgColor = resolvedManagerName ? getInitialsColor(resolvedManagerName) : '#8A8886';
  const isMobile = useIsMobile();
  const officeLocation = sponsor.officeLocation?.trim();
  const streetAddress = sponsor.streetAddress?.trim();
  const postalCode = sponsor.postalCode?.trim();
  const state = sponsor.state?.trim();
  const showOfficeLocation = Boolean(showWorkLocation && officeLocation);

  // Build a combined address from all configured/available address parts.
  // officeLocation is intentionally excluded — it is shown as a separate field.
  const addressParts: string[] = [];
  if (showStreetAddress && streetAddress) addressParts.push(streetAddress);
  if (showPostalCode && postalCode) addressParts.push(postalCode);
  if (showCity && sponsor.city?.trim()) addressParts.push(sponsor.city!.trim());
  if (showState && state) addressParts.push(state);
  if (showCountry && sponsor.country?.trim()) addressParts.push(sponsor.country!.trim());
  const combinedAddress = addressParts.join(', ');
  const hasCombinedAddress = combinedAddress.length > 0;
  const addressMapLink = hasCombinedAddress && externalMapProvider !== 'none'
    ? buildExternalMapLink(externalMapProvider, combinedAddress)
    : undefined;

  const [mapPreviewUrl, setMapPreviewUrl] = React.useState<string | undefined>(undefined);
  const [mapLoading, setMapLoading] = React.useState(false);

  React.useEffect(() => {
    if (!isActive || !hasCombinedAddress || !azureMapsSubscriptionKey) {
      setMapPreviewUrl(undefined);
      setMapLoading(false);
      return;
    }

    const controller = new AbortController();
    setMapLoading(true);

    const geocodeUrl = `https://atlas.microsoft.com/search/address/json?api-version=1.0&subscription-key=${encodeURIComponent(azureMapsSubscriptionKey)}&query=${encodeURIComponent(combinedAddress)}&limit=1`;

    fetch(geocodeUrl, { signal: controller.signal })
      .then(async response => {
        if (!response.ok) {
          throw new Error(`Azure Maps geocode failed: ${response.status}`);
        }
        const payload = (await response.json()) as {
          results?: Array<{
            type?: string;
            entityType?: string;
            position?: { lat?: number; lon?: number };
          }>;
        };
        const result = payload.results?.[0];
        // Allow the map only when the match is specific enough to be meaningful.
        //
        // Precise (always show):
        //   'Point Address'  — exact house number
        //   'Address Range'  — number interpolated along a road
        //   'Street'         — specific road (no house number but still very local)
        //
        // Geography sub-types (show only for city-level precision):
        //   Municipality / MunicipalitySubdivision / Neighbourhood / PostalCodeArea
        //   → "Munich, Germany" resolves here → show ✅
        //
        // Too vague (suppress):
        //   Geography with entityType Country / CountrySubdivision / etc.
        //   → "Germany" alone would land here → hide ❌
        //   POI, Cross Street — wrong or ambiguous location
        const matchType = result?.type;
        const entityType = result?.entityType ?? '';
        const CITY_LEVEL_ENTITY_TYPES = new Set([
          'Municipality', 'MunicipalitySubdivision', 'Neighbourhood', 'PostalCodeArea',
        ]);
        const isPrecise =
          matchType === 'Point Address' ||
          matchType === 'Address Range' ||
          matchType === 'Street' ||
          (matchType === 'Geography' && CITY_LEVEL_ENTITY_TYPES.has(entityType));
        if (!isPrecise) {
          throw new Error(`Azure Maps match too imprecise: ${matchType}/${entityType}`);
        }
        const position = result?.position;
        const lat = position?.lat;
        const lon = position?.lon;
        if (lat === undefined || lon === undefined) {
          throw new Error('No map coordinates returned');
        }
        const staticMapUrl = `https://atlas.microsoft.com/map/static/png?api-version=1.0&subscription-key=${encodeURIComponent(azureMapsSubscriptionKey)}&center=${lon},${lat}&zoom=14&width=560&height=260&pins=default||${lon}%20${lat}`;
        setMapPreviewUrl(staticMapUrl);
        setMapLoading(false);
      })
      .catch(error => {
        if ((error as Error).name === 'AbortError') return;
        setMapPreviewUrl(undefined);
        setMapLoading(false);
      });

    return () => controller.abort();
  }, [isActive, hasCombinedAddress, azureMapsSubscriptionKey, combinedAddress]);

  // Delayed expand: the detail sections slide open after 2 s, mimicking the
  // Microsoft People Card behaviour where the header + actions appear first.
  const [detailsExpanded, setDetailsExpanded] = React.useState(false);
  React.useEffect(() => {
    if (!isActive) { setDetailsExpanded(false); return; }
    const timer = setTimeout(() => setDetailsExpanded(true), 1000);
    return () => clearTimeout(timer);
  }, [isActive]);

  // Show the rich card below (or above when viewport space is tight) the persona
  // tile. We pre-calculate at activation time whether the fully-expanded callout
  // fits below, so the direction is committed before the body animates open and
  // the callout never flips mid-expansion.
  // richCard max-height = min(389px, 80vh); Callout gapSpace = 8px.
  const [calloutHint, setCalloutHint] = React.useState<DirectionalHint>(DirectionalHint.bottomAutoEdge);
  React.useEffect(() => {
    if (!isActive || !cardRef.current) return;
    const rect = cardRef.current.getBoundingClientRect();
    const spaceBelow = window.innerHeight - rect.bottom;
    const expandedHeight = Math.min(389, window.innerHeight * 0.8) + 8; // richCard + gapSpace
    setCalloutHint(spaceBelow >= expandedHeight ? DirectionalHint.bottomAutoEdge : DirectionalHint.topAutoEdge);
  }, [isActive]);

  // The rich card body is defined here so it can be placed inside either
  // a Callout (desktop) or a Panel (mobile) without duplicating the JSX.
  const richBody = (
    <div
      className={styles.richCard}
      style={isMobile ? { width: 'auto', maxHeight: 'none', backgroundColor: 'transparent' } : undefined}
      onMouseEnter={!isMobile ? onActivate : undefined}
      onMouseLeave={!isMobile ? onScheduleDeactivate : undefined}
    >
      {/* ── Header: large avatar + name / title / presence ─── */}
      <div className={styles.richHeader}>
        <div className={styles.richAvatarWrapper}>
          <Persona
            size={PersonaSize.size72}
            initialsColor={bgColor}
            imageInitials={initials}
            imageUrl={showSponsorPhoto ? sponsor.photoUrl : undefined}
            imageShouldFadeIn
            presence={effectivePresence}
            isOutOfOffice={effectiveOof}
            hidePersonaDetails
          />
          {useCustomPresenceDot && (
            <span
              className={styles.richPresenceDot}
              style={{ backgroundColor: customDotColor }}
              aria-hidden="true"
            >
              {isOffline && (
                <svg viewBox="0 0 10 10" width="12" height="12" fill="none" aria-hidden="true">
                  <line x1="1" y1="1" x2="9" y2="9" stroke="#fff" strokeWidth="2" strokeLinecap="round" />
                  <line x1="9" y1="1" x2="1" y2="9" stroke="#fff" strokeWidth="2" strokeLinecap="round" />
                </svg>
              )}
              {isAvailable && (
                <svg viewBox="0 0 12 10" width="13" height="10" fill="none" aria-hidden="true">
                  <polyline points="1,5 5,9 11,1" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                </svg>
              )}
            </span>
          )}
        </div>
        <div className={styles.richHeaderText}>
          <div className={styles.richName}>{resolvedName}</div>
          {/* Job title in header, or department as fallback when job title is hidden */}
          {showSponsorJobTitle && sponsor.jobTitle ? (
            <div className={styles.richJobTitle}>{sponsor.jobTitle}</div>
          ) : !showSponsorJobTitle && showSponsorDepartment && sponsor.department ? (
            <div className={styles.richJobTitle}>{sponsor.department}</div>
          ) : null}
          {presenceLabel && showPresence && sponsor.hasTeams !== false && (
            <div className={styles.richPresenceLabel} style={{ color: presenceColor }}>
              {presenceLabel}
            </div>
          )}
        </div>
      </div>

      {/* ── Action buttons row ───────────────────────────────── */}
      {sponsor.mail && (
        <div className={styles.richActions} role="toolbar" aria-label={strings.ContactActionsAriaLabel}>
          {sponsor.hasTeams !== false && sponsor.mail && (
            <TooltipHost content={guestHasTeamsAccess === false ? fstr('TeamsNotReadyChatTooltip') : strings.ChatTitle.replace('{name}', resolvedName)}>
              <ActionButton
                href={guestHasTeamsAccess === false ? undefined : `https://teams.microsoft.com/l/chat/0/0?tenantId=${encodeURIComponent(hostTenantId)}&users=${encodeURIComponent(sponsor.mail)}`}
                disabled={guestHasTeamsAccess === false}
                iconProps={{ iconName: 'Chat' }}
                target="_blank"
                rel="noreferrer noopener"
                styles={actionButtonStyles}
              />
            </TooltipHost>
          )}
          {sponsor.mail && (
            <TooltipHost content={strings.EmailTitle.replace('{name}', resolvedName)}>
              <ActionButton
                href={`mailto:${sponsor.mail}`}
                iconProps={{ iconName: 'Mail' }}
                styles={actionButtonStyles}
              />
            </TooltipHost>
          )}
          {sponsor.hasTeams !== false && (
            <TooltipHost content={guestHasTeamsAccess === false ? fstr('TeamsNotReadyCallTooltip') : strings.CallTitle.replace('{name}', resolvedName)}>
              <ActionButton
                href={guestHasTeamsAccess === false ? undefined : `https://teams.microsoft.com/l/call/0/0?tenantId=${encodeURIComponent(hostTenantId)}&users=${encodeURIComponent(sponsor.mail)}&withVideo=false`}
                disabled={guestHasTeamsAccess === false}
                iconProps={{ iconName: 'Phone' }}
                target="_blank"
                rel="noreferrer noopener"
                styles={actionButtonStyles}
              />
            </TooltipHost>
          )}
        </div>
      )}

      {/* ── Scrollable detail area (expands after delay) ─────── */}
      <div
        className={`${styles.richCardBody}${isMobile || detailsExpanded ? ` ${styles.richCardBodyExpanded}` : ''}`}
      >

      {/* ── Contact section ─────────────────────────────────── */}
      <div className={styles.richSectionTitle}>{strings.ContactInfoSection}</div>
      <div className={styles.richSection}>
        {sponsor.mail && (
          <div className={`${styles.richInfoRow} ${styles.richInfoRowInteractive}`}>
            <Icon iconName="Mail" className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <Link href={`mailto:${sponsor.mail}`} className={styles.richInfoValue}>{sponsor.mail}</Link>
            </div>
            <CopyButton value={sponsor.mail} ariaLabel={strings.CopyEmailAriaLabel} />
          </div>
        )}
        {showBusinessPhones && sponsor.businessPhones?.map(phone => (
          <div key={phone} className={`${styles.richInfoRow} ${styles.richInfoRowInteractive}`}>
            <Icon iconName="Phone" className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <Link href={`tel:${phone}`} className={styles.richInfoValue}>{phone}</Link>
            </div>
            <CopyButton value={phone} ariaLabel={strings.CopyWorkPhoneAriaLabel} />
          </div>
        ))}
        {showMobilePhone && sponsor.mobilePhone && (
          <div className={`${styles.richInfoRow} ${styles.richInfoRowInteractive}`}>
            <Icon iconName="CellPhone" className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <Link href={`tel:${sponsor.mobilePhone}`} className={styles.richInfoValue}>{sponsor.mobilePhone}</Link>
            </div>
            <CopyButton value={sponsor.mobilePhone} ariaLabel={strings.CopyMobileAriaLabel} />
          </div>
        )}
        {showOfficeLocation && (
          <div className={`${styles.richInfoRow} ${styles.richInfoRowInteractive}`}>
            <Icon iconName="CityNext" className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <div className={styles.richInfoValue}>{officeLocation}</div>
            </div>
            <CopyButton value={officeLocation!} ariaLabel={strings.CopyLocationAriaLabel} />
          </div>
        )}
        {hasCombinedAddress && (
          <>
            <div className={`${styles.richInfoRow} ${styles.richInfoRowInteractive}`}>
              <Icon iconName="MapPin" className={styles.richInfoIcon} aria-hidden="true" />
              <div className={styles.richInfoText}>
                {addressMapLink ? (
                  <Link href={addressMapLink} target="_blank" rel="noreferrer noopener" className={styles.richInfoValue}>
                    {combinedAddress}
                  </Link>
                ) : (
                  <div className={styles.richInfoValue}>{combinedAddress}</div>
                )}
              </div>
              <CopyButton value={combinedAddress} ariaLabel={strings.CopyAddressAriaLabel} />
            </div>
            {addressMapLink && azureMapsSubscriptionKey && (mapLoading || mapPreviewUrl) && (
              <div className={styles.mapPreviewInline}>
                {mapLoading && !mapPreviewUrl && (
                  <div className={styles.mapPreviewStatus}>{strings.AddressMapLoadingLabel}</div>
                )}
                {mapPreviewUrl && (
                  <Link href={addressMapLink} target="_blank" rel="noreferrer noopener">
                    <img
                      src={mapPreviewUrl}
                      alt={strings.AddressMapSectionLabel}
                      className={styles.mapPreviewImage}
                      referrerPolicy="no-referrer"
                    />
                  </Link>
                )}
              </div>
            )}
          </>
        )}
      </div>

      {/* ── Organization section (department only) ─────────────────── */}
      {showSponsorDepartment && sponsor.department && (
        <>
          <div className={styles.richSectionTitle}>{strings.OrganizationSection}</div>
          <div className={styles.richSection}>
            <div className={styles.richInfoRow}>
              <Icon iconName="Org" className={styles.richInfoIcon} aria-hidden="true" />
              <div className={styles.richInfoText}>
                <div className={styles.departmentValue}>{sponsor.department}</div>
              </div>
            </div>
          </div>
        </>
      )}

      {/* ── Reports to section (manager) ───────────────────────────── */}
      {showManager && sponsor.managerDisplayName && (
        <>
          <div className={styles.richSectionTitle}>{strings.ReportsToSection}</div>
          <div className={styles.richSection}>
            <div className={styles.managerRow}>
              <Persona
                size={PersonaSize.size40}
                initialsColor={managerBgColor}
                imageInitials={managerInitials}
                imageUrl={showManagerPhoto ? sponsor.managerPhotoUrl : undefined}
                imageShouldFadeIn
                hidePersonaDetails
              />
              <div className={styles.managerText}>
                <div className={styles.managerName}>{resolvedManagerName}</div>
                {/* Manager job title, or department as fallback if job title is hidden */}
                {showManagerJobTitle && sponsor.managerJobTitle ? (
                  <div className={styles.managerJobTitle}>{sponsor.managerJobTitle}</div>
                ) : !showManagerJobTitle && showManagerDepartment && sponsor.managerDepartment ? (
                  <div className={styles.managerJobTitle}>{sponsor.managerDepartment}</div>
                ) : null}
                {/* Manager department below job title (only when both are shown) */}
                {showManagerJobTitle && showManagerDepartment && sponsor.managerDepartment && (
                  <div className={styles.managerDept}>{sponsor.managerDepartment}</div>
                )}
              </div>
            </div>
          </div>
        </>
      )}
      </div>{/* end richCardBody */}
    </div>
  );

  return (
    <>
      {/* ── Card thumbnail (always visible in the grid) ──────────────── */}
      <div
        ref={cardRef}
        className={`${compact ? styles.cardCompact : styles.card}${readOnly ? ` ${styles.cardReadOnly}` : ''}`}
        onMouseEnter={readOnly ? undefined : onActivate}
        onMouseLeave={readOnly ? undefined : onScheduleDeactivate}
        onFocus={readOnly ? undefined : onActivate}
        onBlur={readOnly ? undefined : onScheduleDeactivate}
        onClick={readOnly ? undefined : onActivate}
        tabIndex={readOnly ? undefined : 0}
        role={readOnly ? undefined : 'button'}
        aria-label={resolvedName}
        aria-haspopup={readOnly ? undefined : 'dialog'}
        aria-expanded={readOnly ? undefined : isActive}
      >
        <div className={compact ? styles.avatarWrapperCompact : styles.avatarWrapper}>
          <Persona
            size={compact ? PersonaSize.size40 : PersonaSize.size72}
            initialsColor={bgColor}
            imageInitials={initials}
            imageUrl={sponsor.photoUrl}
            imageShouldFadeIn
            presence={PersonaPresence.none}
            hidePersonaDetails
          />
        </div>
        <div className={compact ? styles.cardNameCompact : styles.cardName}>
          {resolvedName}
        </div>
      </div>

      {/* ── Rich contact card (Panel on mobile, Callout on desktop) ─── */}
      {!readOnly && isActive && (
        isMobile ? (
          <Panel
            isOpen
            type={PanelType.custom}
            customWidth="100%"
            isLightDismiss
            hasCloseButton
            headerText={resolvedName}
            onDismiss={() => onScheduleDeactivate()}
          >
            {richBody}
          </Panel>
        ) : (
          <Callout
            target={cardRef}
            onDismiss={onScheduleDeactivate}
            directionalHint={calloutHint}
            isBeakVisible={false}
            gapSpace={8}
            role="dialog"
            aria-label={strings.ContactDetailsAriaLabel.replace('{0}', resolvedName)}
            setInitialFocus={false}
          >
            {richBody}
          </Callout>
        )
      )}
    </>
  );
};

export default SponsorCard;
