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

function resolveGeographicLocation(
  city: string | undefined,
  country: string | undefined,
  showCity: boolean,
  showCountry: boolean
): { label?: string; value?: string; copyAriaLabel?: string } {
  const cityValue = showCity ? city?.trim() : undefined;
  const countryValue = showCountry ? country?.trim() : undefined;
  const parts = [cityValue, countryValue].filter((part): part is string => Boolean(part));

  if (parts.length === 0) return {};
  if (parts.length === 2) {
    return {
      label: `${strings.CityFieldLabel} / ${strings.CountryFieldLabel}`,
      value: parts.join(', '),
      copyAriaLabel: `${strings.CopyCityAriaLabel} / ${strings.CopyCountryAriaLabel}`,
    };
  }
  if (cityValue) {
    return {
      label: strings.CityFieldLabel,
      value: cityValue,
      copyAriaLabel: strings.CopyCityAriaLabel,
    };
  }
  return {
    label: strings.CountryFieldLabel,
    value: countryValue,
    copyAriaLabel: strings.CopyCountryAriaLabel,
  };
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
const PRESENCE_LABELS: Record<string, string> = {
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

/**
 * Converts a Graph presence activity token (PascalCase, e.g. InAMeeting) into a
 * localised, human-readable label matching Microsoft's profile card behaviour.
 * All documented tokens are resolved via PRESENCE_LABELS; unknown tokens fall back
 * to a generic PascalCase word-splitter (English only).
 */
function formatPresenceActivity(activity: string): string {
  const normalized = activity.trim();
  if (!normalized || normalized === 'PresenceUnknown') return '';

  // Use the typed map for all documented tokens (also covers availability mirrors).
  const typed = PRESENCE_LABELS[normalized];
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

/** Styles for the stacked (icon-above-label) action buttons in the rich card. */
const actionButtonStyles: IButtonStyles = {
  root: {
    padding: '8px 12px',
    borderRadius: 4,
    minWidth: 60,
    height: 'auto',
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
    flexDirection: 'column',
    alignItems: 'center',
    gap: '4px',
  },
  icon: {
    fontSize: 20,
    lineHeight: '1',
    color: 'var(--themePrimary, #0078d4)',
    margin: 0,
    height: 'auto',
  },
  iconDisabled: {
    color: 'var(--neutralTertiary, #a19f9d)',
  },
  label: {
    fontSize: '10px',
    margin: 0,
    whiteSpace: 'nowrap',
    color: 'var(--neutralSecondary, #666)',
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
  /** Enable structured address rows (street/postal code/state). */
  showStructuredAddress: boolean;
  /** Show the sponsor's street address. */
  showStreetAddress: boolean;
  /** Show the sponsor's postal code. */
  showPostalCode: boolean;
  /** Show the sponsor's state or province. */
  showState: boolean;
  /** Enable inline map preview for address data. */
  showAddressMap: boolean;
  /** Optional Azure Maps subscription key used for inline preview. */
  azureMapsSubscriptionKey: string | undefined;
  /** External map provider used for fallback links. */
  externalMapProvider: 'bing' | 'google' | 'apple' | 'openstreetmap' | 'here';
  /** Show the manager section below the contact details. */
  showManager: boolean;
  /** Show the presence status indicator (dot) and label. */
  showPresence: boolean;
  /** Show the sponsor's job title in the rich card header. */
  showSponsorJobTitle: boolean;
  /** Show the manager's job title in the manager row. */
  showManagerJobTitle: boolean;
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
}

const SponsorCard: React.FC<ISponsorCardProps> = ({
  sponsor,
  hostTenantId,
  isActive,
  onActivate,
  onScheduleDeactivate,
  showBusinessPhones,
  showMobilePhone,
  showWorkLocation,
  showCity,
  showCountry,
  showStructuredAddress,
  showStreetAddress,
  showPostalCode,
  showState,
  showAddressMap,
  azureMapsSubscriptionKey,
  externalMapProvider,
  showManager,
  showPresence,
  showSponsorJobTitle,
  showManagerJobTitle,
  showSponsorDepartment,
  showManagerDepartment,
  useInformalAddress,
  guestHasTeamsAccess,
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
  const showPresenceIndicator = showPresence && sponsor.hasTeams !== false;
  const effectivePresence = showPresenceIndicator ? personaPresence : PersonaPresence.none;
  const effectiveOof = showPresenceIndicator ? personaOof : false;
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
      const base = availability ? (PRESENCE_LABELS[availability] ?? '') : '';
      const suffix = strings.PresenceOutOfOfficeSuffix || ', out of office';
      return base ? `${base}${suffix}` : (strings.PresenceOutOfOffice || 'Out of office');
    }
    if (activity) return formatPresenceActivity(activity);
    return availability ? (PRESENCE_LABELS[availability] ?? '') : undefined;
  }, [sponsor.presence, sponsor.presenceActivity, isOof]);
  const managerInitials = resolvedManagerName ? getInitials(resolvedManagerName) : '';
  const managerBgColor = resolvedManagerName ? getInitialsColor(resolvedManagerName) : '#8A8886';
  const isMobile = useIsMobile();
  const officeLocation = sponsor.officeLocation?.trim();
  const streetAddress = sponsor.streetAddress?.trim();
  const postalCode = sponsor.postalCode?.trim();
  const state = sponsor.state?.trim();
  const geographicLocation = resolveGeographicLocation(sponsor.city, sponsor.country, showCity, showCountry);
  const showGeographicLocation = Boolean(geographicLocation.value);
  const showOfficeLocation = Boolean(showWorkLocation && officeLocation);
  const showStreetAddressRow = Boolean(showStructuredAddress && showStreetAddress && streetAddress);
  const showPostalCodeRow = Boolean(showStructuredAddress && showPostalCode && postalCode);
  const showStateRow = Boolean(showStructuredAddress && showState && state);
  const mapAddressParts = [streetAddress, sponsor.city?.trim(), state, postalCode, sponsor.country?.trim(), officeLocation]
    .filter((part): part is string => Boolean(part));
  const mapAddress = mapAddressParts.join(', ');
  const hasAddressForMap = mapAddress.length > 0;
  const [mapPreviewUrl, setMapPreviewUrl] = React.useState<string | undefined>(undefined);
  const [mapLoading, setMapLoading] = React.useState(false);
  const [mapError, setMapError] = React.useState(false);
  const externalMapLink = hasAddressForMap ? buildExternalMapLink(externalMapProvider, mapAddress) : undefined;

  React.useEffect(() => {
    if (!isActive || !showAddressMap || !hasAddressForMap || !azureMapsSubscriptionKey) {
      setMapPreviewUrl(undefined);
      setMapLoading(false);
      setMapError(false);
      return;
    }

    const controller = new AbortController();
    setMapLoading(true);
    setMapError(false);

    const geocodeUrl = `https://atlas.microsoft.com/search/address/json?api-version=1.0&subscription-key=${encodeURIComponent(azureMapsSubscriptionKey)}&query=${encodeURIComponent(mapAddress)}&limit=1`;

    fetch(geocodeUrl, { signal: controller.signal })
      .then(async response => {
        if (!response.ok) {
          throw new Error(`Azure Maps geocode failed: ${response.status}`);
        }
        const payload = (await response.json()) as {
          results?: Array<{ position?: { lat?: number; lon?: number } }>;
        };
        const position = payload.results?.[0]?.position;
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
        setMapError(true);
      });

    return () => controller.abort();
  }, [isActive, showAddressMap, hasAddressForMap, azureMapsSubscriptionKey, mapAddress]);

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
            imageUrl={sponsor.photoUrl}
            imageShouldFadeIn
            presence={effectivePresence}
            isOutOfOffice={effectiveOof}
            hidePersonaDetails
          />
          {isFocusing && showPresenceIndicator && (
            <span
              className={styles.richPresenceDot}
              style={{ backgroundColor: PRESENCE_COLORS.Focusing }}
              aria-hidden="true"
            />
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
            <TooltipHost content={guestHasTeamsAccess === false ? fstr('TeamsNotReadyChatTooltip') : strings.ChatTitle}>
              <ActionButton
                href={guestHasTeamsAccess === false ? undefined : `https://teams.microsoft.com/l/chat/0/0?tenantId=${encodeURIComponent(hostTenantId)}&users=${encodeURIComponent(sponsor.mail)}`}
                disabled={guestHasTeamsAccess === false}
                iconProps={{ iconName: 'Chat' }}
                text={strings.ChatLabel}
                target="_blank"
                rel="noreferrer noopener"
                styles={actionButtonStyles}
              />
            </TooltipHost>
          )}
          {sponsor.mail && (
            <TooltipHost content={strings.EmailTitle}>
              <ActionButton
                href={`mailto:${sponsor.mail}`}
                iconProps={{ iconName: 'Mail' }}
                text={strings.EmailLabel}
                styles={actionButtonStyles}
              />
            </TooltipHost>
          )}
          {sponsor.hasTeams !== false && (
            <TooltipHost content={guestHasTeamsAccess === false ? fstr('TeamsNotReadyCallTooltip') : strings.CallTitle}>
              <ActionButton
                href={guestHasTeamsAccess === false ? undefined : `https://teams.microsoft.com/l/call/0/0?tenantId=${encodeURIComponent(hostTenantId)}&users=${encodeURIComponent(sponsor.mail)}&withVideo=false`}
                disabled={guestHasTeamsAccess === false}
                iconProps={{ iconName: 'Phone' }}
                text={strings.CallLabel}
                target="_blank"
                rel="noreferrer noopener"
                styles={actionButtonStyles}
              />
            </TooltipHost>
          )}
        </div>
      )}

      {/* ── Contact information section ──────────────────────── */}
      <div className={styles.richSectionTitle}>{strings.ContactInfoSection}</div>
      <div className={styles.richSection}>
        {sponsor.mail && (
          <div className={styles.richInfoRow}>
            <Icon iconName="Mail" className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <div className={styles.richInfoMeta}>{strings.EmailFieldLabel}</div>
              <Link href={`mailto:${sponsor.mail}`} className={styles.richInfoValue}>{sponsor.mail}</Link>
            </div>
            <CopyButton value={sponsor.mail} ariaLabel={strings.CopyEmailAriaLabel} />
          </div>
        )}
        {showBusinessPhones && sponsor.businessPhones?.map(phone => (
          <div key={phone} className={styles.richInfoRow}>
            <Icon iconName="Phone" className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <div className={styles.richInfoMeta}>{strings.WorkPhoneFieldLabel}</div>
              <Link href={`tel:${phone}`} className={styles.richInfoValue}>{phone}</Link>
            </div>
            <CopyButton value={phone} ariaLabel={strings.CopyWorkPhoneAriaLabel} />
          </div>
        ))}
        {showMobilePhone && sponsor.mobilePhone && (
          <div className={styles.richInfoRow}>
            <Icon iconName="CellPhone" className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <div className={styles.richInfoMeta}>{strings.MobileFieldLabel}</div>
              <Link href={`tel:${sponsor.mobilePhone}`} className={styles.richInfoValue}>{sponsor.mobilePhone}</Link>
            </div>
            <CopyButton value={sponsor.mobilePhone} ariaLabel={strings.CopyMobileAriaLabel} />
          </div>
        )}
        {showGeographicLocation && geographicLocation.label && geographicLocation.copyAriaLabel && (
          <div className={styles.richInfoRow}>
            <Icon iconName="MapPin" className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <div className={styles.richInfoMeta}>{geographicLocation.label}</div>
              <div className={styles.richInfoValue}>{geographicLocation.value}</div>
            </div>
            <CopyButton value={geographicLocation.value!} ariaLabel={geographicLocation.copyAriaLabel} />
          </div>
        )}
        {showOfficeLocation && (
          <div className={styles.richInfoRow}>
            <Icon iconName="MapPin" className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <div className={styles.richInfoMeta}>{strings.WorkLocationFieldLabel}</div>
              <div className={styles.richInfoValue}>{officeLocation}</div>
            </div>
            <CopyButton value={officeLocation!} ariaLabel={strings.CopyLocationAriaLabel} />
          </div>
        )}
        {showStreetAddressRow && (
          <div className={styles.richInfoRow}>
            <Icon iconName="MapPin" className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <div className={styles.richInfoMeta}>{strings.StreetAddressFieldLabel}</div>
              <div className={styles.richInfoValue}>{streetAddress}</div>
            </div>
            <CopyButton value={streetAddress!} ariaLabel={strings.CopyStreetAddressAriaLabel} />
          </div>
        )}
        {showPostalCodeRow && (
          <div className={styles.richInfoRow}>
            <Icon iconName="MapPin" className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <div className={styles.richInfoMeta}>{strings.PostalCodeFieldLabel}</div>
              <div className={styles.richInfoValue}>{postalCode}</div>
            </div>
            <CopyButton value={postalCode!} ariaLabel={strings.CopyPostalCodeAriaLabel} />
          </div>
        )}
        {showStateRow && (
          <div className={styles.richInfoRow}>
            <Icon iconName="MapPin" className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <div className={styles.richInfoMeta}>{strings.StateFieldLabel}</div>
              <div className={styles.richInfoValue}>{state}</div>
            </div>
            <CopyButton value={state!} ariaLabel={strings.CopyStateAriaLabel} />
          </div>
        )}
        {showAddressMap && hasAddressForMap && (
          <div className={styles.mapPreviewBlock}>
            <div className={styles.richInfoMeta}>{strings.AddressMapSectionLabel}</div>
            {azureMapsSubscriptionKey && mapLoading && (
              <div className={styles.mapPreviewStatus}>{strings.AddressMapLoadingLabel}</div>
            )}
            {azureMapsSubscriptionKey && mapPreviewUrl && (
              <img
                src={mapPreviewUrl}
                alt={strings.AddressMapSectionLabel}
                className={styles.mapPreviewImage}
                referrerPolicy="no-referrer"
              />
            )}
            {((azureMapsSubscriptionKey && mapError) || !azureMapsSubscriptionKey) && (
              <div className={styles.mapPreviewStatus}>{strings.AddressMapFallbackHint}</div>
            )}
            {externalMapLink && (
              <a
                href={externalMapLink}
                target="_blank"
                rel="noreferrer noopener"
                className={styles.mapPreviewLink}
              >
                {strings.OpenAddressInMapLabel}
              </a>
            )}
          </div>
        )}
      </div>

      {/* ── Organization section (department + manager) ─────────────── */}
      {((showManager && sponsor.managerDisplayName) || (showSponsorDepartment && sponsor.department)) && (
        <>
          <div className={styles.richSectionTitle}>{strings.OrganizationSection}</div>
          <div className={styles.richSection}>
            {/* Sponsor department row */}
            {showSponsorDepartment && sponsor.department && (
              <div className={styles.richInfoRow}>
                <Icon iconName="Org" className={styles.richInfoIcon} aria-hidden="true" />
                <div className={styles.richInfoText}>
                  <div className={styles.richInfoMeta}>{strings.DepartmentLabel}</div>
                  <div className={styles.departmentValue}>{sponsor.department}</div>
                </div>
              </div>
            )}
            {/* Manager row */}
            {showManager && sponsor.managerDisplayName && (
              <div className={styles.managerRow}>
                <Persona
                  size={PersonaSize.size40}
                  initialsColor={managerBgColor}
                  imageInitials={managerInitials}
                  imageUrl={sponsor.managerPhotoUrl}
                  imageShouldFadeIn
                  hidePersonaDetails
                />
                <div className={styles.managerText}>
                  <div className={styles.managerLabel}>{strings.ManagerLabel}</div>
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
            )}
          </div>
        </>
      )}
    </div>
  );

  return (
    <>
      {/* ── Card thumbnail (always visible in the grid) ──────────────── */}
      <div
        ref={cardRef}
        className={styles.card}
        onMouseEnter={onActivate}
        onMouseLeave={onScheduleDeactivate}
        onFocus={onActivate}
        onBlur={onScheduleDeactivate}
        onClick={onActivate}
        tabIndex={0}
        role="button"
        aria-label={resolvedName}
        aria-haspopup="dialog"
        aria-expanded={isActive}
      >
        <div className={styles.avatarWrapper}>
          <Persona
            size={PersonaSize.size72}
            initialsColor={bgColor}
            imageInitials={initials}
            imageUrl={sponsor.photoUrl}
            imageShouldFadeIn
            presence={effectivePresence}
            isOutOfOffice={effectiveOof}
            hidePersonaDetails
          />
          {isFocusing && showPresenceIndicator && (
            <span
              className={styles.presenceDot}
              style={{ backgroundColor: PRESENCE_COLORS.Focusing }}
              aria-hidden="true"
            />
          )}
        </div>
        <div className={styles.cardName}>
          {resolvedName}
        </div>
      </div>

      {/* ── Rich contact card (Panel on mobile, Callout on desktop) ─── */}
      {isActive && (
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
            directionalHint={DirectionalHint.rightTopEdge}
            directionalHintForRTL={DirectionalHint.leftTopEdge}
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
