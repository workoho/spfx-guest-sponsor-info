import * as React from 'react';
import {
  Avatar,
  Button,
  FluentProvider,
  OverlayDrawer,
  DrawerHeader,
  DrawerHeaderTitle,
  DrawerBody,
  Persona,
  Popover,
  PopoverSurface,
  Link,
  Tooltip,
  makeStyles,
  tokens,
  mergeClasses,
} from '@fluentui/react-components';
import type { PresenceBadgeStatus, Theme } from '@fluentui/react-components';
import {
  bundleIcon,
  ChatRegular,
  ChatFilled,
  MailRegular,
  MailFilled,
  CallRegular,
  CallFilled,
  CopyRegular,
  CopyFilled,
  CheckmarkRegular,
  CheckmarkFilled,
  PhoneRegular,
  BuildingRegular,
  LocationRegular,
  DismissRegular,
} from '@fluentui/react-icons';

const ChatIcon = bundleIcon(ChatFilled, ChatRegular);
const MailIcon = bundleIcon(MailFilled, MailRegular);
const CallIcon = bundleIcon(CallFilled, CallRegular);
const CopyIcon = bundleIcon(CopyFilled, CopyRegular);
const CheckmarkIcon = bundleIcon(CheckmarkFilled, CheckmarkRegular);
import * as strings from 'GuestSponsorInfoWebPartStrings';
import { ISponsor } from '../services/ISponsor';
import styles from './GuestSponsorInfo.module.scss';

/** Fluent UI persona colours used as avatar backgrounds when no photo is available. */
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
 * Maps Graph presence availability and activity tokens to Fluent UI v9
 * PresenceBadge status. Focusing maps to do-not-disturb (closest v9 equivalent).
 */
function graphPresenceToPresenceBadge(
  availability: string | undefined,
  activity: string | undefined
): { status: PresenceBadgeStatus; isOutOfOffice: boolean } {
  if (activity === 'OutOfOffice') {
    return { status: 'out-of-office', isOutOfOffice: true };
  }
  if (activity === 'Focusing') {
    return { status: 'do-not-disturb', isOutOfOffice: false };
  }
  switch (availability) {
    case 'Available':
    case 'AvailableIdle':
      return { status: 'available', isOutOfOffice: false };
    case 'Away':
    case 'BeRightBack':
      return { status: 'away', isOutOfOffice: false };
    case 'Busy':
    case 'BusyIdle':
      return { status: 'busy', isOutOfOffice: false };
    case 'DoNotDisturb':
      return { status: 'do-not-disturb', isOutOfOffice: false };
    case 'Offline':
      return { status: 'offline', isOutOfOffice: false };
    default:
      return { status: 'unknown', isOutOfOffice: false };
  }
}

/**
 * Griffel styles for Persona text slots in the rich card header and manager row.
 * Replaces the SCSS classes that previously styled the manual avatar+div structure.
 */
const usePersonaStyles = makeStyles({
  // ── Rich card header (size="huge", 96px avatar) ──────────────────────────
  richName: {
    fontSize: tokens.fontSizeBase400,       // 16px
    fontWeight: tokens.fontWeightSemibold,  // 600
    color: tokens.colorNeutralForeground1,
    display: '-webkit-box' as 'flex',       // line-clamp for long names
    WebkitLineClamp: '2',
    WebkitBoxOrient: 'vertical' as 'horizontal',
    overflow: 'hidden',
  },
  richSecondary: {
    fontSize: tokens.fontSizeBase200,       // 12px — job title or department
    color: tokens.colorNeutralForeground2,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    marginTop: tokens.spacingHorizontalSNudge, // 6px — block 1→2 separator (name → text)
  },
  richTertiary: {
    fontSize: tokens.fontSizeBase200,       // 12px — department (same block as job title)
    fontWeight: tokens.fontWeightRegular,
    color: tokens.colorNeutralForeground3,  // slightly lighter than job title — matches manager style
    // no marginTop — tight within-block spacing (Gestalt proximity)
  },
  richPresenceLine: {
    fontSize: tokens.fontSizeBase300,       // 14px — matches contact info rows
    fontWeight: tokens.fontWeightRegular,
    color: tokens.colorNeutralForeground2,
    marginTop: tokens.spacingHorizontalSNudge, // 6px — block 2→3 separator
  },
  // ── Manager row (size="extra-large", 56px avatar) ────────────────────────
  managerName: {
    fontSize: tokens.fontSizeBase400,       // 16px — matches previous .managerName
    fontWeight: tokens.fontWeightRegular,   // 400
    color: tokens.colorNeutralForeground1,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  managerSecondary: {
    fontSize: tokens.fontSizeBase200,       // 12px — job title
    color: tokens.colorNeutralForeground2,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    marginTop: tokens.spacingHorizontalXS,  // 4px — block 1→2 separator (scaled for smaller avatar)
  },
  managerTertiary: {
    fontSize: tokens.fontSizeBase200,       // 12px — department (same block as job title)
    color: tokens.colorNeutralForeground3,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    // no marginTop — tight within-block spacing (Gestalt proximity)
  },
  // Override the internal avatar↔text gap to match the 24px card-edge padding.
  // Fluent sets --fui-Persona__avatar--spacing on the .fui-Persona__avatar element
  // itself (not inherited from the root), so we must target that element directly.
  // Our compound selector (.richPersona .fui-Persona__avatar) has specificity
  // 0,2,0 vs Fluent's internal single-class 0,1,0 — so we reliably win.
  richPersona: {
    '& .fui-Persona__avatar': {
      '--fui-Persona__avatar--spacing': tokens.spacingHorizontalXXL, // 24px = card edge
    },
  },
  managerPersona: {
    '& .fui-Persona__avatar': {
      '--fui-Persona__avatar--spacing': tokens.spacingHorizontalXXL, // 24px = card edge
    },
  },
});

/** Griffel styles for the icon-only action buttons in the rich card header. */
const useActionButtonStyles = makeStyles({
  actionButton: {
    // 44 × 44 px matches Fluent UI's `size="large"` icon-button footprint and
    // the Teams People Card action button size. The built-in `size` prop on
    // Button changes both padding and min-width/height; override explicitly so
    // we keep `appearance="subtle"` without other `large` layout side-effects
    // (e.g. larger label font). Icon size 24px matches Teams.
    padding: tokens.spacingHorizontalSNudge,
    borderRadius: tokens.borderRadiusMedium,
    minWidth: '44px',
    width: '44px',
    height: '44px',
    color: tokens.colorNeutralForeground2,
    backgroundColor: 'transparent',
    '& .fui-Button__icon': {
      fontSize: '24px',
      width: '24px',
      height: '24px',
    },
    '&:hover': {
      backgroundColor: 'transparent',
    },
    '&:hover:active': {
      backgroundColor: 'transparent',
    },
    // Colour change + filled icon swap only when hovering directly over the icon,
    // not the surrounding padding. The full button area remains clickable.
    '& .fui-Button__icon:hover': {
      color: tokens.colorNeutralForeground2BrandHover,
      '& .fui-Icon-filled': { display: 'inline' },
      '& .fui-Icon-regular': { display: 'none' },
    },
    '& .fui-Button__icon:hover:active': {
      color: tokens.colorNeutralForeground2BrandPressed,
    },
  },
});

/**
 * Small copy-to-clipboard button shown at the trailing edge of each contact row.
 * Switches to a checkmark for 1.5 s after a successful copy.
 */
const CopyButton: React.FC<{ value: string; ariaLabel: string }> = ({ value, ariaLabel }) => {
  const [copied, setCopied] = React.useState(false);
  const actionButtonClasses = useActionButtonStyles();

  const handleCopy = (e: React.MouseEvent<HTMLElement>): void => {
    e.stopPropagation();
    navigator.clipboard.writeText(value).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 1500);
    }).catch(() => { /* clipboard access denied – silently ignore */ });
  };

  return (
    <Tooltip content={copied ? strings.CopiedFeedback : ariaLabel} relationship="label">
      <Button
        appearance="subtle"
        icon={copied ? <CheckmarkIcon /> : <CopyIcon />}
        aria-label={copied ? strings.CopiedFeedback : ariaLabel}
        onClick={handleCopy}
        className={mergeClasses(actionButtonClasses.actionButton, styles.copyButton, copied ? styles.copyButtonCopied : '')}
        size="small"
      />
    </Tooltip>
  );
};

/**
 * Returns true when the primary pointer is coarse (touch device — phone, tablet,
 * Surface in tablet mode, etc.). Determines whether the rich contact card opens
 * in an OverlayDrawer (touch) or a Popover anchored to the card tile (pointer).
 *
 * Using `pointer: coarse` is more reliable than a viewport-width check:
 * - Tablets (iPad, Surface) have wide viewports but need the drawer UX.
 * - A 480px breakpoint would wrongly pick Popover on a landscape phone.
 *
 * Starts as false so it is safe in SSR and test environments (jsdom has no matchMedia).
 * Updates reactively when the pointer type changes (e.g. detaching a keyboard).
 */
function useIsMobile(): boolean {
  const [isMobile, setIsMobile] = React.useState(false);
  React.useEffect(() => {
    if (!window.matchMedia) return;
    const mq = window.matchMedia('(pointer: coarse)');
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
  /** Show the sponsor's department as the third line in the Persona header. */
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
  /**
   * Fluent v9 theme object — passed into a nested FluentProvider inside the
   * Popover/Drawer portal so that design tokens (avatar colours, presence
   * badge, etc.) cascade correctly outside the main FluentProvider DOM tree.
   */
  v9Theme?: Theme;
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
  v9Theme,
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

  const isOof = sponsor.presenceActivity === 'OutOfOffice';
  const { status: presenceBadgeStatus, isOutOfOffice: badgeOof } = graphPresenceToPresenceBadge(
    sponsor.presence, sponsor.presenceActivity
  );
  // Presence badge is shown only inside the rich card header (contact popup),
  // not on the thumbnail tile — consistent with the Teams People Card pattern
  // where the grid view stays clean and presence is revealed on hover/tap.
  const showPresenceBadge = isActive && showPresence && sponsor.hasTeams !== false && !!sponsor.presence;
  const badgeStatus: PresenceBadgeStatus | undefined = showPresenceBadge ? presenceBadgeStatus : undefined;
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
  const isMobile = useIsMobile();
  const actionButtonClasses = useActionButtonStyles();
  const personaClasses = usePersonaStyles();
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

  // Delayed expand: the detail sections slide open ~300 ms after the card
  // appears, matching the Microsoft Teams People Card pattern where the header
  // and action row are visible first and the body expands shortly after.
  // 300 ms = card-enter animation (180 ms) + short settle pause (~120 ms).
  const [detailsExpanded, setDetailsExpanded] = React.useState(false);
  React.useEffect(() => {
    if (!isActive) { setDetailsExpanded(false); return; }
    const timer = setTimeout(() => setDetailsExpanded(true), 300);
    return () => clearTimeout(timer);
  }, [isActive]);

  // The rich card body is defined here so it can be placed inside either
  // a Popover (desktop) or an OverlayDrawer (mobile) without duplicating the JSX.

  // Manager avatar size scales with the number of text rows shown:
  //   2 rows (name + job title + department) → 64 px  (one above extra-large natural)
  //   0–1 rows                              → 56 px  (extra-large natural size)
  const managerThreeLines =
    showManagerJobTitle && showManagerDepartment && !!sponsor.managerDepartment;
  const managerAvatarSize: 56 | 64 = managerThreeLines ? 64 : 56;

  const richBody = (
    <div
      className={mergeClasses(styles.richCard, isMobile && styles.richCardFlat)}
      onMouseEnter={!isMobile ? onActivate : undefined}
      onMouseLeave={!isMobile ? onScheduleDeactivate : undefined}
    >
      {/* ── Header panel: elevated rounded card (avatar + buttons) ─── */}
      <div className={styles.richCardHeaderPanel}>
      <div className={styles.richHeader}>
        <Persona
          size="huge"
          name={resolvedName}
          className={personaClasses.richPersona}
          textAlignment="center"
          secondaryText={
            showSponsorJobTitle && sponsor.jobTitle
              ? { children: sponsor.jobTitle, className: personaClasses.richSecondary }
              : showSponsorDepartment && sponsor.department
                ? { children: sponsor.department, className: personaClasses.richSecondary }
                : presenceLabel && showPresence && sponsor.hasTeams !== false
                  ? { children: presenceLabel, className: personaClasses.richPresenceLine }
                  : undefined
          }
          tertiaryText={
            showSponsorJobTitle && sponsor.jobTitle && showSponsorDepartment && sponsor.department
              ? { children: sponsor.department, className: personaClasses.richTertiary }
              : (showSponsorJobTitle && sponsor.jobTitle || showSponsorDepartment && sponsor.department)
                  && presenceLabel && showPresence && sponsor.hasTeams !== false
                ? { children: presenceLabel, className: personaClasses.richPresenceLine }
                : undefined
          }
          quaternaryText={
            showSponsorJobTitle && sponsor.jobTitle && showSponsorDepartment && sponsor.department
              && presenceLabel && showPresence && sponsor.hasTeams !== false
              ? { children: presenceLabel, className: personaClasses.richPresenceLine }
              : undefined
          }
          primaryText={{ className: personaClasses.richName }}
          avatar={{
            size: 96,
            image: showSponsorPhoto && sponsor.photoUrl ? { src: sponsor.photoUrl } : undefined,
            color: 'colorful',
            badge: badgeStatus ? { status: badgeStatus, outOfOffice: badgeOof } : undefined,
          }}
        />
      </div>{/* end richHeader */}

      {/* ── Action buttons row ───────────────────────────────── */}
      {sponsor.mail && (
        <div className={styles.richActions} role="toolbar" aria-label={strings.ContactActionsAriaLabel}>
          {sponsor.hasTeams !== false && sponsor.mail && (
            <Tooltip
              content={guestHasTeamsAccess === false ? fstr('TeamsNotReadyChatTooltip') : strings.ChatTitle.replace('{name}', resolvedName)}
              relationship="label"
            >
              <Button
                as={guestHasTeamsAccess === false ? 'button' : 'a'}
                href={guestHasTeamsAccess === false ? undefined : `https://teams.microsoft.com/l/chat/0/0?tenantId=${encodeURIComponent(hostTenantId)}&users=${encodeURIComponent(sponsor.mail)}`}
                disabledFocusable={guestHasTeamsAccess === false}
                appearance="subtle"
                icon={<ChatIcon />}
                target="_blank"
                rel="noreferrer noopener"
                className={actionButtonClasses.actionButton}
              />
            </Tooltip>
          )}
          {sponsor.mail && (
            <Tooltip content={strings.EmailTitle.replace('{name}', resolvedName)} relationship="label">
              <Button
                as="a"
                href={`mailto:${sponsor.mail}`}
                appearance="subtle"
                icon={<MailIcon />}
                className={actionButtonClasses.actionButton}
              />
            </Tooltip>
          )}
          {sponsor.hasTeams !== false && (
            <Tooltip
              content={guestHasTeamsAccess === false ? fstr('TeamsNotReadyCallTooltip') : strings.CallTitle.replace('{name}', resolvedName)}
              relationship="label"
            >
              <Button
                as={guestHasTeamsAccess === false ? 'button' : 'a'}
                href={guestHasTeamsAccess === false ? undefined : `https://teams.microsoft.com/l/call/0/0?tenantId=${encodeURIComponent(hostTenantId)}&users=${encodeURIComponent(sponsor.mail)}&withVideo=false`}
                disabledFocusable={guestHasTeamsAccess === false}
                appearance="subtle"
                icon={<CallIcon />}
                target="_blank"
                rel="noreferrer noopener"
                className={actionButtonClasses.actionButton}
              />
            </Tooltip>
          )}
        </div>
      )}

      </div>{/* end richCardHeaderPanel */}

      {/* ── Scrollable detail area (expands after delay) ─────── */}
      <div
        className={mergeClasses(styles.richCardBody, (isMobile || detailsExpanded) && styles.richCardBodyExpanded)}
      >

      {/* ── Contact section ─────────────────────────────────── */}
      <div className={styles.richSectionTitle}>{strings.ContactInfoSection}</div>
      <div className={styles.richSection}>
        {sponsor.mail && (
          <div className={`${styles.richInfoRow} ${styles.richInfoRowInteractive}`}>
            <MailRegular className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <Link href={`mailto:${sponsor.mail}`} className={styles.richInfoValue}>{sponsor.mail}</Link>
            </div>
            <CopyButton value={sponsor.mail} ariaLabel={strings.CopyEmailAriaLabel} />
          </div>
        )}
        {showBusinessPhones && sponsor.businessPhones?.map(phone => (
          <div key={phone} className={`${styles.richInfoRow} ${styles.richInfoRowInteractive}`}>
            <CallRegular className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <Link href={`tel:${phone}`} className={styles.richInfoValue}>{phone}</Link>
            </div>
            <CopyButton value={phone} ariaLabel={strings.CopyWorkPhoneAriaLabel} />
          </div>
        ))}
        {showMobilePhone && sponsor.mobilePhone && (
          <div className={`${styles.richInfoRow} ${styles.richInfoRowInteractive}`}>
            <PhoneRegular className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <Link href={`tel:${sponsor.mobilePhone}`} className={styles.richInfoValue}>{sponsor.mobilePhone}</Link>
            </div>
            <CopyButton value={sponsor.mobilePhone} ariaLabel={strings.CopyMobileAriaLabel} />
          </div>
        )}
        {showOfficeLocation && (
          <div className={`${styles.richInfoRow} ${styles.richInfoRowInteractive}`}>
            <BuildingRegular className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <div className={styles.richInfoValue}>{officeLocation}</div>
            </div>
            <CopyButton value={officeLocation!} ariaLabel={strings.CopyLocationAriaLabel} />
          </div>
        )}
        {hasCombinedAddress && (
          <>
            <div className={`${styles.richInfoRow} ${styles.richInfoRowInteractive}`}>
              <LocationRegular className={styles.richInfoIcon} aria-hidden="true" />
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

      {/* ── Reports to section (manager) ───────────────────────────── */}
      {showManager && sponsor.managerDisplayName && (
        <>
          <div className={styles.richSectionDivider} />
          <div className={styles.richSectionTitle}>{strings.ReportsToSection}</div>
          <div className={styles.richSection}>
            <div className={styles.managerRow}>
              <Persona
                size="extra-large"
                name={resolvedManagerName}
                className={personaClasses.managerPersona}
                textAlignment="center"
                secondaryText={
                  showManagerJobTitle && sponsor.managerJobTitle
                    ? { children: sponsor.managerJobTitle, className: personaClasses.managerSecondary }
                    : !showManagerJobTitle && showManagerDepartment && sponsor.managerDepartment
                      ? { children: sponsor.managerDepartment, className: personaClasses.managerSecondary }
                      : undefined
                }
                tertiaryText={
                  showManagerJobTitle && showManagerDepartment && sponsor.managerDepartment
                    ? { children: sponsor.managerDepartment, className: personaClasses.managerTertiary }
                    : undefined
                }
                primaryText={{ className: personaClasses.managerName }}
                avatar={{
                  size: managerAvatarSize,
                  image: showManagerPhoto && sponsor.managerPhotoUrl ? { src: sponsor.managerPhotoUrl } : undefined,
                  color: 'colorful',
                }}
              />
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
          <Avatar
            size={compact ? 40 : 72}
            name={resolvedName}
            image={showSponsorPhoto && sponsor.photoUrl ? { src: sponsor.photoUrl } : undefined}
            color="colorful"
          />
        </div>
        <div className={compact ? styles.cardNameCompact : styles.cardName}>
          {resolvedName}
        </div>
      </div>

      {/* ── Rich contact card (OverlayDrawer on mobile, Popover on desktop) ─── */}
      {!readOnly && isMobile && (
        <OverlayDrawer
          open={isActive}
          position="bottom"
          onOpenChange={(_, data) => { if (!data.open) onScheduleDeactivate(); }}
        >
          <FluentProvider theme={v9Theme}>
            <DrawerHeader>
              <DrawerHeaderTitle
                action={
                  <Button
                    appearance="subtle"
                    icon={<DismissRegular />}
                    onClick={onScheduleDeactivate}
                    aria-label="Close"
                  />
                }
              >
                {resolvedName}
              </DrawerHeaderTitle>
            </DrawerHeader>
            <DrawerBody>
              {richBody}
            </DrawerBody>
          </FluentProvider>
        </OverlayDrawer>
      )}
      {!readOnly && !isMobile && isActive && (
        <Popover
          open
          positioning={{
            target: cardRef.current,
            position: 'below',
            align: 'start',
            offset: { mainAxis: 8 },
            fallbackPositions: ['above'],
          }}
          onOpenChange={(_, data) => { if (!data.open) onScheduleDeactivate(); }}
        >
          <PopoverSurface
            role="dialog"
            aria-label={strings.ContactDetailsAriaLabel.replace('{0}', resolvedName)}
            style={{ padding: 0, boxShadow: 'none', border: 'none', borderRadius: 0, backgroundColor: 'transparent', overflow: 'visible' }}
            onMouseEnter={onActivate}
            onMouseLeave={onScheduleDeactivate}
          >
            <FluentProvider theme={v9Theme}>
              {richBody}
            </FluentProvider>
          </PopoverSurface>
        </Popover>
      )}
    </>
  );
};

export default SponsorCard;
