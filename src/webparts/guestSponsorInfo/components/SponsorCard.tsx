// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0

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
import { buildExternalMapLink } from '../utils/mapProviderUtils';

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
 * Griffel styles for the card thumbnail tiles visible in the sponsor grid.
 * Covers both the default vertical layout (136px tiles) and the compact
 * horizontal row variant used in narrow SharePoint columns.
 */
const useCardTileStyles = makeStyles({
  card: {
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
    cursor: 'pointer',
    outline: 'none',
    backgroundColor: 'transparent',
    '&:focus-visible': {
      boxShadow: `0 0 0 2px ${tokens.colorStrokeFocus2}`,
    },
  },
  cardReadOnly: {
    cursor: 'default',
  },
  cardCompact: {
    position: 'relative' as const,
    display: 'inline-flex',
    flexDirection: 'row' as const,
    alignItems: 'center' as const,
    gap: tokens.spacingHorizontalMNudge,
    padding: tokens.spacingVerticalSNudge,
    borderRadius: tokens.borderRadiusMedium,
    cursor: 'pointer',
    outline: 'none',
    backgroundColor: 'transparent',
    maxWidth: '100%',
    '&:focus-visible': {
      boxShadow: `0 0 0 2px ${tokens.colorStrokeFocus2}`,
    },
  },
  avatarWrapper: {
    position: 'relative' as const,
    display: 'inline-flex',
  },
  avatarWrapperCompact: {
    position: 'relative' as const,
    display: 'inline-flex',
    flexShrink: 0,
  },
  cardName: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    textAlign: 'center' as const,
    color: tokens.colorNeutralForeground1,
    lineHeight: tokens.lineHeightBase300,
    maxWidth: '100%',
    display: '-webkit-box' as 'flex',
    WebkitLineClamp: '3',
    WebkitBoxOrient: 'vertical' as 'horizontal',
    overflow: 'hidden',
  },
  cardNameCompact: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
    lineHeight: tokens.lineHeightBase300,
    minWidth: 0,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap' as const,
  },
});

/**
 * Griffel styles for the rich contact card (the detail popup / drawer).
 *
 * Cross-class hover effects (e.g. hovering a row reveals the copy button)
 * use CSS custom properties set by the parent row and read by children.
 * This avoids Griffel's limitation of not being able to reference one
 * atomic class from another's descendant selector.
 */
const useRichCardStyles = makeStyles({
  richCard: {
    width: '360px',
    display: 'flex',
    flexDirection: 'column' as const,
    position: 'relative' as const,
    borderRadius: tokens.borderRadiusLarge,
    boxShadow: tokens.shadow16,
    animationName: {
      from: { opacity: 0, transform: 'translateY(-6px) scale(0.98)' },
      to: { opacity: 1, transform: 'translateY(0) scale(1)' },
    },
    animationDuration: '180ms',
    animationTimingFunction: 'cubic-bezier(0.16, 1, 0.3, 1)',
    animationFillMode: 'both',
    '@media (prefers-reduced-motion: reduce)': {
      animationName: 'none',
      animationDuration: '0s',
    },
  },
  richCardFlat: {
    width: 'auto',
    boxShadow: 'none',
    animationName: 'none',
    animationDuration: '0s',
  },
  richCardHeaderPanel: {
    position: 'relative' as const,
    zIndex: 2,
    flexShrink: 0,
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  richCardHeaderPanelFlat: {
    border: 'none',
    boxShadow: 'none',
    borderRadius: '0',
    backgroundColor: 'transparent',
  },
  richCardBody: {
    overflowY: 'auto' as const,
    maxHeight: '0',
    opacity: 0,
    position: 'relative' as const,
    zIndex: 1,
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderTop: 'none',
    borderRadius: `0 0 ${tokens.borderRadiusLarge} ${tokens.borderRadiusLarge}`,
    marginTop: '-8px',
    paddingTop: tokens.spacingVerticalS,
    transitionProperty: 'max-height, opacity',
    transitionDuration: `${tokens.durationSlower}, ${tokens.durationNormal}`,
    transitionTimingFunction: `${tokens.curveEasyEase}, ease-in-out`,
  },
  richCardBodyFlat: {
    border: 'none',
    borderRadius: '0',
    backgroundColor: 'transparent',
    marginTop: '0',
    paddingTop: '0',
  },
  richCardBodyExpanded: {
    maxHeight: 'min(300px, 50vh)',
    opacity: 1,
    paddingBottom: tokens.spacingVerticalXXL,
  },
  richCardBodyExpandedFlat: {
    maxHeight: 'none',
  },
  richHeader: {
    padding: `${tokens.spacingVerticalXXL} ${tokens.spacingHorizontalXXL} 0`,
  },
  richActions: {
    display: 'flex',
    justifyContent: 'flex-start',
    gap: tokens.spacingHorizontalM,
    padding: `0 ${tokens.spacingHorizontalXXL} ${tokens.spacingVerticalL}`,
  },
  richSectionTitle: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: '0.01em',
    padding: `${tokens.spacingVerticalXXL} ${tokens.spacingHorizontalXXL} ${tokens.spacingVerticalXXS}`,
    color: tokens.colorNeutralForeground1,
  },
  richSection: {
    padding: '0',
  },
  richSectionDivider: {
    height: '1px',
    backgroundColor: tokens.colorNeutralStroke2,
    margin: `${tokens.spacingVerticalXXL} ${tokens.spacingHorizontalXXL} 0`,
  },
  richInfoRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalM,
    minHeight: '32px',
    padding: `0 ${tokens.spacingHorizontalXXL}`,
    alignItems: 'center',
    color: 'inherit',
    position: 'relative' as const,
  },
  // Sets CSS custom properties on hover that child elements (richInfoValue,
  // copyButton) read via var(). Only applied on devices with a precise pointer.
  richInfoRowInteractive: {
    '@media (hover: hover)': {
      '&:hover': {
        backgroundColor: tokens.colorNeutralBackground2,
        '--gsi-info-brightness': 'brightness(0.75)',
        '--gsi-copy-opacity': '1',
      },
    },
  },
  richInfoText: {
    flex: '1',
    minWidth: '0',
  },
  richInfoIcon: {
    fontSize: '24px',
    flexShrink: 0,
    width: '24px',
    textAlign: 'center' as const,
    color: tokens.colorNeutralForeground2,
  },
  richInfoValue: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorBrandForeground1,
    overflowWrap: 'break-word' as const,
    textDecoration: 'none',
    '&:hover': {
      textDecoration: 'none',
    },
    // Stretch an invisible click target across the entire row so the user can
    // click anywhere in the row to follow the link.
    '&[href]::before': {
      content: '""',
      position: 'absolute' as const,
      inset: '0',
    },
    // Plain-text values (not links): darken on row hover via the custom property.
    '&:not(a)': {
      filter: 'var(--gsi-info-brightness, none)',
    },
  },
  copyButton: {
    position: 'relative' as const,
    zIndex: 1,
    flexShrink: 0,
    // Read from the parent row's --gsi-copy-opacity custom property, falling
    // back to 0 (hidden) when the row is not hovered.
    opacity: 'var(--gsi-copy-opacity, 0)' as unknown as number,
    transitionProperty: 'opacity',
    transitionDuration: '80ms',
    transitionTimingFunction: 'ease',
    '&:focus-visible': {
      opacity: 1,
    },
  },
  copyButtonCopied: {
    opacity: 1,
    color: tokens.colorStatusSuccessForeground1,
  },
  mapPreviewInline: {
    // 60px = row padding-left (24px) + icon width (24px) + gap (12px) — no spacing token available
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalXXL} ${tokens.spacingVerticalS} 60px`,
    display: 'flex',
    flexDirection: 'column' as const,
    gap: tokens.spacingHorizontalS,
  },
  mapPreviewImage: {
    width: '100%',
    maxWidth: '100%',
    height: 'auto',
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  mapPreviewStatus: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  managerRow: {
    padding: `${tokens.spacingVerticalSNudge} ${tokens.spacingHorizontalXXL} 0`,
  },
});

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
  const richClasses = useRichCardStyles();

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
        className={mergeClasses(actionButtonClasses.actionButton, richClasses.copyButton, copied ? richClasses.copyButtonCopied : '')}
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
  /**
   * Called on click or focus — activates the card immediately without any
   * hover delay, and pins it so a subsequent mouse-leave does not auto-close.
   */
  onActivateNow: () => void;
  /** Called when the mouse/focus leaves this card or its popup. Parent starts the hide timer. */
  onScheduleDeactivate: () => void;
  /**
   * Called when the card is explicitly dismissed (outside-click, Escape,
   * or the mobile drawer’s close button). Always closes regardless of pin state.
   */
  onForceDeactivate: () => void;
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
  externalMapProvider: 'bing' | 'google' | 'apple' | 'openstreetmap' | 'none';
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
  onActivateNow,
  onScheduleDeactivate,
  onForceDeactivate,
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
  const cardClasses = useCardTileStyles();
  const richClasses = useRichCardStyles();
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

  // Pre-calculate whether the popover should open above or below the card tile.
  // This is done once when isActive becomes true — before the Popover is mounted —
  // so the position is stable throughout the expand animation and never flips.
  // pinned: true on the Popover then locks that decision in for the lifetime of
  // the popup (Fluent v9 PositioningProps).
  //
  // Estimated full height of the expanded contact card (header + actions +
  // details + optional map + optional manager section). Conservative upper
  // bound so the card doesn't clip at the bottom on typical screen heights.
  const ESTIMATED_CARD_HEIGHT_PX = 560;
  const [popoverSide, setPopoverSide] = React.useState<'above' | 'below'>('below');
  React.useEffect(() => {
    if (!isActive || !cardRef.current) return;
    const rect = cardRef.current.getBoundingClientRect();
    const spaceBelow = window.innerHeight - rect.bottom;
    const spaceAbove = rect.top;
    // Prefer 'below'. Fall back to 'above' only when there is clearly more
    // room above than below for the fully expanded card.
    setPopoverSide(
      spaceBelow >= ESTIMATED_CARD_HEIGHT_PX || spaceBelow >= spaceAbove
        ? 'below'
        : 'above'
    );
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
      className={mergeClasses(richClasses.richCard, isMobile && richClasses.richCardFlat)}
      onMouseEnter={!isMobile ? onActivate : undefined}
      onMouseLeave={!isMobile ? onScheduleDeactivate : undefined}
    >
      {/* ── Header panel: elevated rounded card (avatar + buttons) ─── */}
      <div className={mergeClasses(richClasses.richCardHeaderPanel, isMobile && richClasses.richCardHeaderPanelFlat)}>
      <div className={richClasses.richHeader}>
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
        <div className={richClasses.richActions} role="toolbar" aria-label={strings.ContactActionsAriaLabel}>
          {sponsor.hasTeams !== false && sponsor.mail && (
            <Tooltip
              content={guestHasTeamsAccess === false ? fstr('TeamsNotReadyChatTooltip') : strings.ChatTitle.replace('{name}', resolvedName)}
              relationship="label"
            >
              <Button
                as={guestHasTeamsAccess === false ? 'button' : 'a'}
                href={guestHasTeamsAccess === false ? undefined : `https://teams.cloud.microsoft/l/chat/0/0?tenantId=${encodeURIComponent(hostTenantId)}&users=${encodeURIComponent(sponsor.mail)}`}
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
                href={guestHasTeamsAccess === false ? undefined : `https://teams.cloud.microsoft/l/call/0/0?tenantId=${encodeURIComponent(hostTenantId)}&users=${encodeURIComponent(sponsor.mail)}&withVideo=false`}
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
        className={mergeClasses(
          richClasses.richCardBody,
          isMobile && richClasses.richCardBodyFlat,
          (isMobile || detailsExpanded) && richClasses.richCardBodyExpanded,
          isMobile && richClasses.richCardBodyExpandedFlat,
        )}
      >

      {/* ── Contact section ─────────────────────────────────── */}
      <div className={richClasses.richSectionTitle}>{strings.ContactInfoSection}</div>
      <div className={richClasses.richSection}>
        {sponsor.mail && (
          <div className={mergeClasses(richClasses.richInfoRow, richClasses.richInfoRowInteractive)}>
            <MailRegular className={richClasses.richInfoIcon} aria-hidden="true" />
            <div className={richClasses.richInfoText}>
              <Link href={`mailto:${sponsor.mail}`} className={richClasses.richInfoValue}>{sponsor.mail}</Link>
            </div>
            <CopyButton value={sponsor.mail} ariaLabel={strings.CopyEmailAriaLabel} />
          </div>
        )}
        {showBusinessPhones && sponsor.businessPhones?.map(phone => (
          <div key={phone} className={mergeClasses(richClasses.richInfoRow, richClasses.richInfoRowInteractive)}>
            <CallRegular className={richClasses.richInfoIcon} aria-hidden="true" />
            <div className={richClasses.richInfoText}>
              <Link href={`tel:${phone}`} className={richClasses.richInfoValue}>{phone}</Link>
            </div>
            <CopyButton value={phone} ariaLabel={strings.CopyWorkPhoneAriaLabel} />
          </div>
        ))}
        {showMobilePhone && sponsor.mobilePhone && (
          <div className={mergeClasses(richClasses.richInfoRow, richClasses.richInfoRowInteractive)}>
            <PhoneRegular className={richClasses.richInfoIcon} aria-hidden="true" />
            <div className={richClasses.richInfoText}>
              <Link href={`tel:${sponsor.mobilePhone}`} className={richClasses.richInfoValue}>{sponsor.mobilePhone}</Link>
            </div>
            <CopyButton value={sponsor.mobilePhone} ariaLabel={strings.CopyMobileAriaLabel} />
          </div>
        )}
        {showOfficeLocation && (
          <div className={mergeClasses(richClasses.richInfoRow, richClasses.richInfoRowInteractive)}>
            <BuildingRegular className={richClasses.richInfoIcon} aria-hidden="true" />
            <div className={richClasses.richInfoText}>
              <div className={richClasses.richInfoValue}>{officeLocation}</div>
            </div>
            <CopyButton value={officeLocation!} ariaLabel={strings.CopyLocationAriaLabel} />
          </div>
        )}
        {hasCombinedAddress && (
          <>
            <div className={mergeClasses(richClasses.richInfoRow, richClasses.richInfoRowInteractive)}>
              <LocationRegular className={richClasses.richInfoIcon} aria-hidden="true" />
              <div className={richClasses.richInfoText}>
                {addressMapLink ? (
                  <Link href={addressMapLink} target="_blank" rel="noreferrer noopener" className={richClasses.richInfoValue}>
                    {combinedAddress}
                  </Link>
                ) : (
                  <div className={richClasses.richInfoValue}>{combinedAddress}</div>
                )}
              </div>
              <CopyButton value={combinedAddress} ariaLabel={strings.CopyAddressAriaLabel} />
            </div>
            {azureMapsSubscriptionKey && (mapLoading || mapPreviewUrl) && (
              <div className={richClasses.mapPreviewInline}>
                {mapLoading && !mapPreviewUrl && (
                  <div className={richClasses.mapPreviewStatus}>{strings.AddressMapLoadingLabel}</div>
                )}
                {mapPreviewUrl && (
                  addressMapLink ? (
                    <Link href={addressMapLink} target="_blank" rel="noreferrer noopener">
                      <img
                        src={mapPreviewUrl}
                        alt={strings.AddressMapSectionLabel}
                        className={richClasses.mapPreviewImage}
                        referrerPolicy="no-referrer"
                      />
                    </Link>
                  ) : (
                    <img
                      src={mapPreviewUrl}
                      alt={strings.AddressMapSectionLabel}
                      className={richClasses.mapPreviewImage}
                      referrerPolicy="no-referrer"
                    />
                  )
                )}
              </div>
            )}
          </>
        )}
      </div>

      {/* ── Reports to section (manager) ───────────────────────────── */}
      {showManager && sponsor.managerDisplayName && (
        <>
          <div className={richClasses.richSectionDivider} />
          <div className={richClasses.richSectionTitle}>{strings.ReportsToSection}</div>
          <div className={richClasses.richSection}>
            <div className={richClasses.managerRow}>
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
        className={mergeClasses(compact ? cardClasses.cardCompact : cardClasses.card, readOnly ? cardClasses.cardReadOnly : '')}
        onMouseEnter={readOnly ? undefined : onActivate}
        onMouseLeave={readOnly ? undefined : onScheduleDeactivate}
        onFocus={readOnly ? undefined : onActivateNow}
        onBlur={readOnly ? undefined : onScheduleDeactivate}
        onClick={readOnly ? undefined : onActivateNow}
        tabIndex={readOnly ? undefined : 0}
        role={readOnly ? undefined : 'button'}
        aria-label={resolvedName}
        aria-haspopup={readOnly ? undefined : 'dialog'}
        aria-expanded={readOnly ? undefined : isActive}
      >
        <div className={compact ? cardClasses.avatarWrapperCompact : cardClasses.avatarWrapper}>
          <Avatar
            size={compact ? 40 : 72}
            name={resolvedName}
            image={showSponsorPhoto && sponsor.photoUrl ? { src: sponsor.photoUrl } : undefined}
            color="colorful"
          />
        </div>
        <div className={compact ? cardClasses.cardNameCompact : cardClasses.cardName}>
          {resolvedName}
        </div>
      </div>

      {/* ── Rich contact card (OverlayDrawer on mobile, Popover on desktop) ─── */}
      {!readOnly && isMobile && (
        <OverlayDrawer
          open={isActive}
          position="bottom"
          onOpenChange={(_, data) => { if (!data.open) onForceDeactivate(); }}
        >
          <FluentProvider theme={v9Theme}>
            <DrawerHeader>
              <DrawerHeaderTitle
                action={
                  <Button
                    appearance="subtle"
                    icon={<DismissRegular />}
                    onClick={onForceDeactivate}
                    aria-label={strings.CloseLabel}
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
            position: popoverSide,
            align: 'start',
            offset: { mainAxis: 8 },
            // pinned prevents Fluent's positioning engine from re-evaluating the
            // flip axis while the card expands — the side chosen above is locked
            // in for the entire lifetime of the popup.
            pinned: true,
          }}
          onOpenChange={(_, data) => { if (!data.open) onForceDeactivate(); }}
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
