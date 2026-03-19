import * as React from 'react';
import { Callout, DirectionalHint, Icon, Panel, PanelType, TooltipHost } from '@fluentui/react';
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

/** Maps Graph presence availability token → display colour. */
const PRESENCE_COLORS: Record<string, string> = {
  Available:       '#107C10',
  AvailableIdle:   '#107C10',
  Away:            '#F7630C',
  BeRightBack:     '#F7630C',
  Busy:            '#D13438',
  BusyIdle:        '#D13438',
  DoNotDisturb:    '#D13438',
  Offline:         '#8A8886',
  PresenceUnknown: '#8A8886',
};

/** Maps Graph presence availability token → human-readable label. */
const PRESENCE_LABELS: Record<string, string> = {
  Available:       strings.PresenceAvailable,
  AvailableIdle:   strings.PresenceAvailableIdle,
  Away:            strings.PresenceAway,
  BeRightBack:     strings.PresenceBeRightBack,
  Busy:            strings.PresenceBusy,
  BusyIdle:        strings.PresenceBusyIdle,
  DoNotDisturb:    strings.PresenceDoNotDisturb,
  Offline:         strings.PresenceOffline,
  PresenceUnknown: '',
};

/**
 * Small copy-to-clipboard button shown at the trailing edge of each contact row.
 * Switches to a checkmark for 1.5 s after a successful copy.
 */
const CopyButton: React.FC<{ value: string; ariaLabel: string }> = ({ value, ariaLabel }) => {
  const [copied, setCopied] = React.useState(false);

  const handleCopy = (e: React.MouseEvent | React.KeyboardEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    navigator.clipboard.writeText(value).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 1500);
    }).catch(() => { /* clipboard access denied – silently ignore */ });
  };

  return (
    <TooltipHost content={copied ? strings.CopiedFeedback : ariaLabel}>
      <button
        type="button"
        className={`${styles.copyButton}${copied ? ` ${styles.copyButtonCopied}` : ''}`}
        onClick={handleCopy}
        onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') handleCopy(e); }}
        aria-label={copied ? strings.CopiedFeedback : ariaLabel}
      >
        <Icon iconName={copied ? 'Accept' : 'Copy'} className={styles.copyIcon} aria-hidden="true" />
      </button>
    </TooltipHost>
  );
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
  /** Show phone numbers (business + mobile) in the contact details section. */
  showPhones: boolean;
  /** Show the work location row in the contact details section. */
  showWorkLocation: boolean;
  /** Show the manager section below the contact details. */
  showManager: boolean;
}

const SponsorCard: React.FC<ISponsorCardProps> = ({
  sponsor,
  hostTenantId,
  isActive,
  onActivate,
  onScheduleDeactivate,
  showPhones,
  showWorkLocation,
  showManager,
}) => {
  const cardRef = React.useRef<HTMLDivElement>(null);
  const initials = getInitials(sponsor.displayName);
  const bgColor = getInitialsColor(sponsor.displayName);
  const presenceColor = sponsor.presence ? (PRESENCE_COLORS[sponsor.presence] ?? '#8A8886') : undefined;
  const presenceLabel = sponsor.presence ? (PRESENCE_LABELS[sponsor.presence] ?? '') : undefined;
  const managerInitials = sponsor.managerDisplayName ? getInitials(sponsor.managerDisplayName) : '';
  const managerBgColor = sponsor.managerDisplayName ? getInitialsColor(sponsor.managerDisplayName) : '#8A8886';
  const isMobile = useIsMobile();

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
          <div className={styles.richAvatar}>
            <div className={styles.initials} style={{ backgroundColor: bgColor, fontSize: '30px' }}>
              {initials}
            </div>
            {sponsor.photoUrl && (
              <img src={sponsor.photoUrl} alt="" className={styles.photo} />
            )}
          </div>
          {presenceColor && sponsor.hasTeams !== false && (
            <span
              className={styles.richPresenceDot}
              style={{ backgroundColor: presenceColor }}
              aria-hidden="true"
            />
          )}
        </div>
        <div className={styles.richHeaderText}>
          <div className={styles.richName}>{sponsor.displayName}</div>
          {sponsor.jobTitle && (
            <div className={styles.richJobTitle}>{sponsor.jobTitle}</div>
          )}
          {sponsor.department && (
            <div className={styles.richDept}>{sponsor.department}</div>
          )}
          {presenceLabel && sponsor.hasTeams !== false && (
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
            <TooltipHost content={strings.ChatTitle}>
              <a
                href={`https://teams.microsoft.com/l/chat/0/0?tenantId=${encodeURIComponent(hostTenantId)}&users=${encodeURIComponent(sponsor.mail)}`}
                className={styles.richAction}
                target="_blank"
                rel="noreferrer noopener"
              >
                <Icon iconName="Chat" className={styles.richActionIcon} aria-hidden="true" />
                <span className={styles.richActionLabel}>{strings.ChatLabel}</span>
              </a>
            </TooltipHost>
          )}
          {sponsor.mail && (
            <TooltipHost content={strings.EmailTitle}>
              <a
                href={`mailto:${sponsor.mail}`}
                className={styles.richAction}
              >
                <Icon iconName="Mail" className={styles.richActionIcon} aria-hidden="true" />
                <span className={styles.richActionLabel}>{strings.EmailLabel}</span>
              </a>
            </TooltipHost>
          )}
          {sponsor.hasTeams !== false && (
            <TooltipHost content={strings.CallTitle}>
              <a
                href={`https://teams.microsoft.com/l/call/0/0?tenantId=${encodeURIComponent(hostTenantId)}&users=${encodeURIComponent(sponsor.mail)}&withVideo=false`}
                className={styles.richAction}
                target="_blank"
                rel="noreferrer noopener"
              >
                <Icon iconName="Phone" className={styles.richActionIcon} aria-hidden="true" />
                <span className={styles.richActionLabel}>{strings.CallLabel}</span>
              </a>
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
              <a href={`mailto:${sponsor.mail}`} className={styles.richInfoValue}>{sponsor.mail}</a>
            </div>
            <CopyButton value={sponsor.mail} ariaLabel={strings.CopyEmailAriaLabel} />
          </div>
        )}
        {showPhones && sponsor.businessPhones?.map(phone => (
          <div key={phone} className={styles.richInfoRow}>
            <Icon iconName="Phone" className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <div className={styles.richInfoMeta}>{strings.WorkPhoneFieldLabel}</div>
              <a href={`tel:${phone}`} className={styles.richInfoValue}>{phone}</a>
            </div>
            <CopyButton value={phone} ariaLabel={strings.CopyWorkPhoneAriaLabel} />
          </div>
        ))}
        {showPhones && sponsor.mobilePhone && (
          <div className={styles.richInfoRow}>
            <Icon iconName="CellPhone" className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <div className={styles.richInfoMeta}>{strings.MobileFieldLabel}</div>
              <a href={`tel:${sponsor.mobilePhone}`} className={styles.richInfoValue}>{sponsor.mobilePhone}</a>
            </div>
            <CopyButton value={sponsor.mobilePhone} ariaLabel={strings.CopyMobileAriaLabel} />
          </div>
        )}
        {showWorkLocation && sponsor.officeLocation && (
          <div className={styles.richInfoRow}>
            <Icon iconName="MapPin" className={styles.richInfoIcon} aria-hidden="true" />
            <div className={styles.richInfoText}>
              <div className={styles.richInfoMeta}>{strings.WorkLocationFieldLabel}</div>
              <div className={styles.richInfoValue}>{sponsor.officeLocation}</div>
            </div>
            <CopyButton value={sponsor.officeLocation} ariaLabel={strings.CopyLocationAriaLabel} />
          </div>
        )}
      </div>

      {/* ── Organization section (manager) ───────────────────── */}
      {showManager && sponsor.managerDisplayName && (
        <>
          <div className={styles.richSectionTitle}>{strings.OrganizationSection}</div>
          <div className={styles.richSection}>
            <div className={styles.managerRow}>
              <div className={styles.managerAvatar}>
                <div
                  className={styles.initials}
                  style={{ backgroundColor: managerBgColor, fontSize: '14px' }}
                >
                  {managerInitials}
                </div>
                {sponsor.managerPhotoUrl && (
                  <img src={sponsor.managerPhotoUrl} alt="" className={styles.photo} />
                )}
              </div>
              <div className={styles.managerText}>
                <div className={styles.managerLabel}>{strings.ManagerLabel}</div>
                <div className={styles.managerName}>{sponsor.managerDisplayName}</div>
                {sponsor.managerJobTitle && (
                  <div className={styles.managerJobTitle}>{sponsor.managerJobTitle}</div>
                )}
              </div>
            </div>
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
        aria-label={sponsor.displayName}
        aria-haspopup="dialog"
        aria-expanded={isActive}
      >
        <div className={styles.avatarWrapper}>
          <div className={styles.avatar}>
            <div className={styles.initials} style={{ backgroundColor: bgColor }}>
              {initials}
            </div>
            {sponsor.photoUrl && (
              <img src={sponsor.photoUrl} alt="" className={styles.photo} />
            )}
          </div>
          {presenceColor && sponsor.hasTeams !== false && (
            <span
              className={styles.presenceDot}
              style={{ backgroundColor: presenceColor }}
              aria-hidden="true"
            />
          )}
        </div>
        <div className={sponsor.jobTitle ? `${styles.cardName} ${styles.cardNameWithJobTitle}` : styles.cardName}>
          {sponsor.displayName}
        </div>
        {sponsor.jobTitle && (
          <div className={styles.cardJobTitle}>{sponsor.jobTitle}</div>
        )}
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
            headerText={sponsor.displayName}
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
            aria-label={strings.ContactDetailsAriaLabel.replace('{0}', sponsor.displayName)}
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
