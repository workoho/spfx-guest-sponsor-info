import * as React from 'react';
import { Callout, DirectionalHint } from '@fluentui/react';
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
  Available:       'Available',
  AvailableIdle:   'Available, Idle',
  Away:            'Away',
  BeRightBack:     'Be Right Back',
  Busy:            'Busy',
  BusyIdle:        'Busy, Idle',
  DoNotDisturb:    'Do Not Disturb',
  Offline:         'Offline',
  PresenceUnknown: '',
};

interface ISponsorCardProps {
  sponsor: ISponsor;
  /** AAD tenant ID of the host tenant — used to build Teams guest-context deep links. */
  hostTenantId: string;
  /** Controlled by the parent SponsorList — true when this card's rich popup should be visible. */
  isActive: boolean;
  /** Called when this card wants to show its popup. Parent cancels any pending hide timer. */
  onActivate: () => void;
  /** Called when the mouse/focus leaves this card or its popup. Parent starts the hide timer. */
  onScheduleDeactivate: () => void;
}

const SponsorCard: React.FC<ISponsorCardProps> = ({
  sponsor,
  hostTenantId,
  isActive,
  onActivate,
  onScheduleDeactivate,
}) => {
  const cardRef = React.useRef<HTMLDivElement>(null);
  const initials = getInitials(sponsor.displayName);
  const bgColor = getInitialsColor(sponsor.displayName);
  const presenceColor = sponsor.presence ? (PRESENCE_COLORS[sponsor.presence] ?? '#8A8886') : undefined;
  const presenceLabel = sponsor.presence ? (PRESENCE_LABELS[sponsor.presence] ?? '') : undefined;
  const managerInitials = sponsor.managerDisplayName ? getInitials(sponsor.managerDisplayName) : '';
  const managerBgColor = sponsor.managerDisplayName ? getInitialsColor(sponsor.managerDisplayName) : '#8A8886';
  const primaryPhone = sponsor.businessPhones?.[0] ?? sponsor.mobilePhone;

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
        tabIndex={0}
        role="button"
        aria-label={sponsor.displayName}
        aria-haspopup="dialog"
        aria-expanded={isActive}
      >
        <div className={styles.avatarWrapper}>
          <div className={styles.avatar}>
            {sponsor.photoUrl ? (
              <img src={sponsor.photoUrl} alt="" className={styles.photo} />
            ) : (
              <div className={styles.initials} style={{ backgroundColor: bgColor }}>
                {initials}
              </div>
            )}
          </div>
          {presenceColor && (
            <span
              className={styles.presenceDot}
              style={{ backgroundColor: presenceColor }}
              aria-hidden="true"
            />
          )}
        </div>
        <div className={styles.cardName}>{sponsor.displayName}</div>
        {sponsor.jobTitle && (
          <div className={styles.cardJobTitle}>{sponsor.jobTitle}</div>
        )}
      </div>

      {/* ── Rich contact card (Fluent UI Callout, renders in portal) ─── */}
      {isActive && (
        <Callout
          target={cardRef}
          onDismiss={onScheduleDeactivate}
          directionalHint={DirectionalHint.rightTopEdge}
          directionalHintForRTL={DirectionalHint.leftTopEdge}
          isBeakVisible={false}
          gapSpace={8}
          role="dialog"
          aria-label={`Contact details for ${sponsor.displayName}`}
          setInitialFocus={false}
        >
          <div
            className={styles.richCard}
            onMouseEnter={onActivate}
            onMouseLeave={onScheduleDeactivate}
          >
            {/* ── Header: large avatar + name / title / presence ─── */}
            <div className={styles.richHeader}>
              <div className={styles.richAvatarWrapper}>
                <div className={styles.richAvatar}>
                  {sponsor.photoUrl ? (
                    <img src={sponsor.photoUrl} alt="" className={styles.photo} />
                  ) : (
                    <div className={styles.initials} style={{ backgroundColor: bgColor, fontSize: '30px' }}>
                      {initials}
                    </div>
                  )}
                </div>
                {presenceColor && (
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
                {presenceLabel && (
                  <div className={styles.richPresenceLabel} style={{ color: presenceColor }}>
                    {presenceLabel}
                  </div>
                )}
              </div>
            </div>

            {/* ── Action buttons row ───────────────────────────────── */}
            {(sponsor.mail || primaryPhone) && (
              <div className={styles.richActions} role="toolbar" aria-label="Contact actions">
                {sponsor.mail && (
                  <a
                    href={`https://teams.cloud.microsoft/l/chat/0/0?users=${encodeURIComponent(sponsor.mail)}`}
                    className={styles.richAction}
                    target="_blank"
                    rel="noreferrer noopener"
                    title="Chat via your home Teams account"
                  >
                    <span className={styles.richActionIcon} aria-hidden="true">💬</span>
                    <span className={styles.richActionLabel}>Chat</span>
                  </a>
                )}
                {sponsor.mail && (
                  <a
                    href={`https://teams.cloud.microsoft/l/chat/0/0?users=${encodeURIComponent(sponsor.mail)}&tenantId=${encodeURIComponent(hostTenantId)}`}
                    className={styles.richAction}
                    target="_blank"
                    rel="noreferrer noopener"
                    title="Chat as guest in sponsor's tenant"
                  >
                    <span className={styles.richActionIcon} aria-hidden="true">💬</span>
                    <span className={styles.richActionLabel}>Chat (guest)</span>
                  </a>
                )}
                {sponsor.mail && (
                  <a
                    href={`mailto:${sponsor.mail}`}
                    className={styles.richAction}
                    title="Send email"
                  >
                    <span className={styles.richActionIcon} aria-hidden="true">✉️</span>
                    <span className={styles.richActionLabel}>Email</span>
                  </a>
                )}
                {primaryPhone && (
                  <a
                    href={`tel:${primaryPhone}`}
                    className={styles.richAction}
                    title="Call"
                  >
                    <span className={styles.richActionIcon} aria-hidden="true">📞</span>
                    <span className={styles.richActionLabel}>Call</span>
                  </a>
                )}
              </div>
            )}

            {/* ── Contact information section ──────────────────────── */}
            <div className={styles.richSectionTitle}>Contact information</div>
            <div className={styles.richSection}>
              {sponsor.mail && (
                <a href={`mailto:${sponsor.mail}`} className={styles.richInfoRow}>
                  <span className={styles.richInfoIcon} aria-hidden="true">✉️</span>
                  <div>
                    <div className={styles.richInfoMeta}>Email</div>
                    <div className={styles.richInfoValue}>{sponsor.mail}</div>
                  </div>
                </a>
              )}
              {sponsor.businessPhones?.map(phone => (
                <a key={phone} href={`tel:${phone}`} className={styles.richInfoRow}>
                  <span className={styles.richInfoIcon} aria-hidden="true">📞</span>
                  <div>
                    <div className={styles.richInfoMeta}>Work phone</div>
                    <div className={styles.richInfoValue}>{phone}</div>
                  </div>
                </a>
              ))}
              {sponsor.mobilePhone && (
                <a href={`tel:${sponsor.mobilePhone}`} className={styles.richInfoRow}>
                  <span className={styles.richInfoIcon} aria-hidden="true">📱</span>
                  <div>
                    <div className={styles.richInfoMeta}>Mobile</div>
                    <div className={styles.richInfoValue}>{sponsor.mobilePhone}</div>
                  </div>
                </a>
              )}
              {sponsor.officeLocation && (
                <div className={styles.richInfoRow}>
                  <span className={styles.richInfoIcon} aria-hidden="true">📍</span>
                  <div>
                    <div className={styles.richInfoMeta}>Work location</div>
                    <div className={styles.richInfoValue}>{sponsor.officeLocation}</div>
                  </div>
                </div>
              )}
            </div>

            {/* ── Organization section (manager) ───────────────────── */}
            {sponsor.managerDisplayName && (
              <>
                <div className={styles.richSectionTitle}>Organization</div>
                <div className={styles.richSection}>
                  <div className={styles.managerRow}>
                    <div className={styles.managerAvatar}>
                      {sponsor.managerPhotoUrl ? (
                        <img src={sponsor.managerPhotoUrl} alt="" className={styles.photo} />
                      ) : (
                        <div
                          className={styles.initials}
                          style={{ backgroundColor: managerBgColor, fontSize: '14px' }}
                        >
                          {managerInitials}
                        </div>
                      )}
                    </div>
                    <div>
                      <div className={styles.managerLabel}>Manager</div>
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
        </Callout>
      )}
    </>
  );
};

export default SponsorCard;
