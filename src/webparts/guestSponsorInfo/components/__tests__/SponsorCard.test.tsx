// Mock the Fluent UI Callout so it renders inline (avoids portal / ResizeObserver issues).
// The Callout's rendered children are still fully exercised.
// NOTE: jest.mock must be placed before imports so the linting rule is satisfied
// (Jest hoists it automatically at runtime regardless of source position).
jest.mock('@fluentui/react', () => ({
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  Callout: ({ children, role }: { children: React.ReactNode; role?: string }) => (
    <div role={role ?? 'dialog'}>{children}</div>
  ),
  DirectionalHint: { rightTopEdge: 0, leftTopEdge: 1 },
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  Icon: ({ iconName, className }: { iconName: string; className?: string }) => (
    <i className={className} data-icon-name={iconName} />
  ),
  // Render children directly; tooltip callout behaviour is tested in Fluent UI itself.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  TooltipHost: ({ children }: { children: React.ReactNode }) => <>{children}</>,
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  Panel: ({ children, isOpen }: { children: React.ReactNode; isOpen?: boolean }) => (
    isOpen ? <div role="dialog">{children}</div> : null
  ),
  PanelType: { custom: 'custom' },
}));

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';
import SponsorCard from '../SponsorCard';
import type { ISponsor } from '../../services/ISponsor';

// ─── Fixtures ──────────────────────────────────────────────────────────────────

const BASE_SPONSOR: ISponsor = {
  id: 'aaaaaaaa-0000-0000-0000-000000000001',
  displayName: 'Alice Smith',
  mail: 'alice@contoso.com',
  jobTitle: 'Project Manager',
  department: 'Engineering',
  officeLocation: 'Berlin',
  businessPhones: ['+49 30 12345678'],
  mobilePhone: undefined,
  photoUrl: undefined,
};

// ─── DOM helpers ───────────────────────────────────────────────────────────────

let container: HTMLDivElement;

beforeEach(() => {
  container = document.createElement('div');
  document.body.appendChild(container);
});

afterEach(() => {
  act(() => { ReactDOM.unmountComponentAtNode(container); });
  container.remove();
});

function render(
  sponsor: ISponsor,
  hostTenantId = 'test-tenant-id',
  isActive = false,
  onActivate = jest.fn(),
  onScheduleDeactivate = jest.fn(),
  showPhones = true,
  showWorkLocation = true,
  showManager = true
): void {
  act(() => {
    ReactDOM.render(
      <SponsorCard
        sponsor={sponsor}
        hostTenantId={hostTenantId}
        isActive={isActive}
        onActivate={onActivate}
        onScheduleDeactivate={onScheduleDeactivate}
        showPhones={showPhones}
        showWorkLocation={showWorkLocation}
        showManager={showManager}
      />,
      container
    );
  });
}

function fireEvent(element: Element, eventName: string): void {
  // React 17 uses event delegation: onMouseEnter is triggered by native 'mouseover'
  // events bubbling to the root container, and onMouseLeave by 'mouseout'.
  // Dispatching the raw 'mouseenter'/'mouseleave' events (which normally don't bubble)
  // would not reach React's root listener even with bubbles:true.
  const nativeEvent =
    eventName === 'mouseenter' ? 'mouseover' :
    eventName === 'mouseleave' ? 'mouseout' :
    eventName;
  act(() => {
    element.dispatchEvent(new MouseEvent(nativeEvent, { bubbles: true, cancelable: true }));
  });
}

// ─── Tests ─────────────────────────────────────────────────────────────────────

describe('SponsorCard', () => {
  describe('basic rendering', () => {
    it('renders the display name', () => {
      render(BASE_SPONSOR);
      expect(container.textContent).toContain('Alice Smith');
    });

    it('renders the job title', () => {
      render(BASE_SPONSOR);
      expect(container.textContent).toContain('Project Manager');
    });

    it('has an accessible button role', () => {
      render(BASE_SPONSOR);
      expect(container.querySelector('[role="button"]')).not.toBeNull();
    });

    it('sets aria-label to the display name', () => {
      render(BASE_SPONSOR);
      const card = container.querySelector('[role="button"]');
      expect(card?.getAttribute('aria-label')).toBe('Alice Smith');
    });

    it('sets aria-expanded=false when not active', () => {
      render(BASE_SPONSOR);
      const card = container.querySelector('[role="button"]');
      expect(card?.getAttribute('aria-expanded')).toBe('false');
    });

    it('sets aria-expanded=true when active', () => {
      render(BASE_SPONSOR, 'test-tenant-id', true);
      const card = container.querySelector('[role="button"]');
      expect(card?.getAttribute('aria-expanded')).toBe('true');
    });
  });

  describe('avatar', () => {
    it('renders coloured initials when no photoUrl is provided', () => {
      render(BASE_SPONSOR);
      // The initials box uses the "initials" CSS class (echoed as-is by styleMock).
      const initialsEl = container.querySelector('[class="initials"]');
      expect(initialsEl).not.toBeNull();
      expect(initialsEl!.textContent).toBe('AS');
    });

    it('uses first-letter + last-word-first-letter for two-part names', () => {
      render({ ...BASE_SPONSOR, displayName: 'John Van Der Berg' });
      expect(container.querySelector('[class="initials"]')!.textContent).toBe('JB');
    });

    it('uses the first two characters for a single-word name', () => {
      render({ ...BASE_SPONSOR, displayName: 'Madonna' });
      expect(container.querySelector('[class="initials"]')!.textContent).toBe('MA');
    });

    it('renders an <img> element overlaid on the initials when photoUrl is provided', () => {
      // Initials are always rendered as the base layer; the photo fades in on top
      // via CSS absolute positioning so there is no pop-in layout shift.
      render({ ...BASE_SPONSOR, photoUrl: 'data:image/jpeg;base64,/9j/4AAQ' });
      const img = container.querySelector('img');
      expect(img).not.toBeNull();
      expect(img!.getAttribute('src')).toBe('data:image/jpeg;base64,/9j/4AAQ');
      // Initials div must still be present (it is the background layer).
      expect(container.querySelector('[class="initials"]')).not.toBeNull();
    });
  });

  describe('contact details overlay', () => {
    it('is not visible before activation', () => {
      render(BASE_SPONSOR);
      expect(container.querySelector('[role="dialog"]')).toBeNull();
    });

    it('appears when isActive=true and contains the email address', () => {
      render(BASE_SPONSOR, 'test-tenant-id', true);
      const dialog = container.querySelector('[role="dialog"]');
      expect(dialog).not.toBeNull();
      expect(dialog!.textContent).toContain('alice@contoso.com');
    });

    it('appears when isActive=true and contains the office phone', () => {
      render(BASE_SPONSOR, 'test-tenant-id', true);
      expect(container.querySelector('[role="dialog"]')!.textContent).toContain('+49 30 12345678');
    });

    it('calls onActivate when mouse enters the card', () => {
      const onActivate = jest.fn();
      render(BASE_SPONSOR, 'test-tenant-id', false, onActivate);
      const card = container.querySelector('[role="button"]') as HTMLElement;
      fireEvent(card, 'mouseenter');
      expect(onActivate).toHaveBeenCalled();
    });

    it('calls onScheduleDeactivate when mouse leaves the card', () => {
      const onScheduleDeactivate = jest.fn();
      render(BASE_SPONSOR, 'test-tenant-id', false, jest.fn(), onScheduleDeactivate);
      const card = container.querySelector('[role="button"]') as HTMLElement;
      fireEvent(card, 'mouseleave');
      expect(onScheduleDeactivate).toHaveBeenCalled();
    });

    it('calls onActivate when the card is clicked (tap on touch devices)', () => {
      const onActivate = jest.fn();
      render(BASE_SPONSOR, 'test-tenant-id', false, onActivate);
      const card = container.querySelector('[role="button"]') as HTMLElement;
      act(() => { card.click(); });
      expect(onActivate).toHaveBeenCalled();
    });

    it('shows the mobile phone when present and business phones are absent', () => {
      const sponsor: ISponsor = {
        ...BASE_SPONSOR,
        businessPhones: [],
        mobilePhone: '+1 555 0199',
      };
      render(sponsor, 'test-tenant-id', true);
      expect(container.querySelector('[role="dialog"]')!.textContent).toContain('+1 555 0199');
    });

    it('does not render email link when mail is absent', () => {
      render({ ...BASE_SPONSOR, mail: undefined }, 'test-tenant-id', true);
      const links = container.querySelectorAll('[role="dialog"] a[href^="mailto:"]');
      expect(links).toHaveLength(0);
    });

    it('renders a Teams chat link and a Teams call link when mail is present', () => {
      const tenantId = 'aaaabbbb-0000-0000-0000-000000000001';
      render(BASE_SPONSOR, tenantId, true);
      const links = Array.from(container.querySelectorAll('[role="dialog"] a[href*="teams.microsoft.com"]'));
      expect(links).toHaveLength(2);

      const chatLink = links.find(l => l.getAttribute('href')!.includes('/l/chat/'));
      const callLink = links.find(l => l.getAttribute('href')!.includes('/l/call/'));
      expect(chatLink).not.toBeNull();
      expect(callLink).not.toBeNull();

      // Chat link: tenantId before users
      const chatHref = chatLink!.getAttribute('href')!;
      expect(chatHref).toContain(`tenantId=${encodeURIComponent(tenantId)}`);
      expect(chatHref).toContain(encodeURIComponent('alice@contoso.com'));
      expect(chatHref.indexOf('tenantId')).toBeLessThan(chatHref.indexOf('users'));

      // Call link: Teams audio call deep link with withVideo=false
      const callHref = callLink!.getAttribute('href')!;
      expect(callHref).toContain(`tenantId=${encodeURIComponent(tenantId)}`);
      expect(callHref).toContain(encodeURIComponent('alice@contoso.com'));
      expect(callHref).toContain('withVideo=false');
      expect(callHref.indexOf('tenantId')).toBeLessThan(callHref.indexOf('users'));
    });

    it('does not render Teams links when mail is absent', () => {
      render({ ...BASE_SPONSOR, mail: undefined }, 'test-tenant-id', true);
      const links = container.querySelectorAll('[role="dialog"] a[href*="teams.microsoft.com"]');
      expect(links).toHaveLength(0);
    });

    it('does not render Teams Chat or Call links when hasTeams is false', () => {
      render({ ...BASE_SPONSOR, hasTeams: false }, 'test-tenant-id', true);
      const links = container.querySelectorAll('[role="dialog"] a[href*="teams.microsoft.com"]');
      expect(links).toHaveLength(0);
    });

    it('still renders the Email link when hasTeams is false', () => {
      render({ ...BASE_SPONSOR, hasTeams: false }, 'test-tenant-id', true);
      const links = container.querySelectorAll('[role="dialog"] a[href^="mailto:"]');
      expect(links.length).toBeGreaterThan(0);
    });
  });

  describe('presence indicator', () => {
    it('renders a presence dot when presence is set', () => {
      render({ ...BASE_SPONSOR, presence: 'Available' });
      expect(container.querySelector('[class="presenceDot"]')).not.toBeNull();
    });

    it('does not render a presence dot when presence is absent', () => {
      render(BASE_SPONSOR);
      expect(container.querySelector('[class="presenceDot"]')).toBeNull();
    });

    it('shows the presence label in the rich card when active', () => {
      render({ ...BASE_SPONSOR, presence: 'Away' }, 'test-tenant-id', true);
      expect(container.querySelector('[role="dialog"]')!.textContent).toContain('Away');
    });

    it('does not render a presence dot when hasTeams is false', () => {
      render({ ...BASE_SPONSOR, presence: 'Available', hasTeams: false });
      expect(container.querySelector('[class="presenceDot"]')).toBeNull();
    });

    it('does not show presence label in rich card when hasTeams is false', () => {
      render({ ...BASE_SPONSOR, presence: 'Away', hasTeams: false }, 'test-tenant-id', true);
      // The presence label text should not appear in the rich card
      const richPresenceLabels = container.querySelectorAll('[class="richPresenceLabel"]');
      expect(richPresenceLabels).toHaveLength(0);
    });
  });

  describe('manager / organisation section', () => {
    it('does not render Organisation section when manager is absent', () => {
      render(BASE_SPONSOR, 'test-tenant-id', true);
      expect(container.querySelector('[role="dialog"]')!.textContent).not.toContain('Organization');
    });

    it('renders manager name when managerDisplayName is set', () => {
      render(
        { ...BASE_SPONSOR, managerDisplayName: 'Bob Jones', managerJobTitle: 'CTO' },
        'test-tenant-id',
        true
      );
      const dialog = container.querySelector('[role="dialog"]')!;
      expect(dialog.textContent).toContain('Bob Jones');
      expect(dialog.textContent).toContain('CTO');
    });
  });

  describe('copy-to-clipboard buttons', () => {
    let clipboardWriteText: jest.Mock;

    beforeEach(() => {
      clipboardWriteText = jest.fn().mockResolvedValue(undefined);
      Object.defineProperty(navigator, 'clipboard', {
        value: { writeText: clipboardWriteText },
        configurable: true,
      });
      jest.useFakeTimers();
    });

    afterEach(() => {
      jest.useRealTimers();
    });

    it('renders a copy button for the email address', () => {
      render(BASE_SPONSOR, 'test-tenant-id', true);
      const dialog = container.querySelector('[role="dialog"]')!;
      const copyBtn = dialog.querySelector('button[aria-label="Copy email address"]');
      expect(copyBtn).not.toBeNull();
    });

    it('renders a copy button for the work phone', () => {
      render(BASE_SPONSOR, 'test-tenant-id', true);
      const dialog = container.querySelector('[role="dialog"]')!;
      const copyBtn = dialog.querySelector('button[aria-label="Copy work phone"]');
      expect(copyBtn).not.toBeNull();
    });

    it('renders a copy button for the office location', () => {
      render(BASE_SPONSOR, 'test-tenant-id', true);
      const dialog = container.querySelector('[role="dialog"]')!;
      const copyBtn = dialog.querySelector('button[aria-label="Copy work location"]');
      expect(copyBtn).not.toBeNull();
    });

    it('copies the email to the clipboard when the copy button is clicked', async () => {
      render(BASE_SPONSOR, 'test-tenant-id', true);
      const dialog = container.querySelector('[role="dialog"]')!;
      const copyBtn = dialog.querySelector('button[aria-label="Copy email address"]') as HTMLElement;
      await act(async () => { copyBtn.click(); });
      expect(clipboardWriteText).toHaveBeenCalledWith('alice@contoso.com');
    });

    it('switches to Copied! state after clicking and reverts after 1500 ms', async () => {
      render(BASE_SPONSOR, 'test-tenant-id', true);
      const dialog = container.querySelector('[role="dialog"]')!;
      const copyBtn = dialog.querySelector('button[aria-label="Copy email address"]') as HTMLElement;

      await act(async () => { copyBtn.click(); });
      expect(copyBtn.getAttribute('aria-label')).toBe('Copied!');

      act(() => { jest.advanceTimersByTime(1500); });
      expect(dialog.querySelector('button[aria-label="Copy email address"]')).not.toBeNull();
    });

    it('renders a copy button for the mobile phone when present', () => {
      const sponsor: ISponsor = { ...BASE_SPONSOR, businessPhones: [], mobilePhone: '+1 555 0199' };
      render(sponsor, 'test-tenant-id', true);
      const dialog = container.querySelector('[role="dialog"]')!;
      expect(dialog.querySelector('button[aria-label="Copy mobile number"]')).not.toBeNull();
    });
  });
});
