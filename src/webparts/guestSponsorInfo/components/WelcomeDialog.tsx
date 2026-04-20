// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0

import * as React from 'react';
import {
  Button,
  Field,
  Input,
  Link,
  Popover,
  PopoverSurface,
  PopoverTrigger,
  PortalMountNodeProvider,
  Radio,
  RadioGroup,
  Text,
  makeStyles,
  mergeClasses,
  tokens,
} from '@fluentui/react-components';
import {
  BeakerRegular,
  CheckmarkRegular,
  ChevronLeftRegular,
  ChevronRightRegular,
  CloudRegular,
  CodeRegular,
  CopyRegular,
  DismissRegular,
} from '@fluentui/react-icons';
import * as strings from 'GuestSponsorInfoWebPartStrings';
import type { IWelcomeSetupConfig } from './IGuestSponsorInfoProps';
import { isValidFunctionUrl, isValidGuid } from '../utils/fieldValidation';
import workohoLogo from '../assets/workoho-default-logo.svg';
import welcomeIllustration from '../assets/welcome-illustration.svg';
import welcomeIllustrationDark from '../assets/welcome-illustration-dark.svg';
import wizardConnectIllustration from '../assets/wizard-connect.svg';
import wizardConnectIllustrationDark from '../assets/wizard-connect-dark.svg';
import wizardSuccessIllustration from '../assets/wizard-success.svg';
import wizardSuccessIllustrationDark from '../assets/wizard-success-dark.svg';

/**
 * URL of the deployment guide on the project website.
 * Shown as a help link in step 2 when the admin chooses the API path.
 */
const GITHUB_SETUP_URL = 'https://guest-sponsor-info.workoho.cloud/setup';

/**
 * Builds the versioned "Deploy to Azure" portal URL for a given semver string.
 * Falls back to the latest release when the version is unavailable.
 */
function buildDeployToAzureUrl(semver: string | undefined): string {
  // Use raw.githubusercontent.com so the Azure Portal can fetch the template
  // without CORS issues. The releases/download redirect does not expose the
  // required CORS headers. Pin to the matching version tag when available;
  // fall back to main for pre-release or unknown versions.
  const ref = semver ? `v${semver}` : 'main';
  const templatePath =
    `https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/${ref}/azure-function/infra/azuredeploy.json`;
  return 'https://portal.azure.com/#create/Microsoft.Template/uri/' + encodeURIComponent(templatePath);
}

/**
 * Builds a versioned raw.githubusercontent.com URL for a setup script.
 * Falls back to the main branch when the version is unavailable.
 */
function buildScriptUrl(semver: string | undefined, scriptName: string): string {
  // Use raw.githubusercontent.com so PowerShell receives plain text, not a
  // binary octet-stream. Pin to the matching version tag when available;
  // fall back to main for pre-release or unknown versions.
  const ref = semver ? `v${semver}` : 'main';
  return `https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/${ref}/azure-function/infra/${scriptName}`;
}

/** Builds the PowerShell one-liner for the App Registration setup script. */
function buildStep1Command(semver: string | undefined): string {
  const url = buildScriptUrl(semver, 'setup-app-registration.ps1');
  return `& ([scriptblock]::Create((iwr '${url}').Content))`;
}

/** Builds the PowerShell one-liner for the Graph permissions setup script. */
function buildStep3Command(semver: string | undefined): string {
  const url = buildScriptUrl(semver, 'setup-graph-permissions.ps1');
  return `& ([scriptblock]::Create((iwr '${url}').Content))`;
}

// ── PowerShell syntax tokenizer ───────────────────────────────────────────────
// Handles the specific one-liner pattern used in the setup commands:
//   & ([scriptblock]::Create((iwr 'URL').Content))
// No external dependency — keeps the bundle lean.
type PsTokenType = 'op' | 'type' | 'method' | 'cmdlet' | 'str' | 'punct' | 'default';
interface IPsToken { t: PsTokenType; v: string; }

/** Cmdlet names and aliases that appear in the setup command one-liners. */
const PS_CMDLETS = new Set(['iwr', 'Invoke-WebRequest']);

function tokenizePwsh(cmd: string): IPsToken[] {
  const result: IPsToken[] = [];
  let i = 0;
  while (i < cmd.length) {
    const ch = cmd[i];
    // Single-quoted string  'text'
    if (ch === "'") {
      const end = cmd.indexOf("'", i + 1);
      if (end !== -1) { result.push({ t: 'str', v: cmd.slice(i, end + 1) }); i = end + 1; continue; }
      result.push({ t: 'default', v: cmd.slice(i) }); break;
    }
    // Type literal  [TypeName]
    if (ch === '[') {
      const end = cmd.indexOf(']', i + 1);
      if (end !== -1) { result.push({ t: 'type', v: cmd.slice(i, end + 1) }); i = end + 1; continue; }
    }
    // Static-method operator  ::MethodName
    if (ch === ':' && cmd[i + 1] === ':') {
      result.push({ t: 'op', v: '::' }); i += 2;
      let j = i; while (j < cmd.length && /\w/.test(cmd[j])) j++;
      if (j > i) { result.push({ t: 'method', v: cmd.slice(i, j) }); i = j; }
      continue;
    }
    // Call operator  &
    if (ch === '&') { result.push({ t: 'op', v: '&' }); i++; continue; }
    // Parentheses and dots
    if (ch === '(' || ch === ')' || ch === '.') { result.push({ t: 'punct', v: ch }); i++; continue; }
    // Word boundary — cmdlet or plain identifier
    if (/\w/.test(ch)) {
      let j = i; while (j < cmd.length && /[\w-]/.test(cmd[j])) j++;
      const word = cmd.slice(i, j);
      result.push({ t: PS_CMDLETS.has(word) ? 'cmdlet' : 'default', v: word }); i = j; continue;
    }
    result.push({ t: 'default', v: ch }); i++;
  }
  return result;
}

const useStyles = makeStyles({
  // ── Inline card wrapper ───────────────────────────────────────────────────
  // The wizard renders as plain DOM content inside the web part zone — no
  // portal, no z-index fight with SharePoint chrome.  A centred card provides
  // the same visual weight as a modal without any of the SPFx caveats.
  root: {
    display: 'flex',
    justifyContent: 'center',
    paddingTop: tokens.spacingVerticalXXL,
    paddingBottom: tokens.spacingVerticalXXL,
  },
  card: {
    maxWidth: '520px',
    width: '100%',
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow16,
    borderRadius: tokens.borderRadiusXLarge,
    // Clip any overflowing children (e.g. long code blocks in the deploy panel)
    // to the card boundary. This also enforces the border-radius visually.
    overflow: 'hidden' as const,
    padding: `${tokens.spacingVerticalXXL} ${tokens.spacingHorizontalXXL}`,
    // Reduce horizontal padding in narrow web-part zones (≤ 400 px viewport)
    // so content never gets squeezed to zero by the 40 px × 2 side padding.
    '@media (max-width: 400px)': {
      padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalL}`,
    },
  },
  // ── Card header row (title + optional dismiss button) ───────────────────────
  cardHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    gap: tokens.spacingHorizontalS,
    marginBottom: tokens.spacingVerticalL,
  },
  wizardTitle: {
    margin: 0,
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase500,
    color: tokens.colorNeutralForeground1,
    display: 'block',
    // flex-shrink so a long title never pushes the dismiss button off-screen.
    flexShrink: 1,
    minWidth: 0,
  },
  stepActions: {
    marginTop: tokens.spacingVerticalL,
    paddingTop: tokens.spacingVerticalM,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  // ── Step progress dots ─────────────────────────────────────────────────────
  stepDots: {
    display: 'flex',
    gap: tokens.spacingHorizontalXS,
    justifyContent: 'center',
    marginBottom: tokens.spacingVerticalL,
  },
  dot: {
    height: '8px',
    width: '8px',
    borderRadius: tokens.borderRadiusCircular,
    backgroundColor: tokens.colorNeutralStroke1,
    transition: 'width 0.2s ease, background-color 0.2s ease',
  },
  dotActive: {
    width: '24px',
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorBrandBackground,
  },
  dotDone: {
    backgroundColor: tokens.colorBrandBackground,
  },
  // ── Illustration placeholder area ─────────────────────────────────────────
  // Each step has a centred icon-in-circle that acts as a placeholder for
  // custom artwork. Replace the icon with an <img> pointing to your SVG asset
  // once the illustrations are ready.
  illustrationWrap: {
    display: 'flex',
    justifyContent: 'center',
    marginBottom: tokens.spacingVerticalL,
    paddingTop: tokens.spacingVerticalS,
  },
  illustrationImg: {
    maxWidth: '200px',
    width: '100%',
    display: 'block',
  },
  illustrationCircle: {
    // ILLUSTRATION PLACEHOLDER — Step 1 (Welcome):
    // Custom artwork suggestion: a stylised "sponsor card" layout showing a guest
    // user icon on the left connected by a dashed arc to two sponsor profile
    // cards on the right, rendered in Workoho brand blue (#0078D4) on a white
    // or light-blue background. Approx. 240 × 140 px SVG.
    width: '72px',
    height: '72px',
    borderRadius: tokens.borderRadiusCircular,
    backgroundColor: tokens.colorBrandBackground2,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
  },
  illustrationCircleSuccess: {
    // ILLUSTRATION PLACEHOLDER — Step 3 (Done):
    // Custom artwork suggestion: a large green checkmark shield with a subtle
    // confetti scatter, indicating successful setup completion.
    width: '72px',
    height: '72px',
    borderRadius: tokens.borderRadiusCircular,
    backgroundColor: tokens.colorPaletteGreenBackground2,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
  },
  illustrationIcon: {
    color: tokens.colorBrandForeground1,
  },
  illustrationIconSuccess: {
    color: tokens.colorPaletteGreenForeground1,
  },
  // ── Typography ─────────────────────────────────────────────────────────────
  body: {
    color: tokens.colorNeutralForeground1,
    lineHeight: tokens.lineHeightBase400,
    display: 'block',
    marginBottom: tokens.spacingVerticalM,
  },
  muted: {
    color: tokens.colorNeutralForeground2,
    lineHeight: tokens.lineHeightBase400,
    display: 'block',
  },
  workohoRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    marginTop: tokens.spacingVerticalXL,
    paddingTop: tokens.spacingVerticalM,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    // Visually de-emphasised: slightly smaller than body text so the attribution
    // reads as a footer note rather than content, without hiding it from review.
    fontSize: tokens.fontSizeBase200,
  },
  workohoLogo: {
    height: '16px',
    display: 'block',
    flexShrink: 0,
  },
  // ── Setup step (step 2) ────────────────────────────────────────────────────
  setupIntro: {
    color: tokens.colorNeutralForeground2,
    lineHeight: tokens.lineHeightBase400,
    display: 'block',
    marginBottom: tokens.spacingVerticalL,
  },
  // ILLUSTRATION PLACEHOLDER — Step 2 (Setup choice):
  // Custom artwork suggestion: a split-screen layout. Left side shows a cloud
  // icon with an Azure Functions "⚡" badge (representing the API path); right
  // side shows a beaker / preview eye with "DEMO" text (representing demo mode).
  // Both sides are separated by a thin vertical divider.  The currently selected
  // side is subtly highlighted with a brand-colour glow.  Approx. 320 × 120 px.
  optionCard: {
    position: 'relative' as const,
    zIndex: 2,
    display: 'flex',
    alignItems: 'flex-start',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusMedium,
    border: `2px solid ${tokens.colorNeutralStroke2}`,
    cursor: 'pointer',
    '&:hover': {
      border: `2px solid ${tokens.colorNeutralStroke1}`,
      backgroundColor: tokens.colorNeutralBackground2,
    },
  },
  optionCardSelected: {
    border: `2px solid ${tokens.colorBrandStroke1}`,
    backgroundColor: tokens.colorBrandBackground2,
    '&:hover': {
      border: `2px solid ${tokens.colorBrandStroke1}`,
      backgroundColor: tokens.colorBrandBackground2,
    },
  },
  // API card when selected: squared bottom corners so the deploy panel attaches flush.
  optionCardSelectedTop: {
    border: `2px solid ${tokens.colorBrandStroke1}`,
    backgroundColor: tokens.colorBrandBackground2,
    borderRadius: `${tokens.borderRadiusMedium} ${tokens.borderRadiusMedium} 0 0`,
    '&:hover': {
      border: `2px solid ${tokens.colorBrandStroke1}`,
      backgroundColor: tokens.colorBrandBackground2,
    },
  },
  optionIcon: {
    color: tokens.colorBrandForeground1,
    flexShrink: 0,
    marginTop: '2px',
  },
  optionText: {
    flex: 1,
  },
  apiFields: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    marginTop: tokens.spacingVerticalM,
    // Indent to align visually with the radio-button text; removed on narrow
    // columns where every pixel is precious for the URL/client-ID inputs.
    paddingLeft: tokens.spacingHorizontalXL,
    '@media (max-width: 400px)': {
      paddingLeft: '0',
    },
  },
  docsLink: {
    display: 'block',
    marginTop: tokens.spacingVerticalXS,
  },
  // ── Deploy panel (step 2, API path selected) — slides out from under the option card ──
  deployPanel: {
    position: 'relative' as const,
    zIndex: 1,
    // Pull the panel 2px up so its border-bottom sits exactly behind the option card's
    // bottom border when collapsed (maxHeight: 0), making the closed state invisible.
    marginTop: '-2px',
    overflow: 'hidden',
    // Explicit width so the flex child never expands beyond the card content area,
    // which is required for text-overflow: ellipsis inside to trigger correctly.
    width: '100%',
    boxSizing: 'border-box' as const,
    maxHeight: '0',
    opacity: 0,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    border: `2px solid ${tokens.colorBrandStroke1}`,
    borderTop: 'none',
    borderRadius: `0 0 ${tokens.borderRadiusMedium} ${tokens.borderRadiusMedium}`,
    paddingLeft: tokens.spacingHorizontalM,
    paddingRight: tokens.spacingHorizontalM,
    transitionProperty: 'max-height, opacity',
    transitionDuration: `${tokens.durationSlower}, ${tokens.durationNormal}`,
    transitionTimingFunction: `${tokens.curveEasyEase}, ease-in-out`,
    '@media (prefers-reduced-motion: reduce)': {
      transitionDuration: '0s',
    },
  },
  deployPanelExpanded: {
    // Step 3 command now renders on 4 lines (pre-formatted) → extra height needed.
    maxHeight: '800px',
    opacity: 1,
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalM,
  },
  // ── 3-step setup guide inside the deploy panel ────────────────────────────
  stepToggle: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    marginTop: tokens.spacingVerticalS,
    width: '100%',
    background: 'none',
    border: 'none',
    paddingTop: tokens.spacingVerticalXXS,
    paddingBottom: tokens.spacingVerticalXXS,
    paddingLeft: 0,
    paddingRight: 0,
    cursor: 'pointer',
    color: 'inherit',
    textAlign: 'left' as const,
  },
  stepToggleLabel: {
    flexGrow: 1,
    color: tokens.colorNeutralForeground1,
  },
  stepChevron: {
    color: tokens.colorNeutralForeground3,
    width: '16px',
    height: '16px',
    transitionProperty: 'transform',
    transitionDuration: tokens.durationNormal,
    transitionTimingFunction: tokens.curveEasyEase,
    flexShrink: 0,
    '@media (prefers-reduced-motion: reduce)': {
      transitionDuration: '0s',
    },
  },
  stepChevronOpen: {
    transform: 'rotate(90deg)',
  },
  stepContent: {
    maxHeight: 0,
    overflow: 'hidden' as const,
    opacity: 0,
    transitionProperty: 'max-height, opacity',
    transitionDuration: `${tokens.durationSlower}, ${tokens.durationNormal}`,
    transitionTimingFunction: `${tokens.curveEasyEase}, ease-in-out`,
    '@media (prefers-reduced-motion: reduce)': {
      transitionDuration: '0s',
    },
  },
  stepContentOpen: {
    maxHeight: '500px',
    opacity: 1,
  },
  stepNumber: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '20px',
    height: '20px',
    borderRadius: tokens.borderRadiusCircular,
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    fontSize: tokens.fontSizeBase100,
    fontWeight: tokens.fontWeightSemibold,
    flexShrink: 0,
  },
  codeWrap: {
    display: 'flex',
    flexDirection: 'column',
    marginTop: tokens.spacingVerticalXS,
    // Always dark — code blocks are terminal-style regardless of the site theme.
    backgroundColor: '#1e1e1e',
    borderRadius: tokens.borderRadiusMedium,
    // clip children to the rounded corners
    overflow: 'hidden' as const,
  },
  codeLangBar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingTop: tokens.spacingVerticalXXS,
    paddingBottom: tokens.spacingVerticalXXS,
    paddingLeft: tokens.spacingHorizontalS,
    backgroundColor: '#2d2d2d',
    borderBottom: '1px solid #454545',
  },
  codeLangLabel: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase100,
    color: '#a0a0a0',
  },
  codeLangActions: {
    display: 'flex',
    alignItems: 'center',
  },
  codeBlock: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase100,
    lineHeight: tokens.lineHeightBase200,
    color: '#d4d4d4',
    // Horizontal scrolling so the user can select and copy the full command.
    overflowX: 'auto' as const,
    whiteSpace: 'nowrap' as const,
    paddingTop: tokens.spacingVerticalXS,
    paddingBottom: tokens.spacingVerticalXS,
    paddingLeft: tokens.spacingHorizontalS,
    paddingRight: tokens.spacingHorizontalS,
  },
  copyButton: {
    flexShrink: 0,
  },
  // ── URL / Client ID field hints (Step 3 — Connect) ───────────────────────
  // Mirrors the visual pattern of the property-pane hints:
  // a small ⓘ icon + label that reveals a Popover on hover.
  urlHintRow: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXXS,
    marginTop: tokens.spacingVerticalXXS,
    background: 'none',
    border: 'none',
    padding: 0,
    cursor: 'help',
    textAlign: 'left' as const,
    // Prevent the flex-column parent from stretching the button to full width.
    alignSelf: 'flex-start',
  },
  urlHintIcon: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorBrandForeground1,
    lineHeight: '1',
    flexShrink: 0,
  },
  urlHintLabel: {
    color: tokens.colorNeutralForeground3,
  },
  // ── PowerShell syntax-highlight token colours (fixed; always on dark bg) ──
  psOp:     { color: '#569cd6' },  // &  ::
  psType:   { color: '#4ec9b0' },  // [scriptblock]
  psMethod: { color: '#dcdcaa' },  // Create
  psCmdlet: { color: '#dcdcaa' },  // iwr
  psStr:    { color: '#ce9178' },  // 'url'
  psPunct:  { color: '#808080' },  // ( )
  // ── Action rows ────────────────────────────────────────────────────────────
  actionsSplit: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    // Wrap when the two buttons no longer fit side-by-side (very narrow zones).
    flexWrap: 'wrap' as const,
    gap: tokens.spacingHorizontalS,
    width: '100%',
  },
  actionsEnd: {
    display: 'flex',
    justifyContent: 'flex-end',
    alignItems: 'center',
    flexWrap: 'wrap' as const,
    gap: tokens.spacingHorizontalS,
    width: '100%',
  },
});

// ─────────────────────────────────────────────────────────────────────────────
// Exported types
// ─────────────────────────────────────────────────────────────────────────────

interface IWelcomeDialogProps {
  open: boolean;
  /**
   * Called when the admin commits their setup choice — on the transition into
   * the Done step. Properties are written here, before the Done screen appears,
   * so the confirmation text is truthful.
   */
  onCommit: (config: IWelcomeSetupConfig) => void;
  /**
   * Called when the user skips the wizard without completing it: X button or
   * "Not now". The host closes the dialog AND opens the property pane so the
   * admin has a direct path to configure the web part manually.
   */
  onSkip: () => void;
  /**
   * Called when the user clicks "Let's go" on the Done step after committing.
   * At this point the property pane is already open (opened by the host on
   * commit), so this callback only needs to close the dialog.
   */
  onDismiss: () => void;
  /**
   * Semver string of the web part (e.g. "0.21.0") used to build the versioned
   * "Deploy to Azure" portal URL in step 2. Falls back to the latest release
   * URL when undefined.
   */
  semver?: string;
  /** When true the wizard switches to dark-mode illustration variants. */
  isDark?: boolean;
}

// ─────────────────────────────────────────────────────────────────────────────
// Step content sub-components
// ─────────────────────────────────────────────────────────────────────────────

/** Step 1 — Welcome intro + Workoho branding. */
const Step1Welcome: React.FC<{ classes: ReturnType<typeof useStyles>; isDark?: boolean }> = ({ classes, isDark }) => (
  <>
    <div className={classes.illustrationWrap}>
      <img src={isDark ? welcomeIllustrationDark : welcomeIllustration} alt="" className={classes.illustrationImg} />
    </div>
    <Text block className={classes.body}>{strings.WelcomeDialogBody}</Text>
    <div className={classes.workohoRow}>
      <Link href="https://workoho.com?utm_source=guest-sponsor-info-webpart&utm_medium=sharepoint-webpart&utm_campaign=setup-wizard&utm_content=branding-logo" target="_blank" rel="noopener noreferrer" aria-label="Workoho">
        <img src={workohoLogo} alt="" className={classes.workohoLogo} />
      </Link>
      <Text className={classes.muted}>
        {strings.WelcomeDialogBroughtToYouBy}{' '}
        <Link href="https://workoho.com?utm_source=guest-sponsor-info-webpart&utm_medium=sharepoint-webpart&utm_campaign=setup-wizard&utm_content=branding-link" target="_blank" rel="noopener noreferrer">
          {strings.WelcomeDialogWorkohoLinkLabel}
        </Link>
      </Text>
    </div>
  </>
);

interface IStep2SetupProps {
  classes: ReturnType<typeof useStyles>;
  choice: 'api' | 'demo';
  deployToAzureUrl: string;
  semver?: string;
  onChoiceChange: (v: 'api' | 'demo') => void;
}

/** Step 2 — Setup choice: API vs. Demo Mode. */
const Step2Setup: React.FC<IStep2SetupProps> = ({
  classes, choice, deployToAzureUrl, semver, onChoiceChange,
}) => {
  // Track which command block just had its content copied to the clipboard.
  // Resets to null after a short delay so the "Copied!" tooltip disappears.
  const [copiedStep, setCopiedStep] = React.useState<1 | 3 | null>(null);

  const step1Cmd = React.useMemo(() => buildStep1Command(semver), [semver]);
  const step3Cmd = React.useMemo(() => buildStep3Command(semver), [semver]);
  const step1Tokens = React.useMemo(() => tokenizePwsh(step1Cmd), [step1Cmd]);
  const step3Tokens = React.useMemo(() => tokenizePwsh(step3Cmd), [step3Cmd]);
  // Raw script URLs — shown in the tooltip and as "View script source" links
  // so admins can inspect the PowerShell before running it.
  const step1Url = React.useMemo(() => buildScriptUrl(semver, 'setup-app-registration.ps1'), [semver]);
  const step3Url = React.useMemo(() => buildScriptUrl(semver, 'setup-graph-permissions.ps1'), [semver]);

  const handleCopy = (cmd: string, step: 1 | 3): void => {
    navigator.clipboard.writeText(cmd).then(() => {
      setCopiedStep(step);
      setTimeout(() => setCopiedStep(null), 2000);
    }).catch(() => { /* clipboard unavailable — silent */ });
  };

  const [openStep, setOpenStep] = React.useState<number | null>(null);
  const toggleStep = (n: number): void =>
    setOpenStep(prev => (prev === n ? null : n));

  return (
    <>
      <Text block className={classes.setupIntro}>{strings.WelcomeDialogSetupIntro}</Text>

      {/* ILLUSTRATION PLACEHOLDER — replace with a split-screen SVG asset (see comment in useStyles) */}
      <RadioGroup value={choice} onChange={(_, d) => onChoiceChange(d.value as 'api' | 'demo')}>
        {/* Option A: Use Guest Sponsor API */}
        <div
          role="presentation"
          className={mergeClasses(classes.optionCard, choice === 'api' && classes.optionCardSelectedTop)}
          onClick={() => onChoiceChange('api')}
        >
          <CloudRegular style={{ width: 24, height: 24 }} className={classes.optionIcon} />
          <div className={classes.optionText}>
            <Radio value="api" label={
              <>
                <Text weight="semibold">{strings.WelcomeDialogOptionApiTitle}</Text>
                <Text size={200} block className={classes.muted}>{strings.WelcomeDialogOptionApiBody}</Text>
              </>
            } />
          </div>
        </div>

        {/* Deploy panel — slides out from under the API option card when selected */}
        <div className={mergeClasses(classes.deployPanel, choice === 'api' && classes.deployPanelExpanded)}>
          {/* Setup guide note — shown first so admins read it before running scripts */}
          <Text size={100} block className={classes.muted}>{strings.WelcomeDialogSetupGuideNote}</Text>
          <Link
            href={GITHUB_SETUP_URL}
            target="_blank"
            rel="noopener noreferrer"
            className={classes.docsLink}
          >
            {strings.WelcomeDialogOptionApiDocsLabel}
          </Link>
          <Text size={200} block className={classes.muted}>{strings.WelcomeDialogDeployNote}</Text>
          <Text size={100} weight="semibold" block className={classes.muted}>{strings.WelcomeDialogSetupStepsOrderHint}</Text>

          {/* ① Create App Registration */}
          <button
            type="button"
            className={classes.stepToggle}
            onClick={() => toggleStep(1)}
            aria-expanded={openStep === 1}
          >
            <span className={classes.stepNumber}>1</span>
            <Text size={200} weight="semibold" className={classes.stepToggleLabel}>
              {strings.WelcomeDialogSetupStep1Label}
            </Text>
            <ChevronRightRegular className={mergeClasses(classes.stepChevron, openStep === 1 && classes.stepChevronOpen)} />
          </button>
          <div className={mergeClasses(classes.stepContent, openStep === 1 && classes.stepContentOpen)}>
            <div className={classes.codeWrap}>
              <div className={classes.codeLangBar}>
                <span className={classes.codeLangLabel}>PowerShell</span>
                <div className={classes.codeLangActions}>
                  <Button
                    as="a"
                    href={step1Url}
                    target="_blank"
                    rel="noopener noreferrer"
                    appearance="transparent"
                    size="small"
                    icon={<CodeRegular />}
                    style={{ color: '#e0e0e0' }}
                  >
                    {strings.ViewScriptSourceLabel}
                  </Button>
                  <Button
                    appearance="transparent"
                    size="small"
                    icon={copiedStep === 1 ? <CheckmarkRegular /> : <CopyRegular />}
                    className={classes.copyButton}
                    style={{ color: '#e0e0e0' }}
                    onClick={() => handleCopy(step1Cmd, 1)}
                  >
                    {copiedStep === 1 ? strings.CopiedToClipboardLabel : strings.CopyToClipboardLabel}
                  </Button>
                </div>
              </div>
              <code className={classes.codeBlock}>{step1Tokens.map((tok, idx) => (
                <span key={idx} className={
                  tok.t === 'op'     ? classes.psOp :
                  tok.t === 'type'   ? classes.psType :
                  tok.t === 'method' ? classes.psMethod :
                  tok.t === 'cmdlet' ? classes.psCmdlet :
                  tok.t === 'str'    ? classes.psStr :
                  tok.t === 'punct'  ? classes.psPunct :
                  undefined
                }>{tok.v}</span>
              ))}</code>
            </div>
          </div>

          {/* ② Deploy to Azure */}
          <button
            type="button"
            className={classes.stepToggle}
            onClick={() => toggleStep(2)}
            aria-expanded={openStep === 2}
          >
            <span className={classes.stepNumber}>2</span>
            <Text size={200} weight="semibold" className={classes.stepToggleLabel}>
              {strings.WelcomeDialogDeployToAzureLabel}
            </Text>
            <ChevronRightRegular className={mergeClasses(classes.stepChevron, openStep === 2 && classes.stepChevronOpen)} />
          </button>
          <div className={mergeClasses(classes.stepContent, openStep === 2 && classes.stepContentOpen)}>
            <Text size={100} block className={classes.muted}>{strings.WelcomeDialogSetupStep2Hint}</Text>
            <Text size={100} block className={classes.muted}>{strings.DeployToAzureClickHint}</Text>
            <Link href={deployToAzureUrl} target="_blank" rel="noopener noreferrer" style={{ display: 'inline-flex' }}>
              <img
                src="https://aka.ms/deploytoazurebutton"
                alt={strings.WelcomeDialogDeployToAzureLabel}
                style={{ display: 'block', maxWidth: '100%' }}
              />
            </Link>
          </div>

          {/* ③ Grant Graph permissions */}
          <button
            type="button"
            className={classes.stepToggle}
            onClick={() => toggleStep(3)}
            aria-expanded={openStep === 3}
          >
            <span className={classes.stepNumber}>3</span>
            <Text size={200} weight="semibold" className={classes.stepToggleLabel}>
              {strings.WelcomeDialogSetupStep3Label}
            </Text>
            <ChevronRightRegular className={mergeClasses(classes.stepChevron, openStep === 3 && classes.stepChevronOpen)} />
          </button>
          <div className={mergeClasses(classes.stepContent, openStep === 3 && classes.stepContentOpen)}>
            <Text size={100} block className={classes.muted}>{strings.WelcomeDialogSetupStep3Hint}</Text>
            <div className={classes.codeWrap}>
              <div className={classes.codeLangBar}>
                <span className={classes.codeLangLabel}>PowerShell</span>
                <div className={classes.codeLangActions}>
                  <Button
                    as="a"
                    href={step3Url}
                    target="_blank"
                    rel="noopener noreferrer"
                    appearance="transparent"
                    size="small"
                    icon={<CodeRegular />}
                    style={{ color: '#e0e0e0' }}
                  >
                    {strings.ViewScriptSourceLabel}
                  </Button>
                  <Button
                    appearance="transparent"
                    size="small"
                    icon={copiedStep === 3 ? <CheckmarkRegular /> : <CopyRegular />}
                    className={classes.copyButton}
                    style={{ color: '#e0e0e0' }}
                    onClick={() => handleCopy(step3Cmd, 3)}
                  >
                    {copiedStep === 3 ? strings.CopiedToClipboardLabel : strings.CopyToClipboardLabel}
                  </Button>
                </div>
              </div>
              <code className={classes.codeBlock}>{step3Tokens.map((tok, idx) => (
                <span key={idx} className={
                  tok.t === 'op'     ? classes.psOp :
                  tok.t === 'type'   ? classes.psType :
                  tok.t === 'method' ? classes.psMethod :
                  tok.t === 'cmdlet' ? classes.psCmdlet :
                  tok.t === 'str'    ? classes.psStr :
                  tok.t === 'punct'  ? classes.psPunct :
                  undefined
                }>{tok.v}</span>
              ))}</code>
            </div>
          </div>
        </div>

        {/* Option B: Demo Mode */}
        <div
          role="presentation"
          className={mergeClasses(classes.optionCard, choice === 'demo' && classes.optionCardSelected)}
          onClick={() => onChoiceChange('demo')}
        >
          <BeakerRegular style={{ width: 24, height: 24 }} className={classes.optionIcon} />
          <div className={classes.optionText}>
            <Radio value="demo" label={
              <>
                <Text weight="semibold">{strings.WelcomeDialogOptionDemoTitle}</Text>
                <Text size={200} block className={classes.muted}>{strings.WelcomeDialogOptionDemoBody}</Text>
              </>
            } />
          </div>
        </div>
      </RadioGroup>
    </>
  );
};

interface IStep3ConnectProps {
  classes: ReturnType<typeof useStyles>;
  isDark?: boolean;
  apiUrl: string;
  clientId: string;
  urlError: string;
  clientIdError: string;
  onApiUrlChange: (v: string) => void;
  onClientIdChange: (v: string) => void;
}

/** Step 3 — Connect Guest Sponsor API (optional credentials, API path only). */
const Step3Connect: React.FC<IStep3ConnectProps> = ({
  classes, isDark, apiUrl, clientId, urlError, clientIdError, onApiUrlChange, onClientIdChange,
}) => {
  const [urlHintOpen, setUrlHintOpen] = React.useState(false);
  const [clientIdHintOpen, setClientIdHintOpen] = React.useState(false);
  return (
    <>
      <div className={classes.illustrationWrap}>
        <img src={isDark ? wizardConnectIllustrationDark : wizardConnectIllustration} alt="" className={classes.illustrationImg} />
      </div>
      <Text block className={classes.setupIntro}>{strings.WelcomeDialogConnectApiIntro}</Text>
      <div className={classes.apiFields}>
        <Field
          label={strings.FunctionUrlFieldLabel}
          validationMessage={urlError || undefined}
          validationState={urlError ? 'error' : 'none'}
        >
          <Input
            value={apiUrl}
            onChange={(_, d) => onApiUrlChange(d.value)}
            placeholder="https://my-app.azurewebsites.net"
            type="url"
          />
        </Field>
        <Popover open={urlHintOpen} onOpenChange={(_, d) => setUrlHintOpen(d.open)}>
          <PopoverTrigger disableButtonEnhancement>
            <button
              type="button"
              className={classes.urlHintRow}
              onMouseEnter={() => setUrlHintOpen(true)}
              onMouseLeave={() => setUrlHintOpen(false)}
            >
              <span className={classes.urlHintIcon}>ⓘ</span>
              <Text size={100} className={classes.urlHintLabel}>{strings.WelcomeDialogUrlHintLabel}</Text>
            </button>
          </PopoverTrigger>
          <PopoverSurface
            onMouseEnter={() => setUrlHintOpen(true)}
            onMouseLeave={() => setUrlHintOpen(false)}
            style={{ maxWidth: '280px' }}
          >
            <Text size={100} block>{strings.WelcomeDialogUrlHintBody}</Text>
          </PopoverSurface>
        </Popover>
        <Field
          label={strings.FunctionClientIdFieldLabel}
          validationMessage={clientIdError || undefined}
          validationState={clientIdError ? 'error' : 'none'}
        >
          <Input
            value={clientId}
            onChange={(_, d) => onClientIdChange(d.value)}
            placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
          />
        </Field>
        <Popover open={clientIdHintOpen} onOpenChange={(_, d) => setClientIdHintOpen(d.open)}>
          <PopoverTrigger disableButtonEnhancement>
            <button
              type="button"
              className={classes.urlHintRow}
              onMouseEnter={() => setClientIdHintOpen(true)}
              onMouseLeave={() => setClientIdHintOpen(false)}
            >
              <span className={classes.urlHintIcon}>ⓘ</span>
              <Text size={100} className={classes.urlHintLabel}>{strings.PpClientIdHintLabel}</Text>
            </button>
          </PopoverTrigger>
          <PopoverSurface
            onMouseEnter={() => setClientIdHintOpen(true)}
            onMouseLeave={() => setClientIdHintOpen(false)}
            style={{ maxWidth: '280px' }}
          >
            <Text size={100} block>{strings.PpClientIdHintBody}</Text>
            <Text size={100} block style={{ fontWeight: 600, marginTop: '4px' }}>
              {'"Guest Sponsor Info \u2013 SharePoint Web Part Auth"'}
            </Text>
            <Link
              href="https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade"
              target="_blank"
              rel="noopener noreferrer"
              style={{ fontSize: tokens.fontSizeBase100, display: 'block', marginTop: '4px' }}
            >
              {strings.FunctionEntraLinkLabel}
            </Link>
          </PopoverSurface>
        </Popover>
      </div>
    </>
  );
};

/** Step 4 — Confirmation (content differs by chosen path and whether credentials were skipped). */
const Step4Done: React.FC<{
  classes: ReturnType<typeof useStyles>;
  isDark?: boolean;
  choice: 'api' | 'demo';
  skippedCredentials?: boolean;
}> = ({ classes, isDark, choice, skippedCredentials }) => (
  <>
    <div className={classes.illustrationWrap}>
      <img src={isDark ? wizardSuccessIllustrationDark : wizardSuccessIllustration} alt="" className={classes.illustrationImg} />
    </div>
    <Text block className={classes.body}>
      {choice === 'api'
        ? (skippedCredentials ? strings.WelcomeDialogDoneApiSkippedBody : strings.WelcomeDialogDoneApiBody)
        : strings.WelcomeDialogDoneDemoBody}
    </Text>
  </>
);

// ─────────────────────────────────────────────────────────────────────────────
// Main dialog component
// ─────────────────────────────────────────────────────────────────────────────

const WelcomeDialog: React.FC<IWelcomeDialogProps> = ({ open, onCommit, onSkip, onDismiss, semver, isDark }) => {
  const classes = useStyles();
  // Card element used as the mount node for Tooltip portals so they render
  // inside the FluentProvider's DOM subtree and inherit CSS token variables.
  // PortalMountNodeContextValue accepts HTMLElement | undefined, not null.
  const [cardMountNode, setCardMountNode] = React.useState<HTMLElement | undefined>(undefined);
  const cardRef = React.useCallback((node: HTMLDivElement | null) => {
    setCardMountNode(node ?? undefined);
  }, []);
  const [step, setStep] = React.useState(0);
  const [choice, setChoice] = React.useState<'api' | 'demo'>('demo');
  const [apiUrl, setApiUrl] = React.useState('');
  const [clientId, setClientId] = React.useState('');
  // True once onCommit has been called — the Done step is a genuine confirmation.
  const [committed, setCommitted] = React.useState(false);
  // True when the user reached Done by skipping the API credentials step.
  const [skippedApiCreds, setSkippedApiCreds] = React.useState(false);

  // Build once — semver is stable for the lifetime of the wizard.
  const deployToAzureUrl = buildDeployToAzureUrl(semver);

  // Derived validation state — computed from current field values so errors
  // appear immediately as the user types, with no extra state needed.
  const trimmedUrl = apiUrl.trim();
  const trimmedClientId = clientId.trim();
  const bothEmpty = trimmedUrl === '' && trimmedClientId === '';
  // Show an error only when the field has content but the format is wrong.
  const urlError = trimmedUrl !== '' && !isValidFunctionUrl(trimmedUrl)
    ? strings.InvalidUrlFormat
    : '';
  const clientIdError = trimmedClientId !== '' && !isValidGuid(trimmedClientId)
    ? strings.InvalidGuidFormat
    : '';
  // Save is only available when something is entered AND all entered values pass.
  const canSave = !bothEmpty && !urlError && !clientIdError;

  // Reset wizard state whenever the dialog is (re-)opened.
  React.useEffect(() => {
    if (open) {
      setStep(0);
      setChoice('demo');
      setApiUrl('');
      setClientId('');
      setCommitted(false);
      setSkippedApiCreds(false);
    }
  }, [open]);

  const handleNext = (): void => {
    if (step === 1) {
      if (choice === 'demo') {
        // Demo: commit now, then show the Done step — properties are written
        // before the confirmation screen appears, so the text is accurate.
        onCommit({ chosenPath: 'demo' });
        setCommitted(true);
        setStep(3);
      } else {
        // API: go to connection details step first.
        setStep(2);
      }
      return;
    }
    if (step === 2) {
      // API Connect: commit now, then show Done.
      // Record whether the user skipped entering credentials so the Done step
      // can show an appropriate message.
      setSkippedApiCreds(bothEmpty);
      onCommit({
        chosenPath: 'api',
        apiUrl: apiUrl.trim(),
        clientId: clientId.trim(),
      });
      setCommitted(true);
      setStep(3);
      return;
    }
    setStep(s => s + 1);
  };

  const handleBack = (): void => {
    // Committed state cannot be undone from within the wizard.
    if (committed) return;
    setStep(s => s - 1);
  };

  // "Let's go" on the Done step — just closes. Data was already committed.
  const handleDismiss = (): void => onDismiss();

  // "Not now" or X — skip without committing; host also opens the property pane.
  const handleSkip = (): void => onSkip();

  const stepTitle =
    step === 0 ? strings.WelcomeDialogTitle :
    step === 1 ? strings.WelcomeDialogSetupTitle :
    step === 2 ? strings.WelcomeDialogConnectApiTitle :
    choice === 'api'
    ? (skippedApiCreds ? strings.WelcomeDialogDoneApiSkippedTitle : strings.WelcomeDialogDoneApiTitle)
    : strings.WelcomeDialogDoneDemoTitle;

  if (!open) return null;

  return (
    <PortalMountNodeProvider value={cardMountNode}>
    <div className={classes.root}>
      <div ref={cardRef} className={classes.card}>
        {/* Card header: wizard title + dismiss (X) button on steps 0-2 */}
        <div className={classes.cardHeader}>
          <Text as="h2" className={classes.wizardTitle}>{stepTitle}</Text>
          {step < 3 && (
            <Button
              appearance="subtle"
              icon={<DismissRegular />}
              aria-label={strings.WelcomeDialogSkipButton}
              onClick={handleSkip}
              style={{ flexShrink: 0, marginTop: '-4px', marginRight: '-8px' }}
            />
          )}
        </div>

        {/* Progress dots — 4 for API path, 3 for Demo path */}
        <div className={classes.stepDots} role="group" aria-label="Step progress">
          {Array.from({ length: choice === 'api' ? 4 : 3 }, (_, i) => {
            const dotStep = choice === 'demo' && step === 3 ? 2 : step;
            return (
              <span
                key={i}
                aria-hidden="true"
                className={mergeClasses(
                  classes.dot,
                  i === dotStep && classes.dotActive,
                  i < dotStep && classes.dotDone,
                )}
              />
            );
          })}
        </div>

        {step === 0 && <Step1Welcome classes={classes} isDark={isDark} />}
        {step === 1 && (
          <Step2Setup
            classes={classes}
            choice={choice}
            deployToAzureUrl={deployToAzureUrl}
            semver={semver}
            onChoiceChange={setChoice}
          />
        )}
        {step === 2 && (
          <Step3Connect
            classes={classes}
            isDark={isDark}
            apiUrl={apiUrl}
            clientId={clientId}
            urlError={urlError}
            clientIdError={clientIdError}
            onApiUrlChange={setApiUrl}
            onClientIdChange={setClientId}
          />
        )}
        {step === 3 && <Step4Done classes={classes} isDark={isDark} choice={choice} skippedCredentials={skippedApiCreds} />}

        <div className={classes.stepActions}>
          {step === 0 && (
            <div className={classes.actionsEnd}>
              <Button appearance="secondary" onClick={handleSkip}>
                {strings.WelcomeDialogSkipButton}
              </Button>
              <Button appearance="primary" icon={<ChevronRightRegular />} iconPosition="after" onClick={handleNext}>
                {strings.WelcomeDialogNextButton}
              </Button>
            </div>
          )}
          {step === 1 && (
            <div className={classes.actionsSplit}>
              <Button appearance="secondary" icon={<ChevronLeftRegular />} onClick={handleBack}>
                {strings.WelcomeDialogBackButton}
              </Button>
              <Button
                appearance="primary"
                icon={choice === 'demo' ? <CheckmarkRegular /> : <ChevronRightRegular />}
                iconPosition="after"
                onClick={handleNext}
              >
                {choice === 'demo' ? strings.WelcomeDialogSaveButton : strings.WelcomeDialogNextButton}
              </Button>
            </div>
          )}
          {step === 2 && (
            <div className={classes.actionsSplit}>
              <Button appearance="secondary" icon={<ChevronLeftRegular />} onClick={handleBack}>
                {strings.WelcomeDialogBackButton}
              </Button>
              {bothEmpty ? (
                // Both fields empty → advance to the Done step (committing with
                // empty credentials) so the user still sees the confirmation screen.
                <Button appearance="primary" icon={<ChevronRightRegular />} iconPosition="after" onClick={handleNext}>
                  {strings.WelcomeDialogSkipApiButton}
                </Button>
              ) : (
                // Something was entered → show Save, disabled until both
                // non-empty fields pass format validation.
                <Button
                  appearance="primary"
                  icon={<CheckmarkRegular />}
                  iconPosition="after"
                  onClick={handleNext}
                  disabled={!canSave}
                >
                  {strings.WelcomeDialogSaveButton}
                </Button>
              )}
            </div>
          )}
          {step === 3 && (
            <div className={classes.actionsEnd}>
              <Button appearance="primary" icon={<ChevronRightRegular />} iconPosition="after" onClick={handleDismiss}>
                {strings.WelcomeDialogDismissButton}
              </Button>
            </div>
          )}
        </div>
      </div>
    </div>
    </PortalMountNodeProvider>
  );
};

export default WelcomeDialog;
