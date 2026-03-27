// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0

import * as React from 'react';
import {
  Button,
  Field,
  Input,
  Link,
  Radio,
  RadioGroup,
  Text,
  makeStyles,
  mergeClasses,
  tokens,
} from '@fluentui/react-components';
import {
  BeakerRegular,
  CheckmarkCircleRegular,
  CheckmarkRegular,
  ChevronLeftRegular,
  ChevronRightRegular,
  CloudRegular,
  DismissRegular,
  PeopleTeamRegular,
} from '@fluentui/react-icons';
import * as strings from 'GuestSponsorInfoWebPartStrings';
import type { IWelcomeSetupConfig } from './IGuestSponsorInfoProps';
import { isValidFunctionUrl, isValidGuid } from '../utils/fieldValidation';
import workohoLogo from '../assets/workoho-default-logo.svg';

/**
 * URL of the Azure Function deployment guide on GitHub.
 * Shown as a help link in step 2 when the admin chooses the API path.
 */
const GITHUB_SETUP_URL = 'https://github.com/workoho/spfx-guest-sponsor-info/blob/main/docs/deployment.md';

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
    marginTop: tokens.spacingVerticalL,
    paddingTop: tokens.spacingVerticalM,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  workohoLogo: {
    height: '20px',
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
    maxHeight: '300px',
    opacity: 1,
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalM,
  },
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
}

// ─────────────────────────────────────────────────────────────────────────────
// Step content sub-components
// ─────────────────────────────────────────────────────────────────────────────

/** Step 1 — Welcome intro + Workoho branding. */
const Step1Welcome: React.FC<{ classes: ReturnType<typeof useStyles> }> = ({ classes }) => (
  <>
    {/* ILLUSTRATION PLACEHOLDER — replace <div> with <img src={yourAsset} alt="" /> */}
    <div className={classes.illustrationWrap}>
      <div className={classes.illustrationCircle}>
        <PeopleTeamRegular style={{ width: 40, height: 40 }} className={classes.illustrationIcon} />
      </div>
    </div>
    <Text block className={classes.body}>{strings.WelcomeDialogBody}</Text>
    <div className={classes.workohoRow}>
      <img src={workohoLogo} alt="Workoho" className={classes.workohoLogo} />
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
  onChoiceChange: (v: 'api' | 'demo') => void;
}

/** Step 2 — Setup choice: API vs. Demo Mode. */
const Step2Setup: React.FC<IStep2SetupProps> = ({
  classes, choice, deployToAzureUrl, onChoiceChange,
}) => (
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
        <Text size={200} block className={classes.muted}>{strings.WelcomeDialogDeployNote}</Text>
        <Link href={deployToAzureUrl} target="_blank" rel="noopener noreferrer">
          <img
            src="https://aka.ms/deploytoazurebutton"
            alt={strings.WelcomeDialogDeployToAzureLabel}
            style={{ display: 'block', maxWidth: '100%' }}
          />
        </Link>
        <Link
          href={GITHUB_SETUP_URL}
          target="_blank"
          rel="noopener noreferrer"
          className={classes.docsLink}
        >
          {strings.WelcomeDialogOptionApiDocsLabel}
        </Link>
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

interface IStep3ConnectProps {
  classes: ReturnType<typeof useStyles>;
  apiUrl: string;
  clientId: string;
  urlError: string;
  clientIdError: string;
  onApiUrlChange: (v: string) => void;
  onClientIdChange: (v: string) => void;
}

/** Step 3 — Connect Guest Sponsor API (optional credentials, API path only). */
const Step3Connect: React.FC<IStep3ConnectProps> = ({
  classes, apiUrl, clientId, urlError, clientIdError, onApiUrlChange, onClientIdChange,
}) => (
  <>
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
    </div>
  </>
);

/** Step 4 — Confirmation (content differs by chosen path and whether credentials were skipped). */
const Step4Done: React.FC<{
  classes: ReturnType<typeof useStyles>;
  choice: 'api' | 'demo';
  skippedCredentials?: boolean;
}> = ({ classes, choice, skippedCredentials }) => (
  <>
    {/* ILLUSTRATION PLACEHOLDER — replace <div> with <img src={yourSuccessAsset} alt="" /> */}
    <div className={classes.illustrationWrap}>
      <div className={classes.illustrationCircleSuccess}>
        <CheckmarkCircleRegular style={{ width: 40, height: 40 }} className={classes.illustrationIconSuccess} />
      </div>
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

const WelcomeDialog: React.FC<IWelcomeDialogProps> = ({ open, onCommit, onSkip, onDismiss, semver }) => {
  const classes = useStyles();
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
    <div className={classes.root}>
      <div className={classes.card}>
        {/* Card header: wizard title + dismiss (X) button on steps 0–2 */}
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

        {step === 0 && <Step1Welcome classes={classes} />}
        {step === 1 && (
          <Step2Setup
            classes={classes}
            choice={choice}
            deployToAzureUrl={deployToAzureUrl}
            onChoiceChange={setChoice}
          />
        )}
        {step === 2 && (
          <Step3Connect
            classes={classes}
            apiUrl={apiUrl}
            clientId={clientId}
            urlError={urlError}
            clientIdError={clientIdError}
            onApiUrlChange={setApiUrl}
            onClientIdChange={setClientId}
          />
        )}
        {step === 3 && <Step4Done classes={classes} choice={choice} skippedCredentials={skippedApiCreds} />}

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
  );
};

export default WelcomeDialog;
