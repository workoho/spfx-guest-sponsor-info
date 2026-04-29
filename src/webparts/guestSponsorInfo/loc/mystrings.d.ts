declare interface IGuestSponsorInfoWebPartStrings {
  BasicGroupName: string;
  TitleFieldLabel: string;
  ShowTitleFieldLabel: string;
  /** Gray placeholder text shown in the title field when no title has been entered yet. */
  TitlePlaceholder: string;
  /** Property pane label for the title size choice group. */
  TitleSizeFieldLabel: string;
  TitleSizeH2Option: string;
  TitleSizeH3Option: string;
  TitleSizeH4Option: string;
  TitleSizeNormalOption: string;
  LoadingMessage: string;
  NoSponsorsTitle: string;
  NoSponsorsMessage: string;
  SponsorUnavailableTitle: string;
  SponsorUnavailableMessage: string;
  ErrorMessageTitle: string;
  ErrorMessage: string;
  EditModePlaceholder: string;
  GuestOnlyPlaceholder: string;
  MockModeFieldLabel: string;
  /** Tooltip text explaining the public demo mode functionality. */
  MockModeFieldTooltip: string;
  MockModeFieldDescription: string;
  MockModePlaceholder: string;
  MockSimulatedHintFieldLabel: string;
  MockSimulatedHintNoneOption: string;
  MockSimulatedHintTeamsAccessPendingOption: string;
  MockSimulatedHintVersionMismatchOption: string;
  MockSimulatedHintSponsorUnavailableOption: string;
  MockSimulatedHintNoSponsorsOption: string;
  /** Slider label for the visible sponsor count. */
  MaxSponsorCountFieldLabel: string;
  /** Tooltip text explaining the sponsor fallback/rotation logic. */
  MaxSponsorCountFieldTooltip: string;

  // Display toggles
  GuestNotificationsGroupName: string;
  GuestNotificationsGroupHint: string;
  ShowTeamsAccessPendingHintLabel: string;
  ShowVersionMismatchHintLabel: string;
  ShowSponsorUnavailableHintLabel: string;
  ShowNoSponsorsHintLabel: string;
  DisplayGroupName: string;
  CardLayoutFieldLabel: string;
  CardLayoutAutoOption: string;
  CardLayoutAutoThresholdFieldLabel: string;
  CardLayoutFullOption: string;
  CardLayoutCompactOption: string;
  ShowBusinessPhonesFieldLabel: string;
  ShowMobilePhoneFieldLabel: string;
  ShowWorkLocationFieldLabel: string;
  ShowCityFieldLabel: string;
  ShowCountryFieldLabel: string;
  ShowStreetAddressFieldLabel: string;
  ShowPostalCodeFieldLabel: string;
  ShowStateFieldLabel: string;
  AzureMapsSubscriptionKeyFieldLabel: string;
  MapProviderModeAutoLabel: string;
  MapProviderModeManualLabel: string;
  MapProviderModeHint: string;
  MapProviderIosLabel: string;
  MapProviderAndroidLabel: string;
  MapProviderWindowsLabel: string;
  MapProviderMacOSLabel: string;
  MapProviderLinuxLabel: string;
  ExternalMapProviderFieldLabel: string;
  MapProviderBingOption: string;
  MapProviderGoogleOption: string;
  MapProviderAppleOption: string;
  MapProviderOpenStreetMapOption: string;
  MapProviderNoneOption: string;
  AddressMapProviderHint: string;
  AzureMapsPreviewHint: string;
  LocationDisplayHintSeparateRows: string;
  AdvancedDisplayGroupName: string;
  ShowManagerFieldLabel: string;
  ShowPresenceFieldLabel: string;
  ShowSponsorPhotoFieldLabel: string;
  ShowManagerPhotoFieldLabel: string;
  UseInformalAddressFieldLabel: string;

  // Sponsor eligibility filter
  SponsorEligibilityGroupName: string;
  SponsorEligibilityGroupHint: string;
  SponsorFilterFieldLabel: string;
  SponsorFilterTeamsOption: string;
  SponsorFilterExchangeOption: string;
  SponsorFilterAnyOption: string;
  RequireUserMailboxFieldLabel: string;

  // Property pane sub-section header labels
  PpSectionLivePage: string;
  PpSectionDemoMode: string;
  PpSectionCardLayout: string;
  PpSectionProfileFields: string;
  PpSectionPhone: string;
  PpSectionWorkLocation: string;
  PpSectionAddress: string;
  PpSectionMapLink: string;
  PpSectionMapPreview: string;

  // Job title and department toggles
  ShowSponsorJobTitleFieldLabel: string;
  ShowManagerJobTitleFieldLabel: string;
  ShowSponsorDepartmentFieldLabel: string;
  ShowManagerDepartmentFieldLabel: string;
  DepartmentLabel: string;

  // Azure Function proxy
  FunctionGroupName: string;
  FunctionUrlFieldLabel: string;
  WebPartClientIdFieldLabel: string;
  SessionCacheTtlMinutesFieldLabel: string;
  /** Short trigger label for the Client ID tooltip. */
  PpClientIdHintLabel: string;
  /** Body text inside the Client ID hover popup. */
  PpClientIdHintBody: string;
  FunctionEntraLinkLabel: string;
  ProxyStatusChecking: string;
  ProxyStatusOk: string;
  ProxyStatusError: string;
  /** Short intro shown above the 3-step setup guide in the property pane. */
  PpSetupIntro: string;
  /** Troubleshooting hint shown in the property pane when the proxy status is 'error'. */
  PpSetupRerunHint: string;

  // Clipboard copy feedback (shared)
  CopyToClipboardLabel: string;
  CopiedToClipboardLabel: string;

  // Author/company information section
  AuthorSectionGroupName: string;
  AuthorSectionCreatedBy: string;
  AuthorSectionIntro: string;
  AuthorSectionWebsiteLinkLabel: string;
  AuthorSectionConsultingText: string;
  AuthorSectionAppWebsiteLabel: string;
  AuthorSectionVersionLabel: string;
  AuthorSectionDeployToAzureLabel: string;
  DeployToAzureClickHint: string;
  AuthorSectionDeploymentGuideLabel: string;
  NewReleaseAvailableLabel: string;

  // Guest Accessibility diagnostics group
  DiagnosticsGroupName: string;
  DiagnosticsGroupHint: string;
  DiagnosticsChecking: string;
  DiagnosticsRetryButtonLabel: string;
  DiagPublicCdnOk: string;
  DiagPrivateCdnInfo: string;
  DiagSiteColCatalogOk: string;
  DiagTenantCatalogEveryoneOk: string;
  DiagTenantCatalogEveryoneDenied: string;
  DiagTenantCatalogEveryoneUnknown: string;
  DiagAssetSourceUnknown: string;
  DiagSiteEveryoneOk: string;
  DiagSiteEveryoneMissing: string;
  DiagSiteEveryoneUnknown: string;

  // Presence labels
  PresenceAvailable: string;
  PresenceAvailableIdle: string;
  PresenceAway: string;
  PresenceBeRightBack: string;
  PresenceBusy: string;
  PresenceBusyIdle: string;
  PresenceDoNotDisturb: string;
  PresenceInAMeeting: string;
  PresenceInACall: string;
  PresencePresenting: string;
  PresenceFocusing: string;
  PresenceOutOfOffice: string;
  PresenceOutOfOfficeSuffix: string;
  PresenceOffline: string;

  // Rich contact card
  ContactDetailsAriaLabel: string;  // contains "{0}" placeholder for the name
  ContactActionsAriaLabel: string;
  CloseLabel: string;
  ChatTitle: string;
  EmailTitle: string;
  CallTitle: string;
  ChatLabel: string;
  EmailLabel: string;
  CallLabel: string;

  // Contact information section
  ContactInfoSection: string;
  EmailFieldLabel: string;
  WorkPhoneFieldLabel: string;
  MobileFieldLabel: string;
  WorkLocationFieldLabel: string;
  AddressFieldLabel: string;
  CityFieldLabel: string;
  CountryFieldLabel: string;
  StreetAddressFieldLabel: string;
  PostalCodeFieldLabel: string;
  StateFieldLabel: string;

  // Organization section
  OrganizationSection: string;
  ReportsToSection: string;
  ManagerLabel: string;

  // Copy-to-clipboard buttons
  CopiedFeedback: string;
  CopyEmailAriaLabel: string;
  CopyWorkPhoneAriaLabel: string;
  CopyMobileAriaLabel: string;
  CopyLocationAriaLabel: string;
  CopyAddressAriaLabel: string;
  CopyCityAriaLabel: string;
  CopyCountryAriaLabel: string;
  CopyStreetAddressAriaLabel: string;
  CopyPostalCodeAriaLabel: string;
  CopyStateAriaLabel: string;
  AddressMapSectionLabel: string;
  AddressMapLoadingLabel: string;
  OpenAddressInMapLabel: string;

  // Teams access pending notice
  TeamsAccessPendingTitle: string;
  TeamsAccessPendingMessage: string;
  TeamsNotReadyChatTooltip: string;
  TeamsNotReadyCallTooltip: string;

  // Version mismatch notice (shown when the web part and Azure Function run different versions)
  VersionMismatchTitle: string;
  VersionMismatchMessage: string;

  // Shown when the Azure Function managed identity is missing required Graph permissions
  InsufficientPermissionsTitle: string;
  InsufficientPermissionsMessage: string;

  // Shown when the sponsor service hits a temporary internal authorization/token state issue
  ServiceAuthorizationIssueTitle: string;
  ServiceAuthorizationIssueMessage: string;

  // First-run welcome dialog (shown once per instance in edit mode)
  WelcomeDialogTitle: string;
  WelcomeDialogBody: string;
  WelcomeDialogBroughtToYouBy: string;
  WelcomeDialogWorkohoLinkLabel: string;
  WelcomeDialogDismissButton: string;
  // Wizard navigation
  WelcomeDialogNextButton: string;
  WelcomeDialogBackButton: string;
  WelcomeDialogSkipButton: string;
  /** Commits the chosen setup path and advances to the Done step. */
  WelcomeDialogSaveButton: string;
  /** Skips the API connect step when both credential fields are empty. */
  WelcomeDialogSkipApiButton: string;
  // Field validation errors (shared between wizard and property pane)
  InvalidUrlFormat: string;
  InvalidGuidFormat: string;
  // Step 2 — setup choice
  WelcomeDialogSetupTitle: string;
  WelcomeDialogSetupIntro: string;
  WelcomeDialogOptionApiTitle: string;
  WelcomeDialogOptionApiBody: string;
  WelcomeDialogOptionApiDocsLabel: string;
  /** Compact preamble shown at the top of the quick-start setup panel, before the 3 steps. */
  WelcomeDialogSetupGuideNote: string;
  WelcomeDialogDeployToAzureLabel: string;
  WelcomeDialogDeployNote: string;
  /** Compact heading for the setup preflight checklist in the wizard. */
  WelcomeDialogPreflightTitle: string;
  /** Checklist item describing automatic tool checks/installation. */
  WelcomeDialogPreflightTools: string;
  /** Checklist item describing Azure sign-in/subscription selection. */
  WelcomeDialogPreflightSignIn: string;
  /** Checklist item describing role checks. */
  WelcomeDialogPreflightRoles: string;
  /** Instruction text above the 3-step accordion (e.g. "Complete the steps in order:"). */
  WelcomeDialogSetupStepsOrderHint: string;
  /** Step label for "Create App Registration" in the 3-step setup guide. */
  WelcomeDialogSetupStep1Label: string;
  /** Hint shown below the step-2 accordion header (e.g. "Uses the App Registration created in step 1"). */
  WelcomeDialogSetupStep2Hint: string;
  /** Step label for "Grant Graph permissions" in the 3-step setup guide. */
  WelcomeDialogSetupStep3Label: string;
  /** Short hint shown above step 1 and 3 code blocks (e.g. "PowerShell:"). */
  WelcomeDialogSetupPwshHint: string;
  /** Code block label for the macOS/Linux shell bootstrap command. */
  WelcomeDialogSetupShellHint: string;
  /** Toggle label shown when the PowerShell command is currently visible. */
  WelcomeDialogUseShellCommandLabel: string;
  /** Toggle label shown when the shell command is currently visible. */
  WelcomeDialogUsePowerShellCommandLabel: string;
  /** Hint shown below step 3 code block (e.g. "Uses the App Registration (step 1) and the Function App URL (step 2)"). */
  WelcomeDialogSetupStep3Hint: string;
  /** Button label for viewing the raw PowerShell script source before running it. */
  ViewScriptSourceLabel: string;
  WelcomeDialogConnectApiTitle: string;
  WelcomeDialogConnectApiIntro: string;
  /** Trigger label for the URL hint Popover below the Function URL field. */
  WelcomeDialogUrlHintLabel: string;
  /** Body text inside the URL hint Popover explaining where to find the Function App URL. */
  WelcomeDialogUrlHintBody: string;
  WelcomeDialogOptionDemoTitle: string;
  WelcomeDialogOptionDemoBody: string;
  WelcomeDialogFunctionUrlRequired: string;
  // Step 3 — confirmation
  WelcomeDialogDoneApiTitle: string;
  WelcomeDialogDoneApiBody: string;
  WelcomeDialogDoneApiSkippedTitle: string;
  WelcomeDialogDoneApiSkippedBody: string;
  WelcomeDialogDoneDemoTitle: string;
  WelcomeDialogDoneDemoBody: string;

  // Informal-address overrides (optional — only provided by locales with T-V distinction)
  LoadingMessageInformal?: string;
  NoSponsorsMessageInformal?: string;
  SponsorUnavailableMessageInformal?: string;
  ErrorMessageInformal?: string;
  TeamsAccessPendingMessageInformal?: string;
  TeamsNotReadyChatTooltipInformal?: string;
  TeamsNotReadyCallTooltipInformal?: string;
}

declare module 'GuestSponsorInfoWebPartStrings' {
  const strings: IGuestSponsorInfoWebPartStrings;
  export = strings;
}
