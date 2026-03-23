declare interface IGuestSponsorInfoWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
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
  MockModePlaceholder: string;
  MockSimulatedHintFieldLabel: string;
  MockSimulatedHintNoneOption: string;
  MockSimulatedHintTeamsAccessPendingOption: string;
  MockSimulatedHintVersionMismatchOption: string;
  MockSimulatedHintSponsorUnavailableOption: string;
  MockSimulatedHintNoSponsorsOption: string;

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
  ExternalMapProviderFieldLabel: string;
  MapProviderBingOption: string;
  MapProviderGoogleOption: string;
  MapProviderAppleOption: string;
  MapProviderOpenStreetMapOption: string;
  MapProviderHereOption: string;
  MapProviderNoneOption: string;
  AddressMapProviderHint: string;
  LocationDisplayHintSeparateRows: string;
  LocationDisplayHintHidden: string;
  AdvancedDisplayGroupName: string;
  ManagerOptionsDisabledHint: string;
  ShowManagerFieldLabel: string;
  ShowPresenceFieldLabel: string;
  ShowSponsorPhotoFieldLabel: string;
  ShowManagerPhotoFieldLabel: string;
  UseInformalAddressFieldLabel: string;
  UseInformalAddressHint: string;

  // Job title and department toggles
  ShowSponsorJobTitleFieldLabel: string;
  ShowManagerJobTitleFieldLabel: string;
  ShowSponsorDepartmentFieldLabel: string;
  ShowManagerDepartmentFieldLabel: string;
  DepartmentLabel: string;

  // Azure Function proxy
  FunctionGroupName: string;
  FunctionUrlFieldLabel: string;
  FunctionClientIdFieldLabel: string;
  ProxyStatusChecking: string;
  ProxyStatusOk: string;
  ProxyStatusError: string;

  // Author/company information section
  AuthorSectionGroupName: string;
  AuthorSectionIntro: string;
  AuthorSectionWebsiteLinkLabel: string;
  AuthorSectionConsultingText: string;
  AuthorSectionPartnerPrefix: string;
  AuthorSectionPartnerLinkLabel: string;
  AuthorSectionPartnerSuffix: string;
  AuthorSectionPartnerTagline: string;
  AuthorSectionSourceCodeLabel: string;
  AuthorSectionVersionLabel: string;

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
