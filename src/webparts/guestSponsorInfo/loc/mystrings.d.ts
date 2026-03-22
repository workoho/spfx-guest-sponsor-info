declare interface IGuestSponsorInfoWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  LoadingMessage: string;
  NoSponsorsMessage: string;
  SponsorUnavailableMessage: string;
  ErrorMessage: string;
  EditModePlaceholder: string;
  GuestOnlyPlaceholder: string;
  MockModeFieldLabel: string;
  MockModePlaceholder: string;

  // Display toggles
  DisplayGroupName: string;
  ShowBusinessPhonesFieldLabel: string;
  ShowMobilePhoneFieldLabel: string;
  ShowWorkLocationFieldLabel: string;
  ShowCityFieldLabel: string;
  ShowCountryFieldLabel: string;
  ShowStructuredAddressFieldLabel: string;
  ShowStreetAddressFieldLabel: string;
  ShowPostalCodeFieldLabel: string;
  ShowStateFieldLabel: string;
  ShowAddressMapFieldLabel: string;
  AzureMapsSubscriptionKeyFieldLabel: string;
  ExternalMapProviderFieldLabel: string;
  MapProviderBingOption: string;
  MapProviderGoogleOption: string;
  MapProviderAppleOption: string;
  MapProviderOpenStreetMapOption: string;
  MapProviderHereOption: string;
  LocationDisplayHintAddressEnabled: string;
  LocationDisplayHintSeparateRows: string;
  LocationDisplayHintHidden: string;
  AddressMapFallbackHint: string;
  AdvancedDisplayGroupName: string;
  ManagerOptionsDisabledHint: string;
  ShowManagerFieldLabel: string;
  ShowPresenceFieldLabel: string;
  UseInformalAddressFieldLabel: string;

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
  CityFieldLabel: string;
  CountryFieldLabel: string;
  StreetAddressFieldLabel: string;
  PostalCodeFieldLabel: string;
  StateFieldLabel: string;

  // Organization section
  OrganizationSection: string;
  ManagerLabel: string;

  // Copy-to-clipboard buttons
  CopiedFeedback: string;
  CopyEmailAriaLabel: string;
  CopyWorkPhoneAriaLabel: string;
  CopyMobileAriaLabel: string;
  CopyLocationAriaLabel: string;
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
  VersionMismatchMessage: string;

  // Shown when the Azure Function managed identity is missing required Graph permissions
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
