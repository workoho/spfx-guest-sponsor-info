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

  // Azure Function proxy
  FunctionGroupName: string;
  FunctionUrlFieldLabel: string;
  FunctionClientIdFieldLabel: string;

  // Presence labels
  PresenceAvailable: string;
  PresenceAvailableIdle: string;
  PresenceAway: string;
  PresenceBeRightBack: string;
  PresenceBusy: string;
  PresenceBusyIdle: string;
  PresenceDoNotDisturb: string;
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

  // Organization section
  OrganizationSection: string;
  ManagerLabel: string;

  // Copy-to-clipboard buttons
  CopiedFeedback: string;
  CopyEmailAriaLabel: string;
  CopyWorkPhoneAriaLabel: string;
  CopyMobileAriaLabel: string;
  CopyLocationAriaLabel: string;
}

declare module 'GuestSponsorInfoWebPartStrings' {
  const strings: IGuestSponsorInfoWebPartStrings;
  export = strings;
}
