// CommonJS equivalent of loc/en-us.js.
// The original locale file uses the AMD define() pattern which jest cannot evaluate
// in a CommonJS test environment. This shim exposes the same strings via module.exports.
module.exports = {
  PropertyPaneDescription: 'Configure the Guest Sponsor Info web part.',
  BasicGroupName: 'Settings',
  TitleFieldLabel: 'Title',
  LoadingMessage: 'Loading your sponsors\u2026',
  NoSponsorsMessage: 'No sponsors found for your account.',
  SponsorUnavailableMessage:
    'Your sponsor/owner is no longer available. Please contact your project stakeholder to have a new sponsor/owner assigned to your account.',
  ErrorMessage: 'Could not load sponsor information. Please try again later.',
  EditModePlaceholder: 'Guest Sponsor Info \u2014 switch to view mode to see sponsors.',
  GuestOnlyPlaceholder: 'This web part is only visible to guest users in view mode.',

  PresenceAvailable: 'Available',
  PresenceAvailableIdle: 'Available, Idle',
  PresenceAway: 'Away',
  PresenceBeRightBack: 'Be Right Back',
  PresenceBusy: 'Busy',
  PresenceBusyIdle: 'Busy, Idle',
  PresenceDoNotDisturb: 'Do Not Disturb',
  PresenceOffline: 'Offline',

  ContactDetailsAriaLabel: 'Contact details for {0}',
  ContactActionsAriaLabel: 'Contact actions',
  ChatTitle: 'Chat via your home Teams account',
  ChatGuestTitle: "Chat as guest in sponsor's tenant",
  EmailTitle: 'Send email',
  CallTitle: 'Call',
  ChatLabel: 'Chat',
  ChatGuestLabel: 'Chat (guest)',
  EmailLabel: 'Email',
  CallLabel: 'Call',

  ContactInfoSection: 'Contact information',
  EmailFieldLabel: 'Email',
  WorkPhoneFieldLabel: 'Work phone',
  MobileFieldLabel: 'Mobile',
  WorkLocationFieldLabel: 'Work location',

  OrganizationSection: 'Organization',
  ManagerLabel: 'Manager',

  CopiedFeedback: 'Copied!',
  CopyEmailAriaLabel: 'Copy email address',
  CopyWorkPhoneAriaLabel: 'Copy work phone',
  CopyMobileAriaLabel: 'Copy mobile number',
  CopyLocationAriaLabel: 'Copy work location',
};
