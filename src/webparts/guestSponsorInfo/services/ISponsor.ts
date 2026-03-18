/** Data model for a single sponsor of the signed-in guest user. */
export interface ISponsor {
  /** Entra object ID of the sponsor user. */
  id: string;
  /** Full display name. */
  displayName: string;
  /** Primary SMTP address. */
  mail?: string;
  /** Job title. */
  jobTitle?: string;
  /** Department name. */
  department?: string;
  /** Office location string. */
  officeLocation?: string;
  /** Work phone numbers. */
  businessPhones?: string[];
  /** Mobile phone number. */
  mobilePhone?: string;
  /** Data URL (base64-encoded JPEG) of the profile photo. Undefined when no photo is available. */
  photoUrl?: string;
  /**
   * Microsoft Graph presence availability string for this user.
   * Possible values: Available, AvailableIdle, Away, BeRightBack, Busy, BusyIdle,
   * DoNotDisturb, Offline, PresenceUnknown.
   * Requires the Presence.Read.All delegated permission.
   */
  presence?: string;
  /** Display name of the sponsor's direct manager. */
  managerDisplayName?: string;
  /** Job title of the sponsor's direct manager. */
  managerJobTitle?: string;
  /** Data URL (base64-encoded JPEG) of the manager's profile photo. */
  managerPhotoUrl?: string;
}
