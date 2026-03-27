// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0

/** Data model for a single sponsor of the signed-in guest user. */
export interface ISponsor {
  /** Entra object ID of the sponsor user. */
  id: string;
  /** Full display name. */
  displayName: string;
  /** Given (first) name — preferred over displayName for rendering. */
  givenName?: string;
  /** Family (last) name — preferred over displayName for rendering. */
  surname?: string;
  /** Primary SMTP address. */
  mail?: string;
  /** Job title. */
  jobTitle?: string;
  /** Department name. */
  department?: string;
  /** Office location string. */
  officeLocation?: string;
  /** Street address. */
  streetAddress?: string;
  /** Postal code. */
  postalCode?: string;
  /** State or province. */
  state?: string;
  /** City. */
  city?: string;
  /** Country or region. */
  country?: string;
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
   * Populated when Presence.Read.All is granted (optional).
   */
  presence?: string;
  /**
   * Microsoft Graph presence activity string for this user.
   * Example values: InAMeeting, InACall, Presenting.
   * Populated when Presence.Read.All is granted (optional).
   */
  presenceActivity?: string;
  /** Display name of the sponsor's direct manager. */
  managerDisplayName?: string;
  /** Given (first) name of the sponsor's direct manager. */
  managerGivenName?: string;
  /** Family (last) name of the sponsor's direct manager. */
  managerSurname?: string;
  /** Job title of the sponsor's direct manager. */
  managerJobTitle?: string;
  /** Department of the sponsor's direct manager. */
  managerDepartment?: string;
  /** Data URL (base64-encoded JPEG) of the manager's profile photo. */
  managerPhotoUrl?: string;
  /**
   * True when the sponsor has an active Microsoft Teams license.
   * False when they do not (hide Teams Chat/Call buttons and presence indicator).
   * Undefined when the license status could not be determined (show everything).
   */
  hasTeams?: boolean;
  /**
   * Entra object ID of the sponsor's direct manager.
   * Present only when the manager relationship could be resolved.
   * Used by the client to progressively load the manager's photo.
   */
  managerId?: string;
}
