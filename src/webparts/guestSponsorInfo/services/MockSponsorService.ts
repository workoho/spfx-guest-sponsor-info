import { ISponsor } from './ISponsor';

/**
 * Realistic-looking but entirely fictitious sponsor records used when demo
 * mode is enabled in the property pane.  No Graph calls are made; these
 * objects are returned directly so that the web part can be previewed in the
 * local workbench without a tenant connection or a guest account.
 */
export const MOCK_SPONSORS: ISponsor[] = [
  {
    id: 'mock-1',
    displayName: 'Anna Müller',
    mail: 'anna.mueller@contoso.com',
    jobTitle: 'IT Manager',
    department: 'Information Technology',
    officeLocation: 'Berlin',
    businessPhones: ['+49 30 12345678'],
    mobilePhone: undefined,
    photoUrl: undefined,
    presence: 'Available',
    managerDisplayName: 'Thomas Schneider',
    managerJobTitle: 'Head of IT',
    managerPhotoUrl: undefined,
  },
  {
    id: 'mock-2',
    displayName: 'James Anderson',
    mail: 'james.anderson@contoso.com',
    jobTitle: 'Project Lead',
    department: 'Business Development',
    officeLocation: 'Munich',
    businessPhones: [],
    mobilePhone: '+49 151 98765432',
    photoUrl: undefined,
    presence: 'Busy',
    managerDisplayName: 'Sarah Webb',
    managerJobTitle: 'VP Business Development',
    managerPhotoUrl: undefined,
  },
];
