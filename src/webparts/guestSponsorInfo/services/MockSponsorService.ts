// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0

import { ISponsor } from './ISponsor';

import annaMuellerPhoto from '../assets/mock/anna-mueller.jpg';
import jamesAndersonPhoto from '../assets/mock/james-anderson.jpg';
import sophieLaurentPhoto from '../assets/mock/sophie-laurent.jpg';
import luisRodriguezPhoto from '../assets/mock/luis-rodriguez.jpg';
import priyaSharmaPhoto from '../assets/mock/priya-sharma.jpg';
import thomasSchneiderPhoto from '../assets/mock/thomas-schneider.jpg';
import sarahWebbPhoto from '../assets/mock/sarah-webb.jpg';
import marcDuboisPhoto from '../assets/mock/marc-dubois.jpg';
import elenaFernandezPhoto from '../assets/mock/elena-fernandez.jpg';
import oliverThompsonPhoto from '../assets/mock/oliver-thompson.jpg';

/**
 * Realistic-looking but entirely fictitious sponsor records used when demo
 * mode is enabled in the property pane.  No Graph calls are made; these
 * objects are returned directly so that the web part can be previewed in the
 * local workbench without a tenant connection or a guest account.
 *
 * Profile photos are AI-generated samples from randomuser.me, which
 * provides portrait images specifically for UI mockups and demos.
 */
export const MOCK_SPONSORS: ISponsor[] = [
  {
    id: 'mock-1',
    displayName: 'Anna Müller',
    givenName: 'Anna',
    surname: 'Müller',
    mail: 'anna.mueller@contoso.com',
    jobTitle: 'IT Manager',
    department: 'Information Technology',
    officeLocation: 'BER-HQ / Bldg A / Floor 4 / A4-12',
    streetAddress: 'Unter den Linden 1',
    postalCode: '10117',
    state: 'Berlin',
    city: 'Berlin',
    country: 'Germany',
    businessPhones: ['+49 30 12345678'],
    mobilePhone: '+49 172 3456789',
    photoUrl: annaMuellerPhoto,
    presence: 'Available',
    presenceActivity: 'Available',
    hasTeams: true,
    managerDisplayName: 'Thomas Schneider',
    managerGivenName: 'Thomas',
    managerSurname: 'Schneider',
    managerJobTitle: 'Head of IT',
    managerDepartment: 'Information Technology',
    managerId: 'mock-mgr-1',
    managerPhotoUrl: thomasSchneiderPhoto,
  },
  {
    id: 'mock-2',
    displayName: 'James Anderson',
    givenName: 'James',
    surname: 'Anderson',
    mail: 'james.anderson@contoso.com',
    jobTitle: 'Project Lead',
    department: 'Business Development',
    officeLocation: 'MUC-03 / Bldg C / Floor 2 / C2-08',
    streetAddress: 'Maximilianstraße 12',
    postalCode: '80539',
    state: 'Bavaria',
    city: 'Munich',
    country: 'Germany',
    businessPhones: ['+49 89 98765432'],
    mobilePhone: '+49 151 98765432',
    photoUrl: jamesAndersonPhoto,
    presence: 'Busy',
    presenceActivity: 'InAMeeting',
    hasTeams: true,
    managerDisplayName: 'Sarah Webb',
    managerGivenName: 'Sarah',
    managerSurname: 'Webb',
    managerJobTitle: 'VP Business Development',
    managerDepartment: 'Business Development',
    managerId: 'mock-mgr-2',
    managerPhotoUrl: sarahWebbPhoto,
  },
  {
    id: 'mock-3',
    displayName: 'Sophie Laurent',
    givenName: 'Sophie',
    surname: 'Laurent',
    mail: 'sophie.laurent@contoso.com',
    jobTitle: 'Corporate Relations Manager',
    department: 'Communications',
    officeLocation: 'PAR-01 / Bldg D / Floor 3 / D3-07',
    streetAddress: 'Avenue des Champs-Élysées 42',
    postalCode: '75008',
    state: 'Île-de-France',
    city: 'Paris',
    country: 'France',
    businessPhones: ['+33 1 23456789'],
    mobilePhone: '+33 6 12345678',
    photoUrl: sophieLaurentPhoto,
    presence: 'Away',
    presenceActivity: 'Away',
    hasTeams: true,
    managerDisplayName: 'Marc Dubois',
    managerGivenName: 'Marc',
    managerSurname: 'Dubois',
    managerJobTitle: 'Chief Communications Officer',
    managerDepartment: 'Communications',
    managerId: 'mock-mgr-3',
    managerPhotoUrl: marcDuboisPhoto,
  },
  {
    id: 'mock-4',
    displayName: 'Luis Rodríguez',
    givenName: 'Luis',
    surname: 'Rodríguez',
    mail: 'luis.rodriguez@contoso.com',
    jobTitle: 'Senior Legal Counsel',
    department: 'Legal',
    officeLocation: 'MAD-02 / Bldg B / Floor 5 / B5-11',
    streetAddress: 'Calle de Alcalá 50',
    postalCode: '28014',
    state: 'Community of Madrid',
    city: 'Madrid',
    country: 'Spain',
    businessPhones: ['+34 91 3456789'],
    mobilePhone: '+34 622 345678',
    photoUrl: luisRodriguezPhoto,
    presence: 'Available',
    presenceActivity: 'Available',
    hasTeams: true,
    managerDisplayName: 'Elena Fernández',
    managerGivenName: 'Elena',
    managerSurname: 'Fernández',
    managerJobTitle: 'General Counsel',
    managerDepartment: 'Legal',
    managerId: 'mock-mgr-4',
    managerPhotoUrl: elenaFernandezPhoto,
  },
  {
    id: 'mock-5',
    displayName: 'Priya Sharma',
    givenName: 'Priya',
    surname: 'Sharma',
    mail: 'priya.sharma@contoso.com',
    jobTitle: 'HR Business Partner',
    department: 'Human Resources',
    officeLocation: 'LON-HQ / Bldg A / Floor 2 / A2-05',
    streetAddress: 'Canary Wharf 25',
    postalCode: 'E14 5AB',
    state: 'England',
    city: 'London',
    country: 'United Kingdom',
    businessPhones: ['+44 20 71234567'],
    mobilePhone: '+44 7700 234567',
    photoUrl: priyaSharmaPhoto,
    presence: 'DoNotDisturb',
    presenceActivity: 'Presenting',
    hasTeams: true,
    managerDisplayName: 'Oliver Thompson',
    managerGivenName: 'Oliver',
    managerSurname: 'Thompson',
    managerJobTitle: 'Chief People Officer',
    managerDepartment: 'Human Resources',
    managerId: 'mock-mgr-5',
    managerPhotoUrl: oliverThompsonPhoto,
  },
];
