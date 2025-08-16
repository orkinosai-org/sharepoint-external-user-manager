export interface IExternalLibrary {
  id: string;
  name: string;
  description: string;
  siteUrl: string;
  externalUsersCount: number;
  lastModified: Date;
  owner: string;
  permissions: 'Read' | 'Contribute' | 'Full Control';
}

export interface IExternalUser {
  id: string;
  email: string;
  displayName: string;
  invitedBy: string;
  invitedDate: Date;
  lastAccess: Date;
  permissions: 'Read' | 'Contribute' | 'Full Control';
}