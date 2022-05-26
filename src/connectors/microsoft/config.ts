// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

const settings: AppSettings = {
  clientId: process.env.MS_CLIENT_ID,
  clientSecret: process.env.MS_CLIENT_SECRET,
  tenantId: process.env.MS_TENANT_ID,
  authTenant: process.env.AUTH_TENANT,
  graphUserScopes: ['user.read', 'mail.read', 'mail.send'],
};

export interface AppSettings {
  clientId: string;
  clientSecret: string;
  tenantId: string;
  authTenant: string;
  graphUserScopes: string[];
}

export default settings;
