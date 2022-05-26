// npm install @azure/identity @microsoft/microsoft-graph-client isomorphic-fetch readline-sync
// npm install -D @microsoft/microsoft-graph-types @types/node @types/readline-sync @types/isomorphic-fetch

// Ref https://docs.microsoft.com/en-us/graph/tutorials/typescript?tabs=aad&tutorial-step=2

import 'isomorphic-fetch';
import { ClientSecretCredential, DeviceCodeCredential, DeviceCodePromptCallback } from '@azure/identity';
import { Client, PageCollection } from '@microsoft/microsoft-graph-client';
import { User, Message } from '@microsoft/microsoft-graph-types';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';

import { AppSettings } from './config';

let _settings: AppSettings | undefined = undefined;
let _deviceCodeCredential: DeviceCodeCredential | undefined = undefined;
let _userClient: Client | undefined = undefined;

export function initializeGraphForUserAuth(settings: AppSettings, deviceCodePrompt: DeviceCodePromptCallback) {
  // Ensure settings isn't null
  if (!settings) {
    throw new Error('Settings cannot be undefined');
  }

  _settings = settings;

  _deviceCodeCredential = new DeviceCodeCredential({
    clientId: settings.clientId,
    tenantId: settings.authTenant,
    userPromptCallback: deviceCodePrompt,
  });

  const authProvider = new TokenCredentialAuthenticationProvider(_deviceCodeCredential, {
    scopes: settings.graphUserScopes,
  });

  _userClient = Client.initWithMiddleware({
    authProvider: authProvider,
  });
}

export async function getUserTokenAsync(): Promise<string> {
  // Ensure credential isn't undefined
  if (!_deviceCodeCredential) {
    throw new Error('Graph has not been initialized for user auth');
  }

  // Ensure scopes isn't undefined
  if (!_settings?.graphUserScopes) {
    throw new Error('Setting "scopes" cannot be undefined');
  }

  // Request token with given scopes
  const response = await _deviceCodeCredential.getToken(_settings?.graphUserScopes);
  return response.token;
}
