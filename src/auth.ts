import { PublicClientApplication, type Configuration } from '@azure/msal-browser'

const clientId = import.meta.env.VITE_MICROSOFT_CLIENT_ID
const tenantId = import.meta.env.VITE_MICROSOFT_TENANT_ID

export const authConfigured = Boolean(clientId && tenantId)

const configuration: Configuration = {
  auth: {
    clientId: clientId || 'configuration-required',
    authority: tenantId
      ? `https://login.microsoftonline.com/${tenantId}`
      : 'https://login.microsoftonline.com/common',
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: 'sessionStorage',
  },
}

export const msalInstance = new PublicClientApplication(configuration)

export const loginRequest = {
  scopes: ['User.Read'],
}
