import { PublicClientApplication, type Configuration } from '@azure/msal-browser'

const clientId = import.meta.env.VITE_MICROSOFT_CLIENT_ID
const tenantId = import.meta.env.VITE_MICROSOFT_TENANT_ID
const redirectUri = import.meta.env.PROD
  ? `${window.location.origin}/ISO-Quality-Control-MDP/`
  : window.location.origin

export const authConfigured = Boolean(clientId && tenantId)

const configuration: Configuration = {
  auth: {
    clientId: clientId || 'configuration-required',
    authority: tenantId
      ? `https://login.microsoftonline.com/${tenantId}`
      : 'https://login.microsoftonline.com/common',
    redirectUri,
    postLogoutRedirectUri: redirectUri,
  },
  cache: {
    cacheLocation: 'sessionStorage',
  },
}

export const msalInstance = new PublicClientApplication(configuration)

export const loginRequest = {
  scopes: ['User.Read', 'Sites.Selected'],
}
