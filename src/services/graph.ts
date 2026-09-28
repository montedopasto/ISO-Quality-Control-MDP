import type { AccountInfo, IPublicClientApplication } from '@azure/msal-browser'
import { sharePointConfig } from '../config'

const graphRoot = 'https://graph.microsoft.com/v1.0'

async function accessToken(instance: IPublicClientApplication, account: AccountInfo) {
  const result = await instance.acquireTokenSilent({
    account,
    scopes: ['User.Read', 'Sites.ReadWrite.All'],
  })
  return result.accessToken
}

async function graphFetch<T>(instance: IPublicClientApplication, account: AccountInfo, path: string): Promise<T> {
  const token = await accessToken(instance, account)
  const response = await fetch(`${graphRoot}${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  })
  if (!response.ok) throw new Error(`Microsoft Graph: ${response.status}`)
  return response.json() as Promise<T>
}

export async function resolveSite(instance: IPublicClientApplication, account: AccountInfo) {
  return graphFetch<{ id: string; displayName: string }>(
    instance,
    account,
    `/sites/${sharePointConfig.hostname}:${sharePointConfig.sitePath}`,
  )
}
