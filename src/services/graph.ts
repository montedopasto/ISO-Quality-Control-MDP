import type { AccountInfo, IPublicClientApplication } from '@azure/msal-browser'
import { sharePointConfig } from '../config'
import type { IsoDocument } from '../types'

const graphRoot = 'https://graph.microsoft.com/v1.0'

async function accessToken(instance: IPublicClientApplication, account: AccountInfo) {
  const result = await instance.acquireTokenSilent({
    account,
    scopes: ['User.Read', 'Sites.Selected'],
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

type GraphDrive = { id: string; name: string; webUrl: string }
type DriveItem = {
  id: string
  name: string
  webUrl: string
  lastModifiedDateTime: string
  lastModifiedBy?: { user?: { displayName?: string } }
  file?: { mimeType: string }
  folder?: { childCount: number }
}

function documentCode(name: string) {
  return name.match(/^[A-Z]{2,6}[.\-_ ]?\d+(?:[.\-_ ]?\d+)?/i)?.[0]?.replace(/[_ ]/g, '.') || name.replace(/\.[^.]+$/, '').slice(0, 18)
}

function documentType(name: string) {
  const extension = name.split('.').pop()?.toUpperCase() || 'FICHEIRO'
  return extension === 'DOCX' || extension === 'DOC' ? 'Word'
    : extension === 'XLSX' || extension === 'XLS' ? 'Excel'
      : extension === 'PPTX' || extension === 'PPT' ? 'PowerPoint'
        : extension === 'PDF' ? 'PDF'
          : extension
}

async function folderFiles(instance: IPublicClientApplication, account: AccountInfo, driveId: string, itemId: string, process: string): Promise<IsoDocument[]> {
  const result = await graphFetch<{ value: DriveItem[] }>(instance, account, `/drives/${driveId}/items/${itemId}/children?$top=999`)
  const documents: IsoDocument[] = []
  for (const item of result.value) {
    if (item.folder) {
      documents.push(...await folderFiles(instance, account, driveId, item.id, process || item.name))
    } else if (item.file) {
      documents.push({
        id: item.id,
        code: documentCode(item.name),
        title: item.name.replace(/\.[^.]+$/, ''),
        type: documentType(item.name),
        process: process || 'Documentos gerais',
        owner: item.lastModifiedBy?.user?.displayName || 'SharePoint',
        version: '—',
        status: 'Em vigor',
        sourceUrl: item.webUrl,
        modifiedAt: item.lastModifiedDateTime,
      })
    }
  }
  return documents
}

export async function loadExistingDocuments(instance: IPublicClientApplication, account: AccountInfo) {
  const site = await graphFetch<{ id: string }>(instance, account, `/sites/${sharePointConfig.hostname}:${sharePointConfig.source.sitePath}`)
  const drives = await graphFetch<{ value: GraphDrive[] }>(instance, account, `/sites/${site.id}/drives`)
  const drive = drives.value.find(item => decodeURI(item.webUrl).includes(sharePointConfig.source.driveWebPath))
    || drives.value.find(item => item.name.toLowerCase().includes('documentos partilhados'))
  if (!drive) throw new Error('Biblioteca Documentos Partilhados não encontrada.')
  const folder = await graphFetch<DriveItem>(instance, account, `/drives/${drive.id}/root:/${encodeURI(sharePointConfig.source.folderPath)}`)
  return folderFiles(instance, account, drive.id, folder.id, '')
}
