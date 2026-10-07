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

async function graphFetch<T>(instance: IPublicClientApplication, account: AccountInfo, path: string, init?: RequestInit): Promise<T> {
  const token = await accessToken(instance, account)
  const response = await fetch(`${graphRoot}${path}`, {
    ...init,
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', ...init?.headers },
  })
  if (!response.ok) throw new Error(`Microsoft Graph: ${response.status}`)
  if (response.status === 204) return undefined as T
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

function documentVersion(name: string) {
  const matches = [...name.matchAll(/(?:^|[\s_.-])(?:ed(?:i[cç][aã]o)?|rev(?:is[aã]o)?|v(?:ers[aã]o)?)[\s_.-]*(\d+(?:[.,]\d+)?)/gi)]
  return matches[matches.length - 1]?.[1]?.replace(',', '.') || '—'
}

function isObsolete(name: string) {
  return /obsolet[oa]?|obsolo|(?:^|[\s_.-])obs(?:$|[\s_.-])/i.test(name)
}

function documentFamily(document: IsoDocument) {
  return `${document.process}::${document.type}::${document.title}`
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/obsolet[oa]?|obsolo|(?:^|[\s_.-])obs(?:$|[\s_.-])/gi, ' ')
    .replace(/(?:^|[\s_.-])(?:ed(?:icao)?|rev(?:isao)?|v(?:ersao)?)[\s_.-]*\d+(?:[.,]\d+)?/gi, ' ')
    .replace(/[\s_.-]+/g, ' ')
    .trim().toLowerCase()
}

function versionValue(version: string) {
  const value = Number.parseFloat(version)
  return Number.isFinite(value) ? value : -1
}

function consolidateVersions(documents: IsoDocument[]) {
  const families = new Map<string, IsoDocument[]>()
  for (const document of documents) {
    const key = documentFamily(document)
    families.set(key, [...(families.get(key) || []), document])
  }
  return [...families.values()].map(versions => {
    const sorted = [...versions].sort((a,b) => {
      if (a.status !== b.status) return a.status === 'Obsoleto' ? 1 : -1
      return versionValue(b.version) - versionValue(a.version)
        || new Date(b.modifiedAt || 0).getTime() - new Date(a.modifiedAt || 0).getTime()
    })
    const [current, ...previousVersions] = sorted
    return previousVersions.length ? {...current, previousVersions} : current
  })
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
        version: documentVersion(item.name),
        status: isObsolete(item.name) ? 'Obsoleto' : 'Em vigor',
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
  const documents = consolidateVersions(await folderFiles(instance, account, drive.id, folder.id, ''))
  try {
    const metadataSite = await resolveSite(instance, account)
    const list = await resolveList(instance, account, metadataSite.id, sharePointConfig.lists.documents)
    const items = await graphFetch<{ value: Array<{ id: string; fields: Record<string, string | boolean | undefined> }> }>(
      instance, account, `/sites/${metadataSite.id}/lists/${list.id}/items?$expand=fields&$top=999`,
    )
    const byFile = new Map(items.value.map(item => [String(item.fields.FileIDAtual || ''), item]))
    return documents.map(document => {
      const item = byFile.get(document.id)
      if (!item) return document
      const fields = item.fields
      return {
        ...document,
        metadataItemId: item.id,
        status: (fields.EstadoAtual as IsoDocument['status']) || document.status,
        nextReview: typeof fields.DataProximaRevisao === 'string' ? fields.DataProximaRevisao : undefined,
        reviewPeriod: typeof fields.PeriodicidadeRevisao === 'string' ? fields.PeriodicidadeRevisao : undefined,
      }
    })
  } catch {
    return documents
  }
}

async function resolveList(instance: IPublicClientApplication, account: AccountInfo, siteId: string, displayName: string) {
  const lists = await graphFetch<{ value: Array<{ id: string; displayName: string }> }>(instance, account, `/sites/${siteId}/lists?$select=id,displayName`)
  const list = lists.value.find(candidate => candidate.displayName === displayName)
  if (!list) throw new Error(`Lista ${displayName} não encontrada.`)
  return list
}

export async function saveDocumentMetadata(instance: IPublicClientApplication, account: AccountInfo, document: IsoDocument) {
  const site = await resolveSite(instance, account)
  const list = await resolveList(instance, account, site.id, sharePointConfig.lists.documents)
  const fields = {
    Title: document.code,
    DocumentoID: document.id,
    Codigo: document.code,
    Titulo: document.title,
    VersaoAtual: document.version,
    EstadoAtual: document.status,
    DataProximaRevisao: document.nextReview || null,
    PeriodicidadeRevisao: document.reviewPeriod || null,
    FileIDAtual: document.id,
    Ativo: document.status !== 'Obsoleto' && document.status !== 'Cancelado',
    UltimaAlteracaoPor: account.name || account.username,
    UltimaAlteracaoEm: new Date().toISOString(),
  }
  if (document.metadataItemId) {
    await graphFetch(instance, account, `/sites/${site.id}/lists/${list.id}/items/${document.metadataItemId}/fields`, {
      method: 'PATCH', body: JSON.stringify(fields),
    })
    return document.metadataItemId
  }
  const created = await graphFetch<{ id: string }>(instance, account, `/sites/${site.id}/lists/${list.id}/items`, {
    method: 'POST', body: JSON.stringify({ fields }),
  })
  return created.id
}
