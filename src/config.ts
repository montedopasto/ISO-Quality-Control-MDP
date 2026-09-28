export const sharePointConfig = {
  hostname: import.meta.env.VITE_SHAREPOINT_HOSTNAME || 'montedopastopt.sharepoint.com',
  sitePath: import.meta.env.VITE_SHAREPOINT_SITE_PATH || '/sites/ISOQualityControlMDP',
  lists: {
    documents: 'ISO_Documentos_Master',
    versions: 'ISO_Versoes',
    approvals: 'ISO_Aprovacoes',
    audit: 'ISO_Auditoria',
    alerts: 'ISO_Alertas',
    processes: 'ISO_Processos',
    departments: 'ISO_Departamentos',
    documentTypes: 'ISO_TiposDocumento',
    users: 'ISO_Utilizadores',
  },
  library: 'ISO_Documentos',
} as const
