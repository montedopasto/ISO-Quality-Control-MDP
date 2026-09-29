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
  source: {
    sitePath: '/sites/MontedoPasto',
    driveWebPath: '/Documentos Partilhados1',
    folderPath: 'documentos partilhados/Documentos Comuns/QUALIDADE - ISO 9001-2015',
    folderUrl: 'https://montedopastopt.sharepoint.com/sites/MontedoPasto/Documentos%20Partilhados1/Forms/AllItems.aspx?id=%2Fsites%2FMontedoPasto%2FDocumentos%20Partilhados1%2Fdocumentos%20partilhados%2FDocumentos%20Comuns%2FQUALIDADE%20%2D%20ISO%209001%2D2015',
  },
} as const
