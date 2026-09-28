export type DocumentStatus =
  | 'Rascunho'
  | 'Em revisão'
  | 'Aguarda aprovação'
  | 'Aprovado'
  | 'Em vigor'
  | 'Obsoleto'
  | 'Rejeitado'
  | 'Cancelado'

export interface IsoDocument {
  id: string
  code: string
  title: string
  type: string
  process: string
  owner: string
  version: string
  status: DocumentStatus
  nextReview?: string
}

export interface IsoAlert {
  id: string
  documentCode: string
  message: string
  level: 'Informação' | 'Aviso' | 'Urgente' | 'Vencido'
  deadline?: string
}
