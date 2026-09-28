import type { IsoAlert, IsoDocument } from '../types'

export const demoDocuments: IsoDocument[] = [
  { id: 'DOC-0001', code: 'PR-001', title: 'Controlo de documentos e registos', type: 'Procedimento', process: 'Gestão da Qualidade', owner: 'José Almanso', version: '01', status: 'Em vigor', nextReview: '2027-09-22' },
  { id: 'DOC-0002', code: 'IT-004', title: 'Tratamento de não conformidades', type: 'Instrução', process: 'Melhoria contínua', owner: 'José Almanso', version: '02', status: 'Aguarda aprovação', nextReview: '2027-03-15' },
  { id: 'DOC-0003', code: 'FR-012', title: 'Registo de auditoria interna', type: 'Formulário', process: 'Auditoria', owner: 'José Almanso', version: '03', status: 'Em revisão', nextReview: '2026-10-08' },
]

export const demoAlerts: IsoAlert[] = [
  { id: 'ALT-001', documentCode: 'FR-012', message: 'Revisão documental aproxima-se do prazo', level: 'Aviso', deadline: '2026-10-08' },
  { id: 'ALT-002', documentCode: 'IT-004', message: 'Documento aguarda decisão do aprovador', level: 'Urgente', deadline: '2026-09-30' },
]
