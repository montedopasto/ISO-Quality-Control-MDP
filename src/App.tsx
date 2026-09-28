import { useMemo, useState } from 'react'
import { useIsAuthenticated, useMsal } from '@azure/msal-react'
import {
  AlertTriangle,
  Bell,
  BookOpenCheck,
  CheckCircle2,
  ChevronRight,
  Clock3,
  FileCheck2,
  FilePlus2,
  Files,
  LayoutDashboard,
  Search,
  Settings,
  ShieldCheck,
  Users,
  LogIn,
  LogOut,
} from 'lucide-react'
import { demoAlerts, demoDocuments } from './data/demo'
import { authConfigured, loginRequest } from './auth'

const navigation = [
  { label: 'Visão geral', icon: LayoutDashboard },
  { label: 'Documentos', icon: Files },
  { label: 'Aprovações', icon: FileCheck2 },
  { label: 'Revisões', icon: BookOpenCheck },
  { label: 'Alertas', icon: Bell, badge: 2 },
  { label: 'Utilizadores', icon: Users },
]

const statusClass: Record<string, string> = {
  'Em vigor': 'status status--green',
  'Aguarda aprovação': 'status status--amber',
  'Em revisão': 'status status--blue',
}

export function App() {
  const [active, setActive] = useState('Visão geral')
  const [query, setQuery] = useState('')
  const { instance, accounts } = useMsal()
  const authenticated = useIsAuthenticated()
  const currentUser = accounts[0]

  async function toggleSession() {
    if (authenticated) {
      await instance.logoutRedirect({ account: currentUser })
      return
    }
    await instance.loginRedirect(loginRequest)
  }

  const documents = useMemo(
    () => demoDocuments.filter((doc) => `${doc.code} ${doc.title} ${doc.owner}`.toLowerCase().includes(query.toLowerCase())),
    [query],
  )

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand__mark">ISO</div>
          <div><strong>Quality Control</strong><span>Monte do Pasto</span></div>
        </div>
        <nav>
          {navigation.map(({ label, icon: Icon, badge }) => (
            <button key={label} className={active === label ? 'nav-item nav-item--active' : 'nav-item'} onClick={() => setActive(label)}>
              <Icon size={19} /><span>{label}</span>{badge && <b>{badge}</b>}
            </button>
          ))}
        </nav>
        <div className="sidebar__footer">
          <button className="nav-item"><Settings size={19} /><span>Configurações</span></button>
          <div className="profile"><div className="avatar">JA</div><div><strong>José Almanso</strong><span>Administrador</span></div></div>
        </div>
      </aside>

      <main>
        <header className="topbar">
          <div><p>Sistema de Gestão da Qualidade</p><h1>{active}</h1></div>
          <div className="topbar__actions">
            <button className="session-button" onClick={toggleSession} disabled={!authConfigured}>
              {authenticated ? <LogOut size={17} /> : <LogIn size={17} />}
              {authenticated ? currentUser?.name || 'Terminar sessão' : 'Entrar com Microsoft'}
            </button>
            <button className="icon-button" aria-label="Alertas"><Bell size={20} /><i /></button>
            <button className="primary"><FilePlus2 size={18} /> Novo documento</button>
          </div>
        </header>

        <section className="content">
          <div className="hero">
            <div><span className="eyebrow"><ShieldCheck size={15} /> ISO 9001</span><h2>Bom trabalho, José.</h2><p>Acompanhe o ciclo documental e as ações que precisam da sua atenção.</p></div>
            <div className="quality-score"><span>Conformidade documental</span><strong>94%</strong><div><i style={{ width: '94%' }} /></div></div>
          </div>

          <div className="metrics">
            <article><span className="metric-icon metric-icon--green"><Files /></span><div><small>Documentos ativos</small><strong>48</strong><em>+3 este mês</em></div></article>
            <article><span className="metric-icon metric-icon--amber"><Clock3 /></span><div><small>Aguardam aprovação</small><strong>6</strong><em>2 urgentes</em></div></article>
            <article><span className="metric-icon metric-icon--blue"><BookOpenCheck /></span><div><small>Em revisão</small><strong>4</strong><em>Dentro do prazo</em></div></article>
            <article><span className="metric-icon metric-icon--red"><AlertTriangle /></span><div><small>Revisões próximas</small><strong>3</strong><em>Próximos 30 dias</em></div></article>
          </div>

          <div className="grid-layout">
            <section className="panel panel--documents">
              <div className="panel__header"><div><h3>Documentos recentes</h3><p>Estado atual do controlo documental</p></div><button>Ver todos <ChevronRight size={16} /></button></div>
              <label className="search"><Search size={18} /><input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Pesquisar por código, título ou responsável" /></label>
              <div className="table-wrap"><table><thead><tr><th>Documento</th><th>Processo</th><th>Versão</th><th>Estado</th><th>Próxima revisão</th></tr></thead><tbody>
                {documents.map((doc) => <tr key={doc.id}><td><strong>{doc.code}</strong><span>{doc.title}</span></td><td>{doc.process}</td><td>v{doc.version}</td><td><span className={statusClass[doc.status] || 'status'}>{doc.status}</span></td><td>{doc.nextReview ? new Intl.DateTimeFormat('pt-PT').format(new Date(doc.nextReview)) : '—'}</td></tr>)}
              </tbody></table></div>
            </section>

            <aside className="panel attention">
              <div className="panel__header"><div><h3>Precisa de atenção</h3><p>Prioridades do momento</p></div></div>
              <div className="attention__list">
                {demoAlerts.map((alert) => <button key={alert.id}><span className={alert.level === 'Urgente' ? 'alert-icon alert-icon--red' : 'alert-icon'}><AlertTriangle size={18} /></span><div><strong>{alert.documentCode}</strong><p>{alert.message}</p><small>{alert.deadline ? `Até ${new Intl.DateTimeFormat('pt-PT').format(new Date(alert.deadline))}` : ''}</small></div><ChevronRight size={17} /></button>)}
              </div>
              <button className="secondary"><CheckCircle2 size={17} /> Abrir centro de tarefas</button>
            </aside>
          </div>
        </section>
      </main>
    </div>
  )
}
