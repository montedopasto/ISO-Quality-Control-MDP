import React from 'react'
import ReactDOM from 'react-dom/client'
import { MsalProvider } from '@azure/msal-react'
import { App } from './App'
import { msalInstance } from './auth'
import './styles.css'

async function bootstrap() {
  await msalInstance.initialize()
  ReactDOM.createRoot(document.getElementById('root')!).render(
    <React.StrictMode>
      <MsalProvider instance={msalInstance}>
        <App />
      </MsalProvider>
    </React.StrictMode>,
  )
}

void bootstrap()
