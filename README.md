# ISO Quality Control MDP

Aplicação web de gestão e controlo documental ISO 9001 do Monte do Pasto.

## Tecnologia

- React + TypeScript + Vite
- Microsoft Entra ID / MSAL para autenticação
- Microsoft Graph para acesso ao SharePoint
- Site: `https://montedopastopt.sharepoint.com/sites/ISOQualityControlMDP`

## Desenvolvimento local

```bash
cp .env.example .env.local
npm install
npm run dev
```

O registo da aplicação no Microsoft Entra ID deve disponibilizar o `client ID` e o `tenant ID` usados nas variáveis de ambiente.

Registo criado no inquilino Monte do Pasto:

- Client ID: `6e4f69e0-d8f5-49d6-99f8-47d4801e2354`
- Tenant ID: `ee417351-ea90-41e0-9147-5ea6ab38ea49`
- SPA redirect local: `http://localhost:5173`
- SPA redirect publicado: `https://montedopasto.github.io/ISO-Quality-Control-MDP/`
- Microsoft Graph: `User.Read` e `Sites.Selected` (acesso limitado ao site da aplicação)

## Regras de auditoria

`ISO_Auditoria` é tratada pela aplicação como append-only: a interface e a camada de dados nunca expõem operações de edição ou eliminação. Qualquer correção origina um novo registo.

## Estado atual

A primeira fundação visual e técnica está criada. Os dados apresentados no dashboard são demonstrativos até a autenticação Microsoft e o acesso ao SharePoint serem ligados.
