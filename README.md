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

## Regras de auditoria

`ISO_Auditoria` é tratada pela aplicação como append-only: a interface e a camada de dados nunca expõem operações de edição ou eliminação. Qualquer correção origina um novo registo.

## Estado atual

A primeira fundação visual e técnica está criada. Os dados apresentados no dashboard são demonstrativos até a autenticação Microsoft e o acesso ao SharePoint serem ligados.
