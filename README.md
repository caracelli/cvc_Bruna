# CVC-Trata-Forms

Automação de e-mail (caixa compartilhada **"Gestão de Acessos"**) + abertura de
chamado no **Jira Service Management**, via **automação de navegador** (Edge),
usando conta corporativa **sem necessidade de admin** nem de API token.

> A Microsoft Graph API foi descartada porque o tenant da CVC exige consentimento
> de administrador para permissões de e-mail. Os arquivos do "Plano A" (Graph)
> ficam arquivados: `auth.py`, `graph_client.py`, `config.py`, `test_conexao.py`.

---

## ✅ Parte de LOGIN (finalizada)

O login dos dois portais é automático — você só faz o **MFA**.

### Como funciona
- O Edge é aberto num **perfil dedicado** (`C:\Users\user\edge_cvc_profile`) em
  **modo depuração** (porta 9222). O Playwright se conecta nele via CDP.
- **Outlook:** preenche e-mail + senha automaticamente, você aprova o **MFA**,
  e em "Continuar conectado?" responde **sempre NÃO**.
- **Jira:** Service Desk → e-mail → "Próximo" → "Continue com a conta da
  Atlassian" → entra pelo **SSO Microsoft** (nunca usa senha da Atlassian).
- **Duas frentes:** se houver credencial salva, faz login automático; se não
  houver (ou se o automático falhar), **espera você logar manualmente**.

### Credenciais (seguras)
A senha fica no `credenciais.xml` (local, no `.gitignore`) **ou** no Gerenciador
de Credenciais do Windows (DPAPI). Para cadastrar/trocar:
```powershell
python salvar_credenciais.py          # digita e-mail + senha (oculta)
python salvar_credenciais.py --apagar # remove
```
> E-mail de login: **nelsondiniz@ext.cvccorp.com.br**

### Rodar o login
```powershell
pip install -r requirements.txt       # primeira vez
python iniciar_sessoes.py             # abre Edge, loga Outlook e depois Jira
```
O Edge fica **aberto** após o login (sobrevive ao fim do script), pronto para o
fluxo. Verificar a qualquer momento:
```powershell
python checar_login.py
```

---

## Fluxo principal (em desenvolvimento)

```powershell
python fluxo.py        # roda em DRY-RUN: mostra o plano, NÃO altera nada
```
Para cada e-mail do remetente **Microsoft Forms** NÃO LIDO na caixa
"Gestão de Acessos": cria chamado no Jira → aplica categoria com o nº do ticket
→ move para "Finalizado" → marca como lido. As ações reais só rodam com
`DRY_RUN = False` (após validação). Se o Jira cair no SSO, o fluxo religa sozinho.

Ao finalizar (modo real), as sessões de Outlook e Jira são encerradas
(`encerrar_sessoes.py`).

---

## Estrutura

| Arquivo                | Função                                                       |
|------------------------|--------------------------------------------------------------|
| `iniciar_sessoes.py`   | **Bootstrap de login** (Outlook → Jira) — rode primeiro      |
| `login_outlook.py`     | Login Microsoft (picker/e-mail/senha/MFA/KMSI=Não)           |
| `login_jira.py`        | Login Jira via SSO Microsoft (mesma janela)                  |
| `credenciais.py` / `salvar_credenciais.py` / `credenciais.xml` | Cofre de credenciais     |
| `checar_login.py`      | Verifica se Outlook e Jira estão logados                     |
| `outlook_web.py`       | Conexão CDP + ler caixa/pastas/e-mails (não lido)            |
| `ler_caixa.py` / `filtrar.py` | Ler e filtrar e-mails                                  |
| `fluxo.py`             | Orquestrador (DRY-RUN por padrão)                            |
| `encerrar_sessoes.py`  | Fecha o Edge = encerra as sessões                            |
| `abrir_edge.ps1` / `abrir_chrome.ps1` | Abrir navegador manualmente (alternativa)    |

## Requisitos
`playwright`, `keyring`, `requests`, `msal` (ver `requirements.txt`).
