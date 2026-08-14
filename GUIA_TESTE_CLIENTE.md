# Guia — Teste na máquina do cliente & handoff (Bruna)

CVC‑Trata‑Forms: automação que lê os e‑mails do **Microsoft Forms** na caixa
compartilhada, abre o chamado no **Jira (via API)**, encaminha pro grupo, marca
como lido e move pra **Finalizados**.

---

## 1. Pré‑requisitos na máquina do cliente
- Windows + **Microsoft Edge** instalado (o app detecta sozinho).
- Rede liberada para `outlook.office.com` e `cvccorp.atlassian.net`.
- Conta Microsoft com acesso à caixa compartilhada **"Gestão de Acessos - Argentina"**.
- Usuário Jira com acesso ao portal **GAAR — Gestión de Accesos**.

## 2. Cada operador precisa do PRÓPRIO token Jira (importante)
O token é **por usuário** — não é compartilhado.
1. Acessar `https://id.atlassian.com/manage-profile/security/api-tokens`
2. **Create API token** → dar um nome → **copiar** o token (só aparece uma vez).
3. Esse token vai no formulário de configuração (passo 3).

> O token usado nos meus testes é meu e **não** vem no pacote (`<token>` vazio).

## 3. Primeiro uso (configuração)
1. Extrair o `CVC-Trata-Forms.zip` numa pasta.
2. Rodar `CVC-Trata-Forms.exe` → como faltam credenciais, ele abre o **formulário**.
3. Preencher:
   - **Outlook:** e‑mail e senha do operador.
   - **Aba Jira:** *Usuário Jira (e‑mail)* + *Token Jira* (do passo 2).
   - **Destinatários** do encaminhamento.
4. **Deixar em DRY‑RUN** (ensaio) para o primeiro teste.
5. Salvar → o app reinicia e pede o **MFA** na janela do Edge (aprovar).

## 4. Progressão de teste (do mais seguro ao real)
| Modo | O que faz | Quando usar |
|---|---|---|
| **DRY‑RUN** (ensaio) | Só lista o que *seria* feito. **Não cria/move nada.** | 1º teste: valida acesso, rede, caixa e portal. |
| **DEMO** (faz‑e‑desfaz) | Cria chamado real → encaminha → move p/ Finalizados → **cancela e restaura**. | Ver o fluxo completo sem deixar resíduo. |
| **REAL** (produção) | Cria + move **definitivo** (não desfaz). | Operação de verdade — ligar quando confiante. |

Trocar de modo pelo formulário (menu do ícone na bandeja → *Configurações*).
Menu → **Ver log** mostra o que aconteceu em cada ciclo.

## 5. IDs do portal — CONFIRMADOS (produção)
Embutidos no pacote (fila **GAAR — Argentina**, mesma já usada em produção):
- `service_desk_id = 1984`  ·  `request_type_id = 7550`  ·  `transicao_cancelar_id = 4`

> Os chamados de teste **GAAR‑60..73** foram criados de verdade neste portal e
> ficaram **Cancelados pelo Solicitante** (o portal do cliente não permite excluir).

## 6. Checklist pré‑envio (segunda de manhã)
- [ ] Token Jira do operador gerado (passo 2).
- [ ] Senha da conta Outlook **válida** (não vencida).
- [ ] Confirmado acesso à caixa compartilhada + ao portal GAAR.
- [ ] **DRY‑RUN** rodou no cliente e listou os Forms não lidos.
- [ ] **1 DEMO** rodou (chamado criado + cancelado, e‑mail movido + restaurado).
- [ ] Decidido quando ligar o modo **REAL**.

## 7. Observações
- **Modo REAL:** faz exatamente `criar chamado → encaminhar → mover p/ Finalizados`
  — as MESMAS operações já validadas no DEMO/produção (GAAR‑60..73); a única diferença
  é que **não desfaz** (não cancela o chamado nem volta o e‑mail). Ou seja, o REAL é o
  DEMO **sem o passo de desfazer**. Ainda assim, vale rodar **1 e‑mail real** e conferir
  antes de deixar automático. Ponto não estressado: processar **vários** e‑mails numa
  passada (é um laço sobre a operação de 1 e‑mail, que está validada).
- O app faz **login do zero a cada abertura** (logoff + login + MFA) — é por design
  ("nunca deixar conectado").
- Config com senha/token fica **só na máquina local** (`dados/config.xml`) — não
  versionar / não sincronizar em nuvem.
