# CONTEXTO PARA O CLAUDE — CVC-Trata-Forms (Gestão de Acessos)

> **Leia este arquivo primeiro.** Ele existe para que uma sessão nova do Claude, em
> **qualquer máquina**, se posicione sobre o projeto sem precisar reconstruir o
> histórico. Ele viaja pelo git (a pasta de memória do Claude é local de cada máquina).
>
> **Última atualização: 2026-08-14** — **Jira via API (REST)** adicionado sobre o HEAD `241e87b`.
> Ao concluir qualquer trabalho relevante, **atualize este arquivo** (seções 6 e 7).

---

## 1. O que é o projeto

Automação para o usuário **Nelson** (conta corporativa CVC) e a usuária final **Bruna**.

Para cada e-mail do remetente **Microsoft Forms** que esteja **não lido** na caixa
compartilhada **"Gestão de Acessos"** (Outlook Web), o app faz o ciclo:

1. Cria um chamado no **Jira Service Management** (portal do cliente)
2. Aplica no e-mail uma **categoria** com o nº do ticket (`JIRA - GAAR-99`)
3. **Encaminha** o e-mail para o grupo responsável, com o nº e o link do chamado
4. **Move** para a subpasta **"Finalizados"**
5. Marca como **lido**

Roda como **ícone de bandeja** (system tray), repetindo o ciclo a cada `intervalo_min`.

### Restrição fundamental (não tente contornar)

A conta CVC **não tem admin nem API token**. A **Microsoft Graph API foi descartada**:
o tenant exige consentimento de administrador para permissões de e-mail. Por isso tudo é
**automação de navegador**: **Edge + Playwright via CDP** (porta 9222), num perfil dedicado.
Os arquivos do "Plano A" (Graph) ficam arquivados na raiz: `auth.py`, `graph_client.py`,
`config.py`, `test_conexao.py` — **código morto, não são o caminho ativo**.

**Login:** o Outlook recebe e-mail+senha preenchidos pelo app; o **MFA é feito pelo usuário
na mão**; em "Continuar conectado?" a resposta é sempre **NÃO**. O Jira entra por **SSO
Microsoft** (nunca senha Atlassian). E-mail de login: `nelsondiniz@ext.cvccorp.com.br`.
Prefixo real dos chamados: **GAAR** (ex.: GAAR-57).

---

## 2. Arquitetura e arquivos

Branch de trabalho: **`Projeto_emails`**. Remote `origin` = `github.com/caracelli/cvc_Bruna`.
Python 3.14. Deps: `playwright`, `keyring`, `pystray`, `Pillow`, `uiautomation`, `requests`, `msal`.

**Existem dois conjuntos de código.** O ativo é `src/cvc_trata_forms/` (camadas). Os `.py`
soltos na raiz são legado/arquivados — **não edite lá achando que é o app**.

```
src/cvc_trata_forms/
  domain/regras.py
  application/processar_acessos.py          <- orquestra o ciclo (FASE 1 e, no DEMO, FASE 2)
  application/sessao/                       <- iniciar_sessoes / encerrar_sessoes / checar_login
  infrastructure/config/config_app.py       <- lê o config.xml em constantes
  infrastructure/credenciais/
  infrastructure/jira/chamado.py            <- cria o chamado e captura o código
  infrastructure/jira/login_jira.py
  infrastructure/outlook/acoes.py           <- categoria, mover, ler/não-ler, ENCAMINHAR
  infrastructure/outlook/login_outlook.py, outlook_web.py
  infrastructure/sistema/                   <- caminhos, diagnostico, fechar_popups, winutil
  presentation/bandeja.py                   <- MAIN (ícone de bandeja)
  presentation/form_config.py               <- formulário de configuração
```

**Entrypoints**
- `python run.py` — lançador com auto-instalação de deps → abre a bandeja
- `python -m cvc_trata_forms` — abre a bandeja
- `app.py` — entrypoint **único do .exe** (PyInstaller), 3 modos por argumento:
  nada = bandeja, `--config` = formulário, `--fechar-popups` = loop do fechador de popups do Edge
- Spec do PyInstaller: `CVC-Trata-Forms.spec`

**Configuração** — `config.xml` na raiz (e `dados\config.xml` ao lado do .exe no cliente).
**Contém senha em texto puro; está no `.gitignore`; NÃO versionar.** Modelo comentado:
`config.example.xml`. Blocos: `<outlook>`, `<jira>` (`url_chamado`, `categoria_prefixo`,
`prefixo_chamado`, `cor_titulo`), `<fluxo>` (`remetente_filtro`, `subpasta_destino`,
`so_nao_lidos`, `dry_run`, `demo`, `demo_qtd`, `intervalo_min`, `invisivel`),
`<encaminhamento>` (`ativo`, `destinatarios`, `assunto_template`, `prefixo_corpo`, `separador`).

⚠️ **Armadilha do formulário:** `form_config._salvar` **reescreve o config.xml inteiro a partir
do template embutido no próprio `form_config.py`**. Campo que não estiver no template **é perdido
ao salvar**. Ao criar uma opção nova no config, adicione-a *também* no `_ler` e no template do form.

**Modos de execução** (prioridade: **DEMO > DRY-RUN**):
- **DEMO self-cleaning** — faz o ciclo REAL e depois **DESFAZ** (cancela o chamado, devolve o
  e-mail para a Inbox como não lido). É o modo de demonstração para a Bruna. `demo_qtd` = quantos
  e-mails processa.
- **DRY-RUN** — não executa as ações.
- **REAL** — desmarcar os dois no formulário (aba Fluxo tem checkbox de DEMO).

**Build:** `powershell -ExecutionPolicy Bypass -File build.ps1`. Usa `packaging\config.xml`
(gitignored) como config que vai dentro do pacote; se não existir, cai no `config.example.xml`
(sem senha). Valida com `--check` e gera o zip. O zip é **force-added** (`git add -f`, pois o
`.gitignore` tem `*.zip`). Ordem correta: **commit do código → build → commit do zip**, para o
exe ter proveniência rastreável.

**Antes de qualquer push:** `ruff check src tests`, `python -m compileall -q src`, `pytest -q`
(hoje 5/5). O CI do GitHub roda ruff; as regras estão fixadas em `pyproject.toml`
(`[tool.ruff.lint] select = E4/E5/E7/E9/F`) — sem isso o default novo do ruff acusa ~224 erros
de estilo e o CI fica vermelho mandando e-mail a cada push.

---

## 3. Aprendizados que NÃO podem ser perdidos

Cada item aqui custou uma sessão inteira de depuração. Antes de "melhorar" algum desses
trechos, entenda por que ele está do jeito que está.

**O Jira agora pode ser via API (REST) — mata a parte mais frágil.** Se o `config.xml`
bloco `<jira>` tem `<usuario>` + `<token>` (constante `JIRA_API_ATIVA`), o app cria e
cancela o chamado pela **JSM REST API** — SEM abrir a aba do Jira, SEM login SSO, SEM
formulário/scraping/tela-pequena/Enter-fallback. Módulo: `infrastructure/jira/chamado_api.py`
(`criar_chamado(dados)->(codigo,link)` e `cancelar_chamado(codigo)`). Parâmetros apurados na
API e validados ao vivo (GAAR-60..62, incl. DEMO completo): serviceDesk **1984**, requestType
**7550**, campos obrigatórios só `summary`+`description`, transição de cancelamento id **4**
("Cancelado pelo Solicitante"). **Token é POR USUÁRIO** (campo no form de config; cada um gera
o seu em id.atlassian.com/manage-profile/security/api-tokens). **Sem token, cai no fluxo de
navegador** (`chamado.py` + `login_jira.py` + PASSO 2 do `iniciar_sessoes`), que continua no
código como **FALLBACK — não apagar**. A restrição "sem API token" do topo é do **Graph
(e-mail)**, NÃO do Jira: o token do Jira é pessoal e não exige admin.

**A máquina do cliente é um notebook de TELA PEQUENA.** Essa foi a causa raiz de vários bugs
que na máquina de dev (tela grande) não reproduziam: o OWA é responsivo e reposiciona tudo.
`iniciar_sessoes._forcar_layout_amplo(page)` força layout desktop via CDP
`Emulation.setDeviceMetricsOverride({width:1600,height:900})`, aplicado no Outlook e no Jira.
Kill-switch: env `CVC_VIEWPORT_OVERRIDE=off`. Para reproduzir tela pequena na dev:
`CVC_WINDOW_SIZE=1024,600`. O log `[layout]` no `cvc_bandeja.log` mostra o `innerWidth` real.

**O clique programático "dá certo" sem submeter, na máquina do cliente.** Vale para o Jira e
para o envio do encaminhamento. Consequências no design:
- No Jira (`chamado.py`): o fallback de **Enter** dispara pela **ausência do código** (o próximo
  passo não acontecer), *não* pela exceção do clique. `_clicar_enviar` só clica;
  `_aguardar_codigo` faz o poll; se o código não vier em ~12s, `_enviar_por_enter` entra.
  Na dev o código aparece direto e o Enter **não** dispara — sem risco de chamado duplicado.
- No encaminhamento (`acoes._enviar_compose`): sucesso = **o compose FECHAR**. Escada:
  clique → `Ctrl+Enter` (com foco no corpo) → `click()` pelo DOM → senão salva diagnóstico
  `encaminhar_nao_enviou` e erra. **No cliente quem envia é o `Ctrl+Enter`.**

**O corpo do encaminhamento é inserido numa ÚNICA operação** (`execCommand insertHTML` em
`acoes._escrever_topo`). Digitar tecla por tecla dava tempo do editor do OWA re-renderizar a
assinatura e **comer a primeira linha**. Digitação virou fallback (`forcar_teclado=True`),
usada só para repor uma linha faltante.

**Não existe mais Ctrl+K.** Era ele que colava o hyperlink na palavra errada ("riginal", de
"Mensagem original") — a seleção era por contagem de teclas e partia de posição diferente
conforme o layout. Hoje a **âncora vai montada no HTML** (o OWA não auto-linka HTML inserido
pronto). Formato do corpo: `<prefixo> <ticket>` + `Link chamado: <a href=url>url</a>`.

**`_conferir_corpo()` só registra, não barra o envio** — e **remove a URL do texto antes de
procurar o número**, porque o ticket aparece dentro da própria URL (sem isso a conferência
ficava cega).

**Popup "Lembrete de anexo"** no envio: `_confirmar_popup_enviar()` espera até 4s, casa o botão
**por conteúdo** (nunca "não enviar") e loga texto + botões. Confirmado no cliente:
`botoes=['enviar','não enviar'] → cliquei em enviar`. Tratado.

**Separador vazio no config significa SEM separador** (commit `b5beb1e`) — é intencional.

**Não re-rodar o ciclo imediatamente quando ele falha.** Já foi tentado (um `continue` em
`bandeja.loop_monitor`) e virou **loop criando chamados** — a falha era na FASE 2 do DEMO, e o
retry refazia o ciclo inteiro. Foi revertido, com comentário no código explicando. O auto-disparo
ao abrir **já funciona** pelo 1º ciclo após o login.

**"2 e-mails por chamado" é correto:** o app envia 1 (o encaminhamento); o 2º é a notificação
automática do Jira ("solicitação criada"). Em teste os dois caem no usuário porque ele é o
solicitante e o destinatário.

**Popups do Edge = não-issue.** Não interferem no ciclo, e em produção o Edge roda invisível.
O fechador (`--fechar-popups --loop`) já roda como rede de segurança. Decidido não investir mais.

**Diagnóstico automático:** `infrastructure/sistema/diagnostico.py` salva **print + HTML + txt**
em `dados\diagnostico\` quando um passo crítico falha (`form_nao_carregou`,
`codigo_nao_capturado`, `mover_inbox_falhou`, `encaminhar_nao_enviou`). É o que substitui o
acesso à máquina do cliente. Ao investigar um problema relatado, **peça essa pasta + o
`cvc_bandeja.log`**.

---

## 4. Como testar nesta base (armadilhas)

⚠️ **Rodar o app do código-fonte estava quebrado na máquina do Nelson**: erro
`Playwright Sync API inside the asyncio loop`. O **.exe funciona normal**. Para testar ao vivo:
use `dist\CVC-Trata-Forms\CVC-Trata-Forms.exe`, copiando o `config.xml` da raiz para
`dist\CVC-Trata-Forms\dados\`.

- **Ao terminar o teste, MATE o processo** — em DEMO ele repete o ciclo a cada 5 min e enche a
  caixa de chamados.
- **Espere a FASE 2 do DEMO terminar** antes de matar. Matar por tempo fixo já deixou um e-mail
  preso em "Finalizados".
- Para limpar o Edge da automação: `fechar_edge_automacao()` em `encerrar_sessoes.py` — mata só
  o `msedge.exe` do perfil `cvc_trata_forms`, não toca no Edge pessoal.
- Padrão de simulação sem OWA real (usado várias vezes, funciona bem): montar uma página HTML
  falsa com os mesmos `aria-label` do OWA e rodar no Edge headless. Valida fluxo de controle
  (escada de envio, inserção do corpo, popup) sem tocar em e-mail de verdade.

---

## 5. Estado atual (2026-08-14)

- **Jira via API (REST) implementado e validado ao vivo** (create+cancel, incl. DEMO completo
  GAAR-62; sem aba/login do Jira). `ruff`/`pytest 5/5`/`compileall` OK. Fluxo de navegador
  mantido como fallback. Ver seção 3.
- **HEAD `241e87b`**, branch `Projeto_emails`, working tree **limpo**, sincronizado com o origin.
- Tag âncora da entrega: **`entrega-bruna-2026-07-31`**.
- Pacotes entregues no origin (buildados de `1de2bdf`):
  - `CVC-Trata-Forms.zip` — pacote da **Bruna**: login e destinatários **vazios** (ela cadastra
    no 1º uso; a bandeja detecta config faltando, avisa, abre o formulário e reinicia o app)
  - `CVC-Trata-Forms-TESTE.zip` — config com os dados do Nelson
  - Ambos em **DEMO**, `demo_qtd=1`, Edge **visível**
- ✅ **O teste do encaminhamento na máquina do cliente foi CONFIRMADO OK pelo usuário em
  13/08/2026.** Era o último item que travava a retomada. **O fluxo de encaminhamento está
  validado em produção — não mexer nele sem motivo.**

**Como a Bruna usa:** `git pull` na `Projeto_emails` → descompacta o zip → roda o
`CVC-Trata-Forms.exe` → aviso de "faltam dados" → preenche e-mail, senha e destinatários no
formulário → Salvar → o app reinicia sozinho e roda o DEMO.

---

## 6. Onde paramos / próximos passos

Em ordem de prioridade:

0. **Jira via API — fechamento (feito o essencial, falta validar o .exe e a PRD).** O código
   está pronto e validado do FONTE. Falta: (a) rodar o **.exe** com token para validar o
   artefato (CONTEXTO diz que a validação de verdade é pelo exe); (b) o pacote sai em
   **DRY-RUN** com `<usuario>`/`<token>` do Jira **vazios** (token por usuário, via form) —
   ao ir para produção, cada usuário informa o token dele no formulário e desmarca DRY-RUN.
1. ⚠️ **TROCAR A SENHA da conta CVC.** O `CVC-Trata-Forms-TESTE.zip` versionado tem a senha em
   **texto puro** e ela está no histórico do GitHub `caracelli/cvc_Bruna` **para sempre**
   (versionar assim foi decisão explícita do usuário, avisado duas vezes). Com o teste
   concluído, trocar a senha é o encerramento correto. Depois, oferecer remover o arquivo do repo.
2. **Rebuild de proveniência** — os zips são de `1de2bdf`, o HEAD é `241e87b` (a diferença é só
   renome de variável + imports para o lint; comportamento idêntico). Rebuildar para exe = HEAD.
3. **Ir para produção** — com o DEMO validado, desmarcar **DEMO + DRY-RUN** no formulário (ou
   rebuild com `demo=false`) e considerar `invisivel=true`. Sugestão em aberto: fazer 2 e 3 num
   único rebuild já em modo produção.
4. **Auto-e-mail de monitoramento** — avisar por e-mail quando um ciclo falhar. O usuário quer;
   **ainda não implementado**.
5. **Instabilidade do "Mover" em sessão FRIA** (`acoes.py`) — no `.exe` (que faz login do
   zero), a árvore do diálogo de mover popula devagar e o `sleep(2)` fixo lia a árvore vazia:
   "Pasta 'Finalizados' não encontrada". **CORRIGIDO (2026-08-14):** `mover_email` e
   `mover_para_inbox` agora usam `_poll_treeitem` (aguarda o item aparecer, ~15s, com fallback
   relaxado) e `_aguardar_dialogo_fechar` no lugar dos sleeps fixos. **Validar no `.exe`**
   (o warm/fonte não reproduz o bug). **VALIDADO ao vivo (DEMO GAAR-64):** move p/
   Finalizados + volta OK. Correlato **CORRIGIDO+VALIDADO:** `abrir_inbox_compartilhada`
   não abria a caixa compartilhada quando ela vinha RECOLHIDA (``) numa sessão fria —
   agora `_expandir_caixa` (expande o cabeçalho da caixa) + `_carregar_mais_pastas` (clica
   "Carregar mais pastas") + poll ~40s; testado no exato estado que falhava (abriu a Inbox
   e revelou Finalizados). Mesmo padrão em `abrir_subpasta`.

**Pendências do usuário, não dá para fazer daqui:** renomear o repo GitHub `cvc_Bruna` para o
nome definitivo (o `gh` não está logado; depois é `git remote set-url origin <nova-url>`) e
renomear a pasta do projeto no disco (exige VS Code fechado + recriar o `.venv` com
`build.ps1 -RecreateVenv`, que tem caminhos absolutos).

**Ideia oferecida e não implementada:** lembrar, dentro da sessão, que o clique é inútil naquela
máquina e ir direto para o `Ctrl+Enter` — economiza ~12s por e-mail.

---

## 7. Regra de manutenção deste arquivo

Ao terminar qualquer trabalho relevante nesta base:

1. Atualize a **seção 5** (estado/HEAD) e a **seção 6** (próximos passos)
2. Se descobrir uma causa raiz nova ou um comportamento contra-intuitivo, registre na
   **seção 3** — é ela que evita refazer depuração já feita
3. Troque a data do cabeçalho e commite junto com a mudança que a motivou
