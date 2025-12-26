# Pulseiras WebApp (Google Apps Script)

WebApp para solicitar, acompanhar e auditar pedidos de pulseiras a partir de uma planilha do Google Sheets. Inclui controle de acesso por perfil (ASSISTENCIA, RECEPCAO e ADM), fluxo completo de tickets e painel resumido.

## Passo a passo de instalação
1. Crie uma planilha nova no Google Sheets (Google Workspace).
2. Abra **Extensões → Apps Script** e copie todo o conteúdo dos arquivos `Code.gs` e `Index.html` para o projeto.
3. No Apps Script, defina o fuso horário correto em **Configurações do projeto** (ex.: `America/Fortaleza`).
4. (Opcional standalone) Em **Propriedades do projeto → Propriedades do script**, defina `DB_SPREADSHEET_ID` com o ID da planilha, caso o script não esteja vinculado.
5. Volte à planilha e, no menu, clique em **Pulseiras → ✅ Setup (criar abas / headers / config)** para criar abas e preencher a `CONFIG` com valores padrão.
6. Ainda no menu, clique em **Pulseiras → ➕ Seed Admin (me tornar ADM)** para inserir seu usuário como administrador ativo.
7. Na aba **USERS**, cadastre assistência, recepção e outros admins (campos obrigatórios: email, nome, perfil, setor, ativo).
8. Publique o WebApp em **Implantar → Implantar como aplicativo da web**, com **Executar como: usuário acessando** e permissões para quem tiver o link ou domínio.
9. Abra o URL do WebApp para usar a interface.

## Configuração de permissões
- O WebApp deve ser publicado com a opção **Executar como usuário acessando** para que `Session.getActiveUser()` traga o e-mail correto.
- Usuários precisam estar cadastrados na aba `USERS` com `ativo = TRUE` e perfil válido (`ASSISTENCIA`, `RECEPCAO` ou `ADM`).
- Planilha deve estar no mesmo domínio do Google Workspace dos usuários.
- Funções administrativas (`apiConfig*`, `apiUpsertUser`, `apiDashboard`) exigem perfil `ADM`.

## Estrutura da planilha
As abas essenciais são criadas pelo `setup()`:
- `CONFIG` (`key | value | updated_at | updated_by`)
- `USERS` (`email | nome | perfil | setor | ativo | updated_at | updated_by`)
- `TICKETS` (fluxo completo do pedido; preenchida apenas pelo sistema)
- `LOGS` (auditoria imutável; preenchida apenas pelo sistema)
- `DASH_CACHE` (cache opcional para KPIs)

## Funções principais
- `doGet()`: renderiza o WebApp com `Index.html`.
- `setup()`: cria abas, cabeçalhos e valores padrão em `CONFIG`.
- `seedMeAsAdmin()`: insere o usuário atual como `ADM` no `USERS`.
- APIs públicas (`apiCreateTicket`, `apiUpdateStatus`, `apiCancelTicket`, `apiListTickets`, `apiDashboard`, etc.) fazem validação de perfil, normalização e registro em `LOGS`.
- Repositório de dados (`sheetRepo_`) centraliza leitura/escrita nas abas, incluindo controle de concorrência via `LockService` nas operações críticas.

## Limitações conhecidas
- Dependência total do Google Sheets; sem conexão, o WebApp não funciona.
- `Session.getActiveUser()` requer Google Workspace e deploy correto; fora disso, o e-mail pode não ser obtido.
- Não há controle de versão de UI no deploy; atualizações exigem nova implantação do WebApp.
- Não há notificações por e-mail ou mobile push embutidas.
