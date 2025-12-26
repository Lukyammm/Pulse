# Pulse (WebApp de Solicitação de Pulseiras)

## 1. Passo a passo de instalação
1. Abra o Google Drive e crie uma nova planilha chamada **Pulse**.
2. No menu `Extensões > App Script`, substitua o conteúdo do projeto pelos arquivos `Code.gs` e `Index.html` deste repositório.
3. Clique em **Publicar > Implantar como aplicativo da Web**.
   - Execute o aplicativo como: *Usuário que acessa*.
   - Quem pode acessar: **Somente na sua organização** (Workspace).
4. Salve a implantação e copie a URL do WebApp para uso interno.
5. Na planilha, rode o menu **Pulseiras > ✅ Setup** para criar abas e cabeçalhos.
6. Rode **Pulseiras > ➕ Seed Admin** para se registrar como ADM inicial.

## 2. Configuração de permissões
- O WebApp deve estar em um domínio Google Workspace; emails externos não acessam.
- A implantação precisa rodar **como o usuário que acessa** para que o backend valide o email real.
- Aba `USERS` controla o acesso: campos `email`, `perfil` (`ASSISTENCIA`, `RECEPCAO`, `ADM`), `setor`, `ativo`.
- Perfis são respeitados no backend: ASSISTÊNCIA só vê/atualiza o que criou no setor; RECEPÇÃO vê fila; ADM vê tudo.

## 3. Funções principais
- `setup()`: cria abas, headers e configura valores padrão.
- `seedMeAsAdmin()`: registra o usuário atual como ADM ativo.
- `apiGetMe()`: retorna o contexto autenticado e o perfil para o front.
- `apiCreateTicket()`, `apiListTickets()`, `apiUpdateStatus()`, `apiCancelTicket()`: CRUD seguro de tickets com regras por perfil.
- `apiDashboard()`: entrega KPIs filtrados conforme o perfil (ADM completo; Recepção geral; Assistência somente setor).
- `apiListUsers()`, `apiUpsertUser()`, `apiConfigGet()`, `apiConfigSet()`: gestão restrita ao perfil ADM.

## 4. Limitações conhecidas
- É necessário estar logado em uma conta do domínio para obter o email via `Session.getActiveUser`.
- Não há testes automatizados; valide o fluxo em um ambiente de homologação antes de produção.
- Sons de notificação dependem de autorização do navegador para áudio.
- O cache de dashboard é em planilha (`DASH_CACHE`); limpe manualmente se alterar cabeçalhos.
