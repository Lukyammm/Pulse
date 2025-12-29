# Pulse WebApp (Solicitação de Pulseiras)

## 1. Passo a passo de instalação
1. Crie uma planilha Google chamada **Pulse**.
2. Em `Extensões > App Script`, substitua o projeto pelos dois arquivos deste repositório (`code.gs` e `ui.html`).
3. Publique em `Implantar > Implantação nova` como **Aplicativo da Web**:
   - Executar como: **Usuário que acessa**.
   - Quem pode acessar: **Sua organização**.
4. Abra a URL do WebApp; o `ensureSetup_()` roda automaticamente e cria abas/cabeçalhos, congela os títulos e grava configurações padrão.
5. Se quiser forçar o processo manualmente, rode o menu `Pulse WebApp > Setup` (se container-bound) ou execute `ensureSetup_()` pelo editor.
6. A conta `admin@local` é criada na primeira execução com senha `admin123`; altere imediatamente na aba `USERS`.

## 2. Configuração de permissões
- As abas criadas são: `USUARIOS`, `SOLICITACOES`, `HISTORICO_STATUS`, `LOGS`, `CONFIG`.
- Perfis disponíveis: `ASSISTENCIA`, `RECEPCAO`, `ADM`.
- Backend valida todo acesso: usuários sem `ativo=true` ou inexistentes são bloqueados.
- Sessões duram 30 minutos (cache) e são renovadas a cada chamada.
- O front-end apenas monta a UI conforme perfil retornado; permissões reais são checadas no servidor.

## 3. Funções principais
- `ensureSetup_()`: cria abas com cabeçalhos corretos e congela linhas de título.
- `api_login(email, senha)`: autentica com hash SHA-256 salgado e retorna token de sessão.
- `api_createSolicitacao(payload, token)`: criação restrita a ASSISTENCIA/ADM.
- `api_listSolicitacoes(token)`: retorna solicitações permitidas pelo perfil (próprias ou todas).
- `api_atualizarStatus(payload, token)`: altera status conforme matriz de permissões.
- `api_dashboard(token)`: KPIs de ADM.
- `api_logout(token)`: encerra sessão removendo token do cache.

## 4. Limitações conhecidas
- Para obter o email real, o WebApp deve rodar em domínio Workspace com execução como usuário que acessa.
- A senha inicial `admin123` precisa ser trocada manualmente atualizando `senha_hash` para o novo hash (use `hashPassword_()` via editor).
- Não há testes automatizados; valide em ambiente controlado.
- Sons/alertas não são carregados (implementação focada em backend seguro e UI Apple-like).
- O layout mobile foi otimizado para telas menores, mas continua com a mesma base funcional do desktop.

## 5. Notas de UI
- A lista de solicitações ocupa toda a largura por padrão; o botão **➕ Nova solicitação** abre um painel lateral (drawer) com o formulário completo, com campos agrupados em duas colunas para preenchimento mais rápido.
- Quando não há tickets, a lista exibe um estado vazio com instruções e CTA para criar a primeira solicitação (apenas para perfis com permissão de criação).
