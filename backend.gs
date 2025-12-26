/**
 * backend.gs
 * Pontos centrais de inicialização e rotas públicas do WebApp.
 */

const APP_META = {
  NAME: 'Pulse WebApp',
  VERSION: '1.0.0',
  TZ: 'America/Sao_Paulo',
  SHEETS: {
    USERS: 'USUARIOS',
    SOLICITACOES: 'SOLICITACOES',
    HISTORICO: 'HISTORICO_STATUS',
    LOGS: 'LOGS',
    CONFIG: 'CONFIG',
  },
  HEADERS: {
    USERS: ['email', 'senha_hash', 'perfil', 'ativo', 'criado_em', 'ultimo_login'],
    SOLICITACOES: [
      'id',
      'data_criacao',
      'solicitante_email',
      'perfil_solicitante',
      'nome_paciente',
      'prontuario',
      'leito',
      'data_nascimento',
      'nome_mae',
      'status',
      'data_status',
      'responsavel',
      'historico',
    ],
    HISTORICO: ['solicitacao_id', 'data', 'status', 'usuario', 'observacao'],
    LOGS: ['data', 'usuario', 'acao', 'detalhe'],
    CONFIG: ['chave', 'valor'],
  },
  STATUS: {
    ABERTO: 'ABERTO',
    RECEBIDO: 'RECEBIDO',
    PRODUCAO: 'EM_PRODUCAO',
    ENVIADO: 'ENVIADO',
    ENTREGUE: 'ENTREGUE',
    CONFIRMADO: 'CONFIRMADO',
    ENCERRADO: 'ENCERRADO',
    CANCELADO: 'CANCELADO',
  },
};

function doGet() {
  ensureSetup_();
  const template = HtmlService.createTemplateFromFile('ui');
  template.APP_NAME = APP_META.NAME;
  template.APP_VERSION = APP_META.VERSION;
  return template
    .evaluate()
    .setTitle(`${APP_META.NAME}`)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function ensureSetup_() {
  const ss = getDb_();
  ensureSheet_(ss, APP_META.SHEETS.USERS, APP_META.HEADERS.USERS);
  ensureSheet_(ss, APP_META.SHEETS.SOLICITACOES, APP_META.HEADERS.SOLICITACOES);
  ensureSheet_(ss, APP_META.SHEETS.HISTORICO, APP_META.HEADERS.HISTORICO);
  ensureSheet_(ss, APP_META.SHEETS.LOGS, APP_META.HEADERS.LOGS);
  ensureSheet_(ss, APP_META.SHEETS.CONFIG, APP_META.HEADERS.CONFIG);
  freezeHeaders_(ss);
  seedAdminUser_();
}

function freezeHeaders_(ss) {
  Object.values(APP_META.SHEETS).forEach((name) => {
    const sh = ss.getSheetByName(name);
    if (sh) sh.setFrozenRows(1);
  });
}

function getDb_() {
  return SpreadsheetApp.getActive();
}

function api_login(email, senha) {
  return login_(email, senha);
}

function api_logout(token) {
  return logout_(token);
}

function api_createSolicitacao(payload, token) {
  const session = validateSession_(token);
  return createSolicitacao_(payload, session);
}

function api_listSolicitacoes(token) {
  const session = validateSession_(token);
  return listSolicitacoes_(session);
}

function api_atualizarStatus(payload, token) {
  const session = validateSession_(token);
  return atualizarStatus_(payload, session);
}

function api_dashboard(token) {
  const session = validateSession_(token);
  return dashboard_(session);
}

function api_getSession(token) {
  return validateSession_(token, true);
}

function includeStyles_() {
  return HtmlService.createHtmlOutputFromFile('css').getContent();
}

function includeScripts_() {
  return HtmlService.createHtmlOutputFromFile('js').getContent();
}
