/**
 * code.gs
 * Backend unificado do Pulse WebApp.
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

const PERMISSIONS = {
  ASSISTENCIA: {
    create: true,
    viewSelf: true,
    updateStatus: ['CONFIRMADO'],
    dashboard: false,
    manageUsers: false,
  },
  RECEPCAO: {
    create: false,
    viewAll: true,
    updateStatus: ['RECEBIDO', 'EM_PRODUCAO', 'ENVIADO'],
    dashboard: false,
    manageUsers: false,
  },
  ADM: {
    create: true,
    viewAll: true,
    updateStatus: ['RECEBIDO', 'EM_PRODUCAO', 'ENVIADO', 'ENTREGUE', 'CONFIRMADO', 'ENCERRADO', 'CANCELADO'],
    dashboard: true,
    manageUsers: true,
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

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  }
  const existingHeaders = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  if (existingHeaders.join('|') !== headers.join('|')) {
    sh.clear();
    sh.appendRow(headers);
  }
  return sh;
}

function seedAdminUser_() {
  const ss = getDb_();
  const sh = ss.getSheetByName(APP_META.SHEETS.USERS);
  const data = sh.getDataRange().getValues();
  if (data.length > 1) return;
  const email = 'admin@local';
  const senha = 'admin123';
  const senhaHash = hashPassword_(senha);
  sh.appendRow([email, senhaHash, 'ADM', true, new Date(), '']);
}

function hashPassword_(senha) {
  const salt = getSalt_();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, senha + salt);
  return Utilities.base64Encode(digest);
}

function getSalt_() {
  const props = PropertiesService.getScriptProperties();
  let salt = props.getProperty('APP_SALT');
  if (!salt) {
    salt = Utilities.getUuid();
    props.setProperty('APP_SALT', salt);
  }
  return salt;
}

function login_(email, senha) {
  const ss = getDb_();
  const sh = ss.getSheetByName(APP_META.SHEETS.USERS);
  const rows = sh.getDataRange().getValues();
  const headers = rows.shift();
  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));
  const user = rows.find((r) => r[idx.email] === email);
  if (!user) throw new Error('Usuário não encontrado.');
  if (String(user[idx.ativo]).toLowerCase() !== 'true') throw new Error('Usuário inativo.');
  const senhaHash = hashPassword_(senha);
  if (senhaHash !== user[idx.senha_hash]) throw new Error('Credenciais inválidas.');
  const perfil = user[idx.perfil];
  const token = Utilities.getUuid();
  const payload = {
    email,
    perfil,
    issued_at: Date.now(),
  };
  CacheService.getScriptCache().put(`session_${token}`, JSON.stringify(payload), 1800);
  sh.getRange(rows.indexOf(user) + 2, idx.ultimo_login + 1).setValue(new Date());
  logAction_(email, 'LOGIN', 'Acesso concedido');
  return {
    token,
    perfil,
    email,
  };
}

function validateSession_(token, silent) {
  if (!token) throw new Error('Sessão expirada.');
  const cache = CacheService.getScriptCache();
  const raw = cache.get(`session_${token}`);
  if (!raw) {
    if (silent) return null;
    throw new Error('Sessão inválida.');
  }
  const session = JSON.parse(raw);
  cache.put(`session_${token}`, raw, 1800);
  return session;
}

function logout_(token) {
  const cache = CacheService.getScriptCache();
  cache.remove(`session_${token}`);
  return true;
}

function ensurePermission_(session, action, allowedStatuses) {
  const profile = session.perfil;
  const rules = PERMISSIONS[profile];
  if (!rules) throw new Error('Perfil sem permissão.');
  if (action === 'create' && !rules.create) throw new Error('Sem permissão para criar.');
  if (action === 'dashboard' && !rules.dashboard) throw new Error('Sem permissão para dashboard.');
  if (action === 'viewAll' && !rules.viewAll && !rules.viewSelf) throw new Error('Sem permissão para visualizar.');
  if (action === 'status') {
    if (!rules.updateStatus || rules.updateStatus.indexOf(allowedStatuses) === -1) {
      throw new Error('Status não autorizado.');
    }
  }
}

function createSolicitacao_(payload, session) {
  ensurePermission_(session, 'create');
  const ss = getDb_();
  const sh = ss.getSheetByName(APP_META.SHEETS.SOLICITACOES);
  const id = Utilities.getUuid();
  const now = new Date();
  const historico = JSON.stringify([{ status: APP_META.STATUS.ABERTO, data: now, usuario: session.email, obs: 'Criado' }]);
  const row = [
    id,
    now,
    session.email,
    session.perfil,
    payload.nome_paciente,
    payload.prontuario,
    payload.leito,
    payload.data_nascimento,
    payload.nome_mae || '',
    APP_META.STATUS.ABERTO,
    now,
    session.email,
    historico,
  ];
  sh.appendRow(row);
  logAction_(session.email, 'CRIAR_SOLICITACAO', id);
  addHistorico_(id, APP_META.STATUS.ABERTO, session.email, 'Criado');
  return { id, status: APP_META.STATUS.ABERTO };
}

function listSolicitacoes_(session) {
  const ss = getDb_();
  const sh = ss.getSheetByName(APP_META.SHEETS.SOLICITACOES);
  const data = sh.getDataRange().getValues();
  const headers = data.shift();
  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));
  const rows = data
    .map((r) => ({
      id: r[idx.id],
      data_criacao: r[idx.data_criacao],
      solicitante_email: r[idx.solicitante_email],
      perfil_solicitante: r[idx.perfil_solicitante],
      nome_paciente: r[idx.nome_paciente],
      prontuario: r[idx.prontuario],
      leito: r[idx.leito],
      data_nascimento: r[idx.data_nascimento],
      nome_mae: r[idx.nome_mae],
      status: r[idx.status],
      data_status: r[idx.data_status],
      responsavel: r[idx.responsavel],
      historico: r[idx.historico] ? JSON.parse(r[idx.historico]) : [],
    }))
    .filter((item) => {
      if (PERMISSIONS[session.perfil]?.viewAll) return true;
      if (PERMISSIONS[session.perfil]?.viewSelf) return item.solicitante_email === session.email;
      return false;
    });
  return rows;
}

function atualizarStatus_(payload, session) {
  ensurePermission_(session, 'status', payload.status);
  const ss = getDb_();
  const sh = ss.getSheetByName(APP_META.SHEETS.SOLICITACOES);
  const data = sh.getDataRange().getValues();
  const headers = data.shift();
  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));
  const rowIndex = data.findIndex((r) => r[idx.id] === payload.id);
  if (rowIndex === -1) throw new Error('Solicitação não encontrada.');
  const sheetRow = rowIndex + 2;
  const historico = data[rowIndex][idx.historico] ? JSON.parse(data[rowIndex][idx.historico]) : [];
  historico.push({ status: payload.status, data: new Date(), usuario: session.email, obs: payload.observacao || '' });
  sh.getRange(sheetRow, idx.status + 1).setValue(payload.status);
  sh.getRange(sheetRow, idx.data_status + 1).setValue(new Date());
  sh.getRange(sheetRow, idx.responsavel + 1).setValue(session.email);
  sh.getRange(sheetRow, idx.historico + 1).setValue(JSON.stringify(historico));
  addHistorico_(payload.id, payload.status, session.email, payload.observacao || '');
  logAction_(session.email, 'ATUALIZAR_STATUS', `${payload.id}:${payload.status}`);
  return { id: payload.id, status: payload.status };
}

function addHistorico_(id, status, usuario, observacao) {
  const ss = getDb_();
  const sh = ss.getSheetByName(APP_META.SHEETS.HISTORICO);
  sh.appendRow([id, new Date(), status, usuario, observacao]);
}

function logAction_(usuario, acao, detalhe) {
  const ss = getDb_();
  const sh = ss.getSheetByName(APP_META.SHEETS.LOGS);
  sh.appendRow([new Date(), usuario, acao, detalhe]);
}

function dashboard_(session) {
  ensurePermission_(session, 'dashboard');
  const dados = listSolicitacoes_(session);
  const total = dados.length;
  const porStatus = dados.reduce((acc, item) => {
    acc[item.status] = (acc[item.status] || 0) + 1;
    return acc;
  }, {});
  const porPerfil = dados.reduce((acc, item) => {
    acc[item.perfil_solicitante] = (acc[item.perfil_solicitante] || 0) + 1;
    return acc;
  }, {});
  return { total, porStatus, porPerfil };
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
