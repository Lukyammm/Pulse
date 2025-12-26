/**
 * data.gs
 * Operações de dados protegidas.
 */

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
