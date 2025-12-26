/** ============================================================================
 *  SISTEMA DE SOLICITAÇÃO DE PULSEIRAS (Apps Script WebApp)
 *  Perfis: ASSISTENCIA | RECEPCAO | ADM
 *  Abas: CONFIG | USERS | TICKETS | LOGS | DASH_CACHE
 *  Autor: você mandou eu entregar pronto, então aqui vai.
 *  ============================================================================ */

const APP = {
  NAME: "Pulseiras",
  VERSION: "1.0.0",
  TZ: Session.getScriptTimeZone() || "America/Fortaleza",
  SHEETS: {
    CONFIG: "CONFIG",
    USERS: "USERS",
    TICKETS: "TICKETS",
    LOGS: "LOGS",
    DASH_CACHE: "DASH_CACHE",
  },
  // Cabeçalhos (linha 1)
  HEADERS: {
    CONFIG: ["key", "value", "updated_at", "updated_by"],
    USERS: ["email", "nome", "perfil", "setor", "ativo", "updated_at", "updated_by"],
    TICKETS: [
      "ticketId",
      "status_atual",
      "paciente_nome",
      "prontuario",
      "leito",
      "nascimento", // ISO yyyy-mm-dd
      "crianca", // TRUE/FALSE
      "mae_nome",
      "solicitante_email",
      "solicitante_nome",
      "setor",
      "recepcao_responsavel",
      "criado_em",
      "recebido_em",
      "em_producao_em",
      "enviado_em",
      "entregue_em",
      "confirmado_em",
      "encerrado_em",
      "cancelado_em",
      "cancelado_motivo",
      "last_updated_em",
      "last_updated_por",
    ],
    LOGS: ["logId", "ticketId", "acao", "de_status", "para_status", "usuario_email", "timestamp", "detalhes_json"],
    DASH_CACHE: ["cache_key", "payload_json", "updated_at"],
  },
  STATUS: {
    ABERTO: "ABERTO",
    RECEBIDO: "RECEBIDO_PELA_RECEPCAO",
    PRODUCAO: "EM_PRODUCAO",
    ENVIADO: "ENVIADO_A_CAMINHO",
    ENTREGUE: "ENTREGUE_DEIXADO",
    CONFIRMADO: "RECEBIDO_PELA_ASSISTENCIA",
    ENCERRADO: "ENCERRADO",
    CANCELADO: "CANCELADO",
  },
  // Status considerados "em aberto" (pra filas/aging)
  OPEN_STATUSES() {
    return [
      APP.STATUS.ABERTO,
      APP.STATUS.RECEBIDO,
      APP.STATUS.PRODUCAO,
      APP.STATUS.ENVIADO,
      APP.STATUS.ENTREGUE,
      APP.STATUS.CONFIRMADO, // (vai auto-encerrar em seguida, mas fica aqui por segurança)
    ];
  },
  // Timestamps por status
  STATUS_TS_COL: {
    ABERTO: "criado_em",
    RECEBIDO_PELA_RECEPCAO: "recebido_em",
    EM_PRODUCAO: "em_producao_em",
    ENVIADO_A_CAMINHO: "enviado_em",
    ENTREGUE_DEIXADO: "entregue_em",
    RECEBIDO_PELA_ASSISTENCIA: "confirmado_em",
    ENCERRADO: "encerrado_em",
    CANCELADO: "cancelado_em",
  },
  DEFAULT_CONFIG: {
    SLA_TOTAL_MIN: "30",
    SLA_RECEPCAO_MIN: "5",
    SLA_ENVIAR_MIN: "15",
    POLL_MS_RECEPCAO: "4000",
    SOUND_ENABLED_DEFAULT: "true",
  },
};

function doGet() {
  const t = HtmlService.createTemplateFromFile("Index");
  t.APP_NAME = APP.NAME;
  t.APP_VERSION = APP.VERSION;
  return t
    .evaluate()
    .setTitle(`${APP.NAME} • WebApp`)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** ========================= MENU / SETUP ========================= **/

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu(APP.NAME)
      .addItem("✅ Setup (criar abas / headers / config)", "setup")
      .addItem("➕ Seed Admin (me tornar ADM)", "seedMeAsAdmin")
      .addSeparator()
      .addItem("🧪 Mock: gerar tickets", "mockGenerateTickets")
      .addToUi();
  } catch (e) {
    // Ignora se não estiver em container-bound
  }
}

function setup() {
  const ss = getDb_();
  ensureSheetWithHeaders_(ss, APP.SHEETS.CONFIG, APP.HEADERS.CONFIG);
  ensureSheetWithHeaders_(ss, APP.SHEETS.USERS, APP.HEADERS.USERS);
  ensureSheetWithHeaders_(ss, APP.SHEETS.TICKETS, APP.HEADERS.TICKETS);
  ensureSheetWithHeaders_(ss, APP.SHEETS.LOGS, APP.HEADERS.LOGS);
  ensureSheetWithHeaders_(ss, APP.SHEETS.DASH_CACHE, APP.HEADERS.DASH_CACHE);

  // Config defaults
  const cfg = sheetRepo_(ss).config;
  Object.keys(APP.DEFAULT_CONFIG).forEach((k) => {
    if (!cfg.get(k)) cfg.set(k, APP.DEFAULT_CONFIG[k], "system@setup");
  });

  // Congelar headers
  Object.values(APP.SHEETS).forEach((name) => {
    const sh = ss.getSheetByName(name);
    if (sh) sh.setFrozenRows(1);
  });

  return true;
}

function seedMeAsAdmin() {
  const me = getUserEmail_();
  if (!me) throw new Error("Não consegui ler seu e-mail. Isso precisa estar em Google Workspace e o WebApp deve rodar como 'usuário acessando'.");
  const ss = getDb_();
  const users = sheetRepo_(ss).users;
  users.upsert({
    email: me,
    nome: "ADMIN",
    perfil: "ADM",
    setor: "GLOBAL",
    ativo: true,
  }, me);
  return true;
}

/** ========================= API PUBLICA (google.script.run) ========================= **/

function apiGetMe() {
  return wrap_(function () {
    const ctx = authContext_();
    return { me: ctx.me, profile: ctx.me.perfil, setor: ctx.me.setor, now: new Date().toISOString() };
  });
}

function apiListTickets(filters) {
  return wrap_(function () {
    const ctx = authContext_();
    filters = filters || {};
    const repo = sheetRepo_(getDb_());
    const rows = repo.tickets.list(filters, ctx);
    return { tickets: rows };
  });
}

function apiGetTicket(ticketId) {
  return wrap_(function () {
    const ctx = authContext_();
    const repo = sheetRepo_(getDb_());
    const ticket = repo.tickets.get(ticketId);
    if (!ticket) throw new Error("Ticket não encontrado.");
    enforceTicketAccess_(ctx, ticket);
    const logs = repo.logs.listByTicket(ticketId);
    return { ticket, logs };
  });
}

function apiCreateTicket(payload) {
  return wrap_(function () {
    const ctx = authContext_();
    requireRole_(ctx, ["ASSISTENCIA", "ADM"]);
    payload = normalizeTicketPayload_(payload || {});
    validateTicketPayload_(payload);

    // setor vem do usuário (segurança)
    payload.setor = ctx.me.setor;
    payload.solicitante_email = ctx.me.email;
    payload.solicitante_nome = ctx.me.nome || ctx.me.email;

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const repo = sheetRepo_(getDb_());
      const created = repo.tickets.create(payload, ctx.me.email);
      repo.logs.append({
        ticketId: created.ticketId,
        acao: "CREATE",
        de_status: "",
        para_status: created.status_atual,
        usuario_email: ctx.me.email,
        detalhes: { payload: payload },
      });
      return { ticketId: created.ticketId };
    } finally {
      lock.releaseLock();
    }
  });
}

function apiUpdateStatus(ticketId, toStatus, extra) {
  return wrap_(function () {
    const ctx = authContext_();
    extra = extra || {};
    const repo = sheetRepo_(getDb_());
    const ticket = repo.tickets.get(ticketId);
    if (!ticket) throw new Error("Ticket não encontrado.");
    enforceTicketAccess_(ctx, ticket, /*allowReceptionAll*/ true);

    const fromStatus = ticket.status_atual;
    toStatus = String(toStatus || "").trim();

    // Cancel é via endpoint próprio
    if (toStatus === APP.STATUS.CANCELADO) throw new Error("Use apiCancelTicket.");

    // Permissões por role + transições válidas
    assertTransitionAllowed_(ctx, fromStatus, toStatus);

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      // Recarrega (evita corrida)
      const t2 = repo.tickets.get(ticketId);
      if (!t2) throw new Error("Ticket não encontrado.");
      const from2 = t2.status_atual;
      if (from2 !== fromStatus) {
        // Se mudou entre ler e atualizar, segue com regra baseada no atual
        assertTransitionAllowed_(ctx, from2, toStatus);
      }

      const updated = repo.tickets.setStatus(ticketId, toStatus, ctx.me.email, {
        recepcao_responsavel: extra.recepcao_responsavel || "",
      });

      repo.logs.append({
        ticketId,
        acao: "STATUS_CHANGE",
        de_status: from2,
        para_status: toStatus,
        usuario_email: ctx.me.email,
        detalhes: { extra },
      });

      // Auto-encerrar quando confirmado pela assistência
      if (toStatus === APP.STATUS.CONFIRMADO) {
        const autoTo = APP.STATUS.ENCERRADO;
        // Apenas se ainda não está encerrado/cancelado
        const t3 = repo.tickets.get(ticketId);
        if (t3 && t3.status_atual === APP.STATUS.CONFIRMADO) {
          repo.tickets.setStatus(ticketId, autoTo, "system@auto", {});
          repo.logs.append({
            ticketId,
            acao: "AUTO_CLOSE",
            de_status: APP.STATUS.CONFIRMADO,
            para_status: APP.STATUS.ENCERRADO,
            usuario_email: "system@auto",
            detalhes: {},
          });
        }
      }

      return { ok: true, updated };
    } finally {
      lock.releaseLock();
    }
  });
}

function apiCancelTicket(ticketId, motivo) {
  return wrap_(function () {
    const ctx = authContext_();
    const repo = sheetRepo_(getDb_());
    const ticket = repo.tickets.get(ticketId);
    if (!ticket) throw new Error("Ticket não encontrado.");
    enforceTicketAccess_(ctx, ticket, true);

    motivo = String(motivo || "").trim();
    if (!motivo) throw new Error("Motivo do cancelamento é obrigatório.");

    // Quem pode cancelar:
    // - ADM sempre
    // - ASSISTENCIA apenas se for dono (já garantido pelo enforceTicketAccess) e ainda não encerrado/cancelado
    // - RECEPCAO pode cancelar (fila) se ainda não encerrado/cancelado
    if (ctx.me.perfil === "ASSISTENCIA" && ticket.solicitante_email !== ctx.me.email) {
      throw new Error("Sem permissão para cancelar ticket de outra pessoa.");
    }
    if ([APP.STATUS.ENCERRADO, APP.STATUS.CANCELADO].includes(ticket.status_atual)) {
      throw new Error("Ticket já finalizado. Não dá pra cancelar.");
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const t2 = repo.tickets.get(ticketId);
      if (!t2) throw new Error("Ticket não encontrado.");
      if ([APP.STATUS.ENCERRADO, APP.STATUS.CANCELADO].includes(t2.status_atual)) {
        throw new Error("Ticket já finalizado. Não dá pra cancelar.");
      }

      repo.tickets.cancel(ticketId, motivo, ctx.me.email);
      repo.logs.append({
        ticketId,
        acao: "CANCEL",
        de_status: t2.status_atual,
        para_status: APP.STATUS.CANCELADO,
        usuario_email: ctx.me.email,
        detalhes: { motivo },
      });
      return { ok: true };
    } finally {
      lock.releaseLock();
    }
  });
}

function apiListUsers() {
  return wrap_(function () {
    const ctx = authContext_();
    requireRole_(ctx, ["ADM"]);
    const repo = sheetRepo_(getDb_());
    return { users: repo.users.list() };
  });
}

function apiUpsertUser(userObj) {
  return wrap_(function () {
    const ctx = authContext_();
    requireRole_(ctx, ["ADM"]);
    userObj = userObj || {};
    const email = String(userObj.email || "").trim().toLowerCase();
    if (!email) throw new Error("email é obrigatório");
    const perfil = String(userObj.perfil || "").trim().toUpperCase();
    if (!["ASSISTENCIA", "RECEPCAO", "ADM"].includes(perfil)) throw new Error("perfil inválido");
    const setor = String(userObj.setor || "").trim();
    if (!setor) throw new Error("setor é obrigatório");
    const ativo = !!userObj.ativo;
    const nome = String(userObj.nome || "").trim();

    const repo = sheetRepo_(getDb_());
    repo.users.upsert({ email, nome, perfil, setor, ativo }, ctx.me.email);
    return { ok: true };
  });
}

function apiConfigGet() {
  return wrap_(function () {
    const ctx = authContext_();
    requireRole_(ctx, ["ADM"]);
    const repo = sheetRepo_(getDb_());
    return { config: repo.config.listAll() };
  });
}

function apiConfigSet(key, value) {
  return wrap_(function () {
    const ctx = authContext_();
    requireRole_(ctx, ["ADM"]);
    key = String(key || "").trim();
    value = String(value || "").trim();
    if (!key) throw new Error("key obrigatório");
    const repo = sheetRepo_(getDb_());
    repo.config.set(key, value, ctx.me.email);
    return { ok: true };
  });
}

function apiDashboard(filters) {
  return wrap_(function () {
    const ctx = authContext_();
    // ADM vê tudo. Recepção vê tudo da fila. Assistência vê do próprio setor.
    filters = filters || {};
    const repo = sheetRepo_(getDb_());
    const tickets = repo.tickets.list(filters, ctx);
    const logs = repo.logs.listRange(filters, ctx);

    const cfg = repo.config.listAll();
    const kpis = computeDashboard_(tickets, logs, cfg);
    return { kpis };
  });
}

function apiExportCsv(filters) {
  return wrap_(function () {
    const ctx = authContext_();
    filters = filters || {};
    const repo = sheetRepo_(getDb_());
    const tickets = repo.tickets.list(filters, ctx);

    const headers = APP.HEADERS.TICKETS;
    const lines = [];
    lines.push(headers.join(","));
    tickets.forEach((t) => {
      lines.push(headers.map((h) => csvEscape_(t[h])).join(","));
    });

    const csv = lines.join("\n");
    const stamp = Utilities.formatDate(new Date(), APP.TZ, "yyyyMMdd-HHmmss");
    return { filename: `tickets-${stamp}.csv`, csv };
  });
}

/** ========================= REPOSITORY / DB ========================= **/

function sheetRepo_(ss) {
  const configSh = ss.getSheetByName(APP.SHEETS.CONFIG);
  const usersSh = ss.getSheetByName(APP.SHEETS.USERS);
  const ticketsSh = ss.getSheetByName(APP.SHEETS.TICKETS);
  const logsSh = ss.getSheetByName(APP.SHEETS.LOGS);

  const config = {
    get(key) {
      const map = readSheetAsObjects_(configSh);
      const found = map.find((r) => r.key === key);
      return found ? String(found.value || "") : "";
    },
    listAll() {
      const rows = readSheetAsObjects_(configSh);
      const out = {};
      rows.forEach((r) => {
        if (r.key) out[r.key] = String(r.value || "");
      });
      return out;
    },
    set(key, value, by) {
      const rows = readSheetAsObjects_(configSh);
      const idx = rows.findIndex((r) => r.key === key);
      const now = new Date().toISOString();
      if (idx >= 0) {
        writeObjectRowByIndex_(configSh, idx, { key, value, updated_at: now, updated_by: by });
      } else {
        appendObjectRow_(configSh, { key, value, updated_at: now, updated_by: by });
      }
    },
  };

  const users = {
    getByEmail(email) {
      email = String(email || "").trim().toLowerCase();
      if (!email) return null;
      const rows = readSheetAsObjects_(usersSh);
      const u = rows.find((r) => String(r.email || "").trim().toLowerCase() === email);
      if (!u) return null;
      u.ativo = truthy_(u.ativo);
      u.perfil = String(u.perfil || "").trim().toUpperCase();
      u.setor = String(u.setor || "").trim();
      return u;
    },
    list() {
      const rows = readSheetAsObjects_(usersSh);
      return rows.map((u) => ({
        email: String(u.email || "").trim().toLowerCase(),
        nome: String(u.nome || ""),
        perfil: String(u.perfil || "").trim().toUpperCase(),
        setor: String(u.setor || ""),
        ativo: truthy_(u.ativo),
        updated_at: u.updated_at || "",
        updated_by: u.updated_by || "",
      }));
    },
    upsert(u, by) {
      const rows = readSheetAsObjects_(usersSh);
      const email = String(u.email || "").trim().toLowerCase();
      const now = new Date().toISOString();
      const obj = {
        email,
        nome: String(u.nome || ""),
        perfil: String(u.perfil || "").trim().toUpperCase(),
        setor: String(u.setor || ""),
        ativo: u.ativo ? "TRUE" : "FALSE",
        updated_at: now,
        updated_by: by,
      };
      const idx = rows.findIndex((r) => String(r.email || "").trim().toLowerCase() === email);
      if (idx >= 0) writeObjectRowByIndex_(usersSh, idx, obj);
      else appendObjectRow_(usersSh, obj);
    },
  };

  const tickets = {
    get(ticketId) {
      ticketId = String(ticketId || "").trim();
      if (!ticketId) return null;
      const rows = readSheetAsObjects_(ticketsSh);
      const t = rows.find((r) => String(r.ticketId || "") === ticketId);
      if (!t) return null;
      // normalize booleans
      t.crianca = truthy_(t.crianca);
      return t;
    },
    list(filters, ctx) {
      const rows = readSheetAsObjects_(ticketsSh);
      let out = rows.map((t) => {
        t.crianca = truthy_(t.crianca);
        return t;
      });

      // Acesso por perfil:
      out = out.filter((t) => {
        if (ctx.me.perfil === "ADM") return true;
        if (ctx.me.perfil === "RECEPCAO") return true; // fila geral
        // ASSISTENCIA: só do setor e/ou próprio solicitante (setor é o principal)
        return String(t.setor || "") === String(ctx.me.setor || "");
      });

      // filtros
      const f = normalizeFilters_(filters || {});
      if (f.status) out = out.filter((t) => String(t.status_atual || "") === f.status);
      if (f.setor) out = out.filter((t) => String(t.setor || "") === f.setor);
      if (f.email) out = out.filter((t) => String(t.solicitante_email || "") === f.email);
      if (f.prontuario) out = out.filter((t) => (String(t.prontuario || "").toLowerCase().includes(f.prontuario)));
      if (f.leito) out = out.filter((t) => (String(t.leito || "").toLowerCase().includes(f.leito)));
      if (f.crianca !== null) out = out.filter((t) => !!t.crianca === f.crianca);

      // range de datas (por criado_em)
      if (f.from || f.to) {
        out = out.filter((t) => {
          const d = parseIso_(t.criado_em);
          if (!d) return false;
          if (f.from && d < f.from) return false;
          if (f.to && d > f.to) return false;
          return true;
        });
      }

      // ordena desc por criado_em
      out.sort((a, b) => {
        const da = parseIso_(a.criado_em)?.getTime() || 0;
        const db = parseIso_(b.criado_em)?.getTime() || 0;
        return db - da;
      });

      return out;
    },
    create(payload, by) {
      const nowIso = new Date().toISOString();
      const id = `${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;

      const obj = {
        ticketId: id,
        status_atual: APP.STATUS.ABERTO,
        paciente_nome: payload.paciente_nome,
        prontuario: payload.prontuario,
        leito: payload.leito,
        nascimento: payload.nascimento, // ISO
        crianca: payload.crianca ? "TRUE" : "FALSE",
        mae_nome: payload.crianca ? payload.mae_nome : "",
        solicitante_email: payload.solicitante_email,
        solicitante_nome: payload.solicitante_nome,
        setor: payload.setor,
        recepcao_responsavel: "",
        criado_em: nowIso,
        recebido_em: "",
        em_producao_em: "",
        enviado_em: "",
        entregue_em: "",
        confirmado_em: "",
        encerrado_em: "",
        cancelado_em: "",
        cancelado_motivo: "",
        last_updated_em: nowIso,
        last_updated_por: by,
      };

      appendObjectRow_(ticketsSh, obj);
      return obj;
    },
    setStatus(ticketId, toStatus, by, extra) {
      const rows = readSheetAsObjects_(ticketsSh);
      const idx = rows.findIndex((r) => String(r.ticketId || "") === String(ticketId));
      if (idx < 0) throw new Error("Ticket não encontrado.");

      const nowIso = new Date().toISOString();
      const current = rows[idx];
      const from = String(current.status_atual || "");

      const upd = Object.assign({}, current);
      upd.status_atual = toStatus;

      // Timestamps
      const tsCol = APP.STATUS_TS_COL[toStatus];
      if (tsCol && !upd[tsCol]) upd[tsCol] = nowIso;

      // Campos auxiliares
      if (extra && extra.recepcao_responsavel) {
        upd.recepcao_responsavel = String(extra.recepcao_responsavel);
      } else if (by && (toStatus === APP.STATUS.RECEBIDO || toStatus === APP.STATUS.PRODUCAO || toStatus === APP.STATUS.ENVIADO)) {
        // se recepção mexeu e ainda não tem responsável, marca
        if (!upd.recepcao_responsavel && by !== "system@auto") upd.recepcao_responsavel = by;
      }

      upd.last_updated_em = nowIso;
      upd.last_updated_por = by;

      writeObjectRowByIndex_(ticketsSh, idx, upd);
      return { from, to: toStatus };
    },
    cancel(ticketId, motivo, by) {
      const rows = readSheetAsObjects_(ticketsSh);
      const idx = rows.findIndex((r) => String(r.ticketId || "") === String(ticketId));
      if (idx < 0) throw new Error("Ticket não encontrado.");

      const nowIso = new Date().toISOString();
      const current = rows[idx];
      const upd = Object.assign({}, current);
      upd.status_atual = APP.STATUS.CANCELADO;
      upd.cancelado_em = nowIso;
      upd.cancelado_motivo = motivo;
      upd.last_updated_em = nowIso;
      upd.last_updated_por = by;

      writeObjectRowByIndex_(ticketsSh, idx, upd);
      return true;
    },
  };

  const logs = {
    append({ ticketId, acao, de_status, para_status, usuario_email, detalhes }) {
      const nowIso = new Date().toISOString();
      const logId = `${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
      const obj = {
        logId,
        ticketId: String(ticketId || ""),
        acao: String(acao || ""),
        de_status: String(de_status || ""),
        para_status: String(para_status || ""),
        usuario_email: String(usuario_email || ""),
        timestamp: nowIso,
        detalhes_json: JSON.stringify(detalhes || {}),
      };
      appendObjectRow_(logsSh, obj);
      return obj;
    },
    listByTicket(ticketId) {
      ticketId = String(ticketId || "").trim();
      const rows = readSheetAsObjects_(logsSh);
      const out = rows.filter((r) => String(r.ticketId || "") === ticketId);
      out.sort((a, b) => (parseIso_(a.timestamp)?.getTime() || 0) - (parseIso_(b.timestamp)?.getTime() || 0));
      return out;
    },
    listRange(filters, ctx) {
      // Para dashboard: logs no range de tempo e respeitando perfil (assistência por setor)
      const f = normalizeFilters_(filters || {});
      const rows = readSheetAsObjects_(logsSh);

      let out = rows;
      if (ctx.me.perfil === "ASSISTENCIA") {
        // restringe via tickets do setor (mais caro, mas funcional)
        const ticketSet = new Set(tickets.list({}, ctx).map((t) => t.ticketId));
        out = out.filter((l) => ticketSet.has(String(l.ticketId || "")));
      }
      if (f.from || f.to) {
        out = out.filter((l) => {
          const d = parseIso_(l.timestamp);
          if (!d) return false;
          if (f.from && d < f.from) return false;
          if (f.to && d > f.to) return false;
          return true;
        });
      }
      return out;
    },
  };

  return { config, users, tickets, logs };
}

/** ========================= AUTH / PERMS ========================= **/

function authContext_() {
  const email = getUserEmail_();
  if (!email) throw new Error("E-mail não identificado. Necessário Google Workspace + WebApp executando como 'usuário acessando'.");
  const repo = sheetRepo_(getDb_());
  const me = repo.users.getByEmail(email);
  if (!me) throw new Error("Sem acesso: usuário não cadastrado na aba USERS.");
  if (!me.ativo) throw new Error("Sem acesso: usuário inativo.");
  if (!["ASSISTENCIA", "RECEPCAO", "ADM"].includes(me.perfil)) throw new Error("Perfil inválido no USERS.");
  return { email, me };
}

function getUserEmail_() {
  // Em Workspace, getActiveUser funciona quando WebApp está como "User accessing".
  // getEffectiveUser pode retornar o owner dependendo do deploy; por isso priorizamos ActiveUser.
  const a = (Session.getActiveUser && Session.getActiveUser().getEmail()) || "";
  const e = (Session.getEffectiveUser && Session.getEffectiveUser().getEmail()) || "";
  return String(a || e || "").trim().toLowerCase();
}

function requireRole_(ctx, allowed) {
  if (!allowed.includes(ctx.me.perfil)) throw new Error("Sem permissão.");
}

function enforceTicketAccess_(ctx, ticket, allowReceptionAll) {
  if (ctx.me.perfil === "ADM") return true;
  if (ctx.me.perfil === "RECEPCAO" && allowReceptionAll) return true;
  if (ctx.me.perfil === "ASSISTENCIA") {
    if (String(ticket.setor || "") !== String(ctx.me.setor || "")) throw new Error("Sem acesso a ticket de outro setor.");
    return true;
  }
  // recepção consultando sem allowReceptionAll
  if (ctx.me.perfil === "RECEPCAO") return true;
  throw new Error("Sem permissão.");
}

function assertTransitionAllowed_(ctx, fromStatus, toStatus) {
  const role = ctx.me.perfil;
  const from = String(fromStatus || "");
  const to = String(toStatus || "");

  if (!Object.values(APP.STATUS).includes(to)) throw new Error("Status destino inválido.");
  if ([APP.STATUS.ENCERRADO, APP.STATUS.CANCELADO].includes(from)) {
    throw new Error("Ticket já finalizado. Não pode mudar status.");
  }

  // Mapa base de transições válidas
  const allowedNext = {};
  allowedNext[APP.STATUS.ABERTO] = [APP.STATUS.RECEBIDO, APP.STATUS.CANCELADO];
  allowedNext[APP.STATUS.RECEBIDO] = [APP.STATUS.PRODUCAO, APP.STATUS.ENVIADO, APP.STATUS.CANCELADO];
  allowedNext[APP.STATUS.PRODUCAO] = [APP.STATUS.ENVIADO, APP.STATUS.CANCELADO];
  allowedNext[APP.STATUS.ENVIADO] = [APP.STATUS.ENTREGUE, APP.STATUS.CONFIRMADO, APP.STATUS.CANCELADO];
  allowedNext[APP.STATUS.ENTREGUE] = [APP.STATUS.CONFIRMADO, APP.STATUS.CANCELADO];
  allowedNext[APP.STATUS.CONFIRMADO] = [APP.STATUS.ENCERRADO]; // auto
  // Encerrado/cancelado não saem

  const baseOk = (allowedNext[from] || []).includes(to);
  if (!baseOk) {
    // ADM pode forçar qualquer transição (mas loga). Aqui vamos permitir apenas ADM.
    if (role === "ADM") return true;
    throw new Error(`Transição inválida: ${from} -> ${to}`);
  }

  // Regras por perfil
  if (role === "ADM") return true;

  if (role === "ASSISTENCIA") {
    // Assistência só cria (ABERTO) e confirma recebimento (CONFIRMADO).
    if (to !== APP.STATUS.CONFIRMADO) {
      throw new Error("Assistência só pode confirmar recebimento (RECEBIDO PELA ASSISTÊNCIA).");
    }
    // Só pode confirmar se já estiver ENVIADO ou ENTREGUE (garantido por baseOk acima)
    return true;
  }

  if (role === "RECEPCAO") {
    // Recepção não confirma recebimento (isso é da assistência)
    if (to === APP.STATUS.CONFIRMADO) throw new Error("Recepção não pode marcar 'Recebido pela assistência'.");
    // Recepção não encerra diretamente (encerramento é automático após confirmação)
    if (to === APP.STATUS.ENCERRADO) throw new Error("Encerramento é automático após confirmação.");
    return true;
  }

  throw new Error("Perfil não suportado.");
}

/** ========================= VALIDATION / NORMALIZATION ========================= **/

function normalizeTicketPayload_(p) {
  const clean = (s) => normalizeStr_(s);
  const isChild = !!p.crianca;

  // nascimento vem como DD/MM/AAAA no front, convertendo pra ISO
  const nascimentoIso = normalizeBirthToIso_(p.nascimento);

  return {
    paciente_nome: clean(p.paciente_nome),
    prontuario: clean(p.prontuario),
    leito: clean(p.leito),
    nascimento: nascimentoIso,
    crianca: isChild,
    mae_nome: clean(p.mae_nome),
  };
}

function validateTicketPayload_(p) {
  if (!p.paciente_nome) throw new Error("Nome do paciente é obrigatório.");
  if (!p.prontuario) throw new Error("Prontuário é obrigatório.");
  if (!p.leito) throw new Error("Leito é obrigatório.");
  if (!p.nascimento) throw new Error("Data de nascimento inválida.");
  if (p.crianca && !p.mae_nome) throw new Error("Nome da mãe é obrigatório para criança.");
}

function normalizeStr_(s) {
  s = String(s || "").trim();
  // remove espaços duplos
  s = s.replace(/\s+/g, " ");
  return s;
}

function normalizeBirthToIso_(input) {
  // aceita "DD/MM/AAAA" ou ISO
  let s = String(input || "").trim();
  if (!s) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return "";
  const dd = parseInt(m[1], 10);
  const mm = parseInt(m[2], 10);
  const yyyy = parseInt(m[3], 10);
  if (yyyy < 1900 || yyyy > 2100) return "";
  if (mm < 1 || mm > 12) return "";
  const maxDay = new Date(yyyy, mm, 0).getDate();
  if (dd < 1 || dd > maxDay) return "";
  const iso = `${yyyy.toString().padStart(4, "0")}-${mm.toString().padStart(2, "0")}-${dd.toString().padStart(2, "0")}`;
  return iso;
}

function normalizeFilters_(filters) {
  const f = Object.assign({}, filters || {});
  const s = (x) => String(x || "").trim();
  const lower = (x) => s(x).toLowerCase();

  f.status = s(f.status);
  f.setor = s(f.setor);
  f.email = lower(f.email);
  f.prontuario = lower(f.prontuario);
  f.leito = lower(f.leito);

  // criança
  if (f.crianca === true || f.crianca === false) {
    // ok
  } else if (s(f.crianca) === "") {
    f.crianca = null;
  } else {
    const v = s(f.crianca).toLowerCase();
    if (v === "true" || v === "1" || v === "sim") f.crianca = true;
    else if (v === "false" || v === "0" || v === "nao" || v === "não") f.crianca = false;
    else f.crianca = null;
  }

  // datas (espera yyyy-mm-dd no frontend)
  f.from = s(f.from) ? new Date(`${s(f.from)}T00:00:00`) : null;
  f.to = s(f.to) ? new Date(`${s(f.to)}T23:59:59`) : null;

  return f;
}

/** ========================= DASHBOARD ========================= **/

function computeDashboard_(tickets, logs, cfg) {
  const now = new Date();

  const slaTotalMin = parseFloat(cfg.SLA_TOTAL_MIN || "30");
  const slaRecepcaoMin = parseFloat(cfg.SLA_RECEPCAO_MIN || "5");
  const slaEnviarMin = parseFloat(cfg.SLA_ENVIAR_MIN || "15");

  const byStatus = {};
  Object.values(APP.STATUS).forEach((st) => (byStatus[st] = 0));
  tickets.forEach((t) => {
    byStatus[t.status_atual] = (byStatus[t.status_atual] || 0) + 1;
  });

  const openNow = tickets.filter((t) => APP.OPEN_STATUSES().includes(t.status_atual)).length;

  // Durations (min)
  const durations = {
    toRecepcao: [],
    toEnviar: [],
    toEncerrar: [],
  };

  const sla = { totalOk: 0, totalN: 0, recepcaoOk: 0, recepcaoN: 0, enviarOk: 0, enviarN: 0 };

  tickets.forEach((t) => {
    const criado = parseIso_(t.criado_em);
    const recebido = parseIso_(t.recebido_em);
    const enviado = parseIso_(t.enviado_em);
    const encerrado = parseIso_(t.encerrado_em);

    if (criado && recebido) {
      const mins = (recebido - criado) / 60000;
      durations.toRecepcao.push(mins);
      sla.recepcaoN++;
      if (mins <= slaRecepcaoMin) sla.recepcaoOk++;
    }

    if (criado && enviado) {
      const mins = (enviado - criado) / 60000;
      durations.toEnviar.push(mins);
      sla.enviarN++;
      if (mins <= slaEnviarMin) sla.enviarOk++;
    }

    if (criado && encerrado) {
      const mins = (encerrado - criado) / 60000;
      durations.toEncerrar.push(mins);
      sla.totalN++;
      if (mins <= slaTotalMin) sla.totalOk++;
    }
  });

  const avg = (arr) => (arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0);

  // Top setores
  const setorCount = {};
  tickets.forEach((t) => {
    const s = String(t.setor || "SEM_SETOR");
    setorCount[s] = (setorCount[s] || 0) + 1;
  });
  const topSetores = Object.keys(setorCount)
    .map((k) => ({ setor: k, total: setorCount[k] }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 10);

  // Horários de pico (por criado_em hora 0-23)
  const hourCount = Array.from({ length: 24 }, () => 0);
  tickets.forEach((t) => {
    const d = parseIso_(t.criado_em);
    if (d) hourCount[d.getHours()]++;
  });

  // Heatmap por dia da semana (0-6) x hora
  const heat = Array.from({ length: 7 }, () => Array.from({ length: 24 }, () => 0));
  tickets.forEach((t) => {
    const d = parseIso_(t.criado_em);
    if (!d) return;
    heat[d.getDay()][d.getHours()]++;
  });

  // Cancelamentos por motivo
  const cancelByReason = {};
  tickets
    .filter((t) => t.status_atual === APP.STATUS.CANCELADO)
    .forEach((t) => {
      const r = String(t.cancelado_motivo || "SEM_MOTIVO").trim() || "SEM_MOTIVO";
      cancelByReason[r] = (cancelByReason[r] || 0) + 1;
    });
  const cancelReasons = Object.keys(cancelByReason).map((k) => ({ motivo: k, total: cancelByReason[k] })).sort((a, b) => b.total - a.total);

  // Crianças x adultos
  const children = tickets.filter((t) => !!t.crianca).length;
  const adults = tickets.length - children;

  // Aging (min) para abertos
  const aging = tickets
    .filter((t) => APP.OPEN_STATUSES().includes(t.status_atual))
    .map((t) => {
      const criado = parseIso_(t.criado_em);
      const mins = criado ? Math.floor((now - criado) / 60000) : 0;
      return { ticketId: t.ticketId, setor: t.setor, status: t.status_atual, minutos: mins, paciente_nome: t.paciente_nome, leito: t.leito };
    })
    .sort((a, b) => b.minutos - a.minutos)
    .slice(0, 50);

  // Atividade de recepcionistas (logs)
  const recActivity = {};
  logs.forEach((l) => {
    const u = String(l.usuario_email || "");
    const acao = String(l.acao || "");
    if (!u) return;
    if (acao === "STATUS_CHANGE" || acao === "CANCEL") {
      recActivity[u] = (recActivity[u] || 0) + 1;
    }
  });
  const topRecepcionistas = Object.keys(recActivity)
    .map((k) => ({ email: k, acoes: recActivity[k] }))
    .sort((a, b) => b.acoes - a.acoes)
    .slice(0, 10);

  // Funil por status (counts)
  const funnel = Object.keys(byStatus).map((k) => ({ status: k, total: byStatus[k] }));

  return {
    totals: {
      total: tickets.length,
      openNow,
      children,
      adults,
    },
    avgTimes: {
      toRecepcaoMin: round_(avg(durations.toRecepcao), 2),
      toEnviarMin: round_(avg(durations.toEnviar), 2),
      toEncerrarMin: round_(avg(durations.toEncerrar), 2),
    },
    sla: {
      slaTotalMin,
      slaRecepcaoMin,
      slaEnviarMin,
      totalPct: sla.totalN ? round_((sla.totalOk / sla.totalN) * 100, 1) : 0,
      recepcaoPct: sla.recepcaoN ? round_((sla.recepcaoOk / sla.recepcaoN) * 100, 1) : 0,
      enviarPct: sla.enviarN ? round_((sla.enviarOk / sla.enviarN) * 100, 1) : 0,
      n: sla,
    },
    charts: {
      byStatus,
      topSetores,
      hourCount,
      heat, // [7][24]
      cancelReasons,
      topRecepcionistas,
      funnel,
    },
    aging,
  };
}

/** ========================= MOCK ========================= **/

function mockGenerateTickets() {
  const ss = getDb_();
  const repo = sheetRepo_(ss);
  const email = getUserEmail_() || "mock@local";
  // precisa existir como ADM pra criar logs/tickets com setor real
  const me = repo.users.getByEmail(email);
  if (!me) throw new Error("Cadastre seu usuário em USERS (ou rode Seed Admin) antes de gerar mock.");

  const setores = ["VASCULAR", "CTI", "CLINICA", "PEDIATRIA", "EMERGENCIA"];
  const nomes = ["Ana", "Bruno", "Carla", "Diego", "Eva", "Felipe", "Gi", "Hugo", "Iara", "João"];
  const leitos = ["101A", "101B", "203", "305", "600.1", "600.2"];
  const prontBase = 100000;

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    for (let i = 0; i < 25; i++) {
      const setor = setores[Math.floor(Math.random() * setores.length)];
      const crianca = Math.random() < 0.2;
      const nascimento = crianca ? "2018-06-10" : "1989-03-22";

      const payload = {
        paciente_nome: `Paciente ${nomes[Math.floor(Math.random() * nomes.length)]} ${i}`,
        prontuario: String(prontBase + i),
        leito: leitos[Math.floor(Math.random() * leitos.length)],
        nascimento,
        crianca,
        mae_nome: crianca ? `Mãe ${nomes[Math.floor(Math.random() * nomes.length)]}` : "",
        solicitante_email: email,
        solicitante_nome: me.nome || email,
        setor,
      };

      // cria direto
      const created = repo.tickets.create(payload, email);
      repo.logs.append({ ticketId: created.ticketId, acao: "CREATE", de_status: "", para_status: created.status_atual, usuario_email: email, detalhes: { mock: true } });

      // random status progress
      const steps = [APP.STATUS.RECEBIDO, APP.STATUS.PRODUCAO, APP.STATUS.ENVIADO, APP.STATUS.ENTREGUE];
      const advance = Math.floor(Math.random() * (steps.length + 1));
      let from = APP.STATUS.ABERTO;
      for (let s = 0; s < advance; s++) {
        const to = steps[s];
        repo.tickets.setStatus(created.ticketId, to, "recepcao@mock", {});
        repo.logs.append({ ticketId: created.ticketId, acao: "STATUS_CHANGE", de_status: from, para_status: to, usuario_email: "recepcao@mock", detalhes: { mock: true } });
        from = to;
      }
    }
  } finally {
    lock.releaseLock();
  }
  return true;
}

/** ========================= HELPERS ========================= **/

function getDb_() {
  // 1) Se for container-bound, usa active
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch (e) {}

  // 2) Se for standalone, pega ID do Script Properties
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty("DB_SPREADSHEET_ID");
  if (!id) throw new Error("Script standalone: defina DB_SPREADSHEET_ID em Script Properties ou use container-bound.");
  return SpreadsheetApp.openById(id);
}

function ensureSheetWithHeaders_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const range = sh.getRange(1, 1, 1, headers.length);
  const values = range.getValues()[0];
  const isEmpty = values.every((v) => !String(v || "").trim());
  if (isEmpty) {
    range.setValues([headers]);
    sh.autoResizeColumns(1, headers.length);
  } else {
    // garante que bate com o esperado (se divergente, não destrói, mas tenta ajustar faltantes)
    const current = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    const need = headers;
    // se a primeira célula não é a mesma, assume que já existe e não mexe
    if (String(current[0] || "") !== String(need[0] || "")) return;
    // adiciona colunas faltantes no final
    for (let i = current.length; i < need.length; i++) {
      sh.getRange(1, i + 1).setValue(need[i]);
    }
  }
}

function readSheetAsObjects_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map((h) => String(h || "").trim());
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const obj = {};
    let empty = true;
    for (let c = 0; c < headers.length; c++) {
      const key = headers[c];
      if (!key) continue;
      const val = row[c];
      if (val !== "" && val !== null && val !== undefined) empty = false;
      obj[key] = val;
    }
    if (!empty) out.push(obj);
  }
  return out;
}

function appendObjectRow_(sh, obj) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map((h) => String(h || "").trim());
  const row = headers.map((h) => (h in obj ? obj[h] : ""));
  sh.appendRow(row);
}

function writeObjectRowByIndex_(sh, idx0Based, obj) {
  // idx0Based refere-se ao array sem header. Linha real = idx+2
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map((h) => String(h || "").trim());
  const row = headers.map((h) => (h in obj ? obj[h] : ""));
  sh.getRange(idx0Based + 2, 1, 1, headers.length).setValues([row]);
}

function parseIso_(iso) {
  const s = String(iso || "").trim();
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function truthy_(v) {
  const s = String(v || "").toLowerCase().trim();
  return s === "true" || s === "1" || s === "sim" || s === "yes";
}

function round_(n, d) {
  const p = Math.pow(10, d || 0);
  return Math.round((Number(n) || 0) * p) / p;
}

function csvEscape_(v) {
  const s = String(v ?? "");
  if (/[,"\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function wrap_(fn) {
  try {
    const data = fn();
    return { ok: true, data };
  } catch (e) {
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}
