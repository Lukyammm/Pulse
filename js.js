const PERMISSIONS = {
  ASSISTENCIA: { updateStatus: ['CONFIRMADO'] },
  RECEPCAO: { updateStatus: ['RECEBIDO', 'EM_PRODUCAO', 'ENVIADO'] },
  ADM: { updateStatus: ['RECEBIDO', 'EM_PRODUCAO', 'ENVIADO', 'ENTREGUE', 'CONFIRMADO', 'ENCERRADO', 'CANCELADO'] },
};

const qs = (sel) => document.querySelector(sel);
const qsa = (sel) => Array.from(document.querySelectorAll(sel));
let sessionToken = null;
let currentSession = null;

function showLogin() {
  qs('#loginModal').classList.remove('hidden');
}

function hideLogin() {
  qs('#loginModal').classList.add('hidden');
}

function setShellVisible(flag) {
  qs('#waiting').classList.toggle('hidden', flag);
  qs('#shell').classList.toggle('hidden', !flag);
  document.body.classList.toggle('desktop', window.innerWidth > 1024 && flag);
}

function renderHome(perfil) {
  const content = qs('#content');
  content.innerHTML = '';
  const card = document.createElement('div');
  card.className = 'card';
  card.innerHTML = `
    <h3>Bem-vindo</h3>
    <p class="muted">Perfil: <strong>${perfil}</strong></p>
    <p>Use o menu inferior para navegar.</p>
  `;
  content.appendChild(card);
  if (perfil === 'ASSISTENCIA' || perfil === 'ADM') {
    renderFormularioSolicitacao();
  }
}

function renderFormularioSolicitacao() {
  const card = document.createElement('div');
  card.className = 'card span-6';
  card.innerHTML = `
    <h3>Nova solicitação</h3>
    <div class="grid-two">
      <div>
        <label>Nome do paciente</label>
        <input id="fNome" type="text" placeholder="Nome completo" />
      </div>
      <div>
        <label>Prontuário</label>
        <input id="fProntuario" type="text" />
      </div>
      <div>
        <label>Leito</label>
        <input id="fLeito" type="text" />
      </div>
      <div>
        <label>Data de nascimento</label>
        <input id="fNascimento" type="date" />
      </div>
      <div>
        <label>Nome da mãe (se criança)</label>
        <input id="fMae" type="text" />
      </div>
    </div>
    <button class="btn primary full" id="btnEnviarSolicitacao">Criar</button>
    <div class="feedback" id="solicitacaoFeedback"></div>
  `;
  qs('#content').appendChild(card);
  card.querySelector('#btnEnviarSolicitacao').addEventListener('click', () => {
    const payload = {
      nome_paciente: qs('#fNome').value.trim(),
      prontuario: qs('#fProntuario').value.trim(),
      leito: qs('#fLeito').value.trim(),
      data_nascimento: qs('#fNascimento').value,
      nome_mae: qs('#fMae').value.trim(),
    };
    qs('#solicitacaoFeedback').textContent = 'Enviando...';
    google.script.run
      .withSuccessHandler(() => {
        qs('#solicitacaoFeedback').textContent = 'Criado com sucesso';
        loadSolicitacoes();
      })
      .withFailureHandler((err) => {
        qs('#solicitacaoFeedback').textContent = err.message || 'Erro';
      })
      .api_createSolicitacao(payload, sessionToken);
  });
}

function renderSolicitacoes(lista) {
  const content = qs('#content');
  content.innerHTML = '';
  const list = document.createElement('div');
  list.className = 'list';
  lista.forEach((item) => {
    const card = document.createElement('div');
    card.className = 'card';
    const hist = item.historico
      .map((h) => `<span class="badge">${h.status} • ${new Date(h.data).toLocaleString('pt-BR')}</span>`)
      .join(' ');
    card.innerHTML = `
      <div class="status-row">
        <div class="tag">${item.status}</div>
        <div class="badge">${item.prontuario || 'Sem prontuário'}</div>
      </div>
      <h3>${item.nome_paciente}</h3>
      <p class="muted">Leito ${item.leito || '-'}</p>
      <p class="muted">Criado por ${item.solicitante_email}</p>
      <div class="chips">${hist}</div>
    `;
    if (canUpdateStatus(item)) {
      const select = document.createElement('select');
      select.innerHTML = statusOptions(item.status)
        .map((s) => `<option value="${s}">${s}</option>`)
        .join('');
      const btn = document.createElement('button');
      btn.textContent = 'Atualizar';
      btn.className = 'btn ghost';
      btn.addEventListener('click', () => {
        btn.disabled = true;
        google.script.run
          .withSuccessHandler(() => {
            btn.disabled = false;
            loadSolicitacoes();
          })
          .withFailureHandler((err) => {
            btn.disabled = false;
            alert(err.message || 'Erro');
          })
          .api_atualizarStatus({ id: item.id, status: select.value }, sessionToken);
      });
      const row = document.createElement('div');
      row.style.marginTop = '10px';
      row.appendChild(select);
      row.appendChild(btn);
      card.appendChild(row);
    }
    list.appendChild(card);
  });
  content.appendChild(list);
}

function canUpdateStatus(item) {
  if (!currentSession) return false;
  const rules = PERMISSIONS[currentSession.perfil];
  if (!rules || !rules.updateStatus) return false;
  return true;
}

function statusOptions(current) {
  const rules = PERMISSIONS[currentSession.perfil];
  if (!rules) return [];
  return rules.updateStatus.filter((s) => s !== current);
}

function renderDashboard(data) {
  const content = qs('#content');
  content.innerHTML = '';
  const card = document.createElement('div');
  card.className = 'card';
  card.innerHTML = `
    <h3>Visão geral</h3>
    <p>Total: <strong>${data.total}</strong></p>
    <div class="grid-two">
      ${Object.entries(data.porStatus)
        .map(([k, v]) => `<div class="badge">${k}: ${v}</div>`)
        .join('')}
    </div>
    <div class="grid-two" style="margin-top:12px;">
      ${Object.entries(data.porPerfil)
        .map(([k, v]) => `<div class="badge">${k}: ${v}</div>`)
        .join('')}
    </div>
  `;
  content.appendChild(card);
}

function loadSolicitacoes() {
  google.script.run
    .withSuccessHandler((res) => renderSolicitacoes(res))
    .withFailureHandler((err) => alert(err.message || 'Erro'))
    .api_listSolicitacoes(sessionToken);
}

function loadDashboard() {
  google.script.run
    .withSuccessHandler((res) => renderDashboard(res))
    .withFailureHandler((err) => alert(err.message || 'Erro'))
    .api_dashboard(sessionToken);
}

function initNav(perfil) {
  qsa('.nav-btn').forEach((btn) => {
    if (btn.dataset.tab === 'dashboard' && perfil !== 'ADM') {
      btn.classList.add('hidden');
    }
    btn.addEventListener('click', () => {
      qsa('.nav-btn').forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      if (btn.dataset.tab === 'home') renderHome(perfil);
      if (btn.dataset.tab === 'solicitacoes') loadSolicitacoes();
      if (btn.dataset.tab === 'dashboard') loadDashboard();
    });
  });
}

function bindAuth() {
  qs('#btnOpenLogin').addEventListener('click', showLogin);
  qs('#pulseDot').addEventListener('click', showLogin);
  qs('#closeLogin').addEventListener('click', hideLogin);
  qs('#btnLogin').addEventListener('click', () => {
    const email = qs('#loginEmail').value.trim();
    const senha = qs('#loginSenha').value;
    qs('#loginFeedback').textContent = 'Validando...';
    google.script.run
      .withSuccessHandler((res) => {
        sessionToken = res.token;
        currentSession = res;
        hideLogin();
        setShellVisible(true);
        qs('#profileLabel').textContent = `${res.email} • ${res.perfil}`;
        initNav(res.perfil);
        renderHome(res.perfil);
      })
      .withFailureHandler((err) => {
        qs('#loginFeedback').textContent = err.message || 'Erro';
      })
      .api_login(email, senha);
  });
  qs('#btnLogout').addEventListener('click', () => {
    google.script.run.withSuccessHandler(() => {
      sessionToken = null;
      currentSession = null;
      setShellVisible(false);
    }).api_logout(sessionToken);
  });
}

document.addEventListener('DOMContentLoaded', () => {
  bindAuth();
});
