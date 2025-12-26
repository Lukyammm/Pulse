/**
 * permissions.gs
 * Controle de permissões por perfil.
 */

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
