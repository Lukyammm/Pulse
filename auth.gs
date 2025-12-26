/**
 * auth.gs
 * Autenticação, hashing de senha e sessão segura.
 */

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
