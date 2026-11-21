/**
 * TABLERO COMERCIAL MOSTO · API (Google Apps Script)
 * ------------------------------------------------------------
 * Solapas del Sheet (las crea setup()):
 *   credenciales : usuario | contraseña | rol (admin/owner/vend) | id vendedor | nombre | activo (SI/NO)
 *   control      : mes (AAAA-MM) | clave variable | variable | id vendedor | vendedor | target | estado (OK/NO)
 *   innovaciones : marca | cuenta como innovación (SI/NO)
 *   config       : clave | valor
 *   datos        : clave | parte | contenido   (lo publica el tablero; no editar a mano)
 *
 * Instalación: pegar este código, ejecutar setup() una vez, y publicar como
 * Aplicación web (Ejecutar como: yo · Acceso: cualquier persona).
 */

const HOJA_CRED = 'credenciales';
const HOJA_CTRL = 'control';
const HOJA_INNOV = 'innovaciones';
const HOJA_CONFIG = 'config';
const HOJA_DATOS = 'datos';
const PARTE = 45000;               // caracteres por celda (límite de Sheets: 50.000)
const DURACION_SESION = 21600;     // 6 horas

// ── instalación ─────────────────────────────────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const crear = (nombre, encabezado) => {
    let h = ss.getSheetByName(nombre);
    if (!h) h = ss.insertSheet(nombre);
    if (h.getLastRow() === 0) { h.getRange(1, 1, 1, encabezado.length).setValues([encabezado]); h.setFrozenRows(1); }
    return h;
  };
  const cred = crear(HOJA_CRED, ['usuario', 'contraseña', 'rol', 'id vendedor', 'nombre', 'activo']);
  const ctrl = crear(HOJA_CTRL, ['mes', 'clave variable', 'variable', 'id vendedor', 'vendedor', 'target', 'estado']);
  ctrl.getRange(1, 1, 2000, 1).setNumberFormat('@');
  crear(HOJA_INNOV, ['marca', 'cuenta como innovación']);
  crear(HOJA_CONFIG, ['clave', 'valor']);
  crear(HOJA_DATOS, ['clave', 'parte', 'contenido']);
  const hay = cred.getLastRow() > 1 && cred.getRange(2, 1, cred.getLastRow() - 1, 3).getValues().some(r => String(r[2]).trim() === 'admin');
  if (!hay) cred.appendRow(['mati', 'cambiar1234', 'admin', '', 'Matías (administrador)', 'SI']);
  if (getConfig_('verRanking') === null) setConfig_('verRanking', 'SI');
  return 'Listo. Usuario admin: mati / cambiar1234 (cambiá la contraseña en la solapa credenciales).';
}

// ── entrada ─────────────────────────────────────────────────
function doGet() {
  return salida_({ok: true, api: 'Tablero Mosto', hora: new Date().toISOString()});
}
function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return salida_({ok: false, error: 'Pedido inválido'}); }
  try {
    switch (req.action) {
      case 'login': return salida_(login_(req));
      case 'datos': return salida_(datos_(sesion_(req.token)));
      case 'publicar': return salida_(publicar_(sesion_(req.token, ['admin']), req));
      case 'control_get': return salida_({ok: true, ctrl: leerControl_(sesion_(req.token))});
      case 'control_set': return salida_(guardarControl_(sesion_(req.token, ['admin', 'owner']), req));
      case 'usuarios_get': sesion_(req.token, ['admin']); return salida_({ok: true, users: leerUsuarios_(), verRanking: getConfig_('verRanking') !== 'NO'});
      case 'usuarios_set': return salida_(guardarUsuarios_(sesion_(req.token, ['admin']), req));
      default: return salida_({ok: false, error: 'Acción desconocida'});
    }
  } catch (err) {
    return salida_({ok: false, error: String(err && err.message ? err.message : err)});
  }
}
function salida_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ── sesiones ────────────────────────────────────────────────
function login_(req) {
  const u = String(req.u || '').trim().toLowerCase(), p = String(req.p || '');
  const user = leerUsuarios_().find(x => x.user === u && x.activo !== false);
  if (!user || user.pass !== p) { Utilities.sleep(600); return {ok: false, error: 'Usuario o contraseña incorrectos.'}; }
  const token = Utilities.getUuid();
  const s = {user: user.user, rol: user.rol, vid: user.vid, nombre: user.nombre};
  CacheService.getScriptCache().put('s_' + token, JSON.stringify(s), DURACION_SESION);
  const d = datos_(s);
  return {ok: true, token: token, rol: s.rol, vid: s.vid, nombre: s.nombre, data: d.data, meta: d.meta, ctrl: d.ctrl, verRanking: d.verRanking};
}
function sesion_(token, roles) {
  const raw = token ? CacheService.getScriptCache().get('s_' + token) : null;
  if (!raw) throw new Error('SESION_VENCIDA');
  const s = JSON.parse(raw);
  if (roles && roles.indexOf(s.rol) < 0) throw new Error('No tenés permiso para esta acción.');
  return s;
}

// ── datos publicados ────────────────────────────────────────
function datos_(s) {
  const clave = s.rol === 'vend' ? 'v:' + s.vid : 'owner';
  const txt = leerBloque_(clave);
  return {ok: true, data: txt ? JSON.parse(txt) : null, meta: leerMeta_(), ctrl: leerControl_(s), verRanking: getConfig_('verRanking') !== 'NO'};
}
function publicar_(s, req) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_DATOS);
    const filas = [];
    Object.keys(req.payloads || {}).forEach(clave => {
      const txt = JSON.stringify(req.payloads[clave]);
      for (let i = 0, n = 0; i < txt.length; i += PARTE, n++) filas.push([clave, n, txt.slice(i, i + PARTE)]);
    });
    if (!filas.length) throw new Error('No llegaron datos para publicar.');
    if (h.getLastRow() > 1) h.getRange(2, 1, h.getLastRow() - 1, 3).clearContent();
    h.getRange(2, 1, filas.length, 3).setValues(filas);
    setConfig_('publicado', new Date().toISOString());
    setConfig_('publicado_por', s.user);
    setConfig_('archivos', (req.meta && req.meta.archivos || []).join(' + '));
    return {ok: true, filas: filas.length, claves: Object.keys(req.payloads).length};
  } finally { lock.releaseLock(); }
}
function leerBloque_(clave) {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_DATOS);
  if (!h || h.getLastRow() < 2) return null;
  const v = h.getRange(2, 1, h.getLastRow() - 1, 3).getValues().filter(r => String(r[0]) === clave);
  if (!v.length) return null;
  v.sort((a, b) => Number(a[1]) - Number(b[1]));
  return v.map(r => String(r[2])).join('');
}
function leerMeta_() {
  return {publicado: getConfig_('publicado'), archivos: getConfig_('archivos')};
}

// ── control (targets, OK/NO, innovaciones) ─────────────────
function mesTxt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  return String(v).trim();
}
function leerControl_(s) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.getSheetByName(HOJA_CTRL);
  const out = {meses: {}, innovOff: [], upd: getConfig_('control_upd') || ''};
  if (h && h.getLastRow() > 1) {
    h.getRange(2, 1, h.getLastRow() - 1, 7).getValues().forEach(r => {
      const mes = mesTxt_(r[0]), k = String(r[1]).trim(), vid = String(r[3]).trim();
      if (!mes || !k || !vid) return;
      if (s.rol === 'vend' && vid !== String(s.vid) && vid !== '*') return;   // '*' = ajuste para todos (ej. variable quitada)
      const m = out.meses[mes] || (out.meses[mes] = {t: {}, off: {}});
      const est = String(r[6]).trim().toUpperCase();
      if (r[5] !== '' && r[5] !== null && !isNaN(Number(r[5]))) { (m.t[k] = m.t[k] || {})[vid] = Number(r[5]); }
      if (est === 'NO') { (m.off[k] = m.off[k] || {})[vid] = 1; }
    });
  }
  const hi = ss.getSheetByName(HOJA_INNOV);
  if (hi && hi.getLastRow() > 1) {
    hi.getRange(2, 1, hi.getLastRow() - 1, 2).getValues().forEach(r => { if (String(r[0]).trim() && String(r[1]).trim().toUpperCase() === 'NO') out.innovOff.push(String(r[0]).trim()); });
  }
  return out;
}
function guardarControl_(s, req) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(HOJA_CTRL);
    const mes = String(req.mes || '').trim();
    if (!/^\d{4}-\d{2}$/.test(mes)) throw new Error('Mes inválido');
    const labels = req.labels || {}, vends = req.vends || {};
    const data = req.data || {t: {}, off: {}};
    // filas de otros meses se conservan
    let resto = [];
    if (h.getLastRow() > 1) resto = h.getRange(2, 1, h.getLastRow() - 1, 7).getValues().filter(r => mesTxt_(r[0]) && mesTxt_(r[0]) !== mes).map(r => { r[0] = mesTxt_(r[0]); return r; });
    const claves = {};
    Object.keys(data.t || {}).forEach(k => Object.keys(data.t[k]).forEach(v => { claves[k + '|' + v] = 1; }));
    Object.keys(data.off || {}).forEach(k => Object.keys(data.off[k]).forEach(v => { claves[k + '|' + v] = 1; }));
    const nuevas = Object.keys(claves).map(kv => {
      const k = kv.split('|')[0], v = kv.split('|')[1];
      const t = data.t && data.t[k] && data.t[k][v] !== undefined ? data.t[k][v] : '';
      const off = data.off && data.off[k] && data.off[k][v];
      return [mes, k, labels[k] || k, v, vends[v] || '', t, off ? 'NO' : 'OK'];
    });
    const todas = resto.concat(nuevas);
    if (h.getLastRow() > 1) h.getRange(2, 1, h.getLastRow() - 1, 7).clearContent();
    if (todas.length) { h.getRange(2, 1, todas.length, 1).setNumberFormat('@'); h.getRange(2, 1, todas.length, 7).setValues(todas); }
    if (req.innovMarcas) {
      const hi = ss.getSheetByName(HOJA_INNOV);
      const off = req.innovOff || [];
      if (hi.getLastRow() > 1) hi.getRange(2, 1, hi.getLastRow() - 1, 2).clearContent();
      const filas = req.innovMarcas.map(m => [m, off.indexOf(m) >= 0 ? 'NO' : 'SI']);
      if (filas.length) hi.getRange(2, 1, filas.length, 2).setValues(filas);
    }
    const upd = new Date().toISOString();
    setConfig_('control_upd', upd);
    return {ok: true, upd: upd, filas: nuevas.length};
  } finally { lock.releaseLock(); }
}

// ── usuarios ────────────────────────────────────────────────
function leerUsuarios_() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CRED);
  if (!h || h.getLastRow() < 2) return [];
  return h.getRange(2, 1, h.getLastRow() - 1, 6).getValues().filter(r => String(r[0]).trim()).map(r => ({
    user: String(r[0]).trim().toLowerCase(), pass: String(r[1]).trim(), rol: normRol_(r[2]),
    vid: String(r[3]).trim() === '' ? null : (isNaN(Number(r[3])) ? String(r[3]).trim() : Number(r[3])),
    nombre: String(r[4]).trim(), activo: String(r[5]).trim().toUpperCase() !== 'NO'
  }));
}
function normRol_(r) {
  const s = String(r || '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  if (/^(admin|administrador|administradora)$/.test(s)) return 'admin';
  if (/^(owner|dueno|duena|propietario)$/.test(s)) return 'owner';
  return 'vend';
}
function guardarUsuarios_(s, req) {
  const users = req.users || [];
  if (!users.some(u => u.rol === 'admin' && u.activo !== false)) throw new Error('Tiene que quedar al menos un administrador activo.');
  const vistos = {};
  users.forEach(u => {
    const k = String(u.user || '').trim().toLowerCase();
    if (!k) throw new Error('Hay un usuario vacío.');
    if (vistos[k]) throw new Error('Usuario repetido: ' + k);
    if (String(u.pass || '').length < 6) throw new Error('La contraseña de ' + k + ' tiene menos de 6 caracteres.');
    vistos[k] = 1;
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CRED);
    if (h.getLastRow() > 1) h.getRange(2, 1, h.getLastRow() - 1, 6).clearContent();
    const filas = users.map(u => [String(u.user).trim().toLowerCase(), String(u.pass), u.rol, u.vid === null || u.vid === undefined ? '' : u.vid, u.nombre || '', u.activo === false ? 'NO' : 'SI']);
    h.getRange(2, 1, filas.length, 6).setValues(filas);
    if (req.verRanking !== undefined) setConfig_('verRanking', req.verRanking ? 'SI' : 'NO');
    return {ok: true, n: filas.length};
  } finally { lock.releaseLock(); }
}

// ── config ──────────────────────────────────────────────────
function getConfig_(k) {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONFIG);
  if (!h || h.getLastRow() < 2) return null;
  const r = h.getRange(2, 1, h.getLastRow() - 1, 2).getValues().find(x => String(x[0]) === k);
  return r ? String(r[1]) : null;
}
function setConfig_(k, v) {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONFIG);
  const n = h.getLastRow();
  if (n > 1) {
    const vals = h.getRange(2, 1, n - 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) if (String(vals[i][0]) === k) { h.getRange(i + 2, 2, 1, 1).setValues([[v]]); return; }
  }
  h.appendRow([k, v]);
}
