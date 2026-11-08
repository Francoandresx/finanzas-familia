/**
 * Usuarios y bitácora de «Finanzas familiares».
 *
 * Es un proyecto de Apps Script APARTE del que ya usas (no toca tu hoja ni tu web app actual).
 * Pasos completos en apps-script/LEEME-usuarios.md.
 *
 * Propiedades del script (⚙ Configuración del proyecto → Propiedades del script):
 *   HOJA_ID  = el ID de tu hoja de Google (lo que va entre /d/ y /edit en su URL)
 *   CLAVE    = la misma clave de acceso que ya usa la app (el valor de ?k=... en tu enlace)
 *
 * Crea dos pestañas nuevas en tu hoja la primera vez que se usa:
 *   USUARIOS  nombre | color | activo        (una fila por persona; edítala a mano)
 *   BITACORA  fecha y hora | usuario | acción | categoría | descripción | monto | fecha del movimiento
 */

var HOJA_USUARIOS = 'USUARIOS';
var HOJA_BITACORA = 'BITACORA';
var COLORES = ['#38D591', '#F2BE22', '#F29727', '#26A699', '#A3365A', '#3B6E8C'];

function props_() { return PropertiesService.getScriptProperties(); }

function hoja_(nombre, cabecera, filasIniciales) {
  var ss = SpreadsheetApp.openById(props_().getProperty('HOJA_ID'));
  var sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.appendRow(cabecera);
    sh.setFrozenRows(1);
    (filasIniciales || []).forEach(function (f) { sh.appendRow(f); });
  }
  return sh;
}

function usuarios_() {
  var sh = hoja_(HOJA_USUARIOS, ['nombre', 'color', 'activo'], [['Franco', COLORES[0], true]]);
  var valores = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < valores.length; i++) {
    var nombre = String(valores[i][0] || '').trim();
    var activo = valores[i][2];
    if (!nombre || activo === false || String(activo).toLowerCase() === 'false' || String(activo).toLowerCase() === 'no') continue;
    out.push({ nombre: nombre, color: String(valores[i][1] || COLORES[out.length % COLORES.length]) });
  }
  return out;
}

function registrar_(d) {
  var nombre = String(d.usuario || '').trim();
  if (!usuarios_().some(function (u) { return u.nombre === nombre; })) throw new Error('Usuario no válido');
  var acciones = { gasto: 1, ingreso: 1, ajuste: 1, cliente: 1, conteo: 1 };
  var accion = acciones[d.accion] ? d.accion : 'otro';
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    hoja_(HOJA_BITACORA, ['fecha y hora', 'usuario', 'acción', 'categoría', 'descripción', 'monto', 'fecha del movimiento'])
      .appendRow([new Date(), nombre, accion, String(d.categoria || ''), String(d.descripcion || '').slice(0, 120), Number(d.monto) || 0, String(d.fecha || '')]);
  } finally { lock.releaseLock(); }
  return { ok: true };
}

function bitacora_(d) {
  var limite = Math.max(1, Math.min(500, Number(d && d.limite) || 200));
  var sh = hoja_(HOJA_BITACORA, ['fecha y hora', 'usuario', 'acción', 'categoría', 'descripción', 'monto', 'fecha del movimiento']);
  var ultima = sh.getLastRow();
  if (ultima < 2) return [];
  var desde = Math.max(2, ultima - limite + 1);
  var filas = sh.getRange(desde, 1, ultima - desde + 1, 7).getValues();
  return filas.reverse().map(function (f) {
    return { ts: f[0] instanceof Date ? f[0].getTime() : 0, usuario: f[1], accion: f[2], categoria: f[3], descripcion: f[4], monto: f[5], fecha: String(f[6] || '') };
  });
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = /^[A-Za-z0-9_$.]+$/.test(p.callback || '') ? p.callback : '';
  var salida;
  try {
    if (String(p.key || '') !== String(props_().getProperty('CLAVE') || '')) throw new Error('Clave no válida');
    var datos = p.data ? JSON.parse(p.data) : {};
    var r;
    if (p.api === 'usuarios') r = usuarios_();
    else if (p.api === 'registrar') r = registrar_(datos);
    else if (p.api === 'bitacora') r = bitacora_(datos);
    else throw new Error('API desconocida');
    salida = { ok: true, data: r };
  } catch (err) {
    salida = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  var txt = JSON.stringify(salida);
  return ContentService.createTextOutput(cb ? cb + '(' + txt + ')' : txt)
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}
