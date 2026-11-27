/**
 * Amex · tarjeta de crédito («Finanzas familiares»).
 *
 * Va en el MISMO proyecto de Apps Script de la hoja del mes (el que responde a leerDatos).
 * La primera vez crea la pestaña «AMEX» en la hoja. No toca EGRESOS, INGRESOS ni CONTEO
 * (salvo instalarCuadreAmex_, que se corre una vez y cambia la fórmula del cuadre).
 *
 * Se registran solas en API_ACCIONES (al final de este archivo): leerAmex, guardarAmex, movAmex.
 *
 * Pestaña AMEX:
 *   C5 Línea total · C6 Lo que debemos hoy (fórmula) · C7 Disponible · C8 Usado
 *   C9 Día de pago · C10 Deuda base (último ajuste) · C11 Base desde (fecha y hora)
 *   C12 Descontado del presupuesto y sin pagar (compras con Amex que ya restaron del saldo
 *       de EGRESOS aunque la plata no salió; el cuadre del conteo lo suma de vuelta)
 *   Desde E6: movimientos  Fecha y hora | Tipo (compra/pago) | Descripción | Monto | Categoría | Usuario
 * Lo que debemos = deuda base + compras − pagos registrados DESPUÉS de «Base desde».
 * Al copiar la hoja para el mes siguiente, la pestaña se copia con todo y sigue funcionando.
 *
 * Cuadre del conteo: EGRESOS!H4 = conteo − saldo − AMEX!C12.
 * Una compra con Amex resta del saldo (MONTO) pero no del efectivo. C12 sube con cada compra
 * que la app registra con pres:true y baja (nunca pasa de lo que debemos) cuando se paga la Amex.
 */

var AMEX_HOJA = 'AMEX';

function amexHoja_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(AMEX_HOJA);
  if (sh && String(sh.getRange('B5').getValue()) === 'Línea total') return sh;
  if (!sh) sh = ss.insertSheet(AMEX_HOJA);
  sh.setHiddenGridlines(true);
  sh.getRange('A1:K200').setBackground('#0C0C0E').setFontColor('#F4F4F6').setFontFamily('Arial');
  sh.setColumnWidth(1, 18); sh.setColumnWidth(2, 250); sh.setColumnWidth(3, 150); sh.setColumnWidth(4, 24);
  sh.setColumnWidth(5, 150); sh.setColumnWidth(6, 80); sh.setColumnWidth(7, 240); sh.setColumnWidth(8, 100);
  sh.setColumnWidth(9, 170); sh.setColumnWidth(10, 90);
  sh.getRange('B2').setValue('💳 Amex · tarjeta de crédito').setFontSize(18).setFontWeight('bold');
  sh.getRange('B3').setValue('No es plata nuestra: lo que se paga con la Amex es una deuda. Esta pestaña la llena la app.')
    .setFontColor('#9A9AA0').setFontSize(10);
  var etiquetas = [['Línea total'], ['Lo que debemos hoy'], ['Disponible en la tarjeta'], ['Usado'], ['Día de pago'], ['Deuda base (último ajuste)'], ['Base desde'], ['Descontado del presupuesto (sin pagar)']];
  sh.getRange('B5:B12').setValues(etiquetas).setFontColor('#B8B8C0');
  sh.getRange('C6').setFormula('=C10+SUMIFS(H6:H,F6:F,"compra",E6:E,">"&C11)-SUMIFS(H6:H,F6:F,"pago",E6:E,">"&C11)');
  sh.getRange('C7').setFormula('=C5-C6');
  sh.getRange('C8').setFormula('=IF(C5>0,C6/C5,0)').setNumberFormat('0%');
  sh.getRange('C5:C7').setNumberFormat('"S/ "#,##0.00');
  sh.getRange('C10').setNumberFormat('"S/ "#,##0.00');
  sh.getRange('C11').setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange('C12').setNumberFormat('"S/ "#,##0.00');
  if (sh.getRange('C12').getValue() === '') sh.getRange('C12').setValue(0);
  sh.getRange('C5:C12').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('C6').setFontColor('#FF8FB1').setFontSize(14);
  sh.getRange('C7').setFontColor('#93E36B');
  sh.getRange('B10:C12').setFontColor('#6E6E76').setFontWeight('normal');
  sh.getRange('E5:J5').setValues([['Fecha y hora', 'Tipo', 'Descripción', 'Monto', 'Categoría', 'Usuario']])
    .setFontWeight('bold').setBackground('#1C1C20').setFontColor('#BEE3F8');
  sh.getRange('E6:E200').setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange('H6:H200').setNumberFormat('"S/ "#,##0.00');
  sh.setFrozenRows(5);
  return sh;
}

function amexMovs_(sh) {
  var ultima = sh.getLastRow();
  if (ultima < 6) return [];
  return sh.getRange(6, 5, ultima - 5, 6).getValues()
    .filter(function (f) { return f[0] !== '' && f[0] !== null; })
    .map(function (f) {
      var ts = f[0] instanceof Date ? f[0] : new Date(f[0]);
      return { ts: ts.toISOString(), t: String(f[1] || ''), d: String(f[2] || ''), m: Number(f[3]) || 0, c: String(f[4] || ''), u: String(f[5] || '') };
    });
}

/** Compras con Amex que ya restaron del saldo de EGRESOS y siguen sin pagarse (0 si no hay pestaña). */
function amexDesc_(sh) {
  sh = sh || SpreadsheetApp.getActive().getSheetByName(AMEX_HOJA);
  if (!sh) return 0;
  return Number(sh.getRange('C12').getValue()) || 0;
}

/** Guarda C12 sin pasar de lo que debemos (C6) ni bajar de 0. */
function amexDescSet_(sh, v) {
  SpreadsheetApp.flush();
  var tope = Number(sh.getRange('C6').getValue()) || 0;
  sh.getRange('C12').setValue(Math.round(Math.max(0, Math.min(Number(v) || 0, tope)) * 100) / 100);
}

function leerAmex() {
  var sh = SpreadsheetApp.getActive().getSheetByName(AMEX_HOJA);
  if (!sh || sh.getRange('C5').getValue() === '') return { vacio: true };
  var v = sh.getRange('C5:C12').getValues().map(function (r) { return r[0]; });
  var base = v[6] instanceof Date ? v[6] : (v[6] ? new Date(v[6]) : null);
  return {
    linea: Number(v[0]) || 0,
    deuda: Number(v[1]) || 0,
    dia: v[4] === '' ? null : Number(v[4]),
    base: Number(v[5]) || 0,
    baseTs: base ? base.toISOString() : null,
    desc: Number(v[7]) || 0,
    movs: amexMovs_(sh).slice(-120)
  };
}

function guardarAmex(d) {
  d = d || {};
  var sh = amexHoja_();
  if (d.linea != null) sh.getRange('C5').setValue(Number(d.linea));
  if (Object.prototype.hasOwnProperty.call(d, 'dia')) sh.getRange('C9').setValue(d.dia == null || d.dia === '' ? '' : Number(d.dia));
  if (d.base != null) {
    sh.getRange('C10').setValue(Number(d.base));
    sh.getRange('C11').setValue(d.baseTs ? new Date(d.baseTs) : new Date());
    amexDescSet_(sh, amexDesc_(sh));
  }
  SpreadsheetApp.flush();
  return leerAmex();
}

/**
 * d = {t:'compra'|'pago', m, d, c, u, ts, pres}
 * pres:true = la compra también se anotó en EGRESOS (restó del saldo): sube C12.
 * Un pago baja C12 hasta lo que quede por deber.
 */
function movAmex(d) {
  d = d || {};
  var tipo = d.t === 'pago' ? 'pago' : 'compra', monto = Number(d.m) || 0;
  if (!(monto > 0)) throw new Error('Monto no válido');
  var sh = amexHoja_(), ts = d.ts ? new Date(d.ts) : new Date();
  var ya = amexMovs_(sh).some(function (m) { return m.ts === ts.toISOString() && m.t === tipo && Math.abs(m.m - monto) < 0.005; });
  if (!ya) {
    var col = sh.getRange('E6:E').getValues(), fila = 6;
    for (var i = 0; i < col.length; i++) { if (col[i][0] === '' || col[i][0] === null) { fila = 6 + i; break; } fila = 7 + i; }
    sh.getRange(fila, 5, 1, 6).setValues([[ts, tipo, String(d.d || ''), monto, String(d.c || ''), String(d.u || '')]]);
    sh.getRange(fila, 6, 1, 4).setFontColor(tipo === 'pago' ? '#93E36B' : '#FF8FB1');
    if (tipo === 'compra' && d.pres) amexDescSet_(sh, amexDesc_(sh) + monto);
    else if (tipo === 'pago') amexDescSet_(sh, amexDesc_(sh));
  }
  SpreadsheetApp.flush();
  return leerAmex();
}

/** Ejecutar una vez desde el editor (▶ crearAmex): crea la pestaña AMEX con los datos iniciales. */
function crearAmex() {
  var sh = amexHoja_();
  if (sh.getRange('C5').getValue() === '') guardarAmex({ linea: 1000, base: 503.05, baseTs: '2026-11-27T16:17:00.000Z', dia: null });
  Logger.log(JSON.stringify(leerAmex()));
}

/**
 * Ejecutar una vez por hoja: deja listo C12 y cambia el cuadre de EGRESOS (H4) para que
 * descuente lo que se pagó con Amex y todavía no salió de la plata.
 */
function instalarCuadreAmex_(desc) {
  var ss = SpreadsheetApp.getActive(), sh = amexHoja_();
  sh.getRange('B12').setValue('Descontado del presupuesto (sin pagar)').setFontColor('#6E6E76').setFontWeight('normal');
  var c = sh.getRange('C12');
  if (c.getValue() === '') c.setValue(Number(desc) || 0);
  c.setNumberFormat('"S/ "#,##0.00').setFontColor('#6E6E76').setFontWeight('bold').setHorizontalAlignment('right');
  var h4 = ss.getSheetByName('EGRESOS').getRange('H4'), antes = h4.getFormula();
  if (antes !== '=ROUND(J4-B4,2)' && antes.indexOf('AMEX!C12') < 0) throw new Error('EGRESOS!H4 no es la fórmula esperada: ' + antes);
  h4.setFormula('=ROUND(J4-B4-N(AMEX!C12),2)');
  SpreadsheetApp.flush();
  Logger.log('H4 antes: ' + antes + ' | ahora: ' + h4.getFormula() + ' = ' + h4.getValue() + ' | C12 = ' + c.getValue());
}
function instalarCuadreNoviembre() { instalarCuadreAmex_(405.7); }
function instalarCuadreDiciembre() { instalarCuadreAmex_(0); }

if (typeof API_ACCIONES !== 'undefined') {
  API_ACCIONES.leerAmex = function (d) { return leerAmex(d); };
  API_ACCIONES.guardarAmex = function (d) { return guardarAmex(d); };
  API_ACCIONES.movAmex = function (d) { return movAmex(d); };
}
