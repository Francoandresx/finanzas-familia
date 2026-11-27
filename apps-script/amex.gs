/**
 * Amex · tarjeta de crédito («Finanzas familiares»).
 *
 * Va en el MISMO proyecto de Apps Script de la hoja del mes (el que responde a leerDatos).
 * La primera vez crea la pestaña «AMEX» en la hoja. No toca EGRESOS, INGRESOS ni CONTEO.
 *
 * Se registran solas en API_ACCIONES (al final de este archivo): leerAmex, guardarAmex, movAmex.
 *
 * Pestaña AMEX:
 *   C5 Línea total · C6 Lo que debemos hoy (fórmula) · C7 Disponible · C8 Usado
 *   C9 Día de pago · C10 Deuda base (último ajuste) · C11 Base desde (fecha y hora)
 *   Desde E6: movimientos  Fecha y hora | Tipo (compra/pago) | Descripción | Monto | Categoría | Usuario
 * Lo que debemos = deuda base + compras − pagos registrados DESPUÉS de «Base desde».
 * Al copiar la hoja para el mes siguiente, la pestaña se copia con todo y sigue funcionando.
 */

var AMEX_HOJA = 'AMEX';

function amexHoja_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(AMEX_HOJA);
  if (sh) return sh;
  sh = ss.insertSheet(AMEX_HOJA);
  sh.setHiddenGridlines(true);
  sh.getRange('A1:K200').setBackground('#0C0C0E').setFontColor('#F4F4F6').setFontFamily('Arial');
  sh.setColumnWidth(1, 18); sh.setColumnWidth(2, 220); sh.setColumnWidth(3, 150); sh.setColumnWidth(4, 24);
  sh.setColumnWidth(5, 150); sh.setColumnWidth(6, 80); sh.setColumnWidth(7, 240); sh.setColumnWidth(8, 100);
  sh.setColumnWidth(9, 170); sh.setColumnWidth(10, 90);
  sh.getRange('B2').setValue('💳 Amex · tarjeta de crédito').setFontSize(18).setFontWeight('bold');
  sh.getRange('B3').setValue('No es plata nuestra: lo que se paga con la Amex es una deuda. Esta pestaña la llena la app.')
    .setFontColor('#9A9AA0').setFontSize(10);
  var etiquetas = [['Línea total'], ['Lo que debemos hoy'], ['Disponible en la tarjeta'], ['Usado'], ['Día de pago'], ['Deuda base (último ajuste)'], ['Base desde']];
  sh.getRange('B5:B11').setValues(etiquetas).setFontColor('#B8B8C0');
  sh.getRange('C6').setFormula('=C10+SUMIFS(H6:H,F6:F,"compra",E6:E,">"&C11)-SUMIFS(H6:H,F6:F,"pago",E6:E,">"&C11)');
  sh.getRange('C7').setFormula('=C5-C6');
  sh.getRange('C8').setFormula('=IF(C5>0,C6/C5,0)').setNumberFormat('0%');
  sh.getRange('C5:C7').setNumberFormat('"S/ "#,##0.00');
  sh.getRange('C10').setNumberFormat('"S/ "#,##0.00');
  sh.getRange('C11').setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange('C5:C11').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('C6').setFontColor('#FF8FB1').setFontSize(14);
  sh.getRange('C7').setFontColor('#93E36B');
  sh.getRange('B10:C11').setFontColor('#6E6E76').setFontWeight('normal');
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

function leerAmex() {
  var sh = SpreadsheetApp.getActive().getSheetByName(AMEX_HOJA);
  if (!sh || sh.getRange('C5').getValue() === '') return { vacio: true };
  var v = sh.getRange('C5:C11').getValues().map(function (r) { return r[0]; });
  var base = v[6] instanceof Date ? v[6] : (v[6] ? new Date(v[6]) : null);
  return {
    linea: Number(v[0]) || 0,
    deuda: Number(v[1]) || 0,
    dia: v[4] === '' ? null : Number(v[4]),
    base: Number(v[5]) || 0,
    baseTs: base ? base.toISOString() : null,
    movs: amexMovs_(sh).slice(-120)
  };
}

function guardarAmex(d) {
  d = d || {};
  {
    var sh = amexHoja_();
    if (d.linea != null) sh.getRange('C5').setValue(Number(d.linea));
    if (Object.prototype.hasOwnProperty.call(d, 'dia')) sh.getRange('C9').setValue(d.dia == null || d.dia === '' ? '' : Number(d.dia));
    if (d.base != null) {
      sh.getRange('C10').setValue(Number(d.base));
      sh.getRange('C11').setValue(d.baseTs ? new Date(d.baseTs) : new Date());
    }
    SpreadsheetApp.flush();
  }
  return leerAmex();
}

function movAmex(d) {
  d = d || {};
  var tipo = d.t === 'pago' ? 'pago' : 'compra', monto = Number(d.m) || 0;
  if (!(monto > 0)) throw new Error('Monto no válido');
  {
    var sh = amexHoja_(), ts = d.ts ? new Date(d.ts) : new Date();
    var ya = amexMovs_(sh).some(function (m) { return m.ts === ts.toISOString() && m.t === tipo && Math.abs(m.m - monto) < 0.005; });
    if (!ya) {
      var col = sh.getRange('E6:E').getValues(), fila = 6;
      for (var i = 0; i < col.length; i++) { if (col[i][0] === '' || col[i][0] === null) { fila = 6 + i; break; } fila = 7 + i; }
      sh.getRange(fila, 5, 1, 6).setValues([[ts, tipo, String(d.d || ''), monto, String(d.c || ''), String(d.u || '')]]);
      sh.getRange(fila, 6, 1, 4).setFontColor(tipo === 'pago' ? '#93E36B' : '#FF8FB1');
    }
    SpreadsheetApp.flush();
  }
  return leerAmex();
}

/** Ejecutar una vez desde el editor (▶ crearAmex): crea la pestaña AMEX con los datos iniciales. */
function crearAmex() {
  var sh = amexHoja_();
  if (sh.getRange('C5').getValue() === '') guardarAmex({ linea: 1000, base: 503.05, baseTs: '2026-11-27T16:17:00.000Z', dia: null });
  Logger.log(JSON.stringify(leerAmex()));
}

if (typeof API_ACCIONES !== 'undefined') {
  API_ACCIONES.leerAmex = function (d) { return leerAmex(d); };
  API_ACCIONES.guardarAmex = function (d) { return guardarAmex(d); };
  API_ACCIONES.movAmex = function (d) { return movAmex(d); };
}
