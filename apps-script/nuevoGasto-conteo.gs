/**
 * Parche para el Apps Script de la web app (/exec) ligada a la hoja «Noviembre».
 * No está publicado. Hay que pegar estas funciones en ese proyecto (reemplazando
 * nuevoGasto y usando saldoEgresosDesdeMontos_ dentro de leerConteo) y volver
 * a implementar la implementación web, misma URL /exec.
 *
 * Descuadre del conteo del 05/11/2026 (leído de la hoja, sin modificarla):
 *   Conteo EGRESOS!J4          S/ 1,172.43
 *   Saldo EGRESOS!B4           S/ 1,351.48
 *   Cuadre EGRESOS!H4 = J4−B4  −S/ 179.05
 *   EGRESOS fila 63, Artículos personales, fecha 05/11:
 *     «⌚️Xiaomi Watch de Anny»
 *     D63 PRESUP. = 179
 *     E63 MONTO   = vacío
 *     E70 = SUM(E63:E69) = 0
 *   B4 = INGRESOS!H14 − (E33+E57+E70+K33+K57+Q14+Q27)
 *   Solo resta columnas MONTO. Los S/ 179 del reloj quedaron como presupuesto,
 *   así que el saldo esperado no bajó. 1,351.48 − 179 = 1,172.48
 *   (quedan S/ 0.05 de redondeo en otros montos).
 *
 * Para cerrar ese descuadre, si el reloj ya se pagó: escribir 179 en E63.
 * No borra D63. Esto no lo hace este archivo; la hoja no se tocó.
 */

function bloqueEgresos_(nombre) {
  var sh = SpreadsheetApp.getActive().getSheetByName('EGRESOS');
  var values = sh.getDataRange().getValues();
  var titleRow = -1;
  for (var r = 0; r < values.length; r++) {
    var blob = values[r].join(' ');
    if (blob.indexOf(nombre) !== -1 && blob.indexOf('FECHA') === -1 && blob.indexOf('Total') === -1) {
      titleRow = r;
      break;
    }
  }
  if (titleRow < 0) throw new Error('No está «' + nombre + '» en EGRESOS.');
  var head = -1;
  for (var r2 = titleRow; r2 < Math.min(values.length, titleRow + 6); r2++) {
    if (String(values[r2][1]).toUpperCase() === 'FECHA') { head = r2; break; }
  }
  if (head < 0) throw new Error('Sin fila FECHA en «' + nombre + '».');
  var presup = -1, monto = -1;
  for (var c = 0; c < values[head].length; c++) {
    var h = String(values[head][c]).toUpperCase();
    if (h.indexOf('PRESUP') === 0) presup = c + 1;
    if (h === 'MONTO') monto = c + 1;
  }
  var first = head + 2; // 1-based, fila debajo del encabezado
  var last = first;
  for (var r3 = head + 1; r3 < values.length; r3++) {
    var etiqueta = String(values[r3][1] || values[r3][2]).toLowerCase();
    if (etiqueta === 'total') break;
    last = r3 + 1;
  }
  return { sh: sh, presup: presup, monto: monto, first: first, last: last };
}

/**
 * p = {categoria, nombre, monto, fecha, pagado}.
 * La web manda pagado:true en categorías sin presupuesto por fila (Ocio y
 * cualquiera que leerGastos marque conPresupuesto:false). Ese monto tiene
 * que ir a MONTO para que B4 lo reste. pagado:false deja el monto en PRESUP.
 */
function nuevoGasto(p) {
  var b = bloqueEgresos_(p.categoria);
  var row = -1;
  for (var r = b.first; r <= b.last; r++) {
    if (!b.sh.getRange(r, 3).getValue()) { row = r; break; }
  }
  if (row < 0) throw new Error('No hay fila libre dentro del bloque de «' + p.categoria + '» (antes del Total).');
  var fecha = p.fecha;
  if (typeof fecha === 'string' && fecha.indexOf('-') > -1) {
    var parts = fecha.split('-');
    fecha = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  }
  b.sh.getRange(row, 2).setValue(fecha);
  b.sh.getRange(row, 3).setValue(p.nombre);
  var monto = Number(p.monto);
  var comoPagado = p.pagado === true || b.presup < 0;
  if (comoPagado) {
    if (b.monto < 0) throw new Error('«' + p.categoria + '» no tiene columna MONTO.');
    b.sh.getRange(row, b.monto).setValue(monto);
  } else {
    b.sh.getRange(row, b.presup).setValue(monto);
  }
  return { ok: true, fila: row, pagado: comoPagado };
}

/** Suma cada MONTO del bloque, incluidas filas nuevas que sigan antes del Total. */
function sumaMontoBloque_(nombre) {
  var b = bloqueEgresos_(nombre);
  if (b.monto < 0) return 0;
  var vals = b.sh.getRange(b.first, b.monto, b.last - b.first + 1, 1).getValues();
  var s = 0;
  vals.forEach(function (cell) { var n = Number(cell[0]); if (isFinite(n)) s += n; });
  return s;
}

/**
 * Usar este número como saldoEgresos en leerConteo (y como cuadre = conteo − saldo).
 * Así un gasto nuevo anotado en MONTO entra aunque no esté en la suma fija de B4.
 */
function saldoEgresosDesdeMontos_() {
  var suma = ['Servicios', 'Salud', 'Artículos personales', 'Mercado', 'Ocio', 'Movilidad']
    .reduce(function (a, n) { return a + sumaMontoBloque_(n); }, 0);
  var koda = Number(SpreadsheetApp.getActive().getSheetByName('KODA').getRange('E22').getValue()) || 0;
  var ingresos = Number(SpreadsheetApp.getActive().getSheetByName('INGRESOS').getRange('H14').getValue()) || 0;
  return Math.round((ingresos - suma - koda) * 100) / 100;
}
