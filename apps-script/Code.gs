/**
 * ABAS-TCT · Backend de pedidos sobre Google Sheets
 * Pégalo en: Hoja ABAS-TCT_Catalogo › Extensiones › Apps Script (reemplaza todo el contenido).
 * Luego: Implementar › Nueva implementación › Aplicación web
 *        Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario › Implementar › Autorizar.
 *
 * Crea y usa estas pestañas:
 *   Pedidos     → un pedido por fila. Aquí validas el pago y confirmas fecha y franja.
 *   Detalle     → una fila por producto pedido (se llena sola).
 *   Consolidado → kilos y bolsas por producto para pedir al proveedor, por fecha de entrega.
 *   Resumen     → ventas, costos y ganancia del mes.
 *   Mercado     → precios de Campo Grande para sugerir márgenes (menú ABAS-TCT › Actualizar precios de mercado).
 *   Ajustes     → delivery, pedido mínimo, días de reparto, franjas y datos de pago.
 *                 Edítalos ahí: la web y el script los leen en cada pedido. No hace falta tocar este código.
 */

var TZ = 'America/Lima';
var TIENDA = 'ABAS-TCT';
// Valores por defecto. Los que valen son los de la pestaña "Ajustes".
var DEFAULTS = { delivery: 12, minimo: 50, anticipacion: 1, sinReparto: 'domingo', whatsapp: '51987609561',
  pagoNumero: '987609561', pagoTitular: 'Dante Daniel Clavijo Bazalar', costoMoto: 0,
  franjas: ['Mañana (9 a 12)', 'Tarde (12 a 4)', 'Noche (4 a 8)'] };
var ESTADOS = ['Pago en validación', 'Pago validado', 'Por preparar', 'Preparado', 'En ruta', 'Entregado', 'Observado', 'Pago rechazado', 'Cancelado'];
var COLS = ['Fecha pedido', 'Código', 'Cliente', 'Celular', 'Distrito', 'Dirección', 'Referencia', 'Productos',
  'Subtotal', 'Delivery', 'Total', 'Op. Yape', 'Franja preferida', 'Entrega propuesta', 'Estado',
  'Entrega confirmada', 'Franja confirmada', 'Nota al cliente', 'WhatsApp al cliente', 'Comentario cliente',
  'Historial', 'Revisión', 'Costo productos', 'Costo delivery real', 'Ganancia', 'Margen %'];
var PAGADOS = ['Pago validado', 'Por preparar', 'Preparado', 'En ruta', 'Entregado'];
// Columnas (1 = A): B Código, D Celular, K Total, O Estado, P Entrega confirmada, Q Franja confirmada, R Nota

function setup() { ensureSheets_(); }

function ensureSheets_() {
  var ss = SpreadsheetApp.getActive();
  var p = ss.getSheetByName('Pedidos');
  if (!p) {
    p = ss.insertSheet('Pedidos');
    p.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1b3a8a').setWrap(true);
    p.setFrozenRows(1); p.setFrozenColumns(3);
    p.getRange('O2:O').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).build());
    p.getRange('P2:P').setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build()).setNumberFormat('ddd dd/mm/yyyy');
    p.getRange('N2:N').setNumberFormat('ddd dd/mm/yyyy');
    p.getRange('A2:A').setNumberFormat('dd/mm/yyyy HH:mm');
    p.getRange('I2:K').setNumberFormat('"S/ "0.00');
    p.getRange('O2:Q').setBackground('#fff4d6');
    p.getRange('R2:R').setBackground('#fff4d6');
    [130, 120, 160, 100, 110, 220, 160, 320, 80, 70, 80, 90, 120, 110, 140, 120, 120, 200, 200, 200, 220, 200]
      .forEach(function (w, i) { p.setColumnWidth(i + 1, w); });
  }
  var a = ss.getSheetByName('Ajustes');
  if (!a) {
    a = ss.insertSheet('Ajustes');
    a.getRange('A1:C1').setValues([['Parámetro', 'Valor', 'Nota']]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1b3a8a');
    a.getRange('A2:C8').setValues([
      ['Delivery (S/)', DEFAULTS.delivery, 'Costo de envío por pedido a toda Lima Metropolitana'],
      ['Pedido mínimo (S/)', DEFAULTS.minimo, 'Incluye el delivery. Ej: 50 = S/ 38 de productos + S/ 12 de envío'],
      ['Días de anticipación', DEFAULTS.anticipacion, '1 = se entrega al día siguiente de la compra'],
      ['Días sin reparto', DEFAULTS.sinReparto, 'Separados por coma. Ej: domingo  ·  domingo, feriado se maneja a mano'],
      ['WhatsApp de la tienda', "'" + DEFAULTS.whatsapp, 'Con 51 adelante'],
      ['Número para pagos', "'" + DEFAULTS.pagoNumero, 'Yape o Plin'],
      ['Titular de la cuenta', DEFAULTS.pagoTitular, 'Se muestra al cliente al pagar']]);
    a.getRange('A9:C9').setValues([['Costo motorizado por pedido (S/)', DEFAULTS.costoMoto, 'Costo real que pagas por envío. Se copia a cada pedido nuevo; ponlo en 0 en los que repartes tú']]);
    a.getRange('A10').setValue('Franjas horarias (una por fila)').setFontWeight('bold');
    a.getRange(11, 1, DEFAULTS.franjas.length, 1).setValues(DEFAULTS.franjas.map(function (f) { return [f]; }));
    a.getRange('B2:B9').setBackground('#fff4d6'); a.getRange('A11:A16').setBackground('#fff4d6');
    a.setColumnWidth(1, 240); a.setColumnWidth(2, 220); a.setColumnWidth(3, 420);
    p.getRange('Q2:Q').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(a.getRange('A11:A16'), true).setAllowInvalid(true).build());
  }
  var d = ss.getSheetByName('Detalle');
  if (!d) {
    d = ss.insertSheet('Detalle');
    d.getRange(1, 1, 1, 10).setValues([['Código', 'Fecha pedido', 'Producto', 'Presentación (g)', 'Bolsas', 'Kg', 'Precio unit.', 'Subtotal', 'Estado', 'Entrega']])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#1b3a8a');
    d.setFrozenRows(1);
    d.getRange('J2:J').setNumberFormat('dd/mm/yyyy');
  }
  var c = ss.getSheetByName('Consolidado');
  if (!c) {
    c = ss.insertSheet('Consolidado');
    c.getRange('A1').setValue('Fecha de entrega a consolidar:').setFontWeight('bold');
    c.getRange('B1').setFormula('=IF(WEEKDAY(TODAY())=7,TODAY()+2,TODAY()+1)').setNumberFormat('ddd dd/mm/yyyy').setBackground('#fff4d6');
    c.getRange('A2').setValue('Incluye pedidos en estado "Pago validado" o "Por preparar" con esa fecha de entrega confirmada.').setFontColor('#5d6577');
    c.getRange('A4').setValue('Por presentación').setFontWeight('bold');
    c.getRange('A5').setFormula('=IFERROR(QUERY(Detalle!A2:J,"select C, D, sum(E), sum(F) where (I = \'Pago validado\' or I = \'Por preparar\') and J = date \'"&TEXT(B1,"yyyy-mm-dd")&"\' group by C, D order by C label C \'Producto\', D \'Presentación (g)\', sum(E) \'Bolsas\', sum(F) \'Kg\'",0),"Sin pedidos para esa fecha")');
    c.getRange('G4').setValue('Total a pedir al proveedor').setFontWeight('bold');
    c.getRange('G5').setFormula('=IFERROR(QUERY(Detalle!A2:J,"select C, sum(F) where (I = \'Pago validado\' or I = \'Por preparar\') and J = date \'"&TEXT(B1,"yyyy-mm-dd")&"\' group by C order by C label C \'Producto\', sum(F) \'Kg total\'",0),"")');
    c.setColumnWidth(1, 220); c.setColumnWidth(7, 220);
  }
  migrate_(p, d, a);
  ensureMercado_();
  var rs = ss.getSheetByName('Resumen');
  if (!rs) {
    rs = ss.insertSheet('Resumen');
    var paid = '{"Pago validado";"Por preparar";"Preparado";"En ruta";"Entregado"}';
    var m = '(TEXT(Pedidos!$A$2:$A,"yyyy-mm")=$B$1)*ISNUMBER(MATCH(Pedidos!$O$2:$O,' + paid + ',0))';
    rs.getRange('A1').setValue('Mes (aaaa-mm)').setFontWeight('bold');
    rs.getRange('B1').setFormula('=TEXT(TODAY(),"yyyy-mm")').setBackground('#fff4d6');
    rs.getRange('A2').setValue('Cuenta solo pedidos con pago validado o posterior. Cambia el mes en B1 (ej. 2026-11).').setFontColor('#5d6577');
    var rows = [
      ['Pedidos pagados', '=SUMPRODUCT(' + m + ')', '0'],
      ['Ventas (productos + delivery cobrado)', '=SUMPRODUCT(' + m + ',IFERROR(Pedidos!$K$2:$K*1,0))', '"S/ "0.00'],
      ['Delivery cobrado', '=SUMPRODUCT(' + m + ',IFERROR(Pedidos!$J$2:$J*1,0))', '"S/ "0.00'],
      ['Costo de productos', '=SUMPRODUCT(' + m + ',IFERROR(Pedidos!$W$2:$W*1,0))', '"S/ "0.00'],
      ['Costo de delivery (motorizado)', '=SUMPRODUCT(' + m + ',IFERROR(Pedidos!$X$2:$X*1,0))', '"S/ "0.00'],
      ['Ganancia bruta', '=B5-B7-B8', '"S/ "0.00'],
      ['Margen sobre ventas', '=IF(B5=0,0,B9/B5)', '0.0%'],
      ['Ticket promedio', '=IF(B4=0,0,B5/B4)', '"S/ "0.00'],
      ['Ganancia promedio por pedido', '=IF(B4=0,0,B9/B4)', '"S/ "0.00']];
    rows.forEach(function (x, i) {
      rs.getRange(4 + i, 1).setValue(x[0]);
      rs.getRange(4 + i, 2).setFormula(x[1]).setNumberFormat(x[2]).setFontWeight('bold');
    });
    rs.setColumnWidth(1, 300); rs.setColumnWidth(2, 140);
  }
  return { p: p, d: d };
}

/* Agrega columnas de costos y ganancia a hojas creadas por versiones anteriores. */
function migrate_(p, d, a) {
  if (a && !a.getRange('A9').getValue()) {
    a.getRange('A9:C9').setValues([['Costo motorizado por pedido (S/)', DEFAULTS.costoMoto, 'Costo real que pagas por envío. Se copia a cada pedido nuevo; ponlo en 0 en los que repartes tú']]);
    a.getRange('B9').setBackground('#fff4d6');
  }
  if (!d.getRange('K1').getValue()) {
    d.getRange('K1:L1').setValues([['Costo unit.', 'Costo total']]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1b3a8a');
    var n = d.getLastRow() - 1;
    if (n > 0) {
      var cm = costMap_(), v = d.getRange(2, 3, n, 3).getValues();
      d.getRange(2, 11, n, 2).setValues(v.map(function (x) {
        var c = unitCost_(cm, x[0], x[1]); return [c, r2_(c * (Number(x[2]) || 0))];
      }));
    }
  }
  if (!p.getRange('W1').getValue()) {
    p.getRange('W1:Z1').setValues([['Costo productos', 'Costo delivery real', 'Ganancia', 'Margen %']]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1b3a8a').setWrap(true);
    for (var r = 2; r <= p.getLastRow(); r++) profitCells_(p, r, 0);
  }
  p.getRange('W2:W').setNumberFormat('"S/ "0.00'); p.getRange('X2:X').setNumberFormat('"S/ "0.00').setBackground('#fff4d6');
  p.getRange('Y2:Y').setNumberFormat('"S/ "0.00'); p.getRange('Z2:Z').setNumberFormat('0.0%');
}
function profitCells_(p, r, moto) {
  p.getRange(r, 23).setFormula('=IF(B' + r + '="","",SUMIF(Detalle!A:A,B' + r + ',Detalle!L:L))');
  p.getRange(r, 24).setValue(moto);
  p.getRange(r, 25).setFormula('=IF(B' + r + '="","",K' + r + '-W' + r + '-N(X' + r + '))');
  p.getRange(r, 26).setFormula('=IF(OR(B' + r + '="",N(K' + r + ')=0),"",Y' + r + '/K' + r + ')');
}
/* Costo actual por producto desde la pestaña Productos: costo/kg usado (I) y empaque por bolsa (N). */
function costMap_() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Productos'), m = {};
  if (!sh) return m;
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (v[i][2]) m[norm_(v[i][2])] = { kg: Number(v[i][8]) || 0, pack: Number(v[i][13]) || 0 };
  return m;
}
function unitCost_(cm, name, g) {
  var c = cm[norm_(name)];
  return c ? r2_(c.kg * (Number(g) || 0) / 1000 + c.pack) : 0;
}

/* Lee la pestaña "Ajustes" (con valores por defecto si falta algo). */
function getSettings_() {
  var a = SpreadsheetApp.getActive().getSheetByName('Ajustes');
  var s = JSON.parse(JSON.stringify(DEFAULTS));
  if (!a) return s;
  var v = a.getRange('B2:B9').getValues().map(function (r) { return r[0]; });
  var num = function (x, def) { var n = Number(String(x).replace(',', '.')); return isNaN(n) || String(x) === '' ? def : n; };
  s.delivery = num(v[0], s.delivery);
  s.minimo = num(v[1], s.minimo);
  s.anticipacion = Math.max(0, Math.round(num(v[2], s.anticipacion)));
  s.sinReparto = String(v[3] || '');
  if (v[4]) s.whatsapp = String(v[4]).replace(/\D/g, '');
  if (v[5]) s.pagoNumero = String(v[5]).replace(/\D/g, '');
  if (v[6]) s.pagoTitular = String(v[6]);
  s.costoMoto = num(v[7], 0);
  var f = a.getRange('A11:A16').getValues().map(function (r) { return String(r[0]).trim(); }).filter(String);
  if (f.length) s.franjas = f;
  return s;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var o = JSON.parse(e.postData.contents || '{}');
    if (o.website) return out_({ ok: true, code: 'X' });            // trampa anti-spam
    var tel = String(o.phone || '').replace(/\D/g, '').slice(-9);
    if (!o.customer || tel.length !== 9 || !o.items || !o.items.length || !o.address) return out_({ ok: false, error: 'Faltan datos del pedido' });
    var sh = ensureSheets_();
    var cfg = getSettings_(), DELIVERY = cfg.delivery, MINIMO = cfg.minimo;
    var prices = priceMap_(), cm = costMap_(), sub = 0, issues = [], lines = [];
    o.items.forEach(function (it) {
      var qty = Math.max(1, Math.round(Number(it.qty) || 0)), g = Math.round(Number(it.g) || 0);
      var real = prices[norm_(it.name) + '|' + g];
      if (real == null) { issues.push('Precio no encontrado: ' + it.name + ' ' + g + ' g'); real = Number(it.price) || 0; }
      sub += real * qty;
      lines.push({ name: it.name, g: g, qty: qty, price: real, cost: unitCost_(cm, it.name, g) });
    });
    sub = r2_(sub);
    var total = r2_(sub + DELIVERY);
    if (total < MINIMO) issues.push('Bajo el mínimo de S/ ' + MINIMO);
    if (o.total != null && Math.abs(Number(o.total) - total) > 0.01) issues.push('La web calculó S/ ' + o.total);
    var code = uniqueCode_(sh.p, String(o.code || '').replace(/[^\w-]/g, '') || ('FS-' + Utilities.formatDate(new Date(), TZ, 'MMdd') + '-' + Math.floor(1000 + Math.random() * 9000)));
    var now = new Date(), prop = proposedDate_(now, cfg);
    var row = sh.p.getLastRow() + 1;
    var text = lines.map(function (l) { return l.qty + ' × ' + l.name + ' ' + gLabel_(l.g); }).join('\n');
    sh.p.getRange(row, 1, 1, 18).setValues([[now, code, String(o.customer).slice(0, 80), "'" + tel, String(o.district || '').slice(0, 60),
      String(o.address || '').slice(0, 160), String(o.ref || '').slice(0, 160), text, sub, DELIVERY, total,
      String(o.op || '').slice(0, 30), String(o.when || ''), prop, 'Pago en validación', prop, '', '']]);
    sh.p.getRange(row, 19).setFormula(waFormula_(row));
    sh.p.getRange(row, 20, 1, 3).setValues([[String(o.notes || '').slice(0, 300), stamp_() + ' Pago en validación', issues.join(' · ')]]);
    if (issues.length) sh.p.getRange(row, 22).setBackground('#fde6ec');
    profitCells_(sh.p, row, cfg.costoMoto);
    var dr = sh.d.getLastRow() + 1;
    var det = lines.map(function (l, i) {
      return [code, now, l.name, l.g, l.qty, r2_(l.g * l.qty / 1000), l.price, r2_(l.price * l.qty), '', '', l.cost, r2_(l.cost * l.qty)];
    });
    sh.d.getRange(dr, 1, det.length, 12).setValues(det);
    for (var i = 0; i < det.length; i++) {
      var r = dr + i;
      sh.d.getRange(r, 9).setFormula('=IFERROR(INDEX(Pedidos!O:O,MATCH(A' + r + ',Pedidos!B:B,0)),"")');
      sh.d.getRange(r, 10).setFormula('=IFERROR(INDEX(Pedidos!P:P,MATCH(A' + r + ',Pedidos!B:B,0)),"")');
    }
    return out_({ ok: true, code: code, subtotal: sub, delivery: DELIVERY, total: total, entrega: fmtDate_(prop) });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var q = (e && e.parameter) || {};
  if (q.action === 'config') { ensureSheets_(); var c = getSettings_(); c.ok = true; return out_(c); }
  if (q.action !== 'status') return out_({ ok: true, tienda: TIENDA });
  var code = String(q.code || '').trim().toUpperCase(), tel4 = String(q.tel || '').replace(/\D/g, '').slice(-4);
  var p = SpreadsheetApp.getActive().getSheetByName('Pedidos');
  if (!p || !code || tel4.length !== 4) return out_({ ok: false, error: 'Datos incompletos' });
  var v = p.getRange(2, 1, Math.max(1, p.getLastRow() - 1), 18).getValues();
  for (var i = v.length - 1; i >= 0; i--) {
    if (String(v[i][1]).toUpperCase() === code && String(v[i][3]).replace(/\D/g, '').slice(-4) === tel4) {
      var fecha = v[i][15] || v[i][13];
      return out_({ ok: true, code: v[i][1], estado: v[i][14], total: v[i][10], entrega: fecha ? fmtDate_(fecha) : '',
        franja: v[i][16] || '', confirmada: ['Pago validado', 'Por preparar', 'Preparado', 'En ruta', 'Entregado'].indexOf(v[i][14]) >= 0, nota: v[i][17] || '' });
    }
  }
  return out_({ ok: false, error: 'No encontramos ese pedido. Revisa el código y los 4 últimos dígitos de tu celular.' });
}

/* Registra en "Historial" cada cambio de estado. */
function onEdit(e) {
  var sh = e.range.getSheet();
  if (sh.getName() !== 'Pedidos' || e.range.getRow() < 2 || e.range.getColumn() !== 15) return;
  var h = sh.getRange(e.range.getRow(), 21);
  h.setValue((h.getValue() ? h.getValue() + '\n' : '') + stamp_() + ' ' + e.value);
}

/* ---------- utilidades ---------- */
function waFormula_(r) {
  var dia = 'CHOOSE(WEEKDAY(f),"domingo","lunes","martes","miércoles","jueves","viernes","sábado")&" "&TEXT(f,"dd/mm")';
  var msg = 'IFS(' +
    'O' + r + '="Pago validado","Hola "&C' + r + '&" 👋 Tu pago de S/ "&TEXT(K' + r + ',"0.00")&" fue validado ✅. Tu pedido "&B' + r + '&" se entregará el "&' + dia + '&IF(q="",""," en el horario "&q)&"."&IF(R' + r + '<>""," "&R' + r + ',"")&" ¡Gracias por comprar en ' + TIENDA + '!",' +
    'O' + r + '="Por preparar","Hola "&C' + r + '&", estamos preparando tu pedido "&B' + r + '&" para el "&' + dia + '&".",' +
    'O' + r + '="En ruta","Hola "&C' + r + '&", tu pedido "&B' + r + '&" ya está en camino 🛵"&IF(q="",""," ("&q&")")&".",' +
    'O' + r + '="Entregado","¡Pedido "&B' + r + '&" entregado! 🎉 Gracias por comprar en ' + TIENDA + '.",' +
    'O' + r + '="Pago rechazado","Hola "&C' + r + '&", no pudimos validar el Yape del pedido "&B' + r + '&". ¿Nos envías la captura con el N° de operación?",' +
    'O' + r + '="Observado","Hola "&C' + r + '&", sobre tu pedido "&B' + r + '&": "&R' + r + ',' +
    'O' + r + '="Cancelado","Hola "&C' + r + '&", tu pedido "&B' + r + '&" fue cancelado."&IF(R' + r + '<>""," "&R' + r + ',""),' +
    'TRUE,"Hola "&C' + r + '&", recibimos tu pedido "&B' + r + '&" por S/ "&TEXT(K' + r + ',"0.00")&". Estamos validando tu pago.")';
  return '=IF(D' + r + '="","",LET(f,IF(P' + r + '="",N' + r + ',P' + r + '),q,Q' + r + ',HYPERLINK("https://wa.me/51"&RIGHT(REGEXREPLACE(D' + r + '&"","\\D",""),9)&"?text="&ENCODEURL(' + msg + '),"Enviar: "&O' + r + ')))';
}
function priceMap_() {
  var t = SpreadsheetApp.getActive().getSheetByName('Tienda');
  var m = {};
  if (!t) return m;
  var v = t.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    var name = v[i][2];
    if (!name || String(v[i][0]).toUpperCase() === 'NO') continue;
    [[5, 6], [7, 8], [9, 10]].forEach(function (c) {
      var g = Math.round(Number(v[i][c[0]]) || 0), pr = Number(v[i][c[1]]);
      if (g && pr) m[norm_(name) + '|' + g] = pr;
    });
  }
  return m;
}
function proposedDate_(now, cfg) {
  cfg = cfg || getSettings_();
  var ymd = Utilities.formatDate(now, TZ, 'yyyy-MM-dd').split('-');
  var d = new Date(Number(ymd[0]), Number(ymd[1]) - 1, Number(ymd[2]), 12, 0, 0);
  d.setDate(d.getDate() + cfg.anticipacion);
  var dias = ['domingo', 'lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado'];
  var off = norm_(cfg.sinReparto).split(/[ ,;]+/).filter(String);
  for (var i = 0; i < 7 && off.indexOf(dias[d.getDay()]) >= 0; i++) d.setDate(d.getDate() + 1);
  return d;
}
function uniqueCode_(p, code) {
  var used = p.getLastRow() > 1 ? p.getRange(2, 2, p.getLastRow() - 1, 1).getValues().map(function (x) { return String(x[0]); }) : [];
  var c = code, n = 1;
  while (used.indexOf(c) >= 0) c = code + '-' + (++n);
  return c;
}
function norm_(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); }
function gLabel_(g) { return g >= 1000 ? (g / 1000 + ' kg').replace('.', ',') : g + ' g'; }
function r2_(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function stamp_() { return Utilities.formatDate(new Date(), TZ, 'dd/MM HH:mm'); }
function fmtDate_(d) {
  var dias = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
  d = new Date(d);
  return dias[d.getDay()] + ' ' + Utilities.formatDate(d, TZ, 'dd/MM');
}
function out_(x) { return ContentService.createTextOutput(JSON.stringify(x)).setMimeType(ContentService.MimeType.JSON); }

/* =================== Análisis de mercado (Campo Grande) =================== */
var MERCADO_FECHA = '07/11/2026';
var GRAMOS = [100, 250, 500, 1000];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('ABAS-TCT')
    .addItem('Actualizar precios de mercado ahora', 'actualizarMercado')
    .addItem('Aplicar márgenes sugeridos', 'aplicarMargenes')
    .addSeparator()
    .addItem('Activar actualización semanal (lunes 6 a. m.)', 'activarSemanal')
    .addToUi();
}

/* Crea la pestaña Mercado, las filas de estrategia en Ajustes y las columnas de sugerencia en Productos. */
function ensureMercado_() {
  var ss = SpreadsheetApp.getActive();
  var m = ss.getSheetByName('Mercado');
  if (!m) {
    m = ss.insertSheet('Mercado');
    m.getRange(1, 1, 1, 8).setValues([['Mi producto', 'URL Campo Grande', '100 g', '250 g', '500 g', '1 kg', 'Actualizado', 'Estado']])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#1b3a8a');
    m.setFrozenRows(1);
    var rows = MERCADO_INICIAL.map(function (r) { return [r[0], r[1], r[2] || '', r[3] || '', r[4] || '', r[5] || '', MERCADO_FECHA, 'Carga inicial']; });
    m.getRange(2, 1, rows.length, 8).setValues(rows);
    m.getRange('C2:F').setNumberFormat('"S/ "0.00');
    m.setColumnWidth(1, 200); m.setColumnWidth(2, 380);
    m.getRange(rows.length + 3, 1).setValue('Para comparar un producto nuevo: escribe su nombre exacto (como en Productos) y pega la URL del producto en Campo Grande.').setFontColor('#5d6577');
  }
  var a = ss.getSheetByName('Ajustes');
  if (a && !a.getRange('A18').getValue()) {
    a.getRange('A18').setValue('Estrategia de precios (análisis de mercado)').setFontWeight('bold');
    a.getRange('A19:C23').setValues([
      ['Competidor de referencia', 'Campo Grande', 'Se actualiza desde la pestaña Mercado (menú ABAS-TCT)'],
      ['Precio bajo el mercado (%)', 5, 'Cuánto más barato que Campo Grande quieres estar'],
      ['Margen mínimo (%)', 15, 'Nunca sugerir menos que esto sobre el costo'],
      ['Margen máximo (%)', 60, 'Nunca sugerir más que esto sobre el costo'],
      ['Aplicar márgenes sugeridos automáticamente', 'NO', 'SI = cada actualización semanal copia el margen sugerido a la columna Margen %']]);
    a.getRange('A25').setValue('Margen cuando no hay precio de mercado (por categoría)').setFontWeight('bold');
    a.getRange('A26:B30').setValues([['Frutos secos', 25], ['Pasas y deshidratados', 30], ['Semillas y granos', 35], ['Especias', 40], ['Conservas', 25]]);
    a.getRange('B20:B23').setBackground('#fff4d6'); a.getRange('B26:B30').setBackground('#fff4d6');
  }
  var p = ss.getSheetByName('Productos');
  if (p && !p.getRange('S1').getValue()) {
    p.getRange('S1:X1').setValues([['Mercado pres. 1', 'Mercado pres. 2', 'Mercado pres. 3', 'Margen sugerido %', 'Precio sugerido pres. 1', 'Tu precio vs mercado']])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#0e7c86').setWrap(true);
    var n = Math.max(150, p.getLastRow());
    var f = [];
    for (var r = 2; r <= n + 1; r++) {
      var mk = function (col) { return '=IF(OR($C' + r + '="",' + col + r + '=""),"",IFERROR(1/(1/INDEX(Mercado!$C:$F,MATCH($C' + r + ',Mercado!$A:$A,0),MATCH(' + col + r + ',{100,250,500,1000},0))),""))'; };
      var c = function (mcol, gcol) { return 'IF(OR(' + mcol + r + '="",' + gcol + r + '=""),1E9,(' + mcol + r + '*(1-Ajustes!$B$20/100)/($I' + r + '*' + gcol + r + '/1000+N($N' + r + '))-1)*100)'; };
      f.push([mk('K'), mk('L'), mk('M'),
        '=IF(OR($C' + r + '="",$I' + r + '=""),"",LET(x,MIN(' + c('S', 'K') + ',' + c('T', 'L') + ',' + c('U', 'M') + '),IF(x>=1E8,IFERROR(VLOOKUP($B' + r + ',Ajustes!$A$26:$B$30,2,FALSE),30),ROUND(MIN(MAX(x,Ajustes!$B$21),Ajustes!$B$22),0))))',
        '=IF(V' + r + '="","",CEILING(($I' + r + '*K' + r + '/1000+N($N' + r + '))*(1+V' + r + '/100),0.5))',
        '=IF(OR(O' + r + '="",S' + r + '=""),"",O' + r + '/S' + r + '-1)']);
    }
    p.getRange(2, 19, f.length, 6).setFormulas(f);
    p.getRange('S2:U').setNumberFormat('"S/ "0.00'); p.getRange('W2:W').setNumberFormat('"S/ "0.00'); p.getRange('X2:X').setNumberFormat('+0%;-0%;0%');
    p.setColumnWidths(19, 6, 105);
  }
  return m;
}

/* Lee los precios actuales de cada URL de la pestaña Mercado. */
function actualizarMercado() {
  ensureSheets_();
  var m = ensureMercado_();
  var last = m.getLastRow();
  var v = m.getRange(2, 1, Math.max(1, last - 1), 2).getValues();
  var idx = [], reqs = [];
  v.forEach(function (row, i) {
    if (row[0] && /^https?:\/\//.test(String(row[1]))) {
      idx.push(i);
      reqs.push({ url: String(row[1]), muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0 (compatible; ABAS-TCT precios)' } });
    }
  });
  var ok = 0, fail = 0, now = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
  for (var b = 0; b < reqs.length; b += 15) {
    var res = UrlFetchApp.fetchAll(reqs.slice(b, b + 15));
    res.forEach(function (resp, k) {
      var r = idx[b + k] + 2, prices = null;
      try { if (resp.getResponseCode() === 200) prices = parseVariaciones_(resp.getContentText()); } catch (e) {}
      if (prices) {
        m.getRange(r, 3, 1, 6).setValues([[prices[100] || '', prices[250] || '', prices[500] || '', prices[1000] || '', now, 'OK']]);
        ok++;
      } else {
        m.getRange(r, 7, 1, 2).setValues([[now, 'No se pudo leer (HTTP ' + resp.getResponseCode() + ')']]);
        fail++;
      }
    });
    Utilities.sleep(800);
  }
  var a = SpreadsheetApp.getActive().getSheetByName('Ajustes');
  if (a && /^s[ií]/i.test(String(a.getRange('B23').getValue()))) aplicarMargenes_(false);
  SpreadsheetApp.getActive().toast('Precios de mercado: ' + ok + ' actualizados, ' + fail + ' con error.', 'ABAS-TCT', 8);
}

function parseVariaciones_(html) {
  var mt = html.match(/data-product_variations="([^"]*)"/);
  if (!mt) return null;
  var json = mt[1].replace(/&quot;/g, '"').replace(/&#039;/g, "'").replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
  var vars = JSON.parse(json), out = {}, n = 0;
  vars.forEach(function (v) {
    var lab = Object.keys(v.attributes || {}).map(function (k) { return v.attributes[k]; }).join(' ').toLowerCase();
    var g = /kg-?\s*1\b/.test(lab) && !/kg-?\s*1\d/.test(lab) ? 1000 : (lab.match(/gr?-?\s*(100|250|500)\b/) || [])[1];
    g = Number(g);
    if (GRAMOS.indexOf(g) >= 0 && v.display_price) { out[g] = Number(v.display_price); n++; }
  });
  return n ? out : null;
}

function aplicarMargenes() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.alert('Aplicar márgenes sugeridos', 'Se copiará el "Margen sugerido %" a la columna "Margen %" de todos los productos activos. ¿Continuar?', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  var n = aplicarMargenes_(true);
  ui.alert(n + ' productos actualizados. Los precios de la web cambiarán en unos minutos.');
}
function aplicarMargenes_(interactive) {
  var p = SpreadsheetApp.getActive().getSheetByName('Productos');
  if (!p) return 0;
  var last = p.getLastRow(), act = p.getRange(2, 1, last - 1, 1).getValues(), cur = p.getRange(2, 10, last - 1, 1).getValues(), sug = p.getRange(2, 22, last - 1, 1).getValues(), n = 0;
  var out = cur.map(function (c, i) {
    var s = sug[i][0];
    if (String(act[i][0]).toUpperCase() !== 'NO' && typeof s === 'number' && s !== c[0]) { n++; return [s]; }
    return [c[0]];
  });
  p.getRange(2, 10, out.length, 1).setValues(out);
  return n;
}

function activarSemanal() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'actualizarMercado') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('actualizarMercado').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).inTimezone(TZ).create();
  SpreadsheetApp.getUi().alert('Listo: los precios de mercado se actualizarán cada lunes a las 6 a. m.');
}

var MERCADO_INICIAL = [["Pecana grande", "https://campograndeperu.com/producto/pecana-pelada-grande/", 6.44, 14.62, 27.54, 51.35], ["Pecana mediana", "https://campograndeperu.com/producto/pecana-pelada-mediana/", 6.19, 13.93, 26, 48.75], ["Pecana partida", "https://campograndeperu.com/producto/pecana-partida/", 6.91, 14.39, 26.76, 49.99], ["Nuez pelada", "https://campograndeperu.com/producto/nuez-pelada-de-nogal-mariposa/", 6.27, 13.07, 24.3, 45.39], ["Almendra cruda", "https://campograndeperu.com/producto/almendra-natural/", 5.33, 12.7, 21.96, 41.67], ["Almendra tostada", "https://campograndeperu.com/producto/almendras-tostadas/", 5.92, 13.65, 23.67, 46.1], ["Maní tostado partido", "https://campograndeperu.com/producto/mani-partido/", 1.95, 4.06, 7.5, 13.92], ["Maní bola", "https://campograndeperu.com/producto/mani-tostado/", 2.25, 4.69, 8.65, 16.06], ["Maní salado", "https://campograndeperu.com/producto/mani-con-sal-frito/", 2.1, 4.38, 8.08, 15], ["Pistacho", "https://campograndeperu.com/producto/pistachos-tostados/", 8.75, 20.9, 40, 70.89], ["Cashew crudo", "https://campograndeperu.com/producto/nuez-de-caju-crudo-cashew/", 6.64, 15.56, 29.28, 55.31], ["Cashew tostado", "https://campograndeperu.com/producto/nuez-de-caju-tostado-cashew/", 6.94, 15.77, 28.91, 53.37], ["Castaña amazónica", "https://campograndeperu.com/producto/castana-primera/", 11.54, 26.79, 50, 93.75], ["Semilla de girasol", "https://campograndeperu.com/producto/semilla-de-girasol-crudo/", 3.06, 6.42, 11.71, 22.13], ["Semilla de girasol tostada", "https://campograndeperu.com/producto/semilla-de-girasol-tostado/", 2.88, 6.05, 11.36, 20.83], ["Semilla de calabaza", "https://campograndeperu.com/producto/semilla-de-calabaza-crudo/", 5.36, 11.23, 20.48, 38.68], ["Semilla de calabaza tostada", "https://campograndeperu.com/producto/semilla-de-calabaza-tostado/", 5.81, 12.18, 22.21, 41.96], ["Avellana", "https://campograndeperu.com/producto/avellanas-tostadas/", 10, 22.83, 42.86, 80.77], ["Pasa importada", "https://campograndeperu.com/producto/pasas-morenas-jumbo/", 2.45, 5.62, 10.38, 19.28], ["Pasa bebé", "https://campograndeperu.com/producto/pasas-morenas-chicas/", 2.15, 4.92, 9.22, 17.35], ["Pasa nacional", "https://campograndeperu.com/producto/pasas-morenas/", 2.79, 6.4, 11.82, 22.59], ["Pasa rubia grande", "https://campograndeperu.com/producto/pasas-rubias/", 4.35, 9.18, 17.02, 32.63], ["Pasa rubia bebé", "https://campograndeperu.com/producto/pasas-rubias/", 4.35, 9.18, 17.02, 32.63], ["Guindón sin pepa", "https://campograndeperu.com/producto/guindones-deshidratados-sin-pepa/", 4.06, 8.27, 15.36, 28.67], ["Damasco", "https://campograndeperu.com/producto/damasco-deshidratado/", 7.5, 17.31, 32.14, 56.25], ["Arándano", "https://campograndeperu.com/producto/arandanos-rojos-deshidratados/", 4.04, 8.75, 15.91, 30.01], ["Dátil natural", "https://campograndeperu.com/producto/datiles-medjool-importado/", 2.72, 5.71, 10.41, 19.67], ["Dátil glaseado", "https://campograndeperu.com/producto/datiles-glaseados/", 3.75, 8.52, 15.63, 28.85], ["Huesillo", "https://campograndeperu.com/producto/huesillos/", 5.77, 12.03, 22.21, 41.25], ["Cereza", "https://campograndeperu.com/producto/cereza-deshidratada-sin-pepa/", 5.62, 11.25, 20.77, 38.57], ["Naranja", "https://campograndeperu.com/producto/naranja-deshidratada/", 6.6, 13.75, 25.39, 47.15], ["Manzana", "https://campograndeperu.com/producto/manzana-deshidratada/", 11.16, 25.57, 47.2, 87.66], ["Mango", "https://campograndeperu.com/producto/mango-deshidratado/", 8.7, 18.12, 33.46, 62.14], ["Piña", "https://campograndeperu.com/producto/pina-deshidratada/", 9.2, 19.17, 35.38, 65.71], ["Higo", "https://campograndeperu.com/producto/higos-secos/", 3.3, 6.88, 12.91, 24.29], ["Fresa deshidratada", "https://campograndeperu.com/producto/fresa-deshidratada/", 11.33, 22.84, 42.27, 78.67], ["Aguaymanto", "https://campograndeperu.com/producto/aguaymanto-deshidratado/", 6.9, 14.71, 27.78, 51.95], ["Coco rallado", "https://campograndeperu.com/producto/coco-rallado-largo/", 3.43, 7.87, 14.75, 27.76], ["Coco fino", "https://campograndeperu.com/producto/coco-rallado-fino/", 3.07, 5.99, 11.28, 21.31], ["Guinda", "https://campograndeperu.com/producto/guinda-deshidratada/", 2.83, 5.9, 10.89, 20.23], ["Avena", "https://campograndeperu.com/producto/hojuela-de-avena/", null, null, 3.16, 5.75], ["Ajonjolí crudo", "https://campograndeperu.com/producto/ajonjoli-crudo/", 1.45, 2.82, 5.07, 9.23], ["Ajonjolí tostado", "https://campograndeperu.com/producto/ajonjoli-tostado/", 1.77, 3.62, 6.71, 12.53], ["Ajonjolí negro", "https://campograndeperu.com/producto/ajonjoli-negro/", 2.46, 4.92, 8.43, 15.95], ["Chía", "https://campograndeperu.com/producto/chia-negra/", 2.5, 5.25, 9.55, 17.5], ["Linaza nacional", "https://campograndeperu.com/producto/linaza-nacional/", 2.62, 6.1, 11.42, 21.45], ["Linaza importada", "https://campograndeperu.com/producto/linaza-importada/", 2.24, 4.2, 6.72, 8], ["Quinua", "https://campograndeperu.com/producto/quinua-blanca/", 2.16, 4.52, 8.24, 15.57], ["Flor de Jamaica", "https://campograndeperu.com/producto/flor-de-jamaica-deshidratada/", 3.26, 6.2, 11.27, 20.67], ["Cacao entero", "https://campograndeperu.com/producto/grano-de-cacao/", 5.7, 11.49, 21.27, 39.58], ["Cáscara de cacao", "https://campograndeperu.com/producto/cascarilla-de-cacao/", null, 2.29, 3.18, 5.9], ["Canela entera", "https://campograndeperu.com/producto/canela-entera/", 14.29, 30, 54.55, 100.01], ["Canela molida", "https://campograndeperu.com/producto/canela-molida/", 9.44, 19.97, 37.09, 69.23], ["Clavo de olor", "https://campograndeperu.com/producto/clavo-de-olor/", 7.02, 15.58, 28.83, 53.66], ["Anís estrella", "https://campograndeperu.com/producto/anis-estrella/", 6.77, 13.65, 25.26, 47], ["Anís en grano", "https://campograndeperu.com/producto/anis-espanol/", 3.9, 7.87, 14.56, 27.09], ["Laurel", "https://campograndeperu.com/producto/hoja-de-laurel/", 8.19, 18.9, 35.1, 65.51], ["Hongo seco", "https://campograndeperu.com/producto/hongos-secos/", 5.09, 10.77, 20, 37.33], ["Orégano fino", "https://campograndeperu.com/producto/oregano-fino/", 2.95, 5.71, 10.41, 19.67], ["Pimienta entera", "https://campograndeperu.com/producto/pimienta-negra-entera/", 5.18, 11.87, 21.92, 40.71], ["Pimienta molida", "https://campograndeperu.com/producto/pimienta-molida/", 8.5, 14.17, 24.29, 42.51], ["Comino molido", "https://campograndeperu.com/producto/comino-molido/", 3.54, 8.05, 14.75, 27.23], ["Cúrcuma", "https://campograndeperu.com/producto/curcuma-molida/", 2.16, 4.35, 8.05, 14.99], ["Pimentón molido", "https://campograndeperu.com/producto/pimenton-molido/", 2.78, 5.78, 10.67, 19.82], ["Romero molido", "https://campograndeperu.com/producto/romero-seco/", 3.4, 7.14, 13.01, 24.58], ["Orégano molido", "https://campograndeperu.com/producto/oregano-molido/", 3.17, 6.85, 12.72, 23.36], ["Bicarbonato", "https://campograndeperu.com/producto/bicarbonato-de-sodio/", null, null, 3.42, 6.27], ["Panela", "https://campograndeperu.com/producto/panela-granulada-organica/", null, null, 6.04, 11.7]];
