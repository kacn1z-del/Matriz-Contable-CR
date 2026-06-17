(function(){
// ════════════════════════════════════════════════════
// PATCH v14.5 — 5 MÓDULOS CONTABLES NUEVOS (TODO EN UN SOLO IIFE)
// Libro Mayor · Balance Comprobación · Flujo de Efectivo
// Conciliación Bancaria · Declaración D-101 (Renta CR)
// ════════════════════════════════════════════════════

// ── Helper: lee una hoja por índice (o la activa) y detecta columnas ──
function leerHojaActiva(sheetIdx) {
  var src = sheetData;
  if (sheetIdx !== undefined && sheetIdx !== null && typeof sheets !== 'undefined') {
    var idx = parseInt(sheetIdx);
    if (idx === activeSheet) src = sheetData;
    else if (sheets[idx]) src = sheets[idx].data || {};
  }
  var headers = [];
  var maxCol = 0, maxRow = 0;
  Object.keys(src).forEach(function(key){
    var m = key.match(/^([A-Z]+)(\d+)/);
    if (!m) return;
    var c = colIndex(m[1]);
    var r = parseInt(m[2]) - 1;
    if (c > maxCol) maxCol = c;
    if (r > maxRow) maxRow = r;
  });
  if (Object.keys(src).length === 0) return { headers:[], rows:[], colMap:{} };

  var headerRow = 0;
  for (var c0 = 0; c0 <= maxCol; c0++) {
    var v = src[cellId(headerRow, c0)];
    if (v) headers[c0] = String(v).toLowerCase().trim();
  }

  var colMap = { fecha:-1, cuenta:-1, debe:-1, haber:-1 };
  headers.forEach(function(h, i){
    if (!h) return;
    if (colMap.fecha < 0 && /fecha/.test(h)) colMap.fecha = i;
    if (colMap.cuenta < 0 && /(cuenta|descripci|concepto)/.test(h)) colMap.cuenta = i;
    if (colMap.debe < 0 && /debe|d[ée]bito/.test(h)) colMap.debe = i;
    if (colMap.haber < 0 && /haber|cr[ée]dito/.test(h)) colMap.haber = i;
  });
  if (colMap.fecha  < 0) colMap.fecha  = 0;
  if (colMap.cuenta < 0) colMap.cuenta = headers.length >= 4 ? 3 : (headers.length >= 2 ? 1 : 0);
  if (colMap.debe   < 0) colMap.debe   = headers.length >= 5 ? 4 : Math.max(0, headers.length - 2);
  if (colMap.haber  < 0) colMap.haber  = headers.length >= 6 ? 5 : Math.max(0, headers.length - 1);

  var rows = [];
  for (var r = headerRow + 1; r <= maxRow; r++) {
    var raw = [];
    var hasData = false;
    for (var c = 0; c <= maxCol; c++) {
      var val = src[cellId(r, c)] || '';
      if (val !== '') hasData = true;
      raw.push(val);
    }
    if (!hasData) continue;
    if (String(raw[colMap.cuenta] || '').toUpperCase().indexOf('TOTAL') === 0) continue;
    rows.push({
      fecha: colMap.fecha  >= 0 ? raw[colMap.fecha]  || '' : '',
      cuenta: colMap.cuenta >= 0 ? raw[colMap.cuenta] || '' : '',
      debe:  parseFloat(String(raw[colMap.debe]  || '0').replace(/[₡,\s]/g,'')) || 0,
      haber: parseFloat(String(raw[colMap.haber] || '0').replace(/[₡,\s]/g,'')) || 0,
      raw: raw
    });
  }
  return { headers: headers, rows: rows, colMap: colMap, maxRow: maxRow, maxCol: maxCol };
}

function fmtCRC(n) {
  return '₡' + Number(n||0).toLocaleString('es-CR', {minimumFractionDigits:2, maximumFractionDigits:2});
}

function populateSheetSelect(selId) {
  var sel = document.getElementById(selId);
  if (!sel) return;
  sel.innerHTML = '';
  (typeof sheets !== 'undefined' ? sheets : [{name:'Hoja 1'}]).forEach(function(s, i){
    var opt = document.createElement('option');
    opt.value = i;
    opt.textContent = s.name || ('Hoja ' + (i+1));
    if (i === (typeof activeSheet !== 'undefined' ? activeSheet : 0)) opt.selected = true;
    sel.appendChild(opt);
  });
}

function descargarBlob(content, filename, mime) {
  var blob = new Blob([content], {type: mime});
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

function nuevoPDF(orientation) {
  var jsPDFLib = window.jspdf && window.jspdf.jsPDF;
  if (!jsPDFLib) { alert('Exportar PDF no está disponible en este momento. Usá Ctrl+P como alternativa.'); return null; }
  return new jsPDFLib({ orientation: orientation||'portrait', unit:'mm', format:'a4' });
}

function pdfHeader(doc, titulo) {
  doc.setFillColor(45,122,12); doc.rect(0,0,doc.internal.pageSize.getWidth(),16,'F');
  doc.setTextColor(255,255,255); doc.setFontSize(14); doc.setFont('helvetica','bold');
  doc.text(titulo, 14, 10);
  doc.setFontSize(8); doc.setFont('helvetica','normal');
  doc.text('Matriz Contable CR · ' + new Date().toLocaleDateString('es-CR'), doc.internal.pageSize.getWidth()-14, 10, {align:'right'});
  doc.setTextColor(0,0,0);
}

// ════════════════════════════════════════════════════
// 1) LIBRO MAYOR AUTOMÁTICO
// ════════════════════════════════════════════════════
var _lmCuentas = {};

function abrirLibroMayor() {
  populateSheetSelect('lm-sheet-sel');
  openModal('modalLibroMayor');
  generarLibroMayor();
}

function generarLibroMayor() {
  var selEl = document.getElementById('lm-sheet-sel');
  var sheetIdx = selEl ? parseInt(selEl.value) : undefined;
  var data = leerHojaActiva(sheetIdx);
  var cont = document.getElementById('lmContenido');
  if (!cont) return;
  if (!data.rows.length) {
    cont.innerHTML = '<p style="color:var(--text3);font-size:12.5px;">No se detectaron movimientos contables en la hoja activa. Cargá una plantilla (Libro Diario) o ingresá datos con columnas Fecha / Cuenta / Debe / Haber.</p>';
    _lmCuentas = {};
    return;
  }
  var cuentas = {};
  data.rows.forEach(function(row){
    var key = row.cuenta || '(sin cuenta)';
    if (!cuentas[key]) cuentas[key] = [];
    cuentas[key].push(row);
  });

  var html = '';
  Object.keys(cuentas).sort().forEach(function(cuenta){
    var movs = cuentas[cuenta];
    var saldo = 0;
    html += '<h4 style="font-size:12.5px;font-weight:800;color:var(--g3);margin:10px 0 6px;">' + cuenta + '</h4>';
    html += '<table class="tbl-v14"><thead><tr><th>Fecha</th><th>Debe</th><th>Haber</th><th>Saldo</th></tr></thead><tbody>';
    movs.forEach(function(m){
      saldo += m.debe - m.haber;
      html += '<tr><td>' + (m.fecha||'') + '</td><td class="num">' + (m.debe?fmtCRC(m.debe):'') + '</td><td class="num">' + (m.haber?fmtCRC(m.haber):'') + '</td><td class="num">' + fmtCRC(saldo) + '</td></tr>';
    });
    html += '<tr class="total-row"><td>TOTAL</td><td class="num">' + fmtCRC(movs.reduce(function(a,m){return a+m.debe;},0)) + '</td><td class="num">' + fmtCRC(movs.reduce(function(a,m){return a+m.haber;},0)) + '</td><td class="num">' + fmtCRC(saldo) + '</td></tr>';
    html += '</tbody></table>';
  });
  cont.innerHTML = html;
  _lmCuentas = cuentas;
}

function exportarLibroMayorCSV() {
  var csv = 'Cuenta,Fecha,Debe,Haber,Saldo\n';
  Object.keys(_lmCuentas).forEach(function(cuenta){
    var saldo = 0;
    _lmCuentas[cuenta].forEach(function(m){
      saldo += m.debe - m.haber;
      csv += '"'+cuenta+'","'+(m.fecha||'')+'",'+m.debe+','+m.haber+','+saldo.toFixed(2)+'\n';
    });
  });
  descargarBlob('\uFEFF'+csv, 'LibroMayor_'+new Date().toISOString().split('T')[0]+'.csv', 'text/csv;charset=utf-8;');
}

function exportarLibroMayorPDF() {
  var doc = nuevoPDF('portrait'); if (!doc) return;
  pdfHeader(doc, 'Libro Mayor Automático');
  var y = 24;
  Object.keys(_lmCuentas).forEach(function(cuenta){
    if (y > 270) { doc.addPage(); y = 14; }
    doc.setFontSize(11); doc.setFont('helvetica','bold'); doc.setTextColor(45,122,12);
    doc.text(cuenta, 14, y); y += 5;
    doc.setFontSize(9); doc.setFont('helvetica','normal'); doc.setTextColor(0,0,0);
    var saldo = 0;
    _lmCuentas[cuenta].forEach(function(m){
      saldo += m.debe - m.haber;
      if (y > 280) { doc.addPage(); y = 14; }
      doc.text(String(m.fecha||''), 14, y);
      doc.text(m.debe?fmtCRC(m.debe):'-', 100, y, {align:'right'});
      doc.text(m.haber?fmtCRC(m.haber):'-', 140, y, {align:'right'});
      doc.text(fmtCRC(saldo), 196, y, {align:'right'});
      y += 5;
    });
    y += 4;
  });
  doc.save('LibroMayor_'+new Date().toISOString().split('T')[0]+'.pdf');
}

// ════════════════════════════════════════════════════
// 2) BALANCE DE COMPROBACIÓN EN TIEMPO REAL
// ════════════════════════════════════════════════════
var _bcCuentas = {};
var _bcTotales = { debe:0, haber:0 };

function abrirBalanceComp() {
  populateSheetSelect('bc-sheet-sel');
  openModal('modalBalanceComp');
  generarBalanceComp();
}

function generarBalanceComp() {
  var selEl = document.getElementById('bc-sheet-sel');
  var sheetIdx = selEl ? parseInt(selEl.value) : undefined;
  var data = leerHojaActiva(sheetIdx);
  var kpisEl = document.getElementById('bcKpis');
  var contEl = document.getElementById('bcContenido');
  if (!kpisEl || !contEl) return;

  if (!data.rows.length) {
    kpisEl.innerHTML = '';
    contEl.innerHTML = '<p style="color:var(--text3);font-size:12.5px;">No se detectaron movimientos contables en la hoja activa.</p>';
    _bcCuentas = {}; _bcTotales = {debe:0,haber:0};
    return;
  }

  var cuentas = {};
  data.rows.forEach(function(row){
    var key = row.cuenta || '(sin cuenta)';
    if (!cuentas[key]) cuentas[key] = { debe:0, haber:0 };
    cuentas[key].debe += row.debe;
    cuentas[key].haber += row.haber;
  });

  var totalDebe = 0, totalHaber = 0;
  var html = '<table class="tbl-v14"><thead><tr><th>Cuenta</th><th>Débitos</th><th>Créditos</th><th>Saldo D</th><th>Saldo A</th></tr></thead><tbody>';
  Object.keys(cuentas).sort().forEach(function(cuenta){
    var c = cuentas[cuenta];
    totalDebe += c.debe; totalHaber += c.haber;
    var saldo = c.debe - c.haber;
    html += '<tr><td>'+cuenta+'</td><td class="num">'+(c.debe?fmtCRC(c.debe):'')+'</td><td class="num">'+(c.haber?fmtCRC(c.haber):'')+'</td>'+
      '<td class="num">'+(saldo>0?fmtCRC(saldo):'')+'</td><td class="num">'+(saldo<0?fmtCRC(-saldo):'')+'</td></tr>';
  });
  var diff = totalDebe - totalHaber;
  html += '<tr class="total-row"><td>TOTALES</td><td class="num">'+fmtCRC(totalDebe)+'</td><td class="num">'+fmtCRC(totalHaber)+'</td><td class="num" colspan="2" style="text-align:right;">'+ (Math.abs(diff)<0.01 ? '✓ Cuadrado' : 'Diferencia: '+fmtCRC(Math.abs(diff))) +'</td></tr>';
  html += '</tbody></table>';
  contEl.innerHTML = html;

  var cuadrado = Math.abs(diff) < 0.01;
  kpisEl.innerHTML = '<div class="v14-kpis">' +
    '<div class="v14-kpi"><div class="v14-kpi-label">Total Débitos</div><div class="v14-kpi-value">'+fmtCRC(totalDebe)+'</div></div>' +
    '<div class="v14-kpi"><div class="v14-kpi-label">Total Créditos</div><div class="v14-kpi-value">'+fmtCRC(totalHaber)+'</div></div>' +
    '<div class="v14-kpi '+(cuadrado?'ok':'bad')+'"><div class="v14-kpi-label">'+(cuadrado?'Balance cuadrado':'Diferencia')+'</div><div class="v14-kpi-value">'+(cuadrado?'✓ OK':fmtCRC(Math.abs(diff)))+'</div></div>' +
    '</div>';

  _bcCuentas = cuentas;
  _bcTotales = { debe: totalDebe, haber: totalHaber };
}

function exportarBalanceCompCSV() {
  var csv = 'Cuenta,Debitos,Creditos,Saldo Deudor,Saldo Acreedor\n';
  Object.keys(_bcCuentas).sort().forEach(function(cuenta){
    var c = _bcCuentas[cuenta];
    var saldo = c.debe - c.haber;
    csv += '"'+cuenta+'",'+c.debe+','+c.haber+','+(saldo>0?saldo:0)+','+(saldo<0?-saldo:0)+'\n';
  });
  descargarBlob('\uFEFF'+csv, 'BalanceComprobacion_'+new Date().toISOString().split('T')[0]+'.csv', 'text/csv;charset=utf-8;');
}

function exportarBalanceCompPDF() {
  var doc = nuevoPDF('portrait'); if (!doc) return;
  pdfHeader(doc, 'Balance de Comprobación');
  var y = 24;
  doc.setFontSize(9); doc.setFont('helvetica','bold');
  doc.text('Cuenta', 14, y); doc.text('Débitos', 110, y, {align:'right'}); doc.text('Créditos', 150, y, {align:'right'}); doc.text('Saldo', 196, y, {align:'right'});
  y += 5;
  doc.setFont('helvetica','normal');
  Object.keys(_bcCuentas).sort().forEach(function(cuenta){
    if (y > 280) { doc.addPage(); y = 14; }
    var c = _bcCuentas[cuenta]; var saldo = c.debe - c.haber;
    doc.text(String(cuenta).substring(0,40), 14, y);
    doc.text(fmtCRC(c.debe), 110, y, {align:'right'});
    doc.text(fmtCRC(c.haber), 150, y, {align:'right'});
    doc.text((saldo>=0?'D ':'A ')+fmtCRC(Math.abs(saldo)), 196, y, {align:'right'});
    y += 5;
  });
  y += 3;
  doc.setFont('helvetica','bold');
  doc.text('TOTALES', 14, y);
  doc.text(fmtCRC(_bcTotales.debe), 110, y, {align:'right'});
  doc.text(fmtCRC(_bcTotales.haber), 150, y, {align:'right'});
  var diff2 = Math.abs(_bcTotales.debe - _bcTotales.haber);
  doc.setTextColor(diff2<0.01?45:220, diff2<0.01?122:38, diff2<0.01?12:38);
  doc.text(diff2<0.01?'✓ Cuadrado':'Dif: '+fmtCRC(diff2), 196, y, {align:'right'});
  doc.save('BalanceComprobacion_'+new Date().toISOString().split('T')[0]+'.pdf');
}

// ════════════════════════════════════════════════════
// 3) FLUJO DE EFECTIVO (DIRECTO E INDIRECTO)
// ════════════════════════════════════════════════════
var _feDirecto = [
  { concepto:'Cobros a clientes', monto:0, tipo:'operacion' },
  { concepto:'Pagos a proveedores', monto:0, tipo:'operacion' },
  { concepto:'Pagos a empleados', monto:0, tipo:'operacion' },
  { concepto:'Compra de activos fijos', monto:0, tipo:'inversion' },
  { concepto:'Préstamos recibidos', monto:0, tipo:'financiamiento' },
  { concepto:'Pago de dividendos', monto:0, tipo:'financiamiento' }
];

function abrirFlujoEfectivo() {
  openModal('modalFlujoEfectivo');
  renderFlujoDirecto();
  renderFlujoIndirecto();
}

function switchFlujoTab(tab) {
  ['directo','indirecto'].forEach(function(t){
    var tabEl = document.getElementById('feTab-'+t);
    var paneEl = document.getElementById('fePane-'+t);
    if (tabEl) tabEl.classList.toggle('active', t===tab);
    if (paneEl) paneEl.classList.toggle('active', t===tab);
  });
}

function renderFlujoDirecto() {
  var body = document.getElementById('feDirectoBody');
  if (!body) return;
  var html = '<table class="tbl-v14"><thead><tr><th>Concepto</th><th>Categoría</th><th>Monto (₡)</th><th></th></tr></thead><tbody>';
  _feDirecto.forEach(function(item, i){
    html += '<tr><td><input class="inp" style="font-size:12px;padding:4px 7px;" value="'+(item.concepto||'').replace(/"/g,'&quot;')+'" oninput="window._feSetConcepto('+i+',this.value)"></td>'+
      '<td><select class="v14-select-inline" onchange="window._feSetTipo('+i+',this.value)">'+
      ['operacion','inversion','financiamiento'].map(function(t){return '<option value="'+t+'"'+(item.tipo===t?' selected':'')+'>'+(t==='operacion'?'Operación':t==='inversion'?'Inversión':'Financiamiento')+'</option>';}).join('')+
      '</select></td>'+
      '<td><input class="inp" type="number" style="text-align:right;" value="'+item.monto+'" oninput="window._feSetMonto('+i+',this.value)"></td>'+
      '<td><button onclick="window._feEliminarFila('+i+')" style="background:none;border:none;color:var(--danger);font-size:15px;cursor:pointer;">×</button></td></tr>';
  });
  html += '</tbody></table>';
  body.innerHTML = html;
}

function agregarFilaFlujoDirecto() {
  _feDirecto.push({ concepto:'Nuevo concepto', monto:0, tipo:'operacion' });
  renderFlujoDirecto();
}

function _feSetConcepto(i, val) { if (_feDirecto[i]) _feDirecto[i].concepto = val; }
function _feSetTipo(i, val) { if (_feDirecto[i]) _feDirecto[i].tipo = val; }
function _feSetMonto(i, val) { if (_feDirecto[i]) _feDirecto[i].monto = parseFloat(val)||0; }
function _feEliminarFila(i) { _feDirecto.splice(i,1); renderFlujoDirecto(); }

function recalcularFlujoEfectivo() {
  renderFlujoDirecto();
  renderFlujoIndirecto();
}

function renderFlujoIndirecto() {
  var body = document.getElementById('feIndirectoBody');
  if (!body) return;
  // Estimar a partir de los nombres de cuenta detectados en la hoja activa
  var data = leerHojaActiva();
  var resultadoNeto = 0, depreciacion = 0, cambioCxC = 0, cambioCxP = 0, cambioInventario = 0;
  data.rows.forEach(function(row){
    var c = (row.cuenta||'').toLowerCase();
    var neto = row.debe - row.haber;
    if (/venta|ingreso/.test(c)) resultadoNeto -= neto; // ventas son crédito, suman al resultado
    if (/costo|gasto|compra/.test(c)) resultadoNeto += neto;
    if (/deprecia/.test(c)) depreciacion += Math.abs(neto);
    if (/cliente|cuenta.*cobrar/.test(c)) cambioCxC += neto;
    if (/proveedor|cuenta.*pagar/.test(c)) cambioCxP -= neto;
    if (/inventario|mercader/.test(c)) cambioInventario += neto;
  });
  var flujoOperacion = resultadoNeto + depreciacion - cambioCxC + cambioCxP - cambioInventario;

  var html = '<table class="tbl-v14"><thead><tr><th>Concepto</th><th>Monto (₡)</th></tr></thead><tbody>'+
    '<tr><td>Resultado neto del período (estimado)</td><td class="num">'+fmtCRC(resultadoNeto)+'</td></tr>'+
    '<tr><td>(+) Depreciación y amortización</td><td class="num">'+fmtCRC(depreciacion)+'</td></tr>'+
    '<tr><td>(–) Aumento en cuentas por cobrar</td><td class="num">'+fmtCRC(-cambioCxC)+'</td></tr>'+
    '<tr><td>(+) Aumento en cuentas por pagar</td><td class="num">'+fmtCRC(cambioCxP)+'</td></tr>'+
    '<tr><td>(–) Aumento en inventarios</td><td class="num">'+fmtCRC(-cambioInventario)+'</td></tr>'+
    '<tr class="total-row"><td>Flujo neto de actividades de operación</td><td class="num">'+fmtCRC(flujoOperacion)+'</td></tr>'+
    '</tbody></table>'+
    '<p style="font-size:11px;color:var(--text3);margin-top:6px;">Estimado automáticamente según nombres de cuenta detectados en la hoja activa (ventas/costos/depreciación/clientes/proveedores/inventario). Ajustá manualmente si tu plan de cuentas usa otra nomenclatura.</p>';
  body.innerHTML = html;
  window._feIndirectoFlujo = flujoOperacion;
}

function exportarFlujoCSV() {
  var csv = 'Concepto,Categoria,Monto\n';
  _feDirecto.forEach(function(item){
    csv += '"'+item.concepto+'","'+item.tipo+'",'+item.monto+'\n';
  });
  descargarBlob('\uFEFF'+csv, 'FlujoEfectivo_'+new Date().toISOString().split('T')[0]+'.csv', 'text/csv;charset=utf-8;');
}

function exportarFlujoPDF() {
  var doc = nuevoPDF('portrait'); if (!doc) return;
  pdfHeader(doc, 'Flujo de Efectivo');
  var y = 24;
  ['operacion','inversion','financiamiento'].forEach(function(cat){
    var items = _feDirecto.filter(function(i){return i.tipo===cat;});
    if (!items.length) return;
    doc.setFontSize(11); doc.setFont('helvetica','bold'); doc.setTextColor(45,122,12);
    doc.text(cat==='operacion'?'Actividades de Operación':cat==='inversion'?'Actividades de Inversión':'Actividades de Financiamiento', 14, y);
    y += 6;
    doc.setFontSize(9); doc.setFont('helvetica','normal'); doc.setTextColor(0,0,0);
    var subtotal = 0;
    items.forEach(function(item){
      if (y > 280) { doc.addPage(); y = 14; }
      doc.text(item.concepto, 14, y);
      doc.text(fmtCRC(item.monto), 196, y, {align:'right'});
      subtotal += item.monto;
      y += 5;
    });
    doc.setFont('helvetica','bold');
    doc.text('Subtotal', 14, y); doc.text(fmtCRC(subtotal), 196, y, {align:'right'});
    y += 8;
  });
  doc.save('FlujoEfectivo_'+new Date().toISOString().split('T')[0]+'.pdf');
}

// ════════════════════════════════════════════════════
// 4) CONCILIACIÓN BANCARIA INTERACTIVA
// ════════════════════════════════════════════════════
var _concAjustes = { banco: [], libros: [] };

function abrirConciliacionV14() {
  openModal('modalConciliacionV14');
  if (!_concAjustes.banco.length) _concAjustes.banco = [{concepto:'Depósitos en tránsito', monto:0},{concepto:'Cheques pendientes de cobro', monto:0}];
  if (!_concAjustes.libros.length) _concAjustes.libros = [{concepto:'Notas de crédito no registradas', monto:0},{concepto:'Notas de débito no registradas', monto:0}];
  renderAjustesConc('banco');
  renderAjustesConc('libros');
  recalcularConciliacionV14();
}

function renderAjustesConc(lado) {
  var cont = document.getElementById(lado==='banco'?'concBancoAjustes':'concLibrosAjustes');
  if (!cont) return;
  var html = '';
  _concAjustes[lado].forEach(function(item, i){
    html += '<div class="v14-conc-row">'+
      '<input type="text" class="inp" value="'+(item.concepto||'').replace(/"/g,'&quot;')+'" oninput="window._concSetConcepto(\''+lado+'\','+i+',this.value)">'+
      '<input type="number" class="inp" value="'+item.monto+'" oninput="window._concSetMonto(\''+lado+'\','+i+',this.value)">'+
      '<button onclick="window._concEliminarAjuste(\''+lado+'\','+i+')">×</button></div>';
  });
  cont.innerHTML = html;
}

function agregarAjusteConc(lado) {
  _concAjustes[lado].push({ concepto:'Ajuste', monto:0 });
  renderAjustesConc(lado);
  recalcularConciliacionV14();
}

function _concSetConcepto(lado, i, val) { if (_concAjustes[lado][i]) { _concAjustes[lado][i].concepto = val; recalcularConciliacionV14(); } }
function _concSetMonto(lado, i, val) { if (_concAjustes[lado][i]) { _concAjustes[lado][i].monto = parseFloat(val)||0; recalcularConciliacionV14(); } }
function _concEliminarAjuste(lado, i) { _concAjustes[lado].splice(i,1); renderAjustesConc(lado); recalcularConciliacionV14(); }

function recalcularConciliacionV14() {
  var saldoBancoBase = parseFloat((document.getElementById('conc-saldo-banco')||{}).value) || 0;
  var saldoLibrosBase = parseFloat((document.getElementById('conc-saldo-libros')||{}).value) || 0;

  var ajusteBanco = _concAjustes.banco.reduce(function(s,a){return s+a.monto;}, 0);
  var ajusteLibros = _concAjustes.libros.reduce(function(s,a){return s+a.monto;}, 0);

  var saldoAjustadoBanco = saldoBancoBase + ajusteBanco;
  var saldoAjustadoLibros = saldoLibrosBase + ajusteLibros;
  var diff = saldoAjustadoBanco - saldoAjustadoLibros;
  var conciliado = Math.abs(diff) < 0.01;

  var kpisEl = document.getElementById('concKpis');
  if (kpisEl) {
    kpisEl.innerHTML =
      '<div class="v14-kpi"><div class="v14-kpi-label">Saldo Ajustado Banco</div><div class="v14-kpi-value">'+fmtCRC(saldoAjustadoBanco)+'</div></div>'+
      '<div class="v14-kpi"><div class="v14-kpi-label">Saldo Ajustado Libros</div><div class="v14-kpi-value">'+fmtCRC(saldoAjustadoLibros)+'</div></div>'+
      '<div class="v14-kpi '+(conciliado?'ok':'bad')+'"><div class="v14-kpi-label">'+(conciliado?'Conciliado':'Diferencia')+'</div><div class="v14-kpi-value">'+(conciliado?'✓ OK':fmtCRC(Math.abs(diff)))+'</div></div>';
  }
  window._concResultado = { saldoAjustadoBanco: saldoAjustadoBanco, saldoAjustadoLibros: saldoAjustadoLibros, diff: diff, conciliado: conciliado };
}

function exportarConciliacionCSV() {
  var r = window._concResultado || {};
  var csv = 'Seccion,Concepto,Monto\n';
  csv += '"Banco","Saldo según estado de cuenta",'+(parseFloat((document.getElementById('conc-saldo-banco')||{}).value)||0)+'\n';
  _concAjustes.banco.forEach(function(a){ csv += '"Banco","'+a.concepto+'",'+a.monto+'\n'; });
  csv += '"Libros","Saldo según libros",'+(parseFloat((document.getElementById('conc-saldo-libros')||{}).value)||0)+'\n';
  _concAjustes.libros.forEach(function(a){ csv += '"Libros","'+a.concepto+'",'+a.monto+'\n'; });
  csv += '"Resultado","Saldo ajustado banco",'+(r.saldoAjustadoBanco||0)+'\n';
  csv += '"Resultado","Saldo ajustado libros",'+(r.saldoAjustadoLibros||0)+'\n';
  csv += '"Resultado","Diferencia",'+(r.diff||0)+'\n';
  descargarBlob('\uFEFF'+csv, 'ConciliacionBancaria_'+new Date().toISOString().split('T')[0]+'.csv', 'text/csv;charset=utf-8;');
}

function exportarConciliacionPDF() {
  var doc = nuevoPDF('portrait'); if (!doc) return;
  pdfHeader(doc, 'Conciliación Bancaria');
  var y = 24;
  doc.setFontSize(11); doc.setFont('helvetica','bold'); doc.setTextColor(45,122,12);
  doc.text('Según Banco', 14, y); y += 6;
  doc.setFontSize(9); doc.setFont('helvetica','normal'); doc.setTextColor(0,0,0);
  doc.text('Saldo según estado de cuenta', 14, y);
  doc.text(fmtCRC(parseFloat((document.getElementById('conc-saldo-banco')||{}).value)||0), 196, y, {align:'right'}); y += 5;
  _concAjustes.banco.forEach(function(a){ doc.text(a.concepto, 14, y); doc.text(fmtCRC(a.monto), 196, y, {align:'right'}); y += 5; });
  y += 4;
  doc.setFontSize(11); doc.setFont('helvetica','bold'); doc.setTextColor(45,122,12);
  doc.text('Según Libros', 14, y); y += 6;
  doc.setFontSize(9); doc.setFont('helvetica','normal'); doc.setTextColor(0,0,0);
  doc.text('Saldo según libros', 14, y);
  doc.text(fmtCRC(parseFloat((document.getElementById('conc-saldo-libros')||{}).value)||0), 196, y, {align:'right'}); y += 5;
  _concAjustes.libros.forEach(function(a){ doc.text(a.concepto, 14, y); doc.text(fmtCRC(a.monto), 196, y, {align:'right'}); y += 5; });
  y += 6;
  var r = window._concResultado || {};
  doc.setFont('helvetica','bold');
  doc.text('Diferencia final', 14, y);
  doc.setTextColor(r.conciliado?45:220, r.conciliado?122:38, r.conciliado?12:38);
  doc.text(r.conciliado?'✓ Conciliado':fmtCRC(Math.abs(r.diff||0)), 196, y, {align:'right'});
  doc.save('ConciliacionBancaria_'+new Date().toISOString().split('T')[0]+'.pdf');
}

// ════════════════════════════════════════════════════
// 5) DECLARACIÓN D-101 — IMPUESTO SOBRE LA RENTA (CR, período fiscal 2026)
// Fuente: Decreto Ejecutivo N° 45333-H (La Gaceta N° 229, 05/12/2025)
// Umbral pyme por renta bruta: ₡119,174,000 anuales (2026)
// ════════════════════════════════════════════════════
var _d101Ingresos = [{concepto:'Ventas / servicios prestados', monto:0}];
var _d101Gastos = [{concepto:'Salarios y cargas sociales', monto:0},{concepto:'Alquiler', monto:0},{concepto:'Servicios públicos', monto:0}];

// Tramos persona física con actividad lucrativa (anual, renta neta) — 2026
var TRAMOS_PF_2026 = [
  { hasta: 4181000,   tasa: 0 },
  { hasta: 6244000,   tasa: 0.10 },
  { hasta: 10412000,  tasa: 0.15 },
  { hasta: 20865000,  tasa: 0.20 },
  { hasta: Infinity,  tasa: 0.25 }
];
// Escala reducida pyme (renta bruta ≤ ₡119,174,000) — tramos referenciales 2026, tasas confirmadas 5/10/15/20%
var TRAMOS_PYME_2026 = [
  { hasta: 5644000,   tasa: 0 },
  { hasta: 8436000,   tasa: 0.05 },
  { hasta: 14066000,  tasa: 0.10 },
  { hasta: 21185000,  tasa: 0.15 },
  { hasta: Infinity,  tasa: 0.20 }
];

function abrirD101() {
  openModal('modalD101');
  renderD101Filas('ingresos');
  renderD101Filas('gastos');
  switchD101Tab('datos');
}

function switchD101Tab(tab) {
  ['datos','ingresos','gastos','calculo'].forEach(function(t){
    var tabEl = document.getElementById('d101Tab-'+t);
    var paneEl = document.getElementById('d101Pane-'+t);
    if (tabEl) tabEl.classList.toggle('active', t===tab);
    if (paneEl) paneEl.classList.toggle('active', t===tab);
  });
  if (tab === 'calculo') calcularD101();
}

function renderD101Filas(tipo) {
  var arr = tipo === 'ingresos' ? _d101Ingresos : _d101Gastos;
  var body = document.getElementById(tipo === 'ingresos' ? 'd101IngresosBody' : 'd101GastosBody');
  if (!body) return;
  var html = '<table class="tbl-v14"><thead><tr><th>Concepto</th><th>Monto (₡)</th><th></th></tr></thead><tbody>';
  arr.forEach(function(item, i){
    html += '<tr><td><input class="inp" style="font-size:12px;padding:4px 7px;" value="'+(item.concepto||'').replace(/"/g,'&quot;')+'" oninput="window._d101SetConcepto(\''+tipo+'\','+i+',this.value)"></td>'+
      '<td><input class="inp" type="number" style="text-align:right;" value="'+item.monto+'" oninput="window._d101SetMonto(\''+tipo+'\','+i+',this.value)"></td>'+
      '<td><button onclick="window._d101EliminarFila(\''+tipo+'\','+i+')" style="background:none;border:none;color:var(--danger);font-size:15px;cursor:pointer;">×</button></td></tr>';
  });
  var total = arr.reduce(function(s,i){return s+i.monto;},0);
  html += '<tr class="total-row"><td>TOTAL</td><td class="num">'+fmtCRC(total)+'</td><td></td></tr>';
  html += '</tbody></table>';
  body.innerHTML = html;
}

function agregarFilaD101(tipo) {
  (tipo === 'ingresos' ? _d101Ingresos : _d101Gastos).push({ concepto:'Nuevo concepto', monto:0 });
  renderD101Filas(tipo);
}
function _d101SetConcepto(tipo, i, val) { var arr = tipo==='ingresos'?_d101Ingresos:_d101Gastos; if (arr[i]) arr[i].concepto = val; }
function _d101SetMonto(tipo, i, val) { var arr = tipo==='ingresos'?_d101Ingresos:_d101Gastos; if (arr[i]) { arr[i].monto = parseFloat(val)||0; renderD101Filas(tipo); } }
function _d101EliminarFila(tipo, i) { var arr = tipo==='ingresos'?_d101Ingresos:_d101Gastos; arr.splice(i,1); renderD101Filas(tipo); }

function calcularImpuestoPorTramos(rentaNeta, tramos) {
  var impuesto = 0, anterior = 0;
  for (var i = 0; i < tramos.length; i++) {
    var tramo = tramos[i];
    if (rentaNeta <= anterior) break;
    var enEsteTramo = Math.min(rentaNeta, tramo.hasta) - anterior;
    if (enEsteTramo > 0) impuesto += enEsteTramo * tramo.tasa;
    anterior = tramo.hasta;
  }
  return impuesto;
}

function calcularD101() {
  var totalIngresos = _d101Ingresos.reduce(function(s,i){return s+i.monto;},0);
  var totalGastos = _d101Gastos.reduce(function(s,i){return s+i.monto;},0);
  var rentaNeta = Math.max(0, totalIngresos - totalGastos);
  var tipo = (document.getElementById('d101-tipo')||{}).value || 'persona_fisica';

  var impuesto = 0, tasaInfo = '';
  if (tipo === 'persona_fisica') {
    impuesto = calcularImpuestoPorTramos(rentaNeta, TRAMOS_PF_2026);
    tasaInfo = 'Tarifa progresiva 0% / 10% / 15% / 20% / 25% sobre renta neta anual (persona física con actividad lucrativa, 2026).';
  } else if (tipo === 'empresa_micro' || tipo === 'empresa_pequena') {
    impuesto = calcularImpuestoPorTramos(rentaNeta, TRAMOS_PYME_2026);
    tasaInfo = 'Escala reducida 0% / 5% / 10% / 15% / 20% — aplica solo si la renta bruta anual no supera ₡119,174,000 (umbral 2026).';
  } else {
    impuesto = rentaNeta * 0.30;
    tasaInfo = 'Tarifa plana del 30% sobre renta neta imponible (renta bruta superior a ₡119,174,000).';
  }

  var body = document.getElementById('d101CalculoBody');
  if (body) {
    body.innerHTML =
      '<div class="v14-kpis">' +
        '<div class="v14-kpi"><div class="v14-kpi-label">Total Ingresos</div><div class="v14-kpi-value">'+fmtCRC(totalIngresos)+'</div></div>' +
        '<div class="v14-kpi"><div class="v14-kpi-label">Total Gastos Deducibles</div><div class="v14-kpi-value">'+fmtCRC(totalGastos)+'</div></div>' +
        '<div class="v14-kpi"><div class="v14-kpi-label">Renta Neta</div><div class="v14-kpi-value">'+fmtCRC(rentaNeta)+'</div></div>' +
        '<div class="v14-kpi ok"><div class="v14-kpi-label">Impuesto Estimado</div><div class="v14-kpi-value">'+fmtCRC(impuesto)+'</div></div>' +
      '</div>' +
      '<div class="v14-d101-sec"><p style="font-size:11.5px;color:var(--text2);line-height:1.5;">'+tasaInfo+'</p>'+
      '<p style="font-size:11px;color:var(--text3);margin-top:6px;">⚠️ Los tramos de la escala reducida (pyme) son referenciales para el período fiscal 2026; verificá los montos exactos vigentes en la tabla oficial de Hacienda antes de presentar la declaración. El umbral de ₡119,174,000 de renta bruta anual y la tasa plana del 30% sí están confirmados por el Decreto 45333-H.</p></div>';
  }
  window._d101Resultado = { totalIngresos: totalIngresos, totalGastos: totalGastos, rentaNeta: rentaNeta, impuesto: impuesto, tipo: tipo };
}

function exportarD101CSV() {
  var r = window._d101Resultado || calcularD101() || window._d101Resultado || {};
  var csv = 'Seccion,Concepto,Monto\n';
  _d101Ingresos.forEach(function(i){ csv += '"Ingresos","'+i.concepto+'",'+i.monto+'\n'; });
  _d101Gastos.forEach(function(i){ csv += '"Gastos","'+i.concepto+'",'+i.monto+'\n'; });
  csv += '"Resultado","Renta Neta",'+(r.rentaNeta||0)+'\n';
  csv += '"Resultado","Impuesto Estimado",'+(r.impuesto||0)+'\n';
  descargarBlob('\uFEFF'+csv, 'D101_'+new Date().toISOString().split('T')[0]+'.csv', 'text/csv;charset=utf-8;');
}

function exportarD101PDF() {
  calcularD101();
  var r = window._d101Resultado || {};
  var doc = nuevoPDF('portrait'); if (!doc) return;
  pdfHeader(doc, 'Declaración D-101 — Impuesto sobre la Renta');
  var y = 24;
  doc.setFontSize(10); doc.setFont('helvetica','normal');
  doc.text('Contribuyente: ' + ((document.getElementById('d101-nombre')||{}).value || '—'), 14, y); y += 6;
  doc.text('Cédula: ' + ((document.getElementById('d101-cedula')||{}).value || '—'), 14, y); y += 6;
  doc.text('Período fiscal: ' + ((document.getElementById('d101-periodo')||{}).value || '—'), 14, y); y += 8;

  doc.setFont('helvetica','bold'); doc.setTextColor(45,122,12);
  doc.text('Ingresos', 14, y); y += 6;
  doc.setFont('helvetica','normal'); doc.setTextColor(0,0,0);
  _d101Ingresos.forEach(function(i){ doc.text(i.concepto, 14, y); doc.text(fmtCRC(i.monto), 196, y, {align:'right'}); y += 5; });
  y += 3;
  doc.setFont('helvetica','bold'); doc.setTextColor(45,122,12);
  doc.text('Gastos Deducibles', 14, y); y += 6;
  doc.setFont('helvetica','normal'); doc.setTextColor(0,0,0);
  _d101Gastos.forEach(function(i){ doc.text(i.concepto, 14, y); doc.text(fmtCRC(i.monto), 196, y, {align:'right'}); y += 5; });
  y += 6;
  doc.setFont('helvetica','bold');
  doc.text('Renta Neta', 14, y); doc.text(fmtCRC(r.rentaNeta||0), 196, y, {align:'right'}); y += 6;
  doc.setTextColor(45,122,12);
  doc.text('Impuesto Estimado', 14, y); doc.text(fmtCRC(r.impuesto||0), 196, y, {align:'right'}); y += 8;
  doc.setFontSize(8); doc.setFont('helvetica','normal'); doc.setTextColor(120,120,120);
  doc.text('Montos referenciales — verificar tramos exactos vigentes en Hacienda antes de presentar.', 14, y);
  doc.save('D101_'+new Date().toISOString().split('T')[0]+'.pdf');
}

// ════════════════════════════════════════════════════
// EXPOSICIÓN GLOBAL DE FUNCIONES (requerido por onclick= en HTML)
// ════════════════════════════════════════════════════
window.leerHojaActiva = leerHojaActiva;
window.fmtCRC = fmtCRC;

window.abrirLibroMayor = abrirLibroMayor;
window.generarLibroMayor = generarLibroMayor;
window.exportarLibroMayorCSV = exportarLibroMayorCSV;
window.exportarLibroMayorPDF = exportarLibroMayorPDF;

window.abrirBalanceComp = abrirBalanceComp;
window.generarBalanceComp = generarBalanceComp;
window.exportarBalanceCompCSV = exportarBalanceCompCSV;
window.exportarBalanceCompPDF = exportarBalanceCompPDF;

window.abrirFlujoEfectivo = abrirFlujoEfectivo;
window.switchFlujoTab = switchFlujoTab;
window.agregarFilaFlujoDirecto = agregarFilaFlujoDirecto;
window._feSetConcepto = _feSetConcepto;
window._feSetTipo = _feSetTipo;
window._feSetMonto = _feSetMonto;
window._feEliminarFila = _feEliminarFila;
window.recalcularFlujoEfectivo = recalcularFlujoEfectivo;
window.exportarFlujoCSV = exportarFlujoCSV;
window.exportarFlujoPDF = exportarFlujoPDF;

window.abrirConciliacionV14 = abrirConciliacionV14;
window.agregarAjusteConc = agregarAjusteConc;
window._concSetConcepto = _concSetConcepto;
window._concSetMonto = _concSetMonto;
window._concEliminarAjuste = _concEliminarAjuste;
window.recalcularConciliacionV14 = recalcularConciliacionV14;
window.exportarConciliacionCSV = exportarConciliacionCSV;
window.exportarConciliacionPDF = exportarConciliacionPDF;

window.abrirD101 = abrirD101;
window.switchD101Tab = switchD101Tab;
window.agregarFilaD101 = agregarFilaD101;
window._d101SetConcepto = _d101SetConcepto;
window._d101SetMonto = _d101SetMonto;
window._d101EliminarFila = _d101EliminarFila;
window.calcularD101 = calcularD101;
window.exportarD101CSV = exportarD101CSV;
window.exportarD101PDF = exportarD101PDF;

// ════════════════════════════════════════════════════
// AUTO-INYECCIÓN: botones en sidebar + items en menú Plantillas
// ════════════════════════════════════════════════════
function inyectarUIv145() {
  // Sección en sidebar
  var sidebar = document.getElementById('sidebar');
  if (sidebar && !document.getElementById('v14-sidebar-sec')) {
    var sec = document.createElement('div');
    sec.id = 'v14-sidebar-sec';
    sec.innerHTML =
      '<div class="sb-div"></div>' +
      '<div class="sb-sec">' +
        '<div class="sb-title">v14 — Nuevas herramientas</div>' +
        '<button class="sb-btn" onclick="abrirLibroMayor()"><span class="si">📗</span><span>Libro Mayor Automático</span></button>' +
        '<button class="sb-btn" onclick="abrirBalanceComp()"><span class="si">📊</span><span>Balance Comprobación (en vivo)</span></button>' +
        '<button class="sb-btn" onclick="abrirFlujoEfectivo()"><span class="si">💧</span><span>Flujo de Efectivo</span></button>' +
        '<button class="sb-btn" onclick="abrirConciliacionV14()"><span class="si">🏦</span><span>Conciliación Bancaria</span></button>' +
        '<button class="sb-btn" onclick="abrirD101()"><span class="si">🇨🇷</span><span>Declaración D-101 (Renta)</span></button>' +
      '</div>';
    // Insertar antes del widget de Sibö si existe, sino al final
    var siboWidget = sidebar.querySelector('.sibo-widget');
    if (siboWidget) sidebar.insertBefore(sec, siboWidget);
    else sidebar.appendChild(sec);
  }

  // Items en el menú dropdown de Plantillas
  var menuPlantillas = document.getElementById('m-plantillas');
  if (menuPlantillas && !document.getElementById('v14-dd-sep')) {
    var sep = document.createElement('div');
    sep.className = 'dd-sep'; sep.id = 'v14-dd-sep';
    var title = document.createElement('div');
    title.className = 'dd-section-title'; title.textContent = 'Herramientas v14';
    var items = [
      ['📗','Libro Mayor Automático','abrirLibroMayor()'],
      ['📊','Balance Comprobación','abrirBalanceComp()'],
      ['💧','Flujo de Efectivo','abrirFlujoEfectivo()'],
      ['🏦','Conciliación Bancaria','abrirConciliacionV14()'],
      ['🇨🇷','Declaración D-101','abrirD101()']
    ];
    menuPlantillas.appendChild(sep);
    menuPlantillas.appendChild(title);
    items.forEach(function(it){
      var div = document.createElement('div');
      div.className = 'dd-item';
      div.setAttribute('onclick', "closeMenu();" + it[2]);
      div.innerHTML = '<span class="di">'+it[0]+'</span>'+it[1];
      menuPlantillas.appendChild(div);
    });
  }
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', inyectarUIv145);
} else {
  inyectarUIv145();
}

})();




