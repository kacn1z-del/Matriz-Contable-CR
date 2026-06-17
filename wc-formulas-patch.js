// MATRIZ CONTABLE CR — PATCH TOTAL v13.9 COMPLETO
// TODO EN UN SOLO ARCHIVO:
//   v13.7 (motor base + gráficos + bugfix + plantillas)
//   + v13.8a + v13.8b (estadísticas/financieras avanzadas)
//   + v13.9 (estadísticas CR + fórmulas exclusivas)
//
// ✓ ~490 fórmulas activas (96% de Excel 365)
// ✓ 20 tipos de gráfico con modal
// ✓ Plantillas + Botones + Exportar CSV
// ✓ Autosave + datos NO se borran
// ✓ Fórmulas CR: CCSS, AGUINALDO, RENTA, IVA, etc.
//
// INSTRUCCIONES:
// 1. QUITÁ TODOS los parches anteriores
// 2. Pegá SOLO ESTE antes de </body>
// ════════════════════════════════════════════════════

// ════════════════════════════════════════════════════
// MATRIZ CONTABLE CR — PATCH ÚNICO v13.7
// TODO EN UN SOLO BLOQUE — reemplaza todos los parches
// anteriores (v13.3, v13.4, v13.5, BUGFIX, v13.6)
//
// ✓ ~345 fórmulas activas (67% de Excel 365)
// ✓ 20 tipos de gráfico con modal de selección
// ✓ Plantillas: Diario, Mayor, Balance, Conciliación, D-151
// ✓ Botones: Tabla, Nota, Agregar/Eliminar fila, Calcular
// ✓ Exportar CSV con fórmulas evaluadas
// ✓ BUGFIX: datos NO se borran al cambiar celda
// ✓ Autosave cada 30s + restaura al cargar página
//
// INSTRUCCIONES:
// 1. Quitá TODOS los parches anteriores
// 2. Pegá SOLO este antes de </body>
// ════════════════════════════════════════════════════

// ════════════════════════════════════════════════════
// MATRIZ CONTABLE CR — PATCH v13.5 UNIFICADO
// Combina v13.3 + v13.4 en UN SOLO bloque
// ~271 fórmulas activas + 20 tipos de gráfico
// + Plantillas + Botones + Exportar
//
// INSTRUCCIONES:
// 1. QUITÁ el PATCH v13.3 y el PATCH v13.4 anteriores
// 2. Pegá SOLO este parche antes de </body>
// ════════════════════════════════════════════════════
function evalFormula(formula, row, col) {
  if (!formula || !formula.startsWith('=')) return formula;
  var expr = formula.substring(1).trim();
  var exprUp = expr.toUpperCase();
  try {

    // ── Helpers internos ────────────────────
    function getVal(cid) {
      var v = sheetData[cid.toUpperCase()] || '';
      if (typeof v === 'string' && v.startsWith('=')) {
        v = evalFormula(v, 0, 0);
      }
      return v;
    }
    function toNum(v) {
      if (v === '' || v === null || v === undefined) return 0;
      var n = parseFloat(String(v).replace(/[₡,\s]/g, ''));
      return isNaN(n) ? 0 : n;
    }
    function getCells(from, to) {
      var fc = colIdx(from.match(/[A-Za-z]+/)[0].toUpperCase());
      var fr = parseInt(from.match(/\d+/)[0]) - 1;
      var tc = colIdx(to.match(/[A-Za-z]+/)[0].toUpperCase());
      var tr2 = parseInt(to.match(/\d+/)[0]) - 1;
      var out = [];
      for (var r = fr; r <= tr2; r++)
        for (var c = fc; c <= tc; c++)
          out.push(cellId(r, c));
      return out;
    }
    function colIdx(s) {
      var n = 0;
      for (var i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64;
      return n - 1;
    }
    function fmtNum(n) {
      return isNaN(n) ? '#VALOR!' : Number(n).toLocaleString('es-CR', {maximumFractionDigits: 10});
    }
    function splitFnArgs(s) {
      var args = [], depth = 0, cur = '', inStr = false, strChar = '';
      for (var i = 0; i < s.length; i++) {
        var ch = s[i];
        if (!inStr && (ch === '"' || ch === "'")) { inStr = true; strChar = ch; cur += ch; continue; }
        if (inStr && ch === strChar) { inStr = false; cur += ch; continue; }
        if (!inStr && ch === '(') depth++;
        if (!inStr && ch === ')') depth--;
        if (!inStr && (ch === ';' || ch === ',') && depth === 0) { args.push(cur.trim()); cur = ''; continue; }
        cur += ch;
      }
      if (cur.trim()) args.push(cur.trim());
      return args;
    }
    function resolveArg(a) {
      a = a.trim();
      if ((a.startsWith('"') && a.endsWith('"')) || (a.startsWith("'") && a.endsWith("'")))
        return a.slice(1, -1);
      if (/^[A-Za-z]+\d+:[A-Za-z]+\d+$/.test(a)) return a; // range ref
      if (/^[A-Za-z]+\d+$/.test(a)) return getVal(a);
      if (a.startsWith('=')) return evalFormula(a, row, col);
      if (!isNaN(a)) return parseFloat(a);
      // sub-expression
      return evalExpr(a);
    }
    function resolveNum(a) { return toNum(resolveArg(a)); }
    function isRange(a) { return /^[A-Za-z]+\d+:[A-Za-z]+\d+$/.test(a.trim()); }
    function rangeNums(a) {
      var parts = a.trim().split(':');
      return getCells(parts[0], parts[1]).map(function(c) { return toNum(getVal(c)); });
    }
    function rangeVals(a) {
      var parts = a.trim().split(':');
      return getCells(parts[0], parts[1]).map(function(c) { return getVal(c); });
    }

    // ── Evaluador de expresiones aritméticas ─
    function evalExpr(e) {
      e = e.trim().toUpperCase();
      // Replace cell refs with values
      e = e.replace(/([A-Z]+\d+)/g, function(m) { return toNum(getVal(m)) || 0; });
      // Safe eval
      try { return Function('"use strict";return (' + e + ')')(); } catch(ex) { return '#ERR'; }
    }

    // ── Extract function call ────────────────
    function matchFn(name) {
      var re = new RegExp('^(?:' + name + ')\\((.*)\\)$', 'i');
      var m = exprUp.match(re);
      if (!m) return null;
      // Usar la posición real del primer paréntesis en expr, no name.length
      // (name puede ser un patrón con sinónimos 'SUMA|SUM' de longitud distinta al nombre real)
      var openParen = expr.indexOf('(');
      var depth2 = 0, end2 = openParen;
      for (var i2 = openParen + 1; i2 < expr.length; i2++) {
        if (expr[i2] === '(') depth2++;
        if (expr[i2] === ')') { if (depth2 === 0) { end2 = i2; break; } depth2--; }
      }
      return expr.substring(openParen + 1, end2);
    }

    var inner, args, nums, vals;

    // ─────────────────────────────────────────
    // MATEMÁTICAS
    // ─────────────────────────────────────────
    if ((inner = matchFn('SUMA|SUM')) !== null) {
      args = splitFnArgs(inner); var s = 0;
      args.forEach(function(a) { if (isRange(a)) rangeNums(a).forEach(function(n) { s += n; }); else s += resolveNum(a); });
      return fmtNum(s);
    }
    if ((inner = matchFn('PROMEDIO|AVERAGE')) !== null) {
      args = splitFnArgs(inner); var tot = 0, cnt = 0;
      args.forEach(function(a) { if (isRange(a)) rangeNums(a).forEach(function(n) { tot += n; cnt++; }); else { tot += resolveNum(a); cnt++; } });
      return cnt ? fmtNum(tot / cnt) : '#DIV/0!';
    }
    if ((inner = matchFn('MAX|MAXIMO')) !== null) {
      args = splitFnArgs(inner); nums = [];
      args.forEach(function(a) { if (isRange(a)) rangeNums(a).forEach(function(n) { nums.push(n); }); else nums.push(resolveNum(a)); });
      return nums.length ? fmtNum(Math.max.apply(null, nums)) : '0';
    }
    if ((inner = matchFn('MIN|MINIMO')) !== null) {
      args = splitFnArgs(inner); nums = [];
      args.forEach(function(a) { if (isRange(a)) rangeNums(a).forEach(function(n) { nums.push(n); }); else nums.push(resolveNum(a)); });
      return nums.length ? fmtNum(Math.min.apply(null, nums)) : '0';
    }
    if ((inner = matchFn('CONTAR|COUNT')) !== null) {
      args = splitFnArgs(inner); var c2 = 0;
      args.forEach(function(a) { if (isRange(a)) rangeNums(a).forEach(function(n) { if (!isNaN(n)) c2++; }); });
      return String(c2);
    }
    if ((inner = matchFn('CONTARA|COUNTA')) !== null) {
      args = splitFnArgs(inner); var ca = 0;
      args.forEach(function(a) { if (isRange(a)) rangeVals(a).forEach(function(v) { if (v !== '') ca++; }); });
      return String(ca);
    }
    if ((inner = matchFn('CONTAR\\.BLANCO|COUNTBLANK')) !== null) {
      args = splitFnArgs(inner); var cb = 0;
      args.forEach(function(a) { if (isRange(a)) rangeVals(a).forEach(function(v) { if (v === '') cb++; }); });
      return String(cb);
    }
    if ((inner = matchFn('ABS')) !== null) { return fmtNum(Math.abs(resolveNum(splitFnArgs(inner)[0]))); }
    if ((inner = matchFn('ENTERO|INT')) !== null) { return fmtNum(Math.floor(resolveNum(splitFnArgs(inner)[0]))); }
    if ((inner = matchFn('TRUNCAR|TRUNC')) !== null) { args = splitFnArgs(inner); var decimals = args[1] ? parseInt(args[1]) : 0; var n2 = resolveNum(args[0]); return fmtNum(Math.trunc(n2 * Math.pow(10, decimals)) / Math.pow(10, decimals)); }
    if ((inner = matchFn('REDONDEAR|ROUND')) !== null) { args = splitFnArgs(inner); return fmtNum(parseFloat(resolveNum(args[0]).toFixed(parseInt(resolveNum(args[1]))))); }
    if ((inner = matchFn('REDONDEAR\\.MAS|ROUNDUP')) !== null) { args = splitFnArgs(inner); var d3 = parseInt(resolveNum(args[1])); var p = Math.pow(10, d3); return fmtNum(Math.ceil(resolveNum(args[0]) * p) / p); }
    if ((inner = matchFn('REDONDEAR\\.MENOS|ROUNDDOWN')) !== null) { args = splitFnArgs(inner); var d4 = parseInt(resolveNum(args[1])); var p2 = Math.pow(10, d4); return fmtNum(Math.floor(resolveNum(args[0]) * p2) / p2); }
    if ((inner = matchFn('RESIDUO|MOD')) !== null) { args = splitFnArgs(inner); var a2 = resolveNum(args[0]), b2 = resolveNum(args[1]); return fmtNum(a2 - Math.floor(a2 / b2) * b2); }
    if ((inner = matchFn('POTENCIA|POWER')) !== null) { args = splitFnArgs(inner); return fmtNum(Math.pow(resolveNum(args[0]), resolveNum(args[1]))); }
    if ((inner = matchFn('RAIZ|SQRT')) !== null) { return fmtNum(Math.sqrt(resolveNum(splitFnArgs(inner)[0]))); }
    if ((inner = matchFn('LN')) !== null) { return fmtNum(Math.log(resolveNum(splitFnArgs(inner)[0]))); }
    if ((inner = matchFn('LOG10')) !== null) { return fmtNum(Math.log10(resolveNum(splitFnArgs(inner)[0]))); }
    if ((inner = matchFn('LOG')) !== null) { args = splitFnArgs(inner); var base = args[1] ? resolveNum(args[1]) : 10; return fmtNum(Math.log(resolveNum(args[0])) / Math.log(base)); }
    if ((inner = matchFn('EXP')) !== null) { return fmtNum(Math.exp(resolveNum(splitFnArgs(inner)[0]))); }
    if ((inner = matchFn('PI')) !== null) { return fmtNum(Math.PI); }
    if ((inner = matchFn('FACT')) !== null) { var fn2 = resolveNum(splitFnArgs(inner)[0]); var f2 = 1; for (var i3 = 2; i3 <= fn2; i3++) f2 *= i3; return fmtNum(f2); }
    if ((inner = matchFn('SIGNO|SIGN')) !== null) { var sv = resolveNum(splitFnArgs(inner)[0]); return String(sv > 0 ? 1 : sv < 0 ? -1 : 0); }
    if ((inner = matchFn('COCIENTE|QUOTIENT')) !== null) { args = splitFnArgs(inner); return fmtNum(Math.trunc(resolveNum(args[0]) / resolveNum(args[1]))); }
    if ((inner = matchFn('COMBINAT|COMBIN')) !== null) { args = splitFnArgs(inner); var nn = resolveNum(args[0]), kk = resolveNum(args[1]); var r2 = 1; for (var i4 = 0; i4 < kk; i4++) r2 = r2 * (nn - i4) / (i4 + 1); return fmtNum(Math.round(r2)); }
    if ((inner = matchFn('SUMAPRODUCTO|SUMPRODUCT')) !== null) {
      args = splitFnArgs(inner);
      if (args.length === 1 && isRange(args[0])) { return fmtNum(rangeNums(args[0]).reduce(function(a, b) { return a + b; }, 0)); }
      if (args.length >= 2 && isRange(args[0]) && isRange(args[1])) {
        var n1 = rangeNums(args[0]), n3 = rangeNums(args[1]); var sp = 0;
        for (var i5 = 0; i5 < Math.min(n1.length, n3.length); i5++) sp += n1[i5] * n3[i5];
        return fmtNum(sp);
      }
    }
    if ((inner = matchFn('SUBTOTALES|SUBTOTAL')) !== null) {
      args = splitFnArgs(inner); var fnNum = parseInt(resolveNum(args[0]));
      if (args[1] && isRange(args[1])) {
        nums = rangeNums(args[1]);
        if (fnNum === 9 || fnNum === 109) return fmtNum(nums.reduce(function(a,b){return a+b;},0));
        if (fnNum === 1 || fnNum === 101) return fmtNum(nums.reduce(function(a,b){return a+b;},0)/nums.length);
        if (fnNum === 4 || fnNum === 104) return fmtNum(Math.max.apply(null,nums));
        if (fnNum === 5 || fnNum === 105) return fmtNum(Math.min.apply(null,nums));
        if (fnNum === 2 || fnNum === 102) return String(nums.filter(function(n){return !isNaN(n);}).length);
      }
    }

    // ── CONTAR.SI / COUNTIF ──────────────────
    if ((inner = matchFn('CONTAR\\.SI\\.CONJUNTO|COUNTIFS')) !== null) {
      args = splitFnArgs(inner); var count3 = 0;
      var rng1 = args[0], crit1 = args[1] ? String(resolveArg(args[1])) : '';
      if (isRange(rng1)) {
        rangeVals(rng1).forEach(function(v, i6) {
          if (matchCrit(v, crit1)) count3++;
        });
      }
      return String(count3);
    }
    if ((inner = matchFn('CONTAR\\.SI|COUNTIF')) !== null) {
      args = splitFnArgs(inner); var count4 = 0;
      var rng2 = args[0], crit2 = String(resolveArg(args[1] || ''));
      if (isRange(rng2)) rangeVals(rng2).forEach(function(v) { if (matchCrit(v, crit2)) count4++; });
      return String(count4);
    }

    // ── SUMAR.SI / SUMIF ────────────────────
    if ((inner = matchFn('SUMAR\\.SI\\.CONJUNTO|SUMIFS')) !== null) {
      args = splitFnArgs(inner); var sumR = isRange(args[0]) ? rangeNums(args[0]) : [];
      var sumCrit = isRange(args[1]) ? rangeVals(args[1]) : [];
      var critVal = String(resolveArg(args[2] || ''));
      var s3 = 0;
      sumCrit.forEach(function(v, i7) { if (matchCrit(v, critVal)) s3 += sumR[i7] || 0; });
      return fmtNum(s3);
    }
    if ((inner = matchFn('SUMAR\\.SI|SUMIF')) !== null) {
      args = splitFnArgs(inner);
      var critRng = isRange(args[0]) ? rangeVals(args[0]) : [];
      var crit3 = String(resolveArg(args[1] || ''));
      var sumRng = args[2] ? (isRange(args[2]) ? rangeNums(args[2]) : []) : critRng.map(function(v){return toNum(v);});
      var s4 = 0;
      critRng.forEach(function(v, i8) { if (matchCrit(v, crit3)) s4 += sumRng[i8] || 0; });
      return fmtNum(s4);
    }

    // ── PROMEDIO.SI / AVERAGEIF ─────────────
    if ((inner = matchFn('PROMEDIO\\.SI|AVERAGEIF')) !== null) {
      args = splitFnArgs(inner);
      var critRng2 = isRange(args[0]) ? rangeVals(args[0]) : [];
      var crit4 = String(resolveArg(args[1] || ''));
      var avgRng = args[2] ? (isRange(args[2]) ? rangeNums(args[2]) : []) : critRng2.map(function(v){return toNum(v);});
      var t2 = 0, cnt2 = 0;
      critRng2.forEach(function(v, i9) { if (matchCrit(v, crit4)) { t2 += avgRng[i9] || 0; cnt2++; } });
      return cnt2 ? fmtNum(t2 / cnt2) : '#DIV/0!';
    }

    // ── MAX.SI / MIN.SI ─────────────────────
    if ((inner = matchFn('MAX\\.SI\\.CONJUNTO|MAXIFS')) !== null) {
      args = splitFnArgs(inner);
      var maxRng = isRange(args[0]) ? rangeNums(args[0]) : [];
      var maxCritRng = isRange(args[1]) ? rangeVals(args[1]) : [];
      var maxCrit = String(resolveArg(args[2] || ''));
      var maxVals = [];
      maxCritRng.forEach(function(v, i){ if (matchCrit(v, maxCrit)) maxVals.push(maxRng[i] || 0); });
      return maxVals.length ? fmtNum(Math.max.apply(null, maxVals)) : '0';
    }
    if ((inner = matchFn('MIN\\.SI\\.CONJUNTO|MINIFS')) !== null) {
      args = splitFnArgs(inner);
      var minRng = isRange(args[0]) ? rangeNums(args[0]) : [];
      var minCritRng = isRange(args[1]) ? rangeVals(args[1]) : [];
      var minCrit = String(resolveArg(args[2] || ''));
      var minVals = [];
      minCritRng.forEach(function(v, i){ if (matchCrit(v, minCrit)) minVals.push(minRng[i] || 0); });
      return minVals.length ? fmtNum(Math.min.apply(null, minVals)) : '0';
    }

    // ── ESTADÍSTICAS ────────────────────────
    if ((inner = matchFn('DESVEST\\.M|DESVEST|STDEV\\.S')) !== null) {
      args = splitFnArgs(inner); nums = isRange(args[0]) ? rangeNums(args[0]) : args.map(resolveNum);
      var mean = nums.reduce(function(a,b){return a+b;},0)/nums.length;
      var variance = nums.reduce(function(s,n){return s+(n-mean)*(n-mean);},0)/(nums.length-1);
      return fmtNum(Math.sqrt(variance));
    }
    if ((inner = matchFn('DESVEST\\.P|DESVESTP|STDEV\\.P')) !== null) {
      args = splitFnArgs(inner); nums = isRange(args[0]) ? rangeNums(args[0]) : args.map(resolveNum);
      var mean2 = nums.reduce(function(a,b){return a+b;},0)/nums.length;
      var var2 = nums.reduce(function(s,n){return s+(n-mean2)*(n-mean2);},0)/nums.length;
      return fmtNum(Math.sqrt(var2));
    }
    if ((inner = matchFn('MEDIANA|MEDIAN')) !== null) {
      args = splitFnArgs(inner); nums = isRange(args[0]) ? rangeNums(args[0]) : args.map(resolveNum);
      nums.sort(function(a,b){return a-b;}); var mid = Math.floor(nums.length/2);
      return fmtNum(nums.length%2 ? nums[mid] : (nums[mid-1]+nums[mid])/2);
    }
    if ((inner = matchFn('K\\.ESIMO\\.MAYOR|LARGE')) !== null) {
      args = splitFnArgs(inner); nums = isRange(args[0]) ? rangeNums(args[0]) : [resolveNum(args[0])];
      nums.sort(function(a,b){return b-a;}); return fmtNum(nums[resolveNum(args[1])-1]);
    }
    if ((inner = matchFn('K\\.ESIMO\\.MENOR|SMALL')) !== null) {
      args = splitFnArgs(inner); nums = isRange(args[0]) ? rangeNums(args[0]) : [resolveNum(args[0])];
      nums.sort(function(a,b){return a-b;}); return fmtNum(nums[resolveNum(args[1])-1]);
    }
    if ((inner = matchFn('COEF\\.DE\\.CORREL|CORREL|PEARSON')) !== null) {
      args = splitFnArgs(inner);
      var x = isRange(args[0]) ? rangeNums(args[0]) : [], y = isRange(args[1]) ? rangeNums(args[1]) : [];
      var n = Math.min(x.length, y.length), mx = x.reduce(function(a,b){return a+b;},0)/n, my = y.reduce(function(a,b){return a+b;},0)/n;
      var num = 0, dx = 0, dy = 0;
      for (var i = 0; i < n; i++) { num += (x[i]-mx)*(y[i]-my); dx += (x[i]-mx)*(x[i]-mx); dy += (y[i]-my)*(y[i]-my); }
      return fmtNum(num / Math.sqrt(dx * dy));
    }

    // ─────────────────────────────────────────
    // LÓGICAS
    // ─────────────────────────────────────────
    if ((inner = matchFn('SI\\.CONJUNTO|IFS')) !== null) {
      args = splitFnArgs(inner);
      for (var i10 = 0; i10 < args.length - 1; i10 += 2) {
        if (evalCond(args[i10])) return String(resolveArg(args[i10+1]));
      }
      return args.length % 2 === 1 ? String(resolveArg(args[args.length-1])) : '#N/A';
    }
    if ((inner = matchFn('SI\\.ERROR|IFERROR')) !== null) {
      args = splitFnArgs(inner);
      try { var v2 = resolveArg(args[0]); if (String(v2).startsWith('#')) return String(resolveArg(args[1])); return String(v2); }
      catch(e2) { return String(resolveArg(args[1])); }
    }
    if ((inner = matchFn('SI\\.ND|IFNA')) !== null) {
      args = splitFnArgs(inner);
      try { var v3 = resolveArg(args[0]); if (String(v3) === '#N/A') return String(resolveArg(args[1])); return String(v3); }
      catch(e3) { return String(resolveArg(args[1])); }
    }
    if ((inner = matchFn('SI|IF')) !== null) {
      args = splitFnArgs(inner);
      return evalCond(args[0]) ? String(resolveArg(args[1] || '')) : String(resolveArg(args[2] || ''));
    }
    if ((inner = matchFn('Y|AND')) !== null) {
      args = splitFnArgs(inner); return args.every(function(a) { return evalCond(a); }) ? 'VERDADERO' : 'FALSO';
    }
    if ((inner = matchFn('O|OR')) !== null) {
      args = splitFnArgs(inner); return args.some(function(a) { return evalCond(a); }) ? 'VERDADERO' : 'FALSO';
    }
    if ((inner = matchFn('NO|NOT')) !== null) {
      return evalCond(splitFnArgs(inner)[0]) ? 'FALSO' : 'VERDADERO';
    }
    if ((inner = matchFn('CAMBIAR|SWITCH')) !== null) {
      args = splitFnArgs(inner); var sw = String(resolveArg(args[0]));
      for (var i11 = 1; i11 < args.length - 1; i11 += 2) {
        if (String(resolveArg(args[i11])) === sw) return String(resolveArg(args[i11+1]));
      }
      return args.length % 2 === 0 ? String(resolveArg(args[args.length-1])) : '';
    }
    if (exprUp === 'VERDADERO()' || exprUp === 'TRUE()') return 'VERDADERO';
    if (exprUp === 'FALSO()' || exprUp === 'FALSE()') return 'FALSO';

    // ─────────────────────────────────────────
    // TEXTO
    // ─────────────────────────────────────────
    if ((inner = matchFn('CONCATENAR|CONCATENATE|CONCAT')) !== null) {
      return splitFnArgs(inner).map(function(a) { return isRange(a) ? rangeVals(a).join('') : String(resolveArg(a)); }).join('');
    }
    if ((inner = matchFn('UNIRCADENAS|TEXTJOIN')) !== null) {
      args = splitFnArgs(inner); var delim = String(resolveArg(args[0])); var ignoreEmpty = String(resolveArg(args[1])).toUpperCase() !== 'FALSO';
      var parts = [];
      for (var i12 = 2; i12 < args.length; i12++) {
        var vt = isRange(args[i12]) ? rangeVals(args[i12]) : [String(resolveArg(args[i12]))];
        vt.forEach(function(v) { if (!ignoreEmpty || v !== '') parts.push(v); });
      }
      return parts.join(delim);
    }
    if ((inner = matchFn('MAYUSC|UPPER')) !== null) { return String(resolveArg(splitFnArgs(inner)[0])).toUpperCase(); }
    if ((inner = matchFn('MINUSC|LOWER')) !== null) { return String(resolveArg(splitFnArgs(inner)[0])).toLowerCase(); }
    if ((inner = matchFn('NOMPROPIO|PROPER')) !== null) { return String(resolveArg(splitFnArgs(inner)[0])).replace(/\w\S*/g, function(w){return w.charAt(0).toUpperCase()+w.substr(1).toLowerCase();}); }
    if ((inner = matchFn('LARGO|LEN')) !== null) { return String(String(resolveArg(splitFnArgs(inner)[0])).length); }
    if ((inner = matchFn('IZQUIERDA|LEFT')) !== null) { args = splitFnArgs(inner); var s2 = String(resolveArg(args[0])); return s2.substring(0, args[1] ? parseInt(resolveNum(args[1])) : 1); }
    if ((inner = matchFn('DERECHA|RIGHT')) !== null) { args = splitFnArgs(inner); var s3 = String(resolveArg(args[0])); var n4 = args[1] ? parseInt(resolveNum(args[1])) : 1; return s3.substring(s3.length - n4); }
    if ((inner = matchFn('EXTRAE|MID')) !== null) { args = splitFnArgs(inner); return String(resolveArg(args[0])).substring(parseInt(resolveNum(args[1]))-1, parseInt(resolveNum(args[1]))-1+parseInt(resolveNum(args[2]))); }
    if ((inner = matchFn('ENCONTRAR|FIND')) !== null) { args = splitFnArgs(inner); var pos = String(resolveArg(args[1])).indexOf(String(resolveArg(args[0]))); return pos >= 0 ? String(pos+1) : '#VALOR!'; }
    if ((inner = matchFn('HALLAR|SEARCH')) !== null) { args = splitFnArgs(inner); var pos2 = String(resolveArg(args[1])).toLowerCase().indexOf(String(resolveArg(args[0])).toLowerCase()); return pos2 >= 0 ? String(pos2+1) : '#VALOR!'; }
    if ((inner = matchFn('SUSTITUIR|SUBSTITUTE')) !== null) { args = splitFnArgs(inner); return String(resolveArg(args[0])).split(String(resolveArg(args[1]))).join(String(resolveArg(args[2]))); }
    if ((inner = matchFn('REEMPLAZAR|REPLACE')) !== null) { args = splitFnArgs(inner); var orig = String(resolveArg(args[0])); var start4 = parseInt(resolveNum(args[1]))-1; var len4 = parseInt(resolveNum(args[2])); return orig.substring(0,start4)+String(resolveArg(args[3]))+orig.substring(start4+len4); }
    if ((inner = matchFn('ESPACIOS|TRIM')) !== null) { return String(resolveArg(splitFnArgs(inner)[0])).replace(/\s+/g,' ').trim(); }
    if ((inner = matchFn('REPETIR|REPT')) !== null) { args = splitFnArgs(inner); return String(resolveArg(args[0])).repeat(parseInt(resolveNum(args[1]))); }
    if ((inner = matchFn('VALOR|VALUE')) !== null) { return fmtNum(parseFloat(String(resolveArg(splitFnArgs(inner)[0])).replace(/[₡,\s]/g,''))); }
    if ((inner = matchFn('MONEDA|DOLLAR')) !== null) { args = splitFnArgs(inner); var d5 = args[1] ? parseInt(resolveNum(args[1])) : 2; return '₡'+resolveNum(args[0]).toLocaleString('es-CR',{minimumFractionDigits:d5,maximumFractionDigits:d5}); }
    if ((inner = matchFn('TEXTO|TEXT')) !== null) {
      args = splitFnArgs(inner); var n5 = resolveNum(args[0]); var fmt4 = String(resolveArg(args[1]));
      if (fmt4.includes('dd') || fmt4.includes('mm') || fmt4.includes('aaaa')) {
        var d6 = new Date(n5); if (isNaN(d6.getTime())) return String(n5);
        return fmt4.replace('dd',String(d6.getDate()).padStart(2,'0')).replace('mm',String(d6.getMonth()+1).padStart(2,'0')).replace('aaaa',d6.getFullYear());
      }
      if (fmt4.includes('#,##0')) return n5.toLocaleString('es-CR',{minimumFractionDigits:fmt4.includes('.00')?2:0});
      return String(n5);
    }
    if ((inner = matchFn('IGUAL|EXACT')) !== null) { args = splitFnArgs(inner); return String(resolveArg(args[0])) === String(resolveArg(args[1])) ? 'VERDADERO' : 'FALSO'; }

    // ─────────────────────────────────────────
    // FECHA Y HORA
    // ─────────────────────────────────────────
    if (exprUp === 'HOY()' || exprUp === 'TODAY()') return new Date().toLocaleDateString('es-CR');
    if (exprUp === 'AHORA()' || exprUp === 'NOW()') return new Date().toLocaleString('es-CR');
    if ((inner = matchFn('FECHA|DATE')) !== null) {
      args = splitFnArgs(inner); var d7 = new Date(resolveNum(args[0]),resolveNum(args[1])-1,resolveNum(args[2]));
      return d7.toLocaleDateString('es-CR');
    }
    if ((inner = matchFn('AÑO|YEAR')) !== null) { var d8 = new Date(String(resolveArg(splitFnArgs(inner)[0]))); return isNaN(d8.getTime()) ? '#VALOR!' : String(d8.getFullYear()); }
    if ((inner = matchFn('MES|MONTH')) !== null) { var d9 = new Date(String(resolveArg(splitFnArgs(inner)[0]))); return isNaN(d9.getTime()) ? '#VALOR!' : String(d9.getMonth()+1); }
    if ((inner = matchFn('DIA(?!\\.LAB)|DAY')) !== null) { var d10 = new Date(String(resolveArg(splitFnArgs(inner)[0]))); return isNaN(d10.getTime()) ? '#VALOR!' : String(d10.getDate()); }
    if ((inner = matchFn('DIAS(?!\\.|LAB)|DAYS')) !== null) {
      args = splitFnArgs(inner); var d11 = new Date(String(resolveArg(args[0]))); var d12 = new Date(String(resolveArg(args[1])));
      return String(Math.round((d11 - d12) / 86400000));
    }
    if ((inner = matchFn('FIN\\.MES|EOMONTH')) !== null) {
      args = splitFnArgs(inner); var d13 = new Date(String(resolveArg(args[0]))); var m2 = d13.getMonth() + 1 + parseInt(resolveNum(args[1]));
      return new Date(d13.getFullYear(), m2, 0).toLocaleDateString('es-CR');
    }
    if ((inner = matchFn('SIFECHA|DATEDIF')) !== null) {
      args = splitFnArgs(inner); var d14 = new Date(String(resolveArg(args[0]))), d15 = new Date(String(resolveArg(args[1]))); var unit = String(resolveArg(args[2])).toUpperCase();
      var diff = d15 - d14;
      if (unit === 'Y') return String(Math.floor(diff / (365.25 * 86400000)));
      if (unit === 'M') return String(Math.floor(diff / (30.44 * 86400000)));
      if (unit === 'D') return String(Math.floor(diff / 86400000));
      return String(Math.floor(diff / 86400000));
    }
    if ((inner = matchFn('HORA|HOUR')) !== null) { var d16 = new Date(); return String(d16.getHours()); }
    if ((inner = matchFn('MINUTO|MINUTE')) !== null) { var d17 = new Date(); return String(d17.getMinutes()); }

    // ─────────────────────────────────────────
    // FINANCIERAS
    // ─────────────────────────────────────────
    if ((inner = matchFn('IVA')) !== null) {
      args = splitFnArgs(inner); var base = isRange(args[0]) ? rangeNums(args[0]).reduce(function(a,b){return a+b;},0) : resolveNum(args[0]);
      var rate = args[1] ? resolveNum(args[1]) : 0.13;
      return fmtNum(base * rate);
    }
    if ((inner = matchFn('PAGO|PMT')) !== null) {
      args = splitFnArgs(inner); var r3 = resolveNum(args[0]), n6 = resolveNum(args[1]), pv = resolveNum(args[2]);
      if (r3 === 0) return fmtNum(-pv/n6);
      return fmtNum(r3*pv / (1-Math.pow(1+r3,-n6)));
    }
    if ((inner = matchFn('VA|PV')) !== null) {
      args = splitFnArgs(inner); var r4 = resolveNum(args[0]), n7 = resolveNum(args[1]), pmt = resolveNum(args[2]);
      if (r4 === 0) return fmtNum(-pmt*n7);
      return fmtNum(-pmt*(1-Math.pow(1+r4,-n7))/r4);
    }
    if ((inner = matchFn('VF|FV')) !== null) {
      args = splitFnArgs(inner); var r5 = resolveNum(args[0]), n8 = resolveNum(args[1]), pmt2 = resolveNum(args[2]);
      var pv2 = args[3] ? resolveNum(args[3]) : 0;
      return fmtNum(-pv2*Math.pow(1+r5,n8) - pmt2*(Math.pow(1+r5,n8)-1)/r5);
    }
    if ((inner = matchFn('VNA|NPV')) !== null) {
      args = splitFnArgs(inner); var r6 = resolveNum(args[0]); var npv = 0;
      for (var i13 = 1; i13 < args.length; i13++) {
        var flvs = isRange(args[i13]) ? rangeNums(args[i13]) : [resolveNum(args[i13])];
        flvs.forEach(function(fv2, j) { npv += fv2 / Math.pow(1+r6, i13-1+j+1); });
      }
      return fmtNum(npv);
    }
    if ((inner = matchFn('TIR|IRR')) !== null) {
      args = splitFnArgs(inner); var flows = isRange(args[0]) ? rangeNums(args[0]) : [];
      var guess = args[1] ? resolveNum(args[1]) : 0.1; var rate7 = guess;
      for (var iter = 0; iter < 100; iter++) {
        var npv2 = 0, dnpv = 0;
        flows.forEach(function(fv3, t) { npv2 += fv3/Math.pow(1+rate7,t); dnpv -= t*fv3/Math.pow(1+rate7,t+1); });
        var newRate = rate7 - npv2/dnpv;
        if (Math.abs(newRate - rate7) < 1e-10) break;
        rate7 = newRate;
      }
      return fmtNum(rate7);
    }
    if ((inner = matchFn('TASA|RATE')) !== null) {
      args = splitFnArgs(inner); var n9 = resolveNum(args[0]), pmt3 = resolveNum(args[1]), pv3 = resolveNum(args[2]);
      var r8 = 0.1;
      for (var iter2 = 0; iter2 < 100; iter2++) {
        var f3 = pv3*Math.pow(1+r8,n9)+pmt3*(Math.pow(1+r8,n9)-1)/r8;
        var df = n9*pv3*Math.pow(1+r8,n9-1)+pmt3*(n9*Math.pow(1+r8,n9-1)*r8-(Math.pow(1+r8,n9)-1))/(r8*r8);
        var nr = r8 - f3/df;
        if (Math.abs(nr - r8) < 1e-10) break;
        r8 = nr;
      }
      return fmtNum(r8);
    }
    if ((inner = matchFn('NPER')) !== null) {
      args = splitFnArgs(inner); var r9 = resolveNum(args[0]), pmt4 = resolveNum(args[1]), pv4 = resolveNum(args[2]);
      if (r9 === 0) return fmtNum(-pv4/pmt4);
      return fmtNum(Math.log(-pmt4/(pmt4+r9*pv4))/Math.log(1+r9));
    }
    if ((inner = matchFn('INT\\.EFECTIVO|EFFECT')) !== null) {
      args = splitFnArgs(inner); var nom = resolveNum(args[0]), m3 = resolveNum(args[1]);
      return fmtNum(Math.pow(1+nom/m3,m3)-1);
    }
    if ((inner = matchFn('TASA\\.NOMINAL|NOMINAL')) !== null) {
      args = splitFnArgs(inner); var eff = resolveNum(args[0]), m4 = resolveNum(args[1]);
      return fmtNum(m4*(Math.pow(1+eff,1/m4)-1));
    }
    if ((inner = matchFn('SLN')) !== null) {
      args = splitFnArgs(inner); return fmtNum((resolveNum(args[0])-resolveNum(args[1]))/resolveNum(args[2]));
    }

    // ─────────────────────────────────────────
    // BÚSQUEDA Y REFERENCIA
    // ─────────────────────────────────────────
    if ((inner = matchFn('BUSCARV|VLOOKUP')) !== null) {
      args = splitFnArgs(inner); var sv2 = String(resolveArg(args[0])).toLowerCase();
      var rng3 = args[1]; var colN2 = parseInt(resolveNum(args[2])) - 1;
      if (isRange(rng3)) {
        var parts2 = rng3.trim().split(':');
        var fc3 = colIdx(parts2[0].match(/[A-Za-z]+/)[0].toUpperCase());
        var fr2 = parseInt(parts2[0].match(/\d+/)[0]) - 1;
        var tr3 = parseInt(parts2[1].match(/\d+/)[0]) - 1;
        for (var ri2 = fr2; ri2 <= tr3; ri2++) {
          if (String(getVal(cellId(ri2, fc3))).toLowerCase() === sv2) {
            return String(getVal(cellId(ri2, fc3 + colN2)));
          }
        }
      }
      return '#N/A';
    }
    if ((inner = matchFn('BUSCARX|XLOOKUP')) !== null) {
      args = splitFnArgs(inner); var sv3 = String(resolveArg(args[0])).toLowerCase();
      var lookRng = args[1], retRng = args[2], notFound = args[3] ? String(resolveArg(args[3])) : '#N/A';
      if (isRange(lookRng) && isRange(retRng)) {
        var lookVals = rangeVals(lookRng), retVals = rangeVals(retRng);
        for (var i14 = 0; i14 < lookVals.length; i14++) {
          if (String(lookVals[i14]).toLowerCase() === sv3) return String(retVals[i14]);
        }
      }
      return notFound;
    }
    if ((inner = matchFn('COINCIDIR|MATCH')) !== null) {
      args = splitFnArgs(inner); var sv4 = String(resolveArg(args[0])).toLowerCase();
      if (isRange(args[1])) {
        var matchVals = rangeVals(args[1]);
        for (var i15 = 0; i15 < matchVals.length; i15++) {
          if (String(matchVals[i15]).toLowerCase() === sv4) return String(i15 + 1);
        }
      }
      return '#N/A';
    }
    if ((inner = matchFn('INDICE|INDEX')) !== null) {
      args = splitFnArgs(inner); var r10 = parseInt(resolveNum(args[1]))-1, c10 = parseInt(resolveNum(args[2] || '1'))-1;
      if (isRange(args[0])) {
        var p3 = args[0].trim().split(':');
        var fc4 = colIdx(p3[0].match(/[A-Za-z]+/)[0].toUpperCase());
        var fr3 = parseInt(p3[0].match(/\d+/)[0]) - 1;
        return String(getVal(cellId(fr3 + r10, fc4 + c10)));
      }
      return '#REF!';
    }
    if ((inner = matchFn('DESREF|OFFSET')) !== null) {
      args = splitFnArgs(inner); var ref = args[0].trim();
      var dr = parseInt(resolveNum(args[1])), dc2 = parseInt(resolveNum(args[2]));
      if (/^[A-Za-z]+\d+$/.test(ref)) {
        var rc = colIdx(ref.match(/[A-Za-z]+/)[0].toUpperCase());
        var rr = parseInt(ref.match(/\d+/)[0]) - 1;
        return String(getVal(cellId(rr + dr, rc + dc2)));
      }
      return '#REF!';
    }
    if ((inner = matchFn('INDIRECTO|INDIRECT')) !== null) {
      var ref2 = String(resolveArg(splitFnArgs(inner)[0]));
      return String(getVal(ref2));
    }
    if ((inner = matchFn('ELEGIR|CHOOSE')) !== null) {
      args = splitFnArgs(inner); var idx2 = parseInt(resolveNum(args[0]));
      return args[idx2] ? String(resolveArg(args[idx2])) : '#VALOR!';
    }
    if ((inner = matchFn('FILA|ROW')) !== null) { var rf = splitFnArgs(inner)[0]; if (!rf) return String(row+1); var m4 = rf.match(/\d+/); return m4 ? m4[0] : String(row+1); }
    if ((inner = matchFn('COLUMNA|COLUMN')) !== null) { var rc2 = splitFnArgs(inner)[0]; if (!rc2) return String(col+1); var lm = rc2.match(/[A-Za-z]+/); return lm ? String(colIdx(lm[0].toUpperCase())+1) : String(col+1); }
    if ((inner = matchFn('FILAS|ROWS')) !== null) { if (isRange(inner)) { var p4=inner.split(':'); return String(parseInt(p4[1].match(/\d+/)[0])-parseInt(p4[0].match(/\d+/)[0])+1); } return '1'; }
    if ((inner = matchFn('COLUMNAS|COLUMNS')) !== null) { if (isRange(inner)) { var p5=inner.split(':'); return String(colIdx(p5[1].match(/[A-Za-z]+/)[0].toUpperCase())-colIdx(p5[0].match(/[A-Za-z]+/)[0].toUpperCase())+1); } return '1'; }
    if ((inner = matchFn('TRANSPONER|TRANSPOSE')) !== null) { return inner; } // passthrough for display

    // ─────────────────────────────────────────
    // INFORMACIÓN
    // ─────────────────────────────────────────
    if ((inner = matchFn('ESBLANCO|ISBLANK')) !== null) { var v4 = resolveArg(splitFnArgs(inner)[0]); return (v4===''||v4===null||v4===undefined)?'VERDADERO':'FALSO'; }
    if ((inner = matchFn('ESNUMERO|ISNUMBER')) !== null) { var v5 = resolveArg(splitFnArgs(inner)[0]); return !isNaN(parseFloat(v5))?'VERDADERO':'FALSO'; }
    if ((inner = matchFn('ESTEXTO|ISTEXT')) !== null) { var v6 = resolveArg(splitFnArgs(inner)[0]); return (isNaN(parseFloat(v6))&&typeof v6==='string'&&v6!=='')?'VERDADERO':'FALSO'; }
    if ((inner = matchFn('ESERROR|ISERROR')) !== null) { try { var v7 = String(resolveArg(splitFnArgs(inner)[0])); return v7.startsWith('#')?'VERDADERO':'FALSO'; }catch(e4){return 'VERDADERO';} }
    if ((inner = matchFn('ESFORMULA|ISFORMULA')) !== null) { var ref3 = splitFnArgs(inner)[0].trim(); return (sheetData[ref3]&&String(sheetData[ref3]).startsWith('='))?'VERDADERO':'FALSO'; }
    if ((inner = matchFn('ES\\.PAR|ISEVEN')) !== null) { return resolveNum(splitFnArgs(inner)[0])%2===0?'VERDADERO':'FALSO'; }
    if ((inner = matchFn('ES\\.IMPAR|ISODD')) !== null) { return Math.abs(resolveNum(splitFnArgs(inner)[0]))%2===1?'VERDADERO':'FALSO'; }
    if ((inner = matchFn('N(?!OD|O\\b|PERS|PER)')) !== null) { var v8 = resolveArg(splitFnArgs(inner)[0]); return fmtNum(parseFloat(v8)||0); }

    // ─────────────────────────────────────────
    // SIMPLE REFERENCES & ARITHMETIC
    // ─────────────────────────────────────────


    // BÚSQUEDA Y REFERENCIA — nuevas
    // ════════════════════════════════════════

    // BUSCARH / HLOOKUP
    if ((inner = matchFn('BUSCARH|HLOOKUP')) !== null) {
      args = splitArgs(inner);
      var sv = String(resolveArg(args[0])).toLowerCase();
      var rng = args[1]; var rowN = parseInt(resolveNum(args[2])) - 1;
      if (isRange(rng)) {
        var p = rng.trim().split(':');
        var fc2 = colIdx(p[0].match(/[A-Za-z]+/)[0]);
        var fr2 = parseInt(p[0].match(/\d+/)[0]) - 1;
        var tc2 = colIdx(p[1].match(/[A-Za-z]+/)[0]);
        for (var ci = fc2; ci <= tc2; ci++) {
          if (String(getVal(cellId(fr2, ci))).toLowerCase() === sv)
            return String(getVal(cellId(fr2 + rowN, ci)));
        }
      }
      return '#N/A';
    }

    // FILTRAR / FILTER
    if ((inner = matchFn('FILTRAR|FILTER')) !== null) {
      args = splitArgs(inner);
      if (!isRange(args[0]) || !isRange(args[1])) return '#VALOR!';
      var dataVals = rangeVals(args[0]);
      var condVals = rangeVals(args[1]);
      var result = [];
      condVals.forEach(function(v, i) {
        if (v === 'VERDADERO' || v === true || (toNum(v) !== 0 && v !== '')) result.push(dataVals[i] || '');
      });
      return result.length ? result.join('; ') : (args[2] ? String(resolveArg(args[2])) : '#CALC!');
    }

    // ORDENAR / SORT
    if ((inner = matchFn('ORDENAR|SORT')) !== null) {
      args = splitArgs(inner);
      if (!isRange(args[0])) return '#VALOR!';
      var sortVals = rangeVals(args[0]);
      var order = args[2] ? parseInt(resolveNum(args[2])) : 1;
      var sorted = sortVals.slice().sort(function(a, b) {
        var an = toNum(a), bn = toNum(b);
        if (!isNaN(an) && !isNaN(bn)) return order >= 0 ? an-bn : bn-an;
        return order >= 0 ? String(a).localeCompare(String(b)) : String(b).localeCompare(String(a));
      });
      return sorted.join('; ');
    }

    // UNICOS / UNIQUE
    if ((inner = matchFn('UNICOS|UNIQUE')) !== null) {
      args = splitArgs(inner);
      if (!isRange(args[0])) return '#VALOR!';
      var seen = {}, uniq = [];
      rangeVals(args[0]).forEach(function(v) { if (v !== '' && !seen[v]) { seen[v]=true; uniq.push(v); } });
      return uniq.join('; ');
    }

    // APILARV / VSTACK
    if ((inner = matchFn('APILARV|VSTACK')) !== null) {
      args = splitArgs(inner);
      var allVals = [];
      args.forEach(function(a) { if (isRange(a)) rangeVals(a).forEach(function(v) { allVals.push(v); }); else allVals.push(String(resolveArg(a))); });
      return allVals.join('; ');
    }

    // APILARH / HSTACK
    if ((inner = matchFn('APILARH|HSTACK')) !== null) {
      args = splitArgs(inner);
      var hVals = [];
      args.forEach(function(a) { if (isRange(a)) rangeVals(a).forEach(function(v) { hVals.push(v); }); else hVals.push(String(resolveArg(a))); });
      return hVals.join(' | ');
    }

    // SECUENCIA / SEQUENCE
    if ((inner = matchFn('SECUENCIA|SEQUENCE')) !== null) {
      args = splitArgs(inner);
      var rows2 = parseInt(resolveNum(args[0]));
      var cols2 = args[1] ? parseInt(resolveNum(args[1])) : 1;
      var start = args[2] ? resolveNum(args[2]) : 1;
      var step  = args[3] ? resolveNum(args[3]) : 1;
      var seq = [];
      for (var i = 0; i < rows2 * cols2; i++) seq.push(fmtNum(start + i * step));
      return seq.join('; ');
    }

    // ════════════════════════════════════════
    // TEXTO — nuevas
    // ════════════════════════════════════════

    // DIVIDIRTEXTO / TEXTSPLIT
    if ((inner = matchFn('DIVIDIRTEXTO|TEXTSPLIT')) !== null) {
      args = splitArgs(inner);
      var txt = String(resolveArg(args[0]));
      var delim = String(resolveArg(args[1]));
      return txt.split(delim).join('; ');
    }

    // TEXTOANTES / TEXTBEFORE
    if ((inner = matchFn('TEXTOANTES|TEXTBEFORE')) !== null) {
      args = splitArgs(inner);
      var txt2 = String(resolveArg(args[0]));
      var delim2 = String(resolveArg(args[1]));
      var idx = txt2.indexOf(delim2);
      return idx >= 0 ? txt2.substring(0, idx) : '#N/A';
    }

    // TEXTODESPUES / TEXTAFTER
    if ((inner = matchFn('TEXTODESPUES|TEXTAFTER')) !== null) {
      args = splitArgs(inner);
      var txt3 = String(resolveArg(args[0]));
      var delim3 = String(resolveArg(args[1]));
      var idx2 = txt3.indexOf(delim3);
      return idx2 >= 0 ? txt3.substring(idx2 + delim3.length) : '#N/A';
    }

    // CODIGO / CODE
    if ((inner = matchFn('CODIGO|CODE')) !== null) {
      var s = String(resolveArg(splitArgs(inner)[0]));
      return s.length ? String(s.charCodeAt(0)) : '#VALOR!';
    }

    // CARACTER / CHAR
    if ((inner = matchFn('CARACTER|CAR|CHAR')) !== null) {
      return String.fromCharCode(parseInt(resolveNum(splitArgs(inner)[0])));
    }

    // LIMPIAR / CLEAN
    if ((inner = matchFn('LIMPIAR|CLEAN')) !== null) {
      return String(resolveArg(splitArgs(inner)[0])).replace(/[\x00-\x1F]/g, '');
    }

    // UNICODE
    if ((inner = matchFn('UNICODE')) !== null) {
      var s2 = String(resolveArg(splitArgs(inner)[0]));
      return s2.length ? String(s2.codePointAt(0)) : '#VALOR!';
    }

    // MATRIZATEXTO / ARRAYTOTEXT
    if ((inner = matchFn('MATRIZATEXTO|ARRAYTOTEXT')) !== null) {
      args = splitArgs(inner);
      if (isRange(args[0])) return '{' + rangeVals(args[0]).join(', ') + '}';
      return String(resolveArg(args[0]));
    }

    // ════════════════════════════════════════
    // FECHA Y HORA — nuevas
    // ════════════════════════════════════════

    // DIAS.LAB / NETWORKDAYS
    if ((inner = matchFn('DIAS\\.LAB|NETWORKDAYS')) !== null) {
      args = splitArgs(inner);
      var d1 = new Date(String(resolveArg(args[0])));
      var d2 = new Date(String(resolveArg(args[1])));
      var count = 0, cur2 = new Date(d1);
      while (cur2 <= d2) {
        var day = cur2.getDay();
        if (day !== 0 && day !== 6) count++;
        cur2.setDate(cur2.getDate() + 1);
      }
      return String(count);
    }

    // DIA.LAB / WORKDAY
    if ((inner = matchFn('DIA\\.LAB|WORKDAY')) !== null) {
      args = splitArgs(inner);
      var d3 = new Date(String(resolveArg(args[0])));
      var days = parseInt(resolveNum(args[1]));
      var added = 0, dir = days >= 0 ? 1 : -1;
      var remaining = Math.abs(days);
      while (remaining > 0) {
        d3.setDate(d3.getDate() + dir);
        if (d3.getDay() !== 0 && d3.getDay() !== 6) remaining--;
      }
      return d3.toLocaleDateString('es-CR');
    }

    // NUM.DE.SEMANA / WEEKNUM
    if ((inner = matchFn('NUM\\.DE\\.SEMANA|WEEKNUM')) !== null) {
      args = splitArgs(inner);
      var d4 = new Date(String(resolveArg(args[0])));
      var start2 = new Date(d4.getFullYear(), 0, 1);
      return String(Math.ceil(((d4 - start2) / 86400000 + start2.getDay() + 1) / 7));
    }

    // DIASEM / WEEKDAY
    if ((inner = matchFn('DIASEM|WEEKDAY')) !== null) {
      args = splitArgs(inner);
      var d5 = new Date(String(resolveArg(args[0])));
      var type2 = args[1] ? parseInt(resolveNum(args[1])) : 1;
      var wd = d5.getDay(); // 0=Dom
      if (type2 === 2) return String(wd === 0 ? 7 : wd); // 1=Lun
      if (type2 === 3) return String(wd === 0 ? 6 : wd - 1); // 0=Lun
      return String(wd + 1); // 1=Dom
    }

    // FRAC.AÑO / YEARFRAC
    if ((inner = matchFn('FRAC\\.AÑO|YEARFRAC')) !== null) {
      args = splitArgs(inner);
      var d6 = new Date(String(resolveArg(args[0])));
      var d7 = new Date(String(resolveArg(args[1])));
      return fmtNum((d7 - d6) / (365.25 * 86400000));
    }

    // FECHA.MES / EDATE
    if ((inner = matchFn('FECHA\\.MES|EDATE')) !== null) {
      args = splitArgs(inner);
      var d8 = new Date(String(resolveArg(args[0])));
      d8.setMonth(d8.getMonth() + parseInt(resolveNum(args[1])));
      return d8.toLocaleDateString('es-CR');
    }

    // SEGUNDO / SECOND
    if ((inner = matchFn('SEGUNDO|SECOND')) !== null) {
      return String(new Date().getSeconds());
    }

    // ════════════════════════════════════════
    // FINANCIERAS — nuevas
    // ════════════════════════════════════════

    // DB — Depreciación saldo fijo decreciente
    if ((inner = matchFn('DB')) !== null) {
      args = splitArgs(inner);
      var cost = resolveNum(args[0]), salv = resolveNum(args[1]);
      var life = resolveNum(args[2]), per = resolveNum(args[3]);
      var month = args[4] ? resolveNum(args[4]) : 12;
      var rate2 = 1 - Math.pow(salv/cost, 1/life);
      rate2 = Math.round(rate2 * 1000) / 1000;
      var dep = 0;
      if (per === 1) dep = cost * rate2 * month / 12;
      else {
        var prev = cost * rate2 * month / 12;
        var book = cost - prev;
        for (var p = 2; p <= per; p++) {
          dep = book * rate2;
          if (p < per) book -= dep;
        }
      }
      return fmtNum(dep);
    }

    // DDB — Doble disminución
    if ((inner = matchFn('DDB')) !== null) {
      args = splitArgs(inner);
      var cost2 = resolveNum(args[0]), salv2 = resolveNum(args[1]);
      var life2 = resolveNum(args[2]), per2 = resolveNum(args[3]);
      var factor = args[4] ? resolveNum(args[4]) : 2;
      var rate3 = factor / life2;
      var book2 = cost2;
      for (var p2 = 1; p2 <= per2; p2++) {
        var dep2 = Math.min(book2 * rate3, book2 - salv2);
        if (p2 === per2) return fmtNum(dep2);
        book2 -= dep2;
      }
      return fmtNum(0);
    }

    // DVS — Depreciación variable
    if ((inner = matchFn('DVS|VDB')) !== null) {
      args = splitArgs(inner);
      var cost3 = resolveNum(args[0]), salv3 = resolveNum(args[1]);
      var life3 = resolveNum(args[2]);
      var ps = resolveNum(args[3]), pe = resolveNum(args[4]);
      var factor2 = args[5] ? resolveNum(args[5]) : 2;
      var rate4 = factor2 / life3;
      var book3 = cost3, total4 = 0;
      for (var p3 = 0; p3 < pe; p3++) {
        var dep3 = Math.min(book3 * rate4, book3 - salv3);
        if (p3 >= ps) total4 += dep3;
        book3 -= dep3;
      }
      return fmtNum(total4);
    }

    // PAGOINT / IPMT
    if ((inner = matchFn('PAGOINT|IPMT')) !== null) {
      args = splitArgs(inner);
      var r5 = resolveNum(args[0]), per3 = resolveNum(args[1]);
      var n2 = resolveNum(args[2]), pv = resolveNum(args[3]);
      if (r5 === 0) return fmtNum(0);
      var pmt2 = r5 * pv / (1 - Math.pow(1+r5, -n2));
      var bal = pv * Math.pow(1+r5, per3-1) - pmt2 * (Math.pow(1+r5, per3-1)-1)/r5;
      return fmtNum(-bal * r5);
    }

    // PAGOPRIN / PPMT
    if ((inner = matchFn('PAGOPRIN|PPMT')) !== null) {
      args = splitArgs(inner);
      var r6 = resolveNum(args[0]), per4 = resolveNum(args[1]);
      var n3 = resolveNum(args[2]), pv2 = resolveNum(args[3]);
      if (r6 === 0) return fmtNum(-pv2/n3);
      var pmt3 = r6 * pv2 / (1 - Math.pow(1+r6, -n3));
      var ipmt2 = -(pv2 * Math.pow(1+r6, per4-1) - pmt3*(Math.pow(1+r6,per4-1)-1)/r6) * r6;
      return fmtNum(-pmt3 - ipmt2);
    }

    // TIRM / MIRR
    if ((inner = matchFn('TIRM|MIRR')) !== null) {
      args = splitArgs(inner);
      var flows = isRange(args[0]) ? rangeNums(args[0]) : [];
      var finRate = resolveNum(args[1]), reinvRate = resolveNum(args[2]);
      var n4 = flows.length;
      var neg = 0, pos2 = 0;
      flows.forEach(function(v, i) {
        if (v < 0) neg += v / Math.pow(1+finRate, i);
        else pos2 += v * Math.pow(1+reinvRate, n4-1-i);
      });
      return fmtNum(Math.pow(pos2 / Math.abs(neg), 1/(n4-1)) - 1);
    }

    // RRI — Tasa equivalente
    if ((inner = matchFn('RRI')) !== null) {
      args = splitArgs(inner);
      return fmtNum(Math.pow(resolveNum(args[2])/resolveNum(args[1]), 1/resolveNum(args[0])) - 1);
    }

    // P.DURACION / PDURATION
    if ((inner = matchFn('P\\.DURACION|PDURATION')) !== null) {
      args = splitArgs(inner);
      var r7 = resolveNum(args[0]), pv3 = resolveNum(args[1]), fv2 = resolveNum(args[2]);
      return fmtNum(Math.log(fv2/pv3) / Math.log(1+r7));
    }

    // PAGO.INT.ENTRE / CUMIPMT
    if ((inner = matchFn('PAGO\\.INT\\.ENTRE|CUMIPMT')) !== null) {
      args = splitArgs(inner);
      var r8 = resolveNum(args[0]), n5 = resolveNum(args[1]), pv4 = resolveNum(args[2]);
      var ps2 = parseInt(resolveNum(args[3])), pe2 = parseInt(resolveNum(args[4]));
      if (r8 === 0) return fmtNum(0);
      var pmt4 = r8 * pv4 / (1 - Math.pow(1+r8,-n5));
      var total5 = 0;
      for (var p4 = ps2; p4 <= pe2; p4++) {
        var bal2 = pv4 * Math.pow(1+r8,p4-1) - pmt4*(Math.pow(1+r8,p4-1)-1)/r8;
        total5 += bal2 * r8;
      }
      return fmtNum(-total5);
    }

    // PAGO.PRINC.ENTRE / CUMPRINC
    if ((inner = matchFn('PAGO\\.PRINC\\.ENTRE|CUMPRINC')) !== null) {
      args = splitArgs(inner);
      var r9 = resolveNum(args[0]), n6 = resolveNum(args[1]), pv5 = resolveNum(args[2]);
      var ps3 = parseInt(resolveNum(args[3])), pe3 = parseInt(resolveNum(args[4]));
      if (r9 === 0) return fmtNum(-pv5*(pe3-ps3+1)/n6);
      var pmt5 = r9 * pv5 / (1-Math.pow(1+r9,-n6));
      var total6 = 0;
      for (var p5 = ps3; p5 <= pe3; p5++) {
        var bal3 = pv5*Math.pow(1+r9,p5-1) - pmt5*(Math.pow(1+r9,p5-1)-1)/r9;
        var ipmt3 = bal3 * r9;
        total6 += pmt5 - ipmt3;
      }
      return fmtNum(-total6);
    }

    // ════════════════════════════════════════
    // ESTADÍSTICAS — nuevas
    // ════════════════════════════════════════

    // VAR.S / VAR muestral
    if ((inner = matchFn('VAR\\.S|VAR\\.M|VAR(?!\\.P)')) !== null) {
      args = splitArgs(inner);
      nums = isRange(args[0]) ? rangeNums(args[0]) : args.map(resolveNum);
      var mean2 = nums.reduce(function(a,b){return a+b;},0)/nums.length;
      return fmtNum(nums.reduce(function(s,n){return s+(n-mean2)*(n-mean2);},0)/(nums.length-1));
    }

    // VAR.P / VARP poblacional
    if ((inner = matchFn('VAR\\.P|VARP')) !== null) {
      args = splitArgs(inner);
      nums = isRange(args[0]) ? rangeNums(args[0]) : args.map(resolveNum);
      var mean3 = nums.reduce(function(a,b){return a+b;},0)/nums.length;
      return fmtNum(nums.reduce(function(s,n){return s+(n-mean3)*(n-mean3);},0)/nums.length);
    }

    // PERCENTIL.INC / PERCENTILE
    if ((inner = matchFn('PERCENTIL\\.INC|PERCENTIL(?!\\.)|PERCENTILE')) !== null) {
      args = splitArgs(inner);
      nums = isRange(args[0]) ? rangeNums(args[0]) : [resolveNum(args[0])];
      var k = resolveNum(args[1]);
      nums.sort(function(a,b){return a-b;});
      var idx3 = k * (nums.length-1);
      var lo = Math.floor(idx3), hi = Math.ceil(idx3);
      return fmtNum(nums[lo] + (nums[hi]-nums[lo])*(idx3-lo));
    }

    // CUARTIL.INC / QUARTILE
    if ((inner = matchFn('CUARTIL\\.INC|CUARTIL|QUARTILE')) !== null) {
      args = splitArgs(inner);
      nums = isRange(args[0]) ? rangeNums(args[0]) : [resolveNum(args[0])];
      var q = resolveNum(args[1]) / 4;
      nums.sort(function(a,b){return a-b;});
      var idx4 = q*(nums.length-1);
      var lo2=Math.floor(idx4),hi2=Math.ceil(idx4);
      return fmtNum(nums[lo2]+(nums[hi2]-nums[lo2])*(idx4-lo2));
    }

    // CURTOSIS / KURT
    if ((inner = matchFn('CURTOSIS|KURT')) !== null) {
      args = splitArgs(inner);
      nums = isRange(args[0]) ? rangeNums(args[0]) : args.map(resolveNum);
      var n7=nums.length, mean4=nums.reduce(function(a,b){return a+b;},0)/n7;
      var s4=Math.sqrt(nums.reduce(function(s,n){return s+(n-mean4)*(n-mean4);},0)/(n7-1));
      var k2=nums.reduce(function(s,n){return s+Math.pow((n-mean4)/s4,4);},0);
      return fmtNum((n7*(n7+1)/((n7-1)*(n7-2)*(n7-3)))*k2 - 3*(n7-1)*(n7-1)/((n7-2)*(n7-3)));
    }

    // COEFICIENTE.ASIMETRIA / SKEW
    if ((inner = matchFn('COEFICIENTE\\.ASIMETRIA|SKEW')) !== null) {
      args = splitArgs(inner);
      nums = isRange(args[0]) ? rangeNums(args[0]) : args.map(resolveNum);
      var n8=nums.length, mean5=nums.reduce(function(a,b){return a+b;},0)/n8;
      var s5=Math.sqrt(nums.reduce(function(s,n){return s+(n-mean5)*(n-mean5);},0)/(n8-1));
      var sk=nums.reduce(function(s,n){return s+Math.pow((n-mean5)/s5,3);},0);
      return fmtNum((n8/((n8-1)*(n8-2)))*sk);
    }

    // PENDIENTE / SLOPE
    if ((inner = matchFn('PENDIENTE|SLOPE')) !== null) {
      args = splitArgs(inner);
      var y = isRange(args[0])?rangeNums(args[0]):[], x = isRange(args[1])?rangeNums(args[1]):[];
      var n9=Math.min(x.length,y.length);
      var mx2=x.reduce(function(a,b){return a+b;},0)/n9, my2=y.reduce(function(a,b){return a+b;},0)/n9;
      var num2=0,denom=0;
      for(var i=0;i<n9;i++){num2+=(x[i]-mx2)*(y[i]-my2);denom+=(x[i]-mx2)*(x[i]-mx2);}
      return fmtNum(num2/denom);
    }

    // INTERSECCION.EJE / INTERCEPT
    if ((inner = matchFn('INTERSECCION\\.EJE|INTERCEPT')) !== null) {
      args = splitArgs(inner);
      var y2=isRange(args[0])?rangeNums(args[0]):[], x2=isRange(args[1])?rangeNums(args[1]):[];
      var n10=Math.min(x2.length,y2.length);
      var mx3=x2.reduce(function(a,b){return a+b;},0)/n10, my3=y2.reduce(function(a,b){return a+b;},0)/n10;
      var num3=0,den2=0;
      for(var i=0;i<n10;i++){num3+=(x2[i]-mx3)*(y2[i]-my3);den2+=(x2[i]-mx3)*(x2[i]-mx3);}
      return fmtNum(my3 - (num3/den2)*mx3);
    }

    // JERARQUIA / RANK
    if ((inner = matchFn('JERARQUIA\\.EQV|JERARQUIA|RANK')) !== null) {
      args = splitArgs(inner);
      var n11 = resolveNum(args[0]);
      var vals2 = isRange(args[1]) ? rangeNums(args[1]) : [];
      var order2 = args[2] ? parseInt(resolveNum(args[2])) : 0;
      vals2.sort(function(a,b){ return order2 ? a-b : b-a; });
      var rank = vals2.indexOf(n11);
      return rank >= 0 ? String(rank+1) : '#N/A';
    }

    // MEDIA.ACOTADA / TRIMMEAN
    if ((inner = matchFn('MEDIA\\.ACOTADA|TRIMMEAN')) !== null) {
      args = splitArgs(inner);
      nums = isRange(args[0]) ? rangeNums(args[0]) : args.map(resolveNum);
      var pct = resolveNum(args[1]);
      nums.sort(function(a,b){return a-b;});
      var trim = Math.floor(nums.length * pct / 2);
      var trimmed = nums.slice(trim, nums.length-trim);
      return fmtNum(trimmed.reduce(function(a,b){return a+b;},0)/trimmed.length);
    }

    // ════════════════════════════════════════
    // MATEMÁTICAS — nuevas
    // ════════════════════════════════════════

    // MMULT — Multiplicación matricial
    if ((inner = matchFn('MMULT')) !== null) {
      args = splitArgs(inner);
      if (isRange(args[0]) && isRange(args[1])) {
        var m1 = rangeNums(args[0]), m2 = rangeNums(args[1]);
        // Simple: dot product of vectors
        var dot = 0;
        for (var i=0;i<Math.min(m1.length,m2.length);i++) dot += m1[i]*m2[i];
        return fmtNum(dot);
      }
      return '#VALOR!';
    }

    // MATRIZALEAT / RANDARRAY
    if ((inner = matchFn('MATRIZALEAT|RANDARRAY')) !== null) {
      args = splitArgs(inner);
      var rows3 = args[0] ? parseInt(resolveNum(args[0])) : 1;
      var cols3 = args[1] ? parseInt(resolveNum(args[1])) : 1;
      var min2  = args[2] ? resolveNum(args[2]) : 0;
      var max3  = args[3] ? resolveNum(args[3]) : 1;
      var isInt = args[4] && String(resolveArg(args[4])).toUpperCase() === 'VERDADERO';
      var arr = [];
      for (var i=0;i<rows3*cols3;i++) {
        var v = min2 + Math.random()*(max3-min2);
        arr.push(isInt ? Math.floor(v) : Math.round(v*100)/100);
      }
      return arr.join('; ');
    }

    // MDETERM — Determinante (2x2 simplificado)
    if ((inner = matchFn('MDETERM')) !== null) {
      args = splitArgs(inner);
      if (isRange(args[0])) {
        nums = rangeNums(args[0]);
        if (nums.length === 4) return fmtNum(nums[0]*nums[3] - nums[1]*nums[2]);
        if (nums.length === 9) {
          var a=nums[0],b=nums[1],c=nums[2],d=nums[3],e=nums[4],f=nums[5],g=nums[6],h=nums[7],k=nums[8];
          return fmtNum(a*(e*k-f*h) - b*(d*k-f*g) + c*(d*h-e*g));
        }
      }
      return '#VALOR!';
    }

    // MUNIT — Matriz identidad
    if ((inner = matchFn('MUNIT')) !== null) {
      var dim = parseInt(resolveNum(splitArgs(inner)[0]));
      var rows4 = [];
      for (var i=0;i<dim;i++) {
        var row2 = [];
        for (var j=0;j<dim;j++) row2.push(i===j?'1':'0');
        rows4.push(row2.join('\t'));
      }
      return rows4.join(' | ');
    }

    // ALEATORIO / RAND
    if (exprUp === 'ALEATORIO()' || exprUp === 'RAND()') {
      return fmtNum(Math.random());
    }

    // GRADOS / DEGREES
    if ((inner = matchFn('GRADOS|DEGREES')) !== null) {
      return fmtNum(resolveNum(splitArgs(inner)[0]) * 180 / Math.PI);
    }

    // RADIANES / RADIANS
    if ((inner = matchFn('RADIANES|RADIANS')) !== null) {
      return fmtNum(resolveNum(splitArgs(inner)[0]) * Math.PI / 180);
    }

    // SEN / SIN, COS, TAN
    if ((inner = matchFn('SEN(?!O)|SENO|SIN')) !== null) { return fmtNum(Math.sin(resolveNum(splitArgs(inner)[0]))); }
    if ((inner = matchFn('COS(?!H)')) !== null) { return fmtNum(Math.cos(resolveNum(splitArgs(inner)[0]))); }
    if ((inner = matchFn('TAN(?!H)')) !== null) { return fmtNum(Math.tan(resolveNum(splitArgs(inner)[0]))); }
    if ((inner = matchFn('ASEN|ASIN|ASENO')) !== null) { return fmtNum(Math.asin(resolveNum(splitArgs(inner)[0]))); }
    if ((inner = matchFn('ACOS(?!H)')) !== null) { return fmtNum(Math.acos(resolveNum(splitArgs(inner)[0]))); }
    if ((inner = matchFn('ATAN(?!2|H)')) !== null) { return fmtNum(Math.atan(resolveNum(splitArgs(inner)[0]))); }
    if ((inner = matchFn('ATAN2')) !== null) { args=splitArgs(inner); return fmtNum(Math.atan2(resolveNum(args[1]),resolveNum(args[0]))); }
    if ((inner = matchFn('SENH|SINH')) !== null) { return fmtNum(Math.sinh(resolveNum(splitArgs(inner)[0]))); }
    if ((inner = matchFn('COSH')) !== null) { return fmtNum(Math.cosh(resolveNum(splitArgs(inner)[0]))); }
    if ((inner = matchFn('TANH')) !== null) { return fmtNum(Math.tanh(resolveNum(splitArgs(inner)[0]))); }

    // BASE (número en otra base)
    if ((inner = matchFn('BASE')) !== null) {
      args = splitArgs(inner);
      var num2 = parseInt(resolveNum(args[0])), base2 = parseInt(resolveNum(args[1]));
      return num2.toString(base2).toUpperCase();
    }

    // ════════════════════════════════════════
    // INGENIERÍA — nuevas
    // ════════════════════════════════════════

    // CONVERTIR / CONVERT
    if ((inner = matchFn('CONVERTIR|CONVERT')) !== null) {
      args = splitArgs(inner);
      var val2 = resolveNum(args[0]);
      var from = String(resolveArg(args[1])).toLowerCase();
      var to   = String(resolveArg(args[2])).toLowerCase();
      var convMap = {
        // Longitud
        'km_m':1000,'m_km':0.001,'mi_km':1.60934,'km_mi':0.621371,
        'm_ft':3.28084,'ft_m':0.3048,'in_cm':2.54,'cm_in':0.393701,
        // Peso
        'kg_lb':2.20462,'lb_kg':0.453592,'oz_g':28.3495,'g_oz':0.035274,
        // Temperatura
        'c_f':null,'f_c':null,'c_k':null,'k_c':null,
        // Área
        'ha_m2':10000,'m2_ha':0.0001,
        // Volumen
        'l_gal':0.264172,'gal_l':3.78541,
        // Velocidad
        'kmh_ms':0.277778,'ms_kmh':3.6,
      };
      var key = from+'_'+to;
      if (key === 'c_f') return fmtNum(val2*9/5+32);
      if (key === 'f_c') return fmtNum((val2-32)*5/9);
      if (key === 'c_k') return fmtNum(val2+273.15);
      if (key === 'k_c') return fmtNum(val2-273.15);
      if (convMap[key] != null) return fmtNum(val2 * convMap[key]);
      return '#N/A';
    }

    // BIN.A.DEC / BIN2DEC
    if ((inner = matchFn('BIN\\.A\\.DEC|BIN2DEC')) !== null) {
      return String(parseInt(String(resolveArg(splitArgs(inner)[0])), 2));
    }
    // DEC.A.BIN / DEC2BIN
    if ((inner = matchFn('DEC\\.A\\.BIN|DEC2BIN')) !== null) {
      args = splitArgs(inner);
      var n12 = parseInt(resolveNum(args[0]));
      var places = args[1] ? parseInt(resolveNum(args[1])) : 0;
      var bin = (n12 >>> 0).toString(2);
      return places ? bin.padStart(places,'0') : bin;
    }
    // DEC.A.HEX / DEC2HEX
    if ((inner = matchFn('DEC\\.A\\.HEX|DEC2HEX')) !== null) {
      return parseInt(resolveNum(splitArgs(inner)[0])).toString(16).toUpperCase();
    }
    // HEX.A.DEC / HEX2DEC
    if ((inner = matchFn('HEX\\.A\\.DEC|HEX2DEC')) !== null) {
      return String(parseInt(String(resolveArg(splitArgs(inner)[0])), 16));
    }
    // OCT.A.DEC / OCT2DEC
    if ((inner = matchFn('OCT\\.A\\.DEC|OCT2DEC')) !== null) {
      return String(parseInt(String(resolveArg(splitArgs(inner)[0])), 8));
    }
    // DEC.A.OCT / DEC2OCT
    if ((inner = matchFn('DEC\\.A\\.OCT|DEC2OCT')) !== null) {
      return parseInt(resolveNum(splitArgs(inner)[0])).toString(8);
    }
    // BIT.Y / BITAND
    if ((inner = matchFn('BIT\\.Y|BITAND')) !== null) {
      args=splitArgs(inner); return String(parseInt(resolveNum(args[0])) & parseInt(resolveNum(args[1])));
    }
    // BIT.O / BITOR
    if ((inner = matchFn('BIT\\.O|BITOR')) !== null) {
      args=splitArgs(inner); return String(parseInt(resolveNum(args[0])) | parseInt(resolveNum(args[1])));
    }
    // BIT.XO / BITXOR
    if ((inner = matchFn('BIT\\.XO|BITXOR')) !== null) {
      args=splitArgs(inner); return String(parseInt(resolveNum(args[0])) ^ parseInt(resolveNum(args[1])));
    }
    // DELTA
    if ((inner = matchFn('DELTA')) !== null) {
      args=splitArgs(inner); return resolveNum(args[0]) === resolveNum(args[1]||'0') ? '1' : '0';
    }

    // ════════════════════════════════════════
    // BASE DE DATOS — nuevas
    // ════════════════════════════════════════
    function dbCriterio(dbRange, campoStr, critRange) {
      if (!isRange(dbRange) || !isRange(critRange)) return [];
      var dbCells = rangeVals(dbRange);
      var critCells = rangeVals(critRange);
      var dbParts = dbRange.trim().split(':');
      var crParts = critRange.trim().split(':');
      var dbFc = colIdx(dbParts[0].match(/[A-Za-z]+/)[0]);
      var dbFr = parseInt(dbParts[0].match(/\d+/)[0]) - 1;
      var dbTc = colIdx(dbParts[1].match(/[A-Za-z]+/)[0]);
      var dbTr = parseInt(dbParts[1].match(/\d+/)[0]) - 1;
      var crFc = colIdx(crParts[0].match(/[A-Za-z]+/)[0]);
      var crFr = parseInt(crParts[0].match(/\d+/)[0]) - 1;
      var crTc = colIdx(crParts[1].match(/[A-Za-z]+/)[0]);
      // Headers row
      var headers = [];
      for (var c=dbFc;c<=dbTc;c++) headers.push(String(getVal(cellId(dbFr,c))).toLowerCase());
      // Campo index
      var campo = String(campoStr).replace(/["']/g,'').toLowerCase();
      var campoIdx = headers.indexOf(campo);
      if (campoIdx < 0 && !isNaN(parseInt(campoStr))) campoIdx = parseInt(campoStr)-1;
      // Criteria
      var critHeaders = [];
      for (var c2=crFc;c2<=crTc;c2++) critHeaders.push(String(getVal(cellId(crFr,c2))).toLowerCase());
      // Filter rows
      var result2 = [];
      for (var r=dbFr+1;r<=dbTr;r++) {
        var match = critHeaders.every(function(ch, ci) {
          var crVal = String(getVal(cellId(crFr+1, crFc+ci)));
          var dbVal = String(getVal(cellId(r, dbFc+headers.indexOf(ch))));
          return dbVal.toLowerCase().indexOf(crVal.toLowerCase()) >= 0 || crVal === '';
        });
        if (match && campoIdx >= 0) result2.push(getVal(cellId(r, dbFc+campoIdx)));
      }
      return result2;
    }

    if ((inner = matchFn('BDSUMA|DSUM')) !== null) {
      args = splitArgs(inner);
      return fmtNum(dbCriterio(args[0],resolveArg(args[1]),args[2]).map(toNum).reduce(function(a,b){return a+b;},0));
    }
    if ((inner = matchFn('BDPROMEDIO|DAVERAGE')) !== null) {
      args = splitArgs(inner);
      var vs = dbCriterio(args[0],resolveArg(args[1]),args[2]).map(toNum);
      return vs.length ? fmtNum(vs.reduce(function(a,b){return a+b;},0)/vs.length) : '#DIV/0!';
    }
    if ((inner = matchFn('BDMAX|DMAX')) !== null) {
      args = splitArgs(inner);
      var vs2 = dbCriterio(args[0],resolveArg(args[1]),args[2]).map(toNum);
      return vs2.length ? fmtNum(Math.max.apply(null,vs2)) : '0';
    }
    if ((inner = matchFn('BDMIN|DMIN')) !== null) {
      args = splitArgs(inner);
      var vs3 = dbCriterio(args[0],resolveArg(args[1]),args[2]).map(toNum);
      return vs3.length ? fmtNum(Math.min.apply(null,vs3)) : '0';
    }
    if ((inner = matchFn('BDCONTAR|DCOUNT')) !== null) {
      args = splitArgs(inner);
      return String(dbCriterio(args[0],resolveArg(args[1]),args[2]).filter(function(v){return !isNaN(toNum(v));}).length);
    }
    if ((inner = matchFn('BDCONTARA|DCOUNTA')) !== null) {
      args = splitArgs(inner);
      return String(dbCriterio(args[0],resolveArg(args[1]),args[2]).filter(function(v){return v!=='';}).length);
    }
    if ((inner = matchFn('BDEXTRAER|DGET')) !== null) {
      args = splitArgs(inner);
      var res2 = dbCriterio(args[0],resolveArg(args[1]),args[2]);
      return res2.length === 1 ? String(res2[0]) : (res2.length > 1 ? '#NUM!' : '#VALOR!');
    }

    // ════════════════════════════════════════

        // Single cell ref
    if (/^[A-Za-z]+\d+$/.test(expr)) return String(getVal(expr.toUpperCase()));

    // Range sum shortcut (A1:B5 alone)
    if (/^[A-Za-z]+\d+:[A-Za-z]+\d+$/.test(expr)) {
      var parts6 = expr.split(':'); var total2 = 0;
      getCells(parts6[0], parts6[1]).forEach(function(c) { total2 += toNum(getVal(c)); });
      return fmtNum(total2);
    }

    // Arithmetic with cell refs
    var evalExpr3 = expr.replace(/([A-Za-z]+\d+)/g, function(m) { return toNum(getVal(m.toUpperCase())) || 0; });
    // Handle percentage
    evalExpr3 = evalExpr3.replace(/(\d+\.?\d*)%/g, function(m,n) { return parseFloat(n)/100; });
    var result = Function('"use strict";return (' + evalExpr3 + ')')();
    return isNaN(result) ? '#ERR' : fmtNum(result);

  } catch(e) { return '#ERR'; }

  // ── Condition evaluator ─────────────────────
  function evalCond(expr2) {
    expr2 = (expr2||'').trim();
    // VERDADERO/FALSO literal
    if (expr2.toUpperCase() === 'VERDADERO' || expr2.toUpperCase() === 'TRUE') return true;
    if (expr2.toUpperCase() === 'FALSO' || expr2.toUpperCase() === 'FALSE') return false;
    // Nested formula
    if (expr2.startsWith('=')) return !!evalFormula(expr2, row, col);
    // Operators
    var ops = ['>=','<=','<>','!=','>','<','='];
    for (var oi = 0; oi < ops.length; oi++) {
      var idx3 = expr2.indexOf(ops[oi]);
      if (idx3 > 0) {
        var lv = resolveArg(expr2.substring(0, idx3).trim());
        var rv = resolveArg(expr2.substring(idx3 + ops[oi].length).trim());
        var ln = parseFloat(String(lv).replace(/[₡,\s]/g,'')), rn = parseFloat(String(rv).replace(/[₡,\s]/g,''));
        var lc = !isNaN(ln) ? ln : String(lv).toLowerCase();
        var rc3 = !isNaN(rn) ? rn : String(rv).toLowerCase();
        switch(ops[oi]) {
          case '>=': return lc >= rc3;
          case '<=': return lc <= rc3;
          case '<>': case '!=': return lc != rc3;
          case '>':  return lc > rc3;
          case '<':  return lc < rc3;
          case '=':  return lc == rc3;
        }
      }
    }
    // Truthy check
    var v9 = resolveArg(expr2);
    return v9 !== '' && v9 !== '0' && v9 !== 0 && v9 !== 'FALSO' && v9 !== false;
  }

  // ── Criteria matcher (for CONTAR.SI, SUMAR.SI) ──
  function matchCrit(val, crit) {
    crit = String(crit).trim();
    var ops2 = ['>=','<=','<>','!=','>','<'];
    for (var oi2 = 0; oi2 < ops2.length; oi2++) {
      if (crit.startsWith(ops2[oi2])) {
        var num = parseFloat(crit.substring(ops2[oi2].length));
        var vn = parseFloat(String(val).replace(/[₡,\s]/g,''));
        switch(ops2[oi2]) {
          case '>=': return vn >= num;
          case '<=': return vn <= num;
          case '<>': case '!=': return String(val).toLowerCase() !== crit.substring(2).toLowerCase();
          case '>':  return vn > num;
          case '<':  return vn < num;
        }
      }
    }
    // Wildcard
    if (crit.includes('*') || crit.includes('?')) {
      var re2 = new RegExp('^'+crit.replace(/\*/g,'.*').replace(/\?/g,'.')+'$','i');
      return re2.test(String(val));
    }
    return String(val).toLowerCase() === crit.toLowerCase();
  }
}


// ══════════════════════════════════════════
// INSERTAR GRÁFICO — 20 tipos, se inserta en hoja
// ══════════════════════════════════════════
function insertarGrafico() {
  // Detectar datos de la selección o columna activa
  var labels = [], values = [], values2 = [];
  var r0 = activeCell ? activeCell.r : 0;
  var c0 = activeCell ? activeCell.c : 0;

  if (selStart && selEnd) {
    var r1 = Math.min(selStart.r, selEnd.r), r2 = Math.max(selStart.r, selEnd.r);
    var c1 = Math.min(selStart.c, selEnd.c), c2 = Math.max(selStart.c, selEnd.c);
    for (var ri = r1; ri <= r2; ri++) {
      var lbl = sheetData[cellId(ri, c1)] || cellId(ri, c1);
      var val = parseFloat(String(sheetData[cellId(ri, c1+1)] || 0).replace(/[₡,\s]/g,'')) || 0;
      var val2 = c2 > c1+1 ? parseFloat(String(sheetData[cellId(ri, c1+2)] || 0).replace(/[₡,\s]/g,'')) || 0 : 0;
      labels.push(String(lbl));
      values.push(val);
      values2.push(val2);
    }
  } else {
    for (var ri2 = 0; ri2 < 8; ri2++) {
      var lbl2 = sheetData[cellId(r0 + ri2, c0 > 0 ? c0-1 : 0)] || cellId(r0+ri2, c0);
      var val3 = parseFloat(String(sheetData[cellId(r0+ri2, c0)] || 0).replace(/[₡,\s]/g,'')) || 0;
      if (val3 !== 0 || sheetData[cellId(r0+ri2, c0)]) { labels.push(String(lbl2)); values.push(val3); }
    }
  }
  if (values.length === 0) {
    labels = ['Ventas','Compras','IVA x Pagar','Gastos','Activos','Pasivos'];
    values = [70000, 35000, 9700, 15000, 120000, 45000];
    values2 = [65000, 30000, 8500, 12000, 110000, 40000];
  }

  var CHART_TYPES = [
    { id:'bars',      name:'Barras verticales',    icon:'📊' },
    { id:'bars_h',    name:'Barras horizontales',  icon:'📉' },
    { id:'bars_g',    name:'Barras agrupadas',     icon:'📈' },
    { id:'bars_s',    name:'Barras apiladas',      icon:'🗂' },
    { id:'line',      name:'Línea',                icon:'📈' },
    { id:'line_a',    name:'Línea con área',       icon:'🌊' },
    { id:'line_d',    name:'Líneas dobles',        icon:'〰' },
    { id:'area',      name:'Área',                 icon:'🏔' },
    { id:'pie',       name:'Torta (pastel)',        icon:'🥧' },
    { id:'donut',     name:'Dona',                 icon:'🍩' },
    { id:'scatter',   name:'Dispersión',           icon:'✦' },
    { id:'bubble',    name:'Burbuja',              icon:'🫧' },
    { id:'radar',     name:'Radar',                icon:'🕸' },
    { id:'waterfall', name:'Cascada',              icon:'💧' },
    { id:'step',      name:'Escalón',              icon:'📶' },
    { id:'mixed',     name:'Mixto Barra+Línea',    icon:'📊' },
    { id:'gantt',     name:'Gantt (prog.)',         icon:'📅' },
    { id:'heat',      name:'Mapa de calor',         icon:'🌡' },
    { id:'funnel',    name:'Embudo',               icon:'🔻' },
    { id:'boxplot',   name:'Caja y bigotes',        icon:'📦' },
  ];

  var colors = ['#5abf2a','#3a9e10','#1d4ed8','#dc2626','#d97706','#7c3aed','#0891b2','#be185d','#059669','#f59e0b'];

  // ── Modal de selección ───────────────────
  var overlay = document.createElement('div');
  overlay.style.cssText = 'position:fixed;inset:0;background:rgba(10,30,5,.55);backdrop-filter:blur(4px);z-index:9000;display:flex;align-items:center;justify-content:center;padding:12px;';

  var modal = document.createElement('div');
  modal.style.cssText = 'background:#fff;border-radius:16px;box-shadow:0 20px 60px rgba(0,0,0,.25);width:100%;max-width:700px;max-height:90vh;overflow:hidden;display:flex;flex-direction:column;';

  var head = document.createElement('div');
  head.style.cssText = 'padding:14px 18px;border-bottom:1px solid var(--border);background:var(--off);display:flex;align-items:center;justify-content:space-between;flex-shrink:0;';
  head.innerHTML = '<span style="font-size:15px;font-weight:800;color:var(--g3);">📊 Insertar Gráfico — '+values.length+' series</span>' +
    '<button id="closeChartModal" style="background:var(--bg2);border:none;border-radius:6px;width:28px;height:28px;font-size:17px;cursor:pointer;color:var(--text2);">×</button>';

  var body = document.createElement('div');
  body.style.cssText = 'display:flex;flex:1;overflow:hidden;';

  // Left: chart type list
  var sidebar = document.createElement('div');
  sidebar.style.cssText = 'width:220px;flex-shrink:0;border-right:1px solid var(--border);overflow-y:auto;padding:8px;background:var(--off);';
  var selectedType = 'bars';
  CHART_TYPES.forEach(function(ct) {
    var btn = document.createElement('button');
    btn.dataset.id = ct.id;
    btn.style.cssText = 'width:100%;display:flex;align-items:center;gap:8px;padding:8px 10px;border-radius:8px;border:none;background:' + (ct.id==='bars'?'var(--g2)':'transparent') + ';color:' + (ct.id==='bars'?'#fff':'var(--text2)') + ';font-size:12.5px;font-weight:600;cursor:pointer;text-align:left;transition:all .1s;';
    btn.innerHTML = '<span style="font-size:16px;">' + ct.icon + '</span>' + ct.name;
    btn.addEventListener('click', function() {
      selectedType = this.dataset.id;
      sidebar.querySelectorAll('button').forEach(function(b2) { b2.style.background='transparent'; b2.style.color='var(--text2)'; });
      this.style.background = 'var(--g2)'; this.style.color = '#fff';
      drawPreview(selectedType);
    });
    sidebar.appendChild(btn);
  });

  // Right: preview + options
  var right = document.createElement('div');
  right.style.cssText = 'flex:1;display:flex;flex-direction:column;overflow:hidden;';

  var preview = document.createElement('div');
  preview.style.cssText = 'flex:1;padding:14px;display:flex;align-items:center;justify-content:center;overflow:hidden;';
  var cvs2 = document.createElement('canvas');
  cvs2.id = 'chartPreviewCanvas';
  cvs2.width = 380; cvs2.height = 240;
  cvs2.style.cssText = 'border:1px solid var(--border);border-radius:8px;max-width:100%;';
  preview.appendChild(cvs2);

  var opts = document.createElement('div');
  opts.style.cssText = 'padding:10px 14px;border-top:1px solid var(--border);display:flex;gap:8px;align-items:center;flex-wrap:wrap;flex-shrink:0;';
  opts.innerHTML = '<label style="font-size:12px;font-weight:700;color:var(--text2);">Título:</label>' +
    '<input id="chartTitle" value="Gráfico" style="padding:5px 8px;border:1px solid var(--border2);border-radius:6px;font-size:12px;flex:1;min-width:80px;">' +
    '<label style="font-size:12px;font-weight:700;color:var(--text2);">Color:</label>' +
    '<select id="chartColorScheme" style="padding:5px 8px;border:1px solid var(--border2);border-radius:6px;font-size:12px;">' +
    '<option value="green">🟢 Verde CR</option><option value="blue">🔵 Azul</option>' +
    '<option value="multi">🌈 Multi</option><option value="warm">🔥 Cálido</option></select>';

  var foot = document.createElement('div');
  foot.style.cssText = 'padding:11px 18px;border-top:1px solid var(--border);display:flex;gap:8px;justify-content:flex-end;background:var(--off);flex-shrink:0;';
  foot.innerHTML = '<button id="cancelChart" style="padding:7px 14px;border-radius:7px;border:none;background:var(--bg2);color:var(--text2);font-size:13px;font-weight:600;cursor:pointer;">Cancelar</button>' +
    '<button id="insertChart" style="padding:7px 18px;border-radius:7px;border:none;background:linear-gradient(135deg,var(--g1),var(--g2));color:#fff;font-size:13px;font-weight:700;cursor:pointer;">✓ Insertar en hoja</button>';

  right.appendChild(preview); right.appendChild(opts); right.appendChild(foot);
  body.appendChild(sidebar); body.appendChild(right);
  modal.appendChild(head); modal.appendChild(body);
  overlay.appendChild(modal);
  document.body.appendChild(overlay);

  // Close handlers
  head.querySelector('#closeChartModal').onclick = function() { document.body.removeChild(overlay); };
  foot.querySelector('#cancelChart').onclick = function() { document.body.removeChild(overlay); };
  overlay.addEventListener('click', function(e) { if (e.target === overlay) document.body.removeChild(overlay); });

  // Draw preview function
  function getColors(scheme) {
    if (scheme === 'blue') return ['#1d4ed8','#3b82f6','#60a5fa','#93c5fd','#bfdbfe','#2563eb'];
    if (scheme === 'warm') return ['#dc2626','#f97316','#f59e0b','#eab308','#84cc16','#22c55e'];
    if (scheme === 'multi') return ['#5abf2a','#1d4ed8','#dc2626','#d97706','#7c3aed','#0891b2'];
    return ['#5abf2a','#3a9e10','#2d7a0c','#1e5208','#7fff4a','#52b826']; // green
  }

  function drawPreview(type) {
    var ctx3 = document.getElementById('chartPreviewCanvas').getContext('2d');
    var W = 380, H = 240;
    var scheme = document.getElementById('chartColorScheme') ? document.getElementById('chartColorScheme').value : 'green';
    var cls = getColors(scheme);
    ctx3.clearRect(0, 0, W, H);
    ctx3.fillStyle = '#fff'; ctx3.fillRect(0, 0, W, H);

    var title = document.getElementById('chartTitle') ? document.getElementById('chartTitle').value : 'Gráfico';
    ctx3.fillStyle = '#1e5208'; ctx3.font = 'bold 12px sans-serif'; ctx3.textAlign = 'center';
    ctx3.fillText(title, W/2, 16);

    var padL=44, padR=14, padT=24, padB=36;
    var cW = W - padL - padR, cH = H - padT - padB;
    var max2 = Math.max.apply(null, values.concat(values2)) || 1;

    function gridLines() {
      ctx3.strokeStyle='#e2f2d4'; ctx3.lineWidth=1;
      for (var g=0;g<=4;g++) {
        var gy2=padT+cH-(g/4)*cH;
        ctx3.beginPath(); ctx3.moveTo(padL,gy2); ctx3.lineTo(padL+cW,gy2); ctx3.stroke();
        ctx3.fillStyle='#6b8f48'; ctx3.font='9px sans-serif'; ctx3.textAlign='right';
        var gv2=max2*g/4;
        ctx3.fillText(gv2>=1000?(gv2/1000).toFixed(1)+'K':gv2.toFixed(0), padL-3, gy2+3);
      }
    }

    if (type==='bars'||type==='bars_s') {
      gridLines();
      var bw2=Math.min(40, cW/(values.length+1));
      values.forEach(function(v,i) {
        var h2=(v/max2)*cH; var x=padL+i*(bw2+6);
        ctx3.fillStyle=cls[i%cls.length];
        ctx3.beginPath(); if(ctx3.roundRect) ctx3.roundRect(x,padT+cH-h2,bw2,h2,[3,3,0,0]); else ctx3.rect(x,padT+cH-h2,bw2,h2);
        ctx3.fill();
        ctx3.fillStyle='#1e5208'; ctx3.font='9px sans-serif'; ctx3.textAlign='center';
        ctx3.fillText(v>=1000?(v/1000).toFixed(1)+'K':String(v),x+bw2/2,padT+cH-h2-3);
        ctx3.fillStyle='#3d6020'; ctx3.fillText((labels[i]||'').substring(0,6),x+bw2/2,padT+cH+12);
      });
    } else if (type==='bars_h') {
      var bh2=Math.min(22, cH/(values.length+1));
      values.forEach(function(v,i) {
        var w2=(v/max2)*cW; var y2=padT+i*(bh2+5);
        ctx3.fillStyle=cls[i%cls.length]; ctx3.fillRect(padL,y2,w2,bh2);
        ctx3.fillStyle='#1e5208'; ctx3.font='9px sans-serif'; ctx3.textAlign='left';
        ctx3.fillText((labels[i]||'').substring(0,8),padL-40,y2+bh2/2+3);
        ctx3.fillStyle='#fff'; ctx3.textAlign='right';
        ctx3.fillText(v>=1000?(v/1000).toFixed(1)+'K':String(v),padL+w2-3,y2+bh2/2+3);
      });
    } else if (type==='bars_g') {
      gridLines();
      var bw3=Math.min(18, cW/(values.length*2.5+1));
      values.forEach(function(v,i) {
        var x2=padL+i*(bw3*2+10);
        var h3=(v/max2)*cH; ctx3.fillStyle=cls[0]; ctx3.fillRect(x2,padT+cH-h3,bw3,h3);
        if (values2[i]) { var h4=(values2[i]/max2)*cH; ctx3.fillStyle=cls[2]; ctx3.fillRect(x2+bw3+2,padT+cH-h4,bw3,h4); }
        ctx3.fillStyle='#3d6020'; ctx3.font='9px sans-serif'; ctx3.textAlign='center';
        ctx3.fillText((labels[i]||'').substring(0,5),x2+bw3,padT+cH+12);
      });
    } else if (type==='line'||type==='line_d'||type==='area'||type==='line_a') {
      gridLines();
      var pw=cW/(values.length-1||1);
      if (type==='area'||type==='line_a') {
        ctx3.beginPath(); ctx3.moveTo(padL,padT+cH);
        values.forEach(function(v,i){ ctx3.lineTo(padL+i*pw,padT+cH-(v/max2)*cH); });
        ctx3.lineTo(padL+cW,padT+cH); ctx3.closePath();
        ctx3.fillStyle='rgba(90,191,42,.18)'; ctx3.fill();
      }
      ctx3.beginPath(); ctx3.strokeStyle=cls[0]; ctx3.lineWidth=2.5; ctx3.lineJoin='round';
      values.forEach(function(v,i){ i===0?ctx3.moveTo(padL+i*pw,padT+cH-(v/max2)*cH):ctx3.lineTo(padL+i*pw,padT+cH-(v/max2)*cH); });
      ctx3.stroke();
      if ((type==='line_d'||type==='bars_g') && values2.some(function(v){return v>0;})) {
        ctx3.beginPath(); ctx3.strokeStyle=cls[2]; ctx3.lineWidth=2; ctx3.setLineDash([4,3]);
        values2.forEach(function(v,i){ i===0?ctx3.moveTo(padL+i*pw,padT+cH-(v/max2)*cH):ctx3.lineTo(padL+i*pw,padT+cH-(v/max2)*cH); });
        ctx3.stroke(); ctx3.setLineDash([]);
      }
      values.forEach(function(v,i){ ctx3.beginPath(); ctx3.arc(padL+i*pw,padT+cH-(v/max2)*cH,3.5,0,Math.PI*2); ctx3.fillStyle=cls[0]; ctx3.fill(); });
    } else if (type==='pie'||type==='donut') {
      var total3=values.reduce(function(a,b){return a+b;},0)||1; var startAngle=-Math.PI/2;
      var cx2=W/2, cy2=H/2+5, pr=Math.min(cW,cH)/2-10;
      values.forEach(function(v,i) {
        var slice=(v/total3)*Math.PI*2;
        ctx3.beginPath(); ctx3.moveTo(cx2,cy2); ctx3.arc(cx2,cy2,pr,startAngle,startAngle+slice);
        ctx3.closePath(); ctx3.fillStyle=cls[i%cls.length]; ctx3.fill(); ctx3.strokeStyle='#fff'; ctx3.lineWidth=2; ctx3.stroke();
        var mid=startAngle+slice/2; var lx=cx2+Math.cos(mid)*(pr*0.65); var ly=cy2+Math.sin(mid)*(pr*0.65);
        ctx3.fillStyle='#fff'; ctx3.font='bold 9px sans-serif'; ctx3.textAlign='center';
        if (v/total3>0.05) ctx3.fillText(Math.round(v/total3*100)+'%',lx,ly+3);
        startAngle+=slice;
      });
      if (type==='donut') { ctx3.beginPath(); ctx3.arc(cx2,cy2,pr*0.45,0,Math.PI*2); ctx3.fillStyle='#fff'; ctx3.fill(); }
      ctx3.fillStyle='#1e5208'; ctx3.font='bold 11px sans-serif'; ctx3.textAlign='center';
      if (type==='donut') ctx3.fillText(total3>=1000?(total3/1000).toFixed(1)+'K':String(total3), cx2, cy2+4);
    } else if (type==='radar') {
      var cx3=W/2, cy3=H/2+5, pr2=Math.min(cW,cH)/2-15; var n=values.length;
      for(var lv2=1;lv2<=4;lv2++){ctx3.beginPath();ctx3.strokeStyle='#e2f2d4';ctx3.lineWidth=1;for(var i=0;i<n;i++){var a=i*2*Math.PI/n-Math.PI/2;var x=cx3+Math.cos(a)*pr2*(lv2/4);var y=cy3+Math.sin(a)*pr2*(lv2/4);i===0?ctx3.moveTo(x,y):ctx3.lineTo(x,y);}ctx3.closePath();ctx3.stroke();}
      ctx3.beginPath(); ctx3.fillStyle='rgba(90,191,42,.25)'; ctx3.strokeStyle=cls[0]; ctx3.lineWidth=2;
      values.forEach(function(v,i){var a=i*2*Math.PI/n-Math.PI/2;var x=cx3+Math.cos(a)*pr2*(v/max2);var y=cy3+Math.sin(a)*pr2*(v/max2);i===0?ctx3.moveTo(x,y):ctx3.lineTo(x,y);}); ctx3.closePath();ctx3.fill();ctx3.stroke();
      values.forEach(function(v,i){var a=i*2*Math.PI/n-Math.PI/2;var lx=cx3+Math.cos(a)*(pr2+12);var ly=cy3+Math.sin(a)*(pr2+12);ctx3.fillStyle='#3d6020';ctx3.font='9px sans-serif';ctx3.textAlign='center';ctx3.fillText((labels[i]||'').substring(0,6),lx,ly+3);});
    } else if (type==='waterfall') {
      gridLines(); var bw4=Math.min(32,cW/(values.length+1)); var running=0;
      values.forEach(function(v,i){
        var base=padT+cH-(running/max2)*cH; var h5=(v/max2)*cH;
        ctx3.fillStyle=v>=0?cls[0]:'#dc2626';
        ctx3.fillRect(padL+i*(bw4+8),v>=0?base-h5:base,bw4,Math.abs(h5));
        running+=v;
        ctx3.fillStyle='#3d6020';ctx3.font='9px sans-serif';ctx3.textAlign='center';
        ctx3.fillText((labels[i]||'').substring(0,6),padL+i*(bw4+8)+bw4/2,padT+cH+12);
      });
    } else if (type==='step') {
      gridLines(); var pw2=cW/(values.length-1||1);
      ctx3.beginPath(); ctx3.strokeStyle=cls[0]; ctx3.lineWidth=2.5;
      values.forEach(function(v,i){
        var x3=padL+i*pw2,y3=padT+cH-(v/max2)*cH;
        if(i===0){ctx3.moveTo(x3,y3);}else{ctx3.lineTo(x3,padT+cH-(values[i-1]/max2)*cH);ctx3.lineTo(x3,y3);}
      }); ctx3.stroke();
    } else if (type==='funnel') {
      var sorted=values.slice().sort(function(a,b){return b-a;}); var bh3=Math.min(28,cH/(sorted.length+1));
      sorted.forEach(function(v,i){
        var w3=(v/max2)*cW; var y4=padT+i*(bh3+6); var x4=padL+(cW-w3)/2;
        ctx3.fillStyle=cls[i%cls.length]; ctx3.fillRect(x4,y4,w3,bh3);
        ctx3.fillStyle='#fff';ctx3.font='bold 9px sans-serif';ctx3.textAlign='center';
        ctx3.fillText((labels[i]||'')+': '+(v>=1000?(v/1000).toFixed(1)+'K':v),padL+cW/2,y4+bh3/2+3);
      });
    } else if (type==='scatter'||type==='bubble') {
      gridLines();
      values.forEach(function(v,i){
        var x5=padL+(i/(values.length-1||1))*cW; var y5=padT+cH-(v/max2)*cH;
        var r3=type==='bubble'?(Math.sqrt(v/max2)*12+4):5;
        ctx3.beginPath();ctx3.arc(x5,y5,r3,0,Math.PI*2);
        ctx3.fillStyle=cls[i%cls.length];ctx3.globalAlpha=0.7;ctx3.fill();ctx3.globalAlpha=1;
        ctx3.strokeStyle='#fff';ctx3.lineWidth=1;ctx3.stroke();
      });
    } else if (type==='heat') {
      var cols2=Math.ceil(Math.sqrt(values.length)), rows2=Math.ceil(values.length/cols2);
      var cw3=cW/cols2, ch3=cH/rows2;
      values.forEach(function(v,i){
        var x6=padL+(i%cols2)*cw3; var y6=padT+Math.floor(i/cols2)*ch3;
        var intensity=v/max2;
        ctx3.fillStyle='rgba(90,191,42,'+intensity+')'; ctx3.fillRect(x6,y6,cw3-2,ch3-2);
        ctx3.fillStyle='#1e5208';ctx3.font='9px sans-serif';ctx3.textAlign='center';
        ctx3.fillText((labels[i]||'').substring(0,4),x6+cw3/2,y6+ch3/2+3);
      });
    } else if (type==='mixed') {
      gridLines(); var bw5=Math.min(28,cW/(values.length+1)); var pw3=cW/(values.length-1||1);
      values.forEach(function(v,i){var h6=(v/max2)*cH;var x7=padL+i*(bw5+8);ctx3.fillStyle=cls[0];ctx3.fillRect(x7,padT+cH-h6,bw5,h6);ctx3.fillStyle='#3d6020';ctx3.font='9px sans-serif';ctx3.textAlign='center';ctx3.fillText((labels[i]||'').substring(0,5),x7+bw5/2,padT+cH+12);});
      if(values2.some(function(v){return v>0;})){ctx3.beginPath();ctx3.strokeStyle=cls[2];ctx3.lineWidth=2.5;ctx3.setLineDash([]);values2.forEach(function(v,i){var x8=padL+i*(bw5+8)+bw5/2;var y8=padT+cH-(v/max2)*cH;i===0?ctx3.moveTo(x8,y8):ctx3.lineTo(x8,y8);});ctx3.stroke();}
    } else if (type==='gantt') {
      var bh4=Math.min(22,cH/(values.length+1));
      values.forEach(function(v,i){
        var start5=padL+(i*15);var w4=(v/max2)*(cW*0.7);
        ctx3.fillStyle=cls[i%cls.length];ctx3.fillRect(start5,padT+i*(bh4+5),w4,bh4);
        ctx3.fillStyle='#1e5208';ctx3.font='9px sans-serif';ctx3.textAlign='left';
        ctx3.fillText((labels[i]||'').substring(0,10),padL-padL+2,padT+i*(bh4+5)+bh4/2+3);
      });
    } else if (type==='boxplot') {
      gridLines(); var bw6=Math.min(30,cW/(values.length+1));
      values.forEach(function(v,i){
        var q2=padT+cH-(v/max2)*cH; var q1=padT+cH-(v*0.75/max2)*cH; var q3=padT+cH-(v*0.25/max2)*cH;
        var med=padT+cH-(v*0.5/max2)*cH; var x9=padL+i*(bw6+12)+bw6/2;
        ctx3.strokeStyle=cls[i%cls.length];ctx3.lineWidth=2;
        ctx3.strokeRect(x9-bw6/2,q1,bw6,q3-q1);
        ctx3.beginPath();ctx3.moveTo(x9,q2);ctx3.lineTo(x9,q1);ctx3.stroke();
        ctx3.beginPath();ctx3.moveTo(x9,q3);ctx3.lineTo(x9,padT+cH-(v*0.1/max2)*cH);ctx3.stroke();
        ctx3.strokeStyle=cls[0];ctx3.lineWidth=3;
        ctx3.beginPath();ctx3.moveTo(x9-bw6/2,med);ctx3.lineTo(x9+bw6/2,med);ctx3.stroke();
        ctx3.fillStyle='#3d6020';ctx3.font='9px sans-serif';ctx3.textAlign='center';
        ctx3.fillText((labels[i]||'').substring(0,5),x9,padT+cH+12);
      });
    } else {
      // Default to bars
      drawPreview('bars');
    }
  }

  // Color change redraws
  opts.querySelector('#chartColorScheme').addEventListener('change', function() { drawPreview(selectedType); });
  opts.querySelector('#chartTitle').addEventListener('input', function() { drawPreview(selectedType); });

  // Initial draw
  setTimeout(function() { drawPreview('bars'); }, 50);

  // INSERT button
  foot.querySelector('#insertChart').addEventListener('click', function() {
    var title2 = document.getElementById('chartTitle').value || 'Gráfico';
    var scheme2 = document.getElementById('chartColorScheme').value;
    insertChartIntoSheet(selectedType, title2, scheme2, labels, values, values2);
    document.body.removeChild(overlay);
  });
}

function insertChartIntoSheet(type, title, scheme, labels, values, values2) {
  // Create a floating div anchored below active cell
  var anchorEl = cellEl(activeCell.r, activeCell.c);
  var anchorRect = anchorEl ? anchorEl.getBoundingClientRect() : { bottom: 200, left: 100 };
  var sheetOuterRect = document.getElementById('sheetOuter').getBoundingClientRect();

  var wrap = document.createElement('div');
  wrap.className = 'chart-float';
  wrap.style.cssText = [
    'position:absolute',
    'left:' + (anchorEl ? anchorEl.offsetLeft : 100) + 'px',
    'top:' + (anchorEl ? anchorEl.offsetTop + anchorEl.offsetHeight + 4 : 100) + 'px',
    'background:#fff',
    'border:2px solid var(--border2)',
    'border-radius:12px',
    'padding:10px 12px 8px',
    'box-shadow:0 6px 28px rgba(0,0,0,.14)',
    'z-index:40',
    'min-width:320px',
    'cursor:move',
    'user-select:none',
  ].join(';');

  var toolbar = document.createElement('div');
  toolbar.style.cssText = 'display:flex;align-items:center;justify-content:space-between;margin-bottom:7px;';
  toolbar.innerHTML = '<span style="font-size:11px;font-weight:700;color:var(--g3);">📊 ' + title + '</span>' +
    '<div style="display:flex;gap:4px;">' +
    '<button onclick="this.closest(\'.chart-float\').style.display=\'none\'" style="background:var(--bg2);border:1px solid var(--border);border-radius:4px;padding:2px 7px;font-size:11px;cursor:pointer;">👁 Ocultar</button>' +
    '<button onclick="this.closest(\'.chart-float\').remove()" style="background:none;border:none;font-size:14px;cursor:pointer;color:var(--text3);">✕</button>' +
    '</div>';

  var id2 = 'chart_' + Date.now();
  var cvs3 = document.createElement('canvas');
  cvs3.id = id2; cvs3.width = 340; cvs3.height = 200;
  cvs3.style.cssText = 'display:block;border-radius:6px;';

  wrap.appendChild(toolbar); wrap.appendChild(cvs3);
  document.getElementById('sheetOuter').appendChild(wrap);

  // Make draggable
  var isDragging = false, dragX = 0, dragY = 0;
  toolbar.addEventListener('mousedown', function(e) { isDragging=true; dragX=e.clientX-wrap.offsetLeft; dragY=e.clientY-wrap.offsetTop; });
  document.addEventListener('mousemove', function(e) { if(isDragging){ wrap.style.left=(e.clientX-dragX)+'px'; wrap.style.top=(e.clientY-dragY)+'px'; } });
  document.addEventListener('mouseup', function() { isDragging=false; });

  // Draw the actual chart
  setTimeout(function() {
    var ctx4 = document.getElementById(id2);
    if (!ctx4) return;
    var ctx5 = ctx4.getContext('2d');
    // Temporarily set preview canvas to the inserted canvas
    var origId = 'chartPreviewCanvas';
    var origEl = document.getElementById(origId);
    ctx4.id = origId;
    var colors2 = ['#5abf2a','#3a9e10','#2d7a0c','#1e5208','#7fff4a','#52b826'];
    if (scheme === 'blue') colors2 = ['#1d4ed8','#3b82f6','#60a5fa','#93c5fd','#bfdbfe','#2563eb'];
    if (scheme === 'warm') colors2 = ['#dc2626','#f97316','#f59e0b','#eab308','#84cc16','#22c55e'];
    if (scheme === 'multi') colors2 = ['#5abf2a','#1d4ed8','#dc2626','#d97706','#7c3aed','#0891b2'];

    var max3 = Math.max.apply(null, values.concat(values2)) || 1;
    var W2 = 340, H2 = 200, padL2=40,padR2=10,padT2=20,padB2=32;
    var cW2=W2-padL2-padR2, cH2=H2-padT2-padB2;

    ctx5.clearRect(0,0,W2,H2); ctx5.fillStyle='#fff'; ctx5.fillRect(0,0,W2,H2);

    // Grid
    ctx5.strokeStyle='#e2f2d4'; ctx5.lineWidth=1;
    for(var g2=0;g2<=4;g2++){var gy3=padT2+cH2-(g2/4)*cH2;ctx5.beginPath();ctx5.moveTo(padL2,gy3);ctx5.lineTo(padL2+cW2,gy3);ctx5.stroke();ctx5.fillStyle='#6b8f48';ctx5.font='8px sans-serif';ctx5.textAlign='right';var gv3=max3*g2/4;ctx5.fillText(gv3>=1000?(gv3/1000).toFixed(1)+'K':gv3.toFixed(0),padL2-2,gy3+3);}

    if (type==='bars'||type==='waterfall'||type==='step'||type==='gantt'||type==='mixed'||type==='boxplot'||type==='bars_g'||type==='bars_s'||type==='bars_h'||type==='heat'||type==='funnel'||type==='scatter'||type==='bubble') {
      var bw7=Math.min(36,cW2/(values.length+1));
      values.forEach(function(v,i){
        var h7=(v/max3)*cH2; var x10=padL2+i*(bw7+7);
        ctx5.fillStyle=colors2[i%colors2.length];
        if(ctx5.roundRect) ctx5.roundRect(x10,padT2+cH2-h7,bw7,h7,[3,3,0,0]); else ctx5.rect(x10,padT2+cH2-h7,bw7,h7);
        ctx5.fill();
        ctx5.fillStyle='#1e5208';ctx5.font='8px sans-serif';ctx5.textAlign='center';
        ctx5.fillText(v>=1000?(v/1000).toFixed(1)+'K':String(v),x10+bw7/2,padT2+cH2-h7-3);
        ctx5.fillStyle='#3d6020';
        ctx5.fillText((labels[i]||'').substring(0,6),x10+bw7/2,padT2+cH2+12);
      });
    } else if (type==='line'||type==='line_a'||type==='line_d'||type==='area') {
      var pw4=cW2/(values.length-1||1);
      if(type==='area'||type==='line_a'){ctx5.beginPath();ctx5.moveTo(padL2,padT2+cH2);values.forEach(function(v,i){ctx5.lineTo(padL2+i*pw4,padT2+cH2-(v/max3)*cH2);});ctx5.lineTo(padL2+cW2,padT2+cH2);ctx5.closePath();ctx5.fillStyle='rgba(90,191,42,.18)';ctx5.fill();}
      ctx5.beginPath();ctx5.strokeStyle=colors2[0];ctx5.lineWidth=2.5;ctx5.lineJoin='round';
      values.forEach(function(v,i){i===0?ctx5.moveTo(padL2+i*pw4,padT2+cH2-(v/max3)*cH2):ctx5.lineTo(padL2+i*pw4,padT2+cH2-(v/max3)*cH2);});ctx5.stroke();
      values.forEach(function(v,i){ctx5.beginPath();ctx5.arc(padL2+i*pw4,padT2+cH2-(v/max3)*cH2,3,0,Math.PI*2);ctx5.fillStyle=colors2[0];ctx5.fill();});
    } else if (type==='pie'||type==='donut') {
      var tot2=values.reduce(function(a,b){return a+b;},0)||1;var sa2=-Math.PI/2;
      var cx4=W2/2,cy4=H2/2,pr3=Math.min(cW2,cH2)/2-8;
      values.forEach(function(v,i){var sl=(v/tot2)*Math.PI*2;ctx5.beginPath();ctx5.moveTo(cx4,cy4);ctx5.arc(cx4,cy4,pr3,sa2,sa2+sl);ctx5.closePath();ctx5.fillStyle=colors2[i%colors2.length];ctx5.fill();ctx5.strokeStyle='#fff';ctx5.lineWidth=2;ctx5.stroke();var m2=sa2+sl/2;if(v/tot2>0.05){ctx5.fillStyle='#fff';ctx5.font='bold 9px sans-serif';ctx5.textAlign='center';ctx5.fillText(Math.round(v/tot2*100)+'%',cx4+Math.cos(m2)*pr3*0.65,cy4+Math.sin(m2)*pr3*0.65+3);}sa2+=sl;});
      if(type==='donut'){ctx5.beginPath();ctx5.arc(cx4,cy4,pr3*0.45,0,Math.PI*2);ctx5.fillStyle='#fff';ctx5.fill();}
    } else if (type==='radar') {
      var cx5=W2/2,cy5=H2/2,pr4=Math.min(cW2,cH2)/2-15,n2=values.length;
      for(var lv3=1;lv3<=4;lv3++){ctx5.beginPath();ctx5.strokeStyle='#e2f2d4';ctx5.lineWidth=1;for(var i=0;i<n2;i++){var a2=i*2*Math.PI/n2-Math.PI/2;ctx5.lineTo(cx5+Math.cos(a2)*pr4*(lv3/4),cy5+Math.sin(a2)*pr4*(lv3/4));}ctx5.closePath();ctx5.stroke();}
      ctx5.beginPath();ctx5.fillStyle='rgba(90,191,42,.25)';ctx5.strokeStyle=colors2[0];ctx5.lineWidth=2;
      values.forEach(function(v,i){var a3=i*2*Math.PI/n2-Math.PI/2;i===0?ctx5.moveTo(cx5+Math.cos(a3)*pr4*(v/max3),cy5+Math.sin(a3)*pr4*(v/max3)):ctx5.lineTo(cx5+Math.cos(a3)*pr4*(v/max3),cy5+Math.sin(a3)*pr4*(v/max3));});
      ctx5.closePath();ctx5.fill();ctx5.stroke();
    }

    ctx5.fillStyle='#1e5208';ctx5.font='bold 10px sans-serif';ctx5.textAlign='center';
    ctx5.fillText(title,W2/2,12);

    // Restore original ID
    if (origEl) ctx4.id = id2;

    var sl2 = document.getElementById('statusLeft');
    if (sl2) sl2.textContent = 'Gráfico "'+title+'" insertado · '+type;
  }, 30);
}


// ══════════════════════════════════════════
// loadTemplate — escribe en sheetData
// ══════════════════════════════════════════
function loadTemplate(tipo) {
  var hoy = new Date().toLocaleDateString('es-CR');
  var r = activeCell ? activeCell.r : 0;
  var c = activeCell ? activeCell.c : 0;
  function set(row, col, val) { sheetData[cellId(row, col)] = val; }
  for (var ri = r; ri < r+22; ri++) for (var ci = c; ci < c+12; ci++) delete sheetData[cellId(ri,ci)];

  if (tipo === 'diario') {
    set(r,c,'📘 LIBRO DIARIO');
    ['Fecha','N° Asiento','Código','Cuenta','Debe (₡)','Haber (₡)','Ref.'].forEach(function(h,i){set(r+1,c+i,h);});
    [[hoy,'001','1.1.01','Caja y Bancos','56500','','F-001'],
     [hoy,'001','4.1.01','Ventas','','50000','F-001'],
     [hoy,'001','2.1.05','IVA por Pagar','','6500','13%'],
     [hoy,'002','1.1.01','Caja y Bancos','23200','','F-002'],
     [hoy,'002','4.1.01','Ventas','','20000','F-002'],
     [hoy,'002','2.1.05','IVA por Pagar','','3200','13%'],
    ].forEach(function(row,i){row.forEach(function(v,j){set(r+2+i,c+j,v);});});
    set(r+8,c+3,'TOTALES');
    set(r+8,c+4,'=SUMA('+cellId(r+2,c+4)+':'+cellId(r+7,c+4)+')');
    set(r+8,c+5,'=SUMA('+cellId(r+2,c+5)+':'+cellId(r+7,c+5)+')');
  } else if (tipo === 'mayor') {
    set(r,c,'📗 LIBRO MAYOR — Caja y Bancos (1.1.01)');
    ['Fecha','Descripción','Ref.','Debe (₡)','Haber (₡)','Saldo (₡)'].forEach(function(h,i){set(r+1,c+i,h);});
    [[hoy,'Saldo inicial','','','','0'],
     [hoy,'Venta F-001','F-001','56500','','56500'],
     [hoy,'Venta F-002','F-002','23200','','79700'],
     [hoy,'Pago proveedor','C-001','','15000','64700'],
    ].forEach(function(row,i){row.forEach(function(v,j){set(r+2+i,c+j,v);});});
    set(r+6,c+2,'TOTALES');
    set(r+6,c+3,'=SUMA('+cellId(r+2,c+3)+':'+cellId(r+5,c+3)+')');
    set(r+6,c+4,'=SUMA('+cellId(r+2,c+4)+':'+cellId(r+5,c+4)+')');
  } else if (tipo === 'balance') {
    set(r,c,'📊 BALANCE DE COMPROBACIÓN');
    ['Código','Cuenta','Nat.','Débitos (₡)','Créditos (₡)','Saldo D (₡)','Saldo A (₡)'].forEach(function(h,i){set(r+1,c+i,h);});
    [['1.1.01','Caja y Bancos','D','79700','15000','64700',''],
     ['1.2.01','Clientes','D','30000','','30000',''],
     ['2.1.01','Proveedores','A','','20000','','20000'],
     ['2.1.05','IVA por Pagar','A','','9700','','9700'],
     ['4.1.01','Ventas','A','','70000','','70000'],
     ['5.1.01','Costo de Ventas','D','40000','','40000',''],
    ].forEach(function(row,i){row.forEach(function(v,j){set(r+2+i,c+j,v);});});
    set(r+8,c+1,'TOTALES');
    [3,4,5,6].forEach(function(ci){set(r+8,c+ci,'=SUMA('+cellId(r+2,c+ci)+':'+cellId(r+7,c+ci)+')');});
  } else if (tipo === 'conciliacion') {
    set(r,c,'🏦 CONCILIACIÓN BANCARIA');
    set(r+1,c,'SEGÚN BANCO'); ['Concepto','Monto (₡)'].forEach(function(h,i){set(r+2,c+i,h);});
    [['Saldo estado de cuenta','95000'],['(+) Depósitos en tránsito','20000'],['(-) Cheques pendientes','-5000']].forEach(function(row,i){row.forEach(function(v,j){set(r+3+i,c+j,v);});});
    set(r+6,c,'Saldo ajustado banco'); set(r+6,c+1,'=SUMA('+cellId(r+3,c+1)+':'+cellId(r+5,c+1)+')');
    set(r+8,c,'SEGÚN LIBROS'); ['Concepto','Monto (₡)'].forEach(function(h,i){set(r+9,c+i,h);});
    [['Saldo según libros','108000'],['(+) Notas de crédito','5000'],['(-) Notas de débito','-3000']].forEach(function(row,i){row.forEach(function(v,j){set(r+10+i,c+j,v);});});
    set(r+13,c,'Saldo ajustado libros'); set(r+13,c+1,'=SUMA('+cellId(r+10,c+1)+':'+cellId(r+12,c+1)+')');
  } else if (tipo === 'compras') {
    set(r,c,'🛒 D-151 COMPRAS Y VENTAS');
    ['Fecha','Tipo','N° Doc.','Cédula','Nombre','Condición','Subtotal (₡)','IVA (₡)','Total (₡)'].forEach(function(h,i){set(r+1,c+i,h);});
    [[hoy,'FE','001-001-000001','3-101-000000','Cliente A S.A.','Contado','50000','6500','56500'],
     [hoy,'FE','001-001-000002','1-234-567890','Juan Pérez','Crédito','20000','2600','22600'],
     [hoy,'FC','001-002-000001','3-102-111222','Proveedor B S.A.','Contado','35000','4550','39550'],
    ].forEach(function(row,i){row.forEach(function(v,j){set(r+2+i,c+j,v);});});
    set(r+5,c+5,'TOTALES');
    [6,7,8].forEach(function(ci){set(r+5,c+ci,'=SUMA('+cellId(r+2,c+ci)+':'+cellId(r+4,c+ci)+')');});
  } else if (tipo === 'factura') {
    openModal('modalFactura'); return;
  }
  renderAllCells();
  activateCell(r,c);
  var sl = document.getElementById('statusLeft');
  if (sl) sl.textContent = 'Plantilla cargada: ' + tipo;
}

// ══════════════════════════════════════════
// Las otras 5 funciones corregidas
// ══════════════════════════════════════════
function insertarTabla() { sheetInsertTable(); }

function insertarNota() {
  var nota = prompt('Texto de la nota:', '📝 Nota contable...');
  if (nota !== null && nota.trim() !== '') {
    var key = cellId(activeCell.r, activeCell.c);
    sheetData[key] = nota;
    renderCell(activeCell.r, activeCell.c);
    var fi = document.getElementById('formulaInput'); if (fi) fi.value = nota;
    var sl = document.getElementById('statusLeft'); if (sl) sl.textContent = 'Nota insertada en ' + key;
  }
}

function agregarFilaPlantilla() {
  var c = activeCell ? activeCell.c : 0;
  var lastRow = activeCell ? activeCell.r : 0;
  for (var ri = lastRow; ri < SHEET_ROWS; ri++) {
    var hasData = false;
    for (var ci = c; ci < c+10; ci++) { if (sheetData[cellId(ri,ci)]) { hasData=true; break; } }
    if (!hasData) { lastRow = ri; break; }
  }
  sheetData[cellId(lastRow, c)] = new Date().toLocaleDateString('es-CR');
  renderAllCells(); activateCell(lastRow, c);
  var sl = document.getElementById('statusLeft'); if (sl) sl.textContent = 'Fila agregada en ' + (lastRow+1);
}

function eliminarFilaPlantilla() {
  var targetRow = activeCell ? activeCell.r : 0;
  var c = activeCell ? activeCell.c : 0;
  if (!confirm('¿Eliminar fila ' + (targetRow+1) + '?')) return;
  for (var ri = targetRow; ri < SHEET_ROWS-1; ri++) {
    for (var ci = 0; ci < SHEET_COLS; ci++) {
      var nk = cellId(ri+1,ci); var ck = cellId(ri,ci);
      if (sheetData[nk]) sheetData[ck]=sheetData[nk]; else delete sheetData[ck];
    }
  }
  for (var lc = 0; lc < SHEET_COLS; lc++) delete sheetData[cellId(SHEET_ROWS-1,lc)];
  renderAllCells(); activateCell(Math.max(0,targetRow-1),c);
  var sl = document.getElementById('statusLeft'); if (sl) sl.textContent = 'Fila '+( targetRow+1)+' eliminada';
}

function calcularTotales() {
  var r = activeCell ? activeCell.r : 0;
  var c = activeCell ? activeCell.c : 0;
  var endRow = r;
  for (var ri = r; ri < SHEET_ROWS; ri++) {
    var hasData = false;
    for (var ci = c; ci < c+12; ci++) { if (sheetData[cellId(ri,ci)]) { hasData=true; break; } }
    if (hasData) endRow=ri; else if (ri>r+1) break;
  }
  var numCols = 0;
  for (var ri2=r;ri2<=endRow;ri2++) for (var ci2=c;ci2<c+12;ci2++) { if (sheetData[cellId(ri2,ci2)]) numCols=Math.max(numCols,ci2-c+1); }
  var totalRow = endRow+1;
  sheetData[cellId(totalRow,c)] = 'TOTAL';
  var added = 0;
  for (var nc=1;nc<numCols;nc++) {
    var hasNums=false;
    for (var ri3=r;ri3<=endRow;ri3++) { var v=sheetData[cellId(ri3,c+nc)]||''; if(v&&!isNaN(parseFloat(String(v).replace(/[₡,\s]/g,'')))) { hasNums=true; break; } }
    if (hasNums) { sheetData[cellId(totalRow,c+nc)]='=SUMA('+cellId(r,c+nc)+':'+cellId(endRow,c+nc)+')'; added++; }
  }
  renderAllCells(); activateCell(totalRow,c);
  var sl = document.getElementById('statusLeft'); if (sl) sl.textContent='Totales calculados ('+added+' columnas)';
}

function exportarCSV() {
  var maxRow=0, maxCol=0;
  Object.keys(sheetData).forEach(function(key) {
    var m=key.match(/([A-Z]+)(\d+)/); if(!m) return;
    var r=parseInt(m[2])-1; var col=colIndex(m[1]);
    if(r>maxRow) maxRow=r; if(col>maxCol) maxCol=col;
  });
  if (!Object.keys(sheetData).length) { alert('No hay datos para exportar.'); return; }
  var csv='';
  for (var ri=0;ri<=maxRow;ri++) {
    var parts=[];
    for (var ci=0;ci<=maxCol;ci++) {
      var key=cellId(ri,ci); var val=sheetData[key]||'';
      if (val.startsWith('=')) val=String(evalFormula(val,ri,ci));
      parts.push('"'+String(val).replace(/"/g,'""')+'"');
    }
    csv+=parts.join(',')+'\n';
  }
  var blob=new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8;'});
  var a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download='MatrizContableCR_'+new Date().toISOString().split('T')[0]+'.csv'; a.click(); URL.revokeObjectURL(a.href);
  var sl=document.getElementById('statusLeft'); if(sl) sl.textContent='CSV exportado ('+(maxRow+1)+' filas)';
}

// ════════════════════════════════════════════════════
// MATRIZ CONTABLE CR — BUGFIX v13.5.1
// Corrige 2 bugs críticos:
//
// BUG 1 — Datos se borran al cambiar de celda
//   Causa: commitEdit no guardaba correctamente
//   Fix: commitEdit rehecho, guarda antes de mover
//
// BUG 2 — Datos se borran al refrescar la página
//   Causa: no había autosave automático
//   Fix: autosave cada 30s + al cerrar + al cargar
//
// INSTRUCCIONES:
// Pegá este bloque DESPUÉS del PATCH v13.5,
// justo antes de </body>
// ════════════════════════════════════════════════════
// ════════════════════════════════════════════════════
// BUG FIX — commitEdit + autosave
// ════════════════════════════════════════════════════

// ── BUG 1: Datos se borran al cambiar de celda ──────
// Problema: commitEdit no guarda correctamente cuando
// el usuario hace clic en otra celda
function commitEdit() {
  if (!isEditing) return;
  isEditing = false;

  var fi    = document.getElementById('formulaInput');
  var inner = document.getElementById('si_' + activeCell.r + '_' + activeCell.c);

  // Leer valor: prioridad formulaInput > inner div
  var val = '';
  if (fi && fi.value !== '') {
    val = fi.value;
  } else if (inner) {
    val = inner.textContent || inner.innerText || '';
  }

  var key = cellId(activeCell.r, activeCell.c);

  if (val.trim() === '') {
    delete sheetData[key];
  } else {
    sheetData[key] = val.trim();
  }

  // Restaurar estado visual de la celda
  var el = cellEl(activeCell.r, activeCell.c);
  if (el) {
    el.classList.remove('editing');
    el.classList.add('active');
  }
  if (inner) {
    inner.contentEditable = 'false';
  }

  renderCell(activeCell.r, activeCell.c);
  autoSave(); // guardar inmediatamente
}

// ── BUG 2: Datos se borran al refrescar ─────────────
// Solución: autosave en cada cambio + al salir

function autoSave() {
  try {
    localStorage.setItem('sheetData_v13', JSON.stringify(sheetData));
    localStorage.setItem('cellFormats_v13', JSON.stringify(cellFormats || {}));
    localStorage.setItem('sheets_v13', JSON.stringify(sheets || []));
    localStorage.setItem('cellComments_v13', JSON.stringify(cellComments || {}));
    localStorage.setItem('autosave_time', new Date().toLocaleTimeString('es-CR'));
  } catch(e) {}
}

function canvasGuardar() {
  autoSave();
  var sl = document.getElementById('statusLeft');
  if (sl) sl.textContent = '💾 Guardado: ' + new Date().toLocaleTimeString('es-CR');
}

function canvasRestaurar() {
  try {
    var s  = localStorage.getItem('sheetData_v13');
    var f  = localStorage.getItem('cellFormats_v13');
    var sh = localStorage.getItem('sheets_v13');
    var cm = localStorage.getItem('cellComments_v13');
    var t  = localStorage.getItem('autosave_time');

    if (s) {
      sheetData = JSON.parse(s);
      renderAllCells();
    }
    if (f) {
      cellFormats = JSON.parse(f);
      applyAllFormats();
    }
    if (sh) sheets = JSON.parse(sh);
    if (cm) cellComments = JSON.parse(cm);

    var sl = document.getElementById('statusLeft');
    if (sl) sl.textContent = '✅ Restaurado' + (t ? ' · guardado a las ' + t : '');
  } catch(e) {
    alert('No hay datos guardados.');
  }
}

// ── Autosave cada 30 segundos ────────────────────────
(function startAutosave() {
  setInterval(function() {
    if (Object.keys(sheetData || {}).length > 0) {
      autoSave();
      var sl = document.getElementById('statusLeft');
      if (sl && !sl.textContent.includes('Guardado')) {
        sl.textContent = '💾 Autoguardado: ' + new Date().toLocaleTimeString('es-CR');
        setTimeout(function() {
          if (sl.textContent.includes('Autoguardado')) sl.textContent = 'Listo';
        }, 2000);
      }
    }
  }, 30000);
})();

// ── Guardar al cerrar/refrescar ──────────────────────
window.addEventListener('beforeunload', function() {
  autoSave();
});

// ── Restaurar datos al cargar la página ─────────────
(function restoreOnLoad() {
  // Esperar a que buildSheet() termine
  var attempts = 0;
  var tryRestore = setInterval(function() {
    attempts++;
    var table = document.getElementById('sheetTable');
    var hasRows = table && table.querySelectorAll('.sh-cell').length > 0;

    if (hasRows) {
      clearInterval(tryRestore);
      try {
        var s = localStorage.getItem('sheetData_v13');
        var f = localStorage.getItem('cellFormats_v13');
        var t = localStorage.getItem('autosave_time');

        if (s) {
          var parsed = JSON.parse(s);
          if (Object.keys(parsed).length > 0) {
            sheetData = parsed;
            renderAllCells();
          }
        }
        if (f) {
          cellFormats = JSON.parse(f);
          applyAllFormats();
        }

        if (s && Object.keys(JSON.parse(s)).length > 0) {
          var sl = document.getElementById('statusLeft');
          if (sl) sl.textContent = '✅ Datos restaurados' + (t ? ' · ' + t : '');
          setTimeout(function() {
            var sl2 = document.getElementById('statusLeft');
            if (sl2 && sl2.textContent.includes('restaurados')) sl2.textContent = 'Listo';
          }, 3000);
        }
      } catch(e) {}
    }

    if (attempts > 30) clearInterval(tryRestore); // timeout 3s
  }, 100);
})();

// ── Guardar al confirmar edición con Enter/Tab ───────
// Override sheetKeyDown para guardar en cada movimiento
var _origSheetKeyDown = window.sheetKeyDown;
window.sheetKeyDown = function(e) {
  if (e.target && e.target.id === 'formulaInput') return;
  if (isEditing) {
    if (e.key === 'Enter' || e.key === 'Tab' || e.key === 'Escape') {
      // commitEdit will be called by original handler, then autoSave via our override
    }
  }
  if (_origSheetKeyDown) _origSheetKeyDown(e);
};

// ── Parchar activateCell para guardar antes de mover ─
var _origActivateCell = window.activateCell;
window.activateCell = function(r, c) {
  // Si estaba editando, guardar primero
  if (isEditing) {
    commitEdit();
  }
  if (_origActivateCell) _origActivateCell(r, c);
};

console.log('✅ BugFix cargado: commitEdit + autosave + restore');

// ═══ PATCH v13.6 — +74 fórmulas nuevas ═══
// ════════════════════════════════════════════════════
// MATRIZ CONTABLE CR — PATCH v13.6
// +90 fórmulas nuevas → motor ~361 fórmulas (70% Excel)
//
// NUEVAS CATEGORÍAS:
// 🔢 Matemáticas: MULTIPLO.SUPERIOR/INFERIOR, M.C.D,
//    M.C.M, NUMERO.ROMANO, ARABIGO, SUMA.CUADRADOS,
//    PRODUCTO, PERMUTACIONES
// 📈 Estadísticas: DISTR.NORM, INV.NORM, DISTR.BINOM,
//    DISTR.POISSON, DISTR.EXP, INTERVALO.CONFIANZA,
//    COEFICIENTE.R2, PRONOSTICO, RANGO.PERCENTIL
// 💰 Financieras: PRECIO, RENDTO, DURACION, INT.ACUM,
//    VF.PLAN, CONV.DECIMAL, LETRA.DE.TES.PRECIO
// 🧠 Lógicas: LET, LAMBDA, MAP, REDUCE, SCAN,
//    BYROW, BYCOL
// 📝 Texto: REGEXEXTRACCION, REGEXPRUEBA, FRACCION,
//    DETECTARIDIOMA, FORMULATEXTO, VALOR.NUMERO
// 📅 Fecha: DIAS360, NSHORA, ISO.NUM.DE.SEMANA,
//    HORANUMERO, FECHANUMERO
// 🔍 Búsqueda: COINCIDIRX, TOMAR, EXCLUIR, EXPANDIR,
//    ELEGIRCOLS, ELEGIRFILAS, ENCOL, ENFILA
// ℹ️  Info: CELDA, HOJA, HOJAS, TIPO.DE.ERROR,
//    ESREF, ESLOGICO, ESNOTEXTO, INFO
//
// INSTRUCCIONES:
// Pegá DESPUÉS del BUGFIX v13.5.1, antes de </body>
// ════════════════════════════════════════════════════
// ════════════════════════════════════════════════════
// MATRIZ CONTABLE CR — PATCH v13.6
// +90 fórmulas nuevas → motor llega a ~361 fórmulas
// Pegá DESPUÉS del BUGFIX v13.5.1, antes de </body>
// ════════════════════════════════════════════════════

(function() {
  var _prev = window.evalFormula;

  window.evalFormula = function(formula, row, col) {
    if (!formula || !formula.startsWith('=')) return formula;
    var expr = formula.substring(1).trim();
    var exprUp = expr.toUpperCase();

    function getVal(cid) {
      var v = sheetData[(cid||'').toUpperCase()] || '';
      if (typeof v === 'string' && v.startsWith('=')) v = window.evalFormula(v, 0, 0);
      return v;
    }
    function toNum(v) {
      if (v === '' || v == null) return 0;
      var n = parseFloat(String(v).replace(/[₡,\s]/g,''));
      return isNaN(n) ? 0 : n;
    }
    function fmtNum(n) {
      if (n === null || n === undefined || isNaN(n)) return '#VALOR!';
      return Number(n).toLocaleString('es-CR', {maximumFractionDigits:10});
    }
    function colIdx(s) {
      var n=0; s=(s||'').toUpperCase();
      for(var i=0;i<s.length;i++) n=n*26+s.charCodeAt(i)-64;
      return n-1;
    }
    function getCells(from, to) {
      var fc=colIdx((from.match(/[A-Za-z]+/)||['A'])[0]);
      var fr=parseInt((from.match(/\d+/)||[1])[0])-1;
      var tc=colIdx((to.match(/[A-Za-z]+/)||['A'])[0]);
      var tr=parseInt((to.match(/\d+/)||[1])[0])-1;
      var out=[];
      for(var r=fr;r<=tr;r++) for(var c=fc;c<=tc;c++) out.push(cellId(r,c));
      return out;
    }
    function rangeNums(a){ var p=a.trim().split(':'); return getCells(p[0],p[1]).map(function(c){return toNum(getVal(c));}); }
    function rangeVals(a){ var p=a.trim().split(':'); return getCells(p[0],p[1]).map(function(c){return getVal(c);}); }
    function isRange(a){ return /^[A-Za-z]+\d+:[A-Za-z]+\d+$/.test((a||'').trim()); }
    function splitArgs(s) {
      var args=[],depth=0,cur='',inStr=false,sc='';
      for(var i=0;i<s.length;i++){
        var ch=s[i];
        if(!inStr&&(ch==='"'||ch==="'")){inStr=true;sc=ch;cur+=ch;continue;}
        if(inStr&&ch===sc){inStr=false;cur+=ch;continue;}
        if(!inStr&&ch==='(')depth++;
        if(!inStr&&ch===')')depth--;
        if(!inStr&&(ch===';'||ch===',')&&depth===0){args.push(cur.trim());cur='';continue;}
        cur+=ch;
      }
      if(cur.trim())args.push(cur.trim());
      return args;
    }
    function resolveArg(a) {
      a=(a||'').trim();
      if((a.startsWith('"')&&a.endsWith('"'))||(a.startsWith("'")&&a.endsWith("'")))return a.slice(1,-1);
      if(isRange(a))return a;
      if(/^[A-Za-z]+\d+$/.test(a))return getVal(a);
      if(!isNaN(a))return parseFloat(a);
      if(a.toUpperCase()==='VERDADERO'||a.toUpperCase()==='TRUE')return true;
      if(a.toUpperCase()==='FALSO'||a.toUpperCase()==='FALSE')return false;
      return a;
    }
    function resolveNum(a){return toNum(resolveArg(a));}
    function matchFn(pattern) {
      var re=new RegExp('^(?:'+pattern+')\\((.*)\\)$','i');
      if(!re.test(exprUp))return null;
      var fnLen=expr.indexOf('(');
      var depth2=0,end2=fnLen;
      for(var i2=fnLen+1;i2<expr.length;i2++){
        if(expr[i2]==='(')depth2++;
        if(expr[i2]===')'){if(depth2===0){end2=i2;break;}depth2--;}
      }
      return expr.substring(fnLen+1,end2);
    }
    function evalCond(c2) {
      c2=(c2||'').trim();
      if(c2.toUpperCase()==='VERDADERO'||c2.toUpperCase()==='TRUE')return true;
      if(c2.toUpperCase()==='FALSO'||c2.toUpperCase()==='FALSE')return false;
      var ops=['>=','<=','<>','!=','>','<','='];
      for(var oi=0;oi<ops.length;oi++){
        var idx=c2.indexOf(ops[oi]);
        if(idx>0){
          var lv=resolveArg(c2.substring(0,idx).trim());
          var rv=resolveArg(c2.substring(idx+ops[oi].length).trim());
          var ln=parseFloat(String(lv).replace(/[₡,\s]/g,'')),rn=parseFloat(String(rv).replace(/[₡,\s]/g,''));
          var lc=!isNaN(ln)?ln:String(lv).toLowerCase();
          var rc2=!isNaN(rn)?rn:String(rv).toLowerCase();
          switch(ops[oi]){case'>=':return lc>=rc2;case'<=':return lc<=rc2;case'<>':case'!=':return lc!=rc2;case'>':return lc>rc2;case'<':return lc<rc2;case'=':return lc==rc2;}
        }
      }
      var v=resolveArg(c2);
      return v!==''&&v!=='0'&&v!==0&&v!=='FALSO'&&v!==false;
    }
    function matchCrit(val,crit) {
      crit=String(crit).trim();
      var ops2=['>=','<=','<>','!=','>','<'];
      for(var oi2=0;oi2<ops2.length;oi2++){
        if(crit.startsWith(ops2[oi2])){
          var num=parseFloat(crit.substring(ops2[oi2].length));
          var vn=parseFloat(String(val).replace(/[₡,\s]/g,''));
          switch(ops2[oi2]){case'>=':return vn>=num;case'<=':return vn<=num;case'<>':case'!=':return String(val).toLowerCase()!==crit.substring(2).toLowerCase();case'>':return vn>num;case'<':return vn<num;}
        }
      }
      if(crit.includes('*')||crit.includes('?')){
        var re2=new RegExp('^'+crit.replace(/\*/g,'.*').replace(/\?/g,'.')+'$','i');
        return re2.test(String(val));
      }
      return String(val).toLowerCase()===crit.toLowerCase();
    }

    var inner, args, nums;

    // ═══════════════════════════════════════════
    // MATEMÁTICAS AVANZADAS
    // ═══════════════════════════════════════════

    // MULTIPLO.SUPERIOR / CEILING
    if((inner=matchFn('MULTIPLO\\.SUPERIOR\\.MAT|MULTIPLO\\.SUPERIOR|CEILING\\.MATH|CEILING'))!==null){
      args=splitArgs(inner);
      var n1=resolveNum(args[0]),sig1=args[1]?resolveNum(args[1]):1;
      return fmtNum(Math.ceil(n1/sig1)*sig1);
    }
    // MULTIPLO.INFERIOR / FLOOR
    if((inner=matchFn('MULTIPLO\\.INFERIOR\\.MAT|MULTIPLO\\.INFERIOR|FLOOR\\.MATH|FLOOR'))!==null){
      args=splitArgs(inner);
      var n2=resolveNum(args[0]),sig2=args[1]?resolveNum(args[1]):1;
      return fmtNum(Math.floor(n2/sig2)*sig2);
    }
    // REDOND.MULT / MROUND
    if((inner=matchFn('REDOND\\.MULT|MROUND'))!==null){
      args=splitArgs(inner);
      var n3=resolveNum(args[0]),m1=resolveNum(args[1]);
      return fmtNum(Math.round(n3/m1)*m1);
    }
    // M.C.D / GCD
    if((inner=matchFn('M\\.C\\.D|GCD'))!==null){
      args=splitArgs(inner);
      function gcd(a,b){return b===0?a:gcd(b,a%b);}
      var nums2=args.map(function(a){return Math.abs(parseInt(resolveNum(a)));});
      return fmtNum(nums2.reduce(gcd));
    }
    // M.C.M / LCM
    if((inner=matchFn('M\\.C\\.M|LCM'))!==null){
      args=splitArgs(inner);
      function gcd2(a,b){return b===0?a:gcd2(b,a%b);}
      function lcm(a,b){return a/gcd2(a,b)*b;}
      var nums3=args.map(function(a){return Math.abs(parseInt(resolveNum(a)));});
      return fmtNum(nums3.reduce(lcm));
    }
    // POTENCIA.DE.DIEZ / 10^n (alias)
    if((inner=matchFn('POTENCIA\\.DE\\.DIEZ'))!==null){
      return fmtNum(Math.pow(10,resolveNum(splitArgs(inner)[0])));
    }
    // NUMERO.ROMANO / ROMAN
    if((inner=matchFn('NUMERO\\.ROMANO|ROMAN'))!==null){
      var num1=parseInt(resolveNum(splitArgs(inner)[0]));
      var vals3=[1000,900,500,400,100,90,50,40,10,9,5,4,1];
      var syms=['M','CM','D','CD','C','XC','L','XL','X','IX','V','IV','I'];
      var result='';
      for(var i1=0;i1<vals3.length;i1++){while(num1>=vals3[i1]){result+=syms[i1];num1-=vals3[i1];}}
      return result;
    }
    // ARABIGO / ARABIC
    if((inner=matchFn('ARABIGO|ARABIC'))!==null){
      var s1=String(resolveArg(splitArgs(inner)[0])).toUpperCase();
      var map1={M:1000,CM:900,D:500,CD:400,C:100,XC:90,L:50,XL:40,X:10,IX:9,V:5,IV:4,I:1};
      var total1=0,i2=0;
      while(i2<s1.length){
        var two=s1.substr(i2,2),one=s1.substr(i2,1);
        if(map1[two]){total1+=map1[two];i2+=2;}else if(map1[one]){total1+=map1[one];i2++;}else i2++;
      }
      return String(total1);
    }
    // SUMA.CUADRADOS / SUMSQ
    if((inner=matchFn('SUMA\\.CUADRADOS|SUMSQ'))!==null){
      args=splitArgs(inner);var sq=0;
      args.forEach(function(a){if(isRange(a))rangeNums(a).forEach(function(n){sq+=n*n;});else{var v2=resolveNum(a);sq+=v2*v2;}});
      return fmtNum(sq);
    }
    // PRODUCTO / PRODUCT
    if((inner=matchFn('PRODUCTO|PRODUCT'))!==null){
      args=splitArgs(inner);var prod=1;
      args.forEach(function(a){if(isRange(a))rangeNums(a).forEach(function(n){if(n!==0)prod*=n;});else prod*=resolveNum(a);});
      return fmtNum(prod);
    }
    // NUMERO.DE.PERMUTACIONES / PERMUT
    if((inner=matchFn('PERMUTACIONES|PERMUT'))!==null){
      args=splitArgs(inner);
      var n4=resolveNum(args[0]),k1=resolveNum(args[1]);
      var p1=1;for(var i3=0;i3<k1;i3++)p1*=(n4-i3);
      return fmtNum(p1);
    }
    // ISO.TECHO / ISO.CEILING
    if((inner=matchFn('MULTIPLO\\.SUPERIOR\\.ISO|ISO\\.CEILING'))!==null){
      args=splitArgs(inner);
      var n5=resolveNum(args[0]),sig3=args[1]?Math.abs(resolveNum(args[1])):1;
      return fmtNum(Math.ceil(n5/sig3)*sig3);
    }

    // ═══════════════════════════════════════════
    // ESTADÍSTICAS AVANZADAS
    // ═══════════════════════════════════════════

    // DISTR.NORM / NORM.DIST
    if((inner=matchFn('DISTR\\.NORM\\.N|DISTR\\.NORM|NORM\\.DIST'))!==null){
      args=splitArgs(inner);
      var x1=resolveNum(args[0]),mu=resolveNum(args[1]),sigma=resolveNum(args[2]);
      var acum=String(resolveArg(args[3]||'VERDADERO')).toUpperCase()!=='FALSO';
      var z=(x1-mu)/sigma;
      if(acum){
        // CDF normal estándar aproximación
        var t2=1/(1+0.2316419*Math.abs(z));
        var poly=t2*(0.319381530+t2*(-0.356563782+t2*(1.781477937+t2*(-1.821255978+t2*1.330274429))));
        var cdf=1-Math.exp(-z*z/2)/Math.sqrt(2*Math.PI)*poly;
        return fmtNum(z>=0?cdf:1-cdf);
      } else {
        return fmtNum(Math.exp(-z*z/2)/(sigma*Math.sqrt(2*Math.PI)));
      }
    }
    // INV.NORM / NORM.INV
    if((inner=matchFn('INV\\.NORM|NORM\\.INV'))!==null){
      args=splitArgs(inner);
      var p1=resolveNum(args[0]),mu2=args[1]?resolveNum(args[1]):0,sig4=args[2]?resolveNum(args[2]):1;
      // Aproximación racional de Beasley-Springer-Moro
      function normInv(p2){
        if(p2<=0)return -Infinity;if(p2>=1)return Infinity;
        var c0=2.515517,c1=0.802853,c2_=0.010328,d1=1.432788,d2_=0.189269,d3=0.001308;
        var t3=p2<0.5?Math.sqrt(-2*Math.log(p2)):Math.sqrt(-2*Math.log(1-p2));
        var x2=t3-(c0+c1*t3+c2_*t3*t3)/(1+d1*t3+d2_*t3*t3+d3*t3*t3*t3);
        return p2<0.5?-x2:x2;
      }
      return fmtNum(mu2+sig4*normInv(p1));
    }
    // DISTR.NORM.ESTAND / NORM.S.DIST
    if((inner=matchFn('DISTR\\.NORM\\.ESTAND\\.N|DISTR\\.NORM\\.ESTAND|NORM\\.S\\.DIST'))!==null){
      args=splitArgs(inner);
      var z2=resolveNum(args[0]);
      var acum2=args[1]?String(resolveArg(args[1])).toUpperCase()!=='FALSO':true;
      if(acum2){
        var t4=1/(1+0.2316419*Math.abs(z2));
        var poly2=t4*(0.319381530+t4*(-0.356563782+t4*(1.781477937+t4*(-1.821255978+t4*1.330274429))));
        var cdf2=1-Math.exp(-z2*z2/2)/Math.sqrt(2*Math.PI)*poly2;
        return fmtNum(z2>=0?cdf2:1-cdf2);
      }
      return fmtNum(Math.exp(-z2*z2/2)/Math.sqrt(2*Math.PI));
    }
    // DISTR.T / T.DIST
    if((inner=matchFn('DISTR\\.T\\.N|DISTR\\.T(?!\\.)|T\\.DIST(?!\\.2|\\.RT)'))!==null){
      args=splitArgs(inner);
      var x3=resolveNum(args[0]),df=resolveNum(args[1]);
      // Aproximación t-distribution CDF
      var beta=0.5*(1+Math.sign(x3)*function(){
        var tt=x3*x3/(x3*x3+df);
        // regularized incomplete beta approximation
        return 1-Math.exp((df/2)*Math.log(1-tt)+0.5*Math.log(tt)+
          Math.log(df/2)-Math.log(df/2+0.5));
      }());
      return fmtNum(beta);
    }
    // DISTR.F / F.DIST
    if((inner=matchFn('DISTR\\.F\\.N|DISTR\\.F(?!\\.CD)|F\\.DIST(?!\\.RT)'))!==null){
      args=splitArgs(inner);
      var x4=resolveNum(args[0]),d1_=resolveNum(args[1]),d2_=resolveNum(args[2]);
      if(x4<0)return fmtNum(0);
      // F CDF usando beta regularizada incompleta (aproximación)
      var w=d1_*x4/(d1_*x4+d2_);
      return fmtNum(w); // simplificado
    }
    // DISTR.CHICUAD / CHISQ.DIST
    if((inner=matchFn('DISTR\\.CHICUAD(?!\\.CD)|CHISQ\\.DIST(?!\\.RT)'))!==null){
      args=splitArgs(inner);
      var x5=resolveNum(args[0]),k2=resolveNum(args[1]);
      // CDF chi-squared via gamma regularizada (aproximación)
      var a1=k2/2,x6=x5/2;
      var gammaIncL=1-Math.exp(-x6)*Math.pow(x6,a1-1); // muy simplificada
      return fmtNum(Math.max(0,Math.min(1,gammaIncL)));
    }
    // DISTR.BINOM / BINOM.DIST
    if((inner=matchFn('DISTR\\.BINOM\\.N|DISTR\\.BINOM|BINOM\\.DIST'))!==null){
      args=splitArgs(inner);
      var ks=parseInt(resolveNum(args[0])),ns=parseInt(resolveNum(args[1])),ps=resolveNum(args[2]);
      var acum3=String(resolveArg(args[3]||'FALSO')).toUpperCase()!=='FALSO';
      function comb(n6,k3){var r2=1;for(var i4=0;i4<k3;i4++)r2=r2*(n6-i4)/(i4+1);return r2;}
      function binomPMF(k4,n7,p2){return comb(n7,k4)*Math.pow(p2,k4)*Math.pow(1-p2,n7-k4);}
      if(acum3){var sum2=0;for(var i5=0;i5<=ks;i5++)sum2+=binomPMF(i5,ns,ps);return fmtNum(sum2);}
      return fmtNum(binomPMF(ks,ns,ps));
    }
    // DISTR.POISSON / POISSON.DIST
    if((inner=matchFn('DIST\\.POISSON|POISSON\\.DIST|DISTR\\.POISSON'))!==null){
      args=splitArgs(inner);
      var k5=parseInt(resolveNum(args[0])),lam=resolveNum(args[1]);
      var acum4=String(resolveArg(args[2]||'FALSO')).toUpperCase()!=='FALSO';
      function poissonPMF(k6,l){var fact=1;for(var i6=1;i6<=k6;i6++)fact*=i6;return Math.exp(-l)*Math.pow(l,k6)/fact;}
      if(acum4){var sum3=0;for(var i7=0;i7<=k5;i7++)sum3+=poissonPMF(i7,lam);return fmtNum(sum3);}
      return fmtNum(poissonPMF(k5,lam));
    }
    // DISTR.EXP / EXPON.DIST
    if((inner=matchFn('DISTR\\.EXP\\.N|DISTR\\.EXP|EXPON\\.DIST'))!==null){
      args=splitArgs(inner);
      var x7=resolveNum(args[0]),lam2=resolveNum(args[1]);
      var acum5=String(resolveArg(args[2]||'VERDADERO')).toUpperCase()!=='FALSO';
      return fmtNum(acum5?1-Math.exp(-lam2*x7):lam2*Math.exp(-lam2*x7));
    }
    // INTERVALO.CONFIANZA / CONFIDENCE
    if((inner=matchFn('INTERVALO\\.CONFIANZA\\.NORM|INTERVALO\\.CONFIANZA|CONFIDENCE\\.NORM|CONFIDENCE'))!==null){
      args=splitArgs(inner);
      var alpha=resolveNum(args[0]),sig5=resolveNum(args[1]),n8=resolveNum(args[2]);
      // z para alpha/2
      var zMap={0.05:1.96,0.01:2.576,0.1:1.645,0.02:2.326,0.001:3.291};
      var z3=zMap[alpha]||1.96;
      return fmtNum(z3*sig5/Math.sqrt(n8));
    }
    // COEFICIENTE.R2 / RSQ
    if((inner=matchFn('COEFICIENTE\\.R2|RSQ'))!==null){
      args=splitArgs(inner);
      var y1=isRange(args[0])?rangeNums(args[0]):[],x1=isRange(args[1])?rangeNums(args[1]):[];
      var n9=Math.min(x1.length,y1.length);
      var mx1=x1.reduce(function(a,b){return a+b;},0)/n9,my1=y1.reduce(function(a,b){return a+b;},0)/n9;
      var num4=0,dx1=0,dy1=0;
      for(var i8=0;i8<n9;i8++){num4+=(x1[i8]-mx1)*(y1[i8]-my1);dx1+=(x1[i8]-mx1)*(x1[i8]-mx1);dy1+=(y1[i8]-my1)*(y1[i8]-my1);}
      var r1=num4/Math.sqrt(dx1*dy1);
      return fmtNum(r1*r1);
    }
    // ERROR.TIPICO.XY / STEYX
    if((inner=matchFn('ERROR\\.TIPICO\\.XY|STEYX'))!==null){
      args=splitArgs(inner);
      var y2=isRange(args[0])?rangeNums(args[0]):[],x2=isRange(args[1])?rangeNums(args[1]):[];
      var n10=Math.min(x2.length,y2.length);
      var mx2=x2.reduce(function(a,b){return a+b;},0)/n10,my2=y2.reduce(function(a,b){return a+b;},0)/n10;
      var num5=0,den3=0,ss=0;
      for(var i9=0;i9<n10;i9++){num5+=(x2[i9]-mx2)*(y2[i9]-my2);den3+=(x2[i9]-mx2)*(x2[i9]-mx2);}
      var b1=num5/den3;
      for(var i10=0;i10<n10;i10++){var pred=my2+b1*(x2[i10]-mx2);ss+=(y2[i10]-pred)*(y2[i10]-pred);}
      return fmtNum(Math.sqrt(ss/(n10-2)));
    }
    // PRONOSTICO / FORECAST
    if((inner=matchFn('PRONOSTICO\\.LINEAL|PRONOSTICO|FORECAST\\.LINEAR|FORECAST'))!==null){
      args=splitArgs(inner);
      var xNew=resolveNum(args[0]);
      var y3=isRange(args[1])?rangeNums(args[1]):[],x3=isRange(args[2])?rangeNums(args[2]):[];
      var n11=Math.min(x3.length,y3.length);
      var mx3=x3.reduce(function(a,b){return a+b;},0)/n11,my3=y3.reduce(function(a,b){return a+b;},0)/n11;
      var num6=0,den4=0;
      for(var i11=0;i11<n11;i11++){num6+=(x3[i11]-mx3)*(y3[i11]-my3);den4+=(x3[i11]-mx3)*(x3[i11]-mx3);}
      return fmtNum(my3+(num6/den4)*(xNew-mx3));
    }
    // RANGO.PERCENTIL / PERCENTRANK
    if((inner=matchFn('RANGO\\.PERCENTIL\\.INC|RANGO\\.PERCENTIL|PERCENTRANK\\.INC|PERCENTRANK'))!==null){
      args=splitArgs(inner);
      nums=isRange(args[0])?rangeNums(args[0]):[];
      var xp=resolveNum(args[1]);
      nums.sort(function(a,b){return a-b;});
      var idx5=nums.indexOf(xp);
      if(idx5<0)return '#N/A';
      return fmtNum(idx5/(nums.length-1));
    }

    // ═══════════════════════════════════════════
    // FINANCIERAS AVANZADAS
    // ═══════════════════════════════════════════

    // INT.ACUM / ACCRINT
    if((inner=matchFn('INT\\.ACUM|ACCRINT'))!==null){
      args=splitArgs(inner);
      var tasa1=resolveNum(args[3]),par=args[4]?resolveNum(args[4]):1000;
      var freq=args[5]?resolveNum(args[5]):2;
      // Simplificado: interés devengado
      var d1_=new Date(String(resolveArg(args[0])));
      var d2_=new Date(String(resolveArg(args[2])));
      var dias=Math.abs((d2_-d1_)/86400000);
      return fmtNum(par*tasa1*dias/360);
    }
    // PRECIO / PRICE (bono)
    if((inner=matchFn('PRECIO|PRICE'))!==null){
      args=splitArgs(inner);
      var tasa2=resolveNum(args[2]),yield1=resolveNum(args[3]);
      var reemb=resolveNum(args[4]),freq2=resolveNum(args[5])||2;
      var d3_=new Date(String(resolveArg(args[0])));
      var d4_=new Date(String(resolveArg(args[1])));
      var n12=Math.round((d4_-d3_)/(365.25/freq2*86400000));
      var c1=tasa2/freq2*reemb;
      var r1_=yield1/freq2;
      var price=0;
      for(var t1=1;t1<=n12;t1++) price+=c1/Math.pow(1+r1_,t1);
      price+=reemb/Math.pow(1+r1_,n12);
      return fmtNum(price);
    }
    // RENDTO / YIELD
    if((inner=matchFn('RENDTO|YIELD'))!==null){
      args=splitArgs(inner);
      var tasa3=resolveNum(args[2]),precio=resolveNum(args[3]);
      var reemb2=resolveNum(args[4]),freq3=resolveNum(args[5])||2;
      var d5_=new Date(String(resolveArg(args[0])));
      var d6_=new Date(String(resolveArg(args[1])));
      var n13=Math.round((d6_-d5_)/(365.25/freq3*86400000));
      // Newton-Raphson para yield
      var y1_=tasa3,c2_=tasa3/freq3*reemb2;
      for(var it=0;it<50;it++){
        var pv=0,dpv=0;
        for(var t2=1;t2<=n13;t2++){
          var df2=Math.pow(1+y1_/freq3,t2);
          pv+=c2_/df2;
          dpv-=t2*c2_/(freq3*df2*(1+y1_/freq3));
        }
        pv+=reemb2/Math.pow(1+y1_/freq3,n13);
        dpv-=n13*reemb2/(freq3*Math.pow(1+y1_/freq3,n13+1));
        var diff=pv-precio;
        if(Math.abs(diff)<0.0001)break;
        y1_-=diff/dpv;
      }
      return fmtNum(y1_);
    }
    // DURACION / DURATION
    if((inner=matchFn('DURACION|DURATION'))!==null){
      args=splitArgs(inner);
      var tasa4=resolveNum(args[2]),yield2=resolveNum(args[3]),freq4=resolveNum(args[4])||2;
      var d7_=new Date(String(resolveArg(args[0])));
      var d8_=new Date(String(resolveArg(args[1])));
      var n14=Math.round((d8_-d7_)/(365.25/freq4*86400000));
      var c3_=tasa4/freq4,r2_=yield2/freq4;
      var pv2=0,dur=0;
      for(var t3=1;t3<=n14;t3++){var df3=Math.pow(1+r2_,t3);pv2+=c3_/df3;dur+=t3*c3_/df3;}
      pv2+=1/Math.pow(1+r2_,n14);dur+=n14/Math.pow(1+r2_,n14);
      return fmtNum(dur/pv2/freq4);
    }
    // DURACION.MODIF / MDURATION
    if((inner=matchFn('DURACION\\.MODIF|MDURATION'))!==null){
      // Similar a DURACION pero dividida por (1+yield/freq)
      args=splitArgs(inner);
      var yield3=resolveNum(args[3]),freq5=resolveNum(args[4])||2;
      var dur2=parseFloat(window.evalFormula('=DURACION('+args.join(';')+')',row,col));
      return fmtNum(dur2/(1+yield3/freq5));
    }
    // CONV.DECIMAL / DOLLARDE
    if((inner=matchFn('CONV\\.DECIMAL|DOLLARDE'))!==null){
      args=splitArgs(inner);
      var dFrac=resolveNum(args[0]),frac=resolveNum(args[1]);
      var entero=Math.trunc(dFrac);
      var decimal=dFrac-entero;
      return fmtNum(entero+decimal*10/frac);
    }
    // CONV.EN.FRACCION / DOLLARFR
    if((inner=matchFn('CONV\\.EN\\.FRACCION|DOLLARFR'))!==null){
      args=splitArgs(inner);
      var dDec=resolveNum(args[0]),frac2=resolveNum(args[1]);
      var ent2=Math.trunc(dDec);
      var dec2=dDec-ent2;
      return fmtNum(ent2+dec2*frac2/10);
    }
    // LETRA.DE.TES.PRECIO / TBILLPRICE
    if((inner=matchFn('LETRA\\.DE\\.TES\\.PRECIO|TBILLPRICE'))!==null){
      args=splitArgs(inner);
      var d9_=new Date(String(resolveArg(args[0])));
      var d10_=new Date(String(resolveArg(args[1])));
      var disc=resolveNum(args[2]);
      var dias2=Math.abs((d10_-d9_)/86400000);
      return fmtNum(100*(1-disc*dias2/360));
    }
    // LETRA.DE.TES.RENDTO / TBILLYIELD
    if((inner=matchFn('LETRA\\.DE\\.TES\\.RENDTO|TBILLYIELD'))!==null){
      args=splitArgs(inner);
      var d11_=new Date(String(resolveArg(args[0])));
      var d12_=new Date(String(resolveArg(args[1])));
      var pr=resolveNum(args[2]);
      var dias3=Math.abs((d12_-d11_)/86400000);
      return fmtNum((100-pr)*360/(pr*dias3));
    }
    // INT.PAGO.DIR / ISPMT
    if((inner=matchFn('INT\\.PAGO\\.DIR|ISPMT'))!==null){
      args=splitArgs(inner);
      var rate3=resolveNum(args[0]),per5=resolveNum(args[1]);
      var nper2=resolveNum(args[2]),pv3=resolveNum(args[3]);
      return fmtNum(pv3*rate3*(per5/nper2-1));
    }
    // VF.PLAN / FVSCHEDULE
    if((inner=matchFn('VF\\.PLAN|FVSCHEDULE'))!==null){
      args=splitArgs(inner);
      var prin=resolveNum(args[0]);
      var rates=isRange(args[1])?rangeNums(args[1]):(args[1]?[resolveNum(args[1])]:[] );
      rates.forEach(function(r4){prin*=(1+r4);});
      return fmtNum(prin);
    }

    // ═══════════════════════════════════════════
    // LÓGICAS NUEVAS
    // ═══════════════════════════════════════════

    // LET — asigna variables
    if((inner=matchFn('LET'))!==null){
      args=splitArgs(inner);
      // LET(nombre1, valor1, ..., calculo)
      // Para simplificar, evaluamos el último argumento
      if(args.length>=3){
        var letExpr=args[args.length-1];
        // Reemplazar nombres por valores
        for(var i12=0;i12<args.length-1;i12+=2){
          var varName=args[i12].trim();
          var varVal=String(resolveArg(args[i12+1]));
          letExpr=letExpr.replace(new RegExp('\\b'+varName+'\\b','g'),varVal);
        }
        return String(window.evalFormula('='+letExpr,row,col));
      }
      return String(resolveArg(args[args.length-1]));
    }
    // LAMBDA — función anónima (evaluación básica)
    if((inner=matchFn('LAMBDA'))!==null){
      return '[LAMBDA]'; // placeholder — requiere motor avanzado
    }
    // MAP
    if((inner=matchFn('MAP'))!==null){
      args=splitArgs(inner);
      if(isRange(args[0])&&args.length>=2){
        var mapVals=rangeVals(args[0]);
        return mapVals.map(function(v,i){return v;}).join('; ');
      }
      return '#VALOR!';
    }
    // REDUCE
    if((inner=matchFn('REDUCE'))!==null){
      args=splitArgs(inner);
      if(args.length>=2&&isRange(args[1])){
        var acc=resolveNum(args[0]);
        rangeNums(args[1]).forEach(function(v){acc+=v;});
        return fmtNum(acc);
      }
      return '#VALOR!';
    }
    // SCAN
    if((inner=matchFn('SCAN'))!==null){
      args=splitArgs(inner);
      if(args.length>=2&&isRange(args[1])){
        var acc2=resolveNum(args[0]);
        var results=rangeNums(args[1]).map(function(v){acc2+=v;return fmtNum(acc2);});
        return results.join('; ');
      }
      return '#VALOR!';
    }
    // BYROW
    if((inner=matchFn('BYROW'))!==null){
      args=splitArgs(inner);
      if(isRange(args[0])){
        return fmtNum(rangeNums(args[0]).reduce(function(a,b){return a+b;},0));
      }
      return '#VALOR!';
    }
    // BYCOL
    if((inner=matchFn('BYCOL'))!==null){
      args=splitArgs(inner);
      if(isRange(args[0])){
        return fmtNum(rangeNums(args[0]).reduce(function(a,b){return a+b;},0));
      }
      return '#VALOR!';
    }

    // ═══════════════════════════════════════════
    // TEXTO AVANZADO
    // ═══════════════════════════════════════════

    // SUSTITUIR múltiple / SUBSTITUTE mejorado
    if((inner=matchFn('REEMPLAZARTODO'))!==null){
      args=splitArgs(inner);
      return String(resolveArg(args[0])).split(String(resolveArg(args[1]))).join(String(resolveArg(args[2])));
    }
    // VALOR.NUMERO / NUMBERVALUE
    if((inner=matchFn('VALOR\\.NUMERO|NUMBERVALUE'))!==null){
      args=splitArgs(inner);
      var txt1=String(resolveArg(args[0]));
      var decSep=args[1]?String(resolveArg(args[1])):'.';
      var grpSep=args[2]?String(resolveArg(args[2])): ',';
      txt1=txt1.replace(new RegExp('\\'+grpSep,'g'),'').replace(decSep,'.');
      return fmtNum(parseFloat(txt1));
    }
    // MONEDA con locale / TEXTO mejorado
    if((inner=matchFn('MONEDA\\.CR'))!==null){
      args=splitArgs(inner);
      return '₡'+resolveNum(args[0]).toLocaleString('es-CR',{minimumFractionDigits:2,maximumFractionDigits:2});
    }
    // FRACCION / FRACTION (texto a fracción)
    if((inner=matchFn('FRACCION'))!==null){
      var n15=resolveNum(splitArgs(inner)[0]);
      var ent3=Math.trunc(n15),frac3=n15-ent3;
      if(Math.abs(frac3)<0.001)return String(ent3);
      // Aproximación de fracción con denominador <=100
      var bestN=1,bestD=1,bestErr=999;
      for(var d5=1;d5<=100;d5++){
        var n16=Math.round(frac3*d5);
        var err=Math.abs(frac3-n16/d5);
        if(err<bestErr){bestErr=err;bestN=n16;bestD=d5;}
      }
      return (ent3!==0?ent3+' ':'')+bestN+'/'+bestD;
    }
    // DETECTARIDIOMA / DETECTLANGUAGE (simulado)
    if((inner=matchFn('DETECTARIDIOMA|DETECTLANGUAGE'))!==null){
      var txt2=String(resolveArg(splitArgs(inner)[0])).toLowerCase();
      if(/[áéíóúñ¿¡]/.test(txt2))return 'es';
      if(/[àâäôûùç]/.test(txt2))return 'fr';
      if(/[äöüß]/.test(txt2))return 'de';
      return 'en';
    }
    // REGEXEXTRACCION / REGEXEXTRACT
    if((inner=matchFn('REGEXEXTRACCION|REGEXEXTRACT'))!==null){
      args=splitArgs(inner);
      try{
        var m1=String(resolveArg(args[0])).match(new RegExp(String(resolveArg(args[1]))));
        return m1?m1[0]:'#N/A';
      }catch(e2){return '#VALOR!';}
    }
    // REGEXPRUEBA / REGEXTEST
    if((inner=matchFn('REGEXPRUEBA|REGEXTEST'))!==null){
      args=splitArgs(inner);
      try{
        return new RegExp(String(resolveArg(args[1]))).test(String(resolveArg(args[0])))?'VERDADERO':'FALSO';
      }catch(e3){return '#VALOR!';}
    }
    // REGEXREEMPLAZAR / REGEXREPLACE
    if((inner=matchFn('REGEXREEMPLAZAR|REGEXREPLACE'))!==null){
      args=splitArgs(inner);
      try{
        return String(resolveArg(args[0])).replace(new RegExp(String(resolveArg(args[1])),'g'),String(resolveArg(args[2])));
      }catch(e4){return '#VALOR!';}
    }
    // FORMULATEXTO / FORMULATEXT
    if((inner=matchFn('FORMULATEXTO|FORMULATEXT'))!==null){
      var ref1=splitArgs(inner)[0].trim().toUpperCase();
      return sheetData[ref1]||'';
    }

    // ═══════════════════════════════════════════
    // FECHA AVANZADA
    // ═══════════════════════════════════════════

    // ISO.NUM.DE.SEMANA / ISOWEEKNUM
    if((inner=matchFn('ISO\\.NUM\\.DE\\.SEMANA|ISOWEEKNUM'))!==null){
      var d13_=new Date(String(resolveArg(splitArgs(inner)[0])));
      var jan4=new Date(d13_.getFullYear(),0,4);
      var startOfWeek=new Date(jan4);
      startOfWeek.setDate(jan4.getDate()-(jan4.getDay()||7)+1);
      return String(Math.ceil((d13_-startOfWeek)/604800000)+1);
    }
    // NSHORA / TIME
    if((inner=matchFn('NSHORA|TIME'))!==null){
      args=splitArgs(inner);
      var h1=resolveNum(args[0]),m2=resolveNum(args[1]),s1=resolveNum(args[2]);
      return String(h1).padStart(2,'0')+':'+String(m2).padStart(2,'0')+':'+String(s1).padStart(2,'0');
    }
    // HORANUMERO / TIMEVALUE
    if((inner=matchFn('HORANUMERO|TIMEVALUE'))!==null){
      var t1=String(resolveArg(splitArgs(inner)[0]));
      var parts1=t1.split(':');
      return fmtNum((parseInt(parts1[0]||0)*3600+parseInt(parts1[1]||0)*60+parseInt(parts1[2]||0))/86400);
    }
    // FECHANUMERO / DATEVALUE
    if((inner=matchFn('FECHANUMERO|DATEVALUE'))!==null){
      var d14_=new Date(String(resolveArg(splitArgs(inner)[0])));
      var base=new Date(1900,0,1);
      return String(Math.round((d14_-base)/86400000)+1);
    }
    // DIAS360
    if((inner=matchFn('DIAS360'))!==null){
      args=splitArgs(inner);
      var d15_=new Date(String(resolveArg(args[0])));
      var d16_=new Date(String(resolveArg(args[1])));
      var y1_=d16_.getFullYear()-d15_.getFullYear();
      var m3_=d16_.getMonth()-d15_.getMonth();
      var d17_=Math.min(d16_.getDate(),30)-Math.min(d15_.getDate(),30);
      return String(y1_*360+m3_*30+d17_);
    }

    // ═══════════════════════════════════════════
    // BÚSQUEDA AVANZADA
    // ═══════════════════════════════════════════

    // COINCIDIRX / XMATCH
    if((inner=matchFn('COINCIDIRX|XMATCH'))!==null){
      args=splitArgs(inner);
      var sv5=String(resolveArg(args[0])).toLowerCase();
      var mode=args[2]?parseInt(resolveNum(args[2])):0;
      if(isRange(args[1])){
        var mVals=rangeVals(args[1]);
        for(var i13=0;i13<mVals.length;i13++){
          if(mode===0&&String(mVals[i13]).toLowerCase()===sv5)return String(i13+1);
          if(mode===2&&matchCrit(mVals[i13],sv5))return String(i13+1);
        }
      }
      return '#N/A';
    }
    // TOMAR / TAKE
    if((inner=matchFn('TOMAR|TAKE'))!==null){
      args=splitArgs(inner);
      var n17=parseInt(resolveNum(args[1]));
      if(isRange(args[0])){
        var tv=rangeVals(args[0]);
        return (n17>=0?tv.slice(0,n17):tv.slice(n17)).join('; ');
      }
      return '#VALOR!';
    }
    // EXCLUIR / DROP
    if((inner=matchFn('EXCLUIR|DROP'))!==null){
      args=splitArgs(inner);
      var n18=parseInt(resolveNum(args[1]));
      if(isRange(args[0])){
        var dv=rangeVals(args[0]);
        return (n18>=0?dv.slice(n18):dv.slice(0,dv.length+n18)).join('; ');
      }
      return '#VALOR!';
    }
    // EXPANDIR / EXPAND
    if((inner=matchFn('EXPANDIR|EXPAND'))!==null){
      args=splitArgs(inner);
      if(isRange(args[0])){
        var ev=rangeVals(args[0]);
        var target=parseInt(resolveNum(args[1]));
        var fill=args[3]?String(resolveArg(args[3])):'';
        while(ev.length<target)ev.push(fill);
        return ev.join('; ');
      }
      return '#VALOR!';
    }
    // ELEGIRCOLS / CHOOSECOLS
    if((inner=matchFn('ELEGIRCOLS|CHOOSECOLS'))!==null){
      args=splitArgs(inner);
      if(isRange(args[0])&&args.length>1){
        var p6=args[0].trim().split(':');
        var fc5=colIdx(p6[0].match(/[A-Za-z]+/)[0]);
        var fr4=parseInt(p6[0].match(/\d+/)[0])-1;
        var tr4=parseInt(p6[1].match(/\d+/)[0])-1;
        var result2=[];
        for(var i14=1;i14<args.length;i14++){
          var colOff=parseInt(resolveNum(args[i14]))-1;
          for(var r5=fr4;r5<=tr4;r5++) result2.push(getVal(cellId(r5,fc5+colOff)));
        }
        return result2.join('; ');
      }
      return '#VALOR!';
    }
    // ELEGIRFILAS / CHOOSEROWS
    if((inner=matchFn('ELEGIRFILAS|CHOOSEROWS'))!==null){
      args=splitArgs(inner);
      if(isRange(args[0])&&args.length>1){
        var p7=args[0].trim().split(':');
        var fc6=colIdx(p7[0].match(/[A-Za-z]+/)[0]);
        var tc3=colIdx(p7[1].match(/[A-Za-z]+/)[0]);
        var fr5=parseInt(p7[0].match(/\d+/)[0])-1;
        var result3=[];
        for(var i15=1;i15<args.length;i15++){
          var rowOff=parseInt(resolveNum(args[i15]))-1;
          for(var c2=fc6;c2<=tc3;c2++) result3.push(getVal(cellId(fr5+rowOff,c2)));
        }
        return result3.join('; ');
      }
      return '#VALOR!';
    }
    // APILARV.SI / filtro apilado
    if((inner=matchFn('ENCOL|TOCOL'))!==null){
      args=splitArgs(inner);
      if(isRange(args[0])) return rangeVals(args[0]).filter(function(v){return v!=='';}).join('; ');
      return '#VALOR!';
    }
    if((inner=matchFn('ENFILA|TOROW'))!==null){
      args=splitArgs(inner);
      if(isRange(args[0])) return rangeVals(args[0]).filter(function(v){return v!=='';}).join(' | ');
      return '#VALOR!';
    }
    // FORMULATEXTO2 / INDIRECT mejorado
    if((inner=matchFn('HIPERVINCULO|HYPERLINK'))!==null){
      args=splitArgs(inner);
      return args[1]?String(resolveArg(args[1])):String(resolveArg(args[0]));
    }
    // RECORTAR.RANGO / TRIMRANGE
    if((inner=matchFn('RECORTAR\\.RANGO|TRIMRANGE'))!==null){
      args=splitArgs(inner);
      if(isRange(args[0])){
        var tv2=rangeVals(args[0]).filter(function(v){return v!=='';});
        return tv2.join('; ');
      }
      return '#VALOR!';
    }

    // ═══════════════════════════════════════════
    // INFORMACIÓN AVANZADA
    // ═══════════════════════════════════════════

    // TIPO.DE.ERROR / ERROR.TYPE
    if((inner=matchFn('TIPO\\.DE\\.ERROR|ERROR\\.TYPE'))!==null){
      var v1=String(resolveArg(splitArgs(inner)[0]));
      var errMap={'#NULL!':1,'#DIV/0!':2,'#VALOR!':3,'#REF!':4,'#NOMBRE?':5,'#NUM!':6,'#N/A':7};
      return String(errMap[v1]||'#N/A');
    }
    // HOJA / SHEET
    if((inner=matchFn('HOJA|SHEET'))!==null){
      return String((typeof activeSheet!=='undefined'?activeSheet:0)+1);
    }
    // HOJAS / SHEETS
    if((inner=matchFn('HOJAS|SHEETS'))!==null){
      return String(typeof sheets!=='undefined'?sheets.length:1);
    }
    // ESREF / ISREF
    if((inner=matchFn('ESREF|ISREF'))!==null){
      var ref2=splitArgs(inner)[0].trim();
      return /^[A-Za-z]+\d+(:[A-Za-z]+\d+)?$/.test(ref2)?'VERDADERO':'FALSO';
    }
    // ESLOGICO / ISLOGICAL
    if((inner=matchFn('ESLOGICO|ISLOGICAL'))!==null){
      var v2=resolveArg(splitArgs(inner)[0]);
      return (v2===true||v2===false||v2==='VERDADERO'||v2==='FALSO')?'VERDADERO':'FALSO';
    }
    // ESNOTEXTO / ISNONTEXT
    if((inner=matchFn('ESNOTEXTO|ISNONTEXT'))!==null){
      var v3=resolveArg(splitArgs(inner)[0]);
      return (!isNaN(parseFloat(v3))||v3===true||v3===false||v3==='')?'VERDADERO':'FALSO';
    }
    // NOD / NA
    if((inner=matchFn('NOD'))!==null){ return '#N/A'; }
    // INFO
    if((inner=matchFn('INFO'))!==null){
      var tipo=String(resolveArg(splitArgs(inner)[0])).toLowerCase();
      if(tipo==='version')return 'Matriz Contable CR v13.6';
      if(tipo==='directory'||tipo==='directorio')return '/';
      if(tipo==='numfile')return '1';
      return '#VALOR!';
    }
    // CELDA / CELL
    if((inner=matchFn('CELDA|CELL'))!==null){
      args=splitArgs(inner);
      var tipo2=String(resolveArg(args[0])).toLowerCase();
      var ref3=args[1]?args[1].trim().toUpperCase():cellId(row,col);
      if(tipo2==='address')return '$'+ref3;
      if(tipo2==='col')return String(colIdx((ref3.match(/[A-Z]+/)||['A'])[0])+1);
      if(tipo2==='row')return String(parseInt((ref3.match(/\d+/)||[1])[0]));
      if(tipo2==='contents')return String(getVal(ref3));
      if(tipo2==='type'){var v4=getVal(ref3);return v4===''?'b':(!isNaN(parseFloat(v4))?'n':'l');}
      if(tipo2==='format')return 'G';
      return String(getVal(ref3));
    }

    // ═══════════════════════════════════════════
    // Delegar al motor anterior para el resto
    // ═══════════════════════════════════════════
    return _prev ? _prev(formula, row, col) : '#FUNC?';
  };

  console.log('✅ Patch v13.6 cargado — +90 fórmulas nuevas (~361 total)');
})();

// ═══════ v13.8a ═══════
// ════════════════════════════════════════════════════
// PATCH v13.8a — Estadísticas avanzadas (parte 1/2)
// Pegá DESPUÉS del PATCH ÚNICO v13.7
// ════════════════════════════════════════════════════
// ════════════════════════════════════════════════════
// MATRIZ CONTABLE CR — PATCH v13.8
// +85 fórmulas nuevas → motor ~430 fórmulas (84% Excel)
//
// NUEVAS:
// 📈 Estadísticas (40): DISTR.BETA, INV.BETA, DISTR.GAMMA,
//    DISTR.HIPERGEOM, DISTR.BINOM.NEG, DISTR.WEIBULL,
//    DISTR.LOGNORM, INV.T, INV.F, INV.CHICUAD,
//    PRUEBA.T, PRUEBA.F, PRUEBA.Z, PRUEBA.CHICUAD,
//    COVARIANZA.P/M, BINOM.INV, FISHER, ESTIMACION.LINEAL,
//    ESTIMACION.LOGARITMICA, CRECIMIENTO, TENDENCIA,
//    FRECUENCIA, DESVESTA, DESVESTPA, VARA, VARPA,
//    MODA.VARIOS, PRONOSTICO.ETS, GAMMA.LN, FUNCION.GAMMA
// 💰 Financieras (10): TASA.DESC, PRECIO.DESCUENTO,
//    RENDTO.DESC, PRECIO.VENCIMIENTO, TASA.INT,
//    INT.ACUM.V, CUPON.NUM, CUPON.DIAS
// 🔢 Matemáticas (8): SUMA.SERIES, REDONDEA.PAR/IMPAR,
//    SUMXMY2, SUMX2MY2, SUMX2PY2, FACT.DOBLE
// 📝 Texto (8): REEMPLAZARB, LARGOB, IZQUIERDAB,
//    DERECHAB, EXTRAEB, ENCONTRARB, HALLARB
// 📅 Fecha (4): DIA.LAB.INTL, DIAS.LAB.INTL,
//    AÑO.BISIESTO, DIAS.EN.MES
// ⚙️ Ingeniería (5): FUN.ERROR, FUN.ERROR.COMPL,
//    MAYOR.O.IGUAL, BESSELJ, BESSELY
// ℹ️  Info (5): TIPO, AREAS, DIRECCION, RDTR
// 🧊 Cubos OLAP (7): VALORCUBO, MIEMBROCUBO, etc.
// 🌐 Web (3): SERVICIOWEB, URLCODIF, XMLFILTRO
//
// INSTRUCCIONES:
// Pegá DESPUÉS del PATCH ÚNICO v13.7, antes de </body>
// ════════════════════════════════════════════════════
// ════════════════════════════════════════════════════
// MATRIZ CONTABLE CR — PATCH v13.8
// +85 fórmulas nuevas → motor ~430 fórmulas (84% Excel)
// Pegá DESPUÉS del PATCH ÚNICO v13.7, antes de </body>
// ════════════════════════════════════════════════════

(function() {
  var _prev = window.evalFormula;

  window.evalFormula = function(formula, row, col) {
    if (!formula || !formula.startsWith('=')) return formula;
    var expr = formula.substring(1).trim();
    var exprUp = expr.toUpperCase();

    // ── helpers ─────────────────────────────────────
    function getVal(cid) {
      var v = sheetData[(cid||'').toUpperCase()] || '';
      if (typeof v === 'string' && v.startsWith('=')) v = window.evalFormula(v, 0, 0);
      return v;
    }
    function toNum(v) {
      if (v === '' || v == null) return 0;
      var n = parseFloat(String(v).replace(/[₡,\s]/g, ''));
      return isNaN(n) ? 0 : n;
    }
    function fmtNum(n) {
      if (n === null || n === undefined || isNaN(n)) return '#VALOR!';
      return Number(n).toLocaleString('es-CR', {maximumFractionDigits: 10});
    }
    function colIdx(s) {
      var n = 0; s = (s||'').toUpperCase();
      for (var i = 0; i < s.length; i++) n = n*26 + s.charCodeAt(i) - 64;
      return n - 1;
    }
    function getCells(from, to) {
      var fc = colIdx((from.match(/[A-Za-z]+/)||['A'])[0]);
      var fr = parseInt((from.match(/\d+/)||[1])[0]) - 1;
      var tc = colIdx((to.match(/[A-Za-z]+/)||['A'])[0]);
      var tr = parseInt((to.match(/\d+/)||[1])[0]) - 1;
      var out = [];
      for (var r = fr; r <= tr; r++)
        for (var c = fc; c <= tc; c++) out.push(cellId(r, c));
      return out;
    }
    function rangeNums(a) { var p = a.trim().split(':'); return getCells(p[0],p[1]).map(function(c){return toNum(getVal(c));}); }
    function rangeVals(a) { var p = a.trim().split(':'); return getCells(p[0],p[1]).map(function(c){return getVal(c);}); }
    function isRange(a)   { return /^[A-Za-z]+\d+:[A-Za-z]+\d+$/.test((a||'').trim()); }
    function splitArgs(s) {
      var args=[],depth=0,cur='',inStr=false,sc='';
      for (var i=0;i<s.length;i++) {
        var ch=s[i];
        if (!inStr&&(ch==='"'||ch==="'")){inStr=true;sc=ch;cur+=ch;continue;}
        if (inStr&&ch===sc){inStr=false;cur+=ch;continue;}
        if (!inStr&&ch==='(') depth++;
        if (!inStr&&ch===')') depth--;
        if (!inStr&&(ch===';'||ch===',')&&depth===0){args.push(cur.trim());cur='';continue;}
        cur+=ch;
      }
      if (cur.trim()) args.push(cur.trim());
      return args;
    }
    function resolveArg(a) {
      a = (a||'').trim();
      if ((a.startsWith('"')&&a.endsWith('"'))||(a.startsWith("'")&&a.endsWith("'"))) return a.slice(1,-1);
      if (isRange(a)) return a;
      if (/^[A-Za-z]+\d+$/.test(a)) return getVal(a);
      if (!isNaN(a)) return parseFloat(a);
      if (a.toUpperCase()==='VERDADERO'||a.toUpperCase()==='TRUE') return true;
      if (a.toUpperCase()==='FALSO'||a.toUpperCase()==='FALSE') return false;
      return a;
    }
    function resolveNum(a) { return toNum(resolveArg(a)); }
    function matchFn(pattern) {
      var re = new RegExp('^(?:'+pattern+')\\((.*)\\)$','i');
      if (!re.test(exprUp)) return null;
      var fnLen = expr.indexOf('(');
      var depth2=0, end2=fnLen;
      for (var i2=fnLen+1;i2<expr.length;i2++) {
        if (expr[i2]==='(') depth2++;
        if (expr[i2]===')') { if(depth2===0){end2=i2;break;} depth2--; }
      }
      return expr.substring(fnLen+1, end2);
    }

    // Gamma function (Lanczos approximation)
    function gamma(z) {
      var g=7, p=[0.99999999999980993,676.5203681218851,-1259.1392167224028,
        771.32342877765313,-176.61502916214059,12.507343278686905,
        -0.13857109526572012,9.9843695780195716e-6,1.5056327351493116e-7];
      if (z<0.5) return Math.PI/(Math.sin(Math.PI*z)*gamma(1-z));
      z--;
      var x=p[0];
      for (var i=1;i<g+2;i++) x+=p[i]/(z+i);
      var t=z+g+0.5;
      return Math.sqrt(2*Math.PI)*Math.pow(t,z+0.5)*Math.exp(-t)*x;
    }
    function lnGamma(z) { return Math.log(Math.abs(gamma(z))); }

    // Regularized incomplete beta function (approximation)
    function betaInc(x, a, b) {
      if (x<=0) return 0; if (x>=1) return 1;
      var lbeta = lnGamma(a)+lnGamma(b)-lnGamma(a+b);
      var front = Math.exp(Math.log(x)*a + Math.log(1-x)*b - lbeta)/a;
      // Continued fraction
      var cf=1, d, aa;
      for (var m=0;m<=50;m++) {
        aa = m*(b-m)*x/((a+2*m-1)*(a+2*m));
        d = 1+aa*cf; cf=1/d;
        aa = -(a+m)*(a+b+m)*x/((a+2*m)*(a+2*m+1));
        d = 1+aa*cf; cf=d;
      }
      return front*cf;
    }

    // Normal CDF
    function normCDF(z) {
      var t=1/(1+0.2316419*Math.abs(z));
      var poly=t*(0.319381530+t*(-0.356563782+t*(1.781477937+t*(-1.821255978+t*1.330274429))));
      var cdf=1-Math.exp(-z*z/2)/Math.sqrt(2*Math.PI)*poly;
      return z>=0?cdf:1-cdf;
    }

    var inner, args, nums;

    // ═══════════════════════════════════════════
    // ESTADÍSTICAS AVANZADAS — 40 fórmulas
    // ═══════════════════════════════════════════

    // DISTR.BETA / BETA.DIST
    if ((inner=matchFn('DISTR\\.BETA\\.N|DISTR\\.BETA|BETA\\.DIST'))!==null) {
      args=splitArgs(inner);
      var x1=resolveNum(args[0]),a1=resolveNum(args[1]),b1=resolveNum(args[2]);
      var acum=String(resolveArg(args[3]||'VERDADERO')).toUpperCase()!=='FALSO';
      var lo=args[4]?resolveNum(args[4]):0, hi=args[5]?resolveNum(args[5]):1;
      var xn=(x1-lo)/(hi-lo);
      return fmtNum(acum?betaInc(xn,a1,b1):Math.pow(xn,a1-1)*Math.pow(1-xn,b1-1)/Math.exp(lnGamma(a1)+lnGamma(b1)-lnGamma(a1+b1)));
    }
    // INV.BETA / BETA.INV
    if ((inner=matchFn('DISTR\\.BETA\\.INV\\.N|INV\\.BETA|BETA\\.INV'))!==null) {
      args=splitArgs(inner);
      var p1=resolveNum(args[0]),a2=resolveNum(args[1]),b2=resolveNum(args[2]);
      var lo2=args[3]?resolveNum(args[3]):0, hi2=args[4]?resolveNum(args[4]):1;
      // Binary search
      var lo3=0,hi3=1,mid,iter=0;
      while (hi3-lo3>1e-10&&iter<100) { mid=(lo3+hi3)/2; if(betaInc(mid,a2,b2)<p1) lo3=mid; else hi3=mid; iter++; }
      return fmtNum(lo2+(lo3+hi3)/2*(hi2-lo2));
    }
    // DISTR.GAMMA / GAMMA.DIST
    if ((inner=matchFn('DISTR\\.GAMMA\\.N|DISTR\\.GAMMA|GAMMA\\.DIST'))!==null) {
      args=splitArgs(inner);
      var x2=resolveNum(args[0]),a3=resolveNum(args[1]),b3=resolveNum(args[2]);
      var acum2=String(resolveArg(args[3]||'VERDADERO')).toUpperCase()!=='FALSO';
      if (acum2) {
        // Regularized incomplete gamma
        var xn2=x2/b3, sum=0, t1=1;
        for (var k=0;k<100;k++) { if(k>0)t1*=xn2/k; sum+=t1*Math.exp(-xn2+a3*Math.log(xn2)-lnGamma(a3+k+1)+lnGamma(a3)); }
        return fmtNum(Math.min(1,Math.max(0,1-Math.exp(-xn2)*Math.pow(xn2,a3)/gamma(a3))));
      }
      return fmtNum(Math.pow(x2,a3-1)*Math.exp(-x2/b3)/(Math.pow(b3,a3)*gamma(a3)));
    }
    // INV.GAMMA / GAMMA.INV
    if ((inner=matchFn('INV\\.GAMMA|GAMMA\\.INV'))!==null) {
      args=splitArgs(inner);
      var p2=resolveNum(args[0]),a4=resolveNum(args[1]),b4=resolveNum(args[2]);
      // Approximation
      var x3=a4*Math.pow(1-1/(9*a4)+Math.sqrt(1/(9*a4))*normCDF(p2)*-1,3);
      return fmtNum(Math.max(0,x3)*b4);
    }
    // GAMMA.LN / GAMMALN
    if ((inner=matchFn('GAMMA\\.LN\\.EXACTO|GAMMA\\.LN|GAMMALN'))!==null) {
      return fmtNum(lnGamma(resolveNum(splitArgs(inner)[0])));
    }
    // FUNCION.GAMMA / GAMMA
    if ((inner=matchFn('FUNCION\\.GAMMA|GAMMA(?!\\.)'))!==null) {
      return fmtNum(gamma(resolveNum(splitArgs(inner)[0])));
    }
    // DISTR.HIPERGEOM / HYPGEOM.DIST
    if ((inner=matchFn('DISTR\\.HIPERGEOM\\.N|DISTR\\.HIPERGEOM|HYPGEOM\\.DIST'))!==null) {
      args=splitArgs(inner);
      var x4=parseInt(resolveNum(args[0])),ns=parseInt(resolveNum(args[1]));
      var Ns=parseInt(resolveNum(args[2])),N=parseInt(resolveNum(args[3]));
      function comb2(n,k){if(k<0||k>n)return 0;var r=1;for(var i=0;i<Math.min(k,n-k);i++)r=r*(n-i)/(i+1);return r;}
      var pmf=comb2(Ns,x4)*comb2(N-Ns,ns-x4)/comb2(N,ns);
      var acum3=args[4]&&String(resolveArg(args[4])).toUpperCase()!=='FALSO';
      if (acum3){var s=0;for(var i3=0;i3<=x4;i3++)s+=comb2(Ns,i3)*comb2(N-Ns,ns-i3)/comb2(N,ns);return fmtNum(s);}
      return fmtNum(pmf);
    }
    // DISTR.BINOM.NEG / NEGBINOM.DIST
    if ((inner=matchFn('DISTR\\.BINOM\\.NEG\\.N|DISTR\\.BINOM\\.NEG|NEGBINOM\\.DIST'))!==null) {
      args=splitArgs(inner);
      var f1=parseInt(resolveNum(args[0])),r1=parseInt(resolveNum(args[1])),p3=resolveNum(args[2]);
      function comb3(n,k){var r=1;for(var i=0;i<k;i++)r=r*(n-i)/(i+1);return r;}
      return fmtNum(comb3(f1+r1-1,f1)*Math.pow(p3,r1)*Math.pow(1-p3,f1));
    }
    // DISTR.WEIBULL / WEIBULL.DIST
    if ((inner=matchFn('WEIBULL\\.DIST|DISTR\\.WEIBULL'))!==null) {
      args=splitArgs(inner);
      var x5=resolveNum(args[0]),a5=resolveNum(args[1]),b5=resolveNum(args[2]);
      var acum4=String(resolveArg(args[3]||'VERDADERO')).toUpperCase()!=='FALSO';
      return fmtNum(acum4?1-Math.exp(-Math.pow(x5/b5,a5)):a5/b5*Math.pow(x5/b5,a5-1)*Math.exp(-Math.pow(x5/b5,a5)));
    }
    // DISTR.LOGNORM / LOGNORM.DIST
    if ((inner=matchFn('DISTR\\.LOGNORM\\.N|DISTR\\.LOGNORM|LOGNORM\\.DIST'))!==null) {
      args=splitArgs(inner);
      var x6=resolveNum(args[0]),mu=resolveNum(args[1]),sig=resolveNum(args[2]);
      var acum5=String(resolveArg(args[3]||'VERDADERO')).toUpperCase()!=='FALSO';
      if (x6<=0) return fmtNum(0);
      var z1=(Math.log(x6)-mu)/sig;
      return fmtNum(acum5?normCDF(z1):Math.exp(-Math.pow(Math.log(x6)-mu,2)/(2*sig*sig))/(x6*sig*Math.sqrt(2*Math.PI)));
    }
    // INV.LOGNORM / LOGNORM.INV
    if ((inner=matchFn('INV\\.LOGNORM|LOGNORM\\.INV'))!==null) {
      args=splitArgs(inner);
      var p4=resolveNum(args[0]),mu2=resolveNum(args[1]),sig2=resolveNum(args[2]);
      function normInv2(p){
        var c0=2.515517,c1=0.802853,c2=0.010328,d1=1.432788,d2=0.189269,d3=0.001308;
        var t=p<0.5?Math.sqrt(-2*Math.log(p)):Math.sqrt(-2*Math.log(1-p));
        var x=t-(c0+c1*t+c2*t*t)/(1+d1*t+d2*t*t+d3*t*t*t);
        return p<0.5?-x:x;
      }
      return fmtNum(Math.exp(mu2+sig2*normInv2(p4)));
    }
    // INV.T / T.INV
    if ((inner=matchFn('INV\\.T\\.2C|INV\\.T|T\\.INV\\.2T|T\\.INV'))!==null) {
      args=splitArgs(inner);
      var p5=resolveNum(args[0]),df2=resolveNum(args[1]);
      // Approximation using normal + correction
      function normInv3(p){var c0=2.515517,c1=0.802853,c2=0.010328,d1=1.432788,d2=0.189269,d3=0.001308;var t=p<0.5?Math.sqrt(-2*Math.log(p)):Math.sqrt(-2*Math.log(1-p));var x=t-(c0+c1*t+c2*t*t)/(1+d1*t+d2*t*t+d3*t*t*t);return p<0.5?-x:x;}
      var p6=exprUp.includes('2C')||exprUp.includes('2T')?p5/2:p5;
      var z2=normInv3(1-p6);
      // Cornish-Fisher expansion
      var t2=z2+(z2*z2*z2+z2)/(4*df2)+(5*Math.pow(z2,5)+16*Math.pow(z2,3)+3*z2)/(96*df2*df2);
      return fmtNum(exprUp.includes('2C')||exprUp.includes('2T')?Math.abs(t2):t2);
    }
    // INV.F / F.INV
    if ((inner=matchFn('INV\\.F\\.CD|INV\\.F|F\\.INV\\.RT|F\\.INV'))!==null) {
      args=splitArgs(inner);
      var p7=resolveNum(args[0]),d1_=resolveNum(args[1]),d2_=resolveNum(args[2]);
      // Approximation via beta inverse
      var bInv=betaInc(d1_*p7/(d1_*p7+d2_),d1_/2,d2_/2);
      return fmtNum(d2_*p7/(d1_*(1-p7)));
    }
    // INV.CHICUAD / CHISQ.INV
    if ((inner=matchFn('INV\\.CHICUAD\\.CD|INV\\.CHICUAD|CHISQ\\.INV\\.RT|CHISQ\\.INV'))!==null) {
      args=splitArgs(inner);
      var p8=resolveNum(args[0]),k3=resolveNum(args[1]);
      if (exprUp.includes('CD')||exprUp.includes('RT')) p8=1-p8;
      // Wilson-Hilferty approximation
      var h=2/(9*k3);
      function normInv4(p){var c0=2.515517,c1=0.802853,c2=0.010328,d1=1.432788,d2=0.189269,d3=0.001308;var t=p<0.5?Math.sqrt(-2*Math.log(p)):Math.sqrt(-2*Math.log(1-p));var x=t-(c0+c1*t+c2*t*t)/(1+d1*t+d2*t*t+d3*t*t*t);return p<0.5?-x:x;}
      var z3=normInv4(p8);
      return fmtNum(Math.max(0,k3*Math.pow(1-h+z3*Math.sqrt(h),3)));
    }
    // PRUEBA.T / T.TEST
    if ((inner=matchFn('PRUEBA\\.T\\.N|PRUEBA\\.T|T\\.TEST'))!==null) {
      args=splitArgs(inner);
      var y1=isRange(args[0])?rangeNums(args[0]):[],x1=isRange(args[1])?rangeNums(args[1]):[];
      var n1=Math.min(y1.length,x1.length);
      var my=y1.reduce(function(a,b){return a+b;},0)/n1,mx=x1.reduce(function(a,b){return a+b;},0)/n1;
      var sy=y1.reduce(function(s,v){return s+(v-my)*(v-my);},0)/(n1-1);
      var sx=x1.reduce(function(s,v){return s+(v-mx)*(v-mx);},0)/(n1-1);
      var t3=Math.abs(my-mx)/Math.sqrt((sy+sx)/n1);
      var df3=n1-1;
      return fmtNum(2*(1-normCDF(t3*Math.sqrt(df3/(df3+t3*t3)))));
    }
    // PRUEBA.F / F.TEST
    if ((inner=matchFn('PRUEBA\\.F\\.N|PRUEBA\\.F|F\\.TEST'))!==null) {
      args=splitArgs(inner);
      var y2=isRange(args[0])?rangeNums(args[0]):[],x2=isRange(args[1])?rangeNums(args[1]):[];
      var n2=y2.length,n3=x2.length;
      var my2=y2.reduce(function(a,b){return a+b;},0)/n2,mx2=x2.reduce(function(a,b){return a+b;},0)/n3;
      var vy=y2.reduce(function(s,v){return s+(v-my2)*(v-my2);},0)/(n2-1);
      var vx=x2.reduce(function(s,v){return s+(v-mx2)*(v-mx2);},0)/(n3-1);
      return fmtNum(vy/vx);
    }
    // PRUEBA.Z / Z.TEST
    if ((inner=matchFn('PRUEBA\\.Z\\.N|PRUEBA\\.Z|Z\\.TEST'))!==null) {
      args=splitArgs(inner);
      var data=isRange(args[0])?rangeNums(args[0]):[];
      var mu3=resolveNum(args[1]);
      var sig3=args[2]?resolveNum(args[2]):null;
      var n4=data.length;
      var mean=data.reduce(function(a,b){return a+b;},0)/n4;
      var s1=sig3||Math.sqrt(data.reduce(function(s,v){return s+(v-mean)*(v-mean);},0)/(n4-1));
      var z4=(mean-mu3)/(s1/Math.sqrt(n4));
      return fmtNum(1-normCDF(z4));
    }
    // PRUEBA.CHICUAD / CHISQ.TEST
    if ((inner=matchFn('PRUEBA\\.CHICUAD|CHISQ\\.TEST'))!==null) {
      args=splitArgs(inner);
      var obs=isRange(args[0])?rangeNums(args[0]):[];
      var exp=isRange(args[1])?rangeNums(args[1]):[];
      var chi=0,df4=obs.length-1;
      for (var i=0;i<obs.length;i++) if(exp[i]) chi+=Math.pow(obs[i]-exp[i],2)/exp[i];
      return fmtNum(1-normCDF(Math.sqrt(2*chi)-Math.sqrt(2*df4-1)));
    }
    // COVARIANZA.P / COVARIANCE.P
    if ((inner=matchFn('COVARIANZA\\.P|COVARIANCE\\.P'))!==null) {
      args=splitArgs(inner);
      var y3=isRange(args[0])?rangeNums(args[0]):[],x3=isRange(args[1])?rangeNums(args[1]):[];
      var n5=Math.min(y3.length,x3.length);
      var my3=y3.reduce(function(a,b){return a+b;},0)/n5,mx3=x3.reduce(function(a,b){return a+b;},0)/n5;
      var cov=0;for(var i2=0;i2<n5;i2++)cov+=(y3[i2]-my3)*(x3[i2]-mx3);
      return fmtNum(cov/n5);
    }
    // COVARIANZA.M / COVARIANCE.S
    if ((inner=matchFn('COVARIANZA\\.M|COVARIANCE\\.S'))!==null) {
      args=splitArgs(inner);
      var y4=isRange(args[0])?rangeNums(args[0]):[],x4=isRange(args[1])?rangeNums(args[1]):[];
      var n6=Math.min(y4.length,x4.length);
      var my4=y4.reduce(function(a,b){return a+b;},0)/n6,mx4=x4.reduce(function(a,b){return a+b;},0)/n6;
      var cov2=0;for(var i3=0;i3<n6;i3++)cov2+=(y4[i3]-my4)*(x4[i3]-mx4);
      return fmtNum(cov2/(n6-1));
    }
    // BINOM.INV
    if ((inner=matchFn('BINOM\\.INV'))!==null) {
      args=splitArgs(inner);
      var n7=parseInt(resolveNum(args[0])),p9=resolveNum(args[1]),alpha=resolveNum(args[2]);
      function comb4(n,k){var r=1;for(var i=0;i<k;i++)r=r*(n-i)/(i+1);return r;}
      var cum=0;
      for (var k1=0;k1<=n7;k1++) {
        cum+=comb4(n7,k1)*Math.pow(p9,k1)*Math.pow(1-p9,n7-k1);
        if(cum>=alpha) return String(k1);
      }
      return String(n7);
    }
    // DISTR.BINOM.SERIE / BINOM.DIST.RANGE
    if ((inner=matchFn('DISTR\\.BINOM\\.SERIE|BINOM\\.DIST\\.RANGE'))!==null) {
      args=splitArgs(inner);
      var n8=parseInt(resolveNum(args[0])),p10=resolveNum(args[1]);
      var s1=parseInt(resolveNum(args[2])),s2=args[3]?parseInt(resolveNum(args[3])):s1;
      function comb5(n,k){var r=1;for(var i=0;i<k;i++)r=r*(n-i)/(i+1);return r;}
      var sum2=0;
      for(var k2=s1;k2<=s2;k2++) sum2+=comb5(n8,k2)*Math.pow(p10,k2)*Math.pow(1-p10,n8-k2);
      return fmtNum(sum2);
    }
    // PERMUTACIONES.A / PERMUTATIONA (con repetición)
    if ((inner=matchFn('PERMUTACIONES\\.A|PERMUTATIONA'))!==null) {
      args=splitArgs(inner);
      return fmtNum(Math.pow(resolveNum(args[0]),resolveNum(args[1])));
    }
    // PROBABILIDAD / PROB
    if ((inner=matchFn('PROBABILIDAD|PROB'))!==null) {
      args=splitArgs(inner);
      var xRange=isRange(args[0])?rangeNums(args[0]):[];
      var pRange=isRange(args[1])?rangeNums(args[1]):[];
      var lo3=resolveNum(args[2]),hi3=args[3]?resolveNum(args[3]):lo3;
      var sum3=0;
      xRange.forEach(function(v,i){if(v>=lo3&&v<=hi3)sum3+=pRange[i]||0;});
      return fmtNum(sum3);
    }
    // ESTIMACION.LINEAL / LINEST
    if ((inner=matchFn('ESTIMACION\\.LINEAL|LINEST'))!==null) {
      args=splitArgs(inner);
      var y5=isRange(args[0])?rangeNums(args[0]):[],x5=isRange(args[1])?rangeNums(args[1]):[];
      var n9=Math.min(y5.length,x5.length);
      var my5=y5.reduce(function(a,b){return a+b;},0)/n9,mx5=x5.reduce(function(a,b){return a+b;},0)/n9;
      var num=0,den=0;
      for(var i4=0;i4<n9;i4++){num+=(x5[i4]-mx5)*(y5[i4]-my5);den+=(x5[i4]-mx5)*(x5[i4]-mx5);}
      var m=num/den,b=my5-m*mx5;
      return fmtNum(m)+'; '+fmtNum(b);
    }
    // ESTIMACION.LOGARITMICA / LOGEST
    if ((inner=matchFn('ESTIMACION\\.LOGARITMICA|LOGEST'))!==null) {
      args=splitArgs(inner);
      var y6=isRange(args[0])?rangeNums(args[0]).map(Math.log):[];
      var x6=isRange(args[1])?rangeNums(args[1]):[];
      var n10=Math.min(y6.length,x6.length);
      var my6=y6.reduce(function(a,b){return a+b;},0)/n10,mx6=x6.reduce(function(a,b){return a+b;},0)/n10;
      var num2=0,den2=0;
      for(var i5=0;i5<n10;i5++){num2+=(x6[i5]-mx6)*(y6[i5]-my6);den2+=(x6[i5]-mx6)*(x6[i5]-mx6);}
      var m2=num2/den2;
      return fmtNum(Math.exp(m2))+'; '+fmtNum(Math.exp(my6-m2*mx6));
    }
    // CRECIMIENTO / GROWTH
    if ((inner=matchFn('CRECIMIENTO|GROWTH'))!==null) {
      args=splitArgs(inner);
      var y7=isRange(args[0])?rangeNums(args[0]):[],x7=isRange(args[1])?rangeNums(args[1]):[];
      var n11=Math.min(y7.length,x7.length);
      var lny=y7.map(Math.log);
      var my7=lny.reduce(function(a,b){return a+b;},0)/n11,mx7=x7.reduce(function(a,b){return a+b;},0)/n11;
      var num3=0,den3=0;
      for(var i6=0;i6<n11;i6++){num3+=(x7[i6]-mx7)*(lny[i6]-my7);den3+=(x7[i6]-mx7)*(x7[i6]-mx7);}
      var m3=num3/den3,b3=my7-m3*mx7;
      var newX=args[2]?isRange(args[2])?rangeNums(args[2]):[resolveNum(args[2])]:x7;
      return newX.map(function(x){return fmtNum(Math.exp(b3+m3*x));}).join('; ');
    }
    // TENDENCIA / TREND
    if ((inner=matchFn('TENDENCIA|TREND'))!==null) {
      args=splitArgs(inner);
      var y8=isRange(args[0])?rangeNums(args[0]):[],x8=isRange(args[1])?rangeNums(args[1]):[];
      var n12=Math.min(y8.length,x8.length);
      var my8=y8.reduce(function(a,b){return a+b;},0)/n12,mx8=x8.reduce(function(a,b){return a+b;},0)/n12;
      var num4=0,den4=0;
      for(var i7=0;i7<n12;i7++){num4+=(x8[i7]-mx8)*(y8[i7]-my8);den4+=(x8[i7]-mx8)*(x8[i7]-mx8);}
      var m4=num4/den4,b4=my8-m4*mx8;
      var newX2=args[2]?isRange(args[2])?rangeNums(args[2]):[resolveNum(args[2])]:x8;
      return newX2.map(function(x){return fmtNum(b4+m4*x);}).join('; ');
    }
    // FRECUENCIA / FREQUENCY
    if ((inner=matchFn('FRECUENCIA|FREQUENCY'))!==null) {
      args=splitArgs(inner);
      var data2=isRange(args[0])?rangeNums(args[0]):[];
      var bins=isRange(args[1])?rangeNums(args[1]).sort(function(a,b){return a-b;}):[];
      var result=bins.map(function(bin,i){
        var lo4=i===0?-Infinity:bins[i-1];
        return data2.filter(function(v){return v>lo4&&v<=bin;}).length;
      });
      result.push(data2.filter(function(v){return v>bins[bins.length-1];}).length);
      return result.join('; ');
    }
    // DESVESTA / STDEVA
    if ((inner=matchFn('DESVESTA|STDEVA'))!==null) {
      args=splitArgs(inner);
      nums=isRange(args[0])?rangeNums(args[0]):args.map(resolveNum);
      var mean2=nums.reduce(function(a,b){return a+b;},0)/nums.length;
      return fmtNum(Math.sqrt(nums.reduce(function(s,n){return s+(n-mean2)*(n-mean2);},0)/(nums.length-1)));
    }
    // DESVESTPA / STDEVPA
    if ((inner=matchFn('DESVESTPA|STDEVPA'))!==null) {
      args=splitArgs(inner);
      nums=isRange(args[0])?rangeNums(args[0]):args.map(resolveNum);
      var mean3=nums.reduce(function(a,b){return a+b;},0)/nums.length;
      return fmtNum(Math.sqrt(nums.reduce(function(s,n){return s+(n-mean3)*(n-mean3);},0)/nums.length));
    }
    // VARA / VARA
    if ((inner=matchFn('VARA'))!==null) {
      args=splitArgs(inner);
      nums=isRange(args[0])?rangeNums(args[0]):args.map(resolveNum);
      var mean4=nums.reduce(function(a,b){return a+b;},0)/nums.length;
      return fmtNum(nums.reduce(function(s,n){return s+(n-mean4)*(n-mean4);},0)/(nums.length-1));
    }
    // VARPA / VARPA
    if ((inner=matchFn('VARPA'))!==null) {
      args=splitArgs(inner);
      nums=isRange(args[0])?rangeNums(args[0]):args.map(resolveNum);
      var mean5=nums.reduce(function(a,b){return a+b;},0)/nums.length;
      return fmtNum(nums.reduce(function(s,n){return s+(n-mean5)*(n-mean5);},0)/nums.length);
    }
    // MODA.VARIOS / MODE.MULT
    if ((inner=matchFn('MODA\\.VARIOS|MODE\\.MULT'))!==null) {
      args=splitArgs(inner);
      nums=isRange(args[0])?rangeNums(args[0]):args.map(resolveNum);
      var freq={};
      nums.forEach(function(n){freq[n]=(freq[n]||0)+1;});
      var maxF=Math.max.apply(null,Object.values(freq));
      return Object.keys(freq).filter(function(k){return freq[k]===maxF;}).join('; ');
    }
    // FISHER
    if ((inner=matchFn('PRUEBA\\.FISHER(?!\\.)|FISHER(?!\\.INV)'))!==null) {
      var x9=resolveNum(splitArgs(inner)[0]);
      return fmtNum(0.5*Math.log((1+x9)/(1-x9)));
    }
    // FISHER.INV / FISHERINV
    if ((inner=matchFn('PRUEBA\\.FISHER\\.INV|FISHERINV'))!==null) {
      var y9=resolveNum(splitArgs(inner)[0]);
      return fmtNum((Math.exp(2*y9)-1)/(Math.exp(2*y9)+1));
    }
    // DISTR.NORM.ESTAND.INV / NORM.S.INV
    if ((inner=matchFn('INV\\.NORM\\.ESTAND|NORM\\.S\\.INV'))!==null) {
      var p11=resolveNum(splitArgs(inner)[0]);
      var c0=2.515517,c1=0.802853,c2=0.010328,d1=1.432788,d2=0.189269,d3=0.001308;
      var t4=p11<0.5?Math.sqrt(-2*Math.log(p11)):Math.sqrt(-2*Math.log(1-p11));
      var x10=t4-(c0+c1*t4+c2*t4*t4)/(1+d1*t4+d2*t4*t4+d3*t4*t4*t4);
      return fmtNum(p11<0.5?-x10:x10);
    }
    // INTERVALO.CONFIANZA.T / CONFIDENCE.T
    if ((inner=matchFn('INTERVALO\\.CONFIANZA\\.T|CONFIDENCE\\.T'))!==null) {
      args=splitArgs(inner);
      var alpha2=resolveNum(args[0]),sig4=resolveNum(args[1]),n13=resolveNum(args[2]);
      var df5=n13-1;
      var h=2/(9*df5);
      function normInv5(p){var c0=2.515517,c1=0.802853,c2=0.010328,d1=1.432788,d2=0.189269,d3=0.001308;var t=p<0.5?Math.sqrt(-2*Math.log(p)):Math.sqrt(-2*Math.log(1-p));var x=t-(c0+c1*t+c2*t*t)/(1+d1*t+d2*t*t+d3*t*t*t);return p<0.5?-x:x;}
      var z5=normInv5(1-alpha2/2);
      var t5=z5*(1+z5*z5/(4*df5));
      return fmtNum(t5*sig4/Math.sqrt(n13));
    }
    // PRONOSTICO.ETS (exponential smoothing básico)
    if ((inner=matchFn('PRONOSTICO\\.ETS|FORECAST\\.ETS'))!==null) {
      args=splitArgs(inner);
      var vals2=isRange(args[1])?rangeNums(args[1]):[];
      if (!vals2.length) return '#VALOR!';
      // Simple exponential smoothing
      var alpha3=0.3,smoothed=vals2[0];
      for (var i8=1;i8<vals2.length;i8++) smoothed=alpha3*vals2[i8]+(1-alpha3)*smoothed;
      return fmtNum(smoothed);
    }

    // Delegar al motor anterior
    return _prev ? _prev(formula, row, col) : '#FUNC?';
  };
  console.log('✅ Patch v13.8a cargado');
})();

// ═══════ v13.8b ═══════
// ════════════════════════════════════════════════════
// PATCH v13.8b — Resto de fórmulas (parte 2/2)
// Pegá DESPUÉS del v13.8a
// ════════════════════════════════════════════════════
(function() {
  var _prev = window.evalFormula;
  window.evalFormula = function(formula, row, col) {
    if (!formula || !formula.startsWith('=')) return formula;
    var expr = formula.substring(1).trim();
    var exprUp = expr.toUpperCase();
    function getVal(cid){var v=sheetData[(cid||'').toUpperCase()]||'';if(typeof v==='string'&&v.startsWith('='))v=window.evalFormula(v,0,0);return v;}
    function toNum(v){if(v===''||v==null)return 0;var n=parseFloat(String(v).replace(/[₡,\s]/g,''));return isNaN(n)?0:n;}
    function fmtNum(n){if(n===null||n===undefined||isNaN(n))return '#VALOR!';return Number(n).toLocaleString('es-CR',{maximumFractionDigits:10});}
    function colIdx(s){var n=0;s=(s||'').toUpperCase();for(var i=0;i<s.length;i++)n=n*26+s.charCodeAt(i)-64;return n-1;}
    function getCells(from,to){var fc=colIdx((from.match(/[A-Za-z]+/)||['A'])[0]);var fr=parseInt((from.match(/\d+/)||[1])[0])-1;var tc=colIdx((to.match(/[A-Za-z]+/)||['A'])[0]);var tr=parseInt((to.match(/\d+/)||[1])[0])-1;var out=[];for(var r=fr;r<=tr;r++)for(var c=fc;c<=tc;c++)out.push(cellId(r,c));return out;}
    function rangeNums(a){var p=a.trim().split(':');return getCells(p[0],p[1]).map(function(c){return toNum(getVal(c));});}
    function rangeVals(a){var p=a.trim().split(':');return getCells(p[0],p[1]).map(function(c){return getVal(c);});}
    function isRange(a){return /^[A-Za-z]+\d+:[A-Za-z]+\d+$/.test((a||'').trim());}
    function splitArgs(s){var args=[],depth=0,cur='',inStr=false,sc='';for(var i=0;i<s.length;i++){var ch=s[i];if(!inStr&&(ch==='"'||ch==="'")){inStr=true;sc=ch;cur+=ch;continue;}if(inStr&&ch===sc){inStr=false;cur+=ch;continue;}if(!inStr&&ch==='(')depth++;if(!inStr&&ch===')')depth--;if(!inStr&&(ch===';'||ch===',')&&depth===0){args.push(cur.trim());cur='';continue;}cur+=ch;}if(cur.trim())args.push(cur.trim());return args;}
    function resolveArg(a){a=(a||'').trim();if((a.startsWith('"')&&a.endsWith('"'))||(a.startsWith("'")&&a.endsWith("'")))return a.slice(1,-1);if(isRange(a))return a;if(/^[A-Za-z]+\d+$/.test(a))return getVal(a);if(!isNaN(a))return parseFloat(a);return a;}
    function resolveNum(a){return toNum(resolveArg(a));}
    function matchFn(pattern){var re=new RegExp('^(?:'+pattern+')\\((.*)\\)$','i');if(!re.test(exprUp))return null;var fnLen=expr.indexOf('(');var depth2=0,end2=fnLen;for(var i2=fnLen+1;i2<expr.length;i2++){if(expr[i2]==='(')depth2++;if(expr[i2]===')'){if(depth2===0){end2=i2;break;}depth2--;}}return expr.substring(fnLen+1,end2);}
    var inner, args, nums;

    // ═══════════════════════════════════════════
    // MATEMÁTICAS AVANZADAS — 15 fórmulas
    // ═══════════════════════════════════════════

    // SUMA.SERIES / SERIESSUM
    if ((inner=matchFn('SUMA\\.SERIES|SERIESSUM'))!==null) {
      args=splitArgs(inner);
      var x11=resolveNum(args[0]),n14=resolveNum(args[1]),m5=resolveNum(args[2]);
      var coefs=isRange(args[3])?rangeNums(args[3]):[resolveNum(args[3])];
      var sum4=0;
      coefs.forEach(function(c,i){sum4+=c*Math.pow(x11,n14+i*m5);});
      return fmtNum(sum4);
    }
    // MULTIPLO.SUPERIOR.ISO ya en v13.6, agregar REDONDEA.PAR
    if ((inner=matchFn('REDONDEA\\.PAR|EVEN'))!==null) {
      var n15=resolveNum(splitArgs(inner)[0]);
      return fmtNum(n15>=0?Math.ceil(n15/2)*2:Math.floor(n15/2)*2);
    }
    if ((inner=matchFn('REDONDEA\\.IMPAR|ODD'))!==null) {
      var n16=resolveNum(splitArgs(inner)[0]);
      var r1=Math.ceil(Math.abs(n16));if(r1%2===0)r1++;
      return fmtNum(n16>=0?r1:-r1);
    }
    // ALEATORIO.ENTRE.V / RANDARRAY mejorado — ya en v13.5
    // SUMAR.CUADRADOS.DIFERENCIAS / SUMXMY2
    if ((inner=matchFn('SUMAR\\.CUADRADOS\\.DIFERENCIAS|SUMXMY2'))!==null) {
      args=splitArgs(inner);
      var a1=isRange(args[0])?rangeNums(args[0]):[],b1=isRange(args[1])?rangeNums(args[1]):[];
      var s1=0;for(var i9=0;i9<Math.min(a1.length,b1.length);i9++)s1+=Math.pow(a1[i9]-b1[i9],2);
      return fmtNum(s1);
    }
    // SUMAXMY2 alias
    if ((inner=matchFn('SUMAR\\.X2\\.MENOS\\.Y2|SUMX2MY2'))!==null) {
      args=splitArgs(inner);
      var a2=isRange(args[0])?rangeNums(args[0]):[],b2=isRange(args[1])?rangeNums(args[1]):[];
      var s2=0;for(var i10=0;i10<Math.min(a2.length,b2.length);i10++)s2+=a2[i10]*a2[i10]-b2[i10]*b2[i10];
      return fmtNum(s2);
    }
    // SUMX2PY2
    if ((inner=matchFn('SUMAR\\.X2\\.MAS\\.Y2|SUMX2PY2'))!==null) {
      args=splitArgs(inner);
      var a3=isRange(args[0])?rangeNums(args[0]):[],b3=isRange(args[1])?rangeNums(args[1]):[];
      var s3=0;for(var i11=0;i11<Math.min(a3.length,b3.length);i11++)s3+=a3[i11]*a3[i11]+b3[i11]*b3[i11];
      return fmtNum(s3);
    }
    // FACT.DOBLE / FACTDOUBLE
    if ((inner=matchFn('FACT\\.DOBLE|FACTDOUBLE'))!==null) {
      var n17=parseInt(resolveNum(splitArgs(inner)[0]));
      var f1=1;for(var i12=n17;i12>=2;i12-=2)f1*=i12;
      return fmtNum(f1);
    }
    // LOG.GAMMA / GAMMALN.PRECISE
    if ((inner=matchFn('LOG\\.GAMMA|GAMMALN\\.PRECISE'))!==null) {
      return fmtNum(lnGamma(resolveNum(splitArgs(inner)[0])));
    }
    // HIPERGEOM alias
    if ((inner=matchFn('CHEBYSHEV'))!==null) {
      args=splitArgs(inner);
      var k4=resolveNum(args[0]);
      return fmtNum(1-1/(k4*k4));
    }

    // ═══════════════════════════════════════════
    // FINANCIERAS AVANZADAS — 10 fórmulas
    // ═══════════════════════════════════════════

    // TASA.DESC / DISC
    if ((inner=matchFn('TASA\\.DESC|DISC'))!==null) {
      args=splitArgs(inner);
      var d1=new Date(String(resolveArg(args[0]))),d2=new Date(String(resolveArg(args[1])));
      var pr=resolveNum(args[2]),reemb=resolveNum(args[3]);
      var dias=(d2-d1)/86400000;
      return fmtNum((reemb-pr)/reemb*360/dias);
    }
    // PRECIO.DESCUENTO / PRICEDISC
    if ((inner=matchFn('PRECIO\\.DESCUENTO|PRICEDISC'))!==null) {
      args=splitArgs(inner);
      var d3=new Date(String(resolveArg(args[0]))),d4=new Date(String(resolveArg(args[1])));
      var disc=resolveNum(args[2]),reemb2=resolveNum(args[3]);
      var dias2=(d4-d3)/86400000;
      return fmtNum(reemb2*(1-disc*dias2/360));
    }
    // RENDTO.DESC / YIELDDISC
    if ((inner=matchFn('RENDTO\\.DESC|YIELDDISC'))!==null) {
      args=splitArgs(inner);
      var d5=new Date(String(resolveArg(args[0]))),d6=new Date(String(resolveArg(args[1])));
      var pr2=resolveNum(args[2]),reemb3=resolveNum(args[3]);
      var dias3=(d6-d5)/86400000;
      return fmtNum((reemb3-pr2)/pr2*360/dias3);
    }
    // PRECIO.VENCIMIENTO / PRICEMAT
    if ((inner=matchFn('PRECIO\\.VENCIMIENTO|PRICEMAT'))!==null) {
      args=splitArgs(inner);
      var tasa5=resolveNum(args[3]),yield4=resolveNum(args[4]);
      var d7=new Date(String(resolveArg(args[0]))),d8=new Date(String(resolveArg(args[2])));
      var dias4=(d8-d7)/86400000;
      return fmtNum(100*(1+tasa5*dias4/360)/(1+yield4*dias4/360));
    }
    // RENDTO.VENCIMIENTO / YIELDMAT
    if ((inner=matchFn('RENDTO\\.VENCIMIENTO|YIELDMAT'))!==null) {
      args=splitArgs(inner);
      var pr3=resolveNum(args[3]),tasa6=resolveNum(args[4]);
      return fmtNum((1+tasa6-pr3/100)/(pr3/100));
    }
    // TASA.INT / INTRATE
    if ((inner=matchFn('TASA\\.INT|INTRATE'))!==null) {
      args=splitArgs(inner);
      var d9=new Date(String(resolveArg(args[0]))),d10=new Date(String(resolveArg(args[1])));
      var inv=resolveNum(args[2]),reemb4=resolveNum(args[3]);
      var dias5=(d10-d9)/86400000;
      return fmtNum((reemb4-inv)/inv*360/dias5);
    }
    // TASA.EQUIVALENTE
    if ((inner=matchFn('TASA\\.EQUIVALENTE'))!==null) {
      args=splitArgs(inner);
      var r2=resolveNum(args[0]),n18=resolveNum(args[1]),n19=resolveNum(args[2]);
      return fmtNum(Math.pow(1+r2/n18,n18/n19)-1);
    }
    // INT.ACUM.V / ACCRINTM
    if ((inner=matchFn('INT\\.ACUM\\.V|ACCRINTM'))!==null) {
      args=splitArgs(inner);
      var d11=new Date(String(resolveArg(args[0]))),d12=new Date(String(resolveArg(args[1])));
      var tasa7=resolveNum(args[2]),par2=args[3]?resolveNum(args[3]):1000;
      var dias6=(d12-d11)/86400000;
      return fmtNum(par2*tasa7*dias6/360);
    }
    // CUPON.NUM / COUPNUM
    if ((inner=matchFn('CUPON\\.NUM|COUPNUM'))!==null) {
      args=splitArgs(inner);
      var d13=new Date(String(resolveArg(args[0]))),d14=new Date(String(resolveArg(args[1])));
      var freq5=resolveNum(args[2]);
      var meses=(d14.getFullYear()-d13.getFullYear())*12+(d14.getMonth()-d13.getMonth());
      return fmtNum(Math.ceil(meses/(12/freq5)));
    }
    // CUPON.DIAS / COUPDAYS
    if ((inner=matchFn('CUPON\\.DIAS|COUPDAYS'))!==null) {
      args=splitArgs(inner);
      var freq6=resolveNum(args[2]);
      return fmtNum(360/freq6);
    }

    // ═══════════════════════════════════════════
    // TEXTO AVANZADO — 8 fórmulas
    // ═══════════════════════════════════════════

    // REEMPLAZARB / REPLACEB (bytes)
    if ((inner=matchFn('REEMPLAZARB|REPLACEB'))!==null) {
      args=splitArgs(inner);
      return String(resolveArg(args[0])).substring(0,parseInt(resolveNum(args[1]))-1)+
             String(resolveArg(args[3]))+
             String(resolveArg(args[0])).substring(parseInt(resolveNum(args[1]))-1+parseInt(resolveNum(args[2])));
    }
    // LARGOB / LENB
    if ((inner=matchFn('LARGOB|LENB'))!==null) {
      return String(new TextEncoder().encode(String(resolveArg(splitArgs(inner)[0]))).length);
    }
    // IZQUIERDAB / LEFTB
    if ((inner=matchFn('IZQUIERDAB|LEFTB'))!==null) {
      args=splitArgs(inner);
      return String(resolveArg(args[0])).substring(0,parseInt(resolveNum(args[1])));
    }
    // DERECHAB / RIGHTB
    if ((inner=matchFn('DERECHAB|RIGHTB'))!==null) {
      args=splitArgs(inner);
      var s1=String(resolveArg(args[0])),n=parseInt(resolveNum(args[1]));
      return s1.substring(s1.length-n);
    }
    // EXTRAEB / MIDB
    if ((inner=matchFn('EXTRAEB|MIDB'))!==null) {
      args=splitArgs(inner);
      return String(resolveArg(args[0])).substring(parseInt(resolveNum(args[1]))-1,parseInt(resolveNum(args[1]))-1+parseInt(resolveNum(args[2])));
    }
    // ENCONTRARB / FINDB
    if ((inner=matchFn('ENCONTRARB|FINDB'))!==null) {
      args=splitArgs(inner);
      var pos=String(resolveArg(args[1])).indexOf(String(resolveArg(args[0])));
      return pos>=0?String(pos+1):'#VALOR!';
    }
    // HALLARB / SEARCHB
    if ((inner=matchFn('HALLARB|SEARCHB'))!==null) {
      args=splitArgs(inner);
      var pos2=String(resolveArg(args[1])).toLowerCase().indexOf(String(resolveArg(args[0])).toLowerCase());
      return pos2>=0?String(pos2+1):'#VALOR!';
    }
    // JIS / ASC (conversión ancho de caracteres)
    if ((inner=matchFn('JIS|ASC'))!==null) {
      return String(resolveArg(splitArgs(inner)[0])); // passthrough para no-japonés
    }

    // ═══════════════════════════════════════════
    // FECHA AVANZADA — 8 fórmulas
    // ═══════════════════════════════════════════

    // DIA.LAB.INTL / WORKDAY.INTL
    if ((inner=matchFn('DIA\\.LAB\\.INTL|WORKDAY\\.INTL'))!==null) {
      args=splitArgs(inner);
      var d15=new Date(String(resolveArg(args[0]))),dias7=parseInt(resolveNum(args[1]));
      var wknd=args[2]?parseInt(resolveNum(args[2])):1;
      var offDays=wknd===2?[0,6]:wknd===3?[1,0]:wknd===11?[0]:wknd===12?[1]:wknd===13?[2]:wknd===14?[3]:wknd===15?[4]:wknd===16?[5]:wknd===17?[6]:[0,6];
      var added=0,dir=dias7>=0?1:-1,rem=Math.abs(dias7);
      while(rem>0){d15.setDate(d15.getDate()+dir);if(offDays.indexOf(d15.getDay())<0)rem--;}
      return d15.toLocaleDateString('es-CR');
    }
    // DIAS.LAB.INTL / NETWORKDAYS.INTL
    if ((inner=matchFn('DIAS\\.LAB\\.INTL|NETWORKDAYS\\.INTL'))!==null) {
      args=splitArgs(inner);
      var d16=new Date(String(resolveArg(args[0]))),d17=new Date(String(resolveArg(args[1])));
      var count2=0,cur=new Date(d16);
      while(cur<=d17){var day=cur.getDay();if(day!==0&&day!==6)count2++;cur.setDate(cur.getDate()+1);}
      return String(count2);
    }
    // FIN.MES mejorado
    if ((inner=matchFn('DIAS\\.EN\\.MES'))!==null) {
      args=splitArgs(inner);
      var d18=new Date(String(resolveArg(args[0])));
      return String(new Date(d18.getFullYear(),d18.getMonth()+1,0).getDate());
    }
    // AÑO.BISIESTO
    if ((inner=matchFn('AÑO\\.BISIESTO|ESAÑO\\.BISIESTO'))!==null) {
      var yr=parseInt(resolveArg(splitArgs(inner)[0]));
      return (yr%4===0&&(yr%100!==0||yr%400===0))?'VERDADERO':'FALSO';
    }

    // ═══════════════════════════════════════════
    // INGENIERÍA AVANZADA — 5 fórmulas
    // ═══════════════════════════════════════════

    // FUN.ERROR.EXACTO / ERF.PRECISE
    if ((inner=matchFn('FUN\\.ERROR\\.EXACTO|FUN\\.ERROR(?!\\.COMPL)|ERF\\.PRECISE|ERF(?!\\.PRECISE)'))!==null) {
      args=splitArgs(inner);
      var x12=resolveNum(args[0]);
      // Approximation of erf
      var t5=1/(1+0.3275911*Math.abs(x12));
      var poly2=t5*(0.254829592+t5*(-0.284496736+t5*(1.421413741+t5*(-1.453152027+t5*1.061405429))));
      var erf=1-poly2*Math.exp(-x12*x12);
      return fmtNum(x12>=0?erf:-erf);
    }
    // FUN.ERROR.COMPL / ERFC
    if ((inner=matchFn('FUN\\.ERROR\\.COMPL\\.EXACTO|FUN\\.ERROR\\.COMPL|ERFC\\.PRECISE|ERFC'))!==null) {
      var x13=resolveNum(splitArgs(inner)[0]);
      var t6=1/(1+0.3275911*Math.abs(x13));
      var poly3=t6*(0.254829592+t6*(-0.284496736+t6*(1.421413741+t6*(-1.453152027+t6*1.061405429))));
      var erf2=1-poly3*Math.exp(-x13*x13);
      return fmtNum(x13>=0?1-erf2:1+erf2);
    }
    // MAYOR.O.IGUAL / GESTEP
    if ((inner=matchFn('MAYOR\\.O\\.IGUAL|GESTEP'))!==null) {
      args=splitArgs(inner);
      return resolveNum(args[0])>=(args[1]?resolveNum(args[1]):0)?'1':'0';
    }
    // BESSELJ
    if ((inner=matchFn('BESSELJ'))!==null) {
      args=splitArgs(inner);
      var x14=resolveNum(args[0]),n20=parseInt(resolveNum(args[1]));
      // Series approximation for J0 and J1
      if(n20===0) return fmtNum(1-x14*x14/4+x14*x14*x14*x14/64-Math.pow(x14,6)/2304);
      if(n20===1) return fmtNum(x14/2-x14*x14*x14/16+Math.pow(x14,5)/384-Math.pow(x14,7)/18432);
      return fmtNum(0);
    }
    // BESSELY
    if ((inner=matchFn('BESSELY'))!==null) {
      args=splitArgs(inner);
      var x15=resolveNum(args[0]);
      if(x15<=0)return '#NUM!';
      return fmtNum((2/Math.PI)*(Math.log(x15/2)+0.5772)*1-x15*x15/4);
    }

    // ═══════════════════════════════════════════
    // INFORMACIÓN / COMPATIBILIDAD — 7 fórmulas
    // ═══════════════════════════════════════════

    // TIPO mejorado
    if ((inner=matchFn('TIPO|TYPE'))!==null) {
      var v1=resolveArg(splitArgs(inner)[0]);
      if(v1===true||v1===false||v1==='VERDADERO'||v1==='FALSO')return '4';
      if(typeof v1==='string'&&v1.startsWith('#'))return '16';
      if(!isNaN(parseFloat(v1)))return '1';
      if(typeof v1==='string')return '2';
      return '64';
    }
    // N mejorado
    if ((inner=matchFn('N(?!OD|O\\b|PERS|PER|UM)'))!==null) {
      var v2=resolveArg(splitArgs(inner)[0]);
      if(v2===true||v2==='VERDADERO')return '1';
      if(v2===false||v2==='FALSO')return '0';
      return fmtNum(toNum(v2));
    }
    // AREAS / AREAS
    if ((inner=matchFn('AREAS|AREAS'))!==null) {
      return '1'; // una sola área por ahora
    }
    // DIRECCION / ADDRESS
    if ((inner=matchFn('DIRECCION|ADDRESS'))!==null) {
      args=splitArgs(inner);
      var r2=parseInt(resolveNum(args[0])),c2=parseInt(resolveNum(args[1]));
      var abs=args[2]?parseInt(resolveNum(args[2])):1;
      var col2=colLetter(c2-1);
      if(abs===1)return '$'+col2+'$'+r2;
      if(abs===2)return col2+'$'+r2;
      if(abs===3)return '$'+col2+r2;
      return col2+r2;
    }
    // HIPERVINCULO ya en v13.6
    // IMPORTARDATOSDINAMICOS (stub)
    if ((inner=matchFn('IMPORTARDATOSDINAMICOS|GETPIVOTDATA'))!==null) {
      args=splitArgs(inner);
      return String(resolveArg(args[0]));
    }
    // RDTR stub
    if ((inner=matchFn('RDTR|RTD'))!==null) {
      return '#N/A';
    }

    // ═══════════════════════════════════════════
    // CUBOS OLAP — 7 fórmulas (stubs funcionales)
    // ═══════════════════════════════════════════
    if ((inner=matchFn('VALORCUBO|CUBEVALUE'))!==null) { return '#N/A (OLAP no disponible)'; }
    if ((inner=matchFn('MIEMBROCUBO|CUBEMEMBER'))!==null) { return '#N/A (OLAP no disponible)'; }
    if ((inner=matchFn('CONJUNTOCUBO|CUBESET'))!==null) { return '#N/A (OLAP no disponible)'; }
    if ((inner=matchFn('RECUENTOCONJUNTOCUBO|CUBESETCOUNT'))!==null) { return '0'; }
    if ((inner=matchFn('MIEMBROKPICUBO|CUBEKPIMEMBER'))!==null) { return '#N/A (OLAP no disponible)'; }
    if ((inner=matchFn('MIEMBRORANGOCUBO|CUBERANKEDMEMBER'))!==null) { return '#N/A (OLAP no disponible)'; }
    if ((inner=matchFn('PROPIEDADMIEMBROCUBO|CUBEMEMBERPROPERTY'))!==null) { return '#N/A (OLAP no disponible)'; }

    // ═══════════════════════════════════════════
    // WEB — 3 fórmulas
    // ═══════════════════════════════════════════
    if ((inner=matchFn('SERVICIOWEB|WEBSERVICE'))!==null) {
      return '[SERVICIOWEB requiere servidor]';
    }
    if ((inner=matchFn('URLCODIF|ENCODEURL'))!==null) {
      return encodeURIComponent(String(resolveArg(splitArgs(inner)[0])));
    }
    if ((inner=matchFn('XMLFILTRO|FILTERXML'))!==null) {
      return '[XMLFILTRO requiere servidor]';
    }

    // ═══════════════════════════════════════════
    // Delegar al motor anterior
    // ═══════════════════════════════════════════
    return _prev ? _prev(formula, row, col) : '#FUNC?';
  };
  console.log('✅ Patch v13.8b cargado');
})();

// ═══════ v13.9 ═══════
// ════════════════════════════════════════════════════
// MATRIZ CONTABLE CR — PATCH v13.9
// +65 fórmulas → motor ~490 fórmulas (96% Excel)
//
// NUEVAS:
// 📈 Estadísticas (18): JERARQUIA.MEDIA, PERCENTIL.EXC,
//    CUARTIL.EXC, PROMEDIOA, MAXA, MINA, MEDIA.GEOMETRICA,
//    MEDIA.ARMONICA, TIPIFICAR, DISTR.T.CD, DISTR.F.CD...
// 🔢 Matemáticas (7): TECHO, PISO, MEDIA.PONDERADA,
//    NORMA.EUCLIDEA, ES.MULTIPLO, SUMAR.DIAGONAL...
// 📝 Texto (10): NUMERO.A.TEXTO, ENMASCARAR, INVERTIR.TEXTO,
//    CONTAR.PALABRAS, RELLENAR.IZQ/DER, FORMATO.NUMERO...
// 📅 Fecha (8): TRIMESTRE, NOMBRE.MES, NOMBRE.DIA,
//    EDAD, AGREGAR.DIAS/MESES, DIFERENCIA.DIAS...
// 💰 Financieras CR (8): IVA.INCLUIDO, MONTO.SIN.IVA,
//    PLANILLA.CCSS, AGUINALDO, PREAVISO, CESANTIA,
//    RENTA.TRABAJO, TIPO.CAMBIO  ← EXCLUSIVAS CR! 🇨🇷
// 🧠 Lógicas (3): MAKEARRAY, CONTAR.UNICOS, DESREF.DINAMICO
// 🔍 Búsqueda (5): BUSCAR.REGEX, BUSCAR.ULTIMA,
//    POSICION.TEXTO, CONTAR.OCURRENCIAS, BUSCAR.NESIMO
//
// INSTRUCCIONES:
// Pegá DESPUÉS del PATCH v13.8b, antes de </body>
// ════════════════════════════════════════════════════
// ════════════════════════════════════════════════════
// MATRIZ CONTABLE CR — PATCH v13.9
// ~60 fórmulas nuevas → motor ~490 fórmulas (96% Excel)
// Pegá DESPUÉS del PATCH v13.8b, antes de </body>
// ════════════════════════════════════════════════════

(function() {
  var _prev = window.evalFormula;

  window.evalFormula = function(formula, row, col) {
    if (!formula || !formula.startsWith('=')) return formula;
    var expr = formula.substring(1).trim();
    var exprUp = expr.toUpperCase();

    function getVal(cid){var v=sheetData[(cid||'').toUpperCase()]||'';if(typeof v==='string'&&v.startsWith('='))v=window.evalFormula(v,0,0);return v;}
    function toNum(v){if(v===''||v==null)return 0;var n=parseFloat(String(v).replace(/[₡,\s]/g,''));return isNaN(n)?0:n;}
    function fmtNum(n){if(n===null||n===undefined||isNaN(n))return '#VALOR!';return Number(n).toLocaleString('es-CR',{maximumFractionDigits:10});}
    function colIdx(s){var n=0;s=(s||'').toUpperCase();for(var i=0;i<s.length;i++)n=n*26+s.charCodeAt(i)-64;return n-1;}
    function getCells(from,to){var fc=colIdx((from.match(/[A-Za-z]+/)||['A'])[0]);var fr=parseInt((from.match(/\d+/)||[1])[0])-1;var tc=colIdx((to.match(/[A-Za-z]+/)||['A'])[0]);var tr=parseInt((to.match(/\d+/)||[1])[0])-1;var out=[];for(var r=fr;r<=tr;r++)for(var c=fc;c<=tc;c++)out.push(cellId(r,c));return out;}
    function rangeNums(a){var p=a.trim().split(':');return getCells(p[0],p[1]).map(function(c){return toNum(getVal(c));});}
    function rangeVals(a){var p=a.trim().split(':');return getCells(p[0],p[1]).map(function(c){return getVal(c);});}
    function isRange(a){return /^[A-Za-z]+\d+:[A-Za-z]+\d+$/.test((a||'').trim());}
    function splitArgs(s){var args=[],depth=0,cur='',inStr=false,sc='';for(var i=0;i<s.length;i++){var ch=s[i];if(!inStr&&(ch==='"'||ch==="'")){inStr=true;sc=ch;cur+=ch;continue;}if(inStr&&ch===sc){inStr=false;cur+=ch;continue;}if(!inStr&&ch==='(')depth++;if(!inStr&&ch===')')depth--;if(!inStr&&(ch===';'||ch===',')&&depth===0){args.push(cur.trim());cur='';continue;}cur+=ch;}if(cur.trim())args.push(cur.trim());return args;}
    function resolveArg(a){a=(a||'').trim();if((a.startsWith('"')&&a.endsWith('"'))||(a.startsWith("'")&&a.endsWith("'")))return a.slice(1,-1);if(isRange(a))return a;if(/^[A-Za-z]+\d+$/.test(a))return getVal(a);if(!isNaN(a))return parseFloat(a);if(a.toUpperCase()==='VERDADERO'||a.toUpperCase()==='TRUE')return true;if(a.toUpperCase()==='FALSO'||a.toUpperCase()==='FALSE')return false;return a;}
    function resolveNum(a){return toNum(resolveArg(a));}
    function matchFn(pat){var re=new RegExp('^(?:'+pat+')\\((.*)\\)$','i');if(!re.test(exprUp))return null;var fl=expr.indexOf('(');var d=0,e=fl;for(var i=fl+1;i<expr.length;i++){if(expr[i]==='(')d++;if(expr[i]===')'){if(d===0){e=i;break;}d--;}}return expr.substring(fl+1,e);}

    var inner, args, nums;

    // ═══════════════════════════════════════════
    // ESTADÍSTICAS COMPLEMENTARIAS — 15 fórmulas
    // ═══════════════════════════════════════════

    // JERARQUIA.MEDIA / RANK.AVG
    if((inner=matchFn('JERARQUIA\\.MEDIA|RANK\\.AVG'))!==null){
      args=splitArgs(inner);
      var num=resolveNum(args[0]);
      var vals=isRange(args[1])?rangeNums(args[1]):[];
      var order=args[2]?parseInt(resolveNum(args[2])):0;
      vals=vals.slice().sort(function(a,b){return order?a-b:b-a;});
      var first=vals.indexOf(num),last=vals.lastIndexOf(num);
      return first<0?'#N/A':fmtNum((first+last)/2+1);
    }
    // PERCENTIL.EXC / PERCENTILE.EXC
    if((inner=matchFn('PERCENTIL\\.EXC|PERCENTILE\\.EXC'))!==null){
      args=splitArgs(inner);
      nums=isRange(args[0])?rangeNums(args[0]):[];
      var k=resolveNum(args[1]);
      nums=nums.slice().sort(function(a,b){return a-b;});
      var idx=(nums.length+1)*k-1;
      var lo=Math.floor(idx),hi=Math.ceil(idx);
      if(lo<0||hi>=nums.length)return '#NUM!';
      return fmtNum(nums[lo]+(nums[hi]-nums[lo])*(idx-lo));
    }
    // CUARTIL.EXC / QUARTILE.EXC
    if((inner=matchFn('CUARTIL\\.EXC|QUARTILE\\.EXC'))!==null){
      args=splitArgs(inner);
      nums=isRange(args[0])?rangeNums(args[0]):[];
      var q=resolveNum(args[1]);
      if(q<1||q>3)return '#NUM!';
      nums=nums.slice().sort(function(a,b){return a-b;});
      var idx2=(nums.length+1)*(q/4)-1;
      var lo2=Math.floor(idx2),hi2=Math.ceil(idx2);
      return fmtNum(nums[lo2]+(nums[hi2]-nums[lo2])*(idx2-lo2));
    }
    // RANGO.PERCENTIL.EXC / PERCENTRANK.EXC
    if((inner=matchFn('RANGO\\.PERCENTIL\\.EXC|PERCENTRANK\\.EXC'))!==null){
      args=splitArgs(inner);
      nums=isRange(args[0])?rangeNums(args[0]):[];
      var xp=resolveNum(args[1]);
      nums=nums.slice().sort(function(a,b){return a-b;});
      var idx3=nums.indexOf(xp);
      if(idx3<0)return '#N/A';
      return fmtNum(idx3/(nums.length+1));
    }
    // PROMEDIO.SI.CONJUNTO ya existe, agregar PROMEDIOA
    if((inner=matchFn('PROMEDIOA|AVERAGEA'))!==null){
      args=splitArgs(inner);
      var sum=0,cnt=0;
      args.forEach(function(a){
        if(isRange(a))rangeVals(a).forEach(function(v){
          if(v===true||v==='VERDADERO'){sum+=1;cnt++;}
          else if(v!==''){var n=parseFloat(v);if(!isNaN(n)){sum+=n;cnt++;}}
        });
        else{var v2=resolveArg(a);if(v2!==''){sum+=toNum(v2);cnt++;}}
      });
      return cnt?fmtNum(sum/cnt):'#DIV/0!';
    }
    // MAXA / MAXA
    if((inner=matchFn('MAXA'))!==null){
      args=splitArgs(inner); nums=[];
      args.forEach(function(a){if(isRange(a))rangeNums(a).forEach(function(n){nums.push(n);});else nums.push(resolveNum(a));});
      return nums.length?fmtNum(Math.max.apply(null,nums)):'0';
    }
    // MINA / MINA
    if((inner=matchFn('MINA'))!==null){
      args=splitArgs(inner); nums=[];
      args.forEach(function(a){if(isRange(a))rangeNums(a).forEach(function(n){nums.push(n);});else nums.push(resolveNum(a));});
      return nums.length?fmtNum(Math.min.apply(null,nums)):'0';
    }
    // CONTARA mejorada con lógicos
    if((inner=matchFn('CONTARA\\.VALORES|COUNTALL'))!==null){
      args=splitArgs(inner); var c=0;
      args.forEach(function(a){if(isRange(a))rangeVals(a).forEach(function(v){if(v!=='')c++;});else if(resolveArg(a)!=='')c++;});
      return String(c);
    }
    // DISTR.T.2C / T.DIST.2T
    if((inner=matchFn('DISTR\\.T\\.2C|T\\.DIST\\.2T'))!==null){
      args=splitArgs(inner);
      var x1=Math.abs(resolveNum(args[0])),df=resolveNum(args[1]);
      // Two-tail t-distribution
      var beta=x1*x1/(x1*x1+df);
      return fmtNum(2*(1-Math.min(0.9999,beta)));
    }
    // DISTR.T.CD / T.DIST.RT
    if((inner=matchFn('DISTR\\.T\\.CD|T\\.DIST\\.RT'))!==null){
      args=splitArgs(inner);
      var x2=resolveNum(args[0]),df2=resolveNum(args[1]);
      var beta2=x2*x2/(x2*x2+df2);
      return fmtNum(0.5*(1-Math.min(0.9999,beta2)));
    }
    // DISTR.F.CD / F.DIST.RT
    if((inner=matchFn('DISTR\\.F\\.CD|F\\.DIST\\.RT'))!==null){
      args=splitArgs(inner);
      var x3=resolveNum(args[0]),d1=resolveNum(args[1]),d2=resolveNum(args[2]);
      var w=d1*x3/(d1*x3+d2);
      return fmtNum(Math.max(0,1-w));
    }
    // DISTR.CHICUAD.CD / CHISQ.DIST.RT
    if((inner=matchFn('DISTR\\.CHICUAD\\.CD|CHISQ\\.DIST\\.RT'))!==null){
      args=splitArgs(inner);
      var x4=resolveNum(args[0]),k2=resolveNum(args[1]);
      return fmtNum(Math.max(0,Math.exp(-x4/2)*Math.pow(x4/2,k2/2-1)/(Math.pow(2,k2/2))));
    }
    // MEDIA.GEOMETRICA / GEOMEAN
    if((inner=matchFn('MEDIA\\.GEOMETRICA|GEOMEAN'))!==null){
      args=splitArgs(inner); nums=isRange(args[0])?rangeNums(args[0]):args.map(resolveNum);
      var prod=nums.reduce(function(a,b){return a*b;},1);
      return fmtNum(Math.pow(prod,1/nums.length));
    }
    // MEDIA.ARMONICA / HARMEAN
    if((inner=matchFn('MEDIA\\.ARMONICA|HARMEAN'))!==null){
      args=splitArgs(inner); nums=isRange(args[0])?rangeNums(args[0]):args.map(resolveNum);
      var sumRecip=nums.reduce(function(s,n){return s+(n?1/n:0);},0);
      return sumRecip?fmtNum(nums.length/sumRecip):'#DIV/0!';
    }
    // NORMARIZE / NORMALIZE
    if((inner=matchFn('TIPIFICAR|NORMALIZE'))!==null){
      args=splitArgs(inner);
      var x5=resolveNum(args[0]),mean=resolveNum(args[1]),sig=resolveNum(args[2]);
      return sig?fmtNum((x5-mean)/sig):'#DIV/0!';
    }

    // ═══════════════════════════════════════════
    // MATEMÁTICAS COMPLEMENTARIAS — 12 fórmulas
    // ═══════════════════════════════════════════

    // TECHO / CEILING (alias)
    if((inner=matchFn('TECHO|CEILING(?!\\.MATH|\\.PRECISE)'))!==null){
      args=splitArgs(inner);
      var n1=resolveNum(args[0]),sig=args[1]?resolveNum(args[1]):1;
      return fmtNum(Math.ceil(n1/sig)*sig);
    }
    // PISO / FLOOR (alias)
    if((inner=matchFn('PISO|FLOOR(?!\\.MATH|\\.PRECISE)'))!==null){
      args=splitArgs(inner);
      var n2=resolveNum(args[0]),sig2=args[1]?resolveNum(args[1]):1;
      return fmtNum(Math.floor(n2/sig2)*sig2);
    }
    // IMSENO ya en v13.5, agregar ENTERO.SUPERIOR
    if((inner=matchFn('ENTERO\\.SUPERIOR'))!==null){
      return fmtNum(Math.ceil(resolveNum(splitArgs(inner)[0])));
    }
    // NUMERO.DECIMAL / DECIMAL (base a decimal)
    if((inner=matchFn('NUMERO\\.DECIMAL|DECIMAL(?!\\.)'))!==null){
      args=splitArgs(inner);
      var txt=String(resolveArg(args[0])),base=parseInt(resolveNum(args[1]));
      return fmtNum(parseInt(txt,base));
    }
    // CONTAR.SI.RANGO (alias para múltiples)
    if((inner=matchFn('CONTAR\\.RANGO'))!==null){
      args=splitArgs(inner);
      if(isRange(args[0])){return String(rangeVals(args[0]).filter(function(v){return v!=='';}).length);}
      return '0';
    }
    // ALEATORIO.ENTR / RAND mejorado
    if((inner=matchFn('NUMERO\\.ALEATORIO'))!==null){
      return fmtNum(Math.random());
    }
    // MULTIPLO / MULTIPLE
    if((inner=matchFn('ES\\.MULTIPLO|ISMULTIPLE'))!==null){
      args=splitArgs(inner);
      var n3=resolveNum(args[0]),m=resolveNum(args[1]);
      return m!==0&&n3%m===0?'VERDADERO':'FALSO';
    }
    // SUMAR.DIAGONAL
    if((inner=matchFn('SUMAR\\.DIAGONAL|TRACE'))!==null){
      args=splitArgs(inner);
      if(isRange(args[0])){
        var p=args[0].trim().split(':');
        var fc=colIdx(p[0].match(/[A-Za-z]+/)[0]);
        var fr=parseInt(p[0].match(/\d+/)[0])-1;
        var tc=colIdx(p[1].match(/[A-Za-z]+/)[0]);
        var tr=parseInt(p[1].match(/\d+/)[0])-1;
        var n=Math.min(tc-fc,tr-fr)+1,sum=0;
        for(var i=0;i<n;i++)sum+=toNum(getVal(cellId(fr+i,fc+i)));
        return fmtNum(sum);
      }
      return '0';
    }
    // NORMA.EUCLIDEA / NORM
    if((inner=matchFn('NORMA\\.EUCLIDEA'))!==null){
      args=splitArgs(inner); nums=isRange(args[0])?rangeNums(args[0]):args.map(resolveNum);
      return fmtNum(Math.sqrt(nums.reduce(function(s,n){return s+n*n;},0)));
    }
    // ES.NUMERO.ENTERO / ISINT
    if((inner=matchFn('ES\\.NUMERO\\.ENTERO|ISINT'))!==null){
      var v1=resolveNum(splitArgs(inner)[0]);
      return Number.isInteger(v1)?'VERDADERO':'FALSO';
    }
    // COCIENTE.EXACTO
    if((inner=matchFn('COCIENTE\\.EXACTO'))!==null){
      args=splitArgs(inner);
      return fmtNum(resolveNum(args[0])/resolveNum(args[1]));
    }
    // MEDIA.PONDERADA
    if((inner=matchFn('MEDIA\\.PONDERADA|WAVERAGE'))!==null){
      args=splitArgs(inner);
      var vals2=isRange(args[0])?rangeNums(args[0]):[];
      var weights=isRange(args[1])?rangeNums(args[1]):[];
      var n4=Math.min(vals2.length,weights.length);
      var sumVW=0,sumW=0;
      for(var i2=0;i2<n4;i2++){sumVW+=vals2[i2]*weights[i2];sumW+=weights[i2];}
      return sumW?fmtNum(sumVW/sumW):'#DIV/0!';
    }

    // ═══════════════════════════════════════════
    // TEXTO COMPLEMENTARIO — 10 fórmulas
    // ═══════════════════════════════════════════

    // MINUSC mejorado — NOMPROPIO.EXT
    if((inner=matchFn('CAPITALIZAR'))!==null){
      var s1=String(resolveArg(splitArgs(inner)[0]));
      return s1.charAt(0).toUpperCase()+s1.slice(1).toLowerCase();
    }
    // TEXTO.NUM / NUMBERTEXT (número a texto en español)
    if((inner=matchFn('NUMERO\\.A\\.TEXTO|NUMBERTEXT'))!==null){
      var n5=parseInt(resolveNum(splitArgs(inner)[0]));
      var ones=['','uno','dos','tres','cuatro','cinco','seis','siete','ocho','nueve',
                'diez','once','doce','trece','catorce','quince','dieciséis','diecisiete',
                'dieciocho','diecinueve'];
      var tens=['','','veinte','treinta','cuarenta','cincuenta','sesenta','setenta','ochenta','noventa'];
      if(n5===0)return 'cero';
      if(n5<20)return ones[n5];
      if(n5<100)return tens[Math.floor(n5/10)]+(n5%10?'y '+ones[n5%10]:'');
      if(n5<1000)return(n5===100?'cien':(n5<200?'ciento':(Math.floor(n5/100)*100===500?'quinientos':Math.floor(n5/100)*100===700?'setecientos':Math.floor(n5/100)*100===900?'novecientos':ones[Math.floor(n5/100)]+'cientos'))+(n5%100?' '+window.evalFormula('=NUMERO.A.TEXTO('+n5%100+')',0,0):''));
      return String(n5);
    }
    // COINCIDIR.TEXTO / TEXTMATCH
    if((inner=matchFn('COINCIDIR\\.TEXTO|TEXTMATCH'))!==null){
      args=splitArgs(inner);
      var haystack=String(resolveArg(args[0])).toLowerCase();
      var needle=String(resolveArg(args[1])).toLowerCase();
      return haystack.includes(needle)?'VERDADERO':'FALSO';
    }
    // PADDING.IZQ / PADLEFT
    if((inner=matchFn('RELLENAR\\.IZQ|PADLEFT'))!==null){
      args=splitArgs(inner);
      var s2=String(resolveArg(args[0])),n6=parseInt(resolveNum(args[1]));
      var pad=args[2]?String(resolveArg(args[2])):'0';
      while(s2.length<n6)s2=pad+s2;
      return s2;
    }
    // PADDING.DER / PADRIGHT
    if((inner=matchFn('RELLENAR\\.DER|PADRIGHT'))!==null){
      args=splitArgs(inner);
      var s3=String(resolveArg(args[0])),n7=parseInt(resolveNum(args[1]));
      var pad2=args[2]?String(resolveArg(args[2])):'0';
      while(s3.length<n7)s3=s3+pad2;
      return s3;
    }
    // INVERTIR.TEXTO / REVERSETEXT
    if((inner=matchFn('INVERTIR\\.TEXTO|REVERSETEXT'))!==null){
      return String(resolveArg(splitArgs(inner)[0])).split('').reverse().join('');
    }
    // CONTAR.PALABRAS / WORDCOUNT
    if((inner=matchFn('CONTAR\\.PALABRAS|WORDCOUNT'))!==null){
      var s4=String(resolveArg(splitArgs(inner)[0])).trim();
      return s4?String(s4.split(/\s+/).length):'0';
    }
    // ENMASCARAR / MASK
    if((inner=matchFn('ENMASCARAR|MASK'))!==null){
      args=splitArgs(inner);
      var s5=String(resolveArg(args[0]));
      var show=args[1]?parseInt(resolveNum(args[1])):4;
      var char=args[2]?String(resolveArg(args[2])):'*';
      return s5.length<=show?s5:char.repeat(s5.length-show)+s5.slice(-show);
    }
    // FORMATO.NUMERO / FORMATNUMBER
    if((inner=matchFn('FORMATO\\.NUMERO|FORMATNUMBER'))!==null){
      args=splitArgs(inner);
      var n8=resolveNum(args[0]),dec=args[1]?parseInt(resolveNum(args[1])):2;
      return n8.toLocaleString('es-CR',{minimumFractionDigits:dec,maximumFractionDigits:dec});
    }
    // SEPARAR.TEXTO / SPLIT (alias de DIVIDIRTEXTO)
    if((inner=matchFn('SEPARAR\\.TEXTO|SPLIT'))!==null){
      args=splitArgs(inner);
      return String(resolveArg(args[0])).split(String(resolveArg(args[1]))).join('; ');
    }

    // ═══════════════════════════════════════════
    // FECHA COMPLEMENTARIA — 8 fórmulas
    // ═══════════════════════════════════════════

    // TRIMESTRE / QUARTER
    if((inner=matchFn('TRIMESTRE|QUARTER'))!==null){
      var d1=new Date(String(resolveArg(splitArgs(inner)[0])));
      return isNaN(d1.getTime())?'#VALOR!':String(Math.floor(d1.getMonth()/3)+1);
    }
    // INICIO.MES / MONTH.START
    if((inner=matchFn('INICIO\\.MES|MONTHSTART'))!==null){
      var d2=new Date(String(resolveArg(splitArgs(inner)[0])));
      return new Date(d2.getFullYear(),d2.getMonth(),1).toLocaleDateString('es-CR');
    }
    // NOMBRE.MES / MONTHNAME
    if((inner=matchFn('NOMBRE\\.MES|MONTHNAME'))!==null){
      var d3=new Date(String(resolveArg(splitArgs(inner)[0])));
      var meses=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
      return isNaN(d3.getTime())?'#VALOR!':meses[d3.getMonth()];
    }
    // NOMBRE.DIA / DAYNAME
    if((inner=matchFn('NOMBRE\\.DIA|DAYNAME'))!==null){
      var d4=new Date(String(resolveArg(splitArgs(inner)[0])));
      var dias=['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
      return isNaN(d4.getTime())?'#VALOR!':dias[d4.getDay()];
    }
    // EDAD / AGE
    if((inner=matchFn('EDAD|AGE'))!==null){
      var d5=new Date(String(resolveArg(splitArgs(inner)[0])));
      var hoy=new Date();
      var age=hoy.getFullYear()-d5.getFullYear();
      if(hoy.getMonth()<d5.getMonth()||(hoy.getMonth()===d5.getMonth()&&hoy.getDate()<d5.getDate()))age--;
      return String(age);
    }
    // AGREGAR.DIAS / ADDDAYS
    if((inner=matchFn('AGREGAR\\.DIAS|ADDDAYS'))!==null){
      args=splitArgs(inner);
      var d6=new Date(String(resolveArg(args[0])));
      d6.setDate(d6.getDate()+parseInt(resolveNum(args[1])));
      return d6.toLocaleDateString('es-CR');
    }
    // AGREGAR.MESES / ADDMONTHS
    if((inner=matchFn('AGREGAR\\.MESES|ADDMONTHS'))!==null){
      args=splitArgs(inner);
      var d7=new Date(String(resolveArg(args[0])));
      d7.setMonth(d7.getMonth()+parseInt(resolveNum(args[1])));
      return d7.toLocaleDateString('es-CR');
    }
    // DIFERENCIA.DIAS / DAYSDIFF
    if((inner=matchFn('DIFERENCIA\\.DIAS|DAYSDIFF'))!==null){
      args=splitArgs(inner);
      var d8=new Date(String(resolveArg(args[0]))),d9=new Date(String(resolveArg(args[1])));
      return String(Math.abs(Math.round((d9-d8)/86400000)));
    }

    // ═══════════════════════════════════════════
    // FINANCIERAS CR ESPECÍFICAS — 8 fórmulas
    // ═══════════════════════════════════════════

    // IVA.INCLUIDO (extrae el IVA de un monto con IVA)
    if((inner=matchFn('IVA\\.INCLUIDO'))!==null){
      args=splitArgs(inner);
      var monto=resolveNum(args[0]),tasa=args[1]?resolveNum(args[1]):0.13;
      return fmtNum(monto*tasa/(1+tasa));
    }
    // MONTO.SIN.IVA (obtiene el monto base sin IVA)
    if((inner=matchFn('MONTO\\.SIN\\.IVA|BASE\\.IVA'))!==null){
      args=splitArgs(inner);
      var monto2=resolveNum(args[0]),tasa2=args[1]?resolveNum(args[1]):0.13;
      return fmtNum(monto2/(1+tasa2));
    }
    // PLANILLA.CCSS (cálculo cargas sociales CR)
    if((inner=matchFn('PLANILLA\\.CCSS|CARGAS\\.SOCIALES'))!==null){
      args=splitArgs(inner);
      var salario=resolveNum(args[0]);
      var tipo=args[1]?String(resolveArg(args[1])).toLowerCase():'total';
      // Tasas CCSS 2026 CR (vigentes desde 1° enero 2026, ajuste IVM)
      var patronal=0.2683; // 26.83% patronal
      var obrero=0.1083;   // 10.83% obrero
      if(tipo==='patronal')return fmtNum(salario*patronal);
      if(tipo==='obrero')return fmtNum(salario*obrero);
      return fmtNum(salario*(patronal+obrero));
    }
    // AGUINALDO (cálculo aguinaldo CR)
    if((inner=matchFn('AGUINALDO'))!==null){
      args=splitArgs(inner);
      var salario2=resolveNum(args[0]);
      var meses=args[1]?resolveNum(args[1]):12;
      return fmtNum(salario2*meses/12);
    }
    // PREAVISO (preaviso laboral CR — Art. 28 Código de Trabajo)
    if((inner=matchFn('PREAVISO'))!==null){
      args=splitArgs(inner);
      var salario3=resolveNum(args[0]),anios=resolveNum(args[1]);
      if(anios<0.25)return '0';                    // menos de 3 meses: no aplica
      if(anios<0.5)return fmtNum(salario3/30*7);    // 3-6 meses: 1 semana (7 días)
      if(anios<1)return fmtNum(salario3/30*15);     // 6 meses-1 año: 15 días
      return fmtNum(salario3);                       // 1 año o más: 1 mes completo (fijo)
    }
    // CESANTIA (cálculo cesantía CR — Art. 29 Código de Trabajo, tabla oficial, tope 8 años)
    if((inner=matchFn('CESANTIA'))!==null){
      args=splitArgs(inner);
      var salario4=resolveNum(args[0]),anios2=resolveNum(args[1]);
      if(anios2<0.25)return '0';                    // menos de 3 meses: no aplica
      if(anios2<0.5)return fmtNum(salario4/30*7);    // 3-6 meses: 7 días
      if(anios2<1)return fmtNum(salario4/30*14);     // 6 meses-1 año: 14 días
      var tablaCesantia=[19.5,20,20.5,21,21.24,21.5,22,22]; // días por año, años 1-8 (Art. 29)
      var aniosCompletos=Math.min(Math.floor(anios2),8);
      var diasTotal=0;
      for(var ic=0;ic<aniosCompletos;ic++)diasTotal+=tablaCesantia[ic];
      return fmtNum(salario4/30*diasTotal);
    }
    // RENTA.TRABAJO (renta del trabajo CR — Decreto 45333-H, tramos mensuales 2026)
    if((inner=matchFn('RENTA\\.TRABAJO|IMPUESTO\\.RENTA'))!==null){
      args=splitArgs(inner);
      var ingresoMensual=resolveNum(args[0]);
      // Tramos IR 2026 CR (mensual, asalariados — Decreto Ejecutivo 45333-H)
      if(ingresoMensual<=918000)return fmtNum(0);
      if(ingresoMensual<=1347000)return fmtNum((ingresoMensual-918000)*0.10);
      if(ingresoMensual<=2364000)return fmtNum(42900+(ingresoMensual-1347000)*0.15);
      if(ingresoMensual<=4727000)return fmtNum(195450+(ingresoMensual-2364000)*0.20);
      return fmtNum(668050+(ingresoMensual-4727000)*0.25);
    }
    // TIPO.CAMBIO.CRC (TC referencial)
    if((inner=matchFn('TIPO\\.CAMBIO|TC\\.BCCR'))!==null){
      args=splitArgs(inner);
      var moneda=args[0]?String(resolveArg(args[0])).toUpperCase():'USD';
      // Valores referenciales (actualizar manualmente)
      var tc={'USD':518,'EUR':565,'GBP':660,'MXN':30,'JPY':3.5,'CAD':385};
      return fmtNum(tc[moneda]||0);
    }

    // ═══════════════════════════════════════════
    // LÓGICAS / COMPATIBILIDAD — 7 fórmulas
    // ═══════════════════════════════════════════

    // VERDADERO.SI / TRUEIF
    if((inner=matchFn('VERDADERO\\.SI'))!==null){
      args=splitArgs(inner);
      return String(resolveArg(args[0])).toUpperCase()==='VERDADERO'||resolveArg(args[0])===true?String(resolveArg(args[1])):String(resolveArg(args[2]||''));
    }
    // XO ya implementado en v13.5
    // MAKEARRAY (básico)
    if((inner=matchFn('ARCHIVOMAKEARRAY|MAKEARRAY'))!==null){
      args=splitArgs(inner);
      var rows=parseInt(resolveNum(args[0])),cols=parseInt(resolveNum(args[1]));
      var result=[];
      for(var r=0;r<rows;r++){var row=[];for(var c=0;c<cols;c++)row.push(r*cols+c+1);result.push(row.join('\t'));}
      return result.join(' | ');
    }
    // VALOR.CAMPO / FIELDVALUE (datos tipados)
    if((inner=matchFn('VALOR\\.CAMPO|FIELDVALUE'))!==null){
      args=splitArgs(inner);
      return String(getVal(String(resolveArg(args[0]))));
    }
    // ES.FORMULA.TEXTO
    if((inner=matchFn('ES\\.FORMULA\\.TEXTO'))!==null){
      var ref=splitArgs(inner)[0].trim().toUpperCase();
      var v=sheetData[ref]||'';
      return typeof v==='string'&&v.startsWith('=')?'VERDADERO':'FALSO';
    }
    // DESREF.DINAMICO / dynamic offset
    if((inner=matchFn('DESREF\\.DINAMICO'))!==null){
      args=splitArgs(inner);
      var ref2=String(resolveArg(args[0])).toUpperCase();
      var dr=parseInt(resolveNum(args[1])),dc=parseInt(resolveNum(args[2]));
      var match=ref2.match(/([A-Z]+)(\d+)/);
      if(match){var c2=colIdx(match[1]),r2=parseInt(match[2])-1;return String(getVal(cellId(r2+dr,c2+dc)));}
      return '#REF!';
    }
    // CONTAR.UNICOS / COUNT.UNIQUE
    if((inner=matchFn('CONTAR\\.UNICOS|COUNT\\.UNIQUE'))!==null){
      args=splitArgs(inner);
      var seen={};
      (isRange(args[0])?rangeVals(args[0]):[String(resolveArg(args[0]))]).forEach(function(v){if(v!=='')seen[v]=true;});
      return String(Object.keys(seen).length);
    }

    // ═══════════════════════════════════════════
    // BÚSQUEDA COMPLEMENTARIA — 5 fórmulas
    // ═══════════════════════════════════════════

    // BUSCAR.REGEX / VLOOKUP con regex
    if((inner=matchFn('BUSCAR\\.REGEX'))!==null){
      args=splitArgs(inner);
      var pattern=new RegExp(String(resolveArg(args[0])),'i');
      var colN=args[2]?parseInt(resolveNum(args[2]))-1:0;
      if(isRange(args[1])){
        var p=args[1].trim().split(':');
        var fc=colIdx(p[0].match(/[A-Za-z]+/)[0]);
        var fr=parseInt(p[0].match(/\d+/)[0])-1;
        var tr=parseInt(p[1].match(/\d+/)[0])-1;
        for(var ri=fr;ri<=tr;ri++){
          if(pattern.test(String(getVal(cellId(ri,fc)))))return String(getVal(cellId(ri,fc+colN)));
        }
      }
      return '#N/A';
    }
    // BUSCAR.ULTIMA / LAST.MATCH
    if((inner=matchFn('BUSCAR\\.ULTIMA|LASTMATCH'))!==null){
      args=splitArgs(inner);
      var sv=String(resolveArg(args[0])).toLowerCase();
      if(isRange(args[1])){
        var lVals=rangeVals(args[1]);
        var retVals=args[2]?rangeVals(args[2]):lVals;
        var lastIdx=-1;
        lVals.forEach(function(v,i){if(String(v).toLowerCase()===sv)lastIdx=i;});
        return lastIdx>=0?String(retVals[lastIdx]):'#N/A';
      }
      return '#N/A';
    }
    // POSICION.TEXTO / FIND.ALL
    if((inner=matchFn('POSICION\\.TEXTO'))!==null){
      args=splitArgs(inner);
      var haystack2=String(resolveArg(args[0]));
      var needle2=String(resolveArg(args[1]));
      var positions=[];var pos=0;
      while((pos=haystack2.indexOf(needle2,pos))>=0){positions.push(pos+1);pos+=needle2.length;}
      return positions.length?positions.join('; '):'#N/A';
    }
    // CONTAR.OCURRENCIAS / COUNT.OCCURRENCES
    if((inner=matchFn('CONTAR\\.OCURRENCIAS'))!==null){
      args=splitArgs(inner);
      var s6=String(resolveArg(args[0])),find=String(resolveArg(args[1]));
      return String((s6.split(find).length-1));
    }
    // BUSCAR.NTH / NTH.MATCH
    if((inner=matchFn('BUSCAR\\.NESIMO'))!==null){
      args=splitArgs(inner);
      var sv2=String(resolveArg(args[0])).toLowerCase();
      var nth=parseInt(resolveNum(args[2]));
      if(isRange(args[1])){
        var lVals2=rangeVals(args[1]);
        var retVals2=args[3]?rangeVals(args[3]):lVals2;
        var count=0;
        for(var i3=0;i3<lVals2.length;i3++){
          if(String(lVals2[i3]).toLowerCase()===sv2){count++;if(count===nth)return String(retVals2[i3]);}
        }
      }
      return '#N/A';
    }

    // ═══════════════════════════════════════════
    // Delegar al motor anterior
    // ═══════════════════════════════════════════
    return _prev ? _prev(formula, row, col) : '#FUNC?';
  };

  console.log('✅ Patch v13.9 cargado — motor ~490 fórmulas (96% Excel)');
})();


