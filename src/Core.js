/* Pure functions. No Google service calls and no formula evaluation with eval(). */
var LH = LH || {};
LH.Core = (function () {
  'use strict';
  function assert(ok, message) { if (!ok) throw new Error(message); }
  function column(n) {
    var s = '';
    for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
    return s;
  }
  function columnNumber(s) { return s.toUpperCase().split('').reduce(function (n, c) { return n * 26 + c.charCodeAt(0) - 64; }, 0); }
  function quote(s) { return "'" + s.replace(/'/g, "''") + "'"; }
  function spreadsheetId(s) {
    var m = String(s).match(/\/spreadsheets\/d\/([\w-]+)/);
    if (m) return m[1];
    assert(/^[\w-]{20,}$/.test(String(s)), 'Expected spreadsheet ID or URL, received: ' + s);
    return String(s);
  }
  function splitOutside(s, separator) {
    var result = [], start = 0, depth = 0, inString = false, inSheet = false;
    for (var i = 0; i < s.length; i++) {
      var c = s[i];
      if (c === '"' && !inSheet) {
        if (inString && s[i + 1] === '"') { i++; continue; }
        inString = !inString; continue;
      }
      if (c === "'" && !inString) {
        if (inSheet && s[i + 1] === "'") { i++; continue; }
        inSheet = !inSheet; continue;
      }
      if (inString || inSheet) continue;
      if (c === '(') depth++;
      if (c === ')') depth--;
      if (c === separator && depth === 0) { result.push(s.slice(start, i)); start = i + 1; }
    }
    result.push(s.slice(start)); return result;
  }
  function imports(formula) {
    var found = [], inString = false;
    for (var i = 0; i < formula.length; i++) {
      if (formula[i] === '"') {
        if (inString && formula[i + 1] === '"') { i++; continue; }
        inString = !inString; continue;
      }
      if (inString || !/^IMPORTRANGE\s*\(/i.test(formula.slice(i)) || (i && /[\w.]/.test(formula[i - 1]))) continue;
      var start = i, open = formula.indexOf('(', i), level = 1, quoted = false, j = open + 1;
      for (; j < formula.length && level; j++) {
        if (formula[j] === '"') {
          if (quoted && formula[j + 1] === '"') { j++; continue; }
          quoted = !quoted;
        } else if (!quoted) {
          if (formula[j] === '(') level++;
          if (formula[j] === ')') level--;
        }
      }
      assert(level === 0, 'Unclosed IMPORTRANGE call');
      var args = splitOutside(formula.slice(open + 1, j - 1), ',');
      if (args.length !== 2) args = splitOutside(formula.slice(open + 1, j - 1), ';');
      assert(args.length === 2, 'IMPORTRANGE must have two arguments');
      found.push({ start: start, end: j, source: args[0].trim(), range: args[1].trim() }); i = j - 1;
    }
    return found;
  }
  function expression(s, resolveReference) {
    return splitOutside(s, '&').map(function (part) {
      part = part.trim();
      if (/^"(?:[^"]|"")*"$/.test(part)) return part.slice(1, -1).replace(/""/g, '"');
      assert(/^(?:(?:'(?:[^']|'')+'|[\w ]+)!\s*)?\$?[A-Z]+\$?\d+$/i.test(part), 'Unsupported import argument: ' + part);
      return String(resolveReference(part));
    }).join('');
  }
  function reference(s, currentSheet) {
    var m = s.match(/^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?(\$?[A-Z]+\$?\d+)$/i);
    assert(m, 'Unsupported cell reference ' + s);
    return { sheet: m[1] ? m[1].replace(/''/g, "'") : (m[2] || currentSheet), cell: m[3].replace(/\$/g, '').toUpperCase() };
  }
  function range(s) {
    var bang = s.lastIndexOf('!'); assert(bang > 0, 'Import range needs an explicit tab: ' + s);
    var tab = s.slice(0, bang), body = s.slice(bang + 1).replace(/\$/g, '').toUpperCase();
    if (tab[0] === "'" && tab[tab.length - 1] === "'") tab = tab.slice(1, -1).replace(/''/g, "'");
    var m = body.match(/^([A-Z]+)(\d+)?(?::([A-Z]+)(\d+)?)?$/);
    assert(m, 'Only A1 column/cell ranges are supported: ' + s);
    var c1 = columnNumber(m[1]), c2 = columnNumber(m[3] || m[1]);
    var r1 = Number(m[2] || 1), r2 = m[3] ? (m[4] ? Number(m[4]) : null) : (m[2] ? Number(m[2]) : null);
    assert(c2 >= c1 && (!r2 || r2 >= r1), 'Reversed range: ' + s);
    return { tab: tab, c1: c1, c2: c2, r1: r1, r2: r2 };
  }
  function compile(formula, calls, resolve) {
    var out = formula;
    calls.slice().reverse().forEach(function (c) { out = out.slice(0, c.start) + resolve(c) + out.slice(c.end); });
    assert(imports(out).length === 0, 'Compilation left an external import'); return out;
  }
  function localRange(dataset, r) {
    var first = dataset.columns.indexOf(r.c1), last = dataset.columns.indexOf(r.c2);
    assert(first >= 0 && last - first === r.c2 - r.c1, 'Non-contiguous cache projection');
    return quote(dataset.cache) + '!' + column(first + 1) + r.r1 + ':' + column(last + 1) + (r.r2 || dataset.rows);
  }
  function latestTab(names, today) {
    var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var best = null, key = '';
    names.forEach(function (s) {
      var m = s.match(/^(\d{1,2}) ([A-Z][a-z]{2}) (\d{2})$/);
      if (!m || months.indexOf(m[2]) < 0) return;
      var y = 2000 + Number(m[3]), mon = months.indexOf(m[2]), day = Number(m[1]);
      var d = new Date(Date.UTC(y, mon, day));
      if (d.getUTCMonth() !== mon || d.getUTCDate() !== day) return;
      var k = d.toISOString().slice(0, 10);
      if (k <= today && k > key) { key = k; best = s; }
    });
    assert(best, 'No dated schedule tab at or before ' + today); return best;
  }
  function rectangular(rows, width, height) {
    return Array.from({ length: height }, function (_, i) { return Array.from({ length: width }, function (_, j) { var v = (rows[i] || [])[j]; return v === undefined || v === null ? '' : v; }); });
  }
  function checkErrors(rows) {
    rows.forEach(function (r, i) { r.forEach(function (v, j) {
      assert(!(typeof v === 'string' && /^#(?:REF!|N\/A$|DIV\/0!|VALUE!|ERROR!|NAME\?|NUM!)/.test(v)), 'Source error at row ' + (i + 1) + ', col ' + (j + 1) + ': ' + v);
    }); });
  }
  return { assert: assert, column: column, columnNumber: columnNumber, quote: quote, spreadsheetId: spreadsheetId,
    splitOutside: splitOutside, imports: imports, expression: expression, reference: reference, range: range,
    compile: compile, localRange: localRange, latestTab: latestTab, rectangular: rectangular, checkErrors: checkErrors };
}());

