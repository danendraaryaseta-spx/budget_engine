var LH = LH || {};
LH.Planner = (function () {
  'use strict';
  function discoverStep(hub, engine, state, deadline) {
    var sheets = engine.getSheets().filter(function (s) { return !/^__(?:src|stage)_/.test(s.getName()); });
    var batch = [];
    while (state.sheet < sheets.length && Date.now() < deadline) {
      var sh = sheets[state.sheet], width = Math.max(1, sh.getLastColumn()), last = sh.getLastRow();
      if (!last || state.row > last) { state.sheet++; state.row = 1; continue; }
      var count = Math.min(last - state.row + 1, Math.max(1, Math.floor(80000 / width)));
      sh.getRange(state.row, 1, count, width).getFormulas().forEach(function (r, i) { r.forEach(function (f, j) {
        if (f && LH.Core.imports(f).length) batch.push({ key: sh.getName() + '!' + LH.Core.column(j + 1) + (state.row + i), sheet: sh.getName(), cell: LH.Core.column(j + 1) + (state.row + i), formula: f, expected: f,
          registry: sh.getName() === LH.Config.registrySheet && j + 1 === LH.Config.registryColumn && state.row + i >= LH.Config.registryFirstRow && state.row + i <= LH.Config.registryLastRow });
      }); });
      state.row += count;
      if (batch.length) {
        var out = LH.IO.ensure(hub, '__Formulas', 2, 2);
        LH.IO.raw(hub.getId(), "'__Formulas'!A" + (out.getLastRow() + 1), batch.map(function (x) { return [x.key, JSON.stringify(x)]; })); batch = [];
      }
      LH.saveState(state);
    }
    if (state.sheet >= sheets.length) { state.phase = 'PLAN'; LH.saveState(state); }
  }
  function build(hub, engine, settings) {
    var C = LH.Core, books = {}, sheets = {}, refs = {}, datasets = {}, templates = LH.IO.readTable(hub, '__Formulas');
    var overrides = JSON.parse(settings.SOURCE_OVERRIDES || '{}');
    function book(id) { return books[id] || (books[id] = SpreadsheetApp.openById(id)); }
    function sourceSheet(id, tab) {
      var key = id + '|' + tab;
      if (!sheets[key]) { sheets[key] = book(id).getSheetByName(tab); C.assert(sheets[key], 'Missing source tab: ' + tab + ' in ' + id); }
      return sheets[key];
    }
    function ref(s, tab) {
      var r = C.reference(s, tab), key = r.sheet + '!' + r.cell;
      if (overrides[key]) return overrides[key];
      if (refs[key] !== undefined) return refs[key];
      var sh = engine.getSheetByName(r.sheet); C.assert(sh, 'Missing reference tab: ' + r.sheet);
      var cell = sh.getRange(r.cell), v = cell.getValue();
      var data = Sheets.Spreadsheets.get(engine.getId(), { ranges: [C.quote(r.sheet) + '!' + r.cell], fields: 'sheets.data.rowData.values.chipRuns' });
      (((((data.sheets || [])[0] || {}).data || [])[0] || {}).rowData || []).forEach(function (row) { (row.values || []).forEach(function (c) {
        (c.chipRuns || []).forEach(function (run) { var uri = (((run.chip || {}).richLinkProperties || {}).uri); if (uri) v = uri; });
      }); });
      var rich = cell.getRichTextValue(); if (rich && rich.getLinkUrl()) v = rich.getLinkUrl();
      C.assert(!(typeof v === 'string' && /^#/.test(v)), 'Import argument has an error at ' + key);
      refs[key] = v; return v;
    }
    function add(id, r, feed) {
      C.assert(id !== engine.getId() && id !== hub.getId(), 'Upstream sources cannot be the engine or hub');
      var key = id + '|' + r.tab, sh = sourceSheet(id, r.tab);
      C.assert(r.c2 <= sh.getMaxColumns() && (!r.r2 || r.r2 <= sh.getMaxRows()), 'Source range exceeds grid: ' + r.tab);
      if (!datasets[key]) datasets[key] = { key: key, sourceId: id, tab: r.tab, columns: [], rows: 1, cache: LH.Config.cachePrefix + LH.IO.hash(key).slice(0, 12), feeds: [] };
      var d = datasets[key];
      for (var c = r.c1; c <= r.c2; c++) if (d.columns.indexOf(c) < 0) d.columns.push(c);
      d.rows = Math.max(d.rows, r.r1, r.r2 || sh.getLastRow());
      if (feed) d.feeds.push(feed);
      return key;
    }
    templates.forEach(function (t) {
      var current = engine.getSheetByName(t.sheet).getRange(t.cell).getFormula();
      C.assert(current === t.expected, 'Formula changed after capture: ' + t.key + '. Review the template before refreshing.');
      if (t.registry) { t.compiled = '="Source connection managed by Apps Script"'; return; }
      t.calls = C.imports(t.formula).map(function (call) {
        call.sourceId = C.spreadsheetId(C.expression(call.source, function (s) { return ref(s, t.sheet); }));
        call.parsed = C.range(C.expression(call.range, function (s) { return ref(s, t.sheet); }));
        call.dataset = add(call.sourceId, call.parsed); return call;
      });
    });
    LH.Config.feeds.forEach(function (feed) {
      var f = JSON.parse(JSON.stringify(feed));
      var id = C.spreadsheetId(f.sourceSetting ? settings[f.sourceSetting] : ref(f.sourceCell, 'Base Assumptions'));
      if (f.latestDated) f.tab = C.latestTab(book(id).getSheets().map(function (s) { return s.getName(); }), settings.RUN_DATE);
      var sh = sourceSheet(id, f.tab), r = C.range(C.quote(f.tab) + '!' + f.range);
      if (f.columnsByHeader) {
        var headers = sh.getRange(f.sourceHeader, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(function (x) { return x.trim(); });
        f.sourceColumns = f.columnsByHeader.map(function (name) { C.assert(headers.filter(function (x) { return x === name; }).length === 1, 'Missing or duplicate header ' + name + ' in ' + f.tab); return headers.indexOf(name) + 1; });
        r.c1 = Math.min.apply(null, f.sourceColumns); r.c2 = Math.max.apply(null, f.sourceColumns);
      } else f.sourceColumns = Array.from({ length: r.c2 - r.c1 + 1 }, function (_, i) { return r.c1 + i; });
      f.firstRow = r.r1; f.rows = Math.max(0, sh.getLastRow() - r.r1 + 1);
      C.assert(f.rows > 0, 'Empty required feed ' + f.key);
      if (f.header) {
        var sourceHeader = sh.getRange(f.header, r.c1, 1, f.sourceColumns.length).getDisplayValues()[0];
        var destHeader = engine.getSheetByName(f.dest).getRange(f.header, f.col, 1, f.sourceColumns.length).getDisplayValues()[0];
        C.assert(JSON.stringify(sourceHeader.map(function (s) { return s.trim(); })) === JSON.stringify(destHeader.map(function (s) { return s.trim(); })), 'Header mismatch: ' + f.key);
      }
      add(id, r, f);
    });
    var list = Object.keys(datasets).map(function (k) {
      var d = datasets[k]; d.columns.sort(function (a, b) { return a - b; });
      C.assert(d.rows * d.columns.length <= LH.Config.maxDatasetCells, 'Source is too large: ' + d.tab + '. Reduce the source range.');
      d.stage = LH.Config.stagePrefix + d.cache.slice(LH.Config.cachePrefix.length); return d;
    });
    templates.forEach(function (t) { if (!t.registry) t.compiled = C.compile(t.formula, t.calls, function (call) { return C.localRange(datasets[call.dataset], call.parsed); }); });
    return { datasets: list, formulas: templates };
  }
  return { discoverStep: discoverStep, build: build };
}());
