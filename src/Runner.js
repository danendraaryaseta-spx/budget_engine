var LH = LH || {};
LH.saveState = function (s) { PropertiesService.getScriptProperties().setProperty('JOB', JSON.stringify(s)); };
LH.state = function () { return JSON.parse(PropertiesService.getScriptProperties().getProperty('JOB') || 'null'); };
LH.settings = function (hub) {
  var sh = hub.getSheetByName('Control'), out = {}; LH.Core.assert(sh, 'Run setup first');
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { out[r[0]] = String(r[1] || ''); }); return out;
};
LH.context = function () {
  var id = PropertiesService.getScriptProperties().getProperty('HUB_ID');
  var hub = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  LH.Core.assert(hub, 'Run setup from the source hub');
  var settings = LH.settings(hub), engineId = LH.Core.spreadsheetId(settings.ENGINE_ID);
  LH.Core.assert(engineId !== hub.getId() && engineId !== LH.Core.spreadsheetId(settings.TEMPLATE_ID), 'Use an engine copy, separate from the hub and original template');
  return { hub: hub, engine: SpreadsheetApp.openById(engineId), settings: settings };
};
LH.log = function (hub, s, message) {
  var sh = LH.IO.ensure(hub, 'Refresh Log', 2, 6);
  if (!sh.getRange('A1').getValue()) LH.IO.raw(hub.getId(), "'Refresh Log'!A1", [['Timestamp (WIB)','Run','Phase','Dataset','Next row','Message']]);
  LH.IO.raw(hub.getId(), "'Refresh Log'!A" + (sh.getLastRow() + 1), [[Utilities.formatDate(new Date(), LH.Config.timezone, 'yyyy-MM-dd HH:mm:ss'),s.id,s.phase,s.dataset || 0,s.nextRow || 1,message]]);
};
function setup() {
  var hub = SpreadsheetApp.getActiveSpreadsheet(); LH.Core.assert(hub, 'Run from the bound script');
  PropertiesService.getScriptProperties().setProperty('HUB_ID', hub.getId()); hub.setSpreadsheetTimeZone(LH.Config.timezone);
  if (!hub.getSheetByName('Control')) {
    var sh = LH.IO.ensure(hub, 'Control', 12, 3);
    LH.IO.raw(hub.getId(), "'Control'!A1", [
      ['Setting','Value','Description'],
      ['TEMPLATE_ID','','Original engine, preserved'],
      ['ENGINE_ID','','Working COPY; populated by createEngineCopy'],
      ['LM_SOURCE_ID','','Source workbook containing Final Proj (Budget)'],
      ['SOURCE_OVERRIDES','{}','JSON: Sheet!Cell => upstream ID when a chip cannot resolve'],
      ['STATUS','NOT INITIALIZED','Refresh status'],
      ['LAST_PUBLISHED_RUN','','Must match the engine __Source_Status'],
      ['LAST_PUBLISHED_AT','','WIB timestamp'],
      ['SOURCE_DATE','','Optional YYYY-MM-DD for dated schedule selection; blank = today']
    ]);
    sh.setFrozenRows(1); sh.getRange('A1:C1').setFontWeight('bold').setBackground('#17365D').setFontColor('#FFFFFF');
    sh.setColumnWidth(1,180); sh.setColumnWidth(2,450); sh.setColumnWidth(3,600);
  }
}
function createEngineCopy() {
  var hub = SpreadsheetApp.getActiveSpreadsheet(), v = LH.settings(hub); LH.Core.assert(!v.ENGINE_ID, 'Engine already configured');
  var copy = DriveApp.getFileById(LH.Core.spreadsheetId(v.TEMPLATE_ID)).makeCopy('[LH] Budget Engine - Source Hub');
  hub.getSheetByName('Control').getRange('B3').setValue(copy.getId()); return copy.getUrl();
}
LH.withLock = function (fn) {
  var lock = LockService.getScriptLock(); LH.Core.assert(lock.tryLock(1000), 'Another refresh is running');
  try { return fn(); } finally { lock.releaseLock(); }
};
function discoverImports() {
  LH.withLock(function () {
    var c = LH.context(), old = LH.state();
    LH.Core.assert(!old || ['FAILED','PUBLISHED','DISCOVERED'].indexOf(old.phase) >= 0, 'A job is already running');
    LH.Core.assert(!LH.IO.readTable(c.hub,'__Formulas').length, 'Templates already captured. Use startRefresh.');
    LH.IO.table(c.hub,'__Formulas',[]);
    var s = { id: Utilities.getUuid(), phase:'DISCOVER', sheet:0, row:1, engineId:c.engine.getId() }; LH.saveState(s); LH.step(c,s);
  });
}
function startRefresh() {
  LH.withLock(function () {
    var c = LH.context(), old = LH.state();
    LH.Core.assert(!old || ['FAILED','PUBLISHED','DISCOVERED'].indexOf(old.phase) >= 0, 'Finish the active job first');
    LH.Core.assert(LH.IO.readTable(c.hub,'__Formulas').length, 'Run discoverImports first');
    var s = { id:Utilities.getUuid(), phase:'PLAN', dataset:0, nextRow:1, manual:0, engineId:c.engine.getId(), runDate:c.settings.SOURCE_DATE || Utilities.formatDate(new Date(),LH.Config.timezone,'yyyy-MM-dd') };
    LH.saveState(s); LH.log(c.hub,s,'Refresh requested'); LH.step(c,s);
  });
}
function continueRefresh() {
  LH.withLock(function () {
    var c = LH.context(), s = LH.state(); LH.Core.assert(s && s.engineId === c.engine.getId(), 'No job for this engine');
    LH.Core.assert(['DISCOVER','PLAN','PULL','MANUAL','PUBLISH','COMMITTING'].indexOf(s.phase) >= 0,'No resumable job: ' + s.phase); LH.step(c,s);
  });
}
LH.step = function (c,s) {
  var deadline = Date.now() + LH.Config.executionMs;
  try {
    if (s.phase === 'COMMITTING') {
      var status=c.engine.getSheetByName('__Source_Status');
      if (status && status.getRange('B1').getValue()===s.id) { LH.finishPublication(c,s); return; }
      // A batchUpdate may have succeeded even if its HTTP response was lost. Check the run ID first.
      s.phase='PUBLISH'; LH.saveState(s);
    }
    if (s.phase === 'DISCOVER') {
      LH.Planner.discoverStep(c.hub,c.engine,s,deadline);
      if (s.phase === 'PLAN') { s.phase = 'DISCOVERED'; LH.saveState(s); LH.log(c.hub,s,'Discovery complete. Run startRefresh.'); } return;
    }
    if (s.phase === 'PLAN') {
      c.settings.RUN_DATE = s.runDate;
      var plan = LH.Planner.build(c.hub,c.engine,c.settings), additions = [];
      plan.datasets.forEach(function (d) { [d.cache,d.stage].forEach(function (name) { additions.push({name:name,rows:d.rows,cols:d.columns.length}); }); });
      LH.Config.manualInputs.forEach(function (m) {
        var r = LH.Core.range(LH.Core.quote(m.sheet)+'!'+m.a1);
        additions.push({name:'__stage_manual_'+m.key,rows:r.r2-r.r1+1,cols:r.c2-r.c1+1});
      });
      var engineAdditions=additions.slice();
      plan.datasets.forEach(function (d) { d.feeds.forEach(function (f) { engineAdditions.push({name:f.dest,rows:f.row+f.rows-1,cols:f.col+f.sourceColumns.length-1}); }); });
      LH.IO.capacity(c.engine,engineAdditions); LH.IO.capacity(c.hub,additions);
      plan.datasets.forEach(function (d) { [c.hub,c.engine].forEach(function (b) { [d.cache,d.stage].forEach(function (name) { LH.IO.ensure(b,name,d.rows,d.columns.length); }); }); });
      LH.IO.table(c.hub,'__Plan',plan.datasets); LH.IO.table(c.hub,'__PendingFormulas',plan.formulas);
      s.phase='PULL'; s.dataset=0; s.nextRow=1; LH.saveState(s); LH.log(c.hub,s,'Planned '+plan.datasets.length+' source tables');
    }
    var datasets = LH.IO.readTable(c.hub,'__Plan');
    while (s.phase === 'PULL' && s.dataset < datasets.length && Date.now() < deadline) {
      var d=datasets[s.dataset], width=d.columns.length;
      var height=Math.min(d.rows-s.nextRow+1,Math.max(1,Math.floor(LH.Config.chunkCells/width)));
      var ranges=d.columns.map(function (col) { return LH.Core.quote(d.tab)+'!'+LH.Core.column(col)+s.nextRow+':'+LH.Core.column(col)+(s.nextRow+height-1); });
      var res=Sheets.Spreadsheets.Values.batchGet(d.sourceId,{ranges:ranges,valueRenderOption:'UNFORMATTED_VALUE',dateTimeRenderOption:'SERIAL_NUMBER'});
      var rows=Array.from({length:height},function (_,i) { return d.columns.map(function (_,j) { return (((res.valueRanges[j]||{}).values||[])[i]||[''])[0]; }); });
      LH.Core.checkErrors(rows);
      LH.IO.writeVerified(c.hub,d.stage,s.nextRow,rows,width);
      var fromHub=LH.IO.values(c.hub.getId(),LH.Core.quote(d.stage)+'!A'+s.nextRow+':'+LH.Core.column(width)+(s.nextRow+height-1),width,height);
      LH.IO.writeVerified(c.engine,d.stage,s.nextRow,fromHub,width);
      s.nextRow+=height;
      if (s.nextRow>d.rows) { s.dataset++; s.nextRow=1; LH.log(c.hub,s,'Staged '+d.tab); } LH.saveState(s);
    }
    if (s.phase==='PULL' && s.dataset===datasets.length) { s.phase='MANUAL'; s.manual=0; s.nextRow=1; LH.saveState(s); }
    while (s.phase==='MANUAL' && s.manual<LH.Config.manualInputs.length && Date.now()<deadline) {
      var m=LH.Config.manualInputs[s.manual], r=LH.Core.range(LH.Core.quote(m.sheet)+'!'+m.a1);
      var w=r.c2-r.c1+1, h=r.r2-r.r1+1, name='manual_'+m.key;
      var seed=PropertiesService.getScriptProperties().getProperty('MANUAL_'+m.key)!=='READY';
      if (seed) {
        LH.Core.assert(!c.engine.getSheetByName(m.sheet).getRange(m.a1).getFormulas().some(function (row) { return row.some(Boolean); }),'Manual input contains formulas: '+m.key);
        LH.IO.ensure(c.hub,name,h,w);
      }
      var n=Math.min(h-s.nextRow+1,Math.max(1,Math.floor(LH.Config.chunkCells/w)));
      var data=seed ? LH.IO.values(c.engine.getId(),LH.Core.quote(m.sheet)+'!'+LH.Core.column(r.c1)+(r.r1+s.nextRow-1)+':'+LH.Core.column(r.c2)+(r.r1+s.nextRow+n-2),w,n)
        : LH.IO.values(c.hub.getId(),LH.Core.quote(name)+'!A'+s.nextRow+':'+LH.Core.column(w)+(s.nextRow+n-1),w,n);
      LH.Core.checkErrors(data);
      if (seed) LH.IO.writeVerified(c.hub,name,s.nextRow,data,w);
      LH.IO.writeVerified(c.engine,'__stage_manual_'+m.key,s.nextRow,data,w);
      s.nextRow+=n;
      if (s.nextRow>h) { PropertiesService.getScriptProperties().setProperty('MANUAL_'+m.key,'READY'); s.manual++; s.nextRow=1; } LH.saveState(s);
    }
    if (s.phase==='MANUAL' && s.manual===LH.Config.manualInputs.length) { s.phase='PUBLISH'; LH.saveState(s); }
    if (s.phase==='PUBLISH' && Date.now()<deadline-15000) LH.publish(c,s,datasets);
    else LH.log(c.hub,s,'Checkpoint saved. Run continueRefresh.');
  } catch (e) {
    if (s.phase!=='COMMITTING' && s.phase!=='PUBLISHED') { s.phase='FAILED'; s.error=e.message; LH.saveState(s); }
    try { LH.log(c.hub,s,'ERROR: '+e.message); } catch (ignored) { console.error(ignored.message); }
    throw e;
  }
};

