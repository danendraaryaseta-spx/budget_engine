var LH = LH || {};
LH.publish = function (c,s,datasets) {
  var C=LH.Core,I=LH.IO, hr=[], er=[], formulas=I.readTable(c.hub,'__PendingFormulas');
  formulas.forEach(function (f) { C.assert(c.engine.getSheetByName(f.sheet).getRange(f.cell).getFormula()===f.expected,'Formula changed during refresh: '+f.key); });
  datasets.forEach(function (d) {
    hr=hr.concat(I.publishRequests(c.hub,d.stage,d.cache,d.rows,d.columns.length));
    er=er.concat(I.publishRequests(c.engine,d.stage,d.cache,d.rows,d.columns.length));
    d.feeds.forEach(function (f) {
      var dest=I.ensure(c.engine,f.dest,f.row+f.rows-1,f.col+f.sourceColumns.length-1), source=c.engine.getSheetByName(d.stage);
      er.push({updateCells:{range:{sheetId:dest.getSheetId(),startRowIndex:f.row-1,startColumnIndex:f.col-1,endColumnIndex:f.col-1+f.sourceColumns.length},fields:'userEnteredValue'}});
      f.sourceColumns.forEach(function (col,j) {
        var sc=d.columns.indexOf(col);
        er.push({copyPaste:{source:{sheetId:source.getSheetId(),startRowIndex:f.firstRow-1,endRowIndex:f.firstRow-1+f.rows,startColumnIndex:sc,endColumnIndex:sc+1},destination:{sheetId:dest.getSheetId(),startRowIndex:f.row-1,endRowIndex:f.row-1+f.rows,startColumnIndex:f.col-1+j,endColumnIndex:f.col+j},pasteType:'PASTE_VALUES'}});
      });
    });
  });
  LH.Config.manualInputs.forEach(function (m) {
    var r=C.range(C.quote(m.sheet)+'!'+m.a1), dest=c.engine.getSheetByName(m.sheet), stage=c.engine.getSheetByName('__stage_manual_'+m.key);
    C.assert(!dest.getRange(m.a1).getFormulas().some(function (row) { return row.some(Boolean); }),'Manual destination now contains formulas: '+m.key);
    er.push({copyPaste:{source:{sheetId:stage.getSheetId(),startRowIndex:0,endRowIndex:r.r2-r.r1+1,startColumnIndex:0,endColumnIndex:r.c2-r.c1+1},destination:{sheetId:dest.getSheetId(),startRowIndex:r.r1-1,endRowIndex:r.r2,startColumnIndex:r.c1-1,endColumnIndex:r.c2},pasteType:'PASTE_VALUES'}});
  });
  formulas.forEach(function (f) {
    var cell=c.engine.getSheetByName(f.sheet).getRange(f.cell);
    er.push({updateCells:{start:{sheetId:cell.getSheet().getSheetId(),rowIndex:cell.getRow()-1,columnIndex:cell.getColumn()-1},rows:[{values:[{userEnteredValue:{formulaValue:f.compiled}}]}],fields:'userEnteredValue'}});
  });
  var status=I.ensure(c.engine,'__Source_Status',4,2), stamp=Utilities.formatDate(new Date(),LH.Config.timezone,'yyyy-MM-dd HH:mm:ss');
  er.push({updateCells:{start:{sheetId:status.getSheetId(),rowIndex:0,columnIndex:0},rows:[['Published run',s.id],['Published at (WIB)',stamp],['Source hub ID',c.hub.getId()]].map(function (r) { return {values:r.map(function (v) { return {userEnteredValue:{stringValue:v}}; })}; }),fields:'userEnteredValue'}});
  // Two workbooks cannot share a transaction. Engine publication is the final atomic commit.
  if (hr.length) Sheets.Spreadsheets.batchUpdate({requests:hr},c.hub.getId());
  s.phase='COMMITTING'; s.publishedAt=stamp; LH.saveState(s);
  Sheets.Spreadsheets.batchUpdate({requests:er},c.engine.getId());
  LH.finishPublication(c,s);
};
LH.finishPublication = function (c,s) {
  var I=LH.IO, formulas=I.readTable(c.hub,'__PendingFormulas');
  // Keep COMMITTING until the metadata is repaired, so a lost response is recoverable.
  formulas.forEach(function (f) { f.expected=f.compiled; delete f.calls; }); I.table(c.hub,'__Formulas',formulas);
  c.hub.getSheetByName('Control').getRange('B6:B8').setValues([['PUBLISHED'],[s.id],[s.publishedAt]]);
  s.phase='PUBLISHED'; LH.saveState(s);
  LH.log(c.hub,s,'Published. Budget formulas calculate from the local values snapshot.');
};
