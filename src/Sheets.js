var LH = LH || {};
LH.IO = (function () {
  'use strict';
  var C = function () { return LH.Core; };
  // SpreadsheetApp can retain a stale tab list after an advanced-service write.
  // Keep owned-tab metadata in the same API as tab creation, never native handles.
  var metadata = {};
  function bookMetadata(book, force) {
    var id=book.getId();
    if (force || !metadata[id]) metadata[id]=Sheets.Spreadsheets.get(id,{fields:'properties.title,sheets.properties(sheetId,title,gridProperties)'});
    return metadata[id];
  }
  function sheet(book, name, force) {
    var found=(bookMetadata(book,force).sheets||[]).filter(function (s) { return s.properties.title===name; });
    return found.length ? found[0].properties : null;
  }
  function rows(book, name, width) {
    var p=sheet(book,name); if (!p) return [];
    var a1=C().quote(name)+'!A1:'+C().column(width||p.gridProperties.columnCount)+p.gridProperties.rowCount;
    return Sheets.Spreadsheets.Values.get(book.getId(),a1,{valueRenderOption:'UNFORMATTED_VALUE',dateTimeRenderOption:'SERIAL_NUMBER'}).values||[];
  }
  function raw(id, a1, values) {
    if (!values.length) return;
    var r=C().range(a1), width=values.reduce(function (n,row) { return Math.max(n,row.length); },0);
    ensure(SpreadsheetApp.openById(id),r.tab,r.r1+values.length-1,r.c1+width-1);
    Sheets.Spreadsheets.Values.update({ values: values, majorDimension: 'ROWS' }, id, a1, { valueInputOption: 'RAW' });
  }
  function values(id, a1, width, height) {
    var r = Sheets.Spreadsheets.Values.get(id, a1, { valueRenderOption: 'UNFORMATTED_VALUE', dateTimeRenderOption: 'SERIAL_NUMBER' });
    return C().rectangular(r.values || [], width, height);
  }
  function ensure(book, name, rows, cols) {
    rows=Math.max(1,rows); cols=Math.max(1,cols);
    var p=sheet(book,name);
    if (!p) p=sheet(book,name,true); // Another invocation may have created it.
    if (!p) {
      try {
        var result=Sheets.Spreadsheets.batchUpdate({requests:[{addSheet:{properties:{title:name,gridProperties:{rowCount:rows,columnCount:cols}}}}]},book.getId());
        p=result.replies[0].addSheet.properties;
        bookMetadata(book).sheets.push({properties:p});
      } catch (e) {
        // A concurrent creator can win between the read and add. Only recover if
        // the server actually contains the requested sheet; propagate other errors.
        p=sheet(book,name,true); if (!p) throw e;
      }
    }
    C().assert(p && p.gridProperties,'Sheets API did not return metadata for '+name);
    var nr=Math.max(rows,p.gridProperties.rowCount),nc=Math.max(cols,p.gridProperties.columnCount);
    if (nr!==p.gridProperties.rowCount || nc!==p.gridProperties.columnCount) {
      Sheets.Spreadsheets.batchUpdate({requests:[{updateSheetProperties:{properties:{sheetId:p.sheetId,gridProperties:{rowCount:nr,columnCount:nc}},fields:'gridProperties.rowCount,gridProperties.columnCount'}}]},book.getId());
      p.gridProperties.rowCount=nr; p.gridProperties.columnCount=nc;
    }
    return p;
  }
  function table(book, name, objects) {
    var p = ensure(book, name, Math.max(2, objects.length + 1), 2);
    var data = [['Key', 'JSON']].concat(objects.map(function (x, i) { return [x.key || String(i + 1), JSON.stringify(x)]; }));
    raw(book.getId(), C().quote(name) + '!A1', data);
    if (p.gridProperties.rowCount>data.length) Sheets.Spreadsheets.Values.clear({},book.getId(),C().quote(name)+'!A'+(data.length+1)+':B'+p.gridProperties.rowCount);
    Sheets.Spreadsheets.batchUpdate({requests:[{updateSheetProperties:{properties:{sheetId:p.sheetId,gridProperties:{frozenRowCount:1}},fields:'gridProperties.frozenRowCount'}}]},book.getId());
    return p;
  }
  function readTable(book, name) {
    return rows(book,name,2).slice(1).filter(function (r) { return r[1]; }).map(function (r) { return JSON.parse(r[1]); });
  }
  function hash(x) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(x)).map(function (v) { return ('0' + ((v + 256) % 256).toString(16)).slice(-2); }).join(''); }
  function writeVerified(book, name, firstRow, rows, width) {
    if (!rows.length) return;
    ensure(book, name, firstRow + rows.length - 1, width);
    var a1 = C().quote(name) + '!A' + firstRow + ':' + C().column(width) + (firstRow + rows.length - 1);
    raw(book.getId(), a1, rows);
    C().assert(hash(rows) === hash(values(book.getId(), a1, width, rows.length)), 'Readback mismatch: ' + name + ' at row ' + firstRow);
  }
  function capacity(book, additions) {
    var meta=bookMetadata(book,true),grids={};
    (meta.sheets||[]).forEach(function (s) { grids[s.properties.title]={rows:s.properties.gridProperties.rowCount,cols:s.properties.gridProperties.columnCount}; });
    additions.forEach(function (a) { var g=grids[a.name]||{rows:0,cols:0}; grids[a.name]={rows:Math.max(a.rows,g.rows),cols:Math.max(a.cols,g.cols)}; });
    var total=Object.keys(grids).reduce(function (n,k) { return n+grids[k].rows*grids[k].cols; },0);
    C().assert(total <= LH.Config.maxWorkbookCells, meta.properties.title + ' would need ' + total + ' cells. Trim unused grids or reduce source scope before retrying.');
    return total;
  }
  function publishRequests(book, stageName, liveName, rows, cols) {
    var live = sheet(book,liveName), stage = sheet(book,stageName);
    C().assert(live && stage, 'Missing staging or live table: ' + liveName);
    return [
      { updateCells: { range: { sheetId: live.sheetId }, fields: 'userEnteredValue' } },
      { copyPaste: { source: { sheetId: stage.sheetId, startRowIndex: 0, endRowIndex: rows, startColumnIndex: 0, endColumnIndex: cols },
        destination: { sheetId: live.sheetId, startRowIndex: 0, endRowIndex: rows, startColumnIndex: 0, endColumnIndex: cols }, pasteType: 'PASTE_VALUES' } }
    ];
  }
  return { raw: raw, values: values, ensure: ensure, sheet:sheet, rows:rows, table: table, readTable: readTable, hash: hash, writeVerified: writeVerified, capacity: capacity, publishRequests: publishRequests };
}());

