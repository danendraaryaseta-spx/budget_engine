var LH = LH || {};
LH.IO = (function () {
  'use strict';
  var C = function () { return LH.Core; };
  function raw(id, a1, values) {
    if (!values.length) return;
    Sheets.Spreadsheets.Values.update({ values: values, majorDimension: 'ROWS' }, id, a1, { valueInputOption: 'RAW' });
  }
  function values(id, a1, width, height) {
    var r = Sheets.Spreadsheets.Values.get(id, a1, { valueRenderOption: 'UNFORMATTED_VALUE', dateTimeRenderOption: 'SERIAL_NUMBER' });
    return C().rectangular(r.values || [], width, height);
  }
  function ensure(book, name, rows, cols) {
    var sh = book.getSheetByName(name);
    if (!sh) {
      Sheets.Spreadsheets.batchUpdate({ requests: [{ addSheet: { properties: { title: name, gridProperties: { rowCount: Math.max(1, rows), columnCount: Math.max(1, cols) } } } }] }, book.getId());
      sh = SpreadsheetApp.openById(book.getId()).getSheetByName(name);
    }
    if (rows > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
    if (cols > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
    return sh;
  }
  function table(book, name, objects) {
    var sh = ensure(book, name, Math.max(2, objects.length + 1), 2);
    var rows = [['Key', 'JSON']].concat(objects.map(function (x, i) { return [x.key || String(i + 1), JSON.stringify(x)]; }));
    raw(book.getId(), C().quote(name) + '!A1', rows);
    if (sh.getLastRow() > rows.length) sh.getRange(rows.length + 1, 1, sh.getLastRow() - rows.length, 2).clearContent();
    sh.setFrozenRows(1); return sh;
  }
  function readTable(book, name) {
    var sh = book.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().filter(function (r) { return r[0]; }).map(function (r) { return JSON.parse(r[0]); });
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
    var total = book.getSheets().reduce(function (n, s) { return n + s.getMaxRows() * s.getMaxColumns(); }, 0);
    additions.forEach(function (a) {
      var s = book.getSheetByName(a.name);
      total += Math.max(a.rows, s ? s.getMaxRows() : 0) * Math.max(a.cols, s ? s.getMaxColumns() : 0) - (s ? s.getMaxRows() * s.getMaxColumns() : 0);
    });
    C().assert(total <= LH.Config.maxWorkbookCells, book.getName() + ' would need ' + total + ' cells. Trim unused grids or reduce source scope before retrying.');
    return total;
  }
  function publishRequests(book, stageName, liveName, rows, cols) {
    var live = book.getSheetByName(liveName), stage = book.getSheetByName(stageName);
    C().assert(live && stage, 'Missing staging or live table: ' + liveName);
    return [
      { updateCells: { range: { sheetId: live.getSheetId() }, fields: 'userEnteredValue' } },
      { copyPaste: { source: { sheetId: stage.getSheetId(), startRowIndex: 0, endRowIndex: rows, startColumnIndex: 0, endColumnIndex: cols },
        destination: { sheetId: live.getSheetId(), startRowIndex: 0, endRowIndex: rows, startColumnIndex: 0, endColumnIndex: cols }, pasteType: 'PASTE_VALUES' } }
    ];
  }
  return { raw: raw, values: values, ensure: ensure, table: table, readTable: readTable, hash: hash, writeVerified: writeVerified, capacity: capacity, publishRequests: publishRequests };
}());

