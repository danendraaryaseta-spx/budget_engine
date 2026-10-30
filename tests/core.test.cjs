const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const ctx = vm.createContext({ console });
vm.runInContext(fs.readFileSync(path.join(__dirname, '../src/Core.js'), 'utf8'), ctx);
const C = ctx.LH.Core;
const plain = x => JSON.parse(JSON.stringify(x));

test('nested import calls are replaced while budget arithmetic and quoted text survive', () => {
  const f = '=ARRAYFORMULA(IFERROR(XLOOKUP($F4:$F&$G4:$G,IMPORTRANGE($B$44,"Rates!B:B")&IMPORTRANGE($B$44,"Rates!C:C"),IMPORTRANGE($B$44,"Rates!E:E"))*JR4:JR*DAY(EOMONTH(KN$3,0))))';
  const calls = C.imports(f);
  assert.equal(calls.length, 3);
  const map = { '"Rates!B:B"': "'__src_test'!A1:A700", '"Rates!C:C"': "'__src_test'!B1:B700", '"Rates!E:E"': "'__src_test'!C1:C700" };
  const compiled = C.compile(f, calls, call => map[call.range]);
  assert.ok(compiled.includes('*JR4:JR*DAY(EOMONTH(KN$3,0))'));
  assert.equal(C.imports(compiled).length, 0);
  assert.equal(C.imports('="IMPORTRANGE(A1,""x"")"').length, 0);
});
test('dynamic month concatenation resolves literal strings and quoted sheet references only', () => {
  assert.equal(C.expression('"\'Air-"&B1&"\'!A5:AR"', r => { assert.equal(r, 'B1'); return 'Oct'; }), "'Air-Oct'!A5:AR");
  assert.equal(C.expression("'Base Assumptions'!$B$44", () => 'source_id'), 'source_id');
  assert.throws(() => C.expression('INDIRECT(A1)', () => ''), /Unsupported/);
  assert.throws(() => C.expression('fetch("https://example.com")', () => ''), /Unsupported/);
});
test('A1 parsing preserves row offsets, casing, apostrophes and finite bounds', () => {
  assert.deepEqual(plain(C.range("'O''Brien Data'!ae1:bf4800")), { tab:"O'Brien Data", c1:31,c2:58,r1:1,r2:4800 });
  assert.deepEqual(plain(C.range('Budget STD!B2:J')), { tab:'Budget STD',c1:2,c2:10,r1:2,r2:null });
  assert.throws(() => C.range('x!D10:A2'), /Reversed/);
  assert.equal(C.column(320), 'LH'); assert.equal(C.columnNumber('LH'), 320);
});
test('compact projections retain contiguous remote blocks and source row numbers', () => {
  const d = { cache:'__src_x', columns:[2,3,4,5,10,11,12],rows:400 };
  assert.equal(C.localRange(d,C.range('x!B3:E')), "'__src_x'!A3:D400");
  assert.equal(C.localRange(d,C.range('x!J:J')), "'__src_x'!E1:E400");
  assert.equal(C.localRange(d,C.range('x!K20:L30')), "'__src_x'!F20:G30");
  assert.throws(() => C.localRange(d,C.range('x!C:I')), /projection/);
});
test('latest schedule ignores future dates, invalid dates, and unrelated tabs', () => {
  assert.equal(C.latestTab(['31 Feb 26','30 Oct 26','31 Oct 26','27 Oct 26','Summary'],'2026-10-30'),'30 Oct 26');
  assert.throws(() => C.latestTab(['31 Oct 26'],'2026-10-30'), /No dated/);
});
test('source errors fail rather than silently becoming a zero budget', () => {
  assert.throws(() => C.checkErrors([[123,'#REF!']]), /Source error/);
  assert.throws(() => C.checkErrors([['#N/A']]), /Source error/);
  C.checkErrors([[0,false,'=1+1','# normal label']]);
  assert.deepEqual(plain(C.rectangular([[1],[2,0]],2,3)),[[1,''],[2,0],['','']]);
});
test('URLs and cell references are parsed without executable argument evaluation', () => {
  assert.equal(C.spreadsheetId('https://docs.google.com/spreadsheets/d/abcdefghijklmnopqrstuvwxyz/edit'), 'abcdefghijklmnopqrstuvwxyz');
  assert.throws(() => C.spreadsheetId('L1M Wallet'), /Expected/);
  assert.deepEqual(plain(C.reference("'O''Brien'!$G$2",'other')), { sheet:"O'Brien",cell:'G2' });
});
