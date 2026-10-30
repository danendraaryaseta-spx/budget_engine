const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
function load(mocks={}) {
  const defaultGet=()=>({properties:{title:'test book'},sheets:['source','Refresh Log','stage','live'].map((title,i)=>({properties:{sheetId:title==='stage'?1:title==='live'?2:10+i,title,gridProperties:{rowCount:100,columnCount:100}}}))});
  const book={getId:()=> 'hub', getSheetByName:()=>{throw new Error('stale native tab list must not be used');}};
  const defaults={get:defaultGet,batchUpdate:()=>({replies:[]}),Values:{update:()=>{},get:()=>({values:[]})}};
  const supplied=mocks.Sheets?.Spreadsheets||{};
  const api={...defaults,...supplied,Values:{...defaults.Values,...supplied.Values}};
  const ctx = vm.createContext({ console, SpreadsheetApp:{openById:()=>book}, ...mocks, Sheets:{Spreadsheets:api} });
  for (const f of ['Core.js','Config.js','Sheets.js','Runner.js','Publish.js']) vm.runInContext(fs.readFileSync(path.join(__dirname,'../src',f),'utf8'),ctx);
  return ctx;
}
test('all source writes use RAW; formula-looking text is kept as data', () => {
  let called;
  const ctx=load({ Sheets:{ Spreadsheets:{ Values:{update:(body,id,range,options)=>{called={body,id,range,options};}}}} });
  ctx.LH.IO.raw('hub',"'source'!A1",[['=HYPERLINK("url")',0,false,46296]]);
  assert.equal(called.options.valueInputOption,'RAW');
  assert.equal(called.body.values[0][0],'=HYPERLINK("url")');
  assert.equal(called.body.values[0][1],0);
});
test('publication clears old table tails and pastes values only', () => {
  const ctx=load({});
  const book={getId:()=> 'hub'};
  const requests=ctx.LH.IO.publishRequests(book,'stage','live',8,3);
  assert.equal(requests[0].updateCells.range.sheetId,2);
  assert.equal(requests[0].updateCells.fields,'userEnteredValue');
  assert.equal(requests[1].copyPaste.pasteType,'PASTE_VALUES');
  assert.equal(requests[1].copyPaste.destination.endRowIndex,8);
});
test('append writes expand table grids instead of failing at the third log row', () => {
  let newRows=0;
  const ctx=load({Sheets:{Spreadsheets:{get:()=>({sheets:[{properties:{sheetId:9,title:'Refresh Log',gridProperties:{rowCount:2,columnCount:2}}}]}),batchUpdate:({requests})=>{newRows=requests[0].updateSheetProperties.properties.gridProperties.rowCount;}}}});
  ctx.LH.IO.raw('hub',"'Refresh Log'!A3",[['event']]);
  assert.equal(newRows,3);
});
test('capacity check blocks expansion before touching the engine', () => {
  const ctx=load({Sheets:{Spreadsheets:{get:()=>({properties:{title:'engine'},sheets:[{properties:{sheetId:1,title:'existing',gridProperties:{rowCount:999000,columnCount:10}}}]})}}});
  const book={getId:()=> 'engine'};
  assert.throws(()=>ctx.LH.IO.capacity(book,[{name:'new',rows:2000,cols:10}]),/would need/);
});
test('lost successful commit response repairs metadata instead of publishing again', () => {
  const ctx=load({}); let repaired=false;
  ctx.LH.finishPublication=()=>{repaired=true;};
  ctx.LH.publish=()=>{throw new Error('must not republish');};
  ctx.LH.IO.sheet=()=>({sheetId:7}); ctx.LH.IO.values=()=>[['run-123']];
  const c={engine:{getId:()=> 'engine'}};
  ctx.LH.step(c,{phase:'COMMITTING',id:'run-123'});
  assert.equal(repaired,true);
});

test('existing server tab is reused even when every native lookup returns null', () => {
  let adds=0;
  const ctx=load({Sheets:{Spreadsheets:{batchUpdate:()=>{adds++;throw new Error('duplicate tab');}}}});
  const book={getId:()=> 'hub',getSheetByName:()=>null};
  const p=ctx.LH.IO.ensure(book,'Refresh Log',2,6);
  assert.equal(p.title,'Refresh Log'); assert.equal(adds,0);
});
test('new tabs use API reply metadata immediately, without a native read-after-write', () => {
  let adds=0;
  const ctx=load({Sheets:{Spreadsheets:{get:()=>({sheets:[]}),batchUpdate:({requests})=>{
    adds++; const p=requests[0].addSheet.properties; return {replies:[{addSheet:{properties:{...p,sheetId:47}}}]};
  }}}});
  const book={getId:()=> 'hub',getSheetByName:()=>null};
  assert.equal(ctx.LH.IO.ensure(book,'new',3,4).sheetId,47);
  assert.equal(ctx.LH.IO.ensure(book,'new',3,4).sheetId,47);
  assert.equal(adds,1);
});
test('concurrent duplicate creation recovers only after server confirms the tab exists', () => {
  let exists=false;
  const p={sheetId:47,title:'new',gridProperties:{rowCount:3,columnCount:4}};
  const ctx=load({Sheets:{Spreadsheets:{get:()=>({sheets:exists?[{properties:p}]:[]}),batchUpdate:()=>{exists=true;throw new Error('name already exists');}}}});
  assert.equal(ctx.LH.IO.ensure({getId:()=> 'hub'},'new',3,4).sheetId,47);
});
test('unrelated creation errors propagate instead of becoming null-sheet errors', () => {
  const ctx=load({Sheets:{Spreadsheets:{get:()=>({sheets:[]}),batchUpdate:()=>{throw new Error('permission denied');}}}});
  assert.throws(()=>ctx.LH.IO.ensure({getId:()=> 'hub'},'new',3,4),/permission denied/);
});
test('refresh logging writes its header and entries without native sheet handles', () => {
  const stored=[]; const updates=[];
  const ctx=load({Utilities:{formatDate:()=> '2026-10-30 17:40:00'},Sheets:{Spreadsheets:{Values:{
    get:()=>({values:stored}),
    update:(body,id,a1)=>{updates.push(a1);stored.push(...body.values);}
  }}}});
  const book={getId:()=> 'hub',getSheetByName:()=>null};
  ctx.LH.log(book,{id:'run',phase:'DISCOVERED'},'complete');
  ctx.LH.log(book,{id:'run',phase:'PLAN'},'next');
  assert.deepEqual(updates,["'Refresh Log'!A1","'Refresh Log'!A2","'Refresh Log'!A3"]);
});
test('recovery from a post-discovery logging failure starts refresh without rescanning', () => {
  const ctx=load({}); let started=0;
  let state={id:'run',engineId:'engine',phase:'FAILED',failedPhase:'DISCOVERED'};
  ctx.LH.state=()=>state; ctx.LH.saveState=s=>{state=s;};
  ctx.LH.context=()=>({engine:{getId:()=> 'engine'}}); ctx.LH.withLock=f=>f();
  ctx.LH.step=()=>{throw new Error('must not rescan');}; ctx.startRefresh=()=>{started++;};
  ctx.resumeSources(); assert.equal(started,1); assert.equal(state.phase,'DISCOVERED');
});
test('a failed source read does not call publish or change published inputs', () => {
  const ctx=load({Sheets:{Spreadsheets:{Values:{batchGet:()=>{throw new Error('source unavailable');}}}}});
  let published=false, logged=false;
  ctx.LH.IO.readTable=()=>[{sourceId:'upstream',tab:'data',columns:[1],rows:3}];
  ctx.LH.publish=()=>{published=true;}; ctx.LH.log=()=>{logged=true;}; ctx.LH.saveState=()=>{};
  const state={phase:'PULL',dataset:0,nextRow:1};
  assert.throws(()=>ctx.LH.step({},state),/source unavailable/);
  assert.equal(published,false); assert.equal(logged,true); assert.equal(state.phase,'FAILED');
});
