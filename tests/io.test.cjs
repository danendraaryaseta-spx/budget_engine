const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
function load(mocks) {
  const ctx = vm.createContext({ console, ...mocks });
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
  const book={getSheetByName:n=>({getSheetId:()=>n==='stage'?1:2})};
  const requests=ctx.LH.IO.publishRequests(book,'stage','live',8,3);
  assert.equal(requests[0].updateCells.range.sheetId,2);
  assert.equal(requests[0].updateCells.fields,'userEnteredValue');
  assert.equal(requests[1].copyPaste.pasteType,'PASTE_VALUES');
  assert.equal(requests[1].copyPaste.destination.endRowIndex,8);
});
test('capacity check blocks expansion before touching the engine', () => {
  const ctx=load({});
  const sheet={getMaxRows:()=>999000,getMaxColumns:()=>10};
  const book={getSheets:()=>[sheet],getSheetByName:()=>null,getName:()=> 'engine'};
  assert.throws(()=>ctx.LH.IO.capacity(book,[{name:'new',rows:2000,cols:10}]),/would need/);
});
test('lost successful commit response repairs metadata instead of publishing again', () => {
  const ctx=load({}); let repaired=false;
  ctx.LH.finishPublication=()=>{repaired=true;};
  ctx.LH.publish=()=>{throw new Error('must not republish');};
  const c={engine:{getSheetByName:()=>({getRange:()=>({getValue:()=> 'run-123'})})}};
  ctx.LH.step(c,{phase:'COMMITTING',id:'run-123'});
  assert.equal(repaired,true);
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
