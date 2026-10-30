const fs = require('node:fs');
const path = require('node:path');
const cp = require('node:child_process');
for (const f of fs.readdirSync(path.join(__dirname, '../src')).filter(x => x.endsWith('.js'))) {
  cp.execFileSync(process.execPath, ['--check', path.join(__dirname, '../src', f)], { stdio: 'inherit' });
}
const manifest = JSON.parse(fs.readFileSync(path.join(__dirname, '../src/appsscript.json')));
if (manifest.runtimeVersion !== 'V8') throw new Error('V8 required');
if (!manifest.dependencies.enabledAdvancedServices.some(s => s.userSymbol === 'Sheets')) throw new Error('Enable advanced Sheets service');
console.log('All Apps Script files parse; Sheets service configured.');
