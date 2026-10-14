// doPost 路由測試：node build/test_dopost.js
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const ctx = {
  ContentService: { MimeType: { JSON: 'json' }, createTextOutput: t => ({ t, setMimeType() { return this; } }) },
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
  CacheService: { getScriptCache: () => ({ get: () => null, put() {}, remove() {} }) }
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(__dirname + '/../apps-script/Code.gs', 'utf8'), ctx);
const call = body => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(body) } }).t);
let r = call({ fn: 'apiBootstrap', args: ['bad-token'] });
assert.strictEqual(r.ok, false); assert.match(r.error, /初始化|重新登入/);
r = call({ fn: 'setup', args: [] });
assert.strictEqual(r.ok, false); assert.match(r.error, /不支援/);
r = call({ fn: 'apiLogout', args: [''] });
assert.deepStrictEqual(r, { ok: true, data: true });
r = JSON.parse(ctx.doPost({ postData: { contents: 'not json' } }).t);
assert.strictEqual(r.ok, false);
const names = Object.keys(ctx.apiFunctions_());
const declared = (fs.readFileSync(__dirname + '/../apps-script/Code.gs', 'utf8').match(/^function (api\w+)\(/gm) || []).map(s => s.slice(9, -1)).filter(n => !n.endsWith('_'));
assert.deepStrictEqual(names.sort(), declared.sort(), '白名單需涵蓋所有 api 函式');
console.log('doPost 測試通過：白名單 ' + names.length + ' 個函式、錯誤處理、非白名單拒絕');
