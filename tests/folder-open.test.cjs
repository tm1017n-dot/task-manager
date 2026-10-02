const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
function helper() {
  const ctx = {};
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync('native/Open_Folder.js','ascii'),ctx);
  return ctx;
}
const link = path => 'workportal-folder://open/'+encodeURIComponent(path);
test('UNC, local and file URI folder paths preserve spaces and Japanese names', () => {
  const h = helper();
  for (const path of ['\\\\server\\share\\日本語 & 100% #資料', 'C:\\資料\\見積'])
    assert.equal(h.folderPath(link(path)), path);
  assert.equal(h.folderPath(link('file:///C:/%E8%B3%87%E6%96%99/a%20b')), 'C:\\資料\\a b');
  assert.equal(h.folderPath(link('file://server/share/a%20b')), '\\\\server\\share\\a b');
});
test('only an existing folder is passed to ShellExecute, never an executable or command', () => {
  const h = helper(), calls = [];
  const shell = {ShellExecute:(...args)=>calls.push(args)};
  h.openFolder(link('C:\\資料'),{FolderExists:()=>true},shell);
  assert.deepEqual(calls[0],['C:\\資料','','','open',1]);
  assert.throws(()=>h.openFolder(link('C:\\tool.exe'),{FolderExists:()=>false},shell),/フォルダー/);
  for(const path of ['cmd.exe','https://example.com','C:\\" & calc.exe','\\\\?\\C:\\data','\\\\.\\pipe\\a','C:\\data:stream'])
    assert.throws(()=>h.folderPath(link(path)));
  assert.throws(()=>h.folderPath('workportal-folder://open/%ZZ'));
  assert.equal(calls.length,1);
});
