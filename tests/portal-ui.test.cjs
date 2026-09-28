const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const light = fs.readFileSync('source/light/index.html', 'utf8');
const dark = fs.readFileSync('source/dark/index.html', 'utf8');

test('both standalone themes carry identical styles, behavior, and controls', () => {
  assert.equal(light.replace('data-theme="light"','data-theme="dark"'),dark);
  for (const source of [light,dark]) {
    assert.match(source, /id="themeToggleBtn"/);
    assert.match(source, /id="outlookSyncBtn"/);
    assert.match(source, /html\[data-theme="dark"\]/);
    assert.match(source, /html\[data-theme="light"\]/);
  }
});

test('theme switch persists and Outlook button hands off only an ID', () => {
  const begin = light.indexOf('function applyTheme(theme)');
  const end = light.indexOf('function refreshCurrentUserControls(){',begin);
  const button = { textContent:'', setAttribute(name,value) {this[name]=value;} };
  const store = new Map();
  const context = {
    document:{documentElement:{dataset:{theme:'light'}}},
    q: () => button,
    localStorage:{setItem:(key,value)=>store.set(key,value)},
    THEME_KEY:'personal_work_portal_theme',
    currentUserId:'u-1', sharedFileHandle:{}, sharedConflict:false,
    sharedSnapshot:()=> 'same', lastSharedSnapshot:'same',
    window:{location:{href:''}}, alert:()=>{throw new Error('unexpected alert')}
  };
  vm.createContext(context);
  vm.runInContext(light.slice(begin,end),context);
  context.switchTheme();
  assert.equal(context.document.documentElement.dataset.theme,'dark');
  assert.equal(store.get('personal_work_portal_theme'),'dark');
  context.syncOutlook();
  assert.equal(context.window.location.href,'workportal-outlook://sync/u-1');
});
