const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const app = fs.readFileSync('source/app/index.html', 'utf8');

test('one HTML carries both palettes and a visible dark scope selection', () => {
  assert.match(app, /id="themeToggleBtn"/);
  assert.match(app, /id="outlookSyncBtn"/);
  assert.match(app, /id="outlookCalendarBtn"/);
  assert.match(app, /html\[data-theme="dark"\]/);
  assert.match(app, /html\[data-theme="light"\]/);
  assert.match(app, /html\[data-theme="dark"\] #homeScopeSwitch button\.active\{background:#a6b9ff;color:#101a36/);
  assert.match(app, /b\.setAttribute\('aria-pressed',String\(selected\)\)/);
  const luminance = hex => {
    const rgb = hex.match(/../g).map(x => parseInt(x, 16) / 255)
      .map(x => x <= 0.04045 ? x / 12.92 : ((x + 0.055) / 1.055) ** 2.4);
    return rgb[0] * 0.2126 + rgb[1] * 0.7152 + rgb[2] * 0.0722;
  };
  const contrast = (a, b) => {
    const [high, low] = [luminance(a), luminance(b)].sort((x, y) => y - x);
    return (high + 0.05) / (low + 0.05);
  };
  assert.ok(contrast('101a36', 'a6b9ff') > 7);
  assert.ok(contrast('dce5f6', '121a29') > 7);
});

test('theme switch persists and Outlook button hands off only an ID', () => {
  const begin = app.indexOf('function applyTheme(theme)');
  const end = app.indexOf('function refreshCurrentUserControls(){',begin);
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
  vm.runInContext(app.slice(begin,end),context);
  context.switchTheme();
  assert.equal(context.document.documentElement.dataset.theme,'dark');
  assert.equal(store.get('personal_work_portal_theme'),'dark');
  context.syncOutlook();
  assert.equal(context.window.location.href,'workportal-outlook://sync/u-1');
  context.chooseOutlookCalendar();
  assert.equal(context.window.location.href,'workportal-outlook://choose/u-1');
});
