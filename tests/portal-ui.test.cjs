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

test('master order survives normalization and moves only within the work category', async () => {
  const state = {workCategories:[{id:'a',name:'総務'},{id:'b',name:'検診'}],
    works:[{id:'w1',name:'予算',categoryId:'a'}, {id:'w2',name:'契約',categoryId:'a'},
      {id:'w3',name:'集計',categoryId:'b'}]};
  let saves = 0;
  const ctx = {state, uid:()=>{throw new Error('unexpected new ID')},
    byId:(arr,id)=>arr.find(x=>x.id===id), workCategory:w=>state.workCategories.find(c=>c.id===w?.categoryId),
    workCategoryId:w=>w?.categoryId||'', nowIso:()=> '2026-10-02T00:00:00Z',
    persist:async()=>{saves++;return true}, logActivity:()=>{}, refreshRelationOptions:()=>{},
    renderAll:()=>{},renderWorkCategoryManager:()=>{},toast:()=>{},esc:x=>String(x)};
  vm.createContext(ctx);
  vm.runInContext(app.slice(app.indexOf('function normalizeWorkCategory('),app.indexOf('function normalizeTask(')),ctx);
  const first = state.works.filter(x=>x.categoryId==='a').sort(ctx.compareMasterOrder);
  await ctx.moveMaster('works', first[1].id, -1);
  assert.equal(saves,1);
  assert.equal(state.works.find(x=>x.id===first[1].id).sortOrder,0);
  assert.equal(state.works.find(x=>x.id==='w3').sortOrder,undefined);
  const restored = state.works.map(x=>ctx.normalizeWork(JSON.parse(JSON.stringify(x))));
  assert.equal(restored.filter(x=>x.categoryId==='a').sort(ctx.compareMasterOrder)[0].id,first[1].id);
  const cats = state.workCategories.slice().sort(ctx.compareMasterOrder);
  await ctx.moveMaster('workCategories',cats[1].id,-1);
  assert.equal(state.workCategories.slice().sort(ctx.compareMasterOrder)[0].id,cats[1].id);
  assert.equal(restored.slice().sort(ctx.compareWorkOrder)[0].categoryId,cats[1].id);
  await ctx.moveMaster('workCategories',cats[1].id,-1);
  assert.equal(saves,2);
});

test('folder links hand off an encoded path while Web links still open in a tab', () => {
  const target = '\\\\server\\share\\日本語 & 100% #資料';
  const state = {links:[{id:'folder',target,category:'共有フォルダ'},
    {id:'web',target:'https://example.com/',category:'Web'}]};
  const opened = [];
  const ctx = {state,byId:(arr,id)=>arr.find(x=>x.id===id),
    linkCategoryKind:name=>name==='共有フォルダ'?'folder':['Web','庁内システム'].includes(name)?'web':'any',
    window:{location:{href:''},open:(...args)=>opened.push(args)},toast:()=>{throw new Error('unexpected toast')}};
  vm.createContext(ctx);
  const begin = app.indexOf('function isWebTarget(');
  const end = app.indexOf('\n',app.indexOf('function useLink(',begin));
  vm.runInContext(app.slice(begin,end),ctx);
  ctx.useLink('folder');
  assert.equal(ctx.window.location.href,'workportal-folder://open/'+encodeURIComponent(target));
  ctx.useLink('web');
  assert.equal(opened[0][0],'https://example.com/');
});

test('link category dropdown always includes shared folders and keeps custom categories', () => {
  assert.match(app, /<select id="linkCategory">/);
  assert.doesNotMatch(app, /list="linkCategorySuggestions"/);
  const elements = {
    linkCategory:{innerHTML:'',value:'Web'},
    linkCategoryCustom:{value:'',required:false,classList:{toggle(name,hidden){this.hidden=hidden;}}}
  };
  const ctx = {q:id=>elements[id],state:{links:[{category:'庁内資料'}]},
    categoryNames:()=>['共有フォルダ','庁内システム','Web','資料','その他','庁内資料'],
    esc:x=>String(x).replace(/&/g,'&amp;').replace(/"/g,'&quot;')};
  vm.createContext(ctx);
  vm.runInContext(app.slice(app.indexOf('function fillLinkCategorySelect('),app.indexOf("function openLink(id=")),ctx);
  ctx.fillLinkCategorySelect('Web');
  assert.match(elements.linkCategory.innerHTML, /value="共有フォルダ">共有フォルダ/);
  assert.match(elements.linkCategory.innerHTML, /value="庁内資料">庁内資料/);
  assert.equal(elements.linkCategory.value,'Web');
  elements.linkCategory.value='共有フォルダ';
  ctx.toggleLinkCategoryInput();
  assert.equal(ctx.selectedLinkCategory(),'共有フォルダ');
  assert.equal(elements.linkCategoryCustom.required,false);
  ctx.fillLinkCategorySelect('既存の独自分類');
  assert.equal(elements.linkCategory.value,'既存の独自分類');
  elements.linkCategory.value='';elements.linkCategoryCustom.value=' 新しい分類 ';
  ctx.toggleLinkCategoryInput();
  assert.equal(elements.linkCategoryCustom.required,true);
  assert.equal(ctx.selectedLinkCategory(),'新しい分類');
});
