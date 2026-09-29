const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function helper(onEnumerate = () => {}) {
  const context = { Enumerator: function (items) {
    onEnumerate(items);
    let index = 0;
    this.atEnd = () => index >= items.length;
    this.item = () => items[index];
    this.moveNext = () => { index++; };
  }};
  vm.createContext(context);
  vm.runInContext(fs.readFileSync('outlook/Sync_Outlook.js', 'ascii'), context);
  return context;
}
function calendar() {
  const list = [];
  const folder = {Items:list, EntryID:'portal-calendar', FolderPath:'\\\\mailbox\\予定表\\業務ポータル',
    StoreID:'mailbox-1', DefaultItemType:1, displayCount:0, Display() { this.displayCount++; }};
  list.Add = type => {
    assert.equal(type, 1);
    const props = new Map();
    const item = {
      Class: 26,
      UserProperties: {
        Find: name => props.get(name) || null,
        Add: name => { const p = { Value: '' }; props.set(name, p); return p; }
      },
      Parent:folder,
      EntryID:'appointment-' + (list.length + 1),
      Save() { if (!list.includes(this)) list.push(this); },
      Move(destination) { const old = this.Parent.Items, index = old.indexOf(this);
        if (index >= 0) old.splice(index, 1); this.Parent = destination;
        destination.Items.push(this); return this; },
      Delete() { list.splice(list.indexOf(this), 1); }
    };
    return item;
  };
  return folder;
}

test('strict JSON parsing rejects executable expressions and malformed data', () => {
  const h = helper();
  assert.equal(h.parseJSON('{"title":"日本語\\n作業","tasks":[1,true,null]}').title, '日本語\n作業');
  assert.throws(() => h.parseJSON('{"tasks":[], "x": (function(){})()}'));
  assert.throws(() => h.parseJSON('{"__proto__": {"polluted": true}}'));
});

test('only active selected user and valid incomplete tasks are eligible', () => {
  const h = helper();
  const data = { format:'personal-work-portal-shared', schemaVersion:6,
    users:[{id:'u-1',active:true}], tasks:[
      {id:'t-1',assigneeId:'u-1',title:'申請',dueDate:'2026-09-30',status:'未着手'},
      {id:'t-2',assigneeId:'u-2',title:'他人',dueDate:'2026-09-30'},
      {id:'t-3',assigneeId:'u-1',title:'完了',dueDate:'2026-09-30',status:'完了'},
      {id:'t-4',assigneeId:'u-1',title:'期限なし',dueDate:''}
    ]};
  const parsed = h.parseJSON(JSON.stringify(data));
  assert.deepEqual(Object.keys(h.desiredTasks(parsed,'u-1')),['t-1']);
  assert.throws(() => h.desiredTasks(parsed,'u-2'));
  parsed.tasks[0].dueDate = '2026-02-31';
  assert.throws(() => h.desiredTasks(parsed,'u-1'));
  parsed.tasks[0].id = '__proto__';
  assert.throws(() => h.desiredTasks(parsed,'u-1'));
});

test('bundled shared sample parses without executing script code', () => {
  const h = helper();
  const sample = fs.readFileSync('source/app/サンプルデータ/共有ダミーデータ.json', 'utf8');
  const data = h.parseJSON(sample);
  assert.equal(data.format,'personal-work-portal-shared');
  assert.ok(Array.isArray(data.tasks));
  assert.ok(Array.isArray(data.users));
});

test('sync errors and removal confirmation are shown in Japanese', () => {
  const h = helper(), c = calendar(), dialogs = [];
  const shell = { Popup(message, timeout, title) {
    dialogs.push({message, title});
    return 7;
  }};
  assert.throws(() => h.parseJSON('{'), /共有JSON/);
  assert.throws(() => h.desiredTasks({}, 'u-1'), /共有JSON/);
  const old = c.Items.Add(1);
  old.UserProperties.Add('WorkPortalTaskId').Value = 't-1';
  old.UserProperties.Add('WorkPortalUserId').Value = 'u-1';
  old.Subject = '古い予定';
  old.Save();
  h.syncCalendar(c, {}, 'u-1', shell);
  assert.match(dialogs[0].message, /対象から外れた連携済み予定/);
  assert.match(dialogs[0].title, /業務ポータル/);
});

test('calendar sync is repeatable, updates dates, and asks before cleanup', () => {
  const h = helper(), c = calendar();
  const shell = { answer:7, Popup() { return this.answer; } };
  const tasks = { 't-1':{title:'確認',date:new Date(2026,8,30)} };
  let result = h.syncCalendar(c,tasks,'u-1',shell);
  assert.equal(result.added,1);
  assert.equal(c.Items.length,1);
  assert.equal(c.Items[0].BusyStatus,0);
  assert.equal(c.Items[0].Body,undefined);
  let saves = 0;
  const originalSave = c.Items[0].Save;
  c.Items[0].Save = function () { saves++; return originalSave.call(this); };
  const someoneElse = c.Items.Add(1);
  someoneElse.UserProperties.Add('WorkPortalTaskId').Value = 'other-task';
  someoneElse.UserProperties.Add('WorkPortalUserId').Value = 'u-2';
  someoneElse.Save();
  result = h.syncCalendar(c,tasks,'u-1',shell);
  assert.equal(result.added,0);
  assert.equal(result.updated,0);
  assert.equal(result.skipped,1);
  assert.equal(saves,0);
  tasks['t-1'].date = new Date(2026,9,1);
  result = h.syncCalendar(c,tasks,'u-1',shell);
  assert.equal(result.updated,1);
  assert.equal(saves,1);
  assert.equal(c.Items[0].Start, h.oleDate(new Date(2026,9,1)));
  assert.equal(c.Items[0].End - c.Items[0].Start, 1);
  result = h.syncCalendar(c,{},'u-1',shell);
  assert.equal(result.pending,1);
  shell.answer = 6;
  result = h.syncCalendar(c,{},'u-1',shell);
  assert.equal(result.removed,1);
  assert.equal(c.Items.length,1);
  assert.equal(c.Items[0],someoneElse);
});

test('a failed Outlook operation identifies its stage and does not remove old appointments', () => {
  const h = helper(), c = calendar();
  const old = c.Items.Add(1);
  old.UserProperties.Add('WorkPortalTaskId').Value = 'old-task';
  old.UserProperties.Add('WorkPortalUserId').Value = 'u-1';
  old.Save();
  c.Items.Add = type => { assert.equal(type, 1); throw new Error('COM failure'); };
  const result = h.syncCalendar(c, {'t-contract':{title:'契約',date:new Date(2026,8,30)}}, 'u-1',
    {Popup() { throw new Error('cleanup must not run'); }});
  assert.equal(result.failed, 1);
  assert.match(result.details[0], /t-contract（予定の作成）: COM failure/);
  assert.equal(result.removed, 0);
  assert.equal(c.Items.length, 1);
});

test('date assignment uses COM date numbers and reports the exact failing property', () => {
  const h = helper(), c = calendar();
  const originalAdd = c.Items.Add;
  c.Items.Add = type => {
    const item = originalAdd(type);
    Object.defineProperty(item, 'Start', {set(value) {
      assert.equal(typeof value, 'number');
      throw new Error('COM date rejected');
    }});
    return item;
  };
  const result = h.syncCalendar(c, {'t-contract':{title:'契約',date:new Date(2026,8,30)}}, 'u-1',
    {Popup() { throw new Error('cleanup must not run'); }});
  assert.equal(result.failed, 1);
  assert.match(result.details[0], /t-contract（開始日の設定）: COM date rejected/);
  assert.equal(h.oleDate(new Date(1899,11,30)), 0);
  assert.equal(h.oleDate(new Date(2026,8,30)), 46295);
});

test('a saved item outside the portal calendar is moved and verified before counting', () => {
  const h = helper(), c = calendar(), defaultCalendar = calendar();
  defaultCalendar.EntryID = 'default-calendar';
  const originalAdd = c.Items.Add;
  c.Items.Add = type => {
    const item = originalAdd(type);
    item.Parent = defaultCalendar;
    item.Save = function () {
      if (!this.Parent.Items.includes(this)) this.Parent.Items.push(this);
    };
    return item;
  };
  const result = h.syncCalendar(c, {'t-contract':{title:'契約',date:new Date(2026,8,30)}}, 'u-1',
    {Popup() { throw new Error('unexpected dialog'); }});
  assert.equal(result.added,1);
  assert.equal(result.failed,0);
  assert.equal(defaultCalendar.Items.length,0);
  assert.equal(c.Items.length,1);
  assert.equal(c.Items[0].Parent,c);
});

test('a save that cannot be read back is reported as failed', () => {
  const h = helper(), c = calendar();
  const originalAdd = c.Items.Add;
  c.Items.Add = type => {
    const item = originalAdd(type);
    item.Save = () => {};
    return item;
  };
  const result = h.syncCalendar(c, {'t-contract':{title:'契約',date:new Date(2026,8,30)}}, 'u-1',
    {Popup() { throw new Error('unexpected dialog'); }});
  assert.equal(result.added,0);
  assert.equal(result.failed,1);
  assert.match(result.details[0], /保存後の再読み取り/);
});

test('a large sync scans the calendar once before and once after the batch', () => {
  const c = calendar();
  let scans = 0;
  const h = helper(items => { if (items === c.Items) scans++; });
  for (let i = 0; i < 200; i++) {
    const unrelated = c.Items.Add(1);
    unrelated.Subject = '別の予定';
    unrelated.Save();
  }
  const tasks = {};
  for (let i = 0; i < 50; i++) tasks['t-' + i] = {title:'作業' + i, date:new Date(2026,8,30)};
  const result = h.syncCalendar(c, tasks, 'u-1', {Popup() { throw new Error('unexpected dialog'); }});
  assert.equal(scans, 2);
  assert.equal(result.added, 50);
  assert.equal(result.failed, 0);
  assert.equal(c.Items.length, 250);
  scans = 0;
  const again = h.syncCalendar(c, tasks, 'u-1', {Popup() { throw new Error('unexpected dialog'); }});
  assert.equal(scans, 1);
  assert.equal(again.skipped, 50);
  assert.equal(again.updated, 0);
});

test('the selected Outlook calendar is remembered and used on the next sync', () => {
  const h = helper(), selected = calendar(), values = new Map();
  const shell = {Popup() { return 1; }, RegRead(key) {
    if (!values.has(key)) throw new Error('missing');
    return values.get(key);
  }, RegWrite(key, value) { values.set(key, value); }};
  const namespace = {
    picks:0,
    PickFolder() { this.picks++; return selected; },
    GetFolderFromID(id, store) {
      assert.equal(id,selected.EntryID); assert.equal(store,selected.StoreID);
      return selected;
    },
    GetDefaultFolder() { return {Folders:{Item() { throw new Error('no legacy folder'); }}}; }
  };
  assert.equal(h.chooseCalendar(namespace,shell,'u-1',false).folder,selected);
  assert.equal(h.chooseCalendar(namespace,shell,'u-1',false).folder,selected);
  assert.equal(namespace.picks,1);
  assert.equal(h.targetRequest('workportal-outlook://choose/u-1').choose,true);
  assert.equal(h.targetRequest('workportal-outlook://sync/u-1').choose,false);
});

test('changing the destination moves only the current user’s tagged appointments', () => {
  const h = helper(), legacy = calendar(), target = calendar();
  legacy.EntryID = 'old-calendar'; target.EntryID = 'new-calendar';
  const owned = legacy.Items.Add(1);
  owned.UserProperties.Add('WorkPortalTaskId').Value = 't-contract';
  owned.UserProperties.Add('WorkPortalUserId').Value = 'u-1'; owned.Save();
  const other = legacy.Items.Add(1);
  other.UserProperties.Add('WorkPortalTaskId').Value = 'other';
  other.UserProperties.Add('WorkPortalUserId').Value = 'u-2'; other.Save();
  const namespace = {PickFolder:()=>target, GetDefaultFolder:()=>({Folders:{Item:()=>legacy}})};
  const shell = {Popup:()=>1, RegWrite() {}};
  const selection = h.chooseCalendar(namespace,shell,'u-1',true);
  assert.equal(selection.migrated,1);
  assert.equal(target.Items.length,1);
  assert.equal(target.Items[0],owned);
  assert.deepEqual([...legacy.Items],[other]);
});
