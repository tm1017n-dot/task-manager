const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function helper() {
  const context = { Enumerator: function (items) {
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
  list.Add = () => {
    const props = new Map();
    const item = {
      Class: 26,
      UserProperties: {
        Find: name => props.get(name) || null,
        Add: name => { const p = { Value: '' }; props.set(name, p); return p; }
      },
      Save() { if (!list.includes(this)) list.push(this); },
      Delete() { list.splice(list.indexOf(this), 1); }
    };
    return item;
  };
  return { Items: list };
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
  const sample = fs.readFileSync('source/light/サンプルデータ/共有ダミーデータ.json', 'utf8');
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
  const old = c.Items.Add();
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
  const someoneElse = c.Items.Add();
  someoneElse.UserProperties.Add('WorkPortalTaskId').Value = 'other-task';
  someoneElse.UserProperties.Add('WorkPortalUserId').Value = 'u-2';
  someoneElse.Save();
  result = h.syncCalendar(c,tasks,'u-1',shell);
  assert.equal(result.added,0);
  assert.equal(result.updated,1);
  tasks['t-1'].date = new Date(2026,9,1);
  h.syncCalendar(c,tasks,'u-1',shell);
  assert.equal(c.Items[0].Start.getDate(),1);
  result = h.syncCalendar(c,{},'u-1',shell);
  assert.equal(result.pending,1);
  shell.answer = 6;
  result = h.syncCalendar(c,{},'u-1',shell);
  assert.equal(result.removed,1);
  assert.equal(c.Items.length,1);
  assert.equal(c.Items[0],someoneElse);
});
