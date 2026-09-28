// Windows Script Host JScript. Keep this file ASCII for WSH on Japanese Windows.
// The browser sends only a user ID; the shared JSON path comes from the local installer.
var PROTOCOL = 'workportal-outlook://sync/';
var REGISTRY = 'HKCU\\Software\\WorkPortalOutlook\\SharedJsonPath';
var TAG = 'WorkPortalTaskId';
var OWNER_TAG = 'WorkPortalUserId';
var CALENDAR_NAME = '\u696d\u52d9\u30dd\u30fc\u30bf\u30eb';
var FORMAT = 'personal-work-portal-shared';

// A small strict JSON parser: WSH JScript does not consistently expose JSON.parse.
// Never eval data from a shared file.
function parseJSON(text) {
  var pos = 0;
  function fail() { throw new Error('Invalid shared JSON at character ' + pos); }
  function space() { while (pos < text.length && /[\x20\t\r\n]/.test(text.charAt(pos))) pos++; }
  function string() {
    var out = '', c, hex, n;
    if (text.charAt(pos++) !== '"') fail();
    while (pos < text.length) {
      c = text.charAt(pos++);
      if (c === '"') return out;
      if (c === '\\') {
        c = text.charAt(pos++);
        if (c === 'u') {
          hex = text.substr(pos, 4);
          if (!/^[0-9a-fA-F]{4}$/.test(hex)) fail();
          out += String.fromCharCode(parseInt(hex, 16)); pos += 4;
        } else {
          n = {'"':'"','\\':'\\','/':'/','b':'\b','f':'\f','n':'\n','r':'\r','t':'\t'};
          if (!Object.prototype.hasOwnProperty.call(n, c)) fail();
          out += n[c];
        }
      } else {
        if (c.charCodeAt(0) < 32) fail();
        out += c;
      }
    }
    fail();
  }
  function value() {
    var c, out, key, match;
    space(); c = text.charAt(pos);
    if (c === '"') return string();
    if (c === '{') {
      pos++; out = {}; space();
      if (text.charAt(pos) === '}') { pos++; return out; }
      while (true) {
        space(); if (text.charAt(pos) !== '"') fail();
        key = string(); space(); if (text.charAt(pos++) !== ':') fail();
        if (key === '__proto__' || key === 'constructor' || key === 'prototype') fail();
        out[key] = value(); space(); c = text.charAt(pos++);
        if (c === '}') return out;
        if (c !== ',') fail();
      }
    }
    if (c === '[') {
      pos++; out = []; space();
      if (text.charAt(pos) === ']') { pos++; return out; }
      while (true) {
        out.push(value()); space(); c = text.charAt(pos++);
        if (c === ']') return out;
        if (c !== ',') fail();
      }
    }
    if (text.substr(pos, 4) === 'true') { pos += 4; return true; }
    if (text.substr(pos, 5) === 'false') { pos += 5; return false; }
    if (text.substr(pos, 4) === 'null') { pos += 4; return null; }
    match = /^-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?/.exec(text.substr(pos));
    if (match) { pos += match[0].length; return Number(match[0]); }
    fail();
  }
  if (text.charAt(0) === '\uFEFF') text = text.substr(1);
  var result = value(); space(); if (pos !== text.length) fail();
  return result;
}

function targetUser(argument) {
  if (argument.substr(0, PROTOCOL.length).toLowerCase() !== PROTOCOL) throw new Error('Invalid launch link.');
  var id = decodeURIComponent(argument.substr(PROTOCOL.length).replace(/\/$/, ''));
  if (!/^[A-Za-z0-9_-]{1,100}$/.test(id)) throw new Error('Invalid user ID.');
  return id;
}
function isoDate(value) {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return null;
  var y = Number(value.substr(0, 4)), m = Number(value.substr(5, 2)), d = Number(value.substr(8, 2));
  var date = new Date(y, m - 1, d);
  if (date.getFullYear() !== y || date.getMonth() !== m - 1 || date.getDate() !== d) return null;
  return date;
}
function safeId(value) {
  return typeof value === 'string' && /^[A-Za-z0-9_-]{1,100}$/.test(value) &&
    value !== '__proto__' && value !== 'constructor' && value !== 'prototype';
}
function desiredTasks(data, userId) {
  if (!data || data.format !== FORMAT || data.schemaVersion !== 6 ||
      !(data.tasks instanceof Array) || !(data.users instanceof Array)) throw new Error('Unsupported shared JSON format or schema.');
  var found = false, i, t, dates = {}, result = {};
  for (i = 0; i < data.users.length; i++) if (data.users[i] && data.users[i].id === userId && data.users[i].active !== false) found = true;
  if (!found) throw new Error('Selected user is not active in the shared JSON.');
  for (i = 0; i < data.tasks.length; i++) {
    t = data.tasks[i];
    if (!t || t.assigneeId !== userId || t.status === '\u5b8c\u4e86' || !t.dueDate) continue;
    if (!safeId(t.id) ||
        typeof t.title !== 'string' || !t.title) throw new Error('Invalid task ID or title.');
    if (Object.prototype.hasOwnProperty.call(result, t.id)) throw new Error('Duplicate task ID in shared JSON.');
    dates[t.id] = isoDate(t.dueDate);
    if (!dates[t.id]) throw new Error('Invalid due date for task ' + t.id);
    result[t.id] = {title:t.title, date:dates[t.id]};
  }
  return result;
}
function readUtf8(path) {
  var stream = new ActiveXObject('ADODB.Stream');
  stream.Type = 2; stream.Charset = 'utf-8'; stream.Open();
  try { stream.LoadFromFile(path); return stream.ReadText(-1); }
  finally { stream.Close(); }
}
function userProperty(item, name) {
  try { return item.UserProperties.Find(name); } catch (e) { return null; }
}
function createCalendar(outlook) {
  var root = outlook.GetNamespace('MAPI').GetDefaultFolder(9);
  try { return root.Folders.Item(CALENDAR_NAME); }
  catch (e) { return root.Folders.Add(CALENDAR_NAME, 9); }
}
function readExisting(calendar, userId) {
  var items = {}, duplicates = 0, item, id, owner, en = new Enumerator(calendar.Items);
  for (; !en.atEnd(); en.moveNext()) {
    item = en.item();
    if (item.Class !== 26) continue;
    id = userProperty(item, TAG); owner = userProperty(item, OWNER_TAG);
    if (!id || !owner || String(owner.Value) !== userId) continue;
    if (!safeId(String(id.Value))) { duplicates++; continue; }
    if (Object.prototype.hasOwnProperty.call(items, String(id.Value))) duplicates++;
    else items[String(id.Value)] = item;
  }
  return {items:items, duplicates:duplicates};
}
function syncCalendar(calendar, tasks, userId, shell) {
  var current = readExisting(calendar, userId), existing = current.items;
  var added = 0, updated = 0, stale = [], failed = 0, id, item, p, date, end, details = [];
  for (id in tasks) if (Object.prototype.hasOwnProperty.call(tasks, id)) {
    try {
      item = Object.prototype.hasOwnProperty.call(existing, id) ? existing[id] : calendar.Items.Add('IPM.Appointment');
      date = tasks[id].date; end = new Date(date.getFullYear(), date.getMonth(), date.getDate() + 1);
      item.Subject = tasks[id].title; item.AllDayEvent = true;
      item.Start = date; item.End = end; item.BusyStatus = 0; item.ReminderSet = false;
      p = userProperty(item, TAG); if (!p) p = item.UserProperties.Add(TAG, 1, false); p.Value = id;
      p = userProperty(item, OWNER_TAG); if (!p) p = item.UserProperties.Add(OWNER_TAG, 1, false); p.Value = userId;
      item.Save();
      if (Object.prototype.hasOwnProperty.call(existing, id)) updated++; else added++;
    } catch (e) { failed++; details.push(id + ': ' + e.message); }
  }
  for (id in existing) if (Object.prototype.hasOwnProperty.call(existing, id) &&
      !Object.prototype.hasOwnProperty.call(tasks, id)) stale.push(existing[id]);
  var removed = 0, preview = [], i, decision;
  if (stale.length && !failed && !current.duplicates) {
    for (i = 0; i < stale.length && i < 8; i++) preview.push(String(stale[i].Subject));
    decision = shell.Popup('No longer assigned/open (' + stale.length + '):\n' + preview.join('\n') +
      '\n\nRemove these linked appointments from the Work Portal calendar?', 0, 'Work Portal Outlook sync', 4 + 32);
    if (decision === 6) for (i = 0; i < stale.length; i++) {
      try { stale[i].Delete(); removed++; } catch (e2) { failed++; details.push('cleanup: ' + e2.message); }
    }
  }
  return {added:added, updated:updated, removed:removed, pending:stale.length-removed,
    duplicates:current.duplicates, failed:failed, details:details};
}
function main() {
  if (WScript.Arguments.length !== 1) throw new Error('Launch this helper from the portal button.');
  var userId = targetUser(String(WScript.Arguments.Item(0)));
  var shell = new ActiveXObject('WScript.Shell');
  var path = shell.RegRead(REGISTRY);
  if (!new ActiveXObject('Scripting.FileSystemObject').FileExists(path)) throw new Error('Shared JSON is missing. Run the installer again.');
  var tasks = desiredTasks(parseJSON(readUtf8(path)), userId);
  var outlook = new ActiveXObject('Outlook.Application');
  var result = syncCalendar(createCalendar(outlook), tasks, userId, shell);
  var message = 'Added: ' + result.added + '\nUpdated: ' + result.updated + '\nRemoved: ' + result.removed +
    '\nPending review: ' + result.pending + '\nDuplicate appointments: ' + result.duplicates + '\nFailed: ' + result.failed;
  if (result.details.length) message += '\n\n' + result.details.slice(0, 5).join('\n');
  shell.Popup(message, 0, 'Work Portal Outlook sync', result.failed ? 16 : 64);
}
if (typeof WScript !== 'undefined') {
  try { main(); }
  catch (error) { new ActiveXObject('WScript.Shell').Popup(String(error.message || error), 0, 'Work Portal Outlook sync', 16); WScript.Quit(1); }
}
