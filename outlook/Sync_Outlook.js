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
  function fail() { throw new Error('\u5171\u6709JSON\u306e\u5f62\u5f0f\u304c\u6b63\u3057\u304f\u3042\u308a\u307e\u305b\u3093\uff08\u6587\u5b57\u4f4d\u7f6e ' + pos + '\uff09\u3002'); }
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
  if (argument.substr(0, PROTOCOL.length).toLowerCase() !== PROTOCOL) throw new Error('\u8d77\u52d5\u30ea\u30f3\u30af\u304c\u6b63\u3057\u304f\u3042\u308a\u307e\u305b\u3093\u3002');
  var id;
  try { id = decodeURIComponent(argument.substr(PROTOCOL.length).replace(/\/$/, '')); }
  catch (e) { throw new Error('\u8d77\u52d5\u30ea\u30f3\u30af\u306e\u5229\u7528\u8005ID\u3092\u8aad\u307f\u53d6\u308c\u307e\u305b\u3093\u3002'); }
  if (!/^[A-Za-z0-9_-]{1,100}$/.test(id)) throw new Error('\u5229\u7528\u8005ID\u304c\u6b63\u3057\u304f\u3042\u308a\u307e\u305b\u3093\u3002');
  return id;
}
function isoDate(value) {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return null;
  var y = Number(value.substr(0, 4)), m = Number(value.substr(5, 2)), d = Number(value.substr(8, 2));
  var date = new Date(y, m - 1, d);
  if (date.getFullYear() !== y || date.getMonth() !== m - 1 || date.getDate() !== d) return null;
  return date;
}
function oleDate(date) {
  // Outlook's COM Date is days since 1899-12-30. Use UTC only to count
  // calendar days; the integer passed to Outlook represents local midnight.
  return (Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) -
    Date.UTC(1899, 11, 30)) / 86400000;
}
function safeId(value) {
  return typeof value === 'string' && /^[A-Za-z0-9_-]{1,100}$/.test(value) &&
    value !== '__proto__' && value !== 'constructor' && value !== 'prototype';
}
function desiredTasks(data, userId) {
  if (!data || data.format !== FORMAT || data.schemaVersion !== 6 ||
      !(data.tasks instanceof Array) || !(data.users instanceof Array)) throw new Error('\u5171\u6709JSON\u306e\u5f62\u5f0f\u307e\u305f\u306f\u7248\u304c\u5bfe\u5fdc\u3057\u3066\u3044\u307e\u305b\u3093\u3002');
  var found = false, i, t, dates = {}, result = {};
  for (i = 0; i < data.users.length; i++) if (data.users[i] && data.users[i].id === userId && data.users[i].active !== false) found = true;
  if (!found) throw new Error('\u9078\u629e\u3057\u305f\u5229\u7528\u8005\u304c\u5171\u6709JSON\u306b\u767b\u9332\u3055\u308c\u3066\u3044\u306a\u3044\u304b\u3001\u7121\u52b9\u306b\u306a\u3063\u3066\u3044\u307e\u3059\u3002');
  for (i = 0; i < data.tasks.length; i++) {
    t = data.tasks[i];
    if (!t || t.assigneeId !== userId || t.status === '\u5b8c\u4e86' || !t.dueDate) continue;
    if (!safeId(t.id) ||
        typeof t.title !== 'string' || !t.title) throw new Error('\u30bf\u30b9\u30afID\u307e\u305f\u306f\u4ef6\u540d\u304c\u6b63\u3057\u304f\u3042\u308a\u307e\u305b\u3093\u3002');
    if (Object.prototype.hasOwnProperty.call(result, t.id)) throw new Error('\u5171\u6709JSON\u5185\u3067\u30bf\u30b9\u30afID\u304c\u91cd\u8907\u3057\u3066\u3044\u307e\u3059\u3002');
    dates[t.id] = isoDate(t.dueDate);
    if (!dates[t.id]) throw new Error('\u30bf\u30b9\u30af\u306e\u671f\u9650\u304c\u6b63\u3057\u304f\u3042\u308a\u307e\u305b\u3093\uff1a' + t.id);
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
  var added = 0, updated = 0, stale = [], failed = 0, id, item, p, date, end, details = [], step;
  for (id in tasks) if (Object.prototype.hasOwnProperty.call(tasks, id)) {
    try {
      step = '\u4e88\u5b9a\u306e\u4f5c\u6210';
      // olAppointmentItem = 1. Use the item type rather than a custom form name.
      item = Object.prototype.hasOwnProperty.call(existing, id) ? existing[id] : calendar.Items.Add(1);
      date = tasks[id].date; end = new Date(date.getFullYear(), date.getMonth(), date.getDate() + 1);
      step = '\u4ef6\u540d\u306e\u8a2d\u5b9a'; item.Subject = tasks[id].title;
      step = '\u7d42\u65e5\u4e88\u5b9a\u306e\u8a2d\u5b9a'; item.AllDayEvent = true;
      step = '\u958b\u59cb\u65e5\u306e\u8a2d\u5b9a'; item.Start = oleDate(date);
      step = '\u7d42\u4e86\u65e5\u306e\u8a2d\u5b9a'; item.End = oleDate(end);
      step = '\u4e88\u5b9a\u306a\u3057\u306e\u8a2d\u5b9a'; item.BusyStatus = 0;
      step = '\u901a\u77e5\u306a\u3057\u306e\u8a2d\u5b9a'; item.ReminderSet = false;
      step = '\u30bf\u30b9\u30afID\u306e\u8a2d\u5b9a';
      p = userProperty(item, TAG); if (!p) p = item.UserProperties.Add(TAG, 1); p.Value = id;
      step = '\u5229\u7528\u8005ID\u306e\u8a2d\u5b9a';
      p = userProperty(item, OWNER_TAG); if (!p) p = item.UserProperties.Add(OWNER_TAG, 1); p.Value = userId;
      step = '\u4e88\u5b9a\u306e\u4fdd\u5b58';
      item.Save();
      if (Object.prototype.hasOwnProperty.call(existing, id)) updated++; else added++;
    } catch (e) { failed++; details.push(id + '\uff08' + step + '\uff09: ' + (e.message || String(e))); }
  }
  for (id in existing) if (Object.prototype.hasOwnProperty.call(existing, id) &&
      !Object.prototype.hasOwnProperty.call(tasks, id)) stale.push(existing[id]);
  var removed = 0, preview = [], i, decision;
  if (stale.length && !failed && !current.duplicates) {
    for (i = 0; i < stale.length && i < 8; i++) preview.push(String(stale[i].Subject));
    decision = shell.Popup('\u5bfe\u8c61\u304b\u3089\u5916\u308c\u305f\u9023\u643a\u6e08\u307f\u4e88\u5b9a\uff08' + stale.length + '\u4ef6\uff09\uff1a\n' + preview.join('\n') +
      '\n\n\u696d\u52d9\u30dd\u30fc\u30bf\u30eb\u306e\u4e88\u5b9a\u8868\u304b\u3089\u3001\u3053\u308c\u3089\u306e\u4e88\u5b9a\u3092\u524a\u9664\u3057\u307e\u3059\u304b\uff1f', 0, '\u696d\u52d9\u30dd\u30fc\u30bf\u30eb\uff1aOutlook\u9023\u643a', 4 + 32);
    if (decision === 6) for (i = 0; i < stale.length; i++) {
      try { stale[i].Delete(); removed++; } catch (e2) { failed++; details.push('\u4e88\u5b9a\u306e\u524a\u9664\uff1a' + e2.message); }
    }
  }
  return {added:added, updated:updated, removed:removed, pending:stale.length-removed,
    duplicates:current.duplicates, failed:failed, details:details};
}
function main() {
  if (WScript.Arguments.length !== 1) throw new Error('\u30dd\u30fc\u30bf\u30eb\u306e\u300cOutlook\u3078\u53cd\u6620\u300d\u30dc\u30bf\u30f3\u304b\u3089\u8d77\u52d5\u3057\u3066\u304f\u3060\u3055\u3044\u3002');
  var userId = targetUser(String(WScript.Arguments.Item(0)));
  var shell = new ActiveXObject('WScript.Shell');
  var path;
  try { path = shell.RegRead(REGISTRY); }
  catch (e) { throw new Error('Outlook\u9023\u643a\u306e\u8a2d\u5b9a\u304c\u3042\u308a\u307e\u305b\u3093\u3002\u30a4\u30f3\u30b9\u30c8\u30fc\u30e9\u30fc\u3092\u5b9f\u884c\u3057\u3066\u304f\u3060\u3055\u3044\u3002'); }
  if (!new ActiveXObject('Scripting.FileSystemObject').FileExists(path)) throw new Error('\u5171\u6709JSON\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002\u9023\u643a\u306e\u8a2d\u5b9a\u3092\u3084\u308a\u76f4\u3057\u3066\u304f\u3060\u3055\u3044\u3002');
  var tasks = desiredTasks(parseJSON(readUtf8(path)), userId);
  var outlook;
  try { outlook = new ActiveXObject('Outlook.Application'); }
  catch (e2) { throw new Error('\u5f93\u6765\u7248Outlook\u3092\u8d77\u52d5\u3067\u304d\u307e\u305b\u3093\u3002\u30a4\u30f3\u30b9\u30c8\u30fc\u30eb\u3068\u5229\u7528\u74b0\u5883\u3092\u78ba\u8a8d\u3057\u3066\u304f\u3060\u3055\u3044\u3002'); }
  var result = syncCalendar(createCalendar(outlook), tasks, userId, shell);
  var message = '\u8ffd\u52a0\uff1a' + result.added + '\u4ef6\n\u66f4\u65b0\uff1a' + result.updated + '\u4ef6\n\u524a\u9664\uff1a' + result.removed +
    '\u4ef6\n\u524a\u9664\u4fdd\u7559\uff1a' + result.pending + '\u4ef6\n\u91cd\u8907\u3092\u691c\u51fa\uff1a' + result.duplicates + '\u4ef6\n\u5931\u6557\uff1a' + result.failed + '\u4ef6';
  if (result.details.length) message += '\n\n' + result.details.slice(0, 5).join('\n');
  shell.Popup(message, 0, '\u696d\u52d9\u30dd\u30fc\u30bf\u30eb\uff1aOutlook\u9023\u643a', result.failed ? 16 : 64);
}
if (typeof WScript !== 'undefined') {
  try { main(); }
  catch (error) { new ActiveXObject('WScript.Shell').Popup(String(error.message || error), 0, '\u696d\u52d9\u30dd\u30fc\u30bf\u30eb\uff1aOutlook\u9023\u643a', 16); WScript.Quit(1); }
}
