// Windows Script Host JScript. Folder-only protocol handler; no command shell.
function folderPath(argument) {
  var prefix = 'workportal-folder://open/', path;
  if (String(argument).substr(0, prefix.length).toLowerCase() !== prefix)
    throw new Error('\u8d77\u52d5\u30ea\u30f3\u30af\u304c\u6b63\u3057\u304f\u3042\u308a\u307e\u305b\u3093\u3002');
  try { path = decodeURIComponent(String(argument).substr(prefix.length)); }
  catch (e) { throw new Error('\u30d5\u30a9\u30eb\u30c0\u30fc\u306e\u30d1\u30b9\u3092\u8aad\u3081\u307e\u305b\u3093\u3002'); }
  if (/^file:\/\//i.test(path)) {
    path = path.substr(7);
    try { path = decodeURIComponent(path); } catch (e2) { throw new Error('\u30d1\u30b9\u304c\u6b63\u3057\u304f\u3042\u308a\u307e\u305b\u3093\u3002'); }
    path = path.replace(/^localhost\//i, '/');
    if (/^\/[A-Za-z]:\//.test(path)) path = path.substr(1);
    else if (!/^[A-Za-z]:\//.test(path)) path = '\\\\' + path.replace(/^\/+/, '');
    path = path.replace(/\//g, '\\');
  }
  if (path.length > 1000 || /[\x00-\x1f"<>|*?]/.test(path) ||
      /^\\\\\.{1,2}\\/.test(path) ||
      !(/^[A-Za-z]:\\/.test(path) || /^\\\\[^\\:]+\\[^\\:]+/.test(path)) ||
      path.substr(2).indexOf(':') !== -1)
    throw new Error('\u30d5\u30a9\u30eb\u30c0\u30fc\u306e\u7d76\u5bfe\u30d1\u30b9\u3092\u6307\u5b9a\u3057\u3066\u304f\u3060\u3055\u3044\u3002');
  return path;
}
function openFolder(argument, fso, shell) {
  var path = folderPath(argument);
  if (!fso.FolderExists(path))
    throw new Error('\u30d5\u30a9\u30eb\u30c0\u30fc\u304c\u898b\u3064\u304b\u3089\u306a\u3044\u304b\u3001\u30a2\u30af\u30bb\u30b9\u3067\u304d\u307e\u305b\u3093\u3002\n' + path);
  shell.ShellExecute(path, '', '', 'open', 1);
}
if (typeof WScript !== 'undefined') {
  try {
    if (WScript.Arguments.length !== 1) throw new Error('\u30dd\u30fc\u30bf\u30eb\u306e\u30ea\u30f3\u30af\u304b\u3089\u958b\u3044\u3066\u304f\u3060\u3055\u3044\u3002');
    openFolder(String(WScript.Arguments.Item(0)), new ActiveXObject('Scripting.FileSystemObject'), new ActiveXObject('Shell.Application'));
  } catch (error) {
    new ActiveXObject('WScript.Shell').Popup(String(error.message || error), 0, '\u696d\u52d9\u30dd\u30fc\u30bf\u30eb\uff1a\u30d5\u30a9\u30eb\u30c0\u30fc\u3092\u958b\u304f', 16);
    WScript.Quit(1);
  }
}
