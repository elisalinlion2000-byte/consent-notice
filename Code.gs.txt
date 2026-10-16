/**
 * 各式同意書撤紙本｜進度表後台（Google Apps Script）
 * 貼到「試算表 → 擴充功能 → Apps Script」，改下面兩個設定後部署為網頁應用程式。
 */
var SHEET_ID   = '1KfQh7660H0KZQPGYXYZfO53Ibw8qpTDyBTBa3vR90ZM'; // 試算表 ID
var SHEET_NAME = '進度表_匯入用';                                    // 分頁名稱
var EDIT_KEY   = '';   // 留空＝不需要密碼，知道網頁的人都能打勾與編輯；若要加密碼，填在這裡（網頁會自動詢問）

// 欄位：A 類別｜B 項目｜C 主管佈達｜D 主管時間｜E 二力佈達｜F 二力時間｜G 產品佈達｜H 產品時間｜I 備註｜J id（自動產生，請勿改）
var COL = { type: 1, name: 2, boss: 3, bossT: 4, lead: 5, leadT: 6, prod: 7, prodT: 8, note: 9, id: 10 };
var TZ = 'Asia/Taipei';

function sheet_() {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('找不到分頁：' + SHEET_NAME);
  return sh;
}
function nowStr_() { return Utilities.formatDate(new Date(), TZ, 'MM/dd HH:mm'); }
function bool_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function time_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'MM/dd HH:mm');
  return v === null || v === undefined ? '' : String(v);
}
function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

/** 第一次使用請在編輯器手動執行一次：補 id、設定時間欄為文字、加上勾選框 */
function setup() {
  var sh = sheet_();
  sh.getRange(1, COL.id).setValue('id');
  fixFormat_(sh);
}
function fixFormat_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return;
  var n = last - 1;
  [COL.bossT, COL.leadT, COL.prodT].forEach(function (c) { sh.getRange(2, c, n, 1).setNumberFormat('@'); });
  var ids = sh.getRange(2, COL.id, n, 1).getValues();
  var data = sh.getRange(2, 1, n, 2).getValues();
  var changed = false;
  for (var i = 0; i < n; i++) {
    if (data[i][0] && data[i][1] && !ids[i][0]) { ids[i][0] = newId_(); changed = true; }
  }
  if (changed) sh.getRange(2, COL.id, n, 1).setValues(ids);
  var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  for (var j = 0; j < n; j++) {
    var row = j + 2, type = String(data[j][0]).trim();
    if (type === '待辦') sh.getRange(row, COL.boss).setDataValidation(rule);
    if (type === '線別') [COL.boss, COL.lead, COL.prod].forEach(function (c) { sh.getRange(row, c).setDataValidation(rule); });
  }
}

function read_() {
  var sh = sheet_();
  var last = sh.getLastRow();
  var out = { todos: [], lines: [] };
  if (last < 2) return out;
  var vals = sh.getRange(2, 1, last - 1, COL.id).getValues();
  var missing = false;
  vals.forEach(function (r) { if (r[0] && r[1] && !r[COL.id - 1]) missing = true; });
  if (missing) { fixFormat_(sh); vals = sh.getRange(2, 1, last - 1, COL.id).getValues(); }
  vals.forEach(function (r) {
    var type = String(r[0]).trim(), id = String(r[COL.id - 1] || '');
    if (!id || !r[1]) return;
    if (type === '待辦') {
      out.todos.push({ id: id, t: String(r[1]), done: bool_(r[COL.boss - 1]), at: time_(r[COL.bossT - 1]) });
    } else if (type === '線別') {
      out.lines.push({
        id: id, n: String(r[1]), note: String(r[COL.note - 1] || ''),
        boss: { done: bool_(r[COL.boss - 1]), at: time_(r[COL.bossT - 1]) },
        lead: { done: bool_(r[COL.lead - 1]), at: time_(r[COL.leadT - 1]) },
        prod: { done: bool_(r[COL.prod - 1]), at: time_(r[COL.prodT - 1]) }
      });
    }
  });
  return out;
}

function findRow_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2 || !id) return 0;
  var ids = sh.getRange(2, COL.id, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}
function clip_(s, n) { return String(s === null || s === undefined ? '' : s).replace(/[\r\n]+/g, ' ').trim().slice(0, n); }

function apply_(b) {
  var sh = sheet_(), rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  var op = b.op, row;
  if (op === 'toggle') {
    row = findRow_(sh, b.id);
    if (!row) throw new Error('找不到這一列，請重新整理頁面');
    var f = { done: [COL.boss, COL.bossT], boss: [COL.boss, COL.bossT], lead: [COL.lead, COL.leadT], prod: [COL.prod, COL.prodT] }[b.field];
    if (!f) throw new Error('欄位錯誤');
    var on = b.value === true;
    sh.getRange(row, f[1]).setNumberFormat('@').setValue(on ? nowStr_() : '');
    sh.getRange(row, f[0]).setDataValidation(rule).setValue(on);
  } else if (op === 'addTodo') {
    var t = clip_(b.t, 120);
    if (!t) throw new Error('內容不能空白');
    var lastTodo = 1, last = sh.getLastRow();
    if (last >= 2) {
      var types = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < types.length; i++) if (String(types[i][0]).trim() === '待辦') lastTodo = i + 2;
    }
    sh.insertRowAfter(lastTodo);
    row = lastTodo + 1;
    sh.getRange(row, COL.bossT).setNumberFormat('@');
    sh.getRange(row, 1, 1, 2).setValues([['待辦', t]]);
    sh.getRange(row, COL.boss).setDataValidation(rule).setValue(false);
    sh.getRange(row, COL.id).setValue(newId_());
  } else if (op === 'addLine') {
    var n = clip_(b.n, 40);
    if (!n) throw new Error('名稱不能空白');
    row = sh.getLastRow() + 1;
    [COL.bossT, COL.leadT, COL.prodT].forEach(function (c) { sh.getRange(row, c).setNumberFormat('@'); });
    sh.getRange(row, 1, 1, 2).setValues([['線別', n]]);
    [COL.boss, COL.lead, COL.prod].forEach(function (c) { sh.getRange(row, c).setDataValidation(rule).setValue(false); });
    sh.getRange(row, COL.id).setValue(newId_());
  } else if (op === 'rename') {
    row = findRow_(sh, b.id);
    if (!row) throw new Error('找不到這一列，請重新整理頁面');
    var type = String(sh.getRange(row, 1).getValue()).trim();
    var v = clip_(b.value, type === '待辦' ? 120 : 40);
    if (!v) throw new Error('內容不能空白');
    sh.getRange(row, COL.name).setValue(v);
  } else if (op === 'delete') {
    row = findRow_(sh, b.id);
    if (row) sh.deleteRow(row);
  } else {
    throw new Error('不支援的動作');
  }
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  try { return json_({ ok: true, data: read_() }); }
  catch (err) { return json_({ ok: false, error: String(err.message || err) }); }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var b = JSON.parse(e.postData.contents);
    if (!b || (EDIT_KEY && b.key !== EDIT_KEY)) return json_({ ok: false, error: 'badkey' });
    lock.waitLock(15000);
    apply_(b);
    return json_({ ok: true, data: read_() });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
