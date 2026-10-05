/**
 * ที่เก็บข้อมูลกลางของหน้า "แฟ้มบันทึกไข้" (Google Sheets + Apps Script Web App)
 * วิธีติดตั้งดูใน SETUP.md
 */
const PASSCODE = 'เปลี่ยนรหัสนี้ก่อน-deploy';   // ต้องตรงกับรหัสที่กรอกในหน้าเว็บ
const DATA_SHEET = 'data';   // เก็บสถานะทั้งหมดเป็น JSON ใน A1
const LOG_SHEET = 'log';     // ตารางอ่านง่าย สร้างใหม่ทุกครั้งที่บันทึก

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function read_() {
  const v = sheet_(DATA_SHEET).getRange('A1').getValue();
  return v ? JSON.parse(v) : null;
}
function merge_(a, b) {
  if (!a) return b;
  if (!b) return a;
  const del = {};
  (a.del || []).concat(b.del || []).forEach(function (id) { del[id] = 1; });
  function uni(x, y) {
    const m = {};
    (x || []).concat(y || []).forEach(function (i) { if (!del[i.id]) m[i.id] = i; });
    return Object.keys(m).map(function (k) { return m[k]; });
  }
  const newer = (b.ts || 0) >= (a.ts || 0) ? b : a;
  return {
    temps: uni(a.temps, b.temps),
    doses: uni(a.doses, b.doses),
    del: Object.keys(del).map(Number),
    sym: newer.sym, chk: newer.chk, interval: newer.interval,
    ts: Math.max(a.ts || 0, b.ts || 0)
  };
}
function writeLog_(d) {
  const sh = sheet_(LOG_SHEET);
  sh.clear();
  const rows = [['ประเภท', 'วันเวลา', 'ค่า', 'ชื่อยา']];
  (d.temps || []).forEach(function (t) { rows.push(['วัดไข้', new Date(t.t), t.v, '']); });
  (d.doses || []).forEach(function (x) { rows.push(['กินยา', new Date(x.t), '', x.name || '']); });
  const body = rows.slice(1).sort(function (p, q) { return p[1] - q[1]; });
  const all = [rows[0]].concat(body);
  sh.getRange(1, 1, all.length, 4).setValues(all);
  sh.getRange(2, 2, Math.max(body.length, 1), 1).setNumberFormat('dd/MM/yyyy HH:mm');
}
function doGet(e) {
  if ((e.parameter.pass || '') !== PASSCODE) return out_({ error: 'bad pass' });
  return out_({ data: read_() });
}
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.pass !== PASSCODE) return out_({ error: 'bad pass' });
    const merged = merge_(read_(), b.data);
    sheet_(DATA_SHEET).getRange('A1').setValue(JSON.stringify(merged));
    writeLog_(merged);
    return out_({ ok: true, data: merged });
  } finally {
    lock.releaseLock();
  }
}
