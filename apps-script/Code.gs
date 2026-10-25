/**
 * ที่นี่ดี — ฝั่งเก็บข้อมูล (Google Apps Script ผูกกับ Google Sheet)
 *
 * วิธีติดตั้ง (ทำครั้งเดียว ดูละเอียดใน README.md):
 * 1. สร้าง Google Sheet ใหม่ → ส่วนขยาย → Apps Script → วางไฟล์นี้แทนของเดิม → บันทึก
 * 2. เลือกฟังก์ชัน setup แล้วกด "เรียกใช้" → อนุญาตสิทธิ์ → ดู "รหัสเข้าแอป" ในบันทึกการดำเนินการ
 * 3. ทำให้ใช้งานได้ → การทำให้ใช้งานได้รายการใหม่ → ประเภท: เว็บแอป
 *    เรียกใช้ในฐานะ: ฉัน · ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน → คัดลอก URL ไปใส่ในแอป
 *
 * ข้อมูลรีวิวอยู่ในชีต "Reviews" (หนึ่งแถวต่อหนึ่งรีวิว) รูปอยู่ในโฟลเดอร์ "ที่นี่ดี รูป" ใน Drive
 */

const SHEET_NAME = 'Reviews';
const FOLDER_NAME = 'ที่นี่ดี รูป';
const HEAD = ['id', 'อัปเดต', 'โดย', 'ประเภท', 'ชื่อ', 'คะแนน', 'ข้อมูล (JSON)'];
const MAX_PHOTO_BYTES = 8 * 1024 * 1024;

/** เรียกครั้งแรกจากหน้า Apps Script เพื่อสร้างชีต โฟลเดอร์รูป และรหัสเข้าแอป */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('KEY')) props.setProperty('KEY', Utilities.getUuid().replace(/-/g, '').slice(0, 10));
  sheet_();
  folder_();
  Logger.log('รหัสเข้าแอป: ' + props.getProperty('KEY'));
  Logger.log('โฟลเดอร์รูป: ' + folder_().getUrl());
}

/** อยากเปลี่ยนรหัส (เช่น เผลอส่งให้คนอื่น) ให้เรียกฟังก์ชันนี้ แล้วใส่รหัสใหม่ในแอปทั้งสองเครื่อง */
function resetKey() {
  PropertiesService.getScriptProperties().deleteProperty('KEY');
  setup();
}

function doGet() {
  return json_({ ok: true, app: 'thineedee', note: 'แอปใช้ POST เท่านั้น ถ้าเห็นข้อความนี้แปลว่า URL ถูกต้อง' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'คำขอไม่ถูกรูปแบบ' }); }
  const key = PropertiesService.getScriptProperties().getProperty('KEY');
  if (!key) return json_({ ok: false, error: 'ยังไม่ได้รัน setup ในหน้า Apps Script' });
  if (req.key !== key) return json_({ ok: false, error: 'รหัสเข้าแอปไม่ถูกต้อง', code: 'KEY' });

  try {
    switch (req.action) {
      case 'ping': return json_({ ok: true });
      case 'list': return json_({ ok: true, items: list_() });
      case 'save': return withLock_(() => json_({ ok: true, item: save_(req.item, req.who) }));
      case 'delete': return withLock_(() => json_({ ok: true, id: remove_(req.id, req.who) }));
      case 'upload': return json_({ ok: true, photo: upload_(req) });
      default: return json_({ ok: false, error: 'ไม่รู้จักคำสั่ง ' + req.action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ---------- actions ---------- */

function list_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, HEAD.length).getValues()
    .map(r => { try { return JSON.parse(r[6]); } catch (e) { return null; } })
    .filter(x => x && x.id && !x.deleted);
}

function save_(item, who) {
  if (!item || !item.name) throw new Error('ต้องมีชื่อสถานที่');
  const sh = sheet_();
  const now = new Date().toISOString();
  item.id = item.id || Utilities.getUuid();
  item.updatedAt = now;
  item.createdAt = item.createdAt || now;
  item.author = item.author || who || '';
  item.editedBy = who || '';
  const row = findRow_(sh, item.id);
  const values = [[item.id, now, who || '', item.type || '', item.name, item.rating || '', JSON.stringify(item)]];
  if (row) sh.getRange(row, 1, 1, HEAD.length).setValues(values);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, HEAD.length).setValues(values);
  return item;
}

function remove_(id, who) {
  const sh = sheet_();
  const row = findRow_(sh, id);
  if (!row) return id;
  let item = {};
  try { item = JSON.parse(sh.getRange(row, 7).getValue()); } catch (e) {}
  (item.photos || []).forEach(p => { try { DriveApp.getFileById(p).setTrashed(true); } catch (e) {} });
  item.deleted = true;
  item.updatedAt = new Date().toISOString();
  item.editedBy = who || '';
  sh.getRange(row, 2, 1, 6).setValues([[item.updatedAt, who || '', item.type || '', '(ลบแล้ว) ' + (item.name || ''), '', JSON.stringify(item)]]);
  return id;
}

function upload_(req) {
  const bytes = Utilities.base64Decode(req.b64 || '');
  if (!bytes.length) throw new Error('ไม่มีข้อมูลรูป');
  if (bytes.length > MAX_PHOTO_BYTES) throw new Error('รูปใหญ่เกินไป');
  const mime = /^image\/(jpeg|png|webp)$/.test(req.mime) ? req.mime : 'image/jpeg';
  const file = folder_().createFile(Utilities.newBlob(bytes, mime, req.name || ('photo-' + Date.now() + '.jpg')));
  // เปิดให้ "ทุกคนที่มีลิงก์ดูได้" เพื่อให้แอปแสดงรูปได้ ลิงก์เป็นรหัสสุ่มยาว คนนอกเดาไม่ได้
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getId();
}

/* ---------- helpers ---------- */

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(5, 260);
    sh.setColumnWidth(7, 400);
  }
  return sh;
}

function folder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  const f = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}

function findRow_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return 0;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
