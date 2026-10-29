/**
 * ระบบกลางเก็บสถานะอนุมัติ (Google Apps Script)
 * - เก็บสถานะ ผู้อนุมัติ 1 / ผู้อนุมัติ 2 / บัญชี ไว้ในชีต "approvals" ของ Google Sheet นี้
 * - ตรวจ PIN ที่ฝั่งเซิร์ฟเวอร์ (PIN เก็บใน Script Properties ไม่อยู่ในโค้ดหน้าเว็บ)
 * - วิธีติดตั้งดูในไฟล์ "วิธีตั้งค่าระบบกลาง.md"
 *
 * Script Properties ที่ต้องตั้ง (Project Settings > Script properties):
 *   PIN_approver1, PIN_approver2, PIN_accountant
 */
const SHEET_NAME = 'approvals';
const FIELDS = ['approver1', 'approver2', 'accountant'];
const TOKEN_TTL = 6 * 60 * 60;   // วินาที (เข้าสู่ระบบอยู่ได้ 6 ชม.)
const MAX_FAILS = 5;             // ใส่ PIN ผิดได้กี่ครั้งต่อบทบาท
const LOCK_SECONDS = 600;        // ถ้าเกิน จะล็อก 10 นาที

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['key'].concat(FIELDS).concat(['updated']));
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll_() {
  const sh = sheet_(), n = sh.getLastRow();
  const out = {};
  if (n < 2) return out;
  sh.getRange(2, 1, n - 1, FIELDS.length + 1).getValues().forEach(r => {
    if (!r[0]) return;
    const rec = {};
    FIELDS.forEach((f, i) => { rec[f] = r[i + 1] ? String(r[i + 1]) : null; });
    out[String(r[0])] = rec;
  });
  return out;
}

function findRow_(sh, key) {
  const n = sh.getLastRow();
  if (n < 2) return -1;
  const keys = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) if (String(keys[i][0]) === key) return i + 2;
  return -1;
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'approvals';
  if (action === 'approvals') return json_({ ok: true, approvals: readAll_() });
  return json_({ ok: false, error: 'unknown_action' });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  if (b.action === 'login') return login_(b);
  if (b.action === 'approve') return approve_(b);
  return json_({ ok: false, error: 'unknown_action' });
}

function login_(b) {
  const role = String(b.role || '');
  if (FIELDS.indexOf(role) < 0) return json_({ ok: false, error: 'pin' });
  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + role;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_FAILS) return json_({ ok: false, error: 'locked' });
  const pin = PropertiesService.getScriptProperties().getProperty('PIN_' + role);
  if (!pin || String(b.pin) !== String(pin)) {
    cache.put(failKey, String(fails + 1), LOCK_SECONDS);
    Utilities.sleep(700);
    return json_({ ok: false, error: 'pin' });
  }
  cache.remove(failKey);
  const token = Utilities.getUuid();
  cache.put('tok_' + token, role, TOKEN_TTL);
  return json_({ ok: true, token: token, role: role });
}

function approve_(b) {
  const role = CacheService.getScriptCache().get('tok_' + String(b.token || ''));
  if (!role) return json_({ ok: false, error: 'auth' });
  const key = String(b.key || '');
  const field = String(b.field || '');
  if (!key || key.length > 120) return json_({ ok: false, error: 'bad_request' });
  if (field !== role) return json_({ ok: false, error: 'forbidden' });   // แต่ละคนแก้ได้เฉพาะของตัวเอง
  const on = !!b.value;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    let row = findRow_(sh, key);
    const rec = {};
    if (row > 0) {
      const v = sh.getRange(row, 1, 1, FIELDS.length + 1).getValues()[0];
      FIELDS.forEach((f, i) => { rec[f] = v[i + 1] ? String(v[i + 1]) : null; });
    } else {
      FIELDS.forEach(f => { rec[f] = null; });
    }
    if (field === 'accountant') {
      if (on && !(rec.approver1 && rec.approver2)) return json_({ ok: false, error: 'need_both' });
    } else if (!on && rec.accountant) {
      return json_({ ok: false, error: 'paid_locked' });
    }
    if (on) { if (!rec[field]) rec[field] = new Date().toISOString(); }   // เวลาใช้ของเซิร์ฟเวอร์
    else rec[field] = null;

    const line = [key].concat(FIELDS.map(f => rec[f] || '')).concat([new Date().toISOString()]);
    if (row > 0) sh.getRange(row, 1, 1, line.length).setValues([line]);
    else sh.appendRow(line);
    return json_({ ok: true, record: rec });
  } finally {
    lock.releaseLock();
  }
}
