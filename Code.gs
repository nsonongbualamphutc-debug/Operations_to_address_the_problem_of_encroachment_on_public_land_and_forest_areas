/**********************************************************************
 * ระบบติดตามการแก้ไขปัญหาการบุกรุกที่ดินสาธารณประโยชน์ และการบุกรุกพื้นที่ป่าไม้
 * จังหวัดหนองบัวลำภู
 * Backend : Google Apps Script + Google Sheets
 * รูปแบบ : JSONP (กัน CORS) | LockService | Upsert ด้วย id | PIN เก็บใน Script Properties
 * --------------------------------------------------------------------
 * วิธีติดตั้ง
 * 1) สร้าง Google Sheet ใหม่ คัดลอก ID จาก URL มาใส่ SHEET_ID ด้านล่าง
 * 2) Extensions > Apps Script วางโค้ดนี้ทั้งหมด
 * 3) รันฟังก์ชัน setup() หนึ่งครั้ง (สร้างชีตทั้งหมด + หัวตาราง)
 * 4) รันฟังก์ชัน seedSampleData() ถ้าต้องการข้อมูลตัวอย่างจากแบบฟอร์มที่ให้มา (ไม่บังคับ)
 * 5) Project Settings (รูปเฟือง) > Script properties > เพิ่มคีย์
 *      ENTRY_PIN = 723901   (หรือรหัสอื่นที่ต้องการ)
 *      SUPER_PIN = <รหัสผู้ดูแล ถ้าต้องการแยกสิทธิ์เพิ่มเติม>
 *    ** ห้ามใส่รหัสในซอร์สโค้ดโดยเด็ดขาด **
 * 6) Deploy > New deployment > Web app
 *      Execute as   : Me
 *      Who has access : Anyone
 *    คัดลอก URL .../exec ไปวางในตัวแปร API ที่ index.html และ input.html
 **********************************************************************/

const SHEET_ID = '1qrQwJ_4reqbACVvMod5PoAvxTMsUkfDFztLVgwb0MQU'; // Sheet ของ O

const SH_DATA          = 'EncroachmentData';
const SH_FOREST_YEARLY = 'ForestArrestHistory';
const SH_COORDINATORS  = 'Coordinators';
const SH_MASTER        = 'Meta';

// ลำดับคอลัมน์ในชีตข้อมูลหลัก (ห้ามสลับลำดับ)
const HEADERS_DATA = [
  'id', 'fiscalYear', 'district', 'type', 'caseTitle', 'actionTaken',
  'status', 'patrolTarget', 'patrolActual', 'caseBacklog', 'resolvedPercent',
  'responsibleOrg', 'updatedAt', 'updatedBy'
];

const HEADERS_FOREST = ['year', 'cases'];
const HEADERS_COORD  = ['id', 'name', 'position', 'org', 'phone'];

const NUM_FIELDS = ['fiscalYear', 'patrolTarget', 'patrolActual', 'caseBacklog', 'resolvedPercent'];

// รายชื่ออำเภอมาตรฐาน (ใช้ตรวจสอบ/แสดงในระบบกรอก)
const DISTRICTS = ['เมืองหนองบัวลำภู', 'นากลาง', 'โนนสัง', 'ศรีบุญเรือง', 'สุวรรณคูหา', 'นาวัง'];
const TYPES = ['ที่ดินสาธารณประโยชน์', 'พื้นที่ป่าไม้'];
const STATUSES = ['อยู่ระหว่างดำเนินการ', 'อยู่ระหว่างไกล่เกลี่ย', 'ฟ้องร้อง/ดำเนินคดี', 'ศาลสั่ง/รอบังคับคดี', 'เสร็จสิ้น'];

/* ================= ติดตั้งครั้งแรก ================= */
function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  ensureSheet_(ss, SH_DATA, HEADERS_DATA);
  ensureSheet_(ss, SH_FOREST_YEARLY, HEADERS_FOREST);
  ensureSheet_(ss, SH_COORDINATORS, HEADERS_COORD);
  ensureSheet_(ss, SH_MASTER, ['key', 'value']);

  return 'setup done — อย่าลืมตั้งค่า ENTRY_PIN ใน Script Properties';
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

/* ================= ข้อมูลตัวอย่าง (ไม่บังคับ) ================= */
function seedSampleData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const shD = ss.getSheetByName(SH_DATA);
  const shF = ss.getSheetByName(SH_FOREST_YEARLY);
  const shC = ss.getSheetByName(SH_COORDINATORS);
  const now = new Date().toISOString();

  // ล้างข้อมูลเดิม (คงหัวตาราง) กันข้อมูลซ้ำเมื่อรันซ้ำ
  [shD, shF, shC].forEach(function(sh){
    var last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, sh.getLastColumn()).clearContent();
  });

  const sample = [
    [1, 2569, 'เมืองหนองบัวลำภู', 'ที่ดินสาธารณประโยชน์', 'บุกรุกที่สำนักงานเร่งรัดพัฒนาชนบทของตำบลหนองภัยศูนย์', 'ตรวจสอบและรอเอกสารเพิ่มเติมจากสำนักงานเร่งรัดการพัฒนาชนบท', 'อยู่ระหว่างดำเนินการ', 0, 0, 0, 100, 'อบต.หนองภัยศูนย์'],
    [2, 2569, 'เมืองหนองบัวลำภู', 'พื้นที่ป่าไม้', 'ติดตามการบุกรุกพื้นที่ป่าไม้ (คดีค้าง 1 คดี)', 'ลาดตระเวนและสอบสวนสำนวนคดี', 'อยู่ระหว่างดำเนินการ', 2, 0, 1, 0, 'ศูนย์ป่าไม้ที่ 6 / ตำรวจภูธร'],
    [3, 2569, 'โนนสัง', 'ที่ดินสาธารณประโยชน์', 'ร้องเรียนบุกรุกที่สาธารณะ ต.บ้านค้อ ม.3 (ผู้บุกรุก 2 ราย)', 'มอบหมายเทศบาลบ้านค้อดำเนินการไกล่เกลี่ย', 'อยู่ระหว่างไกล่เกลี่ย', 0, 0, 0, 100, 'เทศบาลตำบลบ้านค้อ'],
    [4, 2569, 'โนนสัง', 'พื้นที่ป่าไม้', 'ติดตามการบุกรุกพื้นที่ป่าไม้', 'ลาดตระเวนป้องกันและปราบปรามการบุกรุกป่า', 'อยู่ระหว่างดำเนินการ', 4, 0, 0, 0, 'ศูนย์ป่าไม้ที่ 6 / ตำรวจภูธร'],
    [5, 2569, 'ศรีบุญเรือง', 'ที่ดินสาธารณประโยชน์', 'ที่สาธารณประโยชน์บ้านห้วยบง หมู่ที่ 10 ต.โนนสะอาด', 'เทศบาลตำบลโนนสะอาดยื่นคำร้องออก นสล.', 'อยู่ระหว่างดำเนินการ', 0, 0, 0, 100, 'เทศบาลตำบลโนนสะอาด'],
    [6, 2569, 'ศรีบุญเรือง', 'พื้นที่ป่าไม้', 'ติดตามการบุกรุกพื้นที่ป่าไม้ (คดีค้าง 9 คดี)', 'ลาดตระเวนและสอบสวนสำนวนคดี', 'อยู่ระหว่างดำเนินการ', 3, 0, 9, 0, 'ศูนย์ป่าไม้ที่ 6 / ตำรวจภูธร'],
    [7, 2569, 'นากลาง', 'พื้นที่ป่าไม้', 'ติดตามการบุกรุกพื้นที่ป่าไม้ (คดีค้าง 2 คดี)', 'ลาดตระเวนและสอบสวนสำนวนคดี', 'อยู่ระหว่างดำเนินการ', 4, 0, 2, 0, 'ศูนย์ป่าไม้ที่ 6 / ตำรวจภูธร'],
    [8, 2569, 'นาวัง', 'ที่ดินสาธารณประโยชน์', 'ที่สาธารณประโยชน์ ม.6 ตำบลวังปลาป้อม ราษฎรครอบครอง', 'มอบ อบต.ยื่นรังวัดสอบแนวเขต', 'อยู่ระหว่างดำเนินการ', 0, 0, 0, 100, 'อบต.วังปลาป้อม'],
    [9, 2569, 'นาวัง', 'พื้นที่ป่าไม้', 'ติดตามการบุกรุกพื้นที่ป่าไม้ (คดีค้าง 1 คดี)', 'ลาดตระเวนและสอบสวนสำนวนคดี', 'อยู่ระหว่างดำเนินการ', 3, 0, 1, 0, 'ศูนย์ป่าไม้ที่ 6 / ตำรวจภูธร'],
    [10, 2569, 'สุวรรณคูหา', 'พื้นที่ป่าไม้', 'ติดตามการบุกรุกพื้นที่ป่าไม้ (คดีค้าง 2 คดี)', 'ลาดตระเวนและสอบสวนสำนวนคดี', 'อยู่ระหว่างดำเนินการ', 4, 0, 2, 0, 'ศูนย์ป่าไม้ที่ 6 / ตำรวจภูธร']
  ].map(r => r.concat([now, 'ระบบ (ตัวอย่าง)']));

  shD.getRange(2, 1, sample.length, HEADERS_DATA.length).setValues(sample);

  shF.getRange(2, 1, 3, 2).setValues([
    [2567, 9],
    [2568, 0],
    [2569, 6]
  ]);

  shC.getRange(2, 1, 3, 5).setValues([
    [1, 'นายเดชา ภูบัวเพชร', 'ผู้อำนวยการศูนย์ป่าไม้หนองบัวลำภู', 'ศูนย์ป่าไม้หนองบัวลำภู', '081-057-7419'],
    [2, 'นายตวงวิทย์ เชื้อหอม', 'เจ้าพนักงานที่ดินจังหวัดหนองบัวลำภู', 'สำนักงานที่ดินจังหวัดหนองบัวลำภู', '081-661-2772'],
    [3, 'นายสมหมาย วงศ์อุดม', 'ผู้ช่วยจ่าจังหวัด', 'ที่ทำการปกครองจังหวัดหนองบัวลำภู', '063-903-5890']
  ]);

  return 'seed done';
}

/* ================= Router (JSONP) ================= */
function doGet(e) {
  const p = e.parameter || {};
  const cb = p.callback || 'callback';
  let out;
  try {
    const action = p.action || 'init';
    if (action === 'init' || action === 'getData') {
      out = { ok: true, rows: readAllData_(), forest: readForestHistory_(), meta: { districts: DISTRICTS, types: TYPES, statuses: STATUSES } };
    } else if (action === 'save') {
      out = saveRow_(p);
    } else if (action === 'saveBatch') {
      out = saveBatch_(p);
    } else if (action === 'deleteRow') {
      out = deleteRow_(p);
    } else if (action === 'saveForestYear') {
      out = saveForestYear_(p);
    } else if (action === 'getCoordinators') {
      out = getCoordinators_(p);
    } else if (action === 'saveCoordinator') {
      out = saveCoordinator_(p);
    } else if (action === 'ping') {
      out = { ok: true, t: new Date().toISOString() };
    } else {
      out = { ok: false, error: 'unknown action' };
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(out) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/* ================= อ่านข้อมูลหลักทั้งหมด (สาธารณะ ไม่มีข้อมูลส่วนบุคคล) ================= */
function readAllData_() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SH_DATA);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const o = {};
    head.forEach((h, j) => { o[h] = values[i][j]; });
    if (o.id === '' || o.id === null || o.id === undefined) continue;
    NUM_FIELDS.forEach(f => { o[f] = Number(o[f]) || 0; });
    o.id = Number(o.id);
    rows.push(o);
  }
  return rows;
}

function readForestHistory_() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SH_FOREST_YEARLY);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === '' || values[i][0] === null) continue;
    rows.push({ year: Number(values[i][0]), cases: Number(values[i][1]) || 0 });
  }
  return rows;
}

/* ================= บันทึก 1 แถว (upsert ด้วย id) ================= */
function saveRow_(p) {
  if (!checkPin_(p.pin)) return { ok: false, error: 'PIN ไม่ถูกต้อง' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SH_DATA);
    const lastRow = sh.getLastRow();
    const ids = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 1).getValues().flat() : [];

    let id = Number(p.id) || 0;
    let idx = id ? ids.indexOf(id) : -1;
    if (!id) {
      id = ids.length ? Math.max.apply(null, ids.map(Number)) + 1 : 1;
      idx = -1;
    }

    const rec = {
      id: id,
      fiscalYear: num_(p.fiscalYear),
      district: p.district || '',
      type: p.type || '',
      caseTitle: p.caseTitle || '',
      actionTaken: p.actionTaken || '',
      status: p.status || '',
      patrolTarget: num_(p.patrolTarget),
      patrolActual: num_(p.patrolActual),
      caseBacklog: num_(p.caseBacklog),
      resolvedPercent: num_(p.resolvedPercent),
      responsibleOrg: p.responsibleOrg || '',
      updatedAt: new Date().toISOString(),
      updatedBy: p.user || '-'
    };
    const rowArr = HEADERS_DATA.map(h => rec[h]);

    if (idx >= 0) {
      sh.getRange(idx + 2, 1, 1, HEADERS_DATA.length).setValues([rowArr]);
      return { ok: true, mode: 'update', id: id };
    } else {
      sh.appendRow(rowArr);
      return { ok: true, mode: 'insert', id: id };
    }
  } finally {
    lock.releaseLock();
  }
}

/* ================= บันทึกหลายแถวพร้อมกัน (นำเข้าจาก Excel) ================= */
function saveBatch_(p) {
  if (!checkPin_(p.pin)) return { ok: false, error: 'PIN ไม่ถูกต้อง' };
  let list = [];
  try { list = JSON.parse(p.rows || '[]'); } catch (e) { return { ok: false, error: 'รูปแบบข้อมูลไม่ถูกต้อง' }; }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SH_DATA);
    const lastRow = sh.getLastRow();
    let ids = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(Number) : [];
    let nextId = ids.length ? Math.max.apply(null, ids) + 1 : 1;

    let inserted = 0, updated = 0;
    list.forEach(function (p2) {
      const now = new Date().toISOString();
      let id = Number(p2.id) || 0;
      let idx = id ? ids.indexOf(id) : -1;
      if (!id) { id = nextId++; idx = -1; }

      const rec = {
        id: id,
        fiscalYear: num_(p2.fiscalYear),
        district: p2.district || '',
        type: p2.type || '',
        caseTitle: p2.caseTitle || '',
        actionTaken: p2.actionTaken || '',
        status: p2.status || '',
        patrolTarget: num_(p2.patrolTarget),
        patrolActual: num_(p2.patrolActual),
        caseBacklog: num_(p2.caseBacklog),
        resolvedPercent: num_(p2.resolvedPercent),
        responsibleOrg: p2.responsibleOrg || '',
        updatedAt: now,
        updatedBy: p2.user || p.user || 'นำเข้าจาก Excel'
      };
      const rowArr = HEADERS_DATA.map(h => rec[h]);

      if (idx >= 0) {
        sh.getRange(idx + 2, 1, 1, HEADERS_DATA.length).setValues([rowArr]);
        updated++;
      } else {
        sh.appendRow(rowArr);
        ids.push(id);
        inserted++;
      }
    });
    return { ok: true, inserted: inserted, updated: updated };
  } finally {
    lock.releaseLock();
  }
}

/* ================= ลบแถว ================= */
function deleteRow_(p) {
  if (!checkPin_(p.pin)) return { ok: false, error: 'PIN ไม่ถูกต้อง' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SH_DATA);
    const lastRow = sh.getLastRow();
    const ids = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 1).getValues().flat() : [];
    const idx = ids.indexOf(Number(p.id));
    if (idx < 0) return { ok: false, error: 'ไม่พบรายการ' };
    sh.deleteRow(idx + 2);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ================= สถิติจับกุมคดีป่าไม้รายปี ================= */
function saveForestYear_(p) {
  if (!checkPin_(p.pin)) return { ok: false, error: 'PIN ไม่ถูกต้อง' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SH_FOREST_YEARLY);
    const lastRow = sh.getLastRow();
    const years = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(Number) : [];
    const year = Number(p.year);
    const idx = years.indexOf(year);
    if (idx >= 0) {
      sh.getRange(idx + 2, 2, 1, 1).setValue(num_(p.cases));
    } else {
      sh.appendRow([year, num_(p.cases)]);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ================= ผู้ประสานงาน (ข้อมูลอ่อนไหว PDPA — ต้องมี PIN เท่านั้น) ================= */
function getCoordinators_(p) {
  if (!checkPin_(p.pin)) return { ok: false, error: 'PIN ไม่ถูกต้อง' };
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SH_COORDINATORS);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { ok: true, rows: [] };
  const head = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === '' || values[i][0] === null) continue;
    const o = {};
    head.forEach((h, j) => { o[h] = values[i][j]; });
    rows.push(o);
  }
  return { ok: true, rows: rows };
}

function saveCoordinator_(p) {
  if (!checkPin_(p.pin)) return { ok: false, error: 'PIN ไม่ถูกต้อง' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SH_COORDINATORS);
    const lastRow = sh.getLastRow();
    const ids = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 1).getValues().flat() : [];
    let id = Number(p.id) || 0;
    let idx = id ? ids.indexOf(id) : -1;
    if (!id) { id = ids.length ? Math.max.apply(null, ids.map(Number)) + 1 : 1; idx = -1; }
    const rowArr = [id, p.name || '', p.position || '', p.org || '', p.phone || ''];
    if (idx >= 0) sh.getRange(idx + 2, 1, 1, 5).setValues([rowArr]);
    else sh.appendRow(rowArr);
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

/* ================= ตรวจสิทธิ์ PIN (อ่านจาก Script Properties เท่านั้น ไม่มีรหัสในซอร์ส) ================= */
function checkPin_(pin) {
  pin = String(pin || '');
  const props = PropertiesService.getScriptProperties();
  const entryPin = props.getProperty('ENTRY_PIN');
  const superPin = props.getProperty('SUPER_PIN');
  if (entryPin && pin === String(entryPin)) return true;
  if (superPin && pin === String(superPin)) return true;
  return false;
}

/* ================= helpers ================= */
function num_(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
