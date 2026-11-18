/**
 * ระบบจัดการแข่งขันกีฬา "วังยาวเขวาไร่ดอนกลางเกมส์ 2569"
 * Google Apps Script Backend + Google Sheets Database
 *
 * วิธีเริ่มต้นใช้งานใน Google Sheets:
 * 1. เปิด Google Sheets -> ส่วนขยาย (Extensions) -> Apps Script
 * 2. วางโค้ดไฟล์นี้ใน Code.gs และวางโค้ดไฟล์ index.html ในไฟล์ HTML ชื่อ Index
 * 3. รีเฟรชหน้า Google Sheets -> คลิกเมนู "ระบบกีฬา 2569" -> "1. สร้างฐานข้อมูลและเติมข้อมูลจากสูจิบัตร"
 * 4. ตั้งรหัส PIN กรรมการผ่านเมนู "ระบบกีฬา 2569" -> "2. ตั้งรหัส PIN กรรมการ"
 * 5. กด Deploy -> New deployment -> Web app -> Execute as: Me -> Who has access: Anyone
 */

const CACHE_KEY_STATE = 'WKD_SPORTS_STATE_V1';
const CACHE_TTL_SECONDS = 120;
const PROP_ADMIN_PIN_HASH = 'WKD_ADMIN_PIN_HASH';
const PROP_ADMIN_SALT = 'WKD_ADMIN_PIN_SALT';

const SHEET_NAMES = {
  SCHOOLS: 'Schools',
  EVENTS: 'Events',
  MATCHES: 'Matches',
  RANKINGS: 'Rankings',
  CERTIFICATES: 'Certificates'
};

/**
 * สร้างเมนูจัดการบนแถบเครื่องมือ Google Sheets อัตโนมัติเมื่อเปิดชีต
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ระบบกีฬา 2569')
    .addItem('1. สร้างฐานข้อมูลและเติมข้อมูลจากสูจิบัตร', 'setupSheets')
    .addItem('2. ตั้งรหัส PIN กรรมการ/แอดมิน', 'promptSetAdminPin')
    .addItem('3. ล้างแคชข้อมูลเว็บแอป', 'clearAppCache')
    .addToUi();
}

/**
 * จุดเข้าหลักของ Web App (รองรับทั้ง HTML UI และ JSON API สำหรับ Vercel)
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'getInitialData') {
    return ContentService
      .createTextOutput(JSON.stringify(getInitialData()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const template = HtmlService.createTemplateFromFile('Index');
  template.certLookup = (e && e.parameter && e.parameter.cert) ? String(e.parameter.cert) : '';
  return template
    .evaluate()
    .setTitle('วังยาวเขวาไร่ดอนกลางเกมส์ 2569 — ระบบรายงานผลและเกียรติบัตร')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * จุดรับคำขอ POST JSON จากภายนอก (Vercel) สำหรับบันทึกผล/สกอร์ (ตรวจ PIN ฝั่ง Server เสมอ)
 */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) ? e.postData.contents : '{}');
    let res = { ok: false, message: 'Invalid action' };
    if (body.action === 'saveRankingRecord') {
      res = saveRankingRecord(body.payload || {});
    } else if (body.action === 'saveMatchRecord') {
      res = saveMatchRecord(body.payload || {});
    }
    return ContentService
      .createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, message: 'คำขอไม่ถูกต้อง' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * สร้างแผ่นงานทั้ง 5 ชีตและเติมข้อมูลตั้งต้นจากสูจิบัตรปีการศึกษา 2569
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    ensureSheet_(ss, SHEET_NAMES.SCHOOLS, ['school_id', 'school_name', 'short_name', 'group_code', 'type', 'is_host']);
    ensureSheet_(ss, SHEET_NAMES.EVENTS, ['event_id', 'category', 'sport_name', 'sport_en', 'level', 'gender', 'players_per_team', 'status']);
    ensureSheet_(ss, SHEET_NAMES.MATCHES, ['match_id', 'event_id', 'round_name', 'team_a', 'score_a', 'score_b', 'team_b', 'match_status', 'updated_at']);
    ensureSheet_(ss, SHEET_NAMES.RANKINGS, ['event_id', 'rank1_school', 'rank1_athlete', 'rank2_school', 'rank2_athlete', 'rank3a_school', 'rank3a_athlete', 'rank3b_school', 'rank3b_athlete', 'referee_name', 'note', 'updated_at']);
    ensureSheet_(ss, SHEET_NAMES.CERTIFICATES, ['cert_no', 'event_id', 'rank_key', 'rank_label', 'recipient_name', 'school_name', 'sport_name_th', 'sport_name_en', 'issue_date_th', 'created_at']);

    seedSchoolsIfEmpty_(ss.getSheetByName(SHEET_NAMES.SCHOOLS));
    seedEventsIfEmpty_(ss.getSheetByName(SHEET_NAMES.EVENTS));
    seedSampleMatchesIfEmpty_(ss.getSheetByName(SHEET_NAMES.MATCHES));
    seedSampleRankingsIfEmpty_(ss.getSheetByName(SHEET_NAMES.RANKINGS));

    const hasPin = Boolean(PropertiesService.getScriptProperties().getProperty(PROP_ADMIN_PIN_HASH));
    clearAppCache();

    try {
      const pinNote = hasPin
        ? 'ระบบพร้อมใช้งานแล้วค่ะ'
        : 'ขั้นตอนสำคัญ: กรุณากดเมนู "ระบบกีฬา 2569" -> "2. ตั้งรหัส PIN กรรมการ/แอดมิน" ก่อนเริ่มบันทึกคะแนน';
      SpreadsheetApp.getUi().alert('ตั้งค่าฐานข้อมูลทั้ง 5 ชีตจากสูจิบัตรปี 2569 เรียบร้อยแล้ว\n\n' + pinNote);
    } catch (_) {
      // กรณีรันจาก Script Editor โดยไม่ได้เปิดหน้าต่าง Sheets UI
    }
    return { ok: true, message: 'setupSheets completed' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ให้ผู้ดูแลตั้งค่า PIN กรรมการจากเมนูใน Google Sheets
 */
function promptSetAdminPin() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'ตั้งรหัส PIN สำหรับกรรมการ/แอดมิน',
    'กรอกรหัส PIN 4-8 หลักสำหรับใช้ยืนยันตัวตนก่อนบันทึกคะแนนในหน้าเว็บ:',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const pin = res.getResponseText().trim();
  if (pin.length < 4) {
    ui.alert('รหัส PIN ต้องมีอย่างน้อย 4 ตัวอักษรค่ะ');
    return;
  }
  setAdminPin(pin);
  ui.alert('บันทึกรหัส PIN ใหม่แบบเข้ารหัส (SHA-256) เรียบร้อยแล้ว');
}

/**
 * เก็บรหัส PIN แบบ Salted SHA-256 ใน PropertiesService (ไม่เก็บ plaintext)
 */
function setAdminPin(rawPin) {
  const pin = String(rawPin || '').trim();
  if (pin.length < 4) {
    throw new Error('PIN must be at least 4 characters');
  }
  const props = PropertiesService.getScriptProperties();
  const salt = Utilities.getUuid();
  const hash = hashPin_(pin, salt);
  props.setProperty(PROP_ADMIN_SALT, salt);
  props.setProperty(PROP_ADMIN_PIN_HASH, hash);
}

/**
 * ตรวจสอบ PIN ฝั่ง Server
 */
function verifyAdminPin(pinInput) {
  const pin = String(pinInput || '').trim();
  if (!pin) return { ok: false, message: 'กรุณากรอกรหัส PIN กรรมการ' };

  const props = PropertiesService.getScriptProperties();
  const salt = props.getProperty(PROP_ADMIN_SALT);
  const storedHash = props.getProperty(PROP_ADMIN_PIN_HASH);

  if (!salt || !storedHash) {
    return { ok: false, message: 'ยังไม่ได้ตั้งค่ารหัส PIN ใน Google Sheets กรุณาตั้งค่าผ่านเมนู ระบบกีฬา 2569' };
  }

  const candidateHash = hashPin_(pin, salt);
  if (candidateHash === storedHash) {
    return { ok: true, message: 'ยืนยันสิทธิ์กรรมการสำเร็จ' };
  }
  return { ok: false, message: 'รหัส PIN ไม่ถูกต้อง' };
}

/**
 * ดึงข้อมูลทั้งหมดสำหรับหน้าเว็บแบบอ่านรอบเดียว (Batch Read + Cache)
 */
function getInitialData() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY_STATE);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (_) {
      // Cache corrupt -> fallback to sheet read
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schoolsSheet = ss.getSheetByName(SHEET_NAMES.SCHOOLS);
  if (!schoolsSheet) {
    return { ok: false, needsSetup: true, message: 'ยังไม่ได้สร้างตารางฐานข้อมูล กรุณารัน setupSheets() ใน Google Sheets' };
  }

  const schools = readRowsAsObjects_(ss.getSheetByName(SHEET_NAMES.SCHOOLS));
  const events = readRowsAsObjects_(ss.getSheetByName(SHEET_NAMES.EVENTS));
  const matches = readRowsAsObjects_(ss.getSheetByName(SHEET_NAMES.MATCHES));
  const rankings = readRowsAsObjects_(ss.getSheetByName(SHEET_NAMES.RANKINGS));
  const certificates = readRowsAsObjects_(ss.getSheetByName(SHEET_NAMES.CERTIFICATES));

  const payload = {
    ok: true,
    syncedAt: new Date().toISOString(),
    schools: schools,
    events: events,
    matches: matches,
    rankings: rankings,
    certificates: certificates
  };

  try {
    const serialized = JSON.stringify(payload);
    if (serialized.length < 95000) {
      cache.put(CACHE_KEY_STATE, serialized, CACHE_TTL_SECONDS);
    }
  } catch (_) {
    // Ignore cache size limit if exceeded
  }

  return payload;
}

/**
 * บันทึกผลการตัดสินลำดับที่ (1, 2, 3 ร่วม) และสร้างทะเบียนเลขที่เกียรติบัตรอัตโนมัติ
 */
function saveRankingRecord(requestPayload) {
  const payload = requestPayload || {};
  const auth = verifyAdminPin(payload.pin);
  if (!auth.ok) {
    return { ok: false, message: auth.message };
  }

  const eventId = String(payload.eventId || '').trim();
  const rank1School = String(payload.rank1School || '').trim();
  if (!eventId || !rank1School) {
    return { ok: false, message: 'กรุณาเลือกรายการแข่งขันและโรงเรียนที่ชนะเลิศ' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rankingsSheet = ss.getSheetByName(SHEET_NAMES.RANKINGS);
    const eventsSheet = ss.getSheetByName(SHEET_NAMES.EVENTS);
    const certsSheet = ss.getSheetByName(SHEET_NAMES.CERTIFICATES);

    const nowIso = new Date().toISOString();
    const rowData = [
      sanitizeSheetValue_(eventId),
      sanitizeSheetValue_(rank1School),
      sanitizeSheetValue_(payload.rank1Athlete || ('ทีม' + rank1School)),
      sanitizeSheetValue_(payload.rank2School || ''),
      sanitizeSheetValue_(payload.rank2Athlete || ''),
      sanitizeSheetValue_(payload.rank3aSchool || ''),
      sanitizeSheetValue_(payload.rank3aAthlete || ''),
      sanitizeSheetValue_(payload.rank3bSchool || ''),
      sanitizeSheetValue_(payload.rank3bAthlete || ''),
      sanitizeSheetValue_(payload.refereeName || ''),
      sanitizeSheetValue_(payload.note || ''),
      nowIso
    ];

    upsertByKey_(rankingsSheet, 0, eventId, rowData);
    updateEventStatus_(eventsSheet, eventId, 'ตัดสินแล้ว');
    const createdCerts = syncEventCertificates_(certsSheet, eventsSheet, eventId, payload, nowIso);

    clearAppCache();
    return {
      ok: true,
      message: 'บันทึกลำดับคะแนนและอัปเดตทะเบียนเกียรติบัตรใน Google Sheets เรียบร้อยแล้ว',
      certificates: createdCerts,
      data: getInitialData()
    };
  } catch (err) {
    console.error('saveRankingRecord failed:', err);
    return { ok: false, message: 'ไม่สามารถบันทึกข้อมูลได้ กรุณาลองใหม่อีกครั้งหรือติดต่อผู้ดูแลระบบ' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * บันทึกผลสกอร์รายคู่ (Matches)
 */
function saveMatchRecord(requestPayload) {
  const payload = requestPayload || {};
  const auth = verifyAdminPin(payload.pin);
  if (!auth.ok) {
    return { ok: false, message: auth.message };
  }

  const eventId = String(payload.eventId || '').trim();
  const teamA = String(payload.teamA || '').trim();
  const teamB = String(payload.teamB || '').trim();
  if (!eventId || !teamA || !teamB) {
    return { ok: false, message: 'กรุณาระบุรายการกีฬาและชื่อทีมทั้งสองฝั่ง' };
  }

  const scoreA = Math.max(0, parseInt(payload.scoreA, 10) || 0);
  const scoreB = Math.max(0, parseInt(payload.scoreB, 10) || 0);
  const matchId = String(payload.matchId || ('M_' + eventId + '_' + Date.now())).trim();

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const matchesSheet = ss.getSheetByName(SHEET_NAMES.MATCHES);
    const nowIso = new Date().toISOString();

    const rowData = [
      sanitizeSheetValue_(matchId),
      sanitizeSheetValue_(eventId),
      sanitizeSheetValue_(payload.roundName || 'รอบชิงชนะเลิศ'),
      sanitizeSheetValue_(teamA),
      scoreA,
      scoreB,
      sanitizeSheetValue_(teamB),
      sanitizeSheetValue_(payload.matchStatus || 'จบการแข่งขัน'),
      nowIso
    ];

    upsertByKey_(matchesSheet, 0, matchId, rowData);
    clearAppCache();

    return {
      ok: true,
      message: 'บันทึกผลการแข่งขันรายคู่ลง Google Sheets เรียบร้อยแล้ว',
      data: getInitialData()
    };
  } catch (err) {
    console.error('saveMatchRecord failed:', err);
    return { ok: false, message: 'ไม่สามารถบันทึกผลสกอร์ได้ กรุณาลองใหม่อีกครั้ง' };
  } finally {
    lock.releaseLock();
  }
}

function clearAppCache() {
  CacheService.getScriptCache().remove(CACHE_KEY_STATE);
}

/* ==================== Internal Helpers ==================== */

function hashPin_(pin, salt) {
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + ':' + pin,
    Utilities.Charset.UTF_8
  );
  return raw.map(function (b) {
    const v = (b < 0) ? b + 256 : b;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

function sanitizeSheetValue_(val) {
  const str = String(val == null ? '' : val);
  if (/^[=+\-@\t\r]/.test(str)) {
    return "'" + str;
  }
  return str;
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#2E1065')
      .setFontColor('#FDE68A');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readRowsAsObjects_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.join('').trim() === '') continue;
    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      let cell = row[c];
      if (cell instanceof Date) cell = cell.toISOString();
      obj[headers[c]] = cell;
    }
    rows.push(obj);
  }
  return rows;
}

function upsertByKey_(sheet, keyColIndex, keyVal, newRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const keys = sheet.getRange(2, keyColIndex + 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim() === keyVal) {
        sheet.getRange(i + 2, 1, 1, newRow.length).setValues([newRow]);
        return;
      }
    }
  }
  sheet.appendRow(newRow);
}

function updateEventStatus_(eventsSheet, eventId, statusText) {
  const lastRow = eventsSheet.getLastRow();
  if (lastRow < 2) return;
  const values = eventsSheet.getDataRange().getValues();
  const headers = values[0];
  const statusCol = headers.indexOf('status');
  if (statusCol === -1) return;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === eventId) {
      eventsSheet.getRange(i + 1, statusCol + 1).setValue(statusText);
      break;
    }
  }
}

function syncEventCertificates_(certsSheet, eventsSheet, eventId, payload, nowIso) {
  const events = readRowsAsObjects_(eventsSheet);
  const ev = events.find(function (item) { return String(item.event_id) === eventId; }) || {
    sport_name: eventId,
    sport_en: 'SPORTS',
    level: '',
    gender: ''
  };

  const slots = [
    { key: 'rank1', code: '01', label: 'ได้รับรางวัลชนะเลิศ', school: payload.rank1School, athlete: payload.rank1Athlete },
    { key: 'rank2', code: '02', label: 'ได้รับรางวัลรองชนะเลิศ อันดับที่ 1', school: payload.rank2School, athlete: payload.rank2Athlete },
    { key: 'rank3a', code: '03', label: 'ได้รับรางวัลรองชนะเลิศ อันดับที่ 2', school: payload.rank3aSchool, athlete: payload.rank3aAthlete },
    { key: 'rank3b', code: '04', label: 'ได้รับรางวัลรองชนะเลิศ อันดับที่ 2', school: payload.rank3bSchool, athlete: payload.rank3bAthlete }
  ];

  const generated = [];
  slots.forEach(function (slot) {
    if (!slot.school) return;
    const cleanCode = eventId.replace(/[^A-Z0-9]/gi, '').slice(0, 8).toUpperCase();
    const certNo = 'WYK-2569-' + cleanCode + '-' + slot.code;
    const recipient = (slot.athlete && String(slot.athlete).trim()) ? slot.athlete : ('ทีม' + slot.school);
    const fullSportTh = ev.sport_name + (ev.level ? (' ระดับชั้น' + ev.level) : '') + (ev.gender ? (' (' + ev.gender + ')') : '');

    const row = [
      certNo,
      eventId,
      slot.key,
      slot.label,
      sanitizeSheetValue_(recipient),
      sanitizeSheetValue_(slot.school),
      sanitizeSheetValue_(fullSportTh),
      sanitizeSheetValue_(ev.sport_en || 'FOOTBALL'),
      '4 ธันวาคม 2569',
      nowIso
    ];
    upsertByKey_(certsSheet, 0, certNo, row);
    generated.push({ cert_no: certNo, rank_key: slot.key, recipient_name: recipient, school_name: slot.school });
  });
  return generated;
}

function seedSchoolsIfEmpty_(sheet) {
  if (sheet.getLastRow() > 1) return;
  const rows = [
    ['SCH01', 'โรงเรียนบ้านวังกุง', 'ร.ร.บ้านวังกุง', 'สาย A', 'โรงเรียน', 'เจ้าภาพสาย A'],
    ['SCH02', 'โรงเรียนบ้านวังจาน', 'ร.ร.บ้านวังจาน', 'สาย A', 'โรงเรียน', ''],
    ['SCH03', 'โรงเรียนบ้านบะหลวงหนองแวง', 'ร.ร.บ้านบะหลวงหนองแวง', 'สาย A', 'โรงเรียน', ''],
    ['SCH04', 'โรงเรียนบ้านทิพโสต', 'ร.ร.บ้านทิพโสต', 'สาย A', 'โรงเรียน', ''],
    ['SCH05', 'โรงเรียนบ้านหนองสระพังโนนสะอาด', 'ร.ร.บ้านหนองสระพังฯ', 'สาย B', 'โรงเรียน', 'เจ้าภาพสาย B'],
    ['SCH06', 'โรงเรียนหินแห่เสริมศิลป์', 'ร.ร.หินแห่เสริมศิลป์', 'สาย B', 'โรงเรียน', 'สนามกรีฑา'],
    ['SCH07', 'โรงเรียนบ้านโคกกลาง', 'ร.ร.บ้านโคกกลาง', 'สาย B', 'โรงเรียน', ''],
    ['SCH08', 'โรงเรียนบ้านวังยาววิทยายน', 'ร.ร.บ้านวังยาววิทยายน', 'สาย C', 'โรงเรียน', 'สนามหลัก / เจ้าภาพสาย C'],
    ['SCH09', 'โรงเรียนบ้านวังโพน', 'ร.ร.บ้านวังโพน', 'สาย C', 'โรงเรียน', ''],
    ['SCH10', 'โรงเรียนบ้านเหล่าพ่อหา', 'ร.ร.บ้านเหล่าพ่อหา', 'สาย C', 'โรงเรียน', ''],
    ['SCH11', 'โรงเรียนบ้านโนนราษีโคกล่าม', 'ร.ร.บ้านโนนราษีโคกล่าม', 'สาย D', 'โรงเรียน', 'เจ้าภาพสาย D'],
    ['SCH12', 'โรงเรียนบ้านหนองเขื่อน', 'ร.ร.บ้านหนองเขื่อน', 'สาย D', 'โรงเรียน', ''],
    ['SCH13', 'โรงเรียนดอนกลางนุกูลวิทย์', 'ร.ร.ดอนกลางนุกูลวิทย์', 'สาย D', 'โรงเรียน', ''],
    ['KDC01', 'ศพด.อบต.เขวาไร่', 'ศพด.อบต.เขวาไร่', 'อนุบาล', 'ศพด.', ''],
    ['KDC02', 'ศพด.อบต.วังยาว', 'ศพด.อบต.วังยาว', 'อนุบาล', 'ศพด.', ''],
    ['KDC03', 'ศพด.วัดโพธิ์กลาง', 'ศพด.วัดโพธิ์กลาง', 'อนุบาล', 'ศพด.', ''],
    ['KDC04', 'ศพด.วัดสระทอง', 'ศพด.วัดสระทอง', 'อนุบาล', 'ศพด.', ''],
    ['KDC05', 'ศพด.วัดวังจาน', 'ศพด.วัดวังจาน', 'อนุบาล', 'ศพด.', ''],
    ['KDC06', 'ศพด.บ้านหนองสิม', 'ศพด.บ้านหนองสิม', 'อนุบาล', 'ศพด.', ''],
    ['KDC07', 'ศพด.วัดบ้านวังกุง', 'ศพด.วัดบ้านวังกุง', 'อนุบาล', 'ศพด.', ''],
    ['KDC08', 'ศพด.อบต.ดอนกลาง', 'ศพด.อบต.ดอนกลาง', 'อนุบาล', 'ศพด.', '']
  ];
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedEventsIfEmpty_(sheet) {
  if (sheet.getLastRow() > 1) return;
  const rows = [
    // กีฬาฟุตบอล & ฟุตซอล
    ['FB_K_MIX', 'football', 'ฟุตบอลอนุบาล 7 คน', 'KIDS FOOTBALL', 'อนุบาล (อ.1-อ.3)', 'ชาย-หญิง', 7, 'ตัดสินแล้ว'],
    ['FB_7_P13_M', 'football', 'ฟุตบอล 7 คน', 'FOOTBALL 7S', 'ป.1-3', 'ชาย', 12, 'รอแข่งขัน'],
    ['FB_7_P46_M', 'football', 'ฟุตบอล 7 คน', 'FOOTBALL 7S', 'ป.4-6', 'ชาย', 12, 'ตัดสินแล้ว'],
    ['FB_11_P46_M', 'football', 'ฟุตบอล 11 คน', 'FOOTBALL', 'ป.4-6', 'ชาย', 18, 'กำลังแข่งขัน'],
    ['FB_11_M13_M', 'football', 'ฟุตบอล 11 คน', 'FOOTBALL', 'ม.1-3', 'ชาย', 18, 'กำลังแข่งขัน'],
    ['FS_P46_M', 'football', 'ฟุตซอล', 'FUTSAL', 'ป.4-6', 'ชาย', 12, 'ตัดสินแล้ว'],
    ['FS_M13_M', 'football', 'ฟุตซอล', 'FUTSAL', 'ม.1-3', 'ชาย', 12, 'รอแข่งขัน'],
    // วอลเลย์บอล / เซปักตะกร้อ / แชร์บอล
    ['VB_P46_M', 'court', 'วอลเลย์บอล', 'VOLLEYBALL', 'ป.4-6', 'ชาย', 12, 'รอแข่งขัน'],
    ['VB_P46_F', 'court', 'วอลเลย์บอล', 'VOLLEYBALL', 'ป.4-6', 'หญิง', 12, 'ตัดสินแล้ว'],
    ['VB_M13_M', 'court', 'วอลเลย์บอล', 'VOLLEYBALL', 'ม.1-3', 'ชาย', 12, 'รอแข่งขัน'],
    ['VB_M13_F', 'court', 'วอลเลย์บอล', 'VOLLEYBALL', 'ม.1-3', 'หญิง', 12, 'รอแข่งขัน'],
    ['TK_P46_M', 'court', 'เซปักตะกร้อ', 'SEPAK TAKRAW', 'ป.4-6', 'ชาย', 4, 'รอแข่งขัน'],
    ['TK_P46_F', 'court', 'เซปักตะกร้อ', 'SEPAK TAKRAW', 'ป.4-6', 'หญิง', 4, 'รอแข่งขัน'],
    ['TK_M13_M', 'court', 'เซปักตะกร้อ', 'SEPAK TAKRAW', 'ม.1-3', 'ชาย', 4, 'รอแข่งขัน'],
    ['TK_M13_F', 'court', 'เซปักตะกร้อ', 'SEPAK TAKRAW', 'ม.1-3', 'หญิง', 4, 'รอแข่งขัน'],
    ['CB_P13_F', 'court', 'แชร์บอล', 'CHAIRBALL', 'ป.1-3', 'หญิง', 12, 'รอแข่งขัน'],
    // เปตอง
    ['PT_P13_M', 'petanque', 'เปตอง', 'PETANQUE', 'ป.1-3', 'ชาย', 4, 'รอแข่งขัน'],
    ['PT_P13_F', 'petanque', 'เปตอง', 'PETANQUE', 'ป.1-3', 'หญิง', 4, 'รอแข่งขัน'],
    ['PT_P46_M', 'petanque', 'เปตอง', 'PETANQUE', 'ป.4-6', 'ชาย', 4, 'รอแข่งขัน'],
    ['PT_P46_F', 'petanque', 'เปตอง', 'PETANQUE', 'ป.4-6', 'หญิง', 4, 'รอแข่งขัน'],
    ['PT_M13_M', 'petanque', 'เปตอง', 'PETANQUE', 'ม.1-3', 'ชาย', 4, 'รอแข่งขัน'],
    ['PT_M13_F', 'petanque', 'เปตอง', 'PETANQUE', 'ม.1-3', 'หญิง', 4, 'รอแข่งขัน'],
    // กรีฑา (ลู่ / ผลัด)
    ['TR_30_K34_M', 'track', 'วิ่งเร็ว 30 เมตร', 'SPRINT 30M', 'อบ.3-4 ปี', 'ชาย', 1, 'รอแข่งขัน'],
    ['TR_30_K34_F', 'track', 'วิ่งเร็ว 30 เมตร', 'SPRINT 30M', 'อบ.3-4 ปี', 'หญิง', 1, 'รอแข่งขัน'],
    ['TR_30_K56_M', 'track', 'วิ่งเร็ว 30 เมตร', 'SPRINT 30M', 'อบ.5-6 ปี', 'ชาย', 1, 'รอแข่งขัน'],
    ['TR_30_K56_F', 'track', 'วิ่งเร็ว 30 เมตร', 'SPRINT 30M', 'อบ.5-6 ปี', 'หญิง', 1, 'รอแข่งขัน'],
    ['TR_60_P13_M', 'track', 'วิ่งเร็ว 60 เมตร', 'SPRINT 60M', 'ป.1-3', 'ชาย', 1, 'รอแข่งขัน'],
    ['TR_60_P13_F', 'track', 'วิ่งเร็ว 60 เมตร', 'SPRINT 60M', 'ป.1-3', 'หญิง', 1, 'รอแข่งขัน'],
    ['TR_100_P56_M', 'track', 'วิ่งเร็ว 100 เมตร', 'SPRINT 100M', 'ป.5-6', 'ชาย', 1, 'ตัดสินแล้ว'],
    ['TR_100_P56_F', 'track', 'วิ่งเร็ว 100 เมตร', 'SPRINT 100M', 'ป.5-6', 'หญิง', 1, 'รอแข่งขัน'],
    ['TR_100_M13_M', 'track', 'วิ่งเร็ว 100 เมตร', 'SPRINT 100M', 'ม.1-3', 'ชาย', 1, 'รอแข่งขัน'],
    ['TR_100_M13_F', 'track', 'วิ่งเร็ว 100 เมตร', 'SPRINT 100M', 'ม.1-3', 'หญิง', 1, 'รอแข่งขัน'],
    ['RL_8X50_P13_M', 'track', 'วิ่งผลัด 8x50 เมตร', 'RELAY 8X50M', 'ป.1-3', 'ชาย', 9, 'รอแข่งขัน'],
    ['RL_8X50_P13_F', 'track', 'วิ่งผลัด 8x50 เมตร', 'RELAY 8X50M', 'ป.1-3', 'หญิง', 9, 'รอแข่งขัน'],
    ['RL_4X100_P56_M', 'track', 'วิ่งผลัด 4x100 เมตร', 'RELAY 4X100M', 'ป.5-6', 'ชาย', 5, 'รอแข่งขัน'],
    ['RL_4X100_P56_F', 'track', 'วิ่งผลัด 4x100 เมตร', 'RELAY 4X100M', 'ป.5-6', 'หญิง', 5, 'รอแข่งขัน'],
    ['RL_4X100_M13_M', 'track', 'วิ่งผลัด 4x100 เมตร', 'RELAY 4X100M', 'ม.1-3', 'ชาย', 5, 'รอแข่งขัน'],
    ['RL_4X100_M13_F', 'track', 'วิ่งผลัด 4x100 เมตร', 'RELAY 4X100M', 'ม.1-3', 'หญิง', 5, 'รอแข่งขัน'],
    // กีฬาอนุบาลพื้นบ้าน
    ['KG_ZP_K34', 'kindergarten', 'วิ่งเปี้ยว (ช2 ญ2)', 'ZIGZAG RELAY', 'อบ.3-4 ปี', 'ผสม ช2 ญ2', 4, 'รอแข่งขัน'],
    ['KG_ZP_K56', 'kindergarten', 'วิ่งเปี้ยว (ช2 ญ2)', 'ZIGZAG RELAY', 'อบ.5-6 ปี', 'ผสม ช2 ญ2', 4, 'รอแข่งขัน'],
    ['KG_TOW_K34', 'kindergarten', 'ชักกะเย่อ (ช4 ญ4)', 'TUG OF WAR', 'อบ.3-4 ปี', 'ผสม ช4 ญ4', 8, 'รอแข่งขัน'],
    ['KG_TOW_K56', 'kindergarten', 'ชักกะเย่อ (ช4 ญ4)', 'TUG OF WAR', 'อบ.5-6 ปี', 'ผสม ช4 ญ4', 8, 'รอแข่งขัน'],
    ['KG_BB_K34_M', 'kindergarten', 'โยนบอลลงตะกร้า', 'BASKET TOSS', 'อบ.3-4 ปี', 'ชาย', 6, 'รอแข่งขัน'],
    ['KG_BB_K34_F', 'kindergarten', 'โยนบอลลงตะกร้า', 'BASKET TOSS', 'อบ.3-4 ปี', 'หญิง', 6, 'รอแข่งขัน'],
    ['KG_BB_K56_M', 'kindergarten', 'โยนบอลลงตะกร้า', 'BASKET TOSS', 'อบ.5-6 ปี', 'ชาย', 6, 'รอแข่งขัน'],
    ['KG_BB_K56_F', 'kindergarten', 'โยนบอลลงตะกร้า', 'BASKET TOSS', 'อบ.5-6 ปี', 'หญิง', 6, 'รอแข่งขัน']
  ];
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedSampleMatchesIfEmpty_(sheet) {
  if (sheet.getLastRow() > 1) return;
  const nowIso = new Date().toISOString();
  const rows = [
    ['M_FB7_FINAL', 'FB_7_P46_M', 'รอบชิงชนะเลิศ', 'โรงเรียนบ้านวังยาววิทยายน', 3, 1, 'โรงเรียนบ้านวังกุง', 'จบการแข่งขัน', nowIso],
    ['M_FB11_R1', 'FB_11_M13_M', 'พบกันหมด คู่ที่ 1', 'โรงเรียนบ้านวังยาววิทยายน', 2, 2, 'โรงเรียนดอนกลางนุกูลวิทย์', 'กำลังแข่งขัน', nowIso],
    ['M_FS_FINAL', 'FS_P46_M', 'รอบชิงชนะเลิศ', 'โรงเรียนบ้านหนองสระพังโนนสะอาด', 4, 2, 'โรงเรียนดอนกลางนุกูลวิทย์', 'จบการแข่งขัน', nowIso],
    ['M_VB_SF1', 'VB_P46_F', 'รอบรองชนะเลิศ', 'โรงเรียนหินแห่เสริมศิลป์', 2, 0, 'โรงเรียนบ้านโนนราษีโคกล่าม', 'จบการแข่งขัน (เซต)', nowIso]
  ];
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedSampleRankingsIfEmpty_(sheet) {
  if (sheet.getLastRow() > 1) return;
  const nowIso = new Date().toISOString();
  const rows = [
    ['FB_7_P46_M', 'โรงเรียนบ้านวังยาววิทยายน', 'ทีมฟุตบอลโรงเรียนบ้านวังยาววิทยายน', 'โรงเรียนบ้านวังกุง', 'ทีมฟุตบอลโรงเรียนบ้านวังกุง', 'โรงเรียนบ้านหนองสระพังโนนสะอาด', 'ทีมฟุตบอลโรงเรียนบ้านหนองสระพังฯ', 'โรงเรียนบ้านโนนราษีโคกล่าม', 'ทีมฟุตบอลโรงเรียนบ้านโนนราษีฯ', 'นายบรรเจิด กองเกิด', 'รอบชิงชนะเลิศ 3-1', nowIso],
    ['FB_K_MIX', 'ศพด.อบต.วังยาว', 'ทีมฟุตบอลอนุบาล ศพด.อบต.วังยาว', 'โรงเรียนบ้านวังยาววิทยายน', 'ทีมฟุตบอลอนุบาล ร.ร.บ้านวังยาวฯ', 'ศพด.อบต.เขวาไร่', 'ทีมฟุตบอลอนุบาล ศพด.อบต.เขวาไร่', 'ศพด.วัดบ้านวังกุง', 'ทีมฟุตบอลอนุบาล ศพด.วัดบ้านวังกุง', 'นายกิจภูมิ บุญอบ', '', nowIso],
    ['FS_P46_M', 'โรงเรียนบ้านหนองสระพังโนนสะอาด', 'ทีมฟุตซอล ร.ร.บ้านหนองสระพังฯ', 'โรงเรียนดอนกลางนุกูลวิทย์', 'ทีมฟุตซอล ร.ร.ดอนกลางนุกูลวิทย์', 'โรงเรียนบ้านวังโพน', 'ทีมฟุตซอล ร.ร.บ้านวังโพน', 'โรงเรียนบ้านโคกกลาง', 'ทีมฟุตซอล ร.ร.บ้านโคกกลาง', 'นายปวรัฐ บุตรดี', '', nowIso],
    ['VB_P46_F', 'โรงเรียนหินแห่เสริมศิลป์', 'ทีมวอลเลย์บอลหญิง ร.ร.หินแห่เสริมศิลป์', 'โรงเรียนบ้านวังยาววิทยายน', 'ทีมวอลเลย์บอลหญิง ร.ร.บ้านวังยาวฯ', 'โรงเรียนบ้านโนนราษีโคกล่าม', 'ทีมวอลเลย์บอลหญิง ร.ร.บ้านโนนราษีฯ', 'โรงเรียนบ้านทิพโสต', 'ทีมวอลเลย์บอลหญิง ร.ร.บ้านทิพโสต', 'นายมาศพล ประทุมแพง', '', nowIso],
    ['TR_100_P56_M', 'โรงเรียนดอนกลางนุกูลวิทย์', 'เด็กชายกิตติศักดิ์ ใจกล้า', 'โรงเรียนบ้านวังยาววิทยายน', 'เด็กชายศุภกร แสงทอง', 'โรงเรียนบ้านหนองเขื่อน', 'เด็กชายพีรพล สุขดี', '', '', 'นายณัฐพล ผลรุ่ง', '', nowIso]
  ];
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}
