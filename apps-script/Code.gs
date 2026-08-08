/**
 * รับผลแบบประเมินจากเว็บ แล้วบันทึกลง Google Sheet
 * ใช้คู่กับ https://nattakitt-tin.github.io/sikharee-wellness/
 *
 * วิธีติดตั้ง: ดูไฟล์ README.md หัวข้อ "บันทึกผลลง Google Sheet"
 */

/* ---------- ตั้งค่า ---------- */

// ชื่อแท็บที่จะเก็บผล (ถ้ายังไม่มี ระบบจะสร้างให้เอง)
const SHEET_NAME = 'ผลการประเมิน';

// รหัสลับ ต้องตรงกับที่กรอกในหน้า ⚙ ของเว็บ — เปลี่ยนเป็นข้อความของคุณเอง
const SECRET = 'sikharee-2026';

/* ---------- ตัวรับข้อมูล ---------- */

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return reply_({ ok: false, error: 'ไม่มีข้อมูลส่งมา' });
    }

    const data = JSON.parse(e.postData.contents);

    if (String(data.secret || '') !== SECRET) {
      return reply_({ ok: false, error: 'รหัสลับไม่ถูกต้อง' });
    }

    // ป้องกันการเขียนชนกันเมื่อมีคนส่งพร้อมกัน
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = getSheet_();

      // ถ้าเคยบันทึกรหัสอ้างอิงนี้แล้ว ไม่ต้องเขียนซ้ำ (กันเน็ตหลุดแล้วส่งใหม่)
      if (data.id && findById_(sh, data.id)) {
        return reply_({ ok: true, duplicate: true });
      }

      const headers = syncHeaders_(sh, data.fields || {});
      const row = headers.map(function (h) {
        return Object.prototype.hasOwnProperty.call(data.fields, h) ? data.fields[h] : '';
      });
      sh.appendRow(row);

      return reply_({ ok: true, row: sh.getLastRow() });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return reply_({ ok: false, error: String(err) });
  }
}

// เปิด URL ด้วยเบราว์เซอร์เพื่อเช็กว่า deploy สำเร็จ
function doGet() {
  return reply_({ ok: true, message: 'พร้อมรับข้อมูล', sheet: SHEET_NAME });
}

/* ---------- ตัวช่วย ---------- */

function reply_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

/**
 * ทำให้แถวหัวตารางมีครบทุกคอลัมน์ที่ส่งมา
 * ถ้าเว็บเพิ่มคำถามใหม่ คอลัมน์จะถูกต่อท้ายให้อัตโนมัติ โดยข้อมูลเดิมไม่เลื่อน
 */
function syncHeaders_(sh, fields) {
  let headers = [];
  if (sh.getLastRow() > 0 && sh.getLastColumn() > 0) {
    headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
      .map(String).filter(function (h) { return h !== ''; });
  }

  const missing = Object.keys(fields).filter(function (k) {
    return headers.indexOf(k) < 0;
  });

  if (missing.length) {
    headers = headers.concat(missing);
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#e8e6da');
    sh.setFrozenRows(1);
  }
  return headers;
}

function findById_(sh, id) {
  if (sh.getLastRow() < 2) return false;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const col = headers.indexOf('รหัสอ้างอิง') + 1;
  if (col < 1) return false;
  const values = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return true;
  }
  return false;
}
