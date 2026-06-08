/**
 * ════════════════════════════════════════════════════════════════
 *  OPD Daily Auto-Export Script
 *  วิธีใช้: ติดตั้งใน Google Apps Script → ตั้ง Time Trigger ทุกวัน 09:50 น.
 *  ผลลัพธ์: สร้าง/อัปเดตไฟล์ "OPD_Daily_Export.csv" ใน Google Drive
 * ════════════════════════════════════════════════════════════════
 */

// ── CONFIG: แก้ค่าเหล่านี้ให้ตรงกับ Sheet ของคุณ ──────────────
var CONFIG = {
  SPREADSHEET_ID : '1qYdwXuCHDHeHN6a8vU_RVFBYT5MuYq-8K5AMK3cP4DM',
  SHEET_GID      : 2052764970,          // gid ของ tab ที่ต้องการ
  OUTPUT_FILENAME: 'OPD_Daily_Export.csv',
  DRIVE_FOLDER_ID: '',                  // ว่าง = บันทึกใน My Drive หลัก
                                        // ใส่ Folder ID ถ้าต้องการ folder เฉพาะ
};

// ── MAIN FUNCTION (รัน function นี้) ─────────────────────────────
function exportOpdDailyToDrive() {
  try {
    // 1. เปิด Sheet
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheets = ss.getSheets();
    var sheet  = sheets.find(function(s){ return s.getSheetId() === CONFIG.SHEET_GID; });

    if (!sheet) {
      Logger.log('ERROR: ไม่พบ sheet ที่มี gid = ' + CONFIG.SHEET_GID);
      return;
    }

    Logger.log('พบ sheet: ' + sheet.getName());

    // 2. ดึงข้อมูลทั้งหมด
    var range  = sheet.getDataRange();
    var values = range.getValues();

    // 3. แปลงเป็น CSV (handle commas + quotes)
    var csv = values.map(function(row) {
      return row.map(function(cell) {
        var val = String(cell);
        // ถ้ามี comma หรือ quote → ครอบด้วย quotes
        if (val.indexOf(',') !== -1 || val.indexOf('"') !== -1 || val.indexOf('\n') !== -1) {
          return '"' + val.replace(/"/g, '""') + '"';
        }
        return val;
      }).join(',');
    }).join('\n');

    // 4. บันทึกหรืออัปเดตไฟล์ใน Drive
    var folder = CONFIG.DRIVE_FOLDER_ID
      ? DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID)
      : DriveApp.getRootFolder();

    var files = folder.getFilesByName(CONFIG.OUTPUT_FILENAME);
    var file;
    if (files.hasNext()) {
      // อัปเดตไฟล์เดิม
      file = files.next();
      file.setContent(csv);
      Logger.log('อัปเดตไฟล์เดิม: ' + file.getId());
    } else {
      // สร้างไฟล์ใหม่
      file = folder.createFile(CONFIG.OUTPUT_FILENAME, csv, MimeType.PLAIN_TEXT);
      Logger.log('สร้างไฟล์ใหม่: ' + file.getId());
    }

    Logger.log('✅ Export สำเร็จ: ' + values.length + ' rows → ' + CONFIG.OUTPUT_FILENAME);
    Logger.log('Export เวลา: ' + new Date().toLocaleString('th-TH'));

  } catch(e) {
    Logger.log('❌ ERROR: ' + e.message);
  }
}

// ── ตั้ง Time Trigger อัตโนมัติ (รัน function นี้ครั้งเดียว) ──────
function createDailyTrigger() {
  // ลบ trigger เก่าที่มีชื่อเดียวกันก่อน (ป้องกัน duplicate)
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'exportOpdDailyToDrive') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // สร้าง trigger ใหม่: ทุกวัน เวลา 09:45 - 10:00 น.
  ScriptApp.newTrigger('exportOpdDailyToDrive')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .nearMinute(45)
    .create();

  Logger.log('✅ Trigger ตั้งเรียบร้อย: ทุกวัน ~09:45 น.');
}
