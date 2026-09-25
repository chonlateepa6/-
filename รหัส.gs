const EMP_SHEET_NAME = 'employees';
const SPREADSHEET_ID = ''; // ใส่ Spreadsheet ID ของคุณ (ถ้ามี) หรือปล่อยว่างไว้

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ระบบลงเวลาทำงาน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getEmployeeInfo(empId) {
  try {
    const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(EMP_SHEET_NAME);
    if (!sheet) return { found: false };

    const rows = sheet.getDataRange().getValues();
    const cleanInputId = String(empId).replace(/^'/, '').trim().toLowerCase();
    
    let empData = null;
    for (let i = 1; i < rows.length; i++) {
      let sheetEmpId = String(rows[i][0]).replace(/^'/, '').trim().toLowerCase();
      if (sheetEmpId === cleanInputId) {
        empData = {
          found: true,
          name: rows[i][1],
          department: rows[i][2],
          position: rows[i][3] || '-' // ดึงตำแหน่งงานจากชีตพนักงาน (คอลัมน์ D)
        };
        break;
      }
    }

    if (empData) {
      const statusCheck = checkAttendanceStatus(empId);
      empData.hasCheckedIn = statusCheck.hasCheckedIn;
      empData.hasCheckedOut = statusCheck.hasCheckedOut;
      return empData;
    }
    return { found: false };
  } catch (error) {
    return { found: false, error: error.toString() };
  }
}

function checkAttendanceStatus(empId) {
  try {
    const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    const cleanEmpId = String(empId).replace(/^'/, '').trim();
    
    const sheets = ss.getSheets();
    let targetSheetName = null;
    for (let s of sheets) {
      if (s.getName().startsWith(cleanEmpId + ' -')) {
        targetSheetName = s.getName();
        break;
      }
    }

    if (!targetSheetName) return { hasCheckedIn: false, hasCheckedOut: false };
    const sheet = ss.getSheetByName(targetSheetName);

    const rows = sheet.getDataRange().getValues();
    const now = new Date();
    const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy');

    let hasCheckedIn = false;
    let hasCheckedOut = false;

    for (let i = 1; i < rows.length; i++) {
      let rawDate = rows[i][1];
      let rowDateStr = '';
      
      if (rawDate instanceof Date) {
        rowDateStr = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      } else {
        rowDateStr = String(rawDate).replace(/^'/, '').trim();
      }

      const typeVal = String(rows[i][5]).trim();

      if (rowDateStr === todayStr) {
        if (typeVal === 'เข้างาน' || typeVal === 'OT เช้า') {
          hasCheckedIn = true;
        }
        if (typeVal === 'เลิกงาน') {
          hasCheckedOut = true;
        }
      }
    }
    return { hasCheckedIn, hasCheckedOut };
  } catch (error) {
    return { hasCheckedIn: false, hasCheckedOut: false };
  }
}

function evaluateAttendanceTypeAndScore(timeStr) {
  const parts = timeStr.split(':');
  const totalMinutes = parseInt(parts[0]) * 60 + parseInt(parts[1]);

  let type = '';
  let score = 0;
  let status = 'ปกติ';

  if (totalMinutes >= 405 && totalMinutes < 435) {
    type = 'OT เช้า';
    if (totalMinutes >= 405 && totalMinutes <= 415) { score = 5; status = 'ปกติ'; }
    else if (totalMinutes >= 416 && totalMinutes <= 420) { score = 4; status = 'ปกติ'; }
    else if (totalMinutes >= 421 && totalMinutes <= 423) { score = 3; status = 'ปกติ'; }
    else if (totalMinutes >= 424 && totalMinutes <= 426) { score = 2; status = 'ปกติ'; }
    else if (totalMinutes >= 427 && totalMinutes <= 429) { score = 1; status = 'ปกติ'; }
    else { score = 0; status = 'เกินเวลา'; }
  } else {
    type = 'เข้างาน';
    if (totalMinutes >= 455 && totalMinutes <= 465) { score = 5; status = 'ปกติ'; }
    else if (totalMinutes >= 466 && totalMinutes <= 470) { score = 4; status = 'ปกติ'; }
    else if (totalMinutes >= 471 && totalMinutes <= 480) { score = 3; status = 'ปกติ'; }
    else if (totalMinutes >= 481 && totalMinutes <= 485) { score = 2; status = 'สาย'; }
    else if (totalMinutes >= 486 && totalMinutes <= 490) { score = 1; status = 'สาย'; }
    else { score = 0; status = 'สายมาก'; }
  }

  return { type, score, status };
}

function recordAttendance(data) {
  try {
    // บังคับตรวจสอบพิกัด GPS: หากไม่มีพิกัดจะไม่ให้บันทึก
    if (!data.lat || !data.lon) {
      return { 
        status: 'error', 
        message: '⚠️ จำเป็นต้องเปิดและอนุญาตสิทธิ์เข้าถึงตำแหน่ง (GPS) จึงจะสามารถลงเวลาได้' 
      };
    }

    const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    
    const cleanName = String(data.name).replace(/[\\/?*[\]]/g, '').trim();
    const cleanEmpId = String(data.empId).replace(/^'/, '').trim();
    const individualSheetName = cleanEmpId + ' - ' + cleanName;

    let sheet = ss.getSheetByName(individualSheetName);
    
    // โครงสร้างหัวตารางใหม่ เพิ่ม 'position' (ตำแหน่งงาน)
    // 1:timestamp, 2:date, 3:empId, 4:name, 5:department, 6:position, 7:type, 8:time, 9:status, 10:score, 11:computerId, 12:location
    const expectedHeaders = ['timestamp', 'date', 'empId', 'name', 'department', 'position', 'type', 'time', 'status', 'score', 'computerId', 'location'];

    if (!sheet) {
      sheet = ss.insertSheet(individualSheetName);
      sheet.appendRow(expectedHeaders);
      sheet.getRange(1, 1, 1, expectedHeaders.length).setFontWeight('bold').setBackground('#e0f2fe');
    } else {
      let headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn() || expectedHeaders.length);
      let headers = headerRange.getValues()[0];
      if (!headers.includes('position')) {
        // หากชีตเดิมยังไม่มีคอลัมน์ position ให้สร้างชีตใหม่รองรับอัตโนมัติหรือจัดการต่อท้าย
        // เพื่อความเรียบร้อยแนะนำให้ระบบเช็กหัวข้อ
      }
    }
    
    const now = new Date();
    const dateStr = "'" + Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');

    let finalType = data.type;
    let score = '-';
    let status = 'ปกติ';

    if (finalType === 'เลิกงาน') {
      const check = checkAttendanceStatus(data.empId);
      if (!check.hasCheckedIn) {
        return { 
          status: 'error', 
          message: '⚠️ ยังไม่ได้ลงเวลาทำงานของวันนี้ จึงไม่สามารถลงเวลาเลิกงานได้ครับ' 
        };
      }
      if (check.hasCheckedOut) {
        return { 
          status: 'error', 
          message: '⚠️ คุณได้ลงเวลาเลิกงานของวันนี้ไปเรียบร้อยแล้วครับ' 
        };
      }
    } else {
      const check = checkAttendanceStatus(data.empId);
      if (check.hasCheckedIn) {
        return { 
          status: 'error', 
          message: '⚠️ คุณได้ลงเวลาทำงานของวันนี้ไปแล้วครับ' 
        };
      }
      const evaluated = evaluateAttendanceTypeAndScore(timeStr);
      finalType = evaluated.type;
      score = evaluated.score;
      status = evaluated.status;
    }

    const empIdAsText = "'" + cleanEmpId;
    const computerId = data.computerId || 'Unknown PC';
    const locationStr = data.lat + ', ' + data.lon;

    sheet.appendRow([
      now,
      dateStr,
      empIdAsText,
      data.name,
      data.department || '-',
      data.position || '-',
      finalType,
      timeStr,
      status,
      score,
      computerId,
      locationStr
    ]);

    return { status: 'success', message: 'บันทึก "' + finalType + '" สำเร็จ เวลา ' + timeStr };
  } catch (error) {
    return { status: 'error', message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}
