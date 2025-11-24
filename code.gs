// =================================================================
// FULL APPS SCRIPT (Final Version: Complete with Telegram & LINE)
// ระบบลงเวลา แหลมทองกิจเกษตร
// =================================================================

// 1. CONFIGURATION (ตั้งค่าระบบ)
const CONFIG = {
  SPREADSHEET_ID: '1G1Q8DGANy18H8jF4Ir2LYrNPMdmvH2eznvwWBoWzUAk',

  LINE_CONFIG: {
    CHANNEL_ACCESS_TOKEN: 'J0X3cz4db+3+7KX3PvyDB6sHBVCVRDvyMt/8oys54VzyYXhc5pPV1jyu1Meotrsne+tvLaxdDGXDdTPHkLln+YF/AXnjYs8pecszuuF7xabAmobzYp3npqAKg45G3Rbx+nGAF/T2g+lfCjgSFv83qQdB04t89/1O/w1cDnyilFU=' 
  },

  TELEGRAM_CONFIG: {
    token: '8048773789:AAESze8sFkvY_lPc2XsJbHI6gcV0yywv6Jo',
    chatId: '-1002685397724'
  },

  SHEET_NAMES: {
    MAIN_DATA: "ข้อมูล",
    EMPLOYEES: "รายชื่อ",
    DATABASE: "Database",
    ERROR_LOGS: "Error Logs", 
    SECURITY_LOG: "SecurityLog"
  },

  PERFORMANCE: {
    SEARCH_LAST_ROWS: 50,
    LOCK_WAIT_MS: 10000
  }
};

// 2. WEBHOOK & HANDLERS
function doPost(e) {
  const ALLOWED_FUNCTIONS = { getEmployees, getEmployeeStatus, clockIn, clockOut, getEmployeeFromDatabase };
  try {
    const body = (e && e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : {};
    const fn = ALLOWED_FUNCTIONS[body.functionName];
    if (!fn) throw new Error(`Function not allowed.`);
    const data = fn.apply(null, body.args || []);
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', data })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    logError('doPost Error', error.message);
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('ระบบลงเวลา')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 3. LINE PUSH MESSAGE (with Debug Logging)
function sendLinePushMessage(userId, message) {
  console.log(`[DEBUG] sendLinePushMessage called with userId: ${userId}`);
  
  if (!userId) {
    logError('LINE Skipped', 'No UserID provided');
    console.log('[DEBUG] LINE Push skipped - no userId');
    return;
  }
  
  const token = CONFIG.LINE_CONFIG.CHANNEL_ACCESS_TOKEN;
  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = { 'to': userId, 'messages': [{ 'type': 'text', 'text': message }] };
  
  console.log(`[DEBUG] Sending LINE Push to: ${userId.substring(0,10)}...`);
  console.log(`[DEBUG] Message length: ${message.length} chars`);
  
  try {
    const response = UrlFetchApp.fetch(url, {
      'method': 'post',
      'headers': { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token },
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    });
    
    const code = response.getResponseCode();
    const body = response.getContentText();
    
    console.log(`[DEBUG] LINE API Response Code: ${code}`);
    console.log(`[DEBUG] LINE API Response Body: ${body}`);
    
    if (code !== 200) {
      logError('LINE Push Error', `Code: ${code} | Body: ${body}`);
    } else {
      console.log(`[SUCCESS] LINE Push sent to ${userId}`);
    }
  } catch (e) {
    console.error(`[ERROR] LINE Exception: ${e.message}`);
    logError('LINE Exception', e.message);
  }
}

// 4. CLOCK IN/OUT
function clockIn(employee, gps, lineUserId) {
  console.log(`[DEBUG] clockIn called - employee: ${employee}, lineUserId: ${lineUserId || 'NOT PROVIDED'}`);
  return handleClockAction(employee, gps, lineUserId, 'IN');
}

function clockOut(employee, gps, lineUserId) {
  console.log(`[DEBUG] clockOut called - employee: ${employee}, lineUserId: ${lineUserId || 'NOT PROVIDED'}`);
  return handleClockAction(employee, gps, lineUserId, 'OUT');
}

function handleClockAction(employee, gps, lineUserId, type) {
  const lock = LockService.getScriptLock();
  lock.tryLock(CONFIG.PERFORMANCE.LOCK_WAIT_MS);

  try {
    const ss = _getSpreadsheet();
    if (!ss) return [['ERROR', 'DB_ERROR', 'ไม่พบไฟล์ Google Sheet']];
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIN_DATA);
    if (!sheet) return [['ERROR', 'SHEET_ERROR', 'ไม่พบแท็บข้อมูล']];

    const now = new Date();
    const location = getCachedLocation(gps?.[0], gps?.[1]);
    const lastRow = sheet.getLastRow();
    const startRow = Math.max(2, lastRow - CONFIG.PERFORMANCE.SEARCH_LAST_ROWS);

    if (lastRow >= 2) {
      const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 3).getValues();
      let foundPending = false;
      let pendingRow = -1;

      for (let i = data.length - 1; i >= 0; i--) {
        if (data[i][0] === employee && data[i][2] === "") {
          foundPending = true;
          pendingRow = startRow + i;
          break;
        }
      }

      if (type === 'IN' && foundPending) return [['FAILED', 'Warning', 'คุณยังไม่ได้ลงเวลาออกงานครั้งก่อน']];
      if (type === 'OUT' && !foundPending) return [['FAILED', 'Warning', 'ไม่พบเวลาเข้างานล่าสุดของคุณ']];
      
      if (type === 'OUT') {
        const inTime = new Date(sheet.getRange(pendingRow, 2).getValue());
        const hours = ((now - inTime) / 3600000).toFixed(2);
        
        // ✅ แก้ไข: เขียนแยกคอลัมน์ให้ถูกต้อง
        // (C) เวลาออก
        sheet.getRange(pendingRow, 3).setValue(now);
        // (E) location check-out (ข้ามคอลัมน์ D)
        sheet.getRange(pendingRow, 5).setValue(location);
        // (F) ชั่วโมง
        sheet.getRange(pendingRow, 6).setValue(hours);
        
        // ส่ง Telegram
        sendTelegramNotification(formatTelegramClockOut(employee, now, gps?.[0], gps?.[1], location, hours));
        
        // ส่ง LINE (ใช้รูปแบบเดียวกับ Telegram)
        console.log(`[DEBUG] Clock OUT - Preparing LINE notification for userId: ${lineUserId || 'NONE'}`);
        if (lineUserId) {
          const lineMessage = formatTelegramClockOut(employee, now, gps?.[0], gps?.[1], location, hours);
          console.log(`[DEBUG] Calling sendLinePushMessage for Clock OUT`);
          sendLinePushMessage(lineUserId, lineMessage);
        } else {
          console.log(`[WARNING] LINE notification skipped - no lineUserId provided`);
        }
        
        try { TotalHours(employee); } catch(e){}
        return [['SUCCESS', getDate(now), employee, location]];
      }
    } else if (type === 'OUT') {
       return [['FAILED', 'Warning', 'ไม่พบประวัติการเข้างาน']];
    }

    // กรณี Clock IN - เขียนที่ (A)ชื่อ (B)เวลาเข้า (C)ว่าง (D)location check-in (E)ว่าง (F)ว่าง
    sheet.appendRow([employee, now, "", location, "", ""]);
    
    // ส่ง Telegram
    sendTelegramNotification(formatTelegramClockIn(employee, now, gps?.[0], gps?.[1], location));
    
    // ส่ง LINE (ใช้รูปแบบเดียวกับ Telegram)
    console.log(`[DEBUG] Clock IN - Preparing LINE notification for userId: ${lineUserId || 'NONE'}`);
    if (lineUserId) {
      const lineMessage = formatTelegramClockIn(employee, now, gps?.[0], gps?.[1], location);
      console.log(`[DEBUG] Calling sendLinePushMessage for Clock IN`);
      sendLinePushMessage(lineUserId, lineMessage);
    } else {
      console.log(`[WARNING] LINE notification skipped - no lineUserId provided`);
    }
    
    return [['SUCCESS', getDate(now), employee, location]];

  } catch (e) {
    logError(`Clock${type} Error`, e.message);
    return [['ERROR', 'System Error', e.message]];
  } finally {
    lock.releaseLock();
  }
}

// 5. UTILS
function getEmployees() {
  console.log('[DEBUG] getEmployees called');
  const ss = _getSpreadsheet();
  if (!ss) {
    console.log('[DEBUG] Spreadsheet not found');
    return [];
  }
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EMPLOYEES);
  if (!sheet) {
    console.log('[DEBUG] EMPLOYEES sheet not found');
    return [];
  }
  const lastRow = sheet.getLastRow();
  console.log(`[DEBUG] EMPLOYEES sheet lastRow: ${lastRow}`);
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const filtered = data.filter(r => r[0] && r[0].toString().trim() !== "").map(r => [r[0]]);
  console.log(`[DEBUG] Found ${filtered.length} employees`);
  return filtered;
}

function getCachedLocation(lat, lng) {
  if(!lat || !lng) return "ไม่ระบุพิกัด";
  try {
    const cache = CacheService.getScriptCache();
    const key = `loc_${Number(lat).toFixed(4)}_${Number(lng).toFixed(4)}`;
    const cached = cache.get(key);
    if(cached) return cached;
    const res = Maps.newGeocoder().reverseGeocode(lat, lng);
    const loc = res?.results?.[0]?.formatted_address || "ไม่ระบุพิกัด";
    cache.put(key, loc, 21600);
    return loc;
  } catch(e) { return "Error Map API"; }
}

function logError(msg, detail) {
  try {
    const ss = _getSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ERROR_LOGS);
    if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAMES.ERROR_LOGS).appendRow(['Time', 'Msg', 'Detail']);
    sheet.appendRow([new Date(), msg, String(detail)]);
  } catch(e) { console.error(e); }
}

function _getSpreadsheet() { try { return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); } catch(e) { return null; } }
function addZero(i) { return i < 10 ? "0" + i : i; }

function getFullDateTime(d) {
  const day = addZero(d.getDate());
  const month = addZero(d.getMonth() + 1);
  const year = d.getFullYear() + 543; // แปลงเป็น พ.ศ.
  const hours = addZero(d.getHours());
  const minutes = addZero(d.getMinutes());
  const seconds = addZero(d.getSeconds());
  return `วันที่ ${day}/${month}/${year} เวลา ${hours}:${minutes}:${seconds} น.`;
}

function getDate(d) { return `${addZero(d.getDate())}/${addZero(d.getMonth()+1)} ${addZero(d.getHours())}:${addZero(d.getMinutes())}`; }
function formatDateForTelegram(d) { return getDate(d); }

function formatTelegramClockIn(n, d, lat, lng, loc) { 
  const mapLink = lat && lng ? `https://www.google.com/maps?q=${lat},${lng}` : '';
  return `━━━━━━━━━━━━━━━━━━━━
🟢 ลงเวลาเข้างาน
━━━━━━━━━━━━━━━━━━━━
👤 พนักงาน: ${n}
⏰ เวลา: ${getFullDateTime(d)}
📍 ตำแหน่ง:
${loc}
${mapLink ? '\n🗺️ ดูบนแผนที่ (' + mapLink + ')' : ''}`;
}

function formatTelegramClockOut(n, d, lat, lng, loc, h) { 
  const mapLink = lat && lng ? `https://www.google.com/maps?q=${lat},${lng}` : '';
  return `━━━━━━━━━━━━━━━━━━━━
🔴 ลงเวลาเลิกงาน
━━━━━━━━━━━━━━━━━━━━
👤 พนักงาน: ${n}
⏰ เวลา: ${getFullDateTime(d)}
⏱️ ทำงาน: ${h} ชั่วโมง
📍 ตำแหน่ง:
${loc}
${mapLink ? '\n🗺️ ดูบนแผนที่ (' + mapLink + ')' : ''}`;
}

function sendTelegramNotification(msg) {
  if (!CONFIG.TELEGRAM_CONFIG.token) return;
  try {
    UrlFetchApp.fetch(`https://api.telegram.org/bot${CONFIG.TELEGRAM_CONFIG.token}/sendMessage`, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ chat_id: CONFIG.TELEGRAM_CONFIG.chatId, text: msg, parse_mode: 'Markdown' })
    });
  } catch(e){}
}

// 6. EMPLOYEE DATA FUNCTIONS
function TotalHours(employeeName) {
  try {
    const ss = _getSpreadsheet();
    if (!ss) return;
    
    const dataSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIN_DATA);
    if (!dataSheet) return;
    
    const lastRow = dataSheet.getLastRow();
    if (lastRow < 2) return;
    
    // Get all data for this employee (columns A-F: Name, ClockIn, ClockOut, LocIn, LocOut, Hours)
    const allData = dataSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    let totalHours = 0;
    
    // Calculate total hours (column F, index 5)
    allData.forEach(row => {
      if (row[0] === employeeName && row[5] !== "" && !isNaN(parseFloat(row[5]))) {
        totalHours += parseFloat(row[5]);
      }
    });
    
    // Update employee sheet with total hours (column B)
    const empSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EMPLOYEES);
    if (!empSheet) return;
    
    const empLastRow = empSheet.getLastRow();
    if (empLastRow < 2) return;
    
    const empData = empSheet.getRange(2, 1, empLastRow - 1, 2).getValues();
    for (let i = 0; i < empData.length; i++) {
      if (empData[i][0] === employeeName) {
        empSheet.getRange(i + 2, 2).setValue(totalHours.toFixed(2));
        break;
      }
    }
  } catch (e) {
    logError('TotalHours Error', e.message);
  }
}

function getEmployeeStatus(employeeName) {
  try {
    const ss = _getSpreadsheet();
    if (!ss) return { status: 'ERROR', message: 'ไม่พบไฟล์ Google Sheet' };
    
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MAIN_DATA);
    if (!sheet) return { status: 'ERROR', message: 'ไม่พบแท็บข้อมูล' };
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'OUT', message: 'ไม่มีประวัติการทำงาน' };
    
    const startRow = Math.max(2, lastRow - CONFIG.PERFORMANCE.SEARCH_LAST_ROWS);
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 3).getValues();
    
    // Search from bottom to top for most recent record
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i][0] === employeeName) {
        if (data[i][2] === "") {
          // Found pending clock-in (no clock-out time)
          return { 
            status: 'IN', 
            message: 'กำลังทำงานอยู่',
            clockInTime: data[i][1]
          };
        } else {
          // Found completed record
          return { 
            status: 'OUT', 
            message: 'ไม่ได้ทำงาน',
            lastClockOut: data[i][2]
          };
        }
      }
    }
    
    return { status: 'OUT', message: 'ไม่พบประวัติการทำงาน' };
  } catch (e) {
    logError('getEmployeeStatus Error', e.message);
    return { status: 'ERROR', message: e.message };
  }
}

function getEmployeeFromDatabase(employeeName) {
  try {
    const ss = _getSpreadsheet();
    if (!ss) return null;
    
    // Try DATABASE sheet first, fallback to EMPLOYEES sheet
    let dbSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DATABASE);
    if (!dbSheet) {
      dbSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EMPLOYEES);
    }
    if (!dbSheet) return null;
    
    const lastRow = dbSheet.getLastRow();
    if (lastRow < 2) return null;
    
    // Assuming sheet has columns: [Name, LINE User ID, ...]
    const data = dbSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === employeeName && data[i][1]) {
        return {
          name: data[i][0],
          lineUserId: data[i][1]
        };
      }
    }
    
    return null;
  } catch (e) {
    logError('getEmployeeFromDatabase Error', e.message);
    return null;
  }
}