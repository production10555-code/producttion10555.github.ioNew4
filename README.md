/**
 * 🔒 ตั้งค่าความปลอดภัยระบบ 
 */
var ADMIN_PIN = "1234";

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('QR Code Scanner')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); 
}

function doPost(e) {
  var origin = e && e.postData && e.postData.contents ? e.postData.contents : null;
  var rawData = origin;
  if (!rawData && e && e.parameter && e.parameter.jsonData) {
    rawData = e.parameter.jsonData;
  }

  try {
    if (!rawData) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: "⚠️ ไม่พบข้อมูลที่ส่งมา" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    var postData = JSON.parse(rawData);
    var action = postData.action ? postData.action.toString().trim() : "scan";

    if (action === "getLogs") {
      return ContentService.createTextOutput(JSON.stringify({ success: true, logs: getLogsData() }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    if (action === "deleteLog") {
      return ContentService.createTextOutput(JSON.stringify(deleteLogFromServer(postData.timestamp, postData.qrCode)))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    if (action === "adminReset") {
      return ContentService.createTextOutput(JSON.stringify(adminResetLogStatus(postData.qrCode, postData.pin)))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ส่วนของการสแกน
    var scanResult = updateAndFetchDetails(postData.qrData, postData.operatorName, postData.batchId);
    return ContentService.createTextOutput(JSON.stringify(scanResult))
                         .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: "🚨 API Error: " + err.message }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

function updateAndFetchDetails(qrData, operatorName, batchId) {
  var cleanQrData = qrData ? String(qrData).trim() : "";
  var cleanBatch = batchId ? String(batchId).trim() : "";
  
  if (cleanQrData === "") {
    return { success: false, message: "ข้อมูล QR Code ว่างเปล่า" };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Recipe master"); 
    var logSheet = ss.getSheetByName("Mixing_Logs");
    
    if (!sheet) return { success: false, message: "ไม่พบชีท Recipe master" };
    if (!logSheet) {
      logSheet = ss.insertSheet("Mixing_Logs");
      logSheet.appendRow(["วัน-เวลาที่บันทึก", "QR Code ID", "ชื่อสาร (Name)", "ขนาดถุง", "ผู้สแกน", "Batch", "สถานะการสแกน"]);
    }
    
    var data = sheet.getDataRange().getValues();
    var rowData = null;
    
    // ค้นหา QR Code (เปรียบเทียบแบบ String เสมอ)
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim() === cleanQrData) {
        rowData = data[i];
        break;
      }
    }
    
    if (!rowData) return { success: false, message: "ไม่พบรหัส QR: " + cleanQrData };
    
    var currentSubstanceName = String(rowData[1]).trim();
    var bagSize = rowData[3];
    var totalNeeded = parseInt(rowData[4]) || 0;
    
    var logData = logSheet.getDataRange().getValues();
    var timeZone = Session.getScriptTimeZone();
    var todayStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
    var cloudScannedCount = 0;
    
    // นับจำนวนการสแกนในวันนี้
    for (var k = 1; k < logData.length; k++) {
      if (!logData[k][0]) continue; 
      var logDateStr = Utilities.formatDate(new Date(logData[k][0]), timeZone, "yyyy-MM-dd");
      if (logDateStr === todayStr && 
          String(logData[k][2]).trim() === currentSubstanceName && 
          String(logData[k][5]).trim() === "Batch " + cleanBatch && 
          String(logData[k][6]).trim() === "บันทึกสำเร็จ") {
        cloudScannedCount++;
      }
    }
    
    if (totalNeeded > 0 && cloudScannedCount >= totalNeeded) {
      logSheet.appendRow([new Date(), cleanQrData, currentSubstanceName, bagSize, operatorName, "Batch " + cleanBatch, "ปฏิเสธ: เกินโควตา"]);
      return { success: false, message: "สแกนไม่ได้! ครบโควตาแล้ว (" + cloudScannedCount + "/" + totalNeeded + ")" };
    }
    
    logSheet.appendRow([new Date(), cleanQrData, currentSubstanceName, bagSize, operatorName, "Batch " + cleanBatch, "บันทึกสำเร็จ"]);
    
    return {
      success: true,
      message: "บันทึกสำเร็จ! สาร " + currentSubstanceName + " สแกนแล้ว " + (cloudScannedCount + 1) + "/" + totalNeeded + " ถุง"
    };
    
  } catch (err) {
    return { success: false, message: "System Error: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

// ฟังก์ชันอื่นๆ คงเดิมตามที่คุณมี...
function adminResetLogStatus(qrData, inputPin) { /* ...โค้ดเดิมของคุณ... */ return {success: true, message: "รีเซ็ตแล้ว"}; }
function checkAndTriggerDailyBackup(currentSpreadsheet) { /* ...โค้ดเดิมของคุณ... */ }
function getLogsData() { /* ...โค้ดเดิมของคุณ... */ }
function deleteLogFromServer(timestamp, qrCode) { /* ...โค้ดเดิมของคุณ... */ }
