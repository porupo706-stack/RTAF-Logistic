// =============================================================================
//  ระบบจัดการคลังพัสดุ (Google Apps Script) — Code.gs  (ไฟล์ฝั่งเซิร์ฟเวอร์)
//
//  >>> เมื่อนำไปใช้หน่วยงานอื่น แก้ที่บล็อก CONFIG ด้านล่างนี้ที่เดียว <<<
//  ทั้ง Code.gs และ Index.html อ่านค่าจาก CONFIG นี้ (Index.html รับค่าผ่านฟังก์ชัน getClientConfig_)
//
//  ค้นหา [CHANGE] เพื่อดูทุกจุดที่ต้องตรวจ
//    [CHANGE]        = ต้องแก้ทุกหน่วยงาน
//    [CHANGE-ถ้า...] = แก้เฉพาะเมื่อหน่วยงานของคุณต่างจากต้นแบบ
//
//  ชีทที่ระบบต้องมี: ชีทสายงานตาม INVENTORY_SHEETS + ชีทประวัติ + ชีทสรุปรายวัน
//  (ดูโครงคอลัมน์ใน README.md)
// =============================================================================
var CONFIG = {
  // [CHANGE] ชื่อหน่วยงาน (ใช้ที่เมนู ชื่อแท็บ หัวเว็บ หัวรายงานที่พิมพ์ และไฟล์ Excel)
  UNIT_NAME: 'กองบิน ๗',
  // [CHANGE] ชื่อหน่วยงานภาษาอังกฤษ (บรรทัดรองใต้ชื่อระบบที่หัวเว็บ)
  UNIT_NAME_EN: 'Wing7',

  // [CHANGE] Web app URL (ลงท้าย /exec) ที่ได้จากการ Deploy ของคุณ แทน YOUR_DEPLOYMENT_ID
  //          ใช้สร้าง QR Code และปุ่มเปิดระบบเต็มหน้าจอ
  WEB_APP_URL: 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec',

  // [CHANGE] ลิงก์รูปโลโก้หน่วยงาน (ต้องเข้าถึงได้ตลอด)
  LOGO_URL: 'https://YOUR-DOMAIN/path/to/logo.png',

  // [CHANGE-ถ้าสายต่างจากต้นแบบ] ชื่อชีทสายงาน (ต้องตรงกับชื่อชีทใน Google Sheet ทุกตัวอักษร)
  //   เพิ่ม/ลด/เปลี่ยนชื่อสายที่นี่ที่เดียว หน้าเว็บ รายงาน และการค้นหาจะตามอัตโนมัติ
  INVENTORY_SHEETS: ['สายพลา', 'สายสื่อสาร', 'สายช่างโยธา', 'สายขนส่ง'],

  // [CHANGE-ถ้าชื่อสายต่าง] ชีทที่มีช่อง "วันที่ผลิต" และ "เลขทะเบียนรถ" (ไม่มี ให้ใส่ '')
  TRANSPORT_SHEET: 'สายขนส่ง',

  // [CHANGE-ถ้าชื่อชีทต่าง] ชีทระบบ
  HISTORY_SHEET: 'ประวัติรับจ่าย',
  SUMMARY_SHEET: 'สรุปรายวัน',

  // [CHANGE-ตามต้องการ] เกณฑ์ "พัสดุใกล้หมด" (จำนวนที่น้อยกว่าหรือเท่ากับค่านี้)
  LOW_STOCK_THRESHOLD: 5,

  // ---------- รายงาน "บัญชีตรวจสอบพัสดุ (ผนวก ค)" / Excel ----------
  // [CHANGE] ชื่อหน่วย/คลังในหัวกระดาษ (แสดงเป็น "หน่วย <ค่านี้>")
  ANNEX_UNIT_LABEL: 'คลังกลาง',
  // [CHANGE] ยศ/ตำแหน่งผู้ตรวจพัสดุ แสดงท้ายกระดาษทุกหน้า (2 บรรทัด)
  ANNEX_SIGN: [
    'เจ้าหน้าที่ตรวจสอบพัสดุ    น.ต. .........................................',
    '                                พ.อ.อ. .....................................'
  ],
  // [CHANGE-ถ้าต่าง] เดือนที่เริ่มรอบตรวจ (10 = 1 ต.ค.) ใช้เป็นวันเริ่มต้นในหน้าจอ
  ANNEX_FISCAL_START_MONTH: 10,
  // [CHANGE-ถ้าสายต่างจากต้นแบบ] รายการสายในไฟล์ Excel (1 ชีทต่อ 1 สาย)
  //   key   = ชื่อชีทในระบบ (ต้องตรงกับ INVENTORY_SHEETS; สายที่ไม่มี key ตรงกันจะขึ้น "ไม่มี")
  //   sheet = ชื่อชีทใน Excel, title = ชื่อเต็มที่แสดงในหัวกระดาษ
  ANNEX_LINES: [
    { key: 'สายสื่อสาร',    sheet: 'คงคลัง สายสื่อสาร',    title: 'สายสื่อสารอิเล็กทรอนิกส์และการภาพ' },
    { key: 'สายพลา',        sheet: 'คงคลัง สายพลาธิการ',   title: 'สายพลาธิการ' },
    { key: 'สายช่างอากาศ',  sheet: 'คงคลัง สายช่างอากาศ',  title: 'สายช่างอากาศ' },
    { key: 'สายสรรพาวุธ',   sheet: 'คงคลัง สายสรรพาวุธ',   title: 'สายสรรพาวุธ' },
    { key: 'สายขนส่ง',      sheet: 'คงคลัง สายขนส่ง',      title: 'สายขนส่ง' },
    { key: 'สายช่างโยธา',   sheet: 'คงคลัง สายช่างโยธา',   title: 'สายช่างโยธา' },
    { key: 'สายเชื้อเพลิง', sheet: 'คงคลัง สายเชื้อเพลิง', title: 'สายเชื้อเพลิง' },
    { key: 'สายแพทย์',      sheet: 'คงคลัง สายแพทย์',      title: 'สายแพทย์' }
  ]
};

// ค่าที่ส่งให้ Index.html (หน้าเว็บ)
function getClientConfig_() {
  return {
    unitName:       CONFIG.UNIT_NAME,
    unitNameEn:     CONFIG.UNIT_NAME_EN,
    webAppUrl:      CONFIG.WEB_APP_URL,
    logoUrl:        CONFIG.LOGO_URL,
    sheets:         CONFIG.INVENTORY_SHEETS,
    transportSheet: CONFIG.TRANSPORT_SHEET,
    annexUnitLabel: CONFIG.ANNEX_UNIT_LABEL,
    annexSign:      CONFIG.ANNEX_SIGN,
    fyStartMonth:   CONFIG.ANNEX_FISCAL_START_MONTH,
    annexLines:     CONFIG.ANNEX_LINES
  };
}

// ------------------------------------------------------------
//  หาตำแหน่งคอลัมน์จากหัวตาราง (รองรับทั้งชีทที่มี/ไม่มีคอลัมน์ "หน่วยนับ")
//  คืนค่าเป็น index เริ่มที่ 0 ; unit = -1 ถ้าไม่มีคอลัมน์หน่วยนับ
// ------------------------------------------------------------
function detectCols_(rows) {
  var f = {};
  for (var r = 0; r < Math.min(3, rows.length); r++) {
    for (var k = 0; k < rows[r].length; k++) {
      var t = rows[r][k] ? rows[r][k].toString().replace(/\s/g, "") : "";
      if (!t) continue;
      if (t.indexOf("หมายเลข") > -1 && f.id === undefined) f.id = k;
      else if (t === "ชื่อพัสดุ" && f.name === undefined) f.name = k;
      else if (t.indexOf("หน่วยนับ") > -1 && f.unit === undefined) f.unit = k;
      else if (t.indexOf("ที่เก็บ") > -1 && f.loc === undefined) f.loc = k;
      else if (t.indexOf("คงเหลือ") > -1 && f.qty === undefined) f.qty = k;
      else if ((t.indexOf("คิวอาร์") > -1 || t.toUpperCase().indexOf("QR") > -1) && f.qr === undefined) f.qr = k;
    }
  }
  var hasUnit = (f.unit !== undefined);
  var sh = hasUnit ? 1 : 0;           // ถ้ามีคอลัมน์หน่วยนับ คอลัมน์หลังจากนั้นเลื่อนไป 1
  var C = {
    id:   (f.id   !== undefined) ? f.id   : 1,
    name: (f.name !== undefined) ? f.name : 2,
    unit: hasUnit ? f.unit : -1,
    loc:  (f.loc  !== undefined) ? f.loc  : 3 + sh,
    qty:  (f.qty  !== undefined) ? f.qty  : 4 + sh,
    qr:   (f.qr   !== undefined) ? f.qr   : 5 + sh
  };
  C.mfg = C.qr + 1;   // วันที่ผลิต (สายขนส่ง)
  C.reg = C.qr + 2;   // เลขทะเบียนรถ (สายขนส่ง)
  return C;
}

function sheetCols_(sheet) {
  var n = Math.min(3, sheet.getLastRow());
  return detectCols_(sheet.getRange(1, 1, n, sheet.getLastColumn()).getValues());
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('คลังพัสดุ ' + CONFIG.UNIT_NAME)
    .addItem('เปิดระบบจัดการ (เต็มหน้าจอ)', 'openFullScreen') 
    .addItem('เปิดในหน้าต่าง', 'openDialog')
    .addToUi();
}

function openDialog() {
  var template = HtmlService.createTemplateFromFile('Index');
  template.itemId = "";
  template.config = getClientConfig_();
  var html = template.evaluate()
    .setWidth(1200)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'ระบบจัดการคลังพัสดุ ' + CONFIG.UNIT_NAME);
}

function openFullScreen() {
  var url = CONFIG.WEB_APP_URL;
  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:sans-serif;display:flex;flex-direction:column;align-items:center;justify-content:center;height:90vh;gap:16px;background:#f8f9fa;}</style>' +
    '<p style="color:#5f6368;font-size:16px;">กำลังเปิดระบบในแท็บใหม่...</p>' +
    '<a href="' + url + '" target="_blank" ' +
    'style="background:#1a73e8;color:white;padding:12px 28px;border-radius:10px;text-decoration:none;font-size:15px;font-weight:600;">คลิกที่นี่หากไม่เปิดอัตโนมัติ</a>' +
    '<script>window.open("' + url + '", "_blank");<\/script>'
  ).setWidth(400).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'กำลังเปิดระบบ...');
}

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  var rawId = (e.parameter && e.parameter.id) ? e.parameter.id : "";
  template.itemId = rawId.toString().trim().replace(/\s/g, "");
  template.config = getClientConfig_();
  return template.evaluate()
    .setTitle('ระบบจัดการคลังพัสดุ ' + CONFIG.UNIT_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getItemData(itemId) {
  if (!itemId) return null;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var searchId = itemId.toString().trim().replace(/^'+|'+$/g, "");
  var cleanSearch = searchId.replace(/[\s\-]/g, "").toLowerCase();
  var sheetNames = CONFIG.INVENTORY_SHEETS;
  for (var s = 0; s < sheetNames.length; s++) {
    var sheet = ss.getSheetByName(sheetNames[s]);
    if (!sheet) continue;
    var data = sheet.getDataRange().getValues();
    var C = detectCols_(data);
    for (var i = 1; i < data.length; i++) {
      var raw = data[i][C.id];
      if (raw === undefined || raw === null || raw === "") continue;
      var rowId = raw.toString().trim().replace(/^'+|'+$/g, "");
      var cleanRow = rowId.replace(/[\s\-]/g, "").toLowerCase();
      if (rowId === searchId || cleanRow === cleanSearch) {
        var manufRaw = data[i][C.mfg];
        var manufDateStr = '';
        if (manufRaw instanceof Date) {
          manufDateStr = Utilities.formatDate(manufRaw, "GMT+7", "yyyy-MM-dd");
        } else if (manufRaw) {
          manufDateStr = manufRaw.toString().trim();
        }
        var regNumStr = data[i][C.reg] ? data[i][C.reg].toString().trim() : '';
        return {
          sheetName: sheetNames[s],
          row: i + 1,
          id: rowId,
          name: data[i][C.name],
          unit: (C.unit >= 0 && data[i][C.unit]) ? data[i][C.unit].toString().trim() : "",
          location: data[i][C.loc],
          qty: data[i][C.qty],
          mfgDate: manufDateStr,
          regNum: regNumStr  
        };
      }
    }
  }
  return null;
}

function searchParcelByName(query) {
  if (!query) return []; 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];
  var cleanQuery = query.toString().toLowerCase().replace(/[\s\-\.]/g, "");
  var allSheets = ss.getSheets();
  for (var s = 0; s < allSheets.length; s++) {
    var sheet = allSheets[s];
    var sheetName = sheet.getName().trim();
    var isTargetSheet = CONFIG.INVENTORY_SHEETS.some(function(n) { return sheetName.indexOf(n) !== -1; });
    if (!isTargetSheet) continue;
    var data = sheet.getDataRange().getValues();
    var C = detectCols_(data);
    for (var i = 1; i < data.length; i++) {
      var rawId   = data[i][C.id];
      var rawName = data[i][C.name] ? data[i][C.name].toString().trim() : "";
      if (rawName === "ชื่อพัสดุ" || rawName === "") continue;
      if (rawId === undefined || rawId === null || rawId === "") continue;
      var idStr = (typeof rawId === 'number') ? rawId.toFixed(0) : rawId.toString().trim();
      var cleanId   = idStr.toLowerCase().replace(/[\s\-\.]/g, "");
      var cleanName = rawName.toLowerCase().replace(/[\s\-\.]/g, "");
      var regNum = data[i][C.reg] ? data[i][C.reg].toString().toLowerCase().replace(/[\s\-\.]/g, "") : "";
      if (cleanId.indexOf(cleanQuery) > -1 || cleanName.indexOf(cleanQuery) > -1 || regNum.indexOf(cleanQuery) > -1) {
        results.push({
          id: idStr,
          name: rawName,
          sheetName: sheetName,
          qty: data[i][C.qty] ? data[i][C.qty].toString() : "0"
        });
      }
      if (results.length >= 50) return results;
    }
  }
  return results;
}

function updateQty(sheetName, row, newQty, itemName, unitName, newLocation, newManufDate, newRegNum) { 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Error: ไม่พบชีท " + sheetName;
  var C = sheetCols_(sheet);
  var oldQty = sheet.getRange(row, C.qty + 1).getValue();
  oldQty = parseFloat(oldQty) || 0;
  newQty = parseFloat(newQty) || 0;
  var diff = newQty - oldQty;
  sheet.getRange(row, C.qty + 1).setValue(newQty);
  if (newLocation) {
    sheet.getRange(row, C.loc + 1).setValue(newLocation); 
  }
  if (newManufDate && sheetName === CONFIG.TRANSPORT_SHEET) {
    var mdParts = newManufDate.split('-');
    if (mdParts.length === 3) {
      var mdDate = new Date(parseInt(mdParts[0]), parseInt(mdParts[1]) - 1, parseInt(mdParts[2]));
      sheet.getRange(row, C.mfg + 1).setValue(mdDate);
      sheet.getRange(row, C.mfg + 1).setNumberFormat('dd/MM/yyyy');
    }
  }
  var regNumChanged = false;
  if (sheetName === CONFIG.TRANSPORT_SHEET && newRegNum !== null && newRegNum !== undefined) {
    var oldRegNum = sheet.getRange(row, C.reg + 1).getValue();
    oldRegNum = oldRegNum ? oldRegNum.toString().trim() : '';
    var newRegStr = newRegNum ? newRegNum.toString().trim() : '';
    if (oldRegNum !== newRegStr) {
      sheet.getRange(row, C.reg + 1).setValue(newRegStr);
      regNumChanged = true;
    }
  }
  if (diff === 0 && !newLocation && !newManufDate && !regNumChanged) return oldQty;
  var histSheet = ss.getSheetByName(CONFIG.HISTORY_SHEET);
  if (histSheet) {
    var histData = histSheet.getDataRange().getValues();
    var nextSeq = 1;
    if (histData.length > 1) {
      for (var rb = histData.length - 1; rb >= 1; rb--) {
        var sv = parseInt(histData[rb][0]);
        if (!isNaN(sv) && sv > 0) { nextSeq = sv + 1; break; }
      }
    }
    var cleanUnitName = "";
    if (unitName) {
      cleanUnitName = unitName.replace(/^\d+[\s-]*|^\d+/, "").trim();
    } else {
      cleanUnitName = (diff > 0) ? "รับเข้าคลัง" : "-";
    }
    var now = new Date();
    var dateReceive = (diff > 0) ? now : ""; 
    var dateIssue   = (diff < 0) ? now : "";   
    var type   = (diff > 0) ? "รับเข้า" : (diff < 0 ? "จ่ายออก" : "แก้ไขข้อมูล");
    var amount = Math.abs(diff);
    histSheet.appendRow([
      nextSeq, dateReceive, dateIssue, type,
      itemName, cleanUnitName, amount, newQty
    ]);
    var allData = histSheet.getDataRange().getValues();
    var matchRows = [];
    for (var hr = 1; hr < allData.length; hr++) {
      var hName = allData[hr][4] ? allData[hr][4].toString().trim() : "";
      if (hName === itemName.toString().trim()) {
        matchRows.push(hr + 1); 
      }
    }
    for (var ri = 0; ri < matchRows.length; ri++) {
      histSheet.getRange(matchRows[ri], 8).setValue("");
    }
    if (matchRows.length > 0) {
      histSheet.getRange(matchRows[matchRows.length - 1], 8).setValue(newQty);
    }
  }
  return newQty;
}

function generateDailySummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var historySheet = ss.getSheetByName(CONFIG.HISTORY_SHEET);
  var summarySheet = ss.getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!historySheet || !summarySheet) return "Error: ไม่พบชีท";
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var todayTime = today.getTime();
  var historyData = historySheet.getDataRange().getValues();
  if (historyData.length < 2) return "ยังไม่มีข้อมูล";
  var summaryObj = {};
  var foundData = false;
  for (var i = 1; i < historyData.length; i++) {
    var dateIn  = historyData[i][1];
    var dateOut = historyData[i][2];
    var dateVal = (dateIn instanceof Date) ? dateIn : dateOut;
    if (!(dateVal instanceof Date)) continue;
    var recordDate = new Date(dateVal);
    recordDate.setHours(0, 0, 0, 0);
    if (recordDate.getTime() !== todayTime) continue;
    var type   = historyData[i][3];
    var name   = historyData[i][4];
    var amount = parseFloat(historyData[i][6]) || 0;
    if (!name) continue;
    foundData = true;
    if (!summaryObj[name]) {
      summaryObj[name] = { receive: 0, issue: 0, lastBalance: 0 };
    }
    if (type === "รับเข้า") {
      summaryObj[name].receive += amount;
    } else if (type === "จ่ายออก") {
      summaryObj[name].issue += amount;
    }
  }
  if (foundData) {
    var actualQty = {};
    var invSheets = CONFIG.INVENTORY_SHEETS;
    for (var s = 0; s < invSheets.length; s++) {
      var invSheet = ss.getSheetByName(invSheets[s]);
      if (!invSheet) continue;
      var invData = invSheet.getDataRange().getValues();
    var IC = detectCols_(invData);
      for (var r = 1; r < invData.length; r++) {
        var iName = invData[r][IC.name] ? invData[r][IC.name].toString().trim() : "";
        if (iName && iName !== "ชื่อพัสดุ") {
          actualQty[iName] = parseFloat(invData[r][IC.qty]) || 0;
        }
      }
    }
    for (var key in summaryObj) {
      if (actualQty[key] !== undefined) {
        summaryObj[key].lastBalance = actualQty[key];
      }
    }
    var todayStr = Utilities.formatDate(today, "GMT+7", "dd/MM/yyyy");
    var summaryData = summarySheet.getDataRange().getValues();
    for (var j = summaryData.length - 1; j >= 1; j--) {
      var rowDateStr = (summaryData[j][0] instanceof Date)
        ? Utilities.formatDate(summaryData[j][0], "GMT+7", "dd/MM/yyyy")
        : summaryData[j][0].toString();
      if (rowDateStr === todayStr) summarySheet.deleteRow(j + 1);
    }
    var finalData = [];
    for (var itemName in summaryObj) {
      finalData.push([
        todayStr,
        itemName,
        summaryObj[itemName].receive,
        summaryObj[itemName].issue,
        summaryObj[itemName].lastBalance
      ]);
    }
    summarySheet.getRange(summarySheet.getLastRow() + 1, 1, finalData.length, 5).setValues(finalData);
    return "✅ สรุปยอดเรียบร้อย";
  }
  return "ℹ️ ไม่มีรายการของวันนี้";
}

function addNewItem(sheetName, id, name, location, qty, manufactureDate, regNum, unit) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Error: ไม่พบชีท " + sheetName;
  var data = sheet.getDataRange().getValues();
  var C = detectCols_(data);
  var cleanNewId = id.replace(/-/g, "").trim();
  for (var i = 1; i < data.length; i++) {
    if (data[i][C.id] && data[i][C.id].toString().replace(/-/g, "").trim() === cleanNewId) {
      return "Error: หมายเลขพัสดุ " + id + " มีอยู่แล้วในระบบ (ตรงกับ " + data[i][C.id] + ")";
    }
  }
  var lastSeq = 0;
  for (var j = 1; j < data.length; j++) {
    var val = parseFloat(data[j][0]);
    if (!isNaN(val) && val > lastSeq) lastSeq = val;
  }
  var nextSeq = lastSeq + 1;
  var lastDataRow = 1;
  var foundData = false;
  for (var k = 1; k < data.length; k++) {
    var colB = data[k][C.id] ? data[k][C.id].toString().trim() : "";
    if (colB !== "") {
      foundData = true;       
      lastDataRow = k + 1;     
    } else if (foundData) {
      break;              
    }
  }
  sheet.insertRowAfter(lastDataRow);
  var newRow = lastDataRow + 1;
  sheet.getRange(newRow, 1).setValue(nextSeq);
  sheet.getRange(newRow, 2).setNumberFormat("@");
  sheet.getRange(newRow, 2).setValue(id);
  sheet.getRange(newRow, 3).setValue(name);
  sheet.getRange(newRow, 3).setHorizontalAlignment("left");
  if (C.unit >= 0) sheet.getRange(newRow, C.unit + 1).setValue(unit ? unit.toString().trim() : "");
  sheet.getRange(newRow, C.loc + 1).setValue(location);
  sheet.getRange(newRow, C.qty + 1).setValue(qty);
  var scriptUrl = CONFIG.WEB_APP_URL;
  sheet.getRange(newRow, C.qr + 1).setFormula(
    '=IMAGE("https://quickchart.io/qr?text=" & ENCODEURL("' + scriptUrl + '?id=" & B' + newRow + ') & "&size=1000")'
  );
  sheet.setRowHeight(newRow, 140);
  var borderRange = sheet.getRange(newRow, 1, 1, C.qr + 1);
  borderRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  if (sheetName === CONFIG.TRANSPORT_SHEET && manufactureDate) {
    var mdParts = manufactureDate.split('-');
    if (mdParts.length === 3) {
      var mdDate = new Date(parseInt(mdParts[0]), parseInt(mdParts[1]) - 1, parseInt(mdParts[2]));
      sheet.getRange(newRow, C.mfg + 1).setValue(mdDate);
      sheet.getRange(newRow, C.mfg + 1).setNumberFormat('dd/MM/yyyy');
    }
  }
  if (sheetName === CONFIG.TRANSPORT_SHEET && regNum) {
    sheet.getRange(newRow, C.reg + 1).setValue(regNum.toString().trim());
  }
  var histSheet = ss.getSheetByName(CONFIG.HISTORY_SHEET);
  if (histSheet) {
    var lastHistRow = histSheet.getLastRow();
    var nextHistSeq = 1;
    if (lastHistRow > 1) {
      for (var r = lastHistRow; r >= 2; r--) {
        var hval = parseInt(histSheet.getRange(r, 1).getValue());
        if (!isNaN(hval) && hval > 0) {
          nextHistSeq = hval + 1;
          break;
        }
      }
    }
    histSheet.appendRow([
      nextHistSeq, new Date(), "", "รับเข้า",
      name, "รับพัสดุใหม่เข้าคลัง", qty, qty
    ]);
  }
  return "ok";
}

function getDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = CONFIG.INVENTORY_SHEETS;
  var LOW_STOCK_THRESHOLD = CONFIG.LOW_STOCK_THRESHOLD;
  var totalItems = 0;
  var lowStockItems = [];
  for (var s = 0; s < sheetNames.length; s++) {
    var sheet = ss.getSheetByName(sheetNames[s]);
    if (!sheet) continue;
    var data = sheet.getDataRange().getValues();
    var C = detectCols_(data);
    for (var i = 1; i < data.length; i++) {
      if (!data[i][C.id]) continue;
      var rawName = data[i][C.name] ? data[i][C.name].toString().trim() : "";
      if (rawName === "ชื่อพัสดุ" || rawName === "") continue;
      totalItems++;
      var qty = parseFloat(data[i][C.qty]) || 0;
      if (qty <= LOW_STOCK_THRESHOLD) {
        lowStockItems.push({
          id: data[i][C.id].toString().trim(),
          name: rawName,
          qty: qty,
          sheetName: sheetNames[s]
        });
      }
    }
  }
  var todayReceiveItems = {};  
  var todayIssueItems  = {};
  var histSheet = ss.getSheetByName(CONFIG.HISTORY_SHEET);
  if (histSheet) {
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var histData = histSheet.getDataRange().getValues();
    for (var j = 1; j < histData.length; j++) {
      var dateIn  = histData[j][1];
      var dateOut = histData[j][2];
      var dateVal = (dateIn instanceof Date && dateIn !== "") ? dateIn : dateOut;
      if (!(dateVal instanceof Date)) continue;
      var recDate = new Date(dateVal);
      recDate.setHours(0, 0, 0, 0);
      if (recDate.getTime() !== today.getTime()) continue;
      var type     = histData[j][3];
      var itemName = histData[j][4] ? histData[j][4].toString() : "";
      if (type === "รับเข้า")  todayReceiveItems[itemName] = true;
      if (type === "จ่ายออก") todayIssueItems[itemName]  = true;
    }
  }
  return {
    totalItems:    totalItems,
    todayReceive:  Object.keys(todayReceiveItems).length,
    todayIssue:    Object.keys(todayIssueItems).length,
    lowStockItems: lowStockItems,
    threshold:     LOW_STOCK_THRESHOLD
  };
}

function getTotalItemsList(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];
  var sheetNames = sheetName ? [sheetName] : CONFIG.INVENTORY_SHEETS;
  for (var s = 0; s < sheetNames.length; s++) {
    var sheet = ss.getSheetByName(sheetNames[s]);
    if (!sheet) continue;
    var data = sheet.getDataRange().getValues();
    var C = detectCols_(data);
    for (var i = 1; i < data.length; i++) {
      if (!data[i][C.id]) continue;
      var rawName = data[i][C.name] ? data[i][C.name].toString().trim() : "";
      if (rawName === "ชื่อพัสดุ" || rawName === "") continue;
      results.push({
        id:        data[i][C.id].toString().trim(),
        name:      rawName,
        location:  data[i][C.loc] ? data[i][C.loc].toString().trim() : "-",
        unit:      (C.unit >= 0 && data[i][C.unit]) ? data[i][C.unit].toString().trim() : "",
        qty:       parseFloat(data[i][C.qty]) || 0,
        sheetName: sheetNames[s]
      });
    }
  }
  return results;
}

function getHistoryByDate(dateStr, type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var histSheet = ss.getSheetByName(CONFIG.HISTORY_SHEET);
  if (!histSheet) return [];
  var data = histSheet.getDataRange().getValues();
  var results = [];
  var targetDate = new Date(dateStr);
  targetDate.setHours(0, 0, 0, 0);
  var nameToId  = {};
  var nameToQty = {};
  var invSheets = CONFIG.INVENTORY_SHEETS;
  for (var s = 0; s < invSheets.length; s++) {
    var invSheet = ss.getSheetByName(invSheets[s]);
    if (!invSheet) continue;
    var invData = invSheet.getDataRange().getValues();
    var IC = detectCols_(invData);
    for (var ii = 1; ii < invData.length; ii++) {
      var iName = invData[ii][IC.name] ? invData[ii][IC.name].toString().trim() : "";
      var iId   = invData[ii][IC.id] ? invData[ii][IC.id].toString().trim() : "";
      var iQty  = parseFloat(invData[ii][IC.qty]) || 0;
      if (iName && iId) {
        nameToId[iName]  = iId;
        nameToQty[iName] = iQty;
      }
    }
  }
  for (var i = 1; i < data.length; i++) {
    var dateIn  = data[i][1];
    var dateOut = data[i][2];
    var dateVal = (dateIn instanceof Date && dateIn !== "") ? dateIn : dateOut;
    if (!(dateVal instanceof Date)) continue;
    var recDate = new Date(dateVal);
    recDate.setHours(0, 0, 0, 0);
    if (recDate.getTime() !== targetDate.getTime()) continue;
    var rowType = data[i][3] ? data[i][3].toString() : "";
    if (type && rowType !== type) continue;
    var rName = data[i][4] ? data[i][4].toString() : "";
    var histBalance = parseFloat(data[i][7]) || 0;
    var balance = (histBalance > 0) ? histBalance : (nameToQty[rName] !== undefined ? nameToQty[rName] : 0);
    results.push({
      seq:     data[i][0],
      dateIn:  dateIn instanceof Date ? Utilities.formatDate(dateIn, "GMT+7", "dd/MM/yyyy HH:mm") : "",
      dateOut: dateOut instanceof Date ? Utilities.formatDate(dateOut, "GMT+7", "dd/MM/yyyy HH:mm") : "",
      type:    rowType,
      name:    rName,
      id:      nameToId[rName] || "",
      unit:    data[i][5] ? data[i][5].toString() : "",
      amount:  parseFloat(data[i][6]) || 0,
      balance: balance 
    });
  }
  return results;
}

function getMonthActivity(year, month, type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var histSheet = ss.getSheetByName(CONFIG.HISTORY_SHEET);
  if (!histSheet) return [];
  var data = histSheet.getDataRange().getValues();
  var activeDays = {};
  for (var i = 1; i < data.length; i++) {
    var dateIn  = data[i][1];
    var dateOut = data[i][2];
    var dateVal = (dateIn instanceof Date && dateIn !== "") ? dateIn : dateOut;
    if (!(dateVal instanceof Date)) continue;
    var recDate = new Date(dateVal);
    if (recDate.getFullYear() !== year) continue;
    if (recDate.getMonth() !== month) continue;
    var rowType = data[i][3] ? data[i][3].toString() : "";
    if (type && rowType !== type) continue;
    activeDays[recDate.getDate()] = true;
  }
  return Object.keys(activeDays).map(function(d) { return parseInt(d); });
}

function fixHistoryBalance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var histSheet = ss.getSheetByName(CONFIG.HISTORY_SHEET);
  if (!histSheet) return "ไม่พบชีทประวัติรับจ่าย";
  var invSheets = CONFIG.INVENTORY_SHEETS;
  var currentQty = {}; 
  for (var s = 0; s < invSheets.length; s++) {
    var invSheet = ss.getSheetByName(invSheets[s]);
    if (!invSheet) continue;
    var invData = invSheet.getDataRange().getValues();
    var IC = detectCols_(invData);
    for (var ii = 1; ii < invData.length; ii++) {
      var iName = invData[ii][IC.name] ? invData[ii][IC.name].toString().trim() : "";
      var iQty  = parseFloat(invData[ii][IC.qty]) || 0;
      if (iName && iName !== "ชื่อพัสดุ") currentQty[iName] = iQty;
    }
  }
  var histData = histSheet.getDataRange().getValues();
  var lastRowOfItem = {}; 
  for (var r = 1; r < histData.length; r++) {
    var name = histData[r][4] ? histData[r][4].toString().trim() : "";
    if (!name) continue;
    lastRowOfItem[name] = r + 1; 
  }
  var fixCount = 0;
  for (var itemName in lastRowOfItem) {
    if (currentQty[itemName] !== undefined) {
      var targetRow = lastRowOfItem[itemName];
      var existingBalance = histSheet.getRange(targetRow, 8).getValue();
      if (!existingBalance || existingBalance === 0) {
        histSheet.getRange(targetRow, 8).setValue(currentQty[itemName]);
        fixCount++;
        Logger.log("แก้ไข: " + itemName + " แถว " + targetRow + " → " + currentQty[itemName]);
      }
    }
  }
  return "✅ แก้ไขยอดคงเหลือสำเร็จ " + fixCount + " รายการ";
}

function resequenceHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var histSheet = ss.getSheetByName(CONFIG.HISTORY_SHEET);
  if (!histSheet) { Logger.log("ไม่พบชีท"); return; }
  var lastRow = histSheet.getLastRow();
  var seq = 1;
  for (var r = 3; r <= lastRow; r++) { 
    var colB = histSheet.getRange(r, 2).getValue();
    var colC = histSheet.getRange(r, 3).getValue();
    var colE = histSheet.getRange(r, 5).getValue();
    if (colE && colE.toString().trim() !== "") {
      histSheet.getRange(r, 1).setValue(seq);
      seq++;
    } else {
      histSheet.getRange(r, 1).setValue(""); 
    }
  }
  Logger.log("เรียงลำดับใหม่เสร็จ รวม " + (seq - 1) + " รายการ");
}

function fixAllHistoryBalance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var histSheet = ss.getSheetByName(CONFIG.HISTORY_SHEET);
  if (!histSheet) return "ไม่พบชีท";
  var currentQty = {};
  var invSheets = CONFIG.INVENTORY_SHEETS;
  for (var s = 0; s < invSheets.length; s++) {
    var invSheet = ss.getSheetByName(invSheets[s]);
    if (!invSheet) continue;
    var invData = invSheet.getDataRange().getValues();
    var IC = detectCols_(invData);
    for (var i = 1; i < invData.length; i++) {
      var name = invData[i][IC.name] ? invData[i][IC.name].toString().trim() : "";
      var qty  = parseFloat(invData[i][IC.qty]) || 0;
      if (name && name !== "ชื่อพัสดุ") currentQty[name] = qty;
    }
  }
  var allData = histSheet.getDataRange().getValues();
  var lastRowMap = {}; 
  for (var r = 1; r < allData.length; r++) {
    var hName = allData[r][4] ? allData[r][4].toString().trim() : "";
    if (hName) lastRowMap[hName] = r + 1;
  }
  var lastRow = histSheet.getLastRow();
  if (lastRow >= 3) {
    histSheet.getRange(3, 8, lastRow - 2, 1).clearContent();
  }
  var count = 0;
  for (var itemName in lastRowMap) {
    if (currentQty[itemName] !== undefined) {
      histSheet.getRange(lastRowMap[itemName], 8).setValue(currentQty[itemName]);
      count++;
    }
  }
  return "✅ อัปเดตเรียบร้อย " + count + " รายการ";
}

function installOnEditTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "onEditInstallable") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("onEditInstallable")
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  Logger.log("✅ ติดตั้ง trigger เรียบร้อย");
}

function onEditInstallable(e) {
  if (!e) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.HISTORY_SHEET) return;
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(3000)) return; 
    fixAllHistoryBalance();
  } finally {
    lock.releaseLock();
  }
}

function getDailySummaryByRange(dateFromStr, dateToStr) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var histSheet = ss.getSheetByName(CONFIG.HISTORY_SHEET);
  if (!histSheet) return [];
  var actualQty = {};
  var invSheets = CONFIG.INVENTORY_SHEETS;
  for (var s = 0; s < invSheets.length; s++) {
    var invSheet = ss.getSheetByName(invSheets[s]);
    if (!invSheet) continue;
    var invData = invSheet.getDataRange().getValues();
    var IC = detectCols_(invData);
    for (var ii = 1; ii < invData.length; ii++) {
      var iName = invData[ii][IC.name] ? invData[ii][IC.name].toString().trim() : "";
      if (iName && iName !== "ชื่อพัสดุ") {
        actualQty[iName] = parseFloat(invData[ii][IC.qty]) || 0;
      }
    }
  }
  var histData = histSheet.getDataRange().getValues();
  var from = new Date(dateFromStr); from.setHours(0, 0, 0, 0);
  var to   = new Date(dateToStr);   to.setHours(23, 59, 59, 999);
  var dateItemMap = {};
  var dateOrder   = [];
  for (var i = 1; i < histData.length; i++) {
    var dateIn  = histData[i][1];
    var dateOut = histData[i][2];
    var dateVal = (dateIn instanceof Date && dateIn !== "") ? dateIn : dateOut;
    if (!(dateVal instanceof Date)) continue;
    if (dateVal < from || dateVal > to) continue;
    var rowType = histData[i][3] ? histData[i][3].toString() : "";
    if (!rowType || rowType === "แก้ไขข้อมูล") continue;
    var name   = histData[i][4] ? histData[i][4].toString().trim() : "";
    var amount = parseFloat(histData[i][6]) || 0;
    if (!name) continue;
    var dateStr = Utilities.formatDate(new Date(dateVal), "GMT+7", "dd/MM/yyyy");
    var key     = dateStr + "|" + name;
    if (!dateItemMap[key]) {
      dateItemMap[key] = { dateStr: dateStr, name: name, receive: 0, issue: 0 };
      dateOrder.push(key);
    }
    if (rowType === "รับเข้า")  dateItemMap[key].receive += amount;
    if (rowType === "จ่ายออก") dateItemMap[key].issue   += amount;
  }
  var results    = [];
  var lastIdxMap = {};
  for (var di = 0; di < dateOrder.length; di++) {
    var entry = dateItemMap[dateOrder[di]];
    results.push({
      dateStr: entry.dateStr,
      name:    entry.name,
      receive: entry.receive,
      issue:   entry.issue,
      balance: null
    });
    lastIdxMap[entry.name] = results.length - 1;
  }
  for (var itemName in lastIdxMap) {
    var idx = lastIdxMap[itemName];
    if (actualQty[itemName] !== undefined) {
      results[idx].balance = actualQty[itemName];
    }
  }
  return results;
}

function getHistoryByRange(dateFromStr, dateToStr) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var histSheet = ss.getSheetByName(CONFIG.HISTORY_SHEET);
  if (!histSheet) return [];
  var data = histSheet.getDataRange().getValues();
  var from = new Date(dateFromStr); from.setHours(0, 0, 0, 0);
  var to   = new Date(dateToStr);   to.setHours(23, 59, 59, 999);
  var actualQty = {};
  var invSheets = CONFIG.INVENTORY_SHEETS;
  for (var s = 0; s < invSheets.length; s++) {
    var invSheet = ss.getSheetByName(invSheets[s]);
    if (!invSheet) continue;
    var invData = invSheet.getDataRange().getValues();
    var IC = detectCols_(invData);
    for (var ii = 1; ii < invData.length; ii++) {
      var iName = invData[ii][IC.name] ? invData[ii][IC.name].toString().trim() : "";
      if (iName && iName !== "ชื่อพัสดุ") {
        actualQty[iName] = parseFloat(invData[ii][IC.qty]) || 0;
      }
    }
  }
  var results = [];
  for (var i = 1; i < data.length; i++) {
    var dateIn  = data[i][1];
    var dateOut = data[i][2];
    var dateVal = (dateIn instanceof Date && dateIn !== "") ? dateIn : dateOut;
    if (!(dateVal instanceof Date)) continue;
    if (dateVal < from || dateVal > to) continue;
    var rowType = data[i][3] ? data[i][3].toString() : "";
    if (!rowType || rowType === "แก้ไขข้อมูล") continue;
    var rawBal  = data[i][7];
    var balance = (rawBal !== "" && rawBal !== null && !isNaN(parseFloat(rawBal)))
                  ? parseFloat(rawBal) : null;
    results.push({
      dateIn:  dateIn  instanceof Date ? Utilities.formatDate(dateIn,  "GMT+7", "dd/MM/yyyy HH:mm") : "",
      dateOut: dateOut instanceof Date ? Utilities.formatDate(dateOut, "GMT+7", "dd/MM/yyyy HH:mm") : "",
      type:    rowType,
      name:    data[i][4] ? data[i][4].toString() : "",
      unit:    data[i][5] ? data[i][5].toString() : "",
      amount:  parseFloat(data[i][6]) || 0,
      balance: balance
    });
  }
  var lastIdxMap = {};
  for (var ri = 0; ri < results.length; ri++) {
    lastIdxMap[results[ri].name] = ri;
  }
  for (var itemName in lastIdxMap) {
    var idx = lastIdxMap[itemName];
    if ((results[idx].balance === null || results[idx].balance === 0)
        && actualQty[itemName] !== undefined) {
      results[idx].balance = actualQty[itemName];
    }
  }
  return results;
}

function rebuildAllQRCodes() {
  var scriptUrl = CONFIG.WEB_APP_URL;
  var sheetNames = CONFIG.INVENTORY_SHEETS;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  for (var s = 0; s < sheetNames.length; s++) {
    var sheet = ss.getSheetByName(sheetNames[s]);
    if (!sheet) continue;
    var data = sheet.getDataRange().getValues();
    var C = detectCols_(data);
    for (var i = 2; i < data.length; i++) {
      var id = data[i][C.id];
      if (!id) continue;
      sheet.getRange(i + 1, C.qr + 1).setFormula(
        '=IMAGE("https://quickchart.io/qr?text=" & ENCODEURL("' + scriptUrl + '?id=" & B' + (i + 1) + ') & "&size=500")'
      );
    }
    SpreadsheetApp.flush();
  }
  return "✅ อัปเดต QR Code ทุกแถวเรียบร้อย";
}

function debugFindId() {
  var searchId = "7045-35-851-2079";  // [CHANGE-ตามต้องการ] หมายเลขพัสดุตัวอย่างสำหรับดีบั๊ก
  var sheetNames = CONFIG.INVENTORY_SHEETS;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];
  for (var s = 0; s < sheetNames.length; s++) {
    var sheet = ss.getSheetByName(sheetNames[s]);
    if (!sheet) continue;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (!data[i][1]) continue;
      var rawId = data[i][1];
      log.push(sheetNames[s] + " | row " + (i+1) + " | type:" + typeof rawId + " | value:[" + rawId + "]");
    }
  }
  Logger.log(log.join("\n"));
}

function fixColumnBFormat() {
  var sheetNames = CONFIG.INVENTORY_SHEETS;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var totalFixed = 0;
  for (var s = 0; s < sheetNames.length; s++) {
    var sheet = ss.getSheetByName(sheetNames[s]);
    if (!sheet) continue;
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) continue;
    var dataRange = sheet.getRange(3, 2, lastRow - 2, 1);
    dataRange.setNumberFormat("@");
    var values = dataRange.getValues();
    var newValues = [];
    var hasChanges = false;
    for (var i = 0; i < values.length; i++) {
      var raw = values[i][0];
      var newVal = raw;
      if (typeof raw === 'number') {
        var numStr = Math.round(raw).toString();
        var formatted = numStr; 
        if (numStr.length === 13) {
          formatted = numStr.substring(0,4) + "-" + numStr.substring(4,6) + "-" + numStr.substring(6,9) + "-" + numStr.substring(9,13);
        } else if (numStr.length === 12) {
          formatted = numStr.substring(0,4) + "-" + numStr.substring(4,6) + "-" + numStr.substring(6,9) + "-" + numStr.substring(9,12);
        } else if (numStr.length === 11) {
          formatted = numStr.substring(0,4) + "-" + numStr.substring(4,6) + "-" + numStr.substring(6,7) + "-" + numStr.substring(7,11);
        } else {
          Logger.log("ไม่รู้รูปแบบ: " + numStr + " length=" + numStr.length + " row=" + (i+3));
        }
        newVal = formatted;
        totalFixed++;
        hasChanges = true;
        Logger.log("แก้ row " + (i+3) + ": " + raw + " → " + formatted);
      } else if (typeof raw === 'string') {
        newVal = raw.replace(/^'+/, "");
      }
      newValues.push([newVal]);
    }
    if (hasChanges) {
      dataRange.setNumberFormat("@");
      dataRange.setValues(newValues);
      SpreadsheetApp.flush();
    }
  }
  rebuildAllQRCodes();
  return "✅ แก้ไขสำเร็จ " + totalFixed + " รายการ";
}


// ============================================================
//  รายงานแบบ "บัญชีตรวจสอบพัสดุ (ผนวก ค)" / Export Excel
// ============================================================

function getAnnexItems() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = CONFIG.INVENTORY_SHEETS;
  var results = [];
  for (var s = 0; s < sheetNames.length; s++) {
    var sheet = ss.getSheetByName(sheetNames[s]);
    if (!sheet) continue;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    var C = sheetCols_(sheet);
    // อ่านเฉพาะคอลัมน์ B..จำนวนคงเหลือ (ไม่อ่านคอลัมน์ QR) เพื่อให้เร็ว
    var body = sheet.getRange(2, 2, lastRow - 1, C.qty).getValues();   // body[i][k] = คอลัมน์ index k+1
    for (var i = 0; i < body.length; i++) {
      var rawId = body[i][C.id - 1];
      if (!rawId) continue;
      var rawName = body[i][C.name - 1] ? body[i][C.name - 1].toString().trim() : "";
      if (rawName === "ชื่อพัสดุ" || rawName === "") continue;
      var qtyNum = parseFloat(body[i][C.qty - 1]) || 0;
      if (qtyNum <= 0) continue;   // ยอดเป็น 0 ไม่แสดงในรายงานผนวก ค
      var unitVal = (C.unit >= 0 && body[i][C.unit - 1]) ? body[i][C.unit - 1].toString().trim() : "";
      results.push({
        id:        rawId.toString().trim().replace(/^'+|'+$/g, ""),
        name:      rawName,
        unit:      unitVal,
        qty:       qtyNum,
        sheetName: sheetNames[s]
      });
    }
  }
  return results;
}
