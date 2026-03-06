/**
 * Google Apps Script สำหรับ CSD1 Report
 * รองรับ Login ด้วย Google Account (ตรวจสอบอีเมลจากชีท "User_master")
 *
 * วิธีตั้งค่า:
 * 1. สร้าง Google Sheet ใหม่
 * 2. Extensions > Apps Script > วาง code นี้
 * 3. รัน initSheet() เพื่อสร้างหัวตาราง + ชีทผู้ใช้
 * 4. ไปชีท "User_master" เพิ่มอีเมลเจ้าหน้าที่ที่อนุญาต
 * 5. Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Copy URL ไปใส่ในตั้งค่าระบบ (หน้า Login)
 */

// ===== รูปโปรไฟล์ =====
// สร้าง folder ใน Google Drive → Share → Anyone with link → Copy folder ID จาก URL
// URL ตัวอย่าง: https://drive.google.com/drive/folders/ABC123  →  folder ID = ABC123
// ตั้งชื่อรูป: อีเมล.jpg (เช่น example@gmail.com.jpg)
const AVATAR_FOLDER_ID = "1rZtEqOvl9soj_gBR3jcHoL2dB-aNAcOH";

// ===== หัวตาราง (ชีทรายงาน) =====
const HEADERS = [
  "ลำดับ",
  "ประทับเวลา",
  "ผู้รายงาน",
  "อีเมล",
  "ชุดปฏิบัติการ",
  "หัวหน้าชุด",
  "หมายเลขโทรศัพท์",
  "วันที่",
  "เวลา",
  "สถานที่ (จับกุม/หรือตรวจสอบ)",
  "พิกัด Google Maps",
  "การจับกุม ตรวจสอบ",
  "ศาลที่ออกหมายจับ",
  "เลขที่หมายจับ",
  "วันเดือนปีที่ออกหมาย",
  "ประเภทหมายจับ",
  "ชื่อ-สกุล ผู้ต้องหา",
  "อายุ (ปี)",
  "เลขประจำตัวประชาชน (13 หลัก)",
  "ที่อยู่ (ตามบัตร)",
  "ข้อหา (ฐานความผิด)",
  "พร้อมของกลาง (ถ้ามี)",
  "นำส่งพนักงานสอบสวน",
  "ชื่อ-สกุล เจ้าของสถานที่ / พยาน นำตรวจ",
  "อายุ (ปี)",
  "เลขประจำตัวประชาชน (13 หลัก)",
  "ที่อยู่ (ตามบัตร)",
  "ผลการตรวจสอบ",
  "ตรวจค้น จับกุมโดย",
  "ศาลที่ออกหมายค้น",
  "เลขที่หมายค้น",
  "วันเดือนปีที่ออกหมายค้น",
  "ประเภทหมายค้น",
  "หน้างานต่างด้าว",
  "หน้างานห้วงระดม",
  "หน้างานศูนย์",
  "งานออกสื่อ"
];

// ===== หัวตาราง (ชีทผู้ใช้) =====
const USER_HEADERS = ["อีเมล", "ชื่อ-สกุล", "ตำแหน่ง", "ชุดปฏิบัติการ", "เบอร์โทร", "สิทธิ์"];

/**
 * สร้างหัวตารางและชีทผู้ใช้
 */
function initSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // === ชีทรายงาน ===
  let reportSheet = ss.getSheetByName("รายงานผลการปฏิบัติงาน");
  if (!reportSheet) {
    reportSheet = ss.getActiveSheet();
    reportSheet.setName("รายงานผลการปฏิบัติงาน");
  }
  const headerRange = reportSheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setValues([HEADERS]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#1a56db");
  headerRange.setFontColor("#ffffff");
  headerRange.setHorizontalAlignment("center");
  headerRange.setWrap(true);
  reportSheet.setFrozenRows(1);

  // === ชีทผู้ใช้ ===
  let userSheet = ss.getSheetByName("User_master");
  if (!userSheet) {
    userSheet = ss.insertSheet("User_master");
  }
  const userHeaderRange = userSheet.getRange(1, 1, 1, USER_HEADERS.length);
  userHeaderRange.setValues([USER_HEADERS]);
  userHeaderRange.setFontWeight("bold");
  userHeaderRange.setBackground("#059669");
  userHeaderRange.setFontColor("#ffffff");
  userHeaderRange.setHorizontalAlignment("center");
  userHeaderRange.setWrap(true);
  userSheet.setFrozenRows(1);

  userSheet.setColumnWidth(1, 220);
  userSheet.setColumnWidth(2, 250);
  userSheet.setColumnWidth(3, 200);
  userSheet.setColumnWidth(4, 120);
  userSheet.setColumnWidth(5, 140);
  userSheet.setColumnWidth(6, 80);

  if (userSheet.getLastRow() < 2) {
    userSheet.getRange(2, 1, 1, 6).setValues([
      ["example@gmail.com", "ผู้ดูแลระบบ", "admin", "", "", "admin"]
    ]);
  }

  SpreadsheetApp.getUi().alert(
    "สร้างระบบเรียบร้อย!\n\n" +
    "1. ไปชีท 'User_master'\n" +
    "2. ใส่อีเมล Google ของเจ้าหน้าที่ที่อนุญาต\n" +
    "3. Deploy > New deployment > Web app\n" +
    "4. นำ URL ไปตั้งค่าในหน้าเว็บ"
  );
}

/**
 * ค้นหารูปโปรไฟล์จาก Google Drive folder
 * ค้นหาตาม: email, username (ส่วนหน้า @), ชื่อ-สกุล
 * เช่น example@gmail.com.jpg, example.jpg, สมชาย ใจดี.jpg
 */
function getAvatarUrl(email, name) {
  try {
    if (!AVATAR_FOLDER_ID) return "";
    const folder = DriveApp.getFolderById(AVATAR_FOLDER_ID);
    const exts = [".jpg", ".jpeg", ".png", ".JPG", ".JPEG", ".PNG", ".webp"];

    // สร้าง keywords ที่จะค้นหา: email, username, ชื่อ
    var keywords = [email];
    if (email && email.indexOf("@") > -1) {
      keywords.push(email.split("@")[0].toLowerCase());
    }
    if (name) {
      keywords.push(name.trim());
    }

    // 1) ค้นหาแบบ exact filename
    for (var k = 0; k < keywords.length; k++) {
      for (var e = 0; e < exts.length; e++) {
        var files = folder.getFilesByName(keywords[k] + exts[e]);
        if (files.hasNext()) {
          var file = files.next();
          return "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w200";
        }
      }
    }

    // 2) fallback: loop ทุกไฟล์ เทียบ keyword
    var allFiles = folder.getFiles();
    while (allFiles.hasNext()) {
      var f = allFiles.next();
      var fname = f.getName().toLowerCase().replace(/\.[^.]+$/, ""); // ตัดนามสกุลออก
      for (var k2 = 0; k2 < keywords.length; k2++) {
        if (fname === keywords[k2].toLowerCase() || fname.indexOf(keywords[k2].toLowerCase()) > -1) {
          return "https://drive.google.com/thumbnail?id=" + f.getId() + "&sz=w200";
        }
      }
    }
    return "";
  } catch (err) {
    return "";
  }
}

/**
 * รับข้อมูล POST (login / submitReport)
 */
function doPost(e) {
  try {
    let data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return jsonResponse({ status: "error", message: "Invalid JSON" });
    }

    if (data.action === "login") {
      return handleLogin(data);
    } else if (data.action === "getData") {
      return handleGetData(data);
    } else {
      return handleSubmitReport(data);
    }
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

/**
 * ตรวจสอบอีเมลจากชีทผู้ใช้
 */
function handleLogin(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User_master");

  if (!sheet) {
    return jsonResponse({ status: "error", message: "ไม่พบชีท 'User_master' กรุณารัน initSheet() ก่อน" });
  }

  const email = String(data.email || "").trim().toLowerCase();
  if (!email) {
    return jsonResponse({ status: "error", message: "ไม่พบอีเมล" });
  }

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const rowEmail = String(rows[i][0]).trim().toLowerCase();
    if (rowEmail === email) {
      const userName = String(rows[i][1]).trim();
      return jsonResponse({
        status: "success",
        name: userName,
        position: String(rows[i][2]).trim(),
        unit: String(rows[i][3]).trim(),
        phone: String(rows[i][4]).trim(),
        role: String(rows[i][5]).trim() || "user",
        avatarUrl: getAvatarUrl(email, userName)
      });
    }
  }

  return jsonResponse({
    status: "error",
    message: "อีเมล " + email + " ไม่มีสิทธิ์เข้าใช้งาน"
  });
}

/**
 * บันทึกรายงาน
 */
function handleSubmitReport(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("รายงานผลการปฏิบัติงาน") || ss.getActiveSheet();

  const lastRow = sheet.getLastRow();
  const nextNumber = lastRow;

  const row = [
    nextNumber,
    data.timestamp || new Date().toLocaleString('th-TH'),
    data.reporter || "",
    data.email || "",
    data.unit || "",
    data.leaderName || "",
    data.leaderPhone || "",
    data.date || "",
    data.time || "",
    data.location || "",
    data.coordinates || "",
    data.actionType || "",
    data.arrestCourt || "",
    data.arrestNo || "",
    data.warrantDate || "",
    data.warrantType || "",
    data.personName || "",
    data.personAge || "",
    data.personID || "",
    data.personAddress || "",
    data.charges || "",
    data.evidence || "",
    data.investigator || "",
    data.ownerName || "",
    data.ownerAge || "",
    data.ownerID || "",
    data.ownerAddress || "",
    data.inspectionResult || "",
    data.searchBy || "",
    data.searchCourt || "",
    data.searchNo || "",
    data.searchDate || "",
    data.searchWarrantType || "",
    data.workAlien || "",
    data.workMobilize || "",
    data.workCenter || "",
    data.workMedia || ""
  ];

  sheet.appendRow(row);
  return jsonResponse({ status: "success", row: nextNumber });
}

/**
 * ดึงข้อมูลรายงานทั้งหมดสำหรับ Dashboard
 * ตรวจสิทธิ์ก่อน → อ่านชีท → ส่งกลับเป็น array of objects
 */
function handleGetData(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // อ่านข้อมูลรายงาน
  var sheet = ss.getSheetByName("รายงานผลการปฏิบัติงาน");
  if (!sheet || sheet.getLastRow() < 2) {
    return jsonResponse({ status: "success", data: [] });
  }

  var allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  var rows = [];
  for (var i = 0; i < allData.length; i++) {
    var obj = {};
    for (var j = 0; j < HEADERS.length; j++) {
      obj[HEADERS[j]] = allData[i][j] != null ? String(allData[i][j]) : "";
    }
    rows.push(obj);
  }

  return jsonResponse({ status: "success", data: rows });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return jsonResponse({ status: "ok", message: "CSD1 Report API v2.1 (Google Login)", version: "2.1" });
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CSD1 Report')
    .addItem('สร้างหัวตาราง + ชีทผู้ใช้', 'initSheet')
    .addToUi();
}
