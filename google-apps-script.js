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
  "วันที่บันทึก",
  "เวลาบันทึก",
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
const USER_HEADERS = ["อีเมล", "ชื่อ-สกุล", "ตำแหน่ง", "ชุดปฏิบัติการ", "เบอร์โทร", "สิทธิ์", "active"];
const PUBLIC_REPORT_HEADERS = [
  "วันที่",
  "ชุดปฏิบัติการ",
  "การจับกุม ตรวจสอบ",
  "ประเภทหมายจับ",
  "หน้างานต่างด้าว",
  "หน้างานห้วงระดม",
  "หน้างานศูนย์",
  "งานออกสื่อ"
];
const SESSION_CACHE_PREFIX = "CSD1_SESSION_";
const SESSION_TTL_SECONDS = 60 * 60 * 12; // 12 hours
const COMMANDER_INFO_SHEET_NAME = "CommanderInfo";
const COMMANDER_INFO_HEADERS = ["ลำดับ", "ประเภท", "ชื่อ-สกุล", "ตำแหน่ง", "ชุดปฏิบัติการ", "เบอร์โทร", "note"];
const CHANGE_REQUEST_SHEET_NAME = "คำขอแก้ไขรายการ";
const CONNECT_REQUEST_SHEET_NAME = "CSD1 Connect Requests";
const RECYCLE_BIN_SHEET_NAME = "Recycle Bin";
const MEETING_BOOKINGS_SHEET_NAME = "MeetingBookings";
const MEETING_BOOKINGS_HEADERS = ["id", "roomId", "date", "start", "end", "title", "organizer", "phone", "attendees", "note", "equipment", "createdAt", "createdBy"];
const RECYCLED_ROW_COLOR = "#242928";
const EDIT_CHANGED_CELL_COLOR = "#03FF00";
const CONNECT_DOC_NUMBER_PROPERTY_KEY = "CSD1_CONNECT_DOC_NUMBER_NEXT";
const CONNECT_SPREADSHEET_ID_PROPERTY_KEY = "CSD1_CONNECT_SPREADSHEET_ID";
const CONNECT_BACKEND_VERSION = "2026-03-15-connect-sheet-binding-v2";
const CONNECT_SHEET_CELL_CHAR_LIMIT = 45000;
const CONNECT_DOC_NUMBER_MIN = 1;
const CONNECT_DOC_NUMBER_MAX = 9999;
const CONNECT_REQUEST_HEADERS = [
  "request_id",
  "submitted_at_iso",
  "submitted_date_th",
  "submitted_time_th",
  "record_updated_at_iso",
  "submitted_by_email",
  "submitted_by_name",
  "submitted_by_role",
  "submitted_by_position",
  "submitted_by_unit",
  "submitted_by_phone",
  "request_type",
  "request_type_label",
  "unit",
  "commander",
  "doc_number",
  "doc_number_line",
  "case_preset_key",
  "case_detail",
  "case_detail_custom",
  "case_search_term",
  "admin_note_tag",
  "admin_note_tag_label",
  "admin_note_text",
  "request_history_status",
  "request_history_label",
  "status",
  "status_note",
  "phone_network",
  "ais_sub_type",
  "phone_sub_type",
  "phone_imei",
  "phone_numbers_text",
  "phone_numbers_json",
  "ip_entries_text",
  "ip_entries_json",
  "phone_date_start",
  "phone_date_end",
  "bank_code",
  "bank_sub_type",
  "statement_acc_type",
  "bank_accounts_text",
  "bank_accounts_json",
  "bank_promptpay",
  "xxx_amount",
  "xxx_date",
  "xxx_time",
  "bank_account_name",
  "atm_account_no",
  "atm_date",
  "atm_time",
  "atm_location",
  "atm_terminal_id",
  "bank_date_start",
  "bank_date_end",
  "true_id",
  "true_name",
  "true_date_start",
  "true_date_end",
  "tax_type",
  "tax_id",
  "tax_name",
  "tax_year_start",
  "tax_year_end",
  "file_base_name",
  "pdf_file_name",
  "generated_mode",
  "generated_pdf_url",
  "generated_docx_url",
  "generated_warnings_json",
  "raw_state_json",
  "payload_json",
  "status_updated_at_iso",
  "status_updated_by_email",
  "status_updated_by_name",
  "status_updated_by_role",
  "drive_link",
  "file_password"
];
const CHANGE_REQUEST_HEADERS = [
  "requestId",
  "createdAt",
  "requestType",
  "status",
  "requestedByEmail",
  "requestedByName",
  "requestedByUnit",
  "targetNo",
  "targetTimestamp",
  "targetReporter",
  "targetReporterEmail",
  "targetUnit",
  "targetDate",
  "targetTime",
  "targetLocation",
  "targetActionType",
  "targetSuspect",
  "targetWarrantNo",
  "targetCitizenId",
  "deleteReason",
  "note",
  "correctedData",
  "sourceRowJson"
];
const ALLOWED_GOOGLE_CLIENT_IDS = [
  "108806756839-iv4nrrfk4355ogcl2p2ehkh5f6a1u90b.apps.googleusercontent.com"
];
const ALLOW_LOCAL_TOKEN_FALLBACK = true; // ใช้ fallback ชั่วคราวเมื่อ tokeninfo endpoint ขัดข้อง

function formatThaiDateOnlyStandard(value) {
  const dt = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(dt.getTime())) return "";
  const months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
    "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
  const tz = "Asia/Bangkok";
  const day = Number(Utilities.formatDate(dt, tz, "d"));
  const month = Number(Utilities.formatDate(dt, tz, "M"));
  const year = Number(Utilities.formatDate(dt, tz, "yyyy"));
  if (!day || !month || month < 1 || month > 12 || !year) return Utilities.formatDate(dt, tz, "dd/MM/yyyy");
  return day + " " + months[month - 1] + " " + String((year + 543) % 100).padStart(2, "0");
}

function formatTimeOnlyStandard(value) {
  const dt = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(dt.getTime())) return "";
  const tz = "Asia/Bangkok";
  const hh = Utilities.formatDate(dt, tz, "HH");
  const mm = Utilities.formatDate(dt, tz, "mm");
  const ss = Utilities.formatDate(dt, tz, "ss");
  return hh + "." + mm + "." + ss;
}

function normalizeSheetBuddhistYear(yearValue) {
  var year = Number(String(yearValue == null ? "" : yearValue).trim());
  if (!isFinite(year) || year <= 0) return null;
  if (year < 100) year += 2500;
  while (year > 2800) year -= 543;
  if (year < 2400) year += 543;
  return year;
}

function normalizeSheetThaiDateValue(value) {
  var months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  var monthMap = {
    "มกราคม": 1, "ม.ค.": 1, "ม.ค": 1,
    "กุมภาพันธ์": 2, "ก.พ.": 2, "ก.พ": 2,
    "มีนาคม": 3, "มี.ค.": 3, "มี.ค": 3,
    "เมษายน": 4, "เม.ย.": 4, "เม.ย": 4,
    "พฤษภาคม": 5, "พ.ค.": 5, "พ.ค": 5,
    "มิถุนายน": 6, "มิ.ย.": 6, "มิ.ย": 6,
    "กรกฎาคม": 7, "ก.ค.": 7, "ก.ค": 7,
    "สิงหาคม": 8, "ส.ค.": 8, "ส.ค": 8,
    "กันยายน": 9, "ก.ย.": 9, "ก.ย": 9,
    "ตุลาคม": 10, "ต.ค.": 10, "ต.ค": 10,
    "พฤศจิกายน": 11, "พ.ย.": 11, "พ.ย": 11,
    "ธันวาคม": 12, "ธ.ค.": 12, "ธ.ค": 12
  };

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    var tz = "Asia/Bangkok";
    var dayFromDate = Number(Utilities.formatDate(value, tz, "d"));
    var monthFromDate = Number(Utilities.formatDate(value, tz, "M"));
    var yearFromDate = Number(Utilities.formatDate(value, tz, "yyyy")) + 543;
    if (dayFromDate && monthFromDate && yearFromDate) {
      return dayFromDate + " " + months[monthFromDate - 1] + " " + yearFromDate;
    }
  }

  var raw = String(value == null ? "" : value).trim();
  if (!raw) return "";

  function build(dayValue, monthValue, yearValue) {
    var day = Number(dayValue);
    var month = Number(monthValue);
    var yearBE = normalizeSheetBuddhistYear(yearValue);
    if (!day || !month || !yearBE || month < 1 || month > 12 || day < 1 || day > 31) return "";
    return day + " " + months[month - 1] + " " + yearBE;
  }

  var isoMatch = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (isoMatch) return build(isoMatch[3], isoMatch[2], isoMatch[1]);

  var slashMatch = raw.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (slashMatch) return build(slashMatch[1], slashMatch[2], slashMatch[3]);

  var thaiMatch = raw.match(/^(\d{1,2})\s+([^\s]+)\s+(\d{2,4})$/);
  if (thaiMatch) {
    var monthNo = monthMap[thaiMatch[2]] || null;
    return build(thaiMatch[1], monthNo, thaiMatch[3]);
  }

  var fallback = new Date(raw);
  if (!Number.isNaN(fallback.getTime())) {
    return build(fallback.getDate(), fallback.getMonth() + 1, fallback.getFullYear());
  }

  return "";
}

function normalizeSheetTimeValue(value, options) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return Utilities.formatDate(value, "Asia/Bangkok", "HH:mm");
  }

  const raw = String(value == null ? "" : value).trim();
  if (!raw) return "";

  const allowHourOnly = !options || options.allowHourOnly !== false;
  const cleaned = raw
    .replace(/[๐-๙]/g, function(digit) { return String("๐๑๒๓๔๕๖๗๘๙".indexOf(digit)); })
    .replace(/\s+/g, "")
    .replace(/น\.?$/i, "");

  if (!cleaned) return "";

  function buildTime(hhRaw, mmRaw) {
    const hh = Number(hhRaw);
    let mmText = String(mmRaw == null ? "" : mmRaw).trim();
    if (!isFinite(hh)) return "";
    if (!mmText) mmText = "00";
    if (mmText.length === 1) mmText += "0";
    if (mmText.length !== 2) return "";
    const mm = Number(mmText);
    if (!isFinite(mm) || hh < 0 || hh > 23 || mm < 0 || mm > 59) return "";
    return ("0" + hh).slice(-2) + ":" + ("0" + mm).slice(-2);
  }

  const separated = cleaned.match(/^(\d{1,2})[:.](\d{1,2})(?:[:.](\d{1,2}))?$/);
  if (separated) return buildTime(separated[1], separated[2]);

  const compact = cleaned.match(/^(\d{3,4})$/);
  if (compact) {
    const digits = compact[1];
    const hhRaw = digits.length === 3 ? digits.slice(0, 1) : digits.slice(0, 2);
    const mmRaw = digits.length === 3 ? digits.slice(1) : digits.slice(2);
    return buildTime(hhRaw, mmRaw);
  }

  if (allowHourOnly && /^\d{1,2}$/.test(cleaned)) {
    return buildTime(cleaned, "00");
  }

  return "";
}

function formatThaiDateTimeStandard(value) {
  const dt = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(dt.getTime())) return "";
  const months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
    "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
  const tz = "Asia/Bangkok";
  const day = Number(Utilities.formatDate(dt, tz, "d"));
  const month = Number(Utilities.formatDate(dt, tz, "M"));
  const year = Number(Utilities.formatDate(dt, tz, "yyyy"));
  const hh = Utilities.formatDate(dt, tz, "HH");
  const mm = Utilities.formatDate(dt, tz, "mm");
  if (!day || !month || month < 1 || month > 12 || !year) return Utilities.formatDate(dt, tz, "dd/MM/yyyy HH:mm");
  const yearBE = year + 543;
  return day + " " + months[month - 1] + " " + String(yearBE % 100).padStart(2, "0") + " เวลา " + hh + "." + mm + " น.";
}

function setConnectSpreadsheetBinding(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return "";
  const spreadsheetId = normalizeConnectString(ss.getId());
  if (!spreadsheetId) return "";
  PropertiesService.getScriptProperties().setProperty(CONNECT_SPREADSHEET_ID_PROPERTY_KEY, spreadsheetId);
  return spreadsheetId;
}

function getConnectSpreadsheet() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const configuredSpreadsheetId = normalizeConnectString(
    scriptProperties.getProperty(CONNECT_SPREADSHEET_ID_PROPERTY_KEY)
  );

  if (configuredSpreadsheetId) {
    try {
      return SpreadsheetApp.openById(configuredSpreadsheetId);
    } catch (error) {
      Logger.log("getConnectSpreadsheet: openById failed for configured spreadsheetId=" + configuredSpreadsheetId + " error=" + error);
    }
  }

  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (activeSpreadsheet) {
    setConnectSpreadsheetBinding(activeSpreadsheet);
    return activeSpreadsheet;
  }

  throw new Error("ไม่พบ Spreadsheet สำหรับระบบ CSD1 Connect กรุณารัน setConnectSpreadsheetBindingToActive() หรือ initConnectRequestSheet() จากไฟล์ Spreadsheet หลักก่อน");
}

function setConnectSpreadsheetBindingToActive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    const result = {
      ok: false,
      message: "ไม่พบ Spreadsheet ที่เปิดอยู่สำหรับผูกระบบ CSD1 Connect",
      spreadsheetId: "",
      spreadsheetName: ""
    };
    Logger.log("setConnectSpreadsheetBindingToActive => " + JSON.stringify(result));
    return result;
  }
  const spreadsheetId = setConnectSpreadsheetBinding(ss);
  const result = {
    ok: true,
    message: "ผูกระบบ CSD1 Connect กับ Spreadsheet เรียบร้อย",
    spreadsheetId: spreadsheetId,
    spreadsheetName: normalizeConnectString(ss.getName())
  };
  Logger.log("setConnectSpreadsheetBindingToActive => " + JSON.stringify(result));
  return result;
}

function getConnectRequestSheet() {
  return getConnectSpreadsheet().getSheetByName(CONNECT_REQUEST_SHEET_NAME);
}

function ensureConnectRequestSheet(sheet) {
  const ss = getConnectSpreadsheet();
  const targetSheet = sheet || ss.getSheetByName(CONNECT_REQUEST_SHEET_NAME) || ss.insertSheet(CONNECT_REQUEST_SHEET_NAME);
  const requiredColumns = CONNECT_REQUEST_HEADERS.length;
  const currentColumns = targetSheet.getMaxColumns();
  if (currentColumns < requiredColumns) {
    targetSheet.insertColumnsAfter(currentColumns, requiredColumns - currentColumns);
  }
  const headerRange = targetSheet.getRange(1, 1, 1, CONNECT_REQUEST_HEADERS.length);
  headerRange.setValues([CONNECT_REQUEST_HEADERS]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#5b21b6");
  headerRange.setFontColor("#ffffff");
  headerRange.setHorizontalAlignment("center");
  headerRange.setWrap(true);
  targetSheet.setFrozenRows(1);

  var widthMap = {
    1: 210,
    2: 180,
    3: 110,
    4: 110,
    5: 180,
    6: 220,
    7: 180,
    8: 110,
    9: 170,
    10: 120,
    11: 120,
    12: 110,
    13: 130,
    14: 110,
    15: 170,
    16: 90,
    17: 150,
    18: 130,
    19: 360,
    20: 220,
    21: 220,
    22: 120,
    23: 140,
    24: 280,
    25: 130,
    26: 130,
    27: 100,
    28: 220,
    33: 240,
    34: 280,
    35: 300,
    36: 320,
    42: 280,
    43: 320,
    65: 260,
    66: 260,
    67: 130,
    68: 260,
    69: 260,
    70: 220,
    71: 360,
    72: 360,
    73: 180,
    74: 220,
    75: 180,
    76: 120,
    77: 260,
    78: 140
  };

  Object.keys(widthMap).forEach(function(key) {
    targetSheet.setColumnWidth(Number(key), widthMap[key]);
  });

  return targetSheet;
}

function initConnectRequestSheet() {
  setConnectSpreadsheetBinding();
  ensureConnectRequestSheet(getConnectRequestSheet());
  SpreadsheetApp.getUi().alert("สร้างชีท '" + CONNECT_REQUEST_SHEET_NAME + "' เรียบร้อย");
}

function resetConnectDocNumberToOne() {
  const state = setConnectDocNumberNext(1);
  SpreadsheetApp.getUi().alert(
    "รีเซ็ตเลขหนังสือ CSD1 Connect เรียบร้อย\nเลขถัดไป: " + state.nextNumber
  );
}

/**
 * สร้างหัวตารางและชีทผู้ใช้
 */
function initSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setConnectSpreadsheetBinding(ss);

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
  applyReportSheetTimeColumnValidation(reportSheet);

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
  userSheet.setColumnWidth(6, 120);
  userSheet.setColumnWidth(7, 80);

  if (userSheet.getLastRow() < 2) {
    userSheet.getRange(2, 1, 1, 7).setValues([
      ["example@gmail.com", "ผู้ดูแลระบบ", "ผกก.", "ชป.1", "", "super_admin", "TRUE"]
    ]);
  }

  // === ชีท CommanderInfo (ผกก. และ สว.) ===
  let commanderSheet = ss.getSheetByName(COMMANDER_INFO_SHEET_NAME);
  if (!commanderSheet) {
    commanderSheet = ss.insertSheet(COMMANDER_INFO_SHEET_NAME);
  }
  const cmdHeaderRange = commanderSheet.getRange(1, 1, 1, COMMANDER_INFO_HEADERS.length);
  cmdHeaderRange.setValues([COMMANDER_INFO_HEADERS]);
  cmdHeaderRange.setFontWeight("bold");
  cmdHeaderRange.setBackground("#1e3a5f");
  cmdHeaderRange.setFontColor("#ffffff");
  cmdHeaderRange.setHorizontalAlignment("center");
  commanderSheet.setFrozenRows(1);
  commanderSheet.setColumnWidth(1, 80);
  commanderSheet.setColumnWidth(2, 260);
  commanderSheet.setColumnWidth(3, 140);
  commanderSheet.setColumnWidth(4, 120);

  // === ชีทคำขอแก้ไข/ลบ ===
  let requestSheet = ss.getSheetByName(CHANGE_REQUEST_SHEET_NAME);
  if (!requestSheet) {
    requestSheet = ss.insertSheet(CHANGE_REQUEST_SHEET_NAME);
  }
  const requestHeaderRange = requestSheet.getRange(1, 1, 1, CHANGE_REQUEST_HEADERS.length);
  requestHeaderRange.setValues([CHANGE_REQUEST_HEADERS]);
  requestHeaderRange.setFontWeight("bold");
  requestHeaderRange.setBackground("#7c2d12");
  requestHeaderRange.setFontColor("#ffffff");
  requestHeaderRange.setHorizontalAlignment("center");
  requestHeaderRange.setWrap(true);
  requestSheet.setFrozenRows(1);

  // === ชีทคำขอข้อมูลจาก CSD1 Connect ===
  ensureConnectRequestSheet(getConnectRequestSheet());

  // === ชีท Recycle Bin (เก็บแถวที่ลบ) ===
  let recycleSheet = ss.getSheetByName(RECYCLE_BIN_SHEET_NAME);
  if (!recycleSheet) {
    recycleSheet = ss.insertSheet(RECYCLE_BIN_SHEET_NAME);
  }
  ensureRecycleBinSchema(recycleSheet);
  const recycleHeaderRange = recycleSheet.getRange(1, 1, 1, HEADERS.length);
  recycleHeaderRange.setValues([HEADERS]);
  recycleHeaderRange.setFontWeight("bold");
  recycleHeaderRange.setBackground("#7f1d1d");
  recycleHeaderRange.setFontColor("#ffffff");
  recycleHeaderRange.setHorizontalAlignment("center");
  recycleHeaderRange.setWrap(true);

  // === ชีท MeetingBookings ===
  let meetingSheet = ss.getSheetByName(MEETING_BOOKINGS_SHEET_NAME);
  if (!meetingSheet) {
    meetingSheet = ss.insertSheet(MEETING_BOOKINGS_SHEET_NAME);
  }
  const meetingHeaderRange = meetingSheet.getRange(1, 1, 1, MEETING_BOOKINGS_HEADERS.length);
  meetingHeaderRange.setValues([MEETING_BOOKINGS_HEADERS]);
  meetingHeaderRange.setFontWeight("bold");
  meetingHeaderRange.setBackground("#7c3aed");
  meetingHeaderRange.setFontColor("#ffffff");
  meetingHeaderRange.setHorizontalAlignment("center");
  meetingSheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(
    "สร้างระบบเรียบร้อย!\n\n" +
    "1. ไปชีท 'User_master'\n" +
    "2. ใส่อีเมล Google ของเจ้าหน้าที่ที่อนุญาต\n" +
    "3. ระบบได้สร้างชีท '" + CONNECT_REQUEST_SHEET_NAME + "' สำหรับเก็บคำขอจากฟอร์ม CSD1 Connect แล้ว\n" +
    "4. Deploy > New deployment > Web app\n" +
    "5. นำ URL ไปตั้งค่าในหน้าเว็บ"
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

function normalizeEmail(email) {
  return String(email || "").trim().toLowerCase();
}

function getUserMasterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName("User_master");
}

function getChangeRequestSheet(createIfMissing) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CHANGE_REQUEST_SHEET_NAME);
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet(CHANGE_REQUEST_SHEET_NAME);
  }
  if (!sheet) return null;

  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, CHANGE_REQUEST_HEADERS.length).setValues([CHANGE_REQUEST_HEADERS]);
  }
  return sheet;
}

function ensureRecycleBinSchema(sheet) {
  if (!sheet) return;
  const requiredCols = HEADERS.length;
  const currentCols = sheet.getMaxColumns();

  if (currentCols < requiredCols) {
    sheet.insertColumnsAfter(currentCols, requiredCols - currentCols);
  } else if (currentCols > requiredCols) {
    sheet.deleteColumns(requiredCols + 1, currentCols - requiredCols);
  }

  sheet.getRange(1, 1, 1, requiredCols).setValues([HEADERS]);
  sheet.setFrozenRows(1);
}

function getRecycleBinSheet(createIfMissing) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RECYCLE_BIN_SHEET_NAME);
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet(RECYCLE_BIN_SHEET_NAME);
  }
  if (!sheet) return null;

  ensureRecycleBinSchema(sheet);
  return sheet;
}

function normalizeHeaderKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[()_\-./]/g, "");
}

function resolveUserMasterColumns(headerRow) {
  const defaults = { email: 0, name: 1, position: 2, unit: 3, phone: 4, role: 5, active: 6 };
  if (!headerRow || !headerRow.length) return defaults;

  const aliases = {
    email: ["อีเมล", "email", "e-mail", "mail", "gmail"],
    name: ["ชื่อสกุล", "ชื่อ", "name", "fullname", "full name", "full_name"],
    position: ["ตำแหน่ง", "position", "rank"],
    unit: ["ชุดปฏิบัติการ", "ชป", "หน่วย", "unit", "team"],
    phone: ["เบอร์โทร", "เบอร์โทรศัพท์", "โทรศัพท์", "phone", "tel", "mobile"],
    role: ["สิทธิ์", "สิทธิ", "role", "permission", "permissions", "access"],
    active: ["active", "สถานะ", "ใช้งาน", "เปิดใช้"]
  };

  const normalizedAliases = {};
  Object.keys(aliases).forEach(function(key) {
    normalizedAliases[key] = aliases[key].map(normalizeHeaderKey);
  });

  const resolved = Object.assign({}, defaults);
  const found = { email: false, name: false, position: false, unit: false, phone: false, role: false, active: false };
  for (let i = 0; i < headerRow.length; i++) {
    const hk = normalizeHeaderKey(headerRow[i]);
    if (!hk) continue;
    Object.keys(normalizedAliases).forEach(function(key) {
      if (found[key]) return;
      if (normalizedAliases[key].indexOf(hk) !== -1) {
        resolved[key] = i;
        found[key] = true;
      }
    });
  }
  return resolved;
}

function findUserByEmail(email, includeAvatar) {
  const normalizedEmail = normalizeEmail(email);
  if (!normalizedEmail) return null;

  const sheet = getUserMasterSheet();
  if (!sheet) return null;

  const rows = sheet.getDataRange().getValues();
  if (!rows || rows.length < 2) return null;
  const col = resolveUserMasterColumns(rows[0]);

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    let rowEmail = normalizeEmail(row[col.email]);
    if (!rowEmail) {
      for (let c = 0; c < row.length; c++) {
        const candidate = normalizeEmail(row[c]);
        if (candidate === normalizedEmail) {
          rowEmail = candidate;
          break;
        }
      }
    }
    if (rowEmail === normalizedEmail) {
      // active = TRUE/FALSE — ว่าบัญชียังเปิดใช้งานอยู่หรือเปล่า
      const activeVal = col.active < row.length ? String(row[col.active] || "").trim().toUpperCase() : "TRUE";
      if (activeVal === "FALSE") return null; // บัญชีถูกปิดใช้งาน
      const rawRole = String(row[col.role] || "").trim().toLowerCase();
      const normalizedRole = isValidRole(rawRole) ? rawRole : "reporter";
      const userName = String(row[col.name] || "").trim();
      return {
        email: rowEmail,
        name: userName,
        position: String(row[col.position] || "").trim(),
        unit: String(row[col.unit] || "").trim(),
        phone: String(row[col.phone] || "").trim(),
        role: normalizedRole,
        avatarUrl: includeAvatar ? getAvatarUrl(rowEmail, userName) : ""
      };
    }
  }
  return null;
}

function getIdTokenFromPayload(data) {
  if (!data || typeof data !== "object") return "";
  return String(
    data.idToken ||
    data.id_token ||
    data.credential ||
    data.googleCredential ||
    data.googleToken ||
    ""
  ).trim();
}

function verifyGoogleIdToken(idToken, email) {
  const token = String(idToken || "").trim();
  if (!token) {
    return { ok: false, message: "ไม่พบโทเคนสำหรับยืนยันตัวตน" };
  }

  const local = validateIdTokenPayloadLocally(token, email);
  if (local.ok) {
    const remote = verifyGoogleIdTokenRemotely(token, email);
    if (remote.ok) return remote;

    if (ALLOW_LOCAL_TOKEN_FALLBACK) {
      return {
        ok: true,
        warning: "tokeninfo endpoint ขัดข้อง ใช้ local token checks ชั่วคราว"
      };
    }

    return remote;
  }

  const remoteOnLocalFail = verifyGoogleIdTokenRemotely(token, email);
  if (remoteOnLocalFail.ok) {
    return {
      ok: true,
      warning: "local token checks ไม่ผ่าน แต่ tokeninfo endpoint ยืนยันสำเร็จ"
    };
  }

  return local;
}

function decodeJwtPayload(idToken) {
  const parts = String(idToken || "").split(".");
  if (parts.length < 2) return null;

  let payloadPart = String(parts[1] || "").replace(/-/g, "+").replace(/_/g, "/");
  while (payloadPart.length % 4 !== 0) payloadPart += "=";

  try {
    const bytes = Utilities.base64Decode(payloadPart);
    const json = Utilities.newBlob(bytes).getDataAsString("UTF-8");
    const payload = JSON.parse(json || "{}");
    return payload && typeof payload === "object" ? payload : null;
  } catch (err) {
    try {
      const bytes = Utilities.base64DecodeWebSafe(parts[1]);
      const json = Utilities.newBlob(bytes).getDataAsString("UTF-8");
      const payload = JSON.parse(json || "{}");
      return payload && typeof payload === "object" ? payload : null;
    } catch (_) {
      return null;
    }
  }
}

function validateIdTokenPayloadLocally(idToken, email) {
  const payload = decodeJwtPayload(idToken);
  if (!payload) {
    return { ok: false, message: "โทเคนไม่ถูกต้อง" };
  }

  const tokenEmail = normalizeEmail(payload.email);
  const verified = payload.email_verified === true || String(payload.email_verified || "").toLowerCase() === "true";
  const audCandidates = []
    .concat(payload.aud || [])
    .concat(payload.azp || [])
    .map(function(v) { return String(v || "").trim(); })
    .filter(Boolean);
  const iss = String(payload.iss || "");
  const exp = Number(payload.exp || 0);
  const nbf = Number(payload.nbf || 0);
  const nowSec = Math.floor(Date.now() / 1000);

  if (!tokenEmail || tokenEmail !== email) {
    return { ok: false, message: "อีเมลในโทเคนไม่ตรงกับผู้ใช้งาน" };
  }
  if (!verified) {
    return { ok: false, message: "บัญชี Google ยังไม่ยืนยันอีเมล" };
  }
  if (
    ALLOWED_GOOGLE_CLIENT_IDS.length &&
    !audCandidates.some(function(aud) { return ALLOWED_GOOGLE_CLIENT_IDS.indexOf(aud) !== -1; })
  ) {
    return { ok: false, message: "โทเคนมาจากแอปที่ไม่อนุญาต" };
  }
  if (iss !== "https://accounts.google.com" && iss !== "accounts.google.com") {
    return { ok: false, message: "issuer ของโทเคนไม่ถูกต้อง" };
  }
  if (nbf && nowSec + 300 < nbf) {
    return { ok: false, message: "โทเคนยังไม่พร้อมใช้งาน (nbf)" };
  }
  if (!exp || exp <= nowSec) {
    return { ok: false, message: "โทเคนหมดอายุ" };
  }

  return { ok: true };
}

function verifyGoogleIdTokenRemotely(idToken, email) {
  const endpoints = [
    "https://oauth2.googleapis.com/tokeninfo?id_token=",
    "https://www.googleapis.com/oauth2/v3/tokeninfo?id_token="
  ];
  let lastError = "";

  try {
    for (let i = 0; i < endpoints.length; i++) {
      let res;
      try {
        res = UrlFetchApp.fetch(
          endpoints[i] + encodeURIComponent(idToken),
          { muteHttpExceptions: true, followRedirects: true }
        );
      } catch (fetchErr) {
        lastError = String(fetchErr && fetchErr.message ? fetchErr.message : fetchErr);
        continue;
      }

      const code = res.getResponseCode();
      if (code !== 200) {
        lastError = "HTTP " + code;
        continue;
      }

      let payload = {};
      try {
        payload = JSON.parse(res.getContentText() || "{}");
      } catch (parseErr) {
        lastError = "invalid tokeninfo payload";
        continue;
      }

      const tokenEmail = normalizeEmail(payload.email);
      const tokenAudience = String(payload.aud || payload.azp || "").trim();
      if (tokenEmail && tokenEmail === email) {
        if (
          ALLOWED_GOOGLE_CLIENT_IDS.length &&
          tokenAudience &&
          ALLOWED_GOOGLE_CLIENT_IDS.indexOf(tokenAudience) === -1
        ) {
          lastError = "tokeninfo audience mismatch";
          continue;
        }
        return { ok: true };
      }
      lastError = "tokeninfo email mismatch";
    }
  } catch (err) {
    lastError = String(err && err.message ? err.message : err);
  }

  const message = lastError
    ? "ตรวจสอบโทเคนล้มเหลว (" + lastError + ")"
    : "ตรวจสอบโทเคนล้มเหลว";
  return { ok: false, message: message };
}

function createSessionToken(email, role) {
  const token = Utilities.getUuid().replace(/-/g, "") + "." + Utilities.getUuid().replace(/-/g, "");
  const cache = CacheService.getScriptCache();
  const session = {
    email: normalizeEmail(email),
    role: String(role || "user"),
    issuedAt: new Date().toISOString()
  };
  cache.put(SESSION_CACHE_PREFIX + token, JSON.stringify(session), SESSION_TTL_SECONDS);
  return token;
}

function getSessionData(sessionToken) {
  const token = String(sessionToken || "").trim();
  if (!token) return null;

  const raw = CacheService.getScriptCache().get(SESSION_CACHE_PREFIX + token);
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}

function authorizeRequest(data) {
  const email = normalizeEmail(data && data.email);
  const sessionToken = String((data && data.sessionToken) || "").trim();

  if (!email || !sessionToken) {
    return { ok: false, message: "กรุณาเข้าสู่ระบบใหม่" };
  }

  const session = getSessionData(sessionToken);
  if (!session || normalizeEmail(session.email) !== email) {
    return { ok: false, message: "เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่" };
  }

  const user = findUserByEmail(email, false);
  if (!user) {
    return { ok: false, message: "บัญชีผู้ใช้นี้ไม่มีสิทธิ์ใช้งาน" };
  }

  // Extend session age while active.
  CacheService.getScriptCache().put(SESSION_CACHE_PREFIX + sessionToken, JSON.stringify(session), SESSION_TTL_SECONDS);
  return { ok: true, user: user };
}

function normalizeConnectDocNumberNextValue(value) {
  const parsed = Number(String(value == null ? "" : value).trim());
  if (!isFinite(parsed)) return CONNECT_DOC_NUMBER_MIN;
  const normalized = Math.floor(parsed);
  if (normalized < CONNECT_DOC_NUMBER_MIN) return CONNECT_DOC_NUMBER_MIN;
  if (normalized > CONNECT_DOC_NUMBER_MAX + 1) return CONNECT_DOC_NUMBER_MAX + 1;
  return normalized;
}

function validateConnectDocNumberNextInput(nextNumber) {
  const raw = String(nextNumber == null ? "" : nextNumber).trim();
  if (!/^\d+$/.test(raw)) {
    throw new Error("กรุณาระบุเลขถัดไปเป็นตัวเลข 1-" + (CONNECT_DOC_NUMBER_MAX + 1));
  }
  const parsed = Number(raw);
  if (!isFinite(parsed) || parsed < CONNECT_DOC_NUMBER_MIN || parsed > CONNECT_DOC_NUMBER_MAX + 1) {
    throw new Error("เลขถัดไปต้องอยู่ระหว่าง 1-" + (CONNECT_DOC_NUMBER_MAX + 1));
  }
  return Math.floor(parsed);
}

function getConnectDocNumberState() {
  const stored = PropertiesService.getScriptProperties().getProperty(CONNECT_DOC_NUMBER_PROPERTY_KEY);
  const nextNumber = normalizeConnectDocNumberNextValue(stored);
  return {
    nextNumber: nextNumber,
    exhausted: nextNumber > CONNECT_DOC_NUMBER_MAX,
    remaining: Math.max(0, CONNECT_DOC_NUMBER_MAX - nextNumber + 1),
    min: CONNECT_DOC_NUMBER_MIN,
    max: CONNECT_DOC_NUMBER_MAX
  };
}

function setConnectDocNumberNext(nextNumber) {
  const normalized = validateConnectDocNumberNextInput(nextNumber);
  PropertiesService.getScriptProperties().setProperty(CONNECT_DOC_NUMBER_PROPERTY_KEY, String(normalized));
  const state = getConnectDocNumberState();
  Logger.log("[CSD1 Connect] set next doc number = " + JSON.stringify(state));
  return state;
}

function handleReserveConnectDocNumber(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) {
    return jsonResponse({ status: "error", message: auth.message });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const state = getConnectDocNumberState();
    if (state.exhausted) {
      return jsonResponse({
        status: "error",
        message: "เลขหนังสือส่งรันครบ 9999 แล้ว",
        exhausted: true,
        max: CONNECT_DOC_NUMBER_MAX
      });
    }

    const reservedNumber = state.nextNumber;
    const nextNumber = reservedNumber + 1;
    PropertiesService.getScriptProperties().setProperty(CONNECT_DOC_NUMBER_PROPERTY_KEY, String(nextNumber));

    return jsonResponse({
      status: "success",
      docNumber: String(reservedNumber),
      nextNumber: nextNumber > CONNECT_DOC_NUMBER_MAX ? "" : String(nextNumber),
      exhausted: nextNumber > CONNECT_DOC_NUMBER_MAX,
      max: CONNECT_DOC_NUMBER_MAX,
      reservedAt: new Date().toISOString(),
      reservedBy: auth.user.email || ""
    });
  } catch (err) {
    return jsonResponse({
      status: "error",
      message: "ไม่สามารถจองเลขหนังสือส่งได้ กรุณาลองใหม่อีกครั้ง"
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

function readReportRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("รายงานผลการปฏิบัติงาน");
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  const rows = [];
  for (let i = 0; i < allData.length; i++) {
    const rowValues = allData[i];
    const hasAnyValue = rowValues.some(function(cell) {
      return String(cell == null ? "" : cell).trim() !== "";
    });
    if (!hasAnyValue) continue;

    const obj = {};
    const headerSeen = {};
    for (let j = 0; j < HEADERS.length; j++) {
      const header = HEADERS[j];
      const value = rowValues[j] != null ? String(rowValues[j]) : "";

      // Keep backward compatibility by preserving the base key behavior,
      // but also add positional keys to avoid losing duplicated headers.
      obj[header] = value;
      const seenCount = (headerSeen[header] || 0) + 1;
      headerSeen[header] = seenCount;
      obj[header + " #" + seenCount] = value;
    }
    rows.push(obj);
  }
  return rows;
}

function readChangeRequestRows() {
  const sheet = getChangeRequestSheet(false);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, CHANGE_REQUEST_HEADERS.length).getValues();
  return values.map(function(row) {
    const out = {};
    for (let i = 0; i < CHANGE_REQUEST_HEADERS.length; i++) {
      out[CHANGE_REQUEST_HEADERS[i]] = row[i] != null ? String(row[i]) : "";
    }
    return out;
  });
}

function readConnectRequestRows() {
  const sheet = ensureConnectRequestSheet(getConnectRequestSheet());
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, CONNECT_REQUEST_HEADERS.length).getValues();
  return values
    .filter(function(row) {
      return row.some(function(cell) {
        return String(cell == null ? "" : cell).trim() !== "";
      });
    })
    .map(function(row, index) {
      const out = {};
      for (let i = 0; i < CONNECT_REQUEST_HEADERS.length; i++) {
        out[CONNECT_REQUEST_HEADERS[i]] = row[i] != null ? String(row[i]) : "";
      }
      out.__rowIndex = index + 2;
      return out;
    });
}

function toPublicRow(row) {
  const out = {};
  for (let i = 0; i < PUBLIC_REPORT_HEADERS.length; i++) {
    const key = PUBLIC_REPORT_HEADERS[i];
    out[key] = row[key] != null ? String(row[key]) : "";
  }
  return out;
}

function normalizeUnitKey(value) {
  return String(value || "").trim().toLowerCase().replace(/\s+/g, "");
}

function normalizeCourtValue(value) {
  const courtText = String(value || "").replace(/\s+/g, " ").trim();
  if (!courtText) return "";
  if (/^ศาล\s*$/u.test(courtText)) return "";
  if (/^ศาลที่ออกหมาย(?:จับ|ค้น)?\s*$/u.test(courtText)) return "";
  return courtText;
}

function makeChangeRequestTargetKey(targetNo, targetTimestamp, targetReporterEmail, targetUnit) {
  return [
    String(targetNo || "").trim(),
    String(targetTimestamp || "").trim(),
    normalizeEmail(targetReporterEmail || ""),
    normalizeUnitKey(targetUnit || "")
  ].join("|");
}

function getReportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName("รายงานผลการปฏิบัติงาน") || ss.getActiveSheet();
}

function buildReportHeaderColumnMap() {
  const map = {};
  const seen = {};
  for (let i = 0; i < HEADERS.length; i++) {
    const header = String(HEADERS[i] || "");
    if (!header) continue;
    seen[header] = (seen[header] || 0) + 1;
    const aliasKey = header + " #" + seen[header];
    map[aliasKey] = i + 1;
    if (!Object.prototype.hasOwnProperty.call(map, header)) {
      map[header] = i + 1;
    }
  }
  return map;
}

function normalizeChangedFieldsPayload(input) {
  if (!input || typeof input !== "object" || Array.isArray(input)) return {};
  const out = {};
  Object.keys(input).forEach(function(rawKey) {
    const key = String(rawKey || "").trim();
    if (!key) return;
    let value = String(input[rawKey] == null ? "" : input[rawKey]).trim();
    if (key === "เวลา") {
      value = normalizeSheetTimeValue(value) || value;
    }
    if (key === "วันที่" || key === "วันเดือนปีที่ออกหมาย" || key === "วันเดือนปีที่ออกหมายค้น") {
      value = normalizeSheetThaiDateValue(value) || value;
    }
    if (key === "ศาลที่ออกหมายจับ" || key === "ศาลที่ออกหมายค้น") {
      value = normalizeCourtValue(value);
    }
    out[key] = value;
  });
  return out;
}

function parseCorrectedDataPayload(rawValue) {
  const raw = String(rawValue || "").trim();
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === "object") return parsed;
  } catch (_) {}
  return null;
}

function findChangeRequestById(sheet, requestId) {
  if (!sheet || !requestId || sheet.getLastRow() < 2) return null;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, CHANGE_REQUEST_HEADERS.length).getValues();
  for (let i = 0; i < values.length; i++) {
    const rowData = values[i];
    const currentId = String(rowData[0] || "").trim();
    if (currentId !== requestId) continue;
    const out = {};
    for (let j = 0; j < CHANGE_REQUEST_HEADERS.length; j++) {
      out[CHANGE_REQUEST_HEADERS[j]] = rowData[j] != null ? String(rowData[j]) : "";
    }
    return { rowIndex: i + 2, row: out };
  }
  return null;
}

function normalizeConnectString(value) {
  return String(value == null ? "" : value).trim();
}

function normalizeConnectNumber(value) {
  const raw = normalizeConnectString(value);
  if (!raw) return "";
  const parsed = Number(raw);
  return isFinite(parsed) ? String(parsed) : raw;
}

function normalizeConnectStringArray(value) {
  if (!Array.isArray(value)) return [];
  return value.map(normalizeConnectString).filter(function(item) { return item !== ""; });
}

function normalizeConnectObjectArray(value) {
  if (!Array.isArray(value)) return [];
  return value.filter(function(item) { return item && typeof item === "object"; });
}

function parseConnectJsonArray(value) {
  const raw = normalizeConnectString(value);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (_) {
    return [];
  }
}

function safeJsonStringify(value) {
  try {
    return JSON.stringify(value == null ? null : value);
  } catch (_) {
    return "";
  }
}

function truncateSheetCellText(value, limit) {
  const maxChars = Number(limit || CONNECT_SHEET_CELL_CHAR_LIMIT);
  const text = value == null ? "" : String(value);
  if (!maxChars || text.length <= maxChars) return text;
  const suffix = " ...(truncated)";
  const safeLimit = Math.max(0, maxChars - suffix.length);
  return text.slice(0, safeLimit) + suffix;
}

function sanitizeRowValuesForSheet(headers, rowValues) {
  const safeHeaders = Array.isArray(headers) ? headers : [];
  const safeRowValues = Array.isArray(rowValues) ? rowValues.slice() : [];
  const truncatedHeaders = [];

  for (var index = 0; index < safeRowValues.length; index++) {
    const value = safeRowValues[index];
    if (typeof value !== "string") continue;
    if (value.length <= CONNECT_SHEET_CELL_CHAR_LIMIT) continue;
    safeRowValues[index] = truncateSheetCellText(value, CONNECT_SHEET_CELL_CHAR_LIMIT);
    truncatedHeaders.push(safeHeaders[index] || ("column_" + (index + 1)));
  }

  return {
    rowValues: safeRowValues,
    truncatedHeaders: truncatedHeaders
  };
}

function buildCompactConnectRawState(rawState) {
  if (!rawState || typeof rawState !== "object") return rawState;
  return {
    step: rawState.step,
    unit: rawState.unit,
    commander: rawState.commander,
    docNumber: rawState.docNumber,
    caseDetail: rawState.caseDetail,
    casePresetKey: rawState.casePresetKey,
    caseDetailCustom: rawState.caseDetailCustom,
    caseSearchTerm: rawState.caseSearchTerm,
    adminNoteTag: rawState.adminNoteTag,
    adminNoteTagLabel: rawState.adminNoteTagLabel,
    adminNoteText: rawState.adminNoteText,
    requestHistoryStatus: rawState.requestHistoryStatus,
    type: rawState.type,
    phoneNetwork: rawState.phoneNetwork,
    aisSubType: rawState.aisSubType,
    phoneSubType: rawState.phoneSubType,
    phoneImei: rawState.phoneImei,
    phoneNumbers: Array.isArray(rawState.phoneNumbers) ? rawState.phoneNumbers : [],
    ipEntries: Array.isArray(rawState.ipEntries) ? rawState.ipEntries : [],
    phoneDateStart: rawState.phoneDateStart,
    phoneDateEnd: rawState.phoneDateEnd,
    bankCode: rawState.bankCode,
    bankSubType: rawState.bankSubType,
    statementAccType: rawState.statementAccType,
    bankAccounts: Array.isArray(rawState.bankAccounts) ? rawState.bankAccounts : [],
    bankAccountNames: Array.isArray(rawState.bankAccountNames) ? rawState.bankAccountNames : [],
    bankPromptPay: rawState.bankPromptPay,
    xxxAmount: rawState.xxxAmount,
    xxxDate: rawState.xxxDate,
    xxxTime: rawState.xxxTime,
    bankAccountName: rawState.bankAccountName,
    atmAccountNo: rawState.atmAccountNo,
    atmDate: rawState.atmDate,
    atmTime: rawState.atmTime,
    atmLocation: rawState.atmLocation,
    atmTerminalId: rawState.atmTerminalId,
    bankDateStart: rawState.bankDateStart,
    bankDateEnd: rawState.bankDateEnd,
    trueId: rawState.trueId,
    trueName: rawState.trueName,
    trueDateStart: rawState.trueDateStart,
    trueDateEnd: rawState.trueDateEnd,
    taxType: rawState.taxType,
    taxId: rawState.taxId,
    taxName: rawState.taxName,
    taxYearStart: rawState.taxYearStart,
    taxYearEnd: rawState.taxYearEnd,
    requestId: rawState.requestId,
    generatedPdfUrl: rawState.generatedPdfUrl,
    generatedDocxUrl: rawState.generatedDocxUrl,
    generatedWarnings: Array.isArray(rawState.generatedWarnings) ? rawState.generatedWarnings : [],
    generatedMode: rawState.generatedMode
  };
}

function buildCompactConnectPayload(payload) {
  if (!payload || typeof payload !== "object") return payload;
  return {
    requestId: normalizeConnectString(payload.requestId),
    submittedAt: normalizeConnectString(payload.submittedAt),
    submittedBy: payload.submittedBy && typeof payload.submittedBy === "object" ? payload.submittedBy : {},
    type: normalizeConnectString(payload.type),
    typeLabel: normalizeConnectString(payload.typeLabel),
    unit: normalizeConnectString(payload.unit),
    commander: normalizeConnectString(payload.commander),
    docNumber: normalizeConnectString(payload.docNumber),
    casePresetKey: normalizeConnectString(payload.casePresetKey),
    caseDetail: normalizeConnectString(payload.caseDetail),
    caseDetailCustom: normalizeConnectString(payload.caseDetailCustom),
    caseSearchTerm: normalizeConnectString(payload.caseSearchTerm),
    adminNoteTag: normalizeConnectString(payload.adminNoteTag),
    adminNoteTagLabel: normalizeConnectString(payload.adminNoteTagLabel),
    adminNoteText: normalizeConnectString(payload.adminNoteText),
    requestHistoryStatus: normalizeConnectString(payload.requestHistoryStatus),
    requestHistoryLabel: normalizeConnectString(payload.requestHistoryLabel),
    status: normalizeConnectString(payload.status),
    statusNote: normalizeConnectString(payload.statusNote),
    driveLink: normalizeConnectString(payload.driveLink),
    filePassword: payload.filePassword == null ? "" : String(payload.filePassword),
    phone: payload.phone && typeof payload.phone === "object" ? payload.phone : {},
    bank: payload.bank && typeof payload.bank === "object" ? payload.bank : {},
    trueMoney: payload.trueMoney && typeof payload.trueMoney === "object"
      ? (payload.trueMoney || payload.truemoney)
      : {},
    tax: payload.tax && typeof payload.tax === "object" ? payload.tax : {},
    files: payload.files && typeof payload.files === "object" ? {
      fileBaseName: normalizeConnectString(payload.files.fileBaseName),
      pdfFileName: normalizeConnectString(payload.files.pdfFileName),
      generatedMode: normalizeConnectString(payload.files.generatedMode),
      generatedPdfUrl: normalizeConnectString(payload.files.generatedPdfUrl),
      generatedDocxUrl: normalizeConnectString(payload.files.generatedDocxUrl),
      generatedWarnings: Array.isArray(payload.files.generatedWarnings) ? payload.files.generatedWarnings : []
    } : {},
    rawState: buildCompactConnectRawState(payload.rawState)
  };
}

function formatConnectIpEntriesText(entries) {
  return normalizeConnectObjectArray(entries).map(function(entry, index) {
    const parts = [
      "#" + (index + 1),
      normalizeConnectString(entry.ip),
      normalizeConnectString(entry.date),
      normalizeConnectString(entry.time),
      normalizeConnectString(entry.port) ? "port " + normalizeConnectString(entry.port) : ""
    ].filter(function(part) { return part; });
    return parts.join(" | ");
  }).join("\n");
}

function formatConnectBankAccountsText(accounts) {
  return normalizeConnectObjectArray(accounts).map(function(account, index) {
    const accountNo = normalizeConnectString(account.account_no || account.accountNo);
    const accountName = normalizeConnectString(account.account_name || account.accountName);
    return ["#" + (index + 1), accountNo, accountName]
      .filter(function(part) { return part; })
      .join(" | ");
  }).join("\n");
}

function findConnectRequestById(sheet, requestId) {
  const targetSheet = ensureConnectRequestSheet(sheet || getConnectRequestSheet());
  if (!targetSheet || !requestId || targetSheet.getLastRow() < 2) return null;
  const values = targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, CONNECT_REQUEST_HEADERS.length).getValues();
  for (var i = 0; i < values.length; i++) {
    const rowData = values[i];
    const currentId = normalizeConnectString(rowData[0]);
    if (currentId !== requestId) continue;
    const out = {};
    for (var j = 0; j < CONNECT_REQUEST_HEADERS.length; j++) {
      out[CONNECT_REQUEST_HEADERS[j]] = rowData[j] != null ? String(rowData[j]) : "";
    }
    return { rowIndex: i + 2, row: out };
  }
  return null;
}

function normalizeConnectStatus(value) {
  const status = normalizeConnectString(value).toLowerCase();
  return ["pending", "processing", "done", "cancelled"].indexOf(status) !== -1
    ? status
    : "pending";
}

function canAccessConnectRequestRow(authUser, row) {
  const safeUser = authUser || {};
  if (isConnectAdminRole(safeUser.role)) return true;
  const authUnitKey = normalizeUnitKey(safeUser.unit || "");
  const authEmail = normalizeEmail(safeUser.email || "");
  const rowUnitKey = normalizeUnitKey(row && (row.unit || row.submitted_by_unit) || "");
  const rowEmail = normalizeEmail(row && row.submitted_by_email || "");
  if (authUnitKey && rowUnitKey) return authUnitKey === rowUnitKey;
  return !!authEmail && rowEmail === authEmail;
}

function extractDriveFileIdFromUrl(url) {
  const raw = normalizeConnectString(url);
  if (!raw) return "";
  var match = raw.match(/\/file\/d\/([^/?#]+)/i);
  if (match && match[1]) return match[1];
  match = raw.match(/[?&]id=([^&#]+)/i);
  if (match && match[1]) return match[1];
  return "";
}

function getHeaderIgnoreCase(headers, key) {
  const target = String(key || "").toLowerCase();
  if (!headers || typeof headers !== "object" || !target) return "";
  const keys = Object.keys(headers);
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i] || "").toLowerCase() === target) {
      return String(headers[keys[i]] == null ? "" : headers[keys[i]]);
    }
  }
  return "";
}

function getFilenameFromContentDispositionHeader(headerValue) {
  const raw = String(headerValue || "");
  if (!raw) return "";
  var utf8Match = raw.match(/filename\*\s*=\s*UTF-8''([^;]+)/i);
  if (utf8Match && utf8Match[1]) {
    try {
      return decodeURIComponent(utf8Match[1]);
    } catch (_) {}
  }
  var quotedMatch = raw.match(/filename\s*=\s*"([^"]+)"/i);
  if (quotedMatch && quotedMatch[1]) return quotedMatch[1];
  var plainMatch = raw.match(/filename\s*=\s*([^;]+)/i);
  return plainMatch && plainMatch[1] ? plainMatch[1].trim() : "";
}

function getConnectDownloadPayloadFromDriveFile(file) {
  if (!file) throw new Error("ไม่พบไฟล์ใน Google Drive");
  const blob = file.getBlob();
  return {
    sourceFileName: normalizeConnectString(file.getName()) || normalizeConnectString(blob.getName()),
    mimeType: normalizeConnectString(blob.getContentType()) || normalizeConnectString(file.getMimeType()),
    dataBase64: Utilities.base64Encode(blob.getBytes())
  };
}

function getConnectDownloadPayloadFromLink(url) {
  const link = normalizeConnectString(url);
  if (!link) throw new Error("ไม่พบลิงก์ไฟล์ตอบกลับ");

  const driveFileId = extractDriveFileIdFromUrl(link);
  if (driveFileId) {
    const file = DriveApp.getFileById(driveFileId);
    return getConnectDownloadPayloadFromDriveFile(file);
  }

  const response = UrlFetchApp.fetch(link, {
    followRedirects: true,
    muteHttpExceptions: true
  });
  const statusCode = Number(response.getResponseCode());
  if (!isFinite(statusCode) || statusCode < 200 || statusCode >= 300) {
    throw new Error("ดาวน์โหลดไฟล์ตอบกลับไม่สำเร็จ (HTTP " + statusCode + ")");
  }
  const blob = response.getBlob();
  const headers = response.getHeaders();
  return {
    sourceFileName: getFilenameFromContentDispositionHeader(getHeaderIgnoreCase(headers, "content-disposition")) || normalizeConnectString(blob.getName()),
    mimeType: normalizeConnectString(blob.getContentType()),
    dataBase64: Utilities.base64Encode(blob.getBytes())
  };
}

function normalizeDriveNameKey(value) {
  return String(value == null ? "" : value)
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\s+/g, "")
    .trim()
    .toLowerCase();
}

function getConnectPdfRootFolder(createIfMissing) {
  const rootFolderName = "CSD1 Connect PDFs";
  const rootFolders = [];
  const rootIter = DriveApp.getFoldersByName(rootFolderName);
  while (rootIter.hasNext()) {
    rootFolders.push(rootIter.next());
  }
  if (rootFolders.length) return rootFolders[0];
  return createIfMissing ? DriveApp.createFolder(rootFolderName) : null;
}

function getOrCreateConnectSummonRequestFolder(requestId) {
  const cleanRequestId = normalizeConnectString(requestId);
  if (!cleanRequestId) return null;

  const existingFolders = listConnectSummonRequestFolders(cleanRequestId);
  if (existingFolders.length) return existingFolders[0];

  const rootFolder = getConnectPdfRootFolder(true);
  return rootFolder ? rootFolder.createFolder(cleanRequestId) : null;
}

function listConnectSummonRequestFolders(requestId) {
  const cleanRequestId = normalizeConnectString(requestId);
  if (!cleanRequestId) return [];

  const folders = [];
  const seenFolderIds = {};
  const requestIdKey = normalizeDriveNameKey(cleanRequestId);
  const rootFolder = getConnectPdfRootFolder(false);
  const rootFolders = [];
  if (rootFolder) rootFolders.push(rootFolder);
  const rootIter = DriveApp.getFoldersByName("CSD1 Connect PDFs");
  while (rootIter.hasNext()) {
    const nextRoot = rootIter.next();
    const nextRootId = normalizeConnectString(nextRoot.getId());
    if (rootFolders.some(function(folder) { return normalizeConnectString(folder.getId()) === nextRootId; })) continue;
    rootFolders.push(nextRoot);
  }

  for (var rootIndex = 0; rootIndex < rootFolders.length; rootIndex++) {
    const currentRootFolder = rootFolders[rootIndex];
    const reqIter = currentRootFolder.getFoldersByName(cleanRequestId);
    while (reqIter.hasNext()) {
      const reqFolder = reqIter.next();
      const folderId = normalizeConnectString(reqFolder.getId());
      if (folderId && !seenFolderIds[folderId]) {
        seenFolderIds[folderId] = true;
        folders.push(reqFolder);
      }
    }

    const allFolders = currentRootFolder.getFolders();
    while (allFolders.hasNext()) {
      const candidateFolder = allFolders.next();
      const candidateId = normalizeConnectString(candidateFolder.getId());
      if (candidateId && seenFolderIds[candidateId]) continue;
      if (normalizeDriveNameKey(candidateFolder.getName()) !== requestIdKey) continue;
      if (candidateId) seenFolderIds[candidateId] = true;
      folders.push(candidateFolder);
    }
  }
  return folders;
}

function listConnectSummonFilesInFolder(folder) {
  const files = [];
  if (!folder) return files;

  const iter = folder.getFiles();
  while (iter.hasNext()) {
    files.push(iter.next());
  }
  return files;
}

function scoreConnectSummonFileCandidate(file, preferredFileName) {
  if (!file) return -1;

  const cleanPreferredFileName = normalizeConnectString(preferredFileName);
  const preferredFileNameKey = normalizeDriveNameKey(cleanPreferredFileName);
  const fileNameRaw = normalizeConnectString(file.getName());
  const fileNameKey = normalizeDriveNameKey(fileNameRaw);
  const mimeType = normalizeConnectString(file.getMimeType()).toLowerCase();
  let score = 0;

  if (preferredFileNameKey && fileNameKey === preferredFileNameKey) {
    score += 1000;
  }
  if (mimeType === "application/pdf") {
    score += 300;
  }
  if (/\.pdf$/i.test(fileNameRaw)) {
    score += 200;
  }
  if (!score) {
    score += 50;
  }

  try {
    const lastUpdated = file.getLastUpdated();
    if (lastUpdated && !Number.isNaN(lastUpdated.getTime())) {
      score += Math.floor(lastUpdated.getTime() / 1000);
    }
  } catch (_) {}

  return score;
}

function findConnectSummonPdfFileByRequestId(requestId, preferredFileName) {
  const cleanRequestId = normalizeConnectString(requestId);
  if (!cleanRequestId) return null;

  const requestFolders = listConnectSummonRequestFolders(cleanRequestId);
  if (!requestFolders.length) return null;

  let bestFile = null;
  let bestScore = -1;
  for (var folderIndex = 0; folderIndex < requestFolders.length; folderIndex++) {
    var reqFolder = requestFolders[folderIndex];
    const files = listConnectSummonFilesInFolder(reqFolder);
    for (var fileIndex = 0; fileIndex < files.length; fileIndex++) {
      const file = files[fileIndex];
      const score = scoreConnectSummonFileCandidate(file, preferredFileName);
      if (score > bestScore) {
        bestFile = file;
        bestScore = score;
      }
    }
  }
  return bestFile;
}

function ensureConnectRequestGeneratedPdfUrl(sheet, row, rowIndex) {
  if (!row || typeof row !== "object") return "";

  const existingUrl = normalizeConnectString(row.generated_pdf_url);
  if (existingUrl) return existingUrl;

  const requestId = normalizeConnectString(row.request_id);
  if (!requestId) return "";

  const fallbackFile = findConnectSummonPdfFileByRequestId(
    requestId,
    normalizeConnectString(row.pdf_file_name)
  );
  if (!fallbackFile) return "";

  fallbackFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const resolvedUrl = normalizeConnectString(fallbackFile.getUrl());
  if (!resolvedUrl) return "";

  row.generated_pdf_url = resolvedUrl;

  if (sheet && rowIndex > 0) {
    const pdfUrlCol = CONNECT_REQUEST_HEADERS.indexOf("generated_pdf_url") + 1;
    if (pdfUrlCol > 0) sheet.getRange(rowIndex, pdfUrlCol).setValue(resolvedUrl);
  }

  return resolvedUrl;
}

function backfillConnectRequestGeneratedPdfUrls() {
  const sheet = ensureConnectRequestSheet(getConnectRequestSheet());
  const rows = readConnectRequestRows();
  let repaired = 0;
  let alreadyPresent = 0;
  let missingFile = 0;
  const missingRequestIds = [];

  rows.forEach(function(row) {
    if (!row) return;
    const beforeUrl = normalizeConnectString(row.generated_pdf_url);
    if (beforeUrl) {
      alreadyPresent += 1;
      return;
    }
    const afterUrl = ensureConnectRequestGeneratedPdfUrl(sheet, row, Number(row.__rowIndex || 0));
    if (afterUrl) {
      repaired += 1;
    } else {
      missingFile += 1;
      missingRequestIds.push(normalizeConnectString(row.request_id));
    }
  });

  Logger.log("backfillConnectRequestGeneratedPdfUrls => repaired=%s alreadyPresent=%s missingFile=%s", repaired, alreadyPresent, missingFile);
  if (missingRequestIds.length) {
    Logger.log("Missing generated PDF for requestIds: %s", missingRequestIds.join(", "));
  }

  return {
    total: rows.length,
    repaired: repaired,
    alreadyPresent: alreadyPresent,
    missingFile: missingFile,
    missingRequestIds: missingRequestIds
  };
}

function diagnoseConnectSummonPdf(requestId) {
  const cleanRequestId = normalizeConnectString(requestId);
  const result = {
    requestId: cleanRequestId,
    foundRow: false,
    rowIndex: 0,
    sheetGeneratedPdfUrl: "",
    sheetPdfFileName: "",
    rootFolderCount: 0,
    requestFolderCount: 0,
    folderFound: false,
    fileFound: false,
    requestFolderNames: [],
    availableFiles: [],
    fileName: "",
    fileUrl: "",
    repairedUrl: "",
    error: ""
  };

  if (!cleanRequestId) {
    result.error = "missing requestId";
    Logger.log(JSON.stringify(result));
    return result;
  }

  try {
    const sheet = ensureConnectRequestSheet(getConnectRequestSheet());
    const found = findConnectRequestById(sheet, cleanRequestId);
    if (!found || !found.row) {
      result.error = "row not found";
      Logger.log(JSON.stringify(result));
      return result;
    }

    result.foundRow = true;
    result.rowIndex = Number(found.rowIndex || 0);
    result.sheetGeneratedPdfUrl = normalizeConnectString(found.row.generated_pdf_url);
    result.sheetPdfFileName = normalizeConnectString(found.row.pdf_file_name);

    const rootFolders = [];
    const rootIter = DriveApp.getFoldersByName("CSD1 Connect PDFs");
    while (rootIter.hasNext()) {
      rootFolders.push(rootIter.next());
    }
    result.rootFolderCount = rootFolders.length;
    if (!rootFolders.length) {
      result.error = "root folder not found";
      Logger.log(JSON.stringify(result));
      return result;
    }

    const requestFolders = listConnectSummonRequestFolders(cleanRequestId);
    result.requestFolderCount = requestFolders.length;
    result.requestFolderNames = requestFolders.map(function(folder) {
      return normalizeConnectString(folder.getName());
    });
    if (!requestFolders.length) {
      result.error = "request folder not found";
      Logger.log(JSON.stringify(result));
      return result;
    }

    result.folderFound = true;
    result.availableFiles = requestFolders.map(function(folder) {
      return {
        folderName: normalizeConnectString(folder.getName()),
        files: listConnectSummonFilesInFolder(folder).map(function(file) {
          return {
            name: normalizeConnectString(file.getName()),
            mimeType: normalizeConnectString(file.getMimeType())
          };
        })
      };
    });
    const file = findConnectSummonPdfFileByRequestId(cleanRequestId, result.sheetPdfFileName);
    if (!file) {
      result.error = "pdf file not found in request folder";
      Logger.log(JSON.stringify(result));
      return result;
    }

    result.fileFound = true;
    result.fileName = normalizeConnectString(file.getName());
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    result.fileUrl = normalizeConnectString(file.getUrl());

    if (!result.sheetGeneratedPdfUrl && result.fileUrl) {
      const repairedUrl = ensureConnectRequestGeneratedPdfUrl(sheet, found.row, result.rowIndex);
      result.repairedUrl = normalizeConnectString(repairedUrl);
    }
  } catch (error) {
    result.error = error && error.message ? error.message : String(error);
  }

  Logger.log(JSON.stringify(result));
  return result;
}

function toConnectRequestResponse(row) {
  if (!row) return null;
  const phoneNetwork = normalizeConnectString(row.phone_network);
  const aisSubType = normalizeConnectString(row.ais_sub_type);
  const phoneSubType = normalizeConnectString(row.phone_sub_type);
  const phoneImei = normalizeConnectString(row.phone_imei);
  const phoneDateStart = normalizeConnectString(row.phone_date_start);
  const phoneDateEnd = normalizeConnectString(row.phone_date_end);
  const bankCode = normalizeConnectString(row.bank_code);
  const bankSubType = normalizeConnectString(row.bank_sub_type);
  const statementAccType = normalizeConnectString(row.statement_acc_type);
  const bankPromptPay = normalizeConnectString(row.bank_promptpay);
  const xxxAmount = normalizeConnectString(row.xxx_amount);
  const xxxDate = normalizeConnectString(row.xxx_date);
  const xxxTime = normalizeConnectString(row.xxx_time);
  const bankAccountName = normalizeConnectString(row.bank_account_name);
  const atmAccountNo = normalizeConnectString(row.atm_account_no);
  const atmDate = normalizeConnectString(row.atm_date);
  const atmTime = normalizeConnectString(row.atm_time);
  const atmLocation = normalizeConnectString(row.atm_location);
  const atmTerminalId = normalizeConnectString(row.atm_terminal_id);
  const bankDateStart = normalizeConnectString(row.bank_date_start);
  const bankDateEnd = normalizeConnectString(row.bank_date_end);
  const trueId = normalizeConnectString(row.true_id);
  const trueName = normalizeConnectString(row.true_name);
  const trueDateStart = normalizeConnectString(row.true_date_start);
  const trueDateEnd = normalizeConnectString(row.true_date_end);
  const taxType = normalizeConnectString(row.tax_type);
  const taxId = normalizeConnectString(row.tax_id);
  const taxName = normalizeConnectString(row.tax_name);
  const taxYearStart = normalizeConnectNumber(row.tax_year_start);
  const taxYearEnd = normalizeConnectNumber(row.tax_year_end);
  const phoneNumbers = normalizeConnectStringArray(parseConnectJsonArray(row.phone_numbers_json));
  const ipEntries = normalizeConnectObjectArray(parseConnectJsonArray(row.ip_entries_json));
  const bankAccounts = normalizeConnectObjectArray(parseConnectJsonArray(row.bank_accounts_json));
  const phoneNumbersText = normalizeConnectString(row.phone_numbers_text) || phoneNumbers.join("\n");
  const ipEntriesText = normalizeConnectString(row.ip_entries_text) || formatConnectIpEntriesText(ipEntries);
  const bankAccountsText = normalizeConnectString(row.bank_accounts_text) || formatConnectBankAccountsText(bankAccounts);
  return {
    requestId: normalizeConnectString(row.request_id),
    submittedAt: normalizeConnectString(row.submitted_at_iso),
    submittedDateTh: normalizeConnectString(row.submitted_date_th),
    submittedTimeTh: normalizeConnectString(row.submitted_time_th),
    updatedAt: normalizeConnectString(row.record_updated_at_iso),
    submittedByEmail: normalizeConnectString(row.submitted_by_email),
    submittedByName: normalizeConnectString(row.submitted_by_name),
    submittedByRole: normalizeConnectString(row.submitted_by_role),
    submittedByPosition: normalizeConnectString(row.submitted_by_position),
    submittedByUnit: normalizeConnectString(row.submitted_by_unit),
    submittedByPhone: normalizeConnectString(row.submitted_by_phone),
    type: normalizeConnectString(row.request_type),
    typeLabel: normalizeConnectString(row.request_type_label),
    unit: normalizeConnectString(row.unit),
    commander: normalizeConnectString(row.commander),
    docNumber: normalizeConnectString(row.doc_number),
    docNumberLine: normalizeConnectString(row.doc_number_line),
    casePresetKey: normalizeConnectString(row.case_preset_key),
    caseDetail: normalizeConnectString(row.case_detail),
    caseDetailCustom: normalizeConnectString(row.caseDetailCustom),
    caseSearchTerm: normalizeConnectString(row.caseSearchTerm),
    adminNoteTag: normalizeConnectString(row.admin_note_tag),
    adminNoteTagLabel: normalizeConnectString(row.admin_note_tag_label),
    adminNoteText: normalizeConnectString(row.admin_note_text),
    requestHistoryStatus: normalizeConnectString(row.request_history_status),
    requestHistoryLabel: normalizeConnectString(row.request_history_label),
    driveLink: normalizeConnectString(row.drive_link),
    filePassword: row.file_password == null ? "" : String(row.file_password),
    phoneNumbersText: phoneNumbersText,
    ipEntriesText: ipEntriesText,
    bankAccountsText: bankAccountsText,
    phoneNetwork: phoneNetwork,
    aisSubType: aisSubType,
    phoneSubType: phoneSubType,
    phoneImei: phoneImei,
    bankCode: bankCode,
    bankSubType: bankSubType,
    statementAccType: statementAccType,
    bankPromptPay: bankPromptPay,
    bankAccountName: bankAccountName,
    atmAccountNo: atmAccountNo,
    trueId: trueId,
    trueName: trueName,
    taxType: taxType,
    taxId: taxId,
    taxName: taxName,
    status: normalizeConnectStatus(row.status),
    statusNote: normalizeConnectString(row.status_note),
    fileBaseName: normalizeConnectString(row.file_base_name),
    pdfFileName: normalizeConnectString(row.pdf_file_name),
    generatedMode: normalizeConnectString(row.generated_mode),
    generatedPdfUrl: normalizeConnectString(row.generated_pdf_url),
    generatedDocxUrl: normalizeConnectString(row.generated_docx_url),
    phone: {
      network: phoneNetwork,
      aisSubType: aisSubType,
      subType: phoneSubType,
      imei: phoneImei,
      numbers: phoneNumbers,
      ipEntries: ipEntries,
      dateStart: phoneDateStart,
      dateEnd: phoneDateEnd
    },
    bank: {
      code: bankCode,
      subType: bankSubType,
      statementAccType: statementAccType,
      accounts: bankAccounts,
      promptPay: bankPromptPay,
      xxxAmount: xxxAmount,
      xxxDate: xxxDate,
      xxxTime: xxxTime,
      bankAccountName: bankAccountName,
      atmAccountNo: atmAccountNo,
      atmDate: atmDate,
      atmTime: atmTime,
      atmLocation: atmLocation,
      atmTerminalId: atmTerminalId,
      dateStart: bankDateStart,
      dateEnd: bankDateEnd
    },
    trueMoney: {
      id: trueId,
      name: trueName,
      dateStart: trueDateStart,
      dateEnd: trueDateEnd
    },
    tax: {
      type: taxType,
      id: taxId,
      name: taxName,
      yearStart: taxYearStart,
      yearEnd: taxYearEnd
    },
    statusUpdatedAt: normalizeConnectString(row.status_updated_at_iso),
    statusUpdatedByEmail: normalizeConnectString(row.status_updated_by_email),
    statusUpdatedByName: normalizeConnectString(row.status_updated_by_name),
    statusUpdatedByRole: normalizeConnectString(row.status_updated_by_role)
  };
}

function buildConnectRequestRow(payload, authUser, existingRow) {
  const compactPayload = buildCompactConnectPayload(payload);
  const submittedBy = payload && payload.submittedBy && typeof payload.submittedBy === "object" ? payload.submittedBy : {};
  const phone = payload && payload.phone && typeof payload.phone === "object" ? payload.phone : {};
  const bank = payload && payload.bank && typeof payload.bank === "object" ? payload.bank : {};
  const trueMoney = payload && (payload.trueMoney || payload.truemoney) && typeof (payload.trueMoney || payload.truemoney) === "object"
    ? (payload.trueMoney || payload.truemoney)
    : {};
  const tax = payload && payload.tax && typeof payload.tax === "object" ? payload.tax : {};
  const files = payload && payload.files && typeof payload.files === "object" ? payload.files : {};
  const phoneNumbers = normalizeConnectStringArray(phone.numbers);
  const ipEntries = normalizeConnectObjectArray(phone.ipEntries);
  const bankAccounts = normalizeConnectObjectArray(bank.accounts);
  const submittedAtIso = normalizeConnectString(payload && payload.submittedAt) || new Date().toISOString();
  const submittedAtDate = new Date(submittedAtIso);
  const validSubmittedAt = Number.isNaN(submittedAtDate.getTime()) ? new Date() : submittedAtDate;
  const docNumber = normalizeConnectString(payload && payload.docNumber);
  const status = normalizeConnectString(payload && payload.status) || (existingRow ? normalizeConnectString(existingRow.status) : "") || "pending";
  const statusNote = normalizeConnectString(payload && payload.statusNote) || (existingRow ? normalizeConnectString(existingRow.status_note) : "");
  const driveLink = normalizeConnectString(payload && payload.driveLink) || (existingRow ? normalizeConnectString(existingRow.drive_link) : "");
  const filePassword = payload && Object.prototype.hasOwnProperty.call(payload, "filePassword")
    ? String(payload.filePassword == null ? "" : payload.filePassword)
    : (existingRow && existingRow.file_password != null ? String(existingRow.file_password) : "");
  const rowMap = {
    request_id: normalizeConnectString(payload && payload.requestId),
    submitted_at_iso: submittedAtIso,
    submitted_date_th: formatThaiDateOnlyStandard(validSubmittedAt),
    submitted_time_th: formatTimeOnlyStandard(validSubmittedAt),
    record_updated_at_iso: new Date().toISOString(),
    submitted_by_email: normalizeConnectString(authUser && authUser.email) || normalizeConnectString(submittedBy.email),
    submitted_by_name: normalizeConnectString(submittedBy.name) || normalizeConnectString(authUser && authUser.name),
    submitted_by_role: normalizeConnectString(submittedBy.role) || normalizeConnectString(authUser && authUser.role),
    submitted_by_position: normalizeConnectString(submittedBy.position) || normalizeConnectString(authUser && authUser.position),
    submitted_by_unit: normalizeConnectString(submittedBy.unit) || normalizeConnectString(authUser && authUser.unit),
    submitted_by_phone: normalizeConnectString(submittedBy.phone) || normalizeConnectString(authUser && authUser.phone),
    request_type: normalizeConnectString(payload && payload.type),
    request_type_label: normalizeConnectString(payload && payload.typeLabel),
    unit: normalizeConnectString(payload && payload.unit),
    commander: normalizeConnectString(payload && payload.commander),
    doc_number: docNumber,
    doc_number_line: docNumber ? "ตช 0026.21/" + docNumber : "",
    case_preset_key: normalizeConnectString(payload && payload.casePresetKey),
    case_detail: normalizeConnectString(payload && payload.caseDetail),
    case_detail_custom: normalizeConnectString(payload && payload.caseDetailCustom),
    case_search_term: normalizeConnectString(payload && payload.caseSearchTerm),
    admin_note_tag: normalizeConnectString(payload && payload.adminNoteTag),
    admin_note_tag_label: normalizeConnectString(payload && payload.adminNoteTagLabel),
    admin_note_text: normalizeConnectString(payload && payload.adminNoteText),
    request_history_status: normalizeConnectString(payload && payload.requestHistoryStatus),
    request_history_label: normalizeConnectString(payload && payload.requestHistoryLabel),
    status: status,
    status_note: statusNote,
    drive_link: driveLink,
    file_password: filePassword,
    phone_network: normalizeConnectString(phone.network),
    ais_sub_type: normalizeConnectString(phone.aisSubType),
    phone_sub_type: normalizeConnectString(phone.subType),
    phone_imei: normalizeConnectString(phone.imei),
    phone_numbers_text: phoneNumbers.join("\n"),
    phone_numbers_json: safeJsonStringify(phoneNumbers),
    ip_entries_text: formatConnectIpEntriesText(ipEntries),
    ip_entries_json: safeJsonStringify(ipEntries),
    phone_date_start: normalizeConnectString(phone.dateStart),
    phone_date_end: normalizeConnectString(phone.dateEnd),
    bank_code: normalizeConnectString(bank.code),
    bank_sub_type: normalizeConnectString(bank.subType),
    statement_acc_type: normalizeConnectString(bank.statementAccType),
    bank_accounts_text: formatConnectBankAccountsText(bankAccounts),
    bank_accounts_json: safeJsonStringify(bankAccounts),
    bank_promptpay: normalizeConnectString(bank.promptPay),
    xxx_amount: normalizeConnectString(bank.xxxAmount),
    xxx_date: normalizeConnectString(bank.xxxDate),
    xxx_time: normalizeConnectString(bank.xxxTime),
    bank_account_name: normalizeConnectString(bank.bankAccountName),
    atm_account_no: normalizeConnectString(bank.atmAccountNo),
    atm_date: normalizeConnectString(bank.atmDate),
    atm_time: normalizeConnectString(bank.atmTime),
    atm_location: normalizeConnectString(bank.atmLocation),
    atm_terminal_id: normalizeConnectString(bank.atmTerminalId),
    bank_date_start: normalizeConnectString(bank.dateStart),
    bank_date_end: normalizeConnectString(bank.dateEnd),
    true_id: normalizeConnectString(trueMoney.id),
    true_name: normalizeConnectString(trueMoney.name),
    true_date_start: normalizeConnectString(trueMoney.dateStart),
    true_date_end: normalizeConnectString(trueMoney.dateEnd),
    tax_type: normalizeConnectString(tax.type),
    tax_id: normalizeConnectString(tax.id),
    tax_name: normalizeConnectString(tax.name),
    tax_year_start: normalizeConnectString(tax.yearStart),
    tax_year_end: normalizeConnectString(tax.yearEnd),
    file_base_name: normalizeConnectString(files.fileBaseName),
    pdf_file_name: normalizeConnectString(files.pdfFileName),
    generated_mode: normalizeConnectString(files.generatedMode),
    generated_pdf_url: normalizeConnectString(files.generatedPdfUrl),
    generated_docx_url: normalizeConnectString(files.generatedDocxUrl),
    generated_warnings_json: safeJsonStringify(Array.isArray(files.generatedWarnings) ? files.generatedWarnings : []),
    raw_state_json: safeJsonStringify(compactPayload && compactPayload.rawState),
    payload_json: safeJsonStringify(compactPayload),
    status_updated_at_iso: existingRow ? normalizeConnectString(existingRow.status_updated_at_iso) : "",
    status_updated_by_email: existingRow ? normalizeConnectString(existingRow.status_updated_by_email) : "",
    status_updated_by_name: existingRow ? normalizeConnectString(existingRow.status_updated_by_name) : "",
    status_updated_by_role: existingRow ? normalizeConnectString(existingRow.status_updated_by_role) : ""
  };

  return CONNECT_REQUEST_HEADERS.map(function(header) {
    return rowMap.hasOwnProperty(header) ? rowMap[header] : "";
  });
}

function writeConnectRequestRow(sheet, rowValues, existingRowIndex) {
  const targetSheet = ensureConnectRequestSheet(sheet || getConnectRequestSheet());
  if (!targetSheet) {
    throw new Error("ไม่พบชีทสำหรับบันทึกคำขอ CSD1 Connect");
  }
  if (!Array.isArray(rowValues) || rowValues.length !== CONNECT_REQUEST_HEADERS.length) {
    throw new Error("ข้อมูลคำขอไม่ครบตามโครงสร้างชีท");
  }

  const rowIndex = Number(existingRowIndex || 0) > 1
    ? Number(existingRowIndex)
    : targetSheet.getLastRow() + 1;
  const sanitized = sanitizeRowValuesForSheet(CONNECT_REQUEST_HEADERS, rowValues);
  if (sanitized.truncatedHeaders.length) {
    Logger.log("writeConnectRequestRow: truncated oversized fields => " + sanitized.truncatedHeaders.join(", "));
  }
  const writeRange = targetSheet.getRange(rowIndex, 1, 1, CONNECT_REQUEST_HEADERS.length);
  writeRange.setNumberFormat("@");
  writeRange.setValues([sanitized.rowValues]);
  SpreadsheetApp.flush();
  return rowIndex;
}

function probeConnectRequestSheetWrite() {
  const sheet = ensureConnectRequestSheet(getConnectRequestSheet());
  const spreadsheetInfo = getActiveSpreadsheetInfo();
  const requestId = "REQ-PROBE-" + Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMMdd-HHmmss");
  const payload = {
    requestId: requestId,
    submittedAt: new Date().toISOString(),
    submittedBy: {
      email: "probe@local.test",
      name: "Probe Write",
      role: "admin",
      position: "debug",
      unit: "SYSTEM",
      phone: ""
    },
    type: "bank",
    typeLabel: "Probe",
    unit: "SYSTEM",
    commander: "SYSTEM",
    docNumber: "9999",
    casePresetKey: "probe",
    caseDetail: "ทดสอบการเขียนชีทโดยตรงจาก Apps Script",
    caseDetailCustom: "",
    caseSearchTerm: "probe",
    adminNoteTag: "normal",
    adminNoteTagLabel: "งานทั่วไป",
    adminNoteText: "probeConnectRequestSheetWrite",
    requestHistoryStatus: "new",
    requestHistoryLabel: "ยังไม่เคยขอ",
    status: "pending",
    statusNote: "",
    driveLink: "",
    filePassword: "",
    phone: {},
    bank: {},
    trueMoney: {},
    tax: {},
    files: {
      fileBaseName: requestId,
      pdfFileName: requestId + ".pdf",
      generatedMode: "probe",
      generatedPdfUrl: "",
      generatedDocxUrl: "",
      generatedWarnings: []
    },
    rawState: {
      requestId: requestId,
      unit: "SYSTEM",
      type: "bank",
      step: 4
    }
  };

  const row = buildConnectRequestRow(payload, payload.submittedBy, null);
  const rowIndex = writeConnectRequestRow(sheet, row, 0);
  const verified = verifyConnectRequestStored(sheet, requestId);
  const result = {
    ok: !!(verified && verified.row),
    requestId: requestId,
    rowIndex: Number(verified && verified.rowIndex || rowIndex || 0),
    sheetName: CONNECT_REQUEST_SHEET_NAME,
    spreadsheetId: spreadsheetInfo.spreadsheetId,
    spreadsheetName: spreadsheetInfo.spreadsheetName
  };
  Logger.log("probeConnectRequestSheetWrite => " + JSON.stringify(result));
  return result;
}

function verifyConnectRequestStored(sheet, requestId) {
  const cleanRequestId = normalizeConnectString(requestId);
  if (!cleanRequestId) return null;
  return findConnectRequestById(sheet || getConnectRequestSheet(), cleanRequestId);
}

function getActiveSpreadsheetInfo() {
  let ss = null;
  try {
    ss = getConnectSpreadsheet();
  } catch (_) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  if (!ss) {
    return {
      spreadsheetId: "",
      spreadsheetName: ""
    };
  }
  return {
    spreadsheetId: normalizeConnectString(ss.getId()),
    spreadsheetName: normalizeConnectString(ss.getName())
  };
}

function handleGetConnectBackendInfo(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) {
    return jsonResponse({ status: "error", message: auth.message });
  }

  const spreadsheetInfo = getActiveSpreadsheetInfo();
  const configuredSpreadsheetId = normalizeConnectString(
    PropertiesService.getScriptProperties().getProperty(CONNECT_SPREADSHEET_ID_PROPERTY_KEY)
  );

  let sheet = null;
  let sheetReady = false;
  let sheetLastRow = 0;
  let errorMessage = "";

  try {
    sheet = ensureConnectRequestSheet(getConnectRequestSheet());
    sheetReady = !!sheet;
    sheetLastRow = sheet ? Number(sheet.getLastRow() || 0) : 0;
  } catch (error) {
    errorMessage = error && error.message ? error.message : String(error);
  }

  return jsonResponse({
    status: "success",
    backendVersion: CONNECT_BACKEND_VERSION,
    spreadsheetId: spreadsheetInfo.spreadsheetId,
    spreadsheetName: spreadsheetInfo.spreadsheetName,
    configuredSpreadsheetId: configuredSpreadsheetId,
    sheetName: CONNECT_REQUEST_SHEET_NAME,
    sheetReady: sheetReady,
    sheetLastRow: sheetLastRow,
    errorMessage: errorMessage
  });
}

/**
 * รับ PDF base64 → บันทึกลง Drive → คืน URL → อัปเดต generated_pdf_url ในชีท
 * มี Lock ป้องกัน concurrent, ตรวจ file integrity, ยืนยัน URL ใน Sheet
 */
function handleUploadSummonPdf(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) return jsonResponse({ status: "error", message: auth.message });

  const requestId   = normalizeConnectString(data && data.requestId);
  const pdfFileName = normalizeConnectString(data && data.pdfFileName) || (requestId + ".pdf");
  const pdfBase64   = data && data.pdfBase64 ? String(data.pdfBase64) : "";

  if (!requestId) return jsonResponse({ status: "error", message: "ไม่พบ requestId" });
  if (!pdfBase64) return jsonResponse({ status: "error", message: "ไม่พบข้อมูล PDF (base64)" });

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockError) {
    return jsonResponse({ status: "error", message: "ระบบกำลังประมวลผลคำขออื่นอยู่ กรุณาลองใหม่อีกครั้ง" });
  }

  try {
    var reqFolder = getOrCreateConnectSummonRequestFolder(requestId);
    if (!reqFolder) {
      return jsonResponse({ status: "error", message: "ไม่สามารถเข้าถึงโฟลเดอร์จัดเก็บหมายเรียกได้" });
    }

    var existingIter = reqFolder.getFilesByName(pdfFileName);
    while (existingIter.hasNext()) { existingIter.next().setTrashed(true); }

    var pdfBytes = Utilities.base64Decode(pdfBase64);
    var expectedSize = pdfBytes.length;
    var blob = Utilities.newBlob(pdfBytes, "application/pdf", pdfFileName);
    var file = reqFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // ตรวจสอบขนาดไฟล์ที่สร้างจริง
    var actualSize = file.getSize();
    if (actualSize <= 0) {
      file.setTrashed(true);
      return jsonResponse({ status: "error", message: "ไฟล์ที่อัปโหลดมีขนาด 0 ไบต์ อาจเกิดปัญหาระหว่างการบันทึก" });
    }
    if (expectedSize > 0 && Math.abs(actualSize - expectedSize) > 1024) {
      Logger.log("handleUploadSummonPdf: size mismatch expected=" + expectedSize + " actual=" + actualSize);
    }

    var pdfUrl = normalizeConnectString(file.getUrl());
    if (!pdfUrl) {
      file.setTrashed(true);
      return jsonResponse({ status: "error", message: "ไม่สามารถสร้างลิงก์สำหรับไฟล์ที่อัปโหลดได้" });
    }

    // อัปเดต generated_pdf_url + pdf_file_name + record_updated_at_iso ในชีท
    var sheet = ensureConnectRequestSheet(getConnectRequestSheet());
    var found = findConnectRequestById(sheet, requestId);
    var sheetUpdated = false;
    if (found && found.rowIndex > 0) {
      var pdfUrlCol      = CONNECT_REQUEST_HEADERS.indexOf("generated_pdf_url") + 1;
      var updatedAtCol   = CONNECT_REQUEST_HEADERS.indexOf("record_updated_at_iso") + 1;
      var pdfFileNameCol = CONNECT_REQUEST_HEADERS.indexOf("pdf_file_name") + 1;
      var nowIso = new Date().toISOString();
      if (pdfUrlCol > 0)      sheet.getRange(found.rowIndex, pdfUrlCol).setValue(pdfUrl);
      if (updatedAtCol > 0)   sheet.getRange(found.rowIndex, updatedAtCol).setValue(nowIso);
      if (pdfFileNameCol > 0) sheet.getRange(found.rowIndex, pdfFileNameCol).setValue(pdfFileName);
      SpreadsheetApp.flush();

      // ยืนยันว่า URL ถูกเขียนลง Sheet จริง
      if (pdfUrlCol > 0) {
        var writtenUrl = normalizeConnectString(sheet.getRange(found.rowIndex, pdfUrlCol).getValue());
        if (writtenUrl !== pdfUrl) {
          Logger.log("handleUploadSummonPdf: URL verification failed. written=" + writtenUrl + " expected=" + pdfUrl);
          return jsonResponse({
            status: "error",
            message: "บันทึกลิงก์ PDF ลง Google Sheet ไม่สำเร็จ (ตรวจสอบแล้วค่าไม่ตรง)",
            pdfUrl: pdfUrl,
            requestId: requestId
          });
        }
        sheetUpdated = true;
      }
    }

    return jsonResponse({
      status: "success",
      pdfUrl: pdfUrl,
      pdfFileSize: actualSize,
      pdfFileName: pdfFileName,
      requestId: requestId,
      sheetUpdated: sheetUpdated
    });
  } catch (error) {
    Logger.log("handleUploadSummonPdf: " + error);
    return jsonResponse({
      status: "error",
      message: "เกิดข้อผิดพลาดระหว่างอัปโหลด PDF: " + (error && error.message ? error.message : String(error)),
      requestId: requestId
    });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function handleSubmitConnectRequest(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) {
    return jsonResponse({ status: "error", message: auth.message });
  }

  const connectAuthUser = auth.user || {};
  if (!canSubmitConnect(connectAuthUser.role)) {
    return jsonResponse({ status: "error", message: "บัญชีผู้ใช้นี้ไม่มีสิทธิ์ส่งคำขอ CSD1 Connect" });
  }

  const payload = data && data.payload && typeof data.payload === "object" ? data.payload : null;
  if (!payload) {
    return jsonResponse({ status: "error", message: "ไม่พบ payload สำหรับคำขอ CSD1 Connect" });
  }

  const requestId = normalizeConnectString(payload.requestId);
  if (!requestId) {
    return jsonResponse({ status: "error", message: "ไม่พบ requestId" });
  }

  const existingGeneratedPdfUrl = normalizeConnectString(
    payload && payload.files && payload.files.generatedPdfUrl
  );
  const hasClientPdfBase64 = !!normalizeConnectString(
    payload && payload.files && payload.files.pdfBase64
  );

  // อัปโหลด PDF ไป Drive ถ้า client ส่ง base64 มาด้วย (atomic: upload + save ใน call เดียว)
  var pdfUrl = "";
  var pdfFileSize = 0;
  var files = payload.files && typeof payload.files === "object" ? payload.files : {};
  var pdfBase64 = normalizeConnectString(files.pdfBase64);
  var pdfFileName = normalizeConnectString(files.pdfFileName) || (requestId + ".pdf");
  if (pdfBase64) {
    var uploadLock = LockService.getScriptLock();
    try {
      uploadLock.waitLock(30000);
    } catch (lockErr) {
      return jsonResponse({ status: "error", message: "ระบบกำลังประมวลผลคำขออื่นอยู่ กรุณาลองใหม่อีกครั้ง" });
    }
    try {
      var reqFolder = getOrCreateConnectSummonRequestFolder(requestId);
      if (!reqFolder) {
        throw new Error("ไม่สามารถเข้าถึงโฟลเดอร์จัดเก็บหมายเรียกได้");
      }
      var existingIter = reqFolder.getFilesByName(pdfFileName);
      while (existingIter.hasNext()) { existingIter.next().setTrashed(true); }
      var pdfBytes = Utilities.base64Decode(pdfBase64);
      var expectedSize = pdfBytes.length;
      var blob = Utilities.newBlob(pdfBytes, "application/pdf", pdfFileName);
      var file = reqFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      // ตรวจ file integrity
      pdfFileSize = file.getSize();
      if (pdfFileSize <= 0) {
        file.setTrashed(true);
        throw new Error("ไฟล์ที่อัปโหลดมีขนาด 0 ไบต์");
      }
      if (expectedSize > 0 && Math.abs(pdfFileSize - expectedSize) > 1024) {
        Logger.log("handleSubmitConnectRequest: PDF size mismatch expected=" + expectedSize + " actual=" + pdfFileSize);
      }

      pdfUrl = normalizeConnectString(file.getUrl());
      if (!pdfUrl) {
        file.setTrashed(true);
        throw new Error("ไม่สามารถสร้างลิงก์สำหรับไฟล์ที่อัปโหลดได้");
      }

      // อัปเดต payload.files.generatedPdfUrl เพื่อให้ buildConnectRequestRow บันทึก URL
      payload.files.generatedPdfUrl = pdfUrl;
    } catch (uploadErr) {
      Logger.log("handleSubmitConnectRequest: Drive upload failed: " + uploadErr);
      return jsonResponse({
        status: "error",
        message: "ไม่สามารถบันทึกไฟล์หมายเรียกลง Drive ได้: " + (uploadErr && uploadErr.message ? uploadErr.message : String(uploadErr)),
      });
    } finally {
      try { uploadLock.releaseLock(); } catch (_) {}
    }
    // ลบ base64 ออกจาก payload ก่อน save เพื่อไม่ให้ชีทเก็บข้อมูลขนาดใหญ่
    delete payload.files.pdfBase64;
  }

  const resolvedGeneratedPdfUrl = normalizeConnectString(payload.files && payload.files.generatedPdfUrl);
  if (!resolvedGeneratedPdfUrl && !existingGeneratedPdfUrl && !pdfUrl) {
    return jsonResponse({
      status: "error",
      message: hasClientPdfBase64
        ? "ระบบไม่พบ URL ไฟล์หมายเรียกหลังอัปโหลด จึงยังไม่บันทึกคำขอ"
        : "คำขอนี้ยังไม่มีไฟล์หมายเรียกถาวร ระบบจึงยังไม่บันทึกเพื่อป้องกันปัญหาการดาวน์โหลดในภายหลัง",
    });
  }

  const spreadsheetInfo = getActiveSpreadsheetInfo();
  const sheet = ensureConnectRequestSheet(getConnectRequestSheet());
  const found = findConnectRequestById(sheet, requestId);
  const row = buildConnectRequestRow(payload, auth.user, found ? found.row : null);
  let storedRowIndex = 0;
  let mode = "inserted";

  try {
    storedRowIndex = writeConnectRequestRow(sheet, row, found && found.rowIndex > 0 ? found.rowIndex : 0);
    mode = found && found.rowIndex > 0 ? "updated" : "inserted";
  } catch (writeError) {
    Logger.log("handleSubmitConnectRequest: sheet write failed: " + writeError);
    return jsonResponse({
      status: "error",
      message: "บันทึกข้อมูลคำขอลง Google Sheet ไม่สำเร็จ",
      requestId: requestId,
      pdfUrl: pdfUrl,
      sheetName: CONNECT_REQUEST_SHEET_NAME,
      spreadsheetId: spreadsheetInfo.spreadsheetId,
      spreadsheetName: spreadsheetInfo.spreadsheetName
    });
  }

  const verified = verifyConnectRequestStored(sheet, requestId);
  if (!verified || !verified.row) {
    Logger.log("handleSubmitConnectRequest: verification failed for requestId=" + requestId + " rowIndex=" + storedRowIndex);
    return jsonResponse({
      status: "error",
      message: "บันทึกไฟล์ลง Google Drive แล้ว แต่ยืนยันการบันทึกข้อมูลใน Google Sheet ไม่สำเร็จ",
      requestId: requestId,
      pdfUrl: pdfUrl,
      rowIndex: storedRowIndex,
      sheetName: CONNECT_REQUEST_SHEET_NAME,
      spreadsheetId: spreadsheetInfo.spreadsheetId,
      spreadsheetName: spreadsheetInfo.spreadsheetName
    });
  }

  // ยืนยันว่า generated_pdf_url ถูกบันทึกจริง — ป้องกัน link หลุดจาก Sheet
  var finalPdfUrl = pdfUrl || normalizeConnectString(verified.row.generated_pdf_url);
  if (pdfUrl && !normalizeConnectString(verified.row.generated_pdf_url)) {
    // PDF ถูกอัปโหลดแล้วแต่ URL ไม่ถูกบันทึกในชีท — แก้ไขทันที
    var pdfUrlCol = CONNECT_REQUEST_HEADERS.indexOf("generated_pdf_url") + 1;
    if (pdfUrlCol > 0 && verified.rowIndex > 0) {
      sheet.getRange(verified.rowIndex, pdfUrlCol).setValue(pdfUrl);
      SpreadsheetApp.flush();
      Logger.log("handleSubmitConnectRequest: repaired missing generated_pdf_url for requestId=" + requestId);
    }
  }

  return jsonResponse({
    status: "success",
    requestId: requestId,
    rowIndex: Number(verified.rowIndex || storedRowIndex || 0),
    mode: mode,
    sheetName: CONNECT_REQUEST_SHEET_NAME,
    spreadsheetId: spreadsheetInfo.spreadsheetId,
    spreadsheetName: spreadsheetInfo.spreadsheetName,
    pdfUrl: finalPdfUrl,
    pdfFileSize: pdfFileSize || 0
  });
}

function handleGetConnectRequests(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) {
    return jsonResponse({ status: "error", message: auth.message });
  }

  const authUser = auth.user || {};
  const isAdmin = isConnectAdminRole(authUser.role);
  const authUnitKey = normalizeUnitKey(authUser.unit || "");
  const authEmail = normalizeEmail(authUser.email || "");
  const limitRaw = Number(data && data.limit);
  const limit = isFinite(limitRaw) && limitRaw > 0 ? Math.floor(limitRaw) : 0;

  let rows = readConnectRequestRows();
  if (!isAdmin) {
    rows = rows.filter(function(row) {
      return canAccessConnectRequestRow(authUser, row);
    });
  }

  rows.sort(function(a, b) {
    const aTs = String(a.record_updated_at_iso || a.submitted_at_iso || "");
    const bTs = String(b.record_updated_at_iso || b.submitted_at_iso || "");
    if (aTs === bTs) {
      return String(b.request_id || "").localeCompare(String(a.request_id || ""));
    }
    return bTs.localeCompare(aTs);
  });

  if (limit > 0) rows = rows.slice(0, limit);

  const sheet = ensureConnectRequestSheet(getConnectRequestSheet());
  rows.forEach(function(row) {
    if (!row) return;
    ensureConnectRequestGeneratedPdfUrl(sheet, row, Number(row.__rowIndex || 0));
  });

  return jsonResponse({
    status: "success",
    total: rows.length,
    isAdmin: isAdmin,
    data: rows.map(toConnectRequestResponse)
  });
}

function handleDownloadConnectResponseFile(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) {
    return jsonResponse({ status: "error", message: auth.message });
  }

  const requestId = normalizeConnectString(data && data.requestId);
  if (!requestId) {
    return jsonResponse({ status: "error", message: "ไม่พบ requestId" });
  }

  const sheet = ensureConnectRequestSheet(getConnectRequestSheet());
  const found = findConnectRequestById(sheet, requestId);
  if (!found || !found.row) {
    return jsonResponse({ status: "error", message: "ไม่พบคำขอที่ต้องการดาวน์โหลด" });
  }
  if (!canAccessConnectRequestRow(auth.user, found.row)) {
    return jsonResponse({ status: "error", message: "คุณไม่มีสิทธิ์ดาวน์โหลดไฟล์รายการนี้" });
  }

  const driveLink = normalizeConnectString(found.row.drive_link);
  if (!driveLink) {
    return jsonResponse({ status: "error", message: "รายการนี้ยังไม่มีไฟล์ตอบกลับ" });
  }

  try {
    const payload = getConnectDownloadPayloadFromLink(driveLink);
    return jsonResponse({
      status: "success",
      requestId: requestId,
      driveLink: driveLink,
      sourceFileName: payload.sourceFileName,
      mimeType: payload.mimeType,
      dataBase64: payload.dataBase64
    });
  } catch (error) {
    return jsonResponse({ status: "error", message: error && error.message ? error.message : String(error) });
  }
}

function handleDownloadSummonPdf(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) {
    return jsonResponse({ status: "error", message: auth.message });
  }

  const requestId = normalizeConnectString(data && data.requestId);
  if (!requestId) {
    return jsonResponse({ status: "error", message: "ไม่พบ requestId" });
  }

  const sheet = ensureConnectRequestSheet(getConnectRequestSheet());
  const found = findConnectRequestById(sheet, requestId);
  if (!found || !found.row) {
    return jsonResponse({ status: "error", message: "ไม่พบคำขอที่ต้องการดาวน์โหลด" });
  }
  if (!canAccessConnectRequestRow(auth.user, found.row)) {
    return jsonResponse({ status: "error", message: "คุณไม่มีสิทธิ์ดาวน์โหลดไฟล์รายการนี้" });
  }

  let pdfUrl = normalizeConnectString(found.row.generated_pdf_url);
  let fallbackFile = null;
  try {
    fallbackFile = findConnectSummonPdfFileByRequestId(
      requestId,
      normalizeConnectString(found.row.pdf_file_name)
    );
  } catch (lookupError) {
    Logger.log("handleDownloadSummonPdf: fallback lookup failed: " + lookupError);
  }

  if (fallbackFile) {
    try {
      fallbackFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const directPayload = getConnectDownloadPayloadFromDriveFile(fallbackFile);
      const refreshedUrl = normalizeConnectString(fallbackFile.getUrl());
      if (refreshedUrl && refreshedUrl !== pdfUrl) {
        const pdfUrlCol = CONNECT_REQUEST_HEADERS.indexOf("generated_pdf_url") + 1;
        const updatedAtCol = CONNECT_REQUEST_HEADERS.indexOf("record_updated_at_iso") + 1;
        if (pdfUrlCol > 0) sheet.getRange(found.rowIndex, pdfUrlCol).setValue(refreshedUrl);
        if (updatedAtCol > 0) sheet.getRange(found.rowIndex, updatedAtCol).setValue(new Date().toISOString());
        pdfUrl = refreshedUrl;
      }
      return jsonResponse({
        status: "success",
        requestId: requestId,
        sourceFileName: directPayload.sourceFileName,
        mimeType: directPayload.mimeType,
        dataBase64: directPayload.dataBase64
      });
    } catch (directFileError) {
      Logger.log("handleDownloadSummonPdf: direct Drive file download failed: " + directFileError);
    }
  }

  if (!pdfUrl) {
    try {
      pdfUrl = ensureConnectRequestGeneratedPdfUrl(sheet, found.row, found.rowIndex);
    } catch (lookupError) {
      Logger.log("handleDownloadSummonPdf: generated_pdf_url repair failed: " + lookupError);
    }
  }
  if (!pdfUrl) {
    return jsonResponse({ status: "error", message: "รายการนี้ยังไม่มีไฟล์หมายเรียก" });
  }

  try {
    const payload = getConnectDownloadPayloadFromLink(pdfUrl);
    return jsonResponse({
      status: "success",
      requestId: requestId,
      sourceFileName: payload.sourceFileName,
      mimeType: payload.mimeType,
      dataBase64: payload.dataBase64
    });
  } catch (error) {
    try {
      fallbackFile = fallbackFile || findConnectSummonPdfFileByRequestId(
        requestId,
        normalizeConnectString(found.row.pdf_file_name)
      );
      if (fallbackFile) {
        fallbackFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const refreshedUrl = normalizeConnectString(fallbackFile.getUrl());
        if (refreshedUrl && refreshedUrl !== pdfUrl) {
          const pdfUrlCol = CONNECT_REQUEST_HEADERS.indexOf("generated_pdf_url") + 1;
          const updatedAtCol = CONNECT_REQUEST_HEADERS.indexOf("record_updated_at_iso") + 1;
          if (pdfUrlCol > 0) sheet.getRange(found.rowIndex, pdfUrlCol).setValue(refreshedUrl);
          if (updatedAtCol > 0) sheet.getRange(found.rowIndex, updatedAtCol).setValue(new Date().toISOString());
        }
        const payload = getConnectDownloadPayloadFromDriveFile(fallbackFile);
        return jsonResponse({
          status: "success",
          requestId: requestId,
          sourceFileName: payload.sourceFileName,
          mimeType: payload.mimeType,
          dataBase64: payload.dataBase64
        });
      }
    } catch (fallbackError) {
      Logger.log("handleDownloadSummonPdf: fallback download failed: " + fallbackError);
    }
    return jsonResponse({ status: "error", message: error && error.message ? error.message : String(error) });
  }
}

function handleGetConnectRequestRawState(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) {
    return jsonResponse({ status: "error", message: auth.message });
  }

  const requestId = normalizeConnectString(data && data.requestId);
  if (!requestId) {
    return jsonResponse({ status: "error", message: "ไม่พบ requestId" });
  }

  const sheet = ensureConnectRequestSheet(getConnectRequestSheet());
  const found = findConnectRequestById(sheet, requestId);
  if (!found || !found.row) {
    return jsonResponse({ status: "error", message: "ไม่พบคำขอที่ต้องการ" });
  }
  if (!canAccessConnectRequestRow(auth.user, found.row)) {
    return jsonResponse({ status: "error", message: "คุณไม่มีสิทธิ์เข้าถึงรายการนี้" });
  }

  let rawStateJson = normalizeConnectString(found.row.raw_state_json);

  // fallback: ดึง rawState จาก payload_json
  if (!rawStateJson) {
    const payloadJson = normalizeConnectString(found.row.payload_json);
    if (payloadJson) {
      try {
        const payload = JSON.parse(payloadJson);
        if (payload && payload.rawState) {
          rawStateJson = JSON.stringify(payload.rawState);
        }
      } catch (_) {}
    }
  }

  if (!rawStateJson) {
    return jsonResponse({ status: "error", message: "ไม่พบข้อมูล state สำหรับคำขอนี้ (raw_state_json และ payload_json ว่างเปล่า)" });
  }

  return jsonResponse({
    status: "success",
    requestId: requestId,
    rawStateJson: rawStateJson,
  });
}

function handleGetConnectDocIndex(data) {
  var auth = authorizeRequest(data);
  if (!auth.ok) return jsonResponse({ status: "error", message: auth.message });

  var sheet = ensureConnectRequestSheet(getConnectRequestSheet());
  if (!sheet || sheet.getLastRow() < 2) return jsonResponse({ status: "success", items: [] });

  var lastRow = sheet.getLastRow();
  var numRows = lastRow - 1;
  var values = sheet.getRange(2, 1, numRows, CONNECT_REQUEST_HEADERS.length).getValues();

  var idxRequestId  = CONNECT_REQUEST_HEADERS.indexOf("request_id");
  var idxSubmitted  = CONNECT_REQUEST_HEADERS.indexOf("submitted_at_iso");
  var idxDocNumber  = CONNECT_REQUEST_HEADERS.indexOf("doc_number");
  var idxRawState   = CONNECT_REQUEST_HEADERS.indexOf("raw_state_json");

  var items = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var docNum = String(row[idxDocNumber] || "").trim();
    if (!docNum) continue;

    var locked = {};
    var rawStr = String(row[idxRawState] || "");
    if (rawStr) {
      try {
        var raw = JSON.parse(rawStr);
        locked = {
          type:             raw.type            || "",
          investigationType:raw.investigationType|| "",
          phoneNetwork:     raw.phoneNetwork     || "",
          aisSubType:       raw.aisSubType       || "",
          phoneSubType:     raw.phoneSubType     || "",
          phoneImei:        raw.phoneImei        || "",
          phoneNumbers:     Array.isArray(raw.phoneNumbers)     ? raw.phoneNumbers     : [],
          ipEntries:        Array.isArray(raw.ipEntries)        ? raw.ipEntries        : [],
          bankCode:         raw.bankCode         || "",
          bankSubType:      raw.bankSubType      || "",
          statementAccType: raw.statementAccType || "",
          bankAccounts:     Array.isArray(raw.bankAccounts)     ? raw.bankAccounts     : [],
          bankAccountNames: Array.isArray(raw.bankAccountNames) ? raw.bankAccountNames : [],
          bankPromptPay:    raw.bankPromptPay    || "",
          bankAccountName:  raw.bankAccountName  || "",
          atmAccountNo:     raw.atmAccountNo     || "",
          atmDate:          raw.atmDate          || "",
          atmTime:          raw.atmTime          || "",
          atmLocation:      raw.atmLocation      || "",
          atmTerminalId:    raw.atmTerminalId    || "",
          xxxAmount:        raw.xxxAmount        || "",
          xxxDate:          raw.xxxDate          || "",
          xxxTime:          raw.xxxTime          || "",
          xxxAccNo:         raw.xxxAccNo         || "",
          xxxRef:           raw.xxxRef           || "",
          xxxRole:          raw.xxxRole          || "",
          trueId:           raw.trueId           || "",
          trueName:         raw.trueName         || "",
        };
      } catch(_) {}
    }

    var item = { docNum: docNum.padStart(4, "0"), requestId: String(row[idxRequestId] || ""), submittedAt: String(row[idxSubmitted] || "") };
    var keys = Object.keys(locked);
    for (var k = 0; k < keys.length; k++) item[keys[k]] = locked[keys[k]];
    items.push(item);
  }

  return jsonResponse({ status: "success", items: items });
}

function handleLookupConnectDocNumber(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) {
    return jsonResponse({ status: "error", message: auth.message });
  }

  const docNumber = normalizeConnectString(data && data.docNumber);
  if (!docNumber || !/^\d{1,4}$/.test(docNumber)) {
    return jsonResponse({ status: "error", message: "เลขหนังสือไม่ถูกต้อง" });
  }

  const paddedDoc = docNumber.padStart(4, "0");
  const sheet = ensureConnectRequestSheet(getConnectRequestSheet());
  if (!sheet || sheet.getLastRow() < 2) {
    return jsonResponse({ status: "success", found: false });
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, CONNECT_REQUEST_HEADERS.length).getValues();
  for (var i = 0; i < values.length; i++) {
    const row = {};
    for (var j = 0; j < CONNECT_REQUEST_HEADERS.length; j++) {
      row[CONNECT_REQUEST_HEADERS[j]] = values[i][j] != null ? String(values[i][j]) : "";
    }
    const rowDoc = String(row.doc_number || "").trim().padStart(4, "0");
    if (rowDoc === paddedDoc) {
      return jsonResponse({
        status: "success",
        found: true,
        requestId: row.request_id || "",
        submittedAt: row.submitted_at_iso || "",
        submittedDateTh: row.submitted_date_th || "",
      });
    }
  }

  return jsonResponse({ status: "success", found: false });
}

function handleUpdateConnectRequestStatus(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) {
    return jsonResponse({ status: "error", message: auth.message });
  }

  const authUser = auth.user || {};
  if (!isConnectAdminRole(authUser.role)) {
    return jsonResponse({ status: "error", message: "เมนูนี้สำหรับผู้ดูแลระบบเท่านั้น" });
  }

  const requestId = normalizeConnectString(data && data.requestId);
  if (!requestId) {
    return jsonResponse({ status: "error", message: "ไม่พบ requestId" });
  }

  const nextStatus = normalizeConnectStatus(data && data.nextStatus);
  const nextStatusNote = normalizeConnectString(data && data.statusNote);
  const sheet = ensureConnectRequestSheet(getConnectRequestSheet());
  const found = findConnectRequestById(sheet, requestId);
  if (!found || !found.rowIndex) {
    return jsonResponse({ status: "error", message: "ไม่พบคำขอที่ต้องการอัปเดต" });
  }

  const statusCol = CONNECT_REQUEST_HEADERS.indexOf("status") + 1;
  const statusNoteCol = CONNECT_REQUEST_HEADERS.indexOf("status_note") + 1;
  const driveLinkCol = CONNECT_REQUEST_HEADERS.indexOf("drive_link") + 1;
  const filePasswordCol = CONNECT_REQUEST_HEADERS.indexOf("file_password") + 1;
  const updatedAtCol = CONNECT_REQUEST_HEADERS.indexOf("record_updated_at_iso") + 1;
  const statusUpdatedAtCol = CONNECT_REQUEST_HEADERS.indexOf("status_updated_at_iso") + 1;
  const statusUpdatedByEmailCol = CONNECT_REQUEST_HEADERS.indexOf("status_updated_by_email") + 1;
  const statusUpdatedByNameCol = CONNECT_REQUEST_HEADERS.indexOf("status_updated_by_name") + 1;
  const statusUpdatedByRoleCol = CONNECT_REQUEST_HEADERS.indexOf("status_updated_by_role") + 1;
  if (
    statusCol <= 0 ||
    statusNoteCol <= 0 ||
    driveLinkCol <= 0 ||
    filePasswordCol <= 0 ||
    updatedAtCol <= 0 ||
    statusUpdatedAtCol <= 0 ||
    statusUpdatedByEmailCol <= 0 ||
    statusUpdatedByNameCol <= 0 ||
    statusUpdatedByRoleCol <= 0
  ) {
    return jsonResponse({ status: "error", message: "ไม่พบคอลัมน์สถานะในชีต" });
  }

  const updatedAt = new Date().toISOString();
  const updatedByEmail = normalizeConnectString(authUser.email);
  const updatedByName = normalizeConnectString(authUser.name);
  const updatedByRole = normalizeConnectString(authUser.role);
  const driveLink = normalizeConnectString(data && data.driveLink);
  const filePassword = data && Object.prototype.hasOwnProperty.call(data, "filePassword")
    ? String(data.filePassword == null ? "" : data.filePassword)
    : "";
  sheet.getRange(found.rowIndex, statusCol).setValue(nextStatus);
  sheet.getRange(found.rowIndex, statusNoteCol).setValue(nextStatusNote);
  sheet.getRange(found.rowIndex, driveLinkCol).setValue(driveLink);
  sheet.getRange(found.rowIndex, filePasswordCol).setValue(filePassword);
  sheet.getRange(found.rowIndex, updatedAtCol).setValue(updatedAt);
  sheet.getRange(found.rowIndex, statusUpdatedAtCol).setValue(updatedAt);
  sheet.getRange(found.rowIndex, statusUpdatedByEmailCol).setValue(updatedByEmail);
  sheet.getRange(found.rowIndex, statusUpdatedByNameCol).setValue(updatedByName);
  sheet.getRange(found.rowIndex, statusUpdatedByRoleCol).setValue(updatedByRole);

  const updatedRow = Object.assign({}, found.row, {
    status: nextStatus,
    status_note: nextStatusNote,
    drive_link: driveLink,
    file_password: filePassword,
    record_updated_at_iso: updatedAt,
    status_updated_at_iso: updatedAt,
    status_updated_by_email: updatedByEmail,
    status_updated_by_name: updatedByName,
    status_updated_by_role: updatedByRole
  });

  return jsonResponse({
    status: "success",
    requestId: requestId,
    rowIndex: found.rowIndex,
    row: toConnectRequestResponse(updatedRow)
  });
}

function handleDeleteConnectRequest(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) {
    return jsonResponse({ status: "error", message: auth.message });
  }

  const authUser = auth.user || {};
  if (!isConnectAdminRole(authUser.role)) {
    return jsonResponse({ status: "error", message: "เมนูนี้สำหรับผู้ดูแลระบบเท่านั้น" });
  }

  const requestId = normalizeConnectString(data && data.requestId);
  if (!requestId) {
    return jsonResponse({ status: "error", message: "ไม่พบ requestId" });
  }

  const sheet = ensureConnectRequestSheet(getConnectRequestSheet());
  const found = findConnectRequestById(sheet, requestId);
  if (!found || !found.rowIndex) {
    return jsonResponse({ status: "error", message: "ไม่พบคำขอที่ต้องการลบ" });
  }

  sheet.deleteRow(found.rowIndex);
  return jsonResponse({
    status: "success",
    requestId: requestId,
    deleted: true
  });
}

function findReportRowIndexByTarget(sheet, target) {
  if (!sheet || sheet.getLastRow() < 2) return -1;
  const no = String(target.targetNo || target.no || "").trim();
  const ts = String(target.targetTimestamp || target.timestamp || "").trim();
  const reporterEmail = normalizeEmail(target.targetReporterEmail || target.reporterEmail || "");
  const unitKey = normalizeUnitKey(target.targetUnit || target.unit || "");

  const idxNo = HEADERS.indexOf("ลำดับ");
  const idxDate = HEADERS.indexOf("วันที่บันทึก");
  const idxTime = HEADERS.indexOf("เวลาบันทึก");
  const idxEmail = HEADERS.indexOf("อีเมล");
  const idxUnit = HEADERS.indexOf("ชุดปฏิบัติการ");
  if (idxNo < 0 || idxDate < 0 || idxEmail < 0 || idxUnit < 0) return -1;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  for (let i = 0; i < values.length; i++) {
    const rowNo = String(values[i][idxNo] == null ? "" : values[i][idxNo]).trim();
    const rowTs = (String(values[i][idxDate] == null ? "" : values[i][idxDate]).trim()
      + " " + String(idxTime >= 0 && values[i][idxTime] != null ? values[i][idxTime] : "").trim()).trim();
    const rowEmail = normalizeEmail(values[i][idxEmail] == null ? "" : values[i][idxEmail]);
    const rowUnitKey = normalizeUnitKey(values[i][idxUnit] == null ? "" : values[i][idxUnit]);

    if (no && rowNo !== no) continue;
    if (ts && rowTs !== ts) continue;
    if (reporterEmail && rowEmail !== reporterEmail) continue;
    if (unitKey && rowUnitKey !== unitKey) continue;
    return i + 2;
  }
  return -1;
}

// ============================================================
//  ROLE PERMISSION SYSTEM
// ============================================================
const VALID_ROLES = [
  "reporter", "requester", "staff", "supervisor",
  "report_admin", "connect_admin", "meeting_admin",
  "admin", "super_admin"
];

const ROLE_PERMISSIONS_MAP = {
  reporter:      { canReport: true },
  requester:     { canConnect: true },
  staff:         { canReport: true, canConnect: true },
  supervisor:    { canReport: true, canConnect: true, isSupervisor: true },
  report_admin:  { canReport: true, isReportAdmin: true },
  connect_admin: { canConnect: true, isConnectAdmin: true },
  meeting_admin: { canReport: true, canConnect: true, isMeetingAdmin: true },
  admin:         { canReport: true, canConnect: true, isReportAdmin: true, isConnectAdmin: true, isMeetingAdmin: true },
  super_admin:   { canReport: true, canConnect: true, isReportAdmin: true, isConnectAdmin: true, isMeetingAdmin: true, isSuperAdmin: true }
};

// หน้าที่แต่ละ role เข้าได้ (derive จาก role อัตโนมัติ ไม่ต้องตั้งในชีท)
const ALL_PAGES = ["actionreport", "connect", "connect-status", "notification", "history", "dashboard", "compare", "rank", "meeting"];
const ROLE_PAGES_MAP = {
  reporter:      ["actionreport", "history", "dashboard", "compare", "rank", "meeting"],
  requester:     ["connect", "connect-status", "dashboard", "compare", "rank", "meeting"],
  staff:         ["actionreport", "connect", "connect-status", "history", "notification", "dashboard", "compare", "rank", "meeting"],
  supervisor:    ["actionreport", "connect", "connect-status", "history", "dashboard", "compare", "rank", "meeting"],
  report_admin:  ["actionreport", "notification", "dashboard", "compare", "rank", "meeting"],
  connect_admin: ["connect", "connect-status", "dashboard", "compare", "rank", "meeting"],
  meeting_admin: ["actionreport", "connect", "connect-status", "history", "notification", "dashboard", "compare", "rank", "meeting"],
  admin:         ALL_PAGES,
  super_admin:   ALL_PAGES
};

function getRolePages(role) {
  var r = String(role || "").trim().toLowerCase();
  return ROLE_PAGES_MAP[r] || [];
}

function getRolePermissions(role) {
  var r = String(role || "").trim().toLowerCase();
  return ROLE_PERMISSIONS_MAP[r] || {};
}

function isValidRole(role) {
  return VALID_ROLES.indexOf(String(role || "").trim().toLowerCase()) !== -1;
}

function canSubmitReport(role) { return !!getRolePermissions(role).canReport; }
function canSubmitConnect(role) { return !!getRolePermissions(role).canConnect; }
function isReportAdminRole(role) { return !!getRolePermissions(role).isReportAdmin; }
function isConnectAdminRole(role) { return !!getRolePermissions(role).isConnectAdmin; }
function isMeetingAdminRole(role) { return !!getRolePermissions(role).isMeetingAdmin; }
function isSuperAdminRole(role) { return !!getRolePermissions(role).isSuperAdmin; }
function isSupervisorRole(role) { return !!getRolePermissions(role).isSupervisor; }

// Backwards-compatible: admin = anyone with at least one admin flag
function isAdminUserRole(role) {
  var p = getRolePermissions(role);
  return !!(p.isReportAdmin || p.isConnectAdmin || p.isMeetingAdmin || p.isSuperAdmin);
}

// Backwards-compatible: standard = any valid role
function isStandardUserRole(role) {
  return isValidRole(role);
}

/**
 * รับข้อมูล POST (login / submitReport)
 */
function doPost(e) {
  try {
    let data = {};
    try {
      const rawBody = e && e.postData ? e.postData.contents : "";
      data = rawBody ? JSON.parse(rawBody) : {};
    } catch (parseErr) {
      return jsonResponse({ status: "error", message: "Invalid JSON" });
    }

    if (data.action === "login") {
      return handleLogin(data);
    } else if (data.action === "getPublicData") {
      return handleGetPublicData(data);
    } else if (data.action === "getData") {
      return handleGetData(data);
    } else if (data.action === "getUnitHistory") {
      return handleGetUnitHistory(data);
    } else if (data.action === "submitChangeRequest") {
      return handleSubmitChangeRequest(data);
    } else if (data.action === "getAdminNotifications") {
      return handleGetAdminNotifications(data);
    } else if (data.action === "approveChangeRequest") {
      return handleApproveChangeRequest(data);
    } else if (data.action === "getCommanderInfo") {
      return handleGetCommanderInfo();
    } else if (data.action === "getConnectRequests") {
      return handleGetConnectRequests(data);
    } else if (data.action === "getConnectBackendInfo") {
      return handleGetConnectBackendInfo(data);
    } else if (data.action === "submitConnectRequest") {
      return handleSubmitConnectRequest(data);
    } else if (data.action === "updateConnectRequestStatus") {
      return handleUpdateConnectRequestStatus(data);
    } else if (data.action === "deleteConnectRequest") {
      return handleDeleteConnectRequest(data);
    } else if (data.action === "downloadConnectResponseFile") {
      return handleDownloadConnectResponseFile(data);
    } else if (data.action === "downloadSummonPdf") {
      return handleDownloadSummonPdf(data);
    } else if (data.action === "getConnectRequestRawState") {
      return handleGetConnectRequestRawState(data);
    } else if (data.action === "reserveConnectDocNumber") {
      return handleReserveConnectDocNumber(data);
    } else if (data.action === "getConnectDocIndex") {
      return handleGetConnectDocIndex(data);
    } else if (data.action === "lookupConnectDocNumber") {
      return handleLookupConnectDocNumber(data);
    } else if (data.action === "uploadSummonPdf") {
      return handleUploadSummonPdf(data);
    } else if (data.action === "getMeetingBookings") {
      return handleGetMeetingBookings(data);
    } else if (data.action === "submitMeetingBooking") {
      return handleSubmitMeetingBooking(data);
    } else if (data.action === "deleteMeetingBooking") {
      return handleDeleteMeetingBooking(data);
    } else if (data.action === "submit" || !data.action) {
      return handleSubmitReport(data);
    } else {
      return jsonResponse({ status: "error", message: "Unknown action" });
    }
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

/**
 * ตรวจสอบอีเมลจากชีทผู้ใช้
 */
function handleLogin(data) {
  const sheet = getUserMasterSheet();

  if (!sheet) {
    return jsonResponse({ status: "error", message: "ไม่พบชีท 'User_master' กรุณารัน initSheet() ก่อน" });
  }

  const email = normalizeEmail(data.email);
  if (!email) {
    return jsonResponse({ status: "error", message: "ไม่พบอีเมล" });
  }

  const idToken = getIdTokenFromPayload(data);
  const tokenCheck = verifyGoogleIdToken(idToken, email);
  if (!tokenCheck.ok) {
    return jsonResponse({ status: "error", message: tokenCheck.message || "ยืนยันตัวตนไม่สำเร็จ" });
  }
  if (tokenCheck.warning) {
    Logger.log("[Auth Warning] " + tokenCheck.warning + " email=" + email);
  }

  const user = findUserByEmail(email, true);
  if (user) {
    const sessionToken = createSessionToken(user.email, user.role);
    const perms = getRolePermissions(user.role);
    return jsonResponse({
      status: "success",
      name: user.name,
      position: user.position,
      unit: user.unit,
      phone: user.phone,
      role: user.role,
      allowedPages: getRolePages(user.role),
      avatarUrl: user.avatarUrl,
      sessionToken: sessionToken,
      sessionExpiresIn: SESSION_TTL_SECONDS,
      canReport: !!perms.canReport,
      canConnect: !!perms.canConnect,
      isReportAdmin: !!perms.isReportAdmin,
      isConnectAdmin: !!perms.isConnectAdmin,
      isMeetingAdmin: !!perms.isMeetingAdmin,
      isSuperAdmin: !!perms.isSuperAdmin,
      isSupervisor: !!perms.isSupervisor
    });
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
  const auth = authorizeRequest(data);
  if (!auth.ok) {
    return jsonResponse({ status: "error", message: auth.message });
  }

  const authUser = auth.user || {};
  if (!canSubmitReport(authUser.role)) {
    return jsonResponse({ status: "error", message: "บัญชีผู้ใช้นี้ไม่มีสิทธิ์บันทึกรายงาน" });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("รายงานผลการปฏิบัติงาน") || ss.getActiveSheet();

  const lastRow = sheet.getLastRow();
  const nextNumber = lastRow;

  const row = [
    nextNumber,
    data.timestampDate || formatThaiDateOnlyStandard(new Date()),
    data.timestampTime || formatTimeOnlyStandard(new Date()),
    authUser.name || data.reporter || "",
    authUser.email || "",
    data.unit || "",
    data.leaderName || "",
    data.leaderPhone || "",
    normalizeSheetThaiDateValue(data.date || "") || "",
    normalizeSheetTimeValue(data.time || "") || "",
    data.location || "",
    data.coordinates || "",
    data.actionType || "",
    normalizeCourtValue(data.arrestCourt || ""),
    data.arrestNo || "",
    normalizeSheetThaiDateValue(data.warrantDate || "") || "",
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
    normalizeCourtValue(data.searchCourt || ""),
    data.searchNo || "",
    normalizeSheetThaiDateValue(data.searchDate || "") || "",
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
  const auth = authorizeRequest(data);
  if (!auth.ok) {
    return jsonResponse({ status: "error", message: auth.message });
  }

  return jsonResponse({ status: "success", data: readReportRows() });
}

function handleGetUnitHistory(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) {
    return jsonResponse({ status: "error", message: auth.message });
  }

  const authUser = auth.user || {};
  if (!canSubmitReport(authUser.role) && !isReportAdminRole(authUser.role) && !isSupervisorRole(authUser.role)) {
    return jsonResponse({ status: "error", message: "เมนูนี้สำหรับผู้ใช้งานสิทธิ์ User เท่านั้น" });
  }

  const unit = String(authUser.unit || "").trim();
  const unitKey = normalizeUnitKey(unit);
  const authEmail = normalizeEmail(authUser.email);
  const rows = readReportRows();
  const myRows = rows.filter(function(row) {
    return normalizeEmail(row["อีเมล"] || "") === authEmail;
  });

  const changeRows = readChangeRequestRows();
  const latestRequestByTarget = {};
  changeRows.forEach(function(req) {
    const key = makeChangeRequestTargetKey(
      req.targetNo,
      req.targetTimestamp,
      req.targetReporterEmail,
      req.targetUnit
    );
    if (!key.replace(/\|/g, "")) return;
    const prev = latestRequestByTarget[key];
    const reqAt = String(req.createdAt || "");
    const prevAt = prev ? String(prev.createdAt || "") : "";
    if (!prev || reqAt > prevAt) {
      latestRequestByTarget[key] = req;
    }
  });

  // Admin / Super Admin → see ALL rows from every unit
  const isAdmin = isReportAdminRole(authUser.role);

  let unitRows = [];
  if (!isAdmin && unitKey) {
    unitRows = rows.filter(function(row) {
      return normalizeUnitKey(row["ชุดปฏิบัติการ"] || "") === unitKey;
    });
  }

  // Fallback: ถ้าไม่พบข้อมูลจาก unit ให้แสดงข้อมูลที่ผู้ใช้คนนี้เคยกรอกแทน
  const effectiveRows = isAdmin ? rows.slice() : (unitRows.length ? unitRows : myRows).slice();
  effectiveRows.reverse(); // newest first
  const enrichedRows = effectiveRows.map(function(row) {
    const out = Object.assign({}, row);
    const key = makeChangeRequestTargetKey(
      row["ลำดับ"] || "",
      (String(row["วันที่บันทึก"] || "") + " " + (String(row["เวลาบันทึก"] || "").trim()).trim()).trim(),
      row["อีเมล"] || "",
      row["ชุดปฏิบัติการ"] || ""
    );
    const req = latestRequestByTarget[key];
    if (req) {
      out.changeRequestId = req.requestId || "";
      out.changeRequestStatus = req.status || "new";
      out.changeRequestType = req.requestType || "";
      out.changeRequestedAt = req.createdAt || "";
      out.changeRequestedByEmail = req.requestedByEmail || "";
    }
    return out;
  });

  return jsonResponse({
    status: "success",
    unit: unit || "",
    total: enrichedRows.length,
    mineTotal: myRows.length,
    unitTotal: unitRows.length,
    fallbackToMine: !unitRows.length && !!myRows.length,
    data: enrichedRows
  });
}

function handleSubmitChangeRequest(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) {
    return jsonResponse({ status: "error", message: auth.message });
  }

  const authUser = auth.user || {};
  if (!canSubmitReport(authUser.role)) {
    return jsonResponse({ status: "error", message: "บัญชีผู้ใช้นี้ไม่มีสิทธิ์ส่งคำขอแก้ไข/ลบ" });
  }

  const requestType = String(data.requestType || "").trim().toLowerCase();
  if (requestType !== "edit" && requestType !== "delete") {
    return jsonResponse({ status: "error", message: "ประเภทคำขอไม่ถูกต้อง" });
  }

  const target = data && typeof data.target === "object" ? data.target : {};
  const targetNo = String(target.no || "").trim();
  const targetTimestamp = String(target.timestamp || "").trim();
  if (!targetNo && !targetTimestamp) {
    return jsonResponse({ status: "error", message: "ไม่พบรายการอ้างอิงที่ต้องการแจ้งคำขอ" });
  }

  const reportRows = readReportRows();
  const matchedRow = reportRows.find(function(row) {
    const rowNo = String(row["ลำดับ"] || "").trim();
    const rowTs = (String(row["วันที่บันทึก"] || "").trim() + " " + (String(row["เวลาบันทึก"] || "").trim()).trim()).trim();
    if (targetNo && rowNo !== targetNo) return false;
    if (targetTimestamp && rowTs !== targetTimestamp) return false;
    return !!(rowNo || rowTs);
  });

  if (!matchedRow) {
    return jsonResponse({ status: "error", message: "ไม่พบรายการอ้างอิงในระบบ" });
  }

  const targetKey = makeChangeRequestTargetKey(
    String(matchedRow["ลำดับ"] || targetNo || "").trim(),
    String((matchedRow["วันที่บันทึก"] || "") + " " + (matchedRow["เวลาบันทึก"] || "").trim()).trim(),
    String(matchedRow["อีเมล"] || "").trim(),
    String(matchedRow["ชุดปฏิบัติการ"] || "").trim()
  );
  const duplicateExists = readChangeRequestRows().some(function(req) {
    return makeChangeRequestTargetKey(
      req.targetNo,
      req.targetTimestamp,
      req.targetReporterEmail,
      req.targetUnit
    ) === targetKey;
  });
  if (duplicateExists) {
    return jsonResponse({ status: "error", message: "รายการนี้ถูกแจ้งแอดมินแล้ว ไม่สามารถแจ้งซ้ำได้" });
  }

  const deleteReason = String(data.deleteReason || "").trim();
  const allowedDeleteReasons = ["ข้อมูลซ้ำ", "ข้อมูลผิดพลาด ต้องการแก้ไขใหม่เอง"];
  if (requestType === "delete" && allowedDeleteReasons.indexOf(deleteReason) === -1) {
    return jsonResponse({ status: "error", message: "กรุณาเลือกเหตุผลการลบรายการให้ถูกต้อง" });
  }

  const note = String(data.note || "").trim();
  const changedFields = normalizeChangedFieldsPayload(data.changedFields);
  const changedColumns = Object.keys(changedFields);
  const correctedDataRaw = String(data.correctedData || "").trim();
  if (requestType === "edit" && !changedColumns.length && !correctedDataRaw) {
    return jsonResponse({ status: "error", message: "ไม่พบคอลัมน์ที่แก้ไข" });
  }
  const correctedData = requestType === "edit"
    ? (changedColumns.length
      ? JSON.stringify({ changedColumns: changedColumns, changedFields: changedFields })
      : correctedDataRaw)
    : "";

  const sheet = getChangeRequestSheet(true);
  if (!sheet) {
    return jsonResponse({ status: "error", message: "ไม่พบชีทคำขอแก้ไขรายการ" });
  }

  const requestId = Utilities.getUuid();
  const createdAt = new Date().toISOString();
  const row = [
    requestId,
    createdAt,
    requestType,
    "new",
    normalizeEmail(authUser.email),
    String(authUser.name || "").trim(),
    String(authUser.unit || "").trim(),
    String(matchedRow["ลำดับ"] || targetNo || "").trim(),
    String((matchedRow["วันที่บันทึก"] || "") + " " + (matchedRow["เวลาบันทึก"] || "").trim()).trim(),
    String(matchedRow["ผู้รายงาน"] || "").trim(),
    String(matchedRow["อีเมล"] || "").trim(),
    String(matchedRow["ชุดปฏิบัติการ"] || "").trim(),
    String(matchedRow["วันที่"] || "").trim(),
    String(matchedRow["เวลา"] || "").trim(),
    String(matchedRow["สถานที่ (จับกุม/หรือตรวจสอบ)"] || "").trim(),
    String(matchedRow["การจับกุม ตรวจสอบ"] || "").trim(),
    String(matchedRow["ชื่อ-สกุล ผู้ต้องหา"] || "").trim(),
    String(matchedRow["เลขที่หมายจับ"] || matchedRow["เลขที่หมายค้น"] || "").trim(),
    String(
      matchedRow["เลขประจำตัวประชาชน (13 หลัก) #1"] ||
      matchedRow["เลขประจำตัวประชาชน (13 หลัก)"] ||
      ""
    ).trim(),
    requestType === "delete" ? deleteReason : "",
    note,
    correctedData,
    JSON.stringify(matchedRow || {})
  ];

  sheet.appendRow(row);
  return jsonResponse({
    status: "success",
    requestId: requestId,
    createdAt: createdAt
  });
}

function handleGetAdminNotifications(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) {
    return jsonResponse({ status: "error", message: auth.message });
  }

  const authUser = auth.user || {};
  if (!isReportAdminRole(authUser.role)) {
    return jsonResponse({ status: "error", message: "เมนูนี้สำหรับผู้ดูแลระบบเท่านั้น" });
  }

  const rows = readChangeRequestRows();
  rows.sort(function(a, b) {
    const at = String(a.createdAt || "");
    const bt = String(b.createdAt || "");
    if (at === bt) return 0;
    return at > bt ? -1 : 1;
  });

  const notifications = rows.map(function(row) {
    const payload = parseCorrectedDataPayload(row.correctedData || "");
    const changedColumns = payload && Array.isArray(payload.changedColumns)
      ? payload.changedColumns.map(function(col) { return String(col || "").trim(); }).filter(function(col) { return !!col; })
      : [];
    const changedFields = normalizeChangedFieldsPayload(payload && payload.changedFields ? payload.changedFields : {});
    const correctedDataText = payload ? "" : (row.correctedData || "");
    return {
      requestId: row.requestId || "",
      createdAt: row.createdAt || "",
      requestType: row.requestType || "",
      status: row.status || "new",
      requestedByEmail: row.requestedByEmail || "",
      requestedByName: row.requestedByName || "",
      requestedByUnit: row.requestedByUnit || "",
      targetNo: row.targetNo || "",
      targetTimestamp: row.targetTimestamp || "",
      targetReporter: row.targetReporter || "",
      targetReporterEmail: row.targetReporterEmail || "",
      targetUnit: row.targetUnit || "",
      targetDate: row.targetDate || "",
      targetTime: row.targetTime || "",
      targetLocation: row.targetLocation || "",
      targetActionType: row.targetActionType || "",
      targetSuspect: row.targetSuspect || "",
      targetWarrantNo: row.targetWarrantNo || "",
      targetCitizenId: row.targetCitizenId || "",
      deleteReason: row.deleteReason || "",
      note: row.note || "",
      correctedData: correctedDataText,
      changedColumns: changedColumns,
      changedFields: changedFields,
      sourceRowJson: row.sourceRowJson || ""
    };
  });

  return jsonResponse({
    status: "success",
    total: notifications.length,
    data: notifications
  });
}

function handleApproveChangeRequest(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) {
    return jsonResponse({ status: "error", message: auth.message });
  }

  const authUser = auth.user || {};
  if (!isReportAdminRole(authUser.role)) {
    return jsonResponse({ status: "error", message: "เมนูนี้สำหรับผู้ดูแลระบบเท่านั้น" });
  }

  const requestId = String(data.requestId || "").trim();
  if (!requestId) {
    return jsonResponse({ status: "error", message: "ไม่พบรหัสคำขอ" });
  }

  const requestSheet = getChangeRequestSheet(false);
  if (!requestSheet) {
    return jsonResponse({ status: "error", message: "ไม่พบชีทคำขอแก้ไขรายการ" });
  }

  const found = findChangeRequestById(requestSheet, requestId);
  if (!found) {
    return jsonResponse({ status: "error", message: "ไม่พบคำขอที่ต้องการอนุมัติ" });
  }

  const req = found.row || {};
  const requestType = String(req.requestType || "").trim().toLowerCase();
  const currentStatus = String(req.status || "").trim().toLowerCase();
  const decisionRaw = String(data.decision || "").trim().toLowerCase();
  const decision = (decisionRaw === "reject" || decisionRaw === "rejected" || decisionRaw === "deny" || decisionRaw === "decline")
    ? "reject"
    : "approve";

  const isDoneStatus = (
    currentStatus === "done" ||
    currentStatus === "completed" ||
    currentStatus === "resolved" ||
    currentStatus === "closed"
  );
  const isRejectedStatus = (
    currentStatus === "rejected" ||
    currentStatus === "cancelled" ||
    currentStatus === "canceled"
  );
  if (isDoneStatus || isRejectedStatus) {
    return jsonResponse({
      status: "success",
      requestId: requestId,
      requestStatus: isRejectedStatus ? "rejected" : "done",
      updatedColumns: []
    });
  }

  if (requestType !== "edit" && requestType !== "delete") {
    return jsonResponse({ status: "error", message: "ประเภทคำขอไม่ถูกต้อง" });
  }

  const statusCol = CONNECT_REQUEST_HEADERS.indexOf("status") + 1;
  if (decision === "reject") {
    if (statusCol > 0) {
      requestSheet.getRange(found.rowIndex, statusCol).setValue("rejected");
    }
    return jsonResponse({
      status: "success",
      requestId: requestId,
      requestStatus: "rejected",
      updatedColumns: []
    });
  }

  const reportSheet = getReportSheet();
  if (!reportSheet || reportSheet.getLastRow() < 2) {
    return jsonResponse({ status: "error", message: "ไม่พบชีทรายงานหลัก" });
  }

  const reportRowIndex = findReportRowIndexByTarget(reportSheet, {
    targetNo: req.targetNo,
    targetTimestamp: req.targetTimestamp,
    targetReporterEmail: req.targetReporterEmail,
    targetUnit: req.targetUnit
  });
  if (reportRowIndex < 2) {
    return jsonResponse({ status: "error", message: "ไม่พบแถวข้อมูลเดิมที่ต้องการดำเนินการ" });
  }

  const colMap = buildReportHeaderColumnMap();
  let updatedColumns = [];
  if (requestType === "edit") {
    const payload = parseCorrectedDataPayload(req.correctedData || "");
    const changedFields = normalizeChangedFieldsPayload(payload && payload.changedFields ? payload.changedFields : data.changedFields);
    const changedKeys = Object.keys(changedFields);
    if (!changedKeys.length) {
      return jsonResponse({ status: "error", message: "ไม่พบข้อมูลคอลัมน์ที่ต้องแก้ไข" });
    }

    updatedColumns = changedKeys
      .map(function(key) {
        const colIndex = colMap[key];
        if (!colIndex) return null;
        return { key: key, colIndex: colIndex, value: changedFields[key] };
      })
      .filter(function(item) { return !!item; });

    if (!updatedColumns.length) {
      return jsonResponse({ status: "error", message: "ไม่พบคอลัมน์ที่ตรงกับชีทรายงาน" });
    }

    updatedColumns.forEach(function(item) {
      const cell = reportSheet.getRange(reportRowIndex, item.colIndex);
      cell.setValue(item.value);
      cell.setBackground(EDIT_CHANGED_CELL_COLOR);
    });
  } else if (requestType === "delete") {
    const recycleSheet = getRecycleBinSheet(true);
    if (!recycleSheet) {
      return jsonResponse({ status: "error", message: "ไม่พบชีท Recycle Bin" });
    }

    const sourceRowValues = reportSheet.getRange(reportRowIndex, 1, 1, HEADERS.length).getValues()[0];
    const hasAnyValue = sourceRowValues.some(function(cell) {
      return String(cell == null ? "" : cell).trim() !== "";
    });
    if (!hasAnyValue) {
      return jsonResponse({ status: "error", message: "ไม่พบข้อมูลแถวที่ต้องการย้ายไป Recycle Bin" });
    }

    // เก็บสำเนาข้อมูลทั้งแถวลง Recycle Bin (คงค่า "ลำดับ" เดิม)
    recycleSheet.appendRow(sourceRowValues);

    // ล้างข้อมูลเฉพาะคอลัมน์ B เป็นต้นไป (คงเลขลำดับคอลัมน์ A เดิมไว้)
    if (HEADERS.length > 1) {
      reportSheet.getRange(reportRowIndex, 2, 1, HEADERS.length - 1).clearContent();
    }

    // เปลี่ยนสีพื้นหลังทั้งแถวเป็นสีเทาเข้มตามที่กำหนด
    reportSheet.getRange(reportRowIndex, 1, 1, HEADERS.length).setBackground(RECYCLED_ROW_COLOR);
    // ทำให้เลขลำดับในคอลัมน์ A อ่านชัดบนพื้นเข้ม
    reportSheet.getRange(reportRowIndex, 1).setFontColor("#f8fafc");

    updatedColumns = [{ key: "ALL_COLUMNS_EXCEPT_NO", colIndex: 0, value: "" }];
  }

  if (statusCol > 0) {
    requestSheet.getRange(found.rowIndex, statusCol).setValue("done");
  }

  return jsonResponse({
    status: "success",
    requestId: requestId,
    requestStatus: "done",
    updatedColumns: updatedColumns.map(function(item) { return item.key; })
  });
}

function handleGetCommanderInfo() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(COMMANDER_INFO_SHEET_NAME);
  if (!sheet) return jsonResponse({ status: "error", message: "ไม่พบชีท CommanderInfo" });

  const rows = sheet.getDataRange().getValues();
  if (!rows || rows.length < 2) return jsonResponse({ status: "error", message: "ไม่มีข้อมูล" });

  // หา index ของแต่ละคอลัมน์จาก header (รองรับ column เพิ่มเติมโดยอัตโนมัติ)
  const header = rows[0].map(function(h) { return String(h).trim().toLowerCase(); });
  const iType  = header.indexOf("ประเภท");
  const iName  = header.indexOf("ชื่อ-สกุล");
  const iPhone = header.indexOf("เบอร์โทร");
  const iUnit  = header.indexOf("ชุดปฏิบัติการ");
  const iPos   = header.indexOf("ตำแหน่ง");

  var commander = { name: "", phone: "", position: "" };
  var supervisors = [];

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var type     = String(iType  >= 0 ? row[iType]  : "").trim().replace(/\.+$/, '').toLowerCase();
    var name     = String(iName  >= 0 ? row[iName]  : "").trim();
    var phone    = String(iPhone >= 0 ? row[iPhone] : "").trim();
    var unit     = String(iUnit  >= 0 ? row[iUnit]  : "").trim();
    var position = String(iPos   >= 0 ? row[iPos]   : "").trim();
    if (!name) continue;
    if (type === "ผกก") {
      commander = { name: name, phone: phone, position: position };
    } else if (type === "สว") {
      supervisors.push({ name: name, phone: phone, unit: unit, position: position });
    }
  }

  return jsonResponse({ status: "success", commander: commander, supervisors: supervisors });
}

function handleGetPublicData(data) {
  const rows = readReportRows();
  const publicRows = rows.map(function(row) {
    return toPublicRow(row);
  });

  return jsonResponse({ status: "success", data: publicRows });
}

/* ====== MEETING BOOKINGS ====== */

function getMeetingBookingsSheet(createIfMissing) {
  const ss = getConnectSpreadsheet() || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MEETING_BOOKINGS_SHEET_NAME);
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet(MEETING_BOOKINGS_SHEET_NAME);
    sheet.getRange(1, 1, 1, MEETING_BOOKINGS_HEADERS.length).setValues([MEETING_BOOKINGS_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet || null;
}

function handleGetMeetingBookings(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) return jsonResponse({ status: "error", message: auth.message });

  const sheet = getMeetingBookingsSheet(false);
  if (!sheet || sheet.getLastRow() < 2) {
    return jsonResponse({ status: "ok", bookings: [] });
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, MEETING_BOOKINGS_HEADERS.length).getDisplayValues();
  const bookings = rows.map(function (row) {
    var obj = {};
    MEETING_BOOKINGS_HEADERS.forEach(function (h, i) {
      obj[h] = row[i] != null ? String(row[i]) : "";
    });
    if (obj.equipment) {
      try { obj.equipment = JSON.parse(obj.equipment); } catch (_) { obj.equipment = []; }
    } else {
      obj.equipment = [];
    }
    return obj;
  }).filter(function (b) { return b.id; });

  return jsonResponse({ status: "ok", bookings: bookings });
}

function handleSubmitMeetingBooking(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) return jsonResponse({ status: "error", message: auth.message });

  const booking = data && data.booking;
  if (!booking || !booking.id || !booking.roomId || !booking.date || !booking.start || !booking.end || !booking.title || !booking.organizer) {
    return jsonResponse({ status: "error", message: "ข้อมูลการจองไม่ครบ" });
  }

  const sheet = getMeetingBookingsSheet(true);
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (_) {
    return jsonResponse({ status: "error", message: "ระบบกำลังประมวลผล กรุณาลองใหม่" });
  }

  try {
    // ตรวจเวลาซ้ำ
    if (sheet.getLastRow() >= 2) {
      var existing = sheet.getRange(2, 1, sheet.getLastRow() - 1, MEETING_BOOKINGS_HEADERS.length).getValues();
      var roomIdx = MEETING_BOOKINGS_HEADERS.indexOf("roomId");
      var dateIdx = MEETING_BOOKINGS_HEADERS.indexOf("date");
      var startIdx = MEETING_BOOKINGS_HEADERS.indexOf("start");
      var endIdx = MEETING_BOOKINGS_HEADERS.indexOf("end");
      var bStart = normalizeTimeString(booking.start);
      var bEnd = normalizeTimeString(booking.end);
      for (var i = 0; i < existing.length; i++) {
        if (String(existing[i][roomIdx]) === String(booking.roomId) &&
            String(existing[i][dateIdx]) === String(booking.date)) {
          var eStart = normalizeTimeString(existing[i][startIdx]);
          var eEnd = normalizeTimeString(existing[i][endIdx]);
          if (bStart < eEnd && bEnd > eStart) {
            return jsonResponse({ status: "error", message: "เวลาซ้ำกับรายการเดิม (" + eStart + " - " + eEnd + ")" });
          }
        }
      }
    }

    var row = MEETING_BOOKINGS_HEADERS.map(function (h) {
      if (h === "equipment") return JSON.stringify(booking.equipment || []);
      return booking[h] != null ? String(booking[h]) : "";
    });
    sheet.appendRow(row);
    return jsonResponse({ status: "ok", message: "บันทึกสำเร็จ" });
  } finally {
    lock.releaseLock();
  }
}

function handleDeleteMeetingBooking(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) return jsonResponse({ status: "error", message: auth.message });

  var bookingId = data && data.bookingId ? String(data.bookingId) : "";
  if (!bookingId) return jsonResponse({ status: "error", message: "ไม่พบ bookingId" });

  var userEmail = normalizeEmail(auth.user ? auth.user.email : "");
  var isAdmin = auth.user && (auth.user.isMeetingAdmin || auth.user.isSuperAdmin ||
    auth.user.role === "admin" || auth.user.role === "super_admin" || auth.user.role === "meeting_admin");

  var sheet = getMeetingBookingsSheet(false);
  if (!sheet || sheet.getLastRow() < 2) {
    return jsonResponse({ status: "error", message: "ไม่พบรายการ" });
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (_) {
    return jsonResponse({ status: "error", message: "ระบบกำลังประมวลผล กรุณาลองใหม่" });
  }

  try {
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, MEETING_BOOKINGS_HEADERS.length).getValues();
    var idIdx = MEETING_BOOKINGS_HEADERS.indexOf("id");
    var createdByIdx = MEETING_BOOKINGS_HEADERS.indexOf("createdBy");
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][idIdx]) === bookingId) {
        var creator = normalizeEmail(String(rows[i][createdByIdx]));
        if (!isAdmin && creator !== userEmail) {
          return jsonResponse({ status: "error", message: "คุณไม่มีสิทธิ์ลบรายการนี้ (เฉพาะผู้จองเท่านั้น)" });
        }
        sheet.deleteRow(i + 2);
        return jsonResponse({ status: "ok", message: "ลบรายการแล้ว" });
      }
    }
    return jsonResponse({ status: "error", message: "ไม่พบรายการที่ต้องการลบ" });
  } finally {
    lock.releaseLock();
  }
}

function normalizeTimeString(t) {
  if (!t) return '';
  if (typeof t === 'string' && /^\d{2}:\d{2}$/.test(t)) return t;
  var d = new Date(t);
  if (!isNaN(d.getTime())) {
    var h = d.getHours().toString().padStart(2, '0');
    var m = d.getMinutes().toString().padStart(2, '0');
    return h + ':' + m;
  }
  return String(t).slice(-5); // fallback
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return jsonResponse({ status: "ok", message: "CSD1 Report API v2.9 (Public Dashboard + Session Auth + Unit History + Change Requests + Admin Notifications)", version: "2.9" });
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CSD1 Report')
    .addItem('สร้างหัวตาราง + ชีทผู้ใช้', 'initSheet')
    .addItem('สร้างชีท CSD1 Connect Requests', 'initConnectRequestSheet')
    .addItem('รีเซ็ตเลขหนังสือ CSD1 Connect = 1', 'resetConnectDocNumberToOne')
    .addItem('ย้ายข้อมูล: แยกคอลัมน์วันที่/เวลา', 'migrateTimestampColumns')
    .addItem('ปรับรูปแบบคอลัมน์เวลา', 'normalizeReportTimeColumn')
    .addItem('ปรับรูปแบบคอลัมน์วันที่', 'normalizeReportDateColumns')
    .addToUi();
}

function applyReportSheetTimeColumnValidation(sheet) {
  if (!sheet) return;
  const timeCol = HEADERS.indexOf("เวลา") + 1;
  if (timeCol <= 0) return;
  const startRow = 2;
  const lastDataRow = sheet.getLastRow();
  const rowCount = Math.min(Math.max(lastDataRow, 1), 200);
  const range = sheet.getRange(startRow, timeCol, rowCount, 1);
  range.setNumberFormat("@");
  const rule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=OR(INDIRECT(ADDRESS(ROW(),COLUMN()))="",REGEXMATCH(TO_TEXT(INDIRECT(ADDRESS(ROW(),COLUMN()))),"^(?:[01]?\\d|2[0-3]):[0-5]\\d$"))')
    .setAllowInvalid(false)
    .setHelpText("กรอกเวลาเป็นรูปแบบ HH:mm เช่น 14:30")
    .build();
  range.setDataValidation(rule);
}

function normalizeReportTimeColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("รายงานผลการปฏิบัติงาน");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("ไม่พบชีท 'รายงานผลการปฏิบัติงาน'");
    return;
  }

  const timeCol = HEADERS.indexOf("เวลา") + 1;
  const lastRow = sheet.getLastRow();
  applyReportSheetTimeColumnValidation(sheet);
  if (timeCol <= 0 || lastRow < 2) {
    SpreadsheetApp.getUi().alert("ไม่พบข้อมูลคอลัมน์เวลา");
    return;
  }

  const range = sheet.getRange(2, timeCol, lastRow - 1, 1);
  const displayValues = range.getDisplayValues();
  const normalizedValues = [];
  let changedCount = 0;

  for (var i = 0; i < displayValues.length; i++) {
    var current = String(displayValues[i][0] || "").trim();
    var normalized = normalizeSheetTimeValue(current) || current;
    if (normalized !== current) changedCount++;
    normalizedValues.push([normalized]);
  }

  range.setNumberFormat("@");
  range.setValues(normalizedValues);
  SpreadsheetApp.getUi().alert(
    changedCount > 0
      ? "ปรับรูปแบบเวลาเรียบร้อย " + changedCount + " แถว"
      : "ไม่พบข้อมูลเวลาที่ต้องปรับรูปแบบ"
  );
}

function normalizeReportDateColumns() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("รายงานผลการปฏิบัติงาน");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("ไม่พบชีท 'รายงานผลการปฏิบัติงาน'");
    return;
  }

  var targetHeaders = ["วันที่", "วันเดือนปีที่ออกหมาย", "วันเดือนปีที่ออกหมายค้น"];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("ไม่พบข้อมูลคอลัมน์วันที่");
    return;
  }

  var changedCount = 0;
  targetHeaders.forEach(function(headerName) {
    var col = HEADERS.indexOf(headerName) + 1;
    if (col <= 0) return;
    var range = sheet.getRange(2, col, lastRow - 1, 1);
    var displayValues = range.getDisplayValues();
    var normalizedValues = [];
    var hasChanges = false;
    for (var i = 0; i < displayValues.length; i++) {
      var current = String(displayValues[i][0] || "").trim();
      var normalized = normalizeSheetThaiDateValue(current) || current;
      if (normalized !== current) {
        changedCount++;
        hasChanges = true;
      }
      normalizedValues.push([normalized]);
    }
    if (hasChanges) {
      range.setNumberFormat("@");
      range.setValues(normalizedValues);
    }
  });

  SpreadsheetApp.getUi().alert(
    changedCount > 0
      ? "ปรับรูปแบบวันที่เรียบร้อย " + changedCount + " ช่อง"
      : "ไม่พบข้อมูลวันที่ที่ต้องปรับรูปแบบ"
  );
}

function onEdit(e) {
  const range = e && e.range;
  if (!range) return;
  const sheet = range.getSheet();
  if (!sheet || sheet.getName() !== "รายงานผลการปฏิบัติงาน") return;

  const headerName = String(HEADERS[range.getColumn() - 1] || "").trim();
  const isTargetColumn = (
    headerName === "เวลา" ||
    headerName === "วันที่" ||
    headerName === "วันเดือนปีที่ออกหมาย" ||
    headerName === "วันเดือนปีที่ออกหมายค้น"
  );
  if (range.getRow() < 2 || !isTargetColumn || range.getNumColumns() !== 1) return;

  const displayValues = range.getDisplayValues();
  const normalizedValues = [];
  let hasChanges = false;

  for (var i = 0; i < displayValues.length; i++) {
    var current = String(displayValues[i][0] || "").trim();
    var normalized = headerName === "เวลา"
      ? (normalizeSheetTimeValue(current) || current)
      : (normalizeSheetThaiDateValue(current) || current);
    if (normalized !== current) hasChanges = true;
    normalizedValues.push([normalized]);
  }

  range.setNumberFormat("@");
  if (hasChanges) {
    range.setValues(normalizedValues);
  }
}

/**
 * รันครั้งเดียวเพื่อย้ายข้อมูลเก่า:
 *   - ถ้าชีทยังใช้ schema เก่า (คอลัมน์ B = "ประทับเวลา"):
 *       แทรกคอลัมน์ C ใหม่ + แยก timestamp → "วันที่บันทึก" / "เวลาบันทึก"
 *   - ถ้า header ถูกอัพเดทแล้วแต่ข้อมูลยังไม่ถูกแยก:
 *       แยก timestamp ใน column B ที่ยังเป็น format เก่าเท่านั้น
 */
function migrateTimestampColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("รายงานผลการปฏิบัติงาน");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("ไม่พบชีท 'รายงานผลการปฏิบัติงาน'");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("ไม่มีข้อมูลใน Sheet");
    return;
  }

  const headerB = String(sheet.getRange(1, 2).getValue() || "").trim();
  const headerC = String(sheet.getRange(1, 3).getValue() || "").trim();

  // ─── กรณี 1: schema เก่า (B = "ประทับเวลา") → แทรก column C ใหม่ ───
  if (headerB === "ประทับเวลา") {
    sheet.insertColumnAfter(2);
    // อัพเดท header row
    const hRange = sheet.getRange(1, 1, 1, HEADERS.length);
    hRange.setValues([HEADERS]);
    hRange.setFontWeight("bold");
    hRange.setBackground("#1a56db");
    hRange.setFontColor("#ffffff");
    hRange.setHorizontalAlignment("center");
    hRange.setWrap(true);
    sheet.setFrozenRows(1);

    // แยก timestamp: หลัง insertColumnAfter(2), B ยังเป็น old value, C ว่าง
    let count = 0;
    for (let r = 2; r <= lastRow; r++) {
      const old = String(sheet.getRange(r, 2).getValue() || "").trim();
      if (!old) continue;
      const split = _splitTimestamp(old);
      sheet.getRange(r, 2).setValue(split.date);
      sheet.getRange(r, 3).setValue(split.time);
      count++;
    }
    SpreadsheetApp.getUi().alert(
      "เสร็จแล้ว! แทรกคอลัมน์ใหม่และแยก timestamp " + count + " แถว\n" +
      "กรุณา Deploy ใหม่เพื่อให้ GAS ใช้ HEADERS ชุดใหม่"
    );
    return;
  }

  // ─── กรณี 2: header ใหม่แล้ว แต่ข้อมูล B ยังเป็น format เก่า ───
  if (headerB === "วันที่บันทึก" && headerC === "เวลาบันทึก") {
    let count = 0;
    for (let r = 2; r <= lastRow; r++) {
      const valB = String(sheet.getRange(r, 2).getValue() || "").trim();
      const valC = String(sheet.getRange(r, 3).getValue() || "").trim();
      // ข้ามแถวที่มี format ใหม่แล้ว (C มีค่า หรือ B ไม่มี "เวลา")
      if (valC || !valB || valB.indexOf("เวลา") === -1) continue;
      const split = _splitTimestamp(valB);
      sheet.getRange(r, 2).setValue(split.date);
      sheet.getRange(r, 3).setValue(split.time);
      count++;
    }
    SpreadsheetApp.getUi().alert(
      count > 0
        ? "เสร็จแล้ว! แยก timestamp " + count + " แถว"
        : "ไม่พบแถวที่ต้องแปลง (ข้อมูลอาจถูกย้ายไปแล้ว)"
    );
    return;
  }

  SpreadsheetApp.getUi().alert(
    "ไม่รู้จัก schema ปัจจุบัน\nB = '" + headerB + "'\nC = '" + headerC + "'\n" +
    "ลองรัน initSheet() ก่อน แล้วค่อยรัน migrateTimestampColumns() อีกครั้ง"
  );
}

/**
 * แยก timestamp หลาย format → { date: "2 มี.ค. 69", time: "17.14.22" }
 *
 * รองรับ:
 *   "2/3/2026 17:14:22"   (CE year)
 *   "2/3/2569 17:14:22"   (BE year)
 *   "9 มี.ค. 69 เวลา 19.42 น."
 *   "9 มี.ค. 69 เวลา 19.42.00 น."
 */
function _splitTimestamp(raw) {
  const MONTHS = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
    "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];

  const s = String(raw || "").trim();

  // ─── Format: d/M/yyyy HH:mm:ss หรือ d/M/yyyy HH:mm ───
  const slashMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{2}):(\d{2})(?::(\d{2}))?$/);
  if (slashMatch) {
    const day   = Number(slashMatch[1]);
    const month = Number(slashMatch[2]);
    const year  = Number(slashMatch[3]);
    const hh    = slashMatch[4];
    const mm    = slashMatch[5];
    const ss    = slashMatch[6] || "00";

    if (month >= 1 && month <= 12) {
      // ถ้า year > 2500 = ปี พ.ศ. แล้ว, ถ้าน้อยกว่า = ค.ศ. → แปลงเป็น พ.ศ.
      const yearBE = year > 2500 ? year : year + 543;
      const yearShort = String(yearBE % 100).padStart(2, "0");
      return {
        date: day + " " + MONTHS[month - 1] + " " + yearShort,
        time: hh + "." + mm + "." + ss
      };
    }
  }

  // ─── Format หลัก: "9 มี.ค. 69 เวลา 19.42 น." หรือ "...19.42.00 น." ───
  const thaiMatch = s.match(/^(.+?)\s+เวลา\s+(\d{1,2}\.\d{2}(?:\.\d{2})?)\s*น\.?$/);
  if (thaiMatch) {
    const t = thaiMatch[2];
    const timeFull = t.split(".").length === 2 ? t + ".00" : t;
    return { date: thaiMatch[1].trim(), time: timeFull };
  }

  // ─── fallback: ตัดครึ่งตรงช่องว่างสุดท้าย ───
  const spaceIdx = s.lastIndexOf(" ");
  if (spaceIdx > 0) {
    return { date: s.slice(0, spaceIdx).trim(), time: s.slice(spaceIdx + 1).trim() };
  }

  return { date: s, time: "" };
}
