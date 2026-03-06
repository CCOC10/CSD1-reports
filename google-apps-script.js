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
const ALLOWED_GOOGLE_CLIENT_IDS = [
  "108806756839-iv4nrrfk4355ogcl2p2ehkh5f6a1u90b.apps.googleusercontent.com"
];

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

function normalizeEmail(email) {
  return String(email || "").trim().toLowerCase();
}

function getUserMasterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName("User_master");
}

function normalizeHeaderKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[()_\-./]/g, "");
}

function resolveUserMasterColumns(headerRow) {
  const defaults = { email: 0, name: 1, position: 2, unit: 3, phone: 4, role: 5 };
  if (!headerRow || !headerRow.length) return defaults;

  const aliases = {
    email: ["อีเมล", "email", "e-mail", "mail", "gmail"],
    name: ["ชื่อสกุล", "ชื่อ", "name", "fullname", "full name", "full_name"],
    position: ["ตำแหน่ง", "position", "rank"],
    unit: ["ชุดปฏิบัติการ", "ชป", "หน่วย", "unit", "team"],
    phone: ["เบอร์โทร", "เบอร์โทรศัพท์", "โทรศัพท์", "phone", "tel", "mobile"],
    role: ["สิทธิ์", "สิทธิ", "role", "permission", "permissions", "access"]
  };

  const normalizedAliases = {};
  Object.keys(aliases).forEach(function(key) {
    normalizedAliases[key] = aliases[key].map(normalizeHeaderKey);
  });

  const resolved = Object.assign({}, defaults);
  const found = { email: false, name: false, position: false, unit: false, phone: false, role: false };
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
      const userName = String(row[col.name] || "").trim();
      return {
        email: rowEmail,
        name: userName,
        position: String(row[col.position] || "").trim(),
        unit: String(row[col.unit] || "").trim(),
        phone: String(row[col.phone] || "").trim(),
        role: String(row[col.role] || "").trim() || "user",
        avatarUrl: includeAvatar ? getAvatarUrl(rowEmail, userName) : ""
      };
    }
  }
  return null;
}

function verifyGoogleIdToken(idToken, email) {
  const token = String(idToken || "").trim();
  if (!token) {
    return { ok: false, message: "ไม่พบโทเคนสำหรับยืนยันตัวตน" };
  }

  try {
    const res = UrlFetchApp.fetch(
      "https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(token),
      { muteHttpExceptions: true }
    );
    if (res.getResponseCode() !== 200) {
      return { ok: false, message: "ยืนยันตัวตนไม่สำเร็จ" };
    }

    const payload = JSON.parse(res.getContentText() || "{}");
    const tokenEmail = normalizeEmail(payload.email);
    const verified = payload.email_verified === true || String(payload.email_verified || "").toLowerCase() === "true";
    const aud = String(payload.aud || "");

    if (!tokenEmail || tokenEmail !== email) {
      return { ok: false, message: "อีเมลในโทเคนไม่ตรงกับผู้ใช้งาน" };
    }
    if (!verified) {
      return { ok: false, message: "บัญชี Google ยังไม่ยืนยันอีเมล" };
    }
    if (ALLOWED_GOOGLE_CLIENT_IDS.length && ALLOWED_GOOGLE_CLIENT_IDS.indexOf(aud) === -1) {
      return { ok: false, message: "โทเคนมาจากแอปที่ไม่อนุญาต" };
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, message: "ตรวจสอบโทเคนล้มเหลว" };
  }
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
  const email = normalizeEmail(data.email);
  const sessionToken = String(data.sessionToken || "").trim();

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

function readReportRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("รายงานผลการปฏิบัติงาน");
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  const rows = [];
  for (let i = 0; i < allData.length; i++) {
    const obj = {};
    for (let j = 0; j < HEADERS.length; j++) {
      obj[HEADERS[j]] = allData[i][j] != null ? String(allData[i][j]) : "";
    }
    rows.push(obj);
  }
  return rows;
}

function toPublicRow(row) {
  const out = {};
  for (let i = 0; i < PUBLIC_REPORT_HEADERS.length; i++) {
    const key = PUBLIC_REPORT_HEADERS[i];
    out[key] = row[key] != null ? String(row[key]) : "";
  }
  return out;
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

  const tokenCheck = verifyGoogleIdToken(data.idToken, email);
  if (!tokenCheck.ok) {
    return jsonResponse({ status: "error", message: tokenCheck.message || "ยืนยันตัวตนไม่สำเร็จ" });
  }

  const user = findUserByEmail(email, true);
  if (user) {
    const sessionToken = createSessionToken(user.email, user.role);
    return jsonResponse({
      status: "success",
      name: user.name,
      position: user.position,
      unit: user.unit,
      phone: user.phone,
      role: user.role,
      avatarUrl: user.avatarUrl,
      sessionToken: sessionToken,
      sessionExpiresIn: SESSION_TTL_SECONDS
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

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("รายงานผลการปฏิบัติงาน") || ss.getActiveSheet();

  const lastRow = sheet.getLastRow();
  const nextNumber = lastRow;
  const authUser = auth.user;

  const row = [
    nextNumber,
    data.timestamp || new Date().toLocaleString('th-TH'),
    authUser.name || data.reporter || "",
    authUser.email || "",
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
  const auth = authorizeRequest(data);
  if (!auth.ok) {
    return jsonResponse({ status: "error", message: auth.message });
  }

  return jsonResponse({ status: "success", data: readReportRows() });
}

function handleGetPublicData(data) {
  const rows = readReportRows();
  const publicRows = rows.map(function(row) {
    return toPublicRow(row);
  });

  return jsonResponse({ status: "success", data: publicRows });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return jsonResponse({ status: "ok", message: "CSD1 Report API v2.2 (Public Dashboard + Session Auth)", version: "2.2" });
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CSD1 Report')
    .addItem('สร้างหัวตาราง + ชีทผู้ใช้', 'initSheet')
    .addToUi();
}
