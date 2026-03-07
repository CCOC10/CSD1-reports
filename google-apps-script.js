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
const CHANGE_REQUEST_SHEET_NAME = "คำขอแก้ไขรายการ";
const RECYCLE_BIN_SHEET_NAME = "Recycle Bin";
const RECYCLED_ROW_COLOR = "#242928";
const EDIT_CHANGED_CELL_COLOR = "#03FF00";
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

function findReportRowIndexByTarget(sheet, target) {
  if (!sheet || sheet.getLastRow() < 2) return -1;
  const no = String(target.targetNo || target.no || "").trim();
  const ts = String(target.targetTimestamp || target.timestamp || "").trim();
  const reporterEmail = normalizeEmail(target.targetReporterEmail || target.reporterEmail || "");
  const unitKey = normalizeUnitKey(target.targetUnit || target.unit || "");

  const idxNo = HEADERS.indexOf("ลำดับ");
  const idxTs = HEADERS.indexOf("ประทับเวลา");
  const idxEmail = HEADERS.indexOf("อีเมล");
  const idxUnit = HEADERS.indexOf("ชุดปฏิบัติการ");
  if (idxNo < 0 || idxTs < 0 || idxEmail < 0 || idxUnit < 0) return -1;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  for (let i = 0; i < values.length; i++) {
    const rowNo = String(values[i][idxNo] == null ? "" : values[i][idxNo]).trim();
    const rowTs = String(values[i][idxTs] == null ? "" : values[i][idxTs]).trim();
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

function isAdminUserRole(role) {
  return String(role || "").trim().toLowerCase() === "admin";
}

function isStandardUserRole(role) {
  return String(role || "").trim().toLowerCase() === "user";
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
    data.timestamp || formatThaiDateTimeStandard(new Date()),
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
    normalizeCourtValue(data.arrestCourt || ""),
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
    normalizeCourtValue(data.searchCourt || ""),
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

function handleGetUnitHistory(data) {
  const auth = authorizeRequest(data);
  if (!auth.ok) {
    return jsonResponse({ status: "error", message: auth.message });
  }

  const authUser = auth.user || {};
  if (!isStandardUserRole(authUser.role)) {
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

  let unitRows = [];
  if (unitKey) {
    unitRows = rows.filter(function(row) {
      return normalizeUnitKey(row["ชุดปฏิบัติการ"] || "") === unitKey;
    });
  }

  // Fallback: ถ้าไม่พบข้อมูลจาก unit ให้แสดงข้อมูลที่ผู้ใช้คนนี้เคยกรอกแทน
  const effectiveRows = (unitRows.length ? unitRows : myRows).slice();
  effectiveRows.reverse(); // newest first
  const enrichedRows = effectiveRows.map(function(row) {
    const out = Object.assign({}, row);
    const key = makeChangeRequestTargetKey(
      row["ลำดับ"] || "",
      row["ประทับเวลา"] || "",
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
  if (!isStandardUserRole(authUser.role) && !isAdminUserRole(authUser.role)) {
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
    const rowTs = String(row["ประทับเวลา"] || "").trim();
    if (targetNo && rowNo !== targetNo) return false;
    if (targetTimestamp && rowTs !== targetTimestamp) return false;
    return !!(rowNo || rowTs);
  });

  if (!matchedRow) {
    return jsonResponse({ status: "error", message: "ไม่พบรายการอ้างอิงในระบบ" });
  }

  const targetKey = makeChangeRequestTargetKey(
    String(matchedRow["ลำดับ"] || targetNo || "").trim(),
    String(matchedRow["ประทับเวลา"] || targetTimestamp || "").trim(),
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

  const requesterUnit = normalizeUnitKey(String(authUser.unit || "").trim());
  const targetUnitKey = normalizeUnitKey(String(matchedRow["ชุดปฏิบัติการ"] || "").trim());
  if (isStandardUserRole(authUser.role) && requesterUnit && targetUnitKey && requesterUnit !== targetUnitKey) {
    return jsonResponse({ status: "error", message: "ไม่สามารถส่งคำขอข้ามชุดปฏิบัติการได้" });
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
    String(matchedRow["ประทับเวลา"] || targetTimestamp || "").trim(),
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
  if (!isAdminUserRole(authUser.role)) {
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
  if (!isAdminUserRole(authUser.role)) {
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

  const statusCol = CHANGE_REQUEST_HEADERS.indexOf("status") + 1;
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
  return jsonResponse({ status: "ok", message: "CSD1 Report API v2.9 (Public Dashboard + Session Auth + Unit History + Change Requests + Admin Notifications)", version: "2.9" });
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CSD1 Report')
    .addItem('สร้างหัวตาราง + ชีทผู้ใช้', 'initSheet')
    .addToUi();
}
