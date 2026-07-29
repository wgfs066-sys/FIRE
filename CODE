const CONFIG = {
  USER_MAP_ID: "1h09QBxE0gzwAeF6Lb4GlCY3o9YCf6xAr92tz9QvPqcY",     // 人員驗證表
  EXTINGUISHER_ID: "1sV7Me2CZlV0PmY0qX3xYnvsHzZFQ3751IYu6f1vfBoA", // 滅火器表
  SHEET_ID_MAP: "ID",                                             // 驗證表分頁
  SHEET_EXTINGUISHER: "館內滅火器",                                 // 滅火器分頁
  TZ: "Asia/Taipei"                                               //
};

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("滅火器檢查系統")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    // 🌟【關鍵核心】：開啟 XFrame 存取權限，瀏覽器才會允許 iframe 內調用即時相機串流！
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- 1. 登入驗證 ---
function checkLoginById(userId) {
  try {
    const ws = SpreadsheetApp.openById(CONFIG.USER_MAP_ID).getSheetByName(CONFIG.SHEET_ID_MAP); //
    const data = ws.getRange("A2:B" + ws.getLastRow()).getValues();                             //
    const user = data.find(r => String(r[1]).trim() === String(userId).trim());                 //
    
    if (user) {
      return { ok: true, name: user[0], id: user[1] }; //
    }
    return { ok: false, msg: "查無此員工姓名或密碼" };
  } catch (e) {
    return { ok: false, msg: "驗證失敗：" + e.toString() };
  }
}

// --- 2. 修改密碼 / ID ---
function changeUserId(oldId, newId) {
  try {
    const ws = SpreadsheetApp.openById(CONFIG.USER_MAP_ID).getSheetByName(CONFIG.SHEET_ID_MAP); //
    const data = ws.getRange("A2:B" + ws.getLastRow()).getValues();
    const rowIdx = data.findIndex(r => String(r[1]).trim() === String(oldId).trim());
    if (rowIdx > -1) {
      ws.getRange(rowIdx + 2, 2).setValue(String(newId).trim());
      return { ok: true };
    }
    return { ok: false, msg: "舊密碼不對應，請重新嘗試" };
  } catch (e) {
    return { ok: false, msg: String(e) };
  }
}

// --- 3. 取得所有館內滅火器 ---
function getExtinguishers() {
  try {
    const ws = SpreadsheetApp.openById(CONFIG.EXTINGUISHER_ID).getSheetByName(CONFIG.SHEET_EXTINGUISHER);
    const lastRow = ws.getLastRow();
    if (lastRow < 3) return { ok: true, data: [] };
    
    // 第 1-2 列是標題，資料從第 3 列開始
    const data = ws.getRange(3, 1, lastRow - 2, 7).getValues();
    const result = data.map((r, idx) => ({
      row: idx + 3,
      id: String(r[0] || "").trim(),
      floor: String(r[1] || "").trim(),
      location: String(r[2] || "").trim(),
      expiry: r[3] ? formatDate(r[3]) : "",
      pressure: String(r[4] || "").trim(),
      lastDate: r[5] ? formatDate(r[5]) : "",
      lastUser: String(r[6] || "").trim()
    })).filter(item => item.id !== "");
    
    return { ok: true, data: result };
  } catch (e) {
    return { ok: false, msg: String(e) };
  }
}

// --- 4. 確認檢查紀錄 (壓力錶/檢查日/檢查人) ---
function inspectExtinguisher(row, isNormal, userName) {
  try {
    const ws = SpreadsheetApp.openById(CONFIG.EXTINGUISHER_ID).getSheetByName(CONFIG.SHEET_EXTINGUISHER);
    const dateStr = Utilities.formatDate(new Date(), CONFIG.TZ, "yyyy-MM-dd");
    ws.getRange(row, 5).setValue(isNormal ? "正常" : "異常");
    ws.getRange(row, 6).setValue(dateStr);
    ws.getRange(row, 7).setValue(userName);
    return { ok: true };
  } catch (e) {
    return { ok: false, msg: String(e) };
  }
}

// --- 5. 修改滅火器完整資料 ---
function updateExtinguisher(row, info) {
  try {
    const ws = SpreadsheetApp.openById(CONFIG.EXTINGUISHER_ID).getSheetByName(CONFIG.SHEET_EXTINGUISHER);
    ws.getRange(row, 1, 1, 7).setValues([[
      info.id, info.floor, info.location, info.expiry, info.pressure, info.lastDate, info.lastUser
    ]]);
    return { ok: true };
  } catch (e) {
    return { ok: false, msg: String(e) };
  }
}

// --- 6. 新增滅火器 ---
function addExtinguisher(info) {
  try {
    const ws = SpreadsheetApp.openById(CONFIG.EXTINGUISHER_ID).getSheetByName(CONFIG.SHEET_EXTINGUISHER);
    ws.appendRow([info.id, info.floor, info.location, info.expiry, "", "", ""]);
    return { ok: true };
  } catch (e) {
    return { ok: false, msg: String(e) };
  }
}

function formatDate(dateObj) {
  if (!dateObj) return "";
  try {
    if (dateObj instanceof Date) {
      return Utilities.formatDate(dateObj, CONFIG.TZ, "yyyy-MM-dd");
    }
    return String(dateObj).split("T")[0];
  } catch (e) {
    return String(dateObj);
  }
}
