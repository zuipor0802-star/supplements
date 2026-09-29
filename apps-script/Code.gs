/**
 * 保健品管理系統 - Google 試算表專屬後端 API (全面適配版)
 */

const SHEET_ITEMS = "Supplements";
const SHEET_META = "AppMeta";
const SHEET_CHECKINS = "Checkins";
const CHECKIN_HEADERS = ["日期", "品項ID", "品名", "時間", "顆數"];
// 舊版把打卡紀錄整包 JSON 存在 AppMeta 的一格（有 50,000 字元上限），搬移後改名保留為備份
const META_KEY_LEGACY_CHECKINS = "checkins";
const META_KEY_LEGACY_BACKUP = "checkins_legacy_backup";
// 資料版本號：每次儲存 +1；前端上傳時附上所依據的版本，不一致代表其他裝置已更新過
const META_KEY_REVISION = "revision";

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return jsonpResponse(e, { status: "error", message: "伺服器忙碌中，請稍後再試" });
  }

  try {
    let action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getData";
    let payload = null;

    // 解析資料內容 (相容 POST 與 GET)
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (err) {
        // 如果是 URL-encoded
        if (e.parameter && e.parameter.data) {
          payload = JSON.parse(e.parameter.data);
        }
      }
    } else if (e && e.parameter && e.parameter.data) {
      payload = JSON.parse(e.parameter.data);
    }

    if (payload && payload.action) {
      action = payload.action;
    }

    // 1. 讀取資料
    if (action === "getData" || action === "get") {
      return jsonpResponse(e, loadAllData());
    }

    // 2. 儲存資料
    if (action === "saveAll" || action === "save") {
      const dataToSave = (payload && payload.data) ? payload.data : payload;
      if (!dataToSave) {
        return jsonpResponse(e, { status: "error", message: "無有效內容" });
      }
      // 舊版前端不會送 baseRevision，照舊接受；新版前端的版本不一致就拒絕，避免覆蓋其他裝置的修改
      const currentRevision = loadRevision();
      if (payload && Object.prototype.hasOwnProperty.call(payload, "baseRevision") &&
          payload.baseRevision !== currentRevision) {
        return jsonpResponse(e, { status: "conflict", revision: currentRevision, message: "雲端資料已被其他裝置更新" });
      }
      const newRevision = currentRevision + 1;
      saveAllData(dataToSave, newRevision);
      return jsonpResponse(e, { status: "success", revision: newRevision, timestamp: new Date().toISOString() });
    }

    return jsonpResponse(e, { status: "error", message: "未知指令: " + action });

  } catch (err) {
    return jsonpResponse(e, { status: "error", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateSheet(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// 讀取所有資料
function loadAllData() {
  const sheet = getOrCreateSheet(SHEET_ITEMS, [
    "ID", "品名", "單罐規格", "未拆罐數", "散裝顆數", "預計採購罐數", "每日用量", "服用時段", "小備註", "購買連結", "最後更新"
  ]);

  const rows = sheet.getDataRange().getValues();
  const items = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] && !r[1]) continue;

    items.push({
      id: r[0],
      name: r[1],
      capacity: Number(r[2]) || 60,
      bottles: Number(r[3]) || 0,
      loose: Number(r[4]) || 0,
      purchased: Number(r[5]) || 0,
      daily: Number(r[6]) || 1,
      slot: r[7] || "lunch",
      notes: r[8] || "",
      url: r[9] || ""
    });
  }

  // 讀取打卡紀錄 (Checkins 工作表，一次打卡一列)
  let checkins = loadCheckinRows();

  // 讀取提醒時間；Checkins 工作表還沒資料時，沿用舊版 AppMeta 的打卡 JSON
  let reminders = {};
  const metaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_META);
  if (metaSheet) {
    const metaRows = metaSheet.getDataRange().getValues();
    for (let i = 0; i < metaRows.length; i++) {
      if (metaRows[i][0] === META_KEY_LEGACY_CHECKINS && metaRows[i][1] && checkins === null) {
        try { checkins = JSON.parse(metaRows[i][1]); } catch(e){}
      }
      if (metaRows[i][0] === "reminders" && metaRows[i][1]) {
        try { reminders = JSON.parse(metaRows[i][1]); } catch(e){}
      }
    }
  }
  if (!checkins) checkins = {};

  return {
    status: "success",
    data: {
      items: items,
      checkins: checkins,
      reminders: reminders,
      revision: loadRevision()
    }
  };
}

// 儲存所有資料
function loadRevision() {
  const metaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_META);
  if (!metaSheet) return 0;
  const metaRows = metaSheet.getDataRange().getValues();
  for (let i = 0; i < metaRows.length; i++) {
    if (metaRows[i][0] === META_KEY_REVISION) return Number(metaRows[i][1]) || 0;
  }
  return 0;
}

function saveAllData(data, revision) {
  const items = data.items || [];
  const sheet = getOrCreateSheet(SHEET_ITEMS, [
    "ID", "品名", "單罐規格", "未拆罐數", "散裝顆數", "預計採購罐數", "每日用量", "服用時段", "小備註", "購買連結", "最後更新"
  ]);

  // 清除舊資料（保留標題）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 11).clearContent();
  }

  if (items.length > 0) {
    const now = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
    const rows = items.map(item => [
      item.id || Date.now(),
      item.name || "",
      Number(item.capacity) || 60,
      Number(item.bottles) || 0,
      Number(item.loose) || 0,
      Number(item.purchased) || 0,
      Number(item.daily) || 1,
      item.slot || "lunch",
      item.notes || "",
      item.url || "",
      now
    ]);
    sheet.getRange(2, 1, rows.length, 11).setValues(rows);
  }

  if (data.checkins) {
    saveCheckinRows(data.checkins, items);
  }

  // 提醒時間；舊版打卡 JSON 在打卡列寫入後改名為備份，之後不再更新
  const metaSheet = getOrCreateSheet(SHEET_META, ["Key", "JSON"]);
  const metaRows = metaSheet.getDataRange().getValues();
  let legacyCheckins = null;
  let legacyBackup = null;
  for (let i = 0; i < metaRows.length; i++) {
    if (metaRows[i][0] === META_KEY_LEGACY_CHECKINS) legacyCheckins = metaRows[i][1];
    if (metaRows[i][0] === META_KEY_LEGACY_BACKUP) legacyBackup = metaRows[i][1];
  }
  if (data.checkins) {
    if (!legacyBackup) legacyBackup = legacyCheckins;
    legacyCheckins = null;
  }

  // 先一次寫入新內容，再清掉多出來的舊列；中途出錯也不會先把備份清空
  const metaOut = [["Key", "JSON"]];
  metaOut.push([META_KEY_REVISION, revision]);
  if (data.reminders) {
    metaOut.push(["reminders", JSON.stringify(data.reminders)]);
  }
  if (legacyCheckins) {
    metaOut.push([META_KEY_LEGACY_CHECKINS, legacyCheckins]);
  }
  if (legacyBackup) {
    metaOut.push([META_KEY_LEGACY_BACKUP, legacyBackup]);
  }
  ensureRows(metaSheet, metaOut.length);
  metaSheet.getRange(1, 1, metaOut.length, 2).setValues(metaOut);
  clearOutside(metaSheet, metaOut.length, 2);
}

// 清掉 (numRows × numCols) 範圍右邊與下面殘留的舊內容
function clearOutside(sheet, numRows, numCols) {
  const lastCol = sheet.getLastColumn();
  if (lastCol > numCols && numRows > 0) {
    sheet.getRange(1, numCols + 1, numRows, lastCol - numCols).clearContent();
  }
  const staleRows = sheet.getLastRow() - numRows;
  if (staleRows > 0) {
    sheet.getRange(numRows + 1, 1, staleRows, Math.max(numCols, sheet.getLastColumn())).clearContent();
  }
}

// getRange().setValues() 不會自動加列（新工作表預設 1,000 列），寫入前先補足
function ensureRows(sheet, totalRows) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < totalRows) {
    sheet.insertRowsAfter(maxRows, totalRows - maxRows);
  }
}

// 讀取 Checkins 工作表 → { "yyyy-MM-dd": { 品項ID: { time, daily } } }；工作表不存在或沒有資料時回傳 null
function loadCheckinRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CHECKINS);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const tz = ss.getSpreadsheetTimeZone();
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, CHECKIN_HEADERS.length).getValues();
  const checkins = {};

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (r[0] === "" || r[1] === "") continue;

    // 防止試算表把日期、時間字串自動轉成日期物件
    const date = (r[0] instanceof Date) ? Utilities.formatDate(r[0], tz, "yyyy-MM-dd") : String(r[0]);
    const time = (r[3] instanceof Date) ? Utilities.formatDate(r[3], tz, "HH:mm") : String(r[3]);
    const itemId = String(r[1]);

    if (!checkins[date]) checkins[date] = {};
    checkins[date][itemId] = {
      time: time,
      daily: r[4] === "" ? 1 : Number(r[4])
    };
  }
  return checkins;
}

// 以 { 日期: { 品項ID: { time, daily } } } 整份覆寫 Checkins 工作表
function saveCheckinRows(checkins, items) {
  const sheet = getOrCreateSheet(SHEET_CHECKINS, CHECKIN_HEADERS);
  const nameById = {};
  items.forEach(item => { nameById[String(item.id)] = item.name || ""; });

  const rows = [];
  Object.keys(checkins).sort().forEach(date => {
    const day = checkins[date] || {};
    Object.keys(day).forEach(itemId => {
      const rec = day[itemId] || {};
      rows.push([
        date,
        itemId,
        nameById[itemId] || "",
        rec.time || "",
        rec.daily === undefined ? 1 : Number(rec.daily)
      ]);
    });
  });

  // 先寫入新內容，再清掉多出來的舊列；中途出錯也不會先把紀錄清空
  if (rows.length > 0) {
    ensureRows(sheet, rows.length + 1);
    const range = sheet.getRange(2, 1, rows.length, CHECKIN_HEADERS.length);
    // 日期、品項ID、時間設為純文字，避免被自動轉成日期或數字
    range.setNumberFormat("@");
    sheet.getRange(2, 5, rows.length, 1).setNumberFormat("0");
    range.setValues(rows);
  }
  const staleRows = sheet.getLastRow() - (rows.length + 1);
  if (staleRows > 0) {
    sheet.getRange(rows.length + 2, 1, staleRows, CHECKIN_HEADERS.length).clearContent();
  }
}

// 回傳 JSON 或 JSONP
function jsonpResponse(e, data) {
  const callback = e && e.parameter && e.parameter.callback;
  let output = JSON.stringify(data);
  let mime = ContentService.MimeType.JSON;

  if (callback) {
    output = callback + "(" + output + ");";
    mime = ContentService.MimeType.JAVASCRIPT;
  }

  return ContentService.createTextOutput(output).setMimeType(mime);
}
