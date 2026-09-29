/**
 * 保健品管理系統 - Google 試算表專屬後端 API (全面適配版)
 */

const SHEET_ITEMS = "Supplements";
const SHEET_META = "AppMeta";

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(15000);

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
      saveAllData(dataToSave);
      return jsonpResponse(e, { status: "success", timestamp: new Date().toISOString() });
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

  // 讀取打卡與提醒紀錄 (若有)
  let checkins = {};
  let reminders = {};
  const metaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_META);
  if (metaSheet) {
    const metaRows = metaSheet.getDataRange().getValues();
    for (let i = 0; i < metaRows.length; i++) {
      if (metaRows[i][0] === "checkins" && metaRows[i][1]) {
        try { checkins = JSON.parse(metaRows[i][1]); } catch(e){}
      }
      if (metaRows[i][0] === "reminders" && metaRows[i][1]) {
        try { reminders = JSON.parse(metaRows[i][1]); } catch(e){}
      }
    }
  }

  return {
    status: "success",
    data: {
      items: items,
      checkins: checkins,
      reminders: reminders
    }
  };
}

// 儲存所有資料
function saveAllData(data) {
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

  // 備份打卡紀錄與提醒時間
  const metaSheet = getOrCreateSheet(SHEET_META, ["Key", "JSON"]);
  metaSheet.clearContents();
  metaSheet.appendRow(["Key", "JSON"]);
  if (data.checkins) {
    metaSheet.appendRow(["checkins", JSON.stringify(data.checkins)]);
  }
  if (data.reminders) {
    metaSheet.appendRow(["reminders", JSON.stringify(data.reminders)]);
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
