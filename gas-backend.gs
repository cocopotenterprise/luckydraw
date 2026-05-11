// ============================================================
//  Coco 小铺 幸运抽奖查询 — Google Apps Script 后端
//  部署方式：发布 → 部署为 Web 应用程序 → 任何人可访问
// ============================================================

// 修改为你的 Google Sheets ID（网址中 /d/ 后面那一串）
const SHEET_ID = "1JUPM86pRMM0R2Ok9GV-4W0I7dWyxFBKKF1gNlqFP0aM";

function doGet(e) {
  // 处理 CORS 预检（部分浏览器需要）
  const output = handleRequest(e);
  return output;
}

function handleRequest(e) {
  try {
    checkRateLimit();

    const input = (e && e.parameter && e.parameter.q) ? e.parameter.q : "";

    if (!input) {
      return jsonResponse([]);
    }

    const data = getLucky(input);
    return jsonResponse(data);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 主查询逻辑 ────────────────────────────────────────────
function getLucky(input) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("Sheet1");
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  // 读取所有数据（A~G 列）
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const user = normalizeInput(input);

  const result = [];

  for (let i = 0; i < data.length; i++) {
    let dbPhone = String(data[i][0] || "").replace(/\D/g, "");
    dbPhone = stripPrefix(dbPhone);

    const name  = String(data[i][3] || "").trim();
    const first = name ? name.charAt(0).toLowerCase() : "";

    if (dbPhone === user.number && first === user.letter) {
      result.push({
        order:  String(data[i][1] || ""),
        psid:   String(data[i][2] || ""),
        name:   String(data[i][3] || ""),
        amount: String(data[i][4] || ""),
        lucky:  String(data[i][5] || ""),
        count:  String(data[i][6] || "")
      });
    }
  }

  return result;
}

// ── 电话号码标准化 ────────────────────────────────────────
function normalizeInput(input) {
  const str    = String(input).trim();
  const letter = str.slice(-1).toLowerCase();
  const raw    = str.slice(0, -1).replace(/\D/g, "");

  return {
    number: stripPrefix(raw),
    letter: letter
  };
}

function stripPrefix(num) {
  if (num.startsWith("60")) num = num.slice(2);
  if (num.startsWith("0"))  num = num.slice(1);
  if (num.startsWith("65")) num = num.slice(2);
  return num;
}

// ── 防刷（5秒限制）────────────────────────────────────────
function checkRateLimit() {
  const cache   = CacheService.getScriptCache();
  const userKey = Session.getTemporaryActiveUserKey() || "guest";
  const key     = "limit_" + userKey;

  if (cache.get(key)) {
    throw new Error("⛔ 操作太频繁，请5秒后再试");
  }

  cache.put(key, "1", 5);
}
