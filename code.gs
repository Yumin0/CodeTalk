// ============================================================
//  CodeTalk — Google Apps Script 後端程式
//  Phase 1: 接收前端請求 → 呼叫 MaiAgent AI → 回傳白話解釋
//  Phase 2: 把翻譯結果存到 Google 試算表（就像筆記本一樣）
//  Phase 2.5: 新增讀取、更新備註、刪除筆記的功能
//
//  使用前準備：
//  1. 在 GAS 編輯器 → 專案設定 → Script Properties 新增：
//       MAIAGENT_API_KEY  = 你的 API 金鑰
//       SHEET_ID          = 你的 Google 試算表 ID
//  2. 部署成 Web App：
//       執行身分：我
//       存取權限：任何人
//  3. 把 Web App 網址貼到 index.html 的 GAS_URL 變數
// ============================================================

// ── 常數設定 ──────────────────────────────────────────────────────────────
function forceAuth() {
  UrlFetchApp.fetch("https://www.google.com");
}

// MaiAgent AI 服務的網址
const MAIAGENT_BASE_URL = "https://api.maiagent.ai";
const CHATBOT_ID        = "ae1ee433-6f1f-4ad4-9254-a8cbf1d717f8"; // 聊天機器人 ID

// Script Properties 裡存放的金鑰名稱（實際值請在 Script Properties 設定，不要寫在這裡）
const API_KEY_PROP  = "MAIAGENT_API_KEY";
const SHEET_ID_PROP = "SHEET_ID";


// ── 程式入口點 ────────────────────────────────────────────────────────────

/**
 * 處理 GET 請求（用來確認服務是否正常運作）
 * 也處理讀取筆記的請求（action=getNotes）
 */
function doGet(e) {
  // 如果有帶 action 參數，就執行對應的功能
  if (e && e.parameter && e.parameter.action === "getNotes") {
    try {
      var notes = getNotes();
      return buildResponse({ notes: notes });
    } catch (err) {
      Logger.log("getNotes error: " + err.message);
      return buildResponse({ error: err.message }, 500);
    }
  }

  // 沒有 action 就回傳狀態確認
  return buildResponse({ status: "ok", message: "CodeTalk GAS is running." });
}

/**
 * 處理 POST 請求
 *
 * 翻譯：  { action: "translate", mode: "error"|"explain"|"term", content: "..." }
 * 存筆記：{ action: "saveNote",  type: "error"|"code"|"term", input: "...", output: "..." }
 * 讀筆記：{ action: "getNotes" }
 * 更新備註：{ action: "updateNote", id: "...", note: "..." }
 * 刪除筆記：{ action: "deleteNote", id: "..." }
 */
function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action || "translate";

    if (action === "saveNote") {
      // 存入新筆記
      var type   = body.type   || "error";
      var input  = body.input  || "";
      var output = body.output || "";

      if (!input.trim() || !output.trim()) {
        return buildResponse({ error: "input 和 output 不可為空" }, 400);
      }

      saveNote(type, input, output);
      return buildResponse({ success: true });

    } else if (action === "getNotes") {
      // 讀取所有筆記
      var notes = getNotes();
      return buildResponse({ notes: notes });

    } else if (action === "updateNote") {
      // 更新某一筆筆記的備註文字
      var id   = body.id   || "";
      var note = body.note || "";

      if (!id) {
        return buildResponse({ error: "id 不可為空" }, 400);
      }

      updateNote(id, note);
      return buildResponse({ success: true });

    } else if (action === "deleteNote") {
      // 刪除某一筆筆記
      var id = body.id || "";

      if (!id) {
        return buildResponse({ error: "id 不可為空" }, 400);
      }

      deleteNote(id);
      return buildResponse({ success: true });

    } else {
      // action === "translate"（預設：把程式碼翻譯成白話文）
      var mode    = body.mode    || "error";
      var content = body.content || "";

      if (!content.trim()) {
        return buildResponse({ error: "內容不可為空" }, 400);
      }

      var llmResult = callMaiAgent(mode, content);
      return buildResponse({ result: llmResult });
    }

  } catch (err) {
    Logger.log("doPost error: " + err.message);
    return buildResponse({ error: err.message }, 500);
  }
}


// ── MaiAgent AI 呼叫 ────────────────────────────────────────────────────

/**
 * 呼叫 MaiAgent 聊天機器人，回傳 AI 的回覆文字。
 *
 * 驗證方式：Authorization: Api-Key {金鑰}
 * 端點：POST /api/chatbots/{chatbot_id}/completions/
 * 必填欄位：message.content（要問 AI 的內容）
 *
 * AI 的系統提示詞設定在 MaiAgent 後台，不在這裡。
 */
function callMaiAgent(mode, content) {
  var apiKey = PropertiesService.getScriptProperties().getProperty(API_KEY_PROP);
  if (!apiKey) {
    throw new Error("API 金鑰未設定，請至 Script Properties 新增 " + API_KEY_PROP);
  }

  // 在問題前面加上模式標籤，讓 AI 知道要用哪種方式回答
  var modeLabel =
    (mode === "error")   ? "【錯誤訊息模式】"   :
    (mode === "explain") ? "【程式碼解讀模式】" :
    (mode === "term")    ? "【名詞解釋模式】"   :
    "";  // 其他內部用途（例如產生標籤）不需要標籤
  var userMessage = modeLabel ? modeLabel + "\n\n" + content : content;

  // 組合要送給 MaiAgent 的請求內容
  var payload = {
    message: {
      content: userMessage   // 必填：使用者輸入的問題
    },
    isStreaming: false        // GAS 無法處理串流回應，必須設為 false
  };

  var url = MAIAGENT_BASE_URL + "/api/chatbots/" + CHATBOT_ID + "/completions/";

  var options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": "Api-Key " + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true   // 讓我們自己處理 HTTP 錯誤
  };

  var response = UrlFetchApp.fetch(url, options);
  var code     = response.getResponseCode();
  var raw      = response.getContentText();

  Logger.log("MaiAgent status: " + code);
  Logger.log("MaiAgent response: " + raw.substring(0, 800));

  if (code === 401) {
    throw new Error("API 金鑰無效或缺失（401 Unauthorized）");
  }
  if (code !== 200 && code !== 201) {
    throw new Error("MaiAgent 回傳錯誤 " + code + "：" + raw.substring(0, 300));
  }

  var json = JSON.parse(raw);

  // 從 AI 回傳的 JSON 中取出回覆文字（相容不同格式）
  var reply =
    (json.message  && json.message.content)  ? json.message.content  :
    (json.reply)                             ? json.reply             :
    (json.response)                          ? json.response          :
    (json.content)                           ? json.content           :
    raw; // 最後備用：直接回傳原始內容，方便除錯

  return reply;
}


// ── Google 試算表操作 — Phase 2 ────────────────────────────────────────

/**
 * 把一筆翻譯筆記存到 Google 試算表。
 * 試算表 ID 從 Script Properties 讀取（鍵名：SHEET_ID）。
 * 欄位順序：id | type | input | output | tags | created_at | note
 *
 * @param {string} type    "error"（錯誤）| "code"（程式碼）| "term"（名詞）
 * @param {string} input   使用者輸入的原始內容
 * @param {string} output  AI 翻譯的結果
 */
function saveNote(type, input, output) {
  var sheetId = PropertiesService.getScriptProperties().getProperty(SHEET_ID_PROP);
  if (!sheetId) {
    throw new Error("試算表 ID 未設定，請至 Script Properties 新增 SHEET_ID");
  }

  var ss    = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheets()[0]; // 使用第一個工作表

  // 如果試算表是空的，先寫入欄位標題列
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["id", "type", "input", "output", "tags", "created_at", "note"]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold");
  }

  var now       = new Date();
  var id        = now.getTime().toString(); // 用時間戳記當作唯一 ID
  var createdAt = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var tags      = generateTags(type, input, output);

  sheet.appendRow([
    id,
    type,
    input.substring(0, 2000),
    stripMarkdown(output).substring(0, 2000),
    tags,
    createdAt,
    ""   // note — 備註欄位，預設空白，使用者可以之後填入
  ]);
}

/**
 * 讀取試算表裡所有的筆記，依照建立時間由新到舊排列。
 * 回傳一個陣列，每個元素是一筆筆記的資料物件。
 *
 * 就像把整本筆記本的內容讀出來，最新的放最前面。
 */
function getNotes() {
  var sheetId = PropertiesService.getScriptProperties().getProperty(SHEET_ID_PROP);
  if (!sheetId) {
    throw new Error("試算表 ID 未設定，請至 Script Properties 新增 SHEET_ID");
  }

  var ss        = SpreadsheetApp.openById(sheetId);
  var sheet     = ss.getSheets()[0];
  var lastRow   = sheet.getLastRow();

  // 試算表是空的，或只有標題列，就回傳空陣列
  if (lastRow <= 1) {
    return [];
  }

  // 讀取所有資料（從第 2 列開始，跳過標題列）
  var data   = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var notes  = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    // 跳過 id 欄位是空的列（可能是空白列）
    if (!row[0]) continue;

    notes.push({
      id:         String(row[0]),  // 唯一識別碼
      type:       String(row[1]),  // error / code / term
      input:      String(row[2]),  // 使用者輸入的原始內容
      output:     String(row[3]),  // AI 翻譯結果
      tags:       String(row[4]),  // 標籤（逗號分隔）
      created_at: row[5] instanceof Date
                    ? Utilities.formatDate(row[5], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
                    : String(row[5]),  // 建立時間
      note:       String(row[6])   // 使用者的備註
    });
  }

  // 依照建立時間由新到舊排列（最新的排最前面）
  notes.sort(function(a, b) {
    return b.created_at.localeCompare(a.created_at);
  });

  return notes;
}

/**
 * 更新某一筆筆記的備註文字。
 * 就像在筆記本某一頁的空白處，寫下自己的心得。
 *
 * @param {string} id    要更新的筆記 ID
 * @param {string} note  要寫入的備註文字
 */
function updateNote(id, note) {
  var sheetId = PropertiesService.getScriptProperties().getProperty(SHEET_ID_PROP);
  if (!sheetId) {
    throw new Error("試算表 ID 未設定");
  }

  var ss    = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheets()[0];
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    throw new Error("找不到 id: " + id);
  }

  // 讀取 id 欄（第 1 欄），找到對應的列
  var idCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < idCol.length; i++) {
    if (String(idCol[i][0]) === String(id)) {
      // 找到了！更新第 7 欄（note 欄）
      // i+2 是因為：+1 跳過標題列，+1 因為陣列從 0 開始但試算表列從 1 開始
      sheet.getRange(i + 2, 7).setValue(note);
      return;
    }
  }

  throw new Error("找不到 id: " + id);
}

/**
 * 從試算表中刪除指定 id 的那一筆筆記。
 * 就像把筆記本的某一頁撕掉。
 *
 * @param {string} id  要刪除的筆記 ID
 */
function deleteNote(id) {
  var sheetId = PropertiesService.getScriptProperties().getProperty(SHEET_ID_PROP);
  if (!sheetId) {
    throw new Error("試算表 ID 未設定");
  }

  var ss    = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheets()[0];
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    throw new Error("找不到 id: " + id);
  }

  // 讀取 id 欄（第 1 欄），找到對應的列
  var idCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < idCol.length; i++) {
    if (String(idCol[i][0]) === String(id)) {
      // 找到了！刪除整列
      sheet.deleteRow(i + 2);
      return;
    }
  }

  throw new Error("找不到 id: " + id);
}

/**
 * 去除常見的 Markdown 標記符號，讓文字存進試算表時更好閱讀。
 * 處理：標題、粗體、斜體、程式碼區塊、分隔線、清單符號、連結等。
 */
function stripMarkdown(text) {
  return text
    .replace(/```[\s\S]*?```/g, function(m) {        // 程式碼區塊 → 只保留內容
      return m.replace(/```[^\n]*\n?/g, "").trim();
    })
    .replace(/^#{1,6}\s+/gm, "")                     // ## 標題
    .replace(/\*\*(.+?)\*\*/g, "$1")                 // **粗體**
    .replace(/\*(.+?)\*/g, "$1")                     // *斜體*
    .replace(/__(.+?)__/g, "$1")                     // __粗體__
    .replace(/_(.+?)_/g, "$1")                       // _斜體_
    .replace(/`(.+?)`/g, "$1")                       // `行內程式碼`
    .replace(/^[-*]{3,}\s*$/gm, "")                  // --- 分隔線
    .replace(/^[\-\*\+]\s+/gm, "• ")                 // - 清單 → 圓點
    .replace(/^\d+\.\s+/gm, function(m) { return m; }) // 保留數字清單
    .replace(/\[(.+?)\]\(.+?\)/g, "$1")              // [連結文字](網址) → 只保留文字
    .replace(/\n{3,}/g, "\n\n")                      // 多餘的空白列
    .trim();
}

/**
 * 讓 AI 根據翻譯內容，產生 3~5 個英文關鍵字標籤。
 * 就像幫筆記貼上分類貼紙，方便以後搜尋。
 * 失敗時回傳空字串（不影響存檔）。
 */
function generateTags(type, input, output) {
  try {
    var prompt =
      "請根據以下內容產生 3~5 個英文關鍵字標籤，只回傳逗號分隔的關鍵字，不要其他說明。\n\n" +
      "類型：" + type + "\n" +
      "問題：" + input.substring(0, 400) + "\n" +
      "解答：" + output.substring(0, 400);

    var raw = callMaiAgent("tags", prompt);

    // 整理格式：去除換行、中文逗號、空格，合併重複的逗號
    var tags = raw
      .replace(/\n/g, ",")
      .replace(/[，。]/g, ",")
      .replace(/\s+/g, "")
      .replace(/,+/g, ",")
      .replace(/^,|,$/g, "");

    return tags.substring(0, 200);
  } catch (err) {
    Logger.log("generateTags error: " + err.message);
    return "";
  }
}


// ── 回應格式工具 ──────────────────────────────────────────────────────────

/**
 * 建立一個 JSON 格式的回應，讓前端可以正確接收資料。
 * GAS Web App 部署為「任何人」時，CORS 由 Google 自動處理。
 */
function buildResponse(data, statusCode) {
  var output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}


// ── 本機測試用工具 ──────────────────────────────────────────────────────

/**
 * 可在 GAS 編輯器直接執行這個函式，測試 AI 連線是否正常。
 * 部署 Web App 前先跑一次確認。
 */
function testCallMaiAgent() {
  var result = callMaiAgent(
    "error",
    "TypeError: Cannot read properties of undefined (reading 'map')\nat ProductList.jsx:24"
  );
  Logger.log("=== TEST RESULT ===");
  Logger.log(result);
}
