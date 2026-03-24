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
//       DRIVE_FOLDER_ID   = 你的 Google Drive 圖片資料夾 ID（必填，否則圖片會上傳到根目錄）
//                          ➜ 從資料夾網址取得：drive.google.com/drive/folders/{這段就是 ID}
//  2. 部署成 Web App：
//       執行身分：我
//       存取權限：任何人
//  3. 把 Web App 網址貼到 index.html 的 GAS_URL 變數
//  4. 執行 forceAuth() 完成授權（會同時授權 Drive 存取）
//  5. 執行 checkSetup() 確認所有設定正確
// ============================================================

// ── 常數設定 ──────────────────────────────────────────────────────────────

/**
 * 觸發所有必要的授權（UrlFetch、Drive、Spreadsheet）
 * 第一次使用前請在 GAS 編輯器執行此函式
 */
function forceAuth() {
  UrlFetchApp.fetch("https://www.google.com");
  // 觸發 Drive 授權
  DriveApp.getRootFolder();
  // 觸發 Spreadsheet 授權
  var sheetId = PropertiesService.getScriptProperties().getProperty(SHEET_ID_PROP);
  if (sheetId) SpreadsheetApp.openById(sheetId);
  Logger.log("授權完成");
}

/**
 * 撤銷目前的 OAuth 授權 token，讓下次執行時重新申請完整權限
 * 執行此函式後，再執行 forceAuth() 重新授權
 */
function revokeAuth() {
  ScriptApp.invalidateAuth();
  Logger.log("授權已撤銷，請重新執行 forceAuth()");
}

/**
 * 檢查所有 Script Properties 設定是否正確，並驗證 Drive 資料夾存取
 * 設定完 Script Properties 後請執行此函式確認設定無誤
 */
function checkSetup() {
  var props = PropertiesService.getScriptProperties();

  var apiKey   = props.getProperty(API_KEY_PROP);
  var sheetId  = props.getProperty(SHEET_ID_PROP);
  var folderId = props.getProperty(DRIVE_FOLDER_ID_PROP);

  Logger.log("=== CodeTalk 設定檢查 ===");
  Logger.log("MAIAGENT_API_KEY : " + (apiKey   ? "✅ 已設定" : "❌ 未設定"));
  Logger.log("SHEET_ID         : " + (sheetId  ? "✅ 已設定 (" + sheetId + ")" : "❌ 未設定"));
  Logger.log("DRIVE_FOLDER_ID  : " + (folderId ? "✅ 已設定 (" + folderId + ")" : "❌ 未設定（圖片將上傳至根目錄）"));

  if (folderId) {
    try {
      var folder = DriveApp.getFolderById(folderId);
      Logger.log("Drive 資料夾     : ✅ 存取成功 (" + folder.getName() + ")");
    } catch (e) {
      Logger.log("Drive 資料夾     : ❌ 無法存取 — " + e.message);
    }
  }

  if (sheetId) {
    try {
      var ss = SpreadsheetApp.openById(sheetId);
      Logger.log("Google 試算表    : ✅ 存取成功 (" + ss.getName() + ")");
      var mSheet = ss.getSheetByName("Product_Media");
      Logger.log("Product_Media 頁 : " + (mSheet ? "✅ 存在" : "❌ 不存在"));
    } catch (e) {
      Logger.log("Google 試算表    : ❌ 無法存取 — " + e.message);
    }
  }

  Logger.log("========================");
}

// MaiAgent AI 服務的網址
const MAIAGENT_BASE_URL = "https://api.maiagent.ai";
const CHATBOT_ID        = "ae1ee433-6f1f-4ad4-9254-a8cbf1d717f8"; // 聊天機器人 ID

// Script Properties 裡存放的金鑰名稱（實際值請在 Script Properties 設定，不要寫在這裡）
const API_KEY_PROP       = "MAIAGENT_API_KEY";
const SHEET_ID_PROP      = "SHEET_ID";
const DRIVE_FOLDER_ID_PROP = "DRIVE_FOLDER_ID"; // Google Drive 圖片資料夾 ID（選填，未設定則上傳至根目錄）


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

  if (e && e.parameter && e.parameter.action === "getProducts") {
    try {
      var products = getProducts();
      return buildResponse({ products: products });
    } catch (err) {
      Logger.log("getProducts error: " + err.message);
      return buildResponse({ error: err.message }, 500);
    }
  }

  if (e && e.parameter && e.parameter.action === "getTechnologies") {
    try {
      var technologies = getTechnologies();
      return buildResponse({ technologies: technologies });
    } catch (err) {
      Logger.log("getTechnologies error: " + err.message);
      return buildResponse({ error: err.message }, 500);
    }
  }

  if (e && e.parameter && e.parameter.action === "getProductDetail") {
    try {
      var productId = e.parameter.productId || "";
      if (!productId) {
        return buildResponse({ error: "productId 不可為空" }, 400);
      }
      var detail = getProductDetail(productId);
      return buildResponse({ detail: detail });
    } catch (err) {
      Logger.log("getProductDetail error: " + err.message);
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
 * 更新產品：{ action: "updateProductDetail", productId: "...", product: {...}, devNote: {...} }
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

    } else if (action === "updateProductDetail") {
      // 更新產品資料（Product 工作表）及開發筆記（Dev_Notes 工作表）
      var productId = body.productId || "";
      if (!productId) {
        return buildResponse({ error: "productId 不可為空" }, 400);
      }
      updateProductDetail(productId, body.product || {}, body.devNote || {});
      return buildResponse({ success: true });

    } else if (action === "addProduct") {
      // 新增產品（Product 工作表）、開發筆記（Dev_Notes）及新技術標籤（Technologies）
      var pData  = body.product  || {};
      var dnData = body.devNote  || {};
      var newTechs = body.newTechnologies || [];

      if (!pData.name || !String(pData.name).trim()) {
        return buildResponse({ error: "產品名稱不可為空" }, 400);
      }

      var newProductId = addProduct(pData, dnData, newTechs);
      return buildResponse({ success: true, productId: newProductId });

    } else if (action === "uploadProductImage") {
      // 上傳圖片到 Google Drive 並記錄到 Product_Media 工作表
      var productId   = body.productId   || "";
      var imageBase64 = body.imageBase64 || "";
      var mimeType    = body.mimeType    || "image/jpeg";
      var description = body.description || "";
      var order       = body.order !== undefined ? Number(body.order) : 0;

      if (!productId)   return buildResponse({ error: "productId 不可為空" }, 400);
      if (!imageBase64) return buildResponse({ error: "imageBase64 不可為空" }, 400);

      var result = uploadProductImage(productId, imageBase64, mimeType, description, order);
      return buildResponse({ success: true, mediaId: result.mediaId, imageUrl: result.imageUrl });

    } else if (action === "deleteProductMedia") {
      // 從 Product_Media 工作表刪除指定媒體記錄
      var mediaId = body.mediaId || "";
      if (!mediaId) return buildResponse({ error: "mediaId 不可為空" }, 400);
      deleteProductMedia(mediaId);
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
    output.substring(0, 2000),
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


// ── Google 試算表操作 — Product & Technologies ────────────────────────────

/**
 * 讀取 Product 工作表的所有產品資料。
 * 欄位順序：產品ID(A) | 產品名稱(B) | 產品簡介(C) | 技術標籤(D) | 設計者(E) | 部署連結(F)
 */
function getProducts() {
  var sheetId = PropertiesService.getScriptProperties().getProperty(SHEET_ID_PROP);
  if (!sheetId) {
    throw new Error("試算表 ID 未設定，請至 Script Properties 新增 SHEET_ID");
  }

  var ss    = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName("Product");
  if (!sheet) {
    throw new Error("找不到名稱為 'Product' 的工作表");
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var data     = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var products = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0] && !row[1]) continue; // 跳過完全空白的列

    products.push({
      id:         String(row[0]),
      name:       String(row[1]),
      description: String(row[2]),
      tags:       String(row[3]),
      designer:   String(row[4]),
      deployLink: String(row[5])
    });
  }

  return products;
}

/**
 * 讀取 Technologies 工作表的所有技術資料。
 * 欄位順序：技術ID(A) | 技術名稱(B) | 技術分類(C) | 技術說明(D)
 */
function getTechnologies() {
  var sheetId = PropertiesService.getScriptProperties().getProperty(SHEET_ID_PROP);
  if (!sheetId) {
    throw new Error("試算表 ID 未設定，請至 Script Properties 新增 SHEET_ID");
  }

  var ss    = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName("Technologies");
  if (!sheet) {
    throw new Error("找不到名稱為 'Technologies' 的工作表");
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var data  = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var techs = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0] && !row[1]) continue;

    techs.push({
      id:          String(row[0]),
      name:        String(row[1]),
      category:    String(row[2]),
      description: String(row[3])
    });
  }

  return techs;
}


/**
 * 讀取單一產品的完整資料，包含 Dev_Notes 與 Product_Media。
 * 欄位：Product(A-F) + Dev_Notes(A-E) + Product_Media(A-E)
 *
 * @param {string} productId  Product 工作表的產品ID
 */
function getProductDetail(productId) {
  var sheetId = PropertiesService.getScriptProperties().getProperty(SHEET_ID_PROP);
  if (!sheetId) {
    throw new Error("試算表 ID 未設定，請至 Script Properties 新增 SHEET_ID");
  }

  var ss = SpreadsheetApp.openById(sheetId);

  // ── 取得產品基本資料 ──────────────────────────────────────────────────
  var productSheet = ss.getSheetByName("Product");
  if (!productSheet) throw new Error("找不到名稱為 'Product' 的工作表");

  var pLastRow = productSheet.getLastRow();
  var product  = null;
  if (pLastRow > 1) {
    var pData = productSheet.getRange(2, 1, pLastRow - 1, 6).getValues();
    for (var i = 0; i < pData.length; i++) {
      if (String(pData[i][0]) === String(productId)) {
        product = {
          id:          String(pData[i][0]),
          name:        String(pData[i][1]),
          description: String(pData[i][2]),
          tags:        String(pData[i][3]),
          designer:    String(pData[i][4]),
          deployLink:  String(pData[i][5])
        };
        break;
      }
    }
  }
  if (!product) throw new Error("找不到產品 id: " + productId);

  // ── 取得 Dev_Notes（怎麼跟AI溝通、實作遇到的問題、解決方式）────────────
  // 欄位：筆記ID(A) | 產品ID(B) | 怎麼跟AI溝通(C) | 實作時遇到的問題(D) | 解決方式(E)
  var devNotes = [];
  var dnSheet  = ss.getSheetByName("Dev_Notes");
  if (dnSheet && dnSheet.getLastRow() > 1) {
    var dnData    = dnSheet.getRange(2, 1, dnSheet.getLastRow() - 1, 5).getValues();
    for (var j = 0; j < dnData.length; j++) {
      if (String(dnData[j][1]) === String(productId)) {
        devNotes.push({
          id:        String(dnData[j][0]),
          aiTips:    String(dnData[j][2]),
          problems:  String(dnData[j][3]),
          solutions: String(dnData[j][4])
        });
      }
    }
  }

  // ── 取得 Product_Media（圖片） ────────────────────────────────────────
  // 欄位：媒體ID(A) | 產品ID(B) | 圖片連結(C) | 排序(D) | 說明(E)
  var media   = [];
  var mSheet  = ss.getSheetByName("Product_Media");
  if (mSheet && mSheet.getLastRow() > 1) {
    var mData = mSheet.getRange(2, 1, mSheet.getLastRow() - 1, 5).getValues();
    for (var k = 0; k < mData.length; k++) {
      if (String(mData[k][1]) === String(productId)) {
        media.push({
          id:          String(mData[k][0]),
          imageUrl:    String(mData[k][2]),
          order:       Number(mData[k][3]) || 0,
          description: String(mData[k][4])
        });
      }
    }
    media.sort(function(a, b) { return a.order - b.order; });
  }

  return { product: product, devNotes: devNotes, media: media };
}

/**
 * 更新單一產品的基本資料（Product 工作表）與開發筆記（Dev_Notes 工作表）。
 * Dev_Notes 中若找不到對應 productId 的列，就新增一列。
 *
 * @param {string} productId   產品 ID
 * @param {Object} productData 要更新的產品欄位（name/description/tags/designer/deployLink）
 * @param {Object} devNoteData 要更新的開發筆記欄位（aiTips/problems/solutions）
 */
function updateProductDetail(productId, productData, devNoteData) {
  var sheetId = PropertiesService.getScriptProperties().getProperty(SHEET_ID_PROP);
  if (!sheetId) {
    throw new Error("試算表 ID 未設定，請至 Script Properties 新增 SHEET_ID");
  }

  var ss = SpreadsheetApp.openById(sheetId);

  // ── 更新 Product 工作表 ────────────────────────────────────────────────
  var productSheet = ss.getSheetByName("Product");
  if (!productSheet) throw new Error("找不到名稱為 'Product' 的工作表");

  var pLastRow = productSheet.getLastRow();
  if (pLastRow > 1) {
    var pIds = productSheet.getRange(2, 1, pLastRow - 1, 1).getValues();
    for (var i = 0; i < pIds.length; i++) {
      if (String(pIds[i][0]) === String(productId)) {
        var row = i + 2; // +1 for header, +1 for 1-based index
        if (productData.name        !== undefined) productSheet.getRange(row, 2).setValue(productData.name);
        if (productData.description !== undefined) productSheet.getRange(row, 3).setValue(productData.description);
        if (productData.tags        !== undefined) productSheet.getRange(row, 4).setValue(productData.tags);
        if (productData.designer    !== undefined) productSheet.getRange(row, 5).setValue(productData.designer);
        if (productData.deployLink  !== undefined) productSheet.getRange(row, 6).setValue(productData.deployLink);
        break;
      }
    }
  }

  // ── 更新 Dev_Notes 工作表 ──────────────────────────────────────────────
  // 欄位：筆記ID(A) | 產品ID(B) | 怎麼跟AI溝通(C) | 實作時遇到的問題(D) | 解決方式(E)
  var dnSheet = ss.getSheetByName("Dev_Notes");
  if (!dnSheet) return; // 工作表不存在就跳過

  var dnLastRow = dnSheet.getLastRow();
  var found     = false;

  if (dnLastRow > 1) {
    var dnIds = dnSheet.getRange(2, 2, dnLastRow - 1, 1).getValues(); // Column B
    for (var j = 0; j < dnIds.length; j++) {
      if (String(dnIds[j][0]) === String(productId)) {
        var dnRow = j + 2;
        if (devNoteData.aiTips    !== undefined) dnSheet.getRange(dnRow, 3).setValue(devNoteData.aiTips);
        if (devNoteData.problems  !== undefined) dnSheet.getRange(dnRow, 4).setValue(devNoteData.problems);
        if (devNoteData.solutions !== undefined) dnSheet.getRange(dnRow, 5).setValue(devNoteData.solutions);
        found = true;
        break;
      }
    }
  }

  // 找不到對應 productId 的列，新增一列
  if (!found) {
    var newId = new Date().getTime().toString();
    dnSheet.appendRow([
      newId,
      productId,
      devNoteData.aiTips    || "",
      devNoteData.problems  || "",
      devNoteData.solutions || ""
    ]);
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


// ── Google 試算表操作 — 新增產品 ────────────────────────────────────────

/**
 * 取得工作表的下一個流水號 ID（依照 A 欄最大整數值 +1）。
 * 用於確保 Product / Dev_Notes / Technologies 的 ID 欄維持 1, 2, 3... 的順序。
 *
 * @param  {Sheet} sheet  目標工作表
 * @return {number}       下一個可用 ID
 */
function getNextSheetId(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1; // 只有標題列或空表，從 1 開始

  var ids    = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var maxId  = 0;
  for (var i = 0; i < ids.length; i++) {
    var id = parseInt(ids[i][0]);
    if (!isNaN(id) && id > maxId) maxId = id;
  }
  return maxId + 1;
}

/**
 * 新增一筆產品到 Product 工作表、Dev_Notes 工作表，
 * 並把前端帶來的全新技術標籤寫入 Technologies 工作表。
 *
 * @param {Object}   productData      產品欄位（name / description / tags / designer / deployLink）
 * @param {Object}   devNoteData      開發筆記欄位（aiTips / problems / solutions）
 * @param {string[]} newTechnologies  前端新增、尚未存在於 Technologies 工作表的技術標籤名稱陣列
 */
function addProduct(productData, devNoteData, newTechnologies) {
  var sheetId = PropertiesService.getScriptProperties().getProperty(SHEET_ID_PROP);
  if (!sheetId) {
    throw new Error("試算表 ID 未設定，請至 Script Properties 新增 SHEET_ID");
  }

  var ss = SpreadsheetApp.openById(sheetId);

  // ── 新增到 Product 工作表 ──────────────────────────────────────────────
  var productSheet = ss.getSheetByName("Product");
  if (!productSheet) throw new Error("找不到名稱為 'Product' 的工作表");

  var nextProductId = getNextSheetId(productSheet);
  productSheet.appendRow([
    nextProductId,
    productData.name        || "",
    productData.description || "",
    productData.tags        || "",
    productData.designer    || "",
    productData.deployLink  || ""
  ]);

  // ── 新增到 Dev_Notes 工作表 ────────────────────────────────────────────
  // 欄位：筆記ID(A) | 產品ID(B) | 怎麼跟AI溝通(C) | 實作時遇到的問題(D) | 解決方式(E)
  var dnSheet = ss.getSheetByName("Dev_Notes");
  if (dnSheet) {
    var nextNoteId = getNextSheetId(dnSheet);
    dnSheet.appendRow([
      nextNoteId,
      nextProductId,
      devNoteData.aiTips    || "",
      devNoteData.problems  || "",
      devNoteData.solutions || ""
    ]);
  }

  // ── 新增技術標籤到 Technologies 工作表（跳過已存在的） ─────────────────
  // 欄位：技術標籤ID(A) | 技術名稱(B) | 技術分類(C) | 技術說明(D)
  if (newTechnologies && newTechnologies.length > 0) {
    var techSheet = ss.getSheetByName("Technologies");
    if (techSheet) {
      // 先讀取既有技術名稱（小寫比對），避免重複
      var existingNames = [];
      var tLastRow = techSheet.getLastRow();
      if (tLastRow > 1) {
        var tData = techSheet.getRange(2, 2, tLastRow - 1, 1).getValues();
        for (var k = 0; k < tData.length; k++) {
          existingNames.push(String(tData[k][0]).toLowerCase());
        }
      }

      for (var i = 0; i < newTechnologies.length; i++) {
        var techName = String(newTechnologies[i]).trim();
        if (!techName) continue;
        if (existingNames.indexOf(techName.toLowerCase()) !== -1) continue; // 已存在，跳過

        var nextTechId = getNextSheetId(techSheet);
        techSheet.appendRow([nextTechId, techName, "", ""]);
        existingNames.push(techName.toLowerCase()); // 避免同批次內重複
      }
    }
  }

  return nextProductId; // 回傳新產品 ID，供前端後續上傳圖片使用
}


// ── Google Drive 圖片上傳 ──────────────────────────────────────────────────

/**
 * 將 base64 圖片上傳到 Google Drive，設定公開分享，
 * 並將圖片連結記錄到 Product_Media 工作表。
 *
 * @param {string} productId    產品 ID
 * @param {string} imageBase64  Base64 編碼的圖片資料（不含 data URL 前綴）
 * @param {string} mimeType     圖片 MIME 類型（e.g. "image/jpeg"）
 * @param {string} description  圖片說明文字
 * @param {number} order        圖片排列順序
 * @return {{ mediaId: string, imageUrl: string }}
 */
function uploadProductImage(productId, imageBase64, mimeType, description, order) {
  var sheetId = PropertiesService.getScriptProperties().getProperty(SHEET_ID_PROP);
  if (!sheetId) throw new Error("試算表 ID 未設定，請至 Script Properties 新增 SHEET_ID");

  var ss = SpreadsheetApp.openById(sheetId);
  var mSheet = ss.getSheetByName("Product_Media");
  if (!mSheet) throw new Error("找不到 Product_Media 工作表");

  // 組合檔名
  var ext = (mimeType === "image/png") ? ".png" : (mimeType === "image/gif") ? ".gif" : ".jpg";
  var fileName = "product_" + productId + "_" + new Date().getTime() + ext;

  // 建立圖片 Blob 並上傳到 Google Drive
  // 先在根目錄建立（確保有寫入權），再移動到目標資料夾
  var imageBlob = Utilities.newBlob(Utilities.base64Decode(imageBase64), mimeType, fileName);
  var folderId  = PropertiesService.getScriptProperties().getProperty(DRIVE_FOLDER_ID_PROP);
  var file      = DriveApp.createFile(imageBlob);
  if (folderId) {
    try {
      var targetFolder = DriveApp.getFolderById(folderId);
      targetFolder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    } catch (e) {
      // 無法移至目標資料夾（可能只有檢視權），保留在根目錄
      Logger.log("無法移至 DRIVE_FOLDER_ID 資料夾，檔案保留在根目錄：" + e.message);
    }
  }

  // 設定任何人皆可透過連結檢視
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var fileId   = file.getId();
  var imageUrl = "https://drive.google.com/uc?export=view&id=" + fileId;

  // 寫入 Product_Media 工作表
  var nextMediaId = getNextSheetId(mSheet);
  mSheet.appendRow([nextMediaId, productId, imageUrl, order, description]);

  return { mediaId: String(nextMediaId), imageUrl: imageUrl };
}

/**
 * 診斷用：逐步確認 Drive 存取是否正常
 * 先執行這個，確認 Drive 授權和資料夾設定都 OK 後再跑 testUploadProductImage
 */
function debugDriveAccess() {
  Logger.log("=== Drive 存取診斷 ===");

  // 1. 測試基本 Drive 授權
  try {
    var root = DriveApp.getRootFolder();
    Logger.log("✅ Drive 授權正常，根目錄：" + root.getName());
  } catch (e) {
    Logger.log("❌ Drive 基本授權失敗：" + e.message);
    Logger.log("→ 請執行 revokeAuth() 再執行 forceAuth()，然後重試");
    return;
  }

  // 2. 確認 DRIVE_FOLDER_ID 設定
  var folderId = PropertiesService.getScriptProperties().getProperty(DRIVE_FOLDER_ID_PROP);
  if (!folderId) {
    Logger.log("⚠️  DRIVE_FOLDER_ID 未設定，將使用根目錄上傳");
  } else {
    Logger.log("ℹ️  DRIVE_FOLDER_ID = " + folderId);
    try {
      var folder = DriveApp.getFolderById(folderId);
      Logger.log("✅ 資料夾存取正常：" + folder.getName());
    } catch (e) {
      Logger.log("❌ 無法存取指定資料夾：" + e.message);
      Logger.log("→ 請到 Script Properties 更新或刪除 DRIVE_FOLDER_ID");
      return;
    }
  }

  Logger.log("✅ 所有 Drive 檢查通過，可以執行 testUploadProductImage()");
}

/**
 * 測試用：在 GAS 編輯器直接執行此函式來測試圖片上傳
 * 會上傳一張 1x1 的測試圖片到 Drive，確認整個流程是否正常
 */
function testUploadProductImage() {
  // 1x1 透明 PNG 的 base64（不含 data URL 前綴）
  var testBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
  try {
    var result = uploadProductImage("test_product", testBase64, "image/png", "測試圖片", 0);
    Logger.log("✅ 上傳成功！mediaId=" + result.mediaId + "  imageUrl=" + result.imageUrl);
  } catch (e) {
    Logger.log("❌ 上傳失敗：" + e.message);
  }
}

/**
 * 從 Product_Media 工作表刪除指定媒體記錄。
 * 注意：此函式不會從 Google Drive 刪除實際檔案。
 *
 * @param {string} mediaId  要刪除的媒體 ID
 */
function deleteProductMedia(mediaId) {
  var sheetId = PropertiesService.getScriptProperties().getProperty(SHEET_ID_PROP);
  if (!sheetId) throw new Error("試算表 ID 未設定，請至 Script Properties 新增 SHEET_ID");

  var ss = SpreadsheetApp.openById(sheetId);
  var mSheet = ss.getSheetByName("Product_Media");
  if (!mSheet || mSheet.getLastRow() <= 1) return;

  var mLastRow = mSheet.getLastRow();
  var mIds = mSheet.getRange(2, 1, mLastRow - 1, 1).getValues();
  for (var i = 0; i < mIds.length; i++) {
    if (String(mIds[i][0]) === String(mediaId)) {
      mSheet.deleteRow(i + 2);
      return;
    }
  }
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
