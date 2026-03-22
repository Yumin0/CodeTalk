// ============================================================
//  CodeTalk — Google Apps Script Backend
//  Phase 1: Receive request from frontend → call MaiAgent LLM
//           → return plain-language explanation
//
//  Setup checklist:
//  1. In GAS editor → Project Settings → Script Properties, add:
//       MAIAGENT_API_KEY  = your_api_key_here
//  2. Deploy as Web App:
//       Execute as: Me
//       Who has access: Anyone
//  3. Copy the Web App URL into index.html → GAS_URL
// ============================================================

// ── Constants ──────────────────────────────────────────────────────────────
function forceAuth() {
  UrlFetchApp.fetch("https://www.google.com");
}

// MaiAgent chatbot endpoint
// Format: POST /api/chatbots/{chatbot_id}/completions/
const MAIAGENT_BASE_URL = "https://api.maiagent.ai";
const CHATBOT_ID        = "ae1ee433-6f1f-4ad4-9254-a8cbf1d717f8"; // Chatbot ID (from API 串接頁面)

// Script Property key name (actual value stored in Properties, NOT here)
const API_KEY_PROP = "MAIAGENT_API_KEY";

// Script Property key for Google Sheet ID (set via Project Settings → Script Properties)
const SHEET_ID_PROP = "SHEET_ID";


// ── Entry points ────────────────────────────────────────────────────────────

/**
 * Handle GET requests (health check / CORS preflight helper).
 */
function doGet(e) {
  return buildResponse({ status: "ok", message: "CodeTalk GAS is running." });
}

/**
 * Handle POST requests from the frontend.
 *
 * Translate:  { action: "translate", mode: "error"|"explain"|"term", content: "..." }
 * Save note:  { action: "saveNote",  type: "error"|"code"|"term", input: "...", output: "..." }
 */
function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action || "translate";

    if (action === "saveNote") {
      var type   = body.type   || "error";
      var input  = body.input  || "";
      var output = body.output || "";

      if (!input.trim() || !output.trim()) {
        return buildResponse({ error: "input 和 output 不可為空" }, 400);
      }

      saveNote(type, input, output);
      return buildResponse({ success: true });

    } else {
      // action === "translate" (default)
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


// ── MaiAgent API ────────────────────────────────────────────────────────────

/**
 * Call the MaiAgent chatbot API and return the assistant's reply text.
 *
 * MaiAgent authentication : Authorization: Api-Key {key}
 * Endpoint                : POST /api/chatbots/{chatbot_id}/completions/
 * Required body fields    : message.content  (string)
 * Optional body fields    : conversation     (uuid, leave blank → new conversation each time)
 *                           isStreaming       (boolean, keep false for GAS)
 *
 * The system prompt lives inside the MaiAgent chatbot config — NOT here.
 */
function callMaiAgent(mode, content) {
  var apiKey = PropertiesService.getScriptProperties().getProperty(API_KEY_PROP);
  if (!apiKey) {
    throw new Error("API 金鑰未設定，請至 Script Properties 新增 " + API_KEY_PROP);
  }

  // Prepend mode label so the chatbot knows which task to perform
  var modeLabel =
    (mode === "error")   ? "【錯誤訊息模式】"   :
    (mode === "explain") ? "【程式碼解讀模式】" :
    (mode === "term")    ? "【名詞解釋模式】"   :
    "";  // no label for internal calls (e.g. generateTags)
  var userMessage = modeLabel ? modeLabel + "\n\n" + content : content;

  // ── Correct MaiAgent /completions/ request body ──────────────────────────
  var payload = {
    message: {
      content: userMessage   // required: string
      // attachments: []     // optional: file list, not needed for Phase 1
      // sender: ""          // optional: Contact UUID
    },
    // conversation: ""      // optional: UUID — omit to start a fresh conversation each time
    isStreaming: false        // must be false; GAS cannot handle streaming responses
  };
  // ─────────────────────────────────────────────────────────────────────────

  var url = MAIAGENT_BASE_URL + "/api/chatbots/" + CHATBOT_ID + "/completions/";

  var options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": "Api-Key " + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true   // handle HTTP errors ourselves
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

  // ── Extract reply text ───────────────────────────────────────────────────
  // MaiAgent /completions/ typical response shapes (check GAS log to confirm):
  //   { message: { content: "..." } }
  //   { reply: "..." }
  //   { response: "..." }
  var reply =
    (json.message  && json.message.content)  ? json.message.content  :
    (json.reply)                             ? json.reply             :
    (json.response)                          ? json.response          :
    (json.content)                           ? json.content           :
    raw; // last resort: return raw JSON string so we can inspect it

  return reply;
}


// ── Google Sheets — Phase 2 ────────────────────────────────────────────────

/**
 * Save a translation note to the Google Sheet.
 * Sheet ID is read from Script Properties (key: SHEET_ID).
 * Columns: id | type | input | output | tags | created_at | note
 *
 * @param {string} type    "error" | "code" | "term"
 * @param {string} input   original user input
 * @param {string} output  LLM translation result
 */
function saveNote(type, input, output) {
  var sheetId = PropertiesService.getScriptProperties().getProperty(SHEET_ID_PROP);
  if (!sheetId) {
    throw new Error("試算表 ID 未設定，請至 Script Properties 新增 SHEET_ID");
  }

  var ss    = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheets()[0]; // use the first (and only) sheet

  // Write header row if sheet is empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["id", "type", "input", "output", "tags", "created_at", "note"]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold");
  }

  var now       = new Date();
  var id        = now.getTime().toString();
  var createdAt = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var tags      = generateTags(type, input, output);

  sheet.appendRow([
    id,
    type,
    input.substring(0, 2000),
    stripMarkdown(output).substring(0, 2000),
    tags,
    createdAt,
    ""   // note — blank, for user to fill later
  ]);
}

/**
 * Strip common Markdown syntax so output reads cleanly in Google Sheets.
 * Handles: headings, bold, italic, inline code, code fences, hr, list bullets.
 */
function stripMarkdown(text) {
  return text
    .replace(/```[\s\S]*?```/g, function(m) {        // fenced code blocks → keep content only
      return m.replace(/```[^\n]*\n?/g, "").trim();
    })
    .replace(/^#{1,6}\s+/gm, "")                     // ## headings
    .replace(/\*\*(.+?)\*\*/g, "$1")                 // **bold**
    .replace(/\*(.+?)\*/g, "$1")                     // *italic*
    .replace(/__(.+?)__/g, "$1")                     // __bold__
    .replace(/_(.+?)_/g, "$1")                       // _italic_
    .replace(/`(.+?)`/g, "$1")                       // `inline code`
    .replace(/^[-*]{3,}\s*$/gm, "")                  // --- hr
    .replace(/^[\-\*\+]\s+/gm, "• ")                 // - list → bullet
    .replace(/^\d+\.\s+/gm, function(m) { return m; }) // keep numbered lists
    .replace(/\[(.+?)\]\(.+?\)/g, "$1")              // [link text](url) → text
    .replace(/\n{3,}/g, "\n\n")                      // collapse excess blank lines
    .trim();
}

/**
 * Ask LLM to produce 3-5 comma-separated English keyword tags
 * based on the translation type, input, and output.
 * Returns empty string on failure (non-fatal).
 */
function generateTags(type, input, output) {
  try {
    var prompt =
      "請根據以下內容產生 3~5 個英文關鍵字標籤，只回傳逗號分隔的關鍵字，不要其他說明。\n\n" +
      "類型：" + type + "\n" +
      "問題：" + input.substring(0, 400) + "\n" +
      "解答：" + output.substring(0, 400);

    var raw = callMaiAgent("tags", prompt);

    // Normalise: strip newlines / Chinese commas / spaces, collapse multiple commas
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


// ── CORS / response helper ──────────────────────────────────────────────────

/**
 * Build a JSON ContentService response with CORS headers.
 * GAS Web Apps deployed as "Anyone" handle OPTIONS automatically,
 * but we still set Content-Type explicitly.
 */
function buildResponse(data, statusCode) {
  var output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}


// ── Local test helper ───────────────────────────────────────────────────────

/**
 * Run this function manually in the GAS editor to test API connectivity
 * before deploying as a Web App.
 */
function testCallMaiAgent() {
  var result = callMaiAgent(
    "error",
    "TypeError: Cannot read properties of undefined (reading 'map')\nat ProductList.jsx:24"
  );
  Logger.log("=== TEST RESULT ===");
  Logger.log(result);
}
