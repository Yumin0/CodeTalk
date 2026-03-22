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

// Google Sheet ID for Phase 2+ logging (create the sheet, paste its ID here)
const SHEET_ID = "YOUR_GOOGLE_SHEET_ID_HERE";


// ── Entry points ────────────────────────────────────────────────────────────

/**
 * Handle GET requests (health check / CORS preflight helper).
 */
function doGet(e) {
  return buildResponse({ status: "ok", message: "CodeTalk GAS is running." });
}

/**
 * Handle POST requests from the frontend.
 * Expected body: { mode: "error" | "explain", content: "..." }
 */
function doPost(e) {
  try {
    // --- Parse request body ---
    var body = JSON.parse(e.postData.contents);
    var mode    = body.mode    || "error";
    var content = body.content || "";

    if (!content.trim()) {
      return buildResponse({ error: "內容不可為空" }, 400);
    }

    // --- Call MaiAgent ---
    var llmResult = callMaiAgent(mode, content);

    // --- (Phase 2) Log to Sheet ---
    // logToSheet(mode, content, llmResult);

    return buildResponse({ result: llmResult });

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
  var modeLabel   = (mode === "error") ? "【錯誤訊息模式】" : "【程式碼解讀模式】";
  var userMessage = modeLabel + "\n\n" + content;

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


// ── Google Sheets logging (Phase 2) ────────────────────────────────────────

/**
 * Append one row to the CodeTalk Log sheet.
 * Columns: Timestamp | Mode | Input | Output | CharCount
 *
 * Uncomment the call in doPost() when ready for Phase 2.
 */
function logToSheet(mode, input, output) {
  try {
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("Log") || ss.insertSheet("Log");

    // Write header row if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "Mode", "Input", "Output", "InputLength", "OutputLength"]);
      sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
    }

    sheet.appendRow([
      new Date(),
      mode,
      input.substring(0, 1000),   // truncate to avoid cell limit
      output.substring(0, 1000),
      input.length,
      output.length
    ]);
  } catch (err) {
    Logger.log("logToSheet error: " + err.message);
    // Non-fatal — don't fail the main request because of logging
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
