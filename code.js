/**
 * Single-User Dispatch Bot (Math-Friendly Version)
 */

const CONFIG = {
  sheetName: "Заявки",
  triggerStatus: "новая",
  
  cols: {
    id: 1,       // A
    comment: 2,  // B
    status: 3,   // C
    
    // Logging Columns
    logSent: 5,  // E
    logStart: 6, // F
    logEnd: 7    // G
  },

  status: {
    new: "новая",
    work: "в работе",
    done: "завершено"
  },
  
  colors: {
    new: "#FFF2CC",
    work: "#CFE2F3",
    done: "#D9EAD3"
  }
};

/**
 * Trigger: onEdit
 */
function processEdit(e) {
  try {
    if (!e || !e.range) return;
    
    const sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.sheetName) return;
    
    const col = e.range.getColumn();
    if (col < CONFIG.cols.status || e.range.getLastColumn() > CONFIG.cols.status) return;

    const props = PropertiesService.getScriptProperties();
    const token = props.getProperty('TG_TOKEN');
    const chatId = props.getProperty('TG_CHAT_ID');

    if (!token || !chatId) return;

    const startRow = e.range.getRow();
    const numRows = e.range.getNumRows();

    for (let i = 0; i < numRows; i++) {
      processRow(sheet, startRow + i, token, chatId);
    }
  } catch (err) {
    console.error("processEdit failed:", err);
  }
}

function processRow(sheet, row, token, chatId) {
  const statusCell = sheet.getRange(row, CONFIG.cols.status);
  const logSentCell = sheet.getRange(row, CONFIG.cols.logSent);
  
  const currentStatus = String(statusCell.getValue()).trim().toLowerCase();

  if (currentStatus !== CONFIG.status.new) return;
  if (logSentCell.getValue() !== "") return;

  colorRow(sheet, row, CONFIG.colors.new);

  const garage = sheet.getRange(row, CONFIG.cols.id).getValue();
  const comment = sheet.getRange(row, CONFIG.cols.comment).getValue();

  const keyboard = {
    inline_keyboard: [[
      { text: "🚀 Взять в работу", callback_data: `WORK|${row}` }
    ]]
  };

  const msg = formatMessage(garage, comment, CONFIG.status.new);
  
  if (sendTelegram(token, chatId, msg, keyboard)) {
    // Пишем дату как объект и форматируем ячейку
    logSentCell.setValue(new Date());
    logSentCell.setNumberFormat("HH:mm | dd.MM");
  }
}

/**
 * Webhook Handler
 */
function doPost(e) {
  try {
    const update = JSON.parse(e.postData.contents);
    if (update.callback_query) {
      handleCallback(update.callback_query);
    }
    return ContentService.createTextOutput("OK");
  } catch (err) {
    console.error("Webhook error:", err);
    return ContentService.createTextOutput("Error");
  }
}

function handleCallback(cb) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TG_TOKEN');
  
  const [action, rowStr] = cb.data.split("|");
  const row = parseInt(rowStr);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetName);

  let newStatus, newColor, replyMarkup;
  const timeNow = new Date(); // Берем текущее время как объект

  switch (action) {
    case "WORK":
      newStatus = CONFIG.status.work;
      newColor = CONFIG.colors.work;
      replyMarkup = [[ { text: "🏁 Завершить", callback_data: `DONE|${row}` } ]];
      
      // Пишем время начала
      const cellStart = sheet.getRange(row, CONFIG.cols.logStart);
      cellStart.setValue(timeNow);
      cellStart.setNumberFormat("HH:mm | dd.MM");
      break;

    case "DONE":
      newStatus = CONFIG.status.done;
      newColor = CONFIG.colors.done;
      replyMarkup = [];
      
      // Пишем время конца
      const cellEnd = sheet.getRange(row, CONFIG.cols.logEnd);
      cellEnd.setValue(timeNow);
      cellEnd.setNumberFormat("HH:mm | dd.MM");
      break;

    default:
      return;
  }

  sheet.getRange(row, CONFIG.cols.status).setValue(newStatus);
  colorRow(sheet, row, newColor);

  const garage = sheet.getRange(row, CONFIG.cols.id).getValue();
  const comment = sheet.getRange(row, CONFIG.cols.comment).getValue();
  const msg = formatMessage(garage, comment, newStatus);

  try {
    editMessage(token, cb.message.chat.id, cb.message.message_id, msg, replyMarkup);
    answerCallback(token, cb.id, `Статус: ${newStatus}`);
  } catch (e) {
    console.warn("Failed to update TG message:", e);
  }
}

// --- Helpers ---

function colorRow(sheet, row, color) {
  sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground(color);
}

function formatMessage(garage, comment, status) {
  const escape = (text) => String(text || "—").replace(/[_*[\]()~`>#+=|{}.!-]/g, '\\$&');
  return `🚗 *Заявка*\n🏷 Гаражный: ${escape(garage)}\n💬 Коммент: ${escape(comment)}\n📌 Статус: ${escape(status)}`;
}

// --- API Wrappers ---

function sendTelegram(token, chatId, text, markup) {
  return telegramRequest(token, "sendMessage", {
    chat_id: chatId, text: text, parse_mode: "MarkdownV2", reply_markup: markup
  });
}

function editMessage(token, chatId, msgId, text, keyboard) {
  const payload = { chat_id: chatId, message_id: msgId, text: text, parse_mode: "MarkdownV2" };
  if (keyboard) payload.reply_markup = { inline_keyboard: keyboard };
  return telegramRequest(token, "editMessageText", payload);
}

function answerCallback(token, id, text) {
  return telegramRequest(token, "answerCallbackQuery", { callback_query_id: id, text: text });
}

function telegramRequest(token, method, payload) {
  try {
    const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };
    const response = UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/${method}`, options);
    const json = JSON.parse(response.getContentText());
    if (!json.ok) return false;
    return true;
  } catch (e) { return false; }
}

/**
 * Setup Function (Run once)
 */
function setupBot() {
  const token = "YOUR_BOT_TOKEN";
  const chatId = "YOUR_CHAT_ID";
  const webAppUrl = "YOUR_WEB_APP_URL"; // From Deploy -> Web App

  const props = PropertiesService.getScriptProperties();
  props.setProperty('TG_TOKEN', token);
  props.setProperty('TG_CHAT_ID', chatId);
  
  if (webAppUrl && webAppUrl.includes("script.google.com")) {
     UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/setWebhook?url=${webAppUrl}`);
     Logger.log("Webhook set successfully");
  }
}
