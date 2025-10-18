function chatToProTalkBot() {
  let rowToProcess = null; // Объявляем переменные снаружи try
  let targetSheet = null;
  let logsSheet = null;
  let chatId = '';
  let lastRequestText = '';
  try {
    console.log("Начало выполнения скрипта");

    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    if (!dbSheet) throw new Error("Лист 'БД' не найден");
   
    const botToken = dbSheet.getRange("B5").getValue();
    const botId = dbSheet.getRange("B6").getValue();
    const targetSheetName = dbSheet.getRange("B4").getValue();
   
    console.log(`Получены настройки: лист ${targetSheetName}, botId=${botId}, token=${botToken}`);
   
    targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
    if (!targetSheet) throw new Error(`Лист '${targetSheetName}' не найден`);
    logsSheet = SpreadsheetApp.getActive().getSheetByName('LOGS');

    const dataRange = targetSheet.getDataRange();
    const values = dataRange.getValues();
   
    // Поиск строки для обработки
    for (let i = 1; i < values.length; i++) {
      const question = values[i][19];
      const startDate = values[i][23];
     
      if (question && !startDate) {
        rowToProcess = i + 1;
        break;
      }
    }
   
    if (!rowToProcess) {
      console.log("Нет строк для обработки");
      return;
    }
   
    console.log(`Обрабатывается строка ${rowToProcess}`);
    const startDateCell = targetSheet.getRange(rowToProcess, 24);
    startDateCell.setValue(new Date());
   
    const questionCell = targetSheet.getRange(rowToProcess, 20);
    let question = questionCell.getValue();
    if (!question) throw new Error("Пустой вопрос в строке " + rowToProcess);
   
    question = '🏁##' + question;
    const requests = question.split("##").map(q => q.trim()).filter(Boolean);
   
    chatId = "chat_" + Date.now();
    const apiUrl = `https://us1.api.pro-talk.ru/api/v1.0/ask/${botToken}`;
    let lastResponse = null;

    const TIMEOUT_MS = 300000;

    for (const q of requests) {
      try {
        console.log(`Отправка запроса: ${q}`);
        lastRequestText = q;
        const response = UrlFetchApp.fetch(apiUrl, {
          method: "post",
          contentType: "application/json",
          muteHttpExceptions: true,
          timeout: TIMEOUT_MS,
          payload: JSON.stringify({
            bot_id: botId,
            chat_id: chatId,
            message: q
          })
        });
       
        if (response.getResponseCode() !== 200) {
          throw new Error(`HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
        }

        const result = JSON.parse(response.getContentText());
        lastResponse = result.done;
        console.log(`Получен ответ: ${lastResponse}`);
        logProTalkResponse(logsSheet, result, chatId, q);
        Utilities.sleep(1000);

      } catch (e) {
        // Логируем но не прерываем выполнение
        console.error(`Ошибка в запросе: ${e.message}`);
        throw e; // Пробрасываем ошибку в основной catch
      }
    }
   
    const cleanResponse = lastResponse ? lastResponse.replace(/\*/g, '') : '';
    targetSheet.getRange(rowToProcess, 22).setValue(cleanResponse);
    console.log("Обработка завершена успешно");
   
  } catch (e) {
    console.error("Произошла ошибка: " + e.message);
    logProTalkError(logsSheet, chatId, e, lastRequestText);

    // Гарантированная запись ошибки
    if (rowToProcess && targetSheet) {
      try {
        targetSheet.getRange(rowToProcess, 22)
                   .setValue("Разговор не распознан");
      } catch (setError) {
        console.error("Ошибка записи статуса: " + setError.message);
      }
    }
  }
}

function logProTalkResponse(logSheet, responseData, chatId, requestText) {
  try {
    if (!logSheet || !responseData || typeof responseData !== 'object') {
      return;
    }

    const directLogs = Array.isArray(responseData.logs) ? responseData.logs
      : Array.isArray(responseData.log_records) ? responseData.log_records
      : Array.isArray(responseData.debug?.logs) ? responseData.debug.logs
      : [];

    const rows = [];

    if (directLogs.length) {
      directLogs.forEach(log => {
        const tokens = log.tokens || {};
        rows.push([
          log.timestamp ? new Date(log.timestamp) : new Date(),
          log.chat_id || chatId || '',
          log.parent_id || requestText || '',
          log.message || log.text || responseData.done || '',
          log.provider || log.source || responseData.provider || '',
          log.conversation_id || log.thread_id || '',
          log.model || responseData.model || '',
          log.response_id || log.id || '',
          log.total_tokens || tokens.total || '',
          log.prompt_tokens || tokens.prompt || '',
          log.completion_tokens || tokens.completion || '',
          formatAttachmentsForLog(log.attachments),
          log.metadata ? JSON.stringify(log.metadata) : (responseData.metadata ? JSON.stringify(responseData.metadata) : ''),
          log.cluster || log.region || responseData.region || ''
        ]);
      });
    } else {
      rows.push([
        new Date(),
        chatId || '',
        requestText || '',
        responseData.done || responseData.message || '',
        responseData.provider || '',
        responseData.conversation_id || '',
        responseData.model || '',
        responseData.response_id || '',
        responseData.total_tokens || '',
        responseData.prompt_tokens || '',
        responseData.completion_tokens || '',
        '',
        responseData.metadata ? JSON.stringify(responseData.metadata) : '',
        responseData.cluster || responseData.region || ''
      ]);
    }

    if (rows.length) {
      const range = logSheet.getRange(logSheet.getLastRow() + 1, 1, rows.length, rows[0].length);
      range.setValues(rows);
    }
  } catch (error) {
    console.error('Ошибка записи логов ProTalk:', error.message);
  }
}

function logProTalkError(logSheet, chatId, error, requestText) {
  try {
    if (!logSheet || !error) {
      return;
    }
    const row = [[
      new Date(),
      chatId || '',
      '',
      `ERROR: ${error.message}`,
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      requestText ? `request: ${requestText}` : '',
      ''
    ]];
    const range = logSheet.getRange(logSheet.getLastRow() + 1, 1, 1, row[0].length);
    range.setValues(row);
  } catch (rangeError) {
    console.error('Ошибка записи ошибки ProTalk:', rangeError.message);
  }
}

function formatAttachmentsForLog(attachments) {
  if (!Array.isArray(attachments) || attachments.length === 0) {
    return '';
  }
  return attachments.map(att => {
    const name = att?.name || att?.filename || 'attachment';
    const link = att?.link || att?.url || '';
    return link ? `${name}: ${link}` : name;
  }).join('\n');
}
