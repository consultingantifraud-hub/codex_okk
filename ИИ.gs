function chatToProTalkBot() {
  let rowToProcess = null; // Объявляем переменную снаружи try
  try {
    console.log("Начало выполнения скрипта");

    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    if (!dbSheet) throw new Error("Лист 'БД' не найден");

    const botToken = dbSheet.getRange("B5").getValue();
    const botId = dbSheet.getRange("B6").getValue();
    const targetSheetName = dbSheet.getRange("B4").getValue();

    console.log(`Получены настройки: лист ${targetSheetName}, botId=${botId}, token=${botToken}`);

    const targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
    if (!targetSheet) throw new Error(`Лист '${targetSheetName}' не найден`);

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

    const chatId = "chat_" + Date.now();
    const apiUrl = `https://us1.api.pro-talk.ru/api/v1.0/ask/${botToken}`;
    let lastResponse = null;

    const TIMEOUT_MS = 300000;

    for (const q of requests) {
      try {
        console.log(`Отправка запроса: ${q}`);
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

    // Гарантированная запись ошибки
    if (rowToProcess) {
      try {
        const targetSheet = SpreadsheetApp.getActive().getSheetByName(
          SpreadsheetApp.getActive().getSheetByName("БД").getRange("B4").getValue()
        );
        if (targetSheet) {
          targetSheet.getRange(rowToProcess, 22)
                     .setValue("Разговор не распознан");
        }
      } catch (setError) {
        console.error("Ошибка записи статуса: " + setError.message);
      }
    }
  }
}

function processActiveProTalkRow() {
  const spreadsheet = SpreadsheetApp.getActive();
  let targetSheet = null;
  let rowIndex = null;

  try {
    const dbSheet = spreadsheet.getSheetByName("БД");
    if (!dbSheet) {
      throw new Error("Лист 'БД' не найден");
    }

    const botToken = dbSheet.getRange("B5").getValue();
    const botId = dbSheet.getRange("B6").getValue();
    const targetSheetName = dbSheet.getRange("B4").getValue();

    targetSheet = spreadsheet.getSheetByName(targetSheetName);
    if (!targetSheet) {
      throw new Error(`Лист '${targetSheetName}' не найден`);
    }

    const activeSheet = spreadsheet.getActiveSheet();
    if (activeSheet.getSheetId() !== targetSheet.getSheetId()) {
      throw new Error("Активный лист не соответствует листу для обработки ProTalk");
    }

    const activeRange = activeSheet.getActiveRange();
    if (!activeRange) {
      throw new Error("Не выбрана активная строка");
    }

    rowIndex = activeRange.getRow();
    if (rowIndex <= 1) {
      throw new Error("Выберите строку с данными (начиная со 2-й)");
    }

    console.log(`Обработка активной строки ${rowIndex}`);

    const apiUrl = `https://us1.api.pro-talk.ru/api/v1.0/ask/${botToken}`;
    const TIMEOUT_MS = 300000;

    const handleFlow = function(questionColumn, startColumn, resultColumn, flowName) {
      const questionCell = targetSheet.getRange(rowIndex, questionColumn);
      let questionValue = questionCell.getValue();
      const resultCell = targetSheet.getRange(rowIndex, resultColumn);

      if (!questionValue) {
        console.log(`Колонка с вопросом для ${flowName} пуста, обработка пропущена`);
        return;
      }

      targetSheet.getRange(rowIndex, startColumn).setValue(new Date());

      questionValue = '🏁##' + questionValue;
      const requests = questionValue.split("##")
        .map(function(part) { return part.trim(); })
        .filter(function(part) { return part; });

      if (!requests.length) {
        console.log(`Не удалось сформировать запрос для ${flowName}`);
        resultCell.setValue("Разговор не распознан");
        return;
      }

      let lastResponse = null;
      const chatId = 'chat_' + flowName + '_' + Date.now();

      try {
        for (var i = 0; i < requests.length; i++) {
          const message = requests[i];
          console.log(`Отправка запроса (${flowName}): ${message}`);

          const response = UrlFetchApp.fetch(apiUrl, {
            method: "post",
            contentType: "application/json",
            muteHttpExceptions: true,
            timeout: TIMEOUT_MS,
            payload: JSON.stringify({
              bot_id: botId,
              chat_id: chatId,
              message: message
            })
          });

          if (response.getResponseCode() !== 200) {
            throw new Error(`HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
          }

          const result = JSON.parse(response.getContentText());
          lastResponse = result.done;
          console.log(`Получен ответ (${flowName}): ${lastResponse}`);
          Utilities.sleep(1000);
        }

        const cleanResponse = lastResponse ? lastResponse.replace(/\*/g, '') : '';
        resultCell.setValue(cleanResponse);
      } catch (flowError) {
        console.error(`Ошибка ${flowName}: ${flowError.message}`);
        resultCell.setValue("Разговор не распознан");
      }
    };

    handleFlow(20, 24, 22, 'Транскрибация');
    handleFlow(21, 25, 23, 'Оценка');

    console.log("Активная строка обработана");
  } catch (error) {
    console.error("Ошибка обработки активной строки: " + error.message);
    if (targetSheet && rowIndex && rowIndex > 1) {
      try {
        const fallback = "Разговор не распознан";
        const transcriptionCell = targetSheet.getRange(rowIndex, 22);
        if (!transcriptionCell.getValue()) {
          transcriptionCell.setValue(fallback);
        }
        const assessmentCell = targetSheet.getRange(rowIndex, 23);
        if (!assessmentCell.getValue()) {
          assessmentCell.setValue(fallback);
        }
      } catch (writeError) {
        console.error("Ошибка записи статуса: " + writeError.message);
      }
    }
  }
}
