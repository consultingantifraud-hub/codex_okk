function ToProTalkBot() {
  let rowToProcess = null; // Переменная для хранения строки
  let targetSheet = null;
  let logsSheet = null;
  let chatId = '';
  let lastRequestText = '';
  try {
    console.log("Начало выполнения скрипта");
   
    // Получение настроек из листа БД
    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    if (!dbSheet) throw new Error("Лист 'БД' не найден");
   
    const botToken = dbSheet.getRange("B5").getValue();
    const botId = dbSheet.getRange("B6").getValue();
    const targetSheetName = dbSheet.getRange("B4").getValue();
   
    console.log(`Настройки: лист ${targetSheetName}, botId=${botId}, token=${botToken}`);
   
    // Получение целевого листа
    targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
    if (!targetSheet) throw new Error(`Лист '${targetSheetName}' не найден`);
    logsSheet = SpreadsheetApp.getActive().getSheetByName('LOGS');
   
    // Поиск первой незавершенной строки (начиная со второй)
    const dataRange = targetSheet.getDataRange();
    const values = dataRange.getValues();
   
    // Обновленные индексы колонок:
    // U = 20 (индекс 20), Y = 24 (индекс 24)
    for (let i = 1; i < values.length; i++) { 
      const question = values[i][20]; // Колонка U (индекс 20)
      const startDate = values[i][24]; // Колонка Y (индекс 24)
     
      if (question && !startDate) {
        rowToProcess = i + 1; // Номер строки в таблице
        break;
      }
    }
   
    if (!rowToProcess) {
      console.log("Нет строк для обработки");
      return;
    }
   
    console.log(`Обрабатывается строка ${rowToProcess}`);
   
    // Запись даты начала в колонку Y (25-й столбец)
    const startDateCell = targetSheet.getRange(rowToProcess, 25); // Колонка Y
    startDateCell.setValue(new Date());
   
    // Получение вопроса из колонки U (21-й столбец)
    const questionCell = targetSheet.getRange(rowToProcess, 21); // Колонка U
    let question = questionCell.getValue();
    if (!question) throw new Error("Пустой вопрос в строке " + rowToProcess);
   
    // Подготовка запросов
    question = '🏁##' + question;
    const requests = question.split("##").map(q => q.trim()).filter(Boolean);
   
    // Настройки API
    chatId = "chat_" + Date.now();
    const apiUrl = `https://us1.api.pro-talk.ru/api/v1.0/ask/${botToken}`;
    let lastResponse = null;
   
    // Таймаут 5 минут
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
        console.error(`Ошибка в запросе: ${e.message}`);
        throw e; // Пробрасываем в основной catch
      }
    }
   
    // Запись ответа в колонку W (23-й столбец)
    const cleanResponse = lastResponse ? lastResponse.replace(/\*/g, '') : '';
    targetSheet.getRange(rowToProcess, 23).setValue(cleanResponse); // Колонка W
   
    console.log("Обработка завершена успешно");
   
  } catch (e) {
    console.error("Произошла ошибка: " + e.message);
    logProTalkError(logsSheet, chatId, e, lastRequestText);

    // Гарантированная запись ошибки в колонку W
    if (rowToProcess && targetSheet) {
      try {
        targetSheet.getRange(rowToProcess, 23)
                   .setValue("Разговор не распознан"); // Колонка W
      } catch (setError) {
        console.error("Ошибка записи статуса: " + setError.message);
      }
    }
  }
}
