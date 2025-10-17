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
