function chatToProTalkBot() {
  var rowToProcess = null;
  var targetSheet = null;
  try {
    console.log('Начало выполнения скрипта');

    var dbSheet = SpreadsheetApp.getActive().getSheetByName('БД');
    if (!dbSheet) {
      throw new Error("Лист 'БД' не найден");
    }

    var botToken = dbSheet.getRange('B5').getValue();
    var botId = dbSheet.getRange('B6').getValue();
    var targetSheetName = dbSheet.getRange('B4').getValue();

    console.log('Получены настройки: лист ' + targetSheetName + ', botId=' + botId + ', token=' + botToken);

    targetSheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
    if (!targetSheet) {
      throw new Error("Лист '" + targetSheetName + "' не найден");
    }

    var dataRange = targetSheet.getDataRange();
    var values = dataRange.getValues();

    for (var i = 1; i < values.length; i++) {
      var question = values[i][19];
      var startDate = values[i][23];
      if (question && !startDate) {
        rowToProcess = i + 1;
        break;
      }
    }

    if (!rowToProcess) {
      console.log('Нет строк для обработки');
      return;
    }

    console.log('Обрабатывается строка ' + rowToProcess);
    var startDateCell = targetSheet.getRange(rowToProcess, 24);
    startDateCell.setValue(new Date());

    var questionCell = targetSheet.getRange(rowToProcess, 20);
    var questionValue = questionCell.getValue();
    if (!questionValue) {
      throw new Error('Пустой вопрос в строке ' + rowToProcess);
    }

    var questionPrepared = '🏁##' + questionValue;
    var rawRequests = questionPrepared.split('##');
    var requests = [];
    for (var j = 0; j < rawRequests.length; j++) {
      var part = rawRequests[j].trim();
      if (part) {
        requests.push(part);
      }
    }

    var chatId = 'chat_' + Date.now();
    var apiUrl = 'https://us1.api.pro-talk.ru/api/v1.0/ask/' + botToken;
    var lastResponse = null;
    var TIMEOUT_MS = 300000;

    for (var r = 0; r < requests.length; r++) {
      var q = requests[r];
      console.log('Отправка запроса: ' + q);
      var response = UrlFetchApp.fetch(apiUrl, {
        method: 'post',
        contentType: 'application/json',
        muteHttpExceptions: true,
        timeout: TIMEOUT_MS,
        payload: JSON.stringify({
          bot_id: botId,
          chat_id: chatId,
          message: q
        })
      });

      if (response.getResponseCode() !== 200) {
        throw new Error('HTTP ' + response.getResponseCode() + ': ' + response.getContentText());
      }

      var result = JSON.parse(response.getContentText());
      lastResponse = result.done;
      console.log('Получен ответ: ' + lastResponse);
      Utilities.sleep(1000);
    }

    var cleanResponse = lastResponse ? lastResponse.replace(/\*/g, '') : '';
    targetSheet.getRange(rowToProcess, 22).setValue(cleanResponse);
    console.log('Обработка завершена успешно');

  } catch (e) {
    console.error('Произошла ошибка: ' + e.message);
    if (rowToProcess && targetSheet) {
      try {
        targetSheet.getRange(rowToProcess, 22).setValue('Разговор не распознан');
      } catch (setError) {
        console.error('Ошибка записи статуса: ' + setError.message);
      }
    }
    throw e;
  }
}
