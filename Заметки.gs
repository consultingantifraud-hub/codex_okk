function sendNotesToAmoCRM() {
  try {
    Logger.log('--- Начало выполнения скрипта ---');

    var dbSheet = SpreadsheetApp.getActive().getSheetByName('БД');
    if (!dbSheet) {
      throw new Error('Лист "БД" не найден');
    }

    var amoToken = dbSheet.getRange('B2').getValue();
    var amoSubdomainRaw = dbSheet.getRange('B3').getValue();
    var dataSheetName = dbSheet.getRange('B4').getValue();

    if (!amoToken || !amoSubdomainRaw) {
      throw new Error('Не указан токен или поддомен');
    }

    var amoSubdomain = String(amoSubdomainRaw).replace(/\.amocrm\.(ru|eu)$/i, '').trim();
    if (!/^[a-z0-9-]+$/i.test(amoSubdomain)) {
      throw new Error('Некорректный поддомен: ' + amoSubdomain);
    }

    var dataSheet = SpreadsheetApp.getActive().getSheetByName(dataSheetName);
    if (!dataSheet) {
      throw new Error('Лист "' + dataSheetName + '" не найден');
    }

    var lastRow = dataSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('Нет данных для обработки');
      return;
    }

    var dataRange = dataSheet.getRange(2, 1, lastRow - 1, 27);
    var data = dataRange.getValues();

    var filtered = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var aaValue = row[26];
      var zValue = row[25];
      if (aaValue && !zValue) {
        filtered.push({ row: row, index: i + 2 });
      }
    }

    Logger.log('Всего строк: ' + (lastRow - 1));
    Logger.log('Подходит для обработки: ' + filtered.length);

    if (filtered.length === 0) {
      Logger.log('Нет строк для обработки');
      return;
    }

    for (var f = 0; f < filtered.length; f++) {
      var item = filtered[f];
      var rowNumber = item.index;
      var rowData = item.row;

      var dealId = rowData[5];
      var transcription = rowData[21];
      var assessment = rowData[22];

      Logger.log('\nОбработка строки ' + rowNumber + ': dealId=' + dealId);

      try {
        if (typeof transcription === 'string' && transcription.trim()) {
          addNoteToAmo(amoSubdomain, amoToken, dealId, transcription.trim());
          Logger.log('Транскрипция отправлена');
        }

        if (typeof assessment === 'string' && assessment.trim()) {
          addNoteToAmo(amoSubdomain, amoToken, dealId, assessment.trim());
          Logger.log('Оценка отправлена');
        }

        dataSheet.getRange(rowNumber, 26).setValue(new Date());
        Logger.log('Отметка времени установлена в колонку Z');

        Utilities.sleep(500);
      } catch (e) {
        Logger.log('Ошибка в строке ' + rowNumber + ': ' + e.message);
        dataSheet.getRange(rowNumber, 26).setValue(new Date());
        Logger.log('Отметка времени установлена в колонку Z после ошибки');
      }
    }

  } catch (e) {
    Logger.log('*** ГЛОБАЛЬНАЯ ОШИБКА: ' + e.message);
  } finally {
    Logger.log('--- Выполнение скрипта завершено ---');
  }
}

function addNoteToAmo(subdomain, apiKey, dealId, noteText) {
  try {
    Logger.log('--- Начало запроса к amoCRM ---');
    var url = 'https://' + subdomain + '.amocrm.ru/api/v4/leads/' + dealId + '/notes';
    var payload = [{
      note_type: 'common',
      text: noteText,
      created_at: Math.floor(Date.now() / 1000)
    }];

    Logger.log('URL: ' + url);
    Logger.log('Payload: ' + JSON.stringify(payload));

    var options = {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var contentText = response.getContentText();
    var result = contentText ? JSON.parse(contentText) : {};

    Logger.log('Код ответа: ' + responseCode);
    Logger.log('Ответ сервера: ' + JSON.stringify(result));

    if (responseCode >= 400) {
      throw new Error('amoCRM error: ' + (result.detail || contentText));
    }

    return result;

  } catch (e) {
    Logger.log('Ошибка API: ' + e.message);
    throw e;
  }
}
