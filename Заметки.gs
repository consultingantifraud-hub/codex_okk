function sendNotesToAmoCRM() {
  try {
    Logger.log('--- Начало выполнения скрипта ---');
    
    // Загрузка параметров из листа "БД"
    const dbSheet = SpreadsheetApp.getActive().getSheetByName("БД");
    if (!dbSheet) throw new Error('Лист "БД" не найден');
    
    const amoToken = dbSheet.getRange("B2").getValue();
    let amoSubdomain = dbSheet.getRange("B3").getValue();
    const dataSheetName = dbSheet.getRange("B4").getValue();

    // Проверка обязательных параметров
    if (!amoToken || !amoSubdomain) throw new Error('Не указан токен или поддомен');
    amoSubdomain = amoSubdomain.replace(/\.amocrm\.(ru|eu)$/i, '').trim();
    
    if (!/^[a-z0-9-]+$/i.test(amoSubdomain)) {
      throw new Error(`Некорректный поддомен: ${amoSubdomain}`);
    }

    // Загрузка листа с данными
    const dataSheet = SpreadsheetApp.getActive().getSheetByName(dataSheetName);
    if (!dataSheet) throw new Error(`Лист "${dataSheetName}" не найден`);

    // Чтение данных за один запрос
    const lastRow = dataSheet.getLastRow();
    const dataRange = dataSheet.getRange(2, 1, lastRow - 1, 27); // Столбцы A-AA
    const data = dataRange.getValues();
    
    // Фильтрация данных: оставляем только строки где AA заполнено и Z пустое
    const filteredData = data
      .map((row, index) => ({ row, index: index + 2 })) // Сохраняем оригинальный номер строки
      .filter(item => {
        const aaValue = item.row[26]; // Столбец AA (индекс 26 в массиве)
        const zValue = item.row[25];  // Столбец Z (индекс 25 в массиве)
        return (aaValue && !zValue);  // AA заполнено И Z пустое
      });

    Logger.log(`Всего строк: ${lastRow - 1}`);
    Logger.log(`Подходит для обработки: ${filteredData.length}`);

    // Если нет подходящих строк - завершаем выполнение
    if (filteredData.length === 0) {
      Logger.log('Нет строк для обработки');
      return;
    }

    // Обработка отфильтрованных данных
    filteredData.forEach(item => {
      const rowNumber = item.index;
      const rowData = item.row;
      
      const dealId = rowData[5];        // Столбец F
      const transcription = rowData[21]; // Столбец V
      const assessment = rowData[22];   // Столбец W

      Logger.log(`\nОбработка строки ${rowNumber}: dealId=${dealId}`);

      try {
        // Отправка транскрипции
        if (typeof transcription === 'string' && transcription.trim()) {
          addNoteToAmo(amoSubdomain, amoToken, dealId, transcription.trim());
          Logger.log('Транскрипция отправлена');
        }

        // Отправка оценки
        if (typeof assessment === 'string' && assessment.trim()) {
          addNoteToAmo(amoSubdomain, amoToken, dealId, assessment.trim());
          Logger.log('Оценка отправлена');
        }

        // Устанавливаем отметку времени в Z
        dataSheet.getRange(rowNumber, 26).setValue(new Date());
        Logger.log('Отметка времени установлена в колонку Z');

        // Задержка между запросами
        Utilities.sleep(500); // 0.5 секунд

      } catch (e) {
        Logger.log(`Ошибка в строке ${rowNumber}: ${e.message}`);
        // Устанавливаем отметку времени даже при ошибке
        dataSheet.getRange(rowNumber, 26).setValue(new Date());
        Logger.log('Отметка времени установлена в колонку Z после ошибки');
      }
    });

  } catch (e) {
    Logger.log(`*** ГЛОБАЛЬНАЯ ОШИБКА: ${e.message}`);
  } finally {
    Logger.log('--- Выполнение скрипта завершено ---');
  }
}

// Функция отправки заметки в amoCRM
function addNoteToAmo(subdomain, apiKey, dealId, noteText) {
  try {
    Logger.log(`--- Начало запроса к amoCRM ---`);
    const url = `https://${subdomain}.amocrm.ru/api/v4/leads/${dealId}/notes`;
    const payload = [{
      "note_type": "common",
      "text": noteText,
      "created_at": Math.floor(Date.now() / 1000)
    }];

    Logger.log(`URL: ${url}`);
    Logger.log(`Payload: ${JSON.stringify(payload)}`);

    const options = {
      'method': 'post',
      'headers': { 
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const result = JSON.parse(response.getContentText());

    Logger.log(`Код ответа: ${responseCode}`);
    Logger.log(`Ответ сервера: ${JSON.stringify(result)}`);

    if (responseCode >= 400) {
      throw new Error(`amoCRM error: ${result.detail || response.getContentText()}`);
    }

    return result;

  } catch (e) {
    Logger.log(`Ошибка API: ${e.message}`);
    throw e;
  }
}
