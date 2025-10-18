// Конфигурация
function getConfigFromSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = spreadsheet.getSheetByName("БД");
    if (!configSheet) throw new Error('Лист "БД" не найден');
    // Чтение исключений из колонки G
    const excludedFields = readExcludedFields(configSheet);
    return {
      AMO_TOKEN: configSheet.getRange("B2").getValue(),
      AMO_SUBDOMAIN: configSheet.getRange("B3").getValue(),
      SHEET_NAME: configSheet.getRange("B4").getValue(),
      SPREADSHEET_ID: spreadsheet.getId(),
      callStatusMap: Object.fromEntries(
        configSheet.getRange("C:D").getValues()
          .filter(row => row[0] && row[1])
          .map(([id, text]) => [id, text])
      ),
      taskTypesMap: Object.fromEntries(
        configSheet.getRange("E:F").getValues()
          .filter(row => row[0] && row[1])
          .map(([id, text]) => [id, text])
      ),
      excludedFields: excludedFields // Список исключенных полей
    };
  } catch (error) {
    console.error('❌ Ошибка конфигурации:', error.message);
    throw error;
  }
}
const config = getConfigFromSheet();
const userCache = {};
const entityCache = {};
let pipelinesCache = null;
let processedCallIdsCache = null;

// Чтение исключенных полей
function readExcludedFields(sheet) {
  const excludedRange = sheet.getRange("G2:G");
  const values = excludedRange.getValues();
  return values
    .filter(row => row[0] !== "")
    .map(row => row[0].trim());
}

// Форматирование времени
function formatTimestamp(timestamp) {
  if (!timestamp) return "Нет данных";
  const date = new Date(timestamp * 1000);
  return Utilities.formatDate(date, "GMT+3", "dd.MM.yyyy HH:mm");
}

// Проверка обработки звонка
function isCallProcessed(callNoteId) {
  try {
    if (!callNoteId) {
      return false;
    }
    const processedIds = getProcessedCallIds();
    return processedIds.has(String(callNoteId));
  } catch (error) {
    console.error('❌ Ошибка проверки:', error.message);
    return false;
  }
}

// Универсальный запрос к API
function amoRequest(url, method = 'GET', body = null) {
  try {
    const options = {
      method: method,
      headers: { 'Authorization': `Bearer ${config.AMO_TOKEN}` },
      muteHttpExceptions: true
    };
    if (body) options.payload = JSON.stringify(body);
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    if (statusCode === 204) {
      console.log(`✅ Пустой ответ (204) для ${url}`);
      return null;
    }
    if (statusCode !== 200) {
      console.error(`❌ HTTP ${statusCode}: ${response.getContentText()}`);
      return null;
    }
    return JSON.parse(response.getContentText());
  } catch (error) {
    console.error(`❌ Ошибка запроса к ${url}:`, error.message);
    return null;
  }
}

// Получение задач для сделки
function getAmoTasks(leadId) {
  try {
    const url = `https://${config.AMO_SUBDOMAIN}/api/v4/tasks?filter[entity_id]=${leadId}&filter[entity_type]=leads`;
    const tasksResponse = amoRequest(url);
    if (!tasksResponse) return 'Нет задач';
    const tasks = tasksResponse._embedded?.tasks || [];
    return tasks.map(task => {
      const taskType = config.taskTypesMap[task.task_type_id] || `Неизвестный тип (${task.task_type_id})`;
      const dueDate = formatTimestamp(task.complete_till);
      const status = task.is_completed === true ? 'Выполнена' :
        (task.is_completed === false ? 'Не выполнена' : 'Статус не определен');
      return `• ${taskType}: ${task.text || 'Нет текста'}\n  Срок: ${dueDate}\n  Статус: ${status}`;
    }).join('\n') || 'Нет задач';
  } catch (error) {
    console.error(`❌ Ошибка задач для сделки ${leadId}:`, error.message);
    return 'Ошибка при получении задач';
  }
}

// Сбор всех полей из сделки (с фильтрацией)
function getAllEntityFields(lead) {
  try {
    const fields = [];
    // Поля сделки
    fields.push('--- Сделка ---');
    fields.push(`ID: ${lead.id}`);
    fields.push(`Ответственный: ${getAmoUser(lead.responsible_user_id).name}`);
    // Фильтрация кастомных полей сделки
    const customFieldsDeal = (lead.custom_fields_values || [])
      .map(field => {
        const fieldName = field.field_name;
        if (config.excludedFields.includes(fieldName)) return null;
        const values = field.values
          .map(v => v.value)
          .filter(v => v)
          .join(', ');
        return values ? `${fieldName}: ${values}` : null;
      })
      .filter(Boolean);
    fields.push(...customFieldsDeal);
    // Поля контактов
    if (lead._embedded?.contacts?.length) {
      lead._embedded.contacts.forEach(contactId => {
        const contact = getEntityData('contacts', contactId.id);
        if (contact) {
          fields.push('--- Контакт ---');
          fields.push(`ID: ${contact.id}`);
          fields.push(`Ответственный: ${contact.responsible}`);
          // Фильтрация кастомных полей контакта
          const customFieldsContact = (contact.custom_fields_values || [])
            .map(field => {
              const fieldName = field.field_name;
              if (config.excludedFields.includes(fieldName)) return null;
              const values = field.values
                .map(v => v.value)
                .filter(v => v)
                .join(', ');
              return values ? `${fieldName}: ${values}` : null;
            })
            .filter(Boolean);
          fields.push(...customFieldsContact);
        }
      });
    }
    // Поля компаний
    if (lead._embedded?.companies?.length) {
      lead._embedded.companies.forEach(companyId => {
        const company = getEntityData('companies', companyId.id);
        if (company) {
          fields.push('--- Компания ---');
          fields.push(`ID: ${company.id}`);
          fields.push(`Ответственный: ${company.responsible}`);
          // Фильтрация кастомных полей компании
          const customFieldsCompany = (company.custom_fields_values || [])
            .map(field => {
              const fieldName = field.field_name;
              if (config.excludedFields.includes(fieldName)) return null;
              const values = field.values
                .map(v => v.value)
                .filter(v => v)
                .join(', ');
              return values ? `${fieldName}: ${values}` : null;
            })
            .filter(Boolean);
          fields.push(...customFieldsCompany);
        }
      });
    }
    return fields.join('\n') || 'Нет данных';
  } catch (error) {
    console.error('❌ Ошибка сбора полей:', error.message);
    return 'Ошибка';
  }
}

// Получение данных сущности
function getEntityData(entityType, entityId) {
  try {
    const cacheKey = `${entityType}_${entityId}`;
    if (entityCache[cacheKey]) {
      return entityCache[cacheKey];
    }
    const url = `https://${config.AMO_SUBDOMAIN}/api/v4/${entityType}/${entityId}`;
    const entity = amoRequest(url);
    if (!entity) return null;
    const preparedEntity = {
      id: entity.id,
      name: entity.name,
      responsible: getAmoUser(entity.responsible_user_id).name,
      custom_fields_values: entity.custom_fields_values || []
    };
    entityCache[cacheKey] = preparedEntity;
    return preparedEntity;
  } catch (error) {
    console.error(`❌ Ошибка ${entityType} ID ${entityId}:`, error.message);
    return null;
  }
}

// Обработка события
function processEvent(eventData) {
  try {
    console.log(`📞 Обработка события ID: ${eventData.id}`);
    const initialNoteId = eventData.value_after?.[0]?.note?.id;
    if (initialNoteId && isCallProcessed(initialNoteId)) {
      console.log(`🔁 Звонок ${initialNoteId} уже обработан`);
      return;
    }

    let lead;
    switch(eventData.entity_type.toLowerCase()) {
      case 'lead':
        lead = getAmoLead(eventData.entity_id);
        break;
      case 'contact':
        lead = getLeadByContact(eventData.entity_id);
        break;
      case 'company':
        lead = getLeadByCompany(eventData.entity_id);
        break;
      default:
        console.log(`Неизвестный тип сущности: ${eventData.entity_type}`);
        return;
    }

    if (!lead) return;

    // Получение задач и примечаний
    const tasksText = getAmoTasks(lead.id);
    const notes = getAmoNotes('leads', lead.id);
    const fallbackIdentifier = initialNoteId || `event-${eventData.id}`;
    const callNote = findRelevantCallNote(notes, initialNoteId, eventData.created_at);
    if (!callNote) {
      console.log(`⚠️ Звонок для события ${eventData.id} найден без явного примечания, используется идентификатор ${fallbackIdentifier}`);
    }
    const processedIdentifier = (callNote && callNote.id) || fallbackIdentifier;

    if (isCallProcessed(processedIdentifier)) {
      console.log(`🔁 Запись ${processedIdentifier} уже присутствует в таблице`);
      return;
    }

    // Обработка примечаний
    let notesText = notes.map(note => {
      let text = '';
      switch(note.note_type) {
        case 'common':
          text = `📝 Примечание: ${note.params.text}`;
          if (note._embedded?.attachments?.length) {
            text += `\nattachments:\n${note._embedded.attachments
              .map(att => `• ${att.name} (${att.link})`)
              .join('\n')}`;
          }
          break;
        case 'call_out':
          const duration = note.params.duration ? ` (${note.params.duration}с)` : '';
          const result = config.callStatusMap[note.params.call_status] || 'Без результата';
          text = `📞 Звонок${duration}: ${result}`;
          if (note.params.link) text += ` (${note.params.link})`;
          if (note._embedded?.attachments?.length) {
            text += `\nattachments:\n${note._embedded.attachments
              .map(att => `• ${att.name} (${att.link})`)
              .join('\n')}`;
          }
          break;
        case 'task':
          text = `✅ Задача [ID:${note.id}]: ${note.params.text}`;
          text += `\nСтатус: ${note.params.status === 1 ? 'Выполнена' : 'Не выполнена'}`;
          text += `\nСрок: ${new Date(note.params.task_deadline * 1000)}`;
          break;
        case 'mail':
          text = `📧 Почта: ${note.params.subject}`;
          text += `\nОт: ${note.params.from}`;
          text += `\nСообщение: ${note.params.text}`;
          text += `\nСвязь: ${note.params.link}`;
          break;
        case 'chat':
          text = `💬 Чат (${note.params.service}):`;
          text += `\nСообщение: ${note.params.text}`;
          text += `\nСвязь: ${note.params.link}`;
          break;
        default:
          text = `Неизвестный тип: ${note.note_type}`;
      }
      return text;
    }).join('\n').trim() || 'Нет примечаний';

    // Добавляем задачи в примечания
    notesText += `\n📌 Задачи в сделке:\n${tasksText}`;

    // Формируем данные
    const rowData = [
      new Date(eventData.created_at * 1000), // 1. Время
      callNote?.params?.link || 'Нет записи', // 2. Ссылка
      lead.name, // 3. Сделка
      getAmoPipeline(lead.pipeline_id).name, // 4. Воронка
      getAmoStatus(lead.status_id, lead.pipeline_id).name, // 5. Статус
      lead.id, // 6. ID сделки
      lead.pipeline_id, // 7. ID воронки
      lead.status_id, // 8. ID статуса
      lead.responsible_user_id, // 9. ID ответственного
      getAmoUser(lead.responsible_user_id).name, // 10. Ответственный
      notesText, // 11. Примечания + задачи (колонка K)
      callNote?.params?.duration || 0, // 12. Длительность
      config.callStatusMap[callNote?.params?.call_status] || 'Без результата', // 13. Результат
      eventData.type === 'outgoing_call' ? 'Исходящий' : 'Входящий', // 14. Тип
      processedIdentifier, // 15. ID звонка/события
      getAllEntityFields(lead) // 16. Все поля (с фильтрацией)
    ];
    appendToSheet(rowData);
  } catch (error) {
    console.error(`❌ Ошибка в событии ID ${eventData.id}:`, error.message);
  }
}

// Получение сделки с полными данными
function getAmoLead(leadId) {
  try {
    const url = `https://${config.AMO_SUBDOMAIN}/api/v4/leads/${leadId}?with=contacts,companies`;
    const lead = amoRequest(url);
    if (!lead) return null;
    lead.responsible = getAmoUser(lead.responsible_user_id).name;
    return lead;
  } catch (error) {
    console.error(`❌ Ошибка сделки ID ${leadId}:`, error.message);
    return null;
  }
}

// Получение сделок по контакту
function getLeadByContact(contactId) {
  try {
    const url = `https://${config.AMO_SUBDOMAIN}/api/v4/contacts/${contactId}/links`;
    const linksResponse = amoRequest(url);
    const links = linksResponse?._embedded?.links || [];
    
    // Извлекаем ID сделок из связей
    const leadIds = links
      .filter(link => link.to_entity_type === 'leads')
      .map(link => link.to_entity_id);
    
    return leadIds.length > 0 ? getAmoLead(leadIds[0]) : null;
  } catch (error) {
    console.error(`❌ Ошибка поиска сделок для контакта ${contactId}:`, error.message);
    return null;
  }
}

// Получение сделок по компании
function getLeadByCompany(companyId) {
  try {
    const url = `https://${config.AMO_SUBDOMAIN}/api/v4/companies/${companyId}/links`;
    const linksResponse = amoRequest(url);
    const links = linksResponse?._embedded?.links || [];
    
    // Извлекаем ID сделок из связей
    const leadIds = links
      .filter(link => link.to_entity_type === 'leads')
      .map(link => link.to_entity_id);
    
    return leadIds.length > 0 ? getAmoLead(leadIds[0]) : null;
  } catch (error) {
    console.error(`❌ Ошибка поиска сделок для компании ${companyId}:`, error.message);
    return null;
  }
}

// Получение примечаний
function getAmoNotes(entityType, entityId) {
  try {
    const url = `https://${config.AMO_SUBDOMAIN}/api/v4/${entityType}/${entityId}/notes?with=attachments`;
    const response = amoRequest(url);
    return response?._embedded?.notes || [];
  } catch (error) {
    console.error(`❌ Ошибка примечаний для ${entityType} ID ${entityId}:`, error.message);
    return [];
  }
}

// Получение воронки
function getAmoPipeline(pipelineId) {
  try {
    const pipelines = getPipelines();
    return pipelines.find(p => p.id === pipelineId) || { id: pipelineId, name: "Неизвестная воронка" };
  } catch (error) {
    console.error(`❌ Ошибка воронки ID ${pipelineId}:`, error.message);
    return { id: pipelineId, name: "Неизвестная воронка" };
  }
}

// Получение статуса
function getAmoStatus(statusId, pipelineId) {
  try {
    const pipelines = getPipelines();
    const pipeline = pipelines.find(p => p.id === pipelineId);
    const statuses = pipeline?._embedded?.statuses || [];
    return statuses.find(s => s.id === statusId) || { id: statusId, name: "Неизвестный статус" };
  } catch (error) {
    console.error(`❌ Ошибка статуса ID ${statusId}:`, error.message);
    return { id: statusId, name: "Неизвестный статус" };
  }
}

// Получение пользователя
function getAmoUser(userId) {
  try {
    if (userCache[userId]) {
      return userCache[userId];
    }
    const url = `https://${config.AMO_SUBDOMAIN}/api/v4/users/${userId}`;
    const user = amoRequest(url);
    const preparedUser = user ? {
      id: user.id,
      name: user.name || "Неизвестный пользователь"
    } : { id: userId, name: "Неизвестный пользователь" };
    userCache[userId] = preparedUser;
    return preparedUser;
  } catch (error) {
    console.error(`❌ Ошибка пользователя ID ${userId}:`, error.message);
    return { id: userId, name: "Неизвестный пользователь" };
  }
}

// Запись данных в таблицу
function appendToSheet(rowData) {
  try {
    const sheet = SpreadsheetApp.openById(config.SPREADSHEET_ID)
      .getSheetByName(config.SHEET_NAME);
    if (!sheet) throw new Error('Лист не найден');
    sheet.appendRow(rowData);
    console.log('✅ Данные записаны:', rowData);
    markCallAsProcessed(rowData[14]);
  } catch (error) {
    console.error('❌ Ошибка записи:', error.message);
  }
}

function getPipelines() {
  if (pipelinesCache) {
    return pipelinesCache;
  }
  const url = `https://${config.AMO_SUBDOMAIN}/api/v4/leads/pipelines`;
  const response = amoRequest(url);
  pipelinesCache = response?._embedded?.pipelines || [];
  return pipelinesCache;
}

function findRelevantCallNote(notes, expectedId, eventCreatedAt) {
  if (!Array.isArray(notes) || notes.length === 0) {
    return null;
  }

  if (expectedId) {
    const exactMatch = notes.find(function(note) {
      return note.id === expectedId;
    });
    if (exactMatch) {
      return exactMatch;
    }
  }

  const callNotes = notes.filter(function(note) {
    return note.note_type === 'call_out' || note.note_type === 'call_in';
  });

  if (callNotes.length === 0) {
    return null;
  }

  const sorted = callNotes.slice().sort(function(a, b) {
    const aTime = a.created_at || 0;
    const bTime = b.created_at || 0;
    if (eventCreatedAt) {
      return Math.abs(aTime - eventCreatedAt) - Math.abs(bTime - eventCreatedAt);
    }
    return bTime - aTime;
  });

  return sorted[0];
}

// Основная функция синхронизации
function syncEventsToday() {
  try {
    resetProcessedCallIdsCache();
    const timeFrom = Math.floor((new Date().getTime() - 3 * 60 * 1000) / 1000); // 3 минут
    let url = `https://${config.AMO_SUBDOMAIN}/api/v4/events?filter[created_at][from]=${timeFrom}&limit=250`;
    let allEvents = [];
    while (true) {
      const response = amoRequest(url);
      if (!response?._embedded?.events) break;
      allEvents = allEvents.concat(response._embedded.events);
      url = response._links?.next?.href || '';
      if (!url) break;
    }
    const calls = allEvents.filter(e =>
      ['outgoing_call', 'incoming_call'].includes(e.type)
    );
    calls.forEach(event => processEvent(event));
  } catch (error) {
    console.error('❌ Ошибка синхронизации:', error.message);
  }
}

function getProcessedCallIds() {
  if (processedCallIdsCache) {
    return processedCallIdsCache;
  }

  processedCallIdsCache = new Set();

  try {
    const sheet = SpreadsheetApp.openById(config.SPREADSHEET_ID)
      .getSheetByName(config.SHEET_NAME);
    if (!sheet) {
      return processedCallIdsCache;
    }
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return processedCallIdsCache;
    }
    const values = sheet.getRange(2, 15, lastRow - 1, 1).getValues();
    values.forEach(function(row) {
      const value = row[0];
      if (value !== '' && value !== null && value !== undefined) {
        processedCallIdsCache.add(String(value));
      }
    });
  } catch (error) {
    console.error('❌ Ошибка чтения обработанных звонков:', error.message);
  }

  return processedCallIdsCache;
}

function markCallAsProcessed(callNoteId) {
  if (!callNoteId) {
    return;
  }
  try {
    getProcessedCallIds().add(String(callNoteId));
  } catch (error) {
    console.error('❌ Ошибка сохранения ID звонка:', error.message);
  }
}

function resetProcessedCallIdsCache() {
  processedCallIdsCache = null;
}
