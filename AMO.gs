// Конфигурация
function getConfigFromSheet() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = spreadsheet.getSheetByName('БД');
    if (!configSheet) {
      throw new Error('Лист "БД" не найден');
    }

    var lastRow = Math.max(configSheet.getLastRow(), 2);
    var callStatusRange = lastRow > 1 ? configSheet.getRange(2, 3, lastRow - 1, 2) : null;
    var taskTypeRange = lastRow > 1 ? configSheet.getRange(2, 5, lastRow - 1, 2) : null;

    return {
      AMO_TOKEN: String(configSheet.getRange('B2').getValue()).trim(),
      AMO_SUBDOMAIN: String(configSheet.getRange('B3').getValue()).trim(),
      SHEET_NAME: String(configSheet.getRange('B4').getValue()).trim(),
      SPREADSHEET_ID: spreadsheet.getId(),
      callStatusMap: buildDictionaryFromRange(callStatusRange),
      taskTypesMap: buildDictionaryFromRange(taskTypeRange),
      excludedFields: readExcludedFields(configSheet)
    };
  } catch (error) {
    console.error('❌ Ошибка конфигурации:', error.message);
    throw error;
  }
}

function buildDictionaryFromRange(range) {
  if (!range) {
    return {};
  }
  var values = range.getValues();
  var result = {};
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (row[0] && row[1]) {
      result[String(row[0])] = row[1];
    }
  }
  return result;
}

var config = getConfigFromSheet();
var pipelinesCache = { pipelines: null, fetchedAt: 0 };
var userCache = {};

// Чтение исключенных полей
function readExcludedFields(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  var values = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
  var result = [];
  for (var i = 0; i < values.length; i++) {
    var value = values[i][0];
    if (typeof value === 'string') {
      var trimmed = value.trim();
      if (trimmed !== '') {
        result.push(trimmed);
      }
    }
  }
  return result;
}

function getCachedPipelines() {
  var now = Date.now();
  if (!pipelinesCache.pipelines || now - pipelinesCache.fetchedAt > 5 * 60 * 1000) {
    var url = 'https://' + config.AMO_SUBDOMAIN + '/api/v4/leads/pipelines';
    var response = amoRequest(url);
    var pipelines = response && response._embedded ? response._embedded.pipelines : [];
    pipelinesCache.pipelines = pipelines || [];
    pipelinesCache.fetchedAt = now;
  }
  return pipelinesCache.pipelines;
}

// Форматирование времени
function formatTimestamp(timestamp) {
  if (!timestamp) {
    return 'Нет данных';
  }
  var date = new Date(timestamp * 1000);
  return Utilities.formatDate(date, 'GMT+3', 'dd.MM.yyyy HH:mm');
}

// Проверка обработки звонка
function isCallProcessed(callNoteId) {
  if (!callNoteId) {
    return false;
  }

  try {
    var sheet = SpreadsheetApp.openById(config.SPREADSHEET_ID).getSheetByName(config.SHEET_NAME);
    if (!sheet) {
      return false;
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return false;
    }

    var range = sheet.getRange(2, 15, lastRow - 1, 1).getValues();
    for (var i = 0; i < range.length; i++) {
      if (String(range[i][0]) === String(callNoteId)) {
        return true;
      }
    }
    return false;
  } catch (error) {
    console.error('❌ Ошибка проверки:', error.message);
    return false;
  }
}

// Универсальный запрос к API
function amoRequest(url, method, body) {
  if (!method) {
    method = 'GET';
  }
  try {
    var headers = { 'Authorization': 'Bearer ' + config.AMO_TOKEN };
    var options = {
      method: method,
      headers: headers,
      muteHttpExceptions: true
    };
    if (body) {
      headers['Content-Type'] = 'application/json';
      options.payload = JSON.stringify(body);
    }
    var response = UrlFetchApp.fetch(url, options);
    var statusCode = response.getResponseCode();
    if (statusCode === 204) {
      console.log('✅ Пустой ответ (204) для ' + url);
      return null;
    }
    if (statusCode !== 200) {
      console.error('❌ HTTP ' + statusCode + ': ' + response.getContentText());
      return null;
    }
    return JSON.parse(response.getContentText());
  } catch (error) {
    console.error('❌ Ошибка запроса к ' + url + ':', error.message);
    return null;
  }
}

// Получение задач для сделки
function getAmoTasks(leadId) {
  try {
    var url = 'https://' + config.AMO_SUBDOMAIN + '/api/v4/tasks?filter[entity_id]=' + leadId + '&filter[entity_type]=leads';
    var tasksResponse = amoRequest(url);
    if (!tasksResponse) {
      return 'Нет задач';
    }
    var tasks = tasksResponse._embedded && tasksResponse._embedded.tasks ? tasksResponse._embedded.tasks : [];
    if (!tasks || tasks.length === 0) {
      return 'Нет задач';
    }

    var lines = [];
    for (var i = 0; i < tasks.length; i++) {
      var task = tasks[i];
      var taskTypeKey = task.task_type_id;
      var taskType = config.taskTypesMap[String(taskTypeKey)] || ('Неизвестный тип (' + taskTypeKey + ')');
      var dueDate = formatTimestamp(task.complete_till);
      var status;
      if (task.is_completed === true) {
        status = 'Выполнена';
      } else if (task.is_completed === false) {
        status = 'Не выполнена';
      } else {
        status = 'Статус не определен';
      }
      var text = '• ' + taskType + ': ' + (task.text || 'Нет текста') + '\n  Срок: ' + dueDate + '\n  Статус: ' + status;
      lines.push(text);
    }
    return lines.join('\n');
  } catch (error) {
    console.error('❌ Ошибка задач для сделки ' + leadId + ':', error.message);
    return 'Ошибка при получении задач';
  }
}

function shouldSkipField(fieldName) {
  if (!fieldName) {
    return true;
  }
  for (var i = 0; i < config.excludedFields.length; i++) {
    if (config.excludedFields[i] === fieldName) {
      return true;
    }
  }
  return false;
}

function collectCustomFieldValues(fields) {
  if (!fields) {
    return [];
  }
  var result = [];
  for (var i = 0; i < fields.length; i++) {
    var field = fields[i];
    var fieldName = field.field_name;
    if (shouldSkipField(fieldName)) {
      continue;
    }
    var values = field.values || [];
    var collected = [];
    for (var j = 0; j < values.length; j++) {
      if (values[j] && values[j].value) {
        collected.push(values[j].value);
      }
    }
    if (collected.length > 0) {
      result.push(fieldName + ': ' + collected.join(', '));
    }
  }
  return result;
}

// Сбор всех полей из сделки (с фильтрацией)
function getAllEntityFields(lead) {
  try {
    var fields = [];
    fields.push('--- Сделка ---');
    fields.push('ID: ' + lead.id);
    fields.push('Ответственный: ' + getAmoUser(lead.responsible_user_id).name);

    var customFieldsDeal = collectCustomFieldValues(lead.custom_fields_values);
    fields = fields.concat(customFieldsDeal);

    var embedded = lead._embedded || {};
    var contacts = embedded.contacts || [];
    for (var i = 0; i < contacts.length; i++) {
      var contact = getEntityData('contacts', contacts[i].id);
      if (!contact) {
        continue;
      }
      fields.push('--- Контакт ---');
      fields.push('ID: ' + contact.id);
      fields.push('Ответственный: ' + contact.responsible);
      fields = fields.concat(collectCustomFieldValues(contact.custom_fields_values));
    }

    var companies = embedded.companies || [];
    for (var j = 0; j < companies.length; j++) {
      var company = getEntityData('companies', companies[j].id);
      if (!company) {
        continue;
      }
      fields.push('--- Компания ---');
      fields.push('ID: ' + company.id);
      fields.push('Ответственный: ' + company.responsible);
      fields = fields.concat(collectCustomFieldValues(company.custom_fields_values));
    }

    return fields.join('\n');
  } catch (error) {
    console.error('❌ Ошибка сбора полей:', error.message);
    return 'Ошибка';
  }
}

// Получение данных сущности
function getEntityData(entityType, entityId) {
  try {
    var url = 'https://' + config.AMO_SUBDOMAIN + '/api/v4/' + entityType + '/' + entityId;
    var entity = amoRequest(url);
    if (!entity) {
      return null;
    }
    return {
      id: entity.id,
      name: entity.name,
      responsible: getAmoUser(entity.responsible_user_id).name,
      custom_fields_values: entity.custom_fields_values || []
    };
  } catch (error) {
    console.error('❌ Ошибка ' + entityType + ' ID ' + entityId + ':', error.message);
    return null;
  }
}

// Обработка события
function processEvent(eventData) {
  try {
    console.log('📞 Обработка события ID: ' + eventData.id);
    var valueAfter = eventData.value_after || [];
    var noteWrapper = valueAfter.length > 0 ? valueAfter[0] : null;
    var note = noteWrapper && noteWrapper.note ? noteWrapper.note : null;
    var callNoteId = note ? note.id : null;
    if (!callNoteId || isCallProcessed(callNoteId)) {
      return;
    }

    var lead = null;
    var entityType = (eventData.entity_type || '').toLowerCase();
    if (entityType === 'lead') {
      lead = getAmoLead(eventData.entity_id);
    } else if (entityType === 'contact') {
      lead = getLeadByContact(eventData.entity_id);
    } else if (entityType === 'company') {
      lead = getLeadByCompany(eventData.entity_id);
    } else {
      console.log('Неизвестный тип сущности: ' + eventData.entity_type);
      return;
    }

    if (!lead) {
      return;
    }

    var tasksText = getAmoTasks(lead.id);
    var notes = getAmoNotes('leads', lead.id);
    var callNote = null;
    for (var i = 0; i < notes.length; i++) {
      if (notes[i].id === callNoteId) {
        callNote = notes[i];
        break;
      }
    }

    var notesText = buildNotesText(notes, tasksText);

    var callParams = callNote && callNote.params ? callNote.params : {};
    var callDuration = callParams.duration || 0;
    var callResultKey = callParams.call_status;
    var callResult = config.callStatusMap[String(callResultKey)] || 'Без результата';
    var callLink = callParams.link || 'Нет записи';

    var rowData = [
      new Date(eventData.created_at * 1000),
      callLink,
      lead.name,
      getAmoPipeline(lead.pipeline_id).name,
      getAmoStatus(lead.status_id, lead.pipeline_id).name,
      lead.id,
      lead.pipeline_id,
      lead.status_id,
      lead.responsible_user_id,
      getAmoUser(lead.responsible_user_id).name,
      notesText,
      callDuration,
      callResult,
      eventData.type === 'outgoing_call' ? 'Исходящий' : 'Входящий',
      callNoteId,
      getAllEntityFields(lead)
    ];
    appendToSheet(rowData);
  } catch (error) {
    console.error('❌ Ошибка в событии ID ' + eventData.id + ':', error.message);
  }
}

function buildNotesText(notes, tasksText) {
  if (!notes || notes.length === 0) {
    return 'Нет примечаний\n📌 Задачи в сделке:\n' + tasksText;
  }

  var parts = [];
  for (var i = 0; i < notes.length; i++) {
    var note = notes[i];
    var noteText = '';
    var params = note.params || {};
    var attachments = note._embedded && note._embedded.attachments ? note._embedded.attachments : [];
    if (note.note_type === 'common') {
      noteText = '📝 Примечание: ' + (params.text || '');
      noteText += formatAttachments(attachments);
    } else if (note.note_type === 'call_out') {
      var duration = params.duration ? ' (' + params.duration + 'с)' : '';
      var callStatusKey = params.call_status;
      var callStatus = config.callStatusMap[String(callStatusKey)] || 'Без результата';
      noteText = '📞 Звонок' + duration + ': ' + callStatus;
      if (params.link) {
        noteText += ' (' + params.link + ')';
      }
      noteText += formatAttachments(attachments);
    } else if (note.note_type === 'task') {
      noteText = '✅ Задача [ID:' + note.id + ']: ' + (params.text || '');
      noteText += '\nСтатус: ' + (params.status === 1 ? 'Выполнена' : 'Не выполнена');
      noteText += '\nСрок: ' + (params.task_deadline ? new Date(params.task_deadline * 1000) : 'Не указан');
    } else if (note.note_type === 'mail') {
      noteText = '📧 Почта: ' + (params.subject || '');
      noteText += '\nОт: ' + (params.from || '');
      noteText += '\nСообщение: ' + (params.text || '');
      noteText += '\nСвязь: ' + (params.link || '');
    } else if (note.note_type === 'chat') {
      noteText = '💬 Чат (' + (params.service || '') + '):';
      noteText += '\nСообщение: ' + (params.text || '');
      noteText += '\nСвязь: ' + (params.link || '');
    } else {
      noteText = 'Неизвестный тип: ' + note.note_type;
    }
    parts.push(noteText);
  }

  parts.push('📌 Задачи в сделке:\n' + tasksText);
  return parts.join('\n');
}

function formatAttachments(attachments) {
  if (!attachments || attachments.length === 0) {
    return '';
  }
  var lines = ['\nattachments:'];
  for (var i = 0; i < attachments.length; i++) {
    var att = attachments[i];
    lines.push('• ' + (att.name || 'файл') + ' (' + (att.link || '') + ')');
  }
  return '\n' + lines.join('\n');
}

// Получение сделки с полными данными
function getAmoLead(leadId) {
  try {
    var url = 'https://' + config.AMO_SUBDOMAIN + '/api/v4/leads/' + leadId + '?with=contacts,companies';
    var lead = amoRequest(url);
    if (!lead) {
      return null;
    }
    lead.responsible = getAmoUser(lead.responsible_user_id).name;
    return lead;
  } catch (error) {
    console.error('❌ Ошибка сделки ID ' + leadId + ':', error.message);
    return null;
  }
}

// Получение сделок по контакту
function getLeadByContact(contactId) {
  try {
    var url = 'https://' + config.AMO_SUBDOMAIN + '/api/v4/contacts/' + contactId + '/links';
    var linksResponse = amoRequest(url);
    var links = linksResponse && linksResponse._embedded ? linksResponse._embedded.links : [];
    var leadId = null;
    for (var i = 0; i < links.length; i++) {
      if (links[i].to_entity_type === 'leads') {
        leadId = links[i].to_entity_id;
        break;
      }
    }
    return leadId ? getAmoLead(leadId) : null;
  } catch (error) {
    console.error('❌ Ошибка поиска сделок для контакта ' + contactId + ':', error.message);
    return null;
  }
}

// Получение сделок по компании
function getLeadByCompany(companyId) {
  try {
    var url = 'https://' + config.AMO_SUBDOMAIN + '/api/v4/companies/' + companyId + '/links';
    var linksResponse = amoRequest(url);
    var links = linksResponse && linksResponse._embedded ? linksResponse._embedded.links : [];
    var leadId = null;
    for (var i = 0; i < links.length; i++) {
      if (links[i].to_entity_type === 'leads') {
        leadId = links[i].to_entity_id;
        break;
      }
    }
    return leadId ? getAmoLead(leadId) : null;
  } catch (error) {
    console.error('❌ Ошибка поиска сделок для компании ' + companyId + ':', error.message);
    return null;
  }
}

// Получение примечаний
function getAmoNotes(entityType, entityId) {
  try {
    var url = 'https://' + config.AMO_SUBDOMAIN + '/api/v4/' + entityType + '/' + entityId + '/notes?with=attachments';
    var response = amoRequest(url);
    var notes = response && response._embedded ? response._embedded.notes : [];
    return notes || [];
  } catch (error) {
    console.error('❌ Ошибка примечаний для ' + entityType + ' ID ' + entityId + ':', error.message);
    return [];
  }
}

// Получение воронки
function getAmoPipeline(pipelineId) {
  try {
    var pipelines = getCachedPipelines();
    for (var i = 0; i < pipelines.length; i++) {
      if (pipelines[i].id === pipelineId) {
        return pipelines[i];
      }
    }
    return { id: pipelineId, name: 'Неизвестная воронка' };
  } catch (error) {
    console.error('❌ Ошибка воронки ID ' + pipelineId + ':', error.message);
    return { id: pipelineId, name: 'Неизвестная воронка' };
  }
}

// Получение статуса
function getAmoStatus(statusId, pipelineId) {
  try {
    var pipelines = getCachedPipelines();
    for (var i = 0; i < pipelines.length; i++) {
      if (pipelines[i].id === pipelineId) {
        var statuses = pipelines[i]._embedded ? pipelines[i]._embedded.statuses : [];
        for (var j = 0; j < statuses.length; j++) {
          if (statuses[j].id === statusId) {
            return statuses[j];
          }
        }
      }
    }
    return { id: statusId, name: 'Неизвестный статус' };
  } catch (error) {
    console.error('❌ Ошибка статуса ID ' + statusId + ':', error.message);
    return { id: statusId, name: 'Неизвестный статус' };
  }
}

// Получение пользователя
function getAmoUser(userId) {
  var cacheKey = String(userId);
  if (userCache[cacheKey]) {
    return userCache[cacheKey];
  }
  try {
    var url = 'https://' + config.AMO_SUBDOMAIN + '/api/v4/users/' + userId;
    var user = amoRequest(url);
    var normalized = user ? {
      id: user.id,
      name: user.name || 'Неизвестный пользователь'
    } : {
      id: userId,
      name: 'Неизвестный пользователь'
    };
    userCache[cacheKey] = normalized;
    return normalized;
  } catch (error) {
    console.error('❌ Ошибка пользователя ID ' + userId + ':', error.message);
    return { id: userId, name: 'Неизвестный пользователь' };
  }
}

// Запись данных в таблицу
function appendToSheet(rowData) {
  try {
    var sheet = SpreadsheetApp.openById(config.SPREADSHEET_ID).getSheetByName(config.SHEET_NAME);
    if (!sheet) {
      throw new Error('Лист не найден');
    }
    sheet.appendRow(rowData);
    console.log('✅ Данные записаны: ' + JSON.stringify(rowData));
  } catch (error) {
    console.error('❌ Ошибка записи:', error.message);
  }
}

// Основная функция синхронизации
function syncEventsToday() {
  try {
    var timeFrom = Math.floor((new Date().getTime() - 3 * 60 * 1000) / 1000);
    var url = 'https://' + config.AMO_SUBDOMAIN + '/api/v4/events?filter[created_at][from]=' + timeFrom + '&limit=250';
    var allEvents = [];

    while (url) {
      var response = amoRequest(url);
      if (!response) {
        break;
      }
      var events = response._embedded && response._embedded.events ? response._embedded.events : [];
      if (events && events.length) {
        for (var i = 0; i < events.length; i++) {
          allEvents.push(events[i]);
        }
      }
      url = response._links && response._links.next ? response._links.next.href : '';
      if (!url) {
        break;
      }
    }

    for (var j = 0; j < allEvents.length; j++) {
      var event = allEvents[j];
      if (event && (event.type === 'outgoing_call' || event.type === 'incoming_call')) {
        processEvent(event);
      }
    }
  } catch (error) {
    console.error('❌ Ошибка синхронизации:', error.message);
  }
}
