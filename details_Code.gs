/**
 * Code.gs — серверная логика для конструктора «Детали»
 *
 * Полностью совместим с лагерями (project='camp') и интенсивами (project='intensive').
 * Третий проект — project='details' (студия «Детали»).
 *
 * НАСТРОЙКА:
 *   1. Откройте Google Sheets «Детали — Сетка планирования».
 *      Скопируйте ID из URL и подставьте в SHEET_ID ниже.
 *   2. (для ИИ) Project Settings → Script Properties → ANTHROPIC_API_KEY
 *   3. Deploy → New deployment → Web app → Execute as: Me, Who: Anyone
 *      Скопируйте URL вида .../exec — он понадобится для index.html (APPS_SCRIPT_URL).
 *
 * Что лежит в каких листах для Деталей:
 *   – Форматы_детали            — справочник форматов событий (МК Глина, Киновечер, ДР…)
 *   – Педагоги_детали            — справочник педагогов со ставками
 *   – Шаблоны_детали            — переиспользуемые «болванки» событий с описанием
 *   – Опубликованные_детали     — снимки опубликованных календарей (для view/публикации)
 *   – Черновики_детали          — рабочие версии менеджеров
 *   – Сессии_детали             — heartbeat «кто сейчас редактирует»
 *   – События                    — ГЛАВНЫЙ лист с событиями (туда улетает публикация
 *                                  и оттуда — в Альфу через alfa_integration.gs)
 */

// ============================================================
// КОНФИГ
// ============================================================
const SHEET_ID = '1unWnHDYUwnCd9bkdvacEGOpEFFcugMBLSoPuZCMrJIE';

// ----- Листы Деталей -----
const SH_DET_FORMATS    = 'Форматы_детали';
const SH_DET_TEACHERS   = 'Педагоги_детали';
const SH_DET_TEMPLATES  = 'Шаблоны_детали';
const SH_DET_PUBLISHED  = 'Опубликованные_детали';
const SH_DET_DRAFTS     = 'Черновики_детали';
const SH_DET_SESSIONS   = 'Сессии_детали';
const SH_DET_EVENTS     = 'События';   // боевой лист — мост в Альфу

// ----- Старые листы (camp/intensive) — оставлены для совместимости -----
const SHEET_INT_ACTIVITIES  = 'Активности_интенсивы';
const SHEET_INT_DETAILS     = 'Реквизиты_интенсивы';
const SHEET_INT_PUBLISHED   = 'Опубликованные_интенсивы';
const SHEET_INT_DRAFTS      = 'Черновики_интенсивы';
const SHEET_INT_SESSIONS    = 'Сессии_интенсивы';

const SHEET_CAMP_ACTIVITIES = 'Активности';
const SHEET_CAMP_DETAILS    = 'Реквизиты';
const SHEET_CAMP_PUBLISHED  = 'Опубликованные версии';
const SHEET_CAMP_DRAFTS     = 'Черновики';
const SHEET_CAMP_SESSIONS   = 'Сессии';

// ----- Anthropic -----
const ANTHROPIC_API_URL = 'https://api.anthropic.com/v1/messages';
const ANTHROPIC_MODEL = 'claude-sonnet-4-20250514';
const ANTHROPIC_MAX_TOKENS = 2000;

// Сессии: считаем активной если heartbeat был не более 5 минут назад
const SESSION_TIMEOUT_MS = 5 * 60 * 1000;

// Диапазон строк в листе «События» — где могут лежать данные
const EVENTS_DATA_START_ROW = 4;

// ============================================================
// УТИЛИТЫ
// ============================================================
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    initSheetHeaders_(sh, name);
  }
  return sh;
}

function initSheetHeaders_(sh, name) {
  const headers = {
    [SH_DET_FORMATS]:    ['Формат', 'Категория', 'Описание-шаблон', 'Цена по умолчанию', 'Лимит по умолчанию', 'Длительность по умолчанию (ч)', 'Активен'],
    [SH_DET_TEACHERS]:   ['ФИО', 'Ставка, руб', 'ID в Альфе', 'Контакт', 'Специализация', 'Активен'],
    [SH_DET_TEMPLATES]:  ['Название', 'Формат', 'Педагог', 'Описание', 'Цена', 'Лимит', 'Длительность (ч)', 'Теги', 'Использовалось раз'],
    [SH_DET_PUBLISHED]:  ['id', 'name', 'monthKey', 'updated', 'stateJson'],
    [SH_DET_DRAFTS]:     ['id', 'name', 'updated', 'stateJson'],
    [SH_DET_SESSIONS]:   ['sessionId', 'publishId', 'editorName', 'userAgent', 'updated'],

    [SHEET_INT_ACTIVITIES]: ['Статус','Шаблон занятия','Примеры тем','Педагог','Направление','Локация / кабинет','Возраст','Длительность (мин)','Формат','Описание'],
    [SHEET_INT_DETAILS]:    ['Ключ','Значение'],
    [SHEET_INT_PUBLISHED]:  ['id','name','title','updated','stateJson'],
    [SHEET_INT_DRAFTS]:     ['id','name','updated','stateJson'],
    [SHEET_INT_SESSIONS]:   ['sessionId','publishId','editorName','userAgent','updated']
  };
  if (headers[name]) {
    sh.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
    sh.setFrozenRows(1);
  }
}

function getProjectSheets_(project) {
  if (project === 'details') {
    return {
      formats:    SH_DET_FORMATS,
      teachers:   SH_DET_TEACHERS,
      templates:  SH_DET_TEMPLATES,
      published:  SH_DET_PUBLISHED,
      drafts:     SH_DET_DRAFTS,
      sessions:   SH_DET_SESSIONS,
      events:     SH_DET_EVENTS
    };
  }
  if (project === 'intensive') {
    return {
      activities: SHEET_INT_ACTIVITIES,
      details:    SHEET_INT_DETAILS,
      published:  SHEET_INT_PUBLISHED,
      drafts:     SHEET_INT_DRAFTS,
      sessions:   SHEET_INT_SESSIONS
    };
  }
  return {
    activities: SHEET_CAMP_ACTIVITIES,
    details:    SHEET_CAMP_DETAILS,
    published:  SHEET_CAMP_PUBLISHED,
    drafts:     SHEET_CAMP_DRAFTS,
    sessions:   SHEET_CAMP_SESSIONS
  };
}

function readSheetAsObjects_(sh) {
  const range = sh.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.every(c => c === '' || c == null)) continue;
    const obj = {};
    headers.forEach((h, j) => { obj[h] = row[j]; });
    obj._rowIdx = i + 1;
    out.push(obj);
  }
  return out;
}

function findRowById_(sh, id) {
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function nowIso_() { return new Date().toISOString(); }

// ============================================================
// ENTRY POINTS
// ============================================================
function doGet(e) {
  try {
    const params = e.parameter || {};
    const action = params.action || 'all';
    const project = params.project || 'camp';

    if (project === 'details') return doGetDetails_(action, params);

    switch(action) {
      case 'all':           return jsonResponse(getAll_(project));
      case 'library':       return jsonResponse(getLibrary_(project));
      case 'details':       return jsonResponse(getDetails_(project));
      case 'listPublished': return jsonResponse(listPublished_(project));
      case 'getPublished':  return jsonResponse(getPublished_(project, params.id));
      case 'listDrafts':    return jsonResponse(listDrafts_(project));
      case 'loadDraft':     return jsonResponse(loadDraft_(project, params.id));
      case 'listSessions':  return jsonResponse(listSessions_(project, params.publishId));
      default: return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch(err) {
    return jsonResponse({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const project = body.project || 'camp';

    if (project === 'details') return doPostDetails_(action, body);

    switch(action) {
      case 'claude':       return jsonResponse(callClaude_(body.prompt, body.system));
      case 'publish':      return jsonResponse(publishVersion_(project, body));
      case 'unpublish':    return jsonResponse(unpublishVersion_(project, body.id));
      case 'saveDraft':    return jsonResponse(saveDraft_(project, body));
      case 'deleteDraft':  return jsonResponse(deleteDraft_(project, body.id));
      case 'heartbeat':    return jsonResponse(heartbeat_(project, body));
      case 'endSession':   return jsonResponse(endSession_(project, body.sessionId));
      default: return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch(err) {
    return jsonResponse({ ok: false, error: String(err.message || err) });
  }
}

// ============================================================
// DETAILS — ROUTING
// ============================================================
function doGetDetails_(action, params) {
  switch(action) {
    case 'all':           return jsonResponse(detailsGetAll_());
    case 'library':       return jsonResponse(detailsGetLibrary_());
    case 'listPublished': return jsonResponse(detailsListPublished_());
    case 'getPublished':  return jsonResponse(detailsGetPublished_(params.id));
    case 'listDrafts':    return jsonResponse(detailsListDrafts_());
    case 'loadDraft':     return jsonResponse(detailsLoadDraft_(params.id));
    case 'listSessions':  return jsonResponse(listSessions_('details', params.publishId));
    case 'pingEvents':    return jsonResponse(detailsPingEvents_());
    default: return jsonResponse({ ok: false, error: 'Unknown details action: ' + action });
  }
}

function doPostDetails_(action, body) {
  switch(action) {
    case 'claude':              return jsonResponse(callClaude_(body.prompt, body.system));
    case 'publish':             return jsonResponse(detailsPublish_(body));
    case 'publishAndSendToAlfa':return jsonResponse(detailsPublishAndSendToAlfa_(body));
    case 'unpublish':           return jsonResponse(unpublishVersion_('details', body.id));
    case 'saveDraft':           return jsonResponse(detailsSaveDraft_(body));
    case 'deleteDraft':         return jsonResponse(deleteDraft_('details', body.id));
    case 'heartbeat':           return jsonResponse(heartbeat_('details', body));
    case 'endSession':          return jsonResponse(endSession_('details', body.sessionId));
    case 'addFormat':           return jsonResponse(detailsAddFormat_(body));
    case 'addTeacher':          return jsonResponse(detailsAddTeacher_(body));
    case 'saveTemplate':        return jsonResponse(detailsSaveTemplate_(body));
    case 'syncToEvents':        return jsonResponse(detailsSyncToEvents_(body));
    case 'testAlfaConnection':  return jsonResponse(detailsTestAlfa_());
    default: return jsonResponse({ ok: false, error: 'Unknown details action: ' + action });
  }
}

// ============================================================
// DETAILS — DATA
// ============================================================

function detailsGetAll_() {
  return {
    ok: true,
    formats:   detailsGetFormats_(),
    teachers:  detailsGetTeachers_(),
    templates: detailsGetTemplates_()
  };
}

function detailsGetLibrary_() {
  return { ok: true, templates: detailsGetTemplates_() };
}

function detailsGetFormats_() {
  const sh = getSheet_(SH_DET_FORMATS);
  const rows = readSheetAsObjects_(sh);
  return rows
    .filter(r => String(r['Активен'] || '').toLowerCase() !== 'нет')
    .map(r => ({
      name:        String(r['Формат'] || '').trim(),
      category:    String(r['Категория'] || '').trim(),
      template:    String(r['Описание-шаблон'] || '').trim(),
      defaultPrice:    Number(r['Цена по умолчанию']) || null,
      defaultLimit:    Number(r['Лимит по умолчанию']) || null,
      defaultDuration: Number(r['Длительность по умолчанию (ч)']) || null
    }))
    .filter(f => f.name);
}

function detailsGetTeachers_() {
  const sh = getSheet_(SH_DET_TEACHERS);
  const rows = readSheetAsObjects_(sh);
  return rows
    .filter(r => String(r['Активен'] || '').toLowerCase() !== 'нет')
    .map(r => ({
      name:    String(r['ФИО'] || '').trim(),
      rate:    Number(r['Ставка, руб']) || null,
      alfaId:  Number(r['ID в Альфе']) || null,
      contact: String(r['Контакт'] || '').trim(),
      spec:    String(r['Специализация'] || '').trim()
    }))
    .filter(t => t.name);
}

function detailsGetTemplates_() {
  const sh = getSheet_(SH_DET_TEMPLATES);
  const rows = readSheetAsObjects_(sh);
  return rows.map(r => ({
    name:        String(r['Название'] || '').trim(),
    format:      String(r['Формат'] || '').trim(),
    teacher:     String(r['Педагог'] || '').trim(),
    description: String(r['Описание'] || '').trim(),
    price:       Number(r['Цена']) || null,
    limit:       Number(r['Лимит']) || null,
    duration:    Number(r['Длительность (ч)']) || null,
    tags:        String(r['Теги'] || '').trim(),
    usedCount:   Number(r['Использовалось раз']) || 0,
    rowIdx:      r._rowIdx
  })).filter(t => t.name);
}

function detailsAddFormat_(body) {
  const f = body.format || {};
  if (!f.name) return { ok: false, error: 'нет name' };
  const sh = getSheet_(SH_DET_FORMATS);
  sh.appendRow([
    f.name, f.category || '', f.template || '',
    f.defaultPrice || '', f.defaultLimit || '', f.defaultDuration || '',
    'Да'
  ]);
  return { ok: true };
}

function detailsAddTeacher_(body) {
  const t = body.teacher || {};
  if (!t.name) return { ok: false, error: 'нет name' };
  const sh = getSheet_(SH_DET_TEACHERS);
  sh.appendRow([
    t.name, t.rate || '', t.alfaId || '', t.contact || '', t.spec || '',
    'Да'
  ]);
  return { ok: true };
}

function detailsSaveTemplate_(body) {
  const t = body.template || {};
  if (!t.name) return { ok: false, error: 'нет name' };
  const sh = getSheet_(SH_DET_TEMPLATES);
  // Если такое название уже есть — увеличим счётчик использования
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === t.name) {
      const used = Number(data[i][8]) || 0;
      sh.getRange(i + 1, 9).setValue(used + 1);
      return { ok: true, updated: true };
    }
  }
  sh.appendRow([
    t.name, t.format || '', t.teacher || '', t.description || '',
    t.price || '', t.limit || '', t.duration || '', t.tags || '', 1
  ]);
  return { ok: true, created: true };
}

// ============================================================
// DETAILS — PUBLISH / DRAFTS
// ============================================================

function detailsPublish_(body) {
  const id = String(body.id || '').trim();
  const name = String(body.name || '').trim();
  const stateJson = String(body.stateJson || '');
  if (!id) return { ok: false, error: 'нет id' };
  if (!stateJson) return { ok: false, error: 'нет stateJson' };

  const sh = getSheet_(SH_DET_PUBLISHED);

  let monthKey = '';
  let stateObj = null;
  try {
    stateObj = JSON.parse(stateJson);
    monthKey = String(stateObj.monthKey || '');
  } catch(e) {}

  const existingRow = findRowById_(sh, id);
  const rowData = [id, name, monthKey, nowIso_(), stateJson];

  if (existingRow > 0) {
    sh.deleteRow(existingRow);
    sh.appendRow(rowData);
  } else {
    sh.appendRow(rowData);
  }

  // Автоматически синхронизируем в лист «События»
  let syncResult = { ok: true, written: 0 };
  if (stateObj && stateObj.events) {
    try {
      syncResult = syncEventsToSheet_(stateObj);
    } catch(e) {
      syncResult = { ok: false, error: String(e) };
    }
  }

  return { ok: true, id, updated: nowIso_(), sync: syncResult };
}

function detailsListPublished_() {
  const sh = getSheet_(SH_DET_PUBLISHED);
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    out.push({
      id:       String(row[0]),
      name:     String(row[1] || ''),
      monthKey: String(row[2] || ''),
      updated:  row[3] ? String(row[3]) : ''
    });
  }
  out.sort((a, b) => (b.updated || '').localeCompare(a.updated || ''));
  return { ok: true, versions: out };
}

// ============================================================
// DETAILS — ПРЯМАЯ ОТПРАВКА В АЛЬФА CRM
// (использует функции из Alfa.gs: alfaLogin_, alfaCall_,
//  extractCreatedId_, formatTime_, addHours_, toInt_)
// ============================================================

/**
 * Тест прямого соединения с Альфой — для проверки из приложения.
 */
function detailsTestAlfa_() {
  try {
    if (typeof alfaLogin_ !== 'function') {
      return { ok: false, error: 'Файл Alfa.gs не найден или не сохранён в проекте' };
    }
    const token = alfaLogin_();
    if (!token) return { ok: false, error: 'Не получили токен (проверьте API_KEY в Alfa.gs)' };
    const branches = alfaCall_(token, null, '/v2api/branch/index', {});
    const found = (branches.items || []).find(b => b.id === CFG.BRANCH_ID);
    if (!found) return { ok: false, error: 'Филиал id=' + CFG.BRANCH_ID + ' не найден в Альфе' };
    return {
      ok: true,
      branch: { id: found.id, name: found.name },
      lessonType: CFG.LESSON_TYPE,
      roomId: CFG.ROOM_ID
    };
  } catch(e) {
    return { ok: false, error: String(e.message || e) };
  }
}

/**
 * Главная функция «полного цикла» от менеджера:
 *   1. Сохраняет снимок в Опубликованные_детали
 *   2. Синхронизирует ВСЕ события в лист «События»
 *   3. Для каждого события с publishToAlfa=true:
 *      – логинится в Альфу (один раз)
 *      – создаёт группу + урок
 *      – записывает ID в лист «События»
 *   4. Возвращает сводку
 */
function detailsPublishAndSendToAlfa_(body) {
  // --- 1) Парсим состояние ---
  let stateObj;
  try {
    stateObj = JSON.parse(String(body.stateJson || ''));
  } catch(e) {
    return { ok: false, error: 'Невалидный stateJson' };
  }
  if (!stateObj || !stateObj.events) {
    return { ok: false, error: 'Нет событий в состоянии' };
  }

  // --- 2) Сохраняем снимок ---
  const id = String(body.id || '').trim();
  const name = String(body.name || '').trim();
  if (!id) return { ok: false, error: 'Нет id' };

  const sh = getSheet_(SH_DET_PUBLISHED);
  const monthKey = String(stateObj.monthKey || '');
  const stateJson = JSON.stringify(stateObj);
  const existingRow = findRowById_(sh, id);
  const rowData = [id, name, monthKey, nowIso_(), stateJson];
  if (existingRow > 0) { sh.deleteRow(existingRow); sh.appendRow(rowData); }
  else { sh.appendRow(rowData); }

  // --- 3) Синхронизируем события в лист «События» ---
  let syncResult;
  try {
    syncResult = syncEventsToSheet_(stateObj);
  } catch(e) {
    return { ok: false, error: 'Ошибка синхронизации в «События»: ' + (e.message || e) };
  }

  // --- 4) Берём из stateObj события для отправки в Альфу ---
  const toSend = (stateObj.events || []).filter(e => e.publishToAlfa);
  if (toSend.length === 0) {
    return {
      ok: true,
      saved: true,
      sync: syncResult,
      alfa: { sent: 0, errors: 0, skipped: 0, message: 'Ни у одного события не отмечено «Готово в Альфу»' }
    };
  }

  // Проверяем, что Alfa.gs подключён
  if (typeof alfaLogin_ !== 'function') {
    return {
      ok: false,
      error: 'Файл Alfa.gs не найден. Сохраните его в проекте Apps Script рядом с Code.gs.'
    };
  }

  // --- 5) Логинимся в Альфу ---
  let token;
  try {
    token = alfaLogin_();
    if (!token) throw new Error('Не получили токен');
  } catch(e) {
    return { ok: false, error: 'Ошибка логина в Альфу: ' + (e.message || e) };
  }

  // --- 6) Для отправки нам нужно найти строки в листе «События» по eventId ---
  const eventsSh = getSheet_(SH_DET_EVENTS);
  const lastRow = eventsSh.getLastRow();
  const idColIdx = 28;  // AB — Комментарий с [gid:...]
  const rowByEventId = {};
  if (lastRow >= EVENTS_DATA_START_ROW) {
    const idValues = eventsSh.getRange(EVENTS_DATA_START_ROW, idColIdx, lastRow - EVENTS_DATA_START_ROW + 1, 1).getValues();
    idValues.forEach((row, i) => {
      const v = String(row[0] || '');
      const m = v.match(/\[gid:([\w-]+)\]/);
      if (m) rowByEventId[m[1]] = EVENTS_DATA_START_ROW + i;
    });
  }

  // --- 7) Отправляем каждое событие ---
  let sent = 0, errors = 0, skipped = 0;
  const results = [];

  toSend.forEach(ev => {
    const sheetRow = rowByEventId[ev.id];
    if (!sheetRow) {
      errors++;
      results.push({ id: ev.id, name: ev.name, error: 'Не нашли строку в «События»' });
      return;
    }

    try {
      // Проверяем, не отправлено ли уже
      const existingGroupId = eventsSh.getRange(sheetRow, 20).getValue();  // T
      const existingStatus  = eventsSh.getRange(sheetRow, 19).getValue();  // S
      if (existingGroupId) {
        skipped++;
        results.push({ id: ev.id, name: ev.name, skipped: true, groupId: existingGroupId });
        return;
      }
      if (String(existingStatus).trim() === 'Отменено') {
        skipped++;
        return;
      }

      // Готовим payload для Альфы
      const dateObj = new Date(ev.date);
      const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
      const dateStr = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');

      // Название для Альфы — формула в листе уже посчитала, читаем из колонки H
      let alfaName = String(eventsSh.getRange(sheetRow, 8).getValue() || '');
      if (!alfaName) {
        // Резерв: собираем сами
        alfaName = (ev.format || '') + ' "' + (ev.name || '') + '"';
        if (ev.price) alfaName += ' - ' + ev.price + ' руб.';
      }
      const note = String(ev.description || '').slice(0, 1000);

      // Помечаем «Отправляется»
      eventsSh.getRange(sheetRow, 19).setValue('Отправляется');
      SpreadsheetApp.flush();

      // 7.1) Создаём группу
      const groupPayload = {
        name:           alfaName,
        note:           note,
        branch_ids:     [CFG.BRANCH_ID],
        b_date:         dateStr,
        e_date:         dateStr,
        limit:          toInt_(ev.limit) || 12,
        is_public:      1,
        lesson_type_id: CFG.LESSON_TYPE,
      };
      const grpResp = alfaCall_(token, CFG.BRANCH_ID,
        '/v2api/' + CFG.BRANCH_ID + '/customer-group/create', groupPayload);
      const groupId = extractCreatedId_(grpResp);
      if (!groupId) throw new Error('Не получили id группы');

      // 7.2) Создаём урок
      const startStr = formatTime_(ev.timeStart);
      const endStr = addHours_(startStr, ev.duration || 2.5);
      const lessonPayload = {
        branch_id:        CFG.BRANCH_ID,
        room_id:          CFG.ROOM_ID,
        lesson_type_id:   CFG.LESSON_TYPE,
        customer_group_ids: [groupId],
        streaming_type:   0,
        date:             dateStr,
        time_from:        startStr,
        time_to:          endStr,
        note:             String(ev.name || ''),
        status:           1,
      };
      let lessonId = null;
      try {
        const lessonResp = alfaCall_(token, CFG.BRANCH_ID,
          '/v2api/' + CFG.BRANCH_ID + '/lesson/create', lessonPayload);
        lessonId = extractCreatedId_(lessonResp);
      } catch(le) {
        // Группа создана, но урок упал
        eventsSh.getRange(sheetRow, 19).setValue('Ошибка');
        eventsSh.getRange(sheetRow, 20).setValue(groupId);
        eventsSh.getRange(sheetRow, 22).setValue(new Date());
        eventsSh.getRange(sheetRow, 23).setValue('Группа создана (id=' + groupId + '), урок не создан: ' + (le.message || le));
        errors++;
        results.push({ id: ev.id, name: ev.name, error: 'урок: ' + (le.message || le), groupId });
        return;
      }

      // 7.3) Успех — записываем ID
      eventsSh.getRange(sheetRow, 19).setValue('Создано');
      eventsSh.getRange(sheetRow, 20).setValue(groupId);
      eventsSh.getRange(sheetRow, 21).setValue(lessonId || '');
      eventsSh.getRange(sheetRow, 22).setValue(new Date());
      eventsSh.getRange(sheetRow, 23).setValue('');
      SpreadsheetApp.flush();

      sent++;
      results.push({ id: ev.id, name: ev.name, groupId, lessonId });
    } catch(e) {
      errors++;
      try {
        eventsSh.getRange(sheetRow, 19).setValue('Ошибка');
        eventsSh.getRange(sheetRow, 23).setValue(String(e.message || e));
        eventsSh.getRange(sheetRow, 22).setValue(new Date());
      } catch(_) {}
      results.push({ id: ev.id, name: ev.name, error: String(e.message || e) });
    }
  });

  return {
    ok: true,
    saved: true,
    sync: syncResult,
    alfa: {
      sent,
      errors,
      skipped,
      total: toSend.length,
      results
    }
  };
}

function detailsGetPublished_(id) {
  if (!id) return { ok: false, error: 'нет id' };
  const sh = getSheet_(SH_DET_PUBLISHED);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      return {
        ok: true,
        version: {
          id:        String(data[i][0]),
          name:      String(data[i][1] || ''),
          monthKey:  String(data[i][2] || ''),
          updated:   data[i][3] ? String(data[i][3]) : '',
          stateJson: String(data[i][4] || '')
        }
      };
    }
  }
  return { ok: false, error: 'не найдено' };
}

function detailsSaveDraft_(body) {
  const sh = getSheet_(SH_DET_DRAFTS);
  let id = String(body.id || '').trim();
  const name = String(body.name || 'Без названия').trim();
  const stateObj = body.state || {};
  const stateJson = JSON.stringify(stateObj);

  if (!id) {
    id = 'd_' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
  }
  const rowData = [id, name, nowIso_(), stateJson];
  const existing = findRowById_(sh, id);
  if (existing > 0) sh.deleteRow(existing);
  sh.appendRow(rowData);
  return { ok: true, draft: { id, name, updated: nowIso_() } };
}

function detailsListDrafts_() {
  const sh = getSheet_(SH_DET_DRAFTS);
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    out.push({
      id: String(data[i][0]),
      name: String(data[i][1] || ''),
      updated: data[i][2] ? String(data[i][2]) : ''
    });
  }
  out.sort((a, b) => (b.updated || '').localeCompare(a.updated || ''));
  return { ok: true, drafts: out };
}

function detailsLoadDraft_(id) {
  if (!id) return { ok: false, error: 'нет id' };
  const sh = getSheet_(SH_DET_DRAFTS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      let stateObj = {};
      try { stateObj = JSON.parse(String(data[i][3] || '{}')); } catch(e) {}
      return {
        ok: true,
        draft: {
          id: String(data[i][0]),
          name: String(data[i][1] || ''),
          updated: data[i][2] ? String(data[i][2]) : '',
          state: stateObj
        }
      };
    }
  }
  return { ok: false, error: 'не найдено' };
}

// ============================================================
// DETAILS → СОБЫТИЯ (мост в боевую таблицу)
// ============================================================

/**
 * Проверка: лист «События» существует и доступен.
 */
function detailsPingEvents_() {
  try {
    const sh = getSheet_(SH_DET_EVENTS);
    return { ok: true, lastRow: sh.getLastRow(), name: sh.getName() };
  } catch(e) {
    return { ok: false, error: String(e) };
  }
}

/**
 * Ручная синхронизация «опубликованного состояния» в лист «События».
 * Принимает body.stateJson — тот же объект, что и при publish.
 */
function detailsSyncToEvents_(body) {
  let stateObj;
  try {
    stateObj = JSON.parse(String(body.stateJson || ''));
  } catch(e) {
    return { ok: false, error: 'невалидный stateJson' };
  }
  if (!stateObj || !stateObj.events) return { ok: false, error: 'нет events в state' };
  try {
    return syncEventsToSheet_(stateObj);
  } catch(e) {
    return { ok: false, error: String(e) };
  }
}

/**
 * Записывает события в лист «События»:
 *   – находит существующие строки по eventId (если он сохранён в комментарии или
 *     по совпадению дата+время+название) — обновляет их
 *   – новые события добавляет в конец таблицы
 *
 * Сохраняем минимум полей для совместимости с alfa_integration.gs:
 *   A=Дата, C=Время, D=Длит., F=Формат, G=Название, I=Педагог,
 *   J=Цена, K=Лимит, M=Расход на МК, Q=Описание, R=Готово к публикации, S=Статус,
 *   AB=Комментарий (eventId конструктора).
 */
function syncEventsToSheet_(stateObj) {
  const sh = getSheet_(SH_DET_EVENTS);
  if (!sh) throw new Error('Лист «События» не найден');
  const events = (stateObj.events || []).filter(e => e.date);
  if (events.length === 0) return { ok: true, written: 0 };

  // Загружаем существующую таблицу — индексируем по eventId
  const lastRow = sh.getLastRow();
  const existing = {};   // eventId -> rowNum
  if (lastRow >= EVENTS_DATA_START_ROW) {
    const idCol = 28;   // AB — Комментарий, туда пишем eventId
    const idValues = sh.getRange(EVENTS_DATA_START_ROW, idCol, lastRow - EVENTS_DATA_START_ROW + 1, 1).getValues();
    idValues.forEach((row, i) => {
      const v = String(row[0] || '').trim();
      const m = v.match(/\[gid:([\w-]+)\]/);
      if (m) existing[m[1]] = EVENTS_DATA_START_ROW + i;
    });
  }

  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  let written = 0, updated = 0, created = 0;

  events.forEach(ev => {
    const eventId = ev.id || ('e_' + Math.random().toString(36).slice(2, 10));
    let rowNum = existing[eventId];

    if (!rowNum) {
      // Новая строка — добавляем в конец
      rowNum = sh.getLastRow() + 1;
      // Чтобы не терять формулы — копируем формат и формулы предыдущей строки
      if (rowNum > EVENTS_DATA_START_ROW) {
        const src = sh.getRange(EVENTS_DATA_START_ROW, 1, 1, sh.getMaxColumns());
        const dst = sh.getRange(rowNum, 1, 1, sh.getMaxColumns());
        src.copyTo(dst, { formatOnly: true });
      }
      created++;
    } else {
      updated++;
    }

    const dateObj = ev.date ? new Date(ev.date) : null;
    sh.getRange(rowNum, 1).setValue(dateObj);                              // A — Дата
    sh.getRange(rowNum, 1).setNumberFormat('dd.mm.yyyy');

    // B — день недели (формула, не трогаем — она уже стояла в шаблоне)
    if (ev.timeStart) {
      sh.getRange(rowNum, 3).setValue(ev.timeStart);                       // C
      sh.getRange(rowNum, 3).setNumberFormat('hh:mm');
    }
    sh.getRange(rowNum, 4).setValue(ev.duration || 2.5);                   // D
    // E — формула «время окончания» — оставляем

    sh.getRange(rowNum, 6).setValue(ev.format || '');                      // F
    sh.getRange(rowNum, 7).setValue(ev.name || '');                        // G
    // H — формула «Название для Альфы» — оставляем

    sh.getRange(rowNum, 9).setValue(ev.teacher || '');                     // I
    sh.getRange(rowNum, 10).setValue(ev.price != null ? ev.price : '');    // J
    sh.getRange(rowNum, 11).setValue(ev.limit != null ? ev.limit : '');    // K
    // L — формула «Стоимость педагога» — оставляем
    sh.getRange(rowNum, 13).setValue(ev.costPerPerson != null ? ev.costPerPerson : ''); // M
    // N–P — формулы выручка/расход/маржа — оставляем
    sh.getRange(rowNum, 17).setValue(ev.description || '');                // Q

    // R — готово к публикации: НЕ ставим «Да» автоматически.
    // Прямую отправку делает detailsPublishAndSendToAlfa_.
    // Галочка остаётся для ручного режима (если менеджер открывает таблицу).
    if (ev.publishToAlfa && !sh.getRange(rowNum, 18).getValue()) {
      // Ставим только если ячейка пуста — пометка для истории
      sh.getRange(rowNum, 18).setValue('Да');
    }
    // S — статус Альфа: не трогаем (его пишет alfa_integration.gs)

    // AB — Комментарий с eventId
    let comment = ev.comment || '';
    if (!/\[gid:/.test(comment)) {
      comment = (comment ? comment + ' ' : '') + '[gid:' + eventId + ']';
    }
    sh.getRange(rowNum, 28).setValue(comment);

    written++;
  });

  return { ok: true, written, created, updated };
}

// ============================================================
// СТАРЫЕ camp/intensive — БЕЗ ИЗМЕНЕНИЙ
// (нужны для совместимости с camp-generator и intensive-generator)
// ============================================================

function getAll_(project) {
  return {
    ok: true,
    activities: getActivities_(project),
    details: getDetailsRaw_(project)
  };
}
function getLibrary_(project) { return { ok: true, activities: getActivities_(project) }; }
function getActivities_(project) {
  const sheets = getProjectSheets_(project);
  if (!sheets.activities) return [];
  const sh = getSheet_(sheets.activities);
  const rows = readSheetAsObjects_(sh);
  const out = [];
  rows.forEach(r => {
    const status = String(r['Статус'] || '').toLowerCase().trim();
    if (status !== 'утверждено') return;
    out.push({
      template:    String(r['Шаблон занятия'] || '').trim(),
      examples:    String(r['Примеры тем'] || '').trim(),
      teacher:     String(r['Педагог'] || '').trim(),
      direction:   String(r['Направление'] || '').trim(),
      location:    String(r['Локация / кабинет'] || '').trim(),
      age:         String(r['Возраст'] || '').trim(),
      duration:    Number(r['Длительность (мин)']) || 0,
      format:      String(r['Формат'] || '').trim(),
      description: String(r['Описание'] || '').trim()
    });
  });
  return out;
}
function getDetails_(project) { return { ok: true, details: getDetailsRaw_(project) }; }
function getDetailsRaw_(project) {
  const sheets = getProjectSheets_(project);
  if (!sheets.details) return {};
  let sh;
  try { sh = getSheet_(sheets.details); } catch(e) { return {}; }
  const data = sh.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < data.length; i++) {
    const k = String(data[i][0] || '').trim();
    const v = data[i][1];
    if (k) out[k] = v;
  }
  return out;
}

function publishVersion_(project, body) {
  const id = String(body.id || '').trim();
  const name = String(body.name || '').trim();
  const stateJson = String(body.stateJson || '');
  if (!id) return { ok: false, error: 'нет id' };
  if (!stateJson) return { ok: false, error: 'нет stateJson' };
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.published);
  const existingRow = findRowById_(sh, id);
  let title = name;
  try {
    const obj = JSON.parse(stateJson);
    title = String(obj.title || obj.theme || name);
  } catch(e) {}
  const rowData = [id, name, title, nowIso_(), stateJson];
  if (existingRow > 0) { sh.deleteRow(existingRow); sh.appendRow(rowData); }
  else { sh.appendRow(rowData); }
  return { ok: true, id, updated: nowIso_() };
}
function listPublished_(project) {
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.published);
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    out.push({
      id: String(row[0]), name: String(row[1] || ''),
      title: String(row[2] || ''), updated: row[3] ? String(row[3]) : ''
    });
  }
  out.sort((a, b) => (b.updated || '').localeCompare(a.updated || ''));
  return { ok: true, versions: out };
}
function getPublished_(project, id) {
  if (!id) return { ok: false, error: 'нет id' };
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.published);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      return {
        ok: true,
        version: {
          id: String(data[i][0]), name: String(data[i][1] || ''),
          title: String(data[i][2] || ''),
          updated: data[i][3] ? String(data[i][3]) : '',
          stateJson: String(data[i][4] || '')
        }
      };
    }
  }
  return { ok: false, error: 'не найдено' };
}
function unpublishVersion_(project, id) {
  if (!id) return { ok: false, error: 'нет id' };
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.published);
  const row = findRowById_(sh, id);
  if (row > 0) sh.deleteRow(row);
  return { ok: true };
}

function saveDraft_(project, body) {
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.drafts);
  let id = String(body.id || '').trim();
  const name = String(body.name || 'Без названия').trim();
  const stateObj = body.state || {};
  const stateJson = JSON.stringify(stateObj);
  if (!id) id = 'd_' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
  const rowData = [id, name, nowIso_(), stateJson];
  const existing = findRowById_(sh, id);
  if (existing > 0) sh.deleteRow(existing);
  sh.appendRow(rowData);
  return { ok: true, draft: { id, name, updated: nowIso_() } };
}
function listDrafts_(project) {
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.drafts);
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    out.push({
      id: String(data[i][0]), name: String(data[i][1] || ''),
      updated: data[i][2] ? String(data[i][2]) : ''
    });
  }
  out.sort((a, b) => (b.updated || '').localeCompare(a.updated || ''));
  return { ok: true, drafts: out };
}
function loadDraft_(project, id) {
  if (!id) return { ok: false, error: 'нет id' };
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.drafts);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      let stateObj = {};
      try { stateObj = JSON.parse(String(data[i][3] || '{}')); } catch(e) {}
      return {
        ok: true,
        draft: {
          id: String(data[i][0]), name: String(data[i][1] || ''),
          updated: data[i][2] ? String(data[i][2]) : '', state: stateObj
        }
      };
    }
  }
  return { ok: false, error: 'не найдено' };
}
function deleteDraft_(project, id) {
  if (!id) return { ok: false, error: 'нет id' };
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.drafts);
  const row = findRowById_(sh, id);
  if (row > 0) sh.deleteRow(row);
  return { ok: true };
}

// ============================================================
// SESSIONS / HEARTBEAT
// ============================================================
function heartbeat_(project, body) {
  const sessionId = String(body.sessionId || '').trim();
  const publishId = String(body.publishId || '').trim();
  if (!sessionId || !publishId) return { ok: false, error: 'нет sessionId или publishId' };
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.sessions);
  const data = sh.getDataRange().getValues();
  const now = Date.now();
  // Чистим протухшие
  for (let i = data.length - 1; i >= 1; i--) {
    const updated = data[i][4] ? new Date(data[i][4]).getTime() : 0;
    if (now - updated > SESSION_TIMEOUT_MS) sh.deleteRow(i + 1);
  }
  const freshData = sh.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 1; i < freshData.length; i++) {
    if (String(freshData[i][0]) === sessionId) { foundRow = i + 1; break; }
  }
  const rowData = [sessionId, publishId, body.editorName || '', body.userAgent || '', nowIso_()];
  if (foundRow > 0) sh.getRange(foundRow, 1, 1, rowData.length).setValues([rowData]);
  else sh.appendRow(rowData);
  return { ok: true };
}

function endSession_(project, sessionId) {
  if (!sessionId) return { ok: false, error: 'нет sessionId' };
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.sessions);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === sessionId) { sh.deleteRow(i + 1); break; }
  }
  return { ok: true };
}

function listSessions_(project, publishId) {
  if (!publishId) return { ok: true, sessions: [] };
  const sheets = getProjectSheets_(project);
  const sh = getSheet_(sheets.sessions);
  const data = sh.getDataRange().getValues();
  const now = Date.now();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) !== String(publishId)) continue;
    const updated = data[i][4] ? new Date(data[i][4]).getTime() : 0;
    const ageMs = now - updated;
    if (ageMs > SESSION_TIMEOUT_MS) continue;
    out.push({
      sessionId: String(data[i][0]),
      publishId: String(data[i][1]),
      editorName: String(data[i][2] || ''),
      userAgent: String(data[i][3] || ''),
      updated: data[i][4] ? String(data[i][4]) : '',
      ageSec: Math.round(ageMs / 1000)
    });
  }
  return { ok: true, sessions: out };
}

// ============================================================
// CLAUDE PROXY
// ============================================================
function callClaude_(prompt, system) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: false, error: 'ANTHROPIC_API_KEY не настроен в Script Properties' };
  if (!prompt) return { ok: false, error: 'нет prompt' };

  const messages = [{ role: 'user', content: prompt }];
  const payload = {
    model: ANTHROPIC_MODEL,
    max_tokens: ANTHROPIC_MAX_TOKENS,
    messages: messages
  };
  if (system) payload.system = system;

  try {
    const response = UrlFetchApp.fetch(ANTHROPIC_API_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = response.getResponseCode();
    const text = response.getContentText();
    if (code < 200 || code >= 300) {
      return { ok: false, error: 'Anthropic ' + code + ': ' + text.slice(0, 500) };
    }
    const data = JSON.parse(text);
    let outText = '';
    if (Array.isArray(data.content)) {
      data.content.forEach(c => { if (c.type === 'text') outText += c.text; });
    }
    return { ok: true, text: outText };
  } catch(e) {
    return { ok: false, error: String(e.message || e) };
  }
}

// ============================================================
// DEBUG (запускайте вручную из редактора)
// ============================================================

function _testInitDetails() {
  // Создаёт все нужные листы для Деталей
  ['Форматы_детали', 'Педагоги_детали', 'Шаблоны_детали',
   'Опубликованные_детали', 'Черновики_детали', 'Сессии_детали']
   .forEach(name => {
    const sh = getSheet_(name);
    Logger.log('OK: ' + name + ' (' + sh.getLastRow() + ' rows)');
  });
}

function _testSeedDetails() {
  // Заполняет справочники начальными данными
  const formats = getSheet_(SH_DET_FORMATS);
  if (formats.getLastRow() < 2) {
    formats.getRange(2, 1, 16, 7).setValues([
      ['МК Глина', 'Творчество', 'Лепка из глины. Работаем с керамистом.', 65, 12, 2.5, 'Да'],
      ['МК Рисование', 'Творчество', 'Текстурная паста, акрил, холст.', 80, 12, 2.5, 'Да'],
      ['МК Украшения', 'Творчество', 'Бисер, макраме, броши.', 80, 12, 2.5, 'Да'],
      ['МК Бьюти', 'Уход', 'Косметология, уход за кожей.', 95, 10, 2, 'Да'],
      ['МК Кулинарный', 'Творчество', 'Кулинарные мастер-классы.', 90, 10, 2.5, 'Да'],
      ['Киновечер', 'Развлечения', 'Просмотр и обсуждение фильма.', 30, 10, 3, 'Да'],
      ['Книжный клуб', 'Образование', 'Обсуждение книги, лекция.', 30, 12, 2.5, 'Да'],
      ['Настолки', 'Развлечения', '', 30, 12, 3, 'Да'],
      ['Разговорник', 'Образование', '', 30, 10, 2, 'Да'],
      ['Караоке', 'Развлечения', '', 30, 12, 3, 'Да'],
      ['Вечеринка', 'Развлечения', 'Тематическая вечеринка с программой.', 120, 12, 4, 'Да'],
      ['Встреча с экспертом', 'Образование', 'Лекция и обсуждение с приглашённым экспертом.', 50, 12, 2, 'Да'],
      ['Дегустация', 'Развлечения', 'Винная или другая дегустация.', 100, 12, 2.5, 'Да'],
      ['День рождения', 'Частное', 'Закрытое индивидуальное мероприятие.', 500, 12, 3, 'Да'],
      ['Свидание', 'Частное', 'Камерное свидание/съёмка.', '', 2, 2, 'Да'],
      ['МК', 'Творчество', 'Универсальный мастер-класс.', 70, 12, 2.5, 'Да']
    ]);
  }
  const teachers = getSheet_(SH_DET_TEACHERS);
  if (teachers.getLastRow() < 2) {
    teachers.getRange(2, 1, 14, 6).setValues([
      ['Афанасенко Евгения', 100, 57, '', 'Глина / керамика', 'Да'],
      ['Бизунова Ольга', 250, '', '', 'Дегустации', 'Да'],
      ['Власова Юлия', 150, '', '', 'Секс-просвет', 'Да'],
      ['Вязникова Алеся', 500, '', '', 'Вечеринки, ведущая', 'Да'],
      ['Гайкова Тамара', 100, '', '', '', 'Да'],
      ['Костикова Ольга', 100, '', '', 'Украшения', 'Да'],
      ['Крючкова Ирина', 20, '', '', 'Косметология', 'Да'],
      ['Минич Валерия', 50, '', '', 'Книжный клуб', 'Да'],
      ['Парджанадзе Галина', 80, '', '', 'Бисер', 'Да'],
      ['Печковская Юля', 100, '', '', '', 'Да'],
      ['Райкова Ирина', '', '', '', '', 'Да'],
      ['Станченко Екатерина', 100, '', '', '', 'Да'],
      ['Зоркина Наталья', '', '', '', 'Киновечер', 'Да'],
      ['Громова Елена', '', '', '', '', 'Да']
    ]);
  }
  Logger.log('Seeded.');
}
