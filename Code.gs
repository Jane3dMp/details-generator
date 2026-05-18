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
const SH_DET_THUMBS     = 'Превью_кэш'; // кэш base64-картинок для референсов

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
    [SH_DET_TEMPLATES]:  ['Название', 'Формат', 'Педагог', 'Описание', 'Цена', 'Лимит', 'Длительность (ч)', 'Теги', 'Использовалось раз', 'Референсы (JSON)'],
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
    if (project === 'competitors') return doGetCompetitors_(action, params);

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
    if (project === 'competitors') return doPostCompetitors_(action, body);

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
    case 'publicEvents':  return jsonResponse(detailsGetPublicEvents_(params));
    case 'teamEvents':    return jsonResponse(detailsGetTeamEvents_(params));
    case 'bookingCounts': return jsonResponse(detailsBookingCounts_(params));
    case 'weekJson':      return jsonResponse(detailsGetWeekJson_(params));
    case 'postersList':   return jsonResponse(detailsPostersList_());
    case 'posterGet':     return jsonResponse(detailsPosterGet_(params));
    default: return jsonResponse({ ok: false, error: 'Unknown details action: ' + action });
  }
}

function doPostDetails_(action, body) {
  switch(action) {
    case 'claude':              return jsonResponse(callClaude_(body.prompt, body.system));
    case 'publish':             return jsonResponse(detailsPublish_(body));
    case 'publishAndSendToAlfa':return jsonResponse(detailsPublishAndSendToAlfa_(body));
    case 'syncEventsToAlfa':    return jsonResponse(detailsSyncEventsToAlfa_(body));
    case 'archiveInAlfa':       return jsonResponse(detailsArchiveInAlfa_(body));
    case 'linkAlfaGroup':       return jsonResponse(detailsLinkAlfaGroup_(body));
    case 'pullFromAlfa':        return jsonResponse(detailsPullFromAlfa_(body));
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
    case 'fetchOgImage':        return jsonResponse(detailsFetchOgImage_(body));
    case 'pushSnapshot':        return jsonResponse(detailsPushSnapshot_(body));
    case 'rebuildFromAlfaDryRun': return jsonResponse(detailsRebuildFromAlfa_(body, false));
    case 'rebuildFromAlfaApply':  return jsonResponse(detailsRebuildFromAlfa_(body, true));
    case 'importNewFromAlfa':     return jsonResponse(detailsImportNewFromAlfa_(body));
    case 'clearSnapshot':         return jsonResponse(detailsClearSnapshot_(body));
    case 'testTelegramConnection':   return jsonResponse(detailsTestTelegram_());
    case 'previewTelegramPost':      return jsonResponse(detailsPreviewTelegramPost_(body));
    case 'postToTelegram':           return jsonResponse(detailsPostToTelegram_(body));
    case 'postPhotoToTelegram':      return jsonResponse(detailsPostPhotoToTelegram_(body));
    case 'posterSave':               return jsonResponse(detailsPosterSave_(body));
    case 'posterDelete':             return jsonResponse(detailsPosterDelete_(body));
    case 'previewEventDraft':        return jsonResponse(detailsPreviewEventDraft_(body));
    case 'postEventDraftToTelegram': return jsonResponse(detailsPostEventDraftToTelegram_(body));
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
    references:  (function(){
      try {
        const raw = String(r['Референсы (JSON)'] || '').trim();
        if (!raw) return [];
        const arr = JSON.parse(raw);
        return Array.isArray(arr) ? arr : [];
      } catch(e) { return []; }
    })(),
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

  // Сериализуем references (только лёгкие поля — url/note/thumb/domain).
  // thumb может содержать маркер 'thumb:<hash>' или короткий URL — это норм.
  // Полные data:image; URL сюда лучше не пихать — лист имеет 50000 лимит на ячейку.
  const refsArr = Array.isArray(t.references) ? t.references.map(r => ({
    url:    String((r && r.url) || ''),
    note:   String((r && r.note) || ''),
    thumb:  String((r && r.thumb) || ''),
    domain: String((r && r.domain) || '')
  })).filter(r => r.url) : [];
  // Если случайно прилетел гигантский thumb (data:...) — обрезаем его, оставляя только маркер/URL
  refsArr.forEach(r => {
    if (r.thumb && r.thumb.length > 500 && r.thumb.indexOf('data:') === 0) {
      r.thumb = '';   // в кэше thumb-листа всё равно лежит реальная картинка
    }
  });
  let refsJson = refsArr.length ? JSON.stringify(refsArr) : '';
  if (refsJson.length > 49000) refsJson = '';  // безопасность

  // Если такое название уже есть — обновляем references + увеличиваем счётчик
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === t.name) {
      const used = Number(data[i][8]) || 0;
      sh.getRange(i + 1, 9).setValue(used + 1);
      // Если в шаблоне явно переданы НОВЫЕ referenсы — перезаписываем колонку 10
      if (refsArr.length > 0 || (t.references && Array.isArray(t.references))) {
        sh.getRange(i + 1, 10).setValue(refsJson);
      }
      return { ok: true, updated: true, refsCount: refsArr.length };
    }
  }
  sh.appendRow([
    t.name, t.format || '', t.teacher || '', t.description || '',
    t.price || '', t.limit || '', t.duration || '', t.tags || '', 1,
    refsJson
  ]);
  return { ok: true, created: true, refsCount: refsArr.length };
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

// ============================================================
// DETAILS — СИНХРОНИЗАЦИЯ ИЗМЕНЕНИЙ В АЛЬФУ
// (для статуса 'modified' — события уже опубликованы, локально изменились)
// ============================================================

/**
 * Принимает массив событий, для каждого делает customer-group/update +
 * lesson/update в Альфе (используя существующий groupId / lessonId).
 * Также обрабатывает создание новых (status='draft' с publishToAlfa=true)
 * и архивацию (status='archived').
 */
function detailsSyncEventsToAlfa_(body) {
  // Приходит state целиком и список операций:
  //   body.events = [{ id, status: 'create'|'update'|'archive', data: {...} }]
  // ИЛИ body.stateJson + флаг автоопределения по полям.
  let stateObj;
  try { stateObj = JSON.parse(String(body.stateJson || '')); }
  catch(e) { return { ok: false, error: 'невалидный stateJson' }; }
  if (!stateObj || !stateObj.events) return { ok: false, error: 'нет events' };

  // Сохраняем снимок (как обычная публикация — для истории)
  const id = String(body.id || '').trim();
  const name = String(body.name || '').trim();
  if (!id) return { ok: false, error: 'нет id' };
  const sh = getSheet_(SH_DET_PUBLISHED);
  const monthKey = String(stateObj.monthKey || '');
  const stateJson = JSON.stringify(stateObj);
  const existingRow = findRowById_(sh, id);
  const rowData = [id, name, monthKey, nowIso_(), stateJson];
  if (existingRow > 0) { sh.deleteRow(existingRow); sh.appendRow(rowData); }
  else { sh.appendRow(rowData); }

  // Синхронизируем в лист «События»
  let syncResult;
  try { syncResult = syncEventsToSheet_(stateObj); }
  catch(e) { return { ok: false, error: 'ошибка синка в События: ' + (e.message || e) }; }

  // Разделяем события по статусам
  const events = stateObj.events || [];
  const toCreate  = events.filter(e => e.status === 'draft' && e.publishToAlfa);
  const toUpdate  = events.filter(e => e.status === 'modified');
  const toArchive = events.filter(e => e.status === 'archived' && e.alfaGroupId);

  if (toCreate.length === 0 && toUpdate.length === 0 && toArchive.length === 0) {
    return {
      ok: true,
      saved: true,
      sync: syncResult,
      alfa: { created: 0, updated: 0, archived: 0, errors: 0, message: 'Нет изменений для отправки в Альфу' }
    };
  }

  // Проверка Alfa.gs
  if (typeof alfaLogin_ !== 'function') {
    return { ok: false, error: 'Файл Alfa.gs не найден' };
  }
  let token;
  try {
    token = alfaLogin_();
    if (!token) throw new Error('Не получили токен');
  } catch(e) {
    return { ok: false, error: 'Ошибка логина в Альфу: ' + (e.message || e) };
  }

  const eventsSh = getSheet_(SH_DET_EVENTS);
  const rowByEventId = buildEventRowIndex_(eventsSh);

  let created = 0, updated = 0, archived = 0, errors = 0;
  const results = [];

  // 1) CREATE — создаём новые
  toCreate.forEach(ev => {
    try {
      const res = alfaCreateGroupAndLesson_(token, ev, eventsSh, rowByEventId[ev.id]);
      created++;
      results.push({ id: ev.id, name: ev.name, op: 'create', groupId: res.groupId, lessonId: res.lessonId, formUrl: res.formUrl });
    } catch(e) {
      errors++;
      results.push({ id: ev.id, name: ev.name, op: 'create', error: String(e.message || e) });
    }
  });

  // 2) UPDATE — обновляем
  toUpdate.forEach(ev => {
    try {
      const res = alfaUpdateGroupAndLesson_(token, ev, eventsSh, rowByEventId[ev.id]);
      updated++;
      results.push({ id: ev.id, name: ev.name, op: 'update', groupId: res.groupId, lessonId: res.lessonId });
    } catch(e) {
      errors++;
      results.push({ id: ev.id, name: ev.name, op: 'update', error: String(e.message || e) });
    }
  });

  // 3) ARCHIVE — архивируем
  toArchive.forEach(ev => {
    try {
      alfaArchiveGroup_(token, ev.alfaGroupId);
      archived++;
      // Помечаем в листе «События»
      const sheetRow = rowByEventId[ev.id];
      if (sheetRow) {
        eventsSh.getRange(sheetRow, 19).setValue('Отменено');
        eventsSh.getRange(sheetRow, 23).setValue('Архивировано в Альфе ' + new Date().toISOString().slice(0,16));
      }
      results.push({ id: ev.id, name: ev.name, op: 'archive', groupId: ev.alfaGroupId });
    } catch(e) {
      errors++;
      results.push({ id: ev.id, name: ev.name, op: 'archive', error: String(e.message || e) });
    }
  });

  return {
    ok: true,
    saved: true,
    sync: syncResult,
    alfa: { created, updated, archived, errors, total: toCreate.length + toUpdate.length + toArchive.length, results }
  };
}

/**
 * Архивация одной группы (на случай, если событие удалили из конструктора —
 * фронтенд может слать сразу запрос архивации).
 */
function detailsArchiveInAlfa_(body) {
  const groupId = body.groupId;
  if (!groupId) return { ok: false, error: 'нет groupId' };
  if (typeof alfaLogin_ !== 'function') return { ok: false, error: 'Файл Alfa.gs не найден' };
  try {
    const token = alfaLogin_();
    alfaArchiveGroup_(token, groupId);
    return { ok: true, archived: groupId };
  } catch(e) {
    return { ok: false, error: String(e.message || e) };
  }
}

/**
 * Тянет свежие данные группы и связанного урока из Альфы.
 * body.groupIds — массив ID групп Альфы.
 * Возвращает: { ok: true, items: { [groupId]: { name, note, limit, b_date, time_from, duration, lessonId, removed } } }
 *
 * Формат полей:
 *   - name: строка
 *   - note: строка (для Альфы максимум 255 символов)
 *   - limit: число
 *   - b_date: 'yyyy-MM-dd' (как Альфа отдаёт в ответе)
 *   - time_from: 'HH:mm' (выдран из time_from урока, если урок есть)
 *   - duration: минуты (вычисляется из time_from/time_to урока)
 *   - lessonId: id первого найденного урока группы (или null)
 *   - removed: 0 или 1
 */
// dead copy of detailsPullFromAlfa_ removed — see active version below

function detailsBookingCounts_(params) {
  const idsRaw = String(params.groupIds || '').trim();
  if (!idsRaw) return { ok: false, error: 'нет groupIds' };
  const groupIds = idsRaw.split(',').map(x => parseInt(x, 10)).filter(x => x > 0);
  if (groupIds.length === 0) return { ok: false, error: 'пустой groupIds' };

  const cache = CacheService.getScriptCache();
  const sortedKey = 'bookings_v3_' + groupIds.slice().sort((a,b) => a - b).join('_');
  const cacheKey = sortedKey.length > 240 ? sortedKey.slice(0, 240) : sortedKey;
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      const obj = JSON.parse(cached);
      obj.fromCache = true;
      return obj;
    } catch(e) {}
  }

  if (typeof alfaLogin_ !== 'function') return { ok: false, error: 'Alfa.gs не найден' };
  let token;
  try { token = alfaLogin_(); }
  catch(e) { return { ok: false, error: 'Логин в Альфу: ' + (e.message || e) }; }

  const counts = {};
  groupIds.forEach(gid => {
    try {
      const resp = alfaCall_(token, CFG.BRANCH_ID,
        '/v2api/' + CFG.BRANCH_ID + '/customer/index',
        { group_ids: [gid], removed: [0], page: 0 });
      let n = Number(resp.total);
      if (isNaN(n)) n = (resp.items || []).length;
      counts[gid] = n;
    } catch(e) {
      counts[gid] = null;
    }
  });

  const result = { ok: true, counts, fetchedAt: nowIso_() };
  try { cache.put(cacheKey, JSON.stringify(result), 180); } catch(e) {}
  return result;
}
// ============================================================
// REBUILD FROM ALFA — переписать админку из Альфы (Альфа = мастер)
// ============================================================
/**
 * Полная пересборка working_main черновика и шаблонов из данных Альфы.
 *
 * Этапы:
 *   1) Логин в Альфу, тянем все группы + уроки за широкое окно
 *   2) Превращаем каждую активную группу в событие конструктора
 *      (формат подбираем из справочника Форматы_детали)
 *   3) Собираем уникальные шаблоны (формат + название)
 *   4) Если apply=false — возвращаем превью (что произойдёт)
 *   5) Если apply=true:
 *      – бэкапим текущий working_main как backup_before_rebuild_<ts>
 *      – полностью перезаписываем working_main новым набором
 *      – бэкапим Шаблоны_детали в новый лист с timestamp
 *      – полностью перезаписываем Шаблоны_детали
 *      – перезаписываем snapshot_main в Опубликованные_детали
 */
function detailsRebuildFromAlfa_(body, apply) {
  if (typeof alfaLogin_ !== 'function') return { ok: false, error: 'Alfa.gs не найден' };

  let token;
  try { token = alfaLogin_(); }
  catch(e) { return { ok: false, error: 'Логин в Альфу: ' + (e.message || e) }; }

  // 1) Тянем группы
  const allGroups = [];
  let page = 0;
  const seenG = {};
  while (true) {
    let resp;
    try {
      resp = alfaCall_(token, CFG.BRANCH_ID, '/v2api/' + CFG.BRANCH_ID + '/group/index', { page });
    } catch(e) { return { ok: false, error: 'group/index: ' + (e.message || e) }; }
    const items = resp.items || [];
    if (items.length === 0) break;
    let news = 0;
    items.forEach(g => {
      if (!seenG[g.id]) { seenG[g.id] = true; allGroups.push(g); news++; }
    });
    if (news === 0) break;
    page++;
    if (page > 30) break;
  }

  // Только активные нашего филиала
  const activeGroups = allGroups.filter(g => {
    if (g.removed === 1) return false;
    if (g.branch_ids && g.branch_ids.indexOf(CFG.BRANCH_ID) < 0) return false;
    return true;
  });

  // 2) Тянем уроки за широкое окно
  let tz = ''; try { tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone(); } catch(e){}
  if (!tz) tz = 'Europe/Moscow';
  const now = new Date();
  const dateFrom = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth() - 6, 1), tz, 'yyyy-MM-dd');
  const dateTo   = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth() + 6, 28), tz, 'yyyy-MM-dd');
  const lessonByGroupId = {};
  let lpage = 0;
  const seenL = {};
  while (true) {
    let resp;
    try {
      resp = alfaCall_(token, CFG.BRANCH_ID, '/v2api/' + CFG.BRANCH_ID + '/lesson/index',
        { date_from: dateFrom, date_to: dateTo, page: lpage });
    } catch(e) { break; }
    const items = resp.items || [];
    if (items.length === 0) break;
    let news = 0;
    items.forEach(l => {
      if (!seenL[l.id]) {
        seenL[l.id] = true; news++;
        (l.group_ids || []).forEach(gid => {
          // Берём самый ранний урок группы
          if (!lessonByGroupId[gid] || (l.date && l.date < lessonByGroupId[gid].date)) {
            lessonByGroupId[gid] = l;
          }
        });
      }
    });
    if (news === 0) break;
    lpage++;
    if (lpage > 50) break;
  }

  // 3) Список форматов из справочника, отсортированный по длине (длинные раньше)
  const formats = detailsGetFormats_().map(f => f.name).filter(Boolean);
  formats.sort((a, b) => b.length - a.length);

  // 4) Превращаем группы в события
  const events = [];
  const tplsByKey = {};   // 'формат|название' → шаблон
  const stats = {
    totalGroups: allGroups.length,
    activeGroups: activeGroups.length,
    eventsToCreate: 0,
    templatesToCreate: 0,
    formatMatched: 0,
    formatUnknown: 0,
    samples: []
  };

  activeGroups.forEach(g => {
    const parsed = _parseAlfaGroupName_(g.name, formats);
    if (parsed.formatMatched) stats.formatMatched++; else stats.formatUnknown++;

    // Дата группы → события
    let dateStr = '';
    const ld = lessonByGroupId[g.id];
    if (ld && ld.date) {
      dateStr = String(ld.date).slice(0, 10);
    } else if (g.b_date) {
      // b_date в формате 'dd.MM.yyyy' (после твоего предыдущего ответа)
      // или 'yyyy-MM-dd' (бывает по-разному в зависимости от запроса)
      dateStr = normalizeDate_(g.b_date);
    }
    if (!dateStr) {
      // Без даты пропускаем — события без даты не имеют смысла
      return;
    }

    // Время и длительность из урока
    let timeStart = '18:00';
    let durationH = 2.5;
    if (ld) {
      if (ld.time_from) {
        const m = String(ld.time_from).match(/(\d{2}):(\d{2})/);
        if (m) timeStart = m[1] + ':' + m[2];
      }
      if (ld.time_from && ld.time_to) {
        const fT = new Date(ld.time_from), tT = new Date(ld.time_to);
        if (!isNaN(fT) && !isNaN(tT)) {
          const minutes = Math.round((tT - fT) / 60000);
          if (minutes > 0 && minutes < 24 * 60) durationH = +(minutes / 60).toFixed(2);
        }
      }
    }

    const ev = {
      id: 'alfa_' + g.id,                        // стабильный ID на основе группы
      date: dateStr,
      timeStart,
      duration: durationH,
      format: parsed.format,
      name: parsed.name,
      teacher: '',                                // потом руками
      price: parsed.price,
      limit: Number(g.limit) || 12,
      costPerPerson: null,
      costGroup: null,
      description: '',                            // длинное описание — позже маркетологи
      alfaNote: String(g.note || ''),             // короткое — то что в Альфе
      status: 'published',
      publishToAlfa: false,
      publishedToAlfa: true,
      alfaGroupId: g.id,
      alfaLessonId: ld ? ld.id : null,
      formUrl: CFG.ALFA_HOST + '/common/' + CFG.BRANCH_ID + '/lead/create?gid=' + g.id,
      comment: '',
      references: [],
      teamComment: ''
    };
    events.push(ev);

    // Шаблон собираем только если есть и формат и название
    if (parsed.format && parsed.name) {
      const key = (parsed.format + '|' + parsed.name).toLowerCase();
      if (!tplsByKey[key]) {
        tplsByKey[key] = {
          name: parsed.name,
          format: parsed.format,
          teacher: '',
          description: '',
          price: parsed.price,
          limit: Number(g.limit) || 12,
          duration: durationH,
          tags: '',
          usedCount: 1
        };
      } else {
        tplsByKey[key].usedCount++;
      }
    }

    if (stats.samples.length < 8) {
      stats.samples.push({
        gid: g.id,
        original: g.name,
        format: parsed.format,
        name: parsed.name,
        price: parsed.price,
        date: dateStr,
        time: timeStart
      });
    }
  });

  // Сортируем события по дате+времени
  events.sort((a, b) => (a.date + (a.timeStart || '')).localeCompare(b.date + (b.timeStart || '')));
  stats.eventsToCreate = events.length;
  stats.templatesToCreate = Object.keys(tplsByKey).length;

  // 5) Если dry-run — возвращаем превью
  if (!apply) {
    return {
      ok: true,
      dryRun: true,
      stats,
      preview: events.slice(0, 20).map(e => ({
        date: e.date, time: e.timeStart, format: e.format, name: e.name,
        price: e.price, limit: e.limit, gid: e.alfaGroupId
      }))
    };
  }

  // 6) APPLY
  // 6.1) Бэкап working_main (с поддержкой разбитого на куски JSON)
  const draftsSh = getSheet_(SH_DET_DRAFTS);
  const workRow = findRowById_(draftsSh, WORKING_DRAFT_ID);
  const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  if (workRow > 0) {
    const lastCol = Math.max(4, draftsSh.getLastColumn());
    const cells = draftsSh.getRange(workRow, 1, 1, lastCol).getValues()[0];
    // Копируем все колонки начиная с 4 как есть (могут быть chunks)
    const chunkCols = cells.slice(3);
    draftsSh.appendRow([
      'backup_before_rebuild_' + ts,
      'Бэкап перед перезаписью из Альфы (' + ts + ')',
      new Date().toISOString(),
      ...chunkCols
    ]);
    draftsSh.deleteRow(workRow);
  }

  // 6.2) Записываем новый working_main
  const newState = {
    monthKey: '',
    events,
    publishedAt: new Date().toISOString(),
    editorName: 'rebuild from Alfa',
    isWorking: true,
    rebuiltAt: new Date().toISOString()
  };
  // 6.2) Записываем новый working_main с разбиением на куски
  _writeWorkingDraft_(newState, 'Рабочий черновик · перестроен из Альфы ' + ts);

  // 6.3) Бэкап Шаблоны_детали (копируем лист)
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const tplsSh = ss.getSheetByName(SH_DET_TEMPLATES);
  if (tplsSh) {
    const backupName = 'Шаблоны_бэкап_' + ts.slice(0, 10);
    if (!ss.getSheetByName(backupName)) {
      tplsSh.copyTo(ss).setName(backupName);
    }
  }

  // 6.4) Перезаписываем Шаблоны_детали
  const tplOrder = Object.keys(tplsByKey);
  if (tplsSh) {
    const lastRowT = tplsSh.getLastRow();
    if (lastRowT > 1) {
      tplsSh.getRange(2, 1, lastRowT - 1, 9).clearContent();
    }
    const tplRows = tplOrder.map(k => {
      const t = tplsByKey[k];
      return [t.name, t.format, t.teacher, t.description, t.price || '', t.limit || '', t.duration || '', t.tags, t.usedCount];
    });
    if (tplRows.length > 0) {
      tplsSh.getRange(2, 1, tplRows.length, 9).setValues(tplRows);
    }
  }

  // 6.5) Snapshot для view.html и team.html
  const pubSh = getSheet_(SH_DET_PUBLISHED);
  const snapId = 'snapshot_main';
  const snapName = 'Снимок · перестроен из Альфы ' + ts.slice(0, 10);
  const existSnap = findRowById_(pubSh, snapId);
  if (existSnap > 0) pubSh.deleteRow(existSnap);
  pubSh.appendRow([snapId, snapName, '', new Date().toISOString(), JSON.stringify(newState)]);

  return {
    ok: true,
    applied: true,
    stats,
    backups: {
      workingMain: 'backup_before_rebuild_' + ts,
      templates: tplsSh ? ('Шаблоны_бэкап_' + ts.slice(0, 10)) : null
    }
  };
}

/**
 * Парсит имя группы Альфы.
 * Возвращает: { format, name, price, formatMatched }
 *
 * Логика:
 *   1) Откусываем хвост ' - 120 руб.' → price
 *   2) Ищем в начале строки совпадение с одним из формат-имён (по убыванию длины)
 *   3) Если нашли формат — что осталось после формата (минус кавычки) идёт в name
 *   4) Если формата не нашли — name = всё что без хвоста, format = ''
 */
function _parseAlfaGroupName_(rawName, knownFormats) {
  const result = { format: '', name: '', price: null, formatMatched: false };
  if (!rawName) return result;
  let str = String(rawName).trim();

  // 1) Цена
  const priceMatch = str.match(/-\s*(\d+)\s*руб\.?\s*$/i);
  if (priceMatch) {
    result.price = Number(priceMatch[1]);
    str = str.replace(/\s*-\s*\d+\s*руб\.?\s*$/i, '').trim();
  }

  // 2) Ищем формат в начале (knownFormats отсортированы по убыванию длины)
  let matchedFormat = '';
  let rest = str;
  for (const f of knownFormats) {
    if (!f) continue;
    if (str.toLowerCase().indexOf(f.toLowerCase()) === 0) {
      matchedFormat = f;
      rest = str.slice(f.length).trim();
      // Снимаем разделители
      rest = rest.replace(/^[:\-—–\s]+/, '').trim();
      break;
    }
  }

  if (matchedFormat) {
    result.format = matchedFormat;
    result.formatMatched = true;
    // Если остаток в кавычках — выдираем содержимое
    const quoted = rest.match(/^[«"„]([^»"”]+)[»"”]\s*$/);
    if (quoted) rest = quoted[1].trim();
    // Если остаток пустой — бывает «Караоке» без названия. Тогда name=пусто, format=Караоке.
    result.name = rest;
  } else {
    // Формат не определился — кладём всё в name (минус возможные обрамляющие кавычки)
    const quoted = str.match(/^[«"„]([^»"”]+)[»"”]\s*$/);
    result.name = quoted ? quoted[1].trim() : str;
  }

  return result;
}

// ----- Хелперы для работы с Альфой через Code.gs -----

function buildEventRowIndex_(eventsSh) {
  const rowByEventId = {};
  const lastRow = eventsSh.getLastRow();
  if (lastRow >= EVENTS_DATA_START_ROW) {
    const idValues = eventsSh.getRange(EVENTS_DATA_START_ROW, 28, lastRow - EVENTS_DATA_START_ROW + 1, 1).getValues();
    idValues.forEach((row, i) => {
      const v = String(row[0] || '');
      const m = v.match(/\[gid:([\w-]+)\]/);
      if (m) rowByEventId[m[1]] = EVENTS_DATA_START_ROW + i;
    });
  }
  return rowByEventId;
}

function alfaCreateGroupAndLesson_(token, ev, eventsSh, sheetRow) {
  let tz = ''; try { tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone(); } catch(e){}
  if (!tz) tz = 'Europe/Moscow';
  const dateObj = new Date(ev.date);
  const dateStr = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
  // Альфа CRM (S20) ждёт даты в формате dd.MM.yyyy для b_date/e_date группы.
  // КРИТИЧНО: e_date должен быть СТРОГО больше b_date (даже для однодневных событий) — иначе «Некорректный период».
  const dateStrRu = Utilities.formatDate(dateObj, tz, 'dd.MM.yyyy');
  const dateNext = new Date(dateObj.getTime() + 24 * 60 * 60 * 1000);
  const dateStrRuNext = Utilities.formatDate(dateNext, tz, 'dd.MM.yyyy');

  let alfaName = '';
  if (sheetRow) alfaName = String(eventsSh.getRange(sheetRow, 8).getValue() || '');
  if (!alfaName) {
    alfaName = (ev.format || '') + ' "' + (ev.name || '') + '"';
    if (ev.price) alfaName += ' - ' + ev.price + ' руб.';
  }

  const note = String(ev.alfaNote || ev.description || '').slice(0, 255);
  const groupPayload = {
    name: alfaName, note: note, branch_ids: [CFG.BRANCH_ID],
    b_date: dateStrRu, e_date: dateStrRuNext,
    limit: toInt_(ev.limit) || 12, is_public: 1,
    lesson_type_id: CFG.LESSON_TYPE,
  };
  Logger.log('[alfaCreateGroupAndLesson_] ev.date=' + ev.date + ' tz=' + tz);
  Logger.log('[alfaCreateGroupAndLesson_] dateObj=' + dateObj.toISOString());
  Logger.log('[alfaCreateGroupAndLesson_] dateStrRu=' + dateStrRu + ' dateStrRuNext=' + dateStrRuNext);
  Logger.log('[alfaCreateGroupAndLesson_] payload=' + JSON.stringify(groupPayload));
  const grpResp = alfaCall_(token, CFG.BRANCH_ID,
    '/v2api/' + CFG.BRANCH_ID + '/group/create', groupPayload);
  const groupId = extractCreatedId_(grpResp);
  if (!groupId) throw new Error('Не получили id группы');

  // Урок
  const startStr = formatTime_(ev.timeStart);                 // 'HH:mm'
  const endStr = addHours_(startStr, ev.duration || 2.5);     // 'HH:mm'
  const durationMin = Math.round((ev.duration || 2.5) * 60);
  // Формат принят по образцу рабочего проекта Интенсивов:
  //   lesson_date dd.MM.yyyy, time_from/time_to как 'HH:mm' (без секунд),
  //   room_ids массивом, group_ids массивом, поле topic вместо note.
  const lessonPayload = {
    branch_id: CFG.BRANCH_ID,
    room_ids: [CFG.ROOM_ID],
    lesson_type_id: CFG.LESSON_TYPE,
    subject_id: 92,                  // «Детали»
    teacher_ids: [],
    group_ids: [groupId],
    lesson_date: dateStrRu,          // 'dd.MM.yyyy'
    duration: durationMin,
    time_from: startStr,
    time_to:   endStr,
    topic: String(ev.name || ''),
    status: 1,
  };
  Logger.log('[alfaCreateGroupAndLesson_] lessonPayload=' + JSON.stringify(lessonPayload));
  let lessonId = null;
  try {
    const lessonResp = alfaCall_(token, CFG.BRANCH_ID,
      '/v2api/' + CFG.BRANCH_ID + '/lesson/create', lessonPayload);
    lessonId = extractCreatedId_(lessonResp);
  } catch(le) {
    // Группа создана, урок — нет
    if (sheetRow) {
      eventsSh.getRange(sheetRow, 19).setValue('Ошибка');
      eventsSh.getRange(sheetRow, 20).setValue(groupId);
      eventsSh.getRange(sheetRow, 22).setValue(new Date());
      eventsSh.getRange(sheetRow, 23).setValue('Группа создана, урок упал: ' + (le.message || le));
    }
    throw new Error('урок: ' + (le.message || le));
  }

  // Записываем результат в лист «События»
  const formUrl = CFG.ALFA_HOST + '/common/' + CFG.BRANCH_ID + '/lead/create?gid=' + groupId;
  if (sheetRow) {
    eventsSh.getRange(sheetRow, 19).setValue('Создано');
    eventsSh.getRange(sheetRow, 20).setValue(groupId);
    eventsSh.getRange(sheetRow, 21).setValue(lessonId || '');
    eventsSh.getRange(sheetRow, 22).setValue(new Date());
    eventsSh.getRange(sheetRow, 23).setValue('');
    // Ссылка на форму записи в колонку Z (26 — Ссылка на пост — переиспользуем как ссылку на форму)
    eventsSh.getRange(sheetRow, 26).setValue(formUrl);
    SpreadsheetApp.flush();
  }
  return { groupId, lessonId, formUrl };
}

// ============================================================
// ПАТЧ: исправление обновления времени урока в Альфе
// ============================================================
// ПРОБЛЕМА: «Внести правки в Альфу» обновляет название и note (описание)
// группы, но НЕ обновляет время (time_from / time_to / lesson_date)
// и длительность урока.
//
// КОРЕНЬ:
//   1. Если у события нет ev.alfaLessonId — функция вообще не пробует
//      обновить урок (только группу).
//   2. Если обновление урока проваливается — try/catch глотает ошибку:
//      «Не критично — группа обновилась». Время в Альфе остаётся старым,
//      но фронту шлётся ok: true.
//
// КАК ПРИМЕНИТЬ:
// 1. Открой Code.gs в Apps Script
// 2. Ctrl+F → найди: function alfaUpdateGroupAndLesson_(token, ev, eventsSh, sheetRow)
// 3. Замени ВСЮ функцию (от function до закрывающей } перед следующей function)
//    на код ниже:

// ============================================================
// ЗАМЕНИ ВСЮ ФУНКЦИЮ alfaUpdateGroupAndLesson_ В Code.gs НА ЭТУ
// ============================================================
// 1. Открой Code.gs в Apps Script
// 2. Ctrl+F → найди: function alfaUpdateGroupAndLesson_(token, ev, eventsSh, sheetRow)
// 3. Выдели всю функцию ОТ слова `function` ДО закрывающей `}` ПЕРЕД следующей функцией
// 4. Удали и вставь то что ниже (от function до } включительно)
// 5. Ctrl+S → Manage deployments → ✏ → New version → Deploy

function alfaUpdateGroupAndLesson_(token, ev, eventsSh, sheetRow) {
  if (!ev.alfaGroupId) throw new Error('нет alfaGroupId — нечего обновлять');

  let tz = ''; try { tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone(); } catch(e){}
  if (!tz) tz = 'Europe/Moscow';
  const dateObjU = new Date(ev.date);
  const dateStr = Utilities.formatDate(dateObjU, tz, 'yyyy-MM-dd');
  // Альфа CRM (S20): даты в формате dd.MM.yyyy, и e_date должен быть СТРОГО больше b_date
  const dateStrRu = Utilities.formatDate(dateObjU, tz, 'dd.MM.yyyy');
  const dateNextU = new Date(dateObjU.getTime() + 24 * 60 * 60 * 1000);
  const dateStrRuNext = Utilities.formatDate(dateNextU, tz, 'dd.MM.yyyy');

  let alfaName = '';
  if (sheetRow) alfaName = String(eventsSh.getRange(sheetRow, 8).getValue() || '');
  if (!alfaName) {
    alfaName = (ev.format || '') + ' "' + (ev.name || '') + '"';
    if (ev.price) alfaName += ' - ' + ev.price + ' руб.';
  }
  const note = String(ev.alfaNote || ev.description || '').slice(0, 255);

  // 1) Обновляем группу
  const groupPayload = {
    id: ev.alfaGroupId,
    name: alfaName, note: note,
    branch_ids: [CFG.BRANCH_ID],
    b_date: dateStrRu, e_date: dateStrRuNext,
    limit: toInt_(ev.limit) || 12, is_public: 1,
    lesson_type_id: CFG.LESSON_TYPE,
  };
  Logger.log('[alfaUpdate group] → ' + JSON.stringify(groupPayload));
  const grpResp = alfaCall_(token, CFG.BRANCH_ID,
    '/v2api/' + CFG.BRANCH_ID + '/group/update?id=' + ev.alfaGroupId, groupPayload);
  Logger.log('[alfaUpdate group] ← ' + JSON.stringify(grpResp).slice(0, 400));

  // 2) Урок: ищем lessonId если его нет
  let lessonId = ev.alfaLessonId || null;
  if (sheetRow && !lessonId) {
    lessonId = eventsSh.getRange(sheetRow, 21).getValue() || null;
  }
  // Если lessonId всё ещё пуст — спросим у Альфы через lesson/index
  if (!lessonId) {
    try {
      const lessonsResp = alfaCall_(token, CFG.BRANCH_ID,
        '/v2api/' + CFG.BRANCH_ID + '/lesson/index',
        { group_id: ev.alfaGroupId, page: 0 });
      const items = (lessonsResp && lessonsResp.items) || [];
      if (items.length > 0 && items[0].id) {
        lessonId = items[0].id;
        Logger.log('[alfaUpdate] найден lessonId через lesson/index: ' + lessonId);
      } else {
        Logger.log('[alfaUpdate] урок не найден через lesson/index, items=' + JSON.stringify(items).slice(0, 200));
      }
    } catch(e) {
      Logger.log('[alfaUpdate] ошибка поиска lessonId: ' + (e.message || e));
    }
  }

  // 3) Обновляем урок (если нашли ID)
  if (lessonId) {
    const startStr = formatTime_(ev.timeStart);   // "15:00"
    const endStr = addHours_(startStr, ev.duration || 2.5);   // "17:30"
    const durationMin = Math.round((ev.duration || 2.5) * 60);
    // Альфа на вход требует HH:mm:ss (с секундами), хоть в model отдаёт без них
    const startStrFull = startStr + ':00';
    const endStrFull   = endStr   + ':00';
    const lessonPayload = {
      id: lessonId,
      branch_id: CFG.BRANCH_ID,
      room_ids: [CFG.ROOM_ID],
      lesson_type_id: CFG.LESSON_TYPE,
      subject_id: 92,
      teacher_ids: [],
      group_ids: [ev.alfaGroupId],
      lesson_date: dateStrRu,
      duration: durationMin,
      time_from: startStrFull,
      time_to:   endStrFull,
      topic: String(ev.name || ''),
      status: 1,
    };
    Logger.log('[alfaUpdate lesson] → ' + JSON.stringify(lessonPayload));
    const lessonResp = alfaCall_(token, CFG.BRANCH_ID,
      '/v2api/' + CFG.BRANCH_ID + '/lesson/update?id=' + lessonId, lessonPayload);
    Logger.log('[alfaUpdate lesson] ← ' + JSON.stringify(lessonResp).slice(0, 400));

    // Проверяем что Альфа реально обновила
    if (lessonResp && lessonResp.errors &&
        (Array.isArray(lessonResp.errors) ? lessonResp.errors.length > 0
                                          : Object.keys(lessonResp.errors).length > 0)) {
      throw new Error('Альфа отказалась обновлять урок: ' + JSON.stringify(lessonResp.errors).slice(0, 300));
    }
    // Сохраняем найденный lessonId обратно чтобы в следующий раз не искать
    ev.alfaLessonId = lessonId;
  } else {
    // Урока в /lesson/index нет — пробуем обновить РЕГУЛЯРНОЕ расписание группы.
    // Альфа хранит расписание одноразовых событий именно в /regular-lesson/index
    // (то что видно в карточке группы в блоке «Регулярные уроки»).
    Logger.log('[alfaUpdate] обычный урок не найден, пробуем regular-lesson…');
    let regularLessonId = null;
    try {
      // Тянем все regular-lesson (API игнорирует фильтр по group_id)
      // и ищем тот что привязан к нашей группе
      let rpage = 0;
      while (rpage < 30) {
        const rResp = alfaCall_(token, CFG.BRANCH_ID,
          '/v2api/' + CFG.BRANCH_ID + '/regular-lesson/index', { page: rpage });
        const items = (rResp && rResp.items) || [];
        if (items.length === 0) break;
        for (const rl of items) {
          if (rl.related_class === 'Group' && Number(rl.related_id) === Number(ev.alfaGroupId)) {
            regularLessonId = rl.id;
            Logger.log('[alfaUpdate] найдена regular-lesson id=' + regularLessonId + ' для группы ' + ev.alfaGroupId);
            break;
          }
        }
        if (regularLessonId) break;
        rpage++;
      }
    } catch(e) {
      Logger.log('[alfaUpdate] regular-lesson/index ошибка: ' + (e.message || e));
    }

    if (!regularLessonId) {
      throw new Error('У группы id=' + ev.alfaGroupId + ' не найден ни обычный урок, ни регулярное расписание в Альфе. ' +
                      'Время не обновится. Откройте группу в Альфе и добавьте расписание вручную.');
    }

    // Обновляем регулярное расписание
    const startStrR = formatTime_(ev.timeStart);    // "15:00"
    const endStrR   = addHours_(startStrR, ev.duration || 2.5);   // "17:30"
    // Для regular-lesson Альфа на /update ждёт time_from_v и time_to_v в формате "HH:mm"
    // (без секунд — формат тот же что приходит из /index).
    // Поле b_date — реальная дата урока (для одноразовых событий совпадает с e_date).
    const regularPayload = {
      id: regularLessonId,
      time_from_v: startStrR,
      time_to_v:   endStrR,
      b_date:      dateStr,         // 'yyyy-MM-dd'
      e_date:      dateStr,         // одноразовое = b_date
    };
    Logger.log('[alfaUpdate regular] → ' + JSON.stringify(regularPayload));
    try {
      const rUpdResp = alfaCall_(token, CFG.BRANCH_ID,
        '/v2api/' + CFG.BRANCH_ID + '/regular-lesson/update?id=' + regularLessonId, regularPayload);
      Logger.log('[alfaUpdate regular] ← ' + JSON.stringify(rUpdResp).slice(0, 400));
      if (rUpdResp && rUpdResp.errors &&
          (Array.isArray(rUpdResp.errors) ? rUpdResp.errors.length > 0
                                          : Object.keys(rUpdResp.errors).length > 0)) {
        throw new Error('Альфа отказалась обновлять расписание: ' + JSON.stringify(rUpdResp.errors).slice(0, 300));
      }
    } catch(e) {
      throw new Error('Не удалось обновить регулярное расписание группы ' + ev.alfaGroupId + ': ' + (e.message || e));
    }
  }

  if (sheetRow) {
    eventsSh.getRange(sheetRow, 19).setValue('Создано');
    eventsSh.getRange(sheetRow, 22).setValue(new Date());
    eventsSh.getRange(sheetRow, 23).setValue('Обновлено ' + new Date().toISOString().slice(0,16));
    SpreadsheetApp.flush();
  }
  return { groupId: ev.alfaGroupId, lessonId };
}

// 4. Ctrl+S
// 5. Manage deployments → ✏ → New version → Deploy
//
// После деплоя:
//   - Открой событие, поменяй время → жми «Внести правки в Альфу»
//   - Если урок найден и обновлён → время в Альфе изменится
//   - Если урок не найден → получишь конкретный тост с причиной
//   - В Apps Script → Executions можно посмотреть полные логи [alfaUpdate ...]

// ============================================================
// ПАТЧ: исправление архивации группы в Альфе
// ============================================================
// ПРОБЛЕМА: detailsArchiveInAlfa_ возвращает { ok: true } даже когда Альфа
// фактически не архивировала группу. Функция alfaArchiveGroup_ не проверяет
// ответ Альфы — Альфа может вернуть 200 с body { errors: [...] } или
// { errno: 1 }, и это пройдёт как «успех».
//
// КАК ПРИМЕНИТЬ:
// 1. Открой Code.gs в Apps Script
// 2. Ctrl+F → найди: function alfaArchiveGroup_(token, groupId)
// 3. Заменить ВСЮ функцию (всё от `function alfaArchiveGroup_` до закрывающей `}`)
//    на код ниже:

function alfaArchiveGroup_(token, groupId) {
  // Альфа: customer-group/update с removed=1
  const payload = { id: groupId, removed: 1 };
  const path = '/v2api/' + CFG.BRANCH_ID + '/group/update?id=' + groupId;

  Logger.log('[alfaArchive] → POST ' + path + ' payload=' + JSON.stringify(payload));
  const resp = alfaCall_(token, CFG.BRANCH_ID, path, payload);
  Logger.log('[alfaArchive] ← response=' + JSON.stringify(resp).slice(0, 500));

  // Проверяем что Альфа реально архивировала. Возможные форматы ответа:
  //   успех:   { model: { id: 457, removed: 1, ... } }
  //   успех:   { id: 457 }
  //   ошибка:  { errors: { field: "msg" } }   или   { errors: [...] }
  //   ошибка:  { errno: 1, message: "..." }
  //   ошибка:  { code: 500, message: "..." }

  if (!resp || typeof resp !== 'object') {
    throw new Error('Пустой/невалидный ответ от Альфы: ' + JSON.stringify(resp).slice(0, 200));
  }

  // Явные ошибки
  if (resp.errors && (Array.isArray(resp.errors) ? resp.errors.length > 0 : Object.keys(resp.errors).length > 0)) {
    throw new Error('Альфа errors: ' + JSON.stringify(resp.errors).slice(0, 300));
  }
  if (resp.errno && resp.errno !== 0) {
    throw new Error('Альфа errno=' + resp.errno + ': ' + (resp.message || ''));
  }
  if (resp.code && resp.code >= 400) {
    throw new Error('Альфа code=' + resp.code + ': ' + (resp.message || ''));
  }

  // Проверяем что в модели реально removed=1
  const model = resp.model || resp;
  if (model && (model.removed !== undefined) && parseInt(model.removed, 10) !== 1) {
    throw new Error('Альфа приняла запрос но removed=' + model.removed + ' (не 1). model=' + JSON.stringify(model).slice(0, 300));
  }

  return resp;
}

// 4. Ctrl+S
// 5. Manage deployments → ✏ → New version → Deploy
//
// После деплоя — нажми «🗑 Удалить в Альфе» ещё раз:
//   - если Альфа реально что-то отвечает не так → увидишь тост с конкретной ошибкой
//   - если Альфа реально архивирует → группа исчезнет
//
// Дополнительно: после деплоя можно посмотреть Логи (Executions в Apps Script)
// и увидеть строки [alfaArchive] → POST ... и [alfaArchive] ← response= ...
// чтобы понять что именно Альфа возвращает.

// ============================================================
// DETAILS — ПУБЛИЧНЫЕ ЭНДПОЙНТЫ ДЛЯ view.html И team.html
// ============================================================

const PRIVATE_FORMATS = ['День рождения', 'Свидание', 'Частное мероприятие']; // отдаём как заглушку (без подробностей)

/**
 * Для view.html — только опубликованные, будущие.
 * Приватные форматы (ДР, частные, свидания) тоже показываем — но обрезаем поля,
 * чтобы клиент видел только что слот занят. Подробности — только публичные.
 * ВАЖНО: явный whitelist полей — references и teamComment клиентам НЕ отдаём.
 */
function detailsGetPublicEvents_(params) {
  const monthKey = params.month || '';   // 'YYYY-MM' или пусто = все будущие
  const events = collectAllPublishedEvents_();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const filtered = events.filter(ev => {
    const isPrivate = PRIVATE_FORMATS.indexOf(ev.format) >= 0;
    // Для приватных не требуем alfaGroupId и публикации клиентам — они важны как «слот занят»
    if (!isPrivate) {
      if (!ev.alfaGroupId) return false;                   // публичные — только опубликованные в Альфе
      if (ev.publishedToClients === false) return false;   // и с галочкой «Клиентам»
    }
    if (!ev.date) return false;
    const d = new Date(ev.date);
    if (d < today) return false;                           // только будущие
    if (monthKey && ev.date.slice(0, 7) !== monthKey) return false;
    return true;
  });

  // Минимизируем поля. Для приватных — отдаём минимум + флаг isPrivate.
  // Загружаем кэш превью один раз для всех событий (мы только сейчас узнаем
  // что у нас есть приватные/непубличные, но в кэше всё равно лежит общий map).
  const thumbsMapForPublic = thumbCacheLoadAll_();
  const out = filtered.map(ev => {
    const isPrivate = PRIVATE_FORMATS.indexOf(ev.format) >= 0;
    if (isPrivate) {
      return {
        id: ev.id,
        date: ev.date,
        timeStart: ev.timeStart,
        duration: ev.duration,
        isPrivate: true                       // фронт сам нарисует заглушку
        // НЕ отдаём: format, name, teacher, price, limit, description,
        // alfaGroupId, formUrl, references, teamComment
      };
    }
    return {
      id: ev.id, date: ev.date, timeStart: ev.timeStart, duration: ev.duration,
      format: ev.format, name: ev.name, teacher: ev.teacher,
      price: ev.price, limit: ev.limit, description: ev.description,
      alfaGroupId: ev.alfaGroupId,
      formUrl: CFG.ALFA_HOST + '/common/' + CFG.BRANCH_ID + '/lead/create?gid=' + ev.alfaGroupId,
      // Превью референсов — только картинки (thumb), без URL и note.
      // Используется для декоративной полосы под шапкой карточки на view.html.
      // Маркеры 'thumb:<hash>' резолвятся через кэш-лист SH_DET_THUMBS.
      references: resolveThumbs_(
        (Array.isArray(ev.references) ? ev.references : []).filter(r => r && r.thumb),
        thumbsMapForPublic
      ).map(r => ({ thumb: String(r.thumb || '') }))
    };
  });
  // Месяцы, по которым есть события (для пикера)
  const months = [...new Set(out.map(e => e.date.slice(0,7)))].sort();
  return { ok: true, events: out, months };
}

/**
 * Для team.html — все события из последних опубликованных версий и черновиков,
 * включая будущие/прошлые, со статусами и описаниями. Без записи.
 * Маркетологам отдаём всё, включая references и teamComment.
 */
function detailsGetTeamEvents_(params) {
  const monthKey = params.month || '';
  const events = collectAllPublishedEvents_();
  // Также подцепляем события из самых свежих черновиков (по одному на месяц)
  const draftEvents = collectLatestDraftsEvents_();

  // Объединяем по eventId, у опубликованных приоритет
  const byId = {};
  draftEvents.forEach(ev => { byId[ev.id] = { ...ev, _source: 'draft' }; });
  events.forEach(ev => { byId[ev.id] = { ...ev, _source: 'published' }; });

  let arr = Object.values(byId);
  if (monthKey) arr = arr.filter(e => e.date && e.date.slice(0,7) === monthKey);

  // Фильтруем мусор
  arr = arr.filter(e => e.date && (e.format || e.name));

  // Флаг публикации промо/команде. Обратная совместимость: отсутствие = true.
  // Скрываем архивированные события — они уже удалены из Альфы
  arr = arr.filter(e => e.status !== 'archived');
  arr = arr.filter(e => e.publishedToTeam !== false);

  arr.sort((a, b) => (a.date + (a.timeStart || '')).localeCompare(b.date + (b.timeStart || '')));

  // Резолвим thumb:<hash> маркеры на реальные data: URLs (одной пачкой для всех событий)
  const tmap = thumbCacheLoadAll_();
  arr.forEach(ev => {
    if (Array.isArray(ev.references)) {
      ev.references = resolveThumbs_(ev.references, tmap);
    }
  });

  const months = [...new Set(arr.map(e => e.date.slice(0,7)))].sort();
  return { ok: true, events: arr, months };
}

function collectAllPublishedEvents_() {
  // Раньше эта функция сливала события из ВСЕХ опубликованных снимков и брала самый свежий
  // по каждому id. После rebuildFromAlfa у части событий поменялся id (стал 'alfa_<gid>'),
  // и они НЕ перезатирали старые с id 'e_xxx' — отсюда дубли в view.html и team.html.
  //
  // Теперь читаем ТОЛЬКО snapshot_main — это «текущая версия мира» для публикации.
  // Остальные снимки в листе остаются как история (вдруг понадобится откатиться),
  // но в публичную выдачу они не попадают. Один источник правды.
  const sh = getSheet_(SH_DET_PUBLISHED);
  const row = findRowById_(sh, 'snapshot_main');
  if (row < 0) return [];
  const cells = sh.getRange(row, 1, 1, 5).getValues()[0];
  const json = String(cells[4] || '');
  if (!json) return [];
  let st;
  try { st = JSON.parse(json); } catch(e) { return []; }
  const updated = cells[3] ? String(cells[3]) : '';
  const events = (st.events || []).map(ev => ({ ...ev, _publishedAt: updated }));
  return events;
}

function collectLatestDraftsEvents_() {
  const sh = getSheet_(SH_DET_DRAFTS);
  const data = sh.getDataRange().getValues();
  // По одному самому свежему черновику на monthKey
  const latestPerMonth = {};
  // Отдельно — рабочий черновик 'working_main' (у него monthKey может быть пустой)
  let workingMainState = null;
  let workingMainUpdated = '';

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[3]) continue;
    let st;
    try { st = JSON.parse(String(row[3])); } catch(e) { continue; }
    const id = String(row[0] || '');
    const updated = row[2] ? String(row[2]) : '';

    // Рабочий черновик читаем всегда — там 39 событий со всех месяцев
    if (id === 'working_main') {
      if (!workingMainState || workingMainUpdated < updated) {
        workingMainState = st;
        workingMainUpdated = updated;
      }
      continue;
    }

    const mk = String(st.monthKey || '');
    if (!mk) continue;
    if (!latestPerMonth[mk] || latestPerMonth[mk]._updated < updated) {
      latestPerMonth[mk] = { ...st, _updated: updated };
    }
  }
  const out = [];
  Object.values(latestPerMonth).forEach(st => {
    (st.events || []).forEach(ev => out.push(ev));
  });
  // Добавляем события из working_main
  if (workingMainState && Array.isArray(workingMainState.events)) {
    workingMainState.events.forEach(ev => out.push(ev));
  }
  return out;
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
// ============================================================

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

  if (typeof alfaLogin_ !== 'function') {
    return {
      ok: false,
      error: 'Файл Alfa.gs не найден. Сохраните его в проекте Apps Script рядом с Code.gs.'
    };
  }

  let token;
  try {
    token = alfaLogin_();
    if (!token) throw new Error('Не получили токен');
  } catch(e) {
    return { ok: false, error: 'Ошибка логина в Альфу: ' + (e.message || e) };
  }

  const eventsSh = getSheet_(SH_DET_EVENTS);
  const lastRow = eventsSh.getLastRow();
  const idColIdx = 28;
  const rowByEventId = {};
  if (lastRow >= EVENTS_DATA_START_ROW) {
    const idValues = eventsSh.getRange(EVENTS_DATA_START_ROW, idColIdx, lastRow - EVENTS_DATA_START_ROW + 1, 1).getValues();
    idValues.forEach((row, i) => {
      const v = String(row[0] || '');
      const m = v.match(/\[gid:([\w-]+)\]/);
      if (m) rowByEventId[m[1]] = EVENTS_DATA_START_ROW + i;
    });
  }

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
      const existingGroupId = eventsSh.getRange(sheetRow, 20).getValue();
      const existingStatus  = eventsSh.getRange(sheetRow, 19).getValue();
      if (existingGroupId) {
        skipped++;
        results.push({ id: ev.id, name: ev.name, skipped: true, groupId: existingGroupId });
        return;
      }
      if (String(existingStatus).trim() === 'Отменено') {
        skipped++;
        return;
      }

      const dateObj = new Date(ev.date);
      let tz = '';
      try { tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(); } catch(e) {}
      if (!tz) tz = 'Europe/Moscow';
      const dateStr = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
      // Альфа CRM (S20): даты в формате dd.MM.yyyy, и e_date должен быть СТРОГО больше b_date
      const dateStrRu = Utilities.formatDate(dateObj, tz, 'dd.MM.yyyy');
      const dateNext = new Date(dateObj.getTime() + 24 * 60 * 60 * 1000);
      const dateStrRuNext = Utilities.formatDate(dateNext, tz, 'dd.MM.yyyy');

      let alfaName = String(eventsSh.getRange(sheetRow, 8).getValue() || '');
      if (!alfaName) {
        alfaName = (ev.format || '') + ' "' + (ev.name || '') + '"';
        if (ev.price) alfaName += ' - ' + ev.price + ' руб.';
      }
      const note = String(ev.alfaNote || ev.description || '').slice(0, 255);

      eventsSh.getRange(sheetRow, 19).setValue('Отправляется');
      SpreadsheetApp.flush();

      const groupPayload = {
        name:           alfaName,
        note:           note,
        branch_ids:     [CFG.BRANCH_ID],
        b_date:         dateStrRu,
        e_date:         dateStrRuNext,
        limit:          toInt_(ev.limit) || 12,
        is_public:      1,
        lesson_type_id: CFG.LESSON_TYPE,
      };
      const grpResp = alfaCall_(token, CFG.BRANCH_ID,
        '/v2api/' + CFG.BRANCH_ID + '/group/create', groupPayload);
      const groupId = extractCreatedId_(grpResp);
      if (!groupId) throw new Error('Не получили id группы');

      const startStr = formatTime_(ev.timeStart);
      const endStr = addHours_(startStr, ev.duration || 2.5);
      const durationMin = Math.round((ev.duration || 2.5) * 60);
      const lessonPayload = {
        branch_id:        CFG.BRANCH_ID,
        room_ids:         [CFG.ROOM_ID],
        lesson_type_id:   CFG.LESSON_TYPE,
        subject_id:       92,
        teacher_ids:      [],
        group_ids:        [groupId],
        lesson_date:      dateStrRu,
        duration:         durationMin,
        time_from:        startStr,
        time_to:          endStr,
        topic:            String(ev.name || ''),
        status:           1,
      };
      let lessonId = null;
      try {
        const lessonResp = alfaCall_(token, CFG.BRANCH_ID,
          '/v2api/' + CFG.BRANCH_ID + '/lesson/create', lessonPayload);
        lessonId = extractCreatedId_(lessonResp);
      } catch(le) {
        eventsSh.getRange(sheetRow, 19).setValue('Ошибка');
        eventsSh.getRange(sheetRow, 20).setValue(groupId);
        eventsSh.getRange(sheetRow, 22).setValue(new Date());
        eventsSh.getRange(sheetRow, 23).setValue('Группа создана (id=' + groupId + '), урок не создан: ' + (le.message || le));
        errors++;
        results.push({ id: ev.id, name: ev.name, error: 'урок: ' + (le.message || le), groupId });
        return;
      }

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
      try { stateObj = JSON.parse(_joinDraftChunks_(data[i])); } catch(e) {}
      // Резолвим thumb:<hash> маркеры в реальные data: URLs
      try {
        if (stateObj && Array.isArray(stateObj.events)) {
          const tmap = thumbCacheLoadAll_();
          stateObj.events.forEach(ev => {
            if (Array.isArray(ev.references)) {
              ev.references = resolveThumbs_(ev.references, tmap);
            }
          });
        }
      } catch(e) {}
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

function detailsPingEvents_() {
  try {
    const sh = getSheet_(SH_DET_EVENTS);
    return { ok: true, lastRow: sh.getLastRow(), name: sh.getName() };
  } catch(e) {
    return { ok: false, error: String(e) };
  }
}

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
 *   – находит существующие строки по eventId — обновляет их
 *   – новые события добавляет в конец таблицы
 *
 * Колонки:
 *   A=Дата, C=Время, D=Длит., F=Формат, G=Название, I=Педагог,
 *   J=Цена, K=Лимит, M=Расход на МК, Q=Описание, R=Готово к публикации, S=Статус,
 *   AB=Комментарий (eventId конструктора),
 *   AC=Референсы (URL, по одному на строку),
 *   AD=Комментарий команде.
 */
function syncEventsToSheet_(stateObj) {
  const sh = getSheet_(SH_DET_EVENTS);
  if (!sh) throw new Error('Лист «События» не найден');
  const events = (stateObj.events || []).filter(e => e.date);
  if (events.length === 0) return { ok: true, written: 0 };

  // Заголовки для AC/AD — ставим, если их там ещё нет
  try {
    const headerRow = 1;
    const acHeader = String(sh.getRange(headerRow, 29).getValue() || '').trim();
    const adHeader = String(sh.getRange(headerRow, 30).getValue() || '').trim();
    if (!acHeader) sh.getRange(headerRow, 29).setValue('Референсы');
    if (!adHeader) sh.getRange(headerRow, 30).setValue('Комментарий команде');
  } catch(e) { /* не критично */ }

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

  let tz = '';
  try { tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone(); } catch(e) {}
  if (!tz) tz = 'Europe/Moscow';
  let written = 0, updated = 0, created = 0;

  events.forEach(ev => {
    const eventId = ev.id || ('e_' + Math.random().toString(36).slice(2, 10));
    let rowNum = existing[eventId];

    if (!rowNum) {
      rowNum = sh.getLastRow() + 1;
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
    sh.getRange(rowNum, 1).setValue(dateObj);
    sh.getRange(rowNum, 1).setNumberFormat('dd.mm.yyyy');

    if (ev.timeStart) {
      sh.getRange(rowNum, 3).setValue(ev.timeStart);
      sh.getRange(rowNum, 3).setNumberFormat('hh:mm');
    }
    sh.getRange(rowNum, 4).setValue(ev.duration || 2.5);

    sh.getRange(rowNum, 6).setValue(ev.format || '');
    sh.getRange(rowNum, 7).setValue(ev.name || '');

    sh.getRange(rowNum, 9).setValue(ev.teacher || '');
    sh.getRange(rowNum, 10).setValue(ev.price != null ? ev.price : '');
    sh.getRange(rowNum, 11).setValue(ev.limit != null ? ev.limit : '');
    sh.getRange(rowNum, 13).setValue(ev.costPerPerson != null ? ev.costPerPerson : '');
    sh.getRange(rowNum, 17).setValue(ev.description || '');

    if (ev.publishToAlfa && !sh.getRange(rowNum, 18).getValue()) {
      sh.getRange(rowNum, 18).setValue('Да');
    }

    // AB — Комментарий с eventId
    let comment = ev.comment || '';
    if (!/\[gid:/.test(comment)) {
      comment = (comment ? comment + ' ' : '') + '[gid:' + eventId + ']';
    }
    sh.getRange(rowNum, 28).setValue(comment);

    // AC — Референсы (по одному URL на строку, плюс note в скобках)
    const refsArr = Array.isArray(ev.references) ? ev.references : [];
    const refsText = refsArr.map(r => {
      const url = String((r && r.url) || '').trim();
      const note = String((r && r.note) || '').trim();
      if (!url) return '';
      return note ? (url + ' — ' + note) : url;
    }).filter(Boolean).join('\n');
    sh.getRange(rowNum, 29).setValue(refsText);
    if (refsText) sh.getRange(rowNum, 29).setWrap(true);

    // AD — Комментарий команде
    sh.getRange(rowNum, 30).setValue(String(ev.teamComment || ''));
    if (ev.teamComment) sh.getRange(rowNum, 30).setWrap(true);

    written++;
  });

  return { ok: true, written, created, updated };
}

// ============================================================
// FETCH OG:IMAGE — превью для референсов
// Кэш 6 часов в Script Cache.
// ============================================================
/**
 * Перепроверяет все референсы во всех событиях текущего рабочего черновика
 * и догружает превью (thumb) для тех, у кого его нет или они битые.
 *
 * Когда запускать: после деплоя новой версии fetchOgImage, чтобы старые
 * ссылки получили нормальные превью.
 *
 * Запустить из редактора: выпадающий список → _refreshAllRefThumbs → ▶
 */
/**
 * Подтягивает из Альфы только timeStart и duration для всех связанных событий.
 * Безопасно — не трогает другие поля. Использует тот же бэк что pullFromAlfa.
 *
 * Запустить из редактора Apps Script:
 *   выпадающий список → _pullTimesFromAlfa → ▶
 *
 * После выполнения: события в working_main и snapshot_main получат правильное время.
 */
/**
 * Тестовая функция — пробует получить уроки конкретной группы напрямую.
 * Если работает — добавим точечную дозагрузку в _pullTimesFromAlfa.
 *
 * Запустить: выпадающий список → _testLessonsForGroup → ▶
 */
function _testLessonsForGroup() {
  const testGroupId = 334;  // Дракула — у тебя на скрине Альфы видно 10:00-12:00

  if (typeof alfaLogin_ !== 'function') { Logger.log('❌ Alfa.gs не найден'); return; }
  let token; try { token = alfaLogin_(); } catch(e) { Logger.log('❌ Логин: ' + e.message); return; }

  Logger.log('=== Тест 1: lesson/index с group_ids=[' + testGroupId + '] ===');
  try {
    const r1 = alfaCall_(token, CFG.BRANCH_ID,
      '/v2api/' + CFG.BRANCH_ID + '/lesson/index',
      { group_ids: [testGroupId], page: 0 });
    Logger.log('  Items: ' + (r1.items || []).length);
    Logger.log('  Total: ' + (r1.total || 0));
    if (r1.items && r1.items.length > 0) {
      Logger.log('  Первый урок: ' + JSON.stringify(r1.items[0]));
    }
  } catch(e) { Logger.log('  ❌ ' + e.message); }

  Logger.log('');
  Logger.log('=== Тест 2: lesson/index без даты, всё подряд ===');
  try {
    const r2 = alfaCall_(token, CFG.BRANCH_ID,
      '/v2api/' + CFG.BRANCH_ID + '/lesson/index',
      { page: 0 });
    Logger.log('  Items: ' + (r2.items || []).length);
    Logger.log('  Total: ' + (r2.total || 0));
    if (r2.items && r2.items.length > 0) {
      const found = r2.items.filter(l => (l.group_ids || []).indexOf(testGroupId) >= 0);
      Logger.log('  Уроков для группы ' + testGroupId + ': ' + found.length);
      if (found.length > 0) Logger.log('    ' + JSON.stringify(found[0]));
    }
  } catch(e) { Logger.log('  ❌ ' + e.message); }

  Logger.log('');
  Logger.log('=== Тест 3: расширенное окно — год назад до года вперёд ===');
  let tz = ''; try { tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone(); } catch(e){}
  if (!tz) tz = 'Europe/Moscow';
  const now = new Date();
  const dateFrom = Utilities.formatDate(new Date(now.getFullYear() - 1, 0, 1), tz, 'yyyy-MM-dd');
  const dateTo   = Utilities.formatDate(new Date(now.getFullYear() + 1, 11, 31), tz, 'yyyy-MM-dd');
  Logger.log('  Окно: ' + dateFrom + ' .. ' + dateTo);
  let total = 0, foundForGroup = null;
  for (let p = 0; p < 50; p++) {
    try {
      const r = alfaCall_(token, CFG.BRANCH_ID,
        '/v2api/' + CFG.BRANCH_ID + '/lesson/index',
        { date_from: dateFrom, date_to: dateTo, page: p });
      const items = r.items || [];
      if (items.length === 0) break;
      total += items.length;
      const f = items.find(l => (l.group_ids || []).indexOf(testGroupId) >= 0);
      if (f && !foundForGroup) foundForGroup = f;
    } catch(e) { Logger.log('  ❌ страница ' + p + ': ' + e.message); break; }
  }
  Logger.log('  Всего уроков просмотрено: ' + total);
  if (foundForGroup) {
    Logger.log('  ✓ Урок группы ' + testGroupId + ' НАЙДЕН: ' + JSON.stringify(foundForGroup));
  } else {
    Logger.log('  ❌ Урока группы ' + testGroupId + ' НЕТ НИГДЕ в /lesson/index');
    Logger.log('  Это значит: у группы в Альфе РЕАЛЬНО нет уроков, или они хранятся в другом endpoint');
  }
}

/**
 * Простая функция — устанавливает timeStart по формату события.
 * Дефолты подобраны под реальное расписание «Деталей».
 *
 * Запустить из редактора: выпадающий список → _setTimesByFormat → ▶
 *
 * Что делает:
 *   - Для каждого события БЕЗ alfaGroupId проставляет время по формату
 *   - Для событий С alfaGroupId — НЕ ТРОГАЕТ (там должна работать связка с уроком Альфы)
 *   - Логирует все изменения
 *   - Сохраняет в working_main + snapshot_main
 *
 * Если хочешь применить дефолты ВСЕМ (включая связанные с Альфой) —
 * раскомментируй строку с FORCE_ALL = true ниже.
 */
/**
 * Тест: ищем где Альфа хранит время расписания группы.
 * Перебираем известные endpoint'ы.
 */
/**
 * Подтягивает время и длительность из /regular-lesson/index Альфы
 * для всех событий с alfaGroupId. Это правильный источник — он содержит
 * расписание (планируемое время), в отличие от /lesson/index (фактически проведённые).
 *
 * Запустить: выпадающий список → _pullTimesFromRegular → ▶
 */
/**
 * Сбрасывает status с 'modified' на 'published' для всех событий с alfaGroupId.
 * Используется когда мы знаем что данные синхронизированы (например после
 * массовой подтяжки времени из Альфы — сами поля меняются, статус ставится modified,
 * но реально расхождений с Альфой уже нет).
 *
 * Запустить: выпадающий список → _clearModifiedStatuses → ▶
 */
function _clearModifiedStatuses() {
  const draftsSh = getSheet_(SH_DET_DRAFTS);
  const workRow = findRowById_(draftsSh, WORKING_DRAFT_ID);
  if (workRow < 0) { Logger.log('❌ Рабочий черновик не найден'); return; }
  const cells = draftsSh.getRange(workRow, 1, 1, draftsSh.getLastColumn()).getValues()[0];
  let workState;
  try { workState = JSON.parse(_joinDraftChunks_(cells)); }
  catch(e) { Logger.log('❌ Битый JSON: ' + e.message); return; }

  const events = workState.events || [];
  let cleared = 0;
  events.forEach(ev => {
    if (ev.status === 'modified' && ev.alfaGroupId) {
      ev.status = 'published';
      cleared++;
      Logger.log('  ✓ ' + ev.date + ' [' + ev.alfaGroupId + '] ' + (ev.name || ev.format || '?'));
    }
  });

  _writeWorkingDraft_(workState);
  try {
    detailsPushSnapshot_({ stateJson: JSON.stringify(workState) });
    Logger.log('✓ snapshot_main обновлён');
  } catch(e) { Logger.log('⚠ snapshot: ' + e.message); }

  Logger.log('=== ГОТОВО ===');
  Logger.log('Сброшено modified → published: ' + cleared);
  Logger.log('');
  Logger.log('Открой админку/team → Ctrl+Shift+R, статусы должны стать «В Альфе».');
}

function _pullTimesFromRegular() {
  if (typeof alfaLogin_ !== 'function') { Logger.log('❌ Alfa.gs не найден'); return; }
  let token; try { token = alfaLogin_(); } catch(e) { Logger.log('❌ Логин: ' + e.message); return; }

  // Тянем все regular-lesson
  Logger.log('Тяну /regular-lesson/index…');
  const byGroup = {};   // groupId → { time_from, time_to, duration_min }
  let totalRegulars = 0;
  for (let p = 0; p < 30; p++) {
    let r;
    try {
      r = alfaCall_(token, CFG.BRANCH_ID,
        '/v2api/' + CFG.BRANCH_ID + '/regular-lesson/index', { page: p });
    } catch(e) { Logger.log('  ❌ страница ' + p + ': ' + e.message); break; }
    const items = r.items || [];
    if (items.length === 0) break;
    totalRegulars += items.length;
    items.forEach(rl => {
      // Связь с группой может быть через related_class='Group' + related_id
      // или через group_ids[] — берём оба варианта
      const groupIds = [];
      if (rl.related_class === 'Group' && rl.related_id) groupIds.push(rl.related_id);
      if (Array.isArray(rl.group_ids)) rl.group_ids.forEach(g => groupIds.push(g));

      groupIds.forEach(gid => {
        // Берём первое попавшееся расписание для группы (если есть несколько)
        if (byGroup[gid]) return;
        let durationMin = null;
        if (rl.time_from_v && rl.time_to_v) {
          const [fh, fm] = rl.time_from_v.split(':').map(Number);
          const [th, tm] = rl.time_to_v.split(':').map(Number);
          if (!isNaN(fh) && !isNaN(th)) {
            durationMin = (th * 60 + (tm || 0)) - (fh * 60 + (fm || 0));
            if (durationMin < 0) durationMin += 24 * 60;  // через полночь
          }
        }
        byGroup[gid] = {
          time_from: rl.time_from_v || '',
          time_to: rl.time_to_v || '',
          duration_min: durationMin,
          b_date: rl.b_date || ''
        };
      });
    });
    if (items.length < 50) break;
  }
  Logger.log('Всего регулярок: ' + totalRegulars + '. Связано с группами: ' + Object.keys(byGroup).length);

  // Применяем к нашим событиям
  const draftsSh = getSheet_(SH_DET_DRAFTS);
  const workRow = findRowById_(draftsSh, WORKING_DRAFT_ID);
  if (workRow < 0) { Logger.log('❌ Рабочий черновик не найден'); return; }
  const cells = draftsSh.getRange(workRow, 1, 1, draftsSh.getLastColumn()).getValues()[0];
  let workState;
  try { workState = JSON.parse(_joinDraftChunks_(cells)); }
  catch(e) { Logger.log('❌ Битый JSON: ' + e.message); return; }

  const events = workState.events || [];
  let updated = 0, alreadyOk = 0, noRegular = 0, noAlfaGroup = 0;
  events.forEach(ev => {
    if (ev.status === 'archived') return;
    if (!ev.alfaGroupId) { noAlfaGroup++; return; }
    const reg = byGroup[ev.alfaGroupId];
    if (!reg || !reg.time_from) { noRegular++; return; }
    const newTime = reg.time_from;
    const newDur = reg.duration_min ? +(reg.duration_min / 60).toFixed(2) : ev.duration;
    const timeChanged = newTime !== ev.timeStart;
    const durChanged = Math.abs((ev.duration || 0) - newDur) > 0.01;
    if (timeChanged || durChanged) {
      Logger.log('  ' + ev.date + ' [' + ev.alfaGroupId + '] ' +
        (ev.name || ev.format || '?') + ': ' +
        (ev.timeStart || '—') + ' ' + (ev.duration || '?') + 'ч → ' +
        newTime + ' ' + newDur + 'ч');
      ev.timeStart = newTime;
      ev.duration = newDur;
      updated++;
    } else {
      alreadyOk++;
    }
  });

  _writeWorkingDraft_(workState);
  try {
    detailsPushSnapshot_({ stateJson: JSON.stringify(workState) });
    Logger.log('✓ snapshot_main обновлён');
  } catch(e) { Logger.log('⚠ snapshot: ' + e.message); }

  Logger.log('=== ГОТОВО ===');
  Logger.log('Связанных событий: ' + (events.length - noAlfaGroup));
  Logger.log('Обновлено: ' + updated);
  Logger.log('Уже было правильное: ' + alreadyOk);
  Logger.log('Нет расписания в Альфе: ' + noRegular);
  Logger.log('Без alfaGroupId (пропущено): ' + noAlfaGroup);
  Logger.log('');
  Logger.log('Открой админку/team/view → Ctrl+Shift+R');
}

function _testRegularIndex() {
  if (typeof alfaLogin_ !== 'function') { Logger.log('❌ Alfa.gs не найден'); return; }
  let token; try { token = alfaLogin_(); } catch(e) { Logger.log('❌ ' + e.message); return; }

  const endpoints = [
    '/v2api/' + CFG.BRANCH_ID + '/lesson-regular/index',
    '/v2api/' + CFG.BRANCH_ID + '/regular-lesson/index',
    '/v2api/' + CFG.BRANCH_ID + '/schedule/index',
    '/v2api/' + CFG.BRANCH_ID + '/lesson_regular/index',
    '/v2api/' + CFG.BRANCH_ID + '/group/show',
    '/v2api/' + CFG.BRANCH_ID + '/group/get',
    '/v2api/' + CFG.BRANCH_ID + '/timetable/index',
    '/v2api/' + CFG.BRANCH_ID + '/calendar/index'
  ];

  endpoints.forEach(ep => {
    Logger.log('=== ' + ep + ' ===');
    try {
      const r = alfaCall_(token, CFG.BRANCH_ID, ep, { page: 0 });
      Logger.log('  ✓ OK. Ключи: ' + Object.keys(r).join(', '));
      if (r.items && r.items.length > 0) {
        Logger.log('  Items: ' + r.items.length);
        Logger.log('  Пример: ' + JSON.stringify(r.items[0]).slice(0, 400));
      } else if (r.total != null) {
        Logger.log('  total=' + r.total + ', items пусто');
      } else {
        // Это может быть один объект, не список
        Logger.log('  Объект: ' + JSON.stringify(r).slice(0, 400));
      }
    } catch(e) {
      const msg = String(e.message || e);
      // Сокращаем длинный HTML 404
      if (msg.indexOf('404') >= 0) Logger.log('  ❌ 404');
      else if (msg.indexOf('400') >= 0) Logger.log('  ⚠ 400 (нужны параметры): ' + msg.slice(0, 100));
      else Logger.log('  ❌ ' + msg.slice(0, 200));
    }
  });

  Logger.log('');
  Logger.log('=== Тест: group/show?id=334 (как параметр) ===');
  try {
    const r = alfaCall_(token, CFG.BRANCH_ID,
      '/v2api/' + CFG.BRANCH_ID + '/group/show', { id: 334 });
    Logger.log('  ✓ ' + JSON.stringify(r).slice(0, 500));
  } catch(e) { Logger.log('  ❌ ' + String(e.message).slice(0, 200)); }
}

function _setTimesByFormat() {
  const FORCE_ALL = true;  // true = переписать всем, false = только без alfaGroupId

  // Дефолтное время по формату (можешь подправить под себя)
  const TIMES = {
    'Киновечер':                    { time: '19:00', duration: 2.5 },
    'Книжный клуб':                 { time: '13:00', duration: 2 },
    'Винная дегустация':            { time: '19:00', duration: 2 },
    'Вечеринка':                    { time: '19:00', duration: 4 },
    'Встреча с экспертом':          { time: '19:00', duration: 2 },
    'Мастер-класс по глине':        { time: '18:00', duration: 2.5 },
    'Мастер-класс по глине для детей': { time: '11:00', duration: 1.5 },
    'Мастер-класс по живописи':     { time: '18:00', duration: 2.5 },
    'Мастер-класс по макраме':      { time: '18:00', duration: 2.5 },
    'Мастер-класс по бьюти':        { time: '18:00', duration: 2 },
    'Мастер-класс по каллиграфии':  { time: '18:00', duration: 2 },
    'Мастер-класс кулинарный':      { time: '18:00', duration: 2.5 },
    'Мастер-класс':                 { time: '18:00', duration: 2.5 },
    'День рождения':                { time: '15:00', duration: 3 },
    'Свидание':                     { time: '19:00', duration: 3 },
    'Частное мероприятие':          { time: '15:00', duration: 3 }
  };
  const FALLBACK = { time: '18:00', duration: 2.5 };

  const draftsSh = getSheet_(SH_DET_DRAFTS);
  const workRow = findRowById_(draftsSh, WORKING_DRAFT_ID);
  if (workRow < 0) { Logger.log('❌ Рабочий черновик не найден'); return; }
  const cells = draftsSh.getRange(workRow, 1, 1, draftsSh.getLastColumn()).getValues()[0];
  let workState;
  try { workState = JSON.parse(_joinDraftChunks_(cells)); }
  catch(e) { Logger.log('❌ Битый JSON: ' + e.message); return; }

  const events = workState.events || [];
  let changed = 0, skippedAlfa = 0, skippedSame = 0, noFormat = 0;

  events.forEach(ev => {
    if (!ev.format) { noFormat++; return; }
    if (ev.status === 'archived') return;
    if (ev.alfaGroupId && !FORCE_ALL) { skippedAlfa++; return; }

    const def = TIMES[ev.format] || FALLBACK;
    const oldTime = ev.timeStart;
    const oldDur = ev.duration;
    if (oldTime === def.time && Math.abs((oldDur || 0) - def.duration) < 0.01) {
      skippedSame++;
      return;
    }
    Logger.log('  ' + ev.date + ' ' + (ev.format || '?') + ' / ' +
      (ev.name || '') + ': ' + (oldTime || '—') + ' ' + (oldDur || '?') + 'ч → ' +
      def.time + ' ' + def.duration + 'ч');
    ev.timeStart = def.time;
    ev.duration = def.duration;
    changed++;
  });

  _writeWorkingDraft_(workState);
  try {
    detailsPushSnapshot_({ stateJson: JSON.stringify(workState) });
    Logger.log('✓ snapshot_main обновлён');
  } catch(e) { Logger.log('⚠ snapshot_main: ' + e.message); }

  Logger.log('=== ГОТОВО ===');
  Logger.log('Всего событий: ' + events.length);
  Logger.log('Изменено: ' + changed);
  Logger.log('Без формата (пропущено): ' + noFormat);
  Logger.log('Уже было правильное время: ' + skippedSame);
  if (!FORCE_ALL) Logger.log('Связаны с Альфой (пропущено): ' + skippedAlfa);
  Logger.log('');
  Logger.log('Открой админку/team/view → Ctrl+Shift+R, время должно быть по форматам.');
}

function _pullTimesFromAlfa() {
  const draftsSh = getSheet_(SH_DET_DRAFTS);
  const workRow = findRowById_(draftsSh, WORKING_DRAFT_ID);
  if (workRow < 0) {
    Logger.log('❌ Рабочий черновик не найден');
    return;
  }
  const cells = draftsSh.getRange(workRow, 1, 1, draftsSh.getLastColumn()).getValues()[0];
  let workState;
  try { workState = JSON.parse(_joinDraftChunks_(cells)); }
  catch(e) { Logger.log('❌ Битый JSON: ' + e.message); return; }

  const events = workState.events || [];
  // Берём ID групп Альфы из связанных событий
  const groupIds = events
    .filter(e => e.alfaGroupId && e.status !== 'archived')
    .map(e => e.alfaGroupId);

  if (groupIds.length === 0) {
    Logger.log('Нет событий, связанных с Альфой');
    return;
  }

  Logger.log('Запрашиваю данные ' + groupIds.length + ' групп из Альфы…');

  // Используем существующий pullFromAlfa
  const r = detailsPullFromAlfa_({ groupIds });
  if (!r.ok) {
    Logger.log('❌ Ошибка: ' + r.error);
    return;
  }

  let updatedTime = 0, updatedDuration = 0, notFound = 0, alreadyOk = 0, noTimeInAlfa = 0;
  const sampleNoTime = [];
  events.forEach(ev => {
    if (!ev.alfaGroupId || ev.status === 'archived') return;
    const a = r.items[ev.alfaGroupId];
    if (!a || !a.found) {
      notFound++;
      return;
    }
    // Время начала
    if (!a.time_from) {
      // Альфа не отдала time_from — у группы нет уроков в окне ±6 мес
      noTimeInAlfa++;
      if (sampleNoTime.length < 5) {
        sampleNoTime.push({
          gid: ev.alfaGroupId, date: ev.date,
          name: ev.name || ev.format,
          alfa_keys: Object.keys(a).join(','),
          a_full: JSON.stringify(a)
        });
      }
      return;
    }
    if (a.time_from !== ev.timeStart) {
      Logger.log('  ' + ev.date + ' [' + ev.alfaGroupId + '] ' +
        (ev.name || ev.format || '?') + ': ' + (ev.timeStart || '—') + ' → ' + a.time_from);
      ev.timeStart = a.time_from;
      updatedTime++;
    } else {
      alreadyOk++;
    }
    // Длительность (минуты → часы)
    if (a.duration_min != null) {
      const newDur = +(a.duration_min / 60).toFixed(2);
      if (Math.abs((ev.duration || 0) - newDur) > 0.01) {
        ev.duration = newDur;
        updatedDuration++;
      }
    }
  });

  // Сохраняем working_main
  _writeWorkingDraft_(workState);

  // Также обновляем snapshot_main, чтобы view.html и team.html сразу подхватили
  try {
    detailsPushSnapshot_({ stateJson: JSON.stringify(workState) });
    Logger.log('✓ snapshot_main обновлён');
  } catch(e) {
    Logger.log('⚠ Не удалось обновить snapshot_main: ' + e.message);
  }

  Logger.log('=== ГОТОВО ===');
  Logger.log('Связанных событий: ' + groupIds.length);
  Logger.log('Время обновлено: ' + updatedTime);
  Logger.log('Длительность обновлена: ' + updatedDuration);
  Logger.log('Не найдено в Альфе (групп нет): ' + notFound);
  Logger.log('Альфа не отдала time_from (нет урока в окне ±6мес): ' + noTimeInAlfa);
  Logger.log('Уже было правильное: ' + alreadyOk);
  if (sampleNoTime.length > 0) {
    Logger.log('');
    Logger.log('ПРИМЕРЫ событий БЕЗ time_from в Альфе:');
    sampleNoTime.forEach(s => {
      Logger.log('  [' + s.gid + '] ' + s.date + ' ' + s.name);
      Logger.log('    Альфа отдала: ' + s.a_full);
    });
  }
  Logger.log('');
  Logger.log('Открой админку/team/view → Ctrl+Shift+R, время должно стать корректным.');
}

function _refreshAllRefThumbs() {
  const draft = _readWorkingDraft_();
  if (!draft) {
    Logger.log('❌ Рабочий черновик не найден');
    return;
  }
  const workState = draft.state;

  const events = workState.events || [];
  let totalRefs = 0, refreshed = 0, failed = 0, skipped = 0;

  // Сбрасываем кэш по всем URL (чтобы повторный вызов точно сходил на сервер)
  const cache = CacheService.getScriptCache();

  events.forEach(ev => {
    if (!Array.isArray(ev.references)) return;
    ev.references.forEach(ref => {
      if (!ref || !ref.url) return;
      totalRefs++;
      // Пропускаем те, у кого уже есть короткий маркер 'thumb:<hash>' или data: URL
      if (ref.thumb && (ref.thumb.indexOf('thumb:') === 0 || ref.thumb.indexOf('data:') === 0)) {
        skipped++;
        return;
      }
      // Сбрасываем кэш для этого URL
      try {
        const cacheKey = 'og_' + Utilities.base64EncodeWebSafe(ref.url).slice(0, 200);
        cache.remove(cacheKey);
      } catch(e) {}
      // Заново вызываем fetchOgImage
      try {
        const r = detailsFetchOgImage_({ url: ref.url });
        if (r && r.ok && r.thumb && r.thumb !== ref.thumb) {
          ref.thumb = r.thumb;
          ref.domain = r.domain || ref.domain;
          // Успех = либо короткий маркер 'thumb:<hash>' (картинка в кэш-листе),
          // либо data: URL (на случай если кто-то починил без маркера).
          const isInlined = r.thumb.indexOf('thumb:') === 0 || r.thumb.indexOf('data:') === 0;
          if (isInlined) {
            refreshed++;
            Logger.log('  ✓ ' + ref.url.slice(0, 60) + ' → ' + (r.thumb.indexOf('thumb:') === 0 ? r.thumb : 'inlined'));
          } else {
            failed++;
            Logger.log('  ⚠ ' + ref.url.slice(0, 60) + ' → URL (превью не работает: hotlink или > 30KB)');
          }
        } else {
          failed++;
        }
      } catch(e) {
        failed++;
        Logger.log('  ❌ ' + ref.url.slice(0, 60) + ' — ' + e.message);
      }
    });
  });

  // Сохраняем обратно через универсальный хелпер (с разбиением на куски)
  _writeWorkingDraft_(workState);
  const finalJsonLen = JSON.stringify(workState).length;

  Logger.log('=== ГОТОВО ===');
  Logger.log('Всего референсов: ' + totalRefs);
  Logger.log('Уже были инлайн (пропущено): ' + skipped);
  Logger.log('Свежих превью загружено: ' + refreshed);
  Logger.log('Не удалось: ' + failed);
  Logger.log('Размер JSON: ' + finalJsonLen + ' символов');
  Logger.log('');
  Logger.log('Открой админку → Ctrl+Shift+R, превью должны появиться.');
}

/**
 * Тестовая функция для отладки превью референсов.
 * Запусти из редактора Apps Script (выпадающий список → _testFetchOgImage → ▶).
 * После запуска — внизу в журнале появятся строки [fetchOgImage] с диагностикой.
 *
 * Чтобы проверить свою ссылку — замени URL на любую с Pinterest (правый клик
 * на пин → Скопировать ссылку) и запусти ещё раз.
 */
function _testFetchOgImage() {
  // Пример Pinterest ссылки. Замени на любую свою.
  const testUrl = 'https://ru.pinterest.com/pin/345088685816470/';
  Logger.log('=== Тест fetchOgImage для: ' + testUrl);
  // Сбрасываем кэш чтобы тест точно сходил в Pinterest
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'og_' + Utilities.base64EncodeWebSafe(testUrl).slice(0, 200);
    cache.remove(cacheKey);
    Logger.log('Кэш очищен');
  } catch(e) {}

  const r = detailsFetchOgImage_({ url: testUrl });
  Logger.log('---');
  Logger.log('Результат:');
  Logger.log('  ok: ' + r.ok);
  Logger.log('  domain: ' + r.domain);
  Logger.log('  title: ' + (r.title || '').slice(0, 80));
  if (r.thumb) {
    if (r.thumb.indexOf('data:') === 0) {
      Logger.log('  thumb: INLINED data URL длиной ' + r.thumb.length + ' символов ✓');
    } else {
      Logger.log('  thumb URL: ' + r.thumb);
      Logger.log('  ❌ Не инлайн — превью не покажется из-за hotlink-protection.');
      Logger.log('  Смотри строку [fetchOgImage] выше — там написано почему.');
    }
  } else {
    Logger.log('  thumb: пусто');
  }
}

function detailsFetchOgImage_(body) {
  const url = String(body.url || '').trim();
  if (!url) return { ok: false, error: 'нет url' };
  if (!/^https?:\/\//i.test(url)) return { ok: false, error: 'некорректный url' };

  // Кэш по url
  const cache = CacheService.getScriptCache();
  const cacheKey = 'og_' + Utilities.base64EncodeWebSafe(url).slice(0, 200);
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  // СПЕЦСЛУЧАЙ: если URL — это прямая ссылка на картинку (.jpg/.png/.webp/.gif),
  // не парсим HTML, а сразу скачиваем картинку и кодируем в data: URL.
  // Pinterest часто отдаёт прямые URL (i.pinimg.com/736x/.../*.jpg).
  const isDirectImage = /\.(jpe?g|png|webp|gif)(\?.*)?$/i.test(url);
  if (isDirectImage) {
    // Apps Script не имеет глобального URL — парсим домен вручную регуляркой
    let domain = '';
    const dm = url.match(/^https?:\/\/([^\/?#]+)/i);
    if (dm) domain = dm[1].toLowerCase();

    Logger.log('[fetchOgImage] domain распознан как: "' + domain + '"');

    // Для pinimg подменяем размер на /236x/ — превью обычно 10-30KB.
    // Pinterest имеет несколько форматов URL:
    //   /originals/, /1200x/, /736x/, /564x/, /474x/, /236x/
    //   /webp70/<size>/, /<size>/   (webp-вариант)
    // Универсальная подмена: любой <число>x/ → 236x/
    let fetchUrl = url;
    if (domain.indexOf('pinimg') >= 0) {
      fetchUrl = url
        .replace(/\/originals\//, '/236x/')
        .replace(/\/(\d{3,4}x)\//, '/236x/');
      Logger.log('[fetchOgImage] подмена → ' + fetchUrl);
    }

    Logger.log('[fetchOgImage] прямая картинка: ' + url + ' → fetch: ' + fetchUrl);

    let directThumb = url;
    try {
      const imgResp = UrlFetchApp.fetch(fetchUrl, {
        method: 'get', followRedirects: true, muteHttpExceptions: true,
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
          'Accept': 'image/webp,image/png,image/jpeg,image/*,*/*;q=0.8'
        }
      });
      const code = imgResp.getResponseCode();
      Logger.log('[fetchOgImage] код ответа: ' + code);
      if (code < 400) {
        const blob = imgResp.getBlob();
        const bytes = blob.getBytes();
        const ct = blob.getContentType() || 'image/jpeg';
        Logger.log('[fetchOgImage] размер: ' + bytes.length + ' байт, тип: ' + ct);
        // Жёсткий лимит 30KB на исходник (~40KB в base64).
        // Это ниже 50000-символьного лимита Google Sheets cell.
        // С учётом 50+ референсов в одном черновике — лимит должен быть строгим.
        if (bytes.length > 0 && bytes.length <= 30 * 1024) {
          const b64 = Utilities.base64Encode(bytes);
          const dataUrl = 'data:' + ct + ';base64,' + b64;
          // Кладём data URL в отдельный лист «Превью_кэш» по ключу хэшу URL.
          // В референсе сохраняем только короткий маркер 'thumb:<hash>',
          // который займёт ~50 символов вместо 30000+.
          const thumbKey = thumbHash_(fetchUrl);
          thumbCachePut_(thumbKey, dataUrl);
          directThumb = 'thumb:' + thumbKey;
          Logger.log('[fetchOgImage] ✓ inlined в кэш-лист, ключ: ' + thumbKey);
        } else {
          Logger.log('[fetchOgImage] ❌ слишком большая (' + bytes.length + 'B > 30KB), оставляем URL');
        }
      }
    } catch(e) {
      Logger.log('[fetchOgImage] ошибка скачивания: ' + e.message);
    }

    const result = { ok: true, thumb: directThumb, domain: domain, title: '' };
    try {
      const json = JSON.stringify(result);
      if (json.length < 95 * 1024) cache.put(cacheKey, json, 6 * 3600);
    } catch(e) {}
    return result;
  }

  const fallback = (() => {
    try {
      const u = new URL(url);
      return {
        ok: true,
        thumb: 'https://www.google.com/s2/favicons?domain=' + u.hostname + '&sz=128',
        domain: u.hostname,
        title: ''
      };
    } catch(e) {
      return { ok: true, thumb: '', domain: '', title: '' };
    }
  })();

  let html = '';
  try {
    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      followRedirects: true,
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; DetailsBot/1.0; +https://script.google.com)',
        'Accept': 'text/html,application/xhtml+xml',
        'Accept-Language': 'ru,en;q=0.9'
      },
      validateHttpsCertificates: true
    });
    const code = resp.getResponseCode();
    if (code < 200 || code >= 400) {
      cache.put(cacheKey, JSON.stringify(fallback), 6 * 3600);
      return fallback;
    }
    html = resp.getContentText();
    if (!html || html.length < 50) {
      cache.put(cacheKey, JSON.stringify(fallback), 6 * 3600);
      return fallback;
    }
  } catch(e) {
    cache.put(cacheKey, JSON.stringify(fallback), 6 * 3600);
    return fallback;
  }

  // Берём только <head> чтобы не парсить весь документ
  const headMatch = html.match(/<head[^>]*>([\s\S]*?)<\/head>/i);
  const head = headMatch ? headMatch[1] : html.slice(0, 50000);

  let img = matchMeta_(head, 'og:image:secure_url')
         || matchMeta_(head, 'og:image')
         || matchMeta_(head, 'twitter:image')
         || matchMeta_(head, 'twitter:image:src');

  // image_src через <link rel="image_src">
  if (!img) {
    const m = head.match(/<link[^>]+rel=["']?image_src["']?[^>]+href=["']([^"']+)["']/i);
    if (m) img = m[1];
  }

  // Заголовок страницы
  let title = '';
  const tm = head.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  if (tm) title = decodeHtmlEntities_(tm[1]).trim().slice(0, 200);
  if (!title) {
    title = matchMeta_(head, 'og:title') || matchMeta_(head, 'twitter:title') || '';
  }

  // Делаем абсолютный URL
  if (img && !/^https?:\/\//i.test(img)) {
    try {
      const u = new URL(url);
      if (img.startsWith('//')) img = u.protocol + img;
      else if (img.startsWith('/')) img = u.protocol + '//' + u.host + img;
      else img = u.protocol + '//' + u.host + '/' + img;
    } catch(e) { img = ''; }
  }

  let domain = '';
  try { domain = new URL(url).hostname; } catch(e) {}

  // Если картинка с домена с hotlink-protection (Pinterest, Instagram и т.п.) —
  // скачиваем её и кодируем в data: URL, чтобы превью точно работало везде
  // (Pinterest блокирует hotlink через Referer-проверку).
  let finalThumb = img || fallback.thumb;
  let inlineDebug = '';   // для логирования

  if (img) {
    try {
      const imgHost = new URL(img).hostname;
      const isPinimg = /pinimg\.com/i.test(imgHost);
      const needInline = /pinimg\.com|fbcdn\.net|cdninstagram|instagram\.com|fbsbx\.com/i.test(imgHost);

      if (needInline) {
        // Для pinimg подменяем размер на маленький — иначе картинка может быть 500KB+,
        // а у Apps Script лимит на cache 100KB и общий лимит на ответ ~50MB.
        // Pinterest хранит варианты в /originals/, /736x/, /564x/, /474x/, /236x/, /60x60/.
        // /236x/ — обычно 10-30KB, идеально для превью.
        let fetchUrl = img;
        if (isPinimg) {
          fetchUrl = img
            .replace(/\/originals\//, '/236x/')
            .replace(/\/736x\//, '/236x/')
            .replace(/\/564x\//, '/236x/')
            .replace(/\/474x\//, '/236x/');
        }

        const imgResp = UrlFetchApp.fetch(fetchUrl, {
          method: 'get',
          followRedirects: true,
          muteHttpExceptions: true,
          headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'image/webp,image/png,image/jpeg,image/*,*/*;q=0.8',
            'Accept-Language': 'ru,en;q=0.9'
          }
        });
        const code = imgResp.getResponseCode();
        if (code < 400) {
          const blob = imgResp.getBlob();
          const bytes = blob.getBytes();
          const len = bytes.length;
          const ct = blob.getContentType() || 'image/jpeg';
          inlineDebug = 'pinimg fetch: ' + code + ' size=' + len + ' ct=' + ct;
          // Лимит для data: URL — 90KB исходника (~120KB base64) — спокойно влезает в payload.
          if (len > 0 && len <= 90 * 1024) {
            const b64 = Utilities.base64Encode(bytes);
            finalThumb = 'data:' + ct + ';base64,' + b64;
            inlineDebug += ' → inlined';
          } else {
            inlineDebug += ' → too big, fallback to URL (превью НЕ будет видно из-за hotlink-блокировки)';
          }
        } else {
          inlineDebug = 'pinimg fetch HTTP ' + code + ' — будет битая ссылка';
        }
      }
    } catch(e) {
      inlineDebug = 'pinimg error: ' + e.message;
    }
  }

  // Логируем что получилось (увидишь в Apps Script → Журнал выполнения)
  Logger.log('[fetchOgImage] ' + url + ' | img=' + (img || '—').slice(0, 80) +
             ' | thumb=' + (finalThumb.startsWith('data:') ? 'INLINED ' + finalThumb.length + 'B' : finalThumb.slice(0, 80)) +
             (inlineDebug ? ' | ' + inlineDebug : ''));

  const result = {
    ok: true,
    thumb: finalThumb,
    domain: domain,
    title: title
  };
  // Кэш — только если data: URL не слишком большой (cache limit Apps Script ~100KB на ключ)
  try {
    const json = JSON.stringify(result);
    if (json.length < 95 * 1024) {
      cache.put(cacheKey, json, 6 * 3600);
    }
  } catch(e) {}
  return result;
}

function matchMeta_(html, prop) {
  if (!html) return '';
  // property="..." content="..."
  const re1 = new RegExp('<meta[^>]+(?:property|name)\\s*=\\s*["\']' + prop.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '["\'][^>]*content\\s*=\\s*["\']([^"\']+)["\']', 'i');
  const m1 = html.match(re1);
  if (m1) return decodeHtmlEntities_(m1[1].trim());
  // content="..." property="..."
  const re2 = new RegExp('<meta[^>]+content\\s*=\\s*["\']([^"\']+)["\'][^>]*(?:property|name)\\s*=\\s*["\']' + prop.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '["\']', 'i');
  const m2 = html.match(re2);
  if (m2) return decodeHtmlEntities_(m2[1].trim());
  return '';
}

function decodeHtmlEntities_(s) {
  if (!s) return '';
  return String(s)
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&#x2F;/g, '/')
    .replace(/&#(\d+);/g, (_, c) => String.fromCharCode(Number(c)));
}

// ============================================================
// КЭШ ПРЕВЬЮ ДЛЯ РЕФЕРЕНСОВ
// ============================================================
// data: URL картинок может быть 30-50KB — это пробивает 50000-символьный
// лимит ячейки Google Sheets, в которой хранится JSON черновика.
// Решение: храним картинки в отдельном листе SH_DET_THUMBS как
// (key, dataUrl, updated). В референсе хранится только короткий маркер
// 'thumb:<hash>', который занимает 50 символов вместо 30000+.
// При выдаче на фронт (admin/team/view) подменяем маркеры на реальные data: URLs.

function thumbHash_(s) {
  // Простой хэш для ключа в кэше — 12 символов, base64 от первых 9 байт
  // SHA-1 от строки. Достаточно уникален для нашего размера данных.
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, String(s));
  return Utilities.base64EncodeWebSafe(bytes).slice(0, 12).replace(/=/g, '');
}

function thumbCachePut_(key, dataUrl) {
  if (!key || !dataUrl) return;
  // Защита от пробития 50000 на одну ячейку в самом кэше
  if (dataUrl.length > 49000) return;
  const sh = getSheet_(SH_DET_THUMBS);
  // Гарантируем заголовок
  if (sh.getLastRow() === 0) {
    sh.appendRow(['key', 'dataUrl', 'updated']);
  }
  // Ищем существующую строку
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) {
      // Обновляем
      sh.getRange(i + 1, 2).setValue(dataUrl);
      sh.getRange(i + 1, 3).setValue(new Date().toISOString());
      return;
    }
  }
  sh.appendRow([key, dataUrl, new Date().toISOString()]);
}

function thumbCacheGet_(key) {
  if (!key) return '';
  const sh = getSheet_(SH_DET_THUMBS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) {
      return String(data[i][1] || '');
    }
  }
  return '';
}

/**
 * Загружает все превью из кэш-листа в map { key: dataUrl }.
 * Используется для пакетной подмены 'thumb:<key>' маркеров на реальные data: URL.
 */
function thumbCacheLoadAll_() {
  const sh = getSheet_(SH_DET_THUMBS);
  const data = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const k = String(data[i][0] || '').trim();
    const v = String(data[i][1] || '');
    if (k && v) map[k] = v;
  }
  return map;
}

/**
 * Резолвит маркеры 'thumb:<key>' в массиве референсов на реальные data: URL.
 * Применяется при выдаче событий на фронт (admin/team/view).
 * @param {Array} refs - массив { url, note, thumb, domain, ... }
 * @param {Object} thumbsMap - результат thumbCacheLoadAll_()
 * @return {Array} новый массив с подменёнными thumb
 */
function resolveThumbs_(refs, thumbsMap) {
  if (!Array.isArray(refs)) return [];
  if (!thumbsMap) thumbsMap = thumbCacheLoadAll_();
  return refs.map(r => {
    if (!r) return r;
    const t = String(r.thumb || '');
    if (t.indexOf('thumb:') === 0) {
      const key = t.slice(6);
      const dataUrl = thumbsMap[key] || '';
      return Object.assign({}, r, { thumb: dataUrl || r.url });
    }
    return r;
  });
}

// ============================================================
// СТАРЫЕ camp/intensive — БЕЗ ИЗМЕНЕНИЙ
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

  // Google Sheets имеет лимит 50000 символов на ячейку. Если JSON больше —
  // разбиваем его на куски по 49000 символов и пишем в соседние столбцы D, E, F, G…
  // При чтении (loadDraft_) склеиваем обратно через _joinDraftChunks_.
  const CHUNK = 49000;
  const chunks = [];
  for (let i = 0; i < stateJson.length; i += CHUNK) {
    chunks.push(stateJson.slice(i, i + CHUNK));
  }
  if (chunks.length === 0) chunks.push('');

  // Удаляем старую строку — это сразу избавляет от «хвостов» в правых ячейках,
  // которые могли остаться от прошлой большой версии.
  const existing = findRowById_(sh, id);
  if (existing > 0) sh.deleteRow(existing);

  // Строим row: [id, name, updated, chunk1, chunk2, …]
  const rowData = [id, name, nowIso_(), ...chunks];
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
      try { stateObj = JSON.parse(_joinDraftChunks_(data[i])); } catch(e) {}
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

/**
 * Склеивает JSON-строку черновика из всех ячеек начиная с колонки 4 (index 3).
 * Поддерживает старый формат (одна ячейка) и новый разбитый на куски.
 */
function _joinDraftChunks_(row) {
  if (!row || row.length < 4) return '{}';
  const parts = [];
  for (let j = 3; j < row.length; j++) {
    const part = row[j];
    if (part === '' || part == null) break;
    parts.push(String(part));
  }
  return parts.join('') || '{}';
}

/**
 * Читает рабочий черновик 'working_main' и возвращает {row, state}.
 * Учитывает что JSON может быть разбит на несколько ячеек.
 * Возвращает null если строки нет.
 */
function _readWorkingDraft_() {
  const sh = getSheet_(SH_DET_DRAFTS);
  const workRow = findRowById_(sh, WORKING_DRAFT_ID);
  if (workRow < 0) return null;
  const lastCol = Math.max(4, sh.getLastColumn());
  const cells = sh.getRange(workRow, 1, 1, lastCol).getValues()[0];
  let state = {};
  try { state = JSON.parse(_joinDraftChunks_(cells)); } catch(e) {}
  return { row: workRow, sheet: sh, state, cells };
}

/**
 * Записывает workState в рабочий черновик с разбиением на куски.
 * Перед записью удаляет старую строку (чтобы не было «хвостов»).
 */
function _writeWorkingDraft_(workState, name) {
  const sh = getSheet_(SH_DET_DRAFTS);
  const json = JSON.stringify(workState);
  const CHUNK = 49000;
  const chunks = [];
  for (let i = 0; i < json.length; i += CHUNK) {
    chunks.push(json.slice(i, i + CHUNK));
  }
  if (chunks.length === 0) chunks.push('');

  const existing = findRowById_(sh, WORKING_DRAFT_ID);
  if (existing > 0) sh.deleteRow(existing);

  const finalName = name || 'Рабочий черновик · все события';
  const rowData = [WORKING_DRAFT_ID, finalName, new Date().toISOString(), ...chunks];
  sh.appendRow(rowData);
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
// МИГРАЦИЯ В ЕДИНЫЙ РАБОЧИЙ ЧЕРНОВИК
// ============================================================
const WORKING_DRAFT_ID = 'working_main';

function migrateToWorkingDraft() {
  const sh = getSheet_(SH_DET_DRAFTS);
  const data = sh.getDataRange().getValues();
  const allEvents = [];
  const rowsToDelete = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = String(row[0] || '');
    if (!id) continue;
    if (!id.startsWith('imp_')) continue;
    let st;
    try { st = JSON.parse(String(row[3] || '{}')); } catch(e) { continue; }
    (st.events || []).forEach(ev => allEvents.push(ev));
    rowsToDelete.push(i + 1);
  }

  if (allEvents.length === 0) {
    Logger.log('Не найдено импортированных черновиков (imp_*). Может, миграция уже была сделана?');
    return;
  }

  const byId = {};
  allEvents.forEach(ev => {
    if (!byId[ev.id]) byId[ev.id] = ev;
  });
  const merged = Object.values(byId).sort((a, b) => (a.date + (a.timeStart || '')).localeCompare(b.date + (b.timeStart || '')));

  Logger.log('Объединяю ' + allEvents.length + ' событий из ' + rowsToDelete.length + ' черновиков, после dedup: ' + merged.length);

  rowsToDelete.sort((a, b) => b - a).forEach(r => sh.deleteRow(r));

  const existRow = findRowById_(sh, WORKING_DRAFT_ID);
  if (existRow > 0) sh.deleteRow(existRow);

  const stateObj = {
    monthKey: '',
    events: merged,
    publishedAt: new Date().toISOString(),
    editorName: 'миграция',
    isWorking: true
  };
  sh.appendRow([
    WORKING_DRAFT_ID,
    'Рабочий черновик · все события',
    new Date().toISOString(),
    JSON.stringify(stateObj)
  ]);

  Logger.log('---');
  Logger.log('Создан рабочий черновик: ' + WORKING_DRAFT_ID);
  Logger.log('Событий в нём: ' + merged.length);
  Logger.log('Событий по месяцам:');
  const byMonth = {};
  merged.forEach(e => {
    const k = (e.date || '').slice(0, 7);
    byMonth[k] = (byMonth[k] || 0) + 1;
  });
  Object.keys(byMonth).sort().forEach(k => Logger.log('  ' + k + ': ' + byMonth[k]));
  Logger.log('---');
  Logger.log('ГОТОВО. Откройте приложение — рабочий черновик подгрузится автоматически.');
}

// ============================================================
// СИНХРОНИЗАЦИЯ ИЗ АЛЬФЫ → КОНСТРУКТОР
// ============================================================

const IMPORT_UNMATCHED_FROM_ALFA = false;

function syncFromAlfaDryRun() { return _syncFromAlfa_(false); }
function syncFromAlfaApply()  { return _syncFromAlfa_(true);  }

const PRIVATE_FORMATS_FOR_AUTOLINK = ['День рождения', 'Свидание'];

function findSafeAutoLinksDryRun() { return _findSafeAutoLinks_(false); }
function applySafeAutoLinks()      { return _findSafeAutoLinks_(true);  }

/**
 * Исправляет перепутанные поля в рабочем черновике (legacy).
 */
function fixSwappedFields() {
  const draft = _readWorkingDraft_();
  if (!draft) { Logger.log('❌ Рабочий черновик не найден'); return; }
  const workState = draft.state;

  const events = workState.events || [];
  let fixed = 0, skipped = 0;
  events.forEach(ev => {
    const t = ev.teacher;
    const looksLikePrice = (typeof t === 'string' && /^\d+(\.\d+)?$/.test(t.trim())) ||
                           (typeof t === 'number' && t > 0);
    if (!looksLikePrice) {
      skipped++;
      return;
    }
    const priceFromTeacher = Number(t);
    const limitFromPrice = (typeof ev.price === 'number' && ev.price > 0 && ev.price < 50)
                           ? ev.price : null;

    ev.price = priceFromTeacher;
    ev.limit = limitFromPrice || ev.limit;
    ev.teacher = '';
    fixed++;
  });

  workState.events = events;
  _writeWorkingDraft_(workState);

  Logger.log('✅ Исправлено событий: ' + fixed);
  Logger.log('   Пропущено: ' + skipped);
}

/**
 * Пушит присланный со фронта стейт в snapshot_main — публикация для view.html / team.html.
 *
 * Фронт перед вызовом:
 *  1) Проставляет publishedToClients / publishedToTeam у нужного события
 *  2) Сохраняет черновик через saveDraft (working_main)
 *  3) Зовёт этот экшен с текущим stateJson — он перезаписывает snapshot_main.
 *
 * После этого view.html и team.html сразу видят изменения (фильтр по флагам).
 */
function detailsPushSnapshot_(body) {
  const stateJson = String(body.stateJson || '');
  if (!stateJson) return { ok: false, error: 'нет stateJson' };
  let st;
  try { st = JSON.parse(stateJson); } catch(e) { return { ok: false, error: 'битый stateJson' }; }
  if (!Array.isArray(st.events)) st.events = [];

  st.monthKey = '';
  st.publishedAt = new Date().toISOString();

  const id = 'snapshot_main';
  const today = new Date();
  const monthNames = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
  const name = 'Снимок ' + monthNames[today.getMonth()] + ' ' + today.getFullYear() +
               ' · ' + Utilities.formatDate(today, 'Europe/Moscow', 'HH:mm');
  const pubSh = getSheet_(SH_DET_PUBLISHED);
  const existRow = findRowById_(pubSh, id);
  if (existRow > 0) pubSh.deleteRow(existRow);
  pubSh.appendRow([id, name, '', today.toISOString(), JSON.stringify(st)]);

  // Считаем сколько в snapshot будет видно клиентам и команде
  const visClients = st.events.filter(e =>
    e.alfaGroupId && e.publishedToClients !== false &&
    ['День рождения', 'Свидание'].indexOf(e.format) < 0
  ).length;
  const visTeam = st.events.filter(e => e.publishedToTeam !== false).length;

  return { ok: true, snapshotAt: today.toISOString(), visibleClients: visClients, visibleTeam: visTeam };
}

/**
 * Импорт групп из Альфы, которых нет в локальном state.
 * body.knownGroupIds — массив alfaGroupId, которые уже есть у нас (присылает фронт).
 * Возвращает: { ok: true, newEvents: [...] } — события, готовые к добавлению в state.
 *
 * Не пишет ничего в БД — только формирует объекты. Фронт сам решает что добавить.
 */
function detailsImportNewFromAlfa_(body) {
  if (typeof alfaLogin_ !== 'function') return { ok: false, error: 'Alfa.gs не найден' };

  const knownIds = {};
  ((body && body.knownGroupIds) || []).forEach(id => { knownIds[Number(id)] = true; });

  let token;
  try { token = alfaLogin_(); }
  catch(e) { return { ok: false, error: 'Логин: ' + (e.message || e) }; }

  // 1. Тянем все группы постранично
  const allGroups = [];
  const seen = {};
  let page = 0;
  while (true) {
    let resp;
    try {
      resp = alfaCall_(token, CFG.BRANCH_ID, '/v2api/' + CFG.BRANCH_ID + '/group/index', { page });
    } catch(e) { return { ok: false, error: 'group/index: ' + (e.message || e) }; }
    const items = (resp && resp.items) || [];
    if (items.length === 0) break;
    let news = 0;
    items.forEach(g => { if (!seen[g.id]) { seen[g.id] = true; allGroups.push(g); news++; } });
    if (news === 0) break;
    page++;
    if (page > 30) break;
  }

  // 2. Активные нашего филиала, ещё не известные нам
  let tz = ''; try { tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone(); } catch(e){}
  if (!tz) tz = 'Europe/Moscow';
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  const unknownGroups = allGroups.filter(g => {
    if (g.removed === 1) return false;
    if (g.branch_ids && g.branch_ids.indexOf(CFG.BRANCH_ID) < 0) return false;
    if (knownIds[Number(g.id)]) return false;  // у нас уже есть
    return true;
  });

  if (unknownGroups.length === 0) {
    return { ok: true, newEvents: [], message: 'Все группы из Альфы уже есть в админке' };
  }

  // 3. Тянем регулярки чтобы получить дату/время каждой группы
  const regularByGroup = {};
  try {
    let rpage = 0;
    while (rpage < 30) {
      const rr = alfaCall_(token, CFG.BRANCH_ID,
        '/v2api/' + CFG.BRANCH_ID + '/regular-lesson/index', { page: rpage });
      const ritems = (rr && rr.items) || [];
      if (ritems.length === 0) break;
      ritems.forEach(rl => {
        if (rl.related_class !== 'Group') return;
        const gid = Number(rl.related_id);
        if (!gid) return;
        if (regularByGroup[gid]) return;   // берём первое попавшееся
        regularByGroup[gid] = {
          b_date: rl.b_date || '',
          time_from: rl.time_from_v || '',
          time_to: rl.time_to_v || ''
        };
      });
      rpage++;
    }
  } catch(e) { /* не критично */ }

  // 4. Форматы — для подбора format/name из имени группы
  const formats = detailsGetFormats_().map(f => f.name).filter(Boolean);
  formats.sort((a, b) => b.length - a.length);   // длинные сперва

  // 5. Собираем новые события
  const newEvents = [];
  unknownGroups.forEach(g => {
    const reg = regularByGroup[g.id];
    let dateStr = '';
    let timeStart = '';
    let durationH = 2.5;

    if (reg && reg.b_date) {
      dateStr = String(reg.b_date).slice(0, 10);
      if (reg.time_from) {
        const m = reg.time_from.match(/(\d{1,2}):(\d{2})/);
        if (m) timeStart = String(m[1]).padStart(2, '0') + ':' + m[2];
      }
      if (reg.time_from && reg.time_to) {
        const [h1, m1] = reg.time_from.split(':').map(Number);
        const [h2, m2] = reg.time_to.split(':').map(Number);
        if (!isNaN(h1) && !isNaN(h2)) {
          let mins = (h2 * 60 + (m2||0)) - (h1 * 60 + (m1||0));
          if (mins < 0) mins += 24 * 60;
          if (mins > 0) durationH = +(mins / 60).toFixed(2);
        }
      }
    } else if (g.b_date) {
      dateStr = normalizeDate_(g.b_date);
    }

    // Пропускаем группы без даты или с прошедшей датой
    if (!dateStr) return;
    if (dateStr < todayStr) return;

    // Парсим имя
    const parsed = _parseAlfaGroupName_(g.name, formats);

    newEvents.push({
      id: 'alfa_' + g.id,
      date: dateStr,
      timeStart: timeStart || '18:00',
      duration: durationH,
      format: parsed.format,
      name: parsed.name,
      teacher: '',
      price: parsed.price,
      limit: Number(g.limit) || 12,
      costPerPerson: null,
      costGroup: null,
      description: '',
      alfaNote: String(g.note || ''),
      status: 'published',
      publishToAlfa: false,
      publishedToAlfa: true,
      alfaGroupId: g.id,
      alfaLessonId: null,
      formUrl: CFG.ALFA_HOST + '/common/' + CFG.BRANCH_ID + '/lead/create?gid=' + g.id,
      comment: '',
      references: [],
      teamComment: '',
      publishedToClients: true,
      publishedToTeam: true
    });
  });

  newEvents.sort((a, b) => (a.date + a.timeStart).localeCompare(b.date + b.timeStart));

  return {
    ok: true,
    newEvents,
    totalUnknown: unknownGroups.length,
    skippedNoDate: unknownGroups.length - newEvents.length
  };
}

function detailsClearSnapshot_(body) {
  // Удаляет строку snapshot_main из листа SH_DET_PUBLISHED.
  // После этого view.html и team.html отдают пустые списки до следующего pushSnapshot.
  try {
    const pubSh = getSheet_(SH_DET_PUBLISHED);
    const row = findRowById_(pubSh, 'snapshot_main');
    if (row > 0) {
      pubSh.deleteRow(row);
      Logger.log('[clearSnapshot] удалена строка ' + row);
      return { ok: true, deleted: true };
    }
    return { ok: true, deleted: false, reason: 'snapshot_main не найден' };
  } catch(e) {
    return { ok: false, error: String(e.message || e) };
  }
}

function publishWorkingDraftAsSnapshot() {
  const draftsSh = getSheet_(SH_DET_DRAFTS);
  const workRow = findRowById_(draftsSh, WORKING_DRAFT_ID);
  if (workRow < 0) {
    Logger.log('❌ Рабочий черновик не найден');
    return;
  }
  const cells = draftsSh.getRange(workRow, 1, 1, 4).getValues()[0];
  let workState;
  try { workState = JSON.parse(String(cells[3] || '{}')); }
  catch(e) { Logger.log('❌ Битый JSON'); return; }

  const events = workState.events || [];
  Logger.log('Событий в рабочем черновике: ' + events.length);

  const byStatus = { draft: 0, published: 0, modified: 0, archived: 0 };
  events.forEach(ev => {
    const s = ev.status || 'draft';
    byStatus[s] = (byStatus[s] || 0) + 1;
  });
  Logger.log('  ◌ draft: ' + byStatus.draft);
  Logger.log('  ✓ published: ' + byStatus.published);
  Logger.log('  ⏳ modified: ' + byStatus.modified);
  Logger.log('  🗑 archived: ' + byStatus.archived);
  Logger.log('---');

  const id = 'snapshot_main';
  const monthNames = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
  const today = new Date();
  const name = 'Снимок ' + monthNames[today.getMonth()] + ' ' + today.getFullYear() +
               ' · ' + Utilities.formatDate(today, 'Europe/Moscow', 'HH:mm');

  workState.monthKey = '';
  workState.publishedAt = today.toISOString();

  const pubSh = getSheet_(SH_DET_PUBLISHED);
  const existRow = findRowById_(pubSh, id);
  const rowData = [id, name, '', today.toISOString(), JSON.stringify(workState)];
  if (existRow > 0) pubSh.deleteRow(existRow);
  pubSh.appendRow(rowData);

  Logger.log('✅ Снимок сохранён в «Опубликованные_детали» (id=' + id + ')');
  Logger.log('Открывайте view.html и team.html — события появятся сразу.');
  Logger.log('');
  Logger.log('Что увидят клиенты на view.html:');
  const today0 = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const visible = events.filter(ev => {
    if (!ev.alfaGroupId) return false;
    if (['День рождения', 'Свидание'].indexOf(ev.format) >= 0) return false;
    const d = new Date(ev.date);
    return d >= today0;
  });
  Logger.log('  ' + visible.length + ' событий (опубликованных в Альфе, будущих, без частных)');
}

function _findSafeAutoLinks_(apply) {
  if (typeof alfaLogin_ !== 'function') {
    Logger.log('❌ Файл Alfa.gs не найден');
    return;
  }
  const token = alfaLogin_();
  if (!token) { Logger.log('❌ Логин не удался'); return; }

  const allGroups = [];
  let page = 0;
  const seen = {};
  while (true) {
    const resp = alfaCall_(token, CFG.BRANCH_ID,
      '/v2api/' + CFG.BRANCH_ID + '/group/index', { page: page });
    const items = resp.items || [];
    if (items.length === 0) break;
    let news = 0;
    items.forEach(g => {
      if (!seen[g.id] && (!g.removed || g.removed === 0)) {
        seen[g.id] = true;
        allGroups.push(g);
        news++;
      }
    });
    if (news === 0) break;
    page++;
    if (page > 30) break;
  }
  Logger.log('Получено активных групп в Альфе: ' + allGroups.length);

  const alfaByDate = {};
  allGroups.forEach(g => {
    const d = normalizeDate_(g.b_date);
    if (!d) return;
    if (!alfaByDate[d]) alfaByDate[d] = [];
    alfaByDate[d].push(g);
  });

  const draftsSh = getSheet_(SH_DET_DRAFTS);
  const workRow = findRowById_(draftsSh, WORKING_DRAFT_ID);
  if (workRow < 0) { Logger.log('❌ Рабочий черновик не найден'); return; }
  const workDataCells = draftsSh.getRange(workRow, 1, 1, 4).getValues()[0];
  let workState;
  try { workState = JSON.parse(String(workDataCells[3] || '{}')); }
  catch(e) { Logger.log('❌ Битый JSON'); return; }
  const events = workState.events || [];
  const candidates = events.filter(e =>
    !e.alfaGroupId && e.status !== 'archived' && e.date
  );

  Logger.log('Событий-кандидатов (без привязки): ' + candidates.length);
  Logger.log('---');

  const safeLinks = [];
  const unsafeMulti = [];
  const unsafePrivate = [];
  const unsafeMismatch = [];
  const noCandidate = [];

  candidates.sort((a, b) => (a.date + (a.timeStart || '')).localeCompare(b.date + (b.timeStart || '')));

  candidates.forEach(ev => {
    const dayCands = alfaByDate[ev.date] || [];
    if (dayCands.length === 0) { noCandidate.push(ev); return; }
    if (PRIVATE_FORMATS_FOR_AUTOLINK.indexOf(ev.format) >= 0) {
      unsafePrivate.push({ event: ev, candidates: dayCands });
      return;
    }
    if (dayCands.length > 1) {
      unsafeMulti.push({ event: ev, candidates: dayCands });
      return;
    }
    const cand = dayCands[0];
    if (!hasNameOverlap_(ev, cand)) {
      unsafeMismatch.push({ event: ev, candidate: cand });
      return;
    }
    safeLinks.push({ event: ev, group: cand });
  });

  Logger.log('🟢 БЕЗОПАСНО ПРИВЯЗАТЬ АВТОМАТОМ: ' + safeLinks.length);
  safeLinks.forEach(l => {
    Logger.log('  ✓ ' + l.event.date + ' ' + (l.event.timeStart || '') + ' — ' +
               (l.event.name || l.event.format) +
               '   →  gid=' + l.group.id + ' "' + l.group.name + '"');
  });
  Logger.log('');

  if (unsafePrivate.length > 0) {
    Logger.log('🟡 РУЧНОЙ РАЗБОР — ДР / Свидание (' + unsafePrivate.length + '):');
    unsafePrivate.forEach(u => {
      Logger.log('  • ' + u.event.date + ' ' + (u.event.timeStart || '') + ' — ' +
                 u.event.name + ' [' + u.event.format + ']');
      u.candidates.forEach(g => Logger.log('       gid=' + g.id + ' → ' + g.name));
    });
    Logger.log('');
  }

  if (unsafeMulti.length > 0) {
    Logger.log('🟡 РУЧНОЙ РАЗБОР — несколько кандидатов на дату (' + unsafeMulti.length + '):');
    unsafeMulti.forEach(u => {
      Logger.log('  • ' + u.event.date + ' ' + (u.event.timeStart || '') + ' — ' +
                 u.event.name + ' [' + u.event.format + ']');
      u.candidates.forEach(g => Logger.log('       gid=' + g.id + ' → ' + g.name));
    });
    Logger.log('');
  }

  if (unsafeMismatch.length > 0) {
    Logger.log('🔴 НЕ СОВПАДАЮТ ПО НАЗВАНИЮ — пропускаем (' + unsafeMismatch.length + '):');
    unsafeMismatch.forEach(u => {
      Logger.log('  • ' + u.event.date + ' "' + u.event.name + '" [' + u.event.format + ']');
      Logger.log('       Альфа: gid=' + u.candidate.id + ' "' + u.candidate.name + '"');
    });
    Logger.log('');
  }

  if (noCandidate.length > 0) {
    Logger.log('⚪ В Альфе нет групп на эту дату (' + noCandidate.length + '):');
    noCandidate.slice(0, 10).forEach(ev => {
      Logger.log('  • ' + ev.date + ' — ' + ev.name + ' [' + ev.format + ']');
    });
    if (noCandidate.length > 10) Logger.log('  … и ещё ' + (noCandidate.length - 10));
    Logger.log('');
  }

  Logger.log('---');
  Logger.log('📊 ИТОГО:');
  Logger.log('  🟢 Безопасных автопривязок: ' + safeLinks.length);
  Logger.log('  🟡 На ручной разбор: ' + (unsafePrivate.length + unsafeMulti.length));
  Logger.log('  🔴 Не совпало: ' + unsafeMismatch.length);
  Logger.log('  ⚪ Без кандидатов: ' + noCandidate.length);
  Logger.log('---');

  if (!apply) {
    Logger.log('🟡 РЕЖИМ DRY-RUN — НИЧЕГО НЕ ИЗМЕНЕНО.');
    Logger.log('Запустите applySafeAutoLinks(), чтобы применить.');
    return;
  }

  Logger.log('💾 ПРИМЕНЯЮ АВТОПРИВЯЗКИ…');
  let applied = 0;
  safeLinks.forEach(l => {
    const ev = events.find(e => e.id === l.event.id);
    if (!ev) return;
    ev.alfaGroupId = l.group.id;
    ev.formUrl = CFG.ALFA_HOST + '/common/' + CFG.BRANCH_ID + '/lead/create?gid=' + l.group.id;
    ev.status = 'published';
    ev.publishedToAlfa = true;
    applied++;
  });

  workState.events = events;
  workState.lastAutoLinkAt = new Date().toISOString();
  draftsSh.getRange(workRow, 3).setValue(new Date().toISOString());
  draftsSh.getRange(workRow, 4).setValue(JSON.stringify(workState));

  Logger.log('  ✅ Привязано: ' + applied);
  Logger.log('ГОТОВО. Обновите страницу приложения.');
  Logger.log('Остались на ручную привязку: ' + (unsafePrivate.length + unsafeMulti.length));
}

function hasNameOverlap_(ev, group) {
  const stopWords = new Set([
    'мк', 'и', 'в', 'на', 'по', 'с', 'для', 'из', 'к', 'про',
    'руб', 'групповой', 'детей', 'мама', 'ребенок',
    'мастер', 'мастеркласс', 'класс', 'встреча', 'вечер'
  ]);
  const norm = s => String(s || '')
    .toLowerCase()
    .replace(/[«»""''„"`()\[\],\.\-—–]/g, ' ')
    .replace(/\d+\s*руб/g, '')
    .split(/\s+/)
    .filter(w => w.length > 2 && !stopWords.has(w));

  const evWords = new Set([...norm(ev.name), ...norm(ev.format)]);
  const grWords = new Set(norm(group.name));

  for (const w of evWords) if (grWords.has(w)) return true;

  const conceptMap = {
    'глина':       ['глине', 'глины'],
    'рисование':   ['живопись', 'живописи', 'картина'],
    'украшения':   ['макраме', 'брошь', 'бисер'],
    'киновечер':   ['киновечер'],
    'дегустация':  ['дегустация'],
    'кулинарный':  ['кулинарный', 'кулинарного'],
    'караоке':     ['караоке'],
    'вечеринка':   ['вечеринка', 'пижамная'],
  };
  for (const w of evWords) {
    const aliases = conceptMap[w];
    if (aliases) for (const a of aliases) if (grWords.has(a)) return true;
  }
  for (const w of grWords) {
    const aliases = conceptMap[w];
    if (aliases) for (const a of aliases) if (evWords.has(a)) return true;
  }
  return false;
}

function findAlfaCandidates() {
  if (typeof alfaLogin_ !== 'function') {
    Logger.log('❌ Файл Alfa.gs не найден');
    return;
  }
  const token = alfaLogin_();
  if (!token) { Logger.log('❌ Логин не удался'); return; }

  const allGroups = [];
  let page = 0;
  const seen = {};
  while (true) {
    const resp = alfaCall_(token, CFG.BRANCH_ID,
      '/v2api/' + CFG.BRANCH_ID + '/group/index', { page: page });
    const items = resp.items || [];
    if (items.length === 0) break;
    let news = 0;
    items.forEach(g => {
      if (!seen[g.id] && (!g.removed || g.removed === 0)) {
        seen[g.id] = true;
        allGroups.push(g);
        news++;
      }
    });
    if (news === 0) break;
    page++;
    if (page > 30) break;
  }
  Logger.log('Получено групп в Альфе: ' + allGroups.length);

  const alfaByDate = {};
  allGroups.forEach(g => {
    const d = normalizeDate_(g.b_date);
    if (!d) return;
    if (!alfaByDate[d]) alfaByDate[d] = [];
    alfaByDate[d].push(g);
  });

  const draftsSh = getSheet_(SH_DET_DRAFTS);
  const workRow = findRowById_(draftsSh, WORKING_DRAFT_ID);
  if (workRow < 0) { Logger.log('❌ Рабочий черновик не найден'); return; }
  const workData = draftsSh.getRange(workRow, 4).getValue();
  let workState;
  try { workState = JSON.parse(String(workData || '{}')); }
  catch(e) { Logger.log('❌ Битый JSON'); return; }
  const events = (workState.events || []).filter(e => !e.alfaGroupId && e.status !== 'archived');

  Logger.log('Событий в конструкторе без привязки к Альфе: ' + events.length);
  Logger.log('---');
  Logger.log('🔍 КАНДИДАТЫ ДЛЯ ПРИВЯЗКИ:');
  Logger.log('');

  events.sort((a, b) => (a.date + (a.timeStart || '')).localeCompare(b.date + (b.timeStart || '')));

  let withCandidates = 0;
  let withoutCandidates = 0;

  events.forEach(ev => {
    const candidates = alfaByDate[ev.date] || [];
    if (candidates.length === 0) {
      withoutCandidates++;
      return;
    }
    withCandidates++;
    Logger.log('📅 ' + ev.date + ' ' + (ev.timeStart || '—') + ' — ' + (ev.name || ev.format || '?') +
               ' [' + (ev.format || '') + ']');
    candidates.forEach(g => {
      const limit = g.limit ? ' · лимит=' + g.limit : '';
      Logger.log('     gid=' + g.id + ' → ' + g.name + limit);
    });
    Logger.log('');
  });

  Logger.log('---');
  Logger.log('📊 ИТОГО:');
  Logger.log('  С кандидатами: ' + withCandidates);
  Logger.log('  Без кандидатов: ' + withoutCandidates);

  const eventDates = new Set(events.map(e => e.date));
  const orphans = allGroups.filter(g => {
    const d = normalizeDate_(g.b_date);
    return d && !eventDates.has(d);
  });
  if (orphans.length > 0) {
    Logger.log('');
    Logger.log('⚠ В АЛЬФЕ ЕСТЬ ' + orphans.length + ' ГРУПП БЕЗ СОБЫТИЯ В КОНСТРУКТОРЕ НА ТУ ЖЕ ДАТУ:');
    orphans.slice(0, 20).forEach(g => {
      const d = normalizeDate_(g.b_date);
      Logger.log('  • ' + d + ' — ' + g.name + ' (gid=' + g.id + ')');
    });
    if (orphans.length > 20) Logger.log('  … и ещё ' + (orphans.length - 20));
  }
}

function normalizeDate_(v) {
  if (!v) return '';
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, '0');
    const d = String(v.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (m) {
    return m[3] + '-' + String(m[2]).padStart(2,'0') + '-' + String(m[1]).padStart(2,'0');
  }
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) {
    return m[3] + '-' + String(m[1]).padStart(2,'0') + '-' + String(m[2]).padStart(2,'0');
  }
  return '';
}

// ============================================================
// РУЧНАЯ ПРИВЯЗКА К АЛЬФА-ГРУППЕ
// ============================================================
function detailsLinkAlfaGroup_(body) {
  const eventId = String(body.eventId || '').trim();
  const groupId = parseInt(body.groupId, 10);
  if (!eventId || !groupId) return { ok: false, error: 'нет eventId или groupId' };

  if (typeof alfaLogin_ !== 'function') return { ok: false, error: 'Alfa.gs не найден' };
  let token, group;
  try {
    token = alfaLogin_();
  } catch(e) {
    return { ok: false, error: 'Не удалось залогиниться: ' + (e.message || e) };
  }

  return {
    ok: true,
    alfaGroupId: groupId,
    formUrl: CFG.ALFA_HOST + '/common/' + CFG.BRANCH_ID + '/lead/create?gid=' + groupId
  };
}

function discoverAlfaEndpoints() {
  if (typeof alfaLogin_ !== 'function') {
    Logger.log('❌ Файл Alfa.gs не найден');
    return;
  }
  const token = alfaLogin_();
  if (!token) { Logger.log('❌ Логин не удался'); return; }
  Logger.log('✅ Логин успешен');
  Logger.log('---');

  const candidates = [
    '/v2api/' + CFG.BRANCH_ID + '/customer-group/index',
    '/v2api/' + CFG.BRANCH_ID + '/group/index',
    '/v2api/' + CFG.BRANCH_ID + '/customer/index',
    '/v2api/' + CFG.BRANCH_ID + '/lesson/index',
  ];
  candidates.forEach(path => {
    try {
      const resp = UrlFetchApp.fetch(CFG.ALFA_HOST + path, {
        method: 'post',
        contentType: 'application/json',
        headers: { 'X-ALFACRM-TOKEN': token, 'X-APP-KEY': CFG.ALFA_APP_KEY },
        payload: JSON.stringify({}),
        muteHttpExceptions: true,
      });
      const code = resp.getResponseCode();
      const text = resp.getContentText().slice(0, 200);
      if (code === 200) {
        try {
          const json = JSON.parse(resp.getContentText());
          Logger.log('✅ ' + path + ' → 200 (total=' + (json.total || '?') + ', items=' + (json.items ? json.items.length : '?') + ')');
        } catch(e) {
          Logger.log('✅ ' + path + ' → 200 (но не JSON)');
        }
      } else {
        Logger.log('❌ ' + path + ' → ' + code);
      }
    } catch(e) {
      Logger.log('❌ ' + path + ' — exception: ' + e.message);
    }
  });
}

function _syncFromAlfa_(apply) {
  if (typeof alfaLogin_ !== 'function') {
    Logger.log('❌ Файл Alfa.gs не найден');
    return;
  }
  const token = alfaLogin_();
  if (!token) { Logger.log('❌ Не удалось залогиниться в Альфу'); return; }

  Logger.log('🔄 Загружаю группы из Альфы…');
  const allGroups = [];
  let page = 0;
  const seenGroupIds = {};
  while (true) {
    const resp = alfaCall_(token, CFG.BRANCH_ID,
      '/v2api/' + CFG.BRANCH_ID + '/group/index', { page: page });
    const items = resp.items || [];
    if (items.length === 0) break;
    let newCount = 0;
    items.forEach(g => {
      if (!seenGroupIds[g.id]) {
        seenGroupIds[g.id] = true;
        allGroups.push(g);
        newCount++;
      }
    });
    if (newCount === 0) break;
    page++;
    if (page > 30) break;
  }
  Logger.log('  Получено групп в Альфе: ' + allGroups.length);

  const activeGroups = allGroups.filter(g => {
    if (g.removed === 1) return false;
    if (g.branch_ids && g.branch_ids.indexOf(CFG.BRANCH_ID) < 0) return false;
    return true;
  });
  Logger.log('  Из них активных в филиале «Детали»: ' + activeGroups.length);

  const draftsSh = getSheet_(SH_DET_DRAFTS);
  const workRow = findRowById_(draftsSh, WORKING_DRAFT_ID);
  if (workRow < 0) {
    Logger.log('❌ Рабочий черновик не найден. Запустите migrateToWorkingDraft.');
    return;
  }
  const workData = draftsSh.getRange(workRow, 1, 1, 4).getValues()[0];
  let workState;
  try { workState = JSON.parse(String(workData[3] || '{}')); }
  catch(e) { Logger.log('❌ Битый JSON в рабочем черновике'); return; }
  const events = workState.events || [];
  Logger.log('  Событий в конструкторе: ' + events.length);
  Logger.log('---');

  const eventByKey = {};
  events.forEach(ev => {
    const key = makeMatchKey_(ev.date, ev.name);
    if (key) eventByKey[key] = ev;
  });

  Logger.log('🔄 Подгружаю уроки…');
  const lessonsByGroupId = {};
  const now = new Date();
  const dateFrom = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth() - 6, 1), 'Europe/Moscow', 'yyyy-MM-dd');
  const dateTo = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth() + 6, 28), 'Europe/Moscow', 'yyyy-MM-dd');

  let lpage = 0;
  const seenLessonIds = {};
  let totalLessons = 0;
  while (true) {
    const resp = alfaCall_(token, CFG.BRANCH_ID,
      '/v2api/' + CFG.BRANCH_ID + '/lesson/index', { date_from: dateFrom, date_to: dateTo, page: lpage });
    const items = resp.items || [];
    if (items.length === 0) break;
    let newCount = 0;
    items.forEach(l => {
      if (!seenLessonIds[l.id]) {
        seenLessonIds[l.id] = true;
        totalLessons++;
        newCount++;
        (l.customer_group_ids || []).forEach(gid => {
          if (!lessonsByGroupId[gid]) lessonsByGroupId[gid] = [];
          lessonsByGroupId[gid].push(l);
        });
      }
    });
    if (newCount === 0) break;
    lpage++;
    if (lpage > 50) break;
  }
  Logger.log('  Уроков загружено: ' + totalLessons);
  Logger.log('---');

  const matched = [];
  const unmatchedAlfa = [];
  let alfaWithoutDate = 0;

  activeGroups.forEach(g => {
    let groupDate = g.b_date || '';
    const lessons = lessonsByGroupId[g.id] || [];
    if (!groupDate && lessons.length) groupDate = lessons[0].date;
    if (groupDate && groupDate.indexOf('T') > 0) groupDate = groupDate.slice(0, 10);
    if (!groupDate) { alfaWithoutDate++; return; }

    const cleanName = extractNameFromAlfaTitle_(g.name);
    const key = makeMatchKey_(groupDate, cleanName);
    const ev = key ? eventByKey[key] : null;

    if (ev) {
      const conflicts = [];
      const lessonTime = lessons.length && lessons[0].time_from ? lessons[0].time_from.slice(0, 5) : null;
      if (lessonTime && ev.timeStart && ev.timeStart !== lessonTime) {
        conflicts.push({field: 'time', alfa: lessonTime, ours: ev.timeStart});
      }
      if (g.limit && ev.limit && Number(g.limit) !== Number(ev.limit)) {
        conflicts.push({field: 'limit', alfa: g.limit, ours: ev.limit});
      }
      const alfaPrice = extractPriceFromAlfaTitle_(g.name);
      if (alfaPrice && ev.price && Number(alfaPrice) !== Number(ev.price)) {
        conflicts.push({field: 'price', alfa: alfaPrice, ours: ev.price});
      }
      if (ev._matchedAlfaId) {
        return;
      }
      ev._matchedAlfaId = g.id;
      matched.push({group: g, event: ev, conflicts, lessonId: lessons.length ? lessons[0].id : null});
    } else {
      unmatchedAlfa.push({group: g, date: groupDate, cleanName: cleanName});
    }
  });

  const unmatchedOurs = events.filter(ev => !ev._matchedAlfaId);

  Logger.log('📊 РЕЗУЛЬТАТ МАТЧИНГА:');
  Logger.log('  ✅ Сматчено: ' + matched.length);
  Logger.log('  ➕ В Альфе, но нет в конструкторе: ' + unmatchedAlfa.length);
  Logger.log('  ➖ В конструкторе, но нет в Альфе: ' + unmatchedOurs.length);
  Logger.log('---');

  const withConflicts = matched.filter(m => m.conflicts.length > 0);
  Logger.log('⚠ КОНФЛИКТЫ ПОЛЕЙ (' + withConflicts.length + '):');
  withConflicts.forEach(m => {
    Logger.log('  • ' + m.event.date + ' ' + m.event.timeStart + ' — ' + m.event.name + ' (gid=' + m.group.id + ')');
    m.conflicts.forEach(c => {
      Logger.log('      ' + c.field + ': в Альфе=' + c.alfa + ' / у нас=' + c.ours);
    });
  });

  if (!apply) {
    Logger.log('🟡 РЕЖИМ DRY-RUN — изменения НЕ применены.');
    Logger.log('Если результат устраивает — запустите syncFromAlfaApply()');
    return;
  }

  Logger.log('💾 ПРИМЕНЯЮ ИЗМЕНЕНИЯ…');

  matched.forEach(m => {
    const ev = m.event;
    delete ev._matchedAlfaId;
    ev.alfaGroupId = m.group.id;
    ev.alfaLessonId = m.lessonId || null;
    ev.formUrl = CFG.ALFA_HOST + '/common/' + CFG.BRANCH_ID + '/lead/create?gid=' + m.group.id;
    if (!ev.status || ev.status === 'draft') ev.status = 'published';
    if (ev.status === 'published') ev.publishedToAlfa = true;
    if (m.conflicts.length > 0) ev.status = 'modified';
  });

  let imported = 0;
  if (IMPORT_UNMATCHED_FROM_ALFA && unmatchedAlfa.length > 0) {
    unmatchedAlfa.forEach(u => {
      const newEvent = {
        id: 'alfa_' + u.group.id,
        date: u.date,
        timeStart: '18:00',
        duration: 2.5,
        format: '',
        name: u.cleanName || u.group.name,
        teacher: '',
        price: extractPriceFromAlfaTitle_(u.group.name) || null,
        limit: u.group.limit || 12,
        costPerPerson: null,
        description: u.group.note || '',
        status: 'published',
        publishToAlfa: false,
        publishedToAlfa: true,
        alfaGroupId: u.group.id,
        alfaLessonId: null,
        formUrl: CFG.ALFA_HOST + '/common/' + CFG.BRANCH_ID + '/lead/create?gid=' + u.group.id,
        comment: '⚠ импортировано из Альфы — проверьте формат'
      };
      events.push(newEvent);
      imported++;
    });
  }

  events.forEach(ev => delete ev._matchedAlfaId);

  workState.events = events;
  workState.lastSyncedAt = new Date().toISOString();
  draftsSh.getRange(workRow, 3).setValue(new Date().toISOString());
  draftsSh.getRange(workRow, 4).setValue(JSON.stringify(workState));

  Logger.log('  ✅ Сматченным событиям проставлены groupId / formUrl / статус: ' + matched.length);
  if (imported > 0) Logger.log('  ✅ Импортировано из Альфы: ' + imported);
  Logger.log('  💾 Рабочий черновик сохранён.');
}

function makeMatchKey_(dateStr, name) {
  if (!dateStr || !name) return '';
  const d = String(dateStr).slice(0, 10);
  const n = String(name)
    .toLowerCase()
    .replace(/[«»""''„"`]/g, '"')
    .replace(/[\(\)\[\]]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  return d + '|' + n;
}

function extractNameFromAlfaTitle_(title) {
  if (!title) return '';
  const s = String(title);
  const m = s.match(/"([^"]+)"/);
  if (m) return m[1];
  return s.replace(/\s*-\s*\d+\s*руб\.?\s*$/, '').trim();
}

function extractPriceFromAlfaTitle_(title) {
  if (!title) return null;
  const m = String(title).match(/-\s*(\d+)\s*руб/);
  return m ? Number(m[1]) : null;
}

function _testInitDetails() {
  ['Форматы_детали', 'Педагоги_детали', 'Шаблоны_детали',
   'Опубликованные_детали', 'Черновики_детали', 'Сессии_детали']
   .forEach(name => {
    const sh = getSheet_(name);
    Logger.log('OK: ' + name + ' (' + sh.getLastRow() + ' rows)');
  });
}

/**
 * Одноразовая функция: переименовать существующие форматы и добавить новые,
 * чтобы справочник Форматы_детали соответствовал именам в Альфе CRM.
 *
 * Как работает:
 *  1) Бэкапит текущий лист в «Форматы_бэкап_<дата>»
 *  2) Читает текущий список, строит карту по имени
 *  3) Очищает лист от 2-й строки
 *  4) Записывает новый список:
 *     – для переименованных берёт настройки старой записи
 *     – для новых указывает свежие настройки явно
 *     – «Дегустация» отбрасывает (только Винная остаётся)
 *  5) Логирует отчёт
 *
 * Запустить ОДИН РАЗ из редактора Apps Script:
 *   выпадающий список → seedFormatsRebuild → ▶
 */
function seedFormatsRebuild() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SH_DET_FORMATS);
  if (!sh) {
    Logger.log('Лист «' + SH_DET_FORMATS + '» не найден');
    return;
  }

  // 1) Бэкап
  const backupName = 'Форматы_бэкап_' + new Date().toISOString().slice(0, 10);
  if (!ss.getSheetByName(backupName)) {
    sh.copyTo(ss).setName(backupName);
    Logger.log('Бэкап: ' + backupName);
  } else {
    Logger.log('Бэкап «' + backupName + '» уже есть, пропускаю.');
  }

  // 2) Карта по имени
  const lastRow = sh.getLastRow();
  const oldByName = {};
  if (lastRow >= 2) {
    const rows = sh.getRange(2, 1, lastRow - 1, 7).getValues();
    rows.forEach(r => {
      const name = String(r[0] || '').trim();
      if (!name) return;
      oldByName[name] = {
        name,
        category: r[1],
        template: r[2],
        price: r[3],
        limit: r[4],
        duration: r[5],
        active: r[6]
      };
    });
  }

  // 3) Очищаем
  if (lastRow >= 2) {
    sh.getRange(2, 1, lastRow - 1, 7).clearContent();
  }

  // 4) План: [имя_новое, источник_старого_имени_или_null, явные_настройки_если_новый]
  const plan = [
    // Переименование
    { newName: 'Мастер-класс по глине',           oldName: 'МК Глина' },
    { newName: 'Мастер-класс по живописи',        oldName: 'МК Рисование' },
    { newName: 'Мастер-класс по макраме',         oldName: 'МК Украшения' },
    { newName: 'Мастер-класс по бьюти',           oldName: 'МК Бьюти' },
    { newName: 'Мастер-класс кулинарный',         oldName: 'МК Кулинарный' },
    { newName: 'Мастер-класс',                    oldName: 'МК' },
    // Без изменений
    { newName: 'Киновечер',                       oldName: 'Киновечер' },
    { newName: 'Книжный клуб',                    oldName: 'Книжный клуб' },
    { newName: 'Настолки',                        oldName: 'Настолки' },
    { newName: 'Разговорник',                     oldName: 'Разговорник' },
    { newName: 'Караоке',                         oldName: 'Караоке' },
    { newName: 'Вечеринка',                       oldName: 'Вечеринка' },
    { newName: 'Встреча с экспертом',             oldName: 'Встреча с экспертом' },
    { newName: 'День рождения',                   oldName: 'День рождения' },
    { newName: 'Свидание',                        oldName: 'Свидание' },
    // Новые
    {
      newName: 'Мастер-класс по глине для детей',
      oldName: null,
      explicit: { category: 'Творчество', template: 'Лепка из глины для детей.', price: 45, limit: 12, duration: 2.5, active: 'Да' }
    },
    {
      newName: 'Винная дегустация',
      oldName: 'Дегустация', // настройки берём со старой Дегустации
      explicit: { template: 'Винная дегустация.' } // только описание перепишем, остальное со старой
    },
    {
      newName: 'Частное мероприятие',
      oldName: null,
      explicit: { category: 'Частное', template: '', price: '', limit: 12, duration: 3, active: 'Да' }
    }
  ];

  const stats = { renamed: 0, kept: 0, added: 0, dropped: [] };
  const newRows = [];
  const usedOldNames = {};

  plan.forEach(p => {
    const old = p.oldName ? oldByName[p.oldName] : null;
    const ex  = p.explicit || {};
    // Базовые поля — из старой записи (если есть), потом перекрывает explicit
    const row = {
      name:     p.newName,
      category: ex.category !== undefined ? ex.category : (old ? old.category : ''),
      template: ex.template !== undefined ? ex.template : (old ? old.template : ''),
      price:    ex.price    !== undefined ? ex.price    : (old ? old.price    : ''),
      limit:    ex.limit    !== undefined ? ex.limit    : (old ? old.limit    : 12),
      duration: ex.duration !== undefined ? ex.duration : (old ? old.duration : 2.5),
      active:   ex.active   !== undefined ? ex.active   : (old ? old.active   : 'Да')
    };
    newRows.push([row.name, row.category, row.template, row.price, row.limit, row.duration, row.active]);

    if (p.oldName) {
      usedOldNames[p.oldName] = true;
      if (p.oldName === p.newName) stats.kept++;
      else stats.renamed++;
    } else {
      stats.added++;
    }
  });

  // Какие старые форматы не попали в план (= удаляются)
  Object.keys(oldByName).forEach(oldN => {
    if (!usedOldNames[oldN]) stats.dropped.push(oldN);
  });

  // Записываем
  sh.getRange(2, 1, newRows.length, 7).setValues(newRows);

  // 5) Отчёт
  Logger.log('=== ГОТОВО ===');
  Logger.log('Переименовано: ' + stats.renamed);
  Logger.log('Оставлено как есть: ' + stats.kept);
  Logger.log('Добавлено новых: ' + stats.added);
  Logger.log('Удалено: ' + stats.dropped.length + (stats.dropped.length ? ' (' + stats.dropped.join(', ') + ')' : ''));
  Logger.log('Всего теперь форматов: ' + newRows.length);
  Logger.log('');
  Logger.log('СЛЕДУЮЩИЙ ШАГ:');
  Logger.log('1) Открыть админку → Ctrl+Shift+R');
  Logger.log('2) Жмёшь «🔁 Перезагрузить из Альфы»');
  Logger.log('   Теперь форматы из Альфы (Мастер-класс по глине, Вечеринка, …)');
  Logger.log('   будут распознаны и подставлены в админку правильно.');
}

function _testSeedDetails() {
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

// ============================================================
// ИМПОРТ СУЩЕСТВУЮЩИХ СОБЫТИЙ ИЗ ЛИСТА «События»
// ============================================================
function importExistingEvents() {
  const eventsSh = getSheet_(SH_DET_EVENTS);
  if (!eventsSh) {
    Logger.log('Лист «События» не найден');
    return;
  }
  const lastRow = eventsSh.getLastRow();
  if (lastRow < EVENTS_DATA_START_ROW) {
    Logger.log('В листе «События» нет данных');
    return;
  }

  const range = eventsSh.getRange(EVENTS_DATA_START_ROW, 1, lastRow - EVENTS_DATA_START_ROW + 1, 28);
  const rows = range.getValues();
  let tz = '';
  try { tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone(); } catch(e) {}
  if (!tz) tz = 'Europe/Moscow';
  Logger.log('Используем таймзону: ' + tz);

  const eventsByMonth = {};
  const tplsByKey = {};
  const tplOrder = [];

  let imported = 0, skipped = 0;

  rows.forEach((row, idx) => {
    const dateVal = row[0];
    const timeVal = row[2];
    const durVal  = row[3];
    const fmt     = String(row[5] || '').trim();
    const name    = String(row[6] || '').trim();
    const teacher = String(row[8] || '').trim();
    const price   = row[9];
    const limit   = row[10];
    const cpp     = row[12];
    const descr   = String(row[16] || '').trim();
    const status  = String(row[18] || '').trim();
    const groupId = row[19];
    const comment = String(row[27] || '').trim();

    if (!dateVal || (!fmt && !name)) {
      skipped++;
      return;
    }

    let dateStr = '';
    if (dateVal instanceof Date) {
      dateStr = Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
    } else {
      try {
        const d = new Date(1899, 11, 30);
        d.setDate(d.getDate() + Number(dateVal));
        dateStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      } catch(e) { skipped++; return; }
    }

    let timeStr = '18:00';
    if (timeVal instanceof Date) {
      timeStr = Utilities.formatDate(timeVal, tz, 'HH:mm');
    } else if (typeof timeVal === 'number' && timeVal >= 0 && timeVal < 1) {
      const totalMin = Math.round(timeVal * 24 * 60);
      const h = Math.floor(totalMin / 60);
      const m = totalMin % 60;
      timeStr = String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
    } else if (typeof timeVal === 'string' && /^\d{1,2}:\d{2}/.test(timeVal)) {
      const m = timeVal.match(/^(\d{1,2}):(\d{2})/);
      timeStr = String(m[1]).padStart(2, '0') + ':' + m[2];
    }

    let eventId;
    const gidMatch = comment.match(/\[gid:([\w-]+)\]/);
    if (gidMatch) {
      eventId = gidMatch[1];
    } else {
      eventId = 'imp_r' + (EVENTS_DATA_START_ROW + idx);
    }

    const ev = {
      id: eventId,
      date: dateStr,
      timeStart: timeStr,
      duration: Number(durVal) || 2.5,
      format: fmt,
      name: name,
      teacher: teacher,
      price: (typeof price === 'number') ? price : null,
      limit: (typeof limit === 'number') ? limit : null,
      costPerPerson: (typeof cpp === 'number') ? cpp : null,
      description: descr,
      status: groupId ? 'published' : 'draft',
      publishToAlfa: false,
      publishedToAlfa: !!groupId,
      alfaGroupId: groupId || null,
      alfaLessonId: null,
      formUrl: groupId ? (CFG.ALFA_HOST + '/common/' + CFG.BRANCH_ID + '/lead/create?gid=' + groupId) : '',
      comment: comment.replace(/\[gid:[\w-]+\]/g, '').trim()
    };

    const monthKey = dateStr.slice(0, 7);
    if (!eventsByMonth[monthKey]) eventsByMonth[monthKey] = [];
    eventsByMonth[monthKey].push(ev);

    if (fmt && name) {
      const key = (fmt + '|' + name).toLowerCase();
      if (!tplsByKey[key]) {
        tplsByKey[key] = {
          name: name,
          format: fmt,
          teacher: teacher,
          description: descr,
          price: (typeof price === 'number') ? price : null,
          limit: (typeof limit === 'number') ? limit : null,
          duration: Number(durVal) || 2.5,
          tags: '',
          usedCount: 1
        };
        tplOrder.push(key);
      } else {
        tplsByKey[key].usedCount++;
        if (!tplsByKey[key].description && descr) tplsByKey[key].description = descr;
      }
    }

    imported++;
  });

  const tplsSh = getSheet_(SH_DET_TEMPLATES);
  if (tplsSh.getLastRow() > 1) {
    tplsSh.getRange(2, 1, tplsSh.getLastRow() - 1, 9).clearContent();
  }
  const tplRows = tplOrder.map(k => {
    const t = tplsByKey[k];
    return [t.name, t.format, t.teacher, t.description, t.price || '', t.limit || '', t.duration || '', t.tags, t.usedCount];
  });
  if (tplRows.length > 0) {
    tplsSh.getRange(2, 1, tplRows.length, 9).setValues(tplRows);
  }
  Logger.log('Шаблонов записано: ' + tplRows.length);

  const draftsSh = getSheet_(SH_DET_DRAFTS);
  if (draftsSh.getLastRow() > 1) {
    const draftIds = draftsSh.getRange(2, 1, draftsSh.getLastRow() - 1, 1).getValues();
    for (let i = draftIds.length - 1; i >= 0; i--) {
      if (String(draftIds[i][0]).startsWith('imp_')) {
        draftsSh.deleteRow(i + 2);
      }
    }
  }

  const monthNames = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
  const draftsCreated = [];
  Object.keys(eventsByMonth).sort().forEach(monthKey => {
    const events = eventsByMonth[monthKey];
    const [y, m] = monthKey.split('-').map(Number);
    const monthLabel = monthNames[m - 1] + ' ' + y + ' (импорт из таблицы)';
    const draftId = 'imp_' + monthKey;
    const stateObj = {
      monthKey: monthKey,
      events: events,
      publishedAt: new Date().toISOString(),
      editorName: 'импорт'
    };
    draftsSh.appendRow([draftId, monthLabel, new Date().toISOString(), JSON.stringify(stateObj)]);
    draftsCreated.push(monthLabel + ': ' + events.length + ' событий');
  });

  Logger.log('Черновиков создано: ' + draftsCreated.length);
  draftsCreated.forEach(d => Logger.log('  ' + d));
  Logger.log('Всего импортировано событий: ' + imported + ', пропущено: ' + skipped);
}

/**
 * Диагностика подсчёта записанных в группу.
 * Запуск: выпадающий список «функции» в редакторе → _testBookingCounts → ▶
 * Результат: View → Logs (или Cmd/Ctrl+Enter).
 *
 * Цель — найти ПРАВИЛЬНЫЙ способ узнать сколько клиентов записано в конкретную
 * группу Альфы. Пробуем 5 разных гипотез по очереди.
 */
function _testBookingCounts() {
  const TEST_GID = 343;  // каллиграфия — на скрине у юзера видно «0 из 12»

  if (typeof alfaLogin_ !== 'function') { Logger.log('❌ Alfa.gs не найден'); return; }
  let token;
  try { token = alfaLogin_(); }
  catch(e) { Logger.log('❌ Логин: ' + e.message); return; }
  Logger.log('✅ Логин OK, тестируем gid=' + TEST_GID);
  Logger.log('');

  // === Тест 1: смотрим что вообще возвращает /customer/index ===
  Logger.log('═══ Тест 1: /customer/index без фильтра, page=0 ═══');
  try {
    const r = alfaCall_(token, CFG.BRANCH_ID, '/v2api/' + CFG.BRANCH_ID + '/customer/index', { page: 0 });
    Logger.log('Ключи ответа: ' + Object.keys(r).join(', '));
    Logger.log('total: ' + r.total + ', items.length: ' + (r.items || []).length);
    if (r.items && r.items.length > 0) {
      const first = r.items[0];
      Logger.log('Ключи клиента: ' + Object.keys(first).join(', '));
      Logger.log('Клиент целиком: ' + JSON.stringify(first).slice(0, 1500));
    }
  } catch(e) { Logger.log('❌ ' + e.message); }
  Logger.log('');

  // === Тест 2: с заявленным фильтром customer_group_ids ===
  Logger.log('═══ Тест 2: customer_group_ids=[' + TEST_GID + '] ═══');
  try {
    const r = alfaCall_(token, CFG.BRANCH_ID, '/v2api/' + CFG.BRANCH_ID + '/customer/index',
      { customer_group_ids: [TEST_GID], removed: [0], page: 0 });
    Logger.log('total: ' + r.total + ', items.length: ' + (r.items || []).length);
    Logger.log('Если total равен общему числу клиентов из Теста 1 — фильтр игнорируется');
  } catch(e) { Logger.log('❌ ' + e.message); }
  Logger.log('');

  // === Тест 3: альтернативные имена параметра фильтра ===
  Logger.log('═══ Тест 3: альтернативные параметры фильтра ═══');
  const filterVariants = [
    { name: 'study_group_ids', payload: { study_group_ids: [TEST_GID], removed: [0], page: 0 } },
    { name: 'groups',          payload: { groups: [TEST_GID], removed: [0], page: 0 } },
    { name: 'group_ids',       payload: { group_ids: [TEST_GID], removed: [0], page: 0 } },
    { name: 'group_id',        payload: { group_id: TEST_GID, removed: [0], page: 0 } }
  ];
  filterVariants.forEach(v => {
    try {
      const r = alfaCall_(token, CFG.BRANCH_ID, '/v2api/' + CFG.BRANCH_ID + '/customer/index', v.payload);
      Logger.log('  ' + v.name + ': total=' + r.total + ', items=' + (r.items || []).length);
    } catch(e) {
      Logger.log('  ' + v.name + ': ❌ ' + String(e.message || e).slice(0, 80));
    }
  });
  Logger.log('');

  // === Тест 4: альтернативные эндпоинты ===
  Logger.log('═══ Тест 4: альтернативные эндпоинты ═══');
  const endpoints = [
    '/v2api/' + CFG.BRANCH_ID + '/customer-group/index',
    '/v2api/' + CFG.BRANCH_ID + '/customer-tariff/index',
    '/v2api/' + CFG.BRANCH_ID + '/group/show'
  ];
  endpoints.forEach(ep => {
    try {
      const payload = ep.indexOf('group/show') >= 0 ? { id: TEST_GID } : { group_ids: [TEST_GID], page: 0 };
      const r = alfaCall_(token, CFG.BRANCH_ID, ep, payload);
      Logger.log('  ' + ep);
      Logger.log('    Ключи: ' + Object.keys(r).join(', '));
      Logger.log('    total: ' + r.total + ', items: ' + (r.items || []).length);
      if (r.items && r.items.length > 0) {
        Logger.log('    Пример item: ' + JSON.stringify(r.items[0]).slice(0, 400));
      } else if (!r.items) {
        Logger.log('    Содержимое: ' + JSON.stringify(r).slice(0, 400));
      }
    } catch(e) {
      Logger.log('  ' + ep + ': ❌ ' + String(e.message || e).slice(0, 80));
    }
  });
  Logger.log('');

  // === Тест 5: листаем 3 страницы и считаем сами ===
  Logger.log('═══ Тест 5: считаем по customer_group_ids самих клиентов (3 стр) ═══');
  const counts = {};
  let totalScanned = 0;
  const allKeysSeen = {};
  for (let page = 0; page < 3; page++) {
    try {
      const r = alfaCall_(token, CFG.BRANCH_ID, '/v2api/' + CFG.BRANCH_ID + '/customer/index',
        { removed: [0], page: page });
      const items = r.items || [];
      if (items.length === 0) break;
      items.forEach(c => {
        totalScanned++;
        Object.keys(c).forEach(k => { allKeysSeen[k] = true; });
        const groups = c.customer_group_ids || c.groups || c.study_group_ids || c.group_ids || [];
        if (Array.isArray(groups)) {
          groups.forEach(gid => { counts[gid] = (counts[gid] || 0) + 1; });
        }
      });
    } catch(e) { Logger.log('  ❌ page=' + page + ': ' + e.message); break; }
  }
  Logger.log('Все встреченные ключи клиента: ' + Object.keys(allKeysSeen).join(', '));
  Logger.log('Просканировано клиентов: ' + totalScanned);
  Logger.log('Топ-10 групп по числу записей:');
  Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 10)
    .forEach(([gid, n]) => Logger.log('  gid=' + gid + ': ' + n));
  Logger.log('');
  Logger.log('Группа ' + TEST_GID + ': ' + (counts[TEST_GID] || 0) + ' клиентов на 3 стр');
}

function _testProd() {
  const r = detailsBookingCounts_({ groupIds: '343' });
  Logger.log(JSON.stringify(r));
}

// ============================================================
// PULL FROM ALFA — синхронизация в обратную сторону
// ============================================================
function detailsPullFromAlfa_(body) {
  try {
    if (typeof alfaLogin_ !== 'function') {
      return { ok: false, error: 'Файл Alfa.gs не найден' };
    }
    const groupIds = (body && body.groupIds) || [];
    if (!Array.isArray(groupIds) || groupIds.length === 0) {
      return { ok: false, error: 'groupIds пустой' };
    }
    const wantedIds = {};
    groupIds.forEach(id => { wantedIds[Number(id)] = true; });

    const token = alfaLogin_();
    if (!token) return { ok: false, error: 'Не удалось залогиниться в Альфу' };

    // 1. Получаем все активные группы (постранично)
    const allGroups = [];
    const seen = {};
    let page = 0;
    while (true) {
      const resp = alfaCall_(token, CFG.BRANCH_ID,
        '/v2api/' + CFG.BRANCH_ID + '/group/index', { page: page });
      const items = (resp && resp.items) || [];
      if (items.length === 0) break;
      let news = 0;
      items.forEach(g => {
        if (!seen[g.id]) {
          seen[g.id] = true;
          allGroups.push(g);
          news++;
        }
      });
      if (news === 0) break;
      page++;
      if (page > 30) break;
    }

    const groupsById = {};
    allGroups.forEach(g => { groupsById[g.id] = g; });

    // 2. Уроки за следующие 90 дней
    const today = new Date();
    const dateFrom = Utilities.formatDate(today, 'Europe/Moscow', 'yyyy-MM-dd');
    const dateToObj = new Date(today.getTime() + 90 * 24 * 60 * 60 * 1000);
    const dateTo = Utilities.formatDate(dateToObj, 'Europe/Moscow', 'yyyy-MM-dd');

    const lessonsByGroup = {};
    try {
      let lpage = 0;
      while (true) {
        const lr = alfaCall_(token, CFG.BRANCH_ID,
          '/v2api/' + CFG.BRANCH_ID + '/lesson/index',
          { date_from: dateFrom, date_to: dateTo, page: lpage });
        const items = (lr && lr.items) || [];
        if (items.length === 0) break;
        items.forEach(l => {
          const gid = l.group_id || l.subject_id || (l.group && l.group.id);
          if (!gid) return;
          if (!lessonsByGroup[gid]) lessonsByGroup[gid] = [];
          lessonsByGroup[gid].push({
            id: l.id,
            date: l.date || l.lesson_date || '',
            time_from: l.time_from_h || l.time_from || '',
            time_to: l.time_to_h || l.time_to || '',
            room_id: l.room_id || null,
            note: l.note || ''
          });
        });
        lpage++;
        if (lpage > 30) break;
      }
    } catch(eL) {
      Logger.log('lesson/index failed: ' + eL);
    }

    // 2b. Регулярные уроки (тащим целиком — фильтр по group_id в API не работает)
    // Они содержат реальную дату урока (b_date) и время для одноразовых событий.
    // В админке Альфы они показываются как «Регулярные уроки» в карточке группы.
    const regularByGroup = {};
    try {
      let rpage = 0;
      while (true) {
        const rr = alfaCall_(token, CFG.BRANCH_ID,
          '/v2api/' + CFG.BRANCH_ID + '/regular-lesson/index', { page: rpage });
        const ritems = (rr && rr.items) || [];
        if (ritems.length === 0) break;
        ritems.forEach(rl => {
          // Связь с группой через related_class='Group' + related_id
          if (rl.related_class !== 'Group') return;
          const gid = Number(rl.related_id);
          if (!gid) return;
          if (!regularByGroup[gid]) regularByGroup[gid] = [];
          regularByGroup[gid].push({
            id: rl.id,
            b_date: rl.b_date || '',
            e_date: rl.e_date || '',
            time_from: rl.time_from_v || '',
            time_to: rl.time_to_v || '',
            day: rl.day || null,
            is_public: rl.is_public || 0
          });
        });
        rpage++;
        if (rpage > 30) break;
      }
    } catch(eR) {
      Logger.log('regular-lesson/index failed: ' + eR);
    }

    // 3. Собираем результат
    const items = {};
    Object.keys(wantedIds).forEach(gid => {
      const numGid = Number(gid);
      const g = groupsById[numGid];
      if (!g) {
        items[numGid] = { found: false };
        return;
      }
      const isRemoved = (g.removed === 1 || g.removed === true || g.is_archive === 1);
      const lessons = (lessonsByGroup[numGid] || [])
        .slice()
        .sort((a, b) => (a.date + (a.time_from||'')).localeCompare(b.date + (b.time_from||'')));

      // Берём первый регулярный урок этой группы — для одноразовых событий
      // он один и содержит точную дату + время
      const regulars = (regularByGroup[numGid] || [])
        .slice()
        .sort((a, b) => (a.b_date || '').localeCompare(b.b_date || ''));
      const firstRegular = regulars[0] || null;

      // Считаем duration в минутах из time_from/time_to
      let durationMin = null;
      if (firstRegular && firstRegular.time_from && firstRegular.time_to) {
        const [h1, m1] = firstRegular.time_from.split(':').map(Number);
        const [h2, m2] = firstRegular.time_to.split(':').map(Number);
        if (!isNaN(h1) && !isNaN(h2)) {
          durationMin = (h2 * 60 + (m2||0)) - (h1 * 60 + (m1||0));
          if (durationMin < 0) durationMin += 24 * 60;
        }
      }

      items[numGid] = {
        found: !isRemoved,
        archived: isRemoved,
        id: g.id,
        name: g.name || '',
        price: g.price || g.b2c_price || null,
        limit: g.limit || g.max_count || null,
        note: g.note || '',
        // ВЕРХНЕУРОВНЕВЫЕ ПОЛЯ для computeAlfaDiff (фронт смотрит ровно сюда)
        b_date:       firstRegular ? firstRegular.b_date    : (g.b_date || ''),
        time_from:    firstRegular ? firstRegular.time_from : '',
        duration_min: durationMin,
        lessons: lessons,
        regulars: regulars
      };
    });

    return {
      ok: true,
      items: items,
      fetchedAt: new Date().toISOString(),
      counts: {
        wanted: Object.keys(wantedIds).length,
        found: Object.values(items).filter(i => i.found).length,
        archived: Object.values(items).filter(i => i.archived).length,
        missing: Object.values(items).filter(i => !i.found && !i.archived).length
      }
    };
  } catch(e) {
    return { ok: false, error: String(e.message || e) };
  }
}

// ============================================================
// ВОССТАНОВЛЕНИЕ УДАЛЁННЫХ ФУНКЦИЙ + ФИКС alfaUpdate
// ============================================================
//
// При замене функции alfaUpdateGroupAndLesson_ случайно были стёрты
// 18 функций (Telegram + posters cloud + weekJson + draft TG).
//
// Этот файл восстанавливает ВСЁ.
//
// ===== КАК ПРИМЕНИТЬ =====
// 1. Открой Code.gs в Apps Script
// 2. Ctrl+End — перейди в КОНЕЦ файла (увидишь хвост функции
//    с return { ok: true, items: items, ... })
// 3. Поставь пустую строку после закрывающей `}` последней функции
// 4. Скопируй ВЕСЬ ЭТОТ ФАЙЛ (от начала до конца) и вставь
// 5. Ctrl+S → Manage deployments → ✏ → New version → Deploy
//
// Функция alfaUpdateGroupAndLesson_ у тебя в файле УЖЕ есть после
// прошлой замены — её трогать не нужно.
// ============================================================

function tg_getCfg_() {
  const props = PropertiesService.getScriptProperties();
  return {
    token:     props.getProperty('TELEGRAM_BOT_TOKEN') || '',
    channelId: props.getProperty('TELEGRAM_CHANNEL_ID') || ''
  };
}

function tg_apiCall_(method, payload) {
  const cfg = tg_getCfg_();
  if (!cfg.token) throw new Error('TELEGRAM_BOT_TOKEN не задан в Script Properties');
  const url = 'https://api.telegram.org/bot' + cfg.token + '/' + method;
  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  let json;
  try { json = JSON.parse(text); } catch(e) { throw new Error('Telegram вернул не JSON: ' + text.slice(0, 300)); }
  if (code !== 200 || !json.ok) {
    throw new Error('Telegram ' + code + ': ' + (json.description || text.slice(0, 300)));
  }
  return json.result;
}

// === Тест связи с Telegram-ботом (detailsTestTelegram_) ===
function detailsTestTelegram_() {
  try {
    const cfg = tg_getCfg_();
    if (!cfg.token)     return { ok: false, error: 'TELEGRAM_BOT_TOKEN не задан в Script Properties' };
    if (!cfg.channelId) return { ok: false, error: 'TELEGRAM_CHANNEL_ID не задан в Script Properties' };
    // Проверяем токен через getMe
    const me = tg_apiCall_('getMe', {});
    // Проверяем доступ к каналу/чату через getChat
    let chat;
    try {
      chat = tg_apiCall_('getChat', { chat_id: cfg.channelId });
    } catch(e) {
      return {
        ok: false,
        error: 'Бот @' + (me.username || '?') + ' работает, но не может прочитать чат ' + cfg.channelId +
               '. Если это канал — добавьте бота админом. Если личка — отправьте /start от своего аккаунта.\n' +
               'Детали: ' + (e.message || e)
      };
    }
    return {
      ok: true,
      bot: { username: me.username, name: me.first_name },
      chat: { id: chat.id, type: chat.type, title: chat.title || chat.first_name || '' }
    };
  } catch(e) {
    return { ok: false, error: String(e.message || e) };
  }
}

// === Сборка текста поста (tg_buildEventPost_) ===
function tg_buildEventPost_(ev) {
  const monthNamesGen = ['января','февраля','марта','апреля','мая','июня','июля','августа','сентября','октября','ноября','декабря'];
  const dow = ['воскресенье','понедельник','вторник','среда','четверг','пятница','суббота'];
  const d = new Date(ev.date);
  const dateLine = d.getDate() + ' ' + monthNamesGen[d.getMonth()] + ', ' + dow[d.getDay()];
  const time = ev.timeStart ? ev.timeStart : '';

  const lines = [];
  // Заголовок: формат + название
  if (ev.format && ev.name)      lines.push('✨ <b>' + tg_esc_(ev.format) + ': ' + tg_esc_(ev.name) + '</b>');
  else if (ev.name)              lines.push('✨ <b>' + tg_esc_(ev.name) + '</b>');
  else if (ev.format)            lines.push('✨ <b>' + tg_esc_(ev.format) + '</b>');

  // Дата и время
  let when = '📅 ' + dateLine;
  if (time) when += ' · ' + time;
  if (ev.duration) when += ' · ' + ev.duration + ' ч';
  lines.push(when);

  // Педагог
  if (ev.teacher) lines.push('👤 ' + tg_esc_(ev.teacher));

  // Описание (если есть)
  if (ev.description) {
    lines.push('');
    lines.push(tg_esc_(String(ev.description).slice(0, 700)));
  }

  // Цена и места
  const priceLimitParts = [];
  if (ev.price) priceLimitParts.push('💰 ' + ev.price + ' руб.');
  if (ev.limit) priceLimitParts.push('🪑 до ' + ev.limit + ' мест');
  if (priceLimitParts.length > 0) {
    lines.push('');
    lines.push(priceLimitParts.join(' · '));
  }

  // Ссылка на запись
  if (ev.alfaGroupId && CFG && CFG.ALFA_HOST) {
    const formUrl = CFG.ALFA_HOST + '/common/' + CFG.BRANCH_ID + '/lead/create?gid=' + ev.alfaGroupId;
    lines.push('');
    lines.push('👉 <a href="' + formUrl + '">Записаться</a>');
  }

  return lines.join('\n');
}

// === Утилиты Telegram (tg_esc_, tg_findEventById_) ===
function tg_esc_(s) {
  // HTML-режим Telegram требует экранировать только < > &
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

// Находит событие по eventId среди опубликованных + рабочего черновика
function tg_findEventById_(eventId) {
  const allPublished = collectAllPublishedEvents_();
  let found = allPublished.find(e => e.id === eventId);
  if (found) return found;
  const drafts = collectLatestDraftsEvents_();
  found = drafts.find(e => e.id === eventId);
  return found || null;
}

// === Превью поста (detailsPreviewTelegramPost_) ===
function detailsPreviewTelegramPost_(body) {
  const eventId = String(body.eventId || '').trim();
  if (!eventId) return { ok: false, error: 'нет eventId' };
  const ev = tg_findEventById_(eventId);
  if (!ev) return { ok: false, error: 'событие не найдено: ' + eventId };
  const text = tg_buildEventPost_(ev);
  const cfg = tg_getCfg_();
  return {
    ok: true,
    text: text,
    chatId: cfg.channelId || '(не задан)',
    parseMode: 'HTML'
  };
}

// === Постинг текстом одного события (detailsPostToTelegram_) ===
function detailsPostToTelegram_(body) {
  const eventId = String(body.eventId || '').trim();
  const customText = body.customText ? String(body.customText) : null;
  if (!eventId && !customText) return { ok: false, error: 'нужен eventId или customText' };

  const cfg = tg_getCfg_();
  if (!cfg.token || !cfg.channelId) {
    return { ok: false, error: 'Telegram не настроен. Задайте TELEGRAM_BOT_TOKEN и TELEGRAM_CHANNEL_ID в Script Properties.' };
  }

  let text;
  if (customText) {
    text = customText;
  } else {
    const ev = tg_findEventById_(eventId);
    if (!ev) return { ok: false, error: 'событие не найдено: ' + eventId };
    text = tg_buildEventPost_(ev);
  }

  try {
    const result = tg_apiCall_('sendMessage', {
      chat_id: cfg.channelId,
      text: text,
      parse_mode: 'HTML',
      disable_web_page_preview: false,
      link_preview_options: { is_disabled: false }
    });
    // Пробуем построить ссылку на пост (только для каналов с username)
    let postUrl = '';
    if (result.chat && result.chat.username) {
      postUrl = 'https://t.me/' + result.chat.username + '/' + result.message_id;
    }
    return {
      ok: true,
      messageId: result.message_id,
      chatId: result.chat ? result.chat.id : cfg.channelId,
      postUrl: postUrl,
      text: text
    };
  } catch(e) {
    return { ok: false, error: String(e.message || e) };
  }
}

// === JSON недели для агента афиш (detailsGetWeekJson_) ===
function detailsGetWeekJson_(params) {
  const weekStart = String((params && params.weekStart) || '').trim();
  if (!weekStart || !/^\d{4}-\d{2}-\d{2}$/.test(weekStart)) {
    return { ok: false, error: 'нужен weekStart в формате YYYY-MM-DD' };
  }
  // weekEnd = weekStart + 6 дней
  const startDate = new Date(weekStart + 'T00:00:00');
  const endDate = new Date(startDate.getTime() + 6 * 24 * 60 * 60 * 1000);
  const weekEnd = Utilities.formatDate(endDate, 'Europe/Moscow', 'yyyy-MM-dd');

  // Расширенный список приватных форматов (закрытые мероприятия — не для афиш)
  const PRIVATE_FOR_POSTERS = [
    'День рождения', 'Свидание', 'Частное мероприятие',
    'Свадьба', 'Корпоратив', 'Закрытое'
  ];

  // 1. Собираем кандидатов из обоих источников (опубликованные + рабочий черновик)
  const published = collectAllPublishedEvents_();
  const drafts = collectLatestDraftsEvents_();
  const byId = {};
  drafts.forEach(ev => { byId[ev.id] = ev; });
  published.forEach(ev => { byId[ev.id] = ev; });
  let candidates = Object.values(byId);

  // 2. Базовый фильтр: неделя + не архив + не приватный + содержательный
  candidates = candidates.filter(ev => {
    if (!ev.date) return false;
    if (ev.status === 'archived') return false;
    if (PRIVATE_FOR_POSTERS.indexOf(ev.format) >= 0) return false;
    if (ev.date < weekStart || ev.date > weekEnd) return false;
    // Отсекаем "пустышки": нет ни имени, ни описания (или формат "-")
    const fmt = String(ev.format || '').trim();
    const name = String(ev.name || '').trim();
    const desc = String(ev.description || '').trim();
    if (fmt === '-' || fmt === '') return false;
    if (!name && !desc) return false;
    if (name === '-' && !desc) return false;
    return true;
  });

  // 3. Дедупликация: группируем по date+timeStart, оставляем "лучшее" событие
  // "Лучшее" = больше всего полезных данных:
  //   alfa-привязка (alfaGroupId) +20, описание +len/100, референсы +кол-во*5, педагог +5
  const scoreEvent = ev => {
    let score = 0;
    if (ev.alfaGroupId) score += 20;
    if (ev.description) score += Math.min(50, String(ev.description).length / 20);
    if (ev.references && ev.references.length) score += ev.references.length * 5;
    if (ev.teacher) score += 5;
    if (ev.price) score += 2;
    if (ev.id && String(ev.id).startsWith('alfa_')) score += 10; // приоритет alfa-импортов
    return score;
  };

  const groups = {};
  candidates.forEach(ev => {
    const key = ev.date + '|' + (ev.timeStart || '');
    if (!groups[key]) groups[key] = [];
    groups[key].push(ev);
  });

  let events = [];
  Object.keys(groups).forEach(key => {
    const arr = groups[key];
    if (arr.length === 1) {
      events.push(arr[0]);
    } else {
      // Сортируем по score убывая, берём первое
      arr.sort((a, b) => scoreEvent(b) - scoreEvent(a));
      const winner = arr[0];
      // Но если у "проигравшего" есть данные, которых нет у победителя — мерджим
      for (let i = 1; i < arr.length; i++) {
        const loser = arr[i];
        if (!winner.description && loser.description) winner.description = loser.description;
        if (!winner.teacher && loser.teacher) winner.teacher = loser.teacher;
        if (!winner.name && loser.name) winner.name = loser.name;
        if (!winner.format && loser.format) winner.format = loser.format;
        if ((!winner.references || winner.references.length === 0) && loser.references && loser.references.length) {
          winner.references = loser.references;
        }
      }
      events.push(winner);
    }
  });

  // 4. Сортируем
  events.sort((a, b) => (a.date + (a.timeStart || '')).localeCompare(b.date + (b.timeStart || '')));

  // 5. Преобразуем в формат для агента
  const dowNames = ['воскресенье','понедельник','вторник','среда','четверг','пятница','суббота'];
  const out = events.map(ev => {
    const d = new Date(ev.date + 'T00:00:00');
    return {
      id: ev.id,
      date: ev.date,
      dayOfWeek: dowNames[d.getDay()],
      format: ev.format || '',
      name: ev.name || '',
      teacher: ev.teacher || '',
      time: ev.timeStart || '',
      duration: ev.duration || null,
      price: ev.price || null,
      limit: ev.limit || null,
      description: ev.description || '',
      references: (ev.references || []).filter(r => r && r.url).map(r => ({
        url: r.url,
        note: r.note || ''
      })),
      teamComment: ev.teamComment || ''
    };
  });

  return {
    ok: true,
    weekStart: weekStart,
    weekEnd: weekEnd,
    eventsCount: out.length,
    events: out,
    audience: 'Женщины 25-48 из Могилева, ищут состояние и атмосферу, не услугу',
    style: 'Уютно, тепло, по-человечески. Без пафоса. Без штампов вроде "незабываемый вечер творчества".'
  };
}

// === Постинг фото из библиотеки (detailsPostPhotoToTelegram_) ===
function detailsPostPhotoToTelegram_(body) {
  const cfg = tg_getCfg_();
  if (!cfg.token) return { ok: false, error: 'TELEGRAM_BOT_TOKEN не задан' };
  const chatId = String(body.chatId || cfg.channelId || '').trim();
  if (!chatId) return { ok: false, error: 'chatId не задан' };

  let photoBase64 = String(body.photoBase64 || '').trim();
  if (!photoBase64) return { ok: false, error: 'photoBase64 не задан' };
  // Убираем префикс data: если есть
  const m = photoBase64.match(/^data:image\/[a-z]+;base64,(.+)$/i);
  if (m) photoBase64 = m[1];

  const caption = body.caption ? String(body.caption).slice(0, 1024) : ''; // Telegram caption limit

  let blob;
  try {
    const decoded = Utilities.base64Decode(photoBase64);
    blob = Utilities.newBlob(decoded, 'image/png', 'poster.png');
  } catch(e) {
    return { ok: false, error: 'Не удалось декодировать base64: ' + (e.message || e) };
  }

  const url = 'https://api.telegram.org/bot' + cfg.token + '/sendPhoto';
  const formData = {
    chat_id: chatId,
    photo: blob,
    parse_mode: 'HTML'
  };
  if (caption) formData.caption = caption;

  try {
    const resp = UrlFetchApp.fetch(url, {
      method: 'post',
      payload: formData,
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    const text = resp.getContentText();
    let json;
    try { json = JSON.parse(text); } catch(e) {
      return { ok: false, error: 'Telegram вернул не JSON: ' + text.slice(0, 300) };
    }
    if (code !== 200 || !json.ok) {
      return { ok: false, error: 'Telegram ' + code + ': ' + (json.description || text.slice(0, 300)) };
    }
    let postUrl = '';
    if (json.result && json.result.chat && json.result.chat.username) {
      postUrl = 'https://t.me/' + json.result.chat.username + '/' + json.result.message_id;
    }
    return {
      ok: true,
      messageId: json.result.message_id,
      chatId: json.result.chat ? json.result.chat.id : chatId,
      postUrl: postUrl
    };
  } catch(e) {
    return { ok: false, error: String(e.message || e) };
  }
}

// === Облако постеров (postersList/Get/Save/Delete) ===
function detailsPostersList_() {
  try {
    const sh = getPostersSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, items: [] };
    // Читаем только первые 7 столбцов (без payload)
    const data = sh.getRange(2, 1, lastRow - 1, 7).getValues();
    const items = data
      .filter(r => r[0]) // пропуск пустых
      .map(r => ({
        id:           String(r[0]),
        saved_at:     r[1] ? String(r[1]) : '',
        event_name:   String(r[2] || ''),
        event_date:   r[3] ? String(r[3]) : '',
        layout:       String(r[4] || ''),
        accent:       String(r[5] || ''),
        thumbnail_url: String(r[6] || '')
      }));
    items.sort((a, b) => (b.saved_at || '').localeCompare(a.saved_at || ''));
    return { ok: true, items };
  } catch(e) {
    return { ok: false, error: String(e.message || e) };
  }
}

/**
 * GET ?action=posterGet&id=<id>
 * Возвращает один постер с полным payload.
 */
function detailsPosterGet_(params) {
  try {
    const id = String((params && params.id) || '').trim();
    if (!id) return { ok: false, error: 'нет id' };
    const sh = getPostersSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: false, error: 'не найдено' };
    const data = sh.getRange(2, 1, lastRow - 1, 8).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === id) {
        let payload = null;
        try { payload = JSON.parse(String(data[i][7] || 'null')); } catch(e) {}
        return {
          ok: true,
          item: {
            id:            String(data[i][0]),
            saved_at:      data[i][1] ? String(data[i][1]) : '',
            event_name:    String(data[i][2] || ''),
            event_date:    data[i][3] ? String(data[i][3]) : '',
            layout:        String(data[i][4] || ''),
            accent:        String(data[i][5] || ''),
            thumbnail_url: String(data[i][6] || ''),
            payload:       payload
          }
        };
      }
    }
    return { ok: false, error: 'не найдено' };
  } catch(e) {
    return { ok: false, error: String(e.message || e) };
  }
}

/**
 * POST { action: 'posterSave', id, savedAt, eventName, eventDate, layout, accent, thumbnailUrl, payload }
 * Сохраняет/обновляет постер. Если id уже есть — обновляет, иначе создаёт.
 */
function detailsPosterSave_(body) {
  try {
    const id = String(body.id || '').trim();
    if (!id) return { ok: false, error: 'нет id' };

    const sh = getPostersSheet_();
    const payloadJson = body.payload ? JSON.stringify(body.payload) : '';

    // Apps Script лимит ячейки — 50_000 символов. Защищаемся.
    if (payloadJson.length > 49000) {
      return { ok: false, error: 'payload слишком большой (' + payloadJson.length + ' символов, лимит 49000)' };
    }

    const rowData = [
      id,
      String(body.savedAt || nowIso_()),
      String(body.eventName || ''),
      String(body.eventDate || ''),
      String(body.layout || ''),
      String(body.accent || ''),
      String(body.thumbnailUrl || ''),
      payloadJson
    ];

    // Ищем существующую строку
    const lastRow = sh.getLastRow();
    let foundRow = -1;
    if (lastRow >= 2) {
      const idValues = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < idValues.length; i++) {
        if (String(idValues[i][0]) === id) {
          foundRow = i + 2;
          break;
        }
      }
    }

    if (foundRow > 0) {
      sh.getRange(foundRow, 1, 1, 8).setValues([rowData]);
      return { ok: true, id, updated: true };
    } else {
      sh.appendRow(rowData);
      return { ok: true, id, created: true };
    }
  } catch(e) {
    return { ok: false, error: String(e.message || e) };
  }
}

/**
 * POST { action: 'posterDelete', id }
 * Удаляет постер из облака.
 */
function detailsPosterDelete_(body) {
  try {
    const id = String(body.id || '').trim();
    if (!id) return { ok: false, error: 'нет id' };
    const sh = getPostersSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, deleted: false };
    const idValues = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < idValues.length; i++) {
      if (String(idValues[i][0]) === id) {
        sh.deleteRow(i + 2);
        return { ok: true, deleted: true };
      }
    }
    return { ok: true, deleted: false };
  } catch(e) {
    return { ok: false, error: String(e.message || e) };
  }
}

// ============================================================
// БЛОК 2: «На обсуждение» — черновик события в Telegram
// ============================================================
function tg_buildEventDraftPost_(ev) {
  const monthNamesGen = ['января','февраля','марта','апреля','мая','июня','июля','августа','сентября','октября','ноября','декабря'];
  const dow = ['воскресенье','понедельник','вторник','среда','четверг','пятница','суббота'];
  const d = ev.date ? new Date(ev.date) : null;

  const lines = [];

  // Шапка
  lines.push('💡 <b>НА ОБСУЖДЕНИЕ</b>');
  lines.push('');

  // Формат + название
  if (ev.format && ev.name)      lines.push('✨ <b>' + tg_esc_(ev.format) + ': ' + tg_esc_(ev.name) + '</b>');
  else if (ev.name)              lines.push('✨ <b>' + tg_esc_(ev.name) + '</b>');
  else if (ev.format)            lines.push('✨ <b>' + tg_esc_(ev.format) + '</b>');

  // Дата и время
  if (d) {
    const dateLine = d.getDate() + ' ' + monthNamesGen[d.getMonth()] + ', ' + dow[d.getDay()];
    let when = '📅 ' + dateLine;
    if (ev.timeStart) when += ' · ' + ev.timeStart;
    if (ev.duration)  when += ' · ' + ev.duration + ' ч';
    lines.push(when);
  }

  // Педагог
  if (ev.teacher) lines.push('👤 ' + tg_esc_(ev.teacher));

  // Описание (длинное, для команды — без обрезки до 700 как в обычном посте)
  if (ev.description) {
    lines.push('');
    lines.push(tg_esc_(String(ev.description).slice(0, 2000)));
  }

  // Цена и места
  const priceLimitParts = [];
  if (ev.price) priceLimitParts.push('💰 ' + ev.price + ' руб.');
  if (ev.limit) priceLimitParts.push('🪑 до ' + ev.limit + ' мест');
  if (priceLimitParts.length > 0) {
    lines.push('');
    lines.push(priceLimitParts.join(' · '));
  }

  // Краткое для Альфы (если задано отдельно от описания)
  if (ev.alfaNote && ev.alfaNote !== ev.description) {
    lines.push('');
    lines.push('📎 <i>Для Альфы:</i> ' + tg_esc_(String(ev.alfaNote).slice(0, 400)));
  }

  // Референсы
  const refs = (ev.references || []).filter(r => r && r.url);
  if (refs.length > 0) {
    lines.push('');
    lines.push('🔖 <b>Референсы (' + refs.length + '):</b>');
    refs.slice(0, 10).forEach((r, i) => {
      const domain = String(r.domain || '').replace(/^www\./, '');
      const note = String(r.note || '').trim();
      let line = '• <a href="' + tg_esc_(r.url) + '">' + tg_esc_(domain || 'ссылка') + '</a>';
      if (note) line += ' — ' + tg_esc_(note);
      lines.push(line);
    });
    if (refs.length > 10) lines.push('… и ещё ' + (refs.length - 10));
  }

  // Комментарий команде (если есть)
  if (ev.teamComment) {
    lines.push('');
    lines.push('💬 <i>Комментарий:</i> ' + tg_esc_(String(ev.teamComment).slice(0, 500)));
  }

  // Финальный вопрос
  lines.push('');
  lines.push('❓ <b>Есть предложения / пожелания / стиль / цвет?</b>');

  return lines.join('\n');
}

// === Превью поста-черновика ===
function detailsPreviewEventDraft_(body) {
  const eventId = String(body.eventId || '').trim();
  if (!eventId) return { ok: false, error: 'нет eventId' };

  // Сначала ищем в working_main / опубликованных, потом — в body.event (если фронт прислал свежее)
  let ev = tg_findEventById_(eventId);

  // Если фронт прислал свежие данные (можно перетереть несохранённые правки)
  if (body.event && typeof body.event === 'object') {
    ev = body.event;
  }

  if (!ev) return { ok: false, error: 'событие не найдено: ' + eventId };

  // Резолвим thumb-маркеры в реальные data URL для отправки фото
  if (Array.isArray(ev.references)) {
    try {
      const tmap = thumbCacheLoadAll_();
      ev.references = resolveThumbs_(ev.references, tmap);
    } catch(e) {}
  }

  const text = tg_buildEventDraftPost_(ev);
  const cfg = tg_getCfg_();

  // Сразу собираем список референсов с готовыми превью (для модалки на фронте)
  const refsForUI = (ev.references || []).filter(r => r && r.url).map(r => ({
    url: r.url,
    note: r.note || '',
    thumb: r.thumb || '',
    domain: r.domain || ''
  }));

  return {
    ok: true,
    text: text,
    refs: refsForUI,
    chatId: cfg.channelId || '(не задан)',
    parseMode: 'HTML'
  };
}

// === Отправка черновика в Telegram ===
//
// body.eventId — обязательно
// body.event — опционально, свежие данные с фронта (несохранённые правки)
// body.refUrlsToSend — массив URL референсов которые отправить как фото (max 10)
//                     если пусто — отправляем только текст
//
function detailsPostEventDraftToTelegram_(body) {
  const cfg = tg_getCfg_();
  if (!cfg.token)     return { ok: false, error: 'TELEGRAM_BOT_TOKEN не задан' };
  if (!cfg.channelId) return { ok: false, error: 'TELEGRAM_CHANNEL_ID не задан' };

  const eventId = String(body.eventId || '').trim();
  if (!eventId) return { ok: false, error: 'нет eventId' };

  // Берём свежие данные с фронта если переданы
  let ev = body.event && typeof body.event === 'object' ? body.event : tg_findEventById_(eventId);
  if (!ev) return { ok: false, error: 'событие не найдено: ' + eventId };

  // Резолвим thumb-маркеры в data URL — нужно для отправки фото
  if (Array.isArray(ev.references)) {
    try {
      const tmap = thumbCacheLoadAll_();
      ev.references = resolveThumbs_(ev.references, tmap);
    } catch(e) {}
  }

  const text = tg_buildEventDraftPost_(ev);
  const refUrlsToSend = Array.isArray(body.refUrlsToSend) ? body.refUrlsToSend.slice(0, 10) : [];

  try {
    // === ВАРИАНТ 1: только текст (без фото) ===
    if (refUrlsToSend.length === 0) {
      const resp = tg_apiCall_('sendMessage', {
        chat_id: cfg.channelId,
        text: text,
        parse_mode: 'HTML',
        disable_web_page_preview: true
      });
      return {
        ok: true,
        mode: 'text',
        messageId: resp.message_id,
        message: 'Отправлено в «Учительскую» (только текст)'
      };
    }

    // === ВАРИАНТ 2: с фото через sendMediaGroup ===
    // Собираем фото из переданных URL — для каждого ищем data URL в референсах события
    const refsByUrl = {};
    (ev.references || []).forEach(r => { if (r && r.url) refsByUrl[r.url] = r; });

    const mediaList = [];
    const failedUrls = [];

    refUrlsToSend.forEach(url => {
      const r = refsByUrl[url];
      if (!r) return;
      const thumb = String(r.thumb || '');
      // Подходят: data:image/...;base64,... или внешние http(s) URL
      if (thumb.indexOf('data:image') === 0) {
        // base64 data URL — Telegram Bot API НЕ принимает напрямую через JSON в sendMediaGroup
        // Нужно отправлять как multipart с attached file. Это сложно через UrlFetchApp.
        // Альтернатива: использовать оригинальный URL картинки если он публичный.
        // Pinterest URL'ы — публичные (i.pinimg.com), Telegram скачает их сам.
        // Поэтому для media group используем url, не thumb.
        if (url && /^https?:\/\//i.test(url)) {
          mediaList.push({ url: url, note: r.note || '' });
        } else {
          failedUrls.push(url);
        }
      } else if (/^https?:\/\//i.test(thumb)) {
        // thumb уже URL — можно отправить
        mediaList.push({ url: thumb, note: r.note || '' });
      } else if (/^https?:\/\//i.test(url)) {
        // У thumb нет URL, но у url-источника есть — отправим источник
        mediaList.push({ url: url, note: r.note || '' });
      } else {
        failedUrls.push(url);
      }
    });

    if (mediaList.length === 0) {
      // Не получилось ни одного фото — fallback на текст
      const resp = tg_apiCall_('sendMessage', {
        chat_id: cfg.channelId,
        text: text,
        parse_mode: 'HTML',
        disable_web_page_preview: true
      });
      return {
        ok: true,
        mode: 'text-fallback',
        messageId: resp.message_id,
        warning: 'Не удалось прикрепить фото (нет публичных URL у референсов), отправлено только текстом',
        message: 'Отправлено в «Учительскую» (только текст)'
      };
    }

    // Telegram caption лимит: 1024 символа для media с caption.
    // Если текст длиннее — отправим media без caption + отдельный sendMessage с текстом после.
    const CAPTION_LIMIT = 1024;
    let captionForFirst = '';
    let separateMessage = '';

    if (text.length <= CAPTION_LIMIT) {
      captionForFirst = text;
    } else {
      // Caption будет коротким (название + дата + педагог), полный текст — отдельно
      const shortLines = [];
      shortLines.push('💡 <b>НА ОБСУЖДЕНИЕ</b>');
      if (ev.format && ev.name) shortLines.push('<b>' + tg_esc_(ev.format) + ': ' + tg_esc_(ev.name) + '</b>');
      else if (ev.name)         shortLines.push('<b>' + tg_esc_(ev.name) + '</b>');
      if (ev.date) {
        const monthNamesGen = ['января','февраля','марта','апреля','мая','июня','июля','августа','сентября','октября','ноября','декабря'];
        const d = new Date(ev.date);
        let when = '📅 ' + d.getDate() + ' ' + monthNamesGen[d.getMonth()];
        if (ev.timeStart) when += ' · ' + ev.timeStart;
        shortLines.push(when);
      }
      shortLines.push('');
      shortLines.push('↓ подробности и обсуждение ниже ↓');
      captionForFirst = shortLines.join('\n').slice(0, CAPTION_LIMIT);
      separateMessage = text;
    }

    // Формируем массив media для sendMediaGroup
    const media = mediaList.map((m, i) => {
      const item = {
        type: 'photo',
        media: m.url
      };
      if (i === 0 && captionForFirst) {
        item.caption = captionForFirst;
        item.parse_mode = 'HTML';
      }
      return item;
    });

    // Отправляем медиа-группу
    let mediaResult;
    try {
      mediaResult = tg_apiCall_('sendMediaGroup', {
        chat_id: cfg.channelId,
        media: media
      });
    } catch(e) {
      // Если sendMediaGroup провалился (например Telegram не смог скачать картинки) —
      // отправляем хотя бы текст
      const resp = tg_apiCall_('sendMessage', {
        chat_id: cfg.channelId,
        text: text,
        parse_mode: 'HTML',
        disable_web_page_preview: true
      });
      return {
        ok: true,
        mode: 'text-fallback',
        messageId: resp.message_id,
        warning: 'Telegram не смог загрузить фото: ' + (e.message || e) + '. Отправлено только текстом.',
        message: 'Отправлено в «Учительскую» (только текст)'
      };
    }

    // Если был длинный текст — отправляем его отдельным сообщением
    let textMessageId = null;
    if (separateMessage) {
      try {
        const r2 = tg_apiCall_('sendMessage', {
          chat_id: cfg.channelId,
          text: separateMessage,
          parse_mode: 'HTML',
          disable_web_page_preview: true,
          reply_to_message_id: Array.isArray(mediaResult) && mediaResult[0] ? mediaResult[0].message_id : undefined
        });
        textMessageId = r2.message_id;
      } catch(e) { /* не критично */ }
    }

    return {
      ok: true,
      mode: 'media-group',
      photosCount: media.length,
      mediaMessageIds: Array.isArray(mediaResult) ? mediaResult.map(m => m.message_id) : [],
      textMessageId: textMessageId,
      failedCount: failedUrls.length,
      message: 'Отправлено в «Учительскую» (' + media.length + ' фото' + (separateMessage ? ' + текст' : '') + ')'
    };

  } catch(e) {
    return { ok: false, error: String(e.message || e) };
  }
}
