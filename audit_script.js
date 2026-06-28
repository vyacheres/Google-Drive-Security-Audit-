/**
 * Google Drive Security Audit v2.0 — Google Apps Script
 * Аудит прав доступа к Google Диску · Google Drive permissions audit
 *
 * @fileoverview
 * EN: Scans My Drive and Shared Drives (if Advanced Drive Service is enabled),
 *     flags public / link sharing, lists editors and viewers with risk levels,
 *     writes colour-coded results to three tabs, sends an e-mail summary,
 *     supports resumable scans for large drives via PropertiesService.
 *
 * RU: Обходит Мой диск и Общие диски (если включён Advanced Drive Service),
 *     отмечает публичный / ссылочный доступ, выводит редакторов и читателей
 *     с уровнем риска, пишет цветной отчёт на три листа, отправляет
 *     email-сводку и поддерживает возобновляемый обход больших дисков.
 *
 * SETUP / НАСТРОЙКА:
 *   1. Extensions → Apps Script → paste this file → Save
 *   2. (Optional) Services → add "Drive API v2" for Shared Drives + resumable scan
 *   3. Reload the Sheet (F5) → Security menu appears
 *   4. Run "Start Audit" and grant permissions
 */

// ---------------------------------------------------------------------------
// Configuration / Конфигурация
// ---------------------------------------------------------------------------

/**
 * EN: Organization domain without "@" (e.g. "company.com").
 *     If set, editor/viewer rows include only emails whose domain differs.
 *     Public rows are always included regardless of this setting.
 *     Leave '' to include all collaborators (full ACL inventory).
 *
 * RU: Домен организации без "@" (например "company.com").
 *     Если задан — строки редактор/читатель только для адресов вне домена.
 *     Публичные строки выводятся всегда.
 *     Оставьте '' чтобы включать всех участников.
 */
const ORG_DOMAIN = '';

/**
 * EN: E-mail address for the risk summary notification.
 *     Leave '' to use the account that runs the script.
 * RU: Адрес для email-уведомления о рисках.
 *     Оставьте '' — письмо придёт на аккаунт, запустивший скрипт.
 */
const ADMIN_EMAIL = '';

// ---------------------------------------------------------------------------
// Constants / Константы
// ---------------------------------------------------------------------------

const SHEETS = {
  PUBLIC:   '🔴 Публичные',
  EXTERNAL: '🟡 Внешние',
  FULL:     '📋 Полный отчёт',
};

const RISK_COLORS = {
  critical: '#f4cccc', // red   — anyone on the web
  high:     '#fce5cd', // orange — anyone with link
  medium:   '#fff2cc', // yellow — external collaborator
  low:      '#d9ead3', // green  — internal collaborator
  unknown:  '#f3f3f3', // grey   — read error
};

const HEADER_BG    = '#1155cc';
const HEADER_FONT  = '#ffffff';
const COLUMNS      = ['Название файла', 'Тип доступа', 'Уровень риска',
                      'Email с доступом', 'Владелец', 'Последнее изменение', 'Ссылка'];

// Keys used in PropertiesService for resumable scan
const PROP_TOKEN   = 'audit_page_token';
const PROP_ROWS    = 'audit_rows';
const PROP_ACTIVE  = 'audit_in_progress';

// Stop processing new pages after this many ms to stay within the 6-min limit
const SAFE_RUNTIME_MS = 270000; // 4.5 minutes

// ---------------------------------------------------------------------------
// Public entry points / Публичные точки входа
// ---------------------------------------------------------------------------

/**
 * EN: Adds the Security menu when the spreadsheet is opened.
 * RU: Создаёт меню «Безопасность» при открытии таблицы.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛡️ Безопасность')
    .addItem('▶ Запустить аудит прав',            'driveAudit')
    .addSeparator()
    .addItem('⏰ Настроить еженедельный аудит',    'setupWeeklyTrigger')
    .addItem('🗑 Удалить еженедельный аудит',      'removeWeeklyTrigger')
    .addToUi();
}

/**
 * EN: Starts a fresh audit (resets any previous in-progress state).
 * RU: Запускает новый аудит (сбрасывает незавершённый предыдущий).
 */
function driveAudit() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_TOKEN);
  props.deleteProperty(PROP_ROWS);
  props.deleteProperty(PROP_ACTIVE);

  setupSheets_();
  runAuditChunk_();
}

/**
 * EN: Internal continuation handler called by the time-based trigger.
 * RU: Обработчик продолжения, вызываемый триггером по времени.
 */
function continueAudit_() {
  deleteTriggersByHandler_('continueAudit_');
  runAuditChunk_();
}

/**
 * EN: Creates a weekly Monday 9:00 trigger. Removes any previous one first.
 * RU: Создаёт еженедельный триггер на понедельник 9:00. Сначала удаляет старый.
 */
function setupWeeklyTrigger() {
  deleteTriggersByHandler_('driveAudit');
  ScriptApp.newTrigger('driveAudit')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
  toast_('⏰ Триггер создан', 'Еженедельный аудит — каждый понедельник в 09:00');
}

/**
 * EN: Removes all weekly audit triggers.
 * RU: Удаляет все еженедельные триггеры аудита.
 */
function removeWeeklyTrigger() {
  deleteTriggersByHandler_('driveAudit');
  toast_('🗑 Триггер удалён', 'Автоматический запуск отключён');
}

// ---------------------------------------------------------------------------
// Core audit logic / Основная логика аудита
// ---------------------------------------------------------------------------

/**
 * EN: Processes one chunk of files. Schedules continuation if time runs out.
 * RU: Обрабатывает один блок файлов. Планирует продолжение при нехватке времени.
 */
function runAuditChunk_() {
  const startTime = Date.now();
  const props     = PropertiesService.getScriptProperties();
  const rows      = JSON.parse(props.getProperty(PROP_ROWS) || '[]');

  // --- Try Advanced Drive Service (Drive API v2) ---
  // Enables resumable scan and Shared Drives support.
  if (typeof Drive !== 'undefined') {
    let pageToken = props.getProperty(PROP_TOKEN) || undefined;

    while (true) {
      const params = {
        pageSize:                 200,
        fields:                   'nextPageToken, items(id,title,alternateLink,owners,modifiedDate)',
        includeItemsFromAllDrives: true,
        supportsAllDrives:         true,
        corpora:                   'allDrives',
      };
      if (pageToken) params.pageToken = pageToken;

      let response;
      try {
        response = Drive.Files.list(params);
      } catch (apiErr) {
        // API quota or transient error — save and retry later
        props.setProperty(PROP_TOKEN, pageToken || '');
        props.setProperty(PROP_ROWS,  JSON.stringify(rows));
        scheduleContinuation_();
        return;
      }

      (response.items || []).forEach(function (meta) {
        processFileMeta_(meta, rows);
      });

      pageToken = response.nextPageToken;

      if (!pageToken) break; // done

      if (Date.now() - startTime > SAFE_RUNTIME_MS) {
        // Save progress and continue in a new execution
        props.setProperty(PROP_TOKEN, pageToken);
        props.setProperty(PROP_ROWS,  JSON.stringify(rows));
        scheduleContinuation_();
        toast_('⏳ Аудит продолжается…', 'Обработано файлов: ' + rows.length);
        return;
      }
    }

  } else {
    // --- Fallback: DriveApp (My Drive only, not resumable) ---
    const iter = DriveApp.getFiles();
    while (iter.hasNext()) {
      const file = iter.next();
      processFile_(file, null, rows);

      if (Date.now() - startTime > SAFE_RUNTIME_MS) {
        // Cannot resume without a page token — write what we have and warn
        toast_('⚠️ Аудит остановлен по лимиту времени',
               'Включите Advanced Drive Service для возобновляемого обхода. ' +
               'Обработано: ' + rows.length);
        finalizeAudit_(rows);
        return;
      }
    }
  }

  // All pages consumed
  props.deleteProperty(PROP_TOKEN);
  props.deleteProperty(PROP_ROWS);
  props.deleteProperty(PROP_ACTIVE);
  finalizeAudit_(rows);
}

/**
 * EN: Processes a single file retrieved via Drive API v2 metadata object.
 * RU: Обрабатывает файл по метаданным из Drive API v2.
 *
 * @param {{id:string, title:string, alternateLink:string, owners:Array, modifiedDate:string}} meta
 * @param {Array} rows
 */
function processFileMeta_(meta, rows) {
  try {
    const file = DriveApp.getFileById(meta.id);
    processFile_(file, meta, rows);
  } catch (err) {
    rows.push(makeRow_(
      meta.title || '?',
      'Ошибка чтения',
      'unknown',
      String(err.message || err),
      '?',
      '?',
      meta.alternateLink || '?'
    ));
  }
}

/**
 * EN: Inspects one file's sharing settings and pushes result rows.
 * RU: Проверяет настройки доступа одного файла и добавляет строки в отчёт.
 *
 * @param {GoogleAppsScript.Drive.File} file
 * @param {Object|null} meta  Drive API metadata (may be null in DriveApp fallback)
 * @param {Array} rows
 */
function processFile_(file, meta, rows) {
  const name     = file.getName();
  const url      = (meta && meta.alternateLink) || file.getUrl();
  const modified = formatDate_(file.getLastUpdated());
  const owner    = ownerEmail_(file);
  const access   = file.getSharingAccess();

  if (access === DriveApp.Access.ANYONE) {
    rows.push(makeRow_(name, 'ПУБЛИЧНЫЙ (все)',        'critical', '—', owner, modified, url));
  }
  if (access === DriveApp.Access.ANYONE_WITH_LINK) {
    rows.push(makeRow_(name, 'ПУБЛИЧНЫЙ (по ссылке)', 'high',     '—', owner, modified, url));
  }

  file.getEditors().forEach(function (u) {
    const email = u.getEmail();
    if (!includeCollaborator_(email)) return;
    rows.push(makeRow_(name, 'Редактор', riskForEmail_(email), email || '(без email)', owner, modified, url));
  });

  file.getViewers().forEach(function (u) {
    const email = u.getEmail();
    if (!includeCollaborator_(email)) return;
    // Viewers carry slightly lower risk than editors, still flag if external
    const risk = isExternal_(email) ? 'medium' : 'low';
    rows.push(makeRow_(name, 'Читатель', risk, email || '(без email)', owner, modified, url));
  });
}

// ---------------------------------------------------------------------------
// Output / Вывод результатов
// ---------------------------------------------------------------------------

/**
 * EN: Writes results to all three sheets, applies colours, notifies by e-mail.
 * RU: Пишет результаты на три листа, раскрашивает строки, отправляет email.
 *
 * @param {Array} rows
 */
function finalizeAudit_(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const publicRows   = rows.filter(function (r) { return r[2] === 'critical' || r[2] === 'high'; });
  const externalRows = rows.filter(function (r) { return isExternal_(r[3]); });

  writeToSheet_(ss.getSheetByName(SHEETS.PUBLIC),   publicRows);
  writeToSheet_(ss.getSheetByName(SHEETS.EXTERNAL), externalRows);
  writeToSheet_(ss.getSheetByName(SHEETS.FULL),     rows);

  sendNotification_(rows);
  toast_('✅ Аудит завершён', 'Найдено записей: ' + rows.length +
         ' · Публичных файлов: ' + publicRows.length);
}

/**
 * EN: Writes rows to a sheet and applies per-row background colours.
 * RU: Записывает строки на лист и красит фон по уровню риска.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array} rows
 */
function writeToSheet_(sheet, rows) {
  if (!sheet) return;

  if (!rows.length) {
    sheet.getRange(2, 1).setValue('(нет данных)');
    return;
  }

  // Replace raw risk key with human label before writing
  const display = rows.map(function (r) {
    return [r[0], r[1], riskLabel_(r[2]), r[3], r[4], r[5], r[6]];
  });

  const dataRange = sheet.getRange(2, 1, display.length, COLUMNS.length);
  dataRange.setValues(display);

  // Colour each row based on risk level stored in original rows
  rows.forEach(function (r, i) {
    const color = RISK_COLORS[r[2]] || '#ffffff';
    sheet.getRange(i + 2, 1, 1, COLUMNS.length).setBackground(color);
  });

  sheet.autoResizeColumns(1, COLUMNS.length);
}

/**
 * EN: Sends an e-mail summary when critical or high-risk files are found.
 * RU: Отправляет email-сводку при обнаружении критичных или высокорисковых файлов.
 *
 * @param {Array} rows
 */
function sendNotification_(rows) {
  const critical = rows.filter(function (r) { return r[2] === 'critical'; }).length;
  const high     = rows.filter(function (r) { return r[2] === 'high'; }).length;
  if (critical + high === 0) return;

  const to      = ADMIN_EMAIL || Session.getEffectiveUser().getEmail();
  const subject = '[Drive Audit] ' + (critical + high) + ' файлов с высоким риском доступа';
  const body    = [
    'Результаты автоматического аудита Google Drive:',
    '',
    '🔴 Критичных (публично для всех):  ' + critical,
    '🟠 Высоких (доступ по ссылке):     ' + high,
    '📋 Всего записей в отчёте:         ' + rows.length,
    '',
    'Откройте таблицу для просмотра деталей и принятия мер.',
  ].join('\n');

  try {
    MailApp.sendEmail(to, subject, body);
  } catch (e) {
    console.warn('Email notification failed: ' + e);
  }
}

// ---------------------------------------------------------------------------
// Sheet initialisation / Подготовка листов
// ---------------------------------------------------------------------------

/**
 * EN: Creates (or clears) the three result tabs and writes headers.
 * RU: Создаёт (или очищает) три листа отчёта и записывает заголовки.
 */
function setupSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.values(SHEETS).forEach(function (name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    } else {
      sheet.clear();
    }
    const hdr = sheet.getRange(1, 1, 1, COLUMNS.length);
    hdr.setValues([COLUMNS]);
    hdr.setBackground(HEADER_BG);
    hdr.setFontColor(HEADER_FONT);
    hdr.setFontWeight('bold');
    sheet.setFrozenRows(1);
  });
}

// ---------------------------------------------------------------------------
// Helpers / Вспомогательные функции
// ---------------------------------------------------------------------------

/**
 * @param {string} name
 * @param {string} accessType
 * @param {string} risk
 * @param {string} email
 * @param {string} owner
 * @param {string} modified
 * @param {string} url
 * @returns {Array}
 */
function makeRow_(name, accessType, risk, email, owner, modified, url) {
  return [name, accessType, risk, email, owner, modified, url];
}

/**
 * EN: Whether this collaborator should appear in editor/viewer rows.
 * RU: Нужно ли включать этого участника в строки «Редактор»/«Читатель».
 *
 * @param {string} email
 * @returns {boolean}
 */
function includeCollaborator_(email) {
  if (!ORG_DOMAIN) return true;
  if (!email)      return true; // unknown entity — include for manual review
  const at = email.indexOf('@');
  if (at === -1)   return true;
  return email.substring(at + 1).toLowerCase() !== ORG_DOMAIN.toLowerCase();
}

/**
 * EN: Returns true when the email belongs to a domain other than ORG_DOMAIN.
 * RU: Возвращает true, если email относится к домену, отличному от ORG_DOMAIN.
 *
 * @param {string} email
 * @returns {boolean}
 */
function isExternal_(email) {
  if (!email || !ORG_DOMAIN) return false;
  const at = email.indexOf('@');
  if (at === -1) return false;
  return email.substring(at + 1).toLowerCase() !== ORG_DOMAIN.toLowerCase();
}

/**
 * EN: Risk level for a named collaborator (editor gets higher risk than viewer).
 * RU: Уровень риска для редактора (выше, чем для читателя).
 *
 * @param {string} email
 * @returns {string}
 */
function riskForEmail_(email) {
  return isExternal_(email) ? 'high' : 'low';
}

/**
 * @param {string} risk  Internal key
 * @returns {string}     Human-readable label with emoji
 */
function riskLabel_(risk) {
  const LABELS = {
    critical: '🔴 Критичный',
    high:     '🟠 Высокий',
    medium:   '🟡 Средний',
    low:      '🟢 Низкий',
    unknown:  '⚪ Неизвестно',
  };
  return LABELS[risk] || risk;
}

/**
 * @param {GoogleAppsScript.Drive.File} file
 * @returns {string}
 */
function ownerEmail_(file) {
  try {
    const owner = file.getOwner();
    return owner ? (owner.getEmail() || '(нет email)') : '(нет владельца)';
  } catch (e) {
    return '(ошибка)';
  }
}

/**
 * @param {Date|null} date
 * @returns {string}
 */
function formatDate_(date) {
  if (!date) return '?';
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  } catch (e) {
    return '?';
  }
}

/**
 * EN: Schedules continueAudit_ to run 60 seconds from now.
 * RU: Планирует вызов continueAudit_ через 60 секунд.
 */
function scheduleContinuation_() {
  ScriptApp.newTrigger('continueAudit_')
    .timeBased()
    .after(60000)
    .create();
}

/**
 * EN: Deletes all triggers whose handler matches the given function name.
 * RU: Удаляет все триггеры с указанным именем обработчика.
 *
 * @param {string} handlerName
 */
function deleteTriggersByHandler_(handlerName) {
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === handlerName; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
}

/**
 * @param {string} title
 * @param {string} message
 */
function toast_(title, message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title, 8);
}
