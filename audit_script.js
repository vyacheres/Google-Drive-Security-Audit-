/**
 * Google Drive sharing audit for Google Sheets (Apps Script).
 * Аудит прав доступа к Google Диску из таблицы (Google Apps Script).
 *
 * @fileoverview
 * EN: Scans files you can enumerate via DriveApp, flags link/public sharing,
 *     lists editors and viewers (optionally only outside ORG_DOMAIN).
 * RU: Обходит файлы через DriveApp, отмечает публичный/ссылочный доступ,
 *     выводит редакторов и читателей (при необходимости — только вне ORG_DOMAIN).
 */

/**
 * Organization domain without "@", e.g. "company.com".
 * Домен организации без символа @, например "company.com".
 *
 * EN: If non-empty, editor/viewer rows include only emails whose domain is NOT ORG_DOMAIN.
 *     Public / anyone-with-link rows are always included when applicable.
 * RU: Если не пусто — строки «Редактор»/«Читатель» только для адресов вне этого домена.
 *     Строки про публичный доступ по-прежнему выводятся для соответствующих файлов.
 *
 * EN: Leave '' to list every editor and viewer on each file (full ACL inventory).
 * RU: Оставьте '' — в отчёт попадут все редакторы и читатели по каждому файлу.
 */
const ORG_DOMAIN = '';

/**
 * EN: Runs the audit and writes results to the active sheet (row 1 = header).
 * RU: Запускает аудит и записывает результат на активный лист (строка 1 — заголовки).
 */
function driveAudit() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();

  // Column titles / Заголовки колонок отчёта
  const header = ['Название файла', 'Тип доступа', 'Email с доступом', 'Ссылка'];
  const rows = [];

  const files = DriveApp.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    try {
      const access = file.getSharingAccess();
      const url = file.getUrl();
      const name = file.getName();

      // EN: Anyone on the web vs anyone with the link — both are high-risk sharing modes.
      // RU: «Любой в интернете» и «по ссылке» — оба режима повышенного риска.
      if (access === DriveApp.Access.ANYONE || access === DriveApp.Access.ANYONE_WITH_LINK) {
        const label =
          access === DriveApp.Access.ANYONE ? 'ПУБЛИЧНЫЙ (все)' : 'ПУБЛИЧНЫЙ (по ссылке)';
        // No individual email for link-wide access / Для широкого доступа email не применим
        rows.push([name, label, '—', url]);
      }

      // EN: Editors always have at least view rights; list each as its own row.
      // RU: Каждый редактор — отдельная строка (удобно сортировать и фильтровать).
      file.getEditors().forEach(function (e) {
        const email = e.getEmail();
        if (!includeCollaborator(email)) {
          return;
        }
        rows.push([name, 'Редактор', email || '(без email)', url]);
      });

      // EN: Viewers (read-only); same filtering rules as editors.
      // RU: Читатели (только просмотр); те же правила фильтра, что и для редакторов.
      file.getViewers().forEach(function (v) {
        const email = v.getEmail();
        if (!includeCollaborator(email)) {
          return;
        }
        rows.push([name, 'Читатель', email || '(без email)', url]);
      });
    } catch (err) {
      // EN: One bad file must not stop the whole run.
      // RU: Ошибка на одном файле не должна прерывать весь обход.
      rows.push([file.getName(), 'Ошибка чтения', String(err.message || err), file.getUrl()]);
    }
  }

  // EN: Batch write is much faster than appendRow in a loop.
  // RU: Пакетная запись быстрее, чем appendRow в цикле.
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length + 1, header.length).setValues(rows);
  }
}

/**
 * EN: Whether this collaborator should appear in editor/viewer rows.
 * RU: Нужно ли включать этого участника в строки «Редактор»/«Читатель».
 *
 * @param {string} email
 * @returns {boolean}
 */
function includeCollaborator(email) {
  // EN: No domain filter — include everyone (full inventory).
  // RU: Фильтр по домену выключен — показываем всех.
  if (!ORG_DOMAIN) {
    return true;
  }
  // EN: Missing email (e.g. some groups) — still show the row for manual review.
  // RU: Пустой email (иногда группы) — показываем строку для ручной проверки.
  if (!email) {
    return true;
  }
  const at = email.indexOf('@');
  if (at === -1) {
    return true;
  }
  const domain = email.substring(at + 1).toLowerCase();
  return domain !== ORG_DOMAIN.toLowerCase();
}

/**
 * EN: Adds custom menu when the spreadsheet is opened (reload sheet after deploy).
 * RU: Добавляет меню при открытии таблицы (после первой установки скрипта обновите F5).
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛡️ Безопасность')
    .addItem('Запустить аудит прав', 'driveAudit')
    .addToUi();
}
