<div align="center">

# Google Drive Security Audit

### Аудит прав доступа к Google Диску · Google Drive permissions audit

![Google Apps Script](https://img.shields.io/badge/Google-Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![Google Sheets](https://img.shields.io/badge/Google-Sheets-34A853?style=for-the-badge&logo=googlesheets&logoColor=white)
![Version](https://img.shields.io/badge/version-2.0-blue?style=for-the-badge)

**RU:** Цветной многолистовой отчёт о публичном доступе, доступе по ссылке и редакторах/читателях с поддержкой Shared Drives и возобновляемого обхода.  
**EN:** Colour-coded multi-tab report for public / link sharing, editors and viewers — with Shared Drives support and resumable scans.

[Русский документ](#русский) · [English](#english)

</div>

---

<br id="top">

## Русский

<p align="right"><a href="#top">↑ наверх</a> · <a href="#english">English →</a></p>

### Содержание

| Раздел | Описание |
|--------|----------|
| [Зачем](#зачем-это-нужно) | Задача и польза |
| [Как работает](#как-это-работает) | Логика отчёта v2.0 |
| [Установка](#установка) | Пошагово в Таблице |
| [`ORG_DOMAIN`](#настройка-org_domain) | Фильтр «только внешние» |
| [Листы отчёта](#листы-отчёта) | Три вкладки по риску |
| [Уровни риска](#уровни-риска) | Цветовая маркировка |
| [Колонки](#колонки-отчёта) | Поля и смысл |
| [Shared Drives](#поддержка-shared-drives) | Общие диски |
| [Большие диски](#большие-диски--возобновляемый-обход) | Resumable scan |
| [Email-уведомления](#email-уведомления) | Автоматические оповещения |
| [Расписание](#автоматическое-расписание) | Еженедельный триггер |
| [Риски](#ограничения-и-риски) | Лимиты и оговорки |
| [Инцидент](#действия-при-обнаружении-угроз) | Что делать при находке |
| [Скриншоты](#скриншоты) | Внешний вид в Sheets |

### Зачем это нужно

Сотрудники открывают доступ «по ссылке» или конкретным людям; со временем накапливаются объекты с широким или внешним доступом. Скрипт собирает **список для проверки** прямо в Google Таблице, расставляет приоритеты по уровню риска и умеет уведомлять по email.

### Как это работает

В **Extensions → Apps Script** выполняется обход файлов через `DriveApp` (или Drive API v2 при включённом Advanced Drive Service). Проверяется тип общего доступа, списки **редакторов** и **читателей**. Результат пишется **пакетно** на три листа с цветовой маркировкой по уровню риска.

**Что нового в v2.0**

| Улучшение | Описание |
|-----------|----------|
| 🗂 Три листа отчёта | Публичные / Внешние / Полный отчёт |
| 🎨 Цветовая маркировка | Красный → оранжевый → жёлтый → зелёный |
| 👤 Колонка «Владелец» | `file.getOwner()` — кому принадлежит файл |
| 📅 Колонка «Последнее изменение» | `file.getLastUpdated()` в формате дд.мм.гггг |
| 🔁 Возобновляемый обход | `PropertiesService` + триггер для больших дисков |
| 🗄 Shared Drives | Drive API v2 с `includeItemsFromAllDrives: true` |
| 📧 Email-уведомления | Сводка при обнаружении высокого риска |
| ⏰ Еженедельный триггер | Автоматический запуск по расписанию |
| 🐛 Баг-фикс | Исправлен `rows.length + 1` (лишняя пустая строка) |

### Установка

1. Создайте [Google Таблицу](https://sheets.google.com).
2. **Расширения** → **Apps Script**.
3. Вставьте код из [`audit_script.js`](./audit_script.js), сохраните проект.
4. *(Опционально, рекомендуется)* **Services** → найдите **Drive API** → добавьте v2. Это включает поддержку Shared Drives и возобновляемого обхода.
5. Вернитесь в таблицу, **обновите страницу (F5)**.
6. Меню **«🛡️ Безопасность»** → **«▶ Запустить аудит прав»**, выдайте права Диску и Таблице.

> **Примечание:** при первом запуске Google может пометить скрипт как «непроверенный» — это нормально для личного кода.

### Настройка `ORG_DOMAIN`

В начале [`audit_script.js`](./audit_script.js):

| Значение | Поведение |
|----------|-----------|
| `company.com` (без `@`) | В строках редактор/читатель только **внешние** относительно домена; публичные файлы как прежде |
| `''` (пустая строка) | **Полный** список редакторов и читателей по каждому файлу |

### Настройка `ADMIN_EMAIL`

```js
const ADMIN_EMAIL = 'security@company.com';
```

Если оставить пустым — уведомление придёт на аккаунт, запустивший скрипт.

### Листы отчёта

| Лист | Содержимое |
|------|------------|
| 🔴 Публичные | Файлы с `ПУБЛИЧНЫЙ (все)` или `ПУБЛИЧНЫЙ (по ссылке)` |
| 🟡 Внешние | Редакторы и читатели вне домена `ORG_DOMAIN` |
| 📋 Полный отчёт | Все строки |

### Уровни риска

| Цвет | Уровень | Условие |
|------|---------|---------|
| 🔴 Красный | Критичный | Доступ «любой в интернете» |
| 🟠 Оранжевый | Высокий | Доступ по ссылке или внешний редактор |
| 🟡 Жёлтый | Средний | Внешний читатель |
| 🟢 Зелёный | Низкий | Внутренний участник |
| ⚪ Серый | Неизвестно | Ошибка чтения метаданных |

### Колонки отчёта

| Колонка | Источник |
|---------|----------|
| Название файла | `file.getName()` |
| Тип доступа | Режим общего доступа или роль |
| Уровень риска | Автоматически по логике выше |
| Email с доступом | Участник ACL или `—` для публичного |
| Владелец | `file.getOwner().getEmail()` |
| Последнее изменение | `file.getLastUpdated()` |
| Ссылка | `file.getUrl()` |

### Поддержка Shared Drives

При включённом **Drive API v2** скрипт автоматически обходит **все Общие диски** организации, передавая `includeItemsFromAllDrives: true` и `corpora: 'allDrives'`. Без этого — только «Мой диск» через `DriveApp`.

Как включить: **Apps Script → Services → Drive API → v2 → Add**.

### Большие диски — возобновляемый обход

Если файлов много и выполнение упирается в лимит 6 минут:

1. Скрипт сохраняет `pageToken` в `PropertiesService`.
2. Автоматически создаёт триггер-продолжение через 60 секунд.
3. После завершения всех страниц пишет финальный отчёт.

Без Drive API v2 возобновляемый обход недоступен — скрипт записывает обработанное и предупреждает.

### Email-уведомления

Письмо отправляется автоматически после каждого завершённого аудита, **если найдены файлы с критичным или высоким риском**. Для настройки адреса используйте переменную `ADMIN_EMAIL`.

### Автоматическое расписание

Меню **«🛡️ Безопасность»** → **«⏰ Настроить еженедельный аудит»** — создаёт триггер на каждый понедельник в 9:00. Удалить расписание: **«🗑 Удалить еженедельный аудит»**.

### Ограничения и риски

- При отсутствии Drive API v2 — только «Мой диск»; Shared Drives не охвачены.
- Группы и сервисные сущности (без email) включаются в отчёт для ручной проверки.
- Не публикуйте лист с отчётом публично; не храните секреты в репозитории.
- Квоты MailApp: 100 писем/день для стандартных аккаунтов, 1500 для Workspace.

### Действия при обнаружении угроз

1. **Классификация** — критичность данных.
2. **Изоляция** — доступ «Только по приглашению» / отзыв ссылки.
3. **Уведомление** — ИБ / DPO по регламенту при ПДн или секретах.

### Скриншоты

| Меню | Отчёт |
|------|--------|
| <img width="420" alt="Меню безопасности" src="https://github.com/user-attachments/assets/c8fdfa08-78fc-469f-ae3d-9c0e370ae5ff" /> | <img width="420" alt="Пример отчёта" src="https://github.com/user-attachments/assets/7d75f6a2-aa49-438d-a80b-e13bb4994f94" /> |

---

## English

<p align="right"><a href="#top">↑ top</a> · <a href="#русский">← Русский</a></p>

### Contents

| Section | What it covers |
|---------|----------------|
| [Why](#why-use-this) | Problem and outcome |
| [How it works](#how-it-works) | Report logic v2.0 |
| [Setup](#setup) | Sheets + Apps Script |
| [`ORG_DOMAIN`](#org_domain-setting) | External-only filter |
| [Report tabs](#report-tabs) | Three risk-based tabs |
| [Risk levels](#risk-levels) | Colour coding |
| [Columns](#report-columns) | Fields and meaning |
| [Shared Drives](#shared-drives) | Shared drive support |
| [Large drives](#large-drives--resumable-scan) | Resumable scan |
| [Email alerts](#email-alerts) | Automated notifications |
| [Schedule](#automated-schedule) | Weekly trigger |
| [Limits](#limits-and-risks) | Caveats |
| [Incident response](#if-you-find-a-risk) | Practical steps |
| [Screenshots](#screenshots) | UI preview |

### Why use this

Teams often share files with **anyone with the link** or external people. Over time, risky sharing accumulates. This script builds a **colour-coded, prioritised checklist** across three tabs in Google Sheets and sends an email alert when high-risk files are found.

### How it works

The code in **Extensions → Apps Script** walks files via `DriveApp` (or Drive API v2 when the Advanced Drive Service is enabled). It inspects sharing mode, **editors** and **viewers**, and writes rows in **one batch** to three tabs with per-row risk colouring.

**What's new in v2.0**

| Improvement | Details |
|-------------|---------|
| 🗂 Three report tabs | Public / External / Full report |
| 🎨 Colour-coded rows | Red → orange → yellow → green by risk |
| 👤 Owner column | `file.getOwner()` — who owns the file |
| 📅 Last modified column | `file.getLastUpdated()` as dd.MM.yyyy |
| 🔁 Resumable scan | `PropertiesService` + continuation trigger |
| 🗄 Shared Drives | Drive API v2 with `includeItemsFromAllDrives: true` |
| 📧 Email alerts | Summary on high-risk finds |
| ⏰ Weekly trigger | Automated scheduled run |
| 🐛 Bug fix | Removed `rows.length + 1` (extra blank row) |

### Setup

1. Create a [Google Sheet](https://sheets.google.com).
2. **Extensions** → **Apps Script**.
3. Paste [`audit_script.js`](./audit_script.js), save the project.
4. *(Optional, recommended)* **Services** → search for **Drive API** → add v2. This enables Shared Drives and resumable scans.
5. Go back to the sheet and **reload the page (F5)**.
6. Open **"🛡️ Безопасность"** (Security) → **"▶ Запустить аудит прав"**, authorize Drive + Sheets.

> **Note:** Google may show an "unverified app" warning for personal scripts — expected unless published as an add-on.

### `ORG_DOMAIN` setting

At the top of [`audit_script.js`](./audit_script.js):

| Value | Behavior |
|-------|----------|
| `company.com` (no `@`) | Editor/viewer rows list **only** people **outside** that domain; public rows unchanged |
| `''` (empty string) | **Full** editor/viewer inventory per file |

### `ADMIN_EMAIL` setting

```js
const ADMIN_EMAIL = 'security@company.com';
```

Leave empty to send the alert to the account running the script.

### Report tabs

| Tab | Contents |
|-----|----------|
| 🔴 Публичные (Public) | Files with `ПУБЛИЧНЫЙ (все)` or `ПУБЛИЧНЫЙ (по ссылке)` |
| 🟡 Внешние (External) | Editors and viewers outside `ORG_DOMAIN` |
| 📋 Полный отчёт (Full) | Every row |

### Risk levels

| Colour | Level | Condition |
|--------|-------|-----------|
| 🔴 Red | Critical | Anyone on the Internet |
| 🟠 Orange | High | Anyone with the link, or external editor |
| 🟡 Yellow | Medium | External viewer |
| 🟢 Green | Low | Internal collaborator |
| ⚪ Grey | Unknown | Metadata read error |

### Report columns

| Column | Source |
|--------|--------|
| File name | `file.getName()` |
| Access type | Sharing mode or role |
| Risk level | Derived automatically |
| Email | ACL member or `—` for public |
| Owner | `file.getOwner().getEmail()` |
| Last modified | `file.getLastUpdated()` |
| Link | `file.getUrl()` |

### Shared Drives

When **Drive API v2** is enabled, the script automatically walks **all Shared Drives** in the organisation by passing `includeItemsFromAllDrives: true` and `corpora: 'allDrives'`. Without it, only My Drive is scanned via `DriveApp`.

To enable: **Apps Script → Services → Drive API → v2 → Add**.

### Large drives — resumable scan

If the drive has many files and hits the 6-minute execution limit:

1. The script saves the current `pageToken` to `PropertiesService`.
2. It automatically creates a continuation trigger to fire in 60 seconds.
3. After the last page, the final report is written.

Without Drive API v2, resumable scan is unavailable — the script writes what it processed and shows a warning.

### Email alerts

An alert is sent automatically after each completed audit **if critical or high-risk files are found**. Configure the recipient via `ADMIN_EMAIL`.

### Automated schedule

Menu **"🛡️ Безопасность"** → **"⏰ Настроить еженедельный аудит"** creates a trigger for every Monday at 09:00. Remove it with **"🗑 Удалить еженедельный аудит"**.

### Limits and risks

- Without Drive API v2 — My Drive only; Shared Drives are not scanned.
- Groups and service principals (no email) are included for manual review.
- Do not publish the audit sheet publicly; keep secrets out of Git.
- MailApp quota: 100 emails/day on personal accounts, 1 500 on Workspace.

### If you find a risk

1. **Classify** data sensitivity.
2. **Contain** — switch to restricted access / revoke the link.
3. **Notify** security / privacy per your policy if passwords or personal data were exposed.

### Screenshots

| Menu | Report |
|------|--------|
| <img width="420" alt="Security menu in Google Sheets" src="https://github.com/user-attachments/assets/c8fdfa08-78fc-469f-ae3d-9c0e370ae5ff" /> | <img width="420" alt="Sample audit output" src="https://github.com/user-attachments/assets/7d75f6a2-aa49-438d-a80b-e13bb4994f94" /> |

---

<div align="center">

<sub>Add a `LICENSE` file (e.g. MIT) if you redistribute · При распространении добавьте файл `LICENSE` (например MIT).</sub>

</div>
