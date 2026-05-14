<div align="center">

# Google Drive Security Audit

### Аудит прав доступа к Google Диску · Google Drive permissions audit

![Google Apps Script](https://img.shields.io/badge/Google-Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![Google Sheets](https://img.shields.io/badge/Google-Sheets-34A853?style=for-the-badge&logo=googlesheets&logoColor=white)

**RU:** Табличный отчёт о публичном доступе, доступе по ссылке и редакторах/читателях.  
**EN:** Spreadsheet report for public / link sharing plus editors and viewers.

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
| [Как работает](#как-это-работает) | Логика отчёта |
| [Установка](#установка) | Пошагово в Таблице |
| [`ORG_DOMAIN`](#настройка-org_domain) | Фильтр «только внешние» |
| [Отчёт](#интерпретация-отчёта) | Колонки и смысл строк |
| [Риски](#ограничения-и-риски) | Лимиты Apps Script и Диска |
| [Инцидент](#действия-при-обнаружении-угроз) | Что делать при находке |
| [Дальше](#идеи-для-развития) | Возможные улучшения кода |
| [Скриншоты](#скриншоты) | Внешний вид в Sheets |

### Зачем это нужно

Сотрудники открывают доступ «по ссылке» или конкретным людям; со временем накапливаются объекты с широким или внешним доступом. Скрипт собирает **список для проверки** прямо в Google Таблице.

### Как это работает

В **Extensions → Apps Script** выполняется обход файлов через `DriveApp`, проверяется тип общего доступа, списки **редакторов** и **читателей**, результат пишется **пакетно** на активный лист (быстрее, чем `appendRow` в цикле).

**Типы строк в отчёте**

| Тип в колонке «Тип доступа» | Смысл |
|----------------------------|--------|
| `ПУБЛИЧНЫЙ (все)` | Доступ «любой в интернете» |
| `ПУБЛИЧНЫЙ (по ссылке)` | Узнаваемый по ссылке без явного ACL |
| `Редактор` / `Читатель` | Email и роль (см. `ORG_DOMAIN`) |
| `Ошибка чтения` | Сбой при чтении метаданных файла |

### Установка

1. Создайте [Google Таблицу](https://sheets.google.com).
2. **Расширения** → **Apps Script**.
3. Вставьте код из [`audit_script.js`](./audit_script.js), сохраните проект.
4. Вернитесь в таблицу, **обновите страницу (F5)**.
5. Меню **«Безопасность»** → **«Запустить аудит прав»**, выдайте права Диску и Таблице.

> **Примечание:** при первом запуске Google может пометить скрипт как «непроверенный» — это нормально для личного кода, если политика Workspace это допускает.

### Настройка `ORG_DOMAIN`

В начале [`audit_script.js`](./audit_script.js):

| Значение | Поведение |
|----------|-----------|
| `company.com` (без `@`) | В строках редактор/читатель только **внешние** относительно домена; публичные файлы как прежде |
| `''` (пустая строка) | **Полный** список редакторов и читателей по каждому файлу |

### Интерпретация отчёта

Колонки: **Название файла**, **Тип доступа**, **Email с доступом**, **Ссылка**. Имеет смысл сначала отфильтровать публичные строки, затем внешних участников.

### Ограничения и риски

- `DriveApp.getFiles()` — типичный сценарий **«Мой диск»**; Shared Drives и сложные политики могут потребовать [Advanced Drive Service](https://developers.google.com/apps-script/advanced/drive).
- Лимит времени выполнения Apps Script (~6 мин) — при огромном числе файлов нужны пагинация, триггеры или Drive API.
- Группы и сервисные сущности — проверяйте строки вручную.
- Не публикуйте лист с отчётом публично; не храните секреты в репозитории.

### Действия при обнаружении угроз

1. **Классификация** — критичность данных.  
2. **Изоляция** — доступ «Только по приглашению» / отзыв ссылки.  
3. **Уведомление** — ИБ / DPO по регламенту при ПДн или секретах.

### Идеи для развития

Shared Drives, отдельные листы по риску, CSV/email, доменный доступ отдельными строками, возобновляемый обход, [clasp](https://github.com/google/clasp) для Git-воркфлоу.

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
| [How it works](#how-it-works) | Report logic |
| [Setup](#setup) | Sheets + Apps Script |
| [`ORG_DOMAIN`](#org_domain-setting) | External-only filter |
| [Reading the report](#reading-the-report) | Columns and row types |
| [Limits](#limits-and-risks) | Drive + Apps Script caveats |
| [Incident response](#if-you-find-a-risk) | Practical steps |
| [Roadmap](#possible-next-steps) | Enhancements |
| [Screenshots](#screenshots) | UI preview |

### Why use this

Teams often share files with **anyone with the link** or external people. Over time, risky sharing accumulates. This script builds a **reviewable checklist** in Google Sheets.

### How it works

The code in **Extensions → Apps Script** walks files via `DriveApp`, inspects sharing mode, **editors** and **viewers**, and writes rows in **one batch** to the active sheet (faster than per-row `appendRow`).

**Row types (column “Тип доступа” / access type)**

| Value | Meaning |
|-------|---------|
| `ПУБЛИЧНЫЙ (все)` | Literally anyone on the Internet |
| `ПУБЛИЧНЫЙ (по ссылке)` | Broad link-based access |
| `Редактор` / `Читатель` | Editor / viewer email (subject to `ORG_DOMAIN`) |
| `Ошибка чтения` | Could not read that file’s metadata |

### Setup

1. Create a [Google Sheet](https://sheets.google.com).
2. **Extensions** → **Apps Script**.
3. Paste [`audit_script.js`](./audit_script.js), save the project.
4. Go back to the sheet and **reload the page (F5)**.
5. Open **«Безопасность»** (Security) → **«Запустить аудит прав»**, authorize Drive + Sheets.

> **Note:** Google may show an “unverified app” warning for personal scripts—expected unless you publish as an add-on and your Workspace policy allows it.

### `ORG_DOMAIN` setting

At the top of [`audit_script.js`](./audit_script.js):

| Value | Behavior |
|-------|----------|
| `company.com` (no `@`) | Editor/viewer rows list **only** people **outside** that domain; public rows unchanged |
| `''` (empty string) | **Full** editor/viewer inventory per file |

### Reading the report

Columns: **file name**, **access type**, **email**, **link**. Start with public rows, then external collaborators (when a domain filter is set).

### Limits and risks

- `DriveApp.getFiles()` matches the usual **My Drive** owner view; **Shared drives** or strict Workspace rules may need the [Advanced Drive Service](https://developers.google.com/apps-script/advanced/drive).
- Apps Script runtime (~6 minutes)—very large drives may need pagination, time-driven triggers, or the Drive API.
- Groups / service principals may need manual review.
- Do not publish the audit sheet publicly; keep secrets out of Git.

### If you find a risk

1. **Classify** data sensitivity.  
2. **Contain**—switch to restricted access / revoke the link.  
3. **Notify** security / privacy per your policy if passwords or personal data were exposed.

### Possible next steps

Shared drives support, risk tabs, CSV/email export, explicit domain-only rows, resumable scans, [clasp](https://github.com/google/clasp) for source control.

### Screenshots

| Menu | Report |
|------|--------|
| <img width="420" alt="Security menu in Google Sheets" src="https://github.com/user-attachments/assets/c8fdfa08-78fc-469f-ae3d-9c0e370ae5ff" /> | <img width="420" alt="Sample audit output" src="https://github.com/user-attachments/assets/7d75f6a2-aa49-438d-a80b-e13bb4994f94" /> |

---

<div align="center">

<sub>Add a `LICENSE` file (e.g. MIT) if you redistribute · При распространении добавьте файл `LICENSE` (например MIT).</sub>

</div>
