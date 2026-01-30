function driveAudit() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  sheet.appendRow(["Название файла", "Тип доступа", "Email с доступом", "Ссылка"]);
  
  const files = DriveApp.getFiles();
  
  while (files.hasNext()) {
    let file = files.next();
    let access = file.getSharingAccess();
    
    // Проверяем файлы с публичным доступом или доступом по ссылке
    if (access == DriveApp.Access.ANYONE || access == DriveApp.Access.ANYONE_WITH_LINK) {
      sheet.appendRow([file.getName(), "ПУБЛИЧНЫЙ", "Все (по ссылке)", file.getUrl()]);
    }
    
    // Проверка конкретных пользователей (редакторов/читателей)
    let editors = file.getEditors();
    editors.forEach(e => {
      sheet.appendRow([file.getName(), "Редактор", e.getEmail(), file.getUrl()]);
    });
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛡️ Безопасность')
    .addItem('Запустить аудит прав', 'driveAudit')
    .addToUi();
}
