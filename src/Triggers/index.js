/* global App */

/* exported triggerUpdateEveryMonth */
/**
 * Триггер создания новой Таблицы чата
 */
function triggerUpdateEveryMonth() {
  const app = new App();
  const currentBook = DriveApp.getFileById(app.settings.APP_CURRENT_ID);
  const num = Number(app.settings.APP_CURRENT_FILE_NUM) + 1;
  const copy = currentBook.makeCopy(`Таблица чата t.me/google_sheets_pro #${num}`, app.folder);
  copy.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
  app.settings = { ...app.settings, ...{ APP_CURRENT_FILE_NUM: `${num}`, APP_CURRENT_ID: copy.getId() } };
  app.updateStamp({ num, prevUrl: currentBook.getUrl(), prevTitle: currentBook.getName() });
  app.cleanBook();
  app.addNewBlankUserSheet();
  app.orderSheetsByProtections();
  app.generateTOC();
}

/* exported triggerUpdateEveryHour */
/**
 * Триггер постоянного обновления Таблицы
 */
function triggerUpdateEveryHour() {
  const app = new App();
  app.addNewBlankUserSheet();
  app.orderSheetsByProtections();
  app.generateTOC();
}
