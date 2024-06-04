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

/* exported userActionCleanBook */
function userActionCleanBook() {
  const excludes = ['формула по выпадающему списку'];
  new App().cleanBook({ excludes });
}

/* exported cleanOldSheet */
function cleanOldSheet() {
  const app = new App({
    APP_CURRENT_FILE_NUM: '8',
    APP_CURRENT_ID: '1MqeW7LkEUcsDH8lUksXBLHWmSeu8up2Jr_3ujVwE4SE',
    APP_FOLDER_ID: '1mgzpM6dID_GUnzo-aQAAEv3kpFEmtgPx',
    APP_LIST_OF_EXEPTIONS_SHEETS: '["О Таблице"]',
  });
  app.releaseSheets();
  app.generateTOC();
}
