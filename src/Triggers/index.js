/* global App */

/* exported triggerUpdateEveryMonth */
/**
 * Триггер создания новой Таблицы чата.
 * Выполняется ежемесячно.
 * Копирует текущую таблицу, обновляет настройки и подготавливает новую таблицу к использованию.
 */
function triggerUpdateEveryMonth() {
  // Создаем экземпляр приложения App
  const app = new App();
  app.createNextBook();
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

/* exported triggerUpdateEveryMin */
function triggerUpdateEveryMin() {
  const app = new App();
  app.resetName();
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
