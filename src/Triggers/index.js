/* global App */

function copyEveryMonth() {
  // const files = getListFiles_();
  // const lastFile = files.sort((a, b) => b.num - a.num)[0];
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

function reset() {
  new App().settings = {
    APP_CURRENT_FILE_NUM: '7',
    APP_CURRENT_ID: '1fhHRjcFWHOCfx4t56SVI6SnQu0oFBsPvVGoXWIH2Gp4',
    APP_FOLDER_ID: '1mgzpM6dID_GUnzo-aQAAEv3kpFEmtgPx',
    APP_LIST_OF_EXEPTIONS_SHEETS: ['О Таблице'],
  };
}

/* exported triggerUpdateEveryHour */
/**
 * Триггер постоянного обновления Таблицы
 */
function triggerUpdateEveryHour() {
  const app = new App();
  // const book = app.currentBook;
  // cleanEmpties_(book, 'Новый лист для вашего примера');
  // insertNewSheet_(book, 'Новый лист для вашего примера');
  app.addNewBlankUserSheet();
  app.orderSheetsByProtections();
  app.generateTOC();
}

/* exported cleanEmpties_ */
/**
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} book
 * @param {string} sheetNameExclude
 */
function cleanEmpties_(book, sheetNameExclude) {
  book
    .getSheets()
    .forEach(
      (sheet) => sheet.getDataRange().isBlank() && sheet.getName() !== sheetNameExclude && book.deleteSheet(sheet),
    );
}

/* exported sheetsRemoveSheets_ */
/**
 * @param {{
 *   book: globalThis.SpreadsheetApp.Spreadsheet;
 *   filter: sheetsRemoveSheets.filterCallback
 * }} param0
 */
function sheetsRemoveSheets_({ book, filter = () => true }) {
  book
    .getSheets()
    .filter(filter)
    .map((sheet) => sheet.showSheet())
    .forEach((sheet) => {
      book.deleteSheet(sheet);
    });
}

/**
 * Filter callback
 * @callback sheetsRemoveSheets.filterCallback
 * @param {globalThis.SpreadsheetApp.Sheet} sheet
 * @returns {boolean}
 */
