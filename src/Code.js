const __SETTINGS__ = Object.freeze({
  fixedSheetsNames: ['О Таблице'],
});

function getListFiles_() {
  const folder = DriveApp.getFolderById('1mgzpM6dID_GUnzo-aQAAEv3kpFEmtgPx');
  const filesIterator = folder.searchFiles(
    'title contains "Таблица чата t.me/google_sheets_pro #" and mimeType="application/vnd.google-apps.spreadsheet"',
  );
  const files = [];
  while (filesIterator.hasNext()) {
    const file = filesIterator.next();
    const name = file.getName();
    const [_, num] = name.match(/.*?#.*?(\d+)/) || ['', -1];
    files.push({ file, name, num: +num });
  }
  return files;
}

function copyEveryMonth() {
  // const files = getListFiles_();
  // const lastFile = files.sort((a, b) => b.num - a.num)[0];
  const app = new App();
  const currentBook = DriveApp.getFileById(app.settings.APP_CURRENT_ID);
  const num = Number(app.settings.APP_CURRENT_FILE_NUM) + 1;
  const copy = currentBook.makeCopy(`Таблица чата t.me/google_sheets_pro #${num}`, app.folder);
  copy.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
  app.settings = { APP_CURRENT_FILE_NUM: `num`, APP_CURRENT_ID: copy.getId() };
  prepareFile(copy, { num, file: currentBook });
}

function updateEveryHour() {
  // const files = getListFiles_();
  // const lastFile = files.sort((a, b) => b.num - a.num)[0];
  // console.log(lastFile.name);
  const app = new App();
  // cons
  const book = app.currentBook;
  cleanEmpties_(book, 'Новый лист для вашего примера');
  insertNewSheet_(book, 'Новый лист для вашего примера');
  generateTOC_(book);
}

function prepareFile(file, lastFile) {
  const book = SpreadsheetApp.open(file);
  sheetsRemoveSheets_({
    book,
    filter: (sheet) => __SETTINGS__.fixedSheetsNames.indexOf(sheet.getName()) === -1,
  });
  const sheetAbout = book.getSheetByName('О Таблице');
  sheetAbout.getRange('A1:B2').setValues([
    [lastFile.num + 1, ''],
    [lastFile.file.getUrl(), lastFile.file.getName()],
  ]);
  generateTOC_(book);
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

/* exported insertNewSheet_ */
/**
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} book
 * @param {string} sheetName
 */
function insertNewSheet_(book, sheetName) {
  const sheet = book.getSheetByName(sheetName);
  if (sheet)
    if (!sheet.getDataRange().isBlank()) {
      sheet.setName(`${sheet.getName()} [${new Date().getTime()}]`);
    } else return;
  !book.getSheetByName(sheetName) && book.insertSheet(sheetName, 1);
}
