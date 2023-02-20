const __SETTINGS__ = Object.freeze({
  fixedSheetsNames: ['О Таблице', 'Зачем нужна Таблица', 'Как пользоваться Таблицей'],
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
  const files = getListFiles_();
  const lastFile = files.sort((a, b) => b.num - a.num)[0];
  const copy = lastFile.file.makeCopy(`Таблица чата t.me/google_sheets_pro #${lastFile.num + 1}`, folder);
  copy.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
  prepareFile(copy, lastFile);
}

function updateEveryHour() {
  const files = getListFiles_();
  const lastFile = files.sort((a, b) => b.num - a.num)[0];
  console.log(lastFile.name);
  const book = SpreadsheetApp.open(lastFile.file);
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
