/**
 * Пожалуйста, не удаляйте этот код, пока Таблица доступна всем
 */
function onopen() {
  SpreadsheetApp.getUi()
    .createMenu('Инструменты t.me/google_sheets_pro')
    .addItem('Обновить содержание', 'userActionGenerateTOC')
    .addToUi();
}

/**
 * An user action
 */
function userActionGenerateTOC() {
  const book = SpreadsheetApp.getActive();
  generateTOC_(book);
}

/**
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} book
 */
function generateTOC_(book) {
  const excludeSheetNames = ['О Таблице'];
  const sheets = book.getSheets().reduce(
    (a, c) => {
      if (c.getProtections(SpreadsheetApp.ProtectionType.SHEET).length) a.protected.push(c.getName());
      else a.free.push(c.getName());
      return a;
    },
    {
      free: [],
      protected: [],
    },
  );
  const order = [...excludeSheetNames, ...sheets.free, ...sheets.protected];
  order.reverse().forEach((name) => (book.getSheetByName(name).activate(), book.moveActiveSheet(0)));
  const tocBuild = tocBuilder_(book).filter((item) => excludeSheetNames.indexOf(item.name) === -1);

  const firstPage = book.getSheetByName('О Таблице');
  const pos = firstPage.createTextFinder('Содержание').matchCase(true).matchFormulaText(false).findNext();
  if (pos) {
    const range = firstPage.getRange(pos.getRow() + 1, pos.getColumn(), firstPage.getLastRow()).clearContent();
    tocUpdater_(tocBuild, range);
    range.activate();
  }
}
