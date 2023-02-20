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

function generateTOC_(book) {
  const excludeSheetNames = ['О Таблице'];
  const order = [...excludeSheetNames, 'Зачем нужна Таблица', 'Как пользоваться Таблицей'];
  order.reverse().forEach((name) => (book.getSheetByName(name).activate(), book.moveActiveSheet(0)));
  const tocBuild = tocBuilder_(book).filter((item) => excludeSheetNames.indexOf(item.name) === -1);

  const firstPage = book.getSheetByName('О Таблице');
  const pos = firstPage.createTextFinder('Содержание').matchCase(true).matchFormulaText(false).findNext();
  if (pos) {
    const range = firstPage.getRange(pos.getRow() + 2, pos.getColumn(), firstPage.getLastRow()).clearContent();
    tocUpdater_(tocBuild, range);
    range.activate();
  }
}
