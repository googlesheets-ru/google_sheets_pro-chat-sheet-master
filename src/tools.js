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
