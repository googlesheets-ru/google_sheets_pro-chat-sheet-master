/* global App */
/**
 * Удаляет все листы, кроме исключений.
 * @param {{excludes: string[]}} param0
 * @param {string[]} param0.excludes Дополнительный список листов-исключений.
 */
App.prototype.cleanBook = function ({ excludes }) {
  const excludesE = excludes || [];
  const listExceptions = [...this.settings.APP_LIST_OF_EXEPTIONS_SHEETS, ...excludesE];
  const requests = this.book.sheets
    .filter((sheet) => !listExceptions.includes(sheet.properties.title))
    .map((sheet) => {
      const deleteSheetRequest = Sheets.newDeleteSheetRequest();
      deleteSheetRequest.sheetId = sheet.properties.sheetId;
      const request = Sheets.newRequest();
      request.deleteSheet = deleteSheetRequest;
      return request;
    });
  if (requests.length) {
    Sheets.Spreadsheets.batchUpdate({ requests }, this.book.spreadsheetId);
    this._book = undefined;
  }
};
