/* global App */
/**
 * Обновляет кеш информации о книге (`this._book`).
 */
App.prototype.recallBook = function () {
  this._book = Sheets.Spreadsheets.get(this.currentBook.getId(), {
    includeGridData: false,
    fields: 'spreadsheetId,sheets(properties(sheetId,index,title),protectedRanges(range,protectedRangeId))',
  });
};
