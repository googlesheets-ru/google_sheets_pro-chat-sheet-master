/* global App */
/**
 * Добавляет новый лист для примера
 */
App.prototype.addNewBlankUserSheet = function () {
  if (!this.book.sheets.some((sheet) => sheet.properties.title === 'Новый лист для вашего примера')) {
    const addSheetRequest = Sheets.newAddSheetRequest();
    addSheetRequest.properties = {
      index: 1,
      title: 'Новый лист для вашего примера',
    };
    const request = Sheets.newRequest();
    request.addSheet = addSheetRequest;
    const resource = Sheets.newBatchUpdateSpreadsheetRequest();
    resource.requests = [request];
    resource.includeSpreadsheetInResponse = true;
    resource.responseIncludeGridData = false;
    const reply = Sheets.Spreadsheets.batchUpdate(resource, this.settings.APP_CURRENT_ID).replies.find(
      (reply) => reply.addSheet,
    ).addSheet;
    this._book.sheets.splice(1, 0, {
      properties: reply.properties,
    });
    this._book.sheets.forEach((sheet, i) => (sheet.properties.index = i));
  }
};
