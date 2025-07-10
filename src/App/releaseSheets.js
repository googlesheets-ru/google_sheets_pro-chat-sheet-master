/* global App */
/**
 * Снимает защиту со всех листов в книге.
 */
App.prototype.releaseSheets = function () {
  /** @type {GoogleAppsScript.Sheets.Schema.Sheet[]} */
  const sheets = JSON.parse(JSON.stringify(this.book.sheets));
  const requests = [];
  sheets.forEach((sheet) => {
    sheet.protectedRanges?.forEach((protectedRange) => {
      if (protectedRange.protectedRangeId) {
        const deleteProtectedRangeRequest = Sheets.newDeleteProtectedRangeRequest();
        deleteProtectedRangeRequest.protectedRangeId = protectedRange.protectedRangeId;
        const request = Sheets.newRequest();
        request.deleteProtectedRange = deleteProtectedRangeRequest;
        requests.push(request);
      }
    });
  });
  if (requests.length) {
    const resource = Sheets.newBatchUpdateSpreadsheetRequest();
    resource.requests = requests;
    resource.responseIncludeGridData = false;
    Sheets.Spreadsheets.batchUpdate(resource, this.settings.APP_CURRENT_ID);
    this.book.sheets = sheets;
  }
};
