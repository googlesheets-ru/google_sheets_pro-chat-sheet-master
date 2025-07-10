/* global App */
/**
 * Сортирует Таблицу заданным образом.
 * Сначала идут листы из списка исключений `APP_LIST_OF_EXEPTIONS_SHEETS` в том порядке, в котором они указаны в списке.
 * Затем идут защищенные листы.
 * В конце идут все остальные листы.
 *
 * @returns {void} Ничего не возвращает. Обновляет порядок листов в текущей книге.
 */
App.prototype.orderSheetsByProtections = function () {
  const exeptionSheetNames = this.settings.APP_LIST_OF_EXEPTIONS_SHEETS;
  const sorted = JSON.parse(JSON.stringify(this.book.sheets)).sort((a, b) => {
    const excludesA = exeptionSheetNames.indexOf(a.properties.title);
    const excludesB = exeptionSheetNames.indexOf(b.properties.title);
    if (excludesA > -1 && excludesB > -1) return excludesA - excludesB;
    if (excludesA > -1) return -1;
    if (excludesB > -1) return 1;
    const protectedA = a.protectedRanges?.some((r) => Object.keys(r.range).length === 1) ?? false;
    const protectedB = b.protectedRanges?.some((r) => Object.keys(r.range).length === 1) ?? false;
    return protectedA - protectedB;
  });
  const requests = [];
  requests.push(
    ...sorted.map((sheet, index) => {
      const updatePropertiesRequest = Sheets.newUpdateSheetPropertiesRequest();
      updatePropertiesRequest.fields = 'index';
      updatePropertiesRequest.properties = {
        index,
        sheetId: sheet.properties.sheetId,
      };
      const request = Sheets.newRequest();
      request.updateSheetProperties = updatePropertiesRequest;
      return request;
    }),
  );
  const resource = Sheets.newBatchUpdateSpreadsheetRequest();
  resource.requests = requests;
  resource.responseIncludeGridData = false;
  Sheets.Spreadsheets.batchUpdate(resource, this.settings.APP_CURRENT_ID);
  this.book.sheets = sorted;
};
