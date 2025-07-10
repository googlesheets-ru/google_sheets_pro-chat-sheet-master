/* global App */
/**
 * Обновляет содержание
 */
App.prototype.generateTOC = function () {
  const excludeSheetNames = this.settings.APP_LIST_OF_EXEPTIONS_SHEETS;

  const richTextValues = this.book.sheets
    .filter((item) => excludeSheetNames.indexOf(item.properties.title) === -1)
    .map((item) => [
      SpreadsheetApp.newRichTextValue()
        .setText(
          `${item.protectedRanges?.some((r) => Object.keys(r.range).length === 1) ?? false ? '🔏 ' : ''}${
            item.properties.title
          }`,
        )
        .setLinkUrl(`#gid=${item.properties.sheetId}`)
        .build(),
    ]);
  const firstPage = this.currentBook.getSheetByName('О Таблице');
  const pos = firstPage.createTextFinder('Содержание').matchCase(true).matchFormulaText(false).findNext();
  if (pos) {
    const range = firstPage.getRange(pos.getRow() + 1, pos.getColumn(), firstPage.getLastRow()).clearContent();
    range.offset(0, 0, richTextValues.length, richTextValues[0].length).setRichTextValues(richTextValues);
  }
};
