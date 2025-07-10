/* global App */
/**
 * Устанавливает имя книги на основе данных из ячеек.
 */
App.prototype.resetName = function () {
  let title = 'Таблица чата ';
  const bookNameRange = this.currentBook.getRangeByName('BOOK_NAME');
  if (bookNameRange) {
    title = bookNameRange.getValue();
  } else {
    const num = this.currentBook.getSheetByName('О Таблице').getRange('A1').getValue();
    title += `#${num}`;
  }
  DriveApp.getFileById(this.settings.APP_CURRENT_ID).setName(title);
};
