/* global App */
/**
 * Добавляет технический штамп в "О Таблице".
 * @param {{num: number, prevUrl: string, prevTitle: string}} stamp
 * @param {number} stamp.num Новый номер таблицы.
 * @param {string} stamp.prevUrl URL предыдущей таблицы.
 * @param {string} stamp.prevTitle Имя предыдущей таблицы.
 */
App.prototype.updateStamp = function ({ num, prevUrl, prevTitle }) {
  if (this.book.sheets.some((sheet) => sheet.properties.title === 'О Таблице'))
    Sheets.Spreadsheets.Values.update({ values: [[num], [prevUrl, prevTitle]] }, this.book.spreadsheetId, 'О Таблице', {
      valueInputOption: 'RAW',
    });
};
