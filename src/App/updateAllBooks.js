/* global App*/

/**
 * Обновляет ссылки в футере на всех связанных книгах.
 * Находит все таблицы в папке `APP_FOLDER_ID`, и в каждой из них на листе "О Таблице"
 * в ячейке D4 обновляет ссылку на чат специалистов по Apps Script.
 * Также форматирует текст и выводит в консоль отсортированный список обработанных таблиц.
 *
 * Это кастомная функция для канала "Таблицы Гугл"
 */
App.prototype.updateAllBooks = function () {
  const SETTINGS = {
    D2: {
      text: 'Канал "Таблицы Гугл" t.me/GoogleSheets_ru',
      link: 'https://t.me/+lmannExYEyg5OTZi',
      startOffset: 21,
      rangeA1: 'D2',
    },
    D3: {
      text: 'Таблицы и Скрипты Гугл - чат t.me/google_sheets_pro',
      link: 'https://t.me/+pLLUBtcXIqY0MGMy',
      startOffset: 29,
      rangeA1: 'D3',
    },
    D4: {
      text: 'Чат по Apps Script для специалистов t.me/googleappsscriptrc',
      link: 'https://t.me/+7HbI3eq42C80MmMy',
      startOffset: 36,
      rangeA1: 'D4',
    },
  };

  const { rangeA1, text, link, startOffset } = SETTINGS.D4;

  const textStyle = SpreadsheetApp.newTextStyle();
  textStyle.setForegroundColor('#434343');
  textStyle.setFontSize(10);
  textStyle.setItalic(true);
  const ts = textStyle.build();
  console.log(this.settings.APP_FOLDER_ID);
  const books = DriveApp.searchFiles(
    `'${this.settings.APP_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet'`,
  );

  const out = [];

  while (books.hasNext()) {
    const book = SpreadsheetApp.openById(books.next().getId());
    console.log(book.getName());
    const sheet = book.getSheetByName('О Таблице');
    if (sheet) {
      const range = sheet.getRange(rangeA1);
      const item = {};
      item.name = book.getName();
      item.url = book.getUrl();
      item.value = range.getValue();
      item.rtv = range.getRichTextValue().getLinkUrl();
      out.push(item);

      if (item.value !== text || item.rtv !== link) {
        const nrtv = SpreadsheetApp.newRichTextValue();
        nrtv.setText(text);
        nrtv.setTextStyle(ts);
        nrtv.setLinkUrl(startOffset, text.length, link);
        range.setRichTextValue(nrtv.build());
        range.setHorizontalAlignment('right').setVerticalAlignment('middle');
      }
    }
  }

  out
    .sort((a, b) => {
      const aN = Number(a.name.replace(/.*#(\d+).*/, '$1')) || 0;
      const bN = Number(b.name.replace(/.*#(\d+).*/, '$1')) || 0;
      if (aN < bN) return -1;
      if (aN > bN) return 1;
      return 0;
    })
    .forEach((item) => console.log(item));
};
