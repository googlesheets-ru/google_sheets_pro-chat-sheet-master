/* global */
/* exported App */
class App {
  /**
   *
   * @param {App.Settings} settings
   */
  constructor(settings) {
    this._settings = settings;
  }

  /**
   *
   * @param {GoogleAppsScript.Events.DoGet} e
   * @returns {GoogleAppsScript.Content.TextOutput}
   */
  doGet(e) {
    const out = { error: undefined, data: undefined, action: undefined };

    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
  }

  /**
   *
   * @param {GoogleAppsScript.Events.DoPost} e
   * @returns
   */
  doPost(e) {
    // return ContentService.createTextOutput(JSON.stringify({ result: JSON.stringify(e.postData) })).setMimeType(ContentService.MimeType.JSON);
    const out = { error: undefined, data: undefined, action: undefined };
    const contents = JSON.parse(e.postData.contents);

    if (
      contents.access_token &&
      contents.access_token === this.settings.ADMIN_ACCESS_TOKEN &&
      contents.action === 'get_app_current_id'
    ) {
      out.data = {
        APP_CURRENT_ID: this.settings.APP_CURRENT_ID,
      };
      out.action = contents.action;
      return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(
      JSON.stringify({ result: contents, res: contents.access_token }),
    ).setMimeType(ContentService.MimeType.JSON);
  }

  /**
   * @type {GoogleAppsScript.Spreadsheet.Spreadsheet}
   */
  get currentBook() {
    if (!this._currentBook) this._currentBook = SpreadsheetApp.openById(this.settings.APP_CURRENT_ID);
    return this._currentBook;
  }

  set currentBook(book) {
    this._currentBook = book;
  }

  get folder() {
    if (!this._folder) this._folder = DriveApp.getFolderById(this.settings.APP_FOLDER_ID);
    return this._folder;
  }

  /**
   * @returns {App.Settings}
   */
  get settings() {
    if (!this._settings) {
      this._settings = PropertiesService.getScriptProperties().getProperties();
      this._settings.APP_LIST_OF_EXEPTIONS_SHEETS = JSON.parse(this._settings.APP_LIST_OF_EXEPTIONS_SHEETS);
    }
    return this._settings;
  }

  /**
   * @param {App.Settings} settings
   */
  set settings(settings) {
    const data = JSON.parse(JSON.stringify(settings));
    if (this._settings && this._settings.APP_CURRENT_ID !== data.APP_CURRENT_ID) this._book = undefined;
    data.APP_LIST_OF_EXEPTIONS_SHEETS = JSON.stringify(data.APP_LIST_OF_EXEPTIONS_SHEETS || '[]');
    PropertiesService.getScriptProperties().setProperties(data, false);
    this._settings = undefined;
  }

  get book() {
    if (!this._book) this.recallBook();
    return this._book;
  }

  recallBook() {
    this._book = Sheets.Spreadsheets.get(this.currentBook.getId(), {
      includeGridData: false,
      fields: 'spreadsheetId,sheets(properties(sheetId,index,title),protectedRanges(range,protectedRangeId))',
    });
  }

  /**
   * Сортирует Таблицу заданным образом.
   * Сначала идут листы из списка исключений `APP_LIST_OF_EXEPTIONS_SHEETS` в том порядке, в котором они указаны в списке.
   * Затем идут защищенные листы.
   * В конце идут все остальные листы.
   *
   * @returns {void} Ничего не возвращает. Обновляет порядок листов в текущей книге.
   */
  orderSheetsByProtections() {
    // Получаем список имен листов-исключений из настроек
    const exeptionSheetNames = this.settings.APP_LIST_OF_EXEPTIONS_SHEETS;
    // Создаем глубокую копию массива листов и сортируем его
    const sorted = JSON.parse(JSON.stringify(this.book.sheets)).sort((a, b) => {
      // Определяем, является ли лист A исключением
      const excludesA = exeptionSheetNames.indexOf(a.properties.title);
      // Определяем, является ли лист B исключением
      const excludesB = exeptionSheetNames.indexOf(b.properties.title);
      // Если оба листа - исключения, сортируем их по порядку в списке исключений
      if (excludesA > -1 && excludesB > -1) return excludesA - excludesB;
      // Если только лист A - исключение, он идет первым
      if (excludesA > -1) return -1;
      // Если только лист B - исключение, он идет первым
      if (excludesB > -1) return 1;
      // Определяем, защищен ли лист A (защищен, если есть хотя бы один защищенный диапазон, покрывающий весь лист)
      const protectedA = a.protectedRanges?.some((r) => Object.keys(r.range).length === 1) ?? false;
      // Определяем, защищен ли лист B
      const protectedB = b.protectedRanges?.some((r) => Object.keys(r.range).length === 1) ?? false;
      // Сортируем по признаку защищенности (защищенные идут раньше)
      return protectedA - protectedB;
    });
    // Создаем массив запросов для batchUpdate
    const requests = [];
    // Заполняем массив запросами на обновление индекса (позиции) каждого листа
    requests.push(
      ...sorted.map((sheet, index) => {
        // Создаем запрос на обновление свойств листа
        const updatePropertiesRequest = Sheets.newUpdateSheetPropertiesRequest();
        // Указываем, что обновляем только индекс
        updatePropertiesRequest.fields = 'index';
        // Устанавливаем новые свойства: новый индекс и ID листа
        updatePropertiesRequest.properties = {
          index,
          sheetId: sheet.properties.sheetId,
        };
        // Создаем общий запрос
        const request = Sheets.newRequest();
        // Добавляем в него запрос на обновление свойств листа
        request.updateSheetProperties = updatePropertiesRequest;
        return request;
      }),
    );
    // Создаем ресурс для batchUpdate
    const resource = Sheets.newBatchUpdateSpreadsheetRequest();
    // Добавляем массив запросов в ресурс
    resource.requests = requests;
    // Указываем, что не нужно возвращать данные сетки в ответе (для оптимизации)
    resource.responseIncludeGridData = false;
    // Выполняем batchUpdate для обновления позиций листов
    Sheets.Spreadsheets.batchUpdate(resource, this.settings.APP_CURRENT_ID);
    // Обновляем порядок листов в локальном объекте книги
    this.book.sheets = sorted;
  }

  /**
   * Добавляет новый лист для примера
   */
  addNewBlankUserSheet() {
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

      console.log(JSON.stringify(this.book, null, '  '));
    }
  }

  /**
   * Обновляет содержание
   */
  generateTOC() {
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
  }

  /**
   * "Обнуляет" Таблицу
   */
  cleanBook(excludes) {
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
  }

  unlinkForms() {
    this.currentBook.getSheets().forEach((sheet) => {
      const formUrl = sheet.getFormUrl();
      if (formUrl) {
        const form = FormApp.openByUrl(formUrl);
        form.removeDestination();
        const id = form.getId();
        DriveApp.getFileById(id).setTrashed(true);
        console.info(`Unlinked form '${id}' from sheet: ${sheet.getName()}`);
      }
    });
  }

  /**
   * Добавляет технический штамп для Таблицы
   *
   * @param {*} param0
   */
  updateStamp({ num, prevUrl, prevTitle }) {
    if (this.book.sheets.some((sheet) => sheet.properties.title === 'О Таблице'))
      Sheets.Spreadsheets.Values.update(
        { values: [[num], [prevUrl, prevTitle]] },
        this.book.spreadsheetId,
        'О Таблице',
        {
          valueInputOption: 'RAW',
        },
      );
  }

  releaseSheets() {
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
  }

  // updateAllBooks() {
  //   const SETTINGS = {
  //     D2: {
  //       text: 'Канал "Таблицы Гугл" t.me/GoogleSheets_ru',
  //       link: 'https://t.me/+lmannExYEyg5OTZi',
  //       startOffset: 21,
  //       rangeA1: 'D2',
  //     },
  //     D3: {
  //       text: 'Таблицы и Скрипты Гугл - чат t.me/google_sheets_pro',
  //       link: 'https://t.me/+pLLUBtcXIqY0MGMy',
  //       startOffset: 29,
  //       rangeA1: 'D3',
  //     },
  //     D4: {
  //       text: 'Чат по Apps Script для специалистов t.me/googleappsscriptrc',
  //       link: 'https://t.me/+7HbI3eq42C80MmMy',
  //       startOffset: 36,
  //       rangeA1: 'D4',
  //     },
  //   };

  //   const { rangeA1, text, link, startOffset } = SETTINGS.D4;

  //   const textStyle = SpreadsheetApp.newTextStyle();
  //   textStyle.setForegroundColor('#434343');
  //   textStyle.setFontSize(10);
  //   textStyle.setItalic(true);
  //   const ts = textStyle.build();
  //   console.log(this.settings.APP_FOLDER_ID);
  //   const books = DriveApp.searchFiles(
  //     `'${this.settings.APP_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet'`,
  //   );

  //   const out = [];

  //   while (books.hasNext()) {
  //     const book = SpreadsheetApp.openById(books.next().getId());
  //     console.log(book.getName());
  //     const sheet = book.getSheetByName('О Таблице');
  //     if (sheet) {
  //       const range = sheet.getRange(rangeA1);
  //       const item = {};
  //       item.name = book.getName();
  //       item.url = book.getUrl();
  //       item.value = range.getValue();
  //       item.rtv = range.getRichTextValue().getLinkUrl();
  //       out.push(item);

  //       if (item.value !== text || item.rtv !== link) {
  //         const nrtv = SpreadsheetApp.newRichTextValue();
  //         nrtv.setText(text);
  //         nrtv.setTextStyle(ts);
  //         nrtv.setLinkUrl(startOffset, text.length, link);
  //         range.setRichTextValue(nrtv.build());
  //         range.setHorizontalAlignment('right').setVerticalAlignment('middle');
  //       }
  //     }
  //   }

  //   out
  //     .sort((a, b) => {
  //       const aN = Number(a.name.replace(/.*#(\d+).*/, '$1')) || 0;
  //       const bN = Number(b.name.replace(/.*#(\d+).*/, '$1')) || 0;
  //       if (aN < bN) return -1;
  //       if (aN > bN) return 1;
  //       return 0;
  //     })
  //     .forEach((item) => console.log(item));
  // }

  resetName() {
    let title = 'Таблица чата ';
    const bookNameRange = this.currentBook.getRangeByName('BOOK_NAME');
    if (bookNameRange) {
      title = bookNameRange.getValue();
    } else {
      const num = this.currentBook.getSheetByName('О Таблице').getRange('A1').getValue();
      title += `#${num}`;
    }
    DriveApp.getFileById(this.settings.APP_CURRENT_ID).setName(title);
  }

  createNextBook() {
    // Получаем текущий файл таблицы по ID из настроек
    const currentBook = DriveApp.getFileById(this.settings.APP_CURRENT_ID);
    // Увеличиваем номер текущего файла на 1
    const num = Number(this.settings.APP_CURRENT_FILE_NUM) + 1;
    // Создаем копию текущей таблицы с новым именем в указанной папке
    const copy = currentBook.makeCopy(`Таблица чата t.me/google_sheets_pro #${num}`, this.folder);
    copy.getEditors().forEach((editor) => copy.removeEditor(editor));
    // Устанавливаем права доступа на редактирование для всех
    copy.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    if (this.settings.APP_MASTER_EDITOR) {
      copy.addEditor(this.settings.APP_MASTER_EDITOR);
    }
    if (this.settings.APP_EXPERTS_EDITOR) {
      copy.addEditors(this.settings.APP_EXPERTS_EDITOR.split(',').map((email) => email.trim()));
    }
    // Создаем настройки для новой Таблицы
    const settings = { ...this.settings, ...{ APP_CURRENT_FILE_NUM: `${num}`, APP_CURRENT_ID: copy.getId() } };
    const newApp = new App(settings);

    newApp.unlinkForms();

    // Обновляем технический штамп в новой таблице с информацией о предыдущей таблице
    newApp.updateStamp({ num, prevUrl: currentBook.getUrl(), prevTitle: currentBook.getName() });
    // "Обнуляем" новую таблицу (удаляем ненужные листы)
    newApp.cleanBook();
    // Добавляем новый пустой лист для пользователя
    newApp.addNewBlankUserSheet();
    // Сортируем листы в новой таблице
    newApp.orderSheetsByProtections();
    // Генерируем оглавление для новой таблицы
    newApp.generateTOC();

    if (this.settings.APP_MASTER_EDITOR) {
      currentBook.removeEditor(this.settings.APP_MASTER_EDITOR);
    }

    this.releaseSheets();

    // Обновление прошло успешно - обновляем настройки приложения и выходми
    this.settings = settings;
  }
}
