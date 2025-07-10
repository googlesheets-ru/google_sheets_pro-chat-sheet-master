/* global Group */

/* exported App*/
class App {
  /**
   * @param {App.Settings} [settings] Начальные настройки.
   */
  constructor(settings) {
    this._settings = settings;
  }

  /**
   * Текущая книга (таблица).
   * @type {GoogleAppsScript.Spreadsheet.Spreadsheet}
   */
  get currentBook() {
    if (!this._currentBook) this._currentBook = SpreadsheetApp.openById(this.settings.APP_CURRENT_ID);
    return this._currentBook;
  }

  /**
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} book
   */
  set currentBook(book) {
    this._currentBook = book;
  }

  get group() {
    if (!this._group) {
      this._group = new Group(this.settings);
    }
    return this._group;
  }

  /**
   * Родительская папка.
   * @type {GoogleAppsScript.Drive.Folder}
   */
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

  /**
   * Полная информация о книге из Sheets API.
   * @type {GoogleAppsScript.Sheets.Schema.Spreadsheet}
   */
  get book() {
    if (!this._book) this.recallBook();
    return this._book;
  }
}
