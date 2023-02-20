class App {
  constructor() {
    /** @type {App.Settings} */
    // this.settings =
  }

  get currentBook() {
    if (!this._currentBook) this._currentBook = SpreadsheetApp.openById(this.settings.APP_CURRENT_ID);
    return this._currentBook;
  }

  get folder() {
    if (!this._folder) this._folder = DriveApp.getFolderById(this.settings.APP_FOLDER_ID);
    return this._folder;
  }

  get settings() {
    if (!this._settings) this._settings = PropertiesService.getScriptProperties().getProperties();
    return this._settings;
  }

  /**
   * @param {App.Settings} settings
   */
  set settings(settings) {
    console.log(settings);
    PropertiesService.getScriptProperties().setProperties(settings, false);
    this._settings = undefined;
  }
}
