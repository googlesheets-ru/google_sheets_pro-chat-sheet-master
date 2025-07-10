/* global App */
/**
 * Создает и настраивает следующую книгу (таблицу) чата.
 */
App.prototype.createNextBook = function () {
  const currentBook = DriveApp.getFileById(this.settings.APP_CURRENT_ID);
  const num = Number(this.settings.APP_CURRENT_FILE_NUM) + 1;
  const copy = currentBook.makeCopy(`Таблица чата t.me/google_sheets_pro #${num}`, this.folder);
  const fileMetadata = {
    writersCanShare: false,
  };
  Drive.Files.update(fileMetadata, copy.getId());
  copy.getEditors().forEach((editor) => copy.removeEditor(editor));
  copy.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  if (this.settings.APP_MASTER_EDITOR) {
    copy.addEditor(this.settings.APP_MASTER_EDITOR);
  }
  if (this.settings.APP_EXPERTS_EDITOR) {
    copy.addEditors(this.settings.APP_EXPERTS_EDITOR.split(',').map((email) => email.trim()));
  }
  const settings = { ...this.settings, ...{ APP_CURRENT_FILE_NUM: `${num}`, APP_CURRENT_ID: copy.getId() } };
  const prevUrl = currentBook.getUrl();
  const prevTitle = currentBook.getName();
  this.settings = settings;
  this.updateStamp({ num, prevUrl, prevTitle });
  this.unlinkForms();
  this.cleanBook({ excludes: [] });
  this.addNewBlankUserSheet();
  this.orderSheetsByProtections();
  this.generateTOC();
};
