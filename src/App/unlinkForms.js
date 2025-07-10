/* global App */
/**
 * Отвязывает и удаляет все формы, связанные с текущей книгой.
 */
App.prototype.unlinkForms = function () {
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
};
