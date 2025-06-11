/* exported choreUnlinkForms */

/**
 * Отвязывает и удаляет все формы, связанные с указанной Google Таблицей.
 * @see {@link https://stackoverflow.com/a/54587996/1393023}
 */
function choreUnlinkForms() {
  const book = SpreadsheetApp.openById('1EZtaf3Gbnhj3AnaN6Ewlw6iE__4Se2adQcgtKzVkvxs');
  book.getSheets().forEach((sheet) => {
    const formUrl = sheet.getFormUrl();
    if (formUrl) {
      const form = FormApp.openByUrl(formUrl);
      form.removeDestination();
      const id = form.getId();
      DriveApp.getFileById(id).setTrashed(true);
      console.info(`Unlinked form ${form.getId()} from sheet: ${sheet.getName()}`);
    }
  });
}
