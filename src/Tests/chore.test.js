/**
 * @fileoverview Утилиты для обслуживания проекта.
 */
/* exported choreUnlinkForms */

/**
 * Отвязывает и удаляет все формы, связанные с Google Таблицей.
 *
 * Функция итерирует по всем листам в указанной книге. Если лист связан с формой,
 * функция отвязывает форму от листа и перемещает файл формы в корзину.
 * ID книги захардкожен в функции.
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
