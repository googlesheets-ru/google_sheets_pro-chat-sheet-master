/**
 * @fileoverview Пользовательские функции, которые можно вызывать из меню Google Таблиц.
 */
/* global App */

/* exported userActionGenerateTOC */
/**
 * Пользовательское действие для генерации оглавления в определенной книге.
 * ID книги захардкожен.
 */
function userActionGenerateTOC() {
  const TARGET_BOOK_ID = '157EFj6LWune_iP5Lc59aFzaeUWgQgVJV19yP3_a3WQM';

  const app = new App();
  app.currentBook = SpreadsheetApp.openById(TARGET_BOOK_ID);
  app.generateTOC();
}

/* exported userActionUpdateAllBooks */
/**
 * Пользовательское действие для запуска обновления всех книг.
 */
function userActionUpdateAllBooks() {
  new App().updateAllBooks();
}

/* exported userActionCleanBook */
/**
 * Пользовательское действие для очистки текущей книги.
 * Удаляет все листы, кроме списка исключений.
 */
function userActionCleanBook() {
  const excludes = ['формула по выпадающему списку'];
  new App().cleanBook({ excludes });
}

/* exported userActionCleanOldSheet */
/**
 * Пользовательское действие для "обнуления" старого листа.
 * Снимает защиту со всех листов и обновляет оглавление.
 * Использует захардкоженные настройки для конкретной таблицы.
 */
function userActionCleanOldSheet() {
  const app = new App({
    APP_CURRENT_FILE_NUM: '8',
    APP_CURRENT_ID: '1MqeW7LkEUcsDH8lUksXBLHWmSeu8up2Jr_3ujVwE4SE',
    APP_FOLDER_ID: '1mgzpM6dID_GUnzo-aQAAEv3kpFEmtgPx',
    APP_LIST_OF_EXEPTIONS_SHEETS: '["О Таблице"]',
  });
  app.releaseSheets();
  app.generateTOC();
}
