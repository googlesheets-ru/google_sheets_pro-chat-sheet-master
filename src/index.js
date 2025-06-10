/* global App */
/* exported init */
function init(...args) {
  return new App(...args);
}

/* exported customUpdate */
function customUpdate() {
  const app = new App();

  app.currentBook = SpreadsheetApp.openById('157EFj6LWune_iP5Lc59aFzaeUWgQgVJV19yP3_a3WQM');
  app.generateTOC();
}

/* exported updateAllBooks */
function updateAllBooks() {
  new App().updateAllBooks();
}
