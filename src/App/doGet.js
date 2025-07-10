/* global App */
/**
 * Обрабатывает GET-запросы.
 * @param {GoogleAppsScript.Events.DoGet} e Объект события.
 * @returns {GoogleAppsScript.Content.TextOutput} JSON-ответ.
 */
App.prototype.doGet = function (e) {
  const out = { error: undefined, data: undefined, action: undefined };

  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
};
