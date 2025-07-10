/* global App */
/**
 * Обрабатывает POST-запросы.
 * @param {GoogleAppsScript.Events.DoPost} e Объект события.
 * @returns {GoogleAppsScript.Content.TextOutput} JSON-ответ.
 */
App.prototype.doPost = function (e) {
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
  return ContentService.createTextOutput(JSON.stringify({ result: contents, res: contents.access_token })).setMimeType(
    ContentService.MimeType.JSON,
  );
};
