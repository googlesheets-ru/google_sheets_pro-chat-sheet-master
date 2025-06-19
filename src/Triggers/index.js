/**
 * @fileoverview Триггеры, которые выполняются по расписанию.
 */
/* global App */

/* exported triggerUpdateEveryMonth */
/**
 * Триггер создания новой Таблицы чата.
 * Выполняется ежемесячно.
 * Копирует текущую таблицу, обновляет настройки и подготавливает новую таблицу к использованию.
 */
function triggerUpdateEveryMonth() {
  // Создаем экземпляр приложения App
  const app = new App();
  app.createNextBook();
}

/* exported triggerUpdateEveryHour */
/**
 * Триггер ежечасного обновления Таблицы.
 * Добавляет новый пустой лист, сортирует листы и обновляет оглавление.
 */
function triggerUpdateEveryHour() {
  const app = new App();
  app.addNewBlankUserSheet();
  app.orderSheetsByProtections();
  app.generateTOC();
}

/* exported triggerUpdateEveryMin */
/**
 * Триггер ежеминутного обновления Таблицы.
 * Используется для отладки. Переименовывает текущую книгу.
 */
function triggerUpdateEveryMin() {
  const app = new App();
  app.resetName();
}
