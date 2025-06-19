/* global App */

/* exported init */
/**
 * Используется для инициализации в качестве библиотеки. Каждый новый init создает новый экземпляр приложения
 * @param  {...any} args
 * @returns
 */
function init(...args) {
  return new App(...args);
}
