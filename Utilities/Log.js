
/**
 * @param {String} msg
 */
function log(msg) {
  console.log(msg);
}


/**
 * @param {String} msg
 */
function info(msg) {
  log(msg);

  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  activeSpreadsheet.toast(msg, APP_TITLE, 10);
}


/**
 * @param {String} msg
 * @param {Error} e
 */
function err(msg, e) {
  if (e) {
    console.error(`${msg}\u000D\u000A${e}\u000D\u000A${e.stack}`);
    Browser.msgBox(`${msg}\u000D\u000A${e}\u000D\u000A${e.stack}`);
  } else {
    console.error(msg);
    Browser.msgBox(msg);
  }
}