
/**
 * Remove empty sheets.
 */
function onOpen() {
  // Add menu
  const menu = SpreadsheetApp.getUi().createMenu(APP_TITLE);
  menu
    .addItem('Regénérer calendrier', 'generateCalendar')
    .addItem('Mettre à jour la liste des inscrits', 'updatePeopleOnly')
    .addItem('Ajouter un inscrit', 'addOnePerson')
    .addToUi();

  // Remove sheets with default name
  // let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // var sheets = activeSpreadsheet.getSheets();
  // for (var n = 0; n < sheets.length; n++) {
  //   if (sheets[n].getName().startsWith('Feuille')) {
  //     try {
  //       activeSpreadsheet.deleteSheet(sheets[n]);
  //     } catch (err) {
  //       Browser.msgBox('Can\'t delete Sheet named "' + sheets[n].getName() + '" (' + err + ')');
  //     }
  //   }
  // }
}