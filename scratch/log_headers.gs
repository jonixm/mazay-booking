function logHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const splav = ss.getSheetByName('Сплавы');
  if (!splav) {
    console.log('Sheet "Сплавы" not found');
    return;
  }
  const headers = splav.getRange(1, 1, 1, splav.getLastColumn()).getValues()[0];
  console.log('Headers: ' + JSON.stringify(headers));
}
