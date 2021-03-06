// remove blank columns and rows to free up resources from sheets
// use a feather parameter to leave n blank columns and rows

const onOpen = () => {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('__Sheet_Tools__')
      .addSeparator()
      .addItem('Delete extra Rows and Columns (for this sheet)', 'delForThisSheet')
      .addItem('Delete extra Rows and Columns (for all sheets)', 'delForAllSheets')
      .addSeparator()
      .addToUi();
}

const showPrompt = () => {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.prompt(
      'Options',
      'How many columns and rows to leave blank? (Default=0)',
      ui.ButtonSet.OK_CANCEL);
  // Process the user's response.
  var button = result.getSelectedButton();
  var num = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    if(num==''){return 0}
    return num
  } return -1
}

const delForThisSheet = () => {
  const feather = showPrompt()
  const sheet = SpreadsheetApp.getActive().getActiveSheet()
  //if(sheet.getLastRow()>0){
    delBlankColumns(sheet, feather)
    delBlankRows(sheet, feather)
  //}
}

const delForAllSheets = () => {
  const feather = showPrompt()
  const sheets = SpreadsheetApp.getActive().getSheets()
  for(sheet of sheets){
    //if(sheet.getLastRow()>0){
      delBlankColumns(sheet, feather)
      delBlankRows(sheet, feather)
    //}
  } 
}


let delBlankColumns = (sheet, feather = 0) => {
  let currentSheet = sheet || SpreadsheetApp.getActive().getActiveSheet()
  const lastColumn = currentSheet.getLastColumn() || 1
  console.log(currentSheet.getLastColumn())
  const numBlankColumns = currentSheet.getMaxColumns() - lastColumn
  if(numBlankColumns - feather >0){
    currentSheet.deleteColumns(lastColumn+1, numBlankColumns - feather)
  }
}

let delBlankRows = (sheet, feather = 0) => {
  let currentSheet = sheet || SpreadsheetApp.getActive().getActiveSheet()
  const lastRow = currentSheet.getLastRow() || 1
  const numBlankRows = currentSheet.getMaxRows() - lastRow
  if(numBlankRows - feather >0){
    currentSheet.deleteRows(lastRow+1, numBlankRows)
  }
}
