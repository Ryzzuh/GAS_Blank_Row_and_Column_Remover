let delBlankColumns = (sheet, featherAmount) => {
  let currentSheet = sheet || SpreadsheetApp.getActive().getActiveSheet()
  const lastColumn = currentSheet.getLastColumn()
  const numBlankColumns = currentSheet.getMaxColumns() - lastColumn
  if(numBlankColumns>0){
    currentSheet.deleteColumns(lastColumn+1, numBlankColumns)
  }
}

let delBlankRows = (sheet, featherAmount) => {
  let currentSheet = sheet || SpreadsheetApp.getActive().getActiveSheet()
  const lastRow = currentSheet.getLastRow()
  const numBlankRows = currentSheet.getMaxRows() - lastRow
  if(numBlankRows>0){
    currentSheet.deleteRows(lastRow+1, numBlankRows)
  }
}

