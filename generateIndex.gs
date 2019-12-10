function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Index', 'generateIndex')
      .addToUi();
}

function generateIndex() {
  var indexSheetNames = new Array();
  var indexSheetIds = new Array();
  var currentSheet = SpreadsheetApp.getActiveSheet();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  sheets.forEach(function(sheet){
    indexSheetNames.push([sheet.getSheetName()]);
    indexSheetIds.push(['=HYPERLINK("#gid=' 
                        + sheet.getSheetId() 
                        + '")']);
  });
  
  currentSheet.getRange(1,1).setValue('Index').setFontWeight('bold');
  currentSheet.getRange(3,1,indexSheetNames.length,1).setValues(indexSheetNames);
  currentSheet.getRange(3,2,indexSheetIds.length,1).setFormulas(indexSheetIds);
}
