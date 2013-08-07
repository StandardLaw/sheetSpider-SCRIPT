function updateFeederSheet() {
  var sure = Browser.msgBox("Are you sure?", "This action will overwrite all existing values in your feeder sheet with the values currently in the \"Returned Values\" sheet.  Proceed?", Browser.Buttons.OK_CANCEL);
  if (sure == "ok") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var properties = ScriptProperties.getProperties();
    var returnedDataSheet = ss.getSheetByName('Returned Data');
    if (!returnedDataSheet) {
      Browser.msgBox("You must retrieve data from the entity spreadsheets (Step 4) before you can run this step.");
    }
    var properties = ScriptProperties.getProperties();
    var normalizedFeederEntityCol = normalizeHeader(properties.feederEntityCol);
    var feederSheet = getSheetById(ss, properties.feederSheetId);
    var feederSheetName = properties.feederSheetName;
    var feederSheetHeaders = feederSheet.getRange(1,1,1,feederSheet.getLastColumn()).getValues()[0];
    var feederRange = feederSheet.getRange(2, 1, feederSheet.getLastRow(), feederSheet.getLastColumn());
    feederRange.clearContent();
    var returnedDataRange = returnedDataSheet.getRange(2, 1, returnedDataSheet.getLastRow()-1, returnedDataSheet.getLastColumn());
    var returnedData = getRowsData(returnedDataSheet, returnedDataRange);
    var lastCol = feederSheet.getLastColumn();
    feederRange = feederSheet.getRange(2, 1, returnedData.length, lastCol);
    var count = 0;
    var addedCount = 0;
    for (var i = 0; i<returnedData.length; i++) {
      var changeArray = returnedData[i].changeStatus.split("||");
      if (changeArray[0] == 'edited') {   
        count++;
      }
      if (changeArray[0] == 'added') {
        addedCount++;
      }
    }
    
    setRowsData(feederSheet, returnedData);
    var notFoundDataSheet = ss.getSheetByName('Records Not Found');
    var deletedCount = 0;
    if (notFoundDataSheet) {
      var notFoundDataRange = notFoundDataSheet.getRange(2, 1, notFoundDataSheet.getLastRow(), notFoundDataSheet.getLastColumn());
      var notFoundData = getRowsData(notFoundDataSheet, notFoundDataRange);
      for (var i = 0; i<notFoundData.length; i++) {
        if (notFoundData[i].changeStatus == 'deleted') {
          deletedCount++
        }
      }
    }
    sheetSpider_logFeederDataUpdated();
    ss.setActiveSheet(feederSheet);
    feederSheet.setFrozenRows(1);
    Browser.msgBox("The feeder sheet was cleared of values and updated records written to it. This included " + count + " update(s), " + addedCount + " row addition(s), and " + deletedCount + " row deletion(s).");
    properties.stepFourComplete = false;
    ScriptProperties.setProperties(properties);
    onOpen();
  }
  return;
}
