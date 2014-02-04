function sheetSpider_preconfig() {
  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  ScriptProperties.setProperty('ssId', ssId);
  // if you are interested in sharing your complete workflow system for others to copy (with script settings)
  // Select the "Generate preconfig()" option in the menu and
  //#######Paste preconfiguration code below before sharing your system for copy#######
  
  
  
  
  
  //#######End preconfiguration code#######
  ScriptProperties.setProperty('preconfigStatus', 'true'); 
  //Fetch system name, if this script is part of a New Visions system
  var systemName = NVSL.getSystemName();
  if (systemName) {
    ScriptProperties.setProperty('systemName', systemName);
  }
  var institutionalTrackingString = NVSL.checkInstitutionalTrackingCode();
  if (institutionalTrackingString) {
    onOpen();
  }
}


function sheetSpider_extractorWindow () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var propertyString = '';
  var excludedProperties = ['preconfigStatus','sheetSpider_uid','sheetSpider_sid', 'ssKey','primaryFolderKey']
  for (var key in properties) {
    if ((properties[key]!='')&&(excludedProperties.indexOf(key)==-1)) {
      var keyProperty = properties[key].replace(/[/\\*]/g, "\\\\");                                     
      propertyString += "   ScriptProperties.setProperty('" + key + "','" + keyProperty + "');\n";
    }
  }
  var app = UiApp.createApplication().setHeight(500).setTitle("Export preconfig() settings");
  var panel = app.createVerticalPanel().setWidth("100%").setHeight("100%");
  var labelText = "Copying a Google Spreadsheet copies scripts along with it, but without any of the script settings saved.  This normally makes it hard to share full, script-enabled Spreadsheet systems. ";
  labelText += " You can solve this problem by pasting the code below into a script function called \"sheetSpider_preconfig\" (go to sheetSpider in the Script Editor and select \"preconfig.gs\" in the left sidebar) prior to publishing your Spreadsheet for others to copy. \n";
  labelText += " After a user copies your spreadsheet, they will select \"Run initial configuration.\"  This will preconfigure all needed script settings.  If you got this workflow from someone as a copy of a spreadsheet, this has probably already been done for you.";
  var label = app.createLabel(labelText);
  var window = app.createTextArea();
  var codeString = "//This section sets all script properties associated with this sheetSpider profile \n";
  codeString += "var preconfigStatus = ScriptProperties.getProperty('preconfigStatus');\n";
  codeString += "if (preconfigStatus!='true') {\n";
  codeString += propertyString; 
  codeString += "};\n";
  codeString += "ScriptProperties.setProperty('preconfigStatus','true');\n";
  codeString += "var ss = SpreadsheetApp.getActiveSpreadsheet();\n";
  codeString += "if (ss.getSheetByName('Forms in same folder')) { \n";
  codeString += "  sheetSpider_retrieveformurls(); \n";
  codeString += "} \n";
  codeString += "var parentFolder = DocsList.getFileById(ss.getId()).getParents()[0];\n";
  codeString += "var primaryFolder = parentFolder.createFolder('" + properties['primaryFolderName'] + "');\n";
  codeString += "ScriptProperties.setProperty('primaryFolderKey',primaryFolder.getId());\n";
  codeString += "ScriptProperties.setProperty('ssKey',ss.getId());\n";
  if (properties.mode = 'on Google Form submit') {
    codeString += "checkSetFormTrigger();\n";
  }
  codeString += "var ss = SpreadsheetApp.getActiveSpreadsheet();\n";
  var entitySheet = getSheetById(ss, properties.entitySheetId);
  var headers = entitySheet.getRange(1,1,1,entitySheet.getLastColumn()).getValues()[0];
  codeString += "var urlCol = " + (headers.indexOf('Spreadsheet URL') + 1) + ";\n";
  codeString += "var ssIdCol = " + (headers.indexOf('Spreadsheet ID') + 1) + ";\n";
  codeString += "var entitySheet = getSheetById(ss, " + properties.entitySheetId + ");\n";
  codeString += "entitySheet.getRange(2, urlCol, entitySheet.getLastRow(), 1).clear();\n";
  codeString += "entitySheet.getRange(2, ssIdCol, entitySheet.getLastRow(), 1).clear();\n";
  codeString += "ss.toast('Custom sheetSpider preconfiguration ran successfully.  A new folder named " + properties['primaryFolderName'] + " was created in your Drive to hold entity spreadsheets. Please check sheetSpider menu options to confirm system settings.');\n";
  codeString += "provisionSpreadsheetsUi();\n";
  window.setText(codeString).setWidth("100%").setHeight("350px");
  app.add(label);
  panel.add(window);
  app.add(panel);
  ss.show(app);
  return app;
}
