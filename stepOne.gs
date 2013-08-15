var scriptTitle = "sheetSpider Script V1.0.2 (7/22/13)";
var scriptName = "sheetSpider"
var scriptTrackingId = "UA-41943014-1"
// Written by Andrew Stillman for New Visions for Public Schools
// Published under GNU General Public License, version 3 (GPL-3.0)
// See restrictions at http://www.opensource.org/licenses/gpl-3.0.html

var WAITINGICONID = '0B2vrNcqyzernZTI4WGJ4dEcyaDA';

function onInstall() {
  onOpen();
}

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var menuItems = [];
  menuItems.push({name: 'What is sheetSpider?', functionName: 'sheetSpider_whatIs'});
  if (properties.preconfigStatus=="true") {
    menuItems.push({name: 'Step 1: Set Up Entity Sheet', functionName: 'entitySheetSettingsUi'});
    if (properties.stepOneComplete == "true") {
      menuItems.push({name: 'Step 2: Provision Entity Spreadsheets', functionName: 'provisionSpreadsheetsUi'});
      if (properties.stepTwoComplete == "true") {
        menuItems.push({name: 'Step 3: Disaggregate and Push Data', functionName: 'disaggregateAndPushUi'});
        if (properties.stepThreeComplete == "true") {
          menuItems.push({name: 'Step 4: Retrieve Live Data', functionName: 'retrieveLiveData'});
          if (properties.stepFourComplete == "true") {
            menuItems.push({name: 'Update Feeder Sheet With Returned Data', functionName: 'updateFeederSheet'});
          }
        }
        if ((properties.stepThreeComplete == "true") && (properties.mode == "on Google Form submit")) {
          menuItems.push({name: 'Refresh Google Form question with entity names', functionName: 'buildFormQuestion'});
        }
        menuItems.push(null);
        menuItems.push({name:'Export settings', functionName:'sheetSpider_extractorWindow'});
      }
    }
  } else {
    menuItems.push({name:'Run initial configuration',functionName:'sheetSpider_preconfig'});
  }
  ss.addMenu('sheetSpider', menuItems);
}


function entitySheetSettingsUi() {  
  var app = UiApp.createApplication().setTitle('Step 1: Set Up Entity Sheet').setHeight(400);
  var waitingImageUrl = 'https://drive.google.com/uc?export=download&id='+WAITINGICONID;
  var waitingImage = app.createImage(waitingImageUrl).setWidth('150px').setHeight('150px').setId('waitingImage').setVisible(false).setStyleAttribute('position', 'absolute').setStyleAttribute('left', '150px').setStyleAttribute('top', '100px');
  var properties = ScriptProperties.getProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outerScrollPanel = app.createScrollPanel().setHeight('400px');
  var panel = app.createVerticalPanel().setId('panel');
  var sheets = ss.getSheets();
  
  
  //help text
  var helpText = 'Entity names must be UNIQUE.  Each entity will have its own separate spreadsheet provisioned in Step 2. After Step 3, you will be able to push data (manually, or on form submit) based on matching entity name.';
  helpText + '<br/>Examples of possible entities: students, teachers, schools, etc.';
  var helpLabel = app.createHTML(helpText).setStyleAttribute('backgroundColor', 'whiteSmoke').setStyleAttribute('padding', '15px');
  panel.add(helpLabel);
  
  //mode intended
  var modeLabel = app.createLabel("Select how you want to distribute data to entity spreadsheets");
  var modeSelector = app.createListBox().setName('mode');
  modeSelector.addItem('on Google Form submit').addItem('manual push');
  if (properties.mode == 'manual push') {
    modeSelector.setSelectedIndex(1);
  } 
  panel.add(modeLabel).add(modeSelector);
  
  //sheet selectbox
  var sheetSelectLabel = app.createLabel('Select Sheet Containing Unique Entities').setStyleAttribute('marginTop', '10px');
  var sheetSelector = app.createListBox().setName('entitySheetId');
  var sheetIds = [];
  for (var i=0; i<sheets.length; i++) {
    var thisSheetId = sheets[i].getSheetId();
    sheetSelector.addItem(sheets[i].getName(), thisSheetId);
    sheetIds.push(thisSheetId);
  }
  var entitySheetId = properties.entitySheetId;
  var sheet = getSheetById(ss, entitySheetId);
  if ((entitySheetId)&&(sheet)) {
    entitySheetId = parseInt(entitySheetId);
    var index = sheetIds.indexOf(entitySheetId);
    sheetSelector.setSelectedIndex(index);
  } else {
    var newEntitySheet = ss.getSheetByName('SpiderSheetEntities');
    if (!newEntitySheet) {
      newEntitySheet = ss.insertSheet('SpiderSheetEntities');
      sheetSelector.addItem(newEntitySheet.getName(), newEntitySheet.getSheetId());
      sheetIds.push(newEntitySheet.getSheetId());
    }
    entitySheetId = newEntitySheet.getSheetId();
    sheetSelector.setSelectedIndex(sheetIds.indexOf(entitySheetId));
    ScriptProperties.setProperty('entitySheetId', entitySheetId);
  }
  var entitySheetHandler = app.createServerHandler('refreshEntitySheet').addCallbackElement(panel);
  var entitySheetWaitingHandler = app.createClientHandler().forTargets(panel).setStyleAttribute('opacity','0.5').forTargets(waitingImage).setVisible(true);
  sheetSelector.addChangeHandler(entitySheetHandler).addChangeHandler(entitySheetWaitingHandler);
  panel.add(sheetSelectLabel).add(sheetSelector);
  
  //Entity name column selectbox
  var entityColLabel = app.createLabel('Header of Column Containing Unique Entity Names').setStyleAttribute('marginTop', '10px');
  var entityColSelectBox = app.createListBox().setId('entityColSelector').setName('entityCol');
  panel.add(entityColLabel).add(entityColSelectBox);
  
  //Spreadsheet name column selectbox
  var ssNameColLabel = app.createLabel('Header of Column Containing Spreadsheet Names').setStyleAttribute('marginTop', '10px');;
  var ssNameColSelectBox = app.createListBox().setId('ssNameColSelector').setName('ssNameCol');
  panel.add(ssNameColLabel).add(ssNameColSelectBox);
  
  //Spreadsheet editor(s) column selectbox
  var ssEditorsColLabel = app.createLabel('Header of Column Containing Spreadsheet Editors (comma separated)').setStyleAttribute('marginTop', '10px');;
  var ssEditorsColSelectBox = app.createListBox().setId('ssEditorsColSelector').setName('ssEditorsCol');
  panel.add(ssEditorsColLabel).add(ssEditorsColSelectBox);
  
  //Spreadsheet viewer(s) column selectbox
  var ssViewersColLabel = app.createLabel('Header of Column Containing Spreadsheet Viewers (comma separated)').setStyleAttribute('marginTop', '10px');;
  var ssViewersColSelectBox = app.createListBox().setId('ssViewersColSelector').setName('ssViewersCol');
  panel.add(ssViewersColLabel).add(ssViewersColSelectBox);
  
  //primary folder key textbox
  var primaryFolderLabel = app.createLabel('Key of Primary Folder to Contain Spreadsheets').setStyleAttribute('marginTop', '10px');;
  var primaryFolderKeyBox = app.createTextBox().setName('primaryFolderKey').setWidth("100%");
  var primaryFolderKey = properties.primaryFolderKey;
  if (primaryFolderKey) {
    primaryFolderKeyBox.setValue(primaryFolderKey);
  } else {
    var thisSSId = ss.getId();
    var parentFolder = DocsList.getFileById(thisSSId).getParents()[0];
    var parentFolderId = parentFolder.getId();
    primaryFolderKeyBox.setValue(parentFolderId);
  }
  panel.add(primaryFolderLabel).add(primaryFolderKeyBox);
  
  //secondary folder key column listboxbox
  var secondaryFolderColLabel = app.createLabel('Header of Column Containing Key of Secondary Folder for Spreadsheets').setStyleAttribute('marginTop', '10px');;
  var secondaryFolderColSelector = app.createListBox().setId('secondaryFolderColSelector').setName('secondaryFolderCol');
  panel.add(secondaryFolderColLabel).add(secondaryFolderColSelector);
  
  //button
  var saveHandler = app.createServerHandler('saveStepOne').addCallbackElement(panel);
  var button = app.createButton("Save settings", saveHandler).setStyleAttribute('marginTop', '15px');;
  var waitingHandler = app.createClientHandler().forTargets(button).setEnabled(false).forTargets(panel).setStyleAttribute('opacity', '0.5').forTargets(waitingImage).setVisible(true);
  button.addClickHandler(waitingHandler);
  panel.add(button);
  
  var e = new Object();
  e.parameter = new Object();
  e.parameter.entitySheetId = entitySheetId;
  refreshEntitySheet(e);
  outerScrollPanel.add(panel);
  app.add(outerScrollPanel);
  app.add(waitingImage);
  ss.show(app);
  return app;
}


function refreshEntitySheet(e) {
  var app = UiApp.getActiveApplication();
  var waitingImage = app.getElementById('waitingImage').setVisible(false);
  var panel = app.getElementById('panel').setStyleAttribute('opacity','1');
  var properties = ScriptProperties.getProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var entitySheetId = e.parameter.entitySheetId;
  var sheet = getSheetById(ss, entitySheetId);
  var lastRow = sheet.getLastRow();
  //fetch selected sheet headers if they exist.  If not, invent them
  if (lastRow>0) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  } else {
    sheet.setFrozenRows(1);
    var headers = ['Entity Name','Spreadsheet Title', 'Spreadsheet Editor(s)', 'Spreadsheet Viewer(s)', 'Secondary Folder Key'];
    sheet.getRange(1, 1, 1, 5).setValues([headers]);
    sheet.getRange(1, 1).setNote('May not contain commas');
    sheet.getRange(1, 3).setNote('OPTIONAL: Separate multiple with commas');
    sheet.getRange(1, 4).setNote('OPTIONAL: Separate multiple with commas');
    sheet.getRange(1, 5).setNote('OPTIONAL: Folder key is the long unique string in the last argument of a folder\'s URL in Drive');
    var freshSheet = true;
  }
  
  //reset secondary folder col selector
  var secondaryFolderColSelector = app.getElementById('secondaryFolderColSelector');
  secondaryFolderColSelector.clear();
  for (var i=0; i<headers.length; i++) {
    secondaryFolderColSelector.addItem(headers[i]);
  }
  
  //set if preset value exists
  var secondaryFolderCol = properties.secondaryFolderCol;
  if ((secondaryFolderCol)&&(!freshSheet)) {
    var index = headers.indexOf(secondaryFolderCol);
    secondaryFolderColSelector.setSelectedIndex(index);
  } else if (freshSheet) {
    secondaryFolderColSelector.setSelectedIndex(4);
  }
  
  //reset entity col selector
  var entityColSelector = app.getElementById('entityColSelector');
  entityColSelector.clear();
  for (var i=0; i<headers.length; i++) {
    entityColSelector.addItem(headers[i]);
  }
  //set if preset value exists
  var entityCol = properties.entityCol;
  if ((entityCol)&&(!freshSheet)) {
    var index = headers.indexOf(entityCol);
    entityColSelector.setSelectedIndex(index);
  } else if (freshSheet) {
    entityColSelector.setSelectedIndex(0);
  }
  
  //reset ss name col selector
  var ssNameColSelector = app.getElementById('ssNameColSelector');
  ssNameColSelector.clear();
  for (var i=0; i<headers.length; i++) {
    ssNameColSelector.addItem(headers[i]);
  }
  //set if preset value exists
  var ssNameCol = properties.ssNameCol;
  if ((ssNameCol)&&(!freshSheet)) {
    var index = headers.indexOf(ssNameCol);
    ssNameColSelector.setSelectedIndex(index);
  } else if (freshSheet) {
    ssNameColSelector.setSelectedIndex(1);
  }
  
  //reset ss editor col selector
  var ssEditorsColSelector = app.getElementById('ssEditorsColSelector');
  ssEditorsColSelector.clear();
  for (var i=0; i<headers.length; i++) {
    ssEditorsColSelector.addItem(headers[i]);
  }
  
  //set if preset value exists
  var ssEditorsCol = properties.ssEditorsCol;
  if ((ssEditorsCol)&&(!freshSheet)) {
    var index = headers.indexOf(ssEditorsCol);
    ssEditorsColSelector.setSelectedIndex(index);
  } else if (freshSheet) {
    ssEditorsColSelector.setSelectedIndex(2);
  }
  
  //reset ss viewer col selector
  var ssViewersColSelector = app.getElementById('ssViewersColSelector');
  ssViewersColSelector.clear();
  for (var i=0; i<headers.length; i++) {
    ssViewersColSelector.addItem(headers[i]);
  }
  
  //set if preset value exists
  var ssViewersCol = properties.ssViewersCol;
  if ((ssViewersCol)&&(!freshSheet)) {
    var index = headers.indexOf(ssViewersCol);
    ssViewersColSelector.setSelectedIndex(index);
  } else if (freshSheet) {
    ssViewersColSelector.setSelectedIndex(3);
  }
  return app;
}


function saveStepOne(e) {
  var app = UiApp.getActiveApplication();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var entityCol = e.parameter.entityCol;
  var ssNameCol = e.parameter.ssNameCol;
  var ssEditorsCol = e.parameter.ssEditorsCol;
  var ssViewersCol = e.parameter.ssViewersCol;
  var entitySheetId = e.parameter.entitySheetId;
  var primaryFolderKey = e.parameter.primaryFolderKey;
  var secondaryFolderCol = e.parameter.secondaryFolderCol;
  var mode = e.parameter.mode;
  if ((entitySheetId)&&(entityCol)&&(primaryFolderKey!='')&&(secondaryFolderCol)) {
    properties.entityCol = entityCol;
    properties.ssEditorsCol = ssEditorsCol;
    properties.ssViewersCol = ssViewersCol;
    properties.primaryFolderKey = primaryFolderKey;
    properties.primaryFolderName = DocsList.getFolderById(primaryFolderKey).getName();
    properties.secondaryFolderCol = secondaryFolderCol;
    properties.ssNameCol = ssNameCol;
    properties.entitySheetId = entitySheetId;
    properties.mode = mode;
    ScriptProperties.setProperties(properties);
  } else {
    Browser.msgBox('It appears you forgot to enter a value for one of the required fields');
    app.close();
    return app;
  }
  var unique = entitiesAreUnique();
  if (!unique) {
    Browser.msgBox('FYI: Based on these settings, you have duplicate entity names. These are shown in pink on the entity sheet. Please correct this and re-save this step before proceding.');
    app.close();
    return app;
  }
  if (unique == "commas") {
    Browser.msgBox("Entity names may not contain commas.  Please fix and re-save this step before proceding.");
    app.close();
    return app;
  }
  if (properties.stepOneComplete != "true") {
    var firstRun = true;
  }
  properties.stepOneComplete = "true";
  ScriptProperties.setProperties(properties);
  onOpen();
  if (mode == "on Google Form submit") {
    var formUrl = ss.getFormUrl();
    if (!formUrl) {
      Browser.msgBox("You have no form attached to this spreadsheet.  Please correct this before proceding.");
      app.close;
      return app;
    }
  }
  if (firstRun) {
    provisionSpreadsheetsUi();
  }
  app.close();
  return app;
}
