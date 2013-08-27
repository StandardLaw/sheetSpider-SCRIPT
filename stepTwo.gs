function provisionSpreadsheetsUi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle('Step 2: Provision / Update Entity Spreadsheets').setHeight(440);
  var properties = ScriptProperties.getProperties();
  if (properties.mode == 'on Google Form submit') {
    var formUrl = ss.getFormUrl();
    if (!formUrl) {
      Browser.msgBox("You have no form attached to this spreadsheet.  Please correct this before proceding or switch to manual push mode in step 1.");
      app.close;
      return app;
    }
  }
  var waitingImageUrl = 'https://drive.google.com/uc?export=download&id='+WAITINGICONID;
  var waitingImage = app.createImage(waitingImageUrl).setWidth('150px').setHeight('150px').setId('waitingImage').setVisible(false).setStyleAttribute('position', 'absolute').setStyleAttribute('left', '150px').setStyleAttribute('top', '100px');
  var panel = app.createVerticalPanel().setId('panel');
  var sheets = ss.getSheets();
  
  var ssTemplateLabel = app.createLabel('Key of Template Spreadsheet (Optional)');
  var ssTemplateKeyBox = app.createTextBox().setName('ssTemplateKey').setWidth("100%");
  var ssTemplateHelp = app.createLabel('Useful if you want to distribute a pre-formatted or script-enabled spreadsheet to all entities.').setStyleAttribute('color','grey').setStyleAttribute('fontSize','9px');
  var ssTemplateKey = properties.ssTemplateKey;
  if (ssTemplateKey) {
    ssTemplateKeyBox.setValue(ssTemplateKey);
  }
  panel.add(ssTemplateLabel).add(ssTemplateKeyBox).add(ssTemplateHelp);
  
  //Sheet to clone selector
  var feederSheetLabel = app.createLabel('Feeder sheet for entity spreadsheets').setStyleAttribute('marginTop', '10px');
  var feederSheetHelp = app.createLabel('This is the sheet where your aggregate or form data resides.  Entity spreadsheets will be given a new sheet with header structure identical the feeder sheet').setStyleAttribute('fontSize', '9px').setStyleAttribute('color','grey');
  var sheetIds = [];
  var sheetSelector = app.createListBox().setName('feederSheetId');
  for (var i=0; i<sheets.length; i++) {
    var thisSheetId = sheets[i].getSheetId();
    var thisSheetName = sheets[i].getName();
    if (thisSheetId!=parseInt(properties.entitySheetId)) {
      sheetSelector.addItem(sheets[i].getName(), thisSheetId);
      sheetIds.push(thisSheetId);
    }
  }
  if (sheetIds.length==0) {
    Browser.msgBox("You need another sheet in your spreadsheet before this step makes sense...");
    return;
  }
  var feederSheetId = properties.feederSheetId;
  var feederSheet = getSheetById(ss, feederSheetId);
  if ((feederSheetId)&&(feederSheet)) {
    feederSheetId = parseInt(feederSheetId);
  } else {
    if (properties.mode == "on Google Form submit") {
      for (var k=0; k<sheets.length; k++) {
        if (sheets[k].getName().indexOf("Form")!=-1) {
          feederSheetId = sheets[k].getSheetId();
          break;
        }
      }
      if (!feederSheetId) {
        feederSheetId = sheets[0].getSheetId();
      }
    } else {
      feederSheetId = sheets[0].getSheetId();
    }
  }
  var index = sheetIds.indexOf(feederSheetId);
  sheetSelector.setSelectedIndex(index);
  var feederSheetHandler = app.createServerHandler('refreshFeederSheet').addCallbackElement(panel);
  var feederSheetWaitingHandler = app.createClientHandler().forTargets(panel).setStyleAttribute('opacity','0.5').forTargets(waitingImage).setVisible(true);
  sheetSelector.addChangeHandler(feederSheetHandler).addChangeHandler(feederSheetWaitingHandler);
  
  
  panel.add(feederSheetLabel);
  panel.add(sheetSelector).add(feederSheetHelp);
  
  //Entity name column selectbox - feeder sheet
  var feederEntityColLabel = app.createLabel('Header of Column in Feeder Sheet Containing Unique Entity Names').setStyleAttribute('marginTop', '10px');
  var feederEntityColSelectBox = app.createListBox().setId('feederEntityColSelector').setName('feederEntityCol');
  var feederEntityColHelp = app.createLabel('This is the column in the FEEDER sheet that will be used to reference the ENTITY sheet so the script knows where to send the data to.').setStyleAttribute('color','grey').setStyleAttribute('fontSize','9px');
  panel.add(feederEntityColLabel).add(feederEntityColSelectBox).add(feederEntityColHelp);
  
  //button
  var saveHandler = app.createServerHandler('saveStepTwo').addCallbackElement(panel);
  var button = app.createButton("Save settings and provision / update spreadsheets", saveHandler).setStyleAttribute('marginTop', '15px');;
  var waitingHandler = app.createClientHandler().forTargets(button).setEnabled(false).forTargets(panel).setStyleAttribute('opacity', '0.5').forTargets(waitingImage).setVisible(true);
  button.addClickHandler(waitingHandler);
  panel.add(button);
  
  //helptext
  var helpText1 = "It is best to finalize your feeder sheet structure PRIOR to running this step, as (much like editing a Google Form) it will be harder to update all your entity spreadsheets and preserve column order after the fact."
  var helpLabel1 = app.createLabel(helpText1).setStyleAttribute('marginTop', '15px');
  
  var helpText2 = "What happens when this step is run?";
  helpText2 += "<ul><li>Entities without spreadsheets will have spreadsheets created for them, shared with specified collaborators, get a new sheet NAMED identically to the feeder sheet selected here, and with identical column header structure, and spreadsheet keys and links logged in the entity sheet.</li>";
  helpText2 += "<li>Omitting a spreadsheet template will simply result in a blank starter spreadsheet.</li>";
  helpText2 += "<li>Any existing spreadsheets without an existing sheet NAMED IDENTICALLY to the feeder sheet will have one added, along with correct column headers.</li>";
  helpText2 += "<li>Any column headers in the feeder sheet that are missing from the destination spreadsheet will be appended after the last column in the destination sheet.</li></ul>";
  var helpLabel2 = app.createHTML(helpText2).setStyleAttribute('marginTop', '10px').setStyleAttribute('fontSize', '10px');
  
  panel.add(helpLabel1).add(helpLabel2);
  
  var e = new Object();
  e.parameter = new Object();
  e.parameter.feederSheetId = feederSheetId;
  refreshFeederSheet(e);
  app.add(panel);
  app.add(waitingImage);
  ss.show(app);
  return app;
  
}



function refreshFeederSheet(e) {
  var app = UiApp.getActiveApplication();
  var waitingImage = app.getElementById('waitingImage').setVisible(false);
  var panel = app.getElementById('panel').setStyleAttribute('opacity','1');
  var properties = ScriptProperties.getProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var feederSheetId = e.parameter.feederSheetId;
  var sheet = getSheetById(ss, feederSheetId);
  var lastRow = sheet.getLastRow();
  
  //fetch selected sheet headers if they exist.  If not, invent them
  if (lastRow>0) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  } else {
    var headers = ['Dummy Header 1','Dummy Header 2', 'Dummy Header 3'];
    sheet.getRange(1, 1, 1, 3).setValues([headers]);
  }
  
  //reset entity col selector
  var feederEntityColSelector = app.getElementById('feederEntityColSelector');
  feederEntityColSelector.clear();
  var selectorChoices = [];
  for (var i=0; i<headers.length; i++) {
    if (headers[i]!="Timestamp") {
      feederEntityColSelector.addItem(headers[i]);
      selectorChoices.push(headers[i]);
    }
  }
  //set if preset value exists
  var feederEntityCol = properties.feederEntityCol;
  if (feederEntityCol) {
    var index = selectorChoices.indexOf(feederEntityCol);
    feederEntityColSelector.setSelectedIndex(index);
  } 
  return app;
}


function saveStepTwo(e) {
  var app = UiApp.getActiveApplication();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var feederEntityCol = e.parameter.feederEntityCol;
  var feederSheetId = e.parameter.feederSheetId;
  var ssTemplateKey = e.parameter.ssTemplateKey;
  try {
    var feederSheet = getSheetById(ss, feederSheetId);
    var feederSheetName = feederSheet.getName();
  } catch(err) {
    Browser.msgBox('Unable to retrieve feeder sheet. Try this step again.');
    app.close();
    return app;
  }
  if ((feederEntityCol)&&(feederSheetId)&&(feederSheetName)) {
    properties.feederEntityCol = feederEntityCol;
    properties.feederSheetId = feederSheetId;
    properties.feederSheetName = feederSheetName;
  } else {
    Browser.msgBox('It appears you forgot to enter a value for one of the required fields. Please retry step 2 before continuing.');
    app.close();
    return app;
  }
  if (ssTemplateKey) {
    properties.ssTemplateKey = ssTemplateKey;
  } else {
    properties.ssTemplateKey = '';
  }
  ScriptProperties.setProperties(properties);
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
  var provisioned = provisionSpreadsheets();
  if ((properties.stepTwoComplete != "true")&&(provisioned==true)) {
    disaggregateAndPushUi();
  }
  if (provisioned==true) {
    properties.stepTwoComplete = true;
  }
  ScriptProperties.setProperties(properties);
  onOpen();
  app.close();
  return app;
}



function provisionSpreadsheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var thisSSTimezone = ss.getSpreadsheetTimeZone();
  var properties = ScriptProperties.getProperties();
  try {
    var entitySheet = getSheetById(ss, properties.entitySheetId);
    var feederSheet = getSheetById(ss, properties.feederSheetId);
  } catch(err) {
    Browser.msgBox("One of the sheets referenced by this script is missing.  Please revisit steps 1 and 2.");
    return;
  }
  try {
    var entitySheetHeaders = entitySheet.getRange(1,1,1,entitySheet.getLastColumn()).getValues()[0];
    var feederSheetHeaders = feederSheet.getRange(1,1,1,feederSheet.getLastColumn()).getValues()[0];
  } catch(err) {
    Browser.msgBox("One of the sheets referenced by this script in this spreadsheet is empty. Please fix and retry this step.");
    return;
  }
  var indices = fetchEntityIndices(entitySheet, entitySheetHeaders, properties);
  var primaryFolderKey = properties.primaryFolderKey;
  var ssTemplateKey = properties.ssTemplateKey;
  var feederSheetName = feederSheet.getName();
  if (feederSheetName!=properties.feederSheetName) {
    var resp = Browser.msgBox('The script has detected a change of feeder sheet name. Are you sure you want to generate new sheets in all entity spreadsheets?', Browser.Buttons.OK_CANCEL);
    if (resp=="cancel") {
      return;
    }
    ScriptProperties.setProperty('feederSheetName', feederSheetName);
  }
  var lastRow = entitySheet.getLastRow();
  if (lastRow>1) {
    var entityArray = entitySheet.getRange(2, 1, lastRow-1, entitySheet.getLastColumn()).getValues();
    for (var i=0; i<entityArray.length; i++) {
      var thisSSId = entityArray[i][indices.ssIdCol];
      var thisSSUrl = entityArray[i][indices.urlCol];
      var thisEntity = entityArray[i][indices.entityCol];
      var thisSSName = entityArray[i][indices.ssNameCol];
      if (thisSSName == "") {
        Browser.msgBox(thisEntity + " is missing a name. Please fix and run this step again.");
        return false;
      }
      var thisSSEditors = entityArray[i][indices.ssEditorsCol];
      var thisSSViewers = entityArray[i][indices.ssViewersCol];
      var thisSecondaryFolderKey = entityArray[i][indices.secondaryFolderCol];
      if (thisSSId=='') {
        if ((ssTemplateKey)&&(ssTemplateKey!='')) {
          var templateDoc = DocsList.getFileById(ssTemplateKey);
          var copyDoc = templateDoc.makeCopy(thisSSName);
          var copySSId = copyDoc.getId();
          var copySS = SpreadsheetApp.openById(copySSId);
        } else {
          var copySS = SpreadsheetApp.create(thisSSName); 
          var copySSId = copySS.getId();
          var copyDoc = DocsList.getFileById(copySSId);
        }
        call(function(){copySS.setSpreadsheetTimeZone(thisSSTimezone);});
        entitySheet.getRange(i+2, indices.ssIdCol+1).setValue(copySSId);
        entitySheet.getRange(i+2, indices.urlCol+1).setValue(copySS.getUrl());
        var driveRoot = DocsList.getRootFolder();
        copyDoc.removeFromFolder(driveRoot);
      } else {
        var copyDoc = DocsList.getFileById(thisSSId);
        var copySS = SpreadsheetApp.openById(thisSSId);
      }
      checkFixSheetHeaders(copySS, feederSheetName, feederSheetHeaders);
      var approvedParentIds = new Array(); 
      approvedParentIds.push(primaryFolderKey);
      if ((thisSecondaryFolderKey)&&(thisSecondaryFolderKey!="")) {
        thisSecondaryFolderKey = thisSecondaryFolderKey.replace(/\s/g, "").split(",");
        for (var s=0; s<thisSecondaryFolderKey.length; s++) {
          approvedParentIds.push(thisSecondaryFolderKey[s])
        }
      }
      checkFixFileParents(copyDoc, approvedParentIds);
      var approvedEditors = new Array();
      if ((thisSSEditors)&&(thisSSEditors!="")) {
        thisSSEditors = thisSSEditors.replace(/\s/g, "").split(",");
        for (var s=0; s<thisSSEditors.length; s++) {
          approvedEditors.push(thisSSEditors[s])
        }
      }
      var approvedViewers = new Array();
      if ((thisSSViewers)&&(thisSSViewers!="")) {
        thisSSViewers = thisSSViewers.replace(/\s/g, "").split(",");
        for (var s=0; s<thisSSViewers.length; s++) {
          approvedViewers.push(thisSSViewers[s])
        }
      }
      try {
        checkFixFileACLs(copyDoc, approvedViewers, approvedEditors);
      } catch (err) {
      }
      sheetSpider_logEntitySheetProvisioned();
    }
  } else {
    Browser.msgBox("It appears you have no entities listed in the sheet: " + entitySheet.getName() + " Please re-run step 2 once you have added entities.");
    return false;
  }
  return true;
}



function fetchEntityIndices(sheet, headerArray, properties) {
  var indices = new Object();
  if (headerArray.indexOf(properties.entityCol)!=-1) {
    indices.entityCol = headerArray.indexOf(properties.entityCol);
  } else {
    Browser.msgBox('We have detected a change in one of your entity sheet column headers.  Please revisit Step 1');
  }
  if (headerArray.indexOf(properties.ssNameCol)!=-1) {
    indices.ssNameCol = headerArray.indexOf(properties.ssNameCol);
  } else {
    Browser.msgBox('We have detected a change in one of your entity sheet column headers.  Please revisit Step 1');
  }
  if (headerArray.indexOf(properties.ssEditorsCol)!=-1) {
    indices.ssEditorsCol = headerArray.indexOf(properties.ssEditorsCol);
  } else {
    Browser.msgBox('We have detected a change in one of your entity sheet column headers.  Please revisit Step 1');
  }
  if (headerArray.indexOf(properties.ssViewersCol)!=-1) {
    indices.ssViewersCol = headerArray.indexOf(properties.ssViewersCol);
  } else {
    Browser.msgBox('We have detected a change in one of your entity sheet column headers.  Please revisit Step 1');
  }
  if (headerArray.indexOf(properties.secondaryFolderCol)!=-1) { 
    indices.secondaryFolderCol = headerArray.indexOf(properties.secondaryFolderCol);
  } else {
    Browser.msgBox('We have detected a change in one of your entity sheet column headers.  Please revisit Step 1');
  }
  if (headerArray.indexOf("Spreadsheet URL")!=-1) {
    indices.urlCol = headerArray.indexOf("Spreadsheet URL");
  } else {
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn()+1).setValue("Spreadsheet URL").setBackground("black").setFontColor("white").setComment("Don't change this header");
    headerArray.push("Spreadsheet URL");
    indices.urlCol = headerArray.indexOf("Spreadsheet URL");
  }
  if (headerArray.indexOf("Spreadsheet ID")!=-1) {
    indices.ssIdCol = headerArray.indexOf("Spreadsheet ID");
  } else {
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn()+1).setValue("Spreadsheet ID").setBackground("black").setFontColor("white").setComment("Don't change this header");
    headerArray.push("Spreadsheet ID");
    indices.ssIdCol = headerArray.indexOf("Spreadsheet ID");
  }
  return indices;
}
