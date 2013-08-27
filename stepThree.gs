function disaggregateAndPushUi() {
  var app = UiApp.createApplication().setTitle('Step 3: Disaggregate and Push Data').setHeight(450);
  var properties = ScriptProperties.getProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var waitingImageUrl = 'https://drive.google.com/uc?export=download&id='+WAITINGICONID;
  var waitingImage = app.createImage(waitingImageUrl).setWidth('150px').setHeight('150px').setId('waitingImage').setVisible(false).setStyleAttribute('position', 'absolute').setStyleAttribute('left', '150px').setStyleAttribute('top', '100px');
  var panel = app.createVerticalPanel().setId('panel');
  var feederSheetId = properties.feederSheetId;
  if (feederSheetId) {
    feederSheetId = parseInt(feederSheetId);
  } else {
    Browser.msgBox('You have not selected a feeder sheet. Please return to Step 2 and correct this before proceding.');
    app.close();
    return app;
  }
  var feederSheet = getSheetById(ss, feederSheetId);
  ss.setActiveSheet(feederSheet);
  var lastRow = feederSheet.getLastRow();
  if (lastRow>0) {
    var headers = feederSheet.getRange(1, 1, 1, feederSheet.getLastColumn()).getValues()[0];
  } else {
    Browser.msgBox('Your feeder sheet contains no data. Please correct this before proceding.');
    app.close();
    return app;
  }
  
  var e = new Object();
  e.parameter = new Object();
  e.parameter.mode = properties.mode;
  
  //mode selector
  var modeLabel = app.createLabel("Mode you will be using");
  var modeSelector = app.createListBox().setName('mode');
  var modeSelectHandler = app.createServerHandler('refreshFormQuestions').addCallbackElement(panel);
  modeSelector.addChangeHandler(modeSelectHandler);
  modeSelector.addItem('on Google Form submit').addItem('manual push');
  if (properties.mode == 'manual push') {
    modeSelector.setSelectedIndex(1);
    e.parameter.mode = 'manual push';
  } else {
    e.parameter.mode = 'on Google Form submit';
  }
  panel.add(modeLabel).add(modeSelector);
  var entityCol = properties.entityCol;
  
  //form question selector
  var formQuestionLabel = app.createLabel("Form question that contains entity name(s)").setId('formQLabel').setStyleAttribute('marginTop', '10px').setVisible(false);
  var formQuestionSelector = app.createListBox().setName('formQId').setId('formQSelector').setVisible(false).setEnabled(false);
  var formQuestionHelp = app.createLabel("Only username, list, checkbox, and multiple choice questions can be used for this. Form items will populate automatically from entity sheet").setId('formQHelp').setVisible(false).setStyleAttribute('fontSize', '9px').setStyleAttribute('color', 'grey');
  panel.add(formQuestionLabel).add(formQuestionSelector).add(formQuestionHelp);
  var formQId = properties.formQId;
  if (formQId) {
    e.parameter.formQId;
  }
  
  //Overwrite mode listboxbox
  var overwriteModeLabel = app.createLabel("How do you want data pushed to entity spreadsheets?").setId('overwriteModeLabel').setStyleAttribute('marginTop', '10px').setVisible(false);
  var overwriteListBox = app.createListBox().setName('overwriteMode').setId('overwriteModeListBox').setVisible(false).setEnabled(false);
  overwriteListBox.addItem("Append only new unique records");
  overwriteListBox.addItem("Overwrite all existing data");
  panel.add(overwriteModeLabel).add(overwriteListBox);
  var overwriteMode = properties.overwriteMode;
  if ((properties.mode == "manual push") && (overwriteMode=="Overwrite all existing data")) {
    overwriteListBox.setSelectedIndex(1);
    e.parameter.overwriteListBox = "Overwrite all existing data";
  }
  
  //uniqueness criteria
  var uniquenessLabel = app.createLabel("Select uniqueness criterion. Selecting multiple fields will require combined uniqueness.").setStyleAttribute('marginTop', '10px');
  var scrollPanel = app.createScrollPanel().setHeight('150px');
  var checkBoxPanel = app.createVerticalPanel();
  var uniquenessCriterion = properties.uniquenessCriterion;
  if (uniquenessCriterion) {
    uniquenessCriterion = uniquenessCriterion.split("||");
  } else {
    uniquenessCriterion = [];
  }
  
  var handlers = [];
  for (var i=0; i<headers.length; i++) {
    handlers[i] = app.createServerHandler('refreshButton').addCallbackElement(panel);
    var thisCheckBox = app.createCheckBox(headers[i]).setName('header-'+i).addClickHandler(handlers[i]);
    checkBoxPanel.add(thisCheckBox);
    if (uniquenessCriterion.indexOf(headers[i])!=-1) {
      thisCheckBox.setValue(true);
      e.parameter['header-'+i] = 'true';
    }
    if ((uniquenessCriterion.length==0)&&(headers[i]=="Timestamp")) {
      thisCheckBox.setValue(true);
      e.parameter['header-'+i] = 'true';
    }
    if (headers[i]==properties.feederEntityCol) {
      thisCheckBox.setValue(true).setEnabled(false);
      e.parameter['header-'+i] = 'true';
    }
  }
  scrollPanel.add(checkBoxPanel);
  
  //help text
  var helpText = 'When pushing data to entity sheets, the script will only push new rows for which no existing match exists according to the criterion selected below.';
  helpText += 'When retrieving updates to records, the script will use the uniqueness criterion to match rows.  Entity name must always be a part of the uniqueness criterion.';
  var helpLabel = app.createHTML(helpText).setStyleAttribute('backgroundColor', 'whiteSmoke').setStyleAttribute('padding', '15px').setStyleAttribute('marginTop', '10px');
  panel.add(helpLabel);
  
  //button
  var saveHandler = app.createServerHandler('saveStepThree').addCallbackElement(panel);
  var button = app.createButton('Save', saveHandler).setId('button').setEnabled(false);
  var waitingHandler = app.createClientHandler().forTargets(button).setEnabled(false).forTargets(panel).setStyleAttribute('opacity', '0.5').forTargets(waitingImage).setVisible(true);
  button.addClickHandler(waitingHandler);
  
  
  var buttonHelp = app.createLabel('You must select at least one uniqueness criterion in addition to entity name').setId('buttonHelp').setVisible(false).setStyleAttribute('fontSize', '9px').setStyleAttribute('color','grey');
  refreshButton(e);
  refreshFormQuestions(e);
  panel.add(uniquenessLabel);
  panel.add(scrollPanel);
  panel.add(button);
  panel.add(buttonHelp);
  app.add(panel);
  app.add(waitingImage);
  ss.show(app);
  return app;
}


function refreshButton(e) {
  var app = UiApp.getActiveApplication();
  var button = app.getElementById('button');
  var buttonHelp = app.getElementById('buttonHelp');
  var mode = e.parameter.mode;
  var properties = ScriptProperties.getProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var feederSheetId = properties.feederSheetId;
  if (feederSheetId) {
    feederSheetId = parseInt(feederSheetId);
  } else {
    Browser.msgBox('You have not selected a feeder sheet. Please return to Step 2 and correct this before proceding.');
    app.close();
    return app;
  }
  var feederSheet = getSheetById(ss, feederSheetId);
  var lastRow = feederSheet.getLastRow();
  if (lastRow>0) {
    var headers = feederSheet.getRange(1, 1, 1, feederSheet.getLastColumn()).getValues()[0];
  } else {
    Browser.msgBox('Your feeder sheet contains no data. Please correct this before proceding.');
    app.close();
    return app;
  }
  
  var count = 0;
  var selected = [];
  for (var i=0; i<headers.length; i++) {
    var thisBox = e.parameter['header-'+i];
    if (thisBox == "true") {
      selected.push(headers[i])
      count++
    }
  }
  if (count>1) {
    button.setEnabled(true);
    buttonHelp.setVisible(false);
  } else {
    button.setEnabled(false);
    buttonHelp.setVisible(true);
  }
  if (mode == "on Google Form submit") {
    button.setHTML('Save and set form trigger');
  } else {
    button.setHTML('Save and run data push');
  }
  return app;
}


function refreshFormQuestions(e) {
  var app = UiApp.getActiveApplication();
  var properties = ScriptProperties.getProperties();
  var unique = entitiesAreUnique();
  if (!unique) {
    Browser.msgBox('FYI: Based on these settings, you have duplicate entity names. These are shown in pink on the entity sheet. Please correct this before proceding.');
    app.close();
    return app;
  }
  if (unique == "commas") {
    Browser.msgBox("Entity names may not contain commas.  Please fix this before proceding.");
    app.close();
    return app;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formQLabel = app.getElementById('formQLabel');
  var formQuestionSelector = app.getElementById('formQSelector');
  var formQuestionHelp = app.getElementById('formQHelp');
  var mode = e.parameter.mode;
  if (mode == "on Google Form submit") {
    formQLabel.setVisible(true);
    formQuestionSelector.setVisible(true).setEnabled(true);
    formQuestionHelp.setVisible(true);
    var formUrl = ss.getFormUrl();
    if (formUrl) {
      formQuestionSelector.clear();
      var form = FormApp.openByUrl(formUrl);
      var formQIds = [];
      var requiresLogin = form.requiresLogin();
      if ((requiresLogin)&&(properties.feederEntityCol == "Username")) {
        formQuestionSelector.addItem("Username", 'username');
        formQIds.push('Username', 'username');
      }
      var formItems = form.getItems();
      for (var k=0; k<formItems.length; k++) {
        if ((formItems[k].getType()=="CHECKBOX")||(formItems[k].getType()=="LIST")||(formItems[k].getType()=="MULTIPLE_CHOICE")) {
          var thisTitle = formItems[k].getTitle();
          var thisId = formItems[k].getId();
          formQuestionSelector.addItem(thisTitle, thisId);
          formQIds.push(thisTitle, thisId);
        }
      }
      var selectedQId = e.parameter.formQId;
      if (selectedQId) {
        var index = formQIds.indexOf(selectedQId);
        formQuestionSelector.setSelectedIndex(index);
      }
    } else {
      Browser.msgBox("You have no form attached to this spreadsheet.  Please correct this before proceding.");
    }
  } else {
    var overwriteModeLabel = app.getElementById('overwriteModeLabel').setVisible(true);
    var overwriteModeListBox = app.getElementById('overwriteModeListBox').setVisible(true).setEnabled(true);
    var overwriteMode = e.parameter.overwriteMode;
    if (overwriteMode == "Overwrite all existing data") {
      overwriteModeListBox.setSelectedIndex(1);
    }
    formQLabel.setVisible(false);
    formQuestionSelector.setVisible(false).setEnabled(false);
    formQuestionHelp.setVisible(false);
  }
  return app;
}


function saveStepThree(e) {
  var app = UiApp.getActiveApplication();
  var properties = ScriptProperties.getProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  properties.ssKey = ss.getId();
  var feederSheetId = properties.feederSheetId;
  if (feederSheetId) {
    feederSheetId = parseInt(feederSheetId);
  } else {
    Browser.msgBox('You have not selected a feeder sheet. Please return to Step 2 and correct this before proceding.');
    app.close();
    return app;
  }
  var feederSheet = getSheetById(ss, feederSheetId);
  var entitySheet = getSheetById(ss, properties.entitySheetId);
  var lastRow = feederSheet.getLastRow();
  if (lastRow>0) {
    var headers = feederSheet.getRange(1, 1, 1, feederSheet.getLastColumn()).getValues()[0];
  } else {
    Browser.msgBox('Your feeder sheet contains no data. Please correct this before proceding.');
    app.close();
    return app;
  }
  var uniquenessCriterion = [];
  for (var i=0; i<headers.length; i++) {
    var thisBox = e.parameter['header-'+i];
    if (thisBox == "true") {
      uniquenessCriterion.push(headers[i]);
    }   
  }
  uniquenessCriterion = uniquenessCriterion.join('||');
  var mode = e.parameter.mode;
  properties.mode = mode;
  properties.uniquenessCriterion = uniquenessCriterion;
  if (mode == "on Google Form submit") {
    var formQId = e.parameter.formQId;
    properties.formQId = formQId;
    ScriptProperties.setProperties(properties);
    buildFormQuestion();
    checkSetFormTrigger();
  } else {
    properties.formQId = '';
    properties.overwriteMode = e.parameter.overwriteMode;
    checkSetFormTrigger(true);
  }
  properties.stepThreeComplete = "true";
  ScriptProperties.setProperties(properties);
  if (mode == 'manual push') {
    pushData();
  }
  onOpen();
  app.close();
  return app;
}


function buildFormQuestion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var formQId = properties.formQId;
  var entitySheet = getSheetById(ss, properties.entitySheetId);
  var formUrl = ss.getFormUrl();
  if (!formUrl) {
    Browser.msgBox("No form was found attached to this spreadsheet...please fix this if you want to push data on Google Form submission.");
    return;
  }
  var form = FormApp.openByUrl(formUrl);
  if (formQId!='username') {
    var items = [form.getItemById(formQId)];
    for (var i=0; i<items.length; i++) {
      var values = getEntityValues(entitySheet, properties.entityCol);
      if (values.length == 0) {
        values[0] = "No values found in column \"" + questionRanges['header-'+qId] + "\"";
      }
      var type = items[i].getType().toString();
      if (type == "LIST") {
        items[i].asListItem().setChoiceValues(values);
      }
      if (type == "MULTIPLE_CHOICE") {
        items[i].asMultipleChoiceItem().setChoiceValues(values);
      }
      if (type == "CHECKBOX") {
        items[i].asCheckboxItem().setChoiceValues(values);
      }
    }
  }
}


function getEntityValues(sheet, header) {
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var entityData = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  var colIndex = headers.indexOf(header);
  var entityValues = [];
  for (var i=0; i<entityData.length; i++) {
    entityValues.push(entityData[i][colIndex]);
  }
  return entityValues;
}


function checkSetFormTrigger(remove) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i=0; i<triggers.length; i++) {
    if (triggers[i].getHandlerFunction()=="pushData") {
      if (remove) {
        ScriptApp.deleteTrigger(triggers[i]);
      } else {
        return;
      }
    }
  }
  ScriptApp.newTrigger("pushData").forSpreadsheet(ss).onFormSubmit().create();
  return;
}
