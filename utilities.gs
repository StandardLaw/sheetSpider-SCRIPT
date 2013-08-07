function getSheetById(ss, sheetId) {
  var sheets = ss.getSheets();
  for (var i=0; i<sheets.length; i++) {
    var thisSheetId = sheets[i].getSheetId();
    if (thisSheetId==sheetId) {
      return sheets[i];
    }
  }
  return;
}


function checkFixFileParents(file, approvedParentIds) {
  var parents = file.getParents();
  var parentKeys = [];
  try {
    for (var j=0; j<parents.length; j++) {
      if (approvedParentIds.indexOf(parents[j].getId())==-1) {                
        file.removeFromFolder(parents[j]);
      } else {
        parentKeys.push(parents[j].getId())
      }
    }
    for (var i=0; i<approvedParentIds.length; i++) {
      if (parentKeys.indexOf(approvedParentIds[i])==-1) {
        var parentFolder = DocsList.getFolderById(approvedParentIds[i]);
        file.addToFolder(parentFolder);
      }
    }
    return true;
  } catch(err) {
    var test = err.message;
    return false;
  }
}


function checkFixFileACLs(file, approvedViewers, approvedEditors) {
  var viewers = file.getViewers().join(",");
  viewers = viewers.split(",");
  var editors = file.getEditors().join(",");
  editors = editors.split(",");
  viewers = arr_diff(editors, viewers);
  var owner = file.getOwner().toString();
  var fileKey = file.getId();
  var driveFile = DriveApp.getFileById(fileKey);
  var currViewers = [];
  for (var k=0; k<viewers.length; k++) {
    if ((viewers[k]!='')&&(approvedViewers.indexOf(viewers[k].toLowerCase())==-1)&&(approvedEditors.indexOf(viewers[k].toLowerCase())==-1)&&(viewers[k]!=owner)) {
      try {
        call(function() {driveFile.removeViewer(viewers[k].toLowerCase());});
      } catch(err) {
        Logger.log(err.message);
      }
    } else {
      currViewers.push(viewers[k].toLowerCase());
    }
  }
  for (var k=0; k<approvedViewers.length; k++) {
    if ((approvedViewers[k]!='')&&(approvedViewers[k])) {
      if (currViewers.indexOf(approvedViewers[k])==-1) {
        try {
          call(function() {driveFile.addViewer(approvedViewers[k].toLowerCase());});
        } catch(err) {
          Logger.log(err.message);
        }
      }
    }
  }
  var currEditors = [];
  for (var k=0; k<editors.length; k++) {
    if ((editors[k]!='')&&(approvedEditors.indexOf(editors[k].toLowerCase())==-1)&&(editors[k]!=owner)) {
      try {
        call(function() {driveFile.removeEditor(editors[k].toLowerCase());});
      } catch(err) {
        Logger.log(err.message);
      }
    } else {
      currEditors.push(editors[k].toLowerCase().replace(/\s+/g, ''));
    }
  }
  for (var k=0; k<approvedEditors.length; k++) {
    if ((approvedEditors[k]!='')&&(approvedEditors[k])) {
      if (currEditors.indexOf(approvedEditors[k].toLowerCase())==-1) {
        try {
          call(function() {driveFile.addEditor(approvedEditors[k].toLowerCase().replace(/\s+/g, ''));});
        } catch(err) {
          Logger.log(err.message);
        }
      }
    }
  }
  return;
}


function arr_diff(a1, a2)
{
  var a=[], diff=[];
  for(var i=0;i<a1.length;i++)
    a[a1[i]]=true;
  for(var i=0;i<a2.length;i++)
    if(a[a2[i]]) delete a[a2[i]];
  else a[a2[i]]=true;
  for(var k in a)
    diff.push(k);
  return diff;
}


function entitiesAreUnique() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var entitySheetId = properties.entitySheetId;
  var entityCol = properties.entityCol;
  if ((entitySheetId)&&(entityCol)) {
    var entitySheet = getSheetById(ss, entitySheetId);
    var lastCol = entitySheet.getLastColumn();
    if (lastCol>0) {
      var headers = entitySheet.getRange(1, 1, 1, lastCol).getValues()[0];
    } else {
      return true;
    }
    var entityColNum = headers.indexOf(entityCol) + 1;
    var lastRow = entitySheet.getLastRow();
    if (lastRow>1) {
      var values = entitySheet.getRange(2, entityColNum, entitySheet.getLastRow()-1, 1).getValues();
      for (var i=0; i<values.length; i++) {
        var thisElement = values[i][0];
        var commaTest = thisElement.split(",");
        if (commaTest.length>1) {
          return "commas";
        }
        var thisCount = 0;
        var theseIndices = [];
        for (var j=0; j<values.length; j++) {
          if (values[j][0] == thisElement) {
            theseIndices.push(j);
            thisCount++;
          } 
        }
        if (thisCount > 1) {
          for (var k=0; k<theseIndices.length; k++) {
            entitySheet.getRange(theseIndices[k]+2, entityColNum).setBackground("pink"); 
          }
          return false;
        }
      }
    }
  }
  return true;
}


function checkFixSheetHeaders(copySS, feederSheetName, feederSheetHeaders) {
  var sheet = copySS.getSheetByName(feederSheetName);
  if (!sheet) {
    sheet = copySS.insertSheet(feederSheetName, 0);
    sheet.getRange(1, 1, 1, feederSheetHeaders.length).setValues([feederSheetHeaders]).setBackground('#C0C0C0').setFontWeight('bold');
    sheet.getRange(1, 1).setNote('Please do not alter this sheet\'s name or headers as they are required by a script run by ' + Session.getActiveUser().getEmail());
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();
  } else {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastCol==0) {
      lastCol = 1;
    }
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (var i=0; i<feederSheetHeaders.length; i++) {
      if (headers.indexOf(feederSheetHeaders[i])==-1) {
        if (i>0) {
          sheet.insertColumnAfter(i);
        } else {
          sheet.insertColumnBefore(1);
        }
        SpreadsheetApp.flush();
        sheet.getRange(1, i+1).setValue(feederSheetHeaders[i]).setBackground('#C0C0C0').setFontWeight('bold').setNote('Please do not change header value. Column inserted via script at ' + new Date() + ' by ' + Session.getActiveUser().getEmail());
        SpreadsheetApp.flush();
      }
    }
  }
  return sheet;
}


/**
* Invokes a function, performing up to 5 retries with exponential backoff.
* Retries with delays of approximately 1, 2, 4, 8 then 16 seconds for a total of 
* about 32 seconds before it gives up and rethrows the last error. 
* See: https://developers.google.com/google-apps/documents-list/#implementing_exponential_backoff 
* <br>Author: peter.herrmann@gmail.com (Peter Herrmann)
<h3>Examples:</h3>
<pre>//Calls an anonymous function that concatenates a greeting with the current Apps user's email
var example1 = GASRetry.call(function(){return "Hello, " + Session.getActiveUser().getEmail();});
</pre><pre>//Calls an existing function
var example2 = GASRetry.call(myFunction);
</pre><pre>//Calls an anonymous function that calls an existing function with an argument
var example3 = GASRetry.call(function(){myFunction("something")});
</pre><pre>//Calls an anonymous function that invokes DocsList.setTrashed on myFile and logs retries with the Logger.log function.
var example4 = GASRetry.call(function(){myFile.setTrashed(true)}, Logger.log);
</pre>
*
* @param {Function} func The anonymous or named function to call.
* @param {Function} optLoggerFunction Optionally, you can pass a function that will be used to log 
to in the case of a retry. For example, Logger.log (no parentheses) will work.
* @return {*} The value returned by the called function.
*/
function call(func, optLoggerFunction) {
  for (var n=0; n<6; n++) {
    try {
      return func();
    } catch(e) {
      if (optLoggerFunction) {optLoggerFunction("GASRetry " + n + ": " + e)}
      if (n == 5) {
        throw e;
      } 
      Utilities.sleep((Math.pow(2,n)*1000) + (Math.round(Math.random() * 1000)));
    }    
  }
}





function sheetSpider_institutionalTrackingUi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var institutionalTrackingString = UserProperties.getProperty('institutionalTrackingString');
  var eduSetting = UserProperties.getProperty('eduSetting');
  if (!(institutionalTrackingString)) {
    UserProperties.setProperty('institutionalTrackingString', 'not participating');
  }
  var app = UiApp.createApplication().setTitle('Hello there! Help us track the usage of this script').setHeight(400);
  if ((!(institutionalTrackingString))||(!(eduSetting))) {
    var helptext = app.createLabel("You are most likely seeing this prompt because this is the first time you are using a Google Apps script created by New Visions for Public Schools, 501(c)3. If you are using scripts as part of a school or grant-funded program like New Visions' CloudLab, you may wish to track usage rates with Google Analytics. Entering tracking information here will save it to your user credentials and enable tracking for any other New Visions scripts that use this feature. No personal info will ever be collected.").setStyleAttribute('marginBottom', '10px');
  } else {
    var helptext = app.createLabel("If you are using scripts as part of a school or grant-funded program like New Visions' CloudLab, you may wish to track usage rates with Google Analytics. Entering or modifying tracking information here will save it to your user credentials and enable tracking for any other scripts produced by New Visions for Public Schools, 501(c)3, that use this feature. No personal info will ever be collected.").setStyleAttribute('marginBottom', '10px');
  }
  var panel = app.createVerticalPanel();
  var gridPanel = app.createVerticalPanel().setId("gridPanel").setVisible(false);
  var grid = app.createGrid(4,2).setId('trackingGrid').setStyleAttribute('background', 'whiteSmoke').setStyleAttribute('marginTop', '10px');
  var checkHandler = app.createServerHandler('sheetSpider_refreshTrackingGrid').addCallbackElement(panel);
  var checkBox = app.createCheckBox('Participate in institutional usage tracking.  (Only choose this option if you know your institution\'s Google Analytics tracker Id.)').setName('trackerSetting').addValueChangeHandler(checkHandler);  
  var checkBox2 = app.createCheckBox('Let New Visions for Public Schools, 501(c)3 know you\'re an educational user.').setName('eduSetting');  
  if ((institutionalTrackingString == "not participating")||(institutionalTrackingString=='')) {
    checkBox.setValue(false);
  } 
  if (eduSetting=="true") {
    checkBox2.setValue(true);
  }
  var institutionNameFields = [];
  var trackerIdFields = [];
  var institutionNameLabel = app.createLabel('Institution Name');
  var trackerIdLabel = app.createLabel('Google Analytics Tracker Id (UA-########-#)');
  grid.setWidget(0, 0, institutionNameLabel);
  grid.setWidget(0, 1, trackerIdLabel);
  if ((institutionalTrackingString)&&((institutionalTrackingString!='not participating')||(institutionalTrackingString==''))) {
    checkBox.setValue(true);
    gridPanel.setVisible(true);
    var institutionalTrackingObject = Utilities.jsonParse(institutionalTrackingString);
  } else {
    var institutionalTrackingObject = new Object();
  }
  for (var i=1; i<4; i++) {
    institutionNameFields[i] = app.createTextBox().setName('institution-'+i);
    trackerIdFields[i] = app.createTextBox().setName('trackerId-'+i);
    if (institutionalTrackingObject) {
      if (institutionalTrackingObject['institution-'+i]) {
        institutionNameFields[i].setValue(institutionalTrackingObject['institution-'+i]['name']);
        if (institutionalTrackingObject['institution-'+i]['trackerId']) {
          trackerIdFields[i].setValue(institutionalTrackingObject['institution-'+i]['trackerId']);
        }
      }
    }
    grid.setWidget(i, 0, institutionNameFields[i]);
    grid.setWidget(i, 1, trackerIdFields[i]);
  } 
  var help = app.createLabel('Enter up to three institutions, with Google Analytics tracker Id\'s.').setStyleAttribute('marginBottom','5px').setStyleAttribute('marginTop','10px');
  gridPanel.add(help);
  gridPanel.add(grid); 
  panel.add(helptext);
  panel.add(checkBox2);
  panel.add(checkBox);
  panel.add(gridPanel);
  var button = app.createButton("Save settings");
  var saveHandler = app.createServerHandler('sheetSpider_saveInstitutionalTrackingInfo').addCallbackElement(panel);
  button.addClickHandler(saveHandler);
  panel.add(button);
  app.add(panel);
  ss.show(app);
  return app;
}

function sheetSpider_refreshTrackingGrid(e) {
  var app = UiApp.getActiveApplication();
  var gridPanel = app.getElementById("gridPanel");
  var grid = app.getElementById("trackingGrid");
  var setting = e.parameter.trackerSetting;
  if (setting=="true") {
    gridPanel.setVisible(true);
  } else {
    gridPanel.setVisible(false);
  }
  return app;
}

function sheetSpider_saveInstitutionalTrackingInfo(e) {
  var app = UiApp.getActiveApplication();
  var eduSetting = e.parameter.eduSetting;
  var oldEduSetting = UserProperties.getProperty('eduSetting')
  if (eduSetting == "true") {
    UserProperties.setProperty('eduSetting', 'true');
  }
  if ((oldEduSetting)&&(eduSetting=="false")) {
    UserProperties.setProperty('eduSetting', 'false');
  }
  var trackerSetting = e.parameter.trackerSetting;
  if (trackerSetting == "false") {
    UserProperties.setProperty('institutionalTrackingString', 'not participating');
    app.close();
    return app;
  } else {
    var institutionalTrackingObject = new Object;
    for (var i=1; i<4; i++) {
      var checkVal = e.parameter['institution-'+i];
      if (checkVal!='') {
        institutionalTrackingObject['institution-'+i] = new Object();
        institutionalTrackingObject['institution-'+i]['name'] = e.parameter['institution-'+i];
        institutionalTrackingObject['institution-'+i]['trackerId'] = e.parameter['trackerId-'+i];
        if (!(e.parameter['trackerId-'+i])) {
          Browser.msgBox("You entered an institution without a Google Analytics Tracker Id");
          sheetSpider_institutionalTrackingUi()
        }
      }
    }
    var institutionalTrackingString = Utilities.jsonStringify(institutionalTrackingObject);
    UserProperties.setProperty('institutionalTrackingString', institutionalTrackingString);
    app.close();
    return app;
  }
}


// This code was borrowed and modified from the Flubaroo Script author Dave Abouav
// It anonymously tracks script usage to Google Analytics, allowing our non-profit to report our impact to funders
// For original source see http://www.edcode.org

function sheetSpider_createInstitutionalTrackingUrls(institutionTrackingObject, encoded_page_name, encoded_script_name) {
  for (var key in institutionTrackingObject) {
    var utmcc = sheetSpider_createGACookie();
    if (utmcc == null)
    {
      return null;
    }
    var encoded_page_name = encoded_script_name+"/"+encoded_page_name;
    var trackingId = institutionTrackingObject[key].trackerId;
    var ga_url1 = "http://www.google-analytics.com/__utm.gif?utmwv=5.2.2&utmhn=www.sheetSpider-analytics.com&utmcs=-&utmul=en-us&utmje=1&utmdt&utmr=0=";
    var ga_url2 = "&utmac="+trackingId+"&utmcc=" + utmcc + "&utmu=DI~";
    var ga_url_full = ga_url1 + encoded_page_name + "&utmp=" + encoded_page_name + ga_url2;
    
    if (ga_url_full)
    {
      try {
        var response = UrlFetchApp.fetch(ga_url_full);
      } catch(err) {
      }
    }
  }
}



function sheetSpider_createGATrackingUrl(encoded_page_name)
{
  var utmcc = sheetSpider_createGACookie();
  var eduSetting = UserProperties.getProperty('eduSetting');
  if (eduSetting=="true") {
    encoded_page_name = "edu/" + encoded_page_name;
  }
  if (utmcc == null)
  {
    return null;
  }
  
  var ga_url1 = "http://www.google-analytics.com/__utm.gif?utmwv=5.2.2&utmhn=www.sheetSpider-analytics.com&utmcs=-&utmul=en-us&utmje=1&utmdt&utmr=0=";
  var ga_url2 = "&utmac=UA-41943014-1&utmcc=" + utmcc + "&utmu=DI~";
  var ga_url_full = ga_url1 + encoded_page_name + "&utmp=" + encoded_page_name + ga_url2;
  
  return ga_url_full;
}


function sheetSpider_createSystemTrackingUrls(institutionTrackingObject, encoded_system_name, encoded_execution_name) {
  for (var key in institutionTrackingObject) {
    var utmcc = sheetSpider_createGACookie();
    if (utmcc == null)
    {
      return null;
    }
    var trackingId = institutionTrackingObject[key].trackerId;
    var encoded_page_name = encoded_system_name+"/"+encoded_execution_name;
    var ga_url1 = "http://www.google-analytics.com/__utm.gif?utmwv=5.2.2&utmhn=www.cloudlab-systems-analytics.com&utmcs=-&utmul=en-us&utmje=1&utmdt&utmr=0=";
    var ga_url2 = "&utmac="+trackingId+"&utmcc=" + utmcc + "&utmu=DI~";
    var ga_url_full1 = ga_url1 + encoded_page_name + "&utmp=" + encoded_page_name + ga_url2;
    if (ga_url_full1)
    {
      try {
        var response = UrlFetchApp.fetch(ga_url_full1);
      } catch(err) {
      }
    } 
  }
  
  var encoded_page_name = encoded_system_name+"/"+encoded_execution_name;
  var ga_url1 = "http://www.google-analytics.com/__utm.gif?utmwv=5.2.2&utmhn=www.cloudlab-systems-analytics.com&utmcs=-&utmul=en-us&utmje=1&utmdt&utmr=0=";
  var ga_url2 = "&utmac=UA-34521561-1&utmcc=" + utmcc + "&utmu=DI~";
  var ga_url_full2 = ga_url1 + encoded_page_name + "&utmp=" + encoded_page_name + ga_url2;
  if (ga_url_full2)
  {
    try {
      var response = UrlFetchApp.fetch(ga_url_full2);
    }
    catch(err) {
    }
  }
  
}



function sheetSpider_createGACookie()
{
  var a = "";
  var b = "100000000";
  var c = "200000000";
  var d = "";
  
  var dt = new Date();
  var ms = dt.getTime();
  var ms_str = ms.toString();
  
  var sheetSpider_uid = UserProperties.getProperty("sheetSpider_uid");
  if ((sheetSpider_uid == null) || (sheetSpider_uid == ""))
  {
    // shouldn't happen unless user explicitly removed flubaroo_uid from properties.
    return null;
  }
  
  a = sheetSpider_uid.substring(0,9);
  d = sheetSpider_uid.substring(9);
  
  utmcc = "__utma%3D451096098." + a + "." + b + "." + c + "." + d 
  + ".1%3B%2B__utmz%3D451096098." + d + ".1.1.utmcsr%3D(direct)%7Cutmccn%3D(direct)%7Cutmcmd%3D(none)%3B";
  
  return utmcc;
}



function sheetSpider_logEntitySheetProvisioned()
{
  var ga_url = sheetSpider_createGATrackingUrl("Entity%20Spreadsheet%20Provisioned");
  if (ga_url)
  {
    try {
      var response = UrlFetchApp.fetch(ga_url);
    } catch(err) {
    }
  }
  var institutionalTrackingObject = sheetSpider_getInstitutionalTrackerObject();
  if (institutionalTrackingObject) {
    sheetSpider_createInstitutionalTrackingUrls(institutionalTrackingObject,"Entity%20Spreadsheet%20Provisioned", "sheetSpider");
    var systemName = ScriptProperties.getProperty('systemName');
    if (systemName) {
      var encoded_system_name = urlencode(systemName);
      sheetSpider_createSystemTrackingUrls(institutionalTrackingObject, encoded_system_name, "Entity%20Spreadsheet%20Provisioned")
    }
  }
}



function sheetSpider_logFormDataPushed()
{
  var ga_url = sheetSpider_createGATrackingUrl("Form%20Data%20Pushed");
  if (ga_url)
  {
    try {
      var response = UrlFetchApp.fetch(ga_url);
    } catch(err) {
    }
  }
  var institutionalTrackingObject = sheetSpider_getInstitutionalTrackerObject();
  if (institutionalTrackingObject) {
    sheetSpider_createInstitutionalTrackingUrls(institutionalTrackingObject,"Form%20Data%20Pushed", "sheetSpider");
    var systemName = ScriptProperties.getProperty('systemName');
    if (systemName) {
      var encoded_system_name = urlencode(systemName);
      sheetSpider_createSystemTrackingUrls(institutionalTrackingObject, encoded_system_name, "Form%20Data%20Pushed")
    }
  }
}



function sheetSpider_logManualPush()
{
  var ga_url = sheetSpider_createGATrackingUrl("Manual%20Data%20Push");
  if (ga_url)
  {
    try {
      var response = UrlFetchApp.fetch(ga_url);
    } catch(err) {
    }
  }
  var institutionalTrackingObject = sheetSpider_getInstitutionalTrackerObject();
  if (institutionalTrackingObject) {
    sheetSpider_createInstitutionalTrackingUrls(institutionalTrackingObject,"Manual%20Data%20Push", "sheetSpider");
    var systemName = ScriptProperties.getProperty('systemName');
    if (systemName) {
      var encoded_system_name = urlencode(systemName);
      sheetSpider_createSystemTrackingUrls(institutionalTrackingObject, encoded_system_name, "Manual%20Data%20Push")
    }
  }
}



function sheetSpider_logEntityDataReturned()
{
  var ga_url = sheetSpider_createGATrackingUrl("Entity%20Data%20Returned");
  if (ga_url)
  {
    try {
      var response = UrlFetchApp.fetch(ga_url);
    } catch(err) {
    }
  }
  var institutionalTrackingObject = sheetSpider_getInstitutionalTrackerObject();
  if (institutionalTrackingObject) {
    sheetSpider_createInstitutionalTrackingUrls(institutionalTrackingObject,"Entity%20Data%20Returned", "sheetSpider");
    var systemName = ScriptProperties.getProperty('systemName');
    if (systemName) {
      var encoded_system_name = urlencode(systemName);
      sheetSpider_createSystemTrackingUrls(institutionalTrackingObject, encoded_system_name, "Entity%20Data%20Returned")
    }
  }
}



function sheetSpider_logFeederDataUpdated()
{
  var ga_url = sheetSpider_createGATrackingUrl("Feeder%20Data%20Updated");
  if (ga_url)
  {
    try {
      var response = UrlFetchApp.fetch(ga_url);
    } catch(err) {
    }
  }
  var institutionalTrackingObject = sheetSpider_getInstitutionalTrackerObject();
  if (institutionalTrackingObject) {
    sheetSpider_createInstitutionalTrackingUrls(institutionalTrackingObject,"Feeder%20Data%20Updated", "sheetSpider");
    var systemName = ScriptProperties.getProperty('systemName');
    if (systemName) {
      var encoded_system_name = urlencode(systemName);
      sheetSpider_createSystemTrackingUrls(institutionalTrackingObject, encoded_system_name, "Feeder%20Data%20Updated");
    }
  }
}


function sheetSpider_getInstitutionalTrackerObject() {
  var institutionalTrackingString = UserProperties.getProperty('institutionalTrackingString');
  if ((institutionalTrackingString)&&(institutionalTrackingString != "not participating")) {
    var institutionTrackingObject = Utilities.jsonParse(institutionalTrackingString);
    return institutionTrackingObject;
  }
  if (!(institutionalTrackingString)||(institutionalTrackingString='')) {
    sheetSpider_institutionalTrackingUi();
    return;
  }
}


function sheetSpider_logRepeatInstall()
{
  var ga_url = sheetSpider_createGATrackingUrl("Repeat%20Install");
  if (ga_url)
  {
    try {
      var response = UrlFetchApp.fetch(ga_url);
    } catch(err) {
    }
  }
  var institutionalTrackingObject = sheetSpider_getInstitutionalTrackerObject();
  if (institutionalTrackingObject) {
    sheetSpider_createInstitutionalTrackingUrls(institutionalTrackingObject,"Repeat%20Install", "sheetSpider");
    var systemName = ScriptProperties.getProperty('systemName');
    if (systemName) {
      var encoded_system_name = urlencode(systemName);
      sheetSpider_createSystemTrackingUrls(institutionalTrackingObject, encoded_system_name, "Repeat%20Install")
    }
  }
}

function sheetSpider_logFirstInstall()
{
  var ga_url = sheetSpider_createGATrackingUrl("First%20Install");
  if (ga_url)
  { 
    try {
      var response = UrlFetchApp.fetch(ga_url);
    } catch(err) {
    }
  }
  var institutionalTrackingObject = sheetSpider_getInstitutionalTrackerObject();
  if (institutionalTrackingObject) {
    sheetSpider_createInstitutionalTrackingUrls(institutionalTrackingObject,"First%20Install", "sheetSpider");
    var systemName = ScriptProperties.getProperty('systemName');
    if (systemName) {
      var encoded_system_name = urlencode(systemName);
      sheetSpider_createSystemTrackingUrls(institutionalTrackingObject, encoded_system_name, "First%20Install")
    }
  }
}


function setsheetSpiderUid()
{ 
  var sheetSpider_uid = UserProperties.getProperty("sheetSpider_uid");
  if (sheetSpider_uid == null || sheetSpider_uid == "")
  {
    // user has never installed sheetSpider before (in any spreadsheet)
    var dt = new Date();
    var ms = dt.getTime();
    var ms_str = ms.toString();
    
    UserProperties.setProperty("sheetSpider_uid", ms_str);
    sheetSpider_logFirstInstall();
  }
}


function setsheetSpiderSid()
{ 
  var sheetSpider_sid = ScriptProperties.getProperty("sheetSpider_sid");
  if (sheetSpider_sid == null || sheetSpider_sid == "")
  {
    // user has never installed sheetSpider before (in any spreadsheet)
    var dt = new Date();
    var ms = dt.getTime();
    var ms_str = ms.toString();
    ScriptProperties.setProperty("sheetSpider_sid", ms_str);
    var sheetSpider_uid = UserProperties.getProperty("sheetSpider_uid");
    if (sheetSpider_uid != null || sheetSpider_uid != "") {
      sheetSpider_logRepeatInstall();
    }
  }
}
