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
      if ((headers.indexOf(feederSheetHeaders[i])==-1)&&(feederSheetHeaders[i]!="Change Status")) {
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






// This code was borrowed and modified from the Flubaroo Script author Dave Abouav
// It anonymously tracks script usage to Google Analytics, allowing our non-profit to report our impact to funders
// For original source see http://www.edcode.org


function sheetSpider_logEntitySheetProvisioned()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Entity%20Spreadsheet%20Provisioned", scriptName, scriptTrackingId, systemName)
}

function sheetSpider_logFormDataPushed()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Form%20Data%20Pushed", scriptName, scriptTrackingId, systemName)
}

function sheetSpider_logManualPush()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Manual%20Data%20Push", scriptName, scriptTrackingId, systemName)
}

function sheetSpider_logEntityDataReturned()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Entity%20Data%20Returned", scriptName, scriptTrackingId, systemName)
}

function sheetSpider_logFeederDataUpdated()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Feeder%20Data%20Updated", scriptName, scriptTrackingId, systemName)
}

function sheetSpider_logRepeatInstall()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Repeat%20Install", scriptName, scriptTrackingId, systemName)
}

function sheetSpider_logFirstInstall()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("First%20Install", scriptName, scriptTrackingId, systemName)
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
    }else{
      sheetSpider_logFirstInstall();
    }
  }
}


function returnType(obj) {
  return ({}).toString.call(obj).match(/\s([a-zA-Z]+)/)[1].toLowerCase()
}
