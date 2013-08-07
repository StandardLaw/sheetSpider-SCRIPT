function pushData(e) {
    var lock = LockService.getPublicLock();
    var success = lock.tryLock(500);
    var numRecordsPushed = 0;
    var entitiesAffected = [];
  if (e) {
    lock.waitLock(240000);
  } else {
    var app = UiApp.getActiveApplication();
  }
  if (success) {
    var properties = ScriptProperties.getProperties();
    var mode = properties.mode;
    var ssKey = properties.ssKey;
    var ss = SpreadsheetApp.openById(ssKey);
    var feederSheet = getSheetById(ss, properties.feederSheetId);
    var lastRow = feederSheet.getLastRow();
    if (lastRow<2) {
      Browser.msgBox("You have no data in your feeder sheet yet.");
      if (app) {
        app.close();
        return app;
      } else {
        return;
      }
    }
    var feederSheetHeaders = feederSheet.getRange(1, 1, 1, feederSheet.getLastColumn()).getValues()[0];
    var entitySheet = getSheetById(ss, properties.entitySheetId);
    var entitySheetDataRange = entitySheet.getRange(2, 1, entitySheet.getLastRow()-1, entitySheet.getLastColumn());
    var entityData = getRowsData(entitySheet, entitySheetDataRange);
    var entityColNormalized = normalizeHeader(properties.entityCol);
    var feederEntityColNormalized = normalizeHeader(properties.feederEntityCol);
    var uniquenessCriterionNormalized = normalizeHeaders(properties.uniquenessCriterion.split('||'));
    if (mode == "on Google Form submit") {
      var thisRow = SpreadsheetApp.getActiveSheet().getActiveRange().getRow()
      var range = feederSheet.getRange(thisRow, 1, 1, feederSheet.getLastColumn());
      var dataRow = getRowsData(feederSheet, range, 1)[0];   
      var entities = dataRow[feederEntityColNormalized].split(",");
      
      for (var i=0; i<entities.length; i++) {
        dataRow[feederEntityColNormalized] = entities[i];
        var entityKey = fetchEntityKey(entities[i].trim(), entityData, entityColNormalized);
        var destSS = SpreadsheetApp.openById(entityKey);
        var destSheet = checkFixSheetHeaders(destSS, properties.feederSheetName, feederSheetHeaders);
        var destSheetHeaderRange = destSheet.getRange(1, 1, 1, destSheet.getLastColumn());
        if (destSheet.getLastRow()>1) {
          var destSheetDataRange = destSheet.getRange(2, 1, destSheet.getLastRow()-1, destSheet.getLastColumn());
          var destData = getRowsData(destSheet, destSheetDataRange);
          var checkRecord = isRecordUnique(destData, dataRow, uniquenessCriterionNormalized);
        } else {
          var checkRecord = true;
        }
        if (checkRecord) {
          var lastRow = destSheet.getLastRow();
          var lastCol = destSheet.getLastColumn();
          destSheet.insertRowAfter(lastRow);
          var destRange = destSheet.getRange(lastRow+1, 1, 1, lastCol);
          setRowsData(destSheet, [dataRow], destSheetHeaderRange, lastRow+1);
          sheetSpider_logFormDataPushed();
        }
      }
    } else if (properties.overwriteMode == "Append only new unique records") {
      var feederRange = feederSheet.getRange(2, 1, feederSheet.getLastRow(), feederSheet.getLastColumn());
      var feederData = getRowsData(feederSheet, feederRange);
      var allEntities = getEntityValues(entitySheet, properties.entityCol);
      var entitySheetData = new Object();
      var entitySheets = new Object();
      var entityKeys = [];
      for (var s=0; s<allEntities.length; s++) {
        var entityKey = entityData[s]['spreadsheetId'];
        var destSS = SpreadsheetApp.openById(entityKey);
        var destSheet = checkFixSheetHeaders(destSS, properties.feederSheetName, feederSheetHeaders);
        if (destSheet.getLastRow()>1) {
          var destSheetDataRange = destSheet.getRange(2, 1, destSheet.getLastRow()-1, destSheet.getLastColumn());
          var destData = getRowsData(destSheet, destSheetDataRange);
        } else {
          var destData = [];
        }
        entitySheetData[entityKey] = destData;
        entitySheets[entityKey] = destSheet;
        entityKeys.push(entityKey);
      }
      for (var j=0; j<allEntities.length; j++) {
        var thisEntityArray = [];
        for (var k=0; k<feederData.length; k++) {
          var feederDataRow = clone(feederData[k]);   
          var entities = feederDataRow[feederEntityColNormalized].split(",");
          for (var i=0; i<entities.length; i++) {
            if (entities[i].trim() == allEntities[j].trim()) {
              var entityKey = fetchEntityKey(entities[i].trim(), entityData, entityColNormalized);
              feederDataRow[feederEntityColNormalized] = entities[i].trim();
              if (entitySheetData[entityKey].length>0) {
                var checkRecord = isRecordUnique(entitySheetData[entityKey], feederDataRow, uniquenessCriterionNormalized);
              } else {
                var checkRecord = true;
              }
              if (checkRecord) {
                thisEntityArray.push(feederDataRow);
                numRecordsPushed++;
                if (entitiesAffected.indexOf(entities[i].trim())==-1) {
                  entitiesAffected.push(entities[i].trim());
                }
              }
            }
          }
        }
        var destSheet = entitySheets[entityKeys[j]];
        var destSheetHeaderRange = destSheet.getRange(1, 1, 1, destSheet.getLastColumn());
        if (thisEntityArray.length>0) {
          setRowsData(destSheet, thisEntityArray, destSheetHeaderRange, destSheet.getLastRow()+1);
        }
      }
      sheetSpider_logManualPush();
    } else if (properties.overwriteMode == "Overwrite all existing data") { 
      var feederRange = feederSheet.getRange(2, 1, feederSheet.getLastRow(), feederSheet.getLastColumn());
      var feederData = getRowsData(feederSheet, feederRange);
      var allEntities = getEntityValues(entitySheet, properties.entityCol);
      var entitySheets = new Object();
      var entityKeys = [];
      for (var s=0; s<allEntities.length; s++) {
        var entityKey = entityData[s]['spreadsheetId'];
        var destSS = SpreadsheetApp.openById(entityKey);
        var destSheet = checkFixSheetHeaders(destSS, properties.feederSheetName, feederSheetHeaders);
        var thisEntityData = [];
        for (var i=0; i<feederData.length; i++) {
          if (feederData[i][feederEntityColNormalized].split(",").indexOf(allEntities[s])!=-1) {
            thisEntityData.push(feederData[i]);
            numRecordsPushed++;
              if (entitiesAffected.indexOf(allEntities[s].trim())==-1) {
                  entitiesAffected.push(allEntities[s].trim());
                }
          }
        }
        if (destSheet.getLastRow()>1) {
          var destSheetDataRange = destSheet.getRange(2, 1, destSheet.getLastRow()-1, destSheet.getLastColumn());
          destSheetDataRange.clearContent();
        } else {
          var destData = [];
        }
        if (thisEntityData.length>0) {
          var newDestSheetDataRange = destSheet.getRange(2, 1, thisEntityData.length, destSheet.getLastColumn());
          setRowsData(destSheet, thisEntityData);
        }
      }
      sheetSpider_logManualPush();
    }
    lock.releaseLock();
    if (properties.overwriteMode != "on Google Form submit") {
      if (numRecordsPushed>0) {
        Browser.msgBox("Data push completed with " + numRecordsPushed + " records pushed to the following entity Spreadsheets: " + entitiesAffected);
      } else {
        Browser.msgBox("No records were pushed.");
      }
    }
  } else {
    Browser.msgBox('Another process is already attempting to run the disaggregation and push of data.');
  }
}

function isRecordUnique(entityData, dataRow, uniquenessCriterionNormalized) {
  var thisUniqueString = '';
  for (var j=0; j<uniquenessCriterionNormalized.length; j++) {
    var thisValue = '';
    if (uniquenessCriterionNormalized[j] == 'timestamp') {
      thisValue = Number(dataRow[uniquenessCriterionNormalized[j]]);
    } else {
      thisValue = dataRow[uniquenessCriterionNormalized[j]];
    }
    thisUniqueString += thisValue;
  }
  for (var i=0; i<entityData.length; i++) {
    var thatUniqueString = '';
    for (var j=0; j<uniquenessCriterionNormalized.length; j++) {
      var thatValue = '';
      if (uniquenessCriterionNormalized[j] == 'timestamp') {
        thatValue = Number(entityData[i][uniquenessCriterionNormalized[j]]);
      } else {
        thatValue = entityData[i][uniquenessCriterionNormalized[j]];
      }
      thatUniqueString += thatValue;
    }
    if (thisUniqueString == thatUniqueString) {
      return false;
    }
  }
  return true;
}

function clone(obj) {
    if (null == obj || "object" != typeof obj) return obj;
    var copy = obj.constructor();
    for (var attr in obj) {
        if (obj.hasOwnProperty(attr)) copy[attr] = obj[attr];
    }
    return copy;
}


function fetchEntityKey(entity, entityData, entityColNormalized) {
  for (var i=0; i<entityData.length; i++) {
    if (entity == entityData[i][entityColNormalized]) {
      return entityData[i].spreadsheetId;
    }
  }
  return;
}
