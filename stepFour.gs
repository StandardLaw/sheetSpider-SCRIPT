function retrieveLiveData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var normalizedFeederEntityCol = normalizeHeader(properties.feederEntityCol);
  var entitySheet = getSheetById(ss, properties.entitySheetId);
  var feederSheet = getSheetById(ss, properties.feederSheetId);
  var feederSheetName = properties.feederSheetName;
  var feederSheetHeaders = feederSheet.getRange(1,1,1,feederSheet.getLastColumn()).getValues()[0];
  var normalizedFeederSheetHeaders = normalizeHeaders(feederSheetHeaders);
  feederSheetHeaders.push("Change Status");
  var feederRange = feederSheet.getRange(2, 1, feederSheet.getLastRow()-1, feederSheet.getLastColumn());
  var feederSheetData = getRowsData(feederSheet, feederRange);
  feederSheetData.sort(function(a, b){
    var nameA=a[normalizedFeederEntityCol].toLowerCase(), nameB=b[normalizedFeederEntityCol].toLowerCase()
    if (nameA < nameB) //sort string ascending
      return -1 
      if (nameA > nameB)
        return 1
        return 0 //default return value (no sorting)
  });
  var entityRange = entitySheet.getRange(2, 1, entitySheet.getLastRow()-1, entitySheet.getLastColumn());
  var entitySheetData = getRowsData(entitySheet, entityRange);
  var allData = [];
  for (var i=0; i<entitySheetData.length; i++) {
    var thisSSKey = entitySheetData[i]['spreadsheetId'];
    var thisSS = SpreadsheetApp.openById(thisSSKey);
    var thisSheet = checkFixSheetHeaders(thisSS, feederSheetName, feederSheetHeaders);
    var thisLastRow = thisSheet.getLastRow();
    if (thisLastRow > 1) {
      var thisRange = thisSheet.getRange(2, 1, thisSheet.getLastRow()-1, thisSheet.getLastColumn());
      var thisData = getRowsData(thisSheet, thisRange);
      allData = allData.concat(thisData);
    }
  }
  allData.sort(function(a, b){
    var nameA=a[normalizedFeederEntityCol].toLowerCase(), nameB=b[normalizedFeederEntityCol].toLowerCase()
    if (nameA < nameB) //sort string ascending
      return -1 
      if (nameA > nameB)
        return 1
        return 0 //default return value (no sorting)
  });
  var uniquenessCriterion = properties.uniquenessCriterion;
  var dataReturned = markChanges(feederSheetData, allData, uniquenessCriterion);
  var date = new Date();
  date = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), ' - M.d.yy - h:mm:ss a');
  var dataSheet = ss.getSheetByName('Returned Data');
  if (!dataSheet) {
    dataSheet = ss.insertSheet('Returned Data');
  }
  dataSheet.clear().clearNotes();
  dataSheet.getRange(1, 1, 1, feederSheetHeaders.length).setValues([feederSheetHeaders]).setBackground('#C0C0C0').setNote('Data last retrieved at ' + date);
  dataSheet.setFrozenRows(1);
  if (dataReturned.newData.length>0) {
    var maxRows = dataSheet.getMaxRows();
    if (dataReturned.newData.length > maxRows) {
      dataSheet.insertRows(maxRows, dataReturned.newData.length - maxRows + 10);
    }
    var backgroundRange = dataSheet.getRange(2, 1, dataReturned.newData.length, dataSheet.getLastColumn());
    var backgrounds = backgroundRange.getBackgrounds();
    var notes = backgroundRange.getNotes();
    for (var i=0; i<dataReturned.newData.length; i++) {
      var changeArray = dataReturned.newData[i].changeStatus.split("||");
      if (changeArray[0] == 'edited') {
        for (var k=1; k<changeArray.length; k++) {
          var thisIndex = normalizedFeederSheetHeaders.indexOf(changeArray[k]);
          backgrounds[i][thisIndex] = "yellow";
          debugger;
          notes[i][thisIndex] = "The value in this cell is different from the corresponding value in the feeder sheet.  It was likely edited in the entity sheet.";
        }
      }
      if (changeArray[0] == 'added') {
        notes[i][0] = "Values in this row did not match any records found in the feeder sheet.  It's likely they have been added to the entity spreadsheet.";
        for (var j=0; j<backgrounds[0].length; j++) {
          backgrounds[i][j] = "orange";
        }
      }
    }
    setRowsData(dataSheet, dataReturned.newData);
    backgroundRange.setBackgrounds(backgrounds);
    backgroundRange.setNotes(notes);
  }

  if (dataReturned.notFound.length>0) {
    var deletedSheet = ss.getSheetByName('Records Not Found');
    if (!deletedSheet) {
      deletedSheet = ss.insertSheet('Records Not Found');
    }
    deletedSheet.clear();
    deletedSheet.getRange(1, 1, 1, feederSheetHeaders.length).setValues([feederSheetHeaders]).setBackground('#C0C0C0').setNote('The following rows were not found in the entity spreadsheets. Data last retrieved at ' + date);
    deletedSheet.setFrozenRows(1);
    var backgroundRange = deletedSheet.getRange(2, 1, dataReturned.notFound.length, deletedSheet.getLastColumn());
    var backgrounds = backgroundRange.getBackgrounds();
    var notes = backgroundRange.getNotes();
    for (var i=0; i<dataReturned.notFound.length; i++) {
      if (dataReturned.notFound[i].changeStatus.split("||")[0] == 'deleted') {
        notes[i][0] = "Values in this row were found in feeder sheet but NOT in the entity sheets.  Either these records never pushed, or they were deleted from the entity spreadsheet(s) they were pushed to.";
        for (var j=0; j<backgrounds[0].length; j++) {
          backgrounds[i][j] = "pink";
        }
      } else {
        notes[i][feederSheetHeaders.length-1] = "";
        for (var j=0; j<backgrounds[0].length; j++) {
          backgrounds[i][j] = "white";
        }
      }
    }
    setRowsData(deletedSheet, dataReturned.notFound);
    backgroundRange.setBackgrounds(backgrounds);
    backgroundRange.setNotes(notes);
    sheetSpider_logEntityDataReturned();
  } else {
    var deletedSheet = ss.getSheetByName('Records Not Found');
    if (deletedSheet) {
      deletedSheet.clear();
      deletedSheet.getRange(1, 1, 1, feederSheetHeaders.length).setValues([feederSheetHeaders]).setBackground('#C0C0C0').setNote('');
    } 
  }
  if (dataReturned.newData.length > 0) {
    properties.stepFourComplete = true; 
    ScriptProperties.setProperties(properties);
  }
  onOpen();
  Browser.msgBox("Data retrieval from entity spreadsheets completed successfully with " + dataReturned.changeCount + " change(s) detected, " + dataReturned.addedCount + " record(s) added, and " + dataReturned.deletedCount + " record(s) deleted. Please examine the \"Returned Data\" and \"Records Not Found\" sheets and determine whether you want to update the feeder sheet with the returned values.");
  return;
}


function markChanges(referenceData, newData, uniquenessCriterion) {
  var returnObj = new Object();
  var uniquenessCriterionArray = normalizeHeaders(uniquenessCriterion.split("||"));
  var matchedRef = [];
  var allRef = [];
  var changeCount = 0;
  var deletedCount = 0;
  var addedCount = 0;
  for (var i = 0; i<newData.length; i++) {
    var thisRow = clone(newData[i]);
    delete thisRow.changeStatus;
    var thisUnique = getUniqueString(thisRow, uniquenessCriterionArray);
    var found = false;
    for (var j=0; j<referenceData.length; j++) {
      if (allRef.indexOf(j)==-1) {
        allRef.push(j);
      }
      if (matchedRef.indexOf(j)==-1) {
        var referenceRow = referenceData[j];
        var refUnique = getUniqueString(referenceRow, uniquenessCriterionArray);
        if (thisUnique == refUnique) {
          found = true;
          matchedRef.push(j);
          if (JSON.stringify(referenceRow) == JSON.stringify(thisRow)) {
            newData[i].changeStatus = 'unchanged';
          } else {
            changeCount++;
            newData[i].changeStatus = 'edited';
            for (var key in referenceRow) {
              var thisValue = thisRow[key];
              var thisReference = referenceRow[key];
              var type = returnType(thisReference);
              if (type == "date") {
                thisValue = Number(thisValue);
                thisReference = Number(thisReference);
              }
              if(thisValue!=thisReference) {
                newData[i].changeStatus += "||";
                newData[i].changeStatus += key;
              }
            }
          }
          break;
        }
      }
    }
    if (found == false) {
      addedCount++;
      newData[i].changeStatus = 'added';
    }
  }
  var notFound = diffArray(allRef, matchedRef);
  var notFoundData = [];
  for (var i=0; i<notFound.length; i++) {
    notFoundData[i] = referenceData[notFound[i]];
    notFoundData[i].changeStatus = "deleted";
    deletedCount++;
  }
  returnObj.newData = newData;
  returnObj.notFound = notFoundData;
  returnObj.changeCount = changeCount;
  returnObj.deletedCount = deletedCount;
  returnObj.addedCount = addedCount;
  return returnObj;
}


function getUniqueString(rowData, uniquenessKeys) {
  var uniqueString = '';
  for (var i in uniquenessKeys) {
    uniqueString += rowData[uniquenessKeys[i]];
  }
  return uniqueString;
}


function diffArray(a, b) {
  var seen = [], diff = [];
  for ( var i = 0; i < b.length; i++)
    seen[b[i]] = true;
  for ( var i = 0; i < a.length; i++)
    if (!seen[a[i]])
      diff.push(a[i]);
  return diff;
}



function clone(obj) {
    // Handle the 3 simple types, and null or undefined
    if (null == obj || "object" != typeof obj) return obj;

    // Handle Date
    if (obj instanceof Date) {
        var copy = new Date();
        copy.setTime(obj.getTime());
        return copy;
    }

    // Handle Array
    if (obj instanceof Array) {
        var copy = [];
        for (var i = 0, len = obj.length; i < len; i++) {
            copy[i] = clone(obj[i]);
        }
        return copy;
    }

    // Handle Object
    if (obj instanceof Object) {
        var copy = {};
        for (var attr in obj) {
            if (obj.hasOwnProperty(attr)) copy[attr] = clone(obj[attr]);
        }
        return copy;
    }

    throw new Error("Unable to copy obj! Its type isn't supported.");
}
