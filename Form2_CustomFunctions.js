function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Generate Sequence', 'generateSequence2')
      .addItem('Store IDs', 'storeCustomIds2')
      .addToUi();
}



function generateSequence2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourcesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var interimsheet = ss.getSheetByName('HiddenParameterReplicateforLabels');
  // var targetsheet = ss.getSheetByName('SampleIDs for Storage and Labels');

   // Check if the active sheet is Generate SampleIDs
  if (sourcesheet.getName() !== "Generate SampleIDs") {
    // Display an error message if not on Generate SampleIDs
    var ui = SpreadsheetApp.getUi();
    ui.alert('Error', 'This function can only be run on Generate SampleIDs.', ui.ButtonSet.OK);
  return;
  }
  
  // Get the data from column G, called valuesH because of old version didn't feel like changing
  var valuesH = sourcesheet.getRange("G2:G" + sourcesheet.getLastRow()).getValues();

  // Regex pattern to match the specified format "XX_YYYY_yyyy_mm"
  var regexPattern = /^[A-Z]{2}_[A-Z]{3}_\d{4}_\d{2}$/;

  // Filter out values that don't match the regex pattern
  valuesH = valuesH.filter(function(value) {
    return regexPattern.test(value[0]);
  });

  // Log the filtered valuesH array length
  var valuesA = interimsheet.getRange("A2:A" + interimsheet.getLastRow()).getValues();
  var valuesB = interimsheet.getRange("B2:B" + interimsheet.getLastRow()).getValues();
  var valuesC = interimsheet.getRange("C2:C" + interimsheet.getLastRow()).getValues();

  // Loop through the values in column E
  // Loop through the values in column E

  console.log("Length of valuesH array: " + valuesH.length);
  sourcesheet.getRange("H2:I").clearContent();
  
  for (var i = 0; i < valuesH.length; i++) {
    if (i == 0) {
        for (var j = 0; j < valuesA.length; j++) {
            var array = [valuesH[i][0], valuesA[j][0], valuesB[j][0], valuesC[j][0]];
            var newValue = array.join("_");
            sourcesheet.getRange("H" + (j + 2)).setValue(newValue);
            sourcesheet.getRange("I" + (j + 2)).setValue(valuesH[i][0]);
        }
    } else if (i == 1) {
        for (var j = 0; j < valuesA.length; j++) {
            var array = [valuesH[i][0], valuesA[j][0], valuesB[j][0], valuesC[j][0]];
            var newValue = array.join("_");
            sourcesheet.getRange("H" + (j + 102)).setValue(newValue);
            sourcesheet.getRange("I" + (j + 102)).setValue(valuesH[i][0]);
        }
    } else if (i == 2) {
        for (var j = 0; j < valuesA.length; j++) {
            var array = [valuesH[i][0], valuesA[j][0], valuesB[j][0], valuesC[j][0]];
            var newValue = array.join("_");
            sourcesheet.getRange("H" + (j + 202)).setValue(newValue);
            sourcesheet.getRange("I" + (j + 202)).setValue(valuesH[i][0]);
        }
    } else if (i == 3) {
        for (var j = 0; j < valuesA.length; j++) {
            var array = [valuesH[i][0], valuesA[j][0], valuesB[j][0], valuesC[j][0]];
            var newValue = array.join("_");
            sourcesheet.getRange("H" + (j + 302)).setValue(newValue);
            sourcesheet.getRange("I" + (j + 302)).setValue(valuesH[i][0]);
        }
    } else if (i == 4) {
        for (var j = 0; j < valuesA.length; j++) {
            var array = [valuesH[i][0], valuesA[j][0], valuesB[j][0], valuesC[j][0]];
            var newValue = array.join("_");
            sourcesheet.getRange("H" + (j + 402)).setValue(newValue);
            sourcesheet.getRange("I" + (j + 402)).setValue(valuesH[i][0]);
        }
    } else if (i == 5) {
        for (var j = 0; j < valuesA.length; j++) {
            var array = [valuesH[i][0], valuesA[j][0], valuesB[j][0], valuesC[j][0]];
            var newValue = array.join("_");
            sourcesheet.getRange("H" + (j + 502)).setValue(newValue);
            sourcesheet.getRange("I" + (j + 502)).setValue(valuesH[i][0]);
        }
    } else if (i == 6) {
        for (var j = 0; j < valuesA.length; j++) {
            var array = [valuesH[i][0], valuesA[j][0], valuesB[j][0], valuesC[j][0]];
            var newValue = array.join("_");
            sourcesheet.getRange("H" + (j + 602)).setValue(newValue);
            sourcesheet.getRange("I" + (j + 602)).setValue(valuesH[i][0]);
        }
    } else if (i == 7) {
        for (var j = 0; j < valuesA.length; j++) {
            var array = [valuesH[i][0], valuesA[j][0], valuesB[j][0], valuesC[j][0]];
            var newValue = array.join("_");
            sourcesheet.getRange("H" + (j + 702)).setValue(newValue);
            sourcesheet.getRange("I" + (j + 702)).setValue(valuesH[i][0]);
        }
    } else if (i == 8) {
        for (var j = 0; j < valuesA.length; j++) {
            var array = [valuesH[i][0], valuesA[j][0], valuesB[j][0], valuesC[j][0]];
            var newValue = array.join("_");
            sourcesheet.getRange("H" + (j + 802)).setValue(newValue);
            sourcesheet.getRange("I" + (j + 802)).setValue(valuesH[i][0]);
        }
    } else if (i == 9) {
        for (var j = 0; j < valuesA.length; j++) {
            var array = [valuesH[i][0], valuesA[j][0], valuesB[j][0], valuesC[j][0]];
            var newValue = array.join("_");
            sourcesheet.getRange("H" + (j + 902)).setValue(newValue);
            sourcesheet.getRange("I" + (j + 902)).setValue(valuesH[i][0]);
        }
    } else if (i == 10) {
        for (var j = 0; j < valuesA.length; j++) {
            var array = [valuesH[i][0], valuesA[j][0], valuesB[j][0], valuesC[j][0]];
            var newValue = array.join("_");
            sourcesheet.getRange("H" + (j + 1002)).setValue(newValue);
            sourcesheet.getRange("I" + (j + 1002)).setValue(valuesH[i][0]);
        }
    }
  }
}

// active [[GOOGLE FILE ID HERE]]
// desitnation [[GOOGLE FILE ID HERE]]

function storeCustomIds2() {
  var activeSheet = SpreadsheetApp.openById("[[GOOGLE FILE ID HERE]]").getSheetByName("SampleIDs for Storage and Labels");

  // Storing in Parameter database
  // Assuming the custom IDs are generated in columns A through J starting at Row 2
  var startRows = [2, 102, 202, 302, 402, 502, 602, 702, 802, 902, 1002];
  var startCol = 1;
  var numRows = 6; // Take the first six rows of each section
  var numCols = 8;  // Assuming you want to copy 8 columns (A through H)

  var allGeneratedIds = [];
  startRows.forEach(function(startRow) {
    var generatedIdsRange = activeSheet.getRange(startRow, startCol, numRows, numCols);
    var generatedIds = generatedIdsRange.getValues();
    // Filter out rows with empty or "NA" IDs
    generatedIds = generatedIds.filter(row => row[2] !== "" && row[2] !== "NA");
    allGeneratedIds = allGeneratedIds.concat(generatedIds);
  });

  // Open the destination sheet
  var destinationSheet = SpreadsheetApp.openById("[[GOOGLE FILE ID HERE]]").getSheetByName("ParameterDatabase"); 

  // Check for duplicate IDs in column F of the destination sheet
  var destinationValues = destinationSheet.getRange("C2:C" + destinationSheet.getLastRow()).getValues().flat().filter(Boolean);

  // Check for duplicates
  var duplicates = [];
  allGeneratedIds.forEach(function(row) {
    if (destinationValues.indexOf(row[2]) !== -1) {
      duplicates.push(row[2]);
    }
  });

  if (duplicates.length > 0) {
    // Display an error message if duplicates are found
    var ui = SpreadsheetApp.getUi();
    ui.alert('Error', 'Duplicate IDs found in column H. Cannot store duplicates.', ui.ButtonSet.OK);
    return;
  }

  // Determine the destination range based on the last row and last column
  var lastRow = destinationSheet.getLastRow();
  var destinationRange = destinationSheet.getRange(lastRow + 1, 1, allGeneratedIds.length, numCols);

  // Set the values in the corresponding columns in the destination sheet
  destinationRange.setValues(allGeneratedIds);

    // Fetch all rows from activeSheet
  var startRowSam = 2;
  var startCol2 = 1;
  var lastRowSam = activeSheet.getLastRow();
  var numCols2 = 10;

  var generatedIdsRange2 = activeSheet.getRange(startRowSam, startCol2, lastRowSam, numCols2);
  var allGeneratedIds2 = generatedIdsRange2.getValues();

  // Filter out rows with empty or "NA" IDs in the third column
  allGeneratedIds2 = allGeneratedIds2.filter(row => row[2] !== "" && row[2] !== "NA");

  // Log the filtered data to verify
  Logger.log("Filtered allGeneratedIds2: " + JSON.stringify(allGeneratedIds2));

  // Open the destination sheet
  var destinationSheet2 = SpreadsheetApp.openById("[[GOOGLE FILE ID HERE]]").getSheetByName("Sample Tracker");

  // Check for duplicate IDs in column C of the destination sheet
  var destinationRange2 = destinationSheet2.getRange("C4:C" + destinationSheet2.getLastRow());
  var destinationValues2 = destinationRange2.getValues().flat();

  // Filter out empty or falsy values
  destinationValues2 = destinationValues2.filter(Boolean);

  // Log the destination values to verify
  Logger.log("Destination values in column C: " + JSON.stringify(destinationValues2));

  // Check for duplicates
  var duplicates2 = [];
  allGeneratedIds2.forEach(function(row) {
    var idToCheck2 = row[2].toString(); // Ensure row[2] is treated consistently as string
    if (destinationValues2.indexOf(idToCheck2) !== -1) {
      duplicates2.push(idToCheck2);
    }
  });

  // Log duplicates to verify
  Logger.log("Duplicates found: " + JSON.stringify(duplicates2).length);

  if (duplicates2.length > 0) {
    // Display an error message if duplicates are found
    var ui = SpreadsheetApp.getUi();
    ui.alert('Error', 'Duplicate IDs found in column C. Cannot store duplicates in Sampling.', ui.ButtonSet.OK);
    return;
  }

  // Determine the destination range based on the last row and last column
  var lastRow2 = destinationSheet2.getLastRow();
  var destinationRange2 = destinationSheet2.getRange(lastRow2 + 1, 1, allGeneratedIds2.length, numCols2);

  // Set the values in the corresponding columns in the destination sheet
  destinationRange2.setValues(allGeneratedIds2);
  
 // New Section: Update Column H in HiddenSheetConnectedtoOneResponses
  var hiddenSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HiddenSheetConnectedtoOneResponses");
  var hiddenDataRange = hiddenSheet.getRange("A2:H" + hiddenSheet.getLastRow());
  var hiddenData = hiddenDataRange.getValues();

  var updatedColumnH = hiddenData.map(function(row) {
    var hiddenId = row[6]; // Column G (seventh column, index 6)
    if (hiddenId !== "" && hiddenId !== "NA") {
      var idMatch = allGeneratedIds.some(function(idRow) {
        return hiddenId === idRow[2] && row[7] === "N"; // Column H (eighth column, index 7)
      });
      if (idMatch) {
        row[7] = "Y";
      }
    }
    return row;
  });

  // Write only Column H back to the sheet
  hiddenSheet.getRange(2, 8, updatedColumnH.length).setValues(updatedColumnH.map(row => [row[7]]));
}

/*function generateLabels2() {
  var ss = SpreadsheetApp.openById("[[GOOGLE FILE ID HERE]]");
  var sourceSheet = ss.getSheetByName('SampleIDs for Storage and Labels'); 
  var targetSheet = ss.getSheetByName('Labels'); 

  var sourceData = sourceSheet.getDataRange().getValues();
  var targetData = [];
  var row = 0; // Initialize row index

  for (var i = 1; i < sourceData.length; i++) {
    // Check if both columns A:F have values
    if (sourceData[i][5] !== "" && sourceData[i][6] !== "" && sourceData[i][7] !== "" && sourceData[i][0] !== "" && sourceData[i][1] !== "" && sourceData[i[2]] !== "" && sourceData[i][3] !== "" && sourceData[i][4] !== "" && sourceData[i][8] !== "") {
      var label = [
        '\nID: ' + sourceData[i][5] + '\nNumID: '+ sourceData[i][9] + '\nReplicate: ' + sourceData[i][6] + '  Depth: ' + sourceData[i][7] + '\nPI Name: ' + sourceData[i][0] + '  Lake Code: '+ sourceData[i][1]+ '\nDate: ' + sourceData[i][2] + '-' + sourceData[i][3] + '-' + sourceData[i][4] + '\nParameter: ' + sourceData[i][11] + '  Vol: ' + sourceData[i][13] + '\n' // Concatenate values with new lines
      ];

      // Determine the column index (0, 1, or 2)
      var column = (i - 1) % 3;

      // If column is 0, push a new empty array to targetData for a new row
      if (column === 0) {
        targetData.push([]);
        row++;
      }

      // Set the label in the targetData array at the appropriate row and column
      targetData[row - 1][column] = label[0];
    }
  }

  // Check if the last row is incomplete and fill with blank labels
  if (targetData.length > 0) {
    var lastRow = targetData[targetData.length - 1];
    while (lastRow.length < 3) {
      lastRow.push('\nSample ID:\n\nReplicate:_______Depth #:_______\n\nPI Name:' + sourceData[0][0] + '  Lake Code:' + sourceData[0][1] + '\n\nDate:______-____-____\n\nParameter:________\n\n');
    }
  }

  // Clear only the content in the target range, keeping formatting
  targetSheet.getRange(1, 1, targetSheet.getLastRow()+1, targetSheet.getLastColumn()+1).clearContent();

  // Set the new values to the target range
  targetSheet.getRange(1, 1, targetData.length, targetData[0].length).setValues(targetData);
} */

function stringToHash(str) {
  var hash = 0;
  if (str.length == 0) return hash;
  for (var i = 0; i < str.length; i++) {
    var char = str.charCodeAt(i);
    hash = (hash << 5) - hash + char;
    hash = hash & hash; // Convert to 32bit integer
  }
  return hash;
}
