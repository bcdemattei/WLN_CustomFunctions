function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Store Field Data', 'storeFieldData2')
      .addItem('Store Lab Data', 'storeLabData')
      .addToUi();
}

function storeFieldData2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var datasheet = ss.getSheetByName('Field Data Entry Form');
  var FieldMeta = ss.getSheetByName('Field Sampling Metadata');
  
  var existingValues = FieldMeta.getRange("D2:D").getValues();
  var newValue = datasheet.getRange("B2").getValue(); // Value to check for duplicate

  // Check if newValue already exists in column D of FieldMeta
  for (var i = 0; i < existingValues.length; i++) {
    if (existingValues[i][0] == newValue) {
      // If a duplicate is found, alert the user and exit the function
      SpreadsheetApp.getUi().alert("Error: Duplicate value found in Field Sampling Metadata.");
      return;
    }
  }

  // MetaData portion of saving
  var variables = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    datea : datasheet.getRange(5,5), //will need to insert day
    eid : datasheet.getRange(2,2),
    tin : datasheet.getRange(6,5),
    tout : datasheet.getRange(7,5),
    sdep : datasheet.getRange(7,3),
    lat : datasheet.getRange(8,3),
    lon : datasheet.getRange(8,5),
    wcond : datasheet.getRange(10,2),
    sitenot : datasheet.getRange(14,2),
    snowper : datasheet.getRange(20,2),
    snowdep : datasheet.getRange(20,4),
    icethi : datasheet.getRange(20,6),
    snowobs : datasheet.getRange(21,2),
    iceobs : datasheet.getRange(24,2),
    lsa : datasheet.getRange(30,2),
    lsu : datasheet.getRange(30,3),
    lsh : datasheet.getRange(30,4),
    lso : datasheet.getRange(30,5),
    lsoh : datasheet.getRange(30,6),
    lst : datasheet.getRange(30,7),
    lsth : datasheet.getRange(30,8),
    lsr : datasheet.getRange(30,9),
    lna : datasheet.getRange(31,2),
    lnu : datasheet.getRange(31,3),
    lnh : datasheet.getRange(31,4),
    lno : datasheet.getRange(31,5),
    lnoh : datasheet.getRange(31,6),
    lnt : datasheet.getRange(31,7),
    lnth : datasheet.getRange(31,8),
    lnr : datasheet.getRange(31,9),
    lnote : datasheet.getRange(32,2),
    ctdyn : datasheet.getRange(37,2),
    ctdti : datasheet.getRange(37,4),
    secc : datasheet.getRange(37,6),
    watnot : datasheet.getRange(44,2),
    planknet : datasheet.getRange(49,3), //changing this to fine zoop
    plankmsh : datasheet.getRange(49,5), //changing this to fine zoop
    planknote : datasheet.getRange(54,2), //changing this to fine zoop
    cplanknet : datasheet.getRange(59,3), //coarse zoop
    cplankmsh : datasheet.getRange(59,5),
    cplanknote : datasheet.getRange(64,2),  
    sianet : datasheet.getRange(70,3), //This will change with the addition of a coarse zoop
    siamesh : datasheet.getRange(70,5), //This will change with the addition of a coarse zoop
    sianot: datasheet.getRange(74,2), //This will change with the addition of a coarse zoop
    daten : datasheet.getRange(3,4)
  };

  var lastColumn = 1; // Start from column A
  var rowFieldMeta = FieldMeta.getLastRow() + 1; // Increment the last row index
  
  var numberOfVariables = Object.keys(variables).length;
  for (var i = 0; i < numberOfVariables; i++) {
    var variableName = Object.keys(variables)[i];
    var value = variables[variableName].getValue();
    // Check if the range is empty, assign "NA" if empty
    if (!value) {
      value = "NA";
    }
    FieldMeta.getRange(rowFieldMeta, lastColumn).setValue(value);
    lastColumn++; // Increment the column index for the next variable
  }

  //Parameter
  var id = datasheet.getRange("B2").getValue();

  var vars2 = {
    daa1 : datasheet.getRange(5,5),
    daa2 : datasheet.getRange(5,5),
    daa3 : datasheet.getRange(5,5),
    wda1 : datasheet.getRange(41,2),
    wda2 : datasheet.getRange(42,2),
    wda3 : datasheet.getRange(43,2),
    wla1 : datasheet.getRange(41,3),
    wla2 : datasheet.getRange(42,3),
    wla3 : datasheet.getRange(43,3),
    zda1 : datasheet.getRange(51,2),
    zda2 : datasheet.getRange(52,2),
    zda3 : datasheet.getRange(53,2),
    zta1 : datasheet.getRange(51,3),
    zta2 : datasheet.getRange(52,3),
    zta3 : datasheet.getRange(53,3),//end of fine zoop
    czda1 : datasheet.getRange(61,2),
    czda2 : datasheet.getRange(62,2),
    czda3 : datasheet.getRange(63,2),
    czta1 : datasheet.getRange(61,3),
    czta2 : datasheet.getRange(62,3),
    czta3 : datasheet.getRange(63,3),//end of coarse zoop
    sda1 : datasheet.getRange(72,2),//SIA start
    sda2 : datasheet.getRange(73,2),
    sda3 : datasheet.getRange(69,2),
    sta1 : datasheet.getRange(72,3),
    sta2 : datasheet.getRange(73,3),
    sta3 : datasheet.getRange(69,3)
  }

  // Find the next available row for the specified ID
  var dataRange = ParameterData.getDataRange();
  var dataValues = dataRange.getValues();
  var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column C
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 9; c <= 17; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data . Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 9; // Column I
  for (var variableName in vars2) {
      var value = vars2[variableName].getValue();
      // Check if the range is empty, assign "NA" if empty
      if (!value) {
        value = "NA";
      }
      // Store the value in the next available row and column
      ParameterData.getRange(rowIndex, columnIndex).setValue(value);
      ParameterData.getRange(rowIndex + 3, columnIndex).setValue(value);
      // Move onto the next column after 3 variables
      if (variableName.slice(-1) === '3') {
        columnIndex++;
        if (dataValues[i][2] == id){
          rowIndex = i;
        };
      }
      rowIndex++;
    }
    
  // Clear the values in the ranges called in vars2
  for (var variableName in vars2) {
    vars2[variableName].clearContent();
  }

  // Clear the values in the ranges called in variables except for specific cells
  var rangesToClear = [];
  for (var variableName in variables) {
    if (variableName !== "eid" && // Exclude cell B2
        variableName !== "lc" && // Exclude cell C5
        variableName !== "piname") { // Exclude cell C6
      rangesToClear.push(variables[variableName]);
    }
  }

  for (var j = 0; j < rangesToClear.length; j++) {
    rangesToClear[j].clearContent(); // Clear content of each range
  }
}

function storeLabData(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Laboratory Data Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var LabMeta = ss.getSheetByName('Lab Sampling Metadata');

  var existingValues = LabMeta.getRange("D2:D").getValues();
  var newValue = datasheet.getRange("B2").getValue(); // Value to check for duplicate

  // Check if newValue already exists in column D of FieldMeta
  for (var i = 0; i < existingValues.length; i++) {
    if (existingValues[i][0] == newValue) {
      // If a duplicate is found, alert the user and exit the function
      SpreadsheetApp.getUi().alert("Error: Duplicate value found in Lab Sampling Metadata. Please check and if this error continues contact the data manager.");
      return;
    }
  }

  var variables = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    stti : datasheet.getRange(6,5),
    enti : datasheet.getRange(7,5),
    gencom : datasheet.getRange(9,2),
    tpa1 : datasheet.getRange(14,3),
    tpa2 : datasheet.getRange(14,4),
    tpa3 : datasheet.getRange(14,5),
    tpb1 : datasheet.getRange(15,3),
    tpb2 : datasheet.getRange(15,4),
    tpb3 : datasheet.getRange(15,5),
    tpcom : datasheet.getRange(16,2),
    tnha1 : datasheet.getRange(20,3),
    tnha2 : datasheet.getRange(20,4),
    tnha3 : datasheet.getRange(20,5),
    tnhb1 : datasheet.getRange(21,3),
    tnhb2 : datasheet.getRange(21,4),
    tnhb3 : datasheet.getRange(21,5),
    tnhcom : datasheet.getRange(22,2),
    doca1 : datasheet.getRange(26,3),
    doca2 : datasheet.getRange(26,4),
    doca3 : datasheet.getRange(26,5),
    docb1 : datasheet.getRange(27,3),
    docb2 : datasheet.getRange(27,4),
    docb3 : datasheet.getRange(27,5),
    doccom : datasheet.getRange(28,2),
    cola1 : datasheet.getRange(32,3),
    cola2 : datasheet.getRange(32,4),
    cola3 : datasheet.getRange(32,5),
    colb1 : datasheet.getRange(33,3),
    colb2 : datasheet.getRange(33,4),
    colb3 : datasheet.getRange(33,5),
    colcom : datasheet.getRange(34,2),
    pcncom : datasheet.getRange(40,2),
    sfacom : datasheet.getRange(46,2),
    chlacom : datasheet.getRange(52,2),
    mca1 : datasheet.getRange(59,3),
    mca2 : datasheet.getRange(59,4),
    mca3 : datasheet.getRange(59,5),
    mcb1 : datasheet.getRange(60,3),
    mcb2 : datasheet.getRange(60,4),
    mcb3 : datasheet.getRange(60,5),
    mccom : datasheet.getRange(61,2),
    mcccom : datasheet.getRange(67,2),
    phyta1 : datasheet.getRange(71,3),
    phyta2 : datasheet.getRange(71,4),
    phyta3 : datasheet.getRange(71,5),
    phytb1 : datasheet.getRange(72,3),
    phytb2 : datasheet.getRange(72,4),
    phytb3 : datasheet.getRange(72,5),
    phytcom : datasheet.getRange(73,2),
    cpaa1 : datasheet.getRange(77,3),
    cpaa2 : datasheet.getRange(77,4),
    cpaa3 : datasheet.getRange(77,5),
    cpab1 : datasheet.getRange(78,3),
    cpab2 : datasheet.getRange(78,4),
    cpab3 : datasheet.getRange(78,5),
    cpacom : datasheet.getRange(79,2),
    flaa1 : datasheet.getRange(83,3),
    flaa2 : datasheet.getRange(83,4),
    flaa3 : datasheet.getRange(83,5),
    flab1 : datasheet.getRange(84,3),
    flab2 : datasheet.getRange(84,4),
    flab3 : datasheet.getRange(84,5),
    flacom : datasheet.getRange(85,2),
    fada1 : datasheet.getRange(89,3),
    fada2 : datasheet.getRange(89,4),
    fada3 : datasheet.getRange(89,5),
    fadb1 : datasheet.getRange(90,3),
    fadb2 : datasheet.getRange(90,4),
    fadb3 : datasheet.getRange(90,5),
    fadcom : datasheet.getRange(91,2),
    eema1 : datasheet.getRange(95, 3),
    eema2 : datasheet.getRange(95, 4),
    emma3 : datasheet.getRange(95, 5),
    emmb1 : datasheet.getRange(96, 3),
    emmb2 : datasheet.getRange(96, 4),
    emmb3 : datasheet.getRange(96, 5),
    emmcom : datasheet.getRange(97, 2),
    bapa1 : datasheet.getRange(101, 3),
    bapa2 : datasheet.getRange(101, 4),
    bapa3 : datasheet.getRange(101, 5),
    bapb1 : datasheet.getRange(102, 3),
    bapb2 : datasheet.getRange(102, 4),
    bapb3 : datasheet.getRange(102, 5),
    bapcom : datasheet.getRange(103, 2),
    datent : datasheet.getRange(3,4)
  }

  var lastColumn = 1; // Start from column A
  var rowLabMeta = LabMeta.getLastRow() + 1;
  
  var numberOfVariables = Object.keys(variables).length;
  for (var i = 0; i < numberOfVariables; i++) {
    var variableName = Object.keys(variables)[i];
    var value = variables[variableName].getValue();
    // Check if the range is empty, assign "NA" if empty
    if (!value) {
      value = "NA";
    }
    LabMeta.getRange(rowLabMeta, lastColumn).setValue(value);
    lastColumn++; // Increment the column index for the next variable
  }

 var id = datasheet.getRange("B2").getValue();

 var vars2 = {
  pcna1 : datasheet.getRange(38,3),
  pcna2 : datasheet.getRange(38,4),
  pcna3 : datasheet.getRange(38,5),
  pcnb1 : datasheet.getRange(39,3),
  pcnb2 : datasheet.getRange(39,4),
  pcnb3 : datasheet.getRange(39,5),
  sfaa1 : datasheet.getRange(44,3),
  sfaa2 : datasheet.getRange(44,4),
  sfaa3 : datasheet.getRange(44,5),
  sfab1 : datasheet.getRange(45,3),
  sfab2 : datasheet.getRange(45,4),
  sfab3 : datasheet.getRange(45,5),
  chla1 : datasheet.getRange(50,3),
  chla2 : datasheet.getRange(50,4),
  chla3 : datasheet.getRange(50,5),
  chlb1 : datasheet.getRange(51,3),
  chlb2 : datasheet.getRange(51,4),
  chlb3 : datasheet.getRange(51,5),
  mcca1 : datasheet.getRange(65,3),
  mcca2 : datasheet.getRange(65,4),
  mcca3 : datasheet.getRange(65,5),
  mccb1 : datasheet.getRange(66,3),
  mccb2 : datasheet.getRange(66,4),
  mccb3 : datasheet.getRange(66,5)
 }

  var dataRange = ParameterData.getDataRange();
  var dataValues = dataRange.getValues();
  var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column C
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 18; c <= 23; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 18; // Column R
  for (var variableName in vars2) {
    if (variableName.slice(-2, -1) === 'a'){
      var value = vars2[variableName].getValue();
      // Check if the range is empty, assign "NA" if empty
      if (!value) {
        value = "NA";
      }
      // Store the value in the next available row and column
      ParameterData.getRange(rowIndex, columnIndex).setValue(value);
      // Move onto the next column after 3 variables
      if (variableName.slice(-1) === '3') {
        if (dataValues[i][2] == id){
          rowIndex = i;
        };
      }
    }
    if (variableName.slice(-2,-1) === 'b'){
      var value = vars2[variableName].getValue();
      // Check if the range is empty, assign "NA" if empty
      if (!value) {
        value = "NA";
      }
      // Store the value in the next available row and column
      ParameterData.getRange(rowIndex+3, columnIndex).setValue(value);
      // Move onto the next column after 3 variables
      if (variableName.slice(-1) === '3') {
        columnIndex++;
        if (dataValues[i][2] == id){
          rowIndex = i;
        };
      }
    }
    rowIndex++;
  }

    
    // Clear the values in the ranges called in vars2
    for (var variableName in vars2) {
      vars2[variableName].clearContent();
    }

    // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in variables) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date") { // Exclude cell E5
        rangesToClear.push(variables[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }

}

function scrollToValue() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  // Prompt the user to enter the value
  var response = ui.prompt('Enter the value to scroll to:');
  var editedValue = response.getResponseText().trim(); // Trim spaces

  if (editedValue == "") {
    ui.alert("No value entered!");
    return;
  }

  // Define the range to search for the value
  var searchRange = sheet.getRange("I4:I"); // Changed to search in column I from row 4 to the end
  var values = searchRange.getValues();
  
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === editedValue) { // Strict comparison for exact match
      var targetRow = i + 4; // Adjust for starting from row 4
      sheet.setActiveRange(sheet.getRange(targetRow, 9)); // Column I is the 9th column
      return;
    }
  }

  ui.alert("Value not found in the specified range.");
}

//Post-Processing Entry Functions
///Water Chemistry
function enterTP() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    tpa1 : datasheet.getRange(11,3), //First variable to be stored
    tpa2 : datasheet.getRange(11,4),
    tpa3 : datasheet.getRange(11,5),
    tpb1 : datasheet.getRange(12,3),
    tpb2 : datasheet.getRange(12,4),
    tpb3 : datasheet.getRange(12,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 22; c < 23; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 22; // Column V
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterTN() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    tna1 : datasheet.getRange(17,3),
    tna2 : datasheet.getRange(17,4),
    tna3 : datasheet.getRange(17,5),
    tnb1 : datasheet.getRange(18,3),
    tnb2 : datasheet.getRange(18,4),
    tnb3 : datasheet.getRange(18,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 23; c < 24; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 23; // Column W
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterTDP() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    tda1 : datasheet.getRange(23,3),
    tda2 : datasheet.getRange(23,4),
    tda3 : datasheet.getRange(23,5),
    tdb1 : datasheet.getRange(24,3),
    tdb2 : datasheet.getRange(24,4),
    tdb3 : datasheet.getRange(24,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 24; c < 25; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 24; // Column X
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterNOX() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    oxa1 : datasheet.getRange(29,3),
    oxa2 : datasheet.getRange(29,4),
    oxa3 : datasheet.getRange(29,5),
    oxb1 : datasheet.getRange(30,3),
    oxb2 : datasheet.getRange(30,4),
    oxb3 : datasheet.getRange(30,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 25; c < 26; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 25; // Column Y
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterNH4() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    nha1 : datasheet.getRange(35,3),
    nha2 : datasheet.getRange(35,4),
    nha3 : datasheet.getRange(35,5),
    nhb1 : datasheet.getRange(36,3),
    nhb2 : datasheet.getRange(36,4),
    nhb3 : datasheet.getRange(36,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 26; c < 27; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 26; // Column Z
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterDOC() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    dca1 : datasheet.getRange(41,3),
    dca2 : datasheet.getRange(41,4),
    dca3 : datasheet.getRange(41,5),
    dcb1 : datasheet.getRange(42,3),
    dcb2 : datasheet.getRange(42,4),
    dcb3 : datasheet.getRange(42,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 27; c < 28; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 27; // Column AA
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterColor() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    cla1 : datasheet.getRange(47,3),
    cla2 : datasheet.getRange(47,4),
    cla3 : datasheet.getRange(47,5),
    clb1 : datasheet.getRange(48,3),
    clb2 : datasheet.getRange(48,4),
    clb3 : datasheet.getRange(48,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 28; c < 29; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 28; // Column AB
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterPOC() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    pca1 : datasheet.getRange(53,3),
    pca2 : datasheet.getRange(53,4),
    pca3 : datasheet.getRange(53,5),
    pcb1 : datasheet.getRange(54,3),
    pcb2 : datasheet.getRange(54,4),
    pcb3 : datasheet.getRange(54,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 29; c < 30; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 29; // Column AC
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterPON() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    pna1 : datasheet.getRange(59,3),
    pna2 : datasheet.getRange(59,4),
    pna3 : datasheet.getRange(59,5),
    pnb1 : datasheet.getRange(60,3),
    pnb2 : datasheet.getRange(60,4),
    pnb3 : datasheet.getRange(60,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 30; c < 31; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 30; // Column AD
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterSFA() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    sfa1 : datasheet.getRange(65,3),
    sfa2 : datasheet.getRange(65,4),
    sfa3 : datasheet.getRange(65,5),
    sfb1 : datasheet.getRange(66,3),
    sfb2 : datasheet.getRange(66,4),
    sfb3 : datasheet.getRange(66,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 31; c < 32; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 31; // Column AE
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterChl() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    cha1 : datasheet.getRange(71,3),
    cha2 : datasheet.getRange(71,4),
    cha3 : datasheet.getRange(71,5),
    chb1 : datasheet.getRange(72,3),
    chb2 : datasheet.getRange(72,4),
    chb3 : datasheet.getRange(72,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 32; c < 33; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 32; // Column AF
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}
//Plankton Start
function enterMiCo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    mca1 : datasheet.getRange(78,3),
    mca2 : datasheet.getRange(78,4),
    mca3 : datasheet.getRange(78,5),
    mcb1 : datasheet.getRange(79,3),
    mcb2 : datasheet.getRange(79,4),
    mcb3 : datasheet.getRange(79,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 33; c < 34; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 33; // Column AG
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterMCC() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    mda1 : datasheet.getRange(84,3),
    mda2 : datasheet.getRange(84,4),
    mda3 : datasheet.getRange(84,5),
    mdb1 : datasheet.getRange(85,3),
    mdb2 : datasheet.getRange(85,4),
    mdb3 : datasheet.getRange(85,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 34; c < 35; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 34; // Column AH
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterPhyto() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    paa1 : datasheet.getRange(90,3),
    paa2 : datasheet.getRange(90,4),
    paa3 : datasheet.getRange(90,5),
    pab1 : datasheet.getRange(91,3),
    pab2 : datasheet.getRange(91,4),
    pab3 : datasheet.getRange(91,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 35; c < 36; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 35; // Column AI
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterCPA() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    cpa1 : datasheet.getRange(96,3),
    cpa2 : datasheet.getRange(96,4),
    cpa3 : datasheet.getRange(96,5),
    cpb1 : datasheet.getRange(97,3),
    cpb2 : datasheet.getRange(97,4),
    cpb3 : datasheet.getRange(97,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 36; c < 37; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 36; // Column AJ
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterFlagA() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    faa1 : datasheet.getRange(102,3),
    faa2 : datasheet.getRange(102,4),
    faa3 : datasheet.getRange(102,5),
    fab1 : datasheet.getRange(102,3),
    fab2 : datasheet.getRange(102,4),
    fab3 : datasheet.getRange(102,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 37; c < 38; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 37; // Column AK
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterZFA() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    ada1 : datasheet.getRange(108,3),
    ada2 : datasheet.getRange(108,4),
    ada3 : datasheet.getRange(108,5),
    adb1 : datasheet.getRange(109,3),
    adb2 : datasheet.getRange(109,4),
    adb3 : datasheet.getRange(109,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 38; c < 39; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 38; // Column AL
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterBP() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    bpa1 : datasheet.getRange(114,3), //First variable to be stored
    bpa2 : datasheet.getRange(114,4),
    bpa3 : datasheet.getRange(114,5),
    bpb1 : datasheet.getRange(115,3),
    bpb2 : datasheet.getRange(115,4),
    bpb3 : datasheet.getRange(115,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 39; c < 40; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 40; // Column AM
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterEEMS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    eea1 : datasheet.getRange(120,3), //First variable to be stored
    eea2 : datasheet.getRange(120,4),
    eea3 : datasheet.getRange(120,5),
    eeb1 : datasheet.getRange(121,3),
    eeb2 : datasheet.getRange(121,4),
    eeb3 : datasheet.getRange(121,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 40; c < 41; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 40; // Column AN
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterCZoop() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    zoa1 : datasheet.getRange(126,3), //First variable to be stored
    zoa2 : datasheet.getRange(76,8),
    zoa3 : datasheet.getRange(126,5),
    zob1 : datasheet.getRange(127,3),
    zob2 : datasheet.getRange(76,8),
    zob3 : datasheet.getRange(127,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 41; c < 42; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 41; // Column AO
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}

function enterFZoop() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var datasheet = ss.getSheetByName('Sample Post-Processing Entry Form');
  var ParameterData = ss.getSheetByName('ParameterDatabase');
  var id = datasheet.getRange("B2").getValue();

  var tpnm = {
    piname : datasheet.getRange(5,3),
    lc : datasheet.getRange(6,3),
    date : datasheet.getRange(5,5),
    eid : datasheet.getRange(2,2),
    numid : datasheet.getRange(6,5),
    fza1 : datasheet.getRange(132,3), //First variable to be stored
    fza2 : datasheet.getRange(76,8),
    fza3 : datasheet.getRange(132,5),
    fzb1 : datasheet.getRange(133,3),
    fzb2 : datasheet.getRange(76,8),
    fzb3 : datasheet.getRange(133,5)
  }
    var dataRange = ParameterData.getDataRange();
    var dataValues = dataRange.getValues();
    var rowIndex;

  for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][2] == id) { // Assuming ID is in column F
      rowIndex = i + 1; // Add 1 to convert from 0-based index to 1-based row number
      break;
    }
  }
  if (!rowIndex) { // If ID is not found, start from the next empty row
    rowIndex = dataValues.length + 1;
  }

     // Check if the rows already contain data in the specific cells
  var rowContainsData = false;
  for (var r = rowIndex; r < rowIndex + 3; r++) {
    for (var c = 42; c < 43; c++) {
      var cellValue = ParameterData.getRange(r, c).getValue();
      if (cellValue !== "") {
        rowContainsData = true;
        break;
      }
    }
    if (rowContainsData) {
      break;
    }
  }

  // If the row already contains data in the specific cells, alert the user
  if (rowContainsData) {
    var response = SpreadsheetApp.getUi().alert(
      "Row Already Contains Data",
      "The row already contains data. Do you want to proceed and overwrite the existing data?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    // If the user clicks "No" or closes the dialog, exit the script
    if (response != SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
     
  var columnIndex = 42; // Column AP
  var aRowIndex = rowIndex; // Separate row index for 'a' variables
  var bRowIndex = rowIndex + 3; // Separate row index for 'b' variables

  for (var variableName in tpnm) {
    var value = tpnm[variableName].getValue();
    if (!value) {
      value = "NA";
    }

    if (variableName.slice(-2, -1) === 'a') {
      ParameterData.getRange(aRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        aRowIndex = rowIndex; // Reset to the starting row for 'a' variables
      } else {
        aRowIndex++;
      }
    } else if (variableName.slice(-2, -1) === 'b') {
      ParameterData.getRange(bRowIndex, columnIndex).setValue(value);
      if (variableName.slice(-1) === '3') {
        bRowIndex = rowIndex + 3; // Reset to the starting row for 'b' variables
        columnIndex++;
      } else {
        bRowIndex++;
      }
    }
  }
 // Clear the values in the ranges called in variables except for specific cells
    var rangesToClear = [];
    for (var variableName in tpnm) {
      if (variableName !== "eid" && // Exclude cell B2
          variableName !== "lc" && // Exclude cell C5
          variableName !== "piname" && // Exclude cell C6
          variableName !== "date" && // Exclude cell E5
          variableName !== "numid") { // Exclude cell E6
        rangesToClear.push(tpnm[variableName]);
      }
    }

    for (var j = 0; j < rangesToClear.length; j++) {
      rangesToClear[j].clearContent(); // Clear content of each range
    }


}
