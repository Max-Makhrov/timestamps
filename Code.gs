// Sample file for copy: https://docs.google.com/spreadsheets/d/1rV60zGQqIoDedp2wpGrTyEY4xcuXil_gTeIcyguMjAE/copy

var C_ONEDIT_SETS = 
    {
      ///////////// Change \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      
      
      sheetName: ["Sales", "Sales", "Sales", "Sales2", "Sales2"],
      firstRow: ["2", "2", "2", "2", "2"],
      timeStampColumn: ["A", "B", "C", "E", "B"],
      changeColumns: ["D", "H", "D:G;I:", "B:D", "B"],
      triggerType: ["unique number", "create", "update", "each update", "each update"],
      insertType: ["value", "value", "value", "value", "note"],
      lastRow: ["-1", "-1", "-1", "-1", "-1"]
      
      
      ///////////// Change /////////////////////////////////////  
    };


//  _______   _                           
// |__   __| (_)                          
//    | |_ __ _  __ _  __ _  ___ _ __ ___ 
//    | | '__| |/ _` |/ _` |/ _ \ '__/ __|
//    | | |  | | (_| | (_| |  __/ |  \__ \
//    |_|_|  |_|\__, |\__, |\___|_|  |___/
//               __/ | __/ |              
//              |___/ |___/               
// Regerence: https://developers.google.com/apps-script/guides/triggers
// 
// function to test onEdit
function test_onEdit()
{
  var f = SpreadsheetApp.getActive();
  var ee = 
      [
        { range: f.getSheetByName('Sales').getRange('C9:E10') }
        ,{ range: f.getSheetByName('Sales').getRange('H4') }       
        ,{ range: f.getSheetByName('Sales').getRange('D5:D13') }     
        ,{ range: f.getSheetByName('Sales2').getRange('C4:C5') }
        ,{ range: f.getSheetByName('Sales2').getRange('B9:B10') } 
      ];
  for (var i = 0; i < ee.length; i++)
  {
    Logger.log(' ------------------------------- sample No = ' + (i + 1) + ' ------------------------------');      
    var e = ee[i];
    e.value = 'd';
    e.testMode = true;
    var t0 = new Date();
    onEdit(e);
    Logger.log(' ------------------------------- time to run onEdit = ' + (new Date() - t0) + ' ms -----------------');      
  }
  return 0;  
}
// Simple trigger, reserved function **onEdit**. 
// Caution:
//   onEdit(e) is a reserved name for a function. 
//   It is run automatically when the user changes the sheet. 
//   It passes the parameter into the function with notes about the user’s action.
function onEdit(e)
{
  makeTimeStampsRuner_(e); // add more functions with onEdit if needed    
}
// Simple trigger, reserved function **onOpen**. Works when the file is refreshed or opened
function onOpen(e)
{
  createCustomMenu_(); // remove or replace this function when / if not needed
}





//    _____               
//   / ____|              
//  | |     ___  _ __ ___ 
//  | |    / _ \| '__/ _ \
//  | |___| (_) | | |  __/
//   \_____\___/|_|  \___|
//
// Handles all stamps in settings
var C_CORE_STAMP_SETTINGS = 
    {
      a1addresses: {
        delim: ';',
        rangeDelim: ':',
      },
      triggerTypes: {
        // markers = array [ ... , ... ]
        //   [0] checkEmpty:    1 to stop script if all changed values in a row are empty
        //   [1] readStamp:     1 if need to read cell of timestamp or note with timestamp
        //   [2] noOverwrite:   1 if stop execution in case non-empty stamp value
        //   [3] isCumulative:  1 when responses are added to each other
        //   [4] isUniqueNum:   1 if returned type is unique number
        "unique number": {markers: [1, 1, 1, 0, 1, 1]}, 
        "create":        {markers: [1, 1, 1, 0, 0, 0]}, 
        "update":        {markers: [0, 0, 0, 0, 0, 0]}, 
        "each update":   {markers: [0, 1, 0, 1, 0, 0], timeFormat: 'yyyy-MM-dd HH:mm:ss', delimiter: ', '}
      },
      insertTypes: {
        // markers = array [ ... , ... ]
        //   [0] checkIntersect: 1 to stop script if chandes range intersects timestamp range
        //   [1] isNote:         1 if note
        "value": {markers: [1, 0]},
        "note":  {markers: [0, 1], timeFormat: 'yyyy-MM-dd HH:mm:ss', delimiter: '\n'}
      }
    };
function makeTimeStampsRuner_(e)
{  
  // new for this verion:
  //  1) triggerTypes:  [create, update, each update, unique number]
  //  2) insertTypes: [value, note]
  //  3) lastRows:    [500, 0, 500] => hidden functionality, "0" means last row is not counted
  //  + script error handling
  //  + logical error hendling: stamp column intersect change column, all values are empty
  try
  {
    makeTimeStamps_(e);
  }
  catch(err)
  {
    Logger.log(err); 
  }
}
// loop all sets to execute
function makeTimeStamps_(e)
{
  // get data from e object
  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = sheet.getName();
  var sets = C_ONEDIT_SETS;
  // if sheet is among sets
  var indexSets = sets.sheetName.indexOf(sheetName); 
  if (indexSets === -1) { return -1; } // wrong sheet
  // loop sets
  var set = {sheet: sheet, range: range, sheetName: sheetName, value: e.value, oldValue: e.oldValue, testMode: e.testMode};
  // create a loop and try to execute each set
  var res = []; // for tests only
  sets.sheetName.forEach(
    function (sheetNameTrigger, index) {
      if (sheetNameTrigger === sheetName) {
        set.index = index;
        for (var key in sets)
        {
          set[key] = sets[key][index];  // create/rewrite the set      
        }        
        res.push( runTimeStampWithSet_(set) ); 
      }
      else
      {
        res.push(-1); // sheet does not match
      }
    }  
  );  
  // log(res); // leave for test reasons
  return 0;
}
// creates timestamp for one set of rules
function runTimeStampWithSet_(set)
{
  // set
  //   {
  //     "sheet":{},
  //     "range":{},
  //     "sheetName":"Sales",
  //     "value":"d",
  //     "index":0,
  //     "firstRow":3,
  //     "timeStampColumn":"A",
  //     "changeColumns":"D",
  //     "triggerType":"unique number",
  //     "insertType":"value",
  //     "lastRow": -1
  //   }
  // 
  //
  ////////////////////////////////////////////////////////////////////////////////////
  // [ 0 ]. settings checks
  ////////////////////////////////////////////////////////////////////////////////////  
  // check and get trigger type
  var sets = C_CORE_STAMP_SETTINGS; // general settings
  var triggerTypeSets = sets.triggerTypes[set.triggerType];
  if (!triggerTypeSets) { return -2; } // no trigger type in sets
  var insertTypeSets = sets.insertTypes[set.insertType];
  if (!insertTypeSets) { return -3; } // no insert type in sets
  ////////////////////////////////////////////////////////////////////////////////////
  // [ 1 ]. range checks
  ////////////////////////////////////////////////////////////////////////////////////
  // check rows
  var rowEdit = set.range.getRow();
  if (set.hasOwnProperty('firstRow')) {
    if (rowEdit < parseInt(set.firstRow)) { return -4; } // wrong first row
  }
  if (set.hasOwnProperty('lastRow')) {
    if (rowEdit > parseInt(set.lastRow) && parseInt(set.lastRow) > 0) { return -5; } // wrong last row
  }
  // check columns
  var columnsEdit = getTimeStampcolumnsEdit_(set); // [2,3,4]
  var columnsTrigger = getTimeStampColumnsTrigger_(set);
  var triggerColumnsIntersect = isArraysIntersect(columnsEdit, columnsTrigger);
  if (!triggerColumnsIntersect) { return -6; } // edited columns are not set to be triggered   
  // check column stamp intesects changed columns
  var checkIntersect = insertTypeSets.markers[0];
  var columnStamp = numberOfColumnFromText_(set.timeStampColumn); // 3
  if (checkIntersect) {
    if (columnsEdit.indexOf(columnStamp) > -1) { return -7; } // user changes stamp column  
  }
  ////////////////////////////////////////////////////////////////////////////////////
  // [ 2 ]. check if all entered data is empty. get stamps data if needed
  ////////////////////////////////////////////////////////////////////////////////////
  var hEdit = set.range.getHeight();  // edited range size = h
  var wEdit = columnsEdit.length;     // edited range size = w
  // get markers
  var checkEmpty = triggerTypeSets.markers[0];
  var readStamp = triggerTypeSets.markers[1];
  var isNote = insertTypeSets.markers[1];
  // triggerTypeSets
  //   [0] checkEmpty:     1 to stop script if all changed values are empty
  //   [1] readStamp:      1 if need to read cell of timestamp or note with timestamp
  //   [2] noOverwrite:    1 if stop execution in case non-empty stamp value
  //   [3] isCumulative:   1 when responses are added to each other
  // insertTypeSets
  //   [0] checkIntersect: 1 to stop script if chandes range intersects timestamp range
  //   [1] isNote:         1 if note
  var rRead; // range to read
  var dataAll = [];
  var stampValues = []; //  [] if we do not read it
  var dataEdited = [], dataEditedRow = []; // [] if do not read
  if      (isNote === 0 && readStamp === 1 && checkEmpty === 1) {
    Logger.log('read change & stamp cells');
    // decide boundaries
    var colStart = 0, stampIndex = 0, dataIndex = 0;
    var colEditStart =  columnsEdit[0];
    if (colEditStart > columnStamp) { 
      colStart = columnStamp; 
      stampIndex = 0;
      dataIndex = colEditStart - columnStamp;
    }
    else { 
      colStart = colEditStart; 
      stampIndex = columnStamp - colEditStart;
      dataIndex = 0;
    }
    var colEnd = 0;
    var colEditEnd = columnsEdit[columnsEdit.length - 1];
    if (colEditEnd > columnStamp) { 
      colEnd = colEditEnd; }
    else { 
      colEnd = columnStamp; }
    rRead = set.sheet.getRange(rowEdit, colStart, hEdit, colEnd - colStart + 1);
    // >>>>> heavy operation
    // >>>>> in this case cannot ommit this operation even for single-cell entry
    // >>>>> because need to read stamp value, too
    dataAll = rRead.getValues();
    // get stamp data and data edited
    for (var k = 0; k < dataAll.length; k++)
    {
      dataEditedRow = [];
      for (var kk = 0; kk < wEdit; kk++) {
        dataEditedRow.push(dataAll[k][kk + dataIndex]);
      }
      dataEdited.push(dataEditedRow);
      stampValues.push(dataAll[k][stampIndex]);
    }
  }
  else if (isNote === 0 && readStamp === 1 && checkEmpty === 0) {
    Logger.log('read stamp cells');
    rRead = set.sheet.getRange(rowEdit, columnStamp, hEdit);
    // >>>>> heavy operation
    dataAll = rRead.getValues();
    for (var k = 0; k <  wEdit; k++) {
      stampValues.push(dataAll[k][0]);
    }
  }
  else if (checkEmpty === 1)
  {
    Logger.log('read change'); 
    // not read cells if value is known
    if (hEdit === 1 && wEdit === 1 && sets.value) {
      dataAll = [[value]];
    }
    else {
      rRead = set.sheet.getRange(rowEdit, columnsEdit[0], hEdit, wEdit);
      // >>>>> heavy operation
      dataAll = rRead.getValues();      
    }
    dataEdited = dataAll;
  }
  else {
    Logger.log('no read');    
  }
  
  //
  // read note stamp
  if (isNote === 1 && readStamp === 1) {
    Logger.log('read stamp notes');  
    rRead = set.sheet.getRange(rowEdit, columnStamp, hEdit);
    // >>>>> heavy operation
    var notesInfo = rRead.getNotes();
    for (var k = 0; k < notesInfo.length; k++) {
      stampValues.push(notesInfo[k][0]);
    }
  }
  ////////////////////////////////////////////////////////////////////////////////////
  // [ 3 ]. row taks - for individual sets depending on rows
  ////////////////////////////////////////////////////////////////////////////////////
  var timezone = false;
  //    SpreadsheetApp.getActive().getSpreadsheetTimeZone() → will use for cumulative values (works fast)
  var isCumulative = triggerTypeSets.markers[3];
  var timeFormat = triggerTypeSets.timeFormat || insertTypeSets.timeFormat; // take time format from at least 1 set
  var delimiter = insertTypeSets.delimiter || triggerTypeSets.delimiter; // delimiter is insert type priority
  if (isCumulative === 1) { timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone(); }  
  var rowTask = {
    stampValue: false,                       // the current value of stamp for the row,
    editedCells: [],                         // the values user entered, current edit row
    checkEmpty: checkEmpty,                  // stop if all values in edited row are empty
    noOverwrite: triggerTypeSets.markers[2], // if stop execution in case non-empty stamp value
    isCumulative: isCumulative,              // 1 if the result is joined with previous
    delimiter: delimiter,    // for cumulative stamp value: delim stamps
    timeFormat: timeFormat,                  // for cumulative stamp value: format of stamps
    isUniqueNum: triggerTypeSets.markers[4], // 1 if returned type is unique number
    timezone: timezone,
    stamp: new Date()
  }
  // Final notes:
  //    stampValues → the [] array of current stamp values. Source: notes or range. May be empty = no values if not important
  //    dataEdited  → the [[]] array of values -- rows user edited, may be empty if not important
  //    hEdit → the number of rows edited. 
  var data_out = [], nextValue = ''; // the array for all stamp values. Will insert is a single setValues call
  var writeChange = false;
  var isDataEdited = false, isStampValues = false;
  if (dataEdited.length > 0)  { isDataEdited = true; }
  if (stampValues.length > 0) { isStampValues = true; }
  for (var i = 0; i < hEdit; i++)
  {
    rowTask.index = i;
    if (isDataEdited) { rowTask.editedCells = dataEdited[i] }  else { rowTask.editedCells = []; }
    if (isStampValues) { rowTask.stampValue = stampValues[i] } else { rowTask.stampValue = false; }
    nextValue = getTimeStampNextValue_(rowTask);
    if (rowTask.stampValue !== nextValue) { writeChange = true; }
    data_out.push([nextValue]);
  }
  ////////////////////////////////////////////////////////////////////////////////////
  // [ 4 ]. write - update stamp values
  ////////////////////////////////////////////////////////////////////////////////////
  if (set.testMode) {
    Logger.log(data_out);
    Logger.log(writeChange);
    return 0; // do not change in a test mode
  }
  if (!writeChange) { return -8; }
  // use 
  //   columnStamp   → the number of column to unsert stamps
  var rTo = set.sheet.getRange(rowEdit, columnStamp, hEdit);
  if (isNote === 1)
  {
    // wtite to notes
    rTo.setNotes(data_out);
  }
  else
  {
    // write to cells
    rTo.setValues(data_out);
  }
  // return OK status = 0  
  return 0;
}
//
// decide the value to return
function getTimeStampNextValue_(rowTask) {
  var defaultStamp = rowTask.stampValue || ''; // default is stamp value of empty
  // check empty
  if (rowTask.checkEmpty == 1) {
    var textvals = rowTask.editedCells.join('');
    if (textvals === '') { return defaultStamp; } // 
  }
  // check overwrite
  if (rowTask.noOverwrite == 1) {
    if (rowTask.stampValue && rowTask.stampValue !== '') {
      return defaultStamp;
    }   
  }
  // if cumulative
  if (rowTask.isCumulative == 1)
  {
    var newStamp = '' + Utilities.formatDate(rowTask.stamp, rowTask.timezone, rowTask.timeFormat);
    if (defaultStamp === '') { return '' + newStamp; }
    if (typeof defaultStamp.getMonth === 'function') { // if it was date...
      return Utilities.formatDate(defaultStamp, rowTask.timezone, rowTask.timeFormat)
      + rowTask.delimiter + newStamp;
    }
    return defaultStamp + rowTask.delimiter + newStamp; // cumulative value
  }
  // for unique number
  if (rowTask.isUniqueNum) {
    return rowTask.stamp.getTime() + rowTask.index;
  }
  // case if format is defined
  if (rowTask.timeFormat) {
    return Utilities.formatDate(rowTask.stamp, rowTask.timezone, rowTask.timeFormat)
  }
  // basic case
  return rowTask.stamp;
}
//
// returns the list of columns changed by user:
// [1,2,3]
function getTimeStampcolumnsEdit_(set)
{
  var columnChange = set.range.getColumn();
  var widthChange = set.range.getWidth();  
  var lastCol = widthChange - 1 + columnChange;
  var res = [];
  for (var i = columnChange; i <= lastCol; i++)
  {
    res.push(0+i);
  }
  return res;
}
//
// returns the list of columns triggered by script
// [1,2,3]
function getTimeStampColumnsTrigger_(set)
{
  var lastCol = set.sheet.getMaxColumns(); // maximum number of columns in sheet
  var text = set.changeColumns;
  var listOfTriggeredColumns = getColumnsListFromTextNotation_(text, lastCol); // [4, 5, 6, 7, 9, 10]
  return listOfTriggeredColumns; 
}
//
//  text       lasCol              Output
//  H              -               [8]
//  D:G;I:        10               [4, 5, 6, 7, 9, 10]
function getColumnsListFromTextNotation_(text, lastCol)
{
  var sets = C_CORE_STAMP_SETTINGS.a1addresses;
  var delim = sets.delim;
  var texts = text.split(delim);
  var ORes = {}; // object for finding unique keys
  var start, end = false, elts = [];
  for (var i = 0; i < texts.length; i++)
  {
    if (texts[i].substr(texts[i].length - 1) === ':') { end = lastCol; }
    elts = texts[i].split(sets.rangeDelim);
    start = numberOfColumnFromText_(elts[0]);
    if (elts[1]) { end = numberOfColumnFromText_(elts[1]); }
    end = end || start;
    for (var ii = start; ii <= end; ii++)
    {
      ORes[ii] = '';
    }
    end = false;
  }
  // get unique elements and convert them into num
  var res = Object.keys(ORes).map(function(elt) { return parseInt(elt); });
  return res;  
}
// https://stackoverflow.com/a/62906172/5372400
// number of the col from text
function numberOfColumnFromText_(tcol) {
  var a = "A".charCodeAt(0)-1;
  var n = 0, tl = tcol.length;
  var TCOL = tcol.toUpperCase();
  for (var i =0; i < tl; i++)  {        
    n+= Math.pow(26,i);
    n+= (TCOL.charCodeAt(i)-a -1) * Math.pow(26,tl-1-i);  }
  return n;
}  
// check 2 arrays [] [] have common values
// [1, 2, 3] [4, 5, 6] => false
// [1, 2, 3] [4, 5, 1] => true
function isArraysIntersect(arr1, arr2) {
  for (var i = 0; i < arr1.length; i++)
  {
    for (var ii = 0; ii < arr2.length; ii++)
    {
      if (arr1[i] === arr2[ii]) { return true; }
    }    
  }
  return false;
}




//  _____           _        _ _       _   _             
// |_   _|         | |      | | |     | | (_)            
//   | |  _ __  ___| |_ __ _| | | __ _| |_ _  ___  _ __  
//   | | | '_ \/ __| __/ _` | | |/ _` | __| |/ _ \| '_ \ 
//  _| |_| | | \__ \ || (_| | | | (_| | |_| | (_) | | | |
// |_____|_| |_|___/\__\__,_|_|_|\__,_|\__|_|\___/|_| |_|
//  
//  Delete this code if not needed after installation is done
//
// creates a custom menu. 
function createCustomMenu_() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('TimeStamp')
      .addItem('Get the code!', 'logTheCode_')
      .addToUi();  
}
// function for creating settings
function logTheCode_()
{
  var ui = SpreadsheetApp.getUi();
  var html = getTheCode_()
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setHeight(200), 'Paste this code into Tools > Script Editor...');
}
// code generator. Creates code for user to copy into the editor
function getTheCode_() {
  var sheetWithSettings = '_onEditSets_'; // use this constant only once in this function
  var firstRowWithSettings = 2;           // use this constant only once in this function
  var file = SpreadsheetApp.getActive();
  var sheet = file.getSheetByName(sheetWithSettings); 
  var data = sheet.getDataRange().getValues();
  // combine settings
  var myVars = {
    sheetName: {index: 0},
    firstRow: {index: 1},
    timeStampColumn: {index: 2},
    changeColumns: {index: 3},
    triggerType: {index: 4},
    insertType: {index: 5},
    lastRow: {defaults: -1}
  }
  var result = '', results = [];
  for (var elt in myVars) {
    myVars[elt].values = []; // add array for combination of data
  }
  result += '<pre><code>';
  for (var i = firstRowWithSettings-1; i < data.length; i++) {
    for (var ii = 0; ii < data[i].length; ii++) {
      for (var elt in myVars) {
        if (myVars[elt].index === ii)
        {
          myVars[elt].values.push(data[i][ii]);
        }
        else if (ii === 0 && myVars[elt].hasOwnProperty('defaults'))
        {
          myVars[elt].values.push(myVars[elt].defaults);
        } 
      } 
    }
  }
  var results = [];
  for (var elt in myVars) {
    results.push('      ' +  elt + ': ["' + myVars[elt].values.join('", "') + '"]');
  }
  result += results.join(',<br>');
  result += '</code></pre>'
  return result;
}

