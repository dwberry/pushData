var scriptTitle = "pushData V1.4 (11/7/12)";
// Written by Andrew Stillman for New Visions for Public Schools
// Published under GNU General Public License, version 3 (GPL-3.0)
// See restrictions at http://www.opensource.org/licenses/gpl-3.0.html
// Support and contact at http://www.youpd.org/pushdata

var HAMSTERIMAGEURL = "https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/hamsterSync.gif?attachauth=ANoY7coInwbQyw9SsTc07gOK01RkXnQv9Xx0k2mtjXW7KoiGtHYG3mHqZZDv2fFI8SCStyxejKUXivcvLyS6FmMi0rWAb0-GfFmErAIYebtQMZyVXBYd4Pr91b4JVrh25C0nTsYb4Jvpk8AYRdoBxnRTRL8nhL-22zRpHNDE5my1hJIUBgy2vYHmL3WihAxq5cwNGgEtS1Gq8UDNCZRvOQsI4CNiyWfQfLVL61JNi8UHTG1xMSC3Y4JTGhZxYx41RDHXseVqQjMc&attredirects=0";

function onInstall () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [] ;
  menuEntries.push({name: "What is pushData?", functionName: "pushData_whatIs"});
  menuEntries.push({name: "Complete installation", functionName: "onOpen"});
  ss.addMenu("pushData", menuEntries);
}

function pushData_runInstall () {
  //ensure needed triggers are properly installed 
  setPushDataUid();
  pushData_getInstitutionalTrackerObject();
  var triggers = ScriptApp.getScriptTriggers();
  var timeTriggerSet = false;
  var editTriggerSet = false;
  for (var i=0; i<triggers.length; i++) {
    var functionName = triggers[i].getHandlerFunction();
    var triggerSource = triggers[i].getTriggerSource().toString();
    var triggerType = triggers[i].getEventType().toString();
    if ((functionName=="pushData")&&(triggerSource=="CLOCK")&&(triggerType=="CLOCK")) {
      timeTriggerSet = true;
      break;
    }
  }
  if (timeTriggerSet == false) {
    pushData_createTimeTrigger();
  }
  for (var i=0; i<triggers.length; i++) {
    var functionName = triggers[i].getHandlerFunction();
    var triggerSource = triggers[i].getTriggerSource().toString();
    var triggerType = triggers[i].getEventType().toString();
    if ((functionName=="pushData_specialEdit")&&(triggerSource=="SPREADSHEET")&&(triggerType=="ON_EDIT")) {
      editTriggerSet = true;
      break;
    }
  }
  if (editTriggerSet == false) {
    pushData_createEditTrigger();
  }
  ScriptProperties.setProperty("installed", "true");
  onOpen();
}



function onOpen () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [] ;
  var installed = ScriptProperties.getProperty("installed");
  if (installed!="true") {
    menuEntries.push({name: "What is pushData?", functionName: "pushData_whatIs"});
    menuEntries.push({name: "Complete installation", functionName: "pushData_runInstall"});
    ss.addMenu("pushData", menuEntries);
    return;
  }
  menuEntries.push({name: "What is pushData?", functionName: "pushData_whatIs"});
  menuEntries.push({name: "Create new push process", functionName: "pushData_pickerUi"});
  menuEntries.push({name: "Push all (overrides/resets timers)", functionName: "pushDataOverRide"});
  menuEntries.push({name: "Refresh Spreadsheet info", functionName: "pushData_refreshSSKeys"});
  menuEntries.push({name: "Usage tracker settings", functionName: "pushData_institutionalTrackingUi"});
  ss.addMenu("pushData", menuEntries);
  
  //ensure console and readme sheets exist
  var sheets = ss.getSheets();
  var settingsConsoleSet = false;
  var readMeSet = false;
  for (var i=0; i<sheets.length; i++) {
    if (sheets[i].getName()=="Settings Console") {
      settingsConsoleSet = true;
      break;
    }
  }
  for (var i=0; i<sheets.length; i++) {
    if (sheets[i].getName()=="pushData Read Me") {
      readMeSet = true;
      break;
    }
  }
  if (settingsConsoleSet==false) {
    ss.insertSheet("Settings Console");
  }
  if (readMeSet==false) {
    ss.insertSheet("pushData Read Me");
    pushData_setReadMeText();
  }
  //restore headings in case of changes
  pushData_restoreHeadings();
  pushData_refreshSSKeys ();
}


//Function used to create contents of "Read Me" sheet
function pushData_setReadMeText() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("pushData Read Me");
  sheet.insertImage('http://www.youpd.org/sites/default/files/acquia_commons_logo36.png', 1, 1);
  var readMeText = "The hamsterSync script installed on this spreadsheet allows for the time-triggered push of data between spreadsheets in a the same folder in your DocsList, enabling a light-duty database to be built from cloud-based Spreadsheets.  \n \n For example, one might be interested in maintaining a master list of students in a sheet that is used for lookup in as part of the way tallies or triggered emails on forms are done (this is possible using the formMule script: http://www.youpd.org/formmule).  pushData would allow you to push the student list from a single master to multiple destination sheets.  More complex applications are possible, including data warehousing from multiple forms backends back to a single spreadsheet, for example. \n \n Use the \"Settings Console\" sheet to specify the \"Source\" and \"Destination\" spreadsheets, sheet names, and columns to push.   Hovering over each column heading in the \"Settings Console\" will give clues as to how to specify cell values.";
  readMeText += "\n \n A few pointers:\n 1) Column widths are auto-set to be narrow, with no text-wrap in some cases, to allow easier viewing of the whole system.  Double click on cell values with overflow to see the whole value.";
  readMeText += "\n 2) Typing in spreadsheet names and sheet names should trigger the auto-lookup of spreadsheet keys.  Run \"Refresh SSKeys\" if it gets stuck.";
  readMeText += "\n 3) Destination sheets can have formulas to the right of columns that are pushed by this script.";  
  readMeText += "\n 4) Increasing the number of pushed columns will overwrite any existing columns that are to the right of existing pushed values in the destination sheet.";
  readMeText += "\n \n This script was written by Andrew Stillman for New Visions for Public Schools, as an enhancement to school data tracking systems developed at the Academy for Careers in Television and Film.  For support on this script, visit http://www.youpd.org/pushdata";
  sheet.setRowHeight(1, 120);
  sheet.setColumnWidth(1, 800);
  sheet.getRange("A2").setValue(scriptTitle).setFontSize(18);
  sheet.getRange("A3").setValue(readMeText);
}


// Sets an hourly trigger for the pushData function
function pushData_createTimeTrigger() {
  var everyHour = ScriptApp.newTrigger("pushData")
  .timeBased()
  .everyHours(1)
  .create();
}

// Sets on onEdit trigger for the pushData_specialEdit function
// This is used instead of the onEdit function b.c. onEdit is appeared to be unable to call outside the spreadsheet
function pushData_createEditTrigger() {
  var ssKey = SpreadsheetApp.getActiveSpreadsheet().getId();
  var editTrigger = ScriptApp.newTrigger("pushData_specialEdit")
  .forSpreadsheet(ssKey)
  .onEdit()
  .create();
}

// Function called by dropdown menu.  Override variable bypasses all timer checks
function pushDataOverRide () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var override = true;
  var app = UiApp.createApplication().setHeight(220).setWidth(220);
  var panel = app.createVerticalPanel();
  var image = app.createImage(HAMSTERIMAGEURL).setWidth("200px");
  panel.add(image);
  var status = app.createLabel("Pushing data...resetting all timers").setId("status");
  panel.add(status);
  app.add(panel);
  ss.show(app);
  pushData(override);
  return app;
}

// Function used to restore the formats, headers, etc in the console
function pushData_restoreHeadings () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Settings Console");
  sheet.setFrozenRows(1);
  var headerRange = sheet.getRange(1, 1, 1, 11);
  var values = [["Source Spreadsheet Name", "Source Sheet Name", "Source Spreadsheet Key", "Available Columns", "Columns to Push", "Destination Spreadsheet Name", "Destination Sheet Name", "Destination Spreadsheet Key","Columns in Dest Sheet", "Push Interval", "Status Message"]];
  headerRange.setValues(values).setFontSize(8).setWrap(true).setVerticalAlignment("middle").setHorizontalAlignment("center");
  SpreadsheetApp.flush();
  headerRange.setFontStyle("bold");
  sheet.getRange(1, 3).setComment("Do not edit column.  The script will automatically discover the key of the spreadsheet.");
  sheet.getRange(1, 1).setComment("Enter the name of the source spreadsheet exactly as it appears in your DocsList. SS must be in the \"pushData Root\" folder.");
  sheet.getRange(1, 2).setComment("Sheet names are found in the tabs at the bottom of a spreadsheet. Value will default to the top sheet if misspelled or left blank.");
  sheet.getRange(1, 4).setComment("Do not edit column. The script will automatically discover available headers.");
  sheet.getRange(1, 5).setComment("Comma separated list of column headers to push.  BEWARE: If increasing the number of columns, you may overwrite adjacent columns in the destination sheet.");
  sheet.getRange(1, 9).setComment("Do not edit column. Automatically populated upon edit of destination SSName or Sheet Name");
  sheet.getRange(1, 6).setComment("Enter the name of the destination spreadsheet exactly as it appears in your DocsLIst. SS must be in the \"pushData Root\" folder.");
  sheet.getRange(1, 7).setComment("Sheet names are found in the tabs at the bottom of a spreadsheet. Must match the exact spelling and case of the source sheet name.");
  sheet.getRange(1, 8).setComment("Do not edit column. Automatically populated upon edit of destination SSName or Sheet Name");
  sheet.getRange(1, 10).setComment("Set this as a whole number of hours between updates. Defaults to 24 if left blank.");
  sheet.getRange(1, 1, sheet.getLastRow(), 2).setBackground("white").setWrap(false);
  sheet.getRange(1, 3, sheet.getLastRow(), 2).setBackgroundColor("#fff2cc").setWrap(false);
  sheet.getRange(1, 4, sheet.getLastRow(), 1).setFontSize("7").setVerticalAlignment("middle").setHorizontalAlignment("center");
  sheet.getRange(1, 5, sheet.getLastRow(), 1).setFontSize("7").setVerticalAlignment("middle").setHorizontalAlignment("center");
  sheet.getRange(1, 6, sheet.getLastRow(), 2).setWrap(false);
  sheet.getRange(1, 9, sheet.getLastRow(), 1).setFontSize("7").setVerticalAlignment("middle").setHorizontalAlignment("center");
  sheet.getRange(1, 5, sheet.getLastRow(), 1).setBackgroundColor("#ffff00");
  sheet.getRange(1, 8, sheet.getLastRow(), 2).setBackgroundColor("#c9daf8");
  sheet.getRange(1, 10, sheet.getLastRow(), 2).setBackgroundColor("white").setWrap(false);
  
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 70);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 100);
  sheet.setColumnWidth(9, 100);
  sheet.setColumnWidth(10, 50);
  sheet.setColumnWidth(11, 400);
  sheet.hideColumns(3);
  sheet.hideColumns(8);
}


// Function triggered on edit
// Fires off the refreshSSInfo function if active sheet is console
function pushData_specialEdit () {
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var activeSheetName = activeSheet.getName();
  var cell = activeSheet.getActiveCell();
  var col = cell.getColumn();
  var row = cell.getRow();
  if (activeSheetName == "Settings Console") {
    refreshSSInfo (row, col);
  }
  pushData_restoreHeadings();
}

function pushData_refreshSSKeys () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getSheetByName('Settings Console').getDataRange();
  var lastRow = range.getLastRow();
  for (var i=1; i<lastRow; i++) {
    var row = i+1;
    refreshSSInfo(row, 1);
    refreshSSInfo(row, 6);
  }
}


// Populates the source sheet name, headers from the SSKey value in cols 1 and 6
// Does validation for the SSKey and Sheet Name
function refreshSSInfo (row, col) {
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings Console");
  var rootFolderId = getRootFolderId();
  if ((col==1)||(col==2)) {
    var sheetName = activeSheet.getRange(row, 2).getValue();
    var range = activeSheet.getRange(row, 1, 1, 4);
    var SSName = activeSheet.getRange(row, 1).getValue();
    var SSKey = getFileIdFromName(rootFolderId, SSName);
    try
    {
      var SS = SpreadsheetApp.openById(SSKey);
      var SSName = SS.getName();
      var SSURL = SS.getUrl();
      var SSLink = '=hyperlink("'+SSURL+'", "'+ SSName + '")';
      if (sheetName == '') {
        sheetName = SS.getSheets()[0].getName();
      }
      range.clearComment();
      range.setFontColor("black");
    }
    catch(err)
    {
      var SSLink = SSName;
      range.setComment("Spreadsheet not found");
      range.setFontColor("red");
    }
    
    try
    {
      var sheet = SS.getSheetByName(sheetName); 
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].join(", ");
      range.setFontColor("black");
    }
    catch(err)
    {
      var headers = "SHEET EMPTY OR NOT FOUND";
      range.setFontColor("red");
    }
    range.setValues([[SSLink, sheetName, SSKey, headers]])
  }
  if ((col==6)||(col==7)) {
    var sheetName = activeSheet.getRange(row, 7).getValue();
    var range = activeSheet.getRange(row, 6, 1, 4);
    var SSName = activeSheet.getRange(row, 6).getValue();
    var SSKey = getFileIdFromName(rootFolderId, SSName);
    try
    {
      var SS = SpreadsheetApp.openById(SSKey);
      var SSName = SS.getName();
      var SSURL = SS.getUrl();
      var SSLink = '=hyperlink("'+SSURL+'", "'+ SSName + '")';
      if (sheetName == '') {
        sheetName = SS.getSheets()[0].getName();
      }
      range.clearComment();
      range.setFontColor("black");
    }
    catch(err)
    {
      var SSLink = SSName;
      range.setComment("Spreadsheet not found");
      range.setFontColor("red");
    }
    
    try
    {
      var sheet = SS.getSheetByName(sheetName);
      if (sheet.getLastColumn()) {
        var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].join(", ");
      } else {
        var headers = "NO HEADERS";
      }
      range.setFontColor('black');
    }
    catch(err)
    {
      headers = "SHEET NOT FOUND";
      range.setFontColor("red");
    }
    range.setValues([[SSLink, sheetName, SSKey, headers]])
  }
}


//Main function where pushing of data occurs
function pushData (override) {
  setPushDataUid();
  pushData_getInstitutionalTrackerObject();
  pushData_refreshSSKeys();
  var pushStartTime = new Date();
  pushStartTime = parseInt(pushStartTime.getTime()).toFixed();
  // Wrap whole function in a try catch structure 
  // Use catch(err) code to set a temporary trigger to run again in 30 sec.
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Settings Console");
    var range = sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn());
    var objects = pushData_getRowsData(sheet, range);
    var loopIterationStartTime;
    
    //initialize all push times in ScriptProperties if override is true
    if (override==true) {
      for (var i=0; i<objects.length; i++) {
        ScriptProperties.setProperty('lastPushTime-'+i, 'initialize');
      }
    }
    
    //repopulate all links with formulas in the console objects
    for (var i = 0; i<objects.length; i++) {
      var sourceKey = objects[i].sourceSpreadsheetKey;
      var destKey = objects[i].destinationSpreadsheetKey;
      if ((sourceKey!="File not found")&&(destKey!="File not found")) {
        objects[i].sourceSpreadsheetName = pushData_makeLink(sourceKey);
        Utilities.sleep(10);
        objects[i].destinationSpreadsheetName = pushData_makeLink(destKey);
        Utilities.sleep(10);
      }
    }
    
    //begin loop through all console rows
    //check total execution time.  Break loop and schedule rerun in 30 seconds if greater than 270 seconds.
    for (var i = 0; i<objects.length; i++) {
      loopIterationStartTime = new Date();
      loopIterationStartTime = parseInt(loopIterationStartTime.getTime()).toFixed();
      var milliSecondsElapsed = loopIterationStartTime - pushStartTime;
      var secondsElapsed = milliSecondsElapsed/1000;
      if (secondsElapsed > 200) {
        var oldDateObj = new Date();
        var newDateObj = new Date(oldDateObj.getTime() + (0.5)*60000);
        var oneTimeRerunTrigger = ScriptApp.newTrigger('recoverPushData').timeBased().at(newDateObj).create();
        break;
      }
      // set a message on current row status
      var messageCell = sheet.getRange(i+2, 11);
      if (override==true) {
        messageCell.setValue("Overriding set time interval...");
      } else {
        messageCell.setValue("Checking time interval...");
      }
      Utilities.sleep(50);
      // initialize the status message for this loop through
      var message = '';
      
      //check to see if lastPushtime was logged for this row 
      var lastPushTime = ScriptProperties.getProperty('lastPushTime-'+i);
      if (lastPushTime == null) {  //if not, set to initialize
        ScriptProperties.setProperty('lastPushTime-'+i, 'initialize');
      }
      //check to see if pushInterval is set in console
      var pushInterval = objects[i].pushInterval;
      if (pushInterval=='') { // if not, set to 24 hrs
        pushInterval = '24';
        objects[i].pushInterval = '24';
      }
      // set pushinterval as a Script Property for access in checkPushInterval function
      ScriptProperties.setProperty('pushInterval-'+i, pushInterval);
      
      var pushFlag = false;
      var currStatus = objects[i].statusMessage;
      if (sourceKey=="File not found") {
        message += "ERROR: Source SS Sheet not found. Double check that it is correctly named. ";
        messageCell.setValue(message);
        objects[i].statusMessage = message;
        var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
        pushData_setRowsData(sheet, objects, headersRange);
        try {
          pushData_logError()
        }
        catch(err) {
        }
        break;
      }
      
      if (destKey=="File not found") {
        message += "ERROR: Destination SS Sheet not found. Double check that it is correctly named. ";
        messageCell.setValue(message);
        objects[i].statusMessage = message;
        var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
        pushData_setRowsData(sheet, objects, headersRange);
        try {
          pushData_logError()
        } catch(err) {
        }
        break;
      }
      
      if ((checkPushInterval(i))||(override==true)||(currStatus=="rerun")||(currStatus=="Pushing...")) {  //run the pushData only if overRide is true or pushInterval checker returns true
        while (messageCell.getValue()!="Pushing...") {  //while loop ensures that the setValue actually works, doesn't work without.
          messageCell.clear();
          messageCell.setValue("Pushing...");
          pushFlag = true;
          Utilities.sleep(5);
        }
      } else { //assume the row is getting skipped
        while (messageCell.getValue()!="Skipping...already pushed within time interval.") {
          messageCell.clear();
          messageCell.setValue("Skipping...already pushed within time interval.");
          Utilities.sleep(5);
        }
      }
      
      if (pushFlag==true) {
        var sourceSpreadsheetKey = objects[i].sourceSpreadsheetKey;
        //run error checking for bad SSKeys.  Write errors to sheet.
        try 
        {
          var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetKey);
        }
        catch(err)
        {
          message += "ERROR: Source SS not found. Double check that it is correctly named. ";
          objects[i].statusMessage = message;
          var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
          pushData_setRowsData(sheet, objects, headersRange);
          try {
            pushData_logError();
          } catch(err) {
          }
          continue;
        }
        
        
        var sourceSheetName = objects[i].sourceSheetName;
        try 
        {
          var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
        } catch (err) {
          message += "ERROR: Source SS Sheet named \"" + sourceSheetName + "\" not found. Double check that it is correct. ";
          objects[i].statusMessage = message;
          var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
          pushData_setRowsData(sheet, objects, headersRange);
          try {
            pushData_logError()
          } catch(err) {
          }
          continue;
        }
        try
        {
          var sourceRange = sourceSheet.getRange(2, 1, sourceSheet.getLastRow()-1, sourceSheet.getLastColumn());
        }
        catch(err)
        {
          message += "ERROR: Source SS Sheet named \"" + sourceSheetName + "\" contains no data. ";
          objects[i].statusMessage = message;
          var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
          pushData_setRowsData(sheet, objects, headersRange);
          try {
            pushData_logError();
          } catch(err) {
          }
          continue;
        }
        
        var sourceObjects = pushData_getRowsData(sourceSheet, sourceRange);
        var sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0].join(", "); 
        objects[i].availableColumns = sourceHeaders;
        var sourceHeadersToPush = objects[i].columnsToPush;
        if (sourceHeadersToPush!='') {
          sourceHeadersToPush = sourceHeadersToPush.split(",");
        } else {
          sourceHeadersToPush = sourceHeaders.split(",");
          sourceHeadersToPush = [sourceHeadersToPush[0]];
          objects[i].columnsToPush = sourceHeadersToPush[0];
        }
        for (var j=0; j< sourceHeadersToPush.length; j++) {
          sourceHeadersToPush[j] = sourceHeadersToPush[j].replace(/^\s\s*/, '').replace(/\s\s*$/, '');
        }
        sourceHeadersToPush = [sourceHeadersToPush]; 
        var destinationSpreadsheetKey = objects[i].destinationSpreadsheetKey;
        
        try 
        {
          var destinationSpreadsheet = SpreadsheetApp.openById(destinationSpreadsheetKey);
        }
        catch(err)
        {
          message += "ERROR: Destination SS not found. Double check that it is correctly named.";
          objects[i].statusMessage = message;
          var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
          pushData_setRowsData(sheet, objects, headersRange);
          try {
            pushData_logError();
          } catch(err) {
          }
          continue;
        }
        
        
        var destinationSheetName = objects[i].destinationSheetName;
        try
        {
          var destinationSheet = destinationSpreadsheet.getSheetByName(destinationSheetName);
          var lastRow = 0;
          Utilities.sleep(50);
          lastRow = destinationSheet.getLastRow();
          if (lastRow<2) {
            lastRow = 2;
          }
          var destinationRange = destinationSheet.getRange(2, 1, lastRow-1, sourceHeadersToPush[0].length);
          var extraRows = destinationSheet.getMaxRows() - sourceSheet.getLastRow();
          destinationRange.clearContent();
          if (extraRows>0) {
            destinationSheet.deleteRows(sourceSheet.getLastRow()+1, extraRows);
          }
        }
        catch(err)
        {
          message = "ERROR: Destination Sheet named \"" + destinationSheetName + "\" is incorrect.";
          objects[i].statusMessage = message;
          var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
          pushData_setRowsData(sheet, objects, headersRange);
          try {
            pushData_logError();
          } catch(err) {
          }
          continue;
        }
        
        var destinationHeaderRange = destinationSpreadsheet.getSheetByName(destinationSheetName).getRange(1, 1, 1, sourceHeadersToPush[0].length);
        destinationHeaderRange.setValues(sourceHeadersToPush);
        pushData_setRowsData(destinationSheet, sourceObjects, destinationHeaderRange);
        var destinationHeaders = destinationSheet.getRange(1, 1, 1, destinationSheet.getLastColumn()).getValues()[0].join(", ");
        objects[i].columnsInDestSheet = destinationHeaders;  
        var timestamp = new Date();
        destinationSheet.getRange(1, destinationSheet.getLastColumn()).setComment('Note: Columns 1 through ' + destinationSheet.getLastColumn() + ' were last pushed at ' + timestamp + ' from ' + sourceSpreadsheet.getName() + " (" + sourceSpreadsheet.getUrl() + ")" + " by the pushData script installed in " +  ss.getName() + " (" + ss.getUrl() + ")");
        var timeZone = Session.getTimeZone();
        var formatted = Utilities.formatDate(timestamp, timeZone, "M/d/y' 'h:mm:ss' 'a");
        var seconds = parseInt(timestamp.getTime()).toFixed();
        message = lastRow-1 + " rows cleared and " + sourceObjects.length + " rows pushed on " + formatted;
        ScriptProperties.setProperty('lastPushTime-'+i, seconds);
        objects[i].statusMessage = message;
        sheet.getRange(i+2, 11).setValue(message);
        try {
          pushData_logPush();
        } catch(err) {
        }
      } 
    }
    var headersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    if (override==true) {
      Browser.msgBox("Completed " +i+" push processes. Check status messages for info.");
      var app = UiApp.getActiveApplication();
      app.close();
      return app;
    }
    pushData_setRowsData(sheet, objects, headersRange);
  } 
  catch (err) {
    var oldDateObj = new Date();
    var newDateObj = new Date(oldDateObj.getTime() + (0.5)*60000);
    var oneTimeRerunTrigger = ScriptApp.newTrigger('recoverPushData').timeBased().at(newDateObj).create();
    Browser.msgBox("pushData will resume in one minute to avoid server timeout");
  }
}


function pushData_makeLink (docKey) {
  try 
  { 
    //var doc = SpreadsheetApp.openById(docKey);
    var doc = DocsList.getFileById(docKey);
    var docName = doc.getName();
    var docURL = doc.getUrl();
    var link = '=hyperlink("'+docURL+'", "'+ docName + '")';
    return link;
  } 
  catch (err) 
  {
    link = "ERROR: Destination SS not found. Double check that it is correctly named."
    return link;
  }
}


function recoverPushData () {
  var allTriggers = ScriptApp.getScriptTriggers();
  for (var i=0; i<allTriggers.length; i++) {
    if (allTriggers[i].getHandlerFunction()=="recoverPushData") {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }
  pushData();
}


function checkPushInterval(rowIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Settings Console");
  var pushInterval = ScriptProperties.getProperty('pushInterval-'+rowIndex);
  var lastPushTime = ScriptProperties.getProperty('lastPushTime-'+rowIndex);
  if (lastPushTime == 'initialize') {
    return true;
  }
  lastPushTime = +lastPushTime;
  lastPushTime = lastPushTime.toFixed();
  var milliseconds = new Date().getTime();
  milliseconds = parseInt(milliseconds).toFixed();
  var hoursDifference = ((milliseconds - lastPushTime)/1000)/3600;
  var hoursInterval = pushInterval;
  if (hoursDifference>hoursInterval) {
    return true;
  } else {
    return false;
  }
}
