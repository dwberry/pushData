function pushData_pickerUi() {
  var ss = SpreadsheetApp.getActive();
  // SpreadsheetApp returns one spreadsheet Key
  var ssId = ss.getId();
  var thisFile = DocsList.getFileById(ssId);
  // DocsList returns a different key?
  ssId = thisFile.getId();
  var app = UiApp.createApplication().setTitle("Create new push process").setHeight(400);
  var panel = app.createVerticalPanel();
  var scrollpanel = app.createScrollPanel().setHeight("300px");
  var helptext = "Use this panel to insert valid Spreadsheet and sheet names into a new row in the Settings Console.  FYI: You can also type them directly into the console.";
  var helpLabel = app.createLabel(helptext);
  var helpPanel = app.createPopupPanel();
  helpPanel.add(helpLabel);
  panel.add(helpPanel);
  var grid = app.createGrid(7,2).setId("grid");
  grid.setWidget(0, 0, app.createLabel("Source").setWidth("200px").setStyleAttribute("backgroundColor", "whiteSmoke"));
  grid.setWidget(0, 1, app.createLabel("Destination").setWidth("200px").setStyleAttribute("backgroundColor", "whiteSmoke"));
  var rootFolderId = getRootFolderId();
  var rootFolder = DocsList.getFolderById(rootFolderId);
  var files = rootFolder.getFilesByType('spreadsheet');
  var validityHandler = app.createServerHandler('assessValidity').addCallbackElement(panel);
  
  var spinner = app.createImage(HAMSTERIMAGEURL).setWidth(150);
  spinner.setVisible(false);
  spinner.setStyleAttribute("position", "absolute");
  spinner.setStyleAttribute("top", "90px");
  spinner.setStyleAttribute("left", "175px");
  spinner.setId("dialogspinner");
  
  var sourceSSChangeHandler = app.createServerHandler('refreshSourceSheet').addCallbackElement(panel);
  var sourceSSPicker = app.createListBox().setId("sourceSSPicker").setWidth("190px").setName("sourceSSId").addChangeHandler(sourceSSChangeHandler).addChangeHandler(validityHandler);
  sourceSSPicker.addItem("choose source spreadsheet");
  for (var i=0; i<files.length; i++) {
    if (files[i].getId()!=ssId) {
      sourceSSPicker.addItem(files[i].getName(), files[i].getId());
    }
  }
  grid.setWidget(1, 0, sourceSSPicker);
  var destSSChangeHandler = app.createServerHandler('refreshDestinationSheet').addCallbackElement(grid);
  var destSSPicker = app.createListBox().setId("destSSPicker").setWidth("190px").setName("destSSId").addChangeHandler(destSSChangeHandler).addChangeHandler(validityHandler);
  destSSPicker.addItem("choose destination spreadsheet");
  for (var i=0; i<files.length; i++) {
    if (files[i].getId()!=ssId) {
      destSSPicker.addItem(files[i].getName(), files[i].getId());
    }
  }
  grid.setWidget(1, 1, destSSPicker);
  scrollpanel.add(grid);
  panel.add(scrollpanel);
  var insertHandler = app.createServerHandler('insertRow').addCallbackElement(panel);
  var clientHandler = app.createClientHandler().forTargets(panel).setStyleAttribute('opacity', '0.5').forTargets(spinner).setVisible(true);
  var button = app.createButton("Insert into new row in console").setId("button").addClickHandler(insertHandler).setEnabled(false).addClickHandler(clientHandler);
  panel.add(button);
  
  
  
  
  
  app.add(panel);
  app.add(spinner);
  ss.show(app);
  return app;
}


function insertRow(e) {
  var app = UiApp.getActiveApplication();
  var sourceSSId = e.parameter.sourceSSId;
  var sourceSSName = SpreadsheetApp.openById(sourceSSId).getName();
  var sourceSheetName = e.parameter.sourceSheetName;
  var destSSId = e.parameter.destSSId;
  var destSSName = SpreadsheetApp.openById(destSSId).getName();
  var destSheetName = e.parameter.destSheetName;
  var headers = e.parameter.headers;
  headers = headers.split(",");
  var chosenHeaders = [];
  for (var i=0; i<headers.length; i++) {
    if (e.parameter['checkBox-'+i]=="true") {
     chosenHeaders.push(headers[i]);
    }
  }
  var headerString = chosenHeaders.join(", ");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Settings Console");
  var nextRow = sheet.getLastRow()+1;
  var range = sheet.getRange(nextRow, 1, 1, 10);
  var values = [[sourceSSName,sourceSheetName,'','',headerString,destSSName,destSheetName,'','','24']];
  range.setValues(values);
  app.close();
  pushData_refreshSSKeys();
  pushData_restoreHeadings();
  var cell = sheet.getRange(nextRow, 5);
  sheet.setActiveCell(cell);
  return app;
}



function assessValidity(e) {
  var app = UiApp.getActiveApplication();
  var button = app.getElementById("button");
  var sourceSSId = e.parameter.sourceSSId; 
  var sourceSheetName = e.parameter.sourceSheetName;
  var destSSId = e.parameter.destSSId;
  var destSheetName = e.parameter.destSheetName;
  if ((sourceSSId)&&(sourceSSId!="choose source spreadsheet")&&(sourceSheetName)&&(sourceSheetName!="choose source sheet")&&(destSSId)&&(destSSId!="choose source spreadsheet")&&(destSheetName)&&(destSheetName!="choose source sheet")) {
    button.setEnabled(true);
  } else {
    button.setEnabled(false);
  }
  return app;
}


function refreshSourceSheet(e) {
  var ss = SpreadsheetApp.getActive();
  // SpreadsheetApp returns one spreadsheet Key
  var ssId = ss.getId();
  var thisFile = DocsList.getFileById(ssId);
  // DocsList returns a different key?
  ssId = thisFile.getId();
  var app = UiApp.getActiveApplication();
  var grid = app.getElementById("grid");
  var sourceSSId = e.parameter.sourceSSId;
  grid.setWidget(2, 0, app.createLabel(''));
  grid.setWidget(3, 0, app.createLabel(''));
  grid.setWidget(4, 0, app.createLabel(''));
  if (sourceSSId=="choose source spreadsheet") {
    return app
  }
  var ss = SpreadsheetApp.openById(sourceSSId);
  var sheets = ss.getSheets();
  var validityHandler = app.createServerHandler('assessValidity').addCallbackElement(grid);
  var sourceSheetChangeHandler = app.createServerHandler('refreshSourceHeaders').addCallbackElement(grid);
  var sourceSheetPicker = app.createListBox().setWidth("190px").setName("sourceSheetName").addChangeHandler(sourceSheetChangeHandler).addChangeHandler(validityHandler);
  sourceSheetPicker.addItem("choose source sheet");
  for (var i=0; i<sheets.length; i++) {
    sourceSheetPicker.addItem(sheets[i].getName());
  }
  grid.setWidget(2, 0, sourceSheetPicker);
  return app;
}

function refreshSourceHeaders(e) {
  var app = UiApp.getActiveApplication();
  var grid = app.getElementById("grid");
  var sourceSSId = e.parameter.sourceSSId;
  var sourceSheetName = e.parameter.sourceSheetName;
  if ((!(sourceSSId))||(sourceSheetName=="choose source sheet")) {
    grid.setWidget(3,0,app.createLabel(''));
    grid.setWidget(4,0,app.createLabel(''));
    return app;
  }
  var sheet = SpreadsheetApp.openById(sourceSSId).getSheetByName(sourceSheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  if (headers) {
    var label = app.createLabel("Select Columns to Push");
    var scrollPanel = app.createScrollPanel();
    var checkPanel = app.createVerticalPanel();
    var checkBoxes = [];
    for (var i=0; i<headers[0].length; i++) {
      checkBoxes[i] = app.createCheckBox(headers[0][i]).setName('checkBox-'+i);
      checkPanel.add(checkBoxes[i]);
    }
    var headers = app.createTextBox().setVisible(false).setName("headers").setValue(headers[0].join(","));
    checkPanel.add(headers);
    scrollPanel.add(checkPanel);
    grid.setWidget(4,0,scrollPanel);
  } else {
   var label = app.createLabel("No headers available");
  }
  grid.setWidget(3,0,label).setStyleAttribute(3,0,'textAlign', 'left');
  return app;
}


function refreshDestinationSheet(e) {
  var app = UiApp.getActiveApplication();
  var grid = app.getElementById("grid");
  var destSSId = e.parameter.destSSId;
  grid.setWidget(2, 1, app.createLabel(''));
  grid.setWidget(3, 1, app.createLabel(''));
  grid.setWidget(4, 1, app.createLabel(''));
  if (destSSId=="choose destination spreadsheet") {
    return app
  }
  var ss = SpreadsheetApp.openById(destSSId);
  var sheets = ss.getSheets();
  var validityHandler = app.createServerHandler('assessValidity').addCallbackElement(grid);
  var destSheetChangeHandler = app.createServerHandler('refreshDestHeaders').addCallbackElement(grid);
  var destSheetPicker = app.createListBox().setWidth("190px").setName("destSheetName").addChangeHandler(destSheetChangeHandler).addChangeHandler(validityHandler);
  destSheetPicker.addItem('choose destination sheet');
  for (var i=0; i<sheets.length; i++) {
    destSheetPicker.addItem(sheets[i].getName());
  }
  grid.setWidget(2, 1, destSheetPicker);
  return app;
}



function refreshDestHeaders(e) {
  var ss = SpreadsheetApp.getActive();
  // SpreadsheetApp returns one spreadsheet Key
  var ssId = ss.getId();
  var thisFile = DocsList.getFileById(ssId);
  // DocsList returns a different key?
  ssId = thisFile.getId();
  var app = UiApp.getActiveApplication();
  var grid = app.getElementById("grid");
  var destSSId = e.parameter.destSSId;
  var destSheetName = e.parameter.destSheetName;
  if ((!(destSSId))||(destSheetName=="choose destination sheet")) {
    grid.setWidget(3,1,app.createLabel(''));
    grid.setWidget(4,1,app.createLabel(''));
    grid.setWidget(5,1,app.createLabel(''));
    return app;
  }
  var sheet = SpreadsheetApp.openById(destSSId).getSheetByName(destSheetName);
  var lastCol = sheet.getLastColumn();
  if(lastCol) {
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  }
  var label1 = app.createLabel("Columns in Dest Sheet");
  var label2 = app.createLabel("Caveat: This script will overwrite any overlapping columns in your destination sheet").setStyleAttribute('color', 'blue');
  if (headers) {
   var headerLabel = app.createLabel(headers);
  } else {
   var headerLabel = app.createLabel("none");
  }
  grid.setWidget(3,1,label1).setStyleAttribute(3,1,'textAlign', 'left');
  grid.setWidget(4,1,headerLabel).setStyleAttribute(4,1,'textAlign', 'center');
  grid.setWidget(5,1,label2);
  return app;
}
