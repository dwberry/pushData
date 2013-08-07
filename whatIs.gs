function pushData_whatIs() {
  var app = UiApp.createApplication().setHeight(550);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var panel = app.createVerticalPanel();
  var octiGrid = app.createGrid(1, 2);
  var image = app.createImage(this.HAMSTERIMAGEURL);
  image.setHeight("100px");
  var label = app.createLabel("pushData: A utility for pushing tabular data between spreadsheets on time-based triggers.");
  label.setStyleAttribute('fontSize', '1.5em').setStyleAttribute('fontWeight', 'bold');
  octiGrid.setWidget(0, 0, image);
  octiGrid.setWidget(0, 1, label);
  var mainGrid = app.createGrid(4, 1);
  var html = "<h3>Features</h3>";
      html += "<ul><li>Works on other spreadsheets in the SAME FOLDER that you want to establish push relationships between.</li>";
      html += "<li>Auto-discovers source and destination spreadsheets based on file name. This makes systems built with pushData easy to copy and share using a folder-copying utility.</li>";
      html += "<li>Auto-discovers all headers in source and destination sheets.</li>";
      html += "<li>Allows user to specify which headers to push, in any order, from source sheet to destination sheet.</li>"; 
      html += "<li>Syncs each source with its destination on a time trigger of a user-specified number of hours between runs, making it more stable and reliable than IMPORTRANGE function for large data sets.</li>";
      html += "</ul>";
     
  mainGrid.setWidget(0, 0, app.createHTML(html));
  var sponsorLabel = app.createLabel("Brought to you by");
  var sponsorImage = app.createImage("http://www.youpd.org/sites/default/files/acquia_commons_logo36.png");
  var supportLink = app.createAnchor('Watch the tutorial!', 'http://www.youpd.org/pushdata');
  mainGrid.setWidget(1, 0, sponsorLabel);
  mainGrid.setWidget(2, 0, sponsorImage);
  mainGrid.setWidget(3, 0, supportLink);
  app.add(octiGrid);
  panel.add(mainGrid);
  app.add(panel);
  ss.show(app);
  return app;                                                                    
}
