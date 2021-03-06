function sheetSpider_whatIs() {
  var app = UiApp.createApplication().setHeight(550);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var panel = app.createVerticalPanel();
  var octiGrid = app.createGrid(1, 2);
  var image = app.createImage('https://drive.google.com/uc?export=download&id='+this.WAITINGICONID);
  var diagram1 = app.createImage('https://drive.google.com/uc?export=download&id=0B2vrNcqyzernWEFId0JjOWR5dEk');
  var diagram2 = app.createImage('https://drive.google.com/uc?export=download&id=0B2vrNcqyzernWDh0amlhQkdCY2s').setStyleAttribute("marginTop","10px");
  image.setHeight("100px").setWidth("100px");
  var label = app.createLabel("sheetSpider: filter and push form or sheet data out to multiple Google Spreadsheets based on a unique \"entity\" (e.g. student name, teacher name, etc) criterion");
  label.setStyleAttribute('fontSize', '1.5em').setStyleAttribute('fontWeight', 'bold');
  octiGrid.setWidget(0, 0, image);
  octiGrid.setWidget(0, 1, label);
  var mainGrid = app.createGrid(4, 1);
  var scrollPanel = app.createScrollPanel().setHeight("250px").setStyleAttribute("backgroundColor","whiteSmoke").setStyleAttribute('padding', '10px');
  var innerPanel = app.createVerticalPanel();
  var html = "<h3>Features</h3>";
  html += "<ul><li>Sets up and shares \"entity\" spreadsheets with individual collaborators.  Entities can be anything -- student names, course, teachers, etc. -- as long as they have unique names.</li>";
  html += "<li>Allows you to specify the \"feeder\" sheet within the spreadsheet.  This is the sheet where aggregate data lives -- either via Google Form or as a bulk data set.</li>";
  html += "<li>For Google Form mode, sheetSpider automatically populates one listbox, checkbox, or multiple choice question on your form with entity names.  This entity name is used to determine which spreadsheet to send inbound data to.</li>";
  html += "<li>When multiple entity names are selected (i.e. via checkbox question type), data is sent to multiple spreadsheets.</li>";
  html += "<li>Assesses the uniqueness of each record (you can indicate multiple fields as uniqueness criteria) before pushing to the destination sheet.</li>";
  html += "<li>Can be run in manual mode -- i.e. batch disaggregate and push data to entity spreadsheets</li>"
  html += "<li>Allows a live \"pull\" of current data in all entity spreadsheets -- great for harvesting back changes in status, completed records, etc. from constituents (students, teachers, etc.)</li></ul>";
  innerPanel.add(app.createHTML(html));
  innerPanel.add(diagram1);
  innerPanel.add(diagram2);
  scrollPanel.add(innerPanel);
  mainGrid.setWidget(0, 0, scrollPanel);
  var sponsorLabel = app.createLabel("Brought to you by");
  var sponsorImage = app.createImage("http://www.youpd.org/sites/default/files/acquia_commons_logo36.png");
 // var supportLink = app.createAnchor('Watch the tutorial!', 'http://www.youpd.org/sheetspider');
  mainGrid.setWidget(1, 0, sponsorLabel);
  mainGrid.setWidget(2, 0, sponsorImage);
//  mainGrid.setWidget(3, 0, supportLink);
  app.add(octiGrid);
  panel.add(mainGrid);
  app.add(panel);
  ss.show(app);
  return app;                                                                    
}
