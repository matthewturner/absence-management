var GOOGLE_CALENDAR_COLUMN_INDEX = 8;

/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    Logger.log(row);
  }
};

function insertEventForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var numRows = 1;   // Number of rows to process
  var startRow = sheet.getActiveCell().getRowIndex();
  var dataRange = sheet.getRange(startRow, 1, numRows, 5);
  var data = dataRange.getValues();
  var cal = CalendarApp.getDefaultCalendar();
  for (i in data) {
    var row = data[i];
    var title = row[0];
    var tstart = row[2];
    var tstop = new Date(row[4]);
    tstop.setDate(tstop.getDate() + 1)
    var event = cal.createEvent(title, tstart, tstop);
    var calendarRange = sheet.getRange(startRow, GOOGLE_CALENDAR_COLUMN_INDEX);
    calendarRange.setValue(event.getId());
    calendarRange.setBackground("White");
 }
};

function checkEventForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var numRows = 1;
  var startRow = sheet.getActiveCell().getRowIndex();
  var dataRange = sheet.getRange(startRow, 1, numRows, 5);
  var data = dataRange.getValues();
  var calendarRange = sheet.getRange(startRow, GOOGLE_CALENDAR_COLUMN_INDEX);
  var calendarId = calendarRange.getValue();
  var cal = CalendarApp.getDefaultCalendar();
  var event = cal.getEventById(calendarId);
  if (event === null) {
    calendarRange.setBackground("Red");
    return;
  }
  for (i in data) {
    var row = data[i];
    var title = row[0];
    var tstart = new Date(row[2]);
    var tstop = new Date(row[4]);
    tstop.setDate(tstop.getDate() + 1)
    if (event.getStartTime().getTime() == tstart.getTime() && event.getEndTime().getTime() == tstop.getTime()) {
      calendarRange.setBackground("White");
    } else {
      calendarRange.setBackground("Red");
    }
  }
};

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
                   name : "Insert event for active row",
                   functionName : "insertEventForActiveRow"
                 },
                 {
                   name: "Check event for active row",
                   functionName: "checkEventForActiveRow"
                 }];
  sheet.addMenu("Calendar", entries);
};
