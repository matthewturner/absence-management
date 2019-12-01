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
  var numRows = 1;
  var startRow = sheet.getActiveCell().getRowIndex();
  var dataRange = sheet.getRange(startRow, 1, numRows, 5);
  var currentRow = dataRange.getValues()[0];
  var calendarRange = sheet.getRange(startRow, GOOGLE_CALENDAR_COLUMN_INDEX);
  var cal = CalendarApp.getDefaultCalendar();
  var title = currentRow[0];
  var tstart = new Date(currentRow[2]);
  var tstop = new Date(currentRow[4]);
  tstop.setDate(tstop.getDate() + 1)
  var event = cal.createEvent(title, tstart, tstop);
  var calendarRange = sheet.getRange(startRow, GOOGLE_CALENDAR_COLUMN_INDEX);
  calendarRange.setValue(event.getId());
  calendarRange.setBackground("White");
};

function checkEventForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var numRows = 1;
  var startRow = sheet.getActiveCell().getRowIndex();
  var dataRange = sheet.getRange(startRow, 1, numRows, 5);
  var currentRow = dataRange.getValues()[0];
  var calendarRange = sheet.getRange(startRow, GOOGLE_CALENDAR_COLUMN_INDEX);
  var calendarId = calendarRange.getValue();
  var cal = CalendarApp.getDefaultCalendar();
  var title = currentRow[0];
  var tstart = new Date(currentRow[2]);
  var tstop = new Date(currentRow[4]);
  tstop.setDate(tstop.getDate() + 1);
  var event = findEvent(cal, calendarId, tstart, tstop, title);
  if (event === null) {
    calendarRange.setBackground("Red");
    return;
  }
  if (calendarRange.getValue() !== title) {
    calendarRange.setValue(event.getId());
  }
  if (event.getStartTime().getTime() == tstart.getTime() && event.getEndTime().getTime() == tstop.getTime()) {
    calendarRange.setBackground("White");
  } else {
    calendarRange.setBackground("Red");
  }
};

function findEvent(calendar, calendarId, startTime, endTime, title) {
  var event = calendar.getEventById(calendarId);
  if (event !== null) {
    return event;
  }
  var events = calendar.getEvents(startTime, endTime);
  for (var i = 0; i < events.length; i++) {
    var candidateEvent = events[i];
    if (candidateEvent.getTitle() === title) {
      return candidateEvent;
    }
  }
  return null;
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
