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
  var startRow = sheet.getActiveCell().getRowIndex();
  var entry = new AbsenceEntry(sheet, startRow);
  var calendar = CalendarApp.getDefaultCalendar();
  var event = calendar.createEvent(entry.getTitle(), entry.getStartTime(), entry.getEndTime());
  entry.setCalendarId(event.getId());
  entry.clearCalendarConflict();
};

function checkEventForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = sheet.getActiveCell().getRowIndex();
  var calendar = CalendarApp.getDefaultCalendar();
  
  var entry = new AbsenceEntry(sheet, startRow);
  
  var event = entry.findEvent(calendar);
  if (event === null) {
    entry.markCalendarConflict();
    return;
  }
  if (entry.getCalendarId() !== event.getId()) {
    entry.markCalendarConflict();
    return;
  }
  if (entry.getTitle() !== event.getTitle()) {
    entry.markCalendarConflict();
    return;
  }
  if (event.getStartTime().getTime() !== entry.getStartTime().getTime() || event.getEndTime().getTime() !== entry.getEndTime().getTime()) {
    entry.markCalendarConflict();
    return;
  }
  entry.clearCalendarConflict();
};

function syncEventForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = sheet.getActiveCell().getRowIndex();
  var calendar = CalendarApp.getDefaultCalendar();
  
  var entry = new AbsenceEntry(sheet, startRow);
  
  var event = entry.findEvent(calendar);
  if (event === null) {
    var event = calendar.createEvent(entry.getTitle(), entry.getStartTime(), entry.getEndTime());
    entry.setCalendarId(event.getId());
    entry.clearCalendarConflict();
    return;
  }
  
  if (entry.getCalendarId() !== event.getId()) {
    entry.setCalendarId(event.getId());
  }
  
  if (entry.getTitle() !== event.getTitle()) {
    event.setTitle(entry.getTitle());
  }
  
  if (event.getStartTime().getTime() !== entry.getStartTime().getTime() || event.getEndTime().getTime() !== entry.getEndTime().getTime()) {
    event.setTime(entry.getStartTime(), entry.getEndTime());
  }
  
  entry.clearCalendarConflict();
};

var AbsenceEntry = function(sheet, rowIndex) {
  this.sheet = sheet;
  this.rowIndex = rowIndex;
  var numRows = 1;
  this.dataRange = sheet.getRange(rowIndex, 1, numRows, 5);
  this.currentRow = this.dataRange.getValues()[0];
  this.calendarCell = sheet.getRange(rowIndex, GOOGLE_CALENDAR_COLUMN_INDEX);
  
  this.getTitle = function() {
    return this.currentRow[0];
  };
  
  this.getCalendarId = function() {
    var calendarId = this.calendarCell.getValue();
    return calendarId;
  };
  
  this.setCalendarId = function(calendarId) {
    this.calendarCell.setValue(calendarId);
  };
  
  this.getStartTime = function() {
    var tstart = new Date(this.currentRow[2]);
    return tstart;
  };
  
  this.getEndTime = function() {
    var tstop = new Date(this.currentRow[4]);
    tstop.setDate(tstop.getDate() + 1);
    return tstop;
  };
  
  this.markCalendarConflict = function() {
    this.calendarCell.setBackground("Red");
  };
  
  this.clearCalendarConflict = function() {
    this.calendarCell.setBackground("White");
  };
  
  this.findEvent = function(calendar) {
    var event = calendar.getEventById(this.getCalendarId());
    if (event !== null) {
      return event;
    }
    var events = calendar.getEvents(this.getStartTime(), this.getEndTime());
    for (var i = 0; i < events.length; i++) {
      var candidateEvent = events[i];
      if (candidateEvent.getTitle() === this.getTitle()) {
        return candidateEvent;
      }
    }
    return null;
  };
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
                 },
                 {
                   name: "Sync event for active row",
                   functionName: "syncEventForActiveRow"
                 }];
  sheet.addMenu("Calendar", entries);
};
