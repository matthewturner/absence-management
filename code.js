/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var numRows = range.getNumRows();
  var rows = range.getValues();

  for (var i = 0; i < numRows; i++) {
    var row = rows[i];
    Logger.log(row);
  }
};

function insertEventForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowIndex = sheet.getActiveCell().getRowIndex();
  var entry = new AbsenceEntry(sheet, rowIndex);
  var calendar = CalendarApp.getDefaultCalendar();
  var event = calendar.createEvent(entry.getTitle(), entry.getStartTime(), entry.getEndTime());
  entry.setCalendarId(event.getId());
  var calendarType = "google";
  entry.clearCalendarConflict(calendarType);
};

function checkEventForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowIndex = sheet.getActiveCell().getRowIndex();
  var entry = new AbsenceEntry(sheet, rowIndex);
  var calendar = CalendarApp.getDefaultCalendar();
  
  var calendarType = "google";
  checkEventForRow(entry, calendar, calendarType);
};

function checkEventForRow(entry, calendar, calendarType) {
  var event = entry.findEvent(calendar, calendarType);
  if (event === null) {
    entry.markCalendarConflict(calendarType);
    return;
  }
  if (entry.getCalendarId(calendarType) !== event.getId()) {
    entry.markCalendarConflict(calendarType);
    return;
  }
  if (entry.getTitle() !== event.getTitle()) {
    entry.markCalendarConflict(calendarType);
    return;
  }
  if (event.getStartTime().getTime() !== entry.getStartTime().getTime()) {
    entry.markCalendarConflict(calendarType);
    return;
  }
  if (event.getEndTime().getTime() !== entry.getEndTime().getTime()) {
    entry.markCalendarConflict(calendarType);
    return;
  }
  entry.clearCalendarConflict(calendarType);
};

function checkEventsForAllRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var calendar = CalendarApp.getDefaultCalendar();
  var hrCalendar = new HrCalendar(sheet);
  var range = sheet.getDataRange();
  var numRows = range.getNumRows();
  var rows = range.getValues();

  for (var rowIndex = 0; rowIndex < numRows; rowIndex++) {
    var row = rows[rowIndex];
    if (row[0] !== "Bank holiday" && row[0] !== "Type" && row[0] !== "") {
      var entry = new AbsenceEntry(sheet, rowIndex + 1);
      checkEventForRow(entry, calendar, "google");
      checkEventForRow(entry, hrCalendar, "hr");
    }
  }
};

function syncEventForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowIndex = sheet.getActiveCell().getRowIndex();
  var entry = new AbsenceEntry(sheet, rowIndex);
  var calendar = CalendarApp.getDefaultCalendar();
  var calendarType = "google";
  
  var event = entry.findEvent(calendar);
  if (event === null) {
    var event = calendar.createEvent(entry.getTitle(), entry.getStartTime(), entry.getEndTime());
    entry.setCalendarId(event.getId());
    entry.clearCalendarConflict(calendarType);
    return;
  }
  
  if (entry.getCalendarId(calendarType) !== event.getId()) {
    entry.setCalendarId(event.getId());
  }
  
  if (entry.getTitle() !== event.getTitle()) {
    event.setTitle(entry.getTitle());
  }
  
  if (event.getStartTime().getTime() !== entry.getStartTime().getTime() || event.getEndTime().getTime() !== entry.getEndTime().getTime()) {
    event.setTime(entry.getStartTime(), entry.getEndTime());
  }
  
  entry.clearCalendarConflict(calendarType);
};

function deleteEventForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowIndex = sheet.getActiveCell().getRowIndex();
  var entry = new AbsenceEntry(sheet, rowIndex);
  var calendar = CalendarApp.getDefaultCalendar();
  
  var event = entry.findEvent(calendar);
  if (event !== null) {
    event.deleteEvent();
  }
  entry.setCalendarId(null);
  var calendarType = "google";
  entry.clearCalendarConflict(calendarType);
};

function configureActiveRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowIndex = sheet.getActiveCell().getRowIndex();
  var entry = new AbsenceEntry(sheet, rowIndex);
  
  entry.configure();
};

var HrCalendar = function(sheet) {
  this.hrSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet.getSheetName() + " - HR");
  
  this.getEvents = function() {
    var range = this.hrSheet.getDataRange();
    var numRows = range.getNumRows();
    var rows = range.getValues();
    
    var events = [];
    for (var rowIndex = 1; rowIndex < numRows; rowIndex++) {
      var row = rows[rowIndex];
      events.push(new HrCalendarEvent(row));
    }
    return events;
  };
  
  this.allEvents = this.getEvents();
  
  this.getEventById = function(calendarId) {
    // not supported
    return null;
  };
  
  this.getEvents = function(startTime, endTime) {
    return this.allEvents.filter(function(item) {
      if (item.getStartTime().getTime() !== startTime.getTime()) {
        return false;
      }
      if (item.getEndTime().getTime() !== endTime.getTime()) {
        return false;
      }
      return true;
    });
  };
}

var HrCalendarEvent = function(row) {
  this.row = row;
  
  this.getTitle = function() {
    var title = this.row[2];
    return this.row[2];
  };
  
  this.getStartTime = function() {
    var title = new Date(this.row[4]);
    return new Date(this.row[4]);
  };
  
  this.getEndTime = function() {
    var title = new Date(this.row[5]);
    return new Date(this.row[5])
  };
};

var AbsenceEntry = function(sheet, rowIndex) {
  var GOOGLE_CALENDAR_COLUMN_INDEX = 8;
  var HR_CALENDAR_COLUMN_INDEX = 10;
  var NUMBER_OF_ROWS = 1;
  this.sheet = sheet;
  this.rowIndex = rowIndex;
  this.dataRange = sheet.getRange(rowIndex, 1, NUMBER_OF_ROWS, 5);
  this.currentRow = this.dataRange.getValues()[0];
  
  this.getTitle = function() {
    return this.currentRow[0];
  };
  
  this.getCalendarCell = function(calendarType) {
    switch(calendarType) {
      case "google":
        return this.sheet.getRange(rowIndex, GOOGLE_CALENDAR_COLUMN_INDEX);
      case "hr":
        return this.sheet.getRange(rowIndex, HR_CALENDAR_COLUMN_INDEX);
    }   
  };
  
  this.getCalendarId = function(calendarType) {
    var calendarCell = this.getCalendarCell(calendarType);
    var calendarId = calendarCell.getValue();
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
  
  this.markCalendarConflict = function(calendarType) {
    var calendarCell = this.getCalendarCell(calendarType);
    calendarCell.setBackground("Red");
  };
  
  this.clearCalendarConflict = function(calendarType) {
    var calendarCell = this.getCalendarCell(calendarType);
    calendarCell.setBackground("White");
  };
  
  this.configure = function() {
    var startDayCell = sheet.getRange(rowIndex, 2);
    startDayCell.setFormula("=text(C" + rowIndex + ", \"ddd\")");
    var startTimeCell = sheet.getRange(rowIndex, 3);
    startTimeCell.setNumberFormat("dd/MM/YYYY");
    
    var startDayCell = sheet.getRange(rowIndex, 4);
    startDayCell.setFormula("=text(E" + rowIndex + ", \"ddd\")");
    var endTimeCell = sheet.getRange(rowIndex, 5);
    endTimeCell.setNumberFormat("dd/MM/YYYY");
    
    var dayCountCell = sheet.getRange(rowIndex, 6);
    dayCountCell.setFormula("=E" + rowIndex + " - C" + rowIndex + " + 1");
  };
  
  this.findEvent = function(calendar, calendarType) {
    var event = calendar.getEventById(this.getCalendarId(calendarType));
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
  var entries = [
                 {
                   name: "Check all rows",
                   functionName: "checkEventsForAllRows"
                 },
                 {
                   name: "Configure active row",
                   functionName: "configureActiveRow"
                 },
                 {
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
                 },
                 {
                   name: "Delete event for active row",
                   functionName: "deleteEventForActiveRow"
                 }];
  sheet.addMenu("Calendar", entries);
};
