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
  var calendar = new GoogleCalendar(CalendarApp.getDefaultCalendar());
  var event = calendar.createEvent(entry.getTitle(), entry.getStartTime(), entry.getAdjustedEndTime(calendar.getAdjustment()));
  entry.setCalendarId(event.getId());
  entry.clearCalendarConflict(calendar.getType());
};

function checkEventForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowIndex = sheet.getActiveCell().getRowIndex();
  var entry = new AbsenceEntry(sheet, rowIndex);
  var calendar = new GoogleCalendar(CalendarApp.getDefaultCalendar());
  var hrCalendar = new HrCalendar(sheet);

  new Synchronizer(entry, calendar).markSynchronized();
  new Synchronizer(entry, hrCalendar).markSynchronized();
};

function checkEventsForAllRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var calendar = new GoogleCalendar(CalendarApp.getDefaultCalendar());
  var hrCalendar = new HrCalendar(sheet);
  var range = sheet.getDataRange();
  var numRows = range.getNumRows();
  var rows = range.getValues();

  for (var rowIndex = 0; rowIndex < numRows; rowIndex++) {
    var row = rows[rowIndex];
    if (row[0] !== "Bank holiday" && row[0] !== "Type" && row[0] !== "") {
      var entry = new AbsenceEntry(sheet, rowIndex + 1);
      new Synchronizer(entry, calendar).markSynchronized();
      new Synchronizer(entry, hrCalendar).markSynchronized();
    }
  }
};

function syncEventForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowIndex = sheet.getActiveCell().getRowIndex();
  var entry = new AbsenceEntry(sheet, rowIndex);
  var calendar = new GoogleCalendar(CalendarApp.getDefaultCalendar());
  var calendarType = calendar.getType();

  var event = entry.findEvent(calendar);
  if (event === null) {
    var event = calendar.createEvent(entry.getTitle(), entry.getStartTime(), entry.getAdjustedEndTime(calendar.getAdjustment()));
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

  if (event.getStartTime().getTime() !== entry.getStartTime().getTime() || event.getEndTime().getTime() !== entry.getAdjustedEndTime(calendar.getAdjustment()).getTime()) {
    event.setTime(entry.getStartTime(), entry.getEndTime());
  }

  entry.clearCalendarConflict(calendarType);
};

function deleteEventForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowIndex = sheet.getActiveCell().getRowIndex();
  var entry = new AbsenceEntry(sheet, rowIndex);
  var calendar = new GoogleCalendar(CalendarApp.getDefaultCalendar());

  var event = entry.findEvent(calendar);
  if (event !== null) {
    event.deleteEvent();
  }
  entry.setCalendarId(null);
  entry.clearCalendarConflict(calendar.getType());
};

function configureActiveRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowIndex = sheet.getActiveCell().getRowIndex();
  var entry = new AbsenceEntry(sheet, rowIndex);

  entry.configure();
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
      name: "Insert event for active row",
      functionName: "insertEventForActiveRow"
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
