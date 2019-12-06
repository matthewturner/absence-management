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
  if (!authorizeIfRequired()) {
    return;
  }
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowIndex = sheet.getActiveCell().getRowIndex();
  var entry = new AbsenceEntry(sheet, rowIndex);
  var googleCalendar = new GoogleCalendar(CalendarApp.getDefaultCalendar());
  var office365Calendar = new Office365Calendar();

  insertEvent(googleCalendar, entry);
  insertEvent(office365Calendar, entry);
};

function insertEvent(calendar, entry) {
  var event = calendar.createEvent(entry.getTitle(), entry.getStartTime(), entry.getAdjustedEndTime(calendar.getAdjustment()));
  var calendarType = calendar.getType();
  entry.setCalendarId(calendarType, event.getId());
  entry.clearCalendarConflict(calendarType);
};

function checkEventForActiveRow() {
  if (!authorizeIfRequired()) {
    return;
  }
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowIndex = sheet.getActiveCell().getRowIndex();
  var entry = new AbsenceEntry(sheet, rowIndex);
  var googleCalendar = new GoogleCalendar(CalendarApp.getDefaultCalendar());
  var hrCalendar = new HrCalendar(sheet);
  var office365Calendar = new Office365Calendar();
  
  new Synchronizer(entry, googleCalendar).markSynchronized();
  new Synchronizer(entry, hrCalendar).markSynchronized();
  new Synchronizer(entry, office365Calendar).markSynchronized();
};

function checkEventsForAllRows() {
  if (!authorizeIfRequired()) {
    return;
  }
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var googleCalendar = new GoogleCalendar(CalendarApp.getDefaultCalendar());
  var hrCalendar = new HrCalendar(sheet);
  var office365Calendar = new Office365Calendar();
  var range = sheet.getDataRange();
  var numRows = range.getNumRows();
  var rows = range.getValues();

  for (var rowIndex = 0; rowIndex < numRows; rowIndex++) {
    var row = rows[rowIndex];
    switch (row[0]) {
      case "Type":
      case "":
        break;
      case "Bank holiday":
        var entry = new AbsenceEntry(sheet, rowIndex + 1);
        entry.clearCalendarConflict(googleCalendar.getType());
        entry.clearCalendarConflict(office365Calendar.getType());
        entry.clearCalendarConflict(hrCalendar.getType());
        break;
      default:
        var entry = new AbsenceEntry(sheet, rowIndex + 1);
        new Synchronizer(entry, googleCalendar).markSynchronized();
        new Synchronizer(entry, office365Calendar).markSynchronized();
        new Synchronizer(entry, hrCalendar).markSynchronized();
        break;
    }
  }
};

function syncEventForActiveRow() {
  if (!authorizeIfRequired()) {
    return;
  }
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowIndex = sheet.getActiveCell().getRowIndex();
  var entry = new AbsenceEntry(sheet, rowIndex);
  var googleCalendar = new GoogleCalendar(CalendarApp.getDefaultCalendar());
  var office365Calendar = new Office365Calendar();
  
  new Synchronizer(entry, googleCalendar).synchronize();
  new Synchronizer(entry, office365Calendar).synchronize();
};

function deleteEventForActiveRow() {
  if (!authorizeIfRequired()) {
    return;
  }
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowIndex = sheet.getActiveCell().getRowIndex();
  var entry = new AbsenceEntry(sheet, rowIndex);
  var googleCalendar = new GoogleCalendar(CalendarApp.getDefaultCalendar());
  var office365Calendar = new Office365Calendar();
  
  deleteEventIfRequired(googleCalendar, entry);
  deleteEventIfRequired(office365Calendar, entry);
};

function deleteEventIfRequired(calendar, entry) { 
    var event = entry.findEvent(calendar);
    if (event !== null) {
      event.deleteEvent();
    }
    
    var calendarType = calendar.getType();
    entry.setCalendarId(calendarType, null);
    entry.markCalendarConflict(calendarType);
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
                 },
                 {
                   name: "Authorize access to Office 365",
                   functionName: "authorizeIfRequired"
                 },
                 {
                   name: "Logout from Office 365",
                   functionName: "logout"
                 }
                 ];
  sheet.addMenu("Calendar", entries);
};