function getCalendars(sheet) {
  var calendars = [];
  if (Settings.getGoogleCalendarEnabled()) {
    calendars.push(new GoogleCalendar(CalendarApp.getDefaultCalendar()));
  }
  if (Settings.getOffice365CalendarEnabled()) {
    calendars.push(new Office365Calendar());
  }
  if (Settings.getHrCalendarEnabled()) {
    calendars.push(new HrCalendar(sheet));
  }
  return calendars;
};

function checkEventForActiveRow() {
  if (!authorizeIfRequired()) {
    return;
  }

  var sheet = SpreadsheetApp.getActiveSheet();
  var rowIndex = sheet.getActiveCell().getRowIndex();
  var entry = new AbsenceEntry(sheet, rowIndex);

  var calendars = getCalendars(sheet);

  for (var i = 0; i < calendars.length; i++) {
    var calendar = calendars[i];
    new Synchronizer(entry, calendar).markSynchronized();
  }
};

function checkEventsForAllRows() {
  if (!authorizeIfRequired()) {
    return;
  }

  var sheet = SpreadsheetApp.getActiveSheet();
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
        var calendars = getCalendars(sheet);
        for (var i = 0; i < calendars.length; i++) {
          var calendar = calendars[i];
          entry.clearCalendarConflict(calendar.getType());
        }
        break;
      default:
        var entry = new AbsenceEntry(sheet, rowIndex + 1);
        var calendars = getCalendars(sheet);
        for (var i = 0; i < calendars.length; i++) {
          var calendar = calendars[i];
          new Synchronizer(entry, calendar).markSynchronized();
        }
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

  var calendars = getCalendars(sheet);
  for (var i = 0; i < calendars.length; i++) {
    var calendar = calendars[i];
    new Synchronizer(entry, calendar).synchronize();
  }
};

function deleteEventForActiveRow() {
  if (!authorizeIfRequired()) {
    return;
  }
  var sheet = SpreadsheetApp.getActiveSheet();
  var rowIndex = sheet.getActiveCell().getRowIndex();
  var entry = new AbsenceEntry(sheet, rowIndex);

  var calendars = getCalendars(sheet);
  for (var i = 0; i < calendars.length; i++) {
    var calendar = calendars[i];
    deleteEventIfRequired(calendar, entry);
  }
};

function deleteEventIfRequired(calendar, entry) {
  if (calendar.isReadOnly()) {
    return;
  }
  var event = entry.findEvent(calendar);
  if (event !== null) {
    event.deleteEvent();
  }

  var calendarType = calendar.getType();
  entry.setCalendarId(calendarType, null);
  entry.markCalendarConflict(calendarType, "Event missing");
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
    }
  ];
  if (Settings.getOffice365CalendarEnabled()) {
    entries.push({
      name: "Authorize access to Office 365",
      functionName: "authorizeIfRequired"
    });
    entries.push({
      name: "Logout from Office 365",
      functionName: "logout"
    });
  }
  sheet.addMenu("Calendar", entries);
};