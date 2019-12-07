var HrCalendar = function (sheet) {
  this.hrSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet.getSheetName() + " - HR");

  this.getAllEvents = function () {
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

  this.allEvents = this.getAllEvents();

  this.getEventById = function (calendarId) {
    // not supported
    return null;
  };

  this.getEvents = function (startTime, endTime) {
    return this.allEvents.filter(function (item) {
      if (item.getStartTime().getTime() !== startTime.getTime()) {
        return false;
      }
      if (item.getEndTime().getTime() !== endTime.getTime()) {
        return false;
      }
      return true;
    });
  };

  this.requiresDayAdjustment = function () {
    return false;
  };

  this.supportsId = function () {
    return false;
  };

  this.getAdjustment = function () {
    return 0;
  };

  this.getType = function () {
    return "hr";
  };

  this.isReadOnly = function () {
    return true;
  };
}

var HrCalendarEvent = function (row) {
  this.row = row;

  this.getId = function () {
    return "";
  };

  this.getTitle = function () {
    return this.row[2];
  };

  this.getStartTime = function () {
    return new Date(this.row[4]);
  };

  this.getEndTime = function () {
    return new Date(this.row[5])
  };
};