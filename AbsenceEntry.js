var AbsenceEntry = function (sheet, rowIndex) {
    var GOOGLE_CALENDAR_COLUMN_INDEX = 8;
    var HR_CALENDAR_COLUMN_INDEX = 10;
    var NUMBER_OF_ROWS = 1;
    this.sheet = sheet;
    this.rowIndex = rowIndex;
    this.dataRange = sheet.getRange(rowIndex, 1, NUMBER_OF_ROWS, 5);
    this.currentRow = this.dataRange.getValues()[0];
    
    this.getTitle = function () {
      return this.currentRow[0];
    };
    
    this.getCalendarCell = function (calendarType) {
      switch (calendarType) {
        case "google":
          return this.sheet.getRange(rowIndex, GOOGLE_CALENDAR_COLUMN_INDEX);
        case "hr":
          return this.sheet.getRange(rowIndex, HR_CALENDAR_COLUMN_INDEX);
      }   
    };
    
    this.getCalendarId = function (calendarType) {
      var calendarCell = this.getCalendarCell(calendarType);
      var calendarId = calendarCell.getValue();
      return calendarId;
    };
    
    this.setCalendarId = function (calendarType, calendarId) {
      var calendarCell = this.getCalendarCell(calendarType);
      calendarCell.setValue(calendarId);
    };
    
    this.getStartTime = function () {
      var tstart = new Date(this.currentRow[2]);
      return tstart;
    };
    
    this.getEndTime = function () {
      var tstop = new Date(this.currentRow[4]);
      return tstop;
    };
    
    this.getAdjustedEndTime = function (adjustment) {
      var entryEndTime = this.getEndTime();
      if (adjustment > 0) {
        entryEndTime.setDate(entryEndTime.getDate() + adjustment);
      }
      return entryEndTime;
    };
    
    this.markCalendarConflict = function (calendarType) {
      var calendarCell = this.getCalendarCell(calendarType);
      calendarCell.setBackground("Red");
    };
    
    this.clearCalendarConflict = function (calendarType) {
      var calendarCell = this.getCalendarCell(calendarType);
      calendarCell.setBackground("Green");
    };
    
    this.configure = function () {
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
    
    this.findEvent = function (calendar) {
      var calendarType = calendar.getType();
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