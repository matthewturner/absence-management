var GoogleCalendar = function(calendar) {
    this.calendar = calendar;
    
    this.createEvent = function(title, startTime, endTime) {
      return this.calendar.createEvent(title, startTime, endTime);
    }
    
    this.getEventById = function(calendarId) {
      return this.calendar.getEventById(calendarId);
    };
    
    this.getEvents = function(startTime, endTime) {
      return this.calendar.getEvents(startTime, endTime);
    };
    
    this.requiresDayAdjustment = function() {
      return true;
    };
    
    this.supportsId = function() {
      return true;
    };
    
    this.getAdjustment = function() {
      return 1;
    };
    
    this.getType = function() {
      return "google";
    };
  };  