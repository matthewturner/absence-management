var Synchronizer = function(entry, calendar) {
  this.entry = entry;
  this.calendar = calendar;
  
  this.areSynchronized = function() {
    var calendarType = this.calendar.getType();
    var event = this.entry.findEvent(calendar);
    if (event === null) {
      Logger.log("Event missing");
      return false;
    }
    if (this.calendar.supportsId()) {
      if (this.entry.getCalendarId(calendarType) !== event.getId()) {
        Logger.log("Calendar id mismatch");
        return false;
      }
    }
    if (this.entry.getTitle() !== event.getTitle()) {
      Logger.log("Title mismatch");
      return false;
    }
    if (event.getStartTime().getTime() !== this.entry.getStartTime().getTime()) {
      Logger.log("Start date mismatch: " + event.getStartTime().getTime() + " vs " + this.entry.getStartTime().getTime());
      return false;
    }
    
    var eventEndTime = event.getEndTime();
    var entryEndTime = this.entry.getAdjustedEndTime(this.calendar.getAdjustment());
    if (eventEndTime.getTime() !== entryEndTime.getTime()) {
      Logger.log("End date mismatch: " + eventEndTime.getTime() + " vs " + entryEndTime.getTime());
      Logger.log(eventEndTime);
      Logger.log(entryEndTime);
      return false;
    }
    return true;
  };
  
  this.markSynchronized = function() {
    var calendarType = this.calendar.getType();
    if (this.areSynchronized()) {
      this.entry.clearCalendarConflict(calendarType);
    } else {
      this.entry.markCalendarConflict(calendarType);
    }
  };
  
  this.synchronize = function () {
    var calendarType = this.calendar.getType();
    
    var event = this.entry.findEvent(calendar);
    if (event === null) {
      var event = this.calendar.createEvent(this.entry.getTitle(), this.entry.getStartTime(), this.entry.getAdjustedEndTime(this.calendar.getAdjustment()));
      this.entry.setCalendarId(calendarType, event.getId());
      this.entry.clearCalendarConflict(calendarType);
      return;
    }
    
    if (this.entry.getCalendarId(calendarType) !== event.getId()) {
      this.entry.setCalendarId(calendarType, event.getId());
    }
    
    if (this.entry.getTitle() !== event.getTitle()) {
      event.setTitle(this.entry.getTitle());
    }
    
    if (event.getStartTime().getTime() !== this.entry.getStartTime().getTime() || event.getEndTime().getTime() !== this.entry.getAdjustedEndTime(this.calendar.getAdjustment()).getTime()) {
      event.setTime(this.entry.getStartTime(), this.entry.getAdjustedEndTime(this.calendar.getAdjustment()));
    }
    
    this.entry.clearCalendarConflict(calendarType);
  };
};