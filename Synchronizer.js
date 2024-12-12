var Synchronizer = function (entry, calendar) {
  this.entry = entry;
  this.calendar = calendar;

  this.areSynchronized = function () {
    var conflicts = this.getConflicts();
    return conflicts.length > 0;
  };

  this.getConflicts = function () {
    var calendarType = this.calendar.getType();
    var event = this.entry.findEvent(calendar);
    var conflicts = [];
    if (event === null) {
      Logger.log("Event missing");
      conflicts.push("Event missing");
      return conflicts;
    }
    if (this.calendar.supportsId()) {
      if (this.entry.getCalendarId(calendarType) !== event.getId()) {
        Logger.log("Calendar id mismatch");
        conflicts.push("Calendar id mismatch");
      }
    }
    if (this.entry.getTitle() !== event.getTitle()) {
      Logger.log("Title mismatch");
      conflicts.push("Title mismatch");
    }
    if (
      event.getStartTime().getTime() !== this.entry.getStartTime().getTime()
    ) {
      var conflict =
        "Start date mismatch: " +
        event.getStartTime().getTime() +
        " vs " +
        this.entry.getStartTime().getTime();
      Logger.log(conflict);
      conflicts.push(conflict);
    }

    var eventEndTime = event.getEndTime();
    var entryEndTime = this.entry.getAdjustedEndTime(
      this.calendar.getAdjustment()
    );
    if (eventEndTime.getTime() !== entryEndTime.getTime()) {
      var conflict =
        "End date mismatch: " +
        eventEndTime.getTime() +
        " vs " +
        entryEndTime.getTime();
      Logger.log(conflict);
      conflicts.push(conflict);
    }
    return conflicts;
  };

  this.markSynchronized = function () {
    var calendarType = this.calendar.getType();
    var conflicts = this.getConflicts();
    if (conflicts.length === 0) {
      this.entry.clearCalendarConflict(calendarType);
    } else {
      this.entry.markCalendarConflict(calendarType, conflicts.join("\\n"));
    }
  };

  this.synchronize = function () {
    var calendarType = this.calendar.getType();

    var event = this.entry.findEvent(calendar);
    if (event === null) {
      if (calendar.isReadOnly()) {
        this.entry.markCalendarConflict(calendarType, "Event missing");
      } else {
        var event = this.calendar.createEvent(
          this.entry.getTitle(),
          this.entry.getStartTime(),
          this.entry.getAdjustedEndTime(this.calendar.getAdjustment())
        );
        this.entry.setCalendarId(calendarType, event.getId());
        this.entry.clearCalendarConflict(calendarType);
      }
      return;
    }

    if (this.entry.getCalendarId(calendarType) !== event.getId()) {
      this.entry.setCalendarId(calendarType, event.getId());
    }

    if (this.entry.getTitle() !== event.getTitle()) {
      event.setTitle(this.entry.getTitle());
    }

    if (
      event.getStartTime().getTime() !== this.entry.getStartTime().getTime() ||
      event.getEndTime().getTime() !==
        this.entry.getAdjustedEndTime(this.calendar.getAdjustment()).getTime()
    ) {
      event.setTime(
        this.entry.getStartTime(),
        this.entry.getAdjustedEndTime(this.calendar.getAdjustment())
      );
    }

    this.entry.clearCalendarConflict(calendarType);
  };
};
