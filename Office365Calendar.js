function getGraphService() {
  // Create a new service with the given name. The name will be used when
  // persisting the authorized token, so ensure it is unique within the
  // scope of the property store.
  return OAuth2.createService('graph')
      .setAuthorizationBaseUrl("https://login.microsoftonline.com/" + getTenantId() + "/oauth2/v2.0/authorize")
      .setTokenUrl("https://login.microsoftonline.com/" + getTenantId() + "/oauth2/v2.0/token")

      // Set the client ID and secret, from the Google Developers Console.
      .setClientId(getClientId())
      .setClientSecret(getSecret())

      // Set the name of the callback function in the script referenced
      // above that should be invoked to complete the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())
      .setScope('https://outlook.office.com/calendars.readwrite');
};

function logout() {
  var service = getGraphService()
  service.reset();
};

function authorizeIfRequired() {
  var graphService = getGraphService();
  if (graphService.hasAccess()) {
    return true;
  } else {
    var authorizationUrl = graphService.getAuthorizationUrl();
    var template = HtmlService.createTemplate(
        '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
        'Reopen the sidebar when the authorization is complete.');
    template.authorizationUrl = authorizationUrl;
    var page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
    return false;
  }
};

function authCallback(request) {
  var graphService = getGraphService();
  var isAuthorized = graphService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
};

function makeRequest(url, options) {
  var graphService = getGraphService();
  
  if (options === undefined) {
    options = {};
  }
  
  options.headers = {
      Authorization: 'Bearer ' + graphService.getAccessToken()
  };
  
  var response = UrlFetchApp.fetch("https://outlook.office.com/api/v2.0/me/" + url, options);
  return response;
};

function test() {
  var response = makeRequest("calendarview?startDateTime=2019-12-03&endDateTime=2019-12-04&$select=Subject,Start,End,IsAllDay,ShowAs");
  var x = response.getContentText();
  var y = JSON.parse(x);
  var z = 1;
};

function test2() {
  var calendar = new Office365Calendar();
  var events = calendar.getEvents(new Date("2020-01-02"), new Date("2020-01-03"));
  var event = events[0];
  var sd = event.getStartTime();
  var d = new Date(sd);
  var y = new Date("2020-01-01T00:00:00.0000000");
  // var eventById = calendar.getEventById(event.getId());
  var eventByMissingId = calendar.getEventById("xxx");
  var x = 1;
};

function formatDate(date) {
  return date.getFullYear()  + "-" + (date.getMonth() + 1) + "-" + date.getDate()
};

var Office365Calendar = function() {  
  this.createEvent = function(title, startTime, endTime) {
    return null;
  }
  
  this.getEventById = function(calendarId) {
    if (calendarId === null) {
      return null;
    }
    if (calendarId === "") {
      return null;
    }
    
    try {
      var response = makeRequest("events/" + calendarId);
      var content = response.getContentText();
      var eventData = JSON.parse(content);
      return new Office365CalendarEvent(eventData);
    } catch(error) {
      return null;
    }
  };
  
  this.getEvents = function(startTime, endTime) {
    var response = makeRequest("calendarview?startDateTime=" + formatDate(startTime) + "&endDateTime=" + formatDate(endTime) + "&$select=Subject,Start,End,IsAllDay,ShowAs");
    var content = response.getContentText();
    var eventData = JSON.parse(content).value;
    var events = [];
    for(var i = 0; i < eventData.length; i++) {
      events.push(new Office365CalendarEvent(eventData[i]));
    }
    return events;
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
    return "office365";
  };
};

var Office365CalendarEvent = function(data) {
  this.data = data;
  
  this.getId = function() {
    return this.data.Id;
  };
  
  this.getTitle = function() {
    return this.data.Subject;
  };
  
  this.getStartTime = function() {
    return new Date(this.data.Start.DateTime.substring(0, 19));
  };
  
  this.getEndTime = function() {
    return new Date(this.data.End.DateTime.substring(0, 19));
  };
  
  this.setTime = function(startTime, endTime) {
    var payload = { 
            "Start": {
              "DateTime": formatDate(startTime),
              "TimeZone": "UTC"
            },
            "End": {
              "DateTime": formatDate(endTime),
              "TimeZone": "UTC"
            }
          };
    var options =
        {
          "method" : "patch",
          "contentType" : "application/json",
          "payload" : JSON.stringify(payload)
        };

    makeRequest("events/" + this.getId(), options);
  };
};