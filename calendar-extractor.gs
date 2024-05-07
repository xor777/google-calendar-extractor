// add calendar API and Admin SDK API service before running

function listManagerEventsForPeriod() {
  var startDate = '2024-04-29';
  var endDate = '2024-05-05';

  var managers = [
  'email@p2p.org'
  ];

  var ss = SpreadsheetApp.openById('sheet-id-here');
  var sheet = ss.getSheets()[0]; // first sheet in the spreadsheet
  
  sheet.clear();
  sheet.appendRow(["Date", "Manager", "Duration", "Participants", "Event Title"]);
  
  managers.forEach(function(manager) {
    var startDateFormatted = Utilities.formatDate(new Date(startDate), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var endDateFormatted = Utilities.formatDate(new Date(endDate), Session.getScriptTimeZone(), "yyyy-MM-dd");
    fillSheetWithManagerEventsForPeriod(manager, startDateFormatted, endDateFormatted, sheet);
  });
}

function fillSheetWithManagerEventsForPeriod(calendarId, startDate, endDate, sheet) {
  var events = Calendar.Events.list(calendarId, {
    timeMin: startDate + 'T00:00:00Z',
    timeMax: endDate + 'T23:59:59Z',
    singleEvents: true,
    orderBy: 'startTime'
  });
  
  if (events.items && events.items.length > 0) {
    events.items.forEach(function(event) {
      var startString = event.start.dateTime;
      var endString = event.end.dateTime;
      
      if (startString && endString) {
        var start = new Date(startString);
        var end = new Date(endString);
        var duration = (end - start) / (1000 * 60 * 60); // Duration in hours
        var participants = event.attendees ? event.attendees.map(function(attendee) { return attendee.email; }).join(', ') : '';
        var eventTitle = event.summary || ''; 
        
        var hasExternalParticipant = event.attendees && event.attendees.some(function(attendee) {
          return attendee.email.split('@')[1] !== 'p2p.org';
        });
        
        var isNotCancelled = event.status !== 'cancelled';
        
        if (hasExternalParticipant && isNotCancelled) {
          sheet.appendRow([start.toISOString().slice(0, 10), calendarId, duration.toFixed(2), participants, eventTitle]);
          Logger.log('Processed event with external participant(s): ' + eventTitle + ' for calendar: ' + calendarId);
        } else if (isNotCancelled) {
          Logger.log('Skipped event without external participants: ' + eventTitle + ' for calendar: ' + calendarId);
        } else {
          Logger.log('Skipped cancelled event: ' + eventTitle + ' for calendar: ' + calendarId);
        }
      }
    });
  } else {
    Logger.log('No events found for calendar: ' + calendarId);
  }
}
