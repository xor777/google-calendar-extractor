// add calendar API and Admin SDK API service before running

function listManagerEventsForPeriod() {
  // var startDate = '2024-05-06';
  // var endDate = '2024-05-10';

  var managers = [
    'medy.hassaki@p2p.org',
    'jourdan.parkinson@p2p.org',
    'divya.subramaniam@p2p.org',
    'mars.leung@p2p.org',
    'ergun.arduc@p2p.org',
    'd.shmeman@p2p.org',
    'niny.yang@p2p.org',
    'lester.chui@p2p.org',
    'joshua.betancourt@p2p.org'
  ];

  var ss = SpreadsheetApp.openById('1GNG0eBH4VG0ceCrZDyoTqGN4N7gzaisCl3JORjbyHI0');
  var sheet = ss.getSheets()[0]; // first sheet in the spreadsheet

  sheet.clear();
  sheet.appendRow(["Date", "Manager", "Duration", "Participants", "Event Title"]);

  var today = new Date();
  var day = today.getDay();
  var prevMonday = new Date(today);
  prevMonday.setDate(today.getDate() - day - 6);
  var prevFriday = new Date(prevMonday);
  prevFriday.setDate(prevMonday.getDate() + 4);

  var startDateFormatted = Utilities.formatDate(prevMonday, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var endDateFormatted = Utilities.formatDate(prevFriday, Session.getScriptTimeZone(), "yyyy-MM-dd");

  // Log the date range
  Logger.log('Fetching events from ' + startDateFormatted + ' to ' + endDateFormatted);

  managers.forEach(function(manager) {
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
