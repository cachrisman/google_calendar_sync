var PERSONAL_CALENDAR_ID = 'personal_email@gmail.com'
var WORK_CALENDAR_ID = 'me@work.com'
var MONTHS_IN_ADVANCE = 2
var SPREADSHEET_LOG_ID = '1abcdefghijk_lmnopqrstuvwxyzABCDEFHIJKLMNOPQ'

var BetterLogger = BetterLog.useSpreadsheet(SPREADSHEET_LOG_ID)


function syncPersonalCalendar() {
  // Define the calendar event date range to search.
  var today = new Date()
  var futureDate = new Date()
  futureDate.setMonth(futureDate.getMonth() + MONTHS_IN_ADVANCE)
  var lastRun = PropertiesService.getScriptProperties().getProperty('lastRun')
  lastRun = lastRun ? new Date(lastRun) : null

  // Clear the spreadsheet used for logging
  //if (lastRun.getDate() !== today.getDate()) {
    SpreadsheetApp.openById(SPREADSHEET_LOG_ID).getSheetByName('Log').getRange('A:A').offset(1,0).clear()
  //}

  var import_count = 0
  var update_count = 0
  var duplicate_count = 0
  var cancelled_count = 0
  var delete_count = 0

  var personal_events = findEvents(PERSONAL_CALENDAR_ID, today, futureDate)
  var work_events = findEvents(WORK_CALENDAR_ID, today, futureDate)

  var input_count = personal_events.length
  // Loop through all the personal events found and see if they exist on the work calendar
  personal_events.forEach(function(event) {
    BetterLogger.log('Evaluating:\n%s', simplify(event))
    // filter all the work events by the personal event id or if the start datetime, end datetime, summary, and status match.
    var filtered_work_events = work_events.filter(function(e) {
      ids_match = (e.id === event.id)
      details_match = (e.start.dateTime === event.start.dateTime && e.end.dateTime === event.end.dateTime && e.summary === event.summary && e.status === event.status)
      return (ids_match || details_match)
    })

    // if there is not a work event that has the same id as the personal event, import
    // the personal event into the work calendar
    if (filtered_work_events.length === 0) {
      if (event.status === "cancelled") {
        cancelled_count++
        BetterLogger.log('source event cancelled, no destination event found. skipping import')
      } else {
        event = sanitize_event(event)
        try {
          imported_event = Calendar.Events.import(event, WORK_CALENDAR_ID)
          import_count++
          BetterLogger.log('Imported event: \n%s', simplify(imported_event))
        } catch (e) {
          BetterLogger.log('Error attempting to import event:\n%s', JSON.stringify(e.toString(), null, 2))
        }
      }
    }
    // if there is one work event with the same id as the personal event, check if its start time
    // and end time are the same as the personal event
    else if (filtered_work_events.length === 1) {
      var work_event = filtered_work_events[0]
      // if source event is cancelled, delete event from destination calendar
      if (event.status === 'cancelled') {
        if (work_event.status === 'cancelled') {
          cancelled_count++
          BetterLogger.log('Source event removed. Destination event already removed:\n%s', simplify(work_event))
        } else {
          Calendar.Events.remove(WORK_CALENDAR_ID, work_event.id)
          delete_count++
          BetterLogger.log('Original event removed. Deleting event:\n%s', simplify(work_event))
        }
      }
      // if start/end time are the same, don't import as this event has already been imported
      else if (work_event.start.dateTime === event.start.dateTime && work_event.end.dateTime === event.end.dateTime) {
        duplicate_count++
        BetterLogger.log('Skipping import. Duplicate event found:\n%s', simplify(work_event))
      }
      // if start/end times are different, update work event with personal event start/end times
      else if (work_event.start.dateTime != event.start.dateTime || work_event.end.dateTime != event.end.dateTime) {
        event = sanitize_event(event)
        try {
          Calendar.Events.update(event, WORK_CALENDAR_ID, event.id)
          update_count++
          BetterLogger.log('Updated event:\n%s', simplify(event))
        } catch (e) {
          BetterLogger.log('Error attempting to import event:\n%s', JSON.stringify(e.toString(), null, 2))
        }
      }
    }
    // If multiple work events match a single personal event, something is wrong. log the event and skip
    else {
      duplicate_events = filtered_work_events.map(function(event) {return simplify(event)}).join('\n')
      duplicate_count++
      BetterLogger.log('Multiple duplicate events found. Skipping import for\n%s', duplicate_events)
    }
  })

  PropertiesService.getScriptProperties().setProperty('lastRun', today)
  BetterLogger.log('Input ' + input_count + ' events')
  BetterLogger.log('Imported ' + import_count + ' events')
  BetterLogger.log('Updated ' + update_count + ' events')
  BetterLogger.log('Skipped ' + duplicate_count + ' events')
  BetterLogger.log('Deleted ' + delete_count + ' events')
  BetterLogger.log('Already cancelled ' + cancelled_count + ' events')

  var executionTime = ((new Date()).getTime() - today.getTime()) / 1000.0
  BetterLogger.log('Total execution time: ' + executionTime + ' seconds')
}

/**
 * In a given calendar, return any such events found in the specified date range.
 * @param email the email address associated with the calendar from which to find events
 * @param start the starting Date of the range to examine.
 * @param end the ending Date of the range to examine.
 * @return an array of calendar event Objects.
 */
function findEvents(email, start, end) {
  var params = {
    timeMin: formatDate(start),
    timeMax: formatDate(end),
    singleEvents: true,
    showDeleted: true
  }

  BetterLogger.log('findEvents params %s: %s', email, JSON.stringify(params, null, 2))
  var results = []
  try {
    var response = Calendar.Events.list(email, params)
    results = response.items.filter(function(item) {
      // If the event was created by someone other than the calendar owner, only include
      // it if the calendar owner has marked it as 'accepted'.
      if (item.organizer && item.organizer.email != email) {
        // include events from gmail
        if (item.organizer.email === 'unknownorganizer@calendar.google.com') {
          return true
        }
        if (!item.attendees) {
          return false
        }
        var matching = item.attendees.filter(function(attendee) {
          return attendee.self
        })
        return matching.length > 0 && matching[0].status == 'accepted'
      }
      // filter all day events
      if (item.start.hasOwnProperty('date')) {
        return false
      }
      return true
    })
  } catch (e) {
    BetterLogger.log('Error retriving events for %s: %s; skipping', email, e.toString())
    results = []
  }
  BetterLogger.log('findEvents found %s results for %s:', results.length, email)
  results.forEach(function(event) {
    BetterLogger.log([event.summary, event.id, event.start.dateTime, event.end.dateTime, event.status].join('\t'))
  })
  return results
}

/**
 * Return an RFC3339 formated date String corresponding to the given
 * Date object.
 * @param date a Date.
 * @return a formatted date string.
 */
function formatDate(date) {
  return Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd\'T\'HH:mm:ssZ')
}

/**
 * Sanitize an event to to be added to your work calendar.
 * Remove attendees, disable reminders, mark as private, add a tag to the description to help with identification
 * @param event - a calendar event object
 * @return a sanitized event object.
 */
function sanitize_event(event) {
  event.organizer = {
    id: WORK_CALENDAR_ID,
    self: true
  }
  event.attendees = []
  event.reminders = []
  event.reminders.useDefault = false
  event.visibility = 'private'
  event.description = event.description + '\n\nImported with Personal Calendar Sync'
  return event
}

/**
 * Simplify an object for logging
 */
function simplify(obj) {
  return JSON.stringify({
    id: obj.id,
    summary: obj.summary,
    start: obj.start.dateTime,
    end: obj.end.dateTime,
    updated: obj.updated,
    status: obj.status
  }, null, 2)
}
