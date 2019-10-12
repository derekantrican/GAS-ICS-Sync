/*
*=========================================
*               SETTINGS
*=========================================
*/
var targetCalendarName = "Full API TEST";           // The name of the Google Calendar you want to add events to
var sourceCalendarURLs = [""
                         ];            // The ics/ical urls that you want to get events from ["url","url","url"]

var howFrequent = 15;                  // What interval (minutes) to run this script on to check for new events
var addEventsToCalendar = true;        // If you turn this to "false", you can check the log (View > Logs) to make sure your events are being read correctly before turning this on
var modifyExistingEvents = true;       // If you turn this to "false", any event in the feed that was modified after being added to the calendar will not update
var removeEventsFromCalendar = true;   // If you turn this to "true", any event created by the script that is not found in the feed will be removed.
var addAlerts = true;                  // Whether to add the ics/ical alerts as notifications on the Google Calendar events, this will override the standard reminders specified by the target calendar.
var addOrganizerToTitle = false;       // Whether to prefix the event name with the event organiser for further clarity 
var addCalToTitle = false;             // Whether to add the source calendar to title

var addTasks = false;

var emailWhenAdded = false;            // Will email you when an event is added to your calendar
var emailWhenModified = false;         // Will email you when an existing event is updated in your calendar
var email = "";                        // OPTIONAL: If "emailWhenAdded" is set to true, you will need to provide your email
