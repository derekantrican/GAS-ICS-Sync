/*
*=========================================
*          SETTINGS AND FILTERS
*=========================================
*/
const appSettings = {
  howFrequent: 15,                         //What interval (minutes) to run this script to check for new/modified events.  Any integer can be used, but will be rounded up to 5, 10, 15, 30 or to the nearest hour after that.. 60, 120, etc. 1440 (24 hours) is the maximum value.  Anything above that will be replaced with 1440.
  email: "",                               //If Email Summary is true or you want to receive update notifications, you will need to provide your email address.
  emailSummary: false,                     //Will email you when an event is added/modified/removed to/from your calendar.
  customEmailSubject: "",                  //If you want to change the email subject, provide a custom one here. Default: 'GAS-ICS-Sync Execution Summary'.
  dateFormat: ""                           //Custom date format in the email summary (e.g. 'YYYY-MM-DD', 'DD.MM.YYYY', 'MM/DD/YYYY'. Separators are '.', '-' and '/').
};

/*
*=========================================
*         Default Sync Settings
*=========================================
*These settings apply to all calendars unless overriden with calendar-specific settings below.
*/

const defaultSettings = {
  color: null,
  addEventsToCalendar: true,
  modifyExistingEvents: true,
  removePastEventsFromCalendar: true,
  removeEventsFromCalendar: true,
  addOrganizerToTitle: false,
  descriptionAsTitles: false,
  addCalToTitle: false,
  addAlerts: "default",
  addAttendees: false,
  defaultAllDayReminder: null,
  overrideVisibility: "",
  addTasks: false,
  filters: [],
  syncDelayDays: null 
};


/*
*=========================================
*          Calendar Settings
*=========================================
EXAMPLE:
sourceCalendars = [
  {
    //The three fields listed below are required fields for every calendar.  sourceCalendarName must only be used once in the script.
    sourceCalendarName: "Work Calendar",  This is a friendly name for the source calendar you want the script to sync
    sourceURL: "http://someURLforWork.com/someCalendar.ics",   This is the URL of the source calendar
    targetCalendarName: "My Work Calendar", This is the name of the target calendar in Google you want to sync to.  For your personal calendar, use your email address (xyz@gmail.com)

    //Additional properties below can be set from the default list above but customized to be calendar-specific.  These are optional.
    color: 5,
    filers: ['onlyFutureEvents'] //add a comma-separated list of all filter-ids you want to apply
  }  // use a comma after } if there are other calendars.  No comma on last entry.
];
*/


const sourceCalendars = [
  {
  sourceCalendarName: "Calendar x",                    //Required
  sourceURL: "http://CalXURL.ics",                     //Required
  targetCalendarName: "Target Calendar",               //Required
  color: 5
  },
  {
  sourceCalendarName: "Calendar y",                    //Required
  sourceURL: "http://CalYURL.ics",                     //Required
  targetCalendarName: "Target Calendar",               //Required
  color: 5,
  addCalToTitle: false
  }
];

/*
*=========================================
*             Event Filters
*=========================================
Filters for calendar events based on ical properties (RFC 5545).
Define each filter with the following structure and add them to the filters object below:
'uniqueID': {
              parameter: "property",      // Event property to filter by (e.g., "summary", "categories", "dtend", "dtstart").
              type: "include/exclude",    // Whether to include or exclude events matching the criteria.
              comparison: "method",       // Comparison method: "equals", "begins with", "contains", "regex", "<", ">".
                                          // Note: "<", ">" only apply for date/time properties.
              criterias: ["values"],      // Array of values or patterns for comparison.
              offset: number              // (Optional) For date/time properties, specify an offset in days.
            }
*/
const filters = {
  'onlyConfirmed': {
                      parameter: "summary",       // Exclude events whose summary starts with "Pending:" or contains "cancelled".
                      type: "exclude",
                      comparison: "regex",
                      criterias: ["^Pending:", "cancelled"]
                    },
  'onlyMeetings': {
                    parameter: "categories",    // Include only events categorized as "Meetings".
                    type: "include",
                    comparison: "equals",
                    criterias: ["Meetings"]
                  },
  'onlyFutureEvents': {
                        parameter: "dtend",       // Reproduce the deprecated onlyFutureEvents behaviour.
                        type: "include",
                        comparison: ">",
                        offset: 0
                      },
  'x+14': {
            parameter: "dtstart",       // Exclude events starting more than 14 days from now.
            type: "exclude",
            comparison: ">",
            offset: 14
          }
};
