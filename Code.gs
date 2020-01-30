/*
*=========================================
*       INSTALLATION INSTRUCTIONS
*=========================================
*
* 1) Click in the menu "File" > "Make a copy..." and make a copy to your Google Drive
* 2) Changes lines 19-41 to be the settings that you want to use
* 3) Click in the menu "Run" > "Run function" > "install" and authorize the program
*    (For steps to follow in authorization, see this video: https://youtu.be/_5k10maGtek?t=1m22s )
* 4) You can also run "startSync" if you want to sync only once.
*
* **To stop the Script from running click in the menu "Run" > "Run function" > "uninstall"
*
*=========================================
*               SETTINGS
*=========================================
*/

var sourceCalendars = [                // The ics/ical urls that you want to get events from along with their target calendars (list a new row for each mapping of ICS url to Google Calendar)
                                       // For instance: ["https://www.calendarlabs.com/ical-calendar/ics/76/US_Holidays.ics", "US Holidays"]
  ["icsUrl1", "targetCalendar1"],
  ["icsUrl2", "targetCalendar2"],
  ["icsUrl3", "targetCalendar1"]
  
];

var howFrequent = 15;                  // What interval (minutes) to run this script on to check for new events
var onlyFutureEvents = false;          // If you turn this to "true", past events will not be synced (this will also removed past events from the target calendar if removeEventsFromCalendar is true)
var addEventsToCalendar = true;        // If you turn this to "false", you can check the log (View > Logs) to make sure your events are being read correctly before turning this on
var modifyExistingEvents = true;       // If you turn this to "false", any event in the feed that was modified after being added to the calendar will not update
var removeEventsFromCalendar = true;   // If you turn this to "true", any event created by the script that is not found in the feed will be removed.
var addAlerts = true;                  // Whether to add the ics/ical alerts as notifications on the Google Calendar events, this will override the standard reminders specified by the target calendar.
var addOrganizerToTitle = false;       // Whether to prefix the event name with the event organiser for further clarity
var descriptionAsTitles = false;       // Whether to use the ics/ical descriptions as titles (true) or to use the normal titles as titles (false)
var addCalToTitle = false;             // Whether to add the source calendar to title
var addAttendees = false;              // Whether to add the attendee list. If true, duplicate events will be automatically added to the attendees' calendar.
var defaultAllDayReminder = -1;        // Default reminder for all day events in minutes before the day of the event (-1 = no reminder, the value has to be between 0 and 40320)
                                       // See https://github.com/derekantrican/GAS-ICS-Sync/issues/75 for why this is neccessary.
var addTasks = false;

var emailWhenAdded = false;            // Will email you when an event is added to your calendar
var emailWhenModified = false;         // Will email you when an existing event is updated in your calendar
var email = "";                        // OPTIONAL: If "emailWhenAdded" is set to true, you will need to provide your email

/*
*=========================================
*           ABOUT THE AUTHOR
*=========================================
*
* This program was created by Derek Antrican
*
* If you would like to see other programs Derek has made, you can check out
* his website: derekantrican.com or his github: https://github.com/derekantrican
*
*=========================================
*            BUGS/FEATURES
*=========================================
*
* Please report any issues at https://github.com/derekantrican/GAS-ICS-Sync/issues
*
*=========================================
*           $$ DONATIONS $$
*=========================================
*
* If you would like to donate and help Derek keep making awesome programs,
* you can do that here: https://paypal.me/derekantrican
*
*=========================================
*             CONTRIBUTORS
*=========================================
* Andrew Brothers
* Github: https://github.com/agentd00nut
* Twitter: @abrothers656
*
* Joel Balmer
* Github: https://github.com/JoelBalmer
*
* Blackwind
* Github: https://github.com/blackwind
*
* Jonas Geissler
* Github: https://github.com/jonas0b1011001
*/


//=====================================================================================================
//!!!!!!!!!!!!!!!! DO NOT EDIT BELOW HERE UNLESS YOU REALLY KNOW WHAT YOU'RE DOING !!!!!!!!!!!!!!!!!!!!
//=====================================================================================================
function install(){
  try{
    //Delete any already existing triggers so we don't create excessive triggers
    deleteAllTriggers();
    
    if (howFrequent < 1){
      throw "[ERROR] \"howFrequent\" must be greater than 0.";
    }
    else{
      ScriptApp.newTrigger("install").timeBased().after(howFrequent * 60 * 1000).create();//Schedule next Execution
      ScriptApp.newTrigger("startSync").timeBased().after(1000).create();//Start the sync routine
    }
  }catch(e){
    install();//Retry on error
  }
}

function uninstall(){
  deleteAllTriggers();
}

var targetCalendarId;
var targetCalendarName;
var calendarEvents = [];
var calendarEventsIds = [];
var icsEventsIds = [];
var calendarEventsMD5s = [];
var recurringEvents = [];
var startUpdateTime;

function startSync(){
  checkForUpdate();
  
  if (onlyFutureEvents)
    startUpdateTime = new ICAL.Time.fromJSDate(new Date());
  
  //Disable email notification if no mail adress is provided 
  emailWhenAdded = emailWhenAdded && email != "";
  
  sourceCalendars = condenseCalendarMap(sourceCalendars);
  for each (var calendar in sourceCalendars){
    calendarEvents = [];
    targetCalendarName = calendar[0];
    var sourceCalendarURLs = calendar[1];
    var vevents;
    //------------------------ Fetch URL items ------------------------
    var responses = fetchSourceCalendars(sourceCalendarURLs);
    Logger.log("Syncing " + responses.length + " calendars to " + targetCalendarName);
    
    //------------------------ Get target calendar information------------------------
    var targetCalendar = setupTargetCalendar(targetCalendarName);
    targetCalendarId = targetCalendar.id;
    Logger.log("Working on calendar: " + targetCalendarId);
    
    //------------------------ Parse existing events --------------------------
    if(addEventsToCalendar || modifyExistingEvents || removeEventsFromCalendar){
      var eventList = Calendar.Events.list(targetCalendarId, {showDeleted: false, privateExtendedProperty: "fromGAS=true", maxResults: 2500});
      calendarEvents = [].concat(calendarEvents, eventList.items);
      //loop until we received all events
      while(typeof eventList.nextPageToken !== 'undefined'){
        eventList = callWithBackoff(function(){
          return Calendar.Events.list(targetCalendarId, {showDeleted: false, privateExtendedProperty: "fromGAS=true", maxResults: 2500, pageToken: eventList.nextPageToken});
        }, 2);

        if (eventList != null)
          calendarEvents = [].concat(calendarEvents, eventList.items);
      }
      Logger.log("Fetched " + calendarEvents.length + " existing events from " + targetCalendarName);
      for (var i = 0; i < calendarEvents.length; i++){
        if (calendarEvents[i].extendedProperties != null){
          calendarEventsIds[i] = calendarEvents[i].extendedProperties.private["rec-id"] || calendarEvents[i].extendedProperties.private["id"];
          calendarEventsMD5s[i] = calendarEvents[i].extendedProperties.private["MD5"];
        }
      }

      //------------------------ Parse ical events --------------------------
      vevents = parseResponses(responses, icsEventsIds);
      Logger.log("Parsed " + vevents.length + " events from ical sources");
    }
    
    //------------------------ Process ical events ------------------------
    if (addEventsToCalendar || modifyExistingEvents){
      Logger.log("Processing " + vevents.length + " events");
      var calendarTz = Calendar.Settings.get("timezone").value;
      
      vevents.forEach(function(e){
        processEvent(e, calendarTz);
      });

      Logger.log("Done processing events");
    }
    
    //------------------------ Remove old events from calendar ------------------------
    if(removeEventsFromCalendar){
      Logger.log("Checking " + calendarEvents.length + " events for removal");
      processEventCleanup();
      Logger.log("Done checking events for removal");
    }

    //------------------------ Process Tasks ------------------------
    if (addTasks){
      processTasks(responses);
    }

    //------------------------ Add Recurring Event Instances ------------------------
    Logger.log("Processing " + recurringEvents.length + " Recurrence Instances!");
    for each (var recEvent in recurringEvents){
      processEventInstance(recEvent);
    }
  }

  Logger.log("Sync finished!");
}