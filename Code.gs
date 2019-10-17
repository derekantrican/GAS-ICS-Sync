/*
*=========================================
*       INSTALLATION INSTRUCTIONS
*=========================================
*
* 1) Click in the menu "File" > "Make a copy..." and make a copy to your Google Drive
* 2) Changes lines 19-37 to be the settings that you want to use
* 3) Click in the menu "Run" > "Run function" > "Install" and authorize the program
*    (For steps to follow in authorization, see this video: https://youtu.be/_5k10maGtek?t=1m22s )
* 4) You can also run "startSync" if you want to sync only once.
*
* **To stop Script from running click in the menu "Run" > "Run function" > "Uninstall"
*
*=========================================
*               SETTINGS
*=========================================
*/

var sourceCalendars = [                // The ics/ical urls that you want to get events from
  ["targetCalendar",[""]]              //[["targetCalendar1",["url","url"]], ["targetCalendar2",["url","url"]]]
];

var howFrequent = 15;                  // What interval (minutes) to run this script on to check for new events
var onlyFutureEvents = false;          // If you turn this to "true", past events will not be synced (this will also removed past events from the target calendar if removeEventsFromCalendar is true)
var addEventsToCalendar = true;        // If you turn this to "false", you can check the log (View > Logs) to make sure your events are being read correctly before turning this on
var modifyExistingEvents = true;       // If you turn this to "false", any event in the feed that was modified after being added to the calendar will not update
var removeEventsFromCalendar = true;   // If you turn this to "true", any event created by the script that is not found in the feed will be removed.
var addAlerts = true;                  // Whether to add the ics/ical alerts as notifications on the Google Calendar events, this will override the standard reminders specified by the target calendar.
var addOrganizerToTitle = false;       // Whether to prefix the event name with the event organiser for further clarity 
var addCalToTitle = false;             // Whether to add the source calendar to title
var addAttendees = false;              // Whether to add the attendee list. If true, duplicate events will be automatically added to the attendees' calendar.

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
* you can do that here: https://bulkeditcalendarevents.wordpress.com/donate/
*
*=========================================
*             CONTRIBUTORS
*=========================================
* Andrew Brothers
* Github: https://github.com/agentd00nut/
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
function Install(){
  //Delete any already existing triggers so we don't create excessive triggers
  DeleteAllTriggers();

  if (howFrequent < 1){
    throw "[ERROR] \"howFrequent\" must be greater than 0.";
  }
  else{
    ScriptApp.newTrigger("Install").timeBased().after(howFrequent * 60 * 1000).create();//Shedule next Execution
    ScriptApp.newTrigger("startSync").timeBased().after(1).create();//Start the sync routine
  }
}

function Uninstall(){
  DeleteAllTriggers();
}

function startSync(){
  if (onlyFutureEvents)
    var startUpdateTime = new ICAL.Time.fromJSDate(new Date());
  
  for each (var calendar in sourceCalendars){
    var targetCalendarName = calendar[0];
    var sourceCalendarURLs = calendar[1];
    //------------------------ Fetch URL items ------------------------
    var responses = fetchSourceCalendars(sourceCalendarURLs);
    Logger.log("Syncing " + responses.length + " Calendars to " + targetCalendarName);
    
    //------------------------ Get target calendar information------------------------
    var targetCalendar = setupTargetCalendar(targetCalendarName);
    var targetCalendarId = targetCalendar.id;
    Logger.log("Working on calendar: " + targetCalendarId);
    
    //Disable email notification if no mail adress is provided 
    emailWhenAdded = (emailWhenAdded && email != "")
    
    //------------------------ Parse existing events --------------------------
    if(addEventsToCalendar || modifyExistingEvents || removeEventsFromCalendar){ 
      var calendarEvents = Calendar.Events.list(targetCalendarId, {showDeleted: false, privateExtendedProperty: "fromGAS=true"}).items;
      var calendarEventsIds = [] 
      var calendarEventsMD5s = []
      Logger.log("Fetched " + calendarEvents.length + " existing Events from " + targetCalendarName); 
      for (var i = 0; i < calendarEvents.length; i++){
        if (calendarEvents[i].extendedProperties != null){
          calendarEventsIds[i] = calendarEvents[i].extendedProperties.private["rec-id"] || calendarEvents[i].extendedProperties.private["id"];
          calendarEventsMD5s[i] = calendarEvents[i].extendedProperties.private["MD5"];
        }
      }
      //------------------------ Parse ical events --------------------------
      var icsEventsIds = [];
      var vevents = parseResponses(responses, icsEventsIds);
      Logger.log("Parsed " + vevents.length + " events from ical sources.");
    }
    
    //------------------------ Process ical events ------------------------
    if (addEventsToCalendar || modifyExistingEvents){
      Logger.log("Processing " + vevents.length + " Events.");
      var calendarTz = Calendar.Settings.get("timezone").value;
      var calendarUTCOffset = 0;
      var recurringEvents = [];
      
      vevents.forEach(function(e){
        //------------------------ Create the event object ------------------------
        var newEvent = processEvent(e, calendarTz, calendarEventsMD5s, icsEventsIds, startUpdateTime);
        if (newEvent == null)
          return;
        var index = calendarEventsIds.indexOf(newEvent.extendedProperties.private["id"]);
        var needsUpdate = (index > -1);
        
        //------------------------ save instance overrides ------------------------
        //----------- to make sure the parent event is actually created -----------
        if (e.hasProperty('recurrence-id')){
          newEvent.recurringEventId = e.getFirstPropertyValue('recurrence-id').toString();
          Logger.log("Saving event instance for later: " + newEvent.recurringEventId);
          newEvent.extendedProperties.private['rec-id'] = newEvent.extendedProperties.private['id'] + "_" + newEvent.recurringEventId;
          recurringEvents.push(newEvent);
          return;
        }
        else{
          //------------------------ Send event object to gcal ------------------------
          var retries = 0;
          do{
            Utilities.sleep(retries * 100);
            if (needsUpdate){
              if (modifyExistingEvents){
                Logger.log("Updating existing Event " + newEvent.extendedProperties.private["id"]);
                try{
                  newEvent = Calendar.Events.update(newEvent, targetCalendarId, calendarEvents[index].id);
                }
                catch(error){
                  Logger.log("Error, Retrying..." + error);
                }
                if (emailWhenModified)
                  GmailApp.sendEmail(email, "Event \"" + newEvent.summary + "\" modified", "Event was modified in calendar \"" + targetCalendarName + "\" at " + icalEvent.start.toString());
              }
            }
            else{
              if (addEventsToCalendar){
                Logger.log("Adding new Event " + newEvent.extendedProperties.private["id"]);
                try{
                  newEvent = Calendar.Events.insert(newEvent, targetCalendarId);
                }
                catch(error){
                  Logger.log("Error, Retrying..." + error );
                }
                if (emailWhenAdded)
                  GmailApp.sendEmail(email, "New Event \"" + newEvent.summary + "\" added", "New event added to calendar \"" + targetCalendarName + "\" at " + icalEvent.start.toString());
              }
            }
            retries++;
          }while(retries < 5 && (typeof newEvent.etag === "undefined"));
        }
      });
      Logger.log("---done!");
    }
    
    //------------------------ Remove old events from calendar ------------------------
    if(removeEventsFromCalendar){
      Logger.log("Checking " + calendarEvents.length + " events for removal");
      processEventCleanup(calendarEvents, calendarEventsIds, icsEventsIds, targetCalendarId);
      Logger.log("---done!");
    }
    //------------------------ Process Tasks ------------------------
    if (addTasks){
      processTasks(responses);
    }
    //------------------------ Add Recurring Event Instances ------------------------
    Logger.log("---Processing " + recurringEvents.length + " Recurrence Instances!");
    for each (var recEvent in recurringEvents){
      processEventInstance(recEvent, targetCalendarId);
    }
  }
  Logger.log("Sync finished!");
}
