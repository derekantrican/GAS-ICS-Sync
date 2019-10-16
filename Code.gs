/*
*=========================================
*       INSTALLATION INSTRUCTIONS
*=========================================
*
* 1) Click in the menu "File" > "Make a copy..." and make a copy to your Google Drive
* 2) Edit Settings.gs to the settings you want to use
* 3) Enable "Calendar API v3" at "Resources" > "Advanced Google Services" > "Calendar API v3" 
* 4) Click in the menu "Run" > "Run function" > "Install" and authorize the program
*    (For steps to follow in authorization, see this video: https://youtu.be/_5k10maGtek?t=1m22s )
* 5) You can run "startSync" directly if you want to sync only once
*
* **To stop Script from running click in the menu "Run" > "Run function" > "Uninstall"
*
*/
//=====================================================================================================
//!!!!!!!!!!!!!!!! DO NOT EDIT BELOW HERE UNLESS YOU REALLY KNOW WHAT YOU'RE DOING !!!!!!!!!!!!!!!!!!!!
//=====================================================================================================
function Install(){
  //Delete any already existing triggers so we don't create excessive triggers
  DeleteAllTriggers();
  
  //Custom error for restriction here: https://developers.google.com/apps-script/reference/script/clock-trigger-builder#everyMinutes(Integer)
  var validFrequencies = [1, 5, 10, 15, 30];
  if(validFrequencies.indexOf(howFrequent) == -1)
    throw "[ERROR] Invalid value for \"howFrequent\". Must be either 1, 5, 10, 15, or 30";
  
  ScriptApp.newTrigger("startSync").timeBased().everyMinutes(howFrequent).create();
}

function Uninstall(){
  DeleteAllTriggers();
}

function startSync(){
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
        var newEvent = processEvent(e, calendarTz, calendarEventsMD5s);
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