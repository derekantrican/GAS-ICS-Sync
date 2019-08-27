/*
*=========================================
*       INSTALLATION INSTRUCTIONS
*=========================================
*
* 1) Click in the menu "File" > "Make a copy..." and make a copy to your Google Drive
* 2) Changes lines 19-32 to be the settings that you want to use
* 3) Enable "Calendar API v3" at "Resources" > "Advanced Google Services" > "Calendar API v3" 
* 4) Click in the menu "Run" > "Run function" > "Install" and authorize the program
*    (For steps to follow in authorization, see this video: https://youtu.be/_5k10maGtek?t=1m22s )
*
*
* **To stop Script from running click in the menu "Run" > "Run function" > "Uninstall"
*
*=========================================
*               SETTINGS
*=========================================
*/

var targetCalendarName = "Full API TEST";           // The name of the Google Calendar you want to add events to
var sourceCalendarURL = "";            // The ics/ical url that you want to get events from

var howFrequent = 15;                  // What interval (minutes) to run this script on to check for new events
var addEventsToCalendar = true;        // If you turn this to "false", you can check the log (View > Logs) to make sure your events are being read correctly before turning this on
var modifyExistingEvents = true;       // If you turn this to "false", any event in the feed that was modified after being added to the calendar will not update
var removeEventsFromCalendar = true;   // If you turn this to "true", any event in the calendar not found in the feed will be removed.
var addAlerts = true;                  // Whether to add the ics/ical alerts as notifications on the Google Calendar events, this will override the standard reminders specified by the target calendar.
var addOrganizerToTitle = false;       // Whether to prefix the event name with the event organiser for further clarity 
var addTasks = false;

var emailWhenAdded = false;            // Will email you when an event is added to your calendar
var email = "";                        // OPTIONAL: If "emailWhenAdded" is set to true, you will need to provide your email

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

  ScriptApp.newTrigger("main").timeBased().everyMinutes(howFrequent).create();
}

function Uninstall(){
  DeleteAllTriggers();
}

function DeleteAllTriggers(){
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++){
    if (triggers[i].getHandlerFunction() == "main"){
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

var vtimezone;
var targetCalendarId;
var response;

function parseTasks(){
  var taskLists = Tasks.Tasklists.list().items;
  var taskList = taskLists[0];
  
  var existingTasks = Tasks.Tasks.list(taskList.id).items || [];
  var existingTasksIds = []
  Logger.log("Grabbed " + existingTasks.length + " existing Tasks from " + taskList.title);
  for (var i = 0; i < existingTasks.length; i++){
    existingTasksIds[i] = existingTasks[i].id;
  }
  Logger.log("Saved " + existingTasksIds.length + " existing Task IDs");
  
  var icsTasksIds = [];
  
  var jcalData = ICAL.parse(response);
  var component = new ICAL.Component(jcalData);
  
  var vtasks = component.getAllSubcomponents("vtodo");
  vtasks.forEach(function(task){ icsTasksIds.push(task.getFirstPropertyValue('uid').toString()); });
  Logger.log("---Processing " + vtasks.length + " Tasks.");
  
  for each (var task in vtasks){
    var newtask = Tasks.newTask();
    newtask.id = task.getFirstPropertyValue("uid").toString();
    newtask.title = task.getFirstPropertyValue("summary").toString();
    var d = task.getFirstPropertyValue("due").toJSDate();
    newtask.due = (d.getFullYear()) + "-" + ("0"+(d.getMonth()+1)).slice(-2) + "-" + ("0" + d.getDate()).slice(-2) + "T" + ("0" + d.getHours()).slice(-2) + ":" + ("0" + d.getMinutes()).slice(-2) + ":" + ("0" + d.getSeconds()).slice(-2)+"Z";
    Tasks.Tasks.insert(newtask, taskList.id);
  };
  Logger.log("---Done!");
  
  //-------------- Remove old Tasks -----------
  // ID can't be used as identifier as the API reassignes a random id at task creation
  if(removeEventsFromCalendar){
    Logger.log("Checking " + existingTasksIds.length + " tasks for removal");
    for (var i = 0; i < existingTasksIds.length; i++){
      var currentID = existingTasks[i].id;
      var feedIndex = icsTasksIds.indexOf(currentID);
      
      if(feedIndex  == -1){
        Logger.log("Deleting old Task " + currentID);
        Tasks.Tasks.remove(taskList.id, currentID);
      }
    }
    Logger.log("---Done!");
  }
  //----------------------------------------------------------------
}

function main(){
  //Get URL items
  response = UrlFetchApp.fetch(sourceCalendarURL).getContentText();
  
  //Get target calendar information
  var targetCalendar = Calendar.CalendarList.list().items.filter(function(cal) {
  return cal.summary == targetCalendarName;
  })[0];
  
  
  //------------------------ Error checking ------------------------
  if(response.includes("That calendar does not exist"))
    throw "[ERROR] Incorrect ics/ical URL";
  
  if(targetCalendar == null){
    Logger.log("Creating Calendar: " + targetCalendarName);
    targetCalendar = Calendar.newCalendar();
    targetCalendar.summary = targetCalendarName;
    targetCalendar.description = "Created by GAS.";
    targetCalendar.timeZone = Calendar.Calendars.get("primary").timeZone;
    targetCalendar = Calendar.Calendars.insert(targetCalendar);
 }
  //targetCalendarId = targetCalendar.getId()
  targetCalendarId = targetCalendar.id;
  
  Logger.log("Working on calendar: " + targetCalendar.summary + ", ID: " + targetCalendarId)
  
  if (emailWhenAdded && email == "")
    throw "[ERROR] \"emailWhenAdded\" is set to true, but no email is defined";
  //----------------------------------------------------------------
  
  //------------------------ Parse existing events --------------------------
  
  if(addEventsToCalendar || removeEventsFromCalendar){ 
    var calendarEvents = Calendar.Events.list(targetCalendarId, {showDeleted: true}).items; 
    var calendarEventsIds = [] 
    Logger.log("Grabbed " + calendarEvents.length + " existing Events from " + targetCalendarName); 
    for (var i = 0; i < calendarEvents.length; i++){ 
      calendarEventsIds[i] = calendarEvents[i].iCalUID;
    } 
    Logger.log("Saved " + calendarEventsIds.length + " existing Event IDs"); 
  } 

  //------------------------ Parse ics events --------------------------
  var icsEventIds=[];
  
  //Use ICAL.js to parse the data
  var jcalData = ICAL.parse(response);
  var component = new ICAL.Component(jcalData);
  ICAL.helpers.updateTimezones(component);
  vtimezone = component.getAllSubcomponents("vtimezone");
  for each (var tz in vtimezone){
    ICAL.TimezoneService.register(tz);
  }
  var vevents = component.getAllSubcomponents("vevent");
  vevents.forEach(function(event){ icsEventIds.push(event.getFirstPropertyValue('uid').toString()); });
  
  if (addEventsToCalendar || modifyExistingEvents){
    Logger.log("---Processing " + vevents.length + " Events.");
    for each (var event in vevents){
      vevent = new ICAL.Event(event);
      var requiredAction = "skip";
      var index = calendarEventsIds.indexOf(vevent.uid);
      if (index >= 0){
        //check update
        icsModDate = event.getFirstPropertyValue('last-modified') || event.getFirstPropertyValue('created');
        calModDate = new Date(calendarEvents[index].updated);
        if (calendarEvents[index].status == "cancelled"){
          requiredAction = "update";
        }else if (icsModDate === null){
          //manually check if event changed
          if (eventChanged(event, vevent, calendarEvents[index])){
            requiredAction = "update";
          }
        }else if (icsModDate > calModDate){
          requiredAction = "update";
        }else{
          //skip
        }
      }else{
        requiredAction = "insert";
      }
      
      if (requiredAction != "skip"){
        var newEvent = new Event();
        if(vevent.startDate.isDate){
          //All Day Event
          newEvent = {
            start: {
              date: vevent.startDate.toString()
            },
            end: {
              date: vevent.endDate.toString()
            }
          };
        }else{
          //normal Event
          var tzid = vevent.startDate.timezone;
          if (tzids.indexOf(tzid) == -1){
            Logger.log("Timezone " + tzid + " unsupported!");
            if (tzid in tzidreplace){
              tzid = tzidreplace[tzid];
            }else{
              tzid = Calendar.Calendars.get("primary").timeZone; 
            }
            Logger.log("Using Timezone " + tzid + "!");
          };
          newEvent = {
            start: {
              dateTime: vevent.startDate.toString(),
              timeZone: tzid
            },
            end: {
              dateTime: vevent.endDate.toString(),
              timeZone: tzid
            },
          };
        }
        
        newEvent.summary = vevent.summary;
        if (addOrganizerToTitle){
          var organizer = ParseOrganizerName(event.toString());
          
          if (organizer != null)
            newEvent.summary = organizer + ": " + vevent.summary;
        }
        
        newEvent.iCalUID = vevent.uid;
        newEvent.description = vevent.description;
        newEvent.location = vevent.location;
        newEvent.reminders = {
          'useDefault': true
        };
        if (addAlerts){
          var valarms = event.getAllSubcomponents('valarm');
          var overrides = [];
          for each (var valarm in valarms){
            var trigger = valarm.getFirstPropertyValue('trigger').toString();
            if (overrides.length < 5){ //Google supports max 5 reminder-overrides
              overrides.push({'method': 'popup', 'minutes': ParseNotificationTime(trigger)/60});
            }
          }
          if (overrides.length > 0){
            newEvent.reminders = {
              'useDefault': false,
              'overrides': overrides
            };
          }
        }
        var recurrenceRules = event.getAllProperties('rrule');
        var recurrence = [];
        if (recurrenceRules != null)
          for each (var recRule in recurrenceRules){
            recurrence.push("RRULE:" + recRule.getFirstValue().toString());
          }
        var exDatesRegex = RegExp("EXDATE(.*)", "g");
        var exdates = event.toString().match(exDatesRegex);
        if (exdates != null){
          recurrence = recurrence.concat(exdates);
        }
        var rDatesRegex = RegExp("RDATE(.*)", "g");
        var rdates = event.toString().match(rDatesRegex);
        if (rdates != null){
          recurrence = recurrence.concat(rdates);
        }
        newEvent.recurrence = recurrence;
        
        var retries = 0;
        do{
          Utilities.sleep(retries * 100);
          switch (requiredAction){
            case "insert":
              Logger.log("Adding new Event " + newEvent.iCalUID);
              try{
                newEvent = Calendar.Events.insert(newEvent, targetCalendarId);
              }catch(error){
                Logger.log("Error, Retrying..." + error );
              }
              break;
            case "update":
              Logger.log("Updating existing Event!");
              try{
              newEvent = Calendar.Events.update(newEvent, targetCalendarId, calendarEvents[index].id);
              }catch(error){
                Logger.log("Error, Retrying..." + error);
              }
              break;
          }
          retries++;
        }while(retries < 5 && (typeof newEvent.etag === "undefined"))
      }else{
        //Skipping
        Logger.log("Event unchanged. No action required.")
      }
    }
    Logger.log("---done!");
  }
  
  //-------------- Remove old events from calendar -----------
  if(removeEventsFromCalendar){
    Logger.log("Checking " + calendarEvents.length + " events for removal");
    for (var i = 0; i < calendarEvents.length; i++){
      var currentID = calendarEventsIds[i];
      var feedIndex = icsEventIds.indexOf(currentID);
      
      if(feedIndex  == -1 && calendarEvents[i].status != "cancelled"){
        Logger.log("Deleting old Event " + currentID);
        try{
          Calendar.Events.remove(targetCalendarId, calendarEvents[i].id);
        }catch (err){
          Logger.log(err);
        }
      }
    }
    Logger.log("---done!");
  }
  //----------------------------------------------------------------
  if (addTasks)
    parseTasks();
}

function ParseOrganizerName(veventString){
  /*A regex match is necessary here because ICAL.js doesn't let us directly
  * get the "CN" part of an ORGANIZER property. With something like
  * ORGANIZER;CN="Sally Example":mailto:sally@example.com
  * VEVENT.getFirstPropertyValue('organizer') returns "mailto:sally@example.com".
  * Therefore we have to use a regex match on the VEVENT string instead
  */

  var nameMatch = RegExp("ORGANIZER(?:;|:)CN=(.*?):", "g").exec(veventString);
  if (nameMatch != null && nameMatch.length > 1)
    return nameMatch[1];
  else
    return null;
}

function ParseNotificationTime(notificationString){
  //https://www.kanzaki.com/docs/ical/duration-t.html
  var reminderTime = 0;

  //We will assume all notifications are BEFORE the event
  if (notificationString[0] == "+" || notificationString[0] == "-")
    notificationString = notificationString.substr(1);

  notificationString = notificationString.substr(1); //Remove "P" character

  var secondMatch = RegExp("\\d+S", "g").exec(notificationString);
  var minuteMatch = RegExp("\\d+M", "g").exec(notificationString);
  var hourMatch = RegExp("\\d+H", "g").exec(notificationString);
  var dayMatch = RegExp("\\d+D", "g").exec(notificationString);
  var weekMatch = RegExp("\\d+W", "g").exec(notificationString);

  if (weekMatch != null){
    reminderTime += parseInt(weekMatch[0].slice(0, -1)) & 7 * 24 * 60 * 60; //Remove the "W" off the end

    return reminderTime; //Return the notification time in seconds
  }
  else{
    if (secondMatch != null)
      reminderTime += parseInt(secondMatch[0].slice(0, -1)); //Remove the "S" off the end

    if (minuteMatch != null)
      reminderTime += parseInt(minuteMatch[0].slice(0, -1)) * 60; //Remove the "M" off the end

    if (hourMatch != null)
      reminderTime += parseInt(hourMatch[0].slice(0, -1)) * 60 * 60; //Remove the "H" off the end

    if (dayMatch != null)
      reminderTime += parseInt(dayMatch[0].slice(0, -1)) * 24 * 60 * 60; //Remove the "D" off the end

    return reminderTime; //Return the notification time in seconds
  }
}

function eventChanged(event, icsEvent, calEvent){
  if (icsEvent.description != calEvent.description)
    return true;
  if (icsEvent.summary != calEvent.summary)
    return true;
  if (icsEvent.location != calEvent.location)
    return true;
  var startDate = calEvent.start.date || calEvent.start.dateTime;
  startDate = new ICAL.Time().fromJSDate(new Date(startDate), true);
  if (startDate.compare(icsEvent.startDate) != 0)
    return true;
  var endDate = calEvent.end.date || calEvent.end.dateTime;
  endDate = new ICAL.Time().fromJSDate(new Date(endDate), true);
  if (endDate.compare(icsEvent.endDate) != 0)
    return true;
//  Need to manually build the recurrence-array
//  var recurrenceRules = event.getAllProperties('rrule');
//  var recurrence = [];
//  if (recurrenceRules != null)
//    for each (var recRule in recurrenceRules){
//      recurrence.push("RRULE:" + recRule.getFirstValue().toString());
//    }
//  var exDatesRegex = RegExp("EXDATE(.*)", "g");
//  var exdates = event.toString().match(exDatesRegex);
//  if (exdates != null){
//    recurrence = recurrence.concat(exdates);
//  }
//  var rDatesRegex = RegExp("RDATE(.*)", "g");
//  var rdates = event.toString().match(rDatesRegex);
//  if (rdates != null){
//    recurrence = recurrence.concat(rdates);
//  }
//  Logger.log("Comparing recurrence: " + recurrence + " - " + calEvent.recurrence);
//  if (icsEvent.recurrence != calEvent.recurrence)
//    return true;
  
  return false;
}