/*
*=========================================
*       INSTALLATION INSTRUCTIONS
*=========================================
*
* 1) Click in the menu "File" > "Make a copy..." and make a copy to your Google Drive
* 2) Changes lines 19-32 to be the settings that you want to use
* 3) Click in the menu "Run" > "Run function" > "Install" and authorize the program
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
var addAlerts = false;                  // Whether to add the ics/ical alerts as notifications on the Google Calendar events
var addOrganizerToTitle = false;       // Whether to prefix the event name with the event organiser for further clarity 
var descriptionAsTitles = false;       // Whether to use the ics/ical descriptions as titles (true) or to use the normal titles as titles (false)
var defaultDuration = 60;              // Default duration (in minutes) in case the event is missing an end specification in the ICS/ICAL file
var colorizeEvents = true;

var emailWhenAdded = false;            // Will email you when an event is added to your calendar
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

function main(){
  //Get URL items
  var response = UrlFetchApp.fetch(sourceCalendarURL).getContentText();
  
  //Get target calendar information
  var targetCalendar = CalendarApp.getCalendarsByName(targetCalendarName)[0];
  
  
  //------------------------ Error checking ------------------------
  if(response.includes("That calendar does not exist"))
    throw "[ERROR] Incorrect ics/ical URL";
  
  if(targetCalendar == null){
    Logger.log("Creating Calendar: " + targetCalendarName);
    targetCalendar = CalendarApp.createCalendar(targetCalendarName);
    targetCalendar.setSelected(true); //Sets the calendar as "shown" in the Google Calendar UI
  }
  targetCalendarId = targetCalendar.getId()
  
  Logger.log("Working on calendar: " + targetCalendar.getName() + ", ID: " + targetCalendarId)
  
  if (emailWhenAdded && email == "")
    throw "[ERROR] \"emailWhenAdded\" is set to true, but no email is defined";
  //----------------------------------------------------------------
  
  //------------------------ Parse existing events --------------------------
  
  if(addEventsToCalendar || removeEventsFromCalendar){
    var calendarEvents = Calendar.Events.list(targetCalendarId, {showDeleted: false}).items;
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
  var jcalData = ICAL.parse(response);//sourceCalendarString/response
  var component = new ICAL.Component(jcalData);
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
        icsModDate = event.getFirstPropertyValue('last-modified') || event.getFirstPropertyValue('created') || Date.now();
        calModDate = new Date(calendarEvents[index].updated);
        if (icsModDate > calModDate){
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
              tzid = "GMT"; 
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
          'useDefault': false,
          'overrides': [
            {'method': 'popup', 'minutes': 10}
          ]
        };
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
        
        switch (requiredAction){
          case "insert":
            Logger.log("Adding new Event " + newEvent.iCalUID);
            newEvent = Calendar.Events.insert(newEvent, targetCalendarId);
            break;
          case "update":
            Logger.log("Updating existing Event!");
            newEvent = Calendar.Events.update(newEvent, targetCalendarId, calendarEvents[index].id);
            break;
        }
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
      
      if(feedIndex  == -1){
        Logger.log("Deleting old Event " + currentID);
        Calendar.Events.remove(targetCalendarId, calendarEvents[i].id);
      }
    }
    Logger.log("---done!");
  }
  //----------------------------------------------------------------
}

//old
function ConvertToCustomEvent(vevent){

  var duration = vevent.getFirstPropertyValue('duration') || defaultDuration;
  
  if (dtstart.isDate && dtend.isDate)
    event.isAllDay = true;
    
  event.startTime = dtstart;
  
  if (dtend == null)
    event.endTime = new Date(event.startTime.getTime() + duration * 60 * 1000);
  else{
    if (vtimezone != null)
      dtend.zone = new ICAL.Timezone(vtimezone);
      
    event.endTime = dtend;
  }
  
  if (addAlerts){
    var valarms = vevent.getAllSubcomponents('valarm');
    for each (var valarm in valarms){
      var trigger = valarm.getFirstPropertyValue('trigger').toString();
      event.reminderTimes[event.reminderTimes.length++] = ParseNotificationTime(trigger);
    }
  }

  var recurrenceRules = vevent.getAllProperties('rrule');
  event.recurrence = [];
  if (recurrenceRules != null)
    for each (var recRule in recurrenceRules){
      event.recurrence.push("RRULE:" + recRule.getFirstValue().toString());
    }
  var exDatesRegex = RegExp("EXDATE(.*)", "g");
  var exdates = vevent.toString().match(exDatesRegex);
  if (exdates != null){
    event.recurrence = event.recurrence.concat(exdates);
  }
  var rDatesRegex = RegExp("RDATE(.*)", "g");
  var rdates = vevent.toString().match(rDatesRegex);
  if (rdates != null){
    event.recurrence = event.recurrence.concat(rdates);
  }
  Logger.log("TF: " + event.startTime + " - " + event.endTime + " All Day: " + event.isAllDay);
  return event;
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