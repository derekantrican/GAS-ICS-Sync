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

var targetCalendarName = "";           // The name of the Google Calendar you want to add events to
var sourceCalendarURL = "";            // The ics/ical url that you want to get events from

var howFrequent = 30;                  // What interval (minutes) to run this script on to check for new events
var addEventsToCalendar = true;        // If you turn this to "false", you can check the log (View > Logs) to make sure your events are being read correctly before turning this on
var modifyExistingEvents = true;       // If you turn this to "false", any event in the feed that was modified after being added to the calendar will not update
var removeEventsFromCalendar = true;   // If you turn this to "true", any event in the calendar not found in the feed will be removed.
var addAlerts = false;                  // Whether to add the ics/ical alerts as notifications on the Google Calendar events
var addOrganizerToTitle = false;       // Whether to prefix the event name with the event organiser for further clarity 
var descriptionAsTitles = false;       // Whether to use the ics/ical descriptions as titles (true) or to use the normal titles as titles (false)
var defaultDuration = 60;              // Default duration (in minutes) in case the event is missing an end specification in the ICS/ICAL file

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
var tzid;
var targetCalendarId;

function main(){
  CheckForUpdate();
  
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
  
  //------------------------ Parse events --------------------------
  var icsEventIds=[];
  
  //Use ICAL.js to parse the data
  var jcalData = ICAL.parse(response);
  var component = new ICAL.Component(jcalData);
  vtimezone = component.getFirstSubcomponent("vtimezone");
  if (vtimezone == null){
    tzid = "GMT";
  }else{
    tzid = vtimezone.getFirstPropertyValue("tzid") || 'GMT';
  }
  
  //Map the vevents into custom event objects
  var icsEvents = component.getAllSubcomponents("vevent").map(ConvertToCustomEvent);
  icsEvents.forEach(function(event){ icsEventIds.push(event.id); }); //Populate the list of icsEventIds
  //----------------------------------------------------------------
  
  if(addEventsToCalendar || removeEventsFromCalendar){
    var calendarEvents = Calendar.Events.list(targetCalendarId).items;
    var calendarEventsIds = []
    Logger.log("Calendar currently has " + calendarEvents.length + " existing Events.");
    for (var i = 0; i < calendarEvents.length; i++)
      calendarEventsIds[i] = calendarEvents[i].id;
  }
  
  //------------------------ Add events to calendar ----------------
  if (addEventsToCalendar){
    Logger.log("Checking " + icsEvents.length + " Events for creation")
    var existingEvent = null;
    for each (var currEvent in icsEvents){
      //---check if Event needs Update, Insert or Skip
      existingEvent = targetCalendar.getEventById(currEvent.id);
      if (existingEvent){
        //Event-ID present in Calendar -> Update
        if (new Date(currEvent.created) > new Date(existingEvent.getLastUpdated())){
          Logger.log("Updating existing Event!");
          currEvent.created = null;
          currEvent = Calendar.Events.update(currEvent, targetCalendarId, currEvent.id);
          if (emailWhenAdded){
            var mailText = "Event updated in calendar \"" + targetCalendarName + "\":\n\n";
            mailText = mailText + "Title: " + currEvent.summary + "\n";
            if (currEvent.description != null)
              mailText = mailText + "Description: " + currEvent.description + "\n";
            if (currEvent.start.date != null){
              mailText = mailText + "Time: " + Utilities.formatDate(new Date(currEvent.start.date), tzid, "YYYY-MM-dd") + " - " + Utilities.formatDate(new Date(currEvent.end.date), tzid, "YYYY-MM-dd") + "\n";
            }else{
              mailText = mailText + "Time: " + Utilities.formatDate(new Date(currEvent.start.dateTime), tzid, "YYYY-MM-dd HH:mm") + " - " + Utilities.formatDate(new Date(currEvent.end.dateTime), tzid, "YYYY-MM-dd HH:mm") + "\n";
            }
            GmailApp.sendEmail(email, "Event modified", mailText);
          }
        }
        else{
          //Existing Event is up to date
          Logger.log("Skipping Update, Event unchanged! Last Update: " + existingEvent.getLastUpdated());
        }
      }
      else{
        //Event-ID currently not in Calendar -> Insert
        Logger.log("Adding new Event " + currEvent.id);
        currEvent.created = null;
        currEvent = Calendar.Events.insert(currEvent, targetCalendarId);
        if (emailWhenAdded){
          var mailText = "New Event added to calendar \"" + targetCalendarName + "\":\n\n";
          mailText = mailText + "Title: " + currEvent.summary + "\n";
          if (currEvent.description != null)
            mailText = mailText + "Description: " + currEvent.description + "\n";
          if (currEvent.start.date != null){
            mailText = mailText + "Time: " + Utilities.formatDate(new Date(currEvent.start.date), tzid, "YYYY-MM-dd") + " - " + Utilities.formatDate(new Date(currEvent.end.date), tzid, "YYYY-MM-dd") + "\n";
          }else{
            mailText = mailText + "Time: " + Utilities.formatDate(new Date(currEvent.start.dateTime), tzid, "YYYY-MM-dd HH:mm") + " - " + Utilities.formatDate(new Date(currEvent.end.dateTime), tzid, "YYYY-MM-dd HH:mm") + "\n";
          }
          GmailApp.sendEmail(email, "New Event Added", mailText);
        }
      }   
    }
  }
  //----------------------------------------------------------------
  
  
  
  //-------------- Remove old events from calendar -----------  
  Logger.log("Checking " + calendarEvents.length + " events for removal");
  for (var i = 0; i < calendarEvents.length; i++){
    var currentID = calendarEvents[i].id;
    var feedIndex = icsEventIds.indexOf(currentID);
    
    if(removeEventsFromCalendar){
      if(feedIndex  == -1){
        Logger.log("Deleting old Event " + currentID);
        Calendar.Events.remove(targetCalendarId, currentID);
      }
    }
  }
  //----------------------------------------------------------------
}

function ConvertToCustomEvent(vevent){
  var event = new Event();
  event.id = vevent.getFirstPropertyValue('uid').toLowerCase().replace(/[-w_x@y.z,]+/g, '');
  event.created = vevent.getFirstPropertyValue('last-modified') || vevent.getFirstPropertyValue('created');
  
  if (descriptionAsTitles)
    event.summary = vevent.getFirstPropertyValue('description') || '';
  else{
    event.summary = vevent.getFirstPropertyValue('summary') || '';
    event.description = vevent.getFirstPropertyValue('description') || '';
  }
  
  if (addOrganizerToTitle){
    var organizer = ParseOrganizerName(vevent.toString());
    
    if (organizer != null)
      event.summary = organizer + ": " + event.summary;
  }
  
  event.location = vevent.getFirstPropertyValue('location') || '';
  
  var dtstart = vevent.getFirstPropertyValue('dtstart');
  var dtend = vevent.getFirstPropertyValue('dtend');
  var duration = vevent.getFirstPropertyValue('duration') || defaultDuration;
  
  if (dtstart.isDate && dtend.isDate){
    event.start = {date: dtstart.toString(),
                   timeZone: tzid};
    if (dtend == null){
      event.end = event.start;
    }else{
      event.end = {date: dtend.toString(),
                   timeZone: tzid};
    }
  }else{
    event.start = {dateTime: dtstart.toString(),
                   timeZone: tzid};
    
    if (dtend == null){
      event.end = {dateTime: Utilities.formatDate(new Date(event.start.dateTime + duration * 60 * 1000), tzid, "yyyy-MM-dd\'T\'HH:mm:ss"),
                   timeZone: tzid};
    }else{
      event.end = {dateTime: dtend.toString(),
                   timeZone: tzid};
    }
  }
  if (addAlerts){
    var valarms = vevent.getAllSubcomponents('valarm');
    for each (var valarm in valarms){
      var trigger = valarm.getFirstPropertyValue('trigger').toString();
      event.reminders.overrides[event.reminders.overrides.length++] = ParseNotificationTime(trigger);
    }
  }
  //RFC allows multiple Rules, Google supports that as well even though it's not displayed in the event details page
  var recurrenceRules = vevent.getAllProperties('rrule');
  event.recurrence = [];
  if (recurrenceRules != null)
    for each (var recRule in recurrenceRules){
      event.recurrence.push("RRULE:" + recRule.getFirstValue().toString());
    }
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

function CheckForUpdate(){
  var alreadyAlerted = PropertiesService.getScriptProperties().getProperty("alertedForNewVersion");
  if (alreadyAlerted == null){
    try{
      var thisVersion = 2.0;
      var html = UrlFetchApp.fetch("https://github.com/derekantrican/GAS-ICS-Sync/releases");
      var regex = RegExp("<a.*title=\"\\d\\.\\d\">","g");
      var latestRelease = regex.exec(html)[0];
      regex = RegExp("\"(\\d.\\d)\"", "g");
      var latestVersion = Number(regex.exec(latestRelease)[1]);
      
      if (latestVersion > thisVersion){
        if (email != ""){
          GmailApp.sendEmail(email, "New version of GAS-ICS-Sync is available!", "There is a new version of \"GAS-ICS-Sync\". You can see the latest release here: https://github.com/derekantrican/GAS-ICS-Sync/releases");
          
          PropertiesService.getScriptProperties().setProperty("alertedForNewVersion", true);
        }
      }
    }
    catch(e){}
  }
}
