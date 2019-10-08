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

var howFrequent = 15;                  // What interval (minutes) to run this script on to check for new events
var addEventsToCalendar = true;        // If you turn this to "false", you can check the log (View > Logs) to make sure your events are being read correctly before turning this on
var modifyExistingEvents = true;       // If you turn this to "false", any event in the feed that was modified after being added to the calendar will not update
var removeEventsFromCalendar = true;   // If you turn this to "true", any event in the calendar not found in the feed will be removed.
var removeExistingEvents = false;      // If you turn this to "true" and "removeEventsFromCalendar" is "true", any event in the calendar that existed before installing this script will be removed if it is not in the feed. You will also be unable to manually add events via the Google Calendar interface(s).
var addAlerts = true;                  // Whether to add the ics/ical alerts as notifications on the Google Calendar events
var addOrganizerToTitle = false;       // Whether to prefix the event name with the event organiser for further clarity 
var descriptionAsTitles = false;       // Whether to use the ics/ical descriptions as titles (true) or to use the normal titles as titles (false)

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

var targetCalendar;

function main(){
  CheckForUpdate();

  //Get URL items
  var response = UrlFetchApp.fetch(sourceCalendarURL).getContentText();

  //Get target calendar information
  targetCalendar = CalendarApp.getCalendarsByName(targetCalendarName)[0];


  //------------------------ Error checking ------------------------
  if(response.includes("That calendar does not exist"))
    throw "[ERROR] Incorrect ics/ical URL";

  if(targetCalendar == null){
    Logger.log("Creating Calendar: " + targetCalendarName);
    targetCalendar = CalendarApp.createCalendar(targetCalendarName);
    targetCalendar.setTimeZone(CalendarApp.getTimeZone());
    targetCalendar.setSelected(true); //Sets the calendar as "shown" in the Google Calendar UI
  }

  if (emailWhenAdded && email == "")
    throw "[ERROR] \"emailWhenAdded\" is set to true, but no email is defined";
  //----------------------------------------------------------------

  //------------------------ Parse events --------------------------
  var feedEventIds=[];

  //Use ICAL.js to parse the data
  var jcalData = ICAL.parse(response);
  var component = new ICAL.Component(jcalData);
  var vtimezones = component.getAllSubcomponents("vtimezone");
  for each (var tz in vtimezones)
    ICAL.TimezoneService.register(tz);
  
  //Map the vevents into custom event objects
  var events = component.getAllSubcomponents("vevent").map(ConvertToCustomEvent);
  events.forEach(function(event){ feedEventIds.push(event.id); }); //Populate the list of feedEventIds
  //----------------------------------------------------------------
  
  //------------------------ Check results -------------------------
  Logger.log("# of events: " + events.length);
  for each (var event in events){
    Logger.log("Title: " + event.title);
    Logger.log("Id: " + event.id);
    Logger.log("Description: " + event.description);
    Logger.log("Start: " + event.startTime);
    Logger.log("End: " + event.endTime);

    for each (var reminder in event.reminderTimes)
      Logger.log("Reminder: " + reminder + " seconds before");

    Logger.log("");
  }
  //----------------------------------------------------------------

  if(addEventsToCalendar || removeEventsFromCalendar){
    var calendarEvents = targetCalendar.getEvents(new Date(2000,01,01), new Date( 2100,01,01 ))
    var calendarFids = []
    for (var i = 0; i < calendarEvents.length; i++)
      calendarFids[i] = calendarEvents[i].getTag("FID");
  }

  //------------------------ Add events to calendar ----------------
  if (addEventsToCalendar){
    Logger.log("Checking " + events.length + " Events for creation")
    for each (var event in events){
      if (calendarFids.indexOf(event.id) == -1){
        var resultEvent;
        if (event.isAllDay){
          resultEvent = targetCalendar.createAllDayEvent(event.title, 
                                                         event.startTime,
                                                         event.endTime,
                                                         {
                                                           location : event.location, 
                                                           description : event.description
                                                         });
        }
        else{
          resultEvent = targetCalendar.createEvent(event.title, 
                                                   event.startTime,
                                                   event.endTime,
                                                   {
                                                     location : event.location, 
                                                     description : event.description
                                                   });
        }
        
        resultEvent.setTag("FID", event.id);
        Logger.log("   Created: " + event.id);
        
        for each (var reminder in event.reminderTimes)
          resultEvent.addPopupReminder(reminder / 60);
          
        if (emailWhenAdded)
          GmailApp.sendEmail(email, "New Event Added", "New event added to calendar \"" + targetCalendarName + "\" at " + event.startTime);
      }
    }
  }
  //----------------------------------------------------------------



  //-------------- Remove Or modify events from calendar -----------  
  Logger.log("Checking " + calendarEvents.length + " events for removal or modification");
  for (var i = 0; i < calendarEvents.length; i++){
    var tagValue = calendarEvents[i].getTag("FID");
    var feedIndex = feedEventIds.indexOf(tagValue);
    
    if(removeEventsFromCalendar){
      if(feedIndex  == -1 && (removeExistingEvents || tagValue != null)){
        Logger.log("    Deleting " + calendarEvents[i].getTitle());
        calendarEvents[i].deleteEvent();
      }
    }

    if(modifyExistingEvents){
      if(feedIndex != -1){
        var e = calendarEvents[i];
        var fes = events.filter(sameEvent, calendarEvents[i].getTag("FID"));
        
        if(fes.length > 0){
          var fe = fes[0];

          if(e.getStartTime().getTime() != fe.startTime.getTime() ||
            e.getEndTime().getTime() != fe.endTime.getTime()){
            if (fe.isAllDay)
              e.setAllDayDates(fe.startTime, fe.endTime);
            else
              e.setTime(fe.startTime, fe.endTime);
          }
          if(e.getTitle() != fe.title)
            e.setTitle(fe.title);
          if(e.getLocation() != fe.location)
            e.setLocation(fe.location)
          if(e.getDescription() != fe.description)
            e.setDescription(fe.description)

        }
      }
    }
  }
  //----------------------------------------------------------------

}

function ConvertToCustomEvent(vevent){
  var icalEvent = new ICAL.Event(vevent);
  var event = new Event();
  event.id = vevent.getFirstPropertyValue('uid');
  
  if (descriptionAsTitles)
    event.title = vevent.getFirstPropertyValue('description') || '';
  else{
    event.title = vevent.getFirstPropertyValue('summary') || '';
    event.description = vevent.getFirstPropertyValue('description') || '';
  }
  
  if (addOrganizerToTitle){
    var organizer = ParseOrganizerName(vevent.toString());

    if (organizer != null)
      event.title = organizer + ": " + event.title;
  }
  
  event.location = vevent.getFirstPropertyValue('location') || '';
  
  if (icalEvent.startDate.isDate && icalEvent.endDate.isDate)
    event.isAllDay = true;
  
  if (icalEvent.startDate.compare(icalEvent.endDate) == 0 && event.isAllDay){
    //Adjust dtend in case dtstart equals dtend as this is not valid for allday events
    icalEvent.endDate = icalEvent.endDate.adjust(1,0,0,0);
  }
  
  if (!event.isAllDay && (icalEvent.startDate.zone.tzid == "floating" || icalEvent.endDate.zone.tzid == "floating")){
    Logger.log("Floating Time detected");
    var targetTZ = targetCalendar.getTimeZone();
    Logger.log("Adding Event in " + targetTZ);
    // Converting start/end timestamps to UTC
    var utcTZ = ICAL.TimezoneService.get("UTC");
    icalEvent.startDate = icalEvent.startDate.convertToZone(utcTZ);
    icalEvent.endDate = icalEvent.endDate.convertToZone(utcTZ);
    var jsTime = icalEvent.startDate.toJSDate();
    // Calculate targetTZ's UTC-Offset
    var utcTime = new Date(Utilities.formatDate(jsTime, "Etc/GMT", "HH:mm:ss MM/dd/yyyy"));
    var tgtTime = new Date(Utilities.formatDate(jsTime, targetTZ, "HH:mm:ss MM/dd/yyyy"));
    var utcOffset = tgtTime - utcTime;
    // Offset initial timestamps by UTC-Offset
    var startTime = new Date(jsTime.getTime() - utcOffset);
    event.startTime = startTime;
    jsTime = icalEvent.endDate.toJSDate();
    var endTime = new Date(jsTime.getTime() - utcOffset);
    event.endTime = endTime;
  }
  else{
    event.startTime = icalEvent.startDate.toJSDate();
    event.endTime = icalEvent.endDate.toJSDate();
  }
  
  if (addAlerts){
    var valarms = vevent.getAllSubcomponents('valarm');
    for each (var valarm in valarms){
      var trigger = valarm.getFirstPropertyValue('trigger').toString();
      event.reminderTimes[event.reminderTimes.length++] = ParseNotificationTime(trigger);
    }
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

function sameEvent(x){
  return x.id == this;
}

function CheckForUpdate(){
  var alreadyAlerted = PropertiesService.getScriptProperties().getProperty("alertedForNewVersion");
  if (alreadyAlerted == null){
    try{
      var thisVersion = 4.0;
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
