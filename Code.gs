/* --------------- HOW TO INSTALL ---------------
*
* 1) Click in the menu "File" > "Make a copy..." and make a copy to your Google Drive
* 2) Changes lines 11-20 to be the settings that you want to use
* 3) Click in the menu "Run" > "Run function" > "Install" and authorize the program
*    (For steps to follow in authorization, see this video: https://youtu.be/_5k10maGtek?t=1m22s )
* 4) To stop Script from running click in the menu "Edit" > "Current Project's Triggers".  Delete the running trigger.
*/

// --------------- SETTINGS ---------------

var targetCalendarName = "" // The name of the Google Calendar you want to add events to
var sourceCalendarURL = "" // The ics/ical url that you want to get events from


// Currently global settings are applied to all sourceCalendars.

var howFrequent = 15; //What interval (minutes) to run this script on to check for new events
var addEventsToCalendar = true; //If you turn this to "false", you can check the log (View > Logs) to make sure your events are being read correctly before turning this on
var modifyExistingEvents = true; // If you turn this to "false", any event in the feed that was modified after being added to the calendar will not update
var removeEventsFromCalendar = true; //If you turn this to "true", any event in the calendar not found in the feed will be removed.
var addAlerts = true; //Whether to add the ics/ical alerts as notifications on the Google Calendar events
var addOrganizerToTitle = false; //Whether to prefix the event name with the event organiser for further clarity 
var descriptionAsTitles = false; //Whether to use the ics/ical descriptions as titles (true) or to use the normal titles as titles (false)
var defaultDuration = 60; //Default duration (in minutes) in case the event is missing an end specification in the ICS/ICAL file

var emailWhenAdded = false; //Will email you when an event is added to your calendar
var email = ""; //OPTIONAL: If "emailWhenAdded" is set to true, you will need to provide your email

// ----------------------------------------

/* --------------- MISCELLANEOUS ----------
*
* This program was created by Derek Antrican
*
* The code for this program is kept updated here: https://github.com/derekantrican/GAS-ICS-Sync
*
* If you would like to donate and help Derek keep making awesome programs,
* you can do that here: https://bulkeditcalendarevents.wordpress.com/donate/
*
* If you would like to see other programs Derek has made, you can check out
* his website: derekantrican.com or his github: https://github.com/derekantrican
*
*
* Program was modified by Andrew Brothers
* Github: https://github.com/agentd00nut/
* Twitter: @abrothers656
*/



//---------------- DO NOT EDIT BELOW HERE UNLESS YOU REALLY KNOW WHAT YOU'RE DOING --------------------
function Install(){
  ScriptApp.newTrigger("main").timeBased().everyMinutes(howFrequent).create();
}

function main(){

  //Get URL items
  var response = UrlFetchApp.fetch(sourceCalendarURL);
  response = response.getContentText().split("\r\n");

  //Get target calendar information
  var targetCalendar = CalendarApp.getCalendarsByName(targetCalendarName)[0];




  //------------------------ Error checking ------------------------
  if(response[0] == "That calendar does not exist.")
    throw "[ERROR] Incorrect ics/ical URL";

  if(targetCalendar == null){
      Logger.log("Creating Calendar: "+targetCalendarName);
      targetCalendar = CalendarApp.createCalendar(targetCalendarName);
  }

  if (emailWhenAdded && email == "")
    throw "[ERROR] \"emailWhenAdded\" is set to true, but no email is defined";
  //----------------------------------------------------------------

  //------------------------ Parse events --------------------------
  //https://en.wikipedia.org/wiki/ICalendar#Technical_specifications
  //https://tools.ietf.org/html/rfc5545
  //https://www.kanzaki.com/docs/ical

  var parsingEvent = false;
  var parsingNotification = false;
  var currentEvent;
  var events = [];
  var feedEventIds=[];
  var item;
  
  var eventOrganizer = "";
  var eventSummary = "";

  for (var i = 0; i < response.length; i++){
    item = response[i];
    while (i + 1 < response.length && response[i + 1][0] == " ") {
      item += response[i + 1].substr(1);
      i++;
    }
    item = item.replace(/\\n/g, "\n")
               .replace(/\\r/g, "\r")
               .replace(/\\t/g, "\t")
               .replace(/\\(.)/g, "$1");

    if (item == "BEGIN:VEVENT"){
      parsingEvent = true;
      currentEvent = new Event();
    }
    else if (item == "END:VEVENT"){
      currentEvent.title = addOrganizerToTitle ? eventOrganizer + ": " + eventSummary : eventSummary;
      
      if (currentEvent.endTime == null)
        currentEvent.endTime = new Date(currentEvent.startTime.getTime() + defaultDuration * 60 * 1000);
      
      parsingEvent = false;
      events[events.length] = currentEvent;
    }
    else if (item == "BEGIN:VALARM")
      parsingNotification = true;
    else if (item == "END:VALARM")
      parsingNotification = false;
    else if (parsingNotification){
      if (addAlerts){
        if (item.includes("TRIGGER"))
          currentEvent.reminderTimes[currentEvent.reminderTimes.length++] = ParseNotificationTime(item.split("TRIGGER:")[1]);
      }
    }
    else if (parsingEvent){
      if (item.includes("ORGANIZER"))
        eventOrganizer = item.split("ORGANIZER;CN=")[1].split(":")[0];
      
      if (item.includes("SUMMARY") && !descriptionAsTitles)
        eventSummary = item.split("SUMMARY:")[1];

      if (item.includes("DESCRIPTION") && descriptionAsTitles)
        currentEvent.title = item.split("DESCRIPTION:")[1];
      else if (item.includes("DESCRIPTION"))
        currentEvent.description = item.split("DESCRIPTION:")[1];

      if (item.includes("DTSTART"))
        currentEvent.startTime = moment(GetUTCTime(item.split("DTSTART")[1]), "YYYYMMDDTHHmmssZ").toDate();

      if (item.includes("DTEND"))
        currentEvent.endTime = moment(GetUTCTime(item.split("DTEND")[1]), "YYYYMMDDTHHmmssZ").toDate();

      if (item.includes("LOCATION"))
        currentEvent.location = item.split("LOCATION:")[1];

      if (item.includes("UID")){
        currentEvent.id = item.split("UID:")[1];
        feedEventIds.push(currentEvent.id);
      }
    }
  }
  //----------------------------------------------------------------

  //------------------------ Check results -------------------------
  Logger.log("# of events: " + events.length);
  for (var i = 0; i < events.length; i++){
    Logger.log("Title: " + events[i].title);
    Logger.log("Id: " + events[i].id);
    Logger.log("Description: " + events[i].description);
    Logger.log("Start: " + events[i].startTime);
    Logger.log("End: " + events[i].endTime);

    for (var j = 0; j < events[i].reminderTimes.length; j++)
      Logger.log("Reminder: " + events[i].reminderTimes[j] + " seconds before");

    Logger.log("");
  }
  //----------------------------------------------------------------

  if(addEventsToCalendar || removeEventsFromCalendar){
    var calendarEvents = targetCalendar.getEvents(new Date(2000,01,01), new Date( 2100,01,01 ))
    var calendarFids = []
    for(var i=0; i<calendarEvents.length; i++){
      calendarFids[i] = calendarEvents[i].getTag("FID");
    }
  }

  //------------------------ Add events to calendar ----------------
  if (addEventsToCalendar){
    Logger.log("Checking "+events.length+" Events for creation")
    for (var i = 0; i < events.length; i++){
//      Logger.log("Checking: "+ events[i].id);

      if (calendarFids.indexOf( events[i].id ) == -1 ){
        var resultEvent = targetCalendar.createEvent(events[i].title, events[i].startTime, events[i].endTime, {location : events[i].location, description : events[i].description });

        resultEvent.setTag("FID", events[i].id)
        Logger.log("   Created: "+events[i].id);

        for (var j = 0; j < events[i].reminderTimes.length; j++)
          resultEvent.addPopupReminder(events[i].reminderTimes[j] / 60);

        if (emailWhenAdded)
          GmailApp.sendEmail(email, "New Event Added", "New event added to calendar \"" + targetCalendarName + "\" at " + events[i].startTime);
      }
    }
  }
  //----------------------------------------------------------------



  //-------------- Remove Or modify events from calendar -----------
  for(var i=0; i < calendarEvents.length; i++){
    Logger.log("Checking "+calendarEvents.length+" Events For Removal or Modification");
    var tagValue = calendarEvents[i].getTag("FID");
    var feedIndex = feedEventIds.indexOf(tagValue);
    if(removeEventsFromCalendar){
      if(feedIndex  == -1 && tagValue != null){
        Logger.log("    Deleting "+calendarEvents[i].getTitle());
        calendarEvents[i].deleteEvent();
      }
    }

    if(modifyExistingEvents){
      if( feedIndex != -1  ){
        var e = calendarEvents[i];

        var fes = events.filter(sameEvent, calendarEvents[i].getTag("FID"));
        if( fes.length > 0){

          var fe = fes[0];

          if(e.getStartTime() != fe.startTime || e.getEndTime() != fe.endTime)
            e.setTime(fe.startTime, fe.endTime)
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
//function EventExists(calendar, event){
//
//  var events = calendar.getEvents(event.startTime, event.endTime, {search : event.id});
//  return events.length > 0;
//}

function GetUTCTime(parameter){
  parameter = parameter.substr(1); //Remove leading ; or : character
  if (parameter.includes("TZID")){
    var tzid = parameter.split("TZID=")[1].split(":")[0];
    var time = parameter.split(":")[1];
    return moment.tz(time,tzid).tz("Etc/UTC").format("YYYYMMDDTHHmmss") + "Z";
  }
  else
    return parameter;
}
