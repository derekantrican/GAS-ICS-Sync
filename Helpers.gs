String.prototype.includes = function(phrase){ 
  return this.indexOf(phrase) > -1;
}

function DeleteAllTriggers(){
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++){
    if (triggers[i].getHandlerFunction() == "main"){
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function fetchSourceCalendars(){
  var result = []
  for each (var url in sourceCalendarURLs){
    try{
      var urlResponse = UrlFetchApp.fetch(url).getContentText();
      //------------------------ Error checking ------------------------
      if(!urlResponse.includes("BEGIN:VCALENDAR")){
        Logger.log("[ERROR] Incorrect ics/ical URL: " + url);
      }
      else{
        result.push(urlResponse);
      }
    }
    catch (e){
      Logger.log(e);
    }
  }
  return result;
}

function setupTargetCalendar(){
  var targetCalendar = Calendar.CalendarList.list().items.filter(function(cal) {
    return cal.summary == targetCalendarName;
  })[0];
  
  if(targetCalendar == null){
    Logger.log("Creating Calendar: " + targetCalendarName);
    targetCalendar = Calendar.newCalendar();
    targetCalendar.summary = targetCalendarName;
    targetCalendar.description = "Created by GAS.";
    targetCalendar.timeZone = Calendar.Settings.get("timezone").value;
    targetCalendar = Calendar.Calendars.insert(targetCalendar);
  }
  return targetCalendar;
}

function parseResponses(responses, icsEventIds){
  var result = [];
  for each (var resp in responses){
    var jcalData = ICAL.parse(resp);
    var component = new ICAL.Component(jcalData);
    ICAL.helpers.updateTimezones(component);
    var vtimezones = component.getAllSubcomponents("vtimezone");
    for each (var tz in vtimezones){
      ICAL.TimezoneService.register(tz);
    }
    
    var allEvents = component.getAllSubcomponents("vevent");
    var calName = component.getFirstPropertyValue("name");
    if (calName != null)
      allEvents.forEach(function(event){event.addPropertyWithValue("parentCal", calName); });
    result = [].concat(allEvents, result);
  }
  result.forEach(function(event){
    if(event.hasProperty('recurrence-id')){
      icsEventIds.push(event.getFirstPropertyValue('uid').toString() + "_" + event.getFirstPropertyValue('recurrence-id').toString());
    }
    else{
      icsEventIds.push(event.getFirstPropertyValue('uid').toString());
    }
  });
 
  return result;
}

function processEvent(event, calendarTz, calendarEventsMD5s){
  event.removeProperty('dtstamp');
  var icalEvent = new ICAL.Event(event);
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, icalEvent.toString()).toString();
  if(calendarEventsMD5s.indexOf(digest) >= 0){
    Logger.log("Skipping unchanged Event " + event.getFirstPropertyValue('uid').toString());
    return;
  }
  var newEvent = Calendar.newEvent();
  if(icalEvent.startDate.isDate){
    //All Day Event
    if (icalEvent.startDate.compare(icalEvent.endDate) == 0){
      //Adjust dtend in case dtstart equals dtend as this is not valid for allday events
      icalEvent.endDate = icalEvent.endDate.adjust(1,0,0,0);
    }
    newEvent = {
      start: {
        date: icalEvent.startDate.toString()
      },
      end: {
        date: icalEvent.endDate.toString()
      }
    };
  }
  else{
    //normal Event
    var tzid = icalEvent.startDate.timezone;
    if (tzids.indexOf(tzid) == -1){
      Logger.log("Timezone " + tzid + " unsupported!");
      if (tzid in tzidreplace){
        tzid = tzidreplace[tzid];
      }
      else{
        //floating time
        tzid = calendarTz;
      }
      Logger.log("Using Timezone " + tzid + "!");
    };
    newEvent = {
      start: {
        dateTime: icalEvent.startDate.toString(),
        timeZone: tzid
      },
      end: {
        dateTime: icalEvent.endDate.toString(),
        timeZone: tzid
      },
    };
  }
  
  if (addAttendees && event.hasProperty('attendee')){
    newEvent.attendees = [];
    for each (var att in icalEvent.attendees){
      var mail = ParseAttendeeMail(att.toICALString());
      if (mail!=null){
        var newAttendee = {'email':mail};
        var name = ParseAttendeeName(att.toICALString());
        if (name!=null)
          newAttendee['displayName'] = name;
        var resp = ParseAttendeeResp(att.toICALString());
        if (resp!=null)
          newAttendee['responseStatus'] = resp;
        newEvent.attendees.push(newAttendee);
      }
    }
  }
  if (event.hasProperty('status')){
    var status = event.getFirstPropertyValue('status').toString().toLowerCase();
    if (["confirmed", "tentative", "cancelled"].indexOf(status) > -1)
      newEvent.status = status;
  }
  if (event.hasProperty('url')){
    newEvent.source = Calendar.newEventSource()
    newEvent.source.url = event.getFirstPropertyValue('url').toString();
  }
  if (event.hasProperty('sequence')){
    newEvent.sequence = icalEvent.sequence;
  }
  if (event.hasProperty('summary')){
    newEvent.summary = icalEvent.summary;
  }
  if (addOrganizerToTitle){
    var organizer = ParseOrganizerName(event.toString());
    
    if (organizer != null)
      newEvent.summary = organizer + ": " + newEvent.summary;
  }
  
  if (addCalToTitle && event.hasProperty('parentCal')){
    var calName = event.getFirstPropertyValue('parentCal');
    newEvent.summary = "(" + calName + ") " + newEvent.summary;
  }
  
  if (event.hasProperty('description'))
    newEvent.description = icalEvent.description;
  if (event.hasProperty('location'))
    newEvent.location = icalEvent.location;
  if (event.hasProperty('class')){
    var class = event.getFirstPropertyValue('class').toString().toLowerCase();
    if (["default","public","private","confidential"].indexOf(class) > -1)
      newEvent.visibility = class;
  }
  if (event.hasProperty('transp')){
    var transparency = event.getFirstPropertyValue('transp').toString().toLowerCase();
    if(["opaque","transparent"].indexOf(transparency) > -1)
      newEvent.transparency = transparency;
  }
  if (addAlerts){
    var valarms = event.getAllSubcomponents('valarm');
    if (valarms.length == 0){
      newEvent.reminders = {
        'useDefault': true
      };
    }
    else{
      var overrides = [];
      for each (var valarm in valarms){
        var trigger = valarm.getFirstPropertyValue('trigger').toString();
        if (overrides.length < 5){ //Google supports max 5 reminder-overrides
          var timer = ParseNotificationTime(trigger)/60;
          if (0 <= timer <= 40320)
            overrides.push({'method': 'popup', 'minutes': timer});
        }
      }
      if (overrides.length > 0){
        newEvent.reminders = {
          'useDefault': false,
          'overrides': overrides
        };
      }
    }
  }
  
  if (event.hasProperty('rrule') || event.hasProperty('rdate')){
    // Calculate targetTZ's UTC-Offset
    var jsTime = new Date();
    var utcTime = new Date(Utilities.formatDate(jsTime, "Etc/GMT", "HH:mm:ss MM/dd/yyyy"));
    var tgtTime = new Date(Utilities.formatDate(jsTime, calendarTz, "HH:mm:ss MM/dd/yyyy"));
    calendarUTCOffset = tgtTime - utcTime;
    newEvent.recurrence = ParseRecurrenceRule(event, calendarUTCOffset);
  }
  
  newEvent.extendedProperties = {private: {MD5: digest, fromGAS: "true", id: icalEvent.uid}};
  return newEvent;
}

function processEventInstance(recEvent, targetCalendarId){
  Logger.log("-----" + recEvent.recurringEventId.substring(0,10));
  var recIDStart = new Date(recEvent.recurringEventId);
  recIDStart = new ICAL.Time.fromJSDate(recIDStart, true);
  var eventInstanceToPatch = Calendar.Events.list(targetCalendarId, {timeZone:"etc/GMT", singleEvents: true, privateExtendedProperty: "fromGAS=true", privateExtendedProperty: "id=" + recEvent.extendedProperties.private['id']}).items.filter(function(item){
    var origStart = item.originalStartTime.dateTime || item.originalStartTime.date
    var instanceStart = new ICAL.Time.fromString(origStart);
    return (instanceStart.compare(recIDStart) == 0);
  });
  if (eventInstanceToPatch.length == 0){
    Logger.log("No Instance found, skipping!");
  }
  else{
    try{
      Logger.log("Patching event instance " + eventInstanceToPatch[0].id)
      Calendar.Events.patch(recEvent, targetCalendarId, eventInstanceToPatch[0].id);
    }
    catch(error){
      Logger.log(error); 
    }
  }
}

function processEventCleanup(calendarEvents, calendarEventsIds, icsEventsIds, targetCalendarId){
  for (var i = 0; i < calendarEvents.length; i++){
      var currentID = calendarEventsIds[i];
      var feedIndex = icsEventsIds.indexOf(currentID);
      
      if(feedIndex  == -1 && calendarEvents[i].recurringEventId == null){
        Logger.log("Deleting old Event " + currentID);
        try{
          Calendar.Events.remove(targetCalendarId, calendarEvents[i].id);
        }
        catch (err){
          Logger.log(err);
        }
      }
    }
}

function processTasks(responses){
  var taskLists = Tasks.Tasklists.list().items;
  var taskList = taskLists[0];
  
  var existingTasks = Tasks.Tasks.list(taskList.id).items || [];
  var existingTasksIds = []
  Logger.log("Fetched " + existingTasks.length + " existing Tasks from " + taskList.title);
  for (var i = 0; i < existingTasks.length; i++){
    existingTasksIds[i] = existingTasks[i].id;
  }
  
  var icsTasksIds = [];
  var vtasks = [];
  
  for each (var resp in responses){
    var jcalData = ICAL.parse(resp);
    var component = new ICAL.Component(jcalData);
    
    vtasks = [].concat(component.getAllSubcomponents("vtodo"), vtasks);
  }
  vtasks.forEach(function(task){ icsTasksIds.push(task.getFirstPropertyValue('uid').toString()); });
  
  Logger.log("---Processing " + vtasks.length + " Tasks.");
  for each (var task in vtasks){
    var newtask = Tasks.newTask();
    newtask.id = task.getFirstPropertyValue("uid").toString();
    newtask.title = task.getFirstPropertyValue("summary").toString();
    var dueDate = task.getFirstPropertyValue("due").toJSDate();
    newtask.due = (dueDate.getFullYear()) + "-" + ("0"+(dueDate.getMonth()+1)).slice(-2) + "-" + ("0" + dueDate.getDate()).slice(-2) + "T" + ("0" + dueDate.getHours()).slice(-2) + ":" + ("0" + dueDate.getMinutes()).slice(-2) + ":" + ("0" + dueDate.getSeconds()).slice(-2)+"Z";
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

function ParseRecurrenceRule(vevent, utcOffset){
  var recurrenceRules = vevent.getAllProperties('rrule');
  var exRules = vevent.getAllProperties('exrule');//deprecated, for compatibility only
  var exDates = vevent.getAllProperties('exdate');
  var rDates = vevent.getAllProperties('rdate');
  var recurrence = [];
  for each (var recRule in recurrenceRules){
    var recIcal = recRule.toICALString();
    var adjustedTime;
    var untilMatch = RegExp("(.*)(UNTIL=)(\\d\\d\\d\\d)(\\d\\d)(\\d\\d)T(\\d\\d)(\\d\\d)(\\d\\d)(;.*|\\b)", "g").exec(recIcal);
    if (untilMatch != null) {
      adjustedTime = new Date(Date.UTC(parseInt(untilMatch[3],10),parseInt(untilMatch[4], 10)-1,parseInt(untilMatch[5],10), parseInt(untilMatch[6],10), parseInt(untilMatch[7],10), parseInt(untilMatch[8],10)));
      adjustedTime = (Utilities.formatDate(new Date(adjustedTime - utcOffset), "etc/GMT", "YYYYMMdd'T'HHmmss'Z'"));
      recIcal = untilMatch[1] + untilMatch[2] + adjustedTime + untilMatch[9];
    }
    recurrence.push(recIcal);
  }
  for each (var exRule in exRules){
    recurrence.push(exRule.toICALString()); 
  }
  for each (var exDate in exDates){
    recurrence.push(exDate.toICALString());
  }
  for each (var rDate in rDates){
    recurrence.push(rDate.toICALString());
  }
  return recurrence;
}

function ParseAttendeeName(veventString){
  var nameMatch = RegExp("(cn=)([^;$:]*)", "gi").exec(veventString);
  if (nameMatch != null && nameMatch.length > 1)
    return nameMatch[2];
  else
    return null;
}

function ParseAttendeeMail(veventString){
  var mailMatch = RegExp("(:mailto:)([^;$:]*)", "gi").exec(veventString);
  if (mailMatch != null && mailMatch.length > 1)
    return mailMatch[2];
  else
    return null;
}

function ParseAttendeeResp(veventString){
  var respMatch = RegExp("(partstat=)([^;$:]*)", "gi").exec(veventString);
  if (respMatch != null && respMatch.length > 1){
    if (['NEEDS-ACTION'].indexOf(respMatch[2].toUpperCase()) > -1) {
      respMatch[2] = 'needsAction';
    } else if (['ACCEPTED','COMPLETED'].indexOf(respMatch[2].toUpperCase()) > -1) {
      respMatch[2] = 'accepted';
    } else if (['DECLINED'].indexOf(respMatch[2].toUpperCase(respMatch[2].toUpperCase())) > -1) {
      respMatch[2] = 'declined';
    } else if (['DELEGATED','IN-PROCESS','TENTATIVE'].indexOf(respMatch[2].toUpperCase())) {
      respMatch[2] = 'tentative';
    } else {
      respMatch[2] = null;
    }
    return respMatch[2];
  }
  else{
    return null;
  }
}

function ParseOrganizerName(veventString){
  /*A regex match is necessary here because ICAL.js doesn't let us directly
  * get the "CN" part of an ORGANIZER property. With something like
  * ORGANIZER;CN="Sally Example":mailto:sally@example.com
  * VEVENT.getFirstPropertyValue('organizer') returns "mailto:sally@example.com".
  * Therefore we have to use a regex match on the VEVENT string instead
  */
  
  var nameMatch = RegExp("organizer(?:;|:)cn=(.*?):", "gi").exec(veventString);
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
