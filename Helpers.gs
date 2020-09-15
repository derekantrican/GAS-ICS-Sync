/**
 * Takes an intended frequency in minutes and adjusts it to be the closest 
 * acceptable value to use Google "everyMinutes" trigger setting (i.e. one of
 * the following values: 1, 5, 10, 15, 30).
 *
 * @param {integer} The manually set frequency that the user intends to set
 * @return {integer} The closest valid value to the intended frequency setting 
 */
function getValidTriggerFrequency(origFrequency) {
  if (origFrequency == null) 
    return 1; // if param is empty, just return 1, the lowest valid value.

  
  var adjFrequency = Math.round(origFrequency/5) * 5; // Set the number to be the closest divisible-by-5
  adjFrequency = Math.max(adjFrequency, 1); // Make sure the number is at least 1 (0 is not valid for the trigger)
  adjFrequency = Math.min(adjFrequency, 15); // Make sure the number is at most 15 (will check for the 30 value below)
  
  
  if((adjFrequency == 15) && (Math.abs(origFrequency-30) < Math.abs(origFrequency-15)))
    adjFrequency = 30; // If we adjusted to 15, but the original number is actually closer to 30, set it to 30 instead
  
  Logger.log("Intended frequency = "+origFrequency+", Adjusted frequency = "+adjFrequency);
  return adjFrequency;
}

String.prototype.includes = function(phrase){ 
  return this.indexOf(phrase) > -1;
}

/**
 * Takes an array of ICS calendars and target Google calendars and combines them
 *
 * @param {Array.string} calendarMap - User-defined calendar map
 * @return {Array.string} Condensed calendar map
 */
function condenseCalendarMap(calendarMap){
  var result = [];
  for (var mapping of calendarMap){
    var index = -1;
    for (var i = 0; i < result.length; i++){
      if (result[i][0] == mapping[1]){
        index = i;
        break;
      }
    }

    if (index > -1)
      result[index][1].push(mapping[0]);
    else
      result.push([ mapping[1], [ mapping[0] ] ]);
  }

  return result;
}

/**
 * Removes all triggers for the script's 'startSync' and 'install' function.
 */
function deleteAllTriggers(){
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++){
    if (triggers[i].getHandlerFunction() == "startSync" || triggers[i].getHandlerFunction() == "install" || triggers[i].getHandlerFunction() == "main"){
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Gets the ressource from the specified URLs.
 *
 * @param {Array.string} sourceCalendarURLs - Array with URLs to fetch
 * @return {Array.string} The ressources fetched from the specified URLs
 */
function fetchSourceCalendars(sourceCalendarURLs){
  var result = []
  for (var url of sourceCalendarURLs){
    url = url.replace("webcal://", "https://");
    
    callWithBackoff(function() {
      var urlResponse = UrlFetchApp.fetch(url, { 'validateHttpsCertificates' : false });
      if (urlResponse.getResponseCode() == 200){
        var urlContent = urlResponse.getContentText();
        if(!urlContent.includes("BEGIN:VCALENDAR")){
          Logger.log("[ERROR] Incorrect ics/ical URL: " + url);
          return;
        }
        else{
          result.push(urlContent);
          Logger.log("Result: " + result.length);
          return;
        }     
      }
      else{ //Throw here to make callWithBackoff run again
        throw "Error: Encountered " + urlReponse.getReponseCode() + " when accessing " + url; 
      }
    }, 2);
  }
  
  return result;
}

/**
 * Gets the user's Google Calendar with the specified name.
 * A new Calendar will be created if the user does not have a Calendar with the specified name.
 *
 * @param {string} targetCalendarName - The name of the calendar to return
 * @return {Calendar} The calendar retrieved or created
 */
function setupTargetCalendar(targetCalendarName){
  var targetCalendar = Calendar.CalendarList.list().items.filter(function(cal) {
    return cal.summary == targetCalendarName;
  })[0];
  
  if(targetCalendar == null){
    Logger.log("Creating Calendar: " + targetCalendarName);
    targetCalendar = Calendar.newCalendar();
    targetCalendar.summary = targetCalendarName;
    targetCalendar.description = "Created by GAS";
    targetCalendar.timeZone = Calendar.Settings.get("timezone").value;
    targetCalendar = Calendar.Calendars.insert(targetCalendar);
  }

  return targetCalendar;
}

/**
 * Parses all sources using ical.js.
 * Registers all found timezones with TimezoneService.
 * Creates an Array with all events and adds the event-ids to the provided Array.
 *
 * @param {Array.string} responses - Array with all ical sources
 * @return {Array.ICALComponent} Array with all events found
 */
function parseResponses(responses){
  var result = [];
  for (var resp of responses){
    var jcalData = ICAL.parse(resp);
    var component = new ICAL.Component(jcalData);

    ICAL.helpers.updateTimezones(component);
    var vtimezones = component.getAllSubcomponents("vtimezone");
    for (var tz of vtimezones){
      ICAL.TimezoneService.register(tz);
    }
    
    var allEvents = component.getAllSubcomponents("vevent");
    var calName = component.getFirstPropertyValue("x-wr-calname") || component.getFirstPropertyValue("name");
    if (calName != null)
      allEvents.forEach(function(event){event.addPropertyWithValue("parentCal", calName); });

    result = [].concat(allEvents, result);
  }
  
  if (onlyFutureEvents){
    result = result.filter(function(event){
      try{
        if (event.hasProperty('recurrence-id') || event.hasProperty('rrule') || event.hasProperty('rdate') || event.hasProperty('exdate')){
          //Keep recurrences to properly filter them later on
          return true;
        }
        var eventEnde;
        eventEnde = new ICAL.Time.fromString(event.getFirstPropertyValue('dtend').toString(), event.getFirstProperty('dtend'));
        return (eventEnde.compare(startUpdateTime) >= 0);
      }catch(e){
        return true;
      }
    });
  }
  
  result.forEach(function(event){
    if(event.hasProperty('recurrence-id')){
      var recID = new ICAL.Time.fromString(event.getFirstPropertyValue('recurrence-id').toString(), event.getFirstProperty('recurrence-id'));
      var recUTC = recID.convertToZone(ICAL.TimezoneService.get('UTC')).toString();
    
      icsEventsIds.push(event.getFirstPropertyValue('uid').toString() + "_" + recUTC);
    }
    else{
      icsEventsIds.push(event.getFirstPropertyValue('uid').toString());
    }
  });
 
  return result;
}

/**
 * Creates a Google Calendar event and inserts it to the target calendar.
 *
 * @param {ICAL.Component} event - The event to process
 * @param {string} calendarTz - The timezone of the target calendar
 */
function processEvent(event, calendarTz){
  //------------------------ Create the event object ------------------------
  var newEvent = createEvent(event, calendarTz);
  if (newEvent == null)
    return;
  
  var index = calendarEventsIds.indexOf(newEvent.extendedProperties.private["id"]);
  var needsUpdate = index > -1;
  
  //------------------------ Save instance overrides ------------------------
  //----------- To make sure the parent event is actually created -----------
  if (event.hasProperty('recurrence-id')){
    var recID = new ICAL.Time.fromString(event.getFirstPropertyValue('recurrence-id').toString(), event.getFirstProperty('recurrence-id'));
    newEvent.recurringEventId = recID.convertToZone(ICAL.TimezoneService.get('UTC')).toString();
    Logger.log("Saving event instance for later: " + newEvent.recurringEventId);
    newEvent.extendedProperties.private['rec-id'] = newEvent.extendedProperties.private['id'] + "_" + newEvent.recurringEventId;
    recurringEvents.push(newEvent);
    return;
  }
  else{
    //------------------------ Send event object to gcal ------------------------
    if (needsUpdate){
      if (modifyExistingEvents){
        Logger.log("Updating existing event " + newEvent.extendedProperties.private["id"]);
        newEvent = callWithBackoff(function(){
          return Calendar.Events.update(newEvent, targetCalendarId, calendarEvents[index].id);
        }, 2);
        if (newEvent != null && emailSummary){
          modifiedEvents.push([[newEvent.summary, newEvent.start.date||newEvent.start.dateTime], targetCalendarName]);
        }
      }
    }
    else{
      if (addEventsToCalendar){
        Logger.log("Adding new event " + newEvent.extendedProperties.private["id"]);
        newEvent = callWithBackoff(function(){
          return Calendar.Events.insert(newEvent, targetCalendarId);
        }, 2);
        if (newEvent != null && emailSummary){
          addedEvents.push([[newEvent.summary, newEvent.start.date||newEvent.start.dateTime], targetCalendarName]);
        }
      }
    }
  }
}

/**
 * Creates a Google Calendar Event based on the specified ICALEvent.
 * Will return null if the event has not changed since the last sync.
 * If onlyFutureEvents is set to true:
 * -It will return null if the event has already taken place.
 * -Past instances of recurring events will be removed
 *
 * @param {ICAL.Component} event - The event to process
 * @param {string} calendarTz - The timezone of the target calendar
 * @return {?Calendar.Event} The Calendar.Event that will be added to the target calendar
 */
function createEvent(event, calendarTz){
  event.removeProperty('dtstamp');
  var icalEvent = new ICAL.Event(event, {strictExceptions: true});
  if (onlyFutureEvents && checkSkipEvent(event, icalEvent)){
    return;
  }

  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, icalEvent.toString()).toString();
  if(calendarEventsMD5s.indexOf(digest) >= 0){
    Logger.log("Skipping unchanged Event " + event.getFirstPropertyValue('uid').toString());
    return;
  }

  var newEvent = Calendar.newEvent();
  if(icalEvent.startDate.isDate){ //All-day event
    if (icalEvent.startDate.compare(icalEvent.endDate) == 0){
      //Adjust dtend in case dtstart equals dtend as this is not valid for allday events
      icalEvent.endDate = icalEvent.endDate.adjust(1,0,0,0);
    }

    newEvent = {
      start: { date : icalEvent.startDate.toString() },
      end: { date : icalEvent.endDate.toString() }
    };
  }
  else{ //Normal (not all-day) event
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
    }

    newEvent = {
      start: {
        dateTime : icalEvent.startDate.toString(),
        timeZone : tzid
      },
      end: {
        dateTime : icalEvent.endDate.toString(),
        timeZone : tzid
      },
    };
  }
  
  if (addAttendees && event.hasProperty('attendee')){
    newEvent.attendees = [];
    for (var att of icalEvent.attendees){
      var mail = parseAttendeeMail(att.toICALString());
      if (mail != null){
        var newAttendee = {'email' : mail };

        var name = parseAttendeeName(att.toICALString());
        if (name != null)
          newAttendee['displayName'] = name;

        var resp = parseAttendeeResp(att.toICALString());
        if (resp != null)
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

  if (event.hasProperty('url') && event.getFirstPropertyValue('url').toString() != ''){
    newEvent.source = Calendar.newEventSource()
    newEvent.source.url = event.getFirstPropertyValue('url').toString();
  }

  if (event.hasProperty('sequence')){
    //newEvent.sequence = icalEvent.sequence; Currently disabled as it is causing issues with recurrence exceptions
  }

  if (descriptionAsTitles && event.hasProperty('description'))
    newEvent.summary = icalEvent.description;
  else if (event.hasProperty('summary'))
    newEvent.summary = icalEvent.summary;

  if (addOrganizerToTitle && event.hasProperty('organizer')){
    var organizer = event.getFirstProperty('organizer').getParameter('cn').toString();   
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
    var classString = event.getFirstPropertyValue('class').toString().toLowerCase();
    if (["default", "public", "private", "confidential"].indexOf(classString) > -1)
      newEvent.visibility = classString;
  }

  if (event.hasProperty('transp')){
    var transparency = event.getFirstPropertyValue('transp').toString().toLowerCase();
    if(["opaque", "transparent"].indexOf(transparency) > -1)
      newEvent.transparency = transparency;
  }

  if (icalEvent.startDate.isDate){
    if (0 <= defaultAllDayReminder && defaultAllDayReminder <= 40320){
      newEvent.reminders = { 'useDefault' : false, 'overrides' : [{'method' : 'popup', 'minutes' : defaultAllDayReminder}]};//reminder as defined by the user
    }
    else{
      newEvent.reminders = { 'useDefault' : false, 'overrides' : []};//no reminder
    }
  }
  else{
    newEvent.reminders = { 'useDefault' : true, 'overrides' : []};//will set the default reminders as set at calendar.google.com
  }
  
  if (addAlerts){
    var valarms = event.getAllSubcomponents('valarm');
    if (valarms.length > 0){
      var overrides = [];
      for (var valarm of valarms){
        var trigger = valarm.getFirstPropertyValue('trigger').toString();
        try{
          var alarmTime = new ICAL.Time.fromString(trigger);
          trigger = alarmTime.subtractDateTz(icalEvent.startDate).toString();
        }catch(e){}
        if (overrides.length < 5){ //Google supports max 5 reminder-overrides
          var timer = parseNotificationTime(trigger);
          if (0 <= timer && timer <= 40320)
            overrides.push({'method' : 'popup', 'minutes' : timer});
        }
      }

      if (overrides.length > 0){
        newEvent.reminders = {
          'useDefault' : false,
          'overrides' : overrides
        };
      }
    }
  }
  
  if (icalEvent.isRecurring()){
    // Calculate targetTZ's UTC-Offset
    var calendarUTCOffset = 0;
    var jsTime = new Date();
    var utcTime = new Date(Utilities.formatDate(jsTime, "Etc/GMT", "HH:mm:ss MM/dd/yyyy"));
    var tgtTime = new Date(Utilities.formatDate(jsTime, calendarTz, "HH:mm:ss MM/dd/yyyy"));
    calendarUTCOffset = tgtTime - utcTime;
    newEvent.recurrence = parseRecurrenceRule(event, calendarUTCOffset);
  }
  
  newEvent.extendedProperties = { private: { MD5 : digest, fromGAS : "true", id : icalEvent.uid } };
  return newEvent;
}

/**
 * Checks if the provided event has taken place in the past.
 * Removes all past instances of the provided icalEvent object.
 *
 * @param {ICAL.Component} event - The event to process
 * @param {ICAL.Event} icalEvent - The event to process as ICAL.Event object
 * @return {boolean} Wether it's a past event or not
 */
function checkSkipEvent(event, icalEvent){
  if (icalEvent.isRecurrenceException()){
    if((icalEvent.startDate.compare(startUpdateTime) < 0) && (icalEvent.recurrenceId.compare(startUpdateTime) < 0)){
      Logger.log("Skipping past recurrence exception");
      return true; 
    }
  }
  else if(icalEvent.isRecurring()){
    var skip = false; //Indicates if the recurring event and all its instances are in the past
    if (icalEvent.endDate.compare(startUpdateTime) < 0){//Parenting recurring event is the past
      var recur = event.getFirstPropertyValue('rrule');
      var dtstart = event.getFirstPropertyValue('dtstart');
      var iter = recur.iterator(dtstart);
      var newStartDate;
      for (var next = iter.next(); next; next = iter.next()) {
        if (next.compare(startUpdateTime) < 0) {
          continue;
        }
        
        newStartDate = next;
        break;
      }
      
      if (newStartDate != null){//At least one instance is in the future
        var diff = newStartDate.subtractDate(icalEvent.startDate);
        icalEvent.endDate.addDuration(diff);
        var newEndDate = icalEvent.endDate;
        icalEvent.endDate = newEndDate;
        icalEvent.startDate = newStartDate;
        Logger.log("Adjusted RRULE to exclude past instances");
      }
      else{//All instances are in the past
        icalEvent.component.removeProperty('rrule');
        Logger.log("Removed RRULE");
        skip = true;
      }
    }

    //Check and filter recurrence-exceptions
    for (i=0; i<icalEvent.except.length; i++){
      //Exclude the instance if it was moved from future to past
      if((icalEvent.except[i].startDate.compare(startUpdateTime) < 0) && (icalEvent.except[i].recurrenceId.compare(startUpdateTime) >= 0)){
        Logger.log("Creating EXDATE for exception at " + icalEvent.except[i].recurrenceId.toString());
        icalEvent.component.addPropertyWithValue('exdate', icalEvent.except[i].recurrenceId.toString());
      }//Re-add the instance if it is moved from past to future
      else if((icalEvent.except[i].startDate.compare(startUpdateTime) >= 0) && (icalEvent.except[i].recurrenceId.compare(startUpdateTime) < 0)){
        Logger.log("Creating RDATE for exception at " + icalEvent.except[i].recurrenceId.toString());
        icalEvent.component.addPropertyWithValue('rdate', icalEvent.except[i].recurrenceId.toString());
        skip = false;
      }
    }
    
    if(skip){//Completely remove the event as all instances of it are in the past
      icsEventsIds.splice(icsEventsIds.indexOf(event.getFirstPropertyValue('uid').toString()),1);
      Logger.log("Skipping past recurring event " + event.getFirstPropertyValue('uid').toString());
      return true;
    }
  }
  else{//normal events
    if (icalEvent.endDate.compare(startUpdateTime) < 0){
      icsEventsIds.splice(icsEventsIds.indexOf(event.getFirstPropertyValue('uid').toString()),1);
      Logger.log("Skipping previous event " + event.getFirstPropertyValue('uid').toString());
      return true;
    }
  }
  return false;
}

/**
 * Patches an existing event instance with the provided Calendar.Event.
 * The instance that needs to be updated is identified by the recurrence-id of the provided event.
 *
 * @param {Calendar.Event} recEvent - The event instance to process
 */
function processEventInstance(recEvent){
  Logger.log("ID: " + recEvent.extendedProperties.private["id"] + " | Date: "+ recEvent.recurringEventId.substring(0,10));
  var recIDStart = new Date(recEvent.recurringEventId);
  recIDStart = new ICAL.Time.fromJSDate(recIDStart, true);

  var calendarEvents = Calendar.Events.list(targetCalendarId, 
    { timeZone : "etc/GMT", 
      singleEvents : true, 
      privateExtendedProperty : "fromGAS=true", 
      privateExtendedProperty : "id=" + recEvent.extendedProperties.private['id']
    }).items;

  Logger.log("Found " + calendarEvents.length + " possible instances");
  var eventInstanceToPatch = calendarEvents.filter(function(item){
    try{
      var origStart = item.originalStartTime.dateTime || item.originalStartTime.date;
      var instanceStart = new ICAL.Time.fromString(origStart);

      return (instanceStart.compare(recIDStart) == 0);
    }catch(e){
      Logger.log("Error: " + e);
      return 0; 
    }
  });

  if (eventInstanceToPatch.length == 0){
    Logger.log("No Instance matched, adding as new event!");
    try{
      Calendar.Events.insert(recEvent, targetCalendarId);
    }
    catch(error){
      Logger.log(error); 
    }
  }
  else{
    try{
      Logger.log("Patching existing event instance");
      Calendar.Events.patch(recEvent, targetCalendarId, eventInstanceToPatch[0].id);
    }
    catch(error){
      Logger.log(error); 
    }
  }
}

/**
 * Deletes all events from the target calendar that no longer exist in the source calendars.
 * If onlyFutureEvents is set to true, events that have taken place since the last sync are also removed.
 */
function processEventCleanup(){
  for (var i = 0; i < calendarEvents.length; i++){
      var currentID = calendarEventsIds[i];
      var feedIndex = icsEventsIds.indexOf(currentID);
      
      if(feedIndex  == -1 && calendarEvents[i].recurringEventId == null){
        Logger.log("Deleting old event " + currentID);
        try{
          Calendar.Events.remove(targetCalendarId, calendarEvents[i].id);
          if (emailSummary){
            removedEvents.push([[calendarEvents[i].summary, calendarEvents[i].start.date||calendarEvents[i].start.dateTime], targetCalendarName]);
          }
        }
        catch (err){
          Logger.log(err);
        }
      }
    }
}

/**
 * Processes and adds all vtodo components as Tasks to the user's Google Account
 *
 * @param {Array.string} responses - Array with all ical sources
 */
function processTasks(responses){
  var taskLists = Tasks.Tasklists.list().items;
  var taskList = taskLists[0];
  
  var existingTasks = Tasks.Tasks.list(taskList.id).items || [];
  var existingTasksIds = []
  Logger.log("Fetched " + existingTasks.length + " existing tasks from " + taskList.title);
  for (var i = 0; i < existingTasks.length; i++){
    existingTasksIds[i] = existingTasks[i].id;
  }
  
  var icsTasksIds = [];
  var vtasks = [];
  
  for (var resp of responses){
    var jcalData = ICAL.parse(resp);
    var component = new ICAL.Component(jcalData);
    
    vtasks = [].concat(component.getAllSubcomponents("vtodo"), vtasks);
  }

  vtasks.forEach(function(task){ icsTasksIds.push(task.getFirstPropertyValue('uid').toString()); });
  
  Logger.log("\tProcessing " + vtasks.length + " tasks");
  for (var task of vtasks){
    var newtask = Tasks.newTask();
    newtask.id = task.getFirstPropertyValue("uid").toString();
    newtask.title = task.getFirstPropertyValue("summary").toString();
    var dueDate = task.getFirstPropertyValue("due").toJSDate();
    newtask.due = (dueDate.getFullYear()) + "-" + ("0"+(dueDate.getMonth()+1)).slice(-2) + "-" + ("0" + dueDate.getDate()).slice(-2) + "T" + ("0" + dueDate.getHours()).slice(-2) + ":" + ("0" + dueDate.getMinutes()).slice(-2) + ":" + ("0" + dueDate.getSeconds()).slice(-2)+"Z";
    
    Tasks.Tasks.insert(newtask, taskList.id);
  }
  Logger.log("\tDone processing tasks");
  
  //-------------- Remove old Tasks -----------
  // ID can't be used as identifier as the API reassignes a random id at task creation
  if(removeEventsFromCalendar){
    Logger.log("Checking " + existingTasksIds.length + " tasks for removal");
    for (var i = 0; i < existingTasksIds.length; i++){
      var currentID = existingTasks[i].id;
      var feedIndex = icsTasksIds.indexOf(currentID);
      
      if(feedIndex == -1){
        Logger.log("Deleting old task " + currentID);
        Tasks.Tasks.remove(taskList.id, currentID);
      }
    }

    Logger.log("Done removing tasks");
  }
  //----------------------------------------------------------------
}

/**
 * Parses the provided ICAL.Component to find all recurrence rules.
 *
 * @param {ICAL.Component} vevent - The event to parse
 * @param {number} utcOffset - utc offset of the target calendar
 * @return {Array.String} Array with all recurrence components found in the provided event
 */
function parseRecurrenceRule(vevent, utcOffset){
  var recurrenceRules = vevent.getAllProperties('rrule');
  var exRules = vevent.getAllProperties('exrule');//deprecated, for compatibility only
  var exDates = vevent.getAllProperties('exdate');
  var rDates = vevent.getAllProperties('rdate');

  var recurrence = [];
  for (var recRule of recurrenceRules){
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

  for (var exRule of exRules){
    recurrence.push(exRule.toICALString()); 
  }

  for (var exDate of exDates){
    recurrence.push(exDate.toICALString());
  }

  for (var rDate of rDates){
    recurrence.push(rDate.toICALString());
  }

  return recurrence;
}

/**
 * Parses the provided string to find the name of an Attendee.
 * Will return null if no name is found.
 *
 * @param {string} veventString - The string to parse
 * @return {?String} The Attendee's name found in the string, null if no name was found
 */
function parseAttendeeName(veventString){
  var nameMatch = RegExp("(cn=)([^;$:]*)", "gi").exec(veventString);
  if (nameMatch != null && nameMatch.length > 1)
    return nameMatch[2];
  else
    return null;
}

/**
 * Parses the provided string to find the mail adress of an Attendee.
 * Will return null if no mail adress is found.
 *
 * @param {string} veventString - The string to parse
 * @return {?String} The Attendee's mail adress found in the string, null if nothing was found
 */
function parseAttendeeMail(veventString){
  var mailMatch = RegExp("(:mailto:)([^;$:]*)", "gi").exec(veventString);
  if (mailMatch != null && mailMatch.length > 1)
    return mailMatch[2];
  else
    return null;
}

/**
 * Parses the provided string to find the response of an Attendee.
 * Will return null if no response is found or the response string is not supported by google calendar.
 *
 * @param {string} veventString - The string to parse
 * @return {?String} The Attendee's response found in the string, null if nothing was found or unsupported
 */
function parseAttendeeResp(veventString){
  var respMatch = RegExp("(partstat=)([^;$:]*)", "gi").exec(veventString);
  if (respMatch != null && respMatch.length > 1){
    if (['NEEDS-ACTION'].indexOf(respMatch[2].toUpperCase()) > -1) {
      respMatch[2] = 'needsAction';
    } 
    else if (['ACCEPTED', 'COMPLETED'].indexOf(respMatch[2].toUpperCase()) > -1) {
      respMatch[2] = 'accepted';
    } 
    else if (['DECLINED'].indexOf(respMatch[2].toUpperCase()) > -1) {
      respMatch[2] = 'declined';
    } 
    else if (['DELEGATED', 'IN-PROCESS', 'TENTATIVE'].indexOf(respMatch[2].toUpperCase())) {
      respMatch[2] = 'tentative';
    } 
    else {
      respMatch[2] = null;
    }
    return respMatch[2];
  }
  else{
    return null;
  }
}

/**
 * Parses the provided string to find the notification time of an event.
 * Will return 0 by default.
 *
 * @param {string} notificationString - The string to parse
 * @return {number} The notification time in minutes
 */
function parseNotificationTime(notificationString){
  //https://www.kanzaki.com/docs/ical/duration-t.html
  var reminderTime = 0;
  
  //We will assume all notifications are BEFORE the event
  if (notificationString[0] == "+" || notificationString[0] == "-")
    notificationString = notificationString.substr(1);
  
  notificationString = notificationString.substr(1); //Remove "P" character
  
  var minuteMatch = RegExp("\\d+M", "g").exec(notificationString);
  var hourMatch = RegExp("\\d+H", "g").exec(notificationString);
  var dayMatch = RegExp("\\d+D", "g").exec(notificationString);
  var weekMatch = RegExp("\\d+W", "g").exec(notificationString);
  
  if (weekMatch != null){
    reminderTime += parseInt(weekMatch[0].slice(0, -1)) & 7 * 24 * 60; //Remove the "W" off the end
    
    return reminderTime; //Return the notification time in minutes
  }
  else{
    if (minuteMatch != null)
      reminderTime += parseInt(minuteMatch[0].slice(0, -1)); //Remove the "M" off the end
    
    if (hourMatch != null)
      reminderTime += parseInt(hourMatch[0].slice(0, -1)) * 60; //Remove the "H" off the end
    
    if (dayMatch != null)
      reminderTime += parseInt(dayMatch[0].slice(0, -1)) * 24 * 60; //Remove the "D" off the end
    
    return reminderTime; //Return the notification time in minutes
  }
}

/**
* Sends an email summary with added/modified/deleted events.
*/            
function sendSummary() {
  var subject;
  var body;
  
  var subject = `GAS-ICS-Sync Execution Summary: ${addedEvents.length} new, ${modifiedEvents.length} modified, ${removedEvents.length} deleted`;
  addedEvents = condenseCalendarMap(addedEvents);
  modifiedEvents = condenseCalendarMap(modifiedEvents);
  removedEvents = condenseCalendarMap(removedEvents);
  
  body = "GAS-ICS-Sync made the following changes to your calendar:<br/>";
  for (var tgtCal of addedEvents){
    body += `<br/>${tgtCal[0]}: ${tgtCal[1].length} added events<br/><ul>`;
    for (var addedEvent of tgtCal[1]){
      body += "<li>" + addedEvent[0] + " at " + addedEvent[1] + "</li>";
    }
    body += "</ul>";
  }
  
  for (var tgtCal of modifiedEvents){
    body += `<br/>${tgtCal[0]}: ${tgtCal[1].length} modified events<br/><ul>`;
    for (var addedEvent of tgtCal[1]){
      body += "<li>" + addedEvent[0] + " at " + addedEvent[1] + "</li>";
    }
    body += "</ul>";
  }
  
  for (var tgtCal of removedEvents){
    body += `<br/>${tgtCal[0]}: ${tgtCal[1].length} removed events<br/><ul>`;
    for (var addedEvent of tgtCal[1]){
      body += "<li>" + addedEvent[0] + " at " + addedEvent[1] + "</li>";
    }
    body += "</ul>";
  }
  
  body += "<br/><br/>Do you have any problems or suggestions? Contact us at <a href='https://github.com/derekantrican/GAS-ICS-Sync/'>github</a>.";
  var message = {
    to: email,
    subject: subject,
    htmlBody: body,
    name: "GAS-ICS-Sync"
  };
  
  MailApp.sendEmail(message);
}

/**
 * Runs the specified function with exponential backoff and returns the result.
 * Will return null if the function did not succeed afterall.
 *
 * @param {function} func - The function that should be executed
 * @param {Number} maxRetries - How many times the function should try if it fails
 * @return {?Calendar.Event} The Calendar.Event that was added in the calendar, null if func did not complete successfully
 */
function callWithBackoff(func, maxRetries) {
  var tries = 0;
  var result;
  do{
    Utilities.sleep(tries * 100);
    tries++;
    try{
      result = func();
      return result;
    }
    catch(e){
      Logger.log("Error, Retrying..." + e );
    }
  }while(tries <= maxRetries );

  return null;
}

/**
 * Checks for a new version of the script at https://github.com/derekantrican/GAS-ICS-Sync/releases.
 * Will notify the user once if a new version was released.
 */
function checkForUpdate(){
  var alreadyAlerted = PropertiesService.getScriptProperties().getProperty("alertedForNewVersion");
  if (alreadyAlerted == null){
    try{
      var thisVersion = 5.4;
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
