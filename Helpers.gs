/**
 * Formats the date and time according to the format specified in the configuration.
 *
 * @param {string} date The date to be formatted.
 * @return {string} The formatted date string.
 */
function formatDate(date) {
  const year = date.slice(0,4);
  const month = date.slice(5,7);
  const day = date.slice(8,10);
  let formattedDate;

  if (dateFormat == "YYYY/MM/DD") {
    formattedDate = year + "/" + month + "/" + day
  }
  else if (dateFormat == "DD/MM/YYYY") {
    formattedDate = day + "/" + month + "/" + year
  }
  else if (dateFormat == "MM/DD/YYYY") {
    formattedDate = month + "/" + day + "/" + year
  }
  else if (dateFormat == "YYYY-MM-DD") {
    formattedDate = year + "-" + month + "-" + day
  }
  else if (dateFormat == "DD-MM-YYYY") {
    formattedDate = day + "-" + month + "-" + year
  }
  else if (dateFormat == "MM-DD-YYYY") {
    formattedDate = month + "-" + day + "-" + year
  }
  else if (dateFormat == "YYYY.MM.DD") {
    formattedDate = year + "." + month + "." + day
  }
  else if (dateFormat == "DD.MM.YYYY") {
    formattedDate = day + "." + month + "." + year
  }
  else if (dateFormat == "MM.DD.YYYY") {
    formattedDate = month + "." + day + "." + year
  }

  if (date.length < 11) {
    return formattedDate
  }

  const time = date.slice(11,16)
  const timeZone = date.slice(19)

  return formattedDate + " at " + time + " (UTC" + (timeZone == "Z" ? "": timeZone) + ")"
}


/**
 * Takes an intended frequency in minutes and adjusts it to be the closest
 * acceptable value to use Google "everyMinutes" trigger setting (i.e. one of
 * the following values: 1, 5, 10, 15, 30).
 *
 * @param {?integer} The manually set frequency that the user intends to set.
 * @return {integer} The closest valid value to the intended frequency setting. Defaulting to 15 if no valid input is provided.
 */
function getValidTriggerFrequency(origFrequency) {
  if (!origFrequency > 0) {
    Logger.log("No valid frequency specified. Defaulting to 15 minutes.");
    return 15;
  }

  // Limit the original frequency to 1440
  origFrequency = Math.min(origFrequency, 1440);

  var acceptableValues = [5, 10, 15, 30].concat(
    Array.from({ length: 24 }, (_, i) => (i + 1) * 60)
  ); // [5, 10, 15, 30, 60, 120, ..., 1440]

  // Find the smallest acceptable value greater than or equal to the original frequency
  var roundedUpValue = acceptableValues.find(value => value >= origFrequency);

  Logger.log(
    "Intended frequency = " + origFrequency + ", Adjusted frequency = " + roundedUpValue
  );
  return roundedUpValue;
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
      result[index][1].push([mapping[0],mapping[2]]);
    else
      result.push([ mapping[1], [[mapping[0],mapping[2]]] ]);
  }

  return result;
}

/**
 * Removes all triggers for the script's 'startSync' and 'install' function.
 */
function deleteAllTriggers(){
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++){
    if (["startSync","install","main","checkForUpdate"].includes(triggers[i].getHandlerFunction())){
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
  for (var source of sourceCalendarURLs){
    var url = source[0].replace("webcal://", "https://");
    var colorId = source[1];

    try {
      callWithBackoff(function() {
        var urlResponse = UrlFetchApp.fetch(url, { 'validateHttpsCertificates' : false, 'muteHttpExceptions' : true });
        if (urlResponse.getResponseCode() == 200){
          var icsContent = urlResponse.getContentText()
          const icsRegex = RegExp("(BEGIN:VCALENDAR.*?END:VCALENDAR)", "s")
          var urlContent = icsRegex.exec(icsContent);
          if (urlContent == null){
            // Microsoft Outlook has a bug that sometimes results in incorrectly formatted ics files. This tries to fix that problem.
            // Add END:VEVENT for every BEGIN:VEVENT that's missing it
            const veventRegex = /BEGIN:VEVENT(?:(?!END:VEVENT).)*?(?=.BEGIN|.END:VCALENDAR|$)/sg;
            icsContent = icsContent.replace(veventRegex, (match) => match + "\nEND:VEVENT");

            // Add END:VCALENDAR if missing
            if (!icsContent.endsWith("END:VCALENDAR")){
                icsContent += "\nEND:VCALENDAR";
            }
            urlContent = icsRegex.exec(icsContent)
            if (urlContent == null){
              Logger.log("[ERROR] Incorrect ics/ical URL: " + url)
              reportOverallFailure = true;
              return
            }
            Logger.log("[WARNING] Microsoft is incorrectly formatting ics/ical at: " + url)
          }
          result.push([urlContent[0], colorId]);
          return;
        }
        else{ //Throw here to make callWithBackoff run again
          throw "Error: Encountered HTTP error " + urlResponse.getResponseCode() + " when accessing " + url;
        }
      }, defaultMaxRetries);
    }
    catch (e) {
      reportOverallFailure = true;
    }
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
  var targetCalendar = Calendar.CalendarList.list({showHidden: true, maxResults: 250}).items.filter(function(cal) {
    return ((cal.summaryOverride || cal.summary) == targetCalendarName) &&
                (cal.accessRole == "owner" || cal.accessRole == "writer");
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
  for (var itm of responses){
    var resp = itm[0];
    var colorId = itm[1];
    var jcalData = ICAL.parse(resp);
    var component = new ICAL.Component(jcalData);

    ICAL.helpers.updateTimezones(component);
    var vtimezones = component.getAllSubcomponents("vtimezone");
    for (var tz of vtimezones){
      ICAL.TimezoneService.register(tz);
    }

    var allEvents = component.getAllSubcomponents("vevent");
    if (colorId != undefined)
      allEvents.forEach(function(event){event.addPropertyWithValue("color", colorId);});

    var calName = component.getFirstPropertyValue("x-wr-calname") || component.getFirstPropertyValue("name");
    if (calName != null)
      allEvents.forEach(function(event){event.addPropertyWithValue("parentCal", calName); });

    result = [].concat(allEvents, result);
  }

  //No need to process cancelled events as they will be added to gcal's trash anyway
  result = result.filter(function(event){
    try{
      return (event.getFirstPropertyValue('status').toString().toLowerCase() != "cancelled");
    }catch(e){
      return true;
    }
  });

  result = filterResults(result);

  result.forEach(function(event){
    if (!event.hasProperty('uid')){
      event.updatePropertyWithValue('uid', Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, event.toString()).toString(), Utilities.Charset.UTF_8);
    }
    if(event.hasProperty('recurrence-id')){
      let recID = new ICAL.Time.fromString(event.getFirstPropertyValue('recurrence-id').toString(), event.getFirstProperty('recurrence-id'));
      if (event.getFirstProperty('recurrence-id').getParameter('tzid')){
        let recUTCOffset = 0;
        let tz = event.getFirstProperty('recurrence-id').getParameter('tzid').toString();
        if (tz in tzidreplace){
          tz = tzidreplace[tz];
        }
        let jsTime = new Date(event.getFirstPropertyValue('recurrence-id').toString());
        let utcTime = new Date(Utilities.formatDate(jsTime, "Etc/GMT", "HH:mm:ss MM/dd/yyyy"));
        let tgtTime = new Date(Utilities.formatDate(jsTime, tz, "HH:mm:ss MM/dd/yyyy"));
        recUTCOffset = (tgtTime - utcTime)/-1000;
        recID = recID.adjust(0,0,0,recUTCOffset).toString() + "Z";
        event.updatePropertyWithValue('recurrence-id', recID);
      }
      icsEventsIds.push(event.getFirstPropertyValue('uid').toString() + "_" + recID);
    }
    else{
      icsEventsIds.push(event.getFirstPropertyValue('uid').toString());
    }
  });

  return result;
}

/**
 * Applies filters to source events based on filters defined in filters.gs
 *
 * @param {Array.ICALComponent} Array with all events from the source calendars
 * @return {Array.ICALComponent} Array with filtered events
 */
function filterResults(events){
  Logger.log(`Applying ${filters.length} filters on ${events.length} events.`);

  for (var filter of filters){
    filter.parameter = filter.parameter.toLowerCase();
    events = events.filter(function(event){
      try{
        if (["dtstart", "dtend"].includes(filter.parameter)){
          let referenceDate = new ICAL.Time.fromJSDate(new Date(), true).adjust(filter.offset,0,0,0);
          if (event.hasProperty('rrule') || event.hasProperty('rdate')) {
            if ((filter.comparison === ">" && filter.type === "exclude")||(filter.comparison === "<" && filter.type === "include")) {
              event = modifyRecurrenceEnd(event, referenceDate, filter.parameter);
            } else if ((filter.comparison === "<" && filter.type === "exclude")||(filter.comparison === ">" && filter.type === "include")) {
              event = modifyRecurrenceStart(event, referenceDate, filter.parameter);
            }
            return event !== null;
          }
          else{
            let eventTime = new ICAL.Time.fromString(event.getFirstPropertyValue(filter.parameter).toString(), event.getFirstProperty(filter.parameter));
            switch (filter.comparison){
              case ">":
                return ((eventTime.compare(referenceDate) > 0) ^ (filter.type == "exclude"));
              case "<":
                return ((eventTime.compare(referenceDate) < 0) ^ (filter.type == "exclude"));
              case "default":
                return true;
            }
          }
        }
        else{
          let regexString = `${(["equals", "begins with"].includes(filter.comparison)) ? "^" : ""}(${filter.criterias.join("|")})${(filter.comparison == "equals") ? "$" : ""}`;
          let regex = new RegExp(regexString);
          let result = regex.test(event.getFirstPropertyValue(filter.parameter).toString()) ^ (filter.type == "exclude");
          if (!result && event.hasProperty('recurrence-id')){
            let id = event.getFirstPropertyValue('uid');
            Logger.log(`Filtering recurrence instance of ${id} at ${event.getFirstPropertyValue('dtstart').toICALString()}`);
            let indx = events.findIndex((e) => e.getFirstPropertyValue('uid') == id && !e.hasProperty('recurrence-id'));
            if (!events[indx].hasProperty('exdate')){
              events[indx].addProperty(new ICAL.Property('exdate'));
            }
            let exdates = events[indx].getFirstProperty('exdate').getValues().concat(event.getFirstPropertyValue('recurrence-id'));
            events[indx].getFirstProperty('exdate').setValues(exdates);
          }
          return result;
        } 
      }
      catch(e){
        Logger.log(e);
        return (filter.type == "exclude");
      }
    });
  }
  
  Logger.log(`${events.length} events left.`);
  return events;
}

/**
 * Modifies the end of the given recurrence series.
 *
 * @param {ICAL.Component} event - The event to modify
 * @param {ICAL.Time} referenceDate - The new recurrence end date
 * @param {string} filterParameter - The parameter to filter on
 * @return {ICAL.Component|null} The modified event or null if no instances are within the range
 */
function modifyRecurrenceEnd(event, referenceDate, filterParameter) {
  let eventRefDate = new ICAL.Time.fromString(event.getFirstPropertyValue('dtstart').toString(), event.getFirstProperty('dtstart'));
  if (filterParameter.toLowerCase() === "dtend"){
    let eventEnd = new ICAL.Time.fromString(event.getFirstPropertyValue('dtend').toString(), event.getFirstProperty('dtend'));
    var eventDurartion = eventEnd.subtractDate(eventRefDate);
    eventRefDate = eventEnd;
  }
  let icalEvent = new ICAL.Event(event);

  if (eventRefDate.compare(referenceDate) >= 0){
    return null;
  }

  if (event.hasProperty('rrule')){
    let rrule = event.getFirstProperty('rrule');
    let recur = rrule.getFirstValue();
    var dtstart = event.getFirstPropertyValue('dtstart');
    var expand = new ICAL.RecurExpansion({component: event, dtstart: dtstart});
    var next;
    var lastStartDate = null;
    var newCount = 0;
    // Iterate through the recurrence instances to find the last valid one before referenceDate
    while (next = expand.next()) {
      if (filterParameter.toLowerCase() === "dtstart") {
        if (next.compare(referenceDate) > 0) {
          break;
        }
        newCount++;
        lastStartDate = next;
      }
      else if (filterParameter.toLowerCase() === "dtend") {
        let tempEnd = next.clone();
        tempEnd.addDuration(eventDurartion);
        if (tempEnd.compare(referenceDate) > 0) {
          break;
        }
        newCount++;
        lastStartDate = next;
      }
    }

    // Remove EXDATEs that are after the endDate
    var exDates = event.getAllProperties('exdate');
    exDates.forEach(function(e) {
      var ex = new ICAL.Time.fromString(e.getFirstValue().toString(), e);
      if (ex.compare(lastStartDate) > 0) {
        event.removeProperty(e);
      }
      else{
        newCount++
      }
    });

    if (newCount == 0){
      event.removeProperty('rrule');
    }
    else{
      if (recur.isByCount()) {
        recur.count = newCount;
        rrule.setValue(recur);
      }
      else{
        recur.until = referenceDate.clone();
        rrule.setValue(recur);
      }
    }
  }

  // Adjust RDATEs to exclude any dates beyond the endDate
  var rdates = event.getAllProperties('rdate');
  rdates.forEach(function(r) {
    var vals = r.getValues();
    vals = vals.filter(function(v) {
      var valTime = new ICAL.Time.fromString(v.toString(), r);
      if (filterParameter.toLowerCase() === "dtend") {
        valTime.addDuration(eventDurartion);
      }
      return valTime.compare(referenceDate) <= 0;
    });
    if (vals.length === 0) {
      event.removeProperty(r);
    } else if (vals.length === 1) {
      r.setValue(vals[0]);
    } else if (vals.length > 1) {
      r.setValues(vals);
    }
  });

  //Check and filter recurrence-exceptions
  if (filterParameter.toLowerCase() === "dtend"){
    for (let key in icalEvent.exceptions) {
      let recIdEnd = icalEvent.exceptions[key].recurrenceId.clone();
      recIdEnd.addDuration(eventDurartion);
      if((icalEvent.exceptions[key].endDate.compare(referenceDate) > 0) && (recIdEnd.compare(referenceDate) <= 0)){
        icalEvent.component.addPropertyWithValue('exdate', icalEvent.exceptions[key].recurrenceId.toString());
      }
      else if((icalEvent.exceptions[key].endDate.compare(referenceDate) <= 0) && (recIdEnd.compare(referenceDate) > 0)){
        icalEvent.component.addPropertyWithValue('rdate', icalEvent.exceptions[key].recurrenceId.toString());
      }
    }
  }
  else if (filterParameter.toLowerCase() === "dtstart"){
    for (let key in icalEvent.exceptions) {
      if((icalEvent.exceptions[key].startDate.compare(referenceDate) < 0) && (icalEvent.exceptions[key].recurrenceId.compare(referenceDate) >= 0)){
        icalEvent.component.addPropertyWithValue('rdate', icalEvent.exceptions[key].recurrenceId.toString());
      }
      else if((icalEvent.exceptions[key].startDate.compare(referenceDate) >= 0) && (icalEvent.exceptions[key].recurrenceId.compare(referenceDate) < 0)){
        icalEvent.component.addPropertyWithValue('exdate', icalEvent.exceptions[key].recurrenceId.toString());
      }
    }
  }

  return event;
}

/**
 * Modifies the start of the given recurrence series.
 *
 * @param {ICAL.Component} event - The event to modify
 * @param {ICAL.Time} referenceDate - The new recurrence start date
 * @param {string} filterParameter - The parameter to filter on
 * @return {ICAL.Component|null} The modified event or null if no instances are within the range
 */
function modifyRecurrenceStart(event, referenceDate, filterParameter) {
  let eventRefDate = new ICAL.Time.fromString(event.getFirstPropertyValue('dtstart').toString(), event.getFirstProperty('dtstart'));
  if (filterParameter.toLowerCase() === "dtend"){
    let eventEnd = new ICAL.Time.fromString(event.getFirstPropertyValue('dtend').toString(), event.getFirstProperty('dtend'));
    var eventDurartion = eventEnd.subtractDate(eventRefDate);
    eventRefDate = eventEnd;
  }
  let icalEvent = new ICAL.Event(event);

  if (eventRefDate.compare(referenceDate) < 0){
    var dtstart = event.getFirstPropertyValue('dtstart');
    var expand = new ICAL.RecurExpansion({component: event, dtstart: dtstart});
    var next;
    var newStartDate = null;
    var countskipped = 0;
    while (next = expand.next()) {
      if (filterParameter.toLowerCase() === "dtstart"){
        if (next.compare(referenceDate) < 0) {
          countskipped ++;
          continue;
        }
      }
      else if(filterParameter.toLowerCase() === "dtend"){
        let tempEnd = next.clone();
        tempEnd.addDuration(eventDurartion);
        if (tempEnd.compare(referenceDate) < 0) {
          countskipped ++;
          continue;
        }
      } 
      
      newStartDate = next;
      break;
    }
    
    if (newStartDate === null) {
      return null;
    }
    
    var diff = newStartDate.subtractDate(icalEvent.startDate);
    icalEvent.endDate.addDuration(diff);
    var newEndDate = icalEvent.endDate;
    icalEvent.endDate = newEndDate;
    icalEvent.startDate = newStartDate;

    if (event.hasProperty('rrule') ){
      let rrule = event.getFirstProperty('rrule');
      let recur = rrule.getFirstValue();
      var exDates = event.getAllProperties('exdate');
      exDates.forEach(function(e){
        var ex = new ICAL.Time.fromString(e.getFirstValue().toString(), e);
        if (ex < newStartDate){
          event.removeProperty(e);
          if (recur.isByCount()) {
            countskipped++;
          }
        }
      });

      if (recur.isByCount()) {
        recur.count -= countskipped;
        rrule.setValue(recur);
      }
    }
  }

  var rdates = event.getAllProperties('rdate');
  rdates.forEach(function(r){
    var vals = r.getValues();
    vals = vals.filter(function(v){
      var valTime = new ICAL.Time.fromString(v.toString(), r);
      if (filterParameter.toLowerCase() === "dtend") {
        valTime.addDuration(eventDurartion);
      }
      return (valTime.compare(referenceDate) >= 0 && valTime.compare(icalEvent.startDate) > 0)
    });
    if (vals.length == 0){
      event.removeProperty(r);
    }
    else if(vals.length == 1){
      r.setValue(vals[0]);
    }
    else if(vals.length > 1){
      r.setValues(vals);
    }
  });

  //Check and filter recurrence-exceptions
  if (filterParameter.toLowerCase() === "dtend"){
    for (let key in icalEvent.exceptions) {
      let recIdEnd = icalEvent.exceptions[key].recurrenceId.clone();
      recIdEnd.addDuration(eventDurartion);
      //Exclude the instance if it was moved from future to past
      if((icalEvent.exceptions[key].endDate.compare(referenceDate) < 0) && (recIdEnd.compare(referenceDate) >= 0)){
        icalEvent.component.addPropertyWithValue('exdate', icalEvent.exceptions[key].recurrenceId.toString());
      }//Re-add the instance if it is moved from past to future
      else if((icalEvent.exceptions[key].endDate.compare(referenceDate) >= 0) && (recIdEnd.compare(referenceDate) < 0)){
        icalEvent.component.addPropertyWithValue('rdate', icalEvent.exceptions[key].recurrenceId.toString());
      }
    }
  }
  else if (filterParameter.toLowerCase() === "dtstart"){
    for (let key in icalEvent.exceptions) {
      //Exclude the instance if it was moved from future to past
      if((icalEvent.exceptions[key].startDate.compare(referenceDate) < 0) && (icalEvent.exceptions[key].recurrenceId.compare(referenceDate) >= 0)){
        icalEvent.component.addPropertyWithValue('exdate', icalEvent.exceptions[key].recurrenceId.toString());
      }//Re-add the instance if it is moved from past to future
      else if((icalEvent.exceptions[key].startDate.compare(referenceDate) >= 0) && (icalEvent.exceptions[key].recurrenceId.compare(referenceDate) < 0)){
        icalEvent.component.addPropertyWithValue('rdate', icalEvent.exceptions[key].recurrenceId.toString());
      }
    }
  }

  return event;
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
    Logger.log("Saving event instance for later: " + newEvent.recurringEventId);
    recurringEvents.push(newEvent);
    return;
  }
  else{
    //------------------------ Send event object to gcal ------------------------
    if (needsUpdate){
      if (modifyExistingEvents){
        oldEvent = calendarEvents[index]
        Logger.log("Updating existing event " + newEvent.extendedProperties.private["id"]);
        try{
          newEvent = callWithBackoff(function(){
            return Calendar.Events.update(newEvent, targetCalendarId, calendarEvents[index].id);
          }, defaultMaxRetries);
        }
        catch (e){
          Logger.log(`Operation failed with error "${e}"`);
          reportOverallFailure = true;
        }
        if (newEvent != null && emailSummary){
          modifiedEvents.push([[oldEvent.summary, newEvent.summary, oldEvent.start.date||oldEvent.start.dateTime, newEvent.start.date||newEvent.start.dateTime, oldEvent.end.date||oldEvent.end.dateTime, newEvent.end.date||newEvent.end.dateTime, oldEvent.location, newEvent.location, oldEvent.description, newEvent.description], targetCalendarName]);
        }
      }
    }
    else{
      if (addEventsToCalendar){
        Logger.log("Adding new event " + newEvent.extendedProperties.private["id"]);
        try{
          newEvent = callWithBackoff(function(){
            return Calendar.Events.insert(newEvent, targetCalendarId);
          }, defaultMaxRetries);
        }
        catch (e){
          Logger.log(`Operation failed with error "${e}"`);
          reportOverallFailure = true;
        }
        if (newEvent != null && emailSummary){
          addedEvents.push([[newEvent.summary, newEvent.start.date||newEvent.start.dateTime, newEvent.end.date||newEvent.end.dateTime, newEvent.location, newEvent.description], targetCalendarName]);
        }
      }
    }
  }
}

/**
 * Creates a Google Calendar Event based on the specified ICALEvent.
 * Will return null if the event has not changed since the last sync.
 *
 * @param {ICAL.Component} event - The event to process
 * @param {string} calendarTz - The timezone of the target calendar
 * @return {?Calendar.Event} The Calendar.Event that will be added to the target calendar
 */
function createEvent(event, calendarTz){
  event.removeProperty('dtstamp');
  var icalEvent = new ICAL.Event(event);

  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, icalEvent.toString(), Utilities.Charset.UTF_8).toString();
  if(calendarEventsMD5s.indexOf(digest) >= 0){
    Logger.log("Skipping unchanged Event " + event.getFirstPropertyValue('uid').toString());
    return;
  }

  var newEvent =
    callWithBackoff(function() {
        return Calendar.newEvent();
      }, defaultMaxRetries);
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
    newEvent = {
      start: {
        dateTime : icalEvent.startDate.toString(),
        timeZone : validateTimeZone(icalEvent.startDate.timezone || icalEvent.startDate.zone, calendarTz)
      },
      end: {
        dateTime : icalEvent.endDate.toString(),
        timeZone : validateTimeZone(icalEvent.endDate.timezone || icalEvent.endDate.zone, calendarTz)
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

  if (event.hasProperty('url') && event.getFirstPropertyValue('url').toString().substring(0,4) == 'http'){
    newEvent.source = callWithBackoff(function() {
          return Calendar.newEventSource();
        }, defaultMaxRetries);
    newEvent.source.url = event.getFirstPropertyValue('url').toString();
    newEvent.source.title = 'link';
  }

  if (event.hasProperty('sequence')){
    //newEvent.sequence = icalEvent.sequence; Currently disabled as it is causing issues with recurrence exceptions
  }

  if (descriptionAsTitles && event.hasProperty('description'))
    newEvent.summary = icalEvent.description;
  else if (event.hasProperty('summary'))
    newEvent.summary = icalEvent.summary;

  if (event.hasProperty('organizer')){
    var organizerName = event.getFirstProperty('organizer').getParameter('cn');
    var organizerMail = event.getFirstProperty('organizer').getParameter('mailto');
    newEvent.organizer = callWithBackoff(function() {
          return Calendar.newEventOrganizer();
        }, defaultMaxRetries);
    if (organizerName)
      newEvent.organizer.displayName = organizerName.toString();
    if (organizerMail)
      newEvent.organizer.email = organizerMail.toString();

    if (addOrganizerToTitle && organizerName){
        newEvent.summary = organizerName + ": " + newEvent.summary;
    }
  }

  if (addCalToTitle && event.hasProperty('parentCal')){
    var calName = event.getFirstPropertyValue('parentCal');
    newEvent.summary = "(" + calName + ") " + newEvent.summary;
  }

  if (event.hasProperty('description'))
    newEvent.description = icalEvent.description;

  if (event.hasProperty('location'))
    newEvent.location = icalEvent.location;

  var validVisibilityValues = ["default", "public", "private", "confidential"];
  if ( validVisibilityValues.includes(overrideVisibility.toLowerCase()) ) {
    newEvent.visibility = overrideVisibility.toLowerCase();
  } else if (event.hasProperty('class')){
    var classString = event.getFirstPropertyValue('class').toString().toLowerCase();
    if (validVisibilityValues.includes(classString))
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

  switch (addAlerts) {
    case "yes":
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
      break;
    case "no":
      newEvent.reminders = {
        'useDefault' : false,
        'overrides' : []
      };
      break;
    default:
    case "default":
      newEvent.reminders = {
        'useDefault' : true,
        'overrides' : []
      };
      break;
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

  if (event.hasProperty('recurrence-id')){
    newEvent.recurringEventId = event.getFirstPropertyValue('recurrence-id').toString();
    newEvent.extendedProperties.private['rec-id'] = newEvent.extendedProperties.private['id'] + "_" + newEvent.recurringEventId;
  }

  if (event.hasProperty('color')){
    let colorID = event.getFirstPropertyValue('color').toString();
    if (Object.keys(CalendarApp.EventColor).includes(colorID)){
      newEvent.colorId = CalendarApp.EventColor[colorID];
    }else if(Object.values(CalendarApp.EventColor).includes(colorID)){
      newEvent.colorId = colorID;
    }; //else unsupported value
  }

  return newEvent;
}

/**
 * Patches an existing event instance with the provided Calendar.Event.
 * The instance that needs to be updated is identified by the recurrence-id of the provided event.
 *
 * @param {Calendar.Event} recEvent - The event instance to process
 */
function processEventInstance(recEvent){
  Logger.log("ID: " + recEvent.extendedProperties.private["id"] + " | Date: "+ recEvent.recurringEventId);

  var eventInstanceToPatch = callWithBackoff(function(){
    return Calendar.Events.list(targetCalendarId,
      { singleEvents : true,
        privateExtendedProperty : "fromGAS=true",
        privateExtendedProperty : "rec-id=" + recEvent.extendedProperties.private["id"] + "_" + recEvent.recurringEventId
      }).items;
  }, defaultMaxRetries);

  if (eventInstanceToPatch == null || eventInstanceToPatch.length == 0){
    if (recEvent.recurringEventId.length == 10){
      recEvent.recurringEventId += "T00:00:00Z";
    }
    else if (recEvent.recurringEventId.substr(-1) !== "Z"){
      recEvent.recurringEventId += "Z";
    }
    eventInstanceToPatch = callWithBackoff(function(){
       return Calendar.Events.list(targetCalendarId,
        { singleEvents : true,
          orderBy : "startTime",
          maxResults: 1,
          timeMin : recEvent.recurringEventId,
          privateExtendedProperty : "fromGAS=true",
          privateExtendedProperty : "id=" + recEvent.extendedProperties.private["id"]
        }).items;
    }, defaultMaxRetries);
  }

  if (eventInstanceToPatch !== null && eventInstanceToPatch.length == 1){
    if (modifyExistingEvents){
      Logger.log("Updating existing event instance");
      callWithBackoff(function(){
        Calendar.Events.update(recEvent, targetCalendarId, eventInstanceToPatch[0].id);
      }, defaultMaxRetries);
    }
  }
  else{
    if (addEventsToCalendar){
      Logger.log("No Instance matched, adding as new event!");
      callWithBackoff(function(){
        Calendar.Events.insert(recEvent, targetCalendarId);
      }, defaultMaxRetries);
    }
  }
}

/**
 * Deletes all events from the target calendar that no longer exist in the source calendars.
 * If removePastEventsFromCalendar is set to false, events that have taken place will not be removed.
 */
function processEventCleanup(){
  for (var i = 0; i < calendarEvents.length; i++){
      var currentID = calendarEventsIds[i];
      var feedIndex = icsEventsIds.indexOf(currentID);

      if(feedIndex  == -1                                             // Event is no longer in source
        && calendarEvents[i].recurringEventId == null                 // And it's not a recurring event
        && (                                                          // And one of:
          removePastEventsFromCalendar                                // We want to remove past events
          || new Date(calendarEvents[i].start.dateTime) > new Date()  // Or the event is in the future
          || new Date(calendarEvents[i].start.date) > new Date()      // (2 different ways event start can be stored)
        )
      )
      {
        Logger.log("Deleting old event " + currentID);
        try{
          callWithBackoff(function(){
            Calendar.Events.remove(targetCalendarId, calendarEvents[i].id);
          }, defaultMaxRetries);
        }
        catch (e){
          Logger.log(`Operation failed with error "${e}"`);
          reportOverallFailure = true;
        }

        if (emailSummary){
          removedEvents.push([[calendarEvents[i].summary, calendarEvents[i].start.date||calendarEvents[i].start.dateTime, calendarEvents[i].end.date||calendarEvents[i].end.dateTime, calendarEvents[i].location, calendarEvents[i].description], targetCalendarName]);
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
 * Validates provided Timezone descriptor and if needed replaces it with an IANA timezone descriptor.
 *
 * @param {string} tzid - Timezone descriptor to validate
 * @return {string} Valid IANA timezone descriptor
 */
function validateTimeZone(tzid, calendarTz){
  tzid = tzid.toString();
  let IanaTZ;
  if (tzids.indexOf(tzid) == -1){
    if (tzid in tzidreplace){
      IanaTZ = tzidreplace[tzid];
    }
    else{//floating time
      IanaTZ = calendarTz;
    }
    Logger.log("Converting ICS timezone " + tzid + " to Google Calendar (IANA) timezone " + IanaTZ);
  }
  return IanaTZ || tzid;
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
    if (recRule.getParameter('tzid')){
      let tz = recRule.getParameter('tzid').toString();
      if (tz in tzidreplace){
        tz = tzidreplace[tz];
      }
      recRule.setParameter('tzid', tz);
    }
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
    if (exRule.getParameter('tzid')){
      let tz = exRule.getParameter('tzid').toString();
      if (tz in tzidreplace){
        tz = tzidreplace[tz];
      }
      exRule.setParameter('tzid', tz);
    }
    recurrence.push(exRule.toICALString());
  }

  for (var exDate of exDates){
    if (exDate.getParameter('tzid')){
      let tz = exDate.getParameter('tzid').toString();
      if (tz in tzidreplace){
        tz = tzidreplace[tz];
      }
      exDate.setParameter('tzid', tz);
    }
    recurrence.push(exDate.toICALString());
  }

  for (var rDate of rDates){
    if (rDate.getParameter('tzid')){
      let tz = rDate.getParameter('tzid').toString();
      if (tz in tzidreplace){
        tz = tzidreplace[tz];
      }
      rDate.setParameter('tzid', tz);
    }
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

  var subject = `${customEmailSubject ? customEmailSubject : "GAS-ICS-Sync Execution Summary"}: ${addedEvents.length} new, ${modifiedEvents.length} modified, ${removedEvents.length} deleted`;
  addedEvents = condenseCalendarMap(addedEvents);
  modifiedEvents = condenseCalendarMap(modifiedEvents);
  removedEvents = condenseCalendarMap(removedEvents);

  body = "GAS-ICS-Sync made the following changes to your calendar:<br/>";
  for (var tgtCal of addedEvents){
    body += `<br/>${tgtCal[0]}: ${tgtCal[1].length} added events<br/><ul>`;
    for (var addedEvent of tgtCal[1]){
      body += "<li>"
        + "Name: " + addedEvent[0][0] + "<br/>"
        + "Start: " + formatDate(addedEvent[0][1]) + "<br/>"
        + "End: " + formatDate(addedEvent[0][2]) + "<br/>"
        + (addedEvent[0][3] ? ("Location: " + addedEvent[0][3] + "<br/>") : "")
        + (addedEvent[0][4] ? ("Description: " + addedEvent[0][4] + "<br/>") : "")
        + "</li>";
    }
    body += "</ul>";
  }

  for (var tgtCal of modifiedEvents){
    body += `<br/>${tgtCal[0]}: ${tgtCal[1].length} modified events<br/><ul>`;
    for (var modifiedEvent of tgtCal[1]){
      body += "<li>"
        + (modifiedEvent[0][0] != modifiedEvent[0][1] ? ("<del>Name: " + modifiedEvent[0][0] + "</del><br/>") : "")
        + "Name: " + modifiedEvent[0][1] + "<br/>"
        + (modifiedEvent[0][2] != modifiedEvent[0][3] ? ("<del>Start: " + formatDate(modifiedEvent[0][2]) + "</del><br/>") : "")
        + " Start: " + formatDate(modifiedEvent[0][3]) + "<br/>"
        + (modifiedEvent[0][4] != modifiedEvent[0][5] ? ("<del>End: " + formatDate(modifiedEvent[0][4]) + "</del><br/>") : "")
        + " End: " + formatDate(modifiedEvent[0][5]) + "<br/>"
        + (modifiedEvent[0][6] != modifiedEvent[0][7] ? ("<del>Location: " + (modifiedEvent[0][6] ? modifiedEvent[0][6] : "") + "</del><br/>") : "")
        + (modifiedEvent[0][7] ? (" Location: " + modifiedEvent[0][7] + "<br/>") : "")
        + (modifiedEvent[0][8] != modifiedEvent[0][9] ? ("<del>Description: " + (modifiedEvent[0][8] ? modifiedEvent[0][8] : "") + "</del><br/>") : "")
        + (modifiedEvent[0][9] ? (" Description: " + modifiedEvent[0][9] + "<br/>") : "")
        + "</li>";
    }
    body += "</ul>";
  }

  for (var tgtCal of removedEvents){
    body += `<br/>${tgtCal[0]}: ${tgtCal[1].length} removed events<br/><ul>`;
    for (var removedEvent of tgtCal[1]){
      body += "<li>"
        + "<del>Name: " + removedEvent[0][0] + "</del><br/>"
        + "<del>Start: " + formatDate(removedEvent[0][1]) + "</del><br/>"
        + "<del>End: " + formatDate(removedEvent[0][2]) + "</del><br/>"
        + (removedEvent[0][3] ? ("<del>Location: " + removedEvent[0][3] + "</del><br/>") : "")
        + (removedEvent[0][4] ? ("<del>Description: " + removedEvent[0][4] + "</del><br/>") : "")
        + "</li>";
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
var backoffRecoverableErrors = [
  "service invoked too many times in a short time",
  "rate limit exceeded",
  "internal error",
  "http error 403", // forbidden
  "http error 408", // request timeout
  "http error 423", // locked
  "http error 500", // internal server error
  "http error 503", // service unavailable
  "http error 504"  // gateway timeout
];
function callWithBackoff(func, maxRetries) {
  var tries = 0;
  var result;
  while ( tries <= maxRetries ) {
    tries++;
    try{
      result = func();
      return result;
    }
    catch(err){
      err = err.message  || err;
      if ( err.includes("is not a function")  || !backoffRecoverableErrors.some(function(e){
              return err.toLowerCase().includes(e);
            }) ) {
        throw err;
      } else if ( tries > maxRetries) {
        Logger.log(`Error, giving up after trying ${maxRetries} times [${err}]`);
        return null;
      } else {
        Logger.log( "Error, Retrying... [" + err  +"]");
        Utilities.sleep (Math.pow(2,tries)*100) +
                            (Math.round(Math.random() * 100));
      }
    }
  }
  return null;
}

/**
 * Checks for a new version of the script at https://github.com/derekantrican/GAS-ICS-Sync/releases.
 * Will notify the user once if a new version was released.
 */
function checkForUpdate(){
  // No need to check if we can't alert anyway
  if (email == "")
    return;

  var lastAlertedVersion = PropertiesService.getScriptProperties().getProperty("alertedForNewVersion");
  try {
    var thisVersion = 5.8;
    var latestVersion = getLatestVersion();

    if (latestVersion > thisVersion && latestVersion != lastAlertedVersion){
      MailApp.sendEmail(email,
        `Version ${latestVersion} of GAS-ICS-Sync is available! (You have ${thisVersion})`,
        "You can see the latest release here: https://github.com/derekantrican/GAS-ICS-Sync/releases");

      PropertiesService.getScriptProperties().setProperty("alertedForNewVersion", latestVersion);
    }
  }
  catch (e){}

  function getLatestVersion(){
    var json_encoded = UrlFetchApp.fetch("https://api.github.com/repos/derekantrican/GAS-ICS-Sync/releases?per_page=1");
    var json_decoded = JSON.parse(json_encoded);
    var version = json_decoded[0]["tag_name"];
    return Number(version);
  }
}
