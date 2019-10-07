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
