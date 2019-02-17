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

function ParseRecurrence(rrule){
  //[RRULE docs] https://www.kanzaki.com/docs/ical/rrule.html
  //[ICAL.js Recur] http://mozilla-comm.github.io/ical.js/api/ICAL.Recur.html
  //[GAS RecurrenceRule] https://developers.google.com/apps-script/reference/calendar/recurrence-rule
  
  //I have a pending feature request to simply be able to create a RecurrenceRule object from a RRULE string
  //which would elimiate the need for all of the below parsing: https://issuetracker.google.com/issues/124584372

  var recur = ICAL.Recur.fromString(rrule);
  Logger.log(recur.toString());
  Logger.log(recur.parts);
  Logger.log(recur.parts['BYMONTH']);
  
  var eventRecurrence;
  switch (recur.freq){
    //SECONDLY, MINUTELY, and HOURLY are not supported by Google Calendar (or at least the GAS version)
    case "DAILY":
      eventRecurrence = CalendarApp.newRecurrence().addDailyRule();
      break;
    case "WEEKLY":
      eventRecurrence = CalendarApp.newRecurrence().addWeeklyRule();
      break;
    case "MONTHLY":
      eventRecurrence = CalendarApp.newRecurrence().addMonthlyRule();
      break;
    case "YEARLY":
      eventRecurrence = CalendarApp.newRecurrence().addYearlyRule();
      break;
  }
  
  if (recur.parts['BYDAY'] != null){
    var weekdayArray = [];
    for each (var day in recur.parts['BYDAY']){
      switch (day){
        case "SU":
          weekdayArray.push(CalendarApp.Weekday.SUNDAY);
          break;
        case "MO":
          weekdayArray.push(CalendarApp.Weekday.MONDAY);
          break;
        case "TU":
          weekdayArray.push(CalendarApp.Weekday.TUESDAY);
          break;
        case "WE":
          weekdayArray.push(CalendarApp.Weekday.WEDNESDAY);
          break;
        case "TH":
          weekdayArray.push(CalendarApp.Weekday.THURSDAY);
          break;
        case "FR":
          weekdayArray.push(CalendarApp.Weekday.FRIDAY);
          break;
        case "SA":
          weekdayArray.push(CalendarApp.Weekday.SATURDAY);
          break;
      }
    }
    
    eventRecurrence.onlyOnWeekdays(weekdayArray);
  }
  
  if (recur.parts['BYMONTH'] != null){
    var monthArray = [];
    for each (var month in recur.parts['BYMONTH']){
      switch (month){
        case 1:
          monthArray.push(CalendarApp.Month.JANUARY);
          break;
        case 2:
          monthArray.push(CalendarApp.Month.FEBRUARY);
          break;
        case 3:
          monthArray.push(CalendarApp.Month.MARCH);
          break;
        case 4:
          monthArray.push(CalendarApp.Month.APRIL);
          break;
        case 5:
          monthArray.push(CalendarApp.Month.MAY);
          break;
        case 6:
          monthArray.push(CalendarApp.Month.JUNE);
          break;
        case 7:
          monthArray.push(CalendarApp.Month.JULY);
          break;
        case 8:
          monthArray.push(CalendarApp.Month.AUGUST);
          break;
        case 9:
          monthArray.push(CalendarApp.Month.SEPTEMBER);
          break;
        case 10:
          monthArray.push(CalendarApp.Month.OCTOBER);
          break;
        case 11:
          monthArray.push(CalendarApp.Month.NOVEMBER);
          break;
        case 12:
          monthArray.push(CalendarApp.Month.DECEMBER);
          break;
      }
    }
    
    eventRecurrence.onlyInMonths(monthArray);
  }
  
  if (recur.parts['BYMONTHDAY'] != null){
    Logger.log(recur.parts['BYMONTHDAY']);
    //onlyOnMonthDays cannot take negative values (which is the typical convention for RRULE).
    //I have submitted an issue for that here: https://issuetracker.google.com/issues/124579536
    eventRecurrence.onlyOnMonthDays(recur.parts['BYMONTHDAY']);
  }
  
  if (recur.parts['BYYEARDAY'] != null){
    eventRecurrence.onlyOnYearDays(recur.parts['BYYEARDAY']);
  }
  
  if (recur.parts['BYWEEKNO'] != null){
    eventRecurrence.onlyOnWeeks(recur.parts['BYWEEKNO']);
  }
  
  //Todo: need to handle exclusions
  
  eventRecurrence.interval(recur.interval);
  
  if (recur.count != null)
    eventRecurrence.times(recur.count);
  
  if (recur.until != null){
    if (vtimezone != null)
      recur.until.zone = new ICAL.Timezone(vtimezone);
      
    eventRecurrence.until(recur.until.toJSDate());
  }
  
  if (vtimezone != null)
    eventRecurrence.setTimeZone((new ICAL.Timezone(vtimezone)).toString());
  
  return eventRecurrence;
}

function test(){
  var rrule = "FREQ=MONTHLY;BYMONTHDAY=1,3";//"FREQ=WEEKLY;INTERVAL=1;UNTIL=20190506T045959Z;BYDAY=SU,TH,FR,SA";
  var recurrence = ParseRecurrence(rrule);
  
  var targetCalendar = CalendarApp.getCalendarsByName("ICSTEST")[0];
  //targetCalendar.createEventSeries("Test Event", new Date(2019, 1, 15, 15, 0, 0), new Date(2019, 1, 15, 16, 0, 0), recurrence);
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