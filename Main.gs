const defaultMaxRetries = 10; // Maximum number of retries for api functions (with exponential backoff)

function install() {
  // Delete any already existing triggers so we don't create excessive triggers
  deleteAllTriggers();

  // Schedule sync routine to explicitly repeat and schedule the initial sync
  var adjustedMinutes = getValidTriggerFrequency(appSettings.howFrequent);
  if (adjustedMinutes >= 60) {
    ScriptApp.newTrigger("startSync")
      .timeBased()
      .everyHours(adjustedMinutes / 60)
      .create();
  } else {
    ScriptApp.newTrigger("startSync")
      .timeBased()
      .everyMinutes(adjustedMinutes)
      .create();
  }
  ScriptApp.newTrigger("startSync").timeBased().after(1000).create();

  // Schedule sync routine to look for update once per day using everyDays
  ScriptApp.newTrigger("checkForUpdate")
    .timeBased()
    .everyDays(1)
    .create();
}

function uninstall(){
  deleteAllTriggers();
}

// Per-calendar global variables (must be reset before processing each new calendar!)
var calendarEvents = [];
var calendarEventsIds = [];
var icsEventsIds = [];
var calendarEventsMD5s = [];
var recurringEvents = [];
var targetCalendarId;
var vevents = [];
var calendarConfig;

// Per-session global variables (must NOT be reset before processing each new calendar!)
var addedEvents = [];
var modifiedEvents = [];
var removedEvents = [];

// Syncing logic can set this to true to cause the Google Apps Script "Executions" dashboard to report failure
var reportOverallFailure = false;

function startSync(){
  if (PropertiesService.getUserProperties().getProperty('LastRun') > 0 && (new Date().getTime() - PropertiesService.getUserProperties().getProperty('LastRun')) < 360000) {
    Logger.log("Another iteration is currently running! Exiting...");
    return;
  }

  PropertiesService.getUserProperties().setProperty('LastRun', new Date().getTime());

  //Disable email notification if no mail adress is provided
  let emailSummary = appSettings.emailSummary && (appSettings.email != "");

  for (let calendar of sourceCalendars){
    calendarConfig = Object.assign({}, defaultSettings, calendar);
    //------------------------ Reset globals ------------------------
    calendarEvents = [];
    calendarEventsIds = [];
    icsEventsIds = [];
    calendarEventsMD5s = [];
    recurringEvents = [];
    vevents = [];

    //------------------------ Get target calendar information------------------------
    var targetCalendar = setupTargetCalendar();
    targetCalendarId = targetCalendar.id;
    Logger.log(`Syncing '${calendarConfig.sourceCalendarName}' (URL: ${calendarConfig.sourceURL}) to ${calendarConfig.targetCalendarName} (ID: ${targetCalendarId})`);
    //------------------------ Parse existing events --------------------------
    if (calendarConfig.addEventsToCalendar || calendarConfig.modifyExistingEvents || calendarConfig.removeEventsFromCalendar){
      var response = getSourceCalendarEvents();
      getTargetCalendarEvents();
      //------------------------ Parse ical events --------------------------
      vevents = parseSourceCalendarEvents(response);
      Logger.log(`Parsed ${vevents.length} events from ical sources`);
    }

    //------------------------ Process ical events ------------------------
    if (calendarConfig.addEventsToCalendar || calendarConfig.modifyExistingEvents){
      Logger.log(`Processing ${vevents.length} events`);
      var calendarTz =
        callWithBackoff(function(){
          return Calendar.Settings.get("timezone").value;
        }, defaultMaxRetries);

      vevents.forEach(function(e){
        processEvent(e, calendarTz);
      });

      Logger.log("Done processing events");
    }

    //------------------------ Remove old events from calendar ------------------------
    if (calendarConfig.removeEventsFromCalendar){
      Logger.log(`Checking ${calendarEvents.length} events for removal`);
      processEventCleanup();
      Logger.log("Done checking events for removal");
    }

    //------------------------ Process Tasks ------------------------
    if (calendarConfig.addTasks){
      processTasks(responses);
    }

    //------------------------ Add Recurring Event Instances ------------------------
    Logger.log(`Processing ${recurringEvents.length} recurrence instances`);
    for (var recEvent of recurringEvents){
      processEventInstance(recEvent);
    }
  }

  if ((addedEvents.length + modifiedEvents.length + removedEvents.length) > 0 && emailSummary){
    sendSummary();
  }
  Logger.log("Sync finished!");
  PropertiesService.getUserProperties().setProperty('LastRun', 0);

  if (reportOverallFailure) {
    // Cause the Google Apps Script "Executions" dashboard to show a failure
    // (the message text does not seem to be logged anywhere)
    throw new Error('The sync operation produced errors. See log for details.');
  }
}
