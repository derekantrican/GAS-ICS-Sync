/*
*=========================================
*       INSTALLATION INSTRUCTIONS
*=========================================
*
* 1) Make a copy:
*      New Interface: Go to the project overview icon on the left (looks like this: ⓘ), then click the "copy" icon on the top right (looks like two files on top of each other)
*      Old Interface: Click in the menu "File" > "Make a copy..." and make a copy to your Google Drive
* 2) Code Settings: Change lines 25-29 to be the settings that you want to use
* 3) Install:
*      New Interface: Make sure your toolbar says "install" to the right of "Debug", then click "Run"
*      Old Interface: Click "Run" > "Run function" > "install"
* 4) Authorize: You will be prompted to authorize the program and will need to click "Advanced" > "Go to GAS-ICS-Sync (unsafe)"
*    (For steps to follow in authorization, see this video: https://youtu.be/_5k10maGtek?t=1m22s )
* 5) Calendars and Settings: Click on "Deploy" (in the upper-right)-->"Test Deployments". In the dialog box that pops up, click the URL under web app. This should bring up the Calendar Manager webpage. From here you can add/edit/delete calendars and their associated settings.  A file called calendars.json will be saved to your My Drive folder in Google Drive.  You can manually edit the json calendar settings if you don't want to use the html interface.
* 6) You can also run "startSync" if you want to sync only once (New Interface: change the dropdown to the right of "Debug" from "install" to "startSync")
*
* **To stop the Script from running click in the menu "Run" > "Run function" > "uninstall" (New Interface: change the dropdown to the right of "Debug" from "install" to "uninstall")
*
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
* If you would like to donate and support the project,
* you can do that here: https://www.paypal.me/jonasg0b1011001
*
*=========================================
*             CONTRIBUTORS
*=========================================
* Andrew Brothers
* Github: https://github.com/agentd00nut
* Twitter: @abrothers656
*
* Joel Balmer
* Github: https://github.com/JoelBalmer
*
* Blackwind
* Github: https://github.com/blackwind
*
* Jonas Geissler
* Github: https://github.com/jonas0b1011001
*/


//=====================================================================================================
//!!!!!!!!!!!!!!!! DO NOT EDIT BELOW HERE UNLESS YOU REALLY KNOW WHAT YOU'RE DOING !!!!!!!!!!!!!!!!!!!!
//=====================================================================================================

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  var htmlOutput = template.evaluate().setTitle('Calendar Manager');
  return htmlOutput;
}

//this function is only for html to make sure user can only install after App Settings have been set.
var appSettingsFlag = false;
function validateInstall() {
  if (appSettingsFlag) {
    // Check for required settings if necessary
    Logger.log("Install started from html.")
    return true;
  }else{
  Logger.log("Install from html not allowed because no App Settings found.")
  return false;
  }
}

//Save appSettings.json to Google Drive (My Drive "root" directory)
function updateAppSettings(jsonString) {
  const folder = DriveApp.getRootFolder();
  const files = folder.getFilesByName('appSettings.json');

  if (files.hasNext()) {
    const file = files.next();
    file.setContent(jsonString);
  } else {
    //will create appSettings.json file if one doesn't exist
    folder.createFile('appSettings.json', jsonString, MimeType.PLAIN_TEXT);
  }
}

//retrieve App Settings from appSettings.json for use in index.html
function getAppSettings() {
  const folder = DriveApp.getRootFolder();
  const files = folder.getFilesByName('appSettings.json');

  if (files.hasNext()) {
    const file = files.next();
    const jsonContent = file.getBlob().getDataAsString();
    return JSON.parse(jsonContent); // Parse the JSON content
  } else {
    Logger.log("App Settings json file not found.");
    return false;
  }
}

var appSettings = getAppSettings();
var howFrequent = appSettings.howFrequent;
var emailSummary = appSettings.emailSummary === true;
var email = appSettings.email;
var customEmailSubject = appSettings.customEmailSubject;
var dateFormat = appSettings.dateFormat;

// Grab calendars.json from Google Drive (My Drive "root" directory)
function updateCalendars(jsonString) {
  const folder = DriveApp.getRootFolder();
  const files = folder.getFilesByName('calendars.json');

  if (files.hasNext()) {
    const file = files.next();
    //saves new key/value pair or updates existing pair
    file.setContent(jsonString);
  } else {
    //will create calendars.json file if one doesn't exist
    folder.createFile('calendars.json', jsonString, MimeType.PLAIN_TEXT);
  }
}

function getCalendars() {
  const folder = DriveApp.getRootFolder();
  const files = folder.getFilesByName('calendars.json');

  if (files.hasNext()) {
    const file = files.next();
    const jsonContent = file.getBlob().getDataAsString();
    return JSON.parse(jsonContent); // Parse the JSON content
  } else {
    return {}; // Return an empty object if the file doesn't exist yet
  }
}

var defaultMaxRetries = 10; // Maximum number of retries for api functions (with exponential backoff)

function install() {
  // Delete any already existing triggers so we don't create excessive triggers
  deleteAllTriggers();

  // Schedule sync routine to explicitly repeat and schedule the initial sync
  var adjustedMinutes = getValidTriggerFrequency(howFrequent);
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
  Logger.log("Uninstall successful.")
}

var startUpdateTime;

// Per-calendar global variables (must be reset before processing each new calendar!)
var calendarEvents = [];
var calendarEventsIds = [];
var icsEventsIds = [];
var calendarEventsMD5s = [];
var recurringEvents = [];
var targetCalendarId;

// Per-session global variables (must NOT be reset before processing each new calendar!)
var addedEvents = [];
var modifiedEvents = [];
var removedEvents = [];

//Set default values in case they aren't specified in the json (specifically checkboxes=false because json doesn't store false).  Will override later if json value present.
function getDefaultCalendarConfig() {
    return {
        //ParamUpdate
        color: "",
        sourceSyncDelay: "",
        addEventsToCalendar: false,
        modifyExistingEvents: false,
        removeEventsFromCalendar: false,
        onlyFutureEvents: false,
        getPastDaysIfOnlyFutureEvents: "",
        removePastEventsFromCalendar: false,
        addOrganizerToTitle: false,
        descriptionAsTitles: false,
        addCalToTitle: false,
        addAlerts: "yes",
        addAttendees: false,
        defaultAllDayReminder: -1,
        overrideVisibility: "",
        addTasks: false,
    };
}

function startSync(){
  if (PropertiesService.getUserProperties().getProperty('LastRun') > 0 && (new Date().getTime() - PropertiesService.getUserProperties().getProperty('LastRun')) < 360000) {
    Logger.log("Another iteration is currently running! Exiting...");
    return;
  }

  PropertiesService.getUserProperties().setProperty('LastRun', new Date().getTime());
  var currentDate = new Date();
  //Disable email notification if no mail adress is provided
  emailSummary = emailSummary && email != "";

  //retrieve calendar data from calendars.json
  function getCalendarsFromJson() {
    var files = DriveApp.getFilesByName('calendars.json');
    if (!files.hasNext()) {
        throw new Error('calendars.json not found');
    }
    var file = files.next();
    var content = file.getBlob().getDataAsString();
    return JSON.parse(content);
  }
  var sourceCalendars = getCalendarsFromJson();

  for (var key in sourceCalendars){
    if (sourceCalendars.hasOwnProperty(key)) {
      var calendar = sourceCalendars[key];
      var calendarConfig = Object.assign(getDefaultCalendarConfig(), calendar);
    }

    //------------------------ Reset globals ------------------------
    calendarEvents = [];
    calendarEventsIds = [];
    icsEventsIds = [];
    calendarEventsMD5s = [];
    recurringEvents = [];
    var vevents;

if (calendarConfig.onlyFutureEvents) {
    startUpdateTime = new ICAL.Time.fromJSDate(new Date(currentDate.setDate(currentDate.getDate() - calendarConfig.getPastDaysIfOnlyFutureEvents)));
}

//------------------------ Determine whether to sync each calendar based on SyncDelay ------------------------
    let sourceSyncDelay = Number(calendarConfig.sourceSyncDelay)*60*1000;
    let currentTime = Number(new Date().getTime());
    let lastSyncTime = Number(PropertiesService.getUserProperties().getProperty(calendarConfig.sourceCalendarName));
    var lastSyncDelta = currentTime - lastSyncTime;

    if (isNaN(sourceSyncDelay)) {
      Logger.log("Syncing " + calendarConfig.sourceCalendarName + " because no SyncDelay defined.");
    } else if (lastSyncDelta >= sourceSyncDelay) {
      Logger.log("Syncing " + calendarConfig.sourceCalendarName + " because lastSyncDelta ("+ (lastSyncDelta/60/1000).toFixed(1) + ") is greater than sourceSyncDelay (" + (sourceSyncDelay/60/1000).toFixed(0) + ").");
    } else if (lastSyncDelta < sourceSyncDelay) {
      Logger.log("Skipping " + calendarConfig.sourceCalendarName + " because lastSyncDelta ("+ (lastSyncDelta/60/1000).toFixed(1) + ") is less than sourceSyncDelay (" + (sourceSyncDelay/60/1000).toFixed(0) + ").");
      continue;
    }

    //------------------------ Fetch URL items ------------------------
    var responses = fetchSourceCalendars([[calendarConfig.sourceURL, calendarConfig.color]]);
    //Skip the source calendar if a 5xx or 4xx error is returned.  This prevents deleting all of the existing entries if the URL call fails.
    if (responses.length == 0){
      Logger.log("Error Syncing " + calendarConfig.sourceCalendarName + ". Skipping...");
      continue;
      }
    Logger.log("Syncing " + calendarConfig.sourceCalendarName + " calendar to " + calendarConfig.targetCalendarName);

    //------------------------ Get target calendar information------------------------
    var targetCalendar = setupTargetCalendar(calendarConfig.targetCalendarName);
    targetCalendarId = targetCalendar.id;
    Logger.log("Working on target calendar: " + targetCalendarId);

    //------------------------ Parse existing events --------------------------
    if(calendarConfig.addEventsToCalendar || calendarConfig.modifyExistingEvents || calendarConfig.removeEventsFromCalendar){
      var eventList =
        callWithBackoff(function(){
            return Calendar.Events.list(targetCalendarId, {showDeleted: false, privateExtendedProperty: 'fromGAS=' + calendarConfig.sourceCalendarName, maxResults: 2500});
        }, defaultMaxRetries);
      calendarEvents = [].concat(calendarEvents, eventList.items);
      //loop until we received all events
      while(typeof eventList.nextPageToken !== 'undefined'){
        eventList = callWithBackoff(function(){
          return Calendar.Events.list(targetCalendarId, {showDeleted: false, privateExtendedProperty: 'fromGAS=' + calendarConfig.sourceCalendarName, maxResults: 2500, pageToken: eventList.nextPageToken});
        }, defaultMaxRetries);

        if (eventList != null)
          calendarEvents = [].concat(calendarEvents, eventList.items);
      }
      Logger.log("Fetched " + calendarEvents.length + " existing events from " + calendarConfig.targetCalendarName);
      for (var i = 0; i < calendarEvents.length; i++){
        if (calendarEvents[i].extendedProperties != null){
          calendarEventsIds[i] = calendarEvents[i].extendedProperties.private["rec-id"] || calendarEvents[i].extendedProperties.private["id"];
          calendarEventsMD5s[i] = calendarEvents[i].extendedProperties.private["MD5"];
        }
      }

      //------------------------ Parse ical events --------------------------
      vevents = parseResponses(responses, icsEventsIds, calendarConfig);
      Logger.log("Parsed " + vevents.length + " events from ical sources");
    }

    //------------------------ Process ical events ------------------------
    if (calendarConfig.addEventsToCalendar || calendarConfig.modifyExistingEvents){
      Logger.log("Processing " + vevents.length + " events");
      var calendarTz =
        callWithBackoff(function(){
          return Calendar.Settings.get("timezone").value;
        }, defaultMaxRetries);

      vevents.forEach(function(e){
        processEvent(e, calendarTz, targetCalendarId, calendarConfig.sourceCalendarName, calendarConfig);
      });

      Logger.log("Done processing events");
    }

    //------------------------ Remove old events from calendar ------------------------
    if(calendarConfig.removeEventsFromCalendar){
      Logger.log("Checking " + calendarEvents.length + " events for removal");
      processEventCleanup(calendarConfig.sourceURL, calendarConfig.removePastEventsFromCalendar, calendarConfig.targetCalendarName);
      Logger.log("Done checking events for removal");
    }

    //------------------------ Process Tasks ------------------------
    if (calendarConfig.addTasks){
      processTasks(responses);
    }

    //------------------------ Add Recurring Event Instances ------------------------
    Logger.log("Processing " + recurringEvents.length + " Recurrence Instances!");
    for (var recEvent of recurringEvents){
      processEventInstance(recEvent, calendarConfig);
    }
    //Set last sync time for given sourceCalendar
      PropertiesService.getUserProperties().setProperty(calendarConfig.sourceCalendarName, new Date().getTime());
  }

  if ((addedEvents.length + modifiedEvents.length + removedEvents.length) > 0 && emailSummary){
    sendSummary();
  }
  Logger.log("Sync finished!");
  PropertiesService.getUserProperties().setProperty('LastRun', 0);
}
