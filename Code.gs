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

// sourceCalendars contains a named list of URLs for example:
//var sourceCalendars={
//  CalTarget1 :            // The name of the Google Calendar you want to add events to
//    ['URL1'],                // One or more ics/ical url that you want to get events from
//  CalTarget2 :            // The name of the Google Calendar you want to add events to
//    ['URL2',                 // One or more ics/ical url that you want to get events from
//    'URL3']                  // One or more ics/ical url that you want to get events from
//}
var sourceCalendars={
  TargetCalendar :
    ['']
}

var ignoreTitles = [/Stay at/];                  // List of regex patterns that will ignore events using title

var timeout = 275;                     // How long should the script run for
var howFrequent = 4*60;                // What interval (minutes) to run this script on to check for new events
var onlyFuture = false;                // If you turn this to "false", all events (past as well as future will be copied over)
var addEventsToCalendar = true;        // If you turn this to "false", you can check the log (View > Logs) to make sure your events are being read correctly before turning this on
var modifyExistingEvents = true;       // If you turn this to "false", any event in the feed that was modified after being added to the calendar will not update
var removeEventsFromCalendar = true;   // If you turn this to "true", any event in the calendar not found in the feed will be removed.
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
  Uninstall();

  //Custom error for restriction here: https://developers.google.com/apps-script/reference/script/clock-trigger-builder#everyMinutes(Integer)
  var validFrequencies = [1, 5, 10, 15, 30];
  if (Math.floor(howFrequent/60)>=1){
    ScriptApp.newTrigger("Start").timeBased().everyHours(Math.floor(howFrequent/60)).create();
  } else if(validFrequencies.indexOf(howFrequent) != -1) {
    ScriptApp.newTrigger("Start").timeBased().everyMinutes(howFrequent).create();
  } else {
    throw "[ERROR] Invalid value for \"howFrequent\". Must be either 1, 5, 10, 15, 30, or over 60...";
  }
}

function Uninstall(){
  DeleteAllTriggers("Start");
  DeleteAllTriggers("Loop");
}

function Start(){
  CheckForUpdate();
  DeleteAllTriggers("Loop");
  Reset();
  ScriptApp.newTrigger("main").timeBased().everyMinutes(5).create();
}

function Reset(){
  saveProp('iCAL',0);
  saveProp('iURL',0);
  saveProp('iEVT',0);
  saveProp('StartUpdate',GetTime(new Date()));
}


function Loop(){

  try{
    //------------------------ Error checking ------------------------
    if (emailWhenAdded && email == "")
      throw "[ERROR] \"emailWhenAdded\" is set to true, but no email is defined";
    //----------------------------------------------------------------

    //Start time counter
    var TIC = new Date();

    //Retrieve initial counters
    var iCAL = Number(loadProp('iCAL'));
    var iURL = Number(loadProp('iURL'));
    var iEVT = Number(loadProp('iEVT'));
    LogIt("Current indexes: iCAL "+iCAL+"  iURL "+iURL+"  iEVT "+iEVT);

    //Retreve milliseconds from 1970 of StartUpdate
    var StartUpdate = loadProp('StartUpdate');

    //Retrieve all target calendar names
    var targetCalendarNames = Object.keys(sourceCalendars);
    for (; iCAL < targetCalendarNames.length; iCAL++){
      var targetCalendarName = targetCalendarNames[iCAL];
      var sourceCalendarURLs = sourceCalendars[targetCalendarName];

      //Get target calendar information
      var targetCalendar = CalendarApp.getCalendarsByName(targetCalendarName)[0];
      if (targetCalendar == null){
        LogIt("Creating Calendar: " + targetCalendarName);
        targetCalendar = CalendarApp.createCalendar(targetCalendarName);
        targetCalendar.setTimeZone(CalendarApp.getTimeZone());
        targetCalendar.setSelected(false); //Sets the calendar as "hidden" in the Google Calendar UI
      }

      if(addEventsToCalendar || modifyExistingEvents || removeEventsFromCalendar ){
        LogIt("Loading Calendar: " + targetCalendarName);
        var calendarEvents = targetCalendar.getEvents(new Date(2000,01,01), new Date( 2100,01,01 ));
        LogIt("         -> " + calendarEvents.length);
        var calendarFids = [];
        for (var i = 0; i < calendarEvents.length; i++){
          calendarFids[i] = calendarEvents[i].getTag("FID");
        }
      }

      //------------------------ Parse events --------------------------
      for ( ; iURL < sourceCalendarURLs.length; iURL++) {
        var sourceCalendarURL = sourceCalendarURLs[iURL];
        LogIt("URL: " + sourceCalendarURL);

        //Get URL items
        var response = UrlFetchApp.fetch(sourceCalendarURL).getContentText();

        //------------------------ Error checking ------------------------
        if(response.includes("That calendar does not exist")) throw "[ERROR] Incorrect ics/ical URL";

        //Use ICAL.js to parse the data
        var jcalData = ICAL.parse(response);
        var component = new ICAL.Component(jcalData);
        for each (var VTZ in component.getAllSubcomponents("vtimezone")){
          if (VTZ != null){
            ICAL.TimezoneService.register(VTZ);
          }
        }

        //Extract all vevents
        var events_ = component.getAllSubcomponents("vevent");
        LogIt("         -> " + events_.length);

        //Loop through all vevents
        for ( ; iEVT < events_.length; iEVT++) {

          //Convert the vevent into custom event object
          var event = ConvertToCustomEvent(events_[iEVT]);

          //Check for only future events
          if ( onlyFuture && GetTime(event.startTime)<StartUpdate) {
            continue;
          }

          //Check if title contains one of the regex
          var exiting = false;
          for each (var ignoreTitle in ignoreTitles){
            if (ignoreTitle.test(event.title)){
              exiting = true;
              break;
            }
          }
          if (exiting) continue;

          //Look for existing event
          var calendarIndex = calendarFids.indexOf(event.id);

          if (calendarIndex == -1){
            //------------------------ Add events to calendar ----------------
            if (addEventsToCalendar) {
              //          Logger.log("Checking calendar #" + i + " event #" + j + " for creation");
              LogIt("    Creating: " + iEVT + "  -  " + event.title + "  @  " + event.startTime);

              var resultEvent = CreateEvent(targetCalendar,event);

              if(removeEventsFromCalendar){
                //Update status tag
                resultEvent.setTag("Status",StartUpdate);
              }

              if (emailWhenAdded) {
                GmailApp.sendEmail(email, "New Event Added", "New event added to calendar \"" + targetCalendarName + "\" at " + event.startTime);
              }
            }
            //----------------------------------------------------------------
          }
          else {

            //-------------- Modify events from calendar -----------
            if(modifyExistingEvents){
              var e = calendarEvents[calendarIndex];
              LogIt("    Modifying: " + iEVT + "  -  " + event.title + "  @  " + event.startTime);
              // **** BUG when type changes from allday to normal ****
              if(e.isAllDayEvent() != event.isAllDay){
                LogIt("        Changing event type");
                e.deleteEvent();
                calendarEvents[calendarIndex] = CreateEvent(targetCalendar,event);
                e = calendarEvents[calendarIndex];
              }
              if(e.getStartTime().getTime() != event.startTime.getTime() ||
                e.getEndTime().getTime() != event.endTime.getTime()) {
                  e.setTime(event.startTime, event.endTime)
                }
              if(e.getTitle() != event.title) {
                e.setTitle(event.title);
              }
              if(e.getLocation() != event.location){
                e.setLocation(event.location)
              }
              if(e.getDescription() != event.description) {
                e.setDescription(event.description)
              }
              e.removeAllReminders();
              if (addAlerts){
                for each (var reminder in event.reminderTimes) {
                  e.addPopupReminder(reminder / 60);
                }
              }
            }

            if(removeEventsFromCalendar){
              //Update status tag
              e.setTag("Status",StartUpdate);
            }

          }//------ End of if (calendarIndex == -1) --------


          //------------------------ Check results -------------------------
          Logger.log("Title: " + event.title);
          Logger.log("Id: " + event.id);
          Logger.log("Description: " + event.description);
          Logger.log("Start: " + event.startTime);
          Logger.log("End: " + event.endTime);
          for each (var reminder in event.reminderTimes) {
            Logger.log("Reminder: " + reminder + " seconds before");
          }
          Logger.log("");
          //----------------------------------------------------------------


          //------- Checks current time and stops it if overtime --------
          //Stops time counter
          var TOC = new Date();
          if ( timeout * 1000 < Number(TOC)-Number(TIC) ) {
            LogIt("    Overtime : reaching time limit");
            //Update the CAL index
            saveProp('iCAL',iCAL);
            //Update the URL index
            saveProp('iURL',iURL);
            //Update the EVT index
            saveProp('iEVT',iEVT+1);
            return;
          }
          //----------------------------------------------------------------

        } //------------ END OF FOR LOOP EVTS -----------------------

        if (iURL<sourceCalendarURLs.length-1) { //Reset iEVT counter for next URL
          //Reinitialise the EVT index
          iEVT = 0;
        }
        else {
          //Update the URL index
          saveProp('iURL',++iURL);
        }

      } //------------ END OF FOR LOOP URLS -----------------------


      //-------------- Remove events from calendar -----------
      if(removeEventsFromCalendar){
        for ( var iDEL = 0 ; iDEL < calendarEvents.length; iDEL++){
          var e = calendarEvents[iDEL];
          if (GetTime(e.getLastUpdated())<StartUpdate){
            LogIt("    Deleting " + iDEL + "  -  " + e.getTitle());
            calendarEvents[iDEL].deleteEvent();
          }

          //------- Checks current time and stops it if overtime --------
          //Stops time counter
          var TOC = new Date();
          if ( timeout * 1000 < Number(TOC)-Number(TIC) ) {
            Logger.log("    Overtime : reaching time limit");
            //Update the CAL index
            saveProp('iCAL',iCAL);
            return;
          }
        } //------------ END OF FOR LOOP DEL EVENTS -----------------------
      }
    } //------------ END OF FOR LOOP TARGET CALENDAR -----------------------

    LogIt(" FINISHED - CLEANING UP ");
    DeleteAllTriggers("main");
    Reset();

  }
  catch(err){
    LogIt(" ERROR captured - " + err.message);
    //Update the CAL index
    saveProp('iCAL',iCAL);
    //Update the URL index
    saveProp('iURL',iURL);
    //Update the EVT index
    saveProp('iEVT',iEVT);
    DeleteAllTriggers("main");
  }
} //------------ END OF MAIN -----------------------

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
