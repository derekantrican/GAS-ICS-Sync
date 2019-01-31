var defaultSettings = {"targetCalendarName": "",
                       "sourceCalendarURL": "",
                       "howFrequent": 15,
                       "addEventsToCalendar": true,
                       "modifyExistingEvents": true,
                       "removeEventsFromCalendar": true,
                       "addAlerts": true,
                       "addOrganizerToTitle": false,
                       "descriptionAsTitles": false,
                       "defaultDuration": 60,
                       "emailWhenAdded": false,
                       "email": ""
                      };

function loadSetting(settingName){
  return PropertiesService.getScriptProperties().getProperty(settingName);
}

function saveSetting(settingName, settingValue){
  PropertiesService.getScriptProperties().setProperty(settingName, settingValue);
}

function loadAllSettings(){
  var settingsDict = {};

  for each (var setting in getAllSettingNames()){
    var value = PropertiesService.getScriptProperties().getProperty(setting);
    settingsDict[setting] = value;
  }

  return settingsDict;
}

function getAllSettingNames(){
  var settingsKeys = PropertiesService.getScriptProperties().getKeys();
  
  if (settingsKeys == null || settingsKeys.length == 0){
    for (var setting in defaultSettings)
      saveSetting(setting, defaultSettings[setting]);
    
    return PropertiesService.getScriptProperties().getKeys();
  }
  else
    return settingsKeys;
}