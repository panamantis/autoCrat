function autoCrat_removeTimeoutTrigger() {
  var triggers = ScriptApp.getScriptTriggers();
  if (triggers.length>0) {
  for (var i=0; i<triggers.length; i++) {
    var handlerFunction = triggers[i].getHandlerFunction();
    if (handlerFunction=='autoCrat_runMerge') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
}

function autoCrat_preemptTimeout() {
  var date = new Date();
  var newDate = new Date(date);
  newDate.setSeconds(date.getSeconds() + 60);
  ScriptApp.newTrigger('autoCrat_runMerge').timeBased().at(newDate).create();
  Browser.msgBox('Merge paused, restarting in one minute to avoid service timeout.');
}

function urlencode(inNum) {
  // Function to convert non URL compatible characters to URL-encoded characters
  var outNum = 0;     // this will hold the answer
  outNum = escape(inNum); //this will URL Encode the value of inNum replacing whitespaces with %20, etc.
  return outNum;  // return the answer to the cell which has the formula
}


function autoCrat_getInstitutionalTrackerObject() {
  var institutionalTrackingString = UserProperties.getProperty('institutionalTrackingString');
  if ((institutionalTrackingString)&&(institutionalTrackingString != "not participating")) {
    var institutionTrackingObject = Utilities.jsonParse(institutionalTrackingString);
    return institutionTrackingObject;
  }
  if (!(institutionalTrackingString)||(institutionalTrackingString=='')) {
    autoCrat_institutionalTrackingUi();
    return;
  }
}


function autoCrat_institutionalTrackingUi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var institutionalTrackingString = UserProperties.getProperty('institutionalTrackingString');
  var eduSetting = UserProperties.getProperty('eduSetting');
  if (!(institutionalTrackingString)) {
    UserProperties.setProperty('institutionalTrackingString', 'not participating');
  }
  var app = UiApp.createApplication().setTitle('Hello there! Help us track the usage of this script').setHeight(400);
  if ((!(institutionalTrackingString))||(!(eduSetting))) {
    var helptext = app.createLabel("You are most likely seeing this prompt because this is the first time you are using a Google Apps script created by New Visions for Public Schools, 501(c)3. If you are using scripts as part of a school or grant-funded program like New Visions' CloudLab, you may wish to track usage rates with Google Analytics. Entering tracking information here will save it to your user credentials and enable tracking for any other New Visions scripts that use this feature. No personal info will ever be collected.").setStyleAttribute('marginBottom', '10px');
  } else {
  var helptext = app.createLabel("If you are using scripts as part of a school or grant-funded program like New Visions' CloudLab, you may wish to track usage rates with Google Analytics. Entering or modifying tracking information here will save it to your user credentials and enable tracking for any other scripts produced by New Visions for Public Schools, 501(c)3, that use this feature. No personal info will ever be collected.").setStyleAttribute('marginBottom', '10px');
  }
  var panel = app.createVerticalPanel();
  var gridPanel = app.createVerticalPanel().setId("gridPanel").setVisible(false);
  var grid = app.createGrid(4,2).setId('trackingGrid').setStyleAttribute('background', 'whiteSmoke').setStyleAttribute('marginTop', '10px');
  var checkHandler = app.createServerHandler('autoCrat_refreshTrackingGrid').addCallbackElement(panel);
  var checkBox = app.createCheckBox('Participate in institutional usage tracking.  (Only choose this option if you know your institution\'s Google Analytics tracker Id.)').setName('trackerSetting').addValueChangeHandler(checkHandler);  
  var checkBox2 = app.createCheckBox('Let New Visions for Public Schools, 501(c)3 know you\'re an educational user.').setName('eduSetting');  
  if ((institutionalTrackingString == "not participating")||(institutionalTrackingString=='')) {
    checkBox.setValue(false);
  } 
  if (eduSetting=="true") {
    checkBox2.setValue(true);
  }
  var institutionNameFields = [];
  var trackerIdFields = [];
  var institutionNameLabel = app.createLabel('Institution Name');
  var trackerIdLabel = app.createLabel('Google Analytics Tracker Id (UA-########-#)');
  grid.setWidget(0, 0, institutionNameLabel);
  grid.setWidget(0, 1, trackerIdLabel);
  if ((institutionalTrackingString)&&((institutionalTrackingString!='not participating')||(institutionalTrackingString==''))) {
    checkBox.setValue(true);
    gridPanel.setVisible(true);
    var institutionalTrackingObject = Utilities.jsonParse(institutionalTrackingString);
  } else {
    var institutionalTrackingObject = new Object();
  }
  for (var i=1; i<4; i++) {
    institutionNameFields[i] = app.createTextBox().setName('institution-'+i);
    trackerIdFields[i] = app.createTextBox().setName('trackerId-'+i);
    if (institutionalTrackingObject) {
      if (institutionalTrackingObject['institution-'+i]) {
        institutionNameFields[i].setValue(institutionalTrackingObject['institution-'+i]['name']);
        if (institutionalTrackingObject['institution-'+i]['trackerId']) {
          trackerIdFields[i].setValue(institutionalTrackingObject['institution-'+i]['trackerId']);
        }
      }
    }
    grid.setWidget(i, 0, institutionNameFields[i]);
    grid.setWidget(i, 1, trackerIdFields[i]);
  } 
  var help = app.createLabel('Enter up to three institutions, with Google Analytics tracker Id\'s.').setStyleAttribute('marginBottom','5px').setStyleAttribute('marginTop','10px');
  gridPanel.add(help);
  gridPanel.add(grid); 
  panel.add(helptext);
  panel.add(checkBox2);
  panel.add(checkBox);
  panel.add(gridPanel);
  var button = app.createButton("Save settings");
  var saveHandler = app.createServerHandler('autoCrat_saveInstitutionalTrackingInfo').addCallbackElement(panel);
  button.addClickHandler(saveHandler);
  panel.add(button);
  app.add(panel);
  ss.show(app);
  return app;
}

function autoCrat_refreshTrackingGrid(e) {
  var app = UiApp.getActiveApplication();
  var gridPanel = app.getElementById("gridPanel");
  var grid = app.getElementById("trackingGrid");
  var setting = e.parameter.trackerSetting;
  if (setting=="true") {
    gridPanel.setVisible(true);
  } else {
    gridPanel.setVisible(false);
  }
  return app;
}


function autoCrat_saveInstitutionalTrackingInfo(e) {
  var app = UiApp.getActiveApplication();
  var eduSetting = e.parameter.eduSetting;
  var oldEduSetting = UserProperties.getProperty('eduSetting')
  if (eduSetting == "true") {
    UserProperties.setProperty('eduSetting', 'true');
  }
  if ((oldEduSetting)&&(eduSetting=="false")) {
    UserProperties.setProperty('eduSetting', 'false');
  }
  var trackerSetting = e.parameter.trackerSetting;
  if (trackerSetting == "false") {
    UserProperties.setProperty('institutionalTrackingString', 'not participating');
    app.close();
    return app;
  } else {
  var institutionalTrackingObject = new Object;
  for (var i=1; i<4; i++) {
    var checkVal = e.parameter['institution-'+i];
    if (checkVal!='') {
      institutionalTrackingObject['institution-'+i] = new Object();
      institutionalTrackingObject['institution-'+i]['name'] = e.parameter['institution-'+i];
      institutionalTrackingObject['institution-'+i]['trackerId'] = e.parameter['trackerId-'+i];
      if (!(e.parameter['trackerId-'+i])) {
        Browser.msgBox("You entered an institution without a Google Analytics Tracker Id");
        autoCrat_institutionalTrackingUi()
      }
    }
  }
  var institutionalTrackingString = Utilities.jsonStringify(institutionalTrackingObject);
  UserProperties.setProperty('institutionalTrackingString', institutionalTrackingString);
  autoCrat_initialize;
  app.close();
  return app;
}
}


// Some of this code was borrowed and modified from the Flubaroo Script author Dave Abouav
// It anonymously tracks script usage to Google Analytics, allowing our non-profit organization to report the impact of this work to funders
// For original source see http://www.edcode.org

function autoCrat_createInstitutionalTrackingUrls(institutionTrackingObject, encoded_page_name, encoded_script_name) {
  for (var key in institutionTrackingObject) {
   var utmcc = autoCrat_createGACookie();
  if (utmcc == null)
    {
      return null;
    }
  var encoded_page_name = encoded_script_name+"/"+encoded_page_name;
  var trackingId = institutionTrackingObject[key].trackerId;
  var ga_url1 = "http://www.google-analytics.com/__utm.gif?utmwv=5.2.2&utmhn=www.autocrat-analytics.com&utmcs=-&utmul=en-us&utmje=1&utmdt&utmr=0=";
  var ga_url2 = "&utmac="+trackingId+"&utmcc=" + utmcc + "&utmu=DI~";
  var ga_url_full = ga_url1 + encoded_page_name + "&utmp=" + encoded_page_name + ga_url2;
  
  if (ga_url_full)
    {
      var response = UrlFetchApp.fetch(ga_url_full);
    }
  }
}



function autoCrat_createGATrackingUrl(encoded_page_name)
{
  var utmcc = autoCrat_createGACookie();
  var eduSetting = UserProperties.getProperty('eduSetting');
   if (eduSetting=="true") {
    encoded_page_name = "edu/" + encoded_page_name;
  }
  if (utmcc == null)
    {
      return null;
    }
 
  var ga_url1 = "http://www.google-analytics.com/__utm.gif?utmwv=5.2.2&utmhn=www.autocrat-analytics.com&utmcs=-&utmul=en-us&utmje=1&utmdt&utmr=0=";
  var ga_url2 = "&utmac=UA-30983014-1&utmcc=" + utmcc + "&utmu=DI~";
  var ga_url_full = ga_url1 + encoded_page_name + "&utmp=" + encoded_page_name + ga_url2;
  
  return ga_url_full;
}


function autoCrat_createSystemTrackingUrls(institutionTrackingObject, encoded_system_name, encoded_execution_name) {
  for (var key in institutionTrackingObject) {
  var utmcc = autoCrat_createGACookie();
  if (utmcc == null)
    {
      return null;
    }
  var trackingId = institutionTrackingObject[key].trackerId;
  var encoded_page_name = encoded_system_name+"/"+encoded_execution_name;
  var ga_url1 = "http://www.google-analytics.com/__utm.gif?utmwv=5.2.2&utmhn=www.cloudlab-systems-analytics.com&utmcs=-&utmul=en-us&utmje=1&utmdt&utmr=0=";
  var ga_url2 = "&utmac="+trackingId+"&utmcc=" + utmcc + "&utmu=DI~";
  var ga_url_full1 = ga_url1 + encoded_page_name + "&utmp=" + encoded_page_name + ga_url2;
  if (ga_url_full1)
    {
      var response = UrlFetchApp.fetch(ga_url_full1);
    } 
  }
  var encoded_page_name = encoded_system_name+"/"+encoded_execution_name+"/"+trackingId;
  var ga_url1 = "http://www.google-analytics.com/__utm.gif?utmwv=5.2.2&utmhn=www.cloudlab-systems-analytics.com&utmcs=-&utmul=en-us&utmje=1&utmdt&utmr=0=";
  var ga_url2 = "&utmac=UA-34521561-1&utmcc=" + utmcc + "&utmu=DI~";
  var ga_url_full2 = ga_url1 + encoded_page_name + "&utmp=" + encoded_page_name + ga_url2;  
  if (ga_url_full2)
    {
      var response = UrlFetchApp.fetch(ga_url_full2);
    }
}


function autoCrat_createGACookie()
{
  var a = "";
  var b = "100000000";
  var c = "200000000";
  var d = "";

  var dt = new Date();
  var ms = dt.getTime();
  var ms_str = ms.toString();
 
  var autocrat_uid = UserProperties.getProperty("autocrat_uid");
  if ((autocrat_uid == null) || (autocrat_uid == ""))
    {
      // shouldn't happen unless user explicitly removed autocrat_uid from properties.
      return null;
    }
  
  a = autocrat_uid.substring(0,9);
  d = autocrat_uid.substring(9);
  
  utmcc = "__utma%3D451096098." + a + "." + b + "." + c + "." + d 
          + ".1%3B%2B__utmz%3D451096098." + d + ".1.1.utmcsr%3D(direct)%7Cutmccn%3D(direct)%7Cutmcmd%3D(none)%3B";
 
  return utmcc;
}

function autoCrat_logDocCreation()
{
  var ga_url = autoCrat_createGATrackingUrl("Merged%20Doc%20Created");
  if (ga_url)
    {
      var response = UrlFetchApp.fetch(ga_url);
    }
  var institutionalTrackingObject = autoCrat_getInstitutionalTrackerObject();
  if (institutionalTrackingObject) {
    autoCrat_createInstitutionalTrackingUrls(institutionalTrackingObject,"Merged%20Doc%20Created", "autoCrat");
    var systemName = ScriptProperties.getProperty('systemName');
    if (systemName) {
      var encoded_system_name = urlencode(systemName);
      autoCrat_createSystemTrackingUrls(institutionalTrackingObject, encoded_system_name, "Merged%20Doc%20Created")
    }
  }
}


function autoCrat_logFirstInstall()
{
  var ga_url = autoCrat_createGATrackingUrl("First%20Install");
  if (ga_url)
    {
      var response = UrlFetchApp.fetch(ga_url);
    }
    var institutionalTrackingObject = autoCrat_getInstitutionalTrackerObject();
  if (institutionalTrackingObject) {
    autoCrat_createInstitutionalTrackingUrls(institutionalTrackingObject,"First%20Install", "autoCrat");
    var systemName = ScriptProperties.getProperty('systemName');
    if (systemName) {
      var encoded_system_name = urlencode(systemName);
      autoCrat_createSystemTrackingUrls(institutionalTrackingObject, encoded_system_name, "First%20Install")
    }
  }
}



function setAutocratUid()
{ 
  var autocrat_uid = UserProperties.getProperty("autocrat_uid");
  if (autocrat_uid == null || autocrat_uid == "")
    {
      // user has never installed autoCrat before (in any spreadsheet)
      var dt = new Date();
      var ms = dt.getTime();
      var ms_str = ms.toString();
 
      UserProperties.setProperty("autocrat_uid", ms_str);
      autoCrat_logFirstInstall();
    }
}


// utility function to clear all merge status messages and doc links associated with rows in the source sheet
// some may find it useful to set this on a regular time trigger (once a day, for example) within a given workflow
function autoCrat_clearAllFlags() {
  var sheetName = ScriptProperties.getProperty('sheetName');
  var ss = SpreadsheetApp.getActive();
  if ((sheetName)&&(sheetName!='')) {
    var sheet = ss.getSheetByName(sheetName);
    var headers = autoCrat_fetchSheetHeaders(sheetName);
    var lastCol = sheet.getLastColumn();
    var statusCol = headers.indexOf("Document Merge Status");
    var linkCol = headers.indexOf("Link to merged Doc");
    var urlCol = headers.indexOf("Merged Doc URL");
    var lastRow = sheet.getLastRow();
    if ((statusCol!=-1)&&(lastRow>1)) {
      var range = sheet.getRange(3, statusCol+1, lastRow-1, 1).clear();
    }
    if ((linkCol != -1)&&(lastRow>1)) {
      var range = sheet.getRange(3, linkCol+1, lastRow-1, 1).clear();
    }
    if ((urlCol != -1)&&(lastRow>1)) {
      var range = sheet.getRange(3, urlCol+1, lastRow-1, 1).clear();
    }
  }
}
