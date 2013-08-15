var scriptTitle = "autoCrat Script V4.3.1 (4/9/13)";
var scriptName = "autoCrat"
var analyticsId = 'UA-30983014-1'
// Written by Andrew Stillman for New Visions for Public Schools
// Published under GNU General Public License, version 3 (GPL-3.0)
// See restrictions at http://www.opensource.org/licenses/gpl-3.0.html
// Support and contact at http://www.youpd.org/autocrat

//Want to run autoCrat on a time-based trigger?  
//Set time triggers on the autoCrat_onFormSubmit function.


var AUTOCRATIMAGEURL = 'https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/autoCrat_icon.gif?attachauth=ANoY7crV8whTTm4tyWEhFrNIM8Yt6RdYthydKlFA4gpIovpihpZdsviIZ0_D42FJXHSxpZnRFyJdSj7iCS5KMTjv9VYHGfctNojT3Tckh2zJHB5AlEwZIqj2uYdKPz8Zl6JtsUTWIYzLCoCxM-NvPWlji1fL9LGjIx1e-AKmz6qnxq2K_rC9zCiENHHaap9Lyq9W4umEoeYWtqWykApce9wtLXhFYEJ7uLN65vLGDKl5Ao5OHwyTG3COIIij-qsufuMjFr2WtHAK&attredirects=0';

function onInstall() {
  onOpen();
}


//onOpen is part of the Google Apps Script library.  It runs whenever the spreadsheet is opened.  
//Adds script menu to the spreadsheet
//Sometimes this seems not to run by itself when the script is installed
//and needs to be run manually the first time to prompt script authorization.

function onOpen() {
  var ss = SpreadsheetApp.getActive();
  var menuEntries = [];
  menuEntries.push({name: "What is autoCrat?", functionName: "autoCrat_whatIs"});
  menuEntries.push({name: "Run initial configuration", functionName: "autoCrat_preconfig"});
  ss.addMenu("autoCrat", menuEntries);
  autoCrat_initialize();
}


//This function is responsible for configuring the script upon first install.  
//Subsequently, adds the menu of dropdown items based on previous actions by the user, as stored in script properties.
function autoCrat_initialize() {
  var ss = SpreadsheetApp.getActive();
  var menuEntries = [];
  menuEntries.push({name: "What is autoCrat?", functionName: "autoCrat_whatIs"});
  var preconfigStatus = ScriptProperties.getProperty('preconfigStatus');
  if (preconfigStatus) {
    menuEntries.push({name: "Step 1: Choose Template Doc", functionName: "autoCrat_defineTemplate"});
  } else {
    menuEntries.push({name: "Run initial configuration", functionName: "autoCrat_preconfig"});
  }
  var fileId = ScriptProperties.getProperty('fileId');
  if ((fileId)&&(!fileId=="")) {
    menuEntries.push({name: "Step 2: Select Source Data", functionName: "autoCrat_defineSettings"});
    var sheetName = ScriptProperties.getProperty('sheetName');
    var mappingString = ScriptProperties.getProperty('mappingString');
    var fileSetting = ScriptProperties.getProperty('fileSetting');
    var emailSetting = ScriptProperties.getProperty('emailSetting');
    if ((sheetName)&&(!sheetName=="")) {
      menuEntries.push({name: "Step 3: Set Merge Conditions", functionName: "autoCrat_setMergeConditions"});
    }
    menuEntries.push({name: "Step 4: Set Field Mappings", functionName: "autoCrat_mapFields"});
  }
  if ((mappingString)&&(!mappingString=="")) { 
    menuEntries.push({name: "Step 5: Set Merge Type", functionName: "autoCrat_runMergeConsole"});
  }
  if (((fileSetting)||(emailSetting))&&(mappingString)&&(!mappingString=="")) { 
    menuEntries.push({name: "Step 6: Preview/Run Merge", functionName: "autoCrat_runMergePrompt"});
  }
  if (((fileSetting)||(emailSetting))&&(mappingString)&&(!mappingString=="")) { 
    menuEntries.push({name: "Advanced options", functionName: "autoCrat_advanced"});
  }
  ss.addMenu("autoCrat", menuEntries);
  
  //ensure readme sheets exist.  If not, install it and set as active sheet.
  var sheets = ss.getSheets();
  var readMeSet = false;
  for (var i=0; i<sheets.length; i++) {
    if (sheets[i].getName()=="autoCrat Read Me") {
      readMeSet = true;
      break;
    }
  }
  if (readMeSet==false) {
    ss.insertSheet("autoCrat Read Me");
    autoCrat_setReadMeText();
    var sheet = ss.getSheetByName("autoCrat Read Me");
    ss.setActiveSheet(sheet);
  }
  if ((preconfigStatus)&&(!(fileId))) {
    autoCrat_defineTemplate();
  }
}



function autoCrat_advanced() {
  var ss = SpreadsheetApp.getActive();
  var app = UiApp.createApplication().setTitle("Advanced options").setHeight(130).setWidth(290);
  var quitHandler = app.createServerHandler('autoCrat_quitUi');
  var handler2 = app.createServerHandler('autoCrat_extractorWindow');
  var button2 = app.createButton('Package this system for others to copy').addClickHandler(quitHandler).addClickHandler(handler2);
  var handler3 = app.createServerHandler('autoCrat_institutionalTrackingUi');
  var button3 = app.createButton('Manage your usage tracker settings').addClickHandler(quitHandler).addClickHandler(handler3);
  var panel = app.createVerticalPanel();
  panel.add(button2);
  panel.add(button3);
  app.add(panel);
  ss.show(app);
  return app;
}

function autoCrat_quitUi(e) {
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}





function autoCrat_setFormTrigger() {
  var ssKey = SpreadsheetApp.getActiveSpreadsheet().getId();
  ScriptApp.newTrigger('autoCrat_onFormSubmit').forSpreadsheet(ssKey).onFormSubmit().create();
}

//onSubmit is part of the Google Apps Script library.  It runs whenever a Google Form is submitted.  
//Checks to see if a merge method has been set and the "trigger on form submit" option has been saved in the
//and fires off the merge.  
//Note that the merge will only execute for a record if there is no value in the "Merge Status."
//Code for handling this condition is in the runMerge function

function autoCrat_onFormSubmit() {
  var lock = LockService.getPublicLock();
  lock.waitLock(120000);
  var ssKey = ScriptProperties.getProperty('ssKey');
  var ss = SpreadsheetApp.openById(ssKey);
  var formTrigger = ScriptProperties.getProperty('formTrigger');
  var fileSetting = ScriptProperties.getProperty('fileSetting');
  var emailSetting = ScriptProperties.getProperty('emailSetting');
  if ((formTrigger == "true")&&(fileSetting == "true") || (formTrigger == "true")&&(emailSetting == "true")) {
  autoCrat_runMerge();
  }
  lock.releaseLock();
}


//Function to handle the creation of the "Test/Run Merge" GUI panel.

function autoCrat_runMergeConsole() {
  var ss = SpreadsheetApp.getActive();
  var app = UiApp.createApplication();
  app.setTitle("Step 5: Set Merge Type");
  app.setHeight("500");
 
var fileId = ScriptProperties.getProperty('fileId');
  if (!fileId) { 
       Browser.msgBox("You must select a template file before you can run a merge.");
       autoCrat_defineTemplate();
       return;
       }
  var mappingString = ScriptProperties.getProperty('mappingString');
  if (!mappingString) {
       Browser.msgBox("You must map document fields before you can run a merge.");
       autoCrat_mapFields();
       return;
       }

  var sheetName = ScriptProperties.getProperty('sheetName');
  if (!sheetName) {
       Browser.msgBox("You must select a source data sheet before you can run a merge.");
       autoCrat_defineSettings();
       return;
       }

 //create spinner graphic to show upon button click awaiting merge completion
  var refreshPanel = app.createFlowPanel();
  refreshPanel.setId('refreshPanel');
  refreshPanel.setStyleAttribute("width", "100%");
  refreshPanel.setStyleAttribute("height", "500px");
  refreshPanel.setVisible(false);

//Adds the graphic for the waiting period before merge completion. Set invisible until client handler
//is called by button click 
  var spinner = app.createImage(AUTOCRATIMAGEURL);
  spinner.setVisible(false);
  spinner.setStyleAttribute("position", "absolute");
  spinner.setStyleAttribute("top", "220px");
  spinner.setStyleAttribute("left", "220px");
  spinner.setId("dialogspinner");

// Build the panel
  var panel = app.createVerticalPanel().setId("fieldMappingPanel");
  var varScrollPanel = app.createScrollPanel().setHeight("150px").setStyleAttribute('backgroundColor', 'whiteSmoke');
  var scrollPanel = app.createScrollPanel().setHeight("350px");
  panel.setStyleAttribute("width", "100%");
  var folderLabel = app.createLabel().setId("folderLabel").setWidth("100%");
  folderLabel.setText("Select the destination folder(s) for your merged documents. Primary folder must be in the same folder as this spreadsheet. Assign secondary folder(s) using valid folder key(s). Separate multiple with commas. Can be a variable that provides folder key.").setStyleAttribute("clear","right");
  var folderListBox = app.createListBox().setName("destinationFolderId").setId("destFolderListBox").setWidth("210px");
  folderListBox.addItem('Select primary destination folder').setStyleAttribute("clear","right");
  var secondaryFolder = app.createTextBox().setName('secondaryFolderId').setWidth("260px").setValue("Optional: Additional folder key(s) here").setStyleAttribute('color', 'grey');
  var secondaryClickHandler = app.createClientHandler().forTargets(secondaryFolder).setStyleAttribute('color', 'black');
  secondaryFolder.addMouseDownHandler(secondaryClickHandler);
  
// Build listbox for folder destination options.  Limit to first 20 folders to avoid
// Google server errors.
  var parent = DocsList.getFileById(ss.getId()).getParents()[0];
  if (!(parent)) {
   parent = DocsList.getRootFolder();
  }
  var folders = parent.getFolders();
  if (folders.length>0) {
  for (var i = 0; i<folders.length; i++) {
    var name = folders[i].getName();
    var id = folders[i].getId();
    folderListBox.addItem(name, id);
  }
  } else {
    var newFolderId = parent.createFolder("New Merged Document Folder");
    folderListBox.addItem("New Merged Document Folder", newFolderId);
    folderListBox.setSelectedIndex(1);
  }
  
  var fileToFolderCheckBox = app.createCheckBox().setId("fileToFolderCheckBox").setName("fileToFolderCheckValue");
  fileToFolderCheckBox.setText("Save merged files to Docs").setStyleAttribute("clear","right");;

  var fileToFolderCheckBoxFalse = app.createCheckBox().setId("fileToFolderCheckBoxFalse").setName("fileToFolderCheckValueFalse").setVisible(false);
  fileToFolderCheckBoxFalse.setText("Save merged files to Docs").setVisible(false).setStyleAttribute("clear","right"); 

   
  //Preset to previously used folder value if it exists
  var destinationFolderId = ScriptProperties.getProperty('destinationFolderId');
  if ((destinationFolderId)&&(destinationFolderId!='')) {
      //autoCrat_getFolderIndex is a custom built function that looks up where the saved folder is in the list
      var index = autoCrat_getFolderIndex(destinationFolderId)+1;
      folderListBox.setItemSelected(index, true);
      }
  var secondaryFolderId = ScriptProperties.getProperty('secondaryFolderId');
  if ((secondaryFolderId)&&(secondaryFolderId!='')) {
    secondaryFolder.setValue(secondaryFolderId).setStyleAttribute('color','black');
  }
  //build the rest of the panel field and checkboxes.  
  //Checks for pre-set values in Script Properties and pre-populates fields  
  var fileNameLabel = app.createLabel().setId("fileNameLabel");
  var fileNameStringBox = app.createTextBox().setId("fileNameString").setName('fileNameString');
  fileNameStringBox.setWidth("100%");
  var fileNameStringValue = ScriptProperties.getProperty('fileNameString');
  if (fileNameStringValue) {
    fileNameStringBox.setValue(fileNameStringValue);
  }
  var fileNameHelpLabel = app.createLabel().setId("fileNameHelpLabel");
  var sheetName = ScriptProperties.getProperty('sheetName');
  var sheetFieldNames = autoCrat_fetchSheetHeaders(sheetName);
  var normalizedSheetFieldNames = autoCrat_normalizeHeaders(sheetFieldNames);
  var fileNameHelpText = "This setting determines the file name for each merged document.";
  fileNameLabel.setText("File naming convention to use:");
  fileNameHelpLabel.setText(fileNameHelpText);
  fileNameHelpLabel.setStyleAttribute("color","grey");
  var fileTypeLabel = app.createLabel().setText("Select the file type you want to create");
  var fileTypeSelectBox = app.createListBox().setId("fileTypeSelectBox").setName("fileType");
  fileTypeSelectBox.addItem("Google Doc")
                   .addItem("PDF");
  var fileType = ScriptProperties.getProperty('fileType');
  if (fileType=='Google Doc') {
     fileTypeSelectBox.setSelectedIndex(0);
  }
  if (fileType=='PDF') {
     fileTypeSelectBox.setSelectedIndex(1);
  } 
  var linkCheckBox = app.createCheckBox('Save links to merged Docs in spreadsheet').setId('linkToDoc').setName('linkToDoc');
  var linkToDoc = ScriptProperties.getProperty('linkToDoc');
  if (linkToDoc=="true") {
    linkCheckBox.setValue(true);
  }
  var fileId = ScriptProperties.getProperty("fileId");
  var sheetName = ScriptProperties.getProperty("sheetName");
  var mappingString = ScriptProperties.getProperty("mappingString");
  
  //add server and client handlers and callbacks to button
  var saveRunSettingsHandler = app.createServerHandler('autoCrat_saveRunSettings').addCallbackElement(panel);
  var spinnerHandler = app.createClientHandler().forTargets(spinner).setVisible(true).forTargets(panel).setStyleAttribute('opacity', '0.5');
  var button = app.createButton().setId("runMergeButton").addClickHandler(saveRunSettingsHandler);
  button.addClickHandler(spinnerHandler);
  button.setText("Save Settings");

  /*check for pre-existing values and preset checkbox and visibilities accordingly
   checkboxes fire off client handlers to make sub-panels visible when checked
   Below is a bit of ridiculous javascript trickery to work around the dire limitations
   of checkbox handlers, which don't allow for checked-unchecked 
   state to be recognized by the handler...hint, hint Googlers.
   The solution is to create two checkboxes, one for the "checked" state and one for the "unchecked"
   and fire a different handler, alternately hiding one of the two checkboxes and using
   server handlers to reset their checked status.
   The user only ever sees one checkbox
  */
  
  var fileSetting = ScriptProperties.getProperty('fileSetting');
  var fileInfoPanel = app.createVerticalPanel().setId("fileInfoPanel");
   if (fileSetting=="true") {
     fileToFolderCheckBox.setVisible(true).setValue(true);
     fileToFolderCheckBoxFalse.setVisible(false).setValue(true);
     fileInfoPanel.setVisible(true);
  } else {
    fileToFolderCheckBox.setVisible(false).setValue(false);
    fileToFolderCheckBoxFalse.setVisible(true).setValue(false);
    fileInfoPanel.setVisible(false);
  } 


  var fileToEmailCheckBox = app.createCheckBox().setId("fileToEmailCheckBox").setName("fileToEmailCheckValue");
  fileToEmailCheckBox.setText("Send merged files via Email").setEnabled(false);

  var fileToEmailCheckBoxFalse = app.createCheckBox().setId("fileToEmailCheckBoxFalse").setName("fileToEmailCheckValueFalse").setVisible(false);
  fileToEmailCheckBoxFalse.setText("Email and/or share merged documents").setVisible(false); 

  var emailLabel = app.createLabel().setId("emailRecipientsLabel");
  emailLabel.setText("Recepient email addresses:");

  var emailStringBox = app.createTextBox().setId("emailStringBox").setName('emailString');
  emailStringBox.setWidth("100%");
  var emailStringValue = ScriptProperties.getProperty('emailString');
  if (emailStringValue) {
    emailStringBox.setValue(emailStringValue);
  }

  var emailHelpLabel = app.createLabel().setId("emailHelpLabel");
  var emailHelpText = "Emails must be separated by commas.";
  emailHelpLabel.setText(emailHelpText);
  emailHelpLabel.setStyleAttribute("color","grey");

  var emailSubjectLabel = app.createLabel().setText('Email subject:');
 
  var emailSubjectBox = app.createTextBox().setId("emailSubjectBox").setName("emailSubject");
  emailSubjectBox.setWidth("100%");
  var emailSubjectPreset =ScriptProperties.getProperty('emailSubject');
  if (emailSubjectPreset) {
     emailSubjectBox.setValue(emailSubjectPreset);
  } 
  
  var bodyPrefixHelpLabel = app.createLabel().setId("bodyPrefixLabel").setText('Short note to recipients:');
  var bodyPrefixTextArea = app.createTextArea().setId("bodyPrefixTextArea").setName("bodyPrefix");
  bodyPrefixTextArea.setHeight("75px").setWidth("100%");
  
  var bodyPrefix = ScriptProperties.getProperty('bodyPrefix');
  if (bodyPrefix != null) {
  bodyPrefixTextArea.setValue(bodyPrefix);
  }
  var emailAttachmentLabel = app.createLabel().setText("Attachment type:");
  var emailAttachmentListBox = app.createListBox().setId("emailAttachmentListBox").setName("emailAttachment");
  emailAttachmentListBox.addItem("PDF")
                        .addItem("Recipient-view-only Google Doc")
                        .addItem("Recipient-editable Google Doc");
  var emailInfoPanel = app.createVerticalPanel().setId("emailInfoPanel");
  
  var attachmentPreset = ScriptProperties.getProperty('emailAttachment');

  if(attachmentPreset) {
  switch (attachmentPreset) {
    case "PDF":
      emailAttachmentListBox.setSelectedIndex(0);
      break;
    case "Recipient-view-only Google Doc":
      emailAttachmentListBox.setSelectedIndex(1);
      break;
    case "Recipient-editable Google Doc":
      emailAttachmentListBox.setSelectedIndex(2);
      break;
    default:
      emailAttachmentListBox.setSelectedIndex(0);
  }
}

//check for pre-existing value and preset checkbox and visibilities accordingly
  var emailSetting = ScriptProperties.getProperty('emailSetting');
  var emailInfoPanel = app.createVerticalPanel().setId("emailInfoPanel");
   if (emailSetting=="true") {
     fileToEmailCheckBox.setVisible(true).setValue(true);
     fileToEmailCheckBoxFalse.setVisible(false).setValue(true);
     emailInfoPanel.setVisible(true);
  } else {
    fileToEmailCheckBox.setVisible(false).setValue(false);
    fileToEmailCheckBoxFalse.setVisible(true).setValue(false);
    emailInfoPanel.setVisible(false);
  }


// more crazy trickery for checkboxes
  
 var fileUnCheckHandler = app.createClientHandler().forTargets(fileToFolderCheckBox, fileInfoPanel).setVisible(false)
                                                    .forTargets(fileToFolderCheckBoxFalse).setVisible(true)
                                                    .forTargets(fileToEmailCheckBox).setEnabled(false)
                                                    .forTargets(fileToEmailCheckBoxFalse).setEnabled(false)
  var unSetCheck = app.createServerHandler('autoCrat_unsetFileCheck').addCallbackElement(fileToFolderCheckBox);
  var fileCheckHandler = app.createClientHandler().forTargets(fileToFolderCheckBox, fileInfoPanel).setVisible(true)
                                                  .forTargets(fileToFolderCheckBoxFalse).setVisible(false)
                                                  .forTargets(fileToEmailCheckBoxFalse).setEnabled(true)
                                                  .forTargets(fileToEmailCheckBox).setEnabled(true);
  var setCheck = app.createServerHandler('autoCrat_setFileCheck').addCallbackElement(fileToFolderCheckBox);
  
  fileToFolderCheckBox.addClickHandler(unSetCheck).addClickHandler(fileUnCheckHandler);
  fileToFolderCheckBoxFalse.addClickHandler(fileCheckHandler).addClickHandler(setCheck);
  
  fileInfoPanel.setStyleAttribute("width","100%");
  fileInfoPanel.setStyleAttribute("backgroundColor","#F5F5F5");
  fileInfoPanel.setStyleAttribute("padding","5px");
  fileInfoPanel.add(folderLabel);
  var folderPanel = app.createHorizontalPanel();
  folderPanel.add(folderListBox).add(secondaryFolder);
  fileInfoPanel.add(folderPanel);
  fileInfoPanel.add(fileNameLabel);
  fileInfoPanel.add(fileNameStringBox);
  fileInfoPanel.add(fileNameHelpLabel);
  fileInfoPanel.add(fileTypeLabel);
  fileInfoPanel.add(fileTypeSelectBox);
  fileInfoPanel.add(linkCheckBox);
  
  app.add(refreshPanel);
  panel.add(fileToFolderCheckBox); 
  panel.add(fileToFolderCheckBoxFalse);
  panel.add(fileInfoPanel);


 
  var emailUnCheckHandler = app.createClientHandler().forTargets(fileToEmailCheckBox, emailInfoPanel).setVisible(false)
                                                    .forTargets(fileToEmailCheckBoxFalse).setVisible(true);
  var emailUnSetCheck = app.createServerHandler('autoCrat_unSetEmailCheck').addCallbackElement(fileToEmailCheckBox);
  var emailCheckHandler = app.createClientHandler().forTargets(fileToEmailCheckBox, emailInfoPanel).setVisible(true)
                                                  .forTargets(fileToEmailCheckBoxFalse).setVisible(false);
  var emailSetCheck = app.createServerHandler('autoCrat_setEmailCheck').addCallbackElement(fileToEmailCheckBox);
  
  fileToEmailCheckBox.addClickHandler(emailUnSetCheck).addClickHandler(emailUnCheckHandler);
  fileToEmailCheckBoxFalse.addClickHandler(emailCheckHandler).addClickHandler(emailSetCheck);

  if ((fileSetting == "false")||(!fileSetting)) {
  fileToEmailCheckBox.setEnabled(false).setValue(false);
  fileToEmailCheckBoxFalse.setEnabled(false).setValue(false);
  }

  if (fileSetting == "true") {
  fileToEmailCheckBox.setEnabled(true);
  fileToEmailCheckBoxFalse.setEnabled(true);
  }
  
  emailInfoPanel.setStyleAttribute("width","100%");
  emailInfoPanel.setStyleAttribute("backgroundColor","#F5F5F5");
  emailInfoPanel.setStyleAttribute("padding","5px");
  emailInfoPanel.add(emailLabel);
  emailInfoPanel.add(emailStringBox);
  emailInfoPanel.add(emailHelpLabel);
  emailInfoPanel.add(emailSubjectLabel);
  emailInfoPanel.add(emailSubjectBox);
  emailInfoPanel.add(bodyPrefixHelpLabel);
  emailInfoPanel.add(bodyPrefixTextArea);
  emailInfoPanel.add(emailAttachmentLabel);
  emailInfoPanel.add(emailAttachmentListBox);
  
  panel.add(fileToEmailCheckBox); 
  panel.add(fileToEmailCheckBoxFalse);
  panel.add(emailInfoPanel);
   
  var formTrigger = ScriptProperties.getProperty('formTrigger');
  var mergeTriggerCheckBox = app.createCheckBox().setText("Trigger merge on form submit").setName("formTrigger");
  if (formTrigger=="true") {
    mergeTriggerCheckBox.setValue(true);
  }
  panel.add(mergeTriggerCheckBox); 
  panel.add(button);

  //Help text below dynamically loads all field names from the sheet using normalized (camelCase) sheet headers
  var fieldHelpText = "Use these variables to include values from the spreadsheet in any of the fields below.";
  var fieldHelpLabel = app.createLabel().setText(fieldHelpText);
  var fieldHelpTable = app.createFlexTable();
  fieldHelpTable.setWidget(0, 0, app.createLabel("$currDate (adds the current date in mm.dd.yy format)")).setStyleAttribute('color', 'blue');
  for (var i = 0; i<normalizedSheetFieldNames.length; i++) {
    var variable = app.createLabel("$"+normalizedSheetFieldNames[i]).setStyleAttribute('color', 'blue');
    fieldHelpTable.setWidget(i+1, 0, variable)
  }
  var instructions = app.createLabel().setText("Note: Merge will only execute for rows with no entry in the \"Merge Status\" column, and that meet any \"Merge conditions\" you may have set.");
  instructions.setStyleAttribute("font-weight", "bold");
  
  panel.add(instructions);
  app.add(fieldHelpLabel);
  varScrollPanel.add(fieldHelpTable);
  scrollPanel.add(panel);
  app.add(varScrollPanel);
  app.add(refreshPanel);
  app.add(scrollPanel);
  app.add(spinner);
  ss.show(app);
}

// Nutty server handler to always uncheck the ghost checkbox that shows only when true checkbox is clicked

function autoCrat_unSetEmailCheck () {
  var app = UiApp.getActiveApplication();
  var fileToEmailCheckBox = app.getElementById('fileToEmailCheckBox');
  fileToEmailCheckBox.setValue(false);
  var fileToEmailCheckBoxFalse = app.getElementById('fileToEmailCheckBoxFalse');
  fileToEmailCheckBoxFalse.setValue(false);
  return app;
}

// Nutty server handler to always uncheck the ghost checkbox that shows only when false checkbox is clicked

function autoCrat_setEmailCheck () {
  var app = UiApp.getActiveApplication();
  var fileToEmailCheckBox = app.getElementById('fileToEmailCheckBox');
  fileToEmailCheckBox.setValue(true);
  var fileToEmailCheckBoxFalse = app.getElementById('fileToEmailCheckBoxFalse');
  fileToEmailCheckBoxFalse.setValue(false);
  return app;
}

// More of the same craziness

function autoCrat_unsetFileCheck () {
  var app = UiApp.getActiveApplication();
  var fileToFolderCheckBox = app.getElementById('fileToFolderCheckBox');
  fileToFolderCheckBox.setValue(false);
  var fileToFolderCheckBoxFalse = app.getElementById('fileToFolderCheckBoxFalse');
  fileToFolderCheckBoxFalse.setValue(false);
  return app;
}

function autoCrat_setFileCheck () {
  var app = UiApp.getActiveApplication();
  var fileToFolderCheckBox = app.getElementById('fileToFolderCheckBox');
  fileToFolderCheckBox.setValue(true);
  var fileToFolderCheckBoxFalse = app.getElementById('fileToFolderCheckBoxFalse');
  fileToFolderCheckBoxFalse.setValue(false);
  return app;
}


//This function loads all Test/Run Merge settings into script properties
// and does some handling for different user error scenarios

function autoCrat_saveRunSettings(e) {
  var ss = SpreadsheetApp.getActive();
  var app = UiApp.getActiveApplication();
  var destinationFolderId = e.parameter.destinationFolderId;
  var secondaryFolderId = e.parameter.secondaryFolderId;
  if (secondaryFolderId=="Optional: Additional folder key(s) here") {
    secondaryFolderId = "";
  }
  var fileSetting = e.parameter.fileToFolderCheckValue;
  var emailSetting = e.parameter.fileToEmailCheckValue;
  var fileNameString = e.parameter.fileNameString;
  var linkToDoc = e.parameter.linkToDoc;
  var emailString = e.parameter.emailString;
  var fileType = e.parameter.fileType;
  var emailSubject = e.parameter.emailSubject;
  var bodyPrefix = e.parameter.bodyPrefix;
  var emailAttachment = e.parameter.emailAttachment;
  var formTrigger = e.parameter.formTrigger;
  
  if (linkToDoc=="true") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheetName = ScriptProperties.getProperty('sheetName');
    var headers = autoCrat_fetchSheetHeaders(sourceSheetName);
    var sheet = ss.getSheetByName(sourceSheetName);
    var lastCol = sheet.getLastColumn();
    var statusCol = headers.indexOf("Document Merge Status");
    var linkCol = headers.indexOf("Link to merged Doc");
    if (linkCol == -1) {
      sheet.insertColumnBefore(statusCol+1);
      sheet.getRange(1, statusCol+1).setValue("Link to merged Doc").setBackgroundColor("black").setFontColor("white").setComment("Required by autoCrat script. Do not change the text of this header");
      sheet.insertColumnBefore(statusCol+1);
      sheet.getRange(1, statusCol+1).setValue("Merged Doc URL").setBackgroundColor("black").setFontColor("white").setComment("Required by autoCrat script. Do not change the text of this header");
      sheet.insertColumnBefore(statusCol+1);
      sheet.getRange(1, statusCol+1).setValue("Merged Doc ID").setBackgroundColor("black").setFontColor("white").setComment("Required by autoCrat script. Do not change the text of this header");
    }
  } 
  
  if (formTrigger=="true") {
  var triggers = ScriptApp.getScriptTriggers();
  var triggerSetFlag = false;
  for (var i = 0; i<triggers.length; i++) {
    var eventType = triggers[i].getEventType();
    var triggerSource = triggers[i].getTriggerSource();
    var handlerFunction = triggers[i].getHandlerFunction();
    if ((handlerFunction=='autoCrat_onFormSubmit')&&(eventType=="ON_FORM_SUBMIT")&&(triggerSource=="SPREADSHEETS")) {
      triggerSetFlag = true;
      break;
    }
  }
  if (triggerSetFlag==false) {
    autoCrat_setFormTrigger();
  }
  }

//Do a bunch of error handling stuff for all permutations of settings values that don't make sense

  if ((fileSetting=="false")&&(emailSetting=="false")) {
      Browser.msgBox("You must select a merge type before you can run a merge job.");
      autoCrat_runMergeConsole();
      return;
  }

  if ((fileSetting=="true")&&(!fileNameString)) {
       Browser.msgBox("If you are saving your merge job to Docs, you must set a file naming convention.");
       autoCrat_runMergeConsole();
       return;
       }
  
  if ((emailSetting=="true")&&(!emailString)) {
       Browser.msgBox("If you want to email this merge job, you must set at least one recipient email address.");
       autoCrat_runMergeConsole();
       return;
      }
 
  ScriptProperties.setProperty('fileSetting',fileSetting);
  ScriptProperties.setProperty('fileNameString', fileNameString);
  ScriptProperties.setProperty('fileType', fileType);
  ScriptProperties.setProperty('linkToDoc', linkToDoc);
  ScriptProperties.setProperty('emailSetting',emailSetting);
  ScriptProperties.setProperty('emailString', emailString);
  ScriptProperties.setProperty('emailSubject', emailSubject);
  ScriptProperties.setProperty('bodyPrefix', bodyPrefix);
  ScriptProperties.setProperty('emailAttachment', emailAttachment);
  ScriptProperties.setProperty('formTrigger', formTrigger);

  if (destinationFolderId=="Select destination folder") {
    Browser.msgBox("You forgot to choose a destination folder for your merged Docs.");
    autoCrat_runMergeConsole();
  }
  var parent = DocsList.getFileById(ss.getId()).getParents()[0];
  if (!(parent)) {
    parent = DocsList.getRootFolder();
  }
  var folders = parent.getFolders();
  var indexFlag = 1;
  ScriptProperties.setProperty('destinationFolderId', destinationFolderId);
  ScriptProperties.setProperty('secondaryFolderId', secondaryFolderId);
  var destinationFolderName = DocsList.getFolderById(destinationFolderId).getName();
  ScriptProperties.setProperty('destinationFolderName', destinationFolderName);
  if (secondaryFolderId.indexOf("$")!=-1) {
    ScriptProperties.setProperty('secondaryFolderToken', secondaryFolderId);
  }
  autoCrat_initialize();
  autoCrat_runMergePrompt();
  app.close();
  return app; 
}


function autoCrat_runMergePrompt() {
  var app = UiApp.createApplication().setTitle('Step 6: Preview/Run Merge');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = ScriptProperties.getProperty("sheetName");
  var sheet = ss.getSheetByName(sheetName);
  var headers = autoCrat_fetchSheetHeaders(sheetName);
  if (headers.indexOf("")!=-1) {
    Browser.msgBox("It appears one of your merge sheet headers is blank, which the autoCrat does not allow!");
    app.close();
    return app;
  }
  var panel = app.createVerticalPanel();
  var label = app.createLabel("Note that the Google Docs service can cause this script to execute slowly, and there are quotas for the total number of script-generated Docs you can create in a day.  Visit https://docs.google.com/macros/dashboard to learn about your quotas. For large merge jobs, the autoCrat will pause and then automatically restart to avoid a service timeout.");
  panel.add(label);
  //create spinner graphic to show upon button click awaiting merge completion
  var refreshPanel = app.createFlowPanel();
  refreshPanel.setId('refreshPanel');
  refreshPanel.setStyleAttribute("width", "100%");
  refreshPanel.setVisible(false);

//Adds the graphic for the waiting period before merge completion. Set invisible until client handler
//is called by button click 
  var spinner = app.createImage(this.AUTOCRATIMAGEURL).setHeight("220px");
  spinner.setStyleAttribute("opacity", "1");
  spinner.setStyleAttribute("position", "absolute");
  spinner.setStyleAttribute("top", "50px");
  spinner.setStyleAttribute("left", "100px");
  spinner.setId("dialogspinner");
  var waitingLabel1 = app.createLabel("Performing full merge...this sometimes takes a while;)").setStyleAttribute('textAlign', 'center').setVisible(false);
  var waitingLabel2 = app.createLabel("Previewing merge on first row...").setStyleAttribute('textAlign', 'center').setVisible(false);
  refreshPanel.add(spinner);
  refreshPanel.add(waitingLabel1);
  refreshPanel.add(waitingLabel2);
  app.add(refreshPanel);
  
  var horiz = app.createHorizontalPanel();
  var handler1 = app.createServerHandler('autoCrat_runMerge').addCallbackElement(panel);
  var spinnerHandler1 = app.createClientHandler().forTargets(refreshPanel).setVisible(true).forTargets(panel).setVisible(false).forTargets(waitingLabel1).setVisible(true);
  var spinnerHandler2 = app.createClientHandler().forTargets(refreshPanel).setVisible(true).forTargets(panel).setVisible(false).forTargets(waitingLabel2).setVisible(true);
  var button1 = app.createButton('Run merge now').addClickHandler(handler1).addClickHandler(spinnerHandler1);
  var handler2 = app.createServerHandler('autoCrat_previewMerge').addCallbackElement(panel);
  var button2 = app.createButton('Preview on first row only').addClickHandler(handler2).addClickHandler(spinnerHandler2);
  var handler3 = app.createServerHandler('autoCrat_exit').addCallbackElement(panel);
  var button3 = app.createButton('Not now, just keep my settings').addClickHandler(handler3);
  horiz.add(button2);
  horiz.add(button1);
  horiz.add(button3);
  panel.add(horiz);
  app.add(panel);
  ss.show(app);
  return app;
}

function autoCrat_previewMerge() {
  autoCrat_runMerge("true");
}


function autoCrat_exit(e) {
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

// This function is where the actual merge is executed, entirely from values stored in 
// Script Properties.  This is the function that the Google Form submit trigger calls 

function autoCrat_runMerge(preview) {
 var ssKey = ScriptProperties.getProperty('ssKey');
 var ss = SpreadsheetApp.openById(ssKey);
 autoCrat_removeTimeoutTrigger();
//Do some more error handling just in case
 var fileId = ScriptProperties.getProperty('fileId');
  if (!fileId) { 
       Browser.msgBox("You must select a template file before you can run a merge.");
       autoCrat_defineTemplate();
       return;
       }
  var mappingString = ScriptProperties.getProperty('mappingString');
  if (!mappingString) {
       Browser.msgBox("You must map document fields before you can run a merge.");
       autoCrat_defineSettings();
       return;
       }
  var sheetName = ScriptProperties.getProperty('sheetName');
  if (!sheetName) {
       Browser.msgBox("You must select a source data sheet before you can run a merge.");
       autoCrat_defineSettings();
       return;
       }
  
  var sheet = ss.getSheetByName(sheetName);
  var now = new Date();
  var headers = autoCrat_fetchSheetHeaders(sheetName);
  for (var h=0; h<headers.length; h++) {
    if (autoCrat_normalizeHeader(headers[h])=="") {
      Browser.msgBox("Ooops! You must have an illegal header value in column " + (h+1) + ".  Headers cannot be blank or purely numeric.  Please fix.");
      return;
    }
  }
  var statusCol = headers.indexOf("Document Merge Status");
  var lastCol = sheet.getLastColumn();
  var linkCol = headers.indexOf("Link to merged Doc");
  var urlCol = headers.indexOf("Merged Doc URL");
  var docIdCol = headers.indexOf("Merged Doc ID");
  
  //status message will get concatenated as the logic tree progresses through the merge
  var mergeStatusMessage = "";

  var fileSetting = ScriptProperties.getProperty('fileSetting');
  var fileNameString = ScriptProperties.getProperty('fileNameString');
  
  var linkToDoc = ScriptProperties.getProperty('linkToDoc');
  
  
  // error handling
  if ((!fileNameString)&&(fileSetting==true)) {
       Browser.msgBox("You must set a file naming convention before you can run a merge.");
       autoCrat_runMergeConsole();
       return;
       }

  var testOnly = preview;

  // avoids loading any stale settings from old file or email merges
  if (fileSetting == "true") {
  var fileTypeSetting = ScriptProperties.getProperty('fileType');
  var destFolderId = ScriptProperties.getProperty('destinationFolderId');
  var secondaryFolderIdString = ScriptProperties.getProperty('secondaryFolderId');
  var secondaryFolderIdArray = [];
    secondaryFolderIdString = secondaryFolderIdString.replace(/\s+/g, '');
    if ((secondaryFolderIdString)&&(secondaryFolderIdString!='')) {
      secondaryFolderIdArray = secondaryFolderIdString.split(",");
    }
  }
  var emailSetting = ScriptProperties.getProperty('emailSetting');
  if (emailSetting == "true") {
  var emailString = ScriptProperties.getProperty('emailString'); 
  var emailSubject = ScriptProperties.getProperty('emailSubject');
  var emailAttachment = ScriptProperties.getProperty('emailAttachment');
  }
  var row2values = sheet.getRange(2,1,1,sheet.getLastColumn()).getValues()[0];
  if (row2values.indexOf("N/A: This is the formula row.")!=-1) {
    sheet.getRange(2,statusCol+1).setValue("N/A: This is the formula row.").setFontColor('white').setBackground('black');
    Utilities.sleep(5000);
  }
  var normalizedHeaders = autoCrat_normalizeHeaders(headers);
  var mergeTags = "$"+normalizedHeaders.join(",$");
  
  //this array will be used to do replacements in all file name, subject, and email body settings
  mergeTags = mergeTags.split(",");
  var fullRange = sheet.getDataRange();
  var lastCol = fullRange.getLastColumn();
  var lastRow = fullRange.getLastRow();
  // copyId will be used to pass the Doc Id of the merged copy through different branches of the merge logic
  var copyId = "0";
  
  // Load the merge fields from the template
  var mergeFields = autoCrat_fetchDocFields(fileId);
  
  //Create a temporary folder in the case that the user doesn't have
  //the "Save merged files to Docs" option checked.
  //This folder will be kept if they select the shared Google Docs method
  if ((emailSetting == "true") && (fileSetting == "false")) {
    var tempFolderId = DocsList.createFolder('Merged Docs from' + now).getId();
   }
  var count = 0;
  
  //Load conditions and mapping object
  var conditionString = ScriptProperties.getProperty('mergeConditions');
  var mappingString = ScriptProperties.getProperty('mappingString');
  var mappingObject = Utilities.jsonParse(mappingString);
  
  //Load in the range
  var range = sheet.getDataRange();
  var rowValueArray = range.getValues();
  var rowFormatArray = range.getNumberFormats();
  var templateFileType = DocsList.getFileById(fileId).getFileType().toString();
  templateFileType = templateFileType.charAt(0).toUpperCase() + templateFileType.slice(1);
  
  //Commence the merge loop, run through sheet
  for (var i=1; i<lastRow; i++) {
   var loopTime = new Date();
   var timeElapsed = parseInt(loopTime - now);
    if (timeElapsed > 295000) {
      autoCrat_preemptTimeout();
      return;
    }
   var rowValues = rowValueArray[i];
   var rowFormats = rowFormatArray[i];
    
  //reload sheet to ensure status messages update
   var sheet = ss.getSheetByName(sheetName);
 
  // Test conditions on this row
   var conditionTest = autoCrat_evaluateConditions(conditionString, 0, rowValues, normalizedHeaders);
  //Only run a merge for records that have no existing value in the last column ("Document Merge Status"), and that passes condition test
    if ((rowValues[statusCol]=="")&&(conditionTest==true)) {
      var sheet = ss.getSheetByName(sheetName);
  
    //First big branch in the merge logic: If file setting is true, do all necessary things 
   if (fileSetting == "true") {
     // custom function replaces "$variables" with values in a string
     var fileName = autoCrat_replaceStringFields(fileNameString, rowValues, rowFormats, headers, mergeTags);
     var secondaryFolderIds = [];
     if (secondaryFolderIdArray.length>0) {
       for (var z=0; z<secondaryFolderIdArray.length; z++) {
         if (secondaryFolderIdArray[z].indexOf("$")!=-1) {
           secondaryFolderIds.push(autoCrat_replaceStringFields(secondaryFolderIdArray[z], rowValues, rowFormats, headers, mergeTags));
         } else {
           secondaryFolderIds.push(secondaryFolderIdArray[z]);
         }
       }
     }
     //Google Doc created by default.  Custom function to replace all <<merge tags>> with mapped fields
     var copyId = DocsList.getFileById(fileId).makeCopy(fileName).getId();
     if (docIdCol!=-1) {
       sheet.getRange(i+1, docIdCol+1, 1, 1).setValue(copyId);
       SpreadsheetApp.flush();
       rowValues = sheet.getRange(i+1, 1, 1, sheet.getLastColumn()).getValues()[0];
     }
     copyId = autoCrat_makeMergeDoc(copyId, rowValues, rowFormats, destFolderId, secondaryFolderIds, mergeFields, mappingObject);
     try {
       autoCrat_logDocCreation();
     } catch(err) {
     }

     //PDF created only if set
     if (fileTypeSetting == "PDF") {  
       var pdfId = autoCrat_converToPdf(copyId, destFolderId, secondaryFolderIds); 
       mergeStatusMessage += "PDF successfully created,"; 
       autoCrat_trashDoc(copyId);
       copyId = '';
     } else {
       mergeStatusMessage += 'Google ' + templateFileType + ' successfully created,'; 
     }
   }

//handle the case where the email option is chosen without the save in docs option
//set temporary file name and create the merged file in the temporary folder
  if ((emailSetting == "true") && (fileSetting == "false")) {
    var tempFileName = "Merged File #" + (i-1);
    var copyId = DocsList.getFileById(fileId).makeCopy(tempFileName).getId();
    var copyId = autoCrat_makeMergeDoc(copyId, rowValues, rowFormats, tempFolderId, secondaryFolderIds, mergeFields, mappingObject);
  }
  
// 2nd major logic branch in merge
  if (emailSetting == "true") {
  try {
    // replace $variables in strings
     var recipients = autoCrat_replaceStringFields(emailString, rowValues, rowFormats, headers, mergeTags);
    //remove whitespaces from email strings 
    recipients = recipients.replace(/\s+/g, '');
    //remove trailing commas from email strings
    recipients = recipients.replace(/,$/,'');
     var subject = autoCrat_replaceStringFields(emailSubject, rowValues, rowFormats, headers, mergeTags);
     var bodyPrefix = ScriptProperties.getProperty('bodyPrefix');
     var user = Session.getActiveUser().getUserLoginId();
     bodyPrefix = autoCrat_replaceStringFields(bodyPrefix, rowValues, rowFormats, headers, mergeTags);
    //Time to use a switch case. Woot!
     switch (emailAttachment){    
      case "PDF":
         //Hard to grasp, but look to see if PDF already exists from file merge and use it if so
         //attach the file and set email properties
         if (pdfId) {
            var attachment = DocsList.getFileById(pdfId);
            var body = '<table style="border:1px; padding:15px; background-color:#DDDDDD"><tr><td>' + user + ' has attached a PDF file to this email.</td><tr></table><br /><br />';
            body += bodyPrefix;
            MailApp.sendEmail(recipients, subject, body, {htmlBody: body, attachments: attachment});
            mergeStatusMessage += "PDF attached in email to " + recipients;
         }
         // copyId should only exist if PDF option isn't selected in the file method
         //attach the file and set email properties
         if (copyId) {
            var attachment = DocsList.getFileById(copyId).getAs("application/pdf");
            attachment.setName(DocsList.getFileById(copyId).getName() + ".pdf");
            var body = '<table style="border:1px; padding:15px; background-color:#DDDDDD"><tr><td>' + user + ' has attached a PDF file to this email.</td><tr></table><br /><br />';
            body += bodyPrefix;
            MailApp.sendEmail(recipients, subject, body, {htmlBody: body, attachments: attachment});
            mergeStatusMessage += "PDF attached in email to " + recipients;
         }
        break;
         
      case "Recipient-view-only Google Doc":
           var file = DocsList.getFileById(copyId);
         //add email recipients as doc viewers
           file.addViewers(recipients.split(","));
           var docUrl = file.getUrl();
           var docTitle = file.getName();
         // Add a little note on sharing as caring
           var body = '<table style="border:1px; padding:15px; background-color:#DDDDDD"><tr><td>' + user + ' has just shared this view-only Google ' + templateFileType + ' with you:</td><td><a href = "' + docUrl + '">' + docTitle + '</a></td></tr></table><br /><br />';
           body += bodyPrefix;
           MailApp.sendEmail(recipients, subject, body, {htmlBody: body});
           mergeStatusMessage += " View-only " + templateFileType + " shared with " + recipients + " ";
        break;
         
      case "Recipient-editable Google Doc":
           var file = DocsList.getFileById(copyId);
         //add email recipients as doc editors
           file.addEditors(recipients.split(","));
           var docUrl = file.getUrl();
           var docTitle = file.getName();
           var user = Session.getActiveUser().getUserLoginId();
           var body = '<table style="border:1px; padding:15px; background-color:#DDDDDD"><tr><td>' + user + ' has just shared this editable Google ' + templateFileType + ' with you:</td><td><a href = "' + docUrl + '">' +  docTitle + '</a></td></tr></table><br /><br />';
           body += bodyPrefix;
           MailApp.sendEmail(recipients, subject, body, {htmlBody: body});
           mergeStatusMessage += " Editable " + templateFileType + " shared with " + recipients + " ";
        break;
      }
      } catch(err) {
        mergeStatusMessage += err;
      }
    }
      

 //Purge the file if user doesn't want it saved.
 //Leaves files that have been shared as docs in the temporary folder, unless
 //the user has specified the folder
 if ((emailSetting == "true") && (fileSetting == "false") && (emailAttachment=="PDF")){
    mergeStatusMessage += ", file not saved in Docs."
    autoCrat_trashDoc(copyId);
  }
  mergeStatusMessage += now;
      
  if ((linkToDoc=="true")&&((copyId)||(pdfId))) {  
    var range1 = sheet.getRange(i+1, linkCol+1, 1, 1);
    var range2 = sheet.getRange(i+1, urlCol+1, 1, 1);
    var range3 = sheet.getRange(i+1, docIdCol+1, 1, 1);
    var range4 = sheet.getRange(i+1, statusCol+1, 1, 1);
    if (pdfId!='') {
      var mergeFileId = pdfId;
    }
    if (copyId!='') {
      var mergeFileId = copyId;
    }
    var mergeFile = DocsList.getFileById(mergeFileId)
    var url = mergeFile.getUrl();
    var urlValue = [[url]];
    var copyIdValue = [[mergeFileId]];
    var mergeTitle = mergeFile.getName();
    var link = [['=hyperlink("' + url + '", "' + mergeTitle + '")']];
    var mergeStatusMessage = [[mergeStatusMessage]];
    range1.setValues(link);
    range2.setValues(urlValue);
    range3.setValues(copyIdValue);
    range4.setValues(mergeStatusMessage);
  } else {
    var range = sheet.getRange(i+1, statusCol+1);
    range.setValue(mergeStatusMessage);
  }    
  mergeStatusMessage = "";
  count = count+1;
}
  if ((emailSetting == "true") && (fileSetting == "false") && (emailAttachment=="PDF")){
var folder = DocsList.getFolderById(tempFolderId);
    folder.setTrashed(true);
}
    if ((testOnly=="true")&&(count>0)) { break; }
}
  //Extra fancy merge completion confirmation
  //Even fancier: What would it take to have a progress bar embedded in the loop?
  if (count!=0) {
    Browser.msgBox("Merge successfully completed for " + count + " record(s).");
  } else { 
    Browser.msgBox("For some reason, no record(s) were successfully merged.  If you haven't cleared the merge-status messages, that could be the issue.");
  }
}

// This function subs in row values for $variables
function autoCrat_replaceStringFields(string, rowValues, rowFormats, headers, mergeTags) {
  var newString = string;
  var timeZone = Session.getTimeZone();
  for (var i=0; i<headers.length; i++) {
    var thisHeader = headers[i];
    var colNum = autoCrat_getColumnNumberFromHeader(thisHeader, headers);
 if (((rowFormats[colNum-1]=="M/d/yyyy")||(rowFormats[colNum-1]=="MMMM d, yyyy")||(rowFormats[colNum-1]=="M/d/yyyy H:mm:ss"))&&(rowValues[colNum-1]!="")) {
   try {
      var replacementValue = Utilities.formatDate(rowValues[colNum-1], timeZone, rowFormats[colNum-1]);
      }
   catch(err) {
      var date = new Date(rowValues[colNum-1]);
      var colVal = Utilities.formatDate(date, timeZone, rowFormats[colNum-1]);
      }
    } else {
     var replacementValue = rowValues[colNum-1];
    }
    var replaceTag = mergeTags[i];
    replaceTag = replaceTag.replace("$","\\$") + "\\b";
    var find = new RegExp(replaceTag, "g");
    newString = newString.replace(find,replacementValue);
  }
  var currentTime = new Date()
  var month = currentTime.getMonth() + 1;
  var day = currentTime.getDate();
  var year = currentTime.getFullYear();
  newString = newString.replace("$currDate", month+"/"+day+"/"+year);
  return newString;
}


//Creates PDF in a designated folder and returns Id
function autoCrat_converToPdf (copyId, folderId, secondaryFolderIds) {
  var folder = DocsList.getFolderById(folderId);
  var pdfBlob = DocsList.getFileById(copyId).getAs("application/pdf"); 
  var pdfFile = DocsList.createFile(pdfBlob);
  pdfFile.rename(DocsList.getFileById(copyId).getName() + ".pdf");
  pdfFile.addToFolder(folder);
  if (secondaryFolderIds) {
    for (var z=0; z<secondaryFolderIds.length; z++) {
      try {
        var secondaryFolder = DocsList.getFolderById(secondaryFolderIds[z])
        } catch(err) {
          continue;
        }
      pdfFile.addToFolder(secondaryFolder);
    }
  }
  pdfFile.removeFromFolder(DocsList.getRootFolder());
  var pdfId = pdfFile.getId();
  return pdfId;
}

//Trashes given docIdFind and replace
function autoCrat_trashDoc (docId) {
   DocsList.getFileById(docId).setTrashed(true);
}


function autoCrat_makeMergeDoc(copyId, rowValues, rowFormats, folderId, secondaryFolderIds, mergeFields, mappingObject) {
   // Get document template, copy it as a new temp doc, and save the Doc’s id
  var fileType = DocsList.getFileById(copyId).getFileType().toString();
  if ((fileType=="document")||(fileType=="DOCUMENT")) { 
   // Open the temporary document
   var copyDoc = DocumentApp.openById(copyId);
   // Get the document’s body section
   var copyHeader = copyDoc.getHeader();
   var copyBody = copyDoc.getActiveSection();
   var copyFooter = copyDoc.getFooter();
   // Get the mappingString
   for (i in mappingObject) {
   };
  for (i=0; i< mergeFields.length; i++) {
     var normalizedFieldName = autoCrat_normalizeHeader(mergeFields[i]);
     var colNum = mappingObject[normalizedFieldName];
     var timeZone = Session.getTimeZone();
    if ((rowFormats[colNum-1]=="M/d/yyyy")||(rowFormats[colNum-1]=="MMMM d, yyyy")||(rowFormats[colNum-1]=="M/d/yyyy H:mm:ss")) {
      try {
     var colVal = Utilities.formatDate(rowValues[colNum-1], timeZone, rowFormats[colNum-1]);
      }
      catch(err) {
      var date = new Date(rowValues[colNum-1]);
      var colVal = Utilities.formatDate(date, timeZone, rowFormats[colNum-1]);
      }
    } else {
     var colVal = rowValues[colNum-1];
    }
   if (copyHeader) {
     copyHeader.replaceText(mergeFields[i], colVal);
     }
   if (copyBody) {
     copyBody.replaceText(mergeFields[i], colVal);
     }
   if (copyFooter) {
     copyFooter.replaceText(mergeFields[i], colVal);
     }
    }
  
// Save and close the temporary document
   copyDoc.saveAndClose();
  }
  if ((fileType=="spreadsheet")||(fileType=="SPREADSHEET")) {
    var exp = new RegExp(/[<]{2,}\S[^,]*?[>]{2,}/g);
    var ss = SpreadsheetApp.openById(copyId);
    var sheets = ss.getSheets();
    for (var i=0; i<sheets.length; i++) {
     var range = sheets[i].getDataRange();
     var values = range.getValues();
     var formulas = range.getFormulas();
     var formats = range.getNumberFormats();
      for (var j=0; j<values.length; j++) {
        for (var k=0; k<values[j].length; k++) {
          var cellValue = values[j][k].toString();
          var cellFormula = formulas[j][k].toString();
          var tags = cellValue.match(exp);
          if (tags) {
            for (var n=0; n<tags.length; n++) {
              var normalizedFieldName = autoCrat_normalizeHeader(tags[n]);
              var colNum = mappingObject[normalizedFieldName];
                var timeZone = Session.getTimeZone();
                if ((rowFormats[colNum-1]=="M/d/yyyy")||(rowFormats[colNum-1]=="MMMM d, yyyy")||(rowFormats[colNum-1]=="M/d/yyyy H:mm:ss")) {
                  try {
                    var colVal = Utilities.formatDate(rowValues[colNum-1], timeZone, rowFormats[colNum-1]);
                    formats[j][k] = rowFormats[colNum-1];
                  } catch(err) {
                    var date = new Date(rowValues[colNum-1]);
                    var colVal = Utilities.formatDate(date, timeZone, rowFormats[colNum-1]);
                    formats[j][k] = rowFormats[colNum-1];
                  }
                  } else {
                    var colVal = rowValues[colNum-1];
                  }
              if ((cellFormula)&&(cellFormula!='')) {
                cellFormula = cellFormula.replace(tags[n],colVal);
                formulas[j][k] = cellFormula;
              } else {
                cellValue = cellValue.replace(tags[n],colVal);
                values[j][k] = cellValue;
              }
            }
          }
        }
      }
      range.setValues(values);
      for (var j=0; j<formulas.length; j++) {
        for (var k=0; k<formulas[0].length; k++) {
          if (formulas[j][k]!='') {
            sheets[i].getRange(j+1, k+1).setFormula(formulas[j][k]);
          }
        }
      }
      sheets[i].activate();
    }
  }
  // move to folder
  var folder = DocsList.getFolderById(folderId);
  var file = DocsList.getFileById(copyId)
  file.addToFolder(folder);
  if (secondaryFolderIds) {
    for (var z=0; z<secondaryFolderIds.length; z++) {
      try {
        var secondaryFolder = DocsList.getFolderById(secondaryFolderIds[z])
        } catch(err) {
          continue;
        }
      file.addToFolder(secondaryFolder);
    }
  }
  // remove from Drive 
  file.removeFromFolder(DocsList.getRootFolder());
   return copyId;
}


// This function creates the field mapping UI, called from the spreadsheet menu
function autoCrat_mapFields() {
  var ss = SpreadsheetApp.getActive();
  var app = UiApp.createApplication();
  app.setTitle("Step 4: Set Field Mappings");
  app.setHeight("430");
  var panel = app.createScrollPanel().setId("fieldMappingPanel");
  var helpPopup = app.createDecoratedPopupPanel();
  var helpLabel = app.createLabel().setText("If all is working properly, Every <<Merge Tag>> in your document template should be listed on the left. Each tag needs to be \"mapped\" onto a corresponding column in your source data sheet.  IMPORTANT: using non-alpha-numeric characters {(),&%#etc} in tags will cause the merge to fail.");
  helpPopup.add(helpLabel);
  //create grid but without a size, for now
  var tagHeaderLabel = app.createLabel().setText("<<Tag From Doc Template>>");
  var fieldHeaderLabel = app.createLabel().setText("Sheet Header");
  var topGrid = app.createGrid(1,2);
  topGrid.setWidth("100%");
  topGrid.setWidget(0, 0, tagHeaderLabel);
  topGrid.setStyleAttribute(0,0,"backgroundColor", "#cfcfcf");
  topGrid.setStyleAttribute(0,0,"textAlign", "center");
  tagHeaderLabel.setStyleAttribute("fontSize", "16px");
  topGrid.setStyleAttribute(0,0,"width", "48%");
  topGrid.setWidget(0, 1, fieldHeaderLabel);
  topGrid.setStyleAttribute(0,1,"backgroundColor", "#cfcfcf");
  var grid = app.createGrid().setId("fieldMappingGrid");
  topGrid.setStyleAttribute(0,1,"textAlign", "center");
  fieldHeaderLabel.setStyleAttribute("fontSize", "16px");
  topGrid.setStyleAttribute(0,1,"width", "52%");
  grid.setWidth("100%")
  // grab the file id
  var fileId = ScriptProperties.getProperty('fileId');
  // If no template has been set, jump straight to settings UI
  if(!fileId) { 
    Browser.msgBox("You need to choose a template file before you can run a merge!");
    autoCrat_defineTemplate(); 
  }
  // go to file to look for all unique <<mergefields>>
  var docFieldNames = autoCrat_fetchDocFields(fileId);
  if (!docFieldNames) { 
     Browser.msgBox('The selected template contains no merge field tags..eg. <<merge field>>');
     app.close();
     autoCrat_defineTemplate();
  }
  //fetch data sheet from user-determined property
  var sheetName = ScriptProperties.getProperty('sheetName');
    //If not set send user straight to settings UI
    if (!sheetName) {
    Browser.msgBox("You need to choose a template file before you can run a merge!");
    app.close();
    autoCrat_defineTemplate(); 
   }
  //go to data sheet and return all header names
  var sheetFieldNames = autoCrat_fetchSheetHeaders(sheetName);

  if ((docFieldNames)&&(sheetFieldNames)) {
  //resize grid to fit number of unique fields in template
  grid.resize(docFieldNames.length+1, 2);
  //grab already-saved mappings, if they exist.  This is saved as a JSON-like string
  var mappingString = ScriptProperties.getProperty("mappingString");
  var mappingObject = Utilities.jsonParse(mappingString);
  //build Ui elements and assign indexed IDs and names
  // this technique allow the UI to expand to fit the number of tags in the template
    for (i=0; i<docFieldNames.length; i++) {
      var label = app.createLabel().setId("mergefield-" + i).setText(docFieldNames[i]);
      var listBox = app.createListBox().setId("header-" + i).setName("header-" + i);
      listBox.addItem("Choose column");
    for (j=0; j<sheetFieldNames.length; j++) {
      listBox.addItem(sheetFieldNames[j]);
    }
      var fieldName = autoCrat_normalizeHeader(docFieldNames[i]);
      var thisDocFieldName = docFieldNames[i].replace("<<","");
      thisDocFieldName = thisDocFieldName.replace(">>","");
      if (mappingObject) {
        for (var k in mappingObject) {
          //this line handles the case where the template has been edited and a previously mapped field is now missing
          //erases existing mappings and forces the user to re-do them
          if (!mappingObject[fieldName]) { ScriptProperties.setProperty('mappingString','');break;}
          var itemNo = parseInt(mappingObject[fieldName]);
        } 
        if (!mappingObject[fieldName]) { ScriptProperties.setProperty('mappingString',''); break;}  
    } else {
      var itemNo = 0;
      for (var m=0; m<sheetFieldNames.length; m++) {
        var sheetFieldName = autoCrat_normalizeHeader(sheetFieldNames[m]);
        if ((fieldName==sheetFieldName)||(thisDocFieldName==sheetFieldName)) {
          itemNo = m+1;
          break;
        }
      }
    }
      listBox.setSelectedIndex(itemNo); 
      grid.setWidget(i,0,label);
      grid.setStyleAttribute(i, 0, "background", "#F5F5F5");
    grid.setStyleAttribute(i, 0, "width", "50%");
    grid.setStyleAttribute(i, 0, "textAlign", "right");
    grid.setWidget(i,1,listBox);
    grid.setStyleAttribute(i, 1, "background", "#F5F5F5");
    grid.setStyleAttribute(i, 0, "width", "50%");
  }
  
  var spinner = app.createImage(AUTOCRATIMAGEURL).setWidth(150);
  spinner.setVisible(false);
  spinner.setStyleAttribute("position", "absolute");
  spinner.setStyleAttribute("top", "100px");
  spinner.setStyleAttribute("left", "200px");
  spinner.setId("dialogspinner");
  
  var clientHandler = app.createClientHandler().forTargets(spinner).setVisible(true).forTargets(panel).setStyleAttribute("opacity", "0.5");
  var sendHandler = app.createServerHandler('autoCrat_saveMappings').addCallbackElement(grid);
  var button = app.createButton().setId("mappingSubmitButton").addClickHandler(sendHandler).addClickHandler(clientHandler);
  button.setText("Save mappings");
  grid.setWidget(i, 0, app.createLabel().setText("Important: These mappings will hold true only if you don't modify the order of columns in the sheet. \n Dates must be formatted as \"M/d/yyyy\", \"MMMM d, yyyy\", or \"M/d/yyyy H:mm:ss\" using number formats in the spreadsheet").setStyleAttribute("fontSize","9px"));
  grid.setWidget(i, 1, button);
  panel.setStyleAttribute('overflow', 'scroll');
  panel.setHeight("345px");
  panel.add(grid);
  app.add(helpPopup);
  app.add(topGrid);
  app.add(panel);
  app.add(spinner);
  ss.show(app);
}
}

// saves Field mappings to script properties
// as a JSON-like string
function autoCrat_saveMappings(e) {
  var app = UiApp.getActiveApplication();
  var fileId = ScriptProperties.getProperty('fileId');
  var docFieldNames = autoCrat_fetchDocFields(fileId);
  var normalizedFieldNames = autoCrat_normalizeHeaders(docFieldNames);
  var sheetName = ScriptProperties.getProperty('sheetName');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
  var headers = autoCrat_fetchSheetHeaders(sheetName);
  var mappingString = "{";
  var header;
  var column;
  for (i=0; i<docFieldNames.length; i++) {
    header = e.parameter["header-" + i];
    //Handle errors
    if (header=="Choose column") { 
      var errorFlag=true; 
    }    
    if (header) {
    column = autoCrat_getColumnNumberFromHeader(header, headers);
    var fieldName = docFieldNames[i];
      if (fieldName) {
         var fieldName = autoCrat_normalizeHeader(fieldName);
      }
    mappingString += '"' + fieldName + '" : "' + column + '", '; 
    }
  } 
  mappingString += "}";
  var mappingObject = Utilities.jsonParse(mappingString);
  for (var i in mappingObject) {
  }
  
  ScriptProperties.setProperty('mappingString', mappingString);
  if (errorFlag==true) {
     Browser.msgBox("You forgot to assign a column to one or more of your mergefields."); 
     autoCrat_mapFields(); 
     return;
  }
   autoCrat_initialize();
  if (!(ScriptProperties.getProperty('fileSetting'))||!(ScriptProperties.getProperty('emailSetting'))) {
   autoCrat_runMergeConsole();
  }
   app.close();
   return app;
}


//Plucks the column numbers for the mapping string
//Implication: If a user changes the order of columns in the spreadsheet, 
//the mappings will get screwed up.
//Is there a way to erase the mappings or warn the user when column order is changed?

function autoCrat_getColumnNumberFromHeader(header, headers) {
  var colFlag = headers.indexOf(header) + 1;
  return colFlag;
}



function autoCrat_normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = autoCrat_normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}
 
// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function autoCrat_normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!autoCrat_isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && autoCrat_isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function autoCrat_getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (autoCrat_isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function autoCrat_isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}


// Returns true if the character char is alphabetical, false otherwise.
function autoCrat_isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    autoCrat_isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function autoCrat_isDigit(char) {
  return char >= '0' && char <= '9';
  
}

// Given a JavaScript 2d Array, this function returns the transposed table.
// Arguments:
//   - data: JavaScript 2d Array
// Returns a JavaScript 2d Array
// Example: arrayTranspose([[1,2,3],[4,5,6]]) returns [[1,4],[2,5],[3,6]].
function autoCrat_arrayTranspose(data) {
  if (data.length == 0 || data[0].length == 0) {
    return null;
  }

  var ret = [];
  for (var i = 0; i < data[0].length; ++i) {
    ret.push([]);
  }

  for (var i = 0; i < data.length; ++i) {
    for (var j = 0; j < data[i].length; ++j) {
      ret[j][i] = data[i][j];
    }
  }

  return ret;
}


// Grabs the headers from a sheet
function autoCrat_fetchSheetHeaders(sheetName) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  var cols = sheet.getLastColumn();
  var range = sheet.getRange(1,1,1,cols);
  var data = range.getValues();
  var headers = new Array();
  for (i = 0; i<data[0].length; i++) {
    headers[i] = data[0][i];
  }
  return headers;
}
 
//Grabs any <<Merge Tags>> from a document 
function autoCrat_fetchDocFields(fileId) {
  var fileType = DocsList.getFileById(fileId).getFileType().toString();
  if ((fileType=="DOCUMENT")||(fileType=="document")) { 
  var template = DocumentApp.openById(fileId);
  var title = template.getName();
  var fieldExp = "[<]{2,}\\S[^,]*?[>]{2,}";
  var result;
  var matchResults = new Array();
  var headerFieldNames = new Array();
  var bodyFieldNames = new Array();
  var footerFieldNames = new Array();
  
  //get all tags in doc header
  var header = template.getHeader();
  if (header!=null) { matchResults[0] = header.findText(fieldExp);}
  if (matchResults[0]!=null){
  var element = matchResults[0].getElement().asText().getText();
  var start = matchResults[0].getStartOffset()
  var end = matchResults[0].getEndOffsetInclusive()+1;
  var length = end-start;
  headerFieldNames[0] = element.substr(start,length)
    var i = 0;
    while (headerFieldNames[i]) {
      matchResults[i+1] = template.getHeader().findText(fieldExp, matchResults[i]);
      if (matchResults[i+1]) {
      var element = matchResults[i+1].getElement().asText().getText();
      var start = matchResults[i+1].getStartOffset()
      var end = matchResults[i+1].getEndOffsetInclusive()+1;
      var length = end-start;
      headerFieldNames[i+1] = element.substr(start,length);
      }
      i++;
    }
    }
    
   //get all tags in doc body
  matchResults = [];
  var body = template.getActiveSection();
  if (body!=null) { matchResults[0] = body.findText(fieldExp);}
  if (matchResults[0]!=null){
  var element = matchResults[0].getElement().asText().getText();
  var start = matchResults[0].getStartOffset()
  var end = matchResults[0].getEndOffsetInclusive()+1;
  var length = end-start;
  bodyFieldNames[0] = element.substr(start,length)
   var i = 0;
    while (bodyFieldNames[i]) {
      matchResults[i+1] = template.getActiveSection().findText(fieldExp, matchResults[i]);
      if (matchResults[i+1]) {
      var element = matchResults[i+1].getElement().asText().getText();
      var start = matchResults[i+1].getStartOffset()
      var end = matchResults[i+1].getEndOffsetInclusive()+1;
      var length = end-start;
      bodyFieldNames[i+1] = element.substr(start,length);
      }
      i++;
     }
    }
    
   //get all tags in doc footer
  var matchResults = [];
  var footer = template.getFooter();
  if (footer!=null) { matchResults[0] = footer.findText(fieldExp);}
  if (matchResults[0]!=null){
  var element = matchResults[0].getElement().asText().getText();
  var start = matchResults[0].getStartOffset()
  var end = matchResults[0].getEndOffsetInclusive()+1;
  var length = end-start;
  footerFieldNames[0] = element.substr(start,length)
    var i = 0;
    while (footerFieldNames[i]) {
      matchResults[i+1] = template.getFooter().findText(fieldExp, matchResults[i]);
      if (matchResults[i+1]) {
      var element = matchResults[i+1].getElement().asText().getText();
      var start = matchResults[i+1].getStartOffset()
      var end = matchResults[i+1].getEndOffsetInclusive()+1;
      var length = end-start;
      footerFieldNames[i+1] = element.substr(start,length);
      }
      i++;
     }
    }
  var fieldNames = headerFieldNames.concat(bodyFieldNames, footerFieldNames);
  fieldNames = autoCrat_removeDuplicateElement(fieldNames);
 return fieldNames; 
}
if ((fileType=="SPREADSHEET")||(fileType=="spreadsheet")) {
  var ss = SpreadsheetApp.openById(fileId);
  var sheets = ss.getSheets();
  var allTags = [];
  for (var i=0; i<sheets.length; i++) {
    var range = sheets[i].getDataRange();
    var values = range.getValues();
    for (var j=0; j<values.length; j++) {
      for (var k=0; k<values[j].length; k++) {
        var cellValue = values[j][k].toString();
        var exp = new RegExp(/[<]{2,}\S[^,]*?[>]{2,}/g);
        var tags = cellValue.match(exp);
        if (tags) {
          for (var l=0; l<tags.length; l++) {
            allTags.push(tags[l]);
          }
        }
      }
    }
  }
  allTags = autoCrat_removeDuplicateElement(allTags);
  return allTags;
}
}
  

//Takes out any duplicates from an array of values
function autoCrat_removeDuplicateElement(arrayName)
  {
  var newArray=new Array();
  label:for(var i=0; i<arrayName.length;i++ )
  {  
  for(var j=0; j<newArray.length;j++ )
  {
  if(newArray[j]==arrayName[i]) 
  continue label;
  }
  newArray[newArray.length] = arrayName[i];
  }
  return newArray;
  }


//Set merge conditions
function autoCrat_setMergeConditions() {
  var app = UiApp.createApplication().setTitle("Step 3: Set Merge Conditions (optional)");
  var panel = app.createVerticalPanel();
  var helppanel = app.createDecoratedPopupPanel();
  var label = app.createLabel('Use the widget below to set a field value condition that must be met for records to be merged to Docs.  Rows that do not meet the condition will be skipped and given a blank status message.  Leave the condition field blank to ignore.');
  helppanel.add(label);
  panel.add(helppanel);
  var conditionsGrid = app.createGrid(1,5).setId('conditionsGrid');
  var conditionLabel = app.createLabel('Field value');
  var dropdown = app.createListBox().setId('col-0').setName('col-0').setWidth("150px");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheetName = ScriptProperties.getProperty('sheetName');
  if ((sourceSheetName)&&(sourceSheetName!='')) {
     var sourceSheet = ss.getSheetByName(sourceSheetName);
  } else {
     var sourceSheet = ss.getSheets()[0];
  }
  var lastCol = sourceSheet.getLastColumn();
  var headers = autoCrat_fetchSheetHeaders(sourceSheetName);
  if (headers.indexOf("")!=-1) {
    Browser.msgBox("It appears at least one of your merge sheet headers is blank, which the autoCrat does not allow!");
    app.close();
    return app;
  }
  for (var i=0; i<headers.length; i++) {
    if ((headers[i]!="Status")) {
    dropdown.addItem(headers[i]);
    }
  }
  var equalsLabel = app.createLabel('equals');
  var textbox = app.createTextBox().setId('val-0').setName('val-0');
  var conditionHelp = app.createLabel('Leave blank to ignore. Use NULL for empty.  NOT NULL for not empty.').setStyleAttribute('fontSize', '8px');
  conditionsGrid.setWidget(0, 0, conditionLabel);
  conditionsGrid.setWidget(0, 1, dropdown);
  conditionsGrid.setWidget(0, 2, equalsLabel);
  conditionsGrid.setWidget(0, 3, textbox);
  conditionsGrid.setWidget(0, 4, conditionHelp);
  
  var conditionString = ScriptProperties.getProperty('mergeConditions');
  if ((conditionString)&&(conditionString!='')) {
    var conditionObject = Utilities.jsonParse(conditionString);
    var selectedHeader = conditionObject['col-0'];
    var selectedIndex = 0;
    for (var i=0; i<headers.length; i++) {
      if (headers[i]==selectedHeader) {
        selectedIndex = i;
        break;
      }
    }
    dropdown.setSelectedIndex(selectedIndex);
    var selectedValue = conditionObject['val-0'];
    textbox.setValue(selectedValue);
  }
  var spinner = app.createImage(AUTOCRATIMAGEURL).setWidth(150);
  spinner.setVisible(false);
  spinner.setStyleAttribute("position", "absolute");
  spinner.setStyleAttribute("top", "70px");
  spinner.setStyleAttribute("left", "180px");
  spinner.setId("dialogspinner");
  var clienthandler = app.createClientHandler().forTargets(spinner).setVisible(true).forTargets(panel).setStyleAttribute('opacity', '0.5');
  var handler = app.createServerHandler('autoCrat_saveCondition').addCallbackElement(panel);
  var button = app.createButton('Submit').addClickHandler(handler).addClickHandler(clienthandler);
  panel.add(conditionsGrid);
  panel.add(button);
  app.add(panel);
  app.add(spinner);
  ss.show(app);
}


function autoCrat_saveCondition(e) {
  var app = UiApp.getActiveApplication();
  var conditionObject = new Object();
  conditionObject['col-0'] = e.parameter['col-0'];
  conditionObject['val-0'] = e.parameter['val-0'];
  var conditionString = Utilities.jsonStringify(conditionObject);
  ScriptProperties.setProperty('mergeConditions', conditionString);
  if(!(ScriptProperties.getProperty("mappingString"))||(ScriptProperties.getProperty("mappingString")=="")) {
     autoCrat_mapFields();
  }
  app.close();
  return app;
}


//returns true if testval meets the condition 
function autoCrat_evaluateConditions(condString, index, rowData, normalizedHeaders) {
  if ((condString)&&(condString!='')) {
  var condObject = Utilities.jsonParse(condString);
  var i = index;
  var testHeader = autoCrat_normalizeHeader(condObject["col-"+i]);
   var colNum = -1;
  for (var j=0; j < normalizedHeaders.length; j++) {
    if (normalizedHeaders[j]==testHeader) {
      colNum = j;
      break;
    }
  }
  if (colNum == -1) {
    Browser.msgBox("Something is wrong with the merge conditions. Try resetting.");
    return;
  }
  var testVal = rowData[colNum];
  var value = condObject["val-"+i];
  var output = false;
  switch(value)
  {
  case "":
      output = true;
      break;
  case "NULL":
      if((!testVal)||(testVal=='')) {
        output = true;
      }  
    break;
  case "NOT NULL":
    if((testVal)&&(testVal!='')) {
        output = true;
      }  
    break;
  default:
    if(testVal==value) {
        output = true;
      }  
  }
  return output;
} else {
  return true;
}
}


function autoCrat_defineTemplate() {
  setAutocratUid();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //  var headers = ss.getActiveSheet().getRange(1, 1, 1, sheet.getLastColumn());
  var headers = ss.getSheets()[0].getRange(1, 1, 1, ss.getSheets()[0].getLastColumn()).getValues()[0];
  var normalizedHeaders = autoCrat_normalizeHeaders(headers);
  var app = UiApp.createApplication().setWidth(600).setHeight(350);
  var panel = app.createVerticalPanel().setId("panel").setWidth("100%");
  app.setTitle('Step 1: Select a template containing merge tags from your docs');
  panel.add(app.createLabel("Must be a \'Document\' or \'Spreadsheet\' in Google Docs format, with merge tags that contain no special characters or commas, and may NOT begin with numbers. Merge tag format shown below: "));
  var image = app.createImage("https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/download.png?attachauth=ANoY7cpTRdoqkEnOB2Godog9PVqEFovv__DxSqM02KnW3_FNFZZA0CneTMejWaPWWA01AOVfgHhAhYkUkBUfiZQbUm4DiYo82xjSIEYxokxpeHPWChhsI8TetLMKMn7V8inAH7HMCgogCUc5yEtVHUJkus-FbCBGvYmi9KKozeA9cFSu7Q962YAUPp2ft86FC5kMdic_npgrxnaoERIXO6Iw0GUI04KmXseVg4sBioo33Ezz_14u9WPB5TqBW4x6MmSJ1I514Zw3&attredirects=0");
  panel.add(image);    
  var spinner = app.createImage(AUTOCRATIMAGEURL).setWidth(150);
  spinner.setVisible(false);
  spinner.setStyleAttribute("position", "absolute");
  spinner.setStyleAttribute("top", "120px");
  spinner.setStyleAttribute("left", "200px");
  spinner.setId("dialogspinner");
  app.add(panel);
  app.add(spinner);
  var fileId = ScriptProperties.getProperty("fileId");
  if (fileId) {
    refreshTemplatePanel(fileId);
    var formatNote = app.createLabel("Note: If you intend to merge dates, times, or specialized number formats like percentages, it is highly recommended you take a moment to format those values as 'Plain text' or use the =TEXT() function to recalculate them in a new column.").setStyleAttribute("marginTop","10px");
    panel.add(formatNote);    
  } else {
    var chooseHandler = app.createServerHandler('showDocsPicker');
    var chooseButton = app.createButton("Choose template from Drive").addClickHandler(chooseHandler);
    panel.add(chooseButton);
    var tagLabel = app.createLabel("Don't have a template made? Go make one and come back to this step. Here are some suggested merge tags to use in your template Document or Spreadsheet.").setStyleAttribute("marginTop","10px");
    var tagScroll = app.createScrollPanel().setHeight("150px").setStyleAttribute("backgroundColor", "whiteSmoke");
    var tagGrid = app.createGrid(normalizedHeaders.length, 1);
    for (var i=0; i<normalizedHeaders.length; i++) {
      tagGrid.setWidget(i, 0, app.createLabel("<<" + normalizedHeaders[i] + ">>"));
    }
    panel.add(tagLabel);
    tagScroll.add(tagGrid);
    panel.add(tagScroll);
  }
  SpreadsheetApp.getActiveSpreadsheet().show(app);
}


function isNumber(n) {
  return !isNaN(parseFloat(n)) && isFinite(n);
}


function refreshTemplatePanel(fileId) {
    var app = UiApp.getActiveApplication();
    var panel = app.getElementById("panel");
    var spinner = app.getElementById("dialogspinner");
    panel.clear();
    var image = app.createImage("https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/download.png?attachauth=ANoY7cpTRdoqkEnOB2Godog9PVqEFovv__DxSqM02KnW3_FNFZZA0CneTMejWaPWWA01AOVfgHhAhYkUkBUfiZQbUm4DiYo82xjSIEYxokxpeHPWChhsI8TetLMKMn7V8inAH7HMCgogCUc5yEtVHUJkus-FbCBGvYmi9KKozeA9cFSu7Q962YAUPp2ft86FC5kMdic_npgrxnaoERIXO6Iw0GUI04KmXseVg4sBioo33Ezz_14u9WPB5TqBW4x6MmSJ1I514Zw3&attredirects=0");
    panel.setStyleAttribute('opacity','1');
    spinner.setVisible(false);
    var template = DocsList.getFileById(fileId);
    var name = template.getName();
    var url = template.getUrl();
    panel.add(app.createLabel("Currently selected merge template:").setStyleAttribute("marginTop", "10px").setStyleAttribute("width","100%").setStyleAttribute("backgroundColor", "#cfcfcf").setStyleAttribute("fontSize", "16px").setStyleAttribute("padding", "5px"));
    var anchor = app.createAnchor(name, url);
    panel.add(anchor);
    var docFields = autoCrat_fetchDocFields(fileId);
    var scrollpanel = app.createScrollPanel().setHeight(80).setWidth("100%").setStyleAttribute("backgroundColor", "whiteSmoke");
    var illegalTags = false;
    for (var j=0; j<docFields.length; j++) {
      var exp = new RegExp(/[0-9]/);
      var test = docFields[j].substr(2,1);
      if (isNumber(test)==true) {
        illegalTags = true;
      }
    }
    if ((docFields.length>0)&&(illegalTags==false)) {
      var fieldLabel = app.createLabel("Your template contains the following tags: ").setStyleAttribute("marginTop", "10px").setStyleAttribute("width","100%").setStyleAttribute("backgroundColor", "#cfcfcf").setStyleAttribute("fontSize", "16px").setStyleAttribute("padding", "5px");
      var list = app.createFlexTable();
      for (var i=0; i<docFields.length; i++) {
        list.setText(i, 0, docFields[i]);
      }
      panel.add(fieldLabel);
      scrollpanel.add(list);
      panel.add(scrollpanel);
    } 
    if (illegalTags==true) {
      var fieldLabel = app.createLabel("This template contains illegal merge tags. Tags must not start with numbers.").setStyleAttribute("color", "red").setStyleAttribute("marginTop", "10px").setStyleAttribute("width","100%").setStyleAttribute("backgroundColor", "#cfcfcf").setStyleAttribute("fontSize", "16px").setStyleAttribute("padding", "5px");
      panel.add(fieldLabel);
      panel.add(image);
    }
    if (docFields.length==0) {  
      var fieldLabel = app.createLabel("This template contains no merge tags. Tags must contain no special characters, must NOT start with numbers, and must use the style shown below:").setStyleAttribute("color", "red").setStyleAttribute("marginTop", "10px").setStyleAttribute("width","100%").setStyleAttribute("backgroundColor", "#cfcfcf").setStyleAttribute("fontSize", "16px").setStyleAttribute("padding", "5px");
      panel.add(fieldLabel);
      panel.add(image);
    }
    var chooseHandler = app.createServerHandler('showDocsPicker');
    var chooseButton = app.createButton("Choose a different template").addClickHandler(chooseHandler);
    panel.add(chooseButton);
    var nextHandler = app.createServerHandler('next');
    var nextButton = app.createButton("Save settings").addClickHandler(nextHandler);
    if (docFields.length==0) {
      nextButton.setEnabled(false);
    }
    panel.add(nextButton);
}

function showDocsPicker() {
  var app = UiApp.getActiveApplication();
  var panel = app.getElementById("panel");
  var spinner = app.getElementById("dialogspinner");
  var docsPicker = app.createDocsListDialog().addView(UiApp.FileType.DOCUMENTS).addView(UiApp.FileType.SPREADSHEETS).showDocsPicker();
  docsPicker.setWidth(475).setHeight(200);
  var spinnerHandler = app.createClientHandler().forTargets(spinner).setVisible(true).forTargets(panel).setStyleAttribute('opacity', '0.5');
  var selectionHandler = app.createServerHandler('saveDocID');
  docsPicker.addSelectionHandler(selectionHandler).addSelectionHandler(spinnerHandler);
  return app;
}

function saveDocID(e){
    var app = UiApp.getActiveApplication();
    var oldFileId = ScriptProperties.getProperty('fileId');
    var docTemplate = DocsList.getFileById(e.parameter.items[0].id);
    var fileId = docTemplate.getId();
    if ((oldFileId)&&(oldFileId!=fileId)) {
      ScriptProperties.setProperty('mappingString', '');
    }
    var fileName = docTemplate.getName();
    ScriptProperties.setProperty('fileId', fileId);
    ScriptProperties.setProperty('fileName', fileName);
    refreshTemplatePanel(fileId);
    return app;
}

function next() {
  var app = UiApp.getActiveApplication();
  autoCrat_initialize();
  if (!(ScriptProperties.getProperty('sheetName'))||(ScriptProperties.getProperty('sheetName'==''))) {
       autoCrat_defineSettings();
  }
  app.close();
  return app;
}

//responsible for Settings UI
//lots of trickery with client and server change handlers to reload conditional field values
function autoCrat_defineSettings() {
  var ss = SpreadsheetApp.getActive();
  var app = UiApp.createApplication();
  app.setTitle("Step 2: Select Merge Data Source").setHeight(100);
  var panel = app.createVerticalPanel().setId("settingsPanel");
  var sheetLabel = app.createLabel().setId("sheetLabel");
  sheetLabel.setText("Select the sheet in this spreadsheet that contains your merge data");
  
  var sheetListBox = app.createListBox().setName("sheetName").setId("sheetListBox");
  var sheets = ss.getSheets();
  for (i=0; i<sheets.length; i++) {
    var sheetName = sheets[i].getName();
    if (sheetName!="autoCrat Read Me") {
      sheetListBox.addItem(sheetName);
    }
  }

  panel.add(sheetLabel);
  panel.add(sheetListBox);
  
  var spinner = app.createImage(AUTOCRATIMAGEURL).setWidth(75);
  spinner.setVisible(false);
  spinner.setStyleAttribute("position", "absolute");
  spinner.setStyleAttribute("top", "10px");
  spinner.setStyleAttribute("left", "160px");
  spinner.setId("dialogspinner");
  
  var clickHandler = app.createServerHandler('autoCrat_saveSettings').addCallbackElement(panel);
  var spinnerHandler = app.createClientHandler().forTargets(spinner).setVisible(true).forTargets(panel).setStyleAttribute('opacity', '0.5');
  var button = app.createButton("Save").addClickHandler(clickHandler).addClickHandler(spinnerHandler);
   panel.add(button);
  
  app.add(panel);
  app.add(spinner);
  ss.show(app);
  return app;
}

//utility function for listboxes that require index be looked up for preselection of already-saved folders
function autoCrat_getFolderIndex(folderId) {
  var ss = SpreadsheetApp.getActive();
  var parent = DocsList.getFileById(ss.getId()).getParents()[0];
  if (!(parent)) {
    parent = DocsList.getRootFolder();
  }
  var folders = parent.getFolders();
  var indexFlag = 0;
  for (i=0; i<folders.length; i++) {
    if (folderId == folders[i].getId()) {
      indexFlag = i;
      break;
     }
   }
  return indexFlag; 
}


//utility function for listboxes that require index be looked up for preselection of already-saved source sheet
function autoCrat_getSheetIndex(sheetName) {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  var indexFlag = 1;
  for (i=0; i<sheets.length; i++) {
    if (sheetName ==sheets[i].getName()) {
      indexFlag = i;
      break;
     }
   }
  return indexFlag; 
}

//saves the settings from autoCrat_defineSettings function
function autoCrat_saveSettings(e) {
  var ss = SpreadsheetApp.getActive();
  var app = UiApp.getActiveApplication();
  var sheetName = e.parameter.sheetName;
  var sheet = ss.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var headers = autoCrat_fetchSheetHeaders(sheetName);
  if (data=='') {
    Browser.msgBox('Source sheet contains no data');
    autoCrat_defineSettings();
    return;
  }
   // Clear the mapping string property if the sheet name has been changed
   if (!(ScriptProperties.getProperty('sheetName')==sheetName)) {
     ScriptProperties.setProperty('mappingString','');
  }
  ScriptProperties.setProperty('sheetName', sheetName);
  var sheet = ss.getSheetByName(sheetName);
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var statusCol = headers.indexOf("Document Merge Status");
  var clearStatus = e.parameter.clearStatus;
  ScriptProperties.setProperty('clearStatus', true);
  if ((statusCol != -1)&&(clearStatus=="true")) {
    var range = sheet.getRange(2, statusCol+1, lastRow, 1);
    range.clear();
  }
  if (statusCol == -1) {
    var range = sheet.getRange(1, lastCol+1);
    range.setValue("Document Merge Status").setComment("Required by autoCrat script. Do not change the text of this header");
    range.setFontColor("white");
    range.setBackgroundColor("black");
  }
  autoCrat_initialize();
  if (!(ScriptProperties.getProperty('mappingString'))||(ScriptProperties.getProperty('mappingString'==''))) {
  autoCrat_setMergeConditions();
  }
  app.close();
  return app; 
}


//Used to clear the fields below the folder listbox if the value is changed.  Prevents the duplication of field elements in the form.
function autoCrat_reset(){
  var app = UiApp.getActiveApplication();
  app.getElementById('fileListBox').setVisible(false);
  app.getElementById('fileLabel').setVisible(false);
  app.getElementById('settingsButton').setVisible(false);
  return app;
}

//Completes the dropdown for files under the folder dropdown based on the folder selected
function autoCrat_getFiles(e) {
  var app = UiApp.getActiveApplication();
  var spinner = app.getElementById("dialogspinner");
  var panel = app.getElementById("settingsPanel");
  var fileLabel = app.createLabel().setId("fileLabel");
  fileLabel.setText("Select the template file you want to use");
  var fileListBox = app.createListBox().setName("fileName").setId("fileListBox").setWidth("300px");
  fileListBox.addItem('Select template file');
  var folderName = e.parameter.folderName;
  if (!folderName) {
    var tempFolderId = ScriptProperties.getProperty('folderId');
    folderName = DocsList.getFolderById(tempFolderId);
        }
        
    var folders = DocsList.getFolders();
    var indexFlag = 1;
    for (i=0; i < folders.length; i++) {
    if (folders[i].getName()==folderName) {
      indexFlag = i;
      break;
      }
    }  
   var folderId = folders[indexFlag].getId();
  
  var folder = DocsList.getFolderById(folderId);
  var files = folder.getFilesByType("document");
  for (i = 0; i < files.length; i++) {
    var fileName = files[i].getName();
    fileListBox.addItem(fileName);
  }
  app.getElementById('folderListBox').addChangeHandler(app.createServerClickHandler('autoCrat_reset'));
  var settingsSendClientHandler = app.createClientHandler().forTargets(spinner).setVisible(true).forTargets(panel).setStyleAttribute('opacity', "0.5");
  var settingsSendHandler = app.createServerHandler('autoCrat_saveSettings').addCallbackElement(panel);
  var button = app.createButton("Save settings", settingsSendHandler).setId('settingsButton').addClickHandler(settingsSendClientHandler);
  panel.add(fileLabel);
  panel.add(fileListBox);
  panel.add(button);
  spinner.setVisible(false);
  panel.setStyleAttribute("opacity","1");
  
  return app;
}
