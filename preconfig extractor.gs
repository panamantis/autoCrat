function autoCrat_getProperties() {
  var properties =  new Object();
  var formTrigger = ScriptProperties.getProperty('formTrigger');
  var fileType = ScriptProperties.getProperty('fileType');
  var emailAttachment = ScriptProperties.getProperty('emailAttachment');
  var fileNameString = ScriptProperties.getProperty('fileNameString');
  var destinationFolderName = ScriptProperties.getProperty('destinationFolderName');
  var emailSubject = ScriptProperties.getProperty('emailSubject');
  var fileId = ScriptProperties.getProperty('fileId');
  var fileName = ScriptProperties.getProperty('fileName');
  var emailSetting = ScriptProperties.getProperty('emailSetting');
  var bodyPrefix = ScriptProperties.getProperty('bodyPrefix');
  var mergeConditions = ScriptProperties.getProperty('mergeConditions');
  var linkToDoc = ScriptProperties.getProperty('linkToDoc');
  var mappingString = ScriptProperties.getProperty('mappingString');
  var clearStatus = ScriptProperties.getProperty('clearStatus');
  var fileSetting = ScriptProperties.getProperty('fileSetting');
  var sheetName = ScriptProperties.getProperty('sheetName');
  var emailString = ScriptProperties.getProperty('emailString');
  var secondaryFolderToken = ScriptProperties.getProperty('secondaryFolderToken');
  
  
  if (formTrigger) { properties.formTrigger = formTrigger; }
  if (fileType) { properties.fileType = fileType; }
  if (emailAttachment) {properties.emailAttachment = emailAttachment;}
  if (fileNameString) {properties.fileNameString = fileNameString; }
  if (destinationFolderName) {properties.destinationFolderName = destinationFolderName; }
  if (secondaryFolderToken) {properties.secondaryFolderId = secondaryFolderToken; properties.secondaryFolderToken = secondaryFolderToken; }
  if (emailSubject) {properties.emailSubject = emailSubject; }
  if (fileId) {properties.fileId = fileId; }
  if (fileName) {properties.fileName = fileName; }
  if (emailSetting) {properties.emailSetting = emailSetting; }
  if (bodyPrefix) {properties.bodyPrefix = bodyPrefix; }
  if (mergeConditions) {properties.mergeConditions = mergeConditions; }
  if (linkToDoc) {properties.linkToDoc = linkToDoc; }
  if (mappingString) {properties.mappingString = mappingString; }
  if (clearStatus) {properties.clearStatus = clearStatus; }
  if (fileSetting) {properties.fileSetting = fileSetting; }
  if (sheetName) {properties.sheetName = sheetName; }
  if (emailString) {properties.emailString = emailString; }
  return properties;
}


function autoCrat_extractorWindow () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = autoCrat_getProperties();
  var app = UiApp.createApplication().setHeight(500).setWidth(600).setTitle("Export preconfig() settings");
  var panel = app.createVerticalPanel().setWidth("100%").setHeight("100%");
  var labelText = "Copying a Google Spreadsheet copies scripts along with it, but without any of the script settings saved.  This normally makes it hard to share full, script-enabled Spreadsheet systems. ";
  labelText += " You can solve this problem by pasting the code below into a script file called \"paste preconfig here\" (go to Script Editor and look in left sidebar of the autoCrat script) prior to publishing your Spreadsheet for others to copy. \n";
  labelText += " After a user copies your spreadsheet, they will select \"Run initial installation.\"  This will preconfigure all needed script settings.  If you copied this system from someone as a spreadsheet, this has probably already been done for you.";
  var label = app.createLabel(labelText);
  var window = app.createTextArea();
  var codeString = "//This section sets all script properties associated with this autoCrat profile \n";
  codeString += "var preconfigStatus = ScriptProperties.getProperty('preconfigStatus');\n";
  codeString += "if (preconfigStatus!='true') {\n";
  for (var propertyKey in properties) {
    if (propertyKey != "fileId") {
      var propertyVal = properties[propertyKey];
      codeString += "  ScriptProperties.setProperty('" + propertyKey + "', '" + propertyVal + "');\n";
    }
  }
  
 //generate msgbox warning code if automated email or calendar is enabled in template  
    codeString += "\n \n";
    codeString += "  //Custom code to copy document template and create destination folder \n";
    codeString += "  //Note that your template must be set as visible to the user that will be copying this system \n";
    codeString += "  var ss = SpreadsheetApp.getActiveSpreadsheet();\n";
    codeString += "  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();\n";
    codeString += "  var parent = DocsList.getFileById(ssId).getParents()[0];\n";
    codeString += "  var files = parent.getFiles();\n";
    codeString += "  for (var i=0; i<files.length; i++) {\n";
    codeString += "    if (files[i].getName()=='"+properties.fileName+"') {\n";
    codeString += "      var fileId = files[i].getId();\n";
    codeString += "      ScriptProperties.setProperty('fileId',fileId);\n";
    codeString += "      break;\n";
    codeString += "    }\n";
    codeString += "  }\n";
    codeString += "  var newFolder = parent.createFolder('"+ properties.destinationFolderName+"');\n";
    codeString += "  var newFolderId = newFolder.getId();\n";
    codeString += "  ScriptProperties.setProperty('destinationFolderId',newFolderId);\n";
    codeString += "  if (!(fileId)) {\n";
    codeString += "    var originalTemplate = DocsList.getFileById('"+properties.fileId+"');\n";
    codeString += "    var copy = originalTemplate.makeCopy('"+properties.fileName+"')\n";
    codeString += "    var root = copy.getParents()[0];\n";
    codeString += "    copy.removeFromFolder(root);\n";
    codeString += "    copy.addToFolder(parent);\n";
    codeString += "    var copyId = copy.getId();\n";
    codeString += "    ScriptProperties.setProperty('fileId',copyId);\n";
    codeString += "  }\n";
    codeString += "    ss.toast('Custom autoCrat preconfiguration ran successfully.');\n";
    codeString += "}\n";
  window.setText(codeString).setWidth("100%").setHeight("400px");
  app.add(label);
  panel.add(window);
  app.add(panel);
  ss.show(app);
  return app;
}
