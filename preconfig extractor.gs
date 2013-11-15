function autoCrat_extractorWindow () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  var properties = ScriptProperties.getProperties();
  var excludeProperties = ['autocrat_sid', 'preconfigStatus', 'ssId', 'ssKey', 'destinationFolderId', 'fileId', 'formulaTriggerSet'];
  var propertyString = '';
  for (var key in properties) {
    if (excludeProperties.indexOf(key)==-1) {
      var keyProperty = properties[key]; //.replace(/[/\\*]/g, "\\\\");                                     
      propertyString += "   ScriptProperties.setProperty('" + key + "','" + keyProperty + "');\n";
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
  if (properties.formulaTriggerSet == "true") {
    codeString += "   setCopyDownTrigger(); \n";
  }
  if (properties.formTrigger == "true") {
    codeString += "   autoCrat_setFormTrigger(); \n";
  }
  codeString += propertyString;
  codeString += "    ss.toast('Custom autoCrat preconfiguration ran successfully.');\n";
  codeString += "}\n";
  window.setText(codeString).setWidth("100%").setHeight("400px");
  app.add(label);
  panel.add(window);
  app.add(panel);
  ss.show(app);
  return app;
}
