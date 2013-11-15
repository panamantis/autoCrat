function autoCrat_preconfig() {
  // if you are interested in sharing your complete workflow system for others to copy (with script settings)
  // Select the "Generate preconfig()" option in the menu and
  //#######Paste preconfiguration code below before sharing your system for copy#######
  
  
  
  
  
  
  
  //#######End preconfiguration code#######
  autoCrat_logInstall()
  ScriptProperties.setProperty('preconfigStatus', 'true');
  var ssKey = SpreadsheetApp.getActiveSpreadsheet().getId();
  ScriptProperties.setProperty('ssKey', ssKey);
  autoCrat_initialize();
}
