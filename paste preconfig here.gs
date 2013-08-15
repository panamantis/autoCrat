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


//Function used to create contents of "Read Me" sheet
function autoCrat_setReadMeText() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("autoCrat Read Me");
  sheet.insertImage('http://www.youpd.org/sites/default/files/acquia_commons_logo36.png', 1, 1);
  sheet.setRowHeight(1, 100);
  sheet.setColumnWidth(1, 700);
  var readMeText = "Installation and configuration steps...looks hard, but it's really not bad, I promise.  Almost everything you need to do with this script is done through a custom GUI!";
  readMeText += "\n \n In order for the script to work, you will need to create a Google Document or Spreadsheet to use as a template for the merge.";
  readMeText += "\n \n First, create a template Document. Include double bracketed tags for any personalized data you want to populate from the spreadsheet.  Ex) <<First Name>> .  IMPORTANT: Do not use any non-alphanumeric characters in your tags. Other than that, it doesn't matter what you call the fields, because you will map them to your spreadsheet headers in a few steps. The benefit of spelling them the same as your data column headers: they will automatically map to the correct column in your source data.";
  readMeText += "\n \n If you've installed and authorized the script, you should see a new menu item to the right of \"Help\", called \"Document Merge.\"  If you don't see it, try running the onOpen function from the script editor (Tools->Script Editor->Run->onOpen.";
  readMeText += "\n \n In the \"Document Merge\" menu, select \"Select data source and template doc\" and complete the settings.  If you don't have a collection that contains a template file with <<Merge tags>> in it, go back and do this first.";
  readMeText += "\n \n In Step 1: You will first be prompted to choose your merge template from Drive.  A template must be Google Document type file, with <<Merge tags>> that you will use to match to your column headings.";
  readMeText += "\n \n In Step 2: you will be asked to choose which sheet contains your source data for the merge.";
  readMeText += "\n Step 3 will prompt you to \"Set Merge Conditions,\" which means you have the option to require a match to a value in a particular field of your source data before a given row will be merged.  Leaving this setting blank will cause it to be ignored.";
  readMeText += "\n \n Step 4, \"Set Field Mappings\", will ask you to map each <<Merge tag>> to the spreadsheet column you want to use to populate it.  Save the mappings.";
  readMeText += "\n \n In Step 5, \"Set Merge Type\", decide what type of merge you want to try...there are a number of combos and cool possibilities.  Look to the bottom of the panel for a clue as to the $variableNames that are available for any of the fields you want to populate dynamically per row.";
  readMeText += "\n \n Here are some basic options to play with.  Checkbox allows you to test on first-row only if you like.";
  readMeText += "\n \n * ONLY saving merged Docs to a collection, either as PDF or Doc format."
  readMeText += "\n \n * Saving to a collection AND emailing PDF as attachment.";
  readMeText += "\n \n * Saving to a collection AND emailing recipient a link to individually shared Docs as View-only";
  readMeText += "\n \n * Saving to a collection AND emailing recipient a link to individually shared Doc as Editor";
  readMeText += "\n \n For date formatting to be handled in a merge field, you must use the Format->Number menu from the spreadsheet to format any dates.  Currently only three formats are supported:   \"M/d/yyyy\", \"MMMM d, yyyy\", and \"M/d/yyyy H:mm:ss\" ... i.e. 1/30/2012...January 30, 2012, and 1/30/2012 9:32:34.";
  readMeText += "\n \n The document Header and Footer can also contain merge tags!"; 
  readMeText += "\n \n In Step 6, you can run the merge, test on only the first row, or simply save your settings for later.";
  sheet.getRange("A2").setValue(this.scriptTitle).setFontSize(18);
  sheet.getRange("A3").setValue("Support available at http://www.youpd.org/autocrat");
  sheet.getRange("A4").setValue(readMeText);
}
