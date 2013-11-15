var alphabet = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ"];
var FORMULACADDYIMAGEURL = 'https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/formulaCaddy_icon.gif?attachauth=ANoY7cq0pvQqsSiCo2XLcXxjrzNDmlWtNTjFIDM60Wz9CUdiFUYR6UiFB-CF81KKwD7T2EIjdA1JNd65Ndp-d_KypSbOqTv2QdduiiEIwLm3AuaH-iF6kjf5GK-ir7ew5UPWbqiAxl5cVjhvlXZaZGHNpKOb0I78JbuAVcPmoc8uMzAChZ_iHuS_7b6IN_IYF1VgeOWBnIjel6ZCWgGlyfIR65MWvv0bhs1ztQCZRYdQQj96D3ZdcCWugeiHtYCS13_cY-7VU4KT&attredirects=0';
var excludedHeaders = ['Merged Doc ID','Merged Doc URL','Link to merged Doc','Document Merge Status',"Case No"];


function autoCrat_waitForFormulaCaddy(ss) {
  var startTime = new Date();
  startTime = startTime.getTime();
  var formulaCaddySheet = ss.getSheetByName('formulaCaddyStatus');
  if (!formulaCaddySheet) {
    return;
  } else {
    var now = new Date();
    now = now.getTime();
    var caddyStatus = 0;
    while ((caddyStatus == 0)&&((now-startTime)<100000)) {
      caddyStatus = formulaCaddySheet.getRange('A2').getValue();
      Utilities.sleep(100);
      now = new Date();
      now = now.getTime();
    }
  } 
}


function copyDownFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName('formulaCaddyStatus').getRange('B1').setValue(0);
  var properties = ScriptProperties.getProperties();
  var sheet = getSheetById(ss, properties.formSheetId);
  var thisRow = sheet.getActiveRange().getRow();
  var colValues = Utilities.jsonParse(properties.colValues);
  var formulaRow = parseInt(properties.formulaRow);
  var objArray = [];
  for (var key in colValues) {
    var colNum = key.split("-")[1];
    objArray.push({colNum: colNum, type: colValues[key], formula: properties['formula-'+colNum]});
  }
  objArray.sort(function(a,b){return a.colNum-b.colNum});
  for (var i=0; i<objArray.length; i++) {
    var colNum = objArray[i].colNum;
    var type = objArray[i].type;
    var formula = objArray[i].formula;
    var copyCell = sheet.getRange(formulaRow, colNum).setFormula(formula);
    var destCell = sheet.getRange(thisRow, colNum);
    copyCell.copyTo(destCell);
    if (type == 1) {
      var newValue = destCell.getValue();
      destCell.setValue(newValue);
    }
    SpreadsheetApp.flush();
  }
  ss.getSheetByName('formulaCaddyStatus').getRange('B1').setValue(1);
  SpreadsheetApp.flush();
}


function checkCreateSheet(ss) {
  var formulaCaddySheet = ss.getSheetByName('formulaCaddyStatus');
  var topSheet = ss.getSheets()[0];
  if (!formulaCaddySheet) {
    formulaCaddySheet = ss.insertSheet('formulaCaddyStatus');
    formulaCaddySheet.getRange("A1:B1").setValues([["Status","1"]]);
    formulaCaddySheet.getRange("A2:B2").setValues([["Trigger Set", "0"]]);
    formulaCaddySheet.deleteColumns(3, formulaCaddySheet.getMaxColumns()-3);
    formulaCaddySheet.deleteRows(3, formulaCaddySheet.getMaxRows()-3);
    ss.setActiveSheet(topSheet);
    formulaCaddySheet.hideSheet();
  }
  SpreadsheetApp.flush();
  return formulaCaddySheet;
}


function checkSetTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formulaCaddySheet = checkCreateSheet(ss);
  var triggerSet = formulaCaddySheet.getRange("B2").getValue();
  var ssId = ss.getId(); 
  if (triggerSet!=1) {
    setCopyDownTrigger();
  }
  return;
}



function setCopyDownTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  var formulaCaddySheet = checkCreateSheet(ss);
  var triggers = ScriptApp.getProjectTriggers();
  var found = false;
  for (var i=0; i<triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == "copyDownFormulas") {
      found = true;
      break;
    }
  }
  if (!found) {
    ScriptApp.newTrigger('copyDownFormulas').forSpreadsheet(ssId).onFormSubmit().create();
    formulaCaddySheet.getRange("B2").setValue("1");
    ScriptProperties.setProperty('formulaTriggerSet','true');
  }
  return;
}



function detectFormSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  ScriptProperties.setProperty('ssId', ssId);
  var formUrl = ss.getFormUrl();
  if (!formUrl) {
    Browser.msgBox("This feature only works on Spreadsheets with attached forms");
    return;
  }
  var form = FormApp.openByUrl(formUrl);
  var items = form.getItems();
  var sheets = ss.getSheets();
  for (var i=0; i<sheets.length; i++) {
    var thisTopLeftCell = sheets[i].getRange(1, 1);
    var thisTopLeftBgColor = thisTopLeftCell.getBackground();
    var thisTopLeftValue = thisTopLeftCell.getValue();
    var found = false;
    if ((thisTopLeftBgColor == "#DDDDDD")&&(thisTopLeftValue == "Timestamp")) {
      var formSheetId = sheets[i].getSheetId();
      found = true;
      break;
    }
  }
  if (found) {
    ScriptProperties.setProperty('formSheetId', formSheetId);
  } else {
    var error = catchNoFormSheetDetected(ss);
  }
  if ((!error)||(error!=false)) {
    formulaCaddy_createJob();
  }
}


function catchNoFormSheetDetected(ss) {
  var formSheetName = Browser.inputBox('Unable to detect form sheet, please enter the name of the sheet that holds your form responses');
  try {
    var formSheetId = ss.getSheetByName(formSheetName).getSheetId();
  } catch(err) {
    if (formSheetName!="cancel") {
      Browser.msgBox('Unable to find sheet: ' + formSheetName);
      catchNoFormSheetDetected(ss);
    } else {
      return 'error';
    }
  }
  ScriptProperties.setProperty('formSheetId', formSheetId);
  return;
}


function getSheetById(ss, sheetId) {
  var sheets = ss.getSheets();
  for (var i=0; i<sheets.length; i++) {
    if (sheets[i].getSheetId() == sheetId) {
      return sheets[i];
    }
  }
  return;
}


function formulaCaddy_createJob() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var sheets = ss.getSheets();
  var sheetNames = [];
  for (var i=0; i<sheets.length; i++) {
    sheetNames.push(sheets[i].getName());
  }
  var formSheet = getSheetById(ss, properties.formSheetId);
  var activeSheetIndex = formSheet.getIndex();
  var activeSheetName = formSheet.getName();
  var activeSheetSelectIndex = sheetNames.indexOf(activeSheetName);
  var app = UiApp.createApplication().setTitle("Copy down formulas within form response sheet").setHeight(400).setWidth(700);
  var panel = app.createVerticalPanel().setId("panel").setWidth("680px").setHeight("200px");
  var help = "This feature allows you to use additional columns in the form response sheet to calculate or look up values upon form submission. ";
  help += "Only columns that are not part of your form will be available for selection below. Selecting the \"Paste as values?\" checkbox will ";
  help += "paste the calculated values into the form submission row, without formulas."  
  help += "<br/> Be aware that any changes you make to the column order or structure of the form responses sheet will require you to redo these settings."; 
  var helpHtml = app.createHTML(help).setStyleAttribute('marginBottom', '5px');
  panel.add(helpHtml);
  var columns = formSheet.getLastColumn();
  var grid = app.createGrid(4, columns+1).setId('grid');
  var spinner = app.createImage(FORMULACADDYIMAGEURL).setWidth("115px").setId('spinner');
  spinner.setVisible(false);
  spinner.setStyleAttribute("position", "absolute");
  spinner.setStyleAttribute("top", "120px");
  spinner.setStyleAttribute("left", "220px");
  var refreshOpacityHandler = app.createClientHandler().forTargets(panel).setStyleAttribute('opacity', '0.5').forTargets(spinner).setVisible(true);
  var sheetId = formSheet.getSheetId();
  var hiddenSheetIdBox = app.createTextBox().setValue(sheetId).setId('sheetId').setName('sheetId').setVisible(false);  
  var hiddenColNumBox = app.createTextBox().setValue(columns).setId('numCols').setName('numCols').setVisible(false);
  panel.add(hiddenSheetIdBox);
  panel.add(hiddenColNumBox);
  var noDataLabel = app.createLabel("No data in sheet").setId('noDataLabel');
  noDataLabel.setVisible(false);
  panel.add(noDataLabel);
  panel.add(grid);
  app.add(panel);
  app.add(spinner);
  formulaCaddy_returnSheetUi(formSheet, properties);
  ss.show(app);        
}


function formulaCaddy_returnSheetUi(sheet, properties) {
  var app = UiApp.getActiveApplication();
  var panel = app.getElementById('panel');
  panel.setStyleAttribute('opacity','1');
  var spinner = app.getElementById('spinner');
  spinner.setVisible(false);
  var scrollPanel = app.createScrollPanel().setWidth("660px").setStyleAttribute('margin', '10px');
  var grid = app.getElementById('grid');
  var sheetId = sheet.getSheetId();
  var sheetProperties = properties;
  var hiddenSheetIdBox = app.getElementById('sheetId').setValue(sheetId);
  var columns = sheet.getLastColumn();
  var hiddenColNumBox = app.getElementById('numCols').setValue(columns);
  if (sheetProperties['formulaRow']) {
    var row = parseInt(sheetProperties['formulaRow']);
  } else {
    var row = 2;
  }
  var rowsArray = [2,3,4];
  var noDataLabel = app.getElementById('noDataLabel');
  if (columns>0) {
    var colValues = sheetProperties['colValues'];
    if (colValues) {
      colValues = Utilities.jsonParse(colValues);
    } else {
      colValues = new Object();
    }
    noDataLabel.setVisible(false);
    grid.resize(4, columns+1);
    grid.setVisible(true);
    var horizontalPanel = app.createHorizontalPanel().setId('horizPanel');
    var formulaRowLabel = app.createLabel("Row containing master values/formulas");
    var formulaRowSelect = app.createListBox().setId('formulaRowSelect').setName("formulaRow");
    var rowSelectHandler = app.createServerHandler('formulaCaddy_refreshRowFormulas').addCallbackElement(panel);
    var lastRow = sheet.getLastRow();
    for (var i=2; i<lastRow+1; i++) {
      formulaRowSelect.addItem(i);
    }
    var rowSelectIndex = rowsArray.indexOf(row);
    formulaRowSelect.setSelectedIndex(rowSelectIndex);
    grid.setBorderWidth(1).setCellSpacing(0).setStyleAttribute('borderColor','#E5E5E5').setStyleAttribute('opacity', '1');
    formulaRowSelect.addChangeHandler(rowSelectHandler);
    horizontalPanel.add(formulaRowLabel);
    horizontalPanel.add(formulaRowSelect);
    horizontalPanel.add(formulaRowLabel);
    panel.add(horizontalPanel);
   
    var headerRange = sheet.getRange(1,1,1,columns);
    var headers = headerRange.getValues()[0];
    var headerBgs = headerRange.getBackgroundColors()[0];
    grid.setWidget(0, 0, app.createLabel("Column"));
    grid.setWidget(1, 0, app.createLabel("Header"));  
    grid.setWidget(2, 0, app.createLabel("Value/Formula"));  
    grid.setWidget(3, 0, app.createLabel("Paste as values?"));  
    var onButtonHandlers = [];
    var onButtonServerHandlers = [];
    var offButtonHandlers = [];
    var offButtonServerHandlers = [];
    var onButtons = [];
    var offButtons = [];
    var buttonValues = [];
    var formulaLabels = [];
    var asValuesCheckBoxes = [];
    for (var i=0; i<columns; i++) {
      onButtons[i] = app.createButton(this.alphabet[i]).setId('onButton-'+sheetId+'-'+i).setStyleAttribute('background', 'whiteSmoke').setWidth("50px");
      offButtons[i] = app.createButton(this.alphabet[i]).setId('offButton-'+sheetId+'-'+i).setStyleAttribute('background', '#E5E5E5').setStyleAttribute('border', '2px solid grey').setWidth("50px").setVisible(false);
      buttonValues[i] = app.createTextBox().setVisible(false).setText(i+'-off').setName('bv-'+sheetId+'-'+i);
      var buttonPanel = app.createHorizontalPanel();
      buttonPanel.add(onButtons[i])
                 .add(offButtons[i])
                 .add(buttonValues[i])
                 .setStyleAttribute('width',"80px")
                 .setHorizontalAlignment(UiApp.HorizontalAlignment.CENTER);
      var formulas = formulaCaddy_getSheetFormulas(sheet, row);
      var formulaLabel = app.createLabel(formulas[i]).setId('formula-'+i).setStyleAttribute('opacity', '0.5'); 
      asValuesCheckBoxes[i] = app.createCheckBox().setId('asValues-'+sheetId+'-'+i).setName('av-'+sheetId+'-'+i).setEnabled(false).setValue(false);
      grid.setWidget(0, i+1, buttonPanel).setStyleAttribute(0, i+1, 'backgroundColor', 'whiteSmoke').setStyleAttribute(0, i+1, 'textAlign', 'center');
      grid.setWidget(1, i+1, app.createLabel(headers[i]));
      grid.setWidget(2, i+1, formulaLabel);   
      grid.setWidget(3, i+1, asValuesCheckBoxes[i]);
      if ((headerBgs[i] == "#DDDDDD")||(excludedHeaders.indexOf(headers[i])!=-1)) {
        grid.setStyleAttribute(1, i+1, 'backgroundColor', '#DDDDDD');
        onButtons[i].setEnabled(false).setVisible(false);
        offButtons[i].setEnabled(false).setVisible(false);
        asValuesCheckBoxes[i].setEnabled(false).setVisible(false);
      } else { 
        onButtonHandlers[i] = app.createClientHandler().forTargets(onButtons[i]).setVisible(false).forTargets(offButtons[i]).setVisible(true).forTargets(asValuesCheckBoxes[i]).setEnabled(true).forTargets(buttonValues[i]).setText(i+'-on');
        onButtonServerHandlers[i] = app.createServerHandler('toggleOpacity').addCallbackElement(panel);
        offButtonHandlers[i] = app.createClientHandler().forTargets(offButtons[i]).setVisible(false).forTargets(onButtons[i]).setVisible(true).forTargets(asValuesCheckBoxes[i]).setEnabled(false).setValue(false).forTargets(buttonValues[i]).setText(i+'-off');
        offButtonServerHandlers[i] = app.createServerHandler('toggleOpacity').addCallbackElement(panel);
        onButtons[i].addClickHandler(onButtonHandlers[i]).addClickHandler(onButtonServerHandlers[i]);
        offButtons[i].addClickHandler(offButtonHandlers[i]).addClickHandler(offButtonServerHandlers[i]);
      }
      if (colValues['col-'+(i+1)]) {
        onButtons[i].setVisible(false);
        offButtons[i].setVisible(true);
        buttonValues[i].setText(i+'-on');
        formulaLabel.setStyleAttribute('opacity','1');
        asValuesCheckBoxes[i].setEnabled(true);
        if (colValues['col-'+(i+1)]=='1') {
          asValuesCheckBoxes[i].setValue(true);
        }
      }
    }
    scrollPanel.add(grid);
    panel.add(scrollPanel);
    var saveHandler = app.createServerHandler('manualSave').addCallbackElement(panel);
    var saveClientHandler = app.createClientHandler().forTargets(spinner).setVisible(true).forTargets(panel).setStyleAttribute('opacity', '0.5');
    var button = app.createButton("Save settings").setId('button').addClickHandler(saveHandler).addClickHandler(saveClientHandler);
    panel.add(button);
  } else {
    noDataLabel.setVisible(true);
    var saveHandler = app.createServerHandler('manualSave').addCallbackElement(panel);
    var button = app.createButton("Save settings").setId('button').addClickHandler(saveHandler).setVisible(false);
    panel.add(button);
  }
  return app;
}


function manualSave(e) {
  var app = UiApp.getActiveApplication();
  saveformulaCaddySettings(e);
  app.close();
  return app;
}



function waitingIcon() {
  var app = UiApp.createApplication().setHeight(250).setWidth(200);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var waitingImageUrl = this.formulaCadddyIMAGEURL;
  var image = app.createImage(waitingImageUrl).setWidth("125px").setStyleAttribute('marginLeft', '25px');
  app.add(image);
  app.add(app.createLabel('Please be patient as formulaCaddy formulas are recalculated and pasted down their designated columns. For complex spreadsheets this can take some time.'));
  ss.show(app);
  return app;
}

function closeIcon(e) {
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}


function formulaCaddy_getSheetFormulas(sheet, row) {
  var columns = sheet.getLastColumn();
  var range = sheet.getRange(row, 1, 1, columns)
  var formulas = range.getFormulas()[0];
  var values = range.getValues()[0];
  for (var i=0; i<formulas.length; i++) {
    if (formulas[i]=='') {
      if (typeof values[i] == 'string') {
        formulas[i]=values[i].substring(0,15);
      }  else {
        formulas[i]=values[i];
      }
    } else {
      formulas[i] = formulas[i].substring(0,15);
    }
    if (formulas[i].length==15) {
      formulas[i] += "...";
    }
  }
  return formulas;
}
  

function formulaCaddy_refreshRowFormulas(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetProperties = ScriptProperties.getProperties();
  var sheetId = e.parameter.sheetId;
  var colValues = sheetProperties['colValues'];
  if (colValues) {
    colValues = Utilities.jsonParse(colValues);
  } else {
    colValues = new Object();
  }
  var app = UiApp.getActiveApplication();
  var grid = app.getElementById('grid');
  var row = e.parameter["formulaRow"];
  var sheet = getSheetById(ss, sheetId);
  var formulas = formulaCaddy_getSheetFormulas(sheet, row);
  
  for (var i=0; i<formulas.length; i++) {
    var formulaLabel = app.createLabel(formulas[i]).setId('formula-'+i).setStyleAttribute('opacity', '0.35');
    if (colValues['col-'+(i+1)]) {
      formulaLabel.setStyleAttribute('opacity', '1');
    }
    grid.setWidget(2, i+1, formulaLabel); 
  }
  return app;
}



function toggleOpacity(e) {
  var app = UiApp.getActiveApplication();
  var sheetId = e.parameter.sheetId;
  var numCols = e.parameter.numCols;
  for (var i=0; i<numCols; i++ ) {
  var buttonValue = e.parameter['bv-'+sheetId+'-'+i];
  if (buttonValue) {
    buttonValue = buttonValue.split("-");
  }
  var label = app.getElementById('formula-'+buttonValue[0]);
  if (buttonValue[1] == "on") {
    label.setStyleAttribute('opacity','1');
  } else {
    label.setStyleAttribute('opacity','0.5');
  }
}
  return app;
}




function saveformulaCaddySettings(e) {
  var app = UiApp.getActiveApplication();
  var properties = ScriptProperties.getProperties();
  var sheetId = e.parameter.sheetId;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSheetById(ss, sheetId);
  var numCols = e.parameter.numCols;
  var formulaRow = e.parameter['formulaRow']
  properties["formulaRow"] = formulaRow;
  var colValues = {};
    for (var j=0; j<numCols; j++) {
      var buttonValue = e.parameter['bv-'+sheetId+'-'+j];
      if (buttonValue==j+'-on') {
        var asValuesOption = e.parameter['av-'+sheetId+'-'+j];
        if (asValuesOption == "false") {
          colValues["col-" + (j+1)] = 2;
        } else {
          colValues["col-" + (j+1)] = 1; 
        }
        properties["formula-" + (j+1)] = sheet.getRange(formulaRow, j+1).getFormula().toString();
      } 
    }
  properties.colValues = Utilities.jsonStringify(colValues);
  ScriptProperties.setProperties(properties);
  checkSetTrigger();
  app.close();
  return app;
}
