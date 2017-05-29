// F2FLeagueOnOpen

// **********************************************
// function OnOpen()
//
// Function executed everytime the Spreadsheet is
// opened or refreshed
//
// **********************************************

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shtFirst =  ss.getSheets()[0];
  
  ss.setActiveSheet(shtFirst);
  
  var FuncMenuButtons  = [{name: 'Analyze New Match Entry', functionName: 'fcnGameResults'}];
  var StartMenuButtons = [{name: 'Generate Players Card DB', functionName:'fcnGenPlayerCardDB'}, {name:'Delete Players Card DB', functionName:'fcnDelPlayerCardDB'}, {name:'Generate Players Card Pool', functionName:'fcnGenPlayerCardPoolSht'}, {name:'Delete Players Card Pool', functionName:'fcnDelPlayerCardPoolSht'}];
  
  ss.addMenu("General Fctn", FuncMenuButtons);
  ss.addMenu("League Fctn", StartMenuButtons);
}


// **********************************************
// function fcnWeekChange()
//
// When the Week number changes, this function analyzes all
// generates a weekly report 
//
// **********************************************

function onWeekChange(){

  // Opens Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shtCumul = ss.getSheetByName('Cumulative Results');
  var Week = shtCumul.getRange(2,3).getValue();
  var LastWeek = Week - 1;
  var Players = new Array(32);
  
  var EmailSubject = 'Week ' + LastWeek + ' Report';
  var EmailMessage = 'Week ' + LastWeek + ' is now complete and Week '+ Week +' has started. \n\nHere is the week report for Week ' + LastWeek + '.\n\nInsert Report here...\n\n';
  
  // Analyze
  //Players = fcnAnalyzeLossPenalty(ss);
  
  MailApp.sendEmail("gamingleaguemanager@gmail.com", EmailSubject, EmailMessage);
  
}