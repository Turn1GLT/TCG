// F2FLeagueOnOpen

// **********************************************
// function OnOpen()
//
// Function executed everytime the Spreadsheet is
// opened or refreshed
//
// **********************************************

function OnOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var FirstSht =  ss.getSheets()[0];
  
  ss.setActiveSheet(FirstSht);
  
  var FuncMenuButtons = [{name: 'Analyze New Match Entry', functionName: 'fcnGameResults'}, {name: 'Generate Players Card DB', functionName:'fcnGenPlayerCardDB'}, {name:'Delete Players Card DB', functionName:'fcnDelPlayerCardDB'}, {name:'Generate Players Card Pool', functionName:'fcnGenPlayerCardPoolSht'}, {name:'Delete Players Card Pool', functionName:'fcnDelPlayerCardPoolSht'}];
  
  ss.addMenu("General Fctn", FuncMenuButtons);
  }