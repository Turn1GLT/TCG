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
  var StartMenuButtons = [{name: 'Generate Players Card DB', functionName:'fcnGenPlayerCardDB'}, {name:'Generate Players Card Pool', functionName:'fcnGenPlayerCardPoolSht'}, {name:'Delete Players Card DB', functionName:'fcnDelPlayerCardDB'}, {name:'Delete Players Card Pool', functionName:'fcnDelPlayerCardPoolSht'}];
  
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
  var PenaltyTable;
  
  // Players Array to return Penalty Losses
  var PlayerData = new Array(32); // 0= Player Name, 1= Penalty Losses
  for(var plyr = 0; plyr < 32; plyr++){
    PlayerData[plyr] = new Array(2); 
    for (var val = 0; val < 2; val++) PlayerData[plyr][val] = '';
  }
  
  // Analyze
  PlayerData = fcnAnalyzeLossPenalty(ss, Week, PlayerData);
  
  for(var row = 0; row<32; row++){
    if (PlayerData[row][0] != '') Logger.log('Player: %s - Missing: %s',PlayerData[row][0], PlayerData[row][1]);
  }
  
  PenaltyTable = subPlayerPenaltyTable(PlayerData);
  

  var EmailSubject = 'Week ' + LastWeek + ' Report';
  var EmailMessage = 'Week ' + LastWeek + ' is now complete and Week '+ Week +' has started. <br><br>Here is the week report for Week ' + LastWeek + '.<br><br>Insert Report here...<br><br>';
  EmailMessage += PenaltyTable;

  
  MailApp.sendEmail('gamingleaguemanager@gmail.com', EmailSubject, EmailMessage,{name:'TCG Booster League Manager',htmlBody:EmailMessage});
  
}