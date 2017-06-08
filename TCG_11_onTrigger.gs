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
  
  var FuncMenuButtons  = [{name: 'Analyze New Match Entry', functionName: 'fcnMain'}];
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

  // Open Configuration Spreadsheet
  var shtConfig = SpreadsheetApp.openById('1oXXEjOF9EoVxnR8pcmeNBSqJ1V-nPqPYNDwOnHWwznA').getSheetByName('Config');
  
  // League Name EN
  var Location = shtConfig.getRange(11,2).getValue();
  var LeagueTypeEN = shtConfig.getRange(12,2).getValue();
  var LeagueNameEN = shtConfig.getRange(3,2).getValue() + ' ' + LeagueTypeEN;
  
  // League Name FR
  var Location = shtConfig.getRange(11,2).getValue();
  var LeagueTypeFR = shtConfig.getRange(13,2).getValue();
  var LeagueNameFR = LeagueTypeFR + ' ' + shtConfig.getRange(3,2).getValue();
  
  // Open Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shtCumul = ss.getSheetByName('Cumulative Results');
  var Week = shtCumul.getRange(2,3).getValue();
  var LastWeek = Week - 1;
  var WeekName = 'Week'+Week;
  var shtWeek = ss.getSheetByName(WeekName);
  var PenaltyTable;
  var EmailSubject;
  var EmailMessage;
  var MatchPlyd;
  var MatchesPlayed = 0;
  
  // Players Array to return Penalty Losses
  var PlayerData = new Array(32); // 0= Player Name, 1= Penalty Losses
  for(var plyr = 0; plyr < 32; plyr++){
    PlayerData[plyr] = new Array(2); 
    for (var val = 0; val < 2; val++) PlayerData[plyr][val] = '';
  }
  
  // Get Amount of matches played this week.
  MatchPlyd = shtWeek.getRange(5, 5, 32, 1).getValues();
  for(var plyr=0; plyr<32; plyr++){
    if(MatchPlyd[plyr][0] > 0) MatchesPlayed = MatchesPlayed + MatchPlyd[plyr][0];
  }

  // Analyze if Players have missing matches to apply Loss Penalties
  PlayerData = fcnAnalyzeLossPenalty(ss, Week, PlayerData);
  
  for(var row = 0; row<32; row++){
    if (PlayerData[row][0] != '') Logger.log('Player: %s - Missing: %s',PlayerData[row][0], PlayerData[row][1]);
  }
  
  // Populate the Penalty Table for the Weekly Report
  PenaltyTable = subEmailPlayerPenaltyTable(PlayerData);
  
  // Send Weekly Report Email
  EmailSubject = LeagueName +' - Week ' + LastWeek + ' Report';
  EmailMessage = 'Week ' + LastWeek + ' is now complete and Week '+ Week +' has started. <br><br>Here is the week report for Week ' + LastWeek + '.<br><br>' +
    MatchesPlayed +' matches were played this week.<br>'+
      'etc etc etc...<br><br>';
  
  EmailMessage += PenaltyTable;
  
  MailApp.sendEmail('triadgaminglt@gmail.com', EmailSubject, EmailMessage,{name:'Triad Gaming Booster League Manager',htmlBody:EmailMessage});
  
  // Execute Ranking function in Standing tab
  fcnUpdateStandings(ss);
  
  // Copy all data to League Spreadsheet
  fcnCopyStandingsResults(ss, shtConfig);
  
}