// **********************************************
// function OnOpen()
//
// Function executed everytime the Spreadsheet is
// opened or refreshed
//
// **********************************************

function onOpenMinisMaster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var AnalyzeDataMenu  = [];
  AnalyzeDataMenu.push({name: 'Analyze New Match Entry', functionName: 'fcnMain'});
  AnalyzeDataMenu.push({name: 'Reset Match Entries', functionName:'fcnResetLeagueMatch'});
  
  var LeagueMenu = [];
  LeagueMenu.push({name:'Initialize League', functionName:'fcnInitLeague'});
  LeagueMenu.push({name:'Setup Response Sheets',functionName:'fcnSetupResponseSht'});
  LeagueMenu.push({name:'Update Config ID & Links', functionName:'fcnUpdateLinksIDs'});
  LeagueMenu.push(null);
  LeagueMenu.push({name:'Delete Players DB',functionName:'fcnDelPlayerCardDB'});
  LeagueMenu.push({name:'Delete Players Pools', functionName:'fcnDelPlayerCardPoolSht'});

  
  ss.addMenu("Manage League", LeagueMenu);
  ss.addMenu("Process Data", AnalyzeDataMenu);
}

// **********************************************
// function fcnWeekChange()
//
// When the Week number changes, this function analyzes all
// generates a weekly report 
//
// **********************************************

function onWeekChangeMinisMaster(){

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Open Configuration Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // League Name EN
  var Location = shtConfig.getRange(12,2).getValue();
  var LeagueTypeEN = shtConfig.getRange(13,2).getValue();
  var LeagueTypeFR = shtConfig.getRange(14,2).getValue();
  var LeagueNameEN = shtConfig.getRange(3,2).getValue() + ' ' + LeagueTypeEN;
  var LeagueNameFR = LeagueTypeFR + ' ' + shtConfig.getRange(3,2).getValue();
  
  // Open Cumulative Spreadsheet
  var shtCumul = ss.getSheetByName('Cumulative Results');
  var Week = shtCumul.getRange(2,3).getValue();
  var LastWeek = Week - 1;
  var WeekShtName = 'Week'+Week;
  var shtWeek = ss.getSheetByName(WeekShtName);
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
  EmailSubject = LeagueNameEN +' - Week ' + LastWeek + ' Report';
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