// **********************************************
// function OnOpen()
//
// Function executed everytime the Spreadsheet is
// opened or refreshed
//
// **********************************************

function onOpenTCG_Master() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var AnalyzeDataMenu  = [];
  AnalyzeDataMenu.push({name: 'Process New Match Entry', functionName: 'fcnProcessMatchTCG'});
  AnalyzeDataMenu.push({name: 'Reset Match Entries', functionName:'fcnResetLeagueMatch'});
  
  var LeagueMenu = [];
  LeagueMenu.push({name:'Update Config ID & Links', functionName:'fcnUpdateLinksIDs'});
  LeagueMenu.push({name:'Create Match Report Forms', functionName:'fcnCreateReportForm'});
  LeagueMenu.push({name:'Setup Response Sheets',functionName:'fcnSetupResponseSht'});
  LeagueMenu.push({name:'Create Registration Forms', functionName:'fcnCreateRegForm'});
  LeagueMenu.push({name:'Initialize League', functionName:'fcnInitLeague'});
  LeagueMenu.push(null);
  LeagueMenu.push({name:'Generate Card DB',functionName:'fcnGenPlayerCardDB'});
  LeagueMenu.push({name:'Generate Card Pools', functionName:'fcnGenPlayerCardPool'});
  LeagueMenu.push({name:'Generate Starting Pools', functionName:'fcnGenPlayerStartPool'});
  LeagueMenu.push(null);
  LeagueMenu.push({name:'Delete Card DB',functionName:'fcnDelPlayerCardDB'});
  LeagueMenu.push({name:'Delete Card Pools', functionName:'fcnDelPlayerCardPool'});
  LeagueMenu.push({name:'Delete Starting Pools', functionName:'fcnDelPlayerStartPool'});

  
  ss.addMenu("Manage League", LeagueMenu);
  ss.addMenu("Process Data", AnalyzeDataMenu);
}

// **********************************************
// function fcnWeekChangeTCG()
//
// When the Week number changes, this function analyzes all
// generates a weekly report 
//
// **********************************************

function onWeekChangeTCG_Master(){

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Open Configuration Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  var cfgMinGame = shtConfig.getRange(5, 2).getValue();
  
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
  var WeekShtName = 'Week'+LastWeek;
  var shtWeek = ss.getSheetByName(WeekShtName);
  var NbPlayers = shtConfig.getRange(6,2).getValue()+1;
  var PenaltyTable;
  var EmailSubject;
  var EmailMessage;
  var MatchPlyd;
  var MatchPlydStore;
  var MatchesPlayed = 0;
  var MatchPlydStore;
  var MatchesPlayedStore = 0;
  
  // Array to Find Player with Most Matches Played in Store
  var PlayerMostGames = new Array(NbPlayers); 
  for(var i=0; i<NbPlayers; i++){
    PlayerMostGames[i] = new Array (2);// [0]= Player Name, [1]= Data  
  }
  // Array to Find Player with Most Losses
  var PlayerMostLoss = new Array(NbPlayers); // [0]= Player Name, [1]= Data
  for(var i=0; i<NbPlayers; i++){
    PlayerMostLoss[i] = new Array (2);// [0]= Player Name, [1]= Data  
  }
  
  // Players Array to return Penalty Losses
  var PlayerData = new Array(32); // 0= Player Name, 1= Penalty Losses
  for(var plyr = 0; plyr < 32; plyr++){
    PlayerData[plyr] = new Array(2); 
    for (var val = 0; val < 2; val++) PlayerData[plyr][val] = '';
  }
  
  // Modify the Week Number in the Match Report Sheet
  fcnModifyWeekMatchReport(ss, shtConfig);
  
  //Player with Most Games Played in Store
  var Param = 'Store';
  PlayerMostGames = fcnPlayerWithMost(PlayerMostGames, NbPlayers, shtWeek, Param);
  Logger.log(PlayerMostGames[0]);
  Logger.log(PlayerMostGames[1]);
  Logger.log(PlayerMostGames[2]);
  
  // Player with Most Losses
  Param = 'Loss';
  PlayerMostLoss = fcnPlayerWithMost(PlayerMostLoss, NbPlayers, shtWeek, Param);
  Logger.log(PlayerMostLoss[0]);
  Logger.log(PlayerMostLoss[1]);
  Logger.log(PlayerMostLoss[2]);
  
  // Get Amount of matches played this week.
  MatchPlyd = shtWeek.getRange(5, 4, 32, 1).getValues();
  for(var plyr=0; plyr<32; plyr++){
    if(MatchPlyd[plyr][0] > 0) MatchesPlayed += MatchPlyd[plyr][0];
  }
  MatchesPlayed = MatchesPlayed/2;
  
  // Get Amount of matches played at the store this week.
  MatchPlydStore = shtWeek.getRange(5, 9, 32, 1).getValues();
  for(plyr=0; plyr<32; plyr++){
    if(MatchPlydStore[plyr][0] > 0 ) MatchesPlayedStore += MatchPlydStore[plyr][0];
  }
  MatchesPlayedStore = MatchesPlayedStore/2;


  // Send Weekly Report Email
  EmailSubject = LeagueNameEN +" - Week " + LastWeek + " Report";
  
  EmailMessage = "Hello everyone,<br><br>Week " + LastWeek + " is now complete and Week "+ Week +" has started."+
    " <br><br>Here is the week report for Week " + LastWeek + 
      "<br><br><b>Matches Played:</b> " + MatchesPlayed +" matches were played this week."+
        "<br><br><b>Matches Played in Store:</b> " + MatchesPlayedStore +" matches were played at the store this week.";

  // Players Awards
  EmailMessage += '<br><br><font size="3"><b>Week Awards</b></font>';
  // Most Matches Played in Store
  EmailMessage += '<br><br>The player(s) with the most matches played in store this week: ' + 
    '<b>' + PlayerMostGames[0][0] + '</b> with <b>' + PlayerMostGames[0][1] + '</b> games played';
  // Most Losses
  EmailMessage += "<br><br>The player(s) with the most losses this week: " + 
    "<b>" + PlayerMostLoss[0][0] + "</b> with <b>" + PlayerMostLoss[0][1] + "</b> losses";
  
  EmailMessage += "<br><br>The Players mentioned above win a <b>free Standard Legal Booster Pack of their choice</b>";
  
  EmailMessage += "<br><br>Good luck to all player for week "+ Week;
    
  // If there is a minimum games to play per week, generate the Penalty Losses
  if(cfgMinGame > 0){

    // Analyze if Players have missing matches to apply Loss Penalties
    PlayerData = fcnAnalyzeLossPenalty(ss, Week, PlayerData);
    
    // Logs All Players Record
    for(var row = 0; row<32; row++){
      if (PlayerData[row][0] != '') Logger.log('Player: %s - Missing: %s',PlayerData[row][0], PlayerData[row][1]);
    }
    
    // Populate the Penalty Table for the Weekly Report
    PenaltyTable = subEmailPlayerPenaltyTable(PlayerData);  
    // Update the Email message to add the Penalty Losses table
    EmailMessage += PenaltyTable;
  }
  
  MailApp.sendEmail('turn1glt@gmail.com', EmailSubject, EmailMessage,{name:'Turn 1 Gaming League Manager',htmlBody:EmailMessage});
  
  // Execute Ranking function in Standing tab
  fcnUpdateStandings(ss, shtConfig);
  
  // Copy all data to League Spreadsheet
  fcnCopyStandingsResults(ss, shtConfig, LastWeek, 0);
  
}