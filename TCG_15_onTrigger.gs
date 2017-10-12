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
  
  // Email Variables
  var EmailRecipientsEN;
  var EmailSubjectEN;
  var EmailMessageEN;
  var EmailRecipientsFR;
  var EmailSubjectFR;
  var EmailMessageFR;
  var EmailLanguage;  
  var Recipients;
  
  // League Name
  var LeagueLocation = shtConfig.getRange(11,2).getValue();
  var LeagueTypeEN   = shtConfig.getRange(13,2).getValue();
  var LeagueTypeFR   = shtConfig.getRange(14,2).getValue();
  var LeagueNameEN   = LeagueLocation + ' ' + LeagueTypeEN;
  var LeagueNameFR   = LeagueTypeFR + ' ' + LeagueLocation;
  
  // Get Document URLs
  var urlStandingsEN = shtConfig.getRange(17,2).getValue();
  var urlStandingsFR = shtConfig.getRange(20,2).getValue();
  
  // Facebook Page Link
  var urlFacebook = shtConfig.getRange(50, 2).getValue();
  
  // Function Values
  var shtCumul = ss.getSheetByName('Cumulative Results');
  var Week = shtCumul.getRange(2,3).getValue();
  var LastWeek = Week - 1;
  var WeekShtName = 'Week'+LastWeek;
  var shtWeek = ss.getSheetByName(WeekShtName);
  var shtPlayers = ss.getSheetByName('Players');
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  var LocationEmail = shtConfig.getRange(12,2).getValue();
  
  // Function Variables
  var PenaltyTable;
  var WeekData;
  var TotalMatch = 0;
  var TotalWins = 0;
  var TotalLoss = 0;
  var TotalMatchStore = 0;
  var MostParam;
  
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
  MostParam = 'Store';
  PlayerMostGames = fcnPlayerWithMost(PlayerMostGames, NbPlayers, shtWeek, MostParam);
  
  // Player with Most Losses
  MostParam = 'Loss';
  PlayerMostLoss = fcnPlayerWithMost(PlayerMostLoss, NbPlayers, shtWeek, MostParam);
  
 // Verify Week Matches Data Integrity
  WeekData = shtWeek.getRange(5,4,NbPlayers,6).getValues(); //[0]= Matches Played [1]= Wins [2]= Losses [5]= Matches in Store
  // Get Total Matches Played
  for(var plyr=0; plyr<NbPlayers; plyr++){
    if(WeekData[plyr][0] > 0) TotalMatch += WeekData[plyr][0];
  }
  TotalMatch = TotalMatch/2;
  
  // Get Total Wins
  for(plyr=0; plyr<NbPlayers; plyr++){
    if(WeekData[plyr][1] > 0 ) TotalWins += WeekData[plyr][1];
  }
  
  // Get Total Losses
  for(plyr=0; plyr<NbPlayers; plyr++){
    if(WeekData[plyr][2] > 0 ) TotalLoss += WeekData[plyr][2];
  }
  
  // Get Amount of matches played at the store this week.
  for(plyr=0; plyr<NbPlayers; plyr++){
    if(WeekData[plyr][5] > 0 ) TotalMatchStore += WeekData[plyr][5];
  }
 TotalMatchStore = TotalMatchStore/2;
  
  // If All Totals are equal, Week Data is Valid, Send Week Report
  if(TotalMatch == TotalWins &&  TotalMatch == TotalLoss && TotalWins == TotalLoss) {
    
    // Send Weekly Report Email
    EmailSubjectEN = LeagueNameEN +" - Week " + LastWeek + " Report";
    EmailSubjectFR = LeagueNameFR +" - Rapport de la semaine " + LastWeek;
    
    // Generate Week Report Messages
    EmailMessageEN = fcnGenWeekReportMsgEN(EmailMessageEN, LastWeek, Week, TotalMatch, TotalMatchStore, PlayerMostGames, PlayerMostLoss);
    EmailMessageFR = fcnGenWeekReportMsgFR(EmailMessageFR, LastWeek, Week, TotalMatch, TotalMatchStore, PlayerMostGames, PlayerMostLoss);
    
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
      EmailMessageEN += PenaltyTable;
      EmailMessageFR += PenaltyTable;
    }
    
    // English Custom Message
    // Add Standings Link
    EmailMessageEN += "<br><br>Click here to access the League Standings and Results:<br>" + urlStandingsEN ;
    
    // Add Facebook Page Link
    EmailMessageEN += "<br><br>Please join the Community Facebook page to chat with other players and plan matches.<br>" + urlFacebook;
    
    // Turn1 Signature
    EmailMessageEN += "<br><br>Thank you for using TCG Booster League Manager from Turn 1 Gaming Leagues & Tournaments";
    
    
    // French Custom Message
    // Add Standings Link
    EmailMessageFR += "<br><br>Cliquez ici pour accéder aux résutlats et classement de la ligue:<br>" + urlStandingsFR ;
    
    // Add Facebook Page Link
    EmailMessageFR += "<br><br>Joignez vous à la page Facebook de la communauté pour discuter avec les autres joueurs et organiser vos matches.<br>" + urlFacebook;
    
    // Turn1 Signature
    EmailMessageFR += "<br><br>Merci d'utiliser TCG Booster League Manager de Turn 1 Gaming Ligues & Tournois";
    
    // General Recipients
    Recipients = LocationEmail + ', turn1glt@gmail.com';
    
    // Get English Players Email
    EmailLanguage = "English";
    EmailRecipientsEN = subGetEmailRecipients(shtPlayers, NbPlayers, EmailLanguage);
    Logger.log(EmailRecipientsEN);
    
    // Get French Players Email
    EmailLanguage = "Français";
    EmailRecipientsFR = subGetEmailRecipients(shtPlayers, NbPlayers, EmailLanguage);
    Logger.log(EmailRecipientsFR);
    
    // Send English Email
    MailApp.sendEmail(Recipients, EmailSubjectEN,"",{bcc:EmailRecipientsEN,name:'Turn 1 Gaming League Manager',htmlBody:EmailMessageEN});
    
    // Send French Email
    MailApp.sendEmail(Recipients, EmailSubjectFR,"",{bcc:EmailRecipientsFR,name:'Turn 1 Gaming League Manager',htmlBody:EmailMessageFR});
    
    // Execute Ranking function in Standing tab
    fcnUpdateStandings(ss, shtConfig);
    
    // Copy all data to League Spreadsheet
    fcnCopyStandingsResults(ss, shtConfig, LastWeek, 0);
  }
  
  // If Week Match Data is not Valid
  else{
    Logger.log('Week Match Data is not Valid');
    Logger.log('Total Match Played: %s',TotalMatch);
    Logger.log('Total Wins: %s',TotalWins);
    Logger.log('Total Losses: %s',TotalLoss);
  
    // Send Log by email
    var recipient = Session.getActiveUser().getEmail();
    var subject = LeagueNameEN + ' - Week ' + LastWeek + ' - Week Data is not Valid';
    var body = Logger.getLog();
    MailApp.sendEmail(recipient, subject, body)
  }
}