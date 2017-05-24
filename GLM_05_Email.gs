// **********************************************
// function subGetEmailAddress()
//
// This function gets the email addresses from 
// the configuration file
//
// **********************************************

function subGetEmailAddress(shtConfig, WinPlyr, LosPlyr, shtTest){
  
  // Config File Email Address column
  var colEmail = 6;
  var NbPlayers = shtConfig.getRange(16,7).getValue();
  var rowWinr = 0;
  var rowLosr = 0;
  var Addresses = new Array(2);
  
  var PlayerNames = shtConfig.getRange(17,2,NbPlayers,1).getValues();
  
  // Find Players rows
  for (var row = 0; row < NbPlayers; row++){
    if (PlayerNames[row] == WinPlyr) rowWinr = row + 17;
    if (PlayerNames[row] == LosPlyr) rowLosr = row + 17;
    if (rowWinr > 0 && rowLosr > 0) row = NbPlayers + 1;
  }
  
  // Get Email addresses using the players rows
  Addresses[0] = shtConfig.getRange(rowWinr,colEmail).getValue();
  Addresses[1] = shtConfig.getRange(rowLosr,colEmail).getValue();
  
  return Addresses;
}


// **********************************************
// function fcnGenEmailConfirmation()
//
// This function generates the email confirmation 
// after a match report has been submitted
//
// **********************************************

function fcnGenEmailConfirmation(LeagueName, Addresses, MatchData, shtTest) {
  
  var EmailSubject;
  var EmailMessage;
  var EmailName = Addresses[0] + ', ' + Addresses[1];
  
  var MatchID = MatchData[0];
  var Week    = MatchData[1];
  var Winr    = MatchData[2];
  var Losr    = MatchData[3];
 
  if (MatchData[22] == 'Last Card is Masterpiece'){
    var Masterpiece = MatchData[21];
    MatchData[21] = Masterpiece + ' (Masterpiece)' 
  }
  
  EmailSubject = LeagueName + ' - Match Result Received for Week ' + Week ;
  EmailMessage = Winr + " and " + Losr + ",\n\nYour match result has been succesfully received for week " + Week + " in the " + LeagueName + 
    "\n\nHere is the Data for your Match:\n" +
    "\nMatchID: " + MatchData[0] +
    "\nWeek: " + MatchData[1] +
    "\nWinning Player: " + MatchData[2] +
    "\nLosing Player: " + MatchData[3] +
    "\nScore: " + MatchData[4] +
    "\n\nBooster Pack Content" +
    "\n\nSet     : " + MatchData[7] +
    "\nCard 1 : " + MatchData[8] +
    "\nCard 2 : " + MatchData[9] +
    "\nCard 3 : " + MatchData[10] +
    "\nCard 4 : " + MatchData[11] +
    "\nCard 5 : " + MatchData[12] +
    "\nCard 6 : " + MatchData[13] +
    "\nCard 7 : " + MatchData[14] +
    "\nCard 8 : " + MatchData[15] +
    "\nCard 9 : " + MatchData[16] +
    "\nCard 10: " + MatchData[17] +
    "\nCard 11: " + MatchData[18] +
    "\nCard 12: " + MatchData[19] +
    "\nCard 13: " + MatchData[20] +
    "\nCard 14: " + MatchData[21] + 
    "\n\nClick here to access the League Standings and Results: " +
    "\n\nIf you find any discrepancies with your match result and this confirmation, please reply to this message and describe the situation as best as possible so I can make the appropriate correction." +
    "\n\nThank you for using TCG Booster League Manager"
  
  // SENDS THE EMAIL TO THE PERSON WHO MADE THE RESERVATION
  MailApp.sendEmail(EmailName, EmailSubject, EmailMessage);
  //MailApp.sendEmail("gamingleaguemanager@gmail.com", EmailSubject, EmailMessage);
}



















