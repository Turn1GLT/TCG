// **********************************************
// function subGetEmailAddress()
//
// This function gets the email addresses from 
// the configuration file
//
// **********************************************

function subGetEmailAddress(shtConfig, WinPlyr, LosPlyr){
  
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
// function fcnSendConfirmEmail()
//
// This function generates the confirmation email 
// after a match report has been submitted
//
// **********************************************

function fcnSendConfirmEmail(LeagueName, Addresses, MatchData) {
  
  // Variables
  var EmailSubject;
  var EmailMessage;
  var EmailName1 = '';
  var EmailName2 = '';
  
  // Open GLM - Email Templates
  var ssEmail = SpreadsheetApp.openById('15-IjvgcgHWx6nRc0U_Fzg0iUYS_rD6-u5tNZELdZxOo');
  var shtEmailTemplates = ssEmail.getSheetByName('Templates');
  var Headers = shtEmailTemplates.getRange(3,2,25,1).getValues();
  
  // Match Data Assignation
  var MatchID = MatchData[2][0];
  var Week    = MatchData[3][0];
  var Winr    = MatchData[4][0];
  var Losr    = MatchData[5][0];
 
  // Set the EmailName according to the addresses
  if (Addresses[0] != '') EmailName1 = Addresses[0];
  if (Addresses[0] != '') EmailName2 = Addresses[1];
  
  // Add Masterpiece mention if necessary
  if (MatchData[24][2] == 'Last Card is Masterpiece'){
    var Masterpiece = MatchData[23][2];
    MatchData[23][2] += ' (Masterpiece)' 
  }

  // Set Email Subject
  EmailSubject = LeagueName + " - Week " + Week + " - Match Result" ;
  
  // Set Email Subject
  EmailSubject = LeagueName + ' - Week ' + Week + ' - Match Result' ;
  
  // Start of Email Message
  EmailMessage = '<html><body>';
  
  EmailMessage += 'Hi ' + Winr + ' and ' + Losr + ',<br><br>Your match result has been received and succesfully processed for the ' + LeagueName + ', Week ' + Week + 
    '<br><br>Here is your match result:<br><br>';
    
  // Generate Match Data Table
  EmailMessage = subComposeHtmlMsg(EmailMessage, Headers, MatchData,1);
  
  EmailMessage += '<br>Click here to access the League Standings and Results:'+
    '<br>https://docs.google.com/spreadsheets/d/1-p-yXgcXEij_CsYwg7FadKzNwS6E5xiFddGWebpgTDY/edit?usp=sharing'+
      '<br><br>Click here to access your Card Pool:'+
        '<br>https://docs.google.com/spreadsheets/d/1lFiVQaE4_LxOKePdfhhUiBHJq0q3xbzxaDiOVwOQUI8/edit?usp=sharing'+
          '<br><br>Click here to send another Match Report:'+
            '<br>https://goo.gl/forms/jcDtOML96WlNLzVL2'+
              '<br><br>If you find any problems with your match result, please reply to this message and describe the situation as best you can. You will receive a response once it has been processed.'+
                '<br><br>Thank you for using TCG Booster League Manager from Turn 1 Gaming Leagues Applications';
  
  // End of Email Message
  EmailMessage += '</body></html>';
  
  // Sends email to both players with the Match Data
  if (EmailName1 != '') MailApp.sendEmail(EmailName1, EmailSubject, EmailMessage,{name:'TCG Booster League Manager',htmlBody:EmailMessage});
  if (EmailName2 != '') MailApp.sendEmail(EmailName2, EmailSubject, EmailMessage,{name:'TCG Booster League Manager',htmlBody:EmailMessage});
}


// **********************************************
// function fcnSendErrorEmail()
//
// This function generates the error email 
// after a match report has been submitted
//
// **********************************************

function fcnSendErrorEmail(LeagueName, Addresses, MatchData, MatchID, Status) {
  
  // Variables
  var EmailSubject;
  var EmailMessage;
  var EmailName1 = '';
  var EmailName2 = '';
  
  // Open GLM - Email Templates
  var ssEmail = SpreadsheetApp.openById('15-IjvgcgHWx6nRc0U_Fzg0iUYS_rD6-u5tNZELdZxOo');
  var shtEmailTemplates = ssEmail.getSheetByName('Templates');
  var Headers = shtEmailTemplates.getRange(3,2,25,1).getValues();
  
  // Match Data Assignation
  var MatchID = MatchData[2][0];
  var Week    = MatchData[3][0];
  var Winr    = MatchData[4][0];
  var Losr    = MatchData[5][0];
  
  var StatusMsg;
 
  // Set the EmailName according to the addresses
  if (Addresses[0] != '') EmailName1 = Addresses[0];
  if (Addresses[0] != '') EmailName2 = Addresses[1];
  
  // Selects the Appropriate Error Message
  switch (Status[0]){
    case 0: StatusMsg = 'Error 0'; break;
    case 1: StatusMsg = 'Error 1'; break;
  }
  
  // Set Email Subject
  EmailSubject = LeagueName + ' - Week ' + Week + ' - Process Error' ;
  
  // Start of Email Message
  EmailMessage = '<html><body>';
  
  if (MatchID > 0){
    EmailMessage += 'Hi ' + Winr + ' and ' + Losr + ',<br><br>Your match result has been succesfully received for the ' + LeagueName + ', Week ' + Week + 
      "<br><br>We were able to process the match data but an error has been detected in the submitted form.<br>Please contact us to resolve this error as soon as possible<br><br>"+
        "Error Message:<br>" + StatusMsg +
          '<br><br>Here is your match result:<br><br>';
  } 
  
  else {
    EmailMessage += 'Hi ' + Winr + ' and ' + Losr + ',<br><br>Your match result has been succesfully received for the ' + LeagueName + ', Week ' + Week + 
      "<br><br>An error has been detected in the submitted form or in one of the player's record. Unfortunately, this error prevented us to process the match report.<br><br>"+
        "Error Message:<br>" + StatusMsg +
          '<br><br>Here is your match result:<br><br>';
  }
  
  EmailMessage = subComposeHtmlMsg(EmailMessage, Headers, MatchData,StatusMsg);
  
  EmailMessage +='<br>Click here to access the League Standings and Results:'+
    '<br>https://docs.google.com/spreadsheets/d/1-p-yXgcXEij_CsYwg7FadKzNwS6E5xiFddGWebpgTDY/edit?usp=sharing'+
      '<br><br>Click here to access your Card Pool:'+
        '<br>https://docs.google.com/spreadsheets/d/1lFiVQaE4_LxOKePdfhhUiBHJq0q3xbzxaDiOVwOQUI8/edit?usp=sharing'+
          '<br><br>Click here to send another Match Report:'+
            '<br>https://goo.gl/forms/jcDtOML96WlNLzVL2'+
              '<br><br>If you find any problems with your match result, please reply to this message and describe the situation as best you can. You will receive a response once it has been processed.'+
                '<br><br>Thank you for using TCG Booster League Manager from Turn 1 Gaming Leagues Applications';
  
  // End of Email Message
  EmailMessage += '</body></html>';
  
  // Sends email to both players with the Match Data
  if (EmailName1 != '') MailApp.sendEmail(EmailName1, EmailSubject, EmailMessage,{name:'TCG Booster League Manager',htmlBody:EmailMessage});
  if (EmailName2 != '') MailApp.sendEmail(EmailName2, EmailSubject, EmailMessage,{name:'TCG Booster League Manager',htmlBody:EmailMessage});
  MailApp.sendEmail("gamingleaguemanager@gmail.com", EmailSubject, EmailMessage,{name:'TCG Booster League Manager',htmlBody:EmailMessage});
}




// **********************************************
// function subComposeHtmlMsg()
//
// This function generates the HTML table that displays 
// the Match Data and Booster Pack Data
//
// **********************************************

function subComposeHtmlMsg(EmailMessage, Headers, MatchData, Param){
  for(var row=0; row<24; ++row){

    if(row == 1) ++row;
    
    // Start of Match Table
    if(row == 0) {
      EmailMessage += '<table style="border-collapse:collapse;" border = 1 cellpadding = 5><tr>';
    }
    
    // Match Data
    if(row < 7) {
      EmailMessage += '<tr><td>'+Headers[row][0]+'</td><td>'+MatchData[row][0]+'</td></tr>';
    }
    
    // End of first Table
    if(row == 7) EmailMessage += '</table><br>';
    
    // Start of Pack Table
    if(row == 9 && Param == 1) {
      EmailMessage += 'Booster Pack Content<br><br><font size="4"><b>'+MatchData[row][0]+
        '</b></font><br><table style="border-collapse:collapse;" border = 1 cellpadding = 5><th>Item</th><th>Card Number</th><th>Card Name</th><th>Rarity</th>';
    }
    
    // Pack Data
    if(row > 9 && Param == 1) {
      EmailMessage += '<tr><td>'+Headers[row][0]+'</td><td>'+MatchData[row][1]+'</td><td>'+MatchData[row][2]+'</td><td>'+MatchData[row][3]+'</td></tr>';
    }
    
    // If Param is Not 1, Error is Present 
    if(row == 9 && Param != 1) {
      row = 24;
    }
    
  }
  return EmailMessage +'</table>';
}






















