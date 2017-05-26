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
// function fcnGenEmailConfirmation()
//
// This function generates the email confirmation 
// after a match report has been submitted
//
// **********************************************

function fcnGenEmailConfirmation(LeagueName, Addresses, MatchData) {
  
  // Variables
  var EmailSubject;
  var EmailMessage;
  var EmailName;
  
  // Open GLM - Email Templates
  var ssEmail = SpreadsheetApp.openById('15-IjvgcgHWx6nRc0U_Fzg0iUYS_rD6-u5tNZELdZxOo');
  var shtEmailTemplates = ssEmail.getSheetByName('Templates');
  var Headers = shtEmailTemplates.getRange(3, 1, 22, 1).getValues();
  
  // Match Data Assignation
  var MatchID = MatchData[0][0];
  var Week    = MatchData[1][0];
  var Winr    = MatchData[2][0];
  var Losr    = MatchData[3][0];
 
  // Set the EmailName according to the addresses
  if (Addresses[0] != '' && Addresses[1] != '') EmailName = Addresses[0] + ', ' + Addresses[1];
  if (Addresses[0] != '' && Addresses[1] == '') EmailName = Addresses[0];
  if (Addresses[0] == '' && Addresses[1] != '') EmailName = Addresses[1];
  
  // Add Masterpiece mention if necessary
  if (MatchData[22][2] == 'Last Card is Masterpiece'){
    var Masterpiece = MatchData[21][2];
    MatchData[21][2] = Masterpiece + ' (Masterpiece)' 
  }

  // Set Email Subject
  EmailSubject = LeagueName + ' - Week ' + Week + ' - Match Result' ;
  
  // Set Email Message
  EmailMessage = '<html><body>Hi ' + Winr + ' and ' + Losr + ',<br><br>Your match result has been succesfully received for the ' + LeagueName + ', Week ' + Week + 
    '<br><br>Here is your match result:<br><br><table style="border-collapse:collapse;" border = 1 cellpadding = 5><tr>';
    
  EmailMessage = composeHtmlMsg(EmailMessage, Headers, MatchData);
  
  EmailMessage = EmailMessage + '<br>Click here to access the League Standings and Results:'+
    '<br>https://docs.google.com/spreadsheets/d/1-p-yXgcXEij_CsYwg7FadKzNwS6E5xiFddGWebpgTDY/edit?usp=sharing'+
      '<br><br>Click here to access your Card Pool:'+
        '<br>https://docs.google.com/spreadsheets/d/1lFiVQaE4_LxOKePdfhhUiBHJq0q3xbzxaDiOVwOQUI8/edit?usp=sharing'+
          '<br><br>Click here to send another Match Report:'+
            '<br>https://goo.gl/forms/jcDtOML96WlNLzVL2'+
              '<br><br>If you find any problem with your match result and this confirmation, please reply to this message and describe the situation as best as possible so I can make the appropriate correction.'+
                '<br><br>Thank you for using TCG Booster League Manager from Gaming League Manager Applications </body></html>';
  
  // Sends email to both players with the Match Data
  MailApp.sendEmail(EmailName, EmailSubject, EmailMessage,{name:'Gaming League Manager',htmlBody:EmailMessage});
  //MailApp.sendEmail("gamingleaguemanager@gmail.com", EmailSubject, EmailMessage);
}




//-----------------------------------------------------------

function composeHtmlMsg(EmailMessage, Headers, MatchData){
  for(var row=0; row<22; ++row){
    
    // Match Data
    if(row < 5) {
      EmailMessage+='<tr><td>'+Headers[row][0]+'</td><td>'+MatchData[row][0]+'</td></tr>';
    }
    
    // End of first Table
    if(row == 5) EmailMessage+='</table><br>';
    
    //
    if(row == 7) EmailMessage+='Booster Pack Content<br><br><font size="4"><b>'+MatchData[row][0]+'</b></font><br><table style="border-collapse:collapse;" border = 1 cellpadding = 5><th>Item</th><th>Card Number</th><th>Card Name</th><th>Rarity</th>';
    
    // Pack Data
    if(row > 7) {
      EmailMessage+='<tr><td>'+Headers[row][0]+'</td><td>'+MatchData[row][1]+'</td><td>'+MatchData[row][2]+'</td><td>'+MatchData[row][3]+'</td></tr>';
    }
    
  }
  return EmailMessage+'</table>';
}




















