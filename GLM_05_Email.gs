// **********************************************
// function fcnGenEmailConfirmation()
//
// This function generates the email confirmation 
// after a match report has been submitted
//
// **********************************************

function fcnGenEmailConfirmation(MatchData) {
  
  var EmailSubject;
  var EmailMessage;
  var EmailName;
  
  var Week = MatchData[1];
  var Winr = MatchData[2];
  var Losr = MatchData[3];
  
  EmailSubject = "Match Post Confirmation for Week #" + Week ;
  EmailMessage = "Bonjour/Hi  " + Winr + ",\n\nLe PLC de test " + Losr + " n'est pas disponible pour les dates suivantes:\n\nThe " ;
  
  // SENDS THE EMAIL TO THE PERSON WHO MADE THE RESERVATION
  MailApp.sendEmail(EmailName, EmailSubject, EmailMessage);
  MailApp.sendEmail("gamingleaguemanager@gmail.com", EmailSubject, EmailMessage);
}


// **********************************************
// function subGetEmailAddress()
//
// This function gets the email addresses from 
// the configuration file
//
// **********************************************

function subGetEmailAddress(shtConfig){
  
  

}
