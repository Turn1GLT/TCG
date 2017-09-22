// **********************************************
// function fcnRegistrationTCG
//
// This function adds the new player to
// the Player's List and calls other functions
// to create its complete profile
//
// **********************************************

function fcnRegistrationTCG(ss, shtResponse, RowResponse){

  var shtConfig = ss.getSheetByName('Config');
  var shtPlayers = ss.getSheetByName('Players');
  var ssWeekBstrID = shtConfig.getRange(40, 2).getValue();
  
  var PlayerData = new Array(8);
  PlayerData[0] = 0 ; // Function Status
  PlayerData[1] = ''; // Number of Players
  PlayerData[2] = ''; // New Player Full Name
  PlayerData[3] = ''; // New Player First Name
  PlayerData[4] = ''; // New Player Email
  PlayerData[5] = ''; // New Player Phone Number
  PlayerData[6] = ''; // New Player Language
  PlayerData[7] = ''; // New Player DCI Number
  
  
  // Add Player to Player List
  PlayerData = fcnAddPlayerTCG(shtConfig, shtPlayers, shtResponse, RowResponse, PlayerData);
  var NbPlayers  = PlayerData[1];
  var PlayerName = PlayerData[2];
  
  // If Player was succesfully added, Generate Card DB, Generate Card Pool, Modify Match Report Form and Add Player to Weekly Booster
  if(PlayerData[0] == 1) {
    fcnGenPlayerCardDB();
    Logger.log('Card Database Generated'); 
    fcnGenPlayerCardPool();
    Logger.log('Card Pool Generated');
    fcnModifyReportFormTCG(ss, shtConfig, shtPlayers);
    // If Weekly Booster is used, add Player to the list
    if(ssWeekBstrID != ''){
      fcnAddPlayerWeeklyBooster(ssWeekBstrID, NbPlayers, PlayerName);
      Logger.log('Player Added to Weekly Booster');
    }
    
    // Execute Ranking function in Standing tab
    fcnUpdateStandings(ss, shtConfig);
    
    // Copy all data to Standing League Spreadsheet
    fcnCopyStandingsResults(ss, shtConfig, 0, 1);
    
    // Send Confirmation to New Player
    fcnSendNewPlayerConf(shtConfig, PlayerData);
    Logger.log('Confirmation Email Sent');
  }
  
  // Send Log for new Registration
  var recipient = Session.getActiveUser().getEmail();
  var subject = 'Form Log Test';
  var body = Logger.getLog();
  MailApp.sendEmail(recipient, subject, body);
}




// **********************************************
// function fcnAddPlayerTCG_Master
//
// This function adds the new player to
// the Player's List
//
// **********************************************

function fcnAddPlayerTCG(shtConfig, shtPlayers, shtResponse, RowResponse, PlayerData) {

  // Get All Values from Response Sheet
  var EmailAddress = shtResponse.getRange(RowResponse,2).getValue();
  var FirstName = shtResponse.getRange(RowResponse,3).getValue();
  var LastName = shtResponse.getRange(RowResponse,4).getValue();
  var PlayerName = FirstName + ' ' + LastName;
  var Phone = shtResponse.getRange(RowResponse,5).getValue();
  var Language = shtResponse.getRange(RowResponse,6).getValue();
  var DCINum = shtResponse.getRange(RowResponse,7).getValue();
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  
  // Copy Values to Players Sheet at the Next Empty Spot (Number of Players + 3)
  var NextPlayerRow = NbPlayers + 3;
  // Name
  shtPlayers.getRange(NextPlayerRow, 2).setValue(PlayerName);
  Logger.log('Player Name: %s',PlayerName);
  // Email Address
  shtPlayers.getRange(NextPlayerRow, 3).setValue(EmailAddress);
  Logger.log('Email Address: %s',EmailAddress);
  // Language
  shtPlayers.getRange(NextPlayerRow, 4).setValue(Language);
  Logger.log('Language: %s',Language);
  // Phone Number
  shtPlayers.getRange(NextPlayerRow, 5).setValue(Phone);
  Logger.log('Phone: %s',Phone);  
  // DCI Number
  shtPlayers.getRange(NextPlayerRow, 7).setValue(DCINum);
  Logger.log('DCI: %s',DCINum);  Logger.log('-----------------------------');
  
  PlayerData[0] = 1;
  PlayerData[1] = NbPlayers + 1;
  PlayerData[2] = PlayerName;
  PlayerData[3] = FirstName;
  PlayerData[4] = EmailAddress;
  PlayerData[5] = Phone;
  PlayerData[6] = Language;
  PlayerData[7] = DCINum;
  
  return PlayerData;
}


// **********************************************
// function fcnModifyReportFormTCG_Master
//
// This function modifies the Match Report Form
// to add new added players
//
// **********************************************

function fcnModifyReportFormTCG(ss, shtConfig, shtPlayers) {

  var MatchFormEN = FormApp.openById(shtConfig.getRange(36, 2).getValue());
  var FormItemEN = MatchFormEN.getItems();
  var NbFormItem = FormItemEN.length;
  
  var MatchFormFR = FormApp.openById(shtConfig.getRange(37, 2).getValue());
  var FormItemFR = MatchFormFR.getItems();

  // Function Variables
  var ItemTitle;
  var ItemPlayerListEN;
  var ItemPlayerListFR;
  var ItemPlayerChoice;
  
  var NbPlayers = shtPlayers.getRange(2, 1).getValue();
  var Players = shtPlayers.getRange(3, 2, NbPlayers, 1).getValues();
  var ListPlayers = [];
  
  // Loops to Find Players List
  for(var item = 0; item < NbFormItem; item++){
    ItemTitle = FormItemEN[item].getTitle();
    if(ItemTitle == 'Winning Player' || ItemTitle == 'Losing Player'){
      
      // Get the List Item from the Match Report Form
      ItemPlayerListEN = FormItemEN[item].asListItem();
      ItemPlayerListFR = FormItemFR[item].asListItem();
      
      // Build the Player List from the Players Sheet     
      for (i = 0; i < NbPlayers; i++){
        ListPlayers[i] = Players[i][0];
      }
      // Set the Player List to the Match Report Forms
      ItemPlayerListEN.setChoiceValues(ListPlayers);
      ItemPlayerListFR.setChoiceValues(ListPlayers);
      
//      ItemPlayerChoice = ItemPlayerListEN.getChoices();
      
//      Logger.log(ItemTitle);
//      for(var choice = 0; choice < ItemPlayerChoice.length; choice++){
//        Logger.log('Player: %s',ItemPlayerChoice[choice].getValue());
//      }
    }
  }
}

// **********************************************
// function fcnAddPlayerWeeklyBooster
//
// This function adds the new player to
// the Weekly Booster Spreadsheet
//
// **********************************************

function fcnAddPlayerWeeklyBooster(ssWeekBstrID, NbPlayers, PlayerName){
  
  var ssWeekBstr = SpreadsheetApp.openById(ssWeekBstrID);
  var Sheets = ssWeekBstr.getSheets();
  var NumSheets = ssWeekBstr.getNumSheets();
  
  // Function Variables
  var shtWeekBstr;
  var MaxCols;
  var MaxRows;
  var NewCol;
  var Player;
  var PlayerSlots;
  var ColumnAdded = 0;
  
  // Check in first sheet to find where to add the player
  // Get Sheet Data
  shtWeekBstr = Sheets[0];
  MaxCols = shtWeekBstr.getMaxColumns();
  PlayerSlots = MaxCols - 1;
  
  Logger.log('Nb Player   : %s',NbPlayers);
  Logger.log('Player Slots: %s',PlayerSlots);
  
  // If Number of Players is less or equal to Number of Columns, find the first unused column
  if(NbPlayers <= PlayerSlots){
    for(var col = 2; col <= 9; col++){
      Player = shtWeekBstr.getRange(3, col).getValue();
      if(Player == '') {
        NewCol = col;
        col = 10;
        Logger.log('Existing Column Found: %s', NewCol);
      }
    }
  }
  
  // If Number of Players is greater than Number of Columns, add a Column
  if(NbPlayers > PlayerSlots){
    NewCol = MaxCols + 1;
    ColumnAdded = 1;
    Logger.log('New Player added at Column: %s',NewCol);
  }
  

   
  // Loop through all sheets to add the player at the appropriate column
  for(var sheet = 0; sheet < NumSheets; sheet++){
    shtWeekBstr = Sheets[sheet];
    MaxCols = shtWeekBstr.getMaxColumns();
    MaxRows = shtWeekBstr.getMaxRows();
    // Copy the last Column as template
    if(ColumnAdded == 1) {
      shtWeekBstr.insertColumnAfter(MaxCols);
      shtWeekBstr.getRange(1, MaxCols, MaxRows, 1).copyTo(shtWeekBstr.getRange(1, NewCol, MaxRows, 1));
    }
    // Set Player Name
    shtWeekBstr.getRange(3, NewCol).setValue(PlayerName);
  }
}
