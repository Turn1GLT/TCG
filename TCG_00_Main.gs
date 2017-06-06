// **********************************************
// function fcnGameResults()
//
// This function populates the Game Results tab 
// once a player submitted his Form
//
// **********************************************

function fcnGameResults() {
  // Opens Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Sheet to get options
  var shtConfig = SpreadsheetApp.openById('14rR_7-SG9fTi-M7fpS7d6n4XrOlnbKxRW1Ni2ongUVU').getSheetByName('Config');
  var ConfigData = shtConfig.getRange(3,9,26,1).getValues();
  var GameType = shtConfig.getRange(11,2).getValue();
  var LeagueType = shtConfig.getRange(12,2).getValue();
  var LeagueName = shtConfig.getRange(3,2).getValue() + " " + GameType + " " + LeagueType;
  
  // Code Execution Options
  var OptDualSubmission = ConfigData[0][0]; // If Dual Submission is disabled, look for duplicate instead
  var OptPostResult = ConfigData[1][0];
  var OptPlyrMatchValidation = ConfigData[2][0];
  var OptTCGBooster = ConfigData[3][0];
  var OptSendEmail = ConfigData[6][0];
  
  // Columns Values and Parameters
  var ColMatchID = ConfigData[14][0];
  var ColPrcsd = ConfigData[15][0];
  var ColDataConflict = ConfigData[16][0];
  var ColErrorMsg = ConfigData[17][0];
  var ColPrcsdLastVal = ConfigData[18][0];
  var ColMatchIDLastVal = ConfigData[19][0];
  var RspnStartRow = ConfigData[20][0];
  var RspnDataInputs = ConfigData[21][0]; // from Time Stamp to Data Processed
  var NbCards = ConfigData[22][0];

  // Test Sheet (for Debug)
  var shtTest = ss.getSheetByName('Test') ; 
  
  // Form Responses Sheet Variables
  var shtRspn = ss.getSheetByName('Form Responses 16');
  var RspnMaxRows = shtRspn.getMaxRows();
  var RspnMaxCols = shtRspn.getMaxColumns();
  var RspnNextRowPrcss = shtRspn.getRange(1, ColPrcsdLastVal).getValue() + 1;
  var RspnPlyrSubmit;
  var RspnLocation;
  var RspnWeekNum;
  var RspnDataWeek;
  var RspnDataWinr;
  var RspnDataLosr;
  var RspnDataPrcssd = 0;
  var ResponseData;
  var MatchingRspnData;
  
  // Card List Variables
  var CardList = new Array(16); // 0 = Set Name, 1-14 = Card Numbers, 15 = Card 14 is Masterpiece (Y-N)
  var CardName;
  
  // Create Array of 16x4 where each row is Card 1-14 and each column is Card Info
  var PackData = new Array(16); // 0 = Set Name, 1-14 = Card Numbers, 15 = Card 14 is Masterpiece (Y-N)
  for(var cardnum = 0; cardnum < 16; cardnum++){
    PackData[cardnum] = new Array(4); // 0= Card in Pack, 1= Card Number, 2= Card Name, 3= Card Rarity
    for (var val = 0; val < 4; val++) PackData[cardnum][val] = '';
  }

  // Match Data Variables
  var MatchID; 
  var MatchData = new Array(26); // 0 = MatchID, 1 = Week #, 2 = Winning Player, 3 = Losing Player, 4 = Score, 5 = Winner Points, 6 = Loser Points, 7 = Card Set, 8-21 = Cards, 22 = Masterpiece (Y-N), 23 = Reserved, 24 = MatchPostStatus
  // Create Array of 26x4 where each row is Card 1-14 and each column is Card Info. This Info is only used for rows 8-21
  for(var cardnum = 0; cardnum < 26; cardnum++){
    MatchData[cardnum] = new Array(4); // 0= Item Value or Card In Pack, 1= Card Number, 2= Card Name, 3= Card Rarity
    for (var val = 0; val < 4; val++) MatchData[cardnum][val] = '';
  }
  
  // Email Addresses
  var EmailAddresses = new Array(3);
  EmailAddresses[0] = 'gamingleaguemanager@gmail.com';
  EmailAddresses[1] = '';
  EmailAddresses[2] = '';

  // Data Processing Flags
  var Status = new Array(2); // Status[0] = Status Value, Status[1] = Status Message
  Status[0] = 1;
  Status[1] = '';
  
  var DuplicateRspn = -99;
  var MatchingRspn = -98;
  var MatchPostStatus = -97;
  var CardDBUpdated = -96;
  
  Logger.log('Start of Main Function Executed');
  Logger.log('Dual Submission Option: %s',OptDualSubmission);
  Logger.log('Post Results Option: %s',OptPostResult);
  Logger.log('Player Match Validation Option: %s',OptPlyrMatchValidation);
  Logger.log('TCG Option: %s',OptTCGBooster);
  Logger.log('Send Email Option: %s',OptSendEmail);
  
  // Find a Row that is not processed in the Response Sheet (added data)
  for (var RspnRow = RspnNextRowPrcss; RspnRow <= RspnMaxRows; RspnRow++){
   
    // Copy the new response data (from Time Stamp to Data Processed Field
    ResponseData = shtRspn.getRange(RspnRow, 1, 1, RspnDataInputs).getValues();
    
    RspnDataPrcssd = ResponseData[0][25];
    RspnPlyrSubmit = ResponseData[0][1]; // Player Submitting
    RspnLocation   = ResponseData[0][2]; // Match Location (Store Yes or No)
    RspnWeekNum    = ResponseData[0][3]; // Week / Round Number
    RspnDataWinr   = ResponseData[0][4]; // Winning Player
    RspnDataLosr   = ResponseData[0][5]; // Losing Player
    
    // If week number is not empty and Processed is empty and both players are different, Response Data needs to be processed
    if (RspnWeekNum != '' && RspnDataPrcssd == ''){
      
      // If both Players in the response are different, continue
      if (RspnDataWinr != RspnDataLosr){
        
        // Generates the Match ID in advance if data analysis is successful
        MatchID = shtRspn.getRange(1, ColMatchIDLastVal).getValue() + 1;
        
        Logger.log('New Data Found at Row: %s',RspnRow);
        
        // Copy the new response data to Data Array
        ResponseData = shtRspn.getRange(RspnRow, 1, 1, RspnDataInputs).getValues();
        
        // Look for Duplicate Entry (looks in all entries with MatchID and combination of Week Number, Winner and Loser) 
        // Real code will look at Player Posting Data as well
        DuplicateRspn = fcnFindDuplicateData(ss, ConfigData, shtRspn, ResponseData, RspnRow, RspnStartRow, RspnMaxRows, RspnDataInputs, shtTest);  
        
        Logger.log('Duplicate Result: %s', DuplicateRspn);
        
        // FindDuplicateEntry function was executed properly and didn't find any Duplicate entry, continue analyzing the response data
        if (DuplicateRspn == 0){
          
          // If Dual Submission is enabled, Search if the other Entry matching this response has been submitted (must be enabled)
          if (OptDualSubmission == 'Enabled'){
            // function returns row where the matching data was found
            MatchingRspn = fcnFindMatchingData(ss, ConfigData, shtRspn, ResponseData, RspnRow, RspnStartRow, RspnMaxRows, RspnDataInputs, shtTest);
            if (MatchingRspn < 0) DuplicateRspn = 0 - MatchingRspn;
          }
          
          // Search if the other Entry matching this response has been submitted
          if (OptDualSubmission == 'Disabled'){
            MatchingRspn = RspnRow;
          }      
          
          Logger.log('Matching Result: %s', MatchingRspn);
          
          // If the result of the fcnFindMatchingEntry function returns something greater than 0, we found a matching entry, continue analyzing the response data
          if (MatchingRspn > 0){
            
            if (OptPostResult == 'Enabled'){
              
              // Get the Entry Data found at row MatchingRspn
              MatchingRspnData = shtRspn.getRange(MatchingRspn, 1, 1, RspnDataInputs).getValues();
              
              // Execute function to populate Match Result Sheet from processed data
              MatchData = fcnPostMatchResults(ss, ConfigData, shtRspn, ResponseData, MatchingRspnData, MatchID, MatchData, shtTest);
              MatchPostStatus = MatchData[25][0];
              
              Logger.log('Match Post Status: %s',MatchPostStatus);
              
              // If Match was populated in Match Results Tab
              if (MatchPostStatus == 1){
                // Match ID doesn't change because we assumed it was already OK
                
                // Copies all cards added to the Card Database
                if (OptTCGBooster == 'Enabled'){
                  for (var card = 0; card < NbCards; card++){
                    CardList[card] = ResponseData[0][card+7];
                  }
                  // If Pack was opened, Update Card Database and Card Pool for Appropriate player
                  if (CardList[0] != 'No Pack Opened') {
                    PackData = fcnUpdateCardDB(shtConfig, RspnDataLosr, CardList, PackData, shtTest);
                    // Copy all card names to Match Data [7-22]
                    for (var card = 0; card < NbCards; card++){
                      MatchData[card+9][0] = PackData[card][0]; // Card in Pack
                      MatchData[card+9][1] = PackData[card][1]; // Card Number
                      MatchData[card+9][2] = PackData[card][2]; // Card Name
                      MatchData[card+9][3] = PackData[card][3]; // Card Rarity
                      
                      if (PackData[card][2] == 'Card Name not Found for Card Number') {
                        Status = subGenErrorMsg(Status, -60,CardList[card]);
                        PackData[card][2] = Status[1];
                      }
                    }
                  }
                  // for debug
                  //shtTest.getRange(20,1,26,4).setValues(MatchData);
                }
              }
              
              // If MatchPostSuccess = 0, function was executed but was not able to post in the Match Result Tab
              if (MatchPostStatus < 0){
                // Updates the Match ID to an empty value 
                MatchID = '';
                // Generate the Status Message
                Status = subGenErrorMsg(Status, MatchPostStatus,0);
              }
            }
            // If Posting is disabled, generate Match ID for testing        
            if (OptPostResult == 'Disabled'){
              // Match ID doesn't change because we assumed it was already OK
              
            }
            // Set the Data Processed Flag
            RspnDataPrcssd = 1;
          }
          
          // If MatchingEntry = 0, fcnFindMatchingEntry did not find a matching entry, it might be the first response entry
          if (OptDualSubmission == 'Enabled' && MatchingRspn == 0){
            // Generate the Status Message
            Status[0] = '';
            Status[1] = 'Waiting for Other Response Submission';
            // Set the Data Processed Flag
            RspnDataPrcssd = 1;
            
          } 
          
          // If MatchingEntry = -1, fcnFindMatchingEntry was not executed properly, send email to notify
          if (OptDualSubmission == 'Enabled' && MatchingRspn == -1){
            // Set the Status Message
            Status = subGenErrorMsg(Status, MatchingRspn,0);
          }
        }
        
        // If Duplicate is found, send email to notify, set Response Data Processed to -1 to represent the Duplicate Entry
        if (DuplicateRspn > 0){
          
          // Updates the Match ID to an empty value 
          MatchID = '';
          
          // Set the Data Processed Flag
          RspnDataPrcssd = 1;        
          
          // Sets the Status Message
          Status = subGenErrorMsg(Status, -10,DuplicateRspn);
        }
        
        // If FindDuplicateEntry was not executed properly, send email to notify, set Response Data Processed to -2 to represent processing error
        if (DuplicateRspn < 0){
          
          // Updates the Match ID to an empty value 
          MatchID = '';
          
          // Set the Data Processed Flag
          RspnDataPrcssd = 1;  
          
          // Set the Status Message
          Status = subGenErrorMsg(Status, DuplicateRspn,0);
        }
      } 
      
      // If Both Players are the same, report error
      if (RspnDataWinr == RspnDataLosr){
        
        // Updates the Match ID to an empty value 
        MatchID = '';
        
        // Set the Data Processed Flag
        RspnDataPrcssd = 1;  
        
        // Set the Status Message
        Status = subGenErrorMsg(Status, -50,0);
      }
      
      // Call the Email Function, sends Match Data if Send Email Option is Enabled
      if(Status[0] == 1 && Status[1] == '' && OptSendEmail == 'Enabled') {
        // Get Email addresses from Config File
        EmailAddresses = subGetEmailAddress(shtConfig, EmailAddresses, RspnDataWinr, RspnDataLosr);
        fcnSendConfirmEmailEn(shtConfig, LeagueName, EmailAddresses, MatchData);
      }
      
      // If an Error has been detected that prevented to process the Match Data, send available data and Error Message
      if(Status[0] != 1 && Status[1] != 'Waiting for Other Response Submission') {
        
        // Populates Match Data
        MatchData[0][0] = ResponseData[0][0]; // TimeStamp
        MatchData[0][0] = Utilities.formatDate (MatchData[0][0], Session.getScriptTimeZone(), 'YYYY-MM-dd HH:mm:ss');
        
        MatchData[1][0] = ResponseData[0][2];  // Location (Store Y/N)
        MatchData[2][0] = MatchID;             // MatchID
        MatchData[3][0] = ResponseData[0][3];  // Week/Round Number
        MatchData[4][0] = ResponseData[0][4];  // Winning Player
        MatchData[5][0] = ResponseData[0][5];  // Losing Player
        MatchData[6][0] = ResponseData[0][6];  // Score
        
        // Get Player Email Addresses if Send Email Option is Enabled
        if (OptSendEmail == 'Enabled') {
          // Get Email addresses from Config File
          EmailAddresses = subGetEmailAddress(shtConfig, EmailAddresses, RspnDataWinr, RspnDataLosr);
        }
        // Send Error Message
        fcnSendErrorEmail(shtConfig, LeagueName, EmailAddresses, MatchData, MatchID, Status);
      }
      
      // If Player Submitted Feedback, send Feedback to Administrator
      if (ResponseData[0][23] != '') {
        if (EmailAddresses[1] == '' && EmailAddresses[2] == '') EmailAddresses = subGetEmailAddress(shtConfig, EmailAddresses, RspnDataWinr, RspnDataLosr);
        fcnSendFeedbackEmail(LeagueName, EmailAddresses, MatchData, ResponseData[0][23]);
      }
      
      // Set the Match ID (for both Response and Matching Entry), and Updates the Last Match ID generated, 
      if (MatchPostStatus == 1 || OptPostResult == 'Disabled'){
        shtRspn.getRange(RspnRow, ColMatchID).setValue(MatchID);
        shtRspn.getRange(1, ColMatchIDLastVal).setValue(MatchID);
      }
      // Set the Processed Flag and Status Message for the response
      shtRspn.getRange(RspnRow, ColPrcsd).setValue(RspnDataPrcssd);
      shtRspn.getRange(RspnRow, ColPrcsdLastVal).setValue(Status[0]);
      shtRspn.getRange(RspnRow, ColErrorMsg).setValue(Status[1]);
      
      // Set the Matching Response Match ID if Matching Response found
      if (MatchingRspn > 0) shtRspn.getRange(MatchingRspn, ColMatchID).setValue(MatchID);	  
            
    }
    // When Week Number is empty or if the Response Data was processed, we have reached the end of the list, then exit the loop
    if(RspnWeekNum == '' || RspnDataPrcssd == 1) {
      Logger.log('Response Loop exit at Row: %s',RspnRow)
      RspnRow = RspnMaxRows + 1;
    }
  }
  // Execute Ranking function in Standing tab
  fcnUpdateStandings(ss);
  
  // Copy all data to League Spreadsheet
  fcnCopyStandingsResults(ss, shtConfig);

}


