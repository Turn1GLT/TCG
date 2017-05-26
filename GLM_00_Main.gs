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
  var ConfigData = shtConfig.getRange(3, 9, 20, 1).getValues();
  var LeagueName = shtConfig.getRange(2, 2).getValue();
  
  // Code Execution Options
  var OptDualSubmission = ConfigData[0][0]; // If Dual Submission is disabled, look for duplicate instead
  var OptPostResult = ConfigData[1][0];
  var OptPlyrMatchValidation = ConfigData[2][0];
  var OptTCGBooster = ConfigData[3][0];
  
  // Columns Values and Parameters
  var ColMatchID = ConfigData[8][0];
  var ColPrcsd = ConfigData[9][0];
  var ColDataConflict = ConfigData[10][0];
  var ColErrorMsg = ConfigData[11][0];
  var ColPrcsdLastVal = ConfigData[12][0];
  var ColMatchIDLastVal = ConfigData[13][0];
  var RspnStartRow = ConfigData[14][0];
  var RspnDataInputs = ConfigData[15][0]; // from Time Stamp to Data Processed
  var NbCards = ConfigData[16][0];

  // Test Sheet (for Debug)
  var shtTest = ss.getSheetByName('Test') ; 
  
  // Form Responses Sheet Variables
  var shtRspn = ss.getSheetByName('Form Responses 13');
  var RspnMaxRows = shtRspn.getMaxRows();
  var RspnMaxCols = shtRspn.getMaxColumns();
  var RspnNextRowPrcss = shtRspn.getRange(1, ColPrcsdLastVal).getValue() + 1;
  var RspnWeekNum;
  var RspnDataWeek;
  var RspnDataWinr;
  var RspnDataLosr;
  var RspnDataPrcssd = 0;
  var ResponseData;
  var MatchingRspnData;
  
  // Card List Variables
  var CardList = new Array(16);         // 0 = Set Name, 1-14 = Card Numbers, 15 = Card 14 is Masterpiece (Y-N)
  var CardName;
  var PackData = new Array(16); // 0 = Set Name, 1-14 = Card Numbers, 15 = Card 14 is Masterpiece (Y-N)
  
  // Create Array of 16x4 where each row is Card 1-14 and each column is Card Info
  for(var cardnum = 0; cardnum < 16; cardnum++){
    PackData[cardnum] = new Array(4); // 0= Card in Pack, 1= Card Number, 2= Card Name, 3= Card Rarity
    for (var val = 0; val < 4; val++) PackData[cardnum][val] = '';
  }

  // Match Data Variables
  var MatchID; 
  var MatchData = new Array(25); // 0 = MatchID, 1 = Week #, 2 = Winning Player, 3 = Losing Player, 4 = Score, 5 = Winner Points, 6 = Loser Points, 7 = Card Set, 8-21 = Cards, 22 = Masterpiece (Y-N), 23 = Reserved, 24 = MatchPostStatus
  // Create Array of 25x4 where each row is Card 1-14 and each column is Card Info. This Info is only used for rows 8-21
  for(var cardnum = 0; cardnum < 25; cardnum++){
    MatchData[cardnum] = new Array(4); // 0= Card in Pack, 1= Card Number, 2= Card Name, 3= Card Rarity
    for (var val = 0; val < 4; val++) MatchData[cardnum][val] = '';
  }
  
  
  // Email Addresses
  var Addresses = new Array(2);

  // Data Processing Flags
  var StatusMsg = '';
  var DuplicateRspn = -99;
  var MatchingRspn = -98;
  var MatchPostStatus = -97;
  var CardDBUpdated = -96;
  
  Logger.log('Start of Main Function Executed');
  
  Logger.log('Dual Submission Option: %s',OptDualSubmission);
  Logger.log('Post Results Option: %s',OptPostResult);
  Logger.log('Player Match Validation Option: %s',OptPlyrMatchValidation);
  Logger.log('TCG Option: %s',OptTCGBooster);
  
  // Find a Row that is not processed in the Response Sheet (added data)
  for (var RspnRow = RspnNextRowPrcss; RspnRow <= RspnMaxRows; RspnRow++){
    
    // Copy the new response data (from Time Stamp to Data Processed Field
    ResponseData = shtRspn.getRange(RspnRow, 1, 1, RspnDataInputs).getValues();
    
    RspnWeekNum = ResponseData[0][1];
    RspnDataPrcssd = ResponseData[0][23];
    RspnDataWinr  = ResponseData[0][2]; // Winning Player
    RspnDataLosr  = ResponseData[0][3]; // Losing Player 
    
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
              MatchData = fcnPostMatchResults(ss, ConfigData, shtRspn, ResponseData, MatchingRspnData, MatchID, shtTest);
              MatchPostStatus = MatchData[24][0];
              
              Logger.log('Match Post Status: %s',MatchPostStatus);
              
              // If Match was populated in Match Results Tab
              if (MatchPostStatus == 1){
                // Match ID doesn't change because we assumed it was already OK
                
                // Copies all cards added to the Card Database
                if (OptTCGBooster == 'Enabled'){
                  for (var card = 0; card < NbCards; card++){
                    CardList[card] = ResponseData[0][card+5];
                  }
                  if (CardList[0] != 'No Pack Opened') {
                    PackData = fcnUpdateCardDB(RspnDataLosr, CardList, shtTest);
                    // Copy all card names to Match Data [7-22]
                    for (var card = 0; card < NbCards; card++){
                      MatchData[card+7][0] = PackData[card][0]; // Card in Pack
                      MatchData[card+7][1] = PackData[card][1]; // Card Number
                      MatchData[card+7][2] = PackData[card][2]; // Card Name
                      MatchData[card+7][3] = PackData[card][3]; // Card Rarity
                      
                      if (PackData[card][2] == 'Card Name not Found for Card Number') {
                        StatusMsg = 'Card Name not Found for Card Number: ' + CardList[card]; 
                      }
                    }
                  }
                  // for debug
                  shtTest.getRange(20,1,25,4).setValues(MatchData);
                }
                  
                // Send email Confirmation that Response and Entry Data was compiled and posted to the Match Results
                // Get Email addresses from Config File
                Addresses = subGetEmailAddress(shtConfig, RspnDataWinr, RspnDataLosr);
                // Call the Email Function, sends Match Data
                fcnGenEmailConfirmation(LeagueName, Addresses, MatchData);
              }
              
              // If MatchPostSuccess = 0, function was executed but was not able to post in the Match Result Tab
              if (MatchPostStatus < 0){
                // Updates the Match ID to an empty value 
                MatchID = '';
                // Generate the Status Message
                StatusMsg = subGenErrorMsg(MatchPostStatus);
                // Get Email addresses from Config File
                
                // Call the Email Function, sends Match Data
                
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
            StatusMsg = 'Waiting for Other Response Submission'
            // Set the Data Processed Flag
            RspnDataPrcssd = 1;
          } 
          
          // If MatchingEntry = -1, fcnFindMatchingEntry was not executed properly, sends email to notify
          if (OptDualSubmission == 'Enabled' && MatchingRspn < 0){
            // Set the Status Message
            StatusMsg = subGenErrorMsg(MatchingRspn);
            
            // Get Email addresses from Config File
            
            // Call the Email Function, sends Match Data
            
          }
          
        }
        
        // If Duplicate is found, send email to notify, set Response Data Processed to -1 to represent the Duplicate Entry
        if (DuplicateRspn > 0){
          
          // Updates the Match ID to an empty value 
          MatchID = '';
          
          // Set the Data Processed Flag
          RspnDataPrcssd = 1;        
          
          // Sets the Status Message
          StatusMsg = 'Duplicate Entry Found at Row: ' + DuplicateRspn;
                
          // Get Email addresses from Config File
          
          // Call the Email Function, sends Match Data to Organizer
        }
        
        // If FindDuplicateEntry was not executed properly, send email to notify, set Response Data Processed to -2 to represent processing error
        if (DuplicateRspn < 0){
          
          // Updates the Match ID to an empty value 
          MatchID = '';
          
          // Set the Data Processed Flag
          RspnDataPrcssd = 1;  
          
          // Set the Status Message
          StatusMsg = subGenErrorMsg(DuplicateRspn);
 
          // Get Email addresses from Config File
          
          // Call the Email Function, sends Match Data
        }
      } 
      
      // If Both Players are the same, report error
      if (RspnDataWinr == RspnDataLosr){
        
        // Updates the Match ID to an empty value 
        MatchID = '';
        
        // Set the Data Processed Flag
        RspnDataPrcssd = 1;  
        
        // Set the Status Message
        StatusMsg = 'Same Player selected for Win and Loss'; 
        // Get Email addresses from Config File
        
        // Call the Email Function, sends Match Data
      }
      
      // Set the Match ID (for both Response and Matching Entry), and Updates the Last Match ID generated, 
      if (MatchPostStatus == 1 || OptPostResult == 'Disabled'){
        shtRspn.getRange(RspnRow, ColMatchID).setValue(MatchID);
        shtRspn.getRange(1, ColMatchIDLastVal).setValue(MatchID);
      }
      // Set the Processed Flag and Status Message for the response
      shtRspn.getRange(RspnRow, ColPrcsd).setValue(RspnDataPrcssd);
      shtRspn.getRange(RspnRow, ColErrorMsg).setValue(StatusMsg);
      
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
  fcnUpdateStandings(ss, shtConfig);
  
  // Copy all data to League Spreadsheet
  fcnCopyStandingsResults(ss);

}


