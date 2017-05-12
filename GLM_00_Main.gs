

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
  var ConfigSht = SpreadsheetApp.openById('14rR_7-SG9fTi-M7fpS7d6n4XrOlnbKxRW1Ni2ongUVU').getSheetByName('Config');
  var ConfigData = ConfigSht.getRange(3, 9, 20, 1).getValues();
  
  // Code Execution Options
  var OptDualSubmission = ConfigData[0][0]; // If Dual Submission is disabled, look for duplicate instead
  var OptPostResult = ConfigData[1][0];
  var OptPlyrMatchValidation = ConfigData[2][0];
  
  // Columns Values and Parameters
  var ColMatchID = ConfigData[8][0];
  var ColPrcsd = ConfigData[9][0];
  var ColDataConflict = ConfigData[10][0];
  var ColErrorMsg = ConfigData[11][0];
  var ColPrcsdLastVal = ConfigData[12][0];
  var ColMatchIDLastVal = ConfigData[13][0];
  var RspnStartRow = ConfigData[14][0];
  var RspnDataInputs = ConfigData[15][0]; // from Time Stamp to Data Processed

  // Test Sheet (for Debug)
  var TestSht = ss.getSheetByName('Test') ; 
  
  // Form Responses Sheet Variables
  var RspnSht = ss.getSheetByName('Form Responses 13');
  var RspnMaxRows = RspnSht.getMaxRows();
  var RspnMaxCols = RspnSht.getMaxColumns();
  var RspnNextRowPrcss = RspnSht.getRange(1, ColPrcsdLastVal).getValue() + 1;
  var RspnWeekNum;
  var RspnDataWeek;
  var RspnDataWinr;
  var RspnDataLosr;
  var RspnDataPrcssd = 0;
  var ResponseData;
  var MatchingRspnData;
  
  var MatchID; 
  var ErrorMsg = '';

  // Data Processing Flags
  var DuplicateRspn = -1;
  var MatchingRspn = -1;
  var MatchPostStatus = -1;
  
  Logger.log('Start of Main Function Executed');
  
  Logger.log('Dual Submission Option: %s',OptDualSubmission);
  Logger.log('Post Results Option: %s',OptPostResult);
  Logger.log('Player Match Validation Option: %s',OptPlyrMatchValidation);
  
  // Find a Row that is not processed in the Response Sheet (added data)
  for (var RspnRow = RspnNextRowPrcss; RspnRow <= RspnMaxRows; RspnRow++){
    
    // Copy the new response data (from Time Stamp to Data Processed Field
    ResponseData = RspnSht.getRange(RspnRow, 1, 1, RspnDataInputs).getValues();
    
    RspnWeekNum = ResponseData[0][1];
    RspnDataPrcssd = ResponseData[0][23];
    RspnDataWinr  = ResponseData[0][2]; // Winning Player
    RspnDataLosr  = ResponseData[0][3]; // Losing Player 
    
    // If week number is not empty and Processed is empty and both players are different, Response Data needs to be processed
    if (RspnWeekNum != '' && RspnDataPrcssd == ''){
      
      // If both Players in the response are different, continue
      if (RspnDataWinr != RspnDataLosr){
        
        // Generates the Match ID in advance if data analysis is successful
        MatchID = RspnSht.getRange(1, ColMatchIDLastVal).getValue() + 1;
        
        Logger.log('New Data Found at Row: %s',RspnRow);
        
        // Copy the new response data to Data Array
        ResponseData = RspnSht.getRange(RspnRow, 1, 1, RspnDataInputs).getValues();
        
        // Look for Duplicate Entry (looks in all entries with MatchID and combination of Week Number, Winner and Loser) 
        // Real code will look at Player Posting Data as well
        DuplicateRspn = fcnFindDuplicateResponse(ss, ConfigData, RspnSht, ResponseData, RspnRow, RspnStartRow, RspnMaxRows, RspnDataInputs);  
        
        Logger.log('Duplicate Result: %s', DuplicateRspn);
        
        // FindDuplicateEntry function was executed properly and didn't find any Duplicate entry, continue analyzing the response data
        if (DuplicateRspn == 0){
          
          // If Dual Submission is enabled, Search if the other Entry matching this response has been submitted (must be enabled)
          if (OptDualSubmission == 'Enabled'){
            // function returns row where the matching data was found
            MatchingRspn = fcnFindMatchingResponse(ss, ConfigData, RspnSht, ResponseData, RspnRow, RspnStartRow, RspnMaxRows, RspnDataInputs);
          }
          
          // Search if the other Entry matching this response has been submitted
          if (OptDualSubmission == 'Disabled'){
            MatchingRspn = RspnRow;
          }      
          
          Logger.log('Matching Result: %s', MatchingRspn);
          
          // If the result of the fcnFindMatchingEntry function returns something different than -1 and 0, we found a matching entry, continue analyzing the response data
          if (MatchingRspn != -1 && MatchingRspn != 0){
            
            if (OptPostResult == 'Enabled'){
              
              // Get the Entry Data found at row MatchingRspn
              MatchingRspnData = RspnSht.getRange(MatchingRspn, 1, 1, RspnDataInputs).getValues();
              
              // Execute function to populate Match Result Sheet from processed data
              MatchPostStatus = fcnPostMatchResults(ss, ConfigData, RspnSht, ResponseData, MatchingRspnData, MatchID, TestSht);
              Logger.log('Match Post Status: %s',MatchPostStatus);
              
              // If Match was populated in Match Results Tab
              if (MatchPostStatus == 1){
                // Match ID doesn't change because we assumed it was already OK
                
                // Send email Confirmation that Response and Entry Data was compiled and posted to the Match Results
                
              }
              
              // If MatchPostSuccess = 0, function was executed but was not able to post in the Match Result Tab
              if (MatchPostStatus == 0){
                // Updates the Match ID to an empty value 
                MatchID = '';
                // Set the Error Message
                ErrorMsg = 'Not Able to Post Results';
              }
              
              // If MatchPostSuccess = -1, function was not executed properly, sends email to notify
              if (MatchPostStatus == -1){
                // Updates the Match ID to an empty value 
                MatchID = '';
                // Set the Error Message
                ErrorMsg = 'Match Post Not Executed';
                // Get email from Config File
                
                // Call the Email Function, sends Both Response and Entry Data 
                
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
            // Set the Data Processed Flag
            RspnDataPrcssd = 1;
          } 
          
          // If MatchingEntry = -1, fcnFindMatchingEntry was not executed properly, sends email to notify
          if (OptDualSubmission == 'Enabled' && MatchingRspn == -1){
            // Set the Error Message
            ErrorMsg = 'Matching Response Search Not Executed';
            
            // Get email from Config File
            
            // Call the Email Function, sends Both Response and Entry Data 
            
          }
          
        }
        
        // If Duplicate is found, send email to notify, set Response Data Processed to -1 to represent the Duplicate Entry
        if (DuplicateRspn != 0 && DuplicateRspn != -1){
          
          // Updates the Match ID to an empty value 
          MatchID = '';
          
          // Set the Data Processed Flag
          RspnDataPrcssd = 1;        
          
          // Sets the Error Message
          ErrorMsg = 'Duplicate Entry Found at Row: ' + DuplicateRspn;
          
          Logger.log('Duplicate Found');        
          // Get email from Config File
          
          // Call the Email Function, sends Both Response and Entry Data to Organizer
        }
        
        // If FindDuplicateEntry was not executed properly, send email to notify, set Response Data Processed to -2 to represent processing error
        if (DuplicateRspn == -1){
          
          // Updates the Match ID to an empty value 
          MatchID = '';
          
          // Set the Data Processed Flag
          RspnDataPrcssd = 1;  
          
          // Set the Error Message
          ErrorMsg = 'Duplicate Entry Search Not Executed';
          
          Logger.log('Duplicate Not Executed');  
          // Get email from Config File
          
          // Call the Email Function, sends Both Response and Entry Data 
        }
      } 
      
      // If Both Players are the same, report error
      if (RspnDataWinr == RspnDataLosr){
        
        // Updates the Match ID to an empty value 
        MatchID = '';
        
        // Set the Data Processed Flag
        RspnDataPrcssd = 1;  
        
        // Set the Error Message
        ErrorMsg = 'Same Player selected for Win and Loss';
        
        Logger.log('Same Player selected for Win and Loss');  
        // Get email from Config File
        
        // Call the Email Function, sends Both Response and Entry Data 
      }
      
      // Set the Match ID (for both Response and Matching Entry),  and Updates the Last Match ID generated, 
      if (MatchPostStatus == 1 || OptPostResult == 'Disabled'){
        RspnSht.getRange(RspnRow, ColMatchID).setValue(MatchID);
        RspnSht.getRange(1, ColMatchIDLastVal).setValue(MatchID);
      }
      // Set the Processed Flag and Error Message for the response
      RspnSht.getRange(RspnRow, ColPrcsd).setValue(RspnDataPrcssd);
      RspnSht.getRange(RspnRow, ColErrorMsg).setValue(ErrorMsg);
      
      // Set the Matching Response Match ID if Matching Response found
      if (MatchingRspn > 0) RspnSht.getRange(MatchingRspn, ColMatchID).setValue(MatchID);	  
      
    }
    // When Week Number is empty or if the Response Data was processed, we have reached the end of the list, then exit the loop
    if(RspnWeekNum == '' || RspnDataPrcssd == 1) {
      Logger.log('Response Loop exit at Row: %s',RspnRow)
      RspnRow = RspnMaxRows + 1;
    }
  }
  // Execute Ranking function in Standing tab
  fcnUpdateStandings(ss, ConfigSht);
}


