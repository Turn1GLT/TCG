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
  var OptDualSubmission = ConfigSht.getRange(3, 9).getValue(); // If Dual Submission is disabled, look for duplicate insteadS
  var OptPostResult = ConfigSht.getRange(4, 9).getValue();
  
  // Test Sheet (for Debug)
  var TestSht = ss.getSheetByName('Test') ; 
  
  // Columns Values and Parameters
  var ColMatchID = 24;
  var ColPrcsd = 25;
  var ColErrorCode = 26;
  var ColDataConflict = 27;
  var ColPrcsdLastVal = 28;
  var ColMatchIDLastVal = 29;
  var RspnStartRow = 2;
  var RspnDataInputs = 25; // from Time Stamp to Data Processed

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
  
  var MatchID; 
  var LastPrcssdRow;

  // Data Processing Flags
  var MatchingEntryFound = -1;
  var DuplicateEntry = -1;
  var MatchPostSuccess = -1;
  
  Logger.log('Start New Data Loop: %s',RspnNextRowPrcss);
  
  // Find a Row that is not processed in the Response Sheet (added data)
  for (var RspnRow = RspnNextRowPrcss; RspnRow <= RspnMaxRows; RspnRow++){
    
    // Copy the new response data (from Time Stamp to Data Processed Field
    ResponseData = RspnSht.getRange(RspnRow, 1, 1, RspnDataInputs).getValues();
    
    RspnWeekNum = ResponseData[0][1];
    RspnDataPrcssd = ResponseData[0][24];
      
    // If week number is not empty and Processed is empty, Match Data needs to be processed
    if (RspnWeekNum != '' && RspnDataPrcssd == ''){
      
      // Generates the Match ID in advance if data analysis is successful
      MatchID = RspnSht.getRange(1, ColMatchIDLastVal).getValue() + 1;
               
      Logger.log('New Data Found at Row: %s',RspnRow);
                 
      // Copy the new response data to Data Array
      ResponseData = RspnSht.getRange(RspnRow, 1, 1, RspnDataInputs).getValues();
      
      // Look for Duplicate Entry (looks in all entries with MatchID and combination of Week Number, Winner and Loser) 
      // Real code will look at Player Posting Data as well
      DuplicateEntry = fcnFindDuplicateEntry(ss, RspnSht, ResponseData, RspnRow, RspnStartRow, RspnMaxRows, RspnDataInputs);  
      
      // FindDuplicateEntry function was executed properly and didn't find any Duplicate entry, continue analyzing the response data
      if (DuplicateEntry == 0){
      
        // If Dual Submission is enabled, Search if the other Entry matching this response has been submitted (must be enabled)
        if (OptDualSubmission == 'Enabled'){
          // function returns row where the matching data was found
          MatchingEntryFound = fcnFindMatchingEntry(ss, RspnSht, ResponseData, RspnRow, RspnStartRow, RspnMaxRows, RspnDataInputs);
        }
        
        // Search if the other Entry matching this response has been submitted
        if (OptDualSubmission == 'Disabled'){
          MatchingEntryFound = RspnRow;
          RspnSht.getRange(RspnRow, ColMatchID).setValue(MatchID);
        }      
        
        // If the result of the fcnFindMatchingEntry function returns something different than -1 and 0, we found a matching entry, continue analyzing the response data
        if (MatchingEntryFound != -1 && MatchingEntryFound != 0){
          
          // Set Match ID
          RspnSht.getRange(RspnRow, ColMatchID).setValue(MatchID);
          RspnSht.getRange(MatchingEntryFound, ColMatchID).setValue(MatchID);
          
          if (OptPostResult == 'Enabled'){
            
            // Execute function to populate Match Result Sheet from processed data
            MatchPostSuccess = fcnPopMatchResults(ss, RspnSht, ResponseData, MatchID);
            
            // If Match was populated in Match Results Tab
            if (MatchPostSuccess == 1){
              // Updates the Last Match ID generated
              RspnSht.getRange(1, ColMatchIDLastVal).setValue(MatchID);
              // Send email Confirmation that Response and Entry Data was compiled and posted to the Match Results
              
            }
            
            // If MatchPostSuccess = 0, function was executed but was not able to post in the Match Result Tab
            if (MatchPostSuccess == 0){
              // Updates the Error Message Column
              RspnSht.getRange(RspnRow, ColErrorCode).setValue('Not Able to Post Results');
            }
            
            // If MatchPostSuccess = -1, function was not executed properly, sends email to notify
            if (MatchPostSuccess == -1){
              // Updates the Error Message Column
              RspnSht.getRange(RspnRow, ColErrorCode).setValue('Match Post Not Executed');
              // Get email from Config File
              
              // Call the Email Function, sends Both Response and Entry Data 
              
            }
          }
          // If Posting is disabled, generate Match ID for testing        
          if (OptPostResult == 'Disabled'){
            // Update the Last Match ID generated
            RspnSht.getRange(1, ColMatchIDLastVal).setValue(MatchID);
          }
          
          Logger.log('Matching Entry Found: %s',MatchingEntryFound);
          // Set the Data Processed Flag
          RspnDataPrcssd = 1;
          LastPrcssdRow = RspnRow;
        }
        
        // If MatchingEntry = 0, fcnFindMatchingEntry did not find a matching entry, it might be the first response entry
        if (MatchingEntryFound == 0){
         // Set the Data Processed Flag
          RspnDataPrcssd = 1;
          LastPrcssdRow = RspnRow;
        } 
        
        // If MatchingEntry = -1, fcnFindMatchingEntry was not executed properly, sends email to notify
        if (MatchingEntryFound == -1){
          // Updates the Error Message Column
          RspnSht.getRange(RspnRow, ColErrorCode).setValue('Matching Entry Search Not Executed');
          
          // Get email from Config File
          
          // Call the Email Function, sends Both Response and Entry Data 
          
        }
        // When Week Number is empty or if the Response Data was processed, we have reached the end of the list, then exit the loop
        if(RspnWeekNum == '' || RspnDataPrcssd == 1) {
          RspnRow = RspnMaxRows + 1;
        }
      }
      
      // If Duplicate is found, send email to notify, set Response Data Processed to -1 to represent the Duplicate Entry
      if (DuplicateEntry != 0 && DuplicateEntry != -1){
        // Updates the Error Message Column
        RspnSht.getRange(RspnRow, ColErrorCode).setValue('Duplicate Entry Found at Row: ' &DuplicateEntry);
        
        // Get email from Config File
        
        // Call the Email Function, sends Both Response and Entry Data 
        
        // Set the Data Processed Flag
        RspnDataPrcssd = -1;
        LastPrcssdRow = RspnRow;
      }
      
      // If FindDuplicateEntry was not executed properly, send email to notify, set Response Data Processed to -2 to represent processing error
      if (DuplicateEntry == -1){
        
        // Updates the Error Message Column
        RspnSht.getRange(RspnRow, ColErrorCode).setValue('Duplicate Entry Search Not Executed');
        
        // Get email from Config File
        
        // Call the Email Function, sends Both Response and Entry Data 
        
        // Set the Data Processed Flag
        RspnDataPrcssd = -2;
        LastPrcssdRow = RspnRow;
      }
      
    }
    // Set the Processed Flag for that Entry in the spreadsheet and Updates the Last Data Processed
    RspnSht.getRange(RspnRow, ColPrcsd).setValue(RspnDataPrcssd);
    RspnSht.getRange(1, ColPrcsdLastVal).setValue(LastPrcssdRow);
  }
}

