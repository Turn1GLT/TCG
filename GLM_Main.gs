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
  
  // Config Sheet
  var ConfigSht = ss.getSheetByName('Config');
  // Test Sheet (for Debug)
  var TestSht = ss.getSheetByName('Test') ; 
  
  // Columns Values
  var ColMatchID = 24;
  var ColPrcsd = 25;
  var ColDataConflict = 26;
  var ColPrcsdLastVal = 27;
  var ColMatchIDLastVal = 28;

  // Form Responses Sheet Variables
  var RspnSht = ss.getSheetByName('Form Responses 13');
  var RspnMaxRows = RspnSht.getMaxRows();
  var RspnMaxCols = RspnSht.getMaxColumns();
  var RspnStartRow = 2;
  var RspnNextRowPrcss = RspnSht.getRange(1, ColPrcsdLastVal).getValue() + 1;
  var RspnWeekNum;
  var RspnDataWeek;
  var RspnDataWinr;
  var RspnDataLosr;
  var RspnDataInputs = 25; // from Time Stamp to Data Processed
  var RspnDataPrcssd = 0;
  var ResponseData;

  var MatchingEntryFound = -1;
  var DuplicateEntryFound = -1;
  var MatchID;  
  var MatchPostSuccess = -1;
  
  // Options
  var OptDualSubmitEnabled = 1; // If Dual Submission is disabled, look for duplicate insteadS
  var OptDuplDetectEnabled = 0;
  var OptPostResultEnabled = 0;

  Logger.log('Start New Data Loop: %s',RspnNextRowPrcss);
  
  // Find a Row that is not processed in the Response Sheet (added data)
  for (var RspnRow = RspnNextRowPrcss; RspnRow <= RspnMaxRows; RspnRow++){
    
    // Copy the new response data
    ResponseData = RspnSht.getRange(RspnRow, 1, 1, RspnDataInputs).getValues();
    
    RspnWeekNum = ResponseData[0][1];
    RspnDataPrcssd = ResponseData[0][24];
      
    // If week number is not empty and Processed is empty, Match Data needs to be processed
    if (RspnWeekNum != '' && RspnDataPrcssd == ''){
      
      // Generates the Match ID if data analysis is successful
      MatchID = RspnSht.getRange(1, ColMatchIDLastVal).getValue() + 1;
               
      Logger.log('New Data Found at Row: %s',RspnRow);
                 
      // Copy the new response data
      ResponseData = RspnSht.getRange(RspnRow, 1, 1, RspnDataInputs).getValues();
      
      // Looks for Duplicate Entry (looks for MatchID and combination of Week Number, Winner and Loser)
      //DuplicateEntryFound = fcnFindDuplicateEntry(ss, RspnSht, ResponseData, RspnRow, RspnStartRow, RspnMaxRows, RspnDataInputs);
      
      // If Dual Submission is enabled, Search if the other Entry matching this response has been submitted (must be enabled)
      if (OptDualSubmitEnabled == 1){
        // function returns row where the matching data was found
        MatchingEntryFound = fcnFindMatchingEntry(ss, RspnSht, ResponseData, RspnRow, RspnStartRow, RspnMaxRows, RspnDataInputs);
      }

      // Search if the other Entry matching this response has been submitted
      if (OptDualSubmitEnabled == 0){
        MatchingEntryFound = RspnRow;
        RspnSht.getRange(RspnRow, ColMatchID).setValue(MatchID);
      }      
      
      // If the result of the fcnFindMatchingEntry function returns something different than -1 and 0, we found a matching entry
      if (MatchingEntryFound != -1 && MatchingEntryFound != 0){
        
        // Sets Match ID
        RspnSht.getRange(RspnRow, ColMatchID).setValue(MatchID);
        RspnSht.getRange(MatchingEntryFound, ColMatchID).setValue(MatchID);
        
        if (OptPostResultEnabled == 1){

          // Executes function to populate Match Result Sheet from processed data
          MatchPostSuccess = fcnPopMatchResults(ss, RspnSht, ResponseData, MatchID);
          
          // If Match was populated in Match Results Tab
          if (MatchPostSuccess == 1){
            // Updates the Last Match ID generated
            RspnSht.getRange(1, ColMatchIDLastVal).setValue(MatchID);
            // Sends email Confirmation that Response and Entry Data was compiled and posted to the Match Results
          }
          
          // If MatchPostSuccess = 0, function was executed but was not able to post in the Match Result Tab
          if (MatchPostSuccess == 0){
            
          }
          
          // If MatchPostSuccess = -1, function was not executed properly, sends email to notify
          if (MatchPostSuccess == -1){
            // Gets email from Config File
            
            // Calls the Email Function, sends Both Response and Entry Data 
           
          }
        }
        // If Posting is disabled, generate Match ID for testing        
        if (OptPostResultEnabled == 0){
          // Updates the Last Match ID generated
          RspnSht.getRange(1, ColMatchIDLastVal).setValue(MatchID);
        }

        Logger.log('Matching Entry Found: %s',MatchingEntryFound);
        // Sets the Response Data Processed in the Response sheet
        RspnDataPrcssd = 1;
      }

      // If MatchingEntry = 0, fcnFindMatchingEntry did not find a matching entry
      if (MatchingEntryFound == 0){
        RspnDataPrcssd = 1;
      } 
      
      // If MatchingEntry = -1, fcnFindMatchingEntry was not executed properly, sends email to notify
      if (MatchingEntryFound == -1){
        // Gets email from Config File
        
        // Calls the Email Function, sends Both Response and Entry Data 
                    
      }

      // Sets the Processed Flag for that Entry
      RspnSht.getRange(RspnRow, ColPrcsd).setValue(RspnDataPrcssd);
      
      // When Week Number is empty or if the Response Data was processed, we have reached the end of the list, then exit the loop
      if(RspnWeekNum == '' || RspnDataPrcssd == 1) RspnRow = RspnMaxRows + 1;
    }
  }
}

