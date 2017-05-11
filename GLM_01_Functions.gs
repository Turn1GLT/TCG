// **********************************************
// function fcnFindMatchingEntry()
//
// This function searches for the other match entry 
// in the Response Tab. The functions returns the Row number
// where the matching data was found. 
// 
// If no matching data is found, it returns 0;
//
// **********************************************

function fcnFindMatchingResponse(ss, RspnSht, ResponseData, RspnRow, RspnStartRow, RspnMaxRows, RspnDataInputs) {

  // Columns Values and Parameters
  var ColMatchID = 24;
  var ColPrcsd = 25;
  var ColDataConflict = 26;
  var ColErrorMsg = 27;
  var ColPrcsdLastVal = 28;
  var ColMatchIDLastVal = 29;
  var RspnStartRow = 2;
  var RspnDataInputs = 25; // from Time Stamp to Data Processed
  
  var RspnDataWeek;
  var RspnDataWinr;
  var RspnDataLosr;

  var EntryWeek;
  var EntryWinr;
  var EntryLosr;
  var EntryData;
  var EntryPrcssd;
  var EntryMatchID;
  
  var MatchingRow = 0;
  
  var DataConflict = -1;
  
  var TestSht = ss.getSheetByName('Test');
  
  // Loop to find if the other player posted the game results
      for (var EntryRow = RspnStartRow; EntryRow <= RspnMaxRows; EntryRow++){
        
        // Gets Entry Data to analyze
        EntryData = RspnSht.getRange(EntryRow, 1, 1, RspnDataInputs).getValues();

        EntryWeek = EntryData[0][1];
        EntryWinr = EntryData[0][2];
        EntryLosr = EntryData[0][3];
        EntryMatchID = EntryData[0][23];
        EntryPrcssd = EntryData[0][24];

        RspnDataWeek = ResponseData[0][1];
        RspnDataWinr = ResponseData[0][2];
        RspnDataLosr = ResponseData[0][3];
        
        // If both rows are different, Week Number, Player A and Player B are matching, we found the other match to compare data to
        if (EntryRow != RspnRow && EntryPrcssd == 1 && EntryMatchID == '' && RspnDataWeek == EntryWeek && RspnDataWinr == EntryWinr && RspnDataLosr == EntryLosr){
          
          // Compare New Response Data and Entry Data. If Data is not equal to the other, the conflicting Data ID is returned
          DataConflict = subCheckDataConflict(ResponseData, EntryData, 1, RspnDataInputs - 4, TestSht);
          
          // 
          if (DataConflict == 0){
            // Sets Conflict Flag to 'No Conflict'
            RspnSht.getRange(RspnRow, ColDataConflict).setValue('No Conflict');
            RspnSht.getRange(EntryRow, ColDataConflict).setValue('No Conflict');
            
            TestSht.getRange(RspnRow, 1).setValue('Matching Entry Found');
            TestSht.getRange(RspnRow, 2, 1, RspnDataInputs).setValues(ResponseData);
            TestSht.getRange(RspnRow +20, 1).setValue(EntryRow);
            TestSht.getRange(RspnRow +20, 2, 1, RspnDataInputs).setValues(EntryData);
            
            MatchingRow = EntryRow;
          }
          
          // If Data Conflict was detected, sends email to notify Data Conflict
          if (DataConflict != 0 && DataConflict != -1){

            // Sets the Conflict Value to the Data ID value where the conflict was found
            RspnSht.getRange(RspnRow, ColDataConflict).setValue(DataConflict);
            RspnSht.getRange(EntryRow, ColDataConflict).setValue(DataConflict);

            // Gets email from Config File
            
            // Calls the Email Function, sends Both Response and Entry Data and Conflicting Values and Category
            
          }
        }
        // Loop did not find matching data
        else{
          TestSht.getRange(RspnRow, 1).setValue('Matching Entry Not Found');
          TestSht.getRange(RspnRow, 2, 1, RspnDataInputs).setValues(ResponseData);
        }
        
        // Loop reached the end of responses entered or found matching data
        if(EntryWeek == '' || MatchingRow != 0) {
          Logger.log('Find Matching Loop Exits at Row %s',EntryRow);
          EntryRow = RspnMaxRows + 1;
        }
      }

  return MatchingRow;
}

// **********************************************
// function fcnFindMatchingEntry()
//
// This function searches for the other match entry 
// in the Response Tab. The functions returns the Row number
// where the matching data was found. 
// 
// If no matching data is found, it returns 0;
//
// **********************************************

function fcnFindDuplicateResponse(ss, RspnSht, ResponseData, RspnRow, RspnStartRow, RspnMaxRows, RspnDataInputs) {

  // Columns Values and Parameters
  var ColMatchID = 24;
  var ColPrcsd = 25;
  var ColDataConflict = 26;
  var ColErrorMsg = 27;
  var ColPrcsdLastVal = 28;
  var ColMatchIDLastVal = 29;
  var RspnStartRow = 2;
  var RspnDataInputs = 25; // from Time Stamp to Data Processed
  
  // Response Data 
  var RspnDataWeek = ResponseData[0][1];
  var RspnDataWinr = ResponseData[0][2];
  var RspnDataLosr = ResponseData[0][3];

  var EntryWeek;
  var EntryWinr;
  var EntryLosr;
  var EntryData;
  var EntryPrcssd;
  var EntryMatchID;
  
  var DuplicateRow = 0;
  
  var DataConflict = -1;
  
  var TestSht = ss.getSheetByName('Test');
  
  var EntryWeekData = RspnSht.getRange(RspnStartRow, 2, RspnMaxRows-3,1).getValues();
    
  // Loop to find if the other player posted the game results
  for (var EntryRow = RspnStartRow; EntryRow <= RspnMaxRows; EntryRow++){
    
    // Filters only entries of the same week the response was posted
    if (EntryWeekData[EntryRow][0] == RspnDataWeek){
      
      // Gets Entry Data to analyze
      EntryData = RspnSht.getRange(EntryRow, 1, 1, RspnDataInputs).getValues();
      
      EntryWeek = EntryData[0][1];
      EntryWinr = EntryData[0][2];
      EntryLosr = EntryData[0][3];
      EntryMatchID = EntryData[0][23];
      EntryPrcssd = EntryData[0][24];
           
      TestSht.getRange(EntryRow +10, 1).setValue(RspnRow);
      TestSht.getRange(EntryRow +10, 2).setValue(RspnDataWeek); 
      TestSht.getRange(EntryRow +10, 3).setValue(RspnDataWinr);
      TestSht.getRange(EntryRow +10, 4).setValue(RspnDataLosr);
      
      TestSht.getRange(EntryRow +10, 6).setValue(EntryRow);
      TestSht.getRange(EntryRow +10, 7).setValue(EntryWeek);
      TestSht.getRange(EntryRow +10, 8).setValue(EntryWinr);
      TestSht.getRange(EntryRow +10, 9).setValue(EntryLosr);
      
      // If both rows are different, the Data Entry was processed and was compiled in the Match Results (MatchID != '') and Week Number are equal), Look for player entry combination
      if (EntryRow != RspnRow && EntryPrcssd == 1 && EntryMatchID != ''){
        // If combination of players are the same between the entry data and the new response data, duplicate entry was found. Save Row index
        if ((RspnDataWinr == EntryWinr && RspnDataLosr == EntryLosr) || (RspnDataWinr == EntryLosr && RspnDataLosr == EntryWinr)){
          DuplicateRow = EntryRow;
          EntryRow = RspnMaxRows + 1;
        }
      }
    }
    
    // If we do not detect any value in Week Column, we reached the end of the list and skip
    if (EntryRow <= RspnMaxRows && EntryWeekData[EntryRow][0] == ''){
      EntryRow = RspnMaxRows + 1;
    }
  }
  return DuplicateRow;
}

// **********************************************
// function fcnPopMatchResults()
//
// Once both players have submitted their forms 
// the data in the Game Results tab are copied into
// the Week X tab
//
// **********************************************

function fcnPostMatchResults(ss, RspnSht, ResponseData, MatchingRspnData, MatchID, OptDualSubmission, OptPlyrMatchValidation, TestSht) {
  
  // Match Results Sheet Variables
  var RsltSht = ss.getSheetByName('Match Results');
  var RsltShtMaxRows = RsltSht.getMaxRows();
  var RsltShtMaxCol = RsltSht.getMaxColumns();
  var RsltLastResultRowRng = RsltSht.getRange(3, 4);
  var RsltLastResultRow = RsltLastResultRowRng.getValue() + 1;
  var RsltRng = RsltSht.getRange(RsltLastResultRow, 1, 1, RsltShtMaxCol);
  var ResultData = RsltRng.getValues();
  var MatchValidWinr;
  var MatchValidLosr;
  var RsltPlyrDataA;
  var RsltPlyrDataB;
  
  var MatchPostedStatus = 0;
  
  // Sets which Data set is Player A and Player B. Data[0][1] = Player who posted the data
  if (OptDualSubmission == 'Enabled' && ResponseData[0][1] == ResponseData[0][2]) {
    RsltPlyrDataA = ResponseData;
    RsltPlyrDataB = EntryData;
  }
  
  if (OptDualSubmission == 'Enabled' && ResponseData[0][1] == ResponseData[0][3]) {
    RsltPlyrDataA = ResponseData;
    RsltPlyrDataB = EntryData;
  }
  
  // Copies Players Data

  ResultData[0][2]  = ResponseData[0][1]; // Week Number
  ResultData[0][3]  = ResponseData[0][2]; // Winning Player
  ResultData[0][4]  = ResponseData[0][3]; // Losing Player  
  
  // If option is enabled, Validate if players are allowed to post results (look for number of games played versus total amount of games allowed
  if (OptPlyrMatchValidation == 'Enabled'){
    MatchValidWinr = subPlayerMatchValidation(ss, ResultData[0][3], ResultData[0][2], TestSht);
    MatchValidLosr = subPlayerMatchValidation(ss, ResultData[0][4], ResultData[0][2], TestSht);
  }

  // If option is disabled, Consider Matches are valid
  if (OptPlyrMatchValidation == 'Disabled'){
    MatchValidWinr = 1;
    MatchValidLosr = 1;
  }
  
  
  if (MatchValidWinr == 1 && MatchValidLosr == 1){
    // Copies Result Data
    // ResultData[0][0] = Result ID 
    ResultData[0][1]  = MatchID; // Match ID
    ResultData[0][5]  = ResponseData[0][4]; // Score
    ResultData[0][6]  = 2; // Winner Score
    if (ResponseData[0][4] == '2 - 0') ResultData[0][7]  = 0; // Loser Score
    if (ResponseData[0][4] == '2 - 1') ResultData[0][7]  = 1; // Loser Score
    
    
    // Copies Card Data
    ResultData[0][8]  = ResponseData[0][5];  // Expansion Set
    ResultData[0][9]  = ResponseData[0][6];  // Card 1
    ResultData[0][10] = ResponseData[0][7];  // Card 2
    ResultData[0][11] = ResponseData[0][8];  // Card 3
    ResultData[0][12] = ResponseData[0][9];  // Card 4
    ResultData[0][13] = ResponseData[0][10]; // Card 5
    ResultData[0][14] = ResponseData[0][11]; // Card 6
    ResultData[0][15] = ResponseData[0][12]; // Card 8
    ResultData[0][16] = ResponseData[0][13]; // Card 7
    ResultData[0][17] = ResponseData[0][14]; // Card 9
    ResultData[0][18] = ResponseData[0][15]; // Card 10
    ResultData[0][19] = ResponseData[0][16]; // Card 11
    ResultData[0][20] = ResponseData[0][17]; // Card 12
    ResultData[0][21] = ResponseData[0][18]; // Card 13
    ResultData[0][22] = ResponseData[0][19]; // Card 14
    ResultData[0][23] = ResponseData[0][20]; // Card 15 (Regular Foil)
    ResultData[0][24] = ResponseData[0][21]; // Card 16 (Special Foil)  
    
    // Sets Data in Match Result Tab
    RsltRng.setValues(ResultData);
    
    // Updates the 
    MatchPostedStatus = 1;
    RsltLastResultRowRng.setValue(RsltLastResultRow);
    
    // Post Results in Appropriate Week Number for Both Players
    fcnPostResultWeek(ss, ResultData, TestSht);
  }
  
  if (MatchValidWinr != 1 && MatchValidLosr == 1){
    // returns Error that Winning Player match is not legal
    MatchPostedStatus = -2;
  }
  
  if (MatchValidLosr != 1 && MatchValidWinr == 1){
    // returns Error that Losing Player match is not legal
    MatchPostedStatus = -3;
  }
  
  if (MatchValidWinr != 1 && MatchValidLosr != 1){
    // returns Error that Both Players match are not legal
    MatchPostedStatus = -4;
  }
  return MatchPostedStatus;
}

// Response and Entry Data Array

// ResponseData[0][0]  = Time Stamp
// ResponseData[0][1]  = Week Number
// ResponseData[0][2]  = Winning Player
// ResponseData[0][3]  = Losing Player
// ResponseData[0][4]  = Score
// ResponseData[0][5]  = Expansion Set
// ResponseData[0][6]  = Card 1
// ResponseData[0][7]  = Card 2
// ResponseData[0][8]  = Card 3
// ResponseData[0][9]  = Card 4
// ResponseData[0][10] = Card 5
// ResponseData[0][11] = Card 6
// ResponseData[0][12] = Card 7
// ResponseData[0][13] = Card 8
// ResponseData[0][14] = Card 9
// ResponseData[0][15] = Card 10
// ResponseData[0][16] = Card 11
// ResponseData[0][17] = Card 12
// ResponseData[0][18] = Card 13
// ResponseData[0][19] = Card 14
// ResponseData[0][20] = Card 15 (Regular Foil)
// ResponseData[0][21] = Card 16 (Special Foil)
// ResponseData[0][22] = Feedback
// ResponseData[0][23] = MatchID
// ResponseData[0][24] = Data Processed Status                               

// Result Data Array

// ResultData[0][0]  = Result ID
// ResultData[0][1]  = Match ID
// ResultData[0][2]  = Week Number
// ResultData[0][3]  = Winning Player
// ResultData[0][4]  = Losing Player
// ResultData[0][5]  = Score
// ResultData[0][6]  = Winner Score
// ResultData[0][7]  = Loser Score
// ResultData[0][8]  = Expansion Set
// ResultData[0][9]  = Card 1
// ResultData[0][10] = Card 2
// ResultData[0][11] = Card 3
// ResultData[0][12] = Card 4
// ResultData[0][13] = Card 5
// ResultData[0][14] = Card 6
// ResultData[0][15] = Card 8
// ResultData[0][16] = Card 7
// ResultData[0][17] = Card 9
// ResultData[0][18] = Card 10
// ResultData[0][19] = Card 11
// ResultData[0][20] = Card 12
// ResultData[0][21] = Card 13
// ResultData[0][22] = Card 14
// ResultData[0][23] = Card 15 (Regular Foil)
// ResultData[0][24] = Card 16 (Special Foil)

function fcnPostResultWeek(ss, ResultData, TestSht) {

  var ShtWeekRslt;
  var ShtWeekRsltRng;
  var ShtWeekPlyr;
  var ShtWeekWinrData;
  var ShtWeekLosrData;
  var ShtWeekMaxCol;
  var ShtWeekPlyr
  
  var ColPlyr = 1;
  var ColWin = 5;
  var ColLos = 6;
  var ColStartPack1 = 9;
  var ColStartPack2 = 26;
  var ColStartPack3 = 43;
  var ColStartPack4 = 60;
  var ColStartPack5 = 77;
  var ColStartPack6 = 94;
  var PackLength = 17;
  
  var WeekWinrRow = 0;
  var WeekLosrRow = 0;
  
  var MatchWeek = ResultData[0][2];
  var MatchDataWinr = ResultData[0][3];
  var MatchDataLosr = ResultData[0][4];
  
  // Selects the appropriate Week
  if (MatchWeek == 1) ShtWeekRslt = ss.getSheetByName('Week1');
  if (MatchWeek == 2) ShtWeekRslt = ss.getSheetByName('Week2');
  if (MatchWeek == 3) ShtWeekRslt = ss.getSheetByName('Week3');
  if (MatchWeek == 4) ShtWeekRslt = ss.getSheetByName('Week4');
  if (MatchWeek == 5) ShtWeekRslt = ss.getSheetByName('Week5');
  if (MatchWeek == 6) ShtWeekRslt = ss.getSheetByName('Week6');
  if (MatchWeek == 7) ShtWeekRslt = ss.getSheetByName('Week7');

  ShtWeekMaxCol = ShtWeekRslt.getMaxColumns();
  
  Logger.log('Match Week # %s',MatchWeek);
  Logger.log('Max Columns: %s',ShtWeekMaxCol);  
  Logger.log('Win Name:  %s',MatchDataWinr);
  Logger.log('Loss Name: %s',MatchDataLosr);
  

  // Gets All Players Names
  ShtWeekPlyr = ShtWeekRslt.getRange(5,ColPlyr,32,1).getValues();
  
  // Find the Winning and Losing Player in the Week Result Tab
  for (var RsltRow = 5; RsltRow <= 36; RsltRow ++){
    Logger.log('Player Name: %s',ShtWeekPlyr[RsltRow - 5][0]);
    
    if (ShtWeekPlyr[RsltRow - 5][0] == MatchDataWinr) {
      WeekWinrRow = RsltRow;
      Logger.log('Winner at Row: %s',WeekWinrRow);
    }
    if (ShtWeekPlyr[RsltRow - 5][0] == MatchDataLosr) {
      WeekLosrRow = RsltRow;
      Logger.log('Loser at Row: %s',WeekLosrRow);
    }
    
    if (WeekWinrRow != '' && WeekLosrRow != '') {
      // Get Winner and Loser Data 
      ShtWeekWinrData = ShtWeekRslt.getRange(WeekWinrRow,1,1,ShtWeekMaxCol).getValues();
      ShtWeekLosrData = ShtWeekRslt.getRange(WeekLosrRow,1,1,ShtWeekMaxCol).getValues();
      RsltRow = 37;
    }
  }
  
  // Post Winning Player Results

  
  // Post Losing Player Results
  
  
}

