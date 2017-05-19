// **********************************************
// function fcnFindDuplicateData()
//
// This function searches the entry list to find any 
// duplicate responses. To make sure we do not interfere 
// with the fcnFindMatchingData, we look for a non-zero Match ID
//
// The functions returns the Row number where the matching data was found. 
// 
// If no duplicate data is found, it returns 0;
//
// **********************************************

function fcnFindDuplicateData(ss, ConfigData, ShtRspn, ResponseData, RspnRow, RspnStartRow, RspnMaxRows, RspnDataInputs) {

  // Columns Values and Parameters
  var ColMatchID = ConfigData[8][0];
  var ColPrcsd = ConfigData[9][0];
  var ColDataConflict = ConfigData[10][0];
  var ColErrorMsg = ConfigData[11][0];
  var ColPrcsdLastVal = ConfigData[12][0];
  var ColMatchIDLastVal = ConfigData[13][0];
  var RspnStartRow = ConfigData[14][0];
  var RspnDataInputs = ConfigData[15][0]; // from Time Stamp to Data Processed
  
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
  
  var EntryWeekData = ShtRspn.getRange(RspnStartRow, 2, RspnMaxRows-3,1).getValues();
    
  // Loop to find if the other player posted the game results
  for (var EntryRow = RspnStartRow; EntryRow <= RspnMaxRows; EntryRow++){
    
    // Filters only entries of the same week the response was posted
    if (EntryWeekData[EntryRow][0] == RspnDataWeek){
      
      // Gets Entry Data to analyze
      EntryData = ShtRspn.getRange(EntryRow, 1, 1, RspnDataInputs).getValues();
      
      EntryWeek = EntryData[0][1];
      EntryWinr = EntryData[0][2];
      EntryLosr = EntryData[0][3];
      EntryMatchID = EntryData[0][22];
      EntryPrcssd = EntryData[0][23];
      
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
// function fcnFindMatchingData()
//
// This function searches for the other match entry 
// in the Response Tab. The functions returns the Row number
// where the matching data was found. 
// 
// If no matching data is found, it returns 0;
//
// **********************************************

function fcnFindMatchingData(ss, ConfigData, ShtRspn, ResponseData, RspnRow, RspnStartRow, RspnMaxRows, RspnDataInputs) {

  // Columns Values and Parameters
  var ColMatchID = ConfigData[8][0];
  var ColPrcsd = ConfigData[9][0];
  var ColDataConflict = ConfigData[10][0];
  var ColErrorMsg = ConfigData[11][0];
  var ColPrcsdLastVal = ConfigData[12][0];
  var ColMatchIDLastVal = ConfigData[13][0];
  var RspnStartRow = ConfigData[14][0];
  var RspnDataInputs = ConfigData[15][0]; // from Time Stamp to Data Processed
  
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
        EntryData = ShtRspn.getRange(EntryRow, 1, 1, RspnDataInputs).getValues();

        EntryWeek = EntryData[0][1];
        EntryWinr = EntryData[0][2];
        EntryLosr = EntryData[0][3];
        EntryMatchID = EntryData[0][22];
        EntryPrcssd = EntryData[0][23];

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
            ShtRspn.getRange(RspnRow, ColDataConflict).setValue('No Conflict');
            ShtRspn.getRange(EntryRow, ColDataConflict).setValue('No Conflict');
            MatchingRow = EntryRow;
          }
          
          // If Data Conflict was detected, sends email to notify Data Conflict
          if (DataConflict != 0 && DataConflict != -1){

            // Sets the Conflict Value to the Data ID value where the conflict was found
            ShtRspn.getRange(RspnRow, ColDataConflict).setValue(DataConflict);
            ShtRspn.getRange(EntryRow, ColDataConflict).setValue(DataConflict);
          }
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
// function fcnPostMatchResults()
//
// Once both players have submitted their forms 
// the data in the Game Results tab are copied into
// the Week X tab
//
// **********************************************

function fcnPostMatchResults(ss, ConfigData, ShtRspn, ResponseData, MatchingRspnData, MatchID, TestSht) {
  
  // Code Execution Options
  var OptDualSubmission = ConfigData[0][0]; // If Dual Submission is disabled, look for duplicate instead
  var OptPostResult = ConfigData[1][0];
  var OptPlyrMatchValidation = ConfigData[2][0];
  var OptGameTCG = ConfigData[3][0];
  
  // Match Results Sheet Variables
  var RsltSht = ss.getSheetByName('Match Results');
  var RsltShtMaxRows = RsltSht.getMaxRows();
  var RsltShtMaxCol = RsltSht.getMaxColumns();
  var RsltLastResultRowRng = RsltSht.getRange(3, 4);
  var RsltLastResultRow = RsltLastResultRowRng.getValue() + 1;
  var RsltRng = RsltSht.getRange(RsltLastResultRow, 1, 1, RsltShtMaxCol);
  var ResultData = RsltRng.getValues();
  var MatchValidWinr = '';
  var MatchValidLosr = '';
  var RsltPlyrDataA;
  var RsltPlyrDataB;
  
  var MatchPostedStatus = 0;
  
  // Sets which Data set is Player A and Player B. Data[0][1] = Player who posted the data
  if (OptDualSubmission == 'Enabled' && ResponseData[0][1] == ResponseData[0][2]) {
    RsltPlyrDataA = ResponseData;
    RsltPlyrDataB = MatchingRspnData;
  }
  
  if (OptDualSubmission == 'Enabled' && ResponseData[0][1] == ResponseData[0][3]) {
    RsltPlyrDataA = ResponseData;
    RsltPlyrDataB = MatchingRspnData;
  }
  
  // Copies Players Data

  ResultData[0][2]  = ResponseData[0][1]; // Week/Round Number
  ResultData[0][3]  = ResponseData[0][2]; // Winning Player
  ResultData[0][4]  = ResponseData[0][3]; // Losing Player  
  
  // If option is enabled, Validate if players are allowed to post results (look for number of games played versus total amount of games allowed
  if (OptPlyrMatchValidation == 'Enabled'){
    

    MatchValidWinr = subPlayerMatchValidation(ss, ResultData[0][3], TestSht);
    Logger.log('%s Match Validation: %s',ResultData[0][3], MatchValidWinr);
    MatchValidLosr = subPlayerMatchValidation(ss, ResultData[0][4], TestSht);
    Logger.log('%s Match Validation: %s',ResultData[0][4], MatchValidLosr);
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
    if (OptGameTCG == 'Enabled'){
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
      ResultData[0][23] = ResponseData[0][20]; // Card 15 (Special Foil) 
    }
    
    // Sets Data in Match Result Tab
    RsltRng.setValues(ResultData);
    
    // Update the 
    MatchPostedStatus = 1;
    RsltLastResultRowRng.setValue(RsltLastResultRow);
    
    // Post Results in Appropriate Week Number for Both Players
    fcnPostResultWeek(ss, ConfigData, ResultData, TestSht);
  }
  
  // If Match Validation was not successful, generate Error Status
  
  // returns Error that Winning Player is Eliminated from the League
  if (MatchValidWinr == -1 && MatchValidLosr == 1)  MatchPostedStatus = -11;
  
  // returns Error that Winning Player has played too many matches
  if (MatchValidWinr == -2 && MatchValidLosr == 1)  MatchPostedStatus = -12;  
  
  // returns Error that Losing Player is Eliminated from the League
  if (MatchValidLosr == -1 && MatchValidWinr == 1)  MatchPostedStatus = -21;
  
  // returns Error that Losing Player has played too many matches
  if (MatchValidLosr == -2 && MatchValidWinr == 1)  MatchPostedStatus = -22;
  
  // returns Error that Both Players are Eliminated from the League
  if (MatchValidWinr == -1 && MatchValidLosr == -1) MatchPostedStatus = -31;
  
  // returns Error that Winning Player is Eliminated from the League and Losing Player has played too many matches
  if (MatchValidWinr == -1 && MatchValidLosr == -2) MatchPostedStatus = -32;

  // returns Error that Winning Player has player too many matches and Losing Player is Eliminated from the League
  if (MatchValidWinr == -2 && MatchValidLosr == -1) MatchPostedStatus = -33;
  
  // returns Error that Both Players have played too many matches
  if (MatchValidWinr == -2 && MatchValidLosr == -2) MatchPostedStatus = -34;
  
  return MatchPostedStatus;
}


// **********************************************
// function fcnPostResultWeek()
//
// Once the Match Data has been posted in the 
// Match Results Tab, the Week X results are updated
// for each player
//
// **********************************************

function fcnPostResultWeek(ss, ConfigData, ResultData, TestSht) {

  // Code Execution Options
  var OptGameTCG = ConfigData[3][0];
  
  // function variables
  var ShtWeekRslt;
  var ShtWeekRsltRng;
  var ShtWeekPlyr;
  var ShtWeekWinrRec;
  var ShtWeekLosrRec;
  var ShtWeekPackData
  var ShtWeekMaxCol;
  var ShtWeekPlyr
  
  var ColPlyr = 2;
  var ColWin = 5;
  var ColLos = 6;
  var PackLength = 16;
  var NextPackID = 0;
  
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

  // Gets All Players Names
  ShtWeekPlyr = ShtWeekRslt.getRange(5,ColPlyr,32,1).getValues();
  
  // Find the Winning and Losing Player in the Week Result Tab
  for (var RsltRow = 5; RsltRow <= 36; RsltRow ++){
    
    if (ShtWeekPlyr[RsltRow - 5][0] == MatchDataWinr) WeekWinrRow = RsltRow;
    if (ShtWeekPlyr[RsltRow - 5][0] == MatchDataLosr) WeekLosrRow = RsltRow;
    
    if (WeekWinrRow != '' && WeekLosrRow != '') {
      // Get Winner and Loser Match Record 
      ShtWeekWinrRec = ShtWeekRslt.getRange(WeekWinrRow,5,1,2).getValues();
      ShtWeekLosrRec = ShtWeekRslt.getRange(WeekLosrRow,5,1,2).getValues();
      
      // If Game Type is TCG
      if (OptGameTCG == 'Enabled'){
      // Get Loser Pack Data
      ShtWeekPackData = ShtWeekRslt.getRange(WeekLosrRow,8,1,(PackLength*6)+1).getValues();
      }
      RsltRow = 37;
    }
  }
  
  // Update Winning Player Results
  ShtWeekWinrRec[0][0] = ShtWeekWinrRec[0][0] + 1;
  if (ShtWeekWinrRec[0][1] == '') ShtWeekWinrRec[0][1] = 0;  
  
  // Update Losing Player Results
  ShtWeekLosrRec[0][1] = ShtWeekLosrRec[0][1] + 1;
  if (ShtWeekLosrRec[0][0] == '') ShtWeekLosrRec[0][0] = 0;  
  
  // Update the Week Results Sheet
  ShtWeekRslt.getRange(WeekWinrRow,5,1,2).setValues(ShtWeekWinrRec);
  ShtWeekRslt.getRange(WeekLosrRow,5,1,2).setValues(ShtWeekLosrRec);
  
  // If Game Type is TCG and Punishment Pack has been opened, update Punishment Pack Info
  if (OptGameTCG == 'Enabled' && ResultData[0][8] != ''){
      
    // Find the next free Punishment Pack space offset
    if (ShtWeekPackData[0][1]  == '' && NextPackID == 0) NextPackID = 1;
    if (ShtWeekPackData[0][17] == '' && NextPackID == 0) NextPackID = 17;
    if (ShtWeekPackData[0][33] == '' && NextPackID == 0) NextPackID = 33;
    if (ShtWeekPackData[0][49] == '' && NextPackID == 0) NextPackID = 49;
    if (ShtWeekPackData[0][65] == '' && NextPackID == 0) NextPackID = 65;
    if (ShtWeekPackData[0][81] == '' && NextPackID == 0) NextPackID = 81;
    
    ShtWeekPackData[0][0] = ShtWeekPackData[0][0] + 1;
    // Update the Pack data
    for (var PackDataID = 0; PackDataID < PackLength; PackDataID++){
      ShtWeekPackData[0][PackDataID + NextPackID] = ResultData[0][PackDataID + 8];
    }
    // Update the Week Results Sheet with the Pack Info
    ShtWeekRslt.getRange(WeekLosrRow,8,1,(PackLength*6)+1).setValues(ShtWeekPackData);
  }
}

// **********************************************
// function fcnUpdateStandings()
//
// Updates the Standings according to the Win % 
// from the Cumulative Results tab to the Standings Tab
//
// **********************************************

function fcnUpdateStandings(ss){

  var ShtCumul = ss.getSheetByName('Cumulative Results');
  var ShtStand = ss.getSheetByName('Standings');
  
  // Get Player Record Range
  var RngCumul = ShtCumul.getRange(5,2,32,6);
  var RngStand = ShtStand.getRange(6,2,32,6);
  
  // Get Cumulative Results Values and puts them in the Standings Values
  var ValCumul = RngCumul.getValues();
  RngStand.setValues(ValCumul);

  // Sorts the Standings Values by Win % (column 7) and Matches Played (column 4)
  RngStand.sort([{column: 7, ascending: false},{column: 4, ascending: false}]);

}



















