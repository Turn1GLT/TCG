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

function fcnFindMatchingEntry(ss, RspnSht, ResponseData, RspnRow, RspnStartRow, RspnMaxRows, RspnDataInputs) {

  // Columns Values
  var ColMatchID = 24;
  var ColPrcsd = 25;
  var ColErrorCode = 26;
  var ColDataConflict = 27;
  var ColPrcsdLastVal = 28;
  var ColMatchIDLastVal = 29;
  
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
          
          TestSht.getRange(RspnRow +10, 1).setValue(RspnDataWeek);
          TestSht.getRange(RspnRow +10, 2).setValue(RspnDataWinr);
          TestSht.getRange(RspnRow +10, 3).setValue(RspnDataLosr); 
          TestSht.getRange(RspnRow +10, 4).setValue(EntryWeek);
          TestSht.getRange(RspnRow +10, 5).setValue(EntryWinr);
          TestSht.getRange(RspnRow +10, 6).setValue(EntryLosr);

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

function fcnFindDuplicateEntry(ss, RspnSht, ResponseData, RspnRow, RspnStartRow, RspnMaxRows, RspnDataInputs) {

  // Columns Values
  var ColMatchID = 24;
  var ColPrcsd = 25;
  var ColErrorCode = 26;
  var ColDataConflict = 27;
  var ColPrcsdLastVal = 28;
  var ColMatchIDLastVal = 29;
  
  var RspnDataWeek;
  var RspnDataWinr;
  var RspnDataLosr;

  var EntryWeek;
  var EntryWinr;
  var EntryLosr;
  var EntryData;
  var EntryPrcssd;
  var EntryMatchID;
  
  var DuplicateRow = 0;
  
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
    
    // If both rows are different, the Data Entry was processed and was compiled in the Match Results (MatchID != '') and Week Number are equal), Look for player entry combination
    if (EntryRow != RspnRow && EntryPrcssd == 1 && EntryMatchID != '' && RspnDataWeek == EntryWeek){
      // If combination of players are the same between the entry data and the new response data, duplicate entry was found. Save Row index
      if ((RspnDataWinr == EntryWinr && RspnDataLosr == EntryLosr) || (RspnDataWinr == EntryLosr && RspnDataLosr == EntryWinr)){
        DuplicateRow = EntryRow;
        Logger.log('Duplicate entry found at row: %s', DuplicateRow)
        EntryRow = RspnMaxRows + 1;
      }
    }
    return DuplicateRow;
  }
}

// **********************************************
// function fcnPopMatchResults()
//
// Once both players have submitted their forms 
// the data in the Game Results tab are copied into
// the Week X tab
//
// **********************************************

function fcnPopMatchResults(ss, RspnSht, ResponseData, EntryData, MatchID, OptDualSubmission) {
  
  // Match Results Sheet Variables
  var RsltSht = ss.getSheetByName('Match Results');
  var RsltShtMaxRows = RsltSht.getMaxRows();
  var RsltShtMaxCol = RsltSht.getMaxColumns();
  var RsltLastResultRow = RsltSht.getRange(3, 4).getValue() + 1;
  var RsltRng = RsltSht.getRange(RsltLastResultRow, 1, 1, RsltShtMaxCol);
  var ResultData = RsltRng.getValues();
  var MatchValidWinr;
  var MatchValidLosr;
  var RsltPlyrDataA;
  var RsltPlyrDataB;
  
  var MatchResultPopulated = -1;
  
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
  ResultData[0][1]  = ResponseData[0][23]; // Match ID
  ResultData[0][2]  = ResponseData[0][1]; // Week Number
  ResultData[0][3]  = ResponseData[0][2]; // Winning Player
  ResultData[0][4]  = ResponseData[0][3]; // Losing Player  
  
  // Validate if players are allowed to post results (look for number of games played versus total amount of games allowed
  MatchValidWinr = 
  MatchValidLosr = 
  
  
  // Copies Result Data
  // ResultData[0][0] = Result ID 
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
  MatchResultPopulated = 1;
  
  return MatchResultPopulated;
                                   
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
// ResultData[0][6]  = Winner Score - NOT USED
// ResultData[0][7]  = Loser Score - NOT USED
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



