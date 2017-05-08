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
  var ColDataConflict = 26;
  var ColPrcsdLastVal = 27;
  var ColMatchIDLastVal = 28;
  
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

        EntryPrcssd = EntryData[0][24];
        EntryMatchID = EntryData[0][23];
        EntryWeek = EntryData[0][1];
        EntryWinr = EntryData[0][2];
        EntryLosr = EntryData[0][3];

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
          if (DataConflict != 0){

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

}

// **********************************************
// function fcnPopMatchResults()
//
// Once both players have submitted their forms 
// the data in the Game Results tab are copied into
// the Week X tab
//
// **********************************************

function fcnPopMatchResults(ss,RspnSht,ResponseData,MatchData,MatchID) {
  
  // Match Results Sheet Variables
  var RsltSht = ss.getSheetByName('Match Results');
  var RsltShtMaxRows = RsltSht.getMaxRows();
  var RsltShtMaxCol = RsltSht.getMaxColumns();
  var RsltLastMatchRow = RsltSht.getRange(3, 4).getValue() + 1;
  var RsltArray = RsltSht.getRange(RsltLastMatchRow, 1, 1, RsltShtMaxCol);
  var RsltData = RsltArray.getValues();
  var RsltPlyrDataA;
  var RsltPlyrDataB;
  
  var MatchResultPopulated = -1;
  
  // Sets which Data set is Player A and Player B. Data[0][1] = Player who posted the data
  if (ResponseData[0][1] == ResponseData[0][2]) {
    RsltPlyrDataA = ResponseData;
    RsltPlyrDataB = MatchData;
  }
  
  if (ResponseData[0][1] == ResponseData[0][3]) {
    RsltPlyrDataA = ResponseData;
    RsltPlyrDataB = MatchData;
  }
  // Copies Common Data

  
  // Copies Sportsmanship Data and Comments
  
  
  // Copies Card Data
  
  // Sets Data in Match Result
                                   
  return MatchResultPopulated;
                                   
}

                                   
// ResponseData[0][0]  = Week Number
// ResponseData[0][1]  = Player
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

// MatchData[0][1] 	= Match ID
// MatchData[0][2] 	= Week Number
// MatchData[0][3] 	= Winning Player
// MatchData[0][4] 	= Losing Player
// MatchData[0][5] 	= Score
// MatchData[0][8] 	= Expansion Set
// MatchData[0][9] 	= Card 1
// MatchData[0][10] = Card 2
// MatchData[0][11] = Card 3
// MatchData[0][12] = Card 4
// MatchData[0][13] = Card 5
// MatchData[0][14] = Card 6
// MatchData[0][15] = Card 8
// MatchData[0][16] = Card 7
// MatchData[0][17] = Card 9
// MatchData[0][18] = Card 10
// MatchData[0][19] = Card 11
// MatchData[0][20] = Card 12
// MatchData[0][21] = Card 13
// MatchData[0][22] = Card 14
// MatchData[0][23] = Card 15 (Regular Foil)
// MatchData[0][24] = Card 16 (Special Foil)



