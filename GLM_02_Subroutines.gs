// **********************************************
// function subCheckDataConflict()
//
// This function verifies that two arrays of data 
// are the same. If two values are different,
// the function returns the Data ID where they
// differ. If no conflict is found, returns 0;
//
// **********************************************

function subCheckDataConflict(DataArray1, DataArray2, ColStart, ColEnd, shtTest) {
  
  var DataConflict = 0;
  
  // Compare New Response Data and Match Data. If Data is not equal to the other
  for (var j = ColStart; j <= ColEnd; j++){
    
    shtTest.getRange(j+30,1).setValue(DataArray1[0][j]);
    shtTest.getRange(j+30,2).setValue(DataArray2[0][j]);
        
    // If Data Conflict is found, sets the data and sends email
    if (DataArray1[0][j] != DataArray2[0][j]) {
      DataConflict = j;
      shtTest.getRange(j+30,3).setValue('Conflict Detected');
      j = ColEnd + 1;
    }
  }
  return DataConflict;
}

// **********************************************
// function subPlayerMatchValidation()
//
// This function verifies that the player was allowed 
// to play this match. It checks in the total amount of matches
// played by the player to allow the game to be posted
// The function returns 1 if the game is valid and 0 if not valid
//
// **********************************************

function subPlayerMatchValidation(ss, PlayerName, shtTest) {
  
  Logger.log('%s Match Validation executed', PlayerName);
  
  // Opens Cumulative Results tab
  var shtCumul = ss.getSheetByName('Cumulative Results');
    
  // Get Data from Cumulative Results
  var CumulMaxMatch = shtCumul.getRange(3,3).getValue();
  var CumulPlyrData = shtCumul.getRange(5, 1, 32, 9).getValues(); // Data[i][j] i = Player List 1-32, j = ID(0), Name(1), Initials(2), MP(3), W(4), L(5), %(6), Packs(7), Status(8)
  
  var PlayerStatus;
  var PlayerMatchPlayed;
  
  var MatchValid = 0; // Match is invalid by default
  
  // Look for Player Row and if Player is still Active or Eliminated
  for (var i = 0; i < 32; i++) {
    // Player Found, Number of Match Played and Status memorized
    if (PlayerName == CumulPlyrData[i][1]){
      PlayerMatchPlayed = CumulPlyrData[i][3];
      PlayerStatus = CumulPlyrData[i][8];
      //Logger.log('Player Name: %s / MP: %s / Status: %s',CumulPlyrData[i][0], CumulPlyrData[i][3], CumulPlyrData[i][8]);
      i = 32; // Exit Loop
    }
  }

  // If Player is Active and Number of Matches Playes is below or equal to the maximum permitted
  if (PlayerStatus == 'Active' && PlayerMatchPlayed + 1 <= CumulMaxMatch) MatchValid = 1;
  
  // If Player is Eliminated, return -1
  if (PlayerStatus == 'Eliminated') MatchValid = -1;
  
  // If Player has played more games (counting the one to be posted) than permitted, return -2
  if (PlayerMatchPlayed + 1 > CumulMaxMatch && PlayerStatus != 'Eliminated') MatchValid = -2;
  
  return MatchValid;
}

// **********************************************
// function subGenErrorMsg()
//
// This function generates the Error Message according to 
// the value sent in argument
//
// **********************************************

function subGenErrorMsg(ErrorVal) {
  
  var ErrorMsg;
  
  switch (ErrorVal){

    case -99 : ErrorMsg = 'Duplicate Entry Search Not Executed'; break;    
    case -98 : ErrorMsg = 'Matching Response Search Not Executed'; break;  
    case -97 : ErrorMsg = 'Match Results Post Not Executed'; break;  
    case -11 : ErrorMsg = 'Winning Player is Eliminated from League'; break;  
    case -12 : ErrorMsg = 'Winning Player has played too many matches'; break;  
    case -21 : ErrorMsg = 'Losing Player is Eliminated from League'; break;  
    case -22 : ErrorMsg = 'Losing Player has played too many matches'; break;  
    case -31 : ErrorMsg = 'Both Players are Eliminated from the League'; break;  
    case -32 : ErrorMsg = 'Winning Player is Eliminated from the League and Losing Player has played too many matches'; break;
    case -33 : ErrorMsg = 'Winning Player has player too many matches and Losing Player is Eliminated from the League'; break;
    case -34 : ErrorMsg = 'Both Players have played too many matches'; break;
//    case  : ErrorMsg = ''; break;  // Add Error Message for Data Conflict on Dual Submission
//    case  : ErrorMsg = ''; break;
//    case  : ErrorMsg = ''; break;
//    case  : ErrorMsg = ''; break;

}
  
return ErrorMsg;
}

