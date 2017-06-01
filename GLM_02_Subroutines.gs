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
    
//    shtTest.getRange(j+30,1).setValue(DataArray1[0][j]);
//    shtTest.getRange(j+30,2).setValue(DataArray2[0][j]);
        
    // If Data Conflict is found, sets the data and sends email
    if (DataArray1[0][j] != DataArray2[0][j]) {
      DataConflict = j;
//      shtTest.getRange(j+30,3).setValue('Conflict Detected');
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

function subPlayerMatchValidation(ss, PlayerName, MatchValidation, shtTest) {
  
  // Opens Cumulative Results tab
  var shtCumul = ss.getSheetByName('Cumulative Results');
    
  // Get Data from Cumulative Results
  var CumulMaxMatch = shtCumul.getRange(4,3).getValue();
  var CumulPlyrData = shtCumul.getRange(5, 1, 32, 9).getValues(); // Data[i][j] i = Player List 1-32, j = ID(0), Name(1), Initials(2), MP(3), W(4), L(5), %(6), Packs(7), Status(8)
  
  var PlayerStatus;
  var PlayerMatchPlayed;
  
  // Look for Player Row and if Player is still Active or Eliminated
  for (var i = 0; i < 32; i++) {
    // Player Found, Number of Match Played and Status memorized
    if (PlayerName == CumulPlyrData[i][1]){
      PlayerMatchPlayed = CumulPlyrData[i][3];
      PlayerStatus = CumulPlyrData[i][8];
      MatchValidation[1] = PlayerMatchPlayed;
      //Logger.log('Player Name: %s / MP: %s / Status: %s',CumulPlyrData[i][0], CumulPlyrData[i][3], CumulPlyrData[i][8]);
      i = 32; // Exit Loop
    }
  }

  // If Player is Active and Number of Matches Played is below or equal to the maximum permitted
  if (PlayerStatus == 'Active' && PlayerMatchPlayed + 1 <= CumulMaxMatch) MatchValidation[0] = 1;
  
  // If Player is Eliminated, return -1
  if (PlayerStatus == 'Eliminated') MatchValidation[0] = -1;
  
  // If Player has played more games (counting the one to be posted) than permitted, return -2
  if (PlayerMatchPlayed + 1 > CumulMaxMatch && PlayerStatus != 'Eliminated') MatchValidation[0] = -2;
  
  return MatchValidation;
}

// **********************************************
// function subGenErrorMsg()
//
// This function generates the Error Message according to 
// the value sent in argument
//
// **********************************************

function subGenErrorMsg(Status, ErrorVal,Param) {

  switch (ErrorVal){

    case -10 : Status[0] = ErrorVal; Status[1] = 'Duplicate Entry Found at Row ' + Param; break; 
    case -11 : Status[0] = ErrorVal; Status[1] = 'Winning Player is Eliminated from League'; break;  
    case -12 : Status[0] = ErrorVal; Status[1] = 'Winning Player has played too many matches'; break;  
    case -21 : Status[0] = ErrorVal; Status[1] = 'Losing Player is Eliminated from League'; break;  
    case -22 : Status[0] = ErrorVal; Status[1] = 'Losing Player has played too many matches'; break;  
    case -31 : Status[0] = ErrorVal; Status[1] = 'Both Players are Eliminated from the League'; break;  
    case -32 : Status[0] = ErrorVal; Status[1] = 'Winning Player is Eliminated from the League and Losing Player has played too many matches'; break;
    case -33 : Status[0] = ErrorVal; Status[1] = 'Winning Player has player too many matches and Losing Player is Eliminated from the League'; break;
    case -34 : Status[0] = ErrorVal; Status[1] = 'Both Players have played too many matches'; break;
    case -50 : Status[0] = ErrorVal; Status[1] = 'Illegal Match, Same Player selected for Win and Loss'; break; 
    case -60 : Status[0] = ErrorVal; Status[1] = 'Card Name not Found for Card Number: ' + Param; break;  // Card Name not Found
      
    case -97 : Status[0] = ErrorVal; Status[1] = 'Match Results Post Not Executed'; break;   
    case -98 : Status[0] = ErrorVal; Status[1] = 'Matching Response Search Not Executed'; break; 
    case -99 : Status[0] = ErrorVal; Status[1] = 'Duplicate Entry Search Not Executed'; break;    

//    case  : Status[0] = ErrorVal; Status[1] = ''; break;  // Add Error Message for Data Conflict on Dual Submission
//    case  : Status[0] = ErrorVal; Status[1] = ''; break;  // Add Error Message for Data Conflict on Dual Submission
//    case  : Status[0] = ErrorVal; Status[1] = ''; break;  // Add Error Message for Data Conflict on Dual Submission

}
  
return Status;
}

