// **********************************************
// function subCheckDataConflict()
//
// This function verifies that two arrays of data 
// are the same. If two values are different,
// the function returns the Data ID where they
// differ. If no conflict is found, returns 0;
//
// **********************************************

function subCheckDataConflict(DataArray1, DataArray2, ColStart, ColEnd, TestSht) {
  
  var DataConflict = 0;
  
  // Compare New Response Data and Match Data. If Data is not equal to the other
  for (var j = ColStart; j <= ColEnd; j++){
    
    TestSht.getRange(j+30,1).setValue(DataArray1[0][j]);
    TestSht.getRange(j+30,2).setValue(DataArray2[0][j]);
        
    // If Data Conflict is found, sets the data and sends email
    if (DataArray1[0][j] != DataArray2[0][j]) {
      DataConflict = j;
      TestSht.getRange(j+30,3).setValue('Conflict Detected');
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

function subPlayerMatchValidation(ss, PlayerName, CurrentWeek, TestSht) {
  
  // Opens Cumulative Results tab
  var CumulSht = ss.getSheetByName('Cumulative Results');
    
  // Get Data
  var MaxMatch = CumulSht.getRange(3,3).getValue();
  var PlayerData = CumulSht.getRange(5, 2, 32, 3).getValues();
  
  var MatchValid = 0;
  
  for (var CumulRow = 5; CumulRow <= 32; CumulRow++){
    // Enter validation here
  }

  return MatchValid;
}

