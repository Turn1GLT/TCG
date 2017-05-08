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
  
  var DataConflict = -1;
  
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
  return DataConflict + 1;
}
