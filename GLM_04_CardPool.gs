// **********************************************
// function fcnUpdateCardPool()
//
// This function updates the Player card database  
// with the list of cards sent in arguments
//
// **********************************************

function fcnUpdateCardPool(shtCardDB, Player, shtTest){
  
  // Config Spreadsheet
  var ShtConfig = SpreadsheetApp.openById('14rR_7-SG9fTi-M7fpS7d6n4XrOlnbKxRW1Ni2ongUVU').getSheetByName('Config');
  var ssCardPoolID = ShtConfig.getRange(59,2).getValue();
  
  // Card Pool Spreadsheet
  var shtCardPool = SpreadsheetApp.openById(ssCardPoolID).getSheetByName(Player);
  var rngCardPool = shtCardPool.getRange(6, 1, 224, 5); // 0 = Card Qty, 1 = Card Number, 2 = Card Name, 3 = Rarity, 4 = Set Name 
  var CardPool; // Where Card Data will be populated
  
  var CardDBSetTotal = shtCardDB.getRange(2,1,1,48).getValues(); // Gets Sum of all set Quantity, if > 0, set is present in card pool
  var SetData;
  var SetName;
  var colSet;
  var CardNb = 0;
    
  // Clear Player Card Pool
  rngCardPool.clearContent();
  CardPool = rngCardPool.getValues();
    
  // Look for Set with cards present in pool
  for (var col = 0; col <= 48; col++){   
    // if Set Card Quantity > 0, Set has cards in pool, Loop through all cards in Set
    if (CardDBSetTotal[0][col] > 0){
      colSet = col + 1;
      SetName = shtCardDB.getRange(6,colSet+2).getValue();

      // Get all Cards Data from set
      SetData = shtCardDB.getRange(7, colSet, 286, 4).getValues();

      // Loop through each card in Set and get Card Data
      for (var CardID = 1; CardID <= 285; CardID++){
        if (SetData[CardID][0] > 0) {
          CardPool[CardNb][0] = SetData[CardID][0]; // Quantity
          CardPool[CardNb][1] = SetData[CardID][1]; // Card Number (ID)
          CardPool[CardNb][2] = SetData[CardID][2]; // Card Name
          CardPool[CardNb][3] = SetData[CardID][3]; // Card Rarity
          CardPool[CardNb][4] = SetName;            // Set Name    
          CardNb++;
//          shtTest.getRange(CardNb,1).setValue(CardPool[CardNb][0]);
//          shtTest.getRange(CardNb,2).setValue(CardPool[CardNb][1]);
//          shtTest.getRange(CardNb,3).setValue(CardPool[CardNb][2]);
//          shtTest.getRange(CardNb,4).setValue(CardPool[CardNb][3]);
//          shtTest.getRange(CardNb,5).setValue(CardPool[CardNb][4]);
        }
      }
    }
  }
  // Updates the Player Card Pool
  rngCardPool.setValues(CardPool);
  
  // Return Value
}