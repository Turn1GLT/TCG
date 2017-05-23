// **********************************************
// function fcnUpdateCardPool()
//
// This function updates the Player card database  
// with the list of cards sent in arguments
//
// **********************************************

function fcnUpdateCardPool(shtCardDB, Player, TestSht){
  
  // Config Spreadsheet
  var ShtConfig = SpreadsheetApp.openById('14rR_7-SG9fTi-M7fpS7d6n4XrOlnbKxRW1Ni2ongUVU').getSheetByName('Config');
  var CardPoolShtID = ShtConfig.getRange(4, 2).getValue();
  
  // Card Pool Spreadsheet
  var shtCardPool = SpreadsheetApp.openById(CardPoolShtID).getSheetByName(Player);
  var rngCardPool = shtCardPool.getRange(6, 1, 224, 5); // 0 = Card Qty, 1 = Card Number, 2 = Card Name, 3 = Rarity, 4 = Set Name 
  var CardPool; // Where Card Data will be populated
  
  var CardDBSetTotal = shtCardDB.getRange(2,1,1,48).getValues(); // Gets Sum of all set Quantity, if > 0, set is present in card pool
  var SetData;
  var SetName;
  var colSet;
  var CardNb = 0;
  
  Logger.log('Update Card Pool for Player %s', Player);
  
  // Clear Player Card Pool
  rngCardPool.clearContent();
  CardPool = rngCardPool.getValues();
    
  // Look for Set with cards present in pool
  for (var col = 0; col <= 48; col++){   
    // if Set Card Quantity > 0, Set has cards in pool, Loop through all cards in Set
    if (CardDBSetTotal[0][col] > 0){
      colSet = col + 1;
      SetName = shtCardDB.getRange(6,colSet+2).getValue();
      Logger.log('Set Name found for Card Pool: %s',SetName);
      // Get all Cards Data from set
      SetData = shtCardDB.getRange(7, colSet, 286, 4).getValues();
      //TestSht.getRange(1,8,286,4).setValues(SetData);
      // Loop through each card in Set and get Card Data
      for (var CardID = 1; CardID <= 285; CardID++){
        if (SetData[CardID][0] > 0) {
          CardPool[CardNb][0] = SetData[CardID][0]; // Quantity
          CardPool[CardNb][1] = SetData[CardID][1]; // Card Number (ID)
          CardPool[CardNb][2] = SetData[CardID][2]; // Card Name
          CardPool[CardNb][3] = SetData[CardID][3]; // Card Rarity
          CardPool[CardNb][4] = SetName;            // Set Name    
          CardNb++;
          TestSht.getRange(CardNb,5).setValue(CardPool[CardNb][0]);
          TestSht.getRange(CardNb,6).setValue(CardPool[CardNb][1]);
          TestSht.getRange(CardNb,7).setValue(CardPool[CardNb][2]);
          TestSht.getRange(CardNb,8).setValue(CardPool[CardNb][3]);
          TestSht.getRange(CardNb,9).setValue(CardPool[CardNb][4]);
        }
      }
    }
  }
  // Updates the Player Card Pool
  rngCardPool.setValues(CardPool);
  
  // Return Value
}