// **********************************************
// function fcnUpdateCardPool()
//
// This function updates the Player card database  
// with the list of cards sent in arguments
//
// **********************************************

function fcnUpdateCardPool(shtConfig, shtCardDB, Player, shtTest){
  
  // Config Spreadsheet
  var ssCardPoolEnID = shtConfig.getRange(32,2).getValue();
  var ssCardPoolFrID = shtConfig.getRange(33,2).getValue();
  
  // Card Pool Spreadsheet
  var shtCardPoolEn = SpreadsheetApp.openById(ssCardPoolEnID).getSheetByName(Player);
  var CardPoolEnMaxRows = shtCardPoolEn.getMaxRows();
  var shtCardPoolFr = SpreadsheetApp.openById(ssCardPoolFrID).getSheetByName(Player);
  var CardPoolFrMaxRows = shtCardPoolFr.getMaxRows();
  var rngCardPoolEn = shtCardPoolEn.getRange(6, 1, CardPoolEnMaxRows-6, 5); // 0 = Card Qty, 1 = Card Number, 2 = Card Name, 3 = Rarity, 4 = Set Name 
  var rngCardPoolFr = shtCardPoolFr.getRange(6, 1, CardPoolFrMaxRows-6, 5); // 0 = Card Qty, 1 = Card Number, 2 = Card Name, 3 = Rarity, 4 = Set Name 
  var CardPool; // Where Card Data will be populated
  
  var CardDBSetTotal = shtCardDB.getRange(2,1,1,48).getValues(); // Gets Sum of all set Quantity, if > 0, set is present in card pool
  var CardTotal = shtCardDB.getRange(3,7).getValue();
  var SetData;
  var SetName;
  var colSet;
  var CardNb = 0;
    
  // Clear Player Card Pool
  rngCardPoolEn.clearContent();
  CardPool = rngCardPoolEn.getValues();
  rngCardPoolFr.clearContent();
  CardPool = rngCardPoolFr.getValues();
    
  // Look for Set with cards present in pool
  for (var col = 0; col <= 48; col++){   
    // if Set Card Quantity > 0, Set has cards in pool, Loop through all cards in Set
    if (CardDBSetTotal[0][col] > 0){
      colSet = col + 1;
      SetName = shtCardDB.getRange(6,colSet+2).getValue();

      // Get all Cards Data from set
      SetData = shtCardDB.getRange(7, colSet, 300, 4).getValues();
      
      // Loop through each card in Set and get Card Data
      for (var CardID = 1; CardID <= 299; CardID++){
        Logger.log('CardID:%s',CardID);
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
  rngCardPoolEn.setValues(CardPool);
  shtCardPoolEn.getRange(3,1).setValue(CardTotal);
  rngCardPoolFr.setValues(CardPool);
  shtCardPoolFr.getRange(3,1).setValue(CardTotal);
  
  
  // Return Value
}
