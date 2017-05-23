// **********************************************
// function fcnUpdateCardDB()
//
// This function updates the Player card database  
// with the list of cards sent in arguments
//
// **********************************************

function fcnUpdateCardDB(Player, CardList, TestSht){
  
  // Config Spreadsheet
  var ShtConfig = SpreadsheetApp.openById('14rR_7-SG9fTi-M7fpS7d6n4XrOlnbKxRW1Ni2ongUVU').getSheetByName('Config');
  var CardDBShtID = ShtConfig.getRange(3, 2).getValue();
  
  // Player Card DB Spreadsheet
  var shtCardDB = SpreadsheetApp.openById(CardDBShtID).getSheetByName(Player);
  var CardDBSet = shtCardDB.getRange(6,1,1,32).getValues();
  var MstrSet = shtCardDB.getRange(6,33,1,16).getValues();
  var SetNum;
  var CardName;
  var CardID;
  var CardCol = 0;
  var CardQty;
  var MstrCol = 0;
  var CardListSet = CardList[0];
  var UpdateCardDBStatus = new Array(16); // 0 = Set, 1-14 = Card Numbers, 15 = Masterpiece
  
  Logger.log('Card Set: %s',CardListSet);
  
  UpdateCardDBStatus[0] = 1;
  // Find Set Column according to Set in Cardlist (CardList[0]) and get all card quantities (first card starts at row 8, row 7 = card 0)
  for (var SetCol = 0; SetCol <= 31; SetCol++){   
    if (CardListSet == CardDBSet[0][SetCol]){
      CardCol = SetCol+1;
      Logger.log('Card Set Column: %s',CardCol);
      SetCol = 32;
    }
  }
  
  // Loop through each card in CardList and update quantity
  for (var CardListNb = 1; CardListNb <= 14; CardListNb++){
    // Get Card ID Number 
    CardID = CardList[CardListNb];
    
    // Regular cards and non Masterpiece card
    if (CardListNb < 14 || (CardListNb == 14 && CardList[15] == 'No')){
      //
      CardName = shtCardDB.getRange(CardID+7, CardCol).getValue();
      // If Card Name exists, set status to 1 and update card quantity
      if (CardName != ''){
        UpdateCardDBStatus[CardListNb] = 1;
        CardQty = shtCardDB.getRange(CardID+7, CardCol-2).getValue() + 1;
        shtCardDB.getRange(CardID+7, CardCol-2).setValue(CardQty);
      }
      // If Card Name does not exist, set status to 0
      if (CardName == ''){
        UpdateCardDBStatus[CardListNb] = 0;
      }
      //TestSht.getRange(CardListNb,3).setValue(UpdateCardDBStatus[CardListNb]);
    }
    
    // If Last card is a Masterpiece
    if (CardListNb == 14 && CardList[15] == 'Yes'){
      // Get Set Number to find Masterpiece Column
      SetNum = shtCardDB.getRange(4, CardCol).getValue();
      Logger.log('Masterpiece Set Number: %s',SetNum);
      // Set Masterpiece Column according to Set Number
      switch (SetNum){
        case 1 : MstrCol= 35; break;
        case 2 : MstrCol= 35; break;
        case 3 : MstrCol= 39; break;
        case 4 : MstrCol= 39; break;
        case 5 : MstrCol= 43; break;
        case 6 : MstrCol= 43; break;
        case 7 : MstrCol= 47; break;
        case 8 : MstrCol= 47; break;
      }
      CardQty = shtCardDB.getRange(CardID+7, MstrCol-2).getValue() + 1;
      shtCardDB.getRange(CardID+7, MstrCol-2).setValue(CardQty);
    }
  }
  
  
  // Call function to generate clean card pool from Player Card DB
  fcnUpdateCardPool(shtCardDB, Player, TestSht);
  
  // Return Value
  return UpdateCardDBStatus;
}