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
  var ShtCardDB = SpreadsheetApp.openById(CardDBShtID).getSheetByName(Player);
  var CardDBSet = ShtCardDB.getRange(6,1,1,24).getValues();
  var MstrSet = ShtCardDB.getRange(6,25,1,12).getValues();
  var SetNum;
  var CardID;
  var CardCol = 0;
  var CardQty;
  var MstrCol = 0;
  var CardListSet = CardList[0];
    
  Logger.log('Card Set: %s',CardListSet);

  if (CardListSet != 'No Pack Opened'){
    // Find Set Column according to Set in Cardlist (CardList[0]) and get all card quantities (first card starts at row 8, row 7 = card 0)
    for (var SetCol = 0; SetCol <= 23; SetCol++){   
      if (CardListSet == CardDBSet[0][SetCol]){
        CardCol = SetCol+1;
        Logger.log('Card Set Column: %s',CardCol);
        SetCol = 24;
      }
    }
    
    // Loop through each card in CardList and update quantity
    for (var CardListNb = 1; CardListNb <= 14; CardListNb++){
      // Get Card ID Number 
      CardID = CardList[CardListNb];
      
      // Regular cards and non Masterpiece card
      if (CardListNb < 14 || (CardListNb == 14 && CardList[15] == 'No')){
        // Update Quantity for that card
        CardQty = ShtCardDB.getRange(CardID+7, CardCol-2).getValue() + 1;
        ShtCardDB.getRange(CardID+7, CardCol-2).setValue(CardQty);
      }
      
      // If Last card is a Masterpiece
      if (CardListNb == 14 && CardList[15] == 'Yes'){
        // Get Set Number to find Masterpiece Column
        SetNum = ShtCardDB.getRange(4, CardCol).getValue();
        Logger.log('Masterpiece Set Number: %s',SetNum);
        // Set Masterpiece Column according to Set Number
        switch (SetNum){
          case 1 : MstrCol= 27; break;
          case 2 : MstrCol= 27; break;
          case 3 : MstrCol= 30; break;
          case 4 : MstrCol= 30; break;
          case 5 : MstrCol= 33; break;
          case 6 : MstrCol= 33; break;
          case 7 : MstrCol= 36; break;
          case 8 : MstrCol= 36; break;
        }
        CardQty = ShtCardDB.getRange(CardID+7, MstrCol-2).getValue() + 1;
        ShtCardDB.getRange(CardID+7, MstrCol-2).setValue(CardQty);
      }
      TestSht.getRange(CardListNb,1).setValue(CardID);
      TestSht.getRange(CardListNb,2).setValue(CardQty);
    }
    
    
    // Call function to generate clean card pool from Player Card DB
  }
  // Return Value
}