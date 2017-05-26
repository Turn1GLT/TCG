// **********************************************
// function fcnUpdateCardDB()
//
// This function updates the Player card database  
// with the list of cards sent in arguments
//
// **********************************************

function fcnUpdateCardDB(Player, CardList, shtTest){
  
  // Config Spreadsheet
  var ShtConfig = SpreadsheetApp.openById('14rR_7-SG9fTi-M7fpS7d6n4XrOlnbKxRW1Ni2ongUVU').getSheetByName('Config');
  var CardDBShtID = ShtConfig.getRange(3, 2).getValue();
  
  // Player Card DB Spreadsheet
  var shtCardDB = SpreadsheetApp.openById(CardDBShtID).getSheetByName(Player);
  var CardDBSet = shtCardDB.getRange(6,1,1,32).getValues();
  var MstrSet = shtCardDB.getRange(6,33,1,16).getValues();
  var ColCard = 0;
  var SetNum;
  var CardID;
  var CardQty;
  var CardNum;
  var CardName;
  var CardRarity;
  var CardListSet = CardList[0];
  var CardInfo; 
  var PackData = new Array(16); // 0 = Set Name, 1-14 = Card Numbers, 15 = Card 14 is Masterpiece (Y-N)
  
  // Create Array of 16x4 where each row is Card 1-14 and each column is Card Info
  for(var cardnum = 0; cardnum < 16; cardnum++){
    PackData[cardnum] = new Array(4); // 0= Card in Pack, 1= Card Number, 2= Card Name, 3= Card Rarity
    for (var val = 0; val < 4; val++) PackData[cardnum][val] = '';
  }
  
  Logger.log('Card Set: %s',CardListSet);
  
  // Updates the Set Name to return to Main Function
  PackData[0][0] = CardListSet;
  
  // Find Set Column according to Set in Cardlist (CardList[0]) and get all card quantities (first card starts at row 8, row 7 = card 0)
  for (var ColSet = 0; ColSet <= 31; ColSet++){   
    if (CardListSet == CardDBSet[0][ColSet]){
      ColCard = ColSet+1;
      Logger.log('Card Set Column: %s',ColCard);
      ColSet = 32;
    }
  }
  
  // Loop through each card in CardList to find the appropriate column to find card (Masterpiece or not)
  for (var CardListNb = 1; CardListNb <= 14; CardListNb++){
    // Get Card ID Number 
    CardID = CardList[CardListNb];
    
    // Regular cards and non Masterpiece card
    if (CardListNb < 14 || (CardListNb == 14 && CardList[15] == 'No')){
      ColCard = ColCard;
    }
    
    // If Last card is a Masterpiece
    if (CardListNb == 14 && CardList[15] == 'Yes'){
      // Get Set Number to find Masterpiece Column
      SetNum = shtCardDB.getRange(4, ColCard).getValue();
      Logger.log('Masterpiece Set Number: %s',SetNum);
      // Set Masterpiece Column according to Set Number
      switch (SetNum){
        case 1 : ColCard= 35; break;
        case 2 : ColCard= 35; break;
        case 3 : ColCard= 39; break;
        case 4 : ColCard= 39; break;
        case 5 : ColCard= 43; break;
        case 6 : ColCard= 43; break;
        case 7 : ColCard= 47; break;
        case 8 : ColCard= 47; break;
      }
    }
        
    // Get Card Info (Quantity, Card Number, Card Name, Rarity) // [0][0]= Card in Pack, [0][1]= Card Number, [0][2]= Card Name, [0][3]= Card Rarity
    CardInfo = shtCardDB.getRange(CardID+7, ColCard-2,1,4).getValues();
    CardQty  = CardInfo[0][0];
    CardNum  = CardInfo[0][1];
    CardName = CardInfo[0][2];
    CardRarity = CardInfo[0][3];
    
    // If Card Name exists, update card quantity and store Card Info to Pack Data
    if (CardName != ''){
      // Update Card Quantity in Card DB
      shtCardDB.getRange(CardID+7, ColCard-2).setValue(CardQty + 1);

      // Store Card Info to return to Main Function
      PackData[CardListNb][0] = CardListNb;  // Card in Pack
      PackData[CardListNb][1] = CardInfo[0][1]; // Card Number
      PackData[CardListNb][2] = CardInfo[0][2]; // Card Name
      PackData[CardListNb][3] = CardInfo[0][3]; // Card Rarity
      
      // If Last card is not a Masterpiece
      if (CardListNb == 14 && CardList[15] == 'No') PackData[15][2] = 'No Masterpiece';
      
      // If Last card is a Masterpiece
      if (CardListNb == 14 && CardList[15] == 'Yes') PackData[15][2] = 'Last Card is Masterpiece';
            
    }
    
    // If Card Name does not exist, set status to 0
    if (CardName == ''){
      PackData[CardListNb][2] = 'Card Name not Found for Card Number';
    }
    //shtTest.getRange(CardListNb,3).setValue(UpdateCardDBStatus[CardListNb]);
  }

  // Debug
  //shtTest.getRange(1,1,16,4).setValues(PackData);
  
  // Call function to generate clean card pool from Player Card DB
  fcnUpdateCardPool(shtCardDB, Player, shtTest);
  
  // Return Value
  return PackData;
}

