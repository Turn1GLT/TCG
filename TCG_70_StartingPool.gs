// **********************************************
// function fcnGenPlayerStartPool()
//
// This function generates all Starting Card Pool sheets
// for all players from the Config File
//
// **********************************************

function fcnGenPlayerStartPool(){
  
  // Open Start Pool Spreadsheet
  var ssStartPool = SpreadsheetApp.getActiveSpreadsheet();
  var ssStartSheets = ssStartPool.getSheets();
  var NumSheet = ssStartPool.getNumSheets();
  var shtTemplate = ssStartPool.getSheetByName('Template');
  
  // Main Config Spreadsheet
  var MainShtID = shtTemplate.getRange(3, 2).getValue();
  var ssMain = SpreadsheetApp.openById(MainShtID);
  var shtPlayers = ssMain.getSheetByName('Players'); 
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  
  // Routine Variables
  var PlayerFound = 0;
  var PlyrRow;
  var PlyrName;
  var SheetName;
  var shtPlyrStart;
  var StartPoolNumSht;
  
  // Loops through each player starting from the First
  for (var plyr = 1; plyr <= NbPlayers; plyr++){
    
    // Update the Player Row and Get Player Name from Player File
    PlyrRow = plyr + 2; // 2 is the row where the player list starts
    PlyrName = shtPlayers.getRange(PlyrRow, 2).getValue();
    
    // Resets the Player Found flag before searching
    PlayerFound = 0;
    
    // Look if player exists, if yes, skip, if not, create player
    for(var sheet = NumSheet - 2; sheet >= 0; sheet --){
      SheetName = ssStartSheets[sheet].getSheetName();
      Logger.log(SheetName);
      if (SheetName == PlyrName) PlayerFound = 1;
    }
    
    // If Player is not found, add a sheet
    if (PlayerFound == 0){
      // Get the Template sheet index
      StartPoolNumSht = ssStartPool.getNumSheets();
      
      // INSERTS TAB BEFORE "Card DB" TAB
      ssStartPool.insertSheet(PlyrName, StartPoolNumSht-2, {template: shtTemplate});
      shtPlyrStart = ssStartPool.getActiveSheet();
      shtPlyrStart.getRange(1, 2).setValue(PlyrName);
      shtPlyrStart.hideRow(3);
    }
  }
}

// **********************************************
// function fcnDelPlayerStartPool()
//
// This function deletes all Starting Card Pool sheets
//
// **********************************************

function fcnDelPlayerStartPool(){

  // Main Spreadsheet
  var ssStartPool = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtTemplate = ssStartPool.getSheetByName('Template');
  var ssNbSheet = ssStartPool.getNumSheets();
  
  // Routine Variables
  var shtCurr;
  var shtCurrName;
  
  // Activates Template Sheet
  ssStartPool.setActiveSheet(shtTemplate);
  
  for (var sht = 0; sht < ssNbSheet - 1; sht ++){
    shtCurr = ssStartPool.getSheets()[0];
    shtCurrName = shtCurr.getName();
    if( shtCurrName != 'Template') ssStartPool.deleteSheet(shtCurr);
    }
}


// **********************************************
// function fcnPopulatePlayerDB()
//
// This function generates all Starting Card Pool sheets
// for selected player
//
// **********************************************

function fcnPopulatePlayerDB(){
  
  // Open Start Pool Spreadsheet
  var ssStartPool = SpreadsheetApp.getActiveSpreadsheet();
  var shtTemplate = ssStartPool.getSheetByName('Template');
  
  // Player Sheets
  var shtPlayerStart = ssStartPool.getActiveSheet();
  var PlayerName = shtPlayerStart.getSheetName();
  var PlayerStatus = shtPlayerStart.getRange(2, 2).getValue();

  // Main Config Spreadsheet
  var MainShtID = shtPlayerStart.getRange(3, 2).getValue();
  var ssMain = SpreadsheetApp.openById(MainShtID);
  var shtConfig = ssMain.getSheetByName('Config');
  
  // Card DB Spreadsheet
  var CardDBShtID = shtConfig.getRange(31, 2).getValue();
  var ssCardDB = SpreadsheetApp.openById(CardDBShtID);
  var shtPlayerDB = ssCardDB.getSheetByName(PlayerName);
  
  // Card Pool Spreadsheet
  var ssCardPoolEnID = shtConfig.getRange(32,2).getValue();
  var ssCardPoolFrID = shtConfig.getRange(33,2).getValue();
  var shtCardPoolEn;
  var shtCardPoolFr;
  
  // Get Card Pool in array. Column [0] is header, columns [1-6] are boosters
  // Row [0] is Set Name, rows [1-14] are cards, row [15] is Masterpiece, row [16] is Status
  var rngStartCardPool = shtPlayerStart.getRange(5, 1, 17, 7);
  
  // Open Card Pool Sheets for that Player to send in parameter
  shtCardPoolEn = SpreadsheetApp.openById(ssCardPoolEnID).getSheetByName(PlayerName);
  shtCardPoolFr = SpreadsheetApp.openById(ssCardPoolFrID).getSheetByName(PlayerName);
  
  Logger.log(PlayerName);
  
  if (PlayerStatus != 'Starting Pool Complete') fcnPopulateCardDB(shtPlayerStart, shtPlayerDB, rngStartCardPool, PlayerName, shtCardPoolEn, shtCardPoolFr);
  
}

// **********************************************
// function fcnPopulateAllPlayerDB()
//
// This function generates all Starting Card Pool sheets
// for all players
//
// **********************************************

function fcnPopulateAllPlayerDB(){
  
  // Open Start Pool Spreadsheet
  var ssStartPool = SpreadsheetApp.getActiveSpreadsheet();
  var shtTemplate = ssStartPool.getSheetByName('Template');
  var ssStartSheets = ssStartPool.getSheets();
  var ssStartNumSheets = ssStartPool.getNumSheets();
  
  // Main Config Spreadsheet
  var MainShtID = shtTemplate.getRange(3, 2).getValue();
  var ssMain = SpreadsheetApp.openById(MainShtID);
  var shtConfig = ssMain.getSheetByName('Config');
  
  // Card DB Spreadsheet
  var CardDBShtID = shtConfig.getRange(31, 2).getValue();
  var ssCardDB = SpreadsheetApp.openById(CardDBShtID);
  
  // Card Pool Spreadsheet
  var ssCardPoolEnID = shtConfig.getRange(32,2).getValue();
  var ssCardPoolFrID = shtConfig.getRange(33,2).getValue();
  var shtCardPoolEn;
  var shtCardPoolFr;
  
  // Player Sheets
  var shtPlayerStart;
  var PlayerName;
  var PlayerStatus;
  var shtPlayerDB;
  var rngStartCardPool;
  var StartCardPool;
  
  for (var sheet = 0; sheet < ssStartNumSheets; sheet++){
  
    // Get Player sheets
    shtPlayerStart = ssStartPool.getSheets()[sheet];
    PlayerName = shtPlayerStart.getSheetName();
    PlayerStatus = shtPlayerStart.getRange(2, 2).getValue();
    Logger.log(PlayerName);
    
    if(PlayerName != 'Template' || PlayerName != 'Transfer Packs'){
      shtPlayerDB = ssCardDB.getSheetByName(PlayerName);
    
      // Get Card Pool in array. Column [0] is header, columns [1-6] are boosters
      // Row [0] is Set Name, rows [1-14] are cards, row [15] is Masterpiece, row [16] is Status
      rngStartCardPool = shtPlayerStart.getRange(5, 1, 17, 7);
    
      // Open Card Pool Sheets for that Player to send in parameter
      shtCardPoolEn = SpreadsheetApp.openById(ssCardPoolEnID).getSheetByName(PlayerName);
      shtCardPoolFr = SpreadsheetApp.openById(ssCardPoolFrID).getSheetByName(PlayerName);
      
      if (PlayerStatus != 'Starting Pool Complete') fcnPopulateCardDB(shtPlayerStart, shtPlayerDB, rngStartCardPool, PlayerName, shtCardPoolEn, shtCardPoolFr);
    }
  }
}

// **********************************************
// function fcnPopulateCardDB()
//
// This function generates all Starting Card Pool sheets
// for all players from the Config File
//
// **********************************************

function fcnPopulateCardDB(shtPlayerStart, shtPlayerDB, rngStartCardPool, PlayerName, shtCardPoolEn, shtCardPoolFr){
  
  var SetName;
  var CardNum;
  var CardQty;
  var SetNameDB;
  var SetNumDB;
  var SetNumMstr;
  var SetCardDB; // Array 300 rows x 2 columns. Rows = Cards 1-300, Columns 0 = Qty, 1 = Card Number
  
  var rngStatus;
  var StartCardPool = rngStartCardPool.getValues();
  
  var shtPlayerDBMaxCol = shtPlayerDB.getMaxColumns();
  
  // Loop through each Booster
  for (var booster = 1; booster <=6; booster++){
    // Get Set Name of selected booster
    SetName = StartCardPool[0][booster];
    Logger.log('Booster %s Set Name: %s',booster, SetName);
    if(SetName != ''){
      // Get Card DB Card List
      for(var setcol = 1; setcol <= shtPlayerDBMaxCol; setcol++){
        SetNameDB = shtPlayerDB.getRange(6,setcol).getValue();
        Logger.log('Set Name DB: %s',SetNameDB);
        
        // If Set Name is not null
        if (SetName == SetNameDB) {
          // Get Set Card List where:
          // Col[0] = Qty and Col[1] = Card Number
          // Row[0] = Header and Row[1-284] = Card Number
          SetCardDB = shtPlayerDB.getRange(7, setcol-2, 300, 2).getValues();
          
          // Loop through each card to update the quantity for regular cards
          for (var card = 1; card <=14; card++){
            CardNum = StartCardPool[card][booster];
            // First 13 cards
            if(CardNum != '' && card <  14 && SetCardDB[CardNum][1] == CardNum){
              CardQty = SetCardDB[CardNum][0];
              if(CardQty == '') SetCardDB[CardNum][0] = 0;
              SetCardDB[CardNum][0] += 1;
            }
            // Last card if not Masterpiece
            if(CardNum != '' && card == 14 && SetCardDB[CardNum][1] == CardNum && StartCardPool[15][booster] != 'yes'){
              CardQty = SetCardDB[CardNum][0];
              if(CardQty == '') SetCardDB[CardNum][0] = 0;
              SetCardDB[CardNum][0] += 1;
            }
          }
          
          // Update the Card DB for selected Set
          shtPlayerDB.getRange(7, setcol-2, 300, 2).setValues(SetCardDB);
                    
          // Update Masterpiece card
          if(StartCardPool[15][booster] == 'yes'){
            
            // Get Set Number from Card Database for Masterpiece
            SetNumDB =  shtPlayerDB.getRange(4, setcol).getValue();
            Logger.log('Set Num DB: %s', SetNumDB)
            // If Set Number = Even Number, assign the matching Set Number for Masterpieces
            switch(SetNumDB){
              case 2: SetNumDB = 1; break;
              case 4: SetNumDB = 3; break;
              case 6: SetNumDB = 5; break;
              case 8: SetNumDB = 7; break;
            }
            
            Logger.log('Set Num DB: %s', SetNumDB)
            
            CardNum = StartCardPool[14][booster];
            Logger.log('Card Num: %s',CardNum);
            
            if(CardNum != ''){
              // Find Masterpiece Card List with SetNumDB
              for(var setnum = 33; setnum <= shtPlayerDBMaxCol; setnum++){
                // Get Set Number from Masterpiece series
                SetNumMstr = shtPlayerDB.getRange(4,setnum).getValue();
                // When Set Number is found for Masterpiece
                if (SetNumDB == SetNumMstr){
                  Logger.log('Set Number Found: %s at column %s',SetNumMstr, setnum);
                  Logger.log('Cards Columns %s to %s',setnum-2, setnum-1);
                  Logger.log('-----------');
                  SetCardDB = shtPlayerDB.getRange(7, setnum-2, 60, 2).getValues();
                  Logger.log('Card Num DB: %s',SetCardDB[CardNum][1]);
                  CardQty = SetCardDB[CardNum][0];
                  if(CardQty == '') SetCardDB[CardNum][0] = 0;
                  SetCardDB[CardNum][0] += 1;
                  Logger.log('Card Qty: %s',SetCardDB[CardNum][0]);
                  // Update the Card DB for selected Set
                  shtPlayerDB.getRange(7, setnum-2, 60, 2).setValues(SetCardDB);
                  setnum = shtPlayerDBMaxCol + 1;
                }
              }
            }
          }
          // Update the Status
          StartCardPool[16][booster] = 'Posted';
          rngStartCardPool.setValues(StartCardPool);          
          // Exits the Loop
          setcol = shtPlayerDBMaxCol + 1;
        }
      }
    }
  }
  // Update Card Pool Lists
  fcnUpdateCardPool(shtPlayerStart, shtPlayerDB, PlayerName, shtCardPoolEn, shtCardPoolFr);
  
  // Set Player Status to: Starting Pool Complete
  rngStatus = shtPlayerStart.getRange(2, 2);
  rngStatus.setBackground('red');
  rngStatus.setValue('Starting Pool Complete');
}

// **********************************************
// function fcnUpdateCardPool()
//
// This function updates the Player card database  
// with the list of cards sent in arguments
//
// **********************************************

function fcnUpdateCardPool(shtPlayerStart, shtPlayerDB, PlayerName, shtCardPoolEn, shtCardPoolFr){
  
  // Variables
  var CardPoolEnMaxRows = shtCardPoolEn.getMaxRows();
  var CardPoolFrMaxRows = shtCardPoolFr.getMaxRows();
  var rngCardPoolEn = shtCardPoolEn.getRange(6, 1, CardPoolEnMaxRows-6, 5); // 0 = Card Qty, 1 = Card Number, 2 = Card Name, 3 = Rarity, 4 = Set Name 
  var rngCardPoolFr = shtCardPoolFr.getRange(6, 1, CardPoolFrMaxRows-6, 5); // 0 = Card Qty, 1 = Card Number, 2 = Card Name, 3 = Rarity, 4 = Set Name 
  var CardPool; // Where Card Data will be populated
  
  var CardDBSetTotal = shtPlayerDB.getRange(2,1,1,48).getValues(); // Gets Sum of all set Quantity, if > 0, set is present in card pool
  var CardTotal = shtPlayerDB.getRange(3,7).getValue();
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
      SetName = shtPlayerDB.getRange(6,colSet+2).getValue();

      // Get all Cards Data from set
      SetData = shtPlayerDB.getRange(7, colSet, 300, 4).getValues();

      // Loop through each card in Set and get Card Data
      for (var CardID = 1; CardID <= 299; CardID++){
        if (SetData[CardID][0] > 0) {
          CardPool[CardNb][0] = SetData[CardID][0]; // Quantity
          CardPool[CardNb][1] = SetData[CardID][1]; // Card Number (ID)
          CardPool[CardNb][2] = SetData[CardID][2]; // Card Name
          CardPool[CardNb][3] = SetData[CardID][3]; // Card Rarity
          CardPool[CardNb][4] = SetName;            // Set Name    
          CardNb++;
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