// **********************************************
// function fcnGenPlayerCardDB()
//
// This function generates all Card DB for all 
// players from the Config File
//
// **********************************************

function fcnGenPlayerCardDB(){
  
  // Config Spreadsheet
  var shtConfig = SpreadsheetApp.openById('14rR_7-SG9fTi-M7fpS7d6n4XrOlnbKxRW1Ni2ongUVU').getSheetByName('Config');
  var CardDBShtID = shtConfig.getRange(3, 2).getValue();
  var NbPlayers = shtConfig.getRange(16,7).getValue();
  
  // Card DB Spreadsheet
  var ssCardDB = SpreadsheetApp.openById(CardDBShtID);
  var shtCardDB = ssCardDB.getSheetByName('Template');
  var shtCardDBNum;
  var CardDBHeader = shtCardDB.getRange(4,1,4,36).getValues();
    
  var NbSheets = ssCardDB.getNumSheets();
  
  var shtPlyrCardDB;
  var shtPlyrName;
  var PlyrRow
  
  // Loops through each player starting from the first
  for (var plyr = NbPlayers; plyr > 0; plyr--){
    
    // Update the Player Row and Get Player Name from Config File
    PlyrRow = plyr + 16; // 16 is the row where the player list starts
    shtPlyrName = shtConfig.getRange(PlyrRow, 2).getValue();
  
    // INSERTS TAB BEFORE "Card DB" TAB
    ssCardDB.insertSheet(shtPlyrName, 0, {template: shtCardDB});
    shtPlyrCardDB = ssCardDB.getSheets()[0];
        
    // Opens the new sheet and modify appropriate data (Player Name, Header)
    shtPlyrCardDB.getRange(3,3).setValue(shtPlyrName);
    shtPlyrCardDB.getRange(4,1,4,36).setValues(CardDBHeader);
  }
  shtPlyrCardDB = ssCardDB.getSheets()[0];
  ssCardDB.setActiveSheet(shtPlyrCardDB);
}


// **********************************************
// function fcnGenPlayerCardPoolSht()
//
// This function generates all Card DB for all 
// players from the Config File
//
// **********************************************

function fcnGenPlayerCardPoolSht(){
  
  // Config Spreadsheet
  var shtConfig = SpreadsheetApp.openById('14rR_7-SG9fTi-M7fpS7d6n4XrOlnbKxRW1Ni2ongUVU').getSheetByName('Config');
  var CardPoolShtID = shtConfig.getRange(4, 2).getValue();
  var NbPlayers = shtConfig.getRange(16,7).getValue();
  
  // Card Pool Spreadsheet
  var ssCardPool = SpreadsheetApp.openById(CardPoolShtID);
  var shtCardPool = ssCardPool.getSheetByName('Template');
  var shtCardPoolNum;
    
  var NbSheets = ssCardPool.getNumSheets();
  
  var shtPlyrCardPool;
  var shtPlyrName;
  var PlyrRow
  
  // Loops through each player starting from the first
  for (var plyr = NbPlayers; plyr > 0; plyr--){
    
    // Update the Player Row and Get Player Name from Config File
    PlyrRow = plyr + 16; // 16 is the row where the player list starts
    shtPlyrName = shtConfig.getRange(PlyrRow, 2).getValue();
  
    // INSERTS TAB BEFORE "Card DB" TAB
    ssCardPool.insertSheet(shtPlyrName, 0, {template: shtCardPool});
    shtPlyrCardPool = ssCardPool.getSheets()[0];
        
    // Opens the new sheet and modify appropriate data (Player Name, Header)
    shtPlyrCardPool.getRange(2,1).setValue(shtPlyrName);
  }
  shtPlyrCardPool = ssCardPool.getSheets()[0];
  ssCardPool.setActiveSheet(shtPlyrCardPool);
}


// **********************************************
// function fcnDelPlayerCardDB()
//
// This function deletes all Card DB for all 
// players from the Config File
//
// **********************************************

function fcnDelPlayerCardDB(){

  // Config Spreadsheet
  var shtConfig = SpreadsheetApp.openById('14rR_7-SG9fTi-M7fpS7d6n4XrOlnbKxRW1Ni2ongUVU').getSheetByName('Config');
  var CardDBShtID = shtConfig.getRange(3, 2).getValue();
  
  // Card DB Spreadsheet
  var ssCardDB = SpreadsheetApp.openById(CardDBShtID);
  var ssNbSheet = ssCardDB.getNumSheets();
  var shtCurr;
  var shtCurrName;
  var shtTemplate = ssCardDB.getSheetByName('Template');
  
  // Activates Template Sheet
  ssCardDB.setActiveSheet(shtTemplate);
  
  for (var sht = 0; sht < ssNbSheet - 1; sht ++){
    shtCurr = ssCardDB.getSheets()[0];
    shtCurrName = shtCurr.getName();
    if( shtCurrName != 'Template') ssCardDB.deleteSheet(shtCurr);
    }
}

// **********************************************
// function fcnDelPlayerCardPoolSht()
//
// This function deletes all Card DB for all 
// players from the Config File
//
// **********************************************

function fcnDelPlayerCardPoolSht(){

  // Config Spreadsheet
  var shtConfig = SpreadsheetApp.openById('14rR_7-SG9fTi-M7fpS7d6n4XrOlnbKxRW1Ni2ongUVU').getSheetByName('Config');
  var CardPoolShtID = shtConfig.getRange(4, 2).getValue();
  
  // Card Pool Spreadsheet
  var ssCardPool = SpreadsheetApp.openById(CardPoolShtID);
  var ssNbSheet = ssCardPool.getNumSheets();
  var shtCurr;
  var shtCurrName;
  var shtTemplate = ssCardPool.getSheetByName('Template');
  
  // Activates Template Sheet
  ssCardPool.setActiveSheet(shtTemplate);
  
  for (var sht = 0; sht < ssNbSheet - 1; sht ++){
    shtCurr = ssCardPool.getSheets()[0];
    shtCurrName = shtCurr.getName();
    if( shtCurrName != 'Template') ssCardPool.deleteSheet(shtCurr);
  }
}