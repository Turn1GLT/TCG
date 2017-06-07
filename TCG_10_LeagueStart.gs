// **********************************************
// function fcnGenPlayerCardDB()
//
// This function generates all Card DB for all 
// players from the Config File
//
// **********************************************

function fcnGenPlayerCardDB(){
  
  // Config Spreadsheet
  var shtConfig = SpreadsheetApp.openById('1oXXEjOF9EoVxnR8pcmeNBSqJ1V-nPqPYNDwOnHWwznA').getSheetByName('Config');
  var CardDBShtID = shtConfig.getRange(61, 2).getValue();
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
  var shtConfig = SpreadsheetApp.openById('1oXXEjOF9EoVxnR8pcmeNBSqJ1V-nPqPYNDwOnHWwznA').getSheetByName('Config');
  var CardPoolShtEnID = shtConfig.getRange(62, 2).getValue();
  var CardPoolShtFrID = shtConfig.getRange(63, 2).getValue();
  var NbPlayers = shtConfig.getRange(16,7).getValue();
  
  // Card Pool Spreadsheet
  var ssCardPoolEn = SpreadsheetApp.openById(CardPoolShtEnID);
  var ssCardPoolFr = SpreadsheetApp.openById(CardPoolShtFrID);
  var shtCardPoolEn = ssCardPoolEn.getSheetByName('Template');
  var shtCardPoolFr = ssCardPoolFr.getSheetByName('Template');
  var shtCardPoolNum;
  
  var shtPlyrCardPoolEn;
  var shtPlyrCardPoolFr;
  var shtPlyrName;
  var PlyrRow
  
  // Loops through each player starting from the first
  for (var plyr = NbPlayers; plyr > 0; plyr--){
    
    // Update the Player Row and Get Player Name from Config File
    PlyrRow = plyr + 16; // 16 is the row where the player list starts
    shtPlyrName = shtConfig.getRange(PlyrRow, 2).getValue();
  
    // INSERTS TAB BEFORE "Card DB" TAB
    // English Version
    ssCardPoolEn.insertSheet(shtPlyrName, 0, {template: shtCardPoolEn});
    shtPlyrCardPoolEn = ssCardPoolEn.getSheets()[0];
       
    // Opens the new sheet and modify appropriate data (Player Name, Header)
    shtPlyrCardPoolEn.getRange(2,1).setValue(shtPlyrName);
    
    // French Version
    ssCardPoolFr.insertSheet(shtPlyrName, 0, {template: shtCardPoolFr});
    shtPlyrCardPoolFr = ssCardPoolFr.getSheets()[0];
    
    // Opens the new sheet and modify appropriate data (Player Name, Header)
    shtPlyrCardPoolFr.getRange(2,1).setValue(shtPlyrName);    
  }
  // English Version
  shtPlyrCardPoolEn = ssCardPoolEn.getSheets()[0];
  ssCardPoolEn.setActiveSheet(shtPlyrCardPoolEn);
  
  // French Version
  shtPlyrCardPoolFr = ssCardPoolFr.getSheets()[0];
  ssCardPoolFr.setActiveSheet(shtPlyrCardPoolFr);
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
  var shtConfig = SpreadsheetApp.openById('1oXXEjOF9EoVxnR8pcmeNBSqJ1V-nPqPYNDwOnHWwznA').getSheetByName('Config');
  var CardDBShtID = shtConfig.getRange(61, 2).getValue();
  
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
  var shtConfig = SpreadsheetApp.openById('1oXXEjOF9EoVxnR8pcmeNBSqJ1V-nPqPYNDwOnHWwznA').getSheetByName('Config');
  var CardPoolShtIDEn = shtConfig.getRange(62, 2).getValue();
  var CardPoolShtIDFr = shtConfig.getRange(63, 2).getValue();
  
  // Card Pool Spreadsheet
  var ssCardPoolEn = SpreadsheetApp.openById(CardPoolShtIDEn);
  var ssCardPoolFr = SpreadsheetApp.openById(CardPoolShtIDFr);
  var ssNbSheet = ssCardPoolEn.getNumSheets();
  var shtCurrEn;
  var shtCurrNameEn;
  var shtCurrFr;
  var shtCurrNameFr;
  var shtTemplateEn = ssCardPoolEn.getSheetByName('Template');
  var shtTemplateFr = ssCardPoolFr.getSheetByName('Template');
  
  // Activates Template Sheet
  ssCardPoolEn.setActiveSheet(shtTemplateEn);
  ssCardPoolFr.setActiveSheet(shtTemplateFr);
  
  for (var sht = 0; sht < ssNbSheet - 1; sht ++){
    
    // English Version
    shtCurrEn = ssCardPoolEn.getSheets()[0];
    shtCurrNameEn = shtCurrEn.getName();
    if( shtCurrNameEn != 'Template') ssCardPoolEn.deleteSheet(shtCurrEn);
    
    // French Version   
    shtCurrFr = ssCardPoolFr.getSheets()[0];
    shtCurrNameFr = shtCurrFr.getName();
    if( shtCurrNameFr != 'Template') ssCardPoolFr.deleteSheet(shtCurrFr);
  }
}