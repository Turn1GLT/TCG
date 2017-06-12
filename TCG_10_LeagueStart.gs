// **********************************************
// function fcnUpdateLinksIDs()
//
// This function updates all sheets Links and IDs  
// in the Config File
//
// **********************************************

function fcnUpdateLinksIDs(){
  
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Copy Log Spreadsheet
  var shtCopyLogID = shtConfig.getRange(27, 2).getValue();
  
  if (shtCopyLogID != '') {
    var shtCopyLog = SpreadsheetApp.openById(shtCopyLogID).getSheets()[0];
  
    var CopyLogNbFiles = shtCopyLog.getRange(2, 6).getValue();
    var StartRowCopyLog = 5;
    var StartRowConfigId = 30
    var StartRowConfigLink = 17;
    
    var CopyLogVal = shtCopyLog.getRange(StartRowCopyLog, 2, CopyLogNbFiles, 3).getValues();
    
    var FileName;
    var Link;
    var Formula;
    var ConfigRowID = 'Not Found';
    var ConfigRowLk = 'Not Found';
    
    // Loop through all Copied Sheets and get their Link and ID
    for (var row = 0; row < CopyLogNbFiles; row++){
      // Get File Name
      FileName = CopyLogVal[row][0];
      
      switch(FileName){
        case 'Master TCG Booster League' :
          ConfigRowID = StartRowConfigId + 0;
          ConfigRowLk = 'Not Found'; break;
        case 'Master TCG Booster League Card DB' :
          ConfigRowID = StartRowConfigId + 1; 
          ConfigRowLk = 'Not Found'; break;
        case 'Master TCG Booster League Card Pool EN' :
          ConfigRowID = StartRowConfigId + 2; 
          ConfigRowLk = StartRowConfigLink + 1; break;
        case 'Master TCG Booster League Card Pool FR' :
          ConfigRowID = StartRowConfigId + 3; 
          ConfigRowLk = StartRowConfigLink + 4; break;
        case 'Master TCG Booster League Standings EN' :
          ConfigRowID = StartRowConfigId + 4; 
          ConfigRowLk = StartRowConfigLink + 0; break;
        case 'Master TCG Booster League Standings FR' :
          ConfigRowID = StartRowConfigId + 5; 
          ConfigRowLk = StartRowConfigLink + 3; break;
        case 'Master TCG Booster League Match Reporter EN' :
          ConfigRowID = 'Not Found';
          ConfigRowLk = 'Not Found'; break;
        case 'Master TCG Booster League Match Reporter FR' :
          ConfigRowID = 'Not Found';
          ConfigRowLk = 'Not Found'; break;	
        default : 
          ConfigRowID = 'Not Found'; 
          ConfigRowLk = 'Not Found'; break;
      }
      
      // Set tthe Appropriate Sheet ID Value in the Config File
      if (ConfigRowID != 'Not Found') {
        shtConfig.getRange(ConfigRowID, 2).setValue(CopyLogVal[row][2]);
      }
      // Set tthe Appropriate Sheet ID Value in the Config File
      if (ConfigRowLk != 'Not Found') {
        // Opens Spreadsheet by ID
        Link = SpreadsheetApp.openById(CopyLogVal[row][2]).getUrl();
        Logger.log(Link); 
        
        shtConfig.getRange(ConfigRowLk, 2).setValue(Link);
      }
    }
  }
}


// **********************************************
// function fcnGenPlayerCardDB()
//
// This function generates all Card DB for all 
// players from the Config File
//
// **********************************************

function fcnGenPlayerCardDB(){
  
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Card DB Spreadsheet
  var CardDBShtID = shtConfig.getRange(31, 2).getValue();
  var ssCardDB = SpreadsheetApp.openById(CardDBShtID);
  var shtCardDB = ssCardDB.getSheetByName('Template');
  var shtCardDBNum;
  var CardDBHeader = shtCardDB.getRange(4,1,4,48).getValues();

  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var NbPlayers = shtPlayers.getRange(2,6).getValue();
    
  var NbSheets = ssCardDB.getNumSheets();
  
  var shtPlyrCardDB;
  var shtPlyrName;
  var SetNum;
  var PlyrRow;
  
  // Gets the Card Set Data from Config File to Populate the Header
  for (var col = 0; col < 48; col++){
    SetNum = CardDBHeader[0][col];
    switch (SetNum){
      case 1: 
        CardDBHeader[1][col] = shtConfig.getRange(7, 5).getValue();
        if (col < 32) CardDBHeader[2][col] = shtConfig.getRange(7, 6).getValue();
        if (col > 32) CardDBHeader[2][col] = shtConfig.getRange(7, 7).getValue();
        break;
      case 2: 
        CardDBHeader[1][col] = shtConfig.getRange(8, 5).getValue();
        if (col < 32) CardDBHeader[2][col] = shtConfig.getRange(8, 6).getValue();
        if (col > 32) CardDBHeader[2][col] = shtConfig.getRange(8, 7).getValue();
        break;
      case 3: 
        CardDBHeader[1][col] = shtConfig.getRange(9, 5).getValue();
        if (col < 32) CardDBHeader[2][col] = shtConfig.getRange(9, 6).getValue();
        if (col > 32) CardDBHeader[2][col] = shtConfig.getRange(9, 7).getValue();
        break;
      case 4: 
        CardDBHeader[1][col] = shtConfig.getRange(10, 5).getValue();
        if (col < 32) CardDBHeader[2][col] = shtConfig.getRange(10, 6).getValue();
        if (col > 32) CardDBHeader[2][col] = shtConfig.getRange(10, 7).getValue();
        break;
      case 5: 
        CardDBHeader[1][col] = shtConfig.getRange(11, 5).getValue();
        if (col < 32) CardDBHeader[2][col] = shtConfig.getRange(11, 6).getValue();
        if (col > 32) CardDBHeader[2][col] = shtConfig.getRange(11, 7).getValue();
        break;
      case 6: 
        CardDBHeader[1][col] = shtConfig.getRange(12, 5).getValue();
        if (col < 32) CardDBHeader[2][col] = shtConfig.getRange(12, 6).getValue();
        if (col > 32) CardDBHeader[2][col] = shtConfig.getRange(12, 7).getValue();
        break;
      case 7: 
        CardDBHeader[1][col] = shtConfig.getRange(13, 5).getValue();
        if (col < 32) CardDBHeader[2][col] = shtConfig.getRange(13, 6).getValue();
        if (col > 32) CardDBHeader[2][col] = shtConfig.getRange(13, 7).getValue();
        break;
      case 8: 
        CardDBHeader[1][col] = shtConfig.getRange(14, 5).getValue();
        if (col < 32) CardDBHeader[2][col] = shtConfig.getRange(14, 6).getValue();
        if (col > 32) CardDBHeader[2][col] = shtConfig.getRange(14, 7).getValue();
        break;
    }    
  }
  // Set Card Set Names and Codes
  shtCardDB.getRange(4,1,4,48).setValues(CardDBHeader);
  
  // Loops through each player starting from the first
  for (var plyr = NbPlayers; plyr > 0; plyr--){
    
    // Update the Player Row and Get Player Name from Config File
    PlyrRow = plyr + 2; // 2 is the row where the player list starts
    shtPlyrName = shtPlayers.getRange(PlyrRow, 2).getValue();
  
    // INSERTS TAB BEFORE "Card DB" TAB
    ssCardDB.insertSheet(shtPlyrName, 0, {template: shtCardDB});
    shtPlyrCardDB = ssCardDB.getSheets()[0];
        
    // Opens the new sheet and modify appropriate data (Player Name, Header)
    shtPlyrCardDB.getRange(3,3).setValue(shtPlyrName);
    shtPlyrCardDB.getRange(4,1,4,48).setValues(CardDBHeader);
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
    
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Card Pool Spreadsheet
  var CardPoolShtEnID = shtConfig.getRange(32, 2).getValue();
  var CardPoolShtFrID = shtConfig.getRange(33, 2).getValue();
  var ssCardPoolEn = SpreadsheetApp.openById(CardPoolShtEnID);
  var ssCardPoolFr = SpreadsheetApp.openById(CardPoolShtFrID);
  var shtCardPoolEn = ssCardPoolEn.getSheetByName('Template');
  var shtCardPoolFr = ssCardPoolFr.getSheetByName('Template');
  var shtCardPoolNum;
  
  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var NbPlayers = shtPlayers.getRange(2,6).getValue();
  
  var shtPlyrCardPoolEn;
  var shtPlyrCardPoolFr;
  var shtPlyrName;
  var PlyrRow
  
  // Loops through each player starting from the first
  for (var plyr = NbPlayers; plyr > 0; plyr--){
    
    // Update the Player Row and Get Player Name from Config File
    PlyrRow = plyr + 2; // 2 is the row where the player list starts
    shtPlyrName = shtPlayers.getRange(PlyrRow, 2).getValue();
  
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

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Card DB Spreadsheet
  var CardDBShtID = shtConfig.getRange(31, 2).getValue();
  var ssCardDB = SpreadsheetApp.openById(CardDBShtID);
  var shtTemplate = ssCardDB.getSheetByName('Template');
  var ssNbSheet = ssCardDB.getNumSheets();
  
  // Routine Variables
  var shtCurr;
  var shtCurrName;
  
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

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Card Pool Spreadsheet
  var CardPoolShtIDEn = shtConfig.getRange(32, 2).getValue();
  var CardPoolShtIDFr = shtConfig.getRange(33, 2).getValue();
  var ssCardPoolEn = SpreadsheetApp.openById(CardPoolShtIDEn);
  var ssCardPoolFr = SpreadsheetApp.openById(CardPoolShtIDFr);
  var shtTemplateEn = ssCardPoolEn.getSheetByName('Template');
  var shtTemplateFr = ssCardPoolFr.getSheetByName('Template');
  var ssNbSheet = ssCardPoolEn.getNumSheets();
  
  // Routine Variables
  var shtCurrEn;
  var shtCurrNameEn;
  var shtCurrFr;
  var shtCurrNameFr;
  
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