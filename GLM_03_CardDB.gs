// **********************************************
// function fcnGenPlayerCardDB()
//
// This function generates all Card DB for all 
// players from the Config File
//
// **********************************************

function fcnGenPlayerCardDB(){
  
  // Card DB Spreadsheet
  var ssCardDB = SpreadsheetApp.openById('1fzCR0bmX2mGvaRE5empwu1rVMkGTPd5dO6FyQg17I68');
  var ShtCardDB = ssCardDB.getSheetByName('Template');
  var ShtCardDBNum;
  var CardDBHeader = ShtCardDB.getRange(4,1,4,36).getValues();
    
  var NbSheets = ssCardDB.getNumSheets();
  
  // Config Spreadsheet
  var ShtConfig = SpreadsheetApp.openById('14rR_7-SG9fTi-M7fpS7d6n4XrOlnbKxRW1Ni2ongUVU').getSheetByName('Config');
  var NbPlayers = ShtConfig.getRange(16,7).getValue();
  
  var ShtPlyrCardDB;
  var ShtPlyrName;
  var PlyrRow
  
  // Loops through each player starting from the first
  for (var plyr = NbPlayers; plyr > 0; plyr--){
    
    // Update the Player Row and Get Player Name from Config File
    PlyrRow = plyr + 16; // 16 is the row where the player list starts
    ShtPlyrName = ShtConfig.getRange(PlyrRow, 2).getValue();
  
    // INSERTS TAB BEFORE "Card DB" TAB
    ssCardDB.insertSheet(ShtPlyrName, 0, {template: ShtCardDB});
    ShtPlyrCardDB = ssCardDB.getSheets()[0];
        
    // Opens the new sheet and modify appropriate data (Player Name, Header)
    ShtPlyrCardDB.getRange(3,3).setValue(ShtPlyrName);
    ShtPlyrCardDB.getRange(4,1,4,36).setValues(CardDBHeader);
  }
  ShtPlyrCardDB = ssCardDB.getSheets()[0];
  ssCardDB.setActiveSheet(ShtPlyrCardDB);
}
