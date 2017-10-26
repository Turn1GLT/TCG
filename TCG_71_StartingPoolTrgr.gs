// **********************************************
// function onOpenTCGStartPool()
//
// Function executed everytime the Spreadsheet is
// opened or refreshed
//
// **********************************************

function onOpenTCGStartPool_Master() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var FileMenu = [];
  FileMenu.push({name:'Generate Players Starting Pools',functionName:'fcnGenPlayerStartPool'});
  FileMenu.push({name:'Delete Players Starting Pools',  functionName:'fcnDelPlayerStartPool'});
  FileMenu.push(null);
  FileMenu.push({name:'Populate Current Player Card DB', functionName:'fcnPopulatePlayerDB'});
  FileMenu.push({name:'Populate All Players Card DB',    functionName:'fcnPopulateAllPlayerDB'});

  
  ss.addMenu("Starting Pools", FileMenu);
}
