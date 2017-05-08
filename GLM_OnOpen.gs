// F2FLeagueOnOpen

// **********************************************
// function OnOpen()
//
// Function executed everytime the Spreadsheet is
// opened or refreshed
//
// **********************************************

function OnOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var FirstSht =  ss.getSheets()[0];
  
  ss.setActiveSheet(FirstSht);
  
  var FuncMenuButtons = [{name: "Analyze New Match Entry", functionName: "fcnGameResults"}];
  //var SortMenuButtons = [{name: "Sort Deck by Type/Color", functionName: "fcnSortDeckTypeColor"}, {name: "Sort Deck by Type/Card Name", functionName: "fcnSortDeckTypeName"}, {name: "Sort Deck by Color", functionName: "fcnSortDeckColor"}, {name: "Sort Deck by Card Name", functionName: "fcnSortDeckCardName"}, {name: "Sort Deck by Category", functionName: "fcnSortDeckCategory"}, {name: "Sort Deck by Staple", functionName: "fcnSortDeckStaple"}];
  
  ss.addMenu("General Fctn", FuncMenuButtons);
  //ss.addMenu("Sort Fctn", SortMenuButtons);
}