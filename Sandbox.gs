function subCreateArray() {
  
  var CardListData = new Array(16); // // 0 = Set Name, 1-14 = Card Numbers, 15 = Card 14 is Masterpiece (Y-N)
  
  for (var num = 0; num < 16; num++){
    CardListData[num] = new Array(4); // 0= Card in Pack, 1= Card Number, 2= Card Name, 3= Card Rarity
    for (var card = 0; card < 16; card++){
      switch (card){
        case 0: CardListData[num][card] = card; break; // Card in Pack
        case 1: CardListData[num][card] = card; break; // Card Number
        case 2: CardListData[num][card] = card; break; // Card Name
        case 3: CardListData[num][card] = card; break; // Card Rarity
      }
    }
  }
  ss.getSheetByName('Test').getRange(1, 1, 16, 4).setValues(CardListData);
}

// NOT USED
// Function to activate UI

function GameResultButton(){
  showAnchor('Send Match Result','https://goo.gl/forms/jcDtOML96WlNLzVL2');
}

function showAnchor(name,url) {
  var html = '<html><body><a href="'+url+'" target="blank" onclick="google.script.host.close()">'+name+'</a></body></html>';
  var ui = HtmlService.createHtmlOutput(html)
  SpreadsheetApp.getUi().showModelessDialog(ui,'Send Match Result');
}
