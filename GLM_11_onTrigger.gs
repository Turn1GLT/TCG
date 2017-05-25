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