// **********************************************
// function fcnSetUpForm()
//
// This function creates the Registration Form 
// based on the parameters in the Config File
//
// **********************************************

function fcnSetUpForm() {
  
  var ss = SpreadsheetApp.getActive();
  var ssID;
  var shtConfig = ss.getSheetByName('Config');
  var ssSheets;
  var NbSheets;
  var SheetName;
  var shtRegResp;
  
  var form;
  var FormName;
  var FormID;
  var FormItems;
  var FormSubscrID;
  
}