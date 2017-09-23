// **********************************************
// function fcnCreateRegForm()
//
// This function creates the Registration Form 
// based on the parameters in the Config File
//
// **********************************************

function fcnCreateRegForm() {
  
  var ss = SpreadsheetApp.getActive();
  var shtConfig = ss.getSheetByName('Config');
  var shtPlayers = ss.getSheetByName('Players');
  var ssID = shtConfig.getRange(30,2).getValue();
  var ssSheets;
  var NbSheets;
  var SheetName;
  var shtResp;
  
  var formEN;
  var FormIdEN;
  var FormNameEN;
  var FormItemsEN;
  
  var formFR;
  var FormIdFR;
  var FormNameFR;
  var FormItemsFR;

  var RowFormUrlEN = 23;
  var RowFormUrlFR = 24;
  var RowFormIdEN = 38;
  var RowFormIdFR = 39;
  
  var ErrorVal = '';
  
  // Gets the Subscription ID from the Config File
  FormIdEN = shtConfig.getRange(RowFormIdEN, 2).getValue();
  FormIdFR = shtConfig.getRange(RowFormIdFR, 2).getValue();

  // If Form Exists, Log Error Message
  if(FormIdEN != ''){
    ErrorVal = 1;
    Logger.log('Error! FormIdEN already exists. Unlink Response and Delete Form');
  }
  // If Form Exists, Log Error Message
  if(FormIdEN != ''){
    ErrorVal = 1;
    Logger.log('Error! FormIdFR already exists. Unlink Response and Delete Form');
  }

  if (FormIdEN == '' && FormIdFR == ''){
    // Create Forms
    FormNameEN = shtConfig.getRange(3, 2).getValue() + " Registration EN";
    formEN = FormApp.create(FormNameEN).setTitle(FormNameEN);
    
    FormNameFR = shtConfig.getRange(3, 2).getValue() + " Registration FR";
    formFR = FormApp.create(FormNameFR).setTitle(FormNameFR);
    
    
    // Set Registration Email collection
    formEN.setCollectEmail(true);
    formFR.setCollectEmail(true);

    // FIRST NAME    
    formEN.addTextItem()
    .setTitle("First Name")
    .setRequired(true);
        
    formFR.addTextItem()
    .setTitle("Prénom")
    .setRequired(true);
    
    // LAST NAME 
    formEN.addTextItem()
    .setTitle("Last Name")
    .setRequired(true);
    
    formFR.addTextItem()
    .setTitle("Nom de Famille")
    .setRequired(true);
    
    // PHONE NUMBER    
    formEN.addTextItem()
    .setTitle("Phone Number")
    .setRequired(true);
    
    formFR.addTextItem()
    .setTitle("Numéro de téléphone")
    .setRequired(true);
    
    // LANGUAGE
    formEN.addMultipleChoiceItem()
    .setTitle("Communication Language")
    .setHelpText("Which Language do you prefer to use? The application is available in English and French")
    .setRequired(true)
    .setChoiceValues(["English","Français"]);

    formFR.addMultipleChoiceItem()
    .setTitle("Communication Language")
    .setHelpText("Quelle langue préférez-vous utiliser? L'application est disponible en anglais et en français.")
    .setRequired(true)
    .setChoiceValues(["English","Français"]);

    // DCI NUMBER
    formEN.addTextItem()
    .setTitle("DCI Number")
    .setRequired(true);
    
    formFR.addTextItem()
    .setTitle("Numéro DCI")
    .setRequired(true);
  
    // RESPONSE SHEETS
    // Create Response Sheet in Main File and Rename
    
    // English Form
    formEN.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    
    // Find and Rename Response Sheet
    ss = SpreadsheetApp.openById(ssID);
    ssSheets = ss.getSheets();
    ssSheets[0].setName('Registration EN');
    
    // Move Response Sheet to appropriate spot in file
    shtResp = ss.getSheetByName('Registration EN');
    ss.moveActiveSheet(17);
    
    // English Form
    formFR.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    
    // Find and Rename Response Sheet
    ss = SpreadsheetApp.openById(ssID);
    ssSheets = ss.getSheets();
    ssSheets[0].setName('Registration FR');
    
    // Move Response Sheet to appropriate spot in file
    shtResp = ss.getSheetByName('Registration FR');
    ss.moveActiveSheet(18);

    // Set Match Report IDs in Config File
    FormIdEN = formEN.getId();
    shtConfig.getRange(RowFormIdEN, 2).setValue(FormIdEN);
    FormIdFR = formFR.getId();
    shtConfig.getRange(RowFormIdFR, 2).setValue(FormIdFR);
    
    // Create Links to add to Config File  
    urlFormEN = formEN.getPublishedUrl();
    shtConfig.getRange(RowFormUrlEN, 2).setValue(urlFormEN); 
    
    urlFormFR = formFR.getPublishedUrl();
    shtConfig.getRange(RowFormUrlFR, 2).setValue(urlFormFR);
  }
}  