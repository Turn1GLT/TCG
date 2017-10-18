// **********************************************
// function fcnCreateWeekBstrForm()
//
// This function creates the Registration Form 
// based on the parameters in the Config File
//
// **********************************************

function fcnCreateWeekBstrForm() {
  
  var ss = SpreadsheetApp.getActive();
  var shtConfig = ss.getSheetByName('Config');
  var shtPlayers = ss.getSheetByName('Players');
  var ssID = shtConfig.getRange(30,2).getValue();
  var ssSheets;
  var NbSheets;
  var SheetName;
  var shtResp;
  
  var ssWeekBstrID = shtConfig.getRange(40, 2).getValue();
  var shtWeekBstr = SpreadsheetApp.openById(ssWeekBstrID).getSheetByName('Weekly Booster');
  
  var formEN;
  var FormIdEN;
  var FormNameEN;
  var FormItemsEN;
  
  var formFR;
  var FormIdFR;
  var FormNameFR;
  var FormItemsFR;
  
  var WeekNum;
  var WeekArray = new Array(1);
  
  var PlayerNum = shtPlayers.getRange(2,1).getValue();
  var Players;
  var PlayerList;
  var PlayerBstrList;
  
  var ExpansionNum = shtConfig.getRange(6,3).getValue();
  var ExpansionSet = shtConfig.getRange(7,6,ExpansionNum,1).getValues();
  var ExpansionList = new Array(ExpansionNum);

  var SctPackOpenEN;
  var SctPackOpenFR;
  
  var CardValidation;
  
  var ConfirmMsgEN;
  var ConfirmMsgFR;
  var FormUrlEN;
  var FormUrlFR;
  
  var RowFormUrlEN = 27;
  var RowFormUrlFR = 28;
  var RowFormIdEN = 42;
  var RowFormIdFR = 43;
  
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
    
    //---------------------------------------------
    // TITLE SECTION
    // English
    FormNameEN = shtConfig.getRange(3, 2).getValue() + " Weekly Booster EN";
    formEN = FormApp.create(FormNameEN).setTitle(FormNameEN);
    // French    
    FormNameFR = shtConfig.getRange(3, 2).getValue() + " Weekly Booster FR";
    formFR = FormApp.create(FormNameFR).setTitle(FormNameFR);
    
    // Set Match Report Form Description
    //formEN.setDescription("Please enter the Booster information");
    //formEN.setCollectEmail(true);
    
    //formFR.setDescription("SVP, entrez les informations de votre Booster");
    //formFR.setCollectEmail(true);
    
    //---------------------------------------------
    // WEEK NUMBER & PLAYERS SECTION

    // Creates the Week Numbers Array
    for(var i = 0; i < 7; i++){
      WeekArray[i] = i+2;
    }
    
    // Transfers Players Double Array to Single Array
    if (PlayerNum > 0){
      Players = shtPlayers.getRange(3,2,PlayerNum,1).getValues();
      PlayerList = new Array(PlayerNum);
      for(var i = 0; i < PlayerNum; i++){
        PlayerList[i] = Players[i][0];
      }
    }
    
    // English
    formEN.addSectionHeaderItem().setTitle("Week Number & Players");
    // Week
    formEN.addListItem()
    .setTitle("Week")
    .setRequired(true)
    .setChoiceValues(WeekArray);
    
    // Players
    PlayerBstrList = formEN.addListItem()
    .setTitle("Player")
    .setRequired(true);
    if (PlayerNum > 0) PlayerBstrList.setChoiceValues(PlayerList);
     
    // French
    formFR.addSectionHeaderItem().setTitle("Numéro de Semaine & Joueurs");
    // Semaine
    formFR.addListItem()
    .setTitle("Semaine Numéro")
    .setRequired(true)
    .setChoiceValues(WeekArray);
    
    // Joueurs
    PlayerBstrList = formFR.addListItem()
    .setTitle("Joueur")
    .setRequired(true);
    if (PlayerNum > 0) PlayerBstrList.setChoiceValues(PlayerList);

    
    //---------------------------------------------
    // BOOSTER INFORMATION 
    
    // Card Number Validation
    CardValidation = FormApp.createTextValidation()
    .setHelpText("Enter a number between 1 and 300.")
    .requireNumberBetween(1, 300)
    .build();
    
    // EXPANSION SET
    // Transfers Expansion Set Double Array to Single Array
    for(var i = 0; i < ExpansionNum; i++){
      ExpansionList[i] = ExpansionSet[ExpansionNum-1 - i][0];
    }
 
    // English
    formEN.addSectionHeaderItem().setTitle("Booster Info").setHelpText("Enter the card numbers of each card of your pack. The Card Number is the first number in the lower left side corner.");
    
    // English
    formEN.addListItem()
    .setTitle("Expansion Set")
    .setRequired(true)
    .setChoiceValues(ExpansionList);    
   
    // Loop to create first 13 cards of the pack
    for(var card = 1; card<=13; card++){
      formEN.addTextItem()
      .setTitle("Card " + card)
      .setRequired(true)
      .setValidation(CardValidation);
    }
    
    // Create last card to specify Masterpiece Number if applicable 
    formEN.addTextItem()
    .setTitle("Card 14 / Masterpiece")
    .setHelpText("If you opened a Masterpiece, Please enter the card number here (Kaladesh Invention, Amonkhet Invocation)")
    .setRequired(true)
    .setValidation(CardValidation);
    
    // Create Masterpiece Selection
    formEN.addMultipleChoiceItem()
    .setTitle("Masterpiece")
    .setHelpText("Did you open a Masterpiece Foil (Kaladesh Invention, Amonkhet Invocation)")
    .setRequired(true)
    .setChoiceValues(["Yes","No"]);    
    
    
    // French
    formFR.addSectionHeaderItem().setTitle("Booster Info").setHelpText("Entrez le numéro de chaque carte de votre pack. Le numéro de carte est le premier numéro dans le coin inférieur gauche de la carte.");

    // French
    formFR.addListItem()
    .setTitle("Expansion Set")
    .setRequired(true)
    .setChoiceValues(ExpansionList);
    
    // Loop to create first 13 cards of the pack
    for(var card = 1; card<=13; card++){
      formFR.addTextItem()
      .setTitle("Carte " + card)
      .setRequired(true)
      .setValidation(CardValidation);
    }
    
    // Create last card to specify Masterpiece Number if applicable 
    formFR.addTextItem()
    .setTitle("Carte 14 / Masterpiece")
    .setHelpText("Si vous avez ouvert une Masterpiece, SVP, entrez son numéro ici (Kaladesh Invention, Amonkhet Invocation)")
    .setRequired(true)
    .setValidation(CardValidation);
    
    // Create Masterpiece Selection
    formFR.addMultipleChoiceItem()
    .setTitle("Masterpiece")
    .setHelpText("Avez-vous ouvert une Masterpiece (Kaladesh Invention, Amonkhet Invocation)")
    .setRequired(true)
    .setChoiceValues(["Oui","Non"]);  
    
    //---------------------------------------------
    // RESPONSE SHEETS
    // Create Response Sheet in Main File and Rename
    
    // English Form
    formEN.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    
    // Find and Rename Response Sheet
    ss = SpreadsheetApp.openById(ssID);
    ssSheets = ss.getSheets();
    ssSheets[0].setName('WeekBstr EN');
    
    // Move Response Sheet to appropriate spot in file
    shtResp = ss.getSheetByName('WeekBstr EN');
    ss.moveActiveSheet(19);
    
    // English Form
    formFR.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    
    // Find and Rename Response Sheet
    ss = SpreadsheetApp.openById(ssID);
    ssSheets = ss.getSheets();
    ssSheets[0].setName('WeekBstr FR');
    
    // Move Response Sheet to appropriate spot in file
    shtResp = ss.getSheetByName('WeekBstr FR');
    ss.moveActiveSheet(20);
    
    // Set Match Report IDs in Config File
    FormIdEN = formEN.getId();
    shtConfig.getRange(RowFormIdEN, 2).setValue(FormIdEN);
    FormIdFR = formFR.getId();
    shtConfig.getRange(RowFormIdFR, 2).setValue(FormIdFR);
    
    // Create Links to add to Config File
    
    FormUrlEN = formEN.getPublishedUrl();
    FormUrlFR = formFR.getPublishedUrl();
    shtConfig.getRange(RowFormUrlEN, 2).setValue(FormUrlEN); 
    shtConfig.getRange(RowFormUrlFR, 2).setValue(FormUrlFR);
    
    // Copy Link to Weekly Booster Spreadsheet
    // shtWeekBstr.getRange(2, 4).setValue('=HYPERLINK("' + FormUrlEN + '","Add Weekly Booster")');
    shtWeekBstr.getRange(2, 4).setValue('=HYPERLINK("' + FormUrlFR + '","Ajouter Booster de Semaine")');

  }
}