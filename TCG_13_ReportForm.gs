// **********************************************
// function fcnCreateReportForm()
//
// This function creates the Registration Form 
// based on the parameters in the Config File
//
// **********************************************

function fcnCreateReportForm() {
  
  var ss = SpreadsheetApp.getActive();
  var shtConfig = ss.getSheetByName('Config');
  var shtPlayers = ss.getSheetByName('Players');
  var ssID = shtConfig.getRange(30,2).getValue();
  var ssSheets;
  var NbSheets;
  var SheetName;
  var shtResp;
  
  var ssTexts = SpreadsheetApp.openById('1DkSr5HbGqZ_c38DlHKiBhgcBXw3fr3CK9zDE04187fE');
  var shtTxtReport = ssTexts.getSheetByName('Match Report');
  
  var formEN;
  var FormIdEN;
  var FormNameEN;
  var FormItemsEN;
  
  var formFR;
  var FormIdFR;
  var FormNameFR;
  var FormItemsFR;
  
  var OptLocation = shtConfig.getRange(14, 9).getValue();
  var WeekNum = shtConfig.getRange(5,7).getValue();
  var WeekArray = new Array(1); WeekArray[0] = WeekNum;
  
  var PlayerNum = shtPlayers.getRange(2,1).getValue();
  var Players;
  var PlayerList;
  
  var ExpansionNum = shtConfig.getRange(6,3).getValue();
  var ExpansionSet = shtConfig.getRange(7,6,ExpansionNum,1).getValues();
  var ExpansionList = new Array(ExpansionNum);
  
  var PlayerWinList;
  var PlayerLosList;
  var SctPackOpenEN;
  var SctPackOpenFR;
  
  var CardValidation;
  
  var ConfirmMsgEN;
  var ConfirmMsgFR;
  
  var RowFormIdEN = 36;
  var RowFormIdFR = 37;
  
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
    FormNameEN = shtConfig.getRange(3, 2).getValue() + " Match Reporter EN";
    formEN = FormApp.create(FormNameEN).setTitle(FormNameEN);
    // French    
    FormNameFR = shtConfig.getRange(3, 2).getValue() + " Match Reporter FR";
    formFR = FormApp.create(FormNameFR).setTitle(FormNameFR);
    
    // Set Match Report Form Description
    formEN.setDescription("Please enter the following information to submit your match result");
    formEN.setCollectEmail(true);
    
    formFR.setDescription("SVP, entrez les informations suivantes pour soumettre votre rapport de match");
    formFR.setCollectEmail(true);
    
    //---------------------------------------------
    // LOCATION SECTION
    // If Location Bonus is Enabled, add Location Section
    if (OptLocation == 'Enabled'){
      
      // English
      formEN.addPageBreakItem().setTitle("Location")
      formEN.addMultipleChoiceItem()
      .setTitle("Did you play at the store?")
      .setRequired(true)
      .setChoiceValues(["Yes","No"]);
      
      // French
      formFR.addPageBreakItem().setTitle("Localisation")
      formFR.addMultipleChoiceItem()
      .setTitle("Avez-vous joué au magasin?")
      .setRequired(true)
      .setChoiceValues(["Oui","Non"]);
    }
    
    //---------------------------------------------
    // WEEK NUMBER & PLAYERS SECTION
    
    // Transfers Players Double Array to Single Array
    if (PlayerNum > 0){
      Players = shtPlayers.getRange(3,2,PlayerNum,1).getValues();
      PlayerList = new Array(PlayerNum);
      for(var i = 0; i < PlayerNum; i++){
        PlayerList[i] = Players[i][0];
      }
    }
    
    // English
    formEN.addPageBreakItem().setTitle("Week Number & Players");
    // Week
    formEN.addListItem()
    .setTitle("Week")
    .setRequired(true)
    .setChoiceValues(WeekArray);
    
    // Winning Players
    PlayerWinList = formEN.addListItem()
    .setTitle("Winning Player")
    .setRequired(true);
    if (PlayerNum > 0) PlayerWinList.setChoiceValues(PlayerList);
    
    // Losing Players
    PlayerLosList = formEN.addListItem()
    .setTitle("Losing Player")
    .setRequired(true);
    if (PlayerNum > 0) PlayerLosList.setChoiceValues(PlayerList);
    
    // Score
    formEN.addMultipleChoiceItem()
    .setTitle("Score")
    .setRequired(true)
    .setChoiceValues(["2-0","2-1"]);
    
    // French
    formFR.addPageBreakItem().setTitle("Numéro de Semaine & Joueurs");
    // Semaine
    formFR.addListItem()
    .setTitle("Semaine Numéro")
    .setRequired(true)
    .setChoiceValues(WeekArray);
    
    // Joueurs
    PlayerWinList = formFR.addListItem()
    .setTitle("Joueur Gagnant")
    .setRequired(true);
    if (PlayerNum > 0) PlayerWinList.setChoiceValues(PlayerList);
    
    PlayerLosList = formFR.addListItem()
    .setTitle("Joueur Perdant")
    .setRequired(true);
    if (PlayerNum > 0) PlayerLosList.setChoiceValues(PlayerList);
    
    // Score
    formFR.addMultipleChoiceItem()
    .setTitle("Score")
    .setRequired(true)
    .setChoiceValues(["2-0","2-1"]);

    //---------------------------------------------
    // PUNITION PACK SECTION
    
    // English
    formEN.addPageBreakItem().setTitle("Punishment Pack");
    // Pack Opened?
    SctPackOpenEN = formEN.addMultipleChoiceItem().setTitle("Did you open a Punishment Pack?");
    SctPackOpenEN.setChoices([SctPackOpenEN.createChoice("Yes", FormApp.PageNavigationType.CONTINUE), 
                              SctPackOpenEN.createChoice("No", FormApp.PageNavigationType.SUBMIT)]);   
   
    // French
    formFR.addPageBreakItem().setTitle("Pack de Punition");
    // Pack Opened?
    SctPackOpenFR = formFR.addMultipleChoiceItem().setTitle("Avez-vous ouvert un Pack de Punition?");
    SctPackOpenFR.setChoices([SctPackOpenFR.createChoice("Oui", FormApp.PageNavigationType.CONTINUE), 
                              SctPackOpenFR.createChoice("Non", FormApp.PageNavigationType.SUBMIT)]);   
    
    //---------------------------------------------
    // EXPANSION SET
    // Transfers Expansion Set Double Array to Single Array
    for(var i = 0; i < ExpansionNum; i++){
      ExpansionList[i] = ExpansionSet[ExpansionNum-1 - i][0];
    }
    
    // English
    formEN.addPageBreakItem().setTitle("Expansion Set").setHelpText("Please select the expansion set of your punishment pack.");
    formEN.addListItem()
    .setTitle("Expansion Set")
    .setRequired(true)
    .setChoiceValues(ExpansionList);    
    
    // French
    formFR.addPageBreakItem().setTitle("Set d'Expansion").setHelpText("Please select the expansion set of your punishment pack.");
    formFR.addListItem()
    .setTitle("Expansion Set")
    .setRequired(true)
    .setChoiceValues(ExpansionList);

    //---------------------------------------------
    // CARD LIST
    
    // Card Number Validation
    CardValidation = FormApp.createTextValidation()
    .setHelpText("Enter a number between 1 and 100.")
    .requireNumberBetween(1, 100)
    .build();
    
    // English
    formEN.addPageBreakItem().setTitle("Card List").setHelpText("Enter the card numbers of each card of your pack. The Card Number is the first number in the lower left side corner.");
    
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
    formFR.addPageBreakItem().setTitle("Liste de Cartes").setHelpText("Entrez le numéro de chaque carte de votre pack. Le numéro de carte est le premier numéro dans le coin inférieur gauche de la carte.");

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
    // CONFIRMATION MESSAGE
    
    // English
    ConfirmMsgEN = shtTxtReport.getRange(4,2).getValue();
    formEN.setConfirmationMessage(ConfirmMsgEN);
    
    // French
    ConfirmMsgFR = shtTxtReport.getRange(4, 3).getValue();
    formFR.setConfirmationMessage(ConfirmMsgFR);
    
    // RESPONSE SHEETS
    // Create Response Sheet in Main File and Rename
    
    // English Form
    formEN.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    
    // Find and Rename Response Sheet
    ss = SpreadsheetApp.openById(ssID);
    ssSheets = ss.getSheets();
    ssSheets[0].setName('New Responses EN');
    
    // Move Response Sheet to appropriate spot in file
    shtResp = ss.getSheetByName('New Responses EN');
    ss.moveActiveSheet(15);
    
    // English Form
    formFR.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    
    // Find and Rename Response Sheet
    ss = SpreadsheetApp.openById(ssID);
    ssSheets = ss.getSheets();
    ssSheets[0].setName('New Responses FR');
    
    // Move Response Sheet to appropriate spot in file
    shtResp = ss.getSheetByName('New Responses FR');
    ss.moveActiveSheet(16);
    
    // Set Match Report IDs in Config File
    FormIdEN = formEN.getId();
    shtConfig.getRange(RowFormIdEN, 2).setValue(FormIdEN);
    FormIdFR = formFR.getId();
    shtConfig.getRange(RowFormIdFR, 2).setValue(FormIdFR);

  }

}