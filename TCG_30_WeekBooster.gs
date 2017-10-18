// **********************************************
// function fcnWeekBoosterTCG()
//
// This function adds the Booster to the Player
// Card Pool and checks the Main Weekly Booster Table
//
// **********************************************

function fcnWeekBoosterTCG(shtResponse, RowResponse) {

  // Function Sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shtConfig = ss.getSheetByName('Config');
  var shtPlayers = ss.getSheetByName('Players');
  
  // Weekly Booster Sheets
  var ssIdWeekBstr = shtConfig.getRange(40, 2).getValue();
  var ssWeekBstr = SpreadsheetApp.openById(ssIdWeekBstr);
  var shtWeekBstrTable = ssWeekBstr.getSheetByName('Weekly Booster');
  var shtWeekBstrPlyr;
  
  // Function Values
  var NbPlayer = shtPlayers.getRange(2, 1).getValue();
  var PlayerTable = shtWeekBstrTable.getRange(4,2,NbPlayer,1).getValues();
  var rngResponseStatus = shtResponse.getRange(RowResponse,21);

  // Get Data From Response Sheet
  var MaxCol = shtResponse.getMaxColumns();
  var NewResponse = shtResponse.getRange(RowResponse,1,1,MaxCol).getValues(); // [0]= Timestamp [1]= Week [2]=Player [3]= Set [4-17]= Cards1-14 [18]= Masterpiece [19]= Status
  var WeekNum = NewResponse[0][1];
  var Player = NewResponse[0][2];
  var BoosterData = new Array(16);
  for(var i=0; i<16; i++){
    BoosterData[i]= new Array(2);
    BoosterData[i][0] = NewResponse[0][i+3];
  }
      
  // Open Card Pool DB and List Sheets for that Player to send in parameter
  var ssPlyrCardDBId = shtConfig.getRange(31, 2).getValue();
  var ssPlyrCardListEnId = shtConfig.getRange(32, 2).getValue();
  var ssPlyrCardListFrId = shtConfig.getRange(33, 2).getValue();
  var shtPlyrCardDB;
  var shtPlyrCardListEn;
  var shtPlyrCardListFr;
  
  // Function Variables
  var BoosterCheck;
  var RowPlayer;
  var ColWeekCardNum;
  var ValType;
  var ValInt;
  var rngBstrCheck;
  var rngBooster;
  var PopulateStatus = new Array(2);
  var UpdtListStatus;
  var ErrorMsg
  var ConfirmationMsg;
  var Error = new Array(2);
  
  // Error Initialization
  Error[0] = "No Error";
  Error[1] = "Pas d'Erreur";
    
  // Email Addresses Array
  var EmailAddresses = new Array(2); // 0= Language Preference, 1= Email Address
  
  // Create Array of 16x4 where each row is Card 1-14 and each column is Card Info
  var PackData = new Array(16); // 0 = Set Name, 1-14 = Card Numbers, 15 = Card 14 is Masterpiece (Y-N)
  for(var cardnum = 0; cardnum < 16; cardnum++){
    PackData[cardnum] = new Array(4); // 0= Card in Pack, 1= Card Number, 2= Card Name, 3= Card Rarity
    PackData[cardnum][0] = cardnum;
    PackData[cardnum][1] = BoosterData[cardnum][0];
    for (var val = 0; val < 4; val++) {
      if(val > 1) PackData[cardnum][val] = '';
    }
  }
  
  // Update the Response to confirm it was processed
  shtResponse.getRange(RowResponse,20).setValue('=IF(INDIRECT("R[0]C[-19]",FALSE)<>"",1,"")');

  // CODE START ------------------------------------------------------------------------------
  
  // Verify that Booster is allowed for selected week
  // Find Player Row in Player Table
  for(var i = 0; i < NbPlayer; i++){
    if(PlayerTable[i][0] == Player) {
      RowPlayer = i+4;
      i = NbPlayer;
    }
  }
  
  // Check if Booster can be added for select week
  rngBstrCheck = shtWeekBstrTable.getRange(RowPlayer, WeekNum+2);
  BoosterCheck = rngBstrCheck.getValue();
  
  // If there is a value in the Check Table, the Booster is not allowed for selected week
  if(BoosterCheck != '') {
    Error[0] = Player + " already added a booster for week " + WeekNum + ". Please select another week. ";
    Error[1] = Player + " a déjà ajouté un booster pour la semaine " + WeekNum + ". SVP, sélectionnez une autre semaine. ";
  }
  
  // Update the Response to confirm it was processed
  rngResponseStatus.setValue(Error[0]);
  
  // Check if Booster Information is Valid
  if(Error[0] == 'No Error'){
    // Verify all Booster information is present and is an integer
    for (i = 0; i < 16; i++){
      // Verify that Value is a Number
      if(i > 0 && i < 15 && BoosterData[i][0] != ''){
        ValType = typeof(BoosterData[i][0]);
        ValInt = BoosterData[i][0] % 1; 
        if(ValType != 'number' || ValInt != 0) {
          Error[0] = "Card " + i + " does not have a valid number. Value = " + BoosterData[i][0] +" Type: " + ValType;
          Error[1] = "Le numéro de la Carte " + i + " n'est pas valide. Valeur = " + BoosterData[i][0] +" Type: " + ValType;
          i = 16;
        }
      }
    }
  }

  // Update the Response to confirm it was processed
  rngResponseStatus.setValue(Error[0]);

  
  // If All information is present, execute
  if(Error[0] == 'No Error'){
    
    // Open Player Card Pool DB and Lists
    shtPlyrCardDB = SpreadsheetApp.openById(ssPlyrCardDBId).getSheetByName(Player);
    shtPlyrCardListEn = SpreadsheetApp.openById(ssPlyrCardListEnId).getSheetByName(Player);
    shtPlyrCardListFr = SpreadsheetApp.openById(ssPlyrCardListFrId).getSheetByName(Player);
    
    // Add Booster to Card DB and Regenerate Card Pool List
    PackData = fcnPopWeekBstrDB(ss, Player, BoosterData, PackData, shtPlyrCardDB, shtPlyrCardListEn, shtPlyrCardListFr);
    // Function Status is in PackData[0][2]
    PopulateStatus[0] = PackData[0][2];
    PopulateStatus[1] = PackData[0][3];
    Logger.log(PopulateStatus[0]);
    
    // Update the Response to confirm it was processed
    rngResponseStatus.setValue(PopulateStatus[0]);
    
    if(PopulateStatus[0] == 'Card DB Populate: Complete') {
      
      // Send Email to Confirm
      EmailAddresses = subGetEmailAddressSngl(Player, shtPlayers, EmailAddresses);
      
      if(EmailAddresses[0] == 'English')  fcnSendBstrCnfrmEmailEN(Player, WeekNum, EmailAddresses, PackData, shtConfig);
      if(EmailAddresses[0] == 'Français') fcnSendBstrCnfrmEmailFR(Player, WeekNum, EmailAddresses, PackData, shtConfig);
      
      // Update Card Lists
      UpdtListStatus = fcnUpdateBstrCardList(Player, shtPlyrCardDB, shtPlyrCardListEn, shtPlyrCardListFr);
      Logger.log(UpdtListStatus);
      
      // Update the Response to confirm it was processed
      rngResponseStatus.setValue(UpdtListStatus);
      
      // Add Booster Info to Player Weekly Booster Column
      // Open Player Sheet
      shtWeekBstrPlyr = ssWeekBstr.getSheetByName(Player);
      
      // Get Appropriate Week Column
      ColWeekCardNum = 2 + ((WeekNum - 2) * 2);
      // Populate Booster Data Card Names
      for(var card = 1; card < 15; card++){
        BoosterData[card][1] = PackData[card][2];
      }
      shtWeekBstrPlyr.getRange(3, ColWeekCardNum, 16, 2).setValues(BoosterData);
      
      // Add Check to Weekly Booster Table
      rngBstrCheck.setValue('1');
      
      // Update the Response to confirm it was processed
      rngResponseStatus.setValue("Weekly Booster Added");
    }
  }
  
  // If an Error has been detected, send Error Message
  if(Error[0] != 'No Error'){
    EmailAddresses[0] = 'Français';
    EmailAddresses[1] = shtConfig.getRange(12, 2).getValue();
    //EmailAddresses[1] = "turn1glt@gmail.com";
    fcnSendBstrErrorEmailFR(Player, WeekNum, EmailAddresses, PackData, Error, shtConfig);
  }
  
  // If Populate is not Complete, send Status
  if(Error[0] == 'No Error' && PopulateStatus[0] != 'Card DB Populate: Complete'){
    Error = PopulateStatus;
    EmailAddresses[0] = 'Français';
    EmailAddresses[1] = shtConfig.getRange(12, 2).getValue();
    //EmailAddresses[1] = "turn1glt@gmail.com";
    fcnSendBstrErrorEmailFR(Player, WeekNum, EmailAddresses, PackData, Error, shtConfig);
  }
}

// **********************************************
// function fcnPopWeekBstrDB()
//
// This function populates the Weekly Booster 
// to the selected player Card Database
//
// **********************************************

function fcnPopWeekBstrDB(ss, Player, BoosterData, PackData, shtPlyrCardDB, shtPlyrCardListEn, shtPlyrCardListFr){
  
  var SetName;
  var CardNum;
  var CardQty;
  var SetNameDB;
  var SetNumDB;
  var SetNumMstr;
  var SetNameMstr;
  var MasterpiecePresent = 0;
  var MasterpieceValid = 0;
  var SetCardDB; // Array 300 rows x 2 columns. Rows = Cards 1-300, Columns 0 = Qty, 1 = Card Number

  var Status = new Array(2);
  Status[0] = "Card DB Populate: Not Started";
  Status[1] = "Card DB Populate: Pas Démarré";
  
  var PlyrCardDBMaxCol = shtPlyrCardDB.getMaxColumns();
  
  // Get Set Name of selected booster
  SetName = BoosterData[0][0];
  
  // Set the Card Set Name in Pack Data
  PackData[0][0] = "Set Name";
  PackData[0][1] = SetName;
  
  if(BoosterData[15][0] == "Yes" || BoosterData[15][0] == "Oui")  MasterpiecePresent = 1;
  
  // Loop through Card DB to find the Appropriate Set
  for(var setcol = 1; setcol <= PlyrCardDBMaxCol; setcol++){
    SetNameDB = shtPlyrCardDB.getRange(6,setcol).getValue();
    
    // If Set Name is found
    if(SetName == SetNameDB) {
      
      Status[0] = "Card DB Populate: Card Set Found";Logger.log(Status[0]);
      Status[1] = "Card DB Populate: Set de Cartes trouvé";
      
      // Process Masterpiece card if present
      if(MasterpiecePresent == 1){
        Status[0] = "Card DB Populate: Analyzing Masterpiece Validity";Logger.log(Status[0]);
        Status[1] = "Card DB Populate: Analyse de validité de la Masterpiece";
        
        // Get Set Number from Card Database for Masterpiece
        SetNumDB =  shtPlyrCardDB.getRange(4, setcol).getValue();
        
        // If Set Number = Even Number, assign the matching Set Number for Masterpieces
        switch(SetNumDB){
          case 2: SetNumDB = 1; break;
          case 4: SetNumDB = 3; break;
          case 6: SetNumDB = 5; break;
          case 8: SetNumDB = 7; break;
        }
        // Masterpiece Card must be Card 14 in Pack
        CardNum = BoosterData[14][0];
        Logger.log(CardNum);
        if(CardNum != ""){
          // Find Masterpiece Card List with SetNumDB
          for(var Colsetnum = 33; Colsetnum <= PlyrCardDBMaxCol; Colsetnum++){
            // Get Set Number from Masterpiece series
            SetNumMstr = shtPlyrCardDB.getRange(4,Colsetnum).getValue();
            // When Set Number is found for Masterpiece
            Logger.log("%s - %s",SetNumDB,SetNumMstr);
            if(SetNumDB == SetNumMstr){
              Status[0] = "Card DB Populate: Masterpiece Set Found"; Logger.log(Status[0]);
              Status[1] = "Card DB Populate: Set Masterpiece Trouvé"; 
              // Get Masterpiece Set Name
              SetNameMstr = shtPlyrCardDB.getRange(6,Colsetnum).getValue();
              
              // If Masterpiece Set Name is null, set doesn"t have a Masterpiece Series, Reject Pack
              if(SetNameMstr != "" && CardNum <= 54) {
                MasterpieceValid = 1;
                Status[0] = "Card DB Populate: Masterpiece Card Found";
                Status[1] = "Card DB Populate: Carte Masterpiece Trouvée";
              }
              
              // If Masterpiece Set Name is null, set doesn"t have a Masterpiece Series, Reject Pack
              if(SetNameMstr == "") {
                MasterpieceValid = -1;
                Status[0] = "Expansion Set does not have Masterpiece Series";
                Status[1] = "Le Set d'Expansion ne possède pas de série Masterpiece";
                
                // Store Card Info to return to Main Function
                PackData[14][0] = 14;                    // Card in Pack
                PackData[14][1] = BoosterData[14][0]; // Card Number
                PackData[14][2] = "No Masterpiece Set"; // Card Name
                PackData[14][3] = "-"; // Card Rarity
              }
              
              // If Masterpiece Card Number is greater than 54, Card Number is Invalid, Reject Pack
              if(CardNum > 54) {
                MasterpieceValid = -2;
                Status[0] = "Masterpiece Card Number is not valid";
                Status[1] = "Le numéro de Carte Masterpiece n'est pas valide";
                
                // Store Card Info to return to Main Function
                PackData[14][0] = 14;                    // Card in Pack
                PackData[14][1] = BoosterData[14][0]; // Card Number
                PackData[14][2] = "Masterpiece Number Not Valid"; // Card Name
                PackData[14][3] = "-"; // Card Rarity
              }              
              
              // If Masterpiece Set is Valid, Process Masterpiece Card
              if(MasterpieceValid == 1){
                Status[0] = "Card DB Populate: Processing Masterpiece Card";
                Status[1] = "Card DB Populate: Carte Masterpiece en traitement";
                SetCardDB = shtPlyrCardDB.getRange(7, Colsetnum-2, 60, 4).getValues();
                Logger.log("Card Num DB: %s",SetCardDB[CardNum][1]);
                CardQty = SetCardDB[CardNum][0];
                if(CardQty == "") SetCardDB[CardNum][0] = 0;
                SetCardDB[CardNum][0] += 1; 
                
                // Store Card Info to return to Main Function
                PackData[14][0] = 14;                    // Card in Pack
                PackData[14][1] = SetCardDB[CardNum][1]; // Card Number
                PackData[14][2] = SetCardDB[CardNum][2]; // Card Name
                PackData[14][3] = SetCardDB[CardNum][3]; // Card Rarity
                
                // If Masterpiece is present, specify it in Pack Data[15]
                PackData[15][2] = "Masterpiece";

                Logger.log("Card Qty: %s",SetCardDB[CardNum][0]);
                // Update the Card DB for selected Set
                shtPlyrCardDB.getRange(7, Colsetnum-2, 60, 4).setValues(SetCardDB);
                Status[0] = "Card DB Populate: Masterpiece Card List Updated"; Logger.log(Status[0]);
                Status[1] = "Card DB Populate: Liste de Cartes Masterpiece à jour"; 
                Colsetnum = PlyrCardDBMaxCol + 1;
              }
            }
          }
        }
      }  
      
      //Process regular Cards
      // Get Set Card List where:
      // Col[0] = Qty and Col[1] = Card Number
      // Row[0] = Header and Row[1-284] = Card Number
      SetCardDB = shtPlyrCardDB.getRange(7, setcol-2, 300, 4).getValues();
      
      // Loop through each card to update the quantity for regular cards
      for (var card = 1; card <=14; card++){
        CardNum = BoosterData[card][0];
        
        if(MasterpieceValid >= 0){
          // Update Quantity for First 13 cards
          if(CardNum != "" && card < 14 && SetCardDB[CardNum][1] == CardNum){
            CardQty = SetCardDB[CardNum][0];
            if(CardQty == "") SetCardDB[CardNum][0] = 0;
            SetCardDB[CardNum][0] += 1;
          }
          // Last card if not Masterpiece
          if(CardNum != "" && card == 14 && SetCardDB[CardNum][1] == CardNum && MasterpiecePresent != 1){
            CardQty = SetCardDB[CardNum][0];
            if(CardQty == "") SetCardDB[CardNum][0] = 0;
            SetCardDB[CardNum][0] += 1;
            // If Masterpiece is not present, specify it in Pack Data[15]
            PackData[15][2] = "No Masterpiece";
          }
          
          Status[0] = "Card DB Populate: Card Quantities Updated";
          Status[1] = "Card DB Populate: Quantité de Cartes à Jour";
          
          // Update the Card DB for selected Set
          shtPlyrCardDB.getRange(7, setcol-2, 300, 4).setValues(SetCardDB);
          
          Status[0] = "Card DB Populate: Complete";  
          Status[1] = "Card DB Populate: Complété";
        }

        // If Pack Data is null, populate it
        if(PackData[card][2] == ''){
          // Store Card Info to return to Main Function
          PackData[card][0] = card;               // Card in Pack
          PackData[card][1] = SetCardDB[CardNum][1]; // Card Number
          PackData[card][2] = SetCardDB[CardNum][2]; // Card Name
          PackData[card][3] = SetCardDB[CardNum][3]; // Card Rarity
        }
      }
      // Exits the Loop
      setcol = PlyrCardDBMaxCol + 1;
      
    }
  }
  // Send Status through PackData
  PackData[0][2] = Status[0];
  PackData[0][3] = Status[1];
  //  var shtTest = ss.getSheetByName("Test");
  //  shtTest.getRange(1,1,16,4).setValues(PackData);
  return PackData; 
}

// **********************************************
// function fcnUpdateCardList()
//
// This function updates the Player card database  
// with the list of cards sent in arguments
//
// **********************************************

function fcnUpdateBstrCardList(Player, shtPlyrCardDB, shtPlyrCardListEn, shtPlyrCardListFr){
  
  // Variables
  var CardListEnMaxRows = shtPlyrCardListEn.getMaxRows();
  var CardListFrMaxRows = shtPlyrCardListFr.getMaxRows();
  var rngCardListEn = shtPlyrCardListEn.getRange(6, 1, CardListEnMaxRows-6, 5); // 0 = Card Qty, 1 = Card Number, 2 = Card Name, 3 = Rarity, 4 = Set Name 
  var rngCardListFr = shtPlyrCardListFr.getRange(6, 1, CardListFrMaxRows-6, 5); // 0 = Card Qty, 1 = Card Number, 2 = Card Name, 3 = Rarity, 4 = Set Name 
  var CardList; // Where Card Data will be populated
  
  var CardDBSetTotal = shtPlyrCardDB.getRange(2,1,1,48).getValues(); // Gets Sum of all set Quantity, if > 0, set is present in card pool
  var CardTotal = shtPlyrCardDB.getRange(3,7).getValue();
  var SetData;
  var SetName;
  var colSet;
  var CardNb = 0;
  
  var Status;
    
  // Clear Player Card Pool
  rngCardListEn.clearContent();
  CardList = rngCardListEn.getValues();
  rngCardListFr.clearContent();
  CardList = rngCardListFr.getValues();
  
  Status = 'Card List Update: Player Card Lists Cleared';
    
  // Look for Set with cards present in pool
  for (var col = 0; col <= 48; col++){   
    // if Set Card Quantity > 0, Set has cards in pool, Loop through all cards in Set
    if(CardDBSetTotal[0][col] > 0){
      colSet = col + 1;
      SetName = shtPlyrCardDB.getRange(6,colSet+2).getValue();

      // Get all Cards Data from set
      SetData = shtPlyrCardDB.getRange(7, colSet, 300, 4).getValues();

      // Loop through each card in Set and get Card Data
      for (var CardID = 1; CardID <= 299; CardID++){
        if(SetData[CardID][0] > 0) {
          CardList[CardNb][0] = SetData[CardID][0]; // Quantity
          CardList[CardNb][1] = SetData[CardID][1]; // Card Number (ID)
          CardList[CardNb][2] = SetData[CardID][2]; // Card Name
          CardList[CardNb][3] = SetData[CardID][3]; // Card Rarity
          CardList[CardNb][4] = SetName;            // Set Name    
          CardNb++;
        }
      }
    }
    Status = 'Card List Update: Player Cards Processed';
  }
  // Updates the Player Card Pool
  rngCardListEn.setValues(CardList);
  shtPlyrCardListEn.getRange(3,1).setValue(CardTotal);
  rngCardListFr.setValues(CardList);
  shtPlyrCardListFr.getRange(3,1).setValue(CardTotal);
  
  Status = 'Card List Update: Complete';
  
  return Status;
}