

    function onInstall(e) {
      onOpen(e);
    }

function onOpen(e) { 
  var ui = SpreadsheetApp.getUi();
  var ss= SpreadsheetApp.getActiveSpreadsheet();
  var menuName = ui.createMenu("WAREHOUSE MENU")
  
  .addItem("Non Cocoa", "nonCocoa").addSeparator()
  .addItem("Small Scale", "smallScale").addSeparator()
  .addItem("Pre-finance", "preFinance").addSeparator()
  .addItem("Evacuate", "evaCuate").addSeparator()
  .addItem("Add Customer", "addCustomer").addSeparator()
  .addItem("Financial History", "financeHistory").addSeparator()
  .addItem("Small Scale History", "SmallscaleHistory").addSeparator()
  .addItem("Dashboard", "dashBoard")
  .addToUi();
  
 //preFinanceGetFormula("PreFinance");
  preFinanceMS();
  
  
      
}

    function showSidebar() {
      var html = HtmlService.createHtmlOutputFromFile('index')
          .setTitle('G Suite Admin Console')
          .setWidth(400);
      SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
          .showSidebar(html);
    }

 
  


function addCustomer(){
  var ss = SpreadsheetApp.getUi();
  var input = ss.prompt("ADD CUSTOMER", "Please input the NEW customer's name", ss.ButtonSet.OK_CANCEL);
  var custSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("lists");
  var custLr = custSheet.getLastRow()+1;
  //Logger.log(custLr);
  
  if(input.getSelectedButton()==ss.Button.OK){
    var promptText = input.getResponseText().toUpperCase();
    
    custSheet.getRange(custLr, 3).setValue(promptText);
      
  
  
  }else if(input.getSelectedButton()==ss.Button.CANCEL){
  //do nothing
  
  }else if(input.getSelectedButton()==ss.Button.CLOSE){
  //do nothing
  
  }
}



function onEdit(e) {
  //PREFINANCE EDIT
  var row = e.range.getRow();
  var col = e.range.getColumn();

  
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetName();
  if(activeSheet=="Pre-finance"){
    preFinanceGetFormula();}
  if(activeSheet=="Small Scale Receipts"){
     SmallScaleGetFormula();
  }
  
  if(activeSheet=="Non Cocoa"){
     nonCocoaGetFormula();
  }
  
  
    if(activeSheet=="Dashboard"){
      if(col===2 && row===3){
        
        var dashboardSS = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        var startDateGetValue = dashboardSS.getRange("B3").getValue();
        dashboardSS.getRange("D3").setValue(startDateGetValue);
      
      
      }
 
  }
  
  
}




function hideSheets(activeSheet){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName(activeSheet).activate();
  var sheetsCount = ss.getNumSheets();
  var sheets = ss.getSheets();
  for(var i=0;i<sheetsCount;i++){
    var sheet = sheets[i];
    var fetchSheet =  sheet.getName()
    if (fetchSheet != activeSheet){
      sheet.hideSheet();
      //Logger.log(fetchSheet);
    }
   
  }
  
  
}

function nonCocoa(){ hideSheets("Non Cocoa");}
function smallScale(){ hideSheets("Small Scale Receipts");}
function preFinance(){ hideSheets("Pre-finance");}
function evaCuate(){ hideSheets("Evacuate");}
function financeHistory(){ hideSheets("Non Cocoa History");}
function SmallscaleHistory(){ hideSheets("Small-Scale History");}
function dashBoard(){ hideSheets("Dashboard");}



function nonCocoaGetFormula(){

  var ssfetch = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Non Cocoa");
  var lr = ssfetch.getLastRow()-4;
   
  ssfetch.getRange(6, 6, lr, 1).setFormula('=IF(B6="","",if(E6="Advance",D6,0))');
  ssfetch.getRange(6, 7, lr, 1).setFormula('=IF (B6="","",IFERROR(IFS(E6=lists!$A$6,$D6),0))');
  ssfetch.getRange(6, 8, lr, 1).setFormula('=IF(B6="","",if(E6="From Bank",D6,0))');
  ssfetch.getRange(6, 9, lr, 1).setFormula('=IF(B6="","",if(E6="Petty Cash",D6,0))');
  ssfetch.getRange(6, 10, lr, 1).setFormula('=IF(B6="","",if(E6="Cashier Refund",D6,0))');
  
  
  
}



function SmallScaleGetFormula(){

  var ssfetch = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Small Scale Receipts");
  var lr = ssfetch.getLastRow()-4;
   
  ssfetch.getRange(5, 6, lr, 1).setFormula('=iferror(round(I5 / ((G5*64)+H5),2),"")');
  ssfetch.getRange(5, 7, lr, 1).setFormula('=IF(C5>0,INT(K5/63),"")');  
  //ssfetch.getRange(5, 8, lr, 1).setFormula('=IF(B5>0,(IF(E5>0,(INT((((C5-E5)/64)-TRUNC((C5-E5)/64))*64)),((INT(((C5/64)-TRUNC(C5/64))*64))))),"")');
  ssfetch.getRange(5, 8, lr, 1).setFormula('=IF(C5="","",(INT(((K5/63)-INT(K5/63))*63)))');
  ssfetch.getRange(5, 10, lr, 1).setFormula('=IF(C5>0,C5-D5,"")');
  ssfetch.getRange(5, 11, lr, 1).setFormula('=IF(C5>0,J5-E5,"")');
  
}



//PRE-FINANCE FORMULA TO INCLUDE IN THE ONEDIT FUNCTION
function preFinanceGetFormula(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pre-finance");
  var lr = ss.getLastRow()-1;
 
  var getDate = ss.getRange(3, 2, lr, 1);

   ss.getRange(3, 1,lr,1).setFormula('=IF(C3="","",COUNTIF($C$3:C3 ,C3))');
   ss.getRange(3, 4,lr,1).setFormula('=IF(C3="","",C3&A3)');
   ss.getRange(3, 10,lr,1).setFormula('=iferror(round(R3 / (Q3),2),"")');
   ss.getRange(3, 12,lr,1).setFormula('=IF(K3 ="","",K3-R3)');
   ss.getRange(3, 13,lr,1).setFormula('=IF(F3 = "","",IF(C3>0,(ROUND((F3/H3),2)),""))');
   ss.getRange(3, 14,lr,1).setFormula('=IF(C3>0,(ROUND(((M3-8)/100)*(E3-(H3-I3)))),"")');
   ss.getRange(3, 15,lr,1).setFormula('=IF(C3>0,INT(Q3/63),"")');   
   ss.getRange(3, 16,lr,1).setFormula('=IF(C3="","",(INT(((Q3/63)-INT(Q3/63))*63)))');
   ss.getRange(3, 17,lr,1).setFormula('=IF(C3>0,(E3-(H3-I3))-N3,"")');
   
   ss.getRange(3, 19,lr,1).setFormula('=IF(K3="","",IF((C3&1)=D3,L3,(VLOOKUP(C3&(A3-1),$D$3:$S,16,0) + VLOOKUP(D3,$D$3:$L,9,0))))');
  
}


function preFinanceMS(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PreFinance");
  var lr = ss.getLastRow()-2;
  var lc = ss.getLastColumn();
  //var getAdjustedValuesForBalancing = ss.getRange(3, 12,87,8).getValues();
 
  var getDate = ss.getRange(3, 2, lr, 1);

   ss.getRange(3, 1,lr,1).setFormula('=IF(C3="","",COUNTIF($C$3:C3 ,C3))');
   ss.getRange(3, 4,lr,1).setFormula('=IF(C3="","",C3&A3)');
   ss.getRange(3, 10,lr,1).setFormula('=iferror(round(R3 / (Q3),2),"")');
   ss.getRange(3, 12,lr,1).setFormula('K3-R3');
   ss.getRange(3, 13,lr,1).setFormula('=IF(F3 = "","",IF(C3>0,(ROUND((F3/H3),2)),""))');
   ss.getRange(3, 14,lr,1).setFormula('=IF(C3>0,(ROUND(((M3-8)/100)*(E3-(H3-I3)))),"")');
   ss.getRange(3, 15,lr,1).setFormula('=IF(C3>0,INT(Q3/63),"")');   
   ss.getRange(3, 16,lr,1).setFormula('=IF(C3="","",(INT(((Q3/63)-INT(Q3/63))*63)))');
   ss.getRange(3, 17,lr,1).setFormula('=IF(C3>0,(E3-(H3-I3))-N3,"")');
   
   ss.getRange(3, 19,lr,1).setFormula('IF((C3&1)=D3,L3,(VLOOKUP(C3&(A3-1),$D$3:$S,16,0) + VLOOKUP(D3,$D$3:$L,9,0)))');
  //IF((C71&1)=D71,L71,(VLOOKUP(C71&(A71-1),$D$3:$S,16,0) + VLOOKUP(D71,$D$3:$L,9,0)))
  
  
}


function finalTransfer() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pre-finance");
  var ssFetch = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Non Cocoa");
  var lc = ss.getLastColumn();
  var lr = ss.getLastRow()-1;
  var labels = ss.getRange(1, 1, 1,lc).getValues();
  var preNames = ss.getRange(3, 3, lr,1);
  
 
  

  
  labels.forEach(function(label,j){
   
    
  var checkValues = combineValues("Date","Non Cocoa");
   
    if (lr == 1){ ss.getRange(3, 2, checkValues.length,1).setValues(checkValues);}else {
       ss.getRange(lr+1, 2, checkValues.length,1).setValues(checkValues);}
      
      //formulaCopy();
  
      
      
      
   
  
  })
  
  
   labels.forEach(function(label,i){
  var checkValues = combineValues("Name","Non Cocoa");
     if (lr == 1){ ss.getRange(3, 3, checkValues.length,1).setValues(checkValues);}else{
     ss.getRange(lr+1, 3, checkValues.length,1).setValues(checkValues);}
    
  
  
  })
    labels.forEach(function(label,i){
  var checkValues = combineValues("Amount","Non Cocoa");
     
    if (lr == 1){ ss.getRange(3, 11, checkValues.length,1).setValues(checkValues);}else{
      ss.getRange(lr+1, 11, checkValues.length,1).setValues(checkValues);}
  
  
  }) 
    
    //formulaCopy();
    //moveToMasterSheets();
    moveToSmallScale();
    moveToPrefinance(); 
    moveToFinanceSheet(); 
    preFinanceMS();
  
    
    }
 




function combineValues(label,sheetName) {
  
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var combValues = [];
  var tempValues = colValues(label,sheetName);
       
     combValues = combValues.concat(tempValues);
     //Logger.log(combValues);
          
      return combValues;
 
} 



function colValues(label,sheetName){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var colIndex = titleIndex(label,sheetName);
  var numRows = ss.getLastRow()-5;
  var sheetVal = ss.getRange(6, colIndex, numRows, 1).getValues();
  var resultArr = [];
  var filteredVal;

  var looper = ss.getRange(6, 5, numRows, 1).getValues();
  for(i=0; i<sheetVal.length; i++){
    
     
    if(looper[i] == 'Prefinance'){
       //Logger.log(sheetVal[i]);
      filteredVal = sheetVal[i];
      resultArr= resultArr.concat([sheetVal[i]]);
    }
    
      }
  
  return resultArr;


}



function titleIndex(label,sheetName) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var lc = ss.getLastColumn();
  var rangeValues = ss.getRange(5, 1,1,lc).getValues()[0];
  var index = rangeValues.indexOf(label)+1;
  return index;
}



function moveToPrefinance() {sheetTransfer("Pre-finance","PreFinance",2,3);}
function moveToFinanceSheet() {sheetTransfer("Non Cocoa","Finance",5,6);}
function moveToSmallScale() {sheetTransfer("Small Scale Receipts","SmallScale",3,5);}
//function moveToWareEvacuate() {sheetTransfer("Evacuate","wareEvacuate",3,4);}
//function moveToNew() {sheetTransfer("Sheet31","Pre-finance",1,3);}


function sheetTransfer(fetch,master,rowCount,startRow){
  
  var ssfetch = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(fetch);
  var ssMaster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(master);
  var fetchLr = ssfetch.getLastRow()-rowCount;
  var masterLr = ssMaster.getLastRow();
 // Logger.log(masterLr);
  var fetchLc = ssfetch.getLastColumn();
  var masterLc = ssMaster.getLastColumn();
  var fetchContents = ssfetch.getRange(startRow, 1, fetchLr, fetchLc);
  var fetchValues = fetchContents.getValues();
  var sheetFormulas = fetchContents.getFormulas();
  
  for(i=0;i<fetchValues.length;i++){
    ssMaster.appendRow(fetchValues[i]);}
        
  //Logger.log(masterLr);
  fetchContents.clearContent();
   // fetchContents.setFormulas(sheetFormulas);
    

}

