

  function onOpen(e) {
    SpreadsheetApp.getUi().createAddonMenu()
    .addItem('Add Entry', 'showSidebar')
    .addItem('Build CRM - Click Once!', 'buildCRM')
    .addToUi();
  }
  
  function onInstall(e) {
    onOpen(e);
  }
  
  function showSidebar() {
    var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Add New Customer');

    
    SpreadsheetApp.getUi().showSidebar(ui);
  }

  function buildCRM() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); 
    var crm3 = ss.insertSheet("CRM 2.0"); 
    crm3.getRange(1,1).setFormula("=Today()");
    crm3.getRange(2,1,1,11).setValues([["Name","Company","Address","Referral/Referred","Number","Email","Medium","Stage","Date - 1st Contact","Date - Last Contact","Coac"]]);
    
    crm3.getRange(3,1).setValue("Closed").setFontColor("#FFFFFF") && crm3.getRange(3,1,1,100).setBackground("#6fa8dc");
    crm3.getRange(4,1).setValue("Account").setFontColor("#FFFFFF") && crm3.getRange(4,1,1,100).setBackground("#93c47d");
    crm3.getRange(5,1).setValue("Qualified").setFontColor("#FFFFFF") && crm3.getRange(5,1,1,100).setBackground("#ffd966");
  }
  
function returnRow (searchTerm,searchColumn,sheet) {
  var i = 1;
  for (i;i<=500;i++) {
    var test = sheet.getRange(i,searchColumn,1,1).getValue();
    if(test==searchTerm) {
      Logger.log("Match is made at row " + i);
      return i;
    }
  }
}


function newEnter(test) {
  
     var ss = SpreadsheetApp.getActiveSpreadsheet();
     var crm2 = ss.getSheetByName("CRM 2.0");
     var ui = SpreadsheetApp.getUi();
     var run=true; 
  
      var closedRow = returnRow("Closed",1,crm2);
      var accountRow = returnRow("Account",1,crm2);
      var qualifiedRow = returnRow("Qualified",1,crm2);
      var stage = test[10]; 
      var address = test[5]+test[6]+test[7];
      var rr = test[8]+test[9];
  
    for(var i=1;i<crm2.getLastRow()+1;i++) {
       if(test[0]==crm2.getRange(i,1).getValue()) {
         run = false;
         Logger.log('we here');
         var response = ui.alert('Error. Client Already Exists!');
         break;
       }
     }
      
      //Gotta add COAC & Reasons for COAC 
      
      var testStageObject = [[test[0],test[2],address,rr,test[4],test[1],test[11],test[10],test[12],test[12],test[13]]];
      
      if(stage == "Closed" && run==true) {
        crm2.insertRowBefore(accountRow) && crm2.getRange(accountRow,1,1,25).clearFormat();
        crm2.getRange(accountRow,1,1,11).setValues(testStageObject);
        createDoc(test[0],ui,test[16],test[14],test[15]);
      } else if (stage == "Account" && run==true) {
        crm2.insertRowBefore(qualifiedRow) && crm2.getRange(qualifiedRow,1,1,25).clearFormat();
        crm2.getRange(qualifiedRow,1,1,11).setValues(testStageObject);
        createDoc(test[0],ui,test[16],test[14],test[15]);
      } else if (stage == "Qualified" && run==true) {
        Logger.log("Pass");
        crm2.getRange(crm2.getLastRow() + 1,1,1,11).setValues(testStageObject);
        createDoc(test[0],ui,test[16],test[14],test[15]);
      }
      
     
      
    
  }

function createDoc(x,ui,personalFacts,coacItems,objections) {
 //Strategy Report
      var doc = DocumentApp.create("Strategy Report: " +x);
      var response = ui.alert('New Strategy Report Made!');
      var body = doc.getBody();
      var title = body.insertParagraph(0, doc.getName()).setHeading(DocumentApp.ParagraphHeading.HEADING1);
      body.insertParagraph(1, "Personal Facts").setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.insertParagraph(2, personalFacts);
      body.insertParagraph(3, "COAC Items").setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.insertParagraph(5, coacItems);
      body.insertParagraph(6, "Objections").setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.insertParagraph(8, objections);
  
  //Add paragraph text ; new paragraph
}
  




