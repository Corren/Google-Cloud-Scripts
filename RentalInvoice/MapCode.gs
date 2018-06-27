function createAllInvoices(month,printPDF) {
  var errorLog = "";
  var errorFlag = false;
  var ss = SpreadsheetApp.getActive();
  var allNames = "";
  var counter =0;
 // var n=0;
  for(var n in ss.getSheets()){
    var sheet = ss.getSheets()[n];// look at every sheet in spreadsheet
    var name = sheet.getName();//get name
    if(name !== "Index" && name !== "Template" && name !== "Master Template"){ // exclude some names
      allNames += name+",";
      
      var error = createInvoice(month, name, printPDF,false);
      if (error !== "No errors detected")
      {
        errorFlag = true;
        errorLog += counter + " " +name +error + "\n";
        if (counter % 5 == 0)
        {
          emailMe(errorLog,"Every 5 rentals - update");
        }
      }
      counter++;
    //  testWait();
    }
  }
  var now = new Date();
  var message = "Done creating invoices for "+month+" month generated on "+now+"\n\n";
  var subject = "";
  if (errorFlag)
  {
    subject = "Errors in creating invoices for "+month+" month generated on "+now;
    message += errorLog;
  }else
  {
    subject = message;
  }
  emailMe(message,subject);
}



//function createAndEmailLastMonthInvoice()
//{
//  createInvoice("last", null,false,true);
//}




//    REQUIRED: string month: must be "current" or "last"
//    OPTIONAL: string tabName: the exact spelling of the tab you want to generate the invoice 
//              default: if no tab name provided, it runs on the current tab
//    OPTIONAL: bool printPDF: true or false. This indicates if the rental invoice PDF gets emailed to the printer
//              default: set to false by default
//    OPTIONAL: bool emailReport: true or false. This indicates if the rental invoice gets emailed. 
//              default: set to false by default
//function createInvoice(month, tabName, printPDF,emailReport)
function checkAllTabs() {
  var errorLog = "";
  var errorFlag = false;
  var ss = SpreadsheetApp.getActive();
  var allNames = "";
 // var n=0;
  
  var RENTAL_DATA = "1NKVar1I2BhSJd5LdAT5g37w8syj9H22k9BV73wu0hgw";
  var rentalSheet = SpreadsheetApp.openById(RENTAL_DATA); // this gets you the rental log spreadsheet 
  
  for(var n in ss.getSheets()){
    var sheet = ss.getSheets()[n];// look at every sheet in spreadsheet
    var name = sheet.getName();//get name
   //
    if(name !== "Index" && name !== "Template" && name !== "Master Template"){ // exclude some names
     // var error = createInvoice(month, name, printPDF,false);
// var activeSheet = rentalSheet.getSheetByName(tabName);
      var activeSheet = rentalSheet.getSheetByName(name);
      var status = activeSheet.getRange("K2").getValue();
    //  var status = statusRange;
      
      
      
      var month = activeSheet.getRange("L2").getValue();
      
 //     emailMe("checkAllTabs ran","min script ran for "+name+" K2:" +status+" month:"+month);
      
      if (status == "*Report Pending*" && (month == "Last" || month == "Current"))
      {
        activeSheet.getRange("K2").setValue("*Report Generating*");
        testWait();
        createInvoice(month, name, false,true);
      //  emailMe("Set value min script ran for "+name,"checkAllTabs ran");
      }
    //  
    }
  }
  
  
 // emailMe("done","done wtih checkAllTabs");
}



function onEditInvoice(event) {
  // assumes source data in sheet named Needed
  // target sheet of move to named Acquired
  // test column with yes/no is col 4 or D
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = event.source.getActiveSheet();
  var r = event.source.getActiveRange();
  var sheetUpdated = s.getName();

  
  if(sheetUpdated!== "Index" && s.getName() !== "Template" && s.getName() !== "Master Template" && r.getColumn() == 11 && r.getRow() == 2 && r.getValue() == "*Report Pending*")
  {
    emailMe("Trigger ran","Trigger ran "+sheetUpdated);
//    createLastMonthInvoice();
    createInvoice("Last", sheetUpdated, false,true);
  }

}