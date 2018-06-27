function getAllTabNames(month,printPDF) {
  var errorLog = "";
  var errorFlag = false;
  var ss = SpreadsheetApp.getActive();
  var allNames = "";
 // var n=0;
  var message = "Done creating invoices for "+month+" month \n\n";
  for(var n in ss.getSheets()){
    var sheet = ss.getSheets()[n];// look at every sheet in spreadsheet
    var name = sheet.getName();//get name
    if(name !== "Index" && name !== "Template" && name !== "Master Template"){ // exclude some names
      message += name +"\n";
    }
  }
  var now = new Date();
  
  var subject = "list of all tab names";

  emailMe(message,subject);
}

function justThisOne()
{
  var tabName = SpreadsheetApp.getActiveSheet().getName();
  
  
  createInvoice("Last", tabName, false,true);
}