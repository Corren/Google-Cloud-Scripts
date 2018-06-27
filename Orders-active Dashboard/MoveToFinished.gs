function onEdit(event) {
  // assumes source data in sheet named Needed
  // target sheet of move to named Acquired
  // test column with yes/no is col 4 or D
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = event.source.getActiveSheet();
  var r = event.source.getActiveRange();
  

  if(s.getName() == "Active Orders" && r.getColumn() == 12 && (r.getValue() == "Yes" || r.getValue() == "Completed")) 
  {
    var lastRow = r.getRow(); // gets the last row number
    var numColumns = s.getLastColumn(); // get the last colum number
    var targetSheet = ss.getSheetByName("Finished Orders"); 
    var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1); // gets the value of the first last row on the finished tab
    target.setNote('Last modified: ' + new Date());
     
    var lastCell = s.getRange(lastRow, numColumns); // gets the range of the value of how many rows to transfer for the single order
    
    // gets the range of these items
    var range_ORDER_NUMBER = s.getRange(lastRow, 1);
    var range_COMPANY_NAME = s.getRange(lastRow, 4);
    var range_SALES_SHEET_URL = s.getRange(lastRow, 14);
    var range_TECH = s.getRange(lastRow, 11);
    
    // gets the value fo these items
    var ORDER_NUMBER = range_ORDER_NUMBER.getValue();
    var COMPANY_NAME = range_COMPANY_NAME.getValue();
    var SALES_SHEET_URL = range_SALES_SHEET_URL.getValue();
    var TECH = range_TECH.getValue();
    
    var lastCellValue = lastCell.getValue(); // gets that value
    var rows2Move = 2;  // sets the default to 2
    
    if (lastCellValue > rows2Move) // if the number of rows is more than 2
    {
      rows2Move = lastCellValue; // update the number of rows
    } 
    
    
    
    
        
     ///////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV///////
    ////////////////////// Starts forming Email //////////////////////////
    //////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV////////

    // This is to send an email to the customer
    var subject = "Order: " + ORDER_NUMBER + " Finished for " + COMPANY_NAME + " by "+ TECH;
    var message = "Adepto,\n\n";

    message += TECH + " has finished order " +ORDER_NUMBER+ " for " + COMPANY_NAME + "\n\n";
    message += "Here is the sales sheet for the order:\n";

    message += SALES_SHEET_URL;
    message += "\n\nIt has been removed from the biomed queue on the TV.\n\n";
    message += "\n\n Thank you,\nYour Automated Ordering System";

    var email = "billing@adeptomed.com";
    MailApp.sendEmail(email, subject, message); // {bcc: "rgaform@adeptomed.com", noReply: true}


    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^///////
    //////// End Copies all relevent info to the Biotech tab ///////////
    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^/////// 

    
    
    
    
    
    
    
    
    
    
    


    s.getRange(lastRow, 1, rows2Move, numColumns).moveTo(target); // moves the rows to the finished tab
    s.deleteRow(lastRow); // removes them from the original tab
    
    

    
    
    
  }
  else if(s.getName() == "Active Orders" && r.getColumn() == 12 && r.getValue() == "In progress") 
  {
    var lastRow = r.getRow(); // gets the last row number
    s.getRange(lastRow, 12).setBackground("GREEN"); // sets that cell to green
  }
  else if(s.getName() == "Active Orders" && r.getColumn() == 12 && r.getValue() == "No") 
  {
    var lastRow = r.getRow(); // gets the last row number
    s.getRange(lastRow, 12).setBackground("RED"); // sets that cell to green
  }
}