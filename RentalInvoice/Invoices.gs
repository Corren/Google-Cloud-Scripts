// This function generates the monthly rental bill for the desired customer
// The customer must have their own tab on the rental log
// The rental log must be in the specified format with rentals starting on line 4
// The contents of the first 5 coloms must not be blank on each row to be considered valid
// input:
//    REQUIRED: string month: must be "Current" or "Last"
//    OPTIONAL: string tabName: the exact spelling of the tab you want to generate the invoice 
//              default: if no tab name provided, it runs on the current tab
//    OPTIONAL: bool printPDF: true or false. This indicates if the rental invoice PDF gets emailed to the printer
//              default: set to false by default
//    OPTIONAL: bool emailReport: true or false. This indicates if the rental invoice gets emailed. 
//              default: set to false by default
function createInvoice(month, tabName, printPDF, emailReport,errorOverride) {
   Logger.log("===Starting generateInvoice function======");
    // Google Doc id from the document template and folders
    var SOURCE_TEMPLATE = "1Cq-yXXtgd6GMETpJQ8S0FL0A-NXZmUDb9f6Scz0ZVk4"; // test template
    var RENTAL_DATA = "1NKVar1I2BhSJd5LdAT5g37w8syj9H22k9BV73wu0hgw";
    var PAST_INVOICE_FOLDER_ID = "1ltr1mn3-m7Shyq0kBi_Ac2SOyd9a1ePW";
    var EXTRA_CHARGE_FEE = 1;

    var summary = "";
    ///////////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//////////////////
    /////////////////// Start getting Variables from the rental log //////////////////////
    ///////////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//////////////////
    var tabNameParameter = tabName;
    // Refference number logic
    var rentalSheet = SpreadsheetApp.openById(RENTAL_DATA); // this gets you the rental log spreadsheet 
    var indexTab = rentalSheet.getSheetByName('Index'); // gets the tab that holds the refference counter
    var tabNameActive = "";

    if (tabName == null) {
        var activeSheet = SpreadsheetApp.getActiveSheet();
        //      tabNameActive = activeSheet.getRange("B2").getValue();
      Logger.log("activeSheet tab: null");
      Logger.log("activeSheet tab set to current tab: "+activeSheet.getName());
    } else {
        var activeSheet = rentalSheet.getSheetByName(tabName);
      Logger.log("activeSheet tab: "+tabName);
    }
     Logger.log("activeSheet: "+activeSheet);


    if (emailReport == true) {
        var sendEmailReport = true;
    } else {
        var sendEmailReport = false;
    }
   Logger.log("sendEmailReport: "+sendEmailReport);

    if (printPDF == true) {
        var print = true; // should the item be printed to the HP printer automatically
    } else {
        var print = false; // should the item be printed to the HP printer automatically
    }
   Logger.log("print: "+print);

    // This is an override that prevents the generated invoice from showing red on out of bounds rentals
    if (errorOverride == true) {
        var override = true;
    } else {
        var override = false;
    }
   Logger.log("Override: "+override);

    if (sendEmailReport) {
        var REF_NUM = indexTab.getRange("B3").getValue(); // gets the next available refference number
        indexTab.getRange('B3').setValue(REF_NUM + 1); // increases the refference number counter

        var FACILITY = activeSheet.getRange("B2").getValue(); // gets the facility name
        tabName = FACILITY;

        var CONTACT = activeSheet.getRange("E2").getValue(); // gets the facility name
        var CONTACT_EMAIL = activeSheet.getRange("H2").getValue(); // gets customer email
        var PO = activeSheet.getRange("J2").getValue(); // gets the facility po
    }
  Logger.log("get [Apply $1+/Mo] logic");
  
  
  var EXTRA_CHARGE = activeSheet.getRange("K1").getValue(); // gets the facility name
  if (EXTRA_CHARGE == "Apply $1+/Mo")
  {var extraChargeAmount = EXTRA_CHARGE_FEE;
  }
  else
  {var extraChargeAmount = 0;
  }
  Logger.log("extra charge read as ["+EXTRA_CHARGE+"] and set to: "+extraChargeAmount);
  
  Logger.log("get po logic");
    var PO_TITLE = ""; // If there is no PO, then PO is not in the invoice title
    var error = false; // error flag to check rental parameters
    var errorMessage = "";
    var errorLog = "";
    if (sendEmailReport) {
        // This checks if there is a PO, if so, add the PO lable in the invoice title
        if (!activeSheet.getRange("J2").isBlank()) {
            PO_TITLE = " PO: " + PO;
        }

        // This is the title for invoice folder
        var FACILITY_PO_INVOICE_FOLDER_NAME = FACILITY + PO_TITLE;
    }
   Logger.log("date logic");
    // last month date logic
    var date = new Date();
    var currentYear = date.getYear();
    var currentDay = date.getDate();
    var currentMonth = date.getMonth();

    if (month == "Current") {
      Logger.log("month: "+month);
        var billingYear = currentYear;
        var billingMonth = currentMonth;

        // Billing date range logic
        // new Date(year, month, day, hours, minutes, seconds, milliseconds);
        var millisecondsPerDay = 24 * 3600 * 1000;
        var firstDayOfNextMonth = new Date(currentYear, currentMonth + 1, 1, 0, 0, 0, 0);
        // This gets the last day of last month regaurdless of number of days in that month 

        var lastDayOfBilling = new Date(firstDayOfNextMonth.getTime() - 1 * (millisecondsPerDay));
        var net30DueDate = new Date(firstDayOfNextMonth.getTime() + 29 * (millisecondsPerDay));
    } else if (month == "twoMonthsAgo"){
      
      Logger.log("month: "+month);
            if (currentMonth == 0) // if you are billing dec in jan
        {
            var billingYear = currentYear - 1;
            var billingMonth = 11; // months are from 0(jan) - 11(dec)
        }else if (currentMonth == 1){
                    var billingYear = currentYear - 1;
            var billingMonth = 10; // months are from 0(jan) - 11(dec)
        }else {
            var billingYear = currentYear;
           // var billingMonth = date.getMonth() - 1; /////////for 1 month ago
            var billingMonth = date.getMonth() - 2;/////////for 2 month ago
        }

        // Billing date range logic
        // new Date(year, month, day, hours, minutes, seconds, milliseconds);
        var millisecondsPerDay = 24 * 3600 * 1000;
     //   var firstDayOfCurrentMonth = new Date(currentYear, currentMonth, 1, 0, 0, 0, 0);/////////for 1 month ago
      var firstDayOfCurrentMonth = new Date(currentYear, currentMonth-1, 1, 0, 0, 0, 0);/////////for 2 month ago
        // This gets the last day of last month regaurdless of number of days in that month 
     //   var lastDayOfBilling = new Date(firstDayOfCurrentMonth.getTime() - 1 * (millisecondsPerDay)); // change this value to get the last day /////////for 1 month ago
     //   var net30DueDate = new Date(firstDayOfCurrentMonth.getTime() + 29 * (millisecondsPerDay)); /////////for 1 month ago
        var lastDayOfBilling = new Date(firstDayOfCurrentMonth.getTime() - 1 * (millisecondsPerDay)); /////////for 2 month ago
        var net30DueDate = new Date(firstDayOfCurrentMonth.getTime() + 29 * (millisecondsPerDay));/////////for 2 month ago
    }  else if (month == "Last"){
      Logger.log("month: "+month);
        if (currentMonth == 0) // if you are billing dec in jan
        {
            var billingYear = currentYear - 1;
            var billingMonth = 11; // months are from 0(jan) - 11(dec)
        } else {
            var billingYear = currentYear;
            var billingMonth = date.getMonth() - 1; /////////for 1 month ago
        }

        // Billing date range logic
        // new Date(year, month, day, hours, minutes, seconds, milliseconds);
        var millisecondsPerDay = 24 * 3600 * 1000;
        var firstDayOfCurrentMonth = new Date(currentYear, currentMonth, 1, 0, 0, 0, 0);/////////for 1 month ago
        // This gets the last day of last month regaurdless of number of days in that month 
        var lastDayOfBilling = new Date(firstDayOfCurrentMonth.getTime() - 1 * (millisecondsPerDay)); // change this value to get the last day /////////for 1 month ago
        var net30DueDate = new Date(firstDayOfCurrentMonth.getTime() + 29 * (millisecondsPerDay)); /////////for 1 month ago
    }else
    {
      Logger.log("error! month [%s] not valid",month);
    }
      

    // sets the first day of the month for billing
    var firstDayOfBilling = new Date(billingYear, billingMonth, 1, 0, 0, 0, 0);

    // rental counters
    var itemCounter = 0;
    var grandRentalTotal = 0;

    if (sendEmailReport) {
        // Invoice creation logic
        var INVOICE_NAME = "Rental_Invoice_for_" + FACILITY + PO_TITLE + "_Date:" + (parseInt(billingMonth) + 1) + "-" + billingYear + "_REF:" + REF_NUM;
        var spreadsheetTemplate = SpreadsheetApp.openById(SOURCE_TEMPLATE); // gets the template for the order sheet
        var invoiceSheet = spreadsheetTemplate.copy(INVOICE_NAME); // creates a title for the new sheet
        var pdfName = INVOICE_NAME + ".pdf";
      
      try{
         Logger.log("Checking if the extra charge does not apply and removing warrning");
      if (extraChargeAmount == 0)
      {
        invoiceSheet.getActiveSheet().getRange('B12').setValue("");
        Logger.log("Removed extra charge warrning");
      }
      else
      {
        Logger.log("Left extra charge warrning");
      }}catch(e)
      {
        Logger.log("Error while fixing the disclaimer: "+e);
      }
    }
    // gets the total number of rows with rentals on them
    var lastRowRental = activeSheet.getLastRow();
    var firstRowRental = 4; // the first row that rentals are logged
    var firstInvoiceCol = 2; // first colum that has data (item colum)
    var lastRow = 9; // last row on the invoice with data in it

    ///////////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV/////////////
    //////////////////////////// start error ////////////////////////////////
    ///////////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV/////////////

      Logger.log("starting to search each row: ");

    // adds each item to rental invoice if applicable
    for (var row = firstRowRental; row <= lastRowRental; row++) {
            Logger.log("row: %s of "+lastRowRental, row);

        // logic for blank dates and dates in the month following
        if (activeSheet.getRange(row, 4).isBlank()) {
            var itemReturnDate = lastDayOfBilling;
        } else {
            var itemReturnDate = new Date(activeSheet.getRange(row, 4).getValue());
        }

        if (itemReturnDate.valueOf() > lastDayOfBilling) {
            itemReturnDate = lastDayOfBilling;
        }

        // makes sure that the reurn date is not before the rental period and check to make sure that the row is not blank
        var itemShipDate = new Date(activeSheet.getRange(row, 3).getValue());
        if (itemShipDate.valueOf() < firstDayOfBilling) {
            itemShipDate = firstDayOfBilling;
          
        }
      Logger.log("itemShipDate: "+itemShipDate);
      Logger.log("itemReturnDate: "+itemReturnDate);
       Logger.log("firstDayOfBilling: "+firstDayOfBilling);
       Logger.log("lastDayOfBilling: "+lastDayOfBilling);

        if (itemReturnDate.valueOf() >= firstDayOfBilling && itemShipDate.valueOf() <= lastDayOfBilling) {
          Logger.log("item is valid for the selected billing month: "+itemShipDate);
            // gets all the rental data from the sheet
            var itemModel = activeSheet.getRange(row, 1).getValue();
            var itemSN = activeSheet.getRange(row, 2).getValue();
          
            var itemMonthlyRate = Number(activeSheet.getRange(row, 5).getValue())+extraChargeAmount;
            var itemDailyRate = (itemMonthlyRate) / 30;

            //days billed in month logic
            var daysBilled = daysBetween(itemShipDate, itemReturnDate);

            //      if (daysBilled > 30)
            //      {
            //        daysBilled = 30;
            //      }
            var totalCostForMonth = Number((daysBilled) * Number(itemDailyRate));
            itemCounter++;

            // sets the values on the invoice sheet
            if (sendEmailReport) {

              
                var itemRange = invoiceSheet.getActiveSheet().getRange(lastRow, firstInvoiceCol, 1, 9);
                itemRange.setValues([
                    [itemCounter, itemModel, itemSN, itemShipDate, itemReturnDate, daysBilled, itemMonthlyRate,  (itemDailyRate), "$" + totalCostForMonth.toFixed(4)]
                ]);

                if (itemCounter % 2 == 0) {
                    // fills every odd number with grey  (255, 255, 255)(239, 239, 239)
                    itemRange.setBackgroundRGB(255, 255, 255); // fills the first 10 cells with black fill
                } else {
                    // fills every odd number with grey 
                    itemRange.setBackgroundRGB(239, 239, 239); // fills the first 10 cells with grey fill
                }
            }
            grandRentalTotal += totalCostForMonth;

            if (itemMonthlyRate < 10 || itemMonthlyRate > 280 || daysBilled > 31) {
                error = true;
                if (itemMonthlyRate < 10 || itemMonthlyRate > 280) {
                    errorMessage += " Error SN: " + itemSN + " monthly rate out of range at $" + itemMonthlyRate + "/month.\n";
                    errorLog += tabName + " SN: " + itemSN + " monthly rate out of range at $" + itemMonthlyRate + "/month.\n";
                }

                if (daysBilled > 31) {
                    errorMessage += " Error SN: " + itemSN + " being billed for over 31 days.\n";
                    errorLog += tabName + " SN: " + itemSN + " being billed for over 31 days.\n";
                }
              
                // fills the out or range items with red fill, override means dont fill the invoice out of bounds with red fill
                if (sendEmailReport && !override) {
                    itemRange.setBackgroundRGB(255, 0, 0); // fills the first 10 cells with red fill
                }
              testWait();
              addTabColor(tabNameParameter, "red"); // changes tab to red since there was an error
            }
          else
          {
            
          }
            if (sendEmailReport) {
                // This is inserting the amount of rows needed, if any. Using requiredRows value. After the last row of data.
                invoiceSheet.insertRowsAfter(lastRow, 1);
                lastRow++;
            }
        }
      if (error) {addTabColor(tabNameParameter, "green");
                       Logger.log("error on parameters: ");} // changes tab color to green since it is all clear
    }

    // This sets the values for the invoice from the rental log
    if (sendEmailReport) {
        testWait();
        invoiceSheet.getActiveSheet().getRange('E2:E6').setValues([
            ["[" + pdfName + "]"],
            [FACILITY],
            [PO],
            [CONTACT],
            [CONTACT_EMAIL]
        ]);
        invoiceSheet.getActiveSheet().getRange('J3:J6').setValues([
            [firstDayOfBilling],
            [lastDayOfBilling],
            [net30DueDate],
            ["$" + grandRentalTotal.toFixed(2)]
        ]);

        invoiceSheet.getActiveSheet().getRange('K6').setValue(itemCounter);

     

        ///////////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^/////////////
        ///////////////////////////// End error ////////////////////////////////
        ///////////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^/////////////


        ///////////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV/////////////
        //////////// Start moving the invoice to the facility/po folder ////////////////
        ///////////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV/////////////
 Logger.log("Starting to move invoice folder");
        // gets the folder containing all the past invoices
        var par_fdr = DriveApp.getFolderById(PAST_INVOICE_FOLDER_ID);

        // This checks if the folder for the facility/po exist. if not, creates it
        try { // gets the folder for the facility/po if exist
            var newFdr = par_fdr.getFoldersByName(FACILITY_PO_INVOICE_FOLDER_NAME).next();
        } catch (e) { // creates the folder for the facility/po if not exist
            var newFdr = par_fdr.createFolder(FACILITY_PO_INVOICE_FOLDER_NAME);
        }

        // Move to (jonathan's) My Drive/Monthly Rental Invoices/Past Invoices/ FacilityPO
        var invoiceSheetFile = DriveApp.getFileById(invoiceSheet.getId()); // gets the new invoice file created
        newFdr.addFile(invoiceSheetFile); // adds the invoice to the facility/po folder
        DriveApp.getRootFolder().removeFile(invoiceSheetFile); // removes the copy of invoice after it was moved
        var INVOICE_URL = invoiceSheetFile.getUrl(); // gets the URL for that invoice
 Logger.log("Moved to invoice folder");

        ///////////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^/////////////
        ///////////// End moving the invoice to the facility/po folder ////////////////
        ///////////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^/////////////

        activeSheet.getRange("K2:L2").setValues([
            [month + " Months Invoice:", INVOICE_URL]
        ]);
      Logger.log("set URL for new sheet: "+ INVOICE_URL);


        //////////////////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV///////
        ///////////////////////////////// Starts forming Email //////////////////////////
        /////////////////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV////////

        var finished = false;
        while (!finished) {
            if (invoiceSheet.getActiveSheet().getRange('J6').getValue() == "*pending*") {
                testWait();
            } else /// this makes sure that the email is not sent until the total has been updated.
            {
                finished = true;
                // Make zee PDF, currently called "Weekly status.pdf"
                // When I'm smart, filename will include a date and project name
                var pdf = DriveApp.getFileById(invoiceSheetFile.getId()).getAs('application/pdf').getBytes();

                var attach = {
                    fileName: pdfName,
                    content: pdf,
                    mimeType: 'application/pdf'
                };

                // This is to send an email to the customer
                var subject = "";
                if (error) // if there was an error in the rental range, then add alert to the email
                {
                    subject += "ALERT! OUT OF RANGE: ";
                }
                subject += "\nAdepto Medical IV Pump Rental Bill";
                var message = FACILITY + ":\n\n";

                message += "Contact: " + CONTACT + "\n";
                message += "Email: " + CONTACT_EMAIL + "\n\n";

                if (error) // if there was an error in the rental range, then add alert to the email
                {
                    message += errorMessage;
                }

                message += "\n\nPlease reference attachment [" + pdfName + "] for itemized rental breakdown of your monthly bill. ";
                message += "The total balance for the month is $" + grandRentalTotal.toFixed(2) + ".\n\n";

                message += "If you have any questions or concerns regarding your bill, please feel free to give us a call at (913) 261-9933 or email us at rentals@adeptomed.com";
                message += "\n\nThank you,\nAdepto Team";

                var email = "jonathan@adeptomed.com,marni@adeptomed.com,brandi@adeptomed.com";

                MailApp.sendEmail(email, subject, message, {
                    attachments: [attach]
                }); // {bcc: "rgaform@adeptomed.com", noReply: true}
                if (grandRentalTotal == 0)// Dont print the invoice if the total is $0
                {print = false;}
                if (print) {
                    var printer_email = "adeptoinvoices@hpeprint.com";
                    MailApp.sendEmail(printer_email, " ", "", {
                        attachments: [attach]
                    });
                }

            }
        }
    } else {
    }
    summary = ",Items," + itemCounter + ",Total for "+month + " month,"+grandRentalTotal.toFixed(2);
    /////////////////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^///////
    ////////////////////////////////// End sending email //////////////////////////
    /////////////////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^/////// 
    if (true) {
            Logger.log("Finished with report. summary: "+summary);
     // emailMe("Finished with report. summary: "+summary+"\n\n --------------------- \n\n "+Logger.getLog(),"Done with ["+tabName+"]")
        return summary;
    }

    //  if (error)
    //  {
    //    return errorLog;
    // }
    // else
    //  {
    //   return "No errors detected";
    //  }
}