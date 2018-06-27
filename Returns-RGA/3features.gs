function RGAsubmission(e) {
    Logger.log("====================== Starting RGAsubmission Function ======================");
    // This is to generate the ID number
    var columnNumber = 88; // this is the first column that holds the next RGA number
    var sp = PropertiesService.getScriptProperties();
    var sheet = SpreadsheetApp.getActiveSheet()
    var ss = SpreadsheetApp.openById("1dxt6c62NJrjiEOa_Ye5xAze14UQTgLBn26IunymDAQc"); // this selects the spreadsheet by its ID#

    var s = SpreadsheetApp.getActiveSheet();
    var headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];

    // added these variables to easily send to the fedex label function
    var timestamp = e.namedValues[headers[0]].toString();
    var customerEmail = e.namedValues[headers[1]].toString();
    var customerCompany = e.namedValues[headers[2]].toString();
    var customerCompanySanitized = customerCompany.replace(/[^a-zA-Z ]/g, "");
    var customerContact = e.namedValues[headers[3]].toString();
    var customerContactSanitized = customerContact.replace(/[^a-zA-Z ]/g, "");
    var customerAddress1 = e.namedValues[headers[4]].toString();
    var customerAddress2 = e.namedValues[headers[5]].toString();
    var customerCity = e.namedValues[headers[6]].toString();
    var customerState = e.namedValues[headers[7]].toString();
    var customerZip = e.namedValues[headers[8]].toString();
    var customerLabelsRequested = e.namedValues[headers[9]].toString();
    var customerPhone = e.namedValues[headers[10]].toString();
    var customerApprovedCost = e.namedValues[headers[11]].toString();
    var customerPO = e.namedValues[headers[12]].toString();
    var RGAType = e.namedValues[headers[13]].toString(); // gets the RGA catagory
    var shipSpeed = "FEDEX_GROUND";
    var pumpIsHot = false;

    //  var catagory = e.namedValues[headers[11]].toString();
    var patientIncident = false;
    // This sorts the responce by RGA type
    if (RGAType.indexOf("Rental Returns") >= 0) {
        var rga = sheet.getRange("CL1").getValue(); // gets the next rental rga #
        e.range.offset(0, columnNumber + 1, 1, 1).setValue(rga);
        var tab = ss.getSheetByName('50000 - Rental Returns'); //replace with source Sheet tab name
        Logger.log("Tab: " + '50000 - Rental Returns');
    } else if (RGAType.indexOf("Service (if you need a device repaired or P.M./ Re-certified )") >= 0) {
        var rga = sheet.getRange("CM1").getValue(); // gets the next service rga #
        e.range.offset(0, columnNumber + 2, 1, 1).setValue(rga);
        var tab = ss.getSheetByName('60000 - Services'); //replace with source Sheet tab name
        Logger.log("Tab: " + '60000 - Services');
    } else if (RGAType.indexOf("Warranty (on serviced, repaired, or purchased pumps from Adepto)") >= 0) {
        var rga = sheet.getRange("CN1").getValue(); // gets the next warranty rga #
        e.range.offset(0, columnNumber + 3, 1, 1).setValue(rga);
        var tab = ss.getSheetByName('90000 - Warranty/Malfunctions'); //replace with source Sheet tab name
        Logger.log("Tab: " + '90000 - Warranty/Malfunctions');
    }

    Logger.log("");
    Logger.log("---- Start Form variables ----");
    Logger.log("timestamp: " + timestamp);
    Logger.log("customerEmail: " + customerEmail);
    Logger.log("customerCompany: " + customerCompany);
    Logger.log("customerCompanySanitized: " + customerCompanySanitized);
    Logger.log("customerContact: " + customerContact);
    Logger.log("customerContactSanitized: " + customerContactSanitized);
    Logger.log("customerAddress1: " + customerAddress1);
    Logger.log("customerAddress2: " + customerAddress2);
    Logger.log("customerCity: " + customerCity);
    Logger.log("customerState: " + customerState);
    Logger.log("customerZip: " + customerZip);
    Logger.log("customerLabelsRequested: " + customerLabelsRequested);
    Logger.log("customerPhone: " + customerPhone);
    Logger.log("customerApprovedCost: " + customerApprovedCost);
    Logger.log("customerPO: " + customerPO);
    Logger.log("RGAType: " + RGAType);
    Logger.log("rga: " + rga);
    Logger.log("shipSpeed: " + shipSpeed);
    Logger.log("---- End Form variables ----");
    Logger.log("");
    /////////////////// Code to copy data to the next tab //////////////////////

    var requiredRows = 0; // holds the number of how many rows will be inserted (1 per device) on the respective tab
    //  var model = 9; // model # in col 9, S/N in col 10

    var items = ""; // going to store all the devices and serial numbers (for the email to the customer)
    var modelCol = 14; // colum that holds the first model #. model # in col 14, S/N in col 15 (note: col A = col 0, not col 1)
    // This loops through all 15 possible devices returns the number of how many are 
    for (var l = 0; l < 15; l++) {
        var colToNextModel = l * 5;
        if (e.namedValues[headers[modelCol + l * 5]].toString()) {
            requiredRows++;
            items += e.namedValues[headers[modelCol + colToNextModel]].toString() + " - " + e.namedValues[headers[modelCol + colToNextModel + 1]].toString() + "\n";
            Logger.log("Checking if pump is 'hot'");
            if (hotPumpCheck(e.namedValues[headers[modelCol + colToNextModel]].toString())) // check if the pump is on the hit pump list
            {
                pumpIsHot = true;
                if (customerState != "KS" && customerState != "MO" && customerState != "NE" && customerState != "IA") {
                    Logger.log("Pump is hot and not in one of the 4 local states. 1 day ground shipping applied.");
                    Logger.log("");
                    shipSpeed = "FEDEX_NEXT_DAY_END_OF_DAY";
                } else {
                    Logger.log("Pump is hot but it is in a state that has 1 day ground shipping so shipping has not changed.");
                    Logger.log("");
                }
            }
        }
    }

    Logger.log("");
    Logger.log("Number of rows required to insert into the respective RGA tab: " + requiredRows);
    Logger.log("");
    Logger.log("Checking for patient incidents");
    // this checks it there was a PATIENT INCIDENT
    var patientIncidentDevices = "";
    for (var rr2 = 0; rr2 < requiredRows; rr2++) {
        var colToNextModel = rr2 * 5;
        if (e.namedValues[headers[modelCol + (colToNextModel) + 2]].toString().indexOf("PATIENT INCIDENT") > -1) {
            patientIncident = true;

            patientIncidentDevices += e.namedValues[headers[modelCol + (colToNextModel)]].toString() + " - "; // Model  - Device 1
            patientIncidentDevices += e.namedValues[headers[modelCol + (colToNextModel) + 1]].toString() + " - "; // Serial Number - Device 1
            patientIncidentDevices += e.namedValues[headers[modelCol + (colToNextModel) + 2]].toString() + "\n"; // Description of problem(s) with Device 1

            tab = ss.getSheetByName('911 - PATIENT INCIDENT'); //replace with source Sheet tab name
            Logger.log("patient incident device: " + patientIncidentDevices);
        }
    }
    Logger.log("");
    var lastRow = tab.getLastRow() + 1; // returns integer — the last row of the sheet that contains content
    Logger.log("Adding 1 new line to the RGA tab (1 line before every order - gets filled black to seperate the RGAs)");
    tab.insertRowsAfter(lastRow, 1); // inserts a new line after the active last row so we never run out of rows on the RGA form.
    Logger.log("");


    // This is inserting the amount of rows needed, if any. Using requiredRows * 2 value (we add a blank line between each row so times 2). After the last row of data.
    if (requiredRows > 0) {
        tab.insertRowsAfter(lastRow, requiredRows * 2);
        Logger.log("Adding " + requiredRows * 2 + " rows to the RGA tab ");
    }

    var blackOutRow = tab.getLastRow() + 1; // returns the next row to write date to
    Logger.log("");
    Logger.log("Adding black fill to seperate the RGAs");
    // loops through and fills the line between RGA with a black empty fill  
    for (var lines2fill = 1; lines2fill <= 20; lines2fill++) {
        tab.getRange(blackOutRow, lines2fill).setBackgroundRGB(0, 0, 0); // fills the first 20 cells with black fill
    }

    Logger.log("");
    Logger.log("Now filling out the respective RGA tab with the form info");
    // this loops through and fills all the extra lines
    for (var rr = 0; rr < requiredRows; rr++) {
        var colToNextModel = rr * 5;
        var modelSN = e.namedValues[headers[modelCol + (colToNextModel)]].toString() + " - " + e.namedValues[headers[modelCol + (colToNextModel) + 1]].toString(); // [model - sn]
        var newLine = [e.namedValues[headers[modelCol + (colToNextModel) + 1]].toString(),
            //modelSN, // model - s/n (additional refference at the begining) // line replaced with just the S/N b/c andrew wants to cpy it with out the scrolling over to the right
            rga, // RGA number assigned
            e.namedValues[headers[0]].toString(), // Timestamp
            e.namedValues[headers[1]].toString(), // Email Address
            e.namedValues[headers[2]].toString(), // Company name
            e.namedValues[headers[3]].toString(), // Your name or main contact
            e.namedValues[headers[4]].toString() + "\n" + // Address 1 (return address)
            e.namedValues[headers[5]].toString() + "\n" + // Address 2
            e.namedValues[headers[6]].toString() + "\n" + // City
            e.namedValues[headers[7]].toString() + "\n" + // State
            e.namedValues[headers[8]].toString() + "\n" + // Zip
            e.namedValues[headers[9]].toString(), // # of fedex labels needed
            e.namedValues[headers[10]].toString(), // Contact phone number
            e.namedValues[headers[11]].toString(), // Approved repair cost per device (default is $35)
            e.namedValues[headers[12]].toString(), // PO # for your records (optional)
            e.namedValues[headers[13]].toString(), // Category  (Please fill out form again for each category. ie. rental return and service request require 2 separate RGA submissions)
            e.namedValues[headers[modelCol + (colToNextModel)]].toString(), // Model  - Device 1
            e.namedValues[headers[modelCol + (colToNextModel) + 1]].toString(), // Serial Number - Device 1
            e.namedValues[headers[modelCol + (colToNextModel) + 2]].toString(), // Description of problem(s) with Device 1
            e.namedValues[headers[modelCol + (colToNextModel) + 3]].toString()
        ]; // Additional notes for device 1




        var afterLastRow = tab.getLastRow() + 1; // returns the next row to write date to
        // moved above the for loop
        // loops through and fills the line between RGA with a black empty fill  
        //  for (var lines2fill =1; lines2fill<=20;lines2fill++)
        //  {
        //    tab.getRange(afterLastRow, lines2fill).setBackgroundRGB(0,0,0); // fills the first 20 cells with black fill
        //  }

        afterLastRow += 1; // This increments the next row to write data to

        Logger.log("Writing out: " + newLine);
        // loops through the array and writes each item to its respective cell 1 at a time    
        // I have lines2write + 6 as there are 6 columns that are internal notes coloms that are filled before the data
        for (var lines2write = 0; lines2write <= 14; lines2write++) // 0-14 as there are 15 possible pumps they can send in
        {
            tab.getRange(afterLastRow, 6 + lines2write).setValue(newLine[lines2write]); // inserts the data into the next line 
        }
    }
    Logger.log("Done adding form info the the RGA tab");
    /////////////////// end of tab data copy ///////////////////////////////////


    ////////////// Stats forming Email ///////////////////////////////////////
    sp.setProperty("Event Id", rga + 1);
    Logger.log("");
    Logger.log("Forming email to send to the customer");
    // This is to send an email to the customer
    var subject = "Adepto Medical RGA: " + rga;
    var message = "Dear " + customerCompany + ",\n\n";
    if (patientIncident) {
        message += "\nWARNING: PLEASE CALL US WITH MORE INFORMATION ABOUT THE PATIENT INCIDENT DEVICE.\n\n";
    }
    message += "Thank you for the opportunity to serve you.\n";
    message += "Below are devices we received information on, should you have any additional questions or concerns please contact us at service@adeptomed.com or call us at (913) 261-9933 and reference your RGA Number.\n\n";

    // Remember to replace this email address with your own email address


    // The variable e holds all the form values in an array.
    // Loop through the array and append values to the body.
    // message += e.namedValues[headers[2]].toString()+":\n\n"; // company name

    // message += "You have successfully submitted a service request for the following:\n";
    /*  
      var items = "";// going to store all the devices and serial numbers
      var modelCol = 14; // model # in col 14, S/N in col 15 (note: col A = col 0, not col 1)

      
      // This loops through all 15 possible devices and adds them to the email if they are not empty
      for (var x = 0; x < 15; x++)
      {
        if (e.namedValues[headers[modelCol]].toString())
        {
          items += e.namedValues[headers[modelCol]].toString() + " - " +e.namedValues[headers[modelCol+1]].toString() + "\n";
          modelCol+=5;
        }
      }
      */

    var fileName = "FedEx_Shipping_Label_RGA_" + rga + ".pdf";

    message += items + "\n\n";
    Logger.log("");
    Logger.log("Checking if the customer is in a valid state to get the shipping label");
    if (customerState != 'Other' && customerState != 'HI' && customerState != 'AK' && customerLabelsRequested != "No") {
        Logger.log("Customer is in a valid state.");
        message += "Use the attached FedEx shipping label [" + fileName + "] to send the items to the following address:\n\n";
        message += "  Adepto Medical \n  ATTN: RGA: " + rga + "\n  3120 Terrace St. \n  Kansas City, MO 64111 \n\n";
        message += 'Please make sure that your package is no heaver than 50 LBS and no larger than 24" x 24" x 24" in dimensions.\n';
        message += "While we will take care of the shipping cost on our end, please make sure you package the items well to prevent any damages.\n";
        message += "For best packaging practices, please see http://www.fedex.com/ca_english/shippingguide/preparepackage/ \n\n";
    } else {
        Logger.log("Customer is not in a valid state or has selected not to recieve a label.");
        message += "Send the items to the following address:\n\n";
        message += "  Adepto Medical \n  ATTN: RGA: " + rga + "\n  3120 Terrace St. \n  Kansas City, MO 64111 \n\n";
    }
    message += "Thank you, \nAdepto Medical\n\n";
    message += "Please do not respond to this email as it is unmonitored. If you have any questions, please feel free to give us a call at (913) 261-9933 or email us at sales@adeptomed.com";

    // Send the email

    if (patientIncident) {
        var emergencyEmails = "tommy@adeptomed.com,clint@adeptomed.com,andrew@adeptomed.com,jonathan@adeptomed.com";
        var emergencySubject = "PATIENT INCIDENT from " + customerCompany + " RGA: " + rga;
        var emergencyMessage = "Look out for these devices:\n\n" + patientIncidentDevices;
        MailApp.sendEmail(emergencyEmails, emergencySubject, emergencyMessage, {
            noReply: true
        });
    }

    //  if (email == "tommy@adeptomed.com" || email == "jonathan@adeptomed.com")
    //  {// send with fedex label

    Logger.log("");
    Logger.log("");
    Logger.log("Attempting to get FedEx base64 label");
    try {
        if (customerLabelsRequested.indexOf("No") >= 0) {
            //no label
        } else {

            var base64 = GetFedExLabel(customerCompanySanitized, customerPhone, customerAddress1, customerAddress2, customerCity, customerState, customerZip, rga, shipSpeed, timestamp);
            //converts base64 into a blob and saves as a pdf
            var blob = Utilities.newBlob(Utilities.base64Decode(base64), 'application/PDF', fileName);
            var dir = DriveApp.getFolderById("1drWVHX_SluCr42LTO9k7qkpvASHPEylB");
            var file = dir.createFile(blob);
        }

    } catch (e) {
        emailMe(e + "\n\n\n\n-------------\n\n Log message: " + Logger.getLog(), "error creating FedEx label");
        Logger.log("Error getting the FedEx label");
    } finally {
        Logger.log("");
        Logger.log("");
        Logger.log("Now sending the customer the email");
        try {
            if (customerLabelsRequested.indexOf("No") >= 0) {
                Logger.log("");
                Logger.log("-- sending with no label");
                MailApp.sendEmail(customerEmail, subject, message, {
                    bcc: "rgaform@adeptomed.com,jonathan@adeptomed.com,brandi@adeptomed.com,clint@adeptomed.com,tommy@adeptomed.com,marni@adeptomed.com",
                    noReply: true
                });
                //no label
            } else {
                Logger.log("");
                Logger.log("-- sending with label");
                MailApp.sendEmail(customerEmail, subject, message, {
                    bcc: "rgaform@adeptomed.com,jonathan@adeptomed.com,brandi@adeptomed.com,clint@adeptomed.com,tommy@adeptomed.com,marni@adeptomed.com",
                    noReply: true,
                    attachments: [blob]
                });
            }
        } catch (e) {
            Logger.log("There was an error sending the attachment.");
            emailMe(Logger.getLog(), "Error sending RGA email");

        }
        //  }else{
        //      MailApp.sendEmail(email, subject, message, {bcc: "rgaform@adeptomed.com,jonathan@adeptomed.com,brandi@adeptomed.com,clint@adeptomed.com,tommy@adeptomed.com,marni@adeptomed.com", noReply: true}); // {bcc: "rgaform@adeptomed.com", noReply: true}
        //  }

        // MailApp.sendEmail("jonathan@adeptomed.com", "jFilter-"+subject, message, {noReply: true});

        // This constant is written in column C for rows for which an email
        // has been sent successfully.
        var EMAIL_SENT = "EMAIL_SENT";
        e.range.offset(0, columnNumber, 1, 1).setValue(EMAIL_SENT);
        Logger.log("");
        Logger.log("====================== Finishing RGAsubmission Function ======================");
    }
    ////////////// End of sending email /////////////////////////
}

function SanitizeString(word) {
    Logger.log("Sanatizing [%s]", word);
    var cleaned = word.replace(/[^a-zA-Z ]/g, "");
    Logger.log("Cleaned to [%s]", word);
    return cleaned;
}