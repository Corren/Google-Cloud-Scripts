// known issues:
// the notes section for the accessories on the sales sheet duplicates the notes for the pump.
// this is because i do not have a notes section on the oprder form.


function myFunction(e) {

    // Google Doc id from the document template
    // (Get ids from the URL)
    //var SOURCE_TEMPLATE = "1lP_vNLwlp2uEDg_VORWNBVrKiYbe__tgFrdN1UUkkqM"; // offical template
    var SOURCE_TEMPLATE = "1o3JufTOGLpRFoQnImI3GoT9uYDH3-mquXw7CjqUq4UU"; // test template

    var FORM_DATA = "1oUFi-zQm-XfvVwvHyZu_cArykys5KEo5GoGzKmu99fA";


    ///////////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV/////////////
    /////////////////// Start getting Variables from the form //////////////////////
    ///////////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV/////////////
    var sp = PropertiesService.getScriptProperties();
    var sheet = SpreadsheetApp.getActiveSheet()
    var ss = SpreadsheetApp.openById(FORM_DATA); // this selects the spreadsheet by its ID#

    var s = SpreadsheetApp.getActiveSheet();
    var headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0]; // gets all the content from the form

    var ORDER_NUMBER = sheet.getRange("A1").getValue(); // gets the next available order number
    e.range.offset(0, -1, 1, 1).setValue(ORDER_NUMBER); // sets the first cell to the order number

    var TIME_STAMP = e.namedValues[headers[1]].toString(); // gets the DATE_CREATED
    var REP_CREATED = e.namedValues[headers[2]].toString(); // gets the REP_CREATED
    var COMPANY_NAME = e.namedValues[headers[3]].toString(); // gets the company name
    var BILLING_ADDRESS = e.namedValues[headers[4]].toString(); // gets the BILLING_ADDRESS
    var NEW_CUSTOMER = e.namedValues[headers[5]].toString(); // gets the NEW_CUSTOMER
    var POINT_OF_CONTACT = e.namedValues[headers[6]].toString(); // gets the POINT_OF_CONTACT
    var ATTN = e.namedValues[headers[7]].toString();
    var PO = e.namedValues[headers[8]].toString();
    var RGA = e.namedValues[headers[9]].toString();
    var SHIP_IMPORTANCE = e.namedValues[headers[10]].toString();
    var SHIPPING_ADDRESS = e.namedValues[headers[11]].toString();
    var SHIPPING_METHOD = e.namedValues[headers[12]].toString();
    var SHIPPING_CHARGE = e.namedValues[headers[13]].toString();
    var BLIND_SHIPPING = e.namedValues[headers[14]].toString();
    var CUSTOMER_SHIPPING_ACCOUNT = e.namedValues[headers[15]].toString();
    var SHIPPING_TRACKING_EMAIL = e.namedValues[headers[16]].toString();
    var PUMP_1_ORDER_TYPE = e.namedValues[headers[17]].toString();
    var QUANTITY_OF_PUMP_1 = Number(e.namedValues[headers[18]].toString());
    var PUMP_1_MODEL = e.namedValues[headers[19]].toString();
    var PRICE_FOR_PUMP_1 = e.namedValues[headers[20]].toString();
    var NOTES_FOR_PUMP_1 = e.namedValues[headers[21]].toString();
    var ORDER_TYPE_FOR_ACCESSORY_1 = e.namedValues[headers[22]].toString();
    var QUANTITY_FOR_ACCESSORY_1 = Number(e.namedValues[headers[23]].toString());
    var ACCESSORY_1 = e.namedValues[headers[24]].toString();
    var PRICE_FOR_ACCESSORY_1 = e.namedValues[headers[25]].toString();
    var ADD_2ND_PUMP_ACCESSORY = e.namedValues[headers[26]].toString();
    var PUMP_2_ORDER_TYPE = e.namedValues[headers[27]].toString();
    var QUANTITY_OF_PUMP_2 = Number(e.namedValues[headers[28]].toString());
    var PUMP_2_MODEL = e.namedValues[headers[29]].toString();
    var PRICE_FOR_PUMP_2 = e.namedValues[headers[30]].toString();
    var NOTES_FOR_PUMP_2 = e.namedValues[headers[31]].toString();
    var ORDER_TYPE_FOR_ACCESSORY_2 = e.namedValues[headers[32]].toString();
    var QUANTITY_FOR_ACCESSORY_2 = Number(e.namedValues[headers[33]].toString());
    var ACCESSORY_2 = e.namedValues[headers[34]].toString();
    var PRICE_FOR_ACCESSORY_2 = e.namedValues[headers[35]].toString();
    var ADD_3RD_PUMP_ACCESSORY = e.namedValues[headers[36]].toString();
    var PUMP_3_ORDER_TYPE = e.namedValues[headers[37]].toString();
    var QUANTITY_OF_PUMP_3 = Number(e.namedValues[headers[38]].toString());
    var PUMP_3_MODEL = e.namedValues[headers[39]].toString();
    var PRICE_FOR_PUMP_3 = e.namedValues[headers[40]].toString();
    var NOTES_FOR_PUMP_3 = e.namedValues[headers[41]].toString();
    var ORDER_TYPE_FOR_ACCESSORY_3 = e.namedValues[headers[42]].toString();
    var QUANTITY_FOR_ACCESSORY_3 = Number(e.namedValues[headers[43]].toString());
    var ACCESSORY_3 = e.namedValues[headers[44]].toString();
    var PRICE_FOR_ACCESSORY_3 = e.namedValues[headers[45]].toString();
    var ADD_4TH_PUMP_ACCESSORY = e.namedValues[headers[46]].toString();
    var PUMP_4_ORDER_TYPE = e.namedValues[headers[47]].toString();
    var QUANTITY_OF_PUMP_4 = Number(e.namedValues[headers[48]].toString());
    var PUMP_4_MODEL = e.namedValues[headers[49]].toString();
    var PRICE_FOR_PUMP_4 = e.namedValues[headers[50]].toString();
    var NOTES_FOR_PUMP_4 = e.namedValues[headers[51]].toString();
    var ORDER_TYPE_FOR_ACCESSORY_4 = e.namedValues[headers[52]].toString();
    var QUANTITY_FOR_ACCESSORY_4 = Number(e.namedValues[headers[53]].toString());
    var ACCESSORY_4 = e.namedValues[headers[54]].toString();
    var PRICE_FOR_ACCESSORY_4 = e.namedValues[headers[55]].toString();
    var ADD_5TH_PUMP_ACCESSORY = e.namedValues[headers[56]].toString();
    var PUMP_5_ORDER_TYPE = e.namedValues[headers[57]].toString();
    var QUANTITY_OF_PUMP_5 = Number(e.namedValues[headers[58]].toString());
    var PUMP_5_MODEL = e.namedValues[headers[59]].toString();
    var PRICE_FOR_PUMP_5 = e.namedValues[headers[60]].toString();
    var NOTES_FOR_PUMP_5 = e.namedValues[headers[61]].toString();
    var ORDER_TYPE_FOR_ACCESSORY_5 = e.namedValues[headers[62]].toString();
    var QUANTITY_FOR_ACCESSORY_5 = Number(e.namedValues[headers[63]].toString());
    var ACCESSORY_5 = e.namedValues[headers[64]].toString();
    var PRICE_FOR_ACCESSORY_5 = e.namedValues[headers[65]].toString();
    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^///////  
    ///////// End getting Variables from the form //////////////////////
    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^/////// 	  


    ///////////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV/////////////
    ///// Start to Find out how many pumps/accessories got ordered ///////
    ///////////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV/////////////
    var requiredRows = 1; // holds the number of how many rows will be inserted (1 per device/accessory pair)
    var add_more = 26; // add more pumps in col 26
    // This loops through all 15 possible devices returns the number of how many are 
    for (var l = 0; l < 4; l++) // there is a limit of 5 pumps/accessories so we loop and check all, the first one is assumed
    {
        if (e.namedValues[headers[add_more + l * 11]].toString().indexOf("Yes") > -1) // checks every 11 rows for more pumps/accessories
        {
            requiredRows++;
        }
    }
    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^///////
    ///// End to Find out how many pumps/accessories got ordered ///////
    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^///////  



    ///////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV///////
    //////// Start Create and fill a new copy of the template ////////////
    //////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV////////
    var spreadsheetTemplate = SpreadsheetApp.openById(SOURCE_TEMPLATE); // gets the template for the order sheet
    var newSS = spreadsheetTemplate.copy("Order " + ORDER_NUMBER + " for " + COMPANY_NAME); // creates a title for the new sheet

    // This sets the vales from the form
    // Row 1
    newSS.getActiveSheet().getRange('B1').setValue(ORDER_NUMBER);
    newSS.getActiveSheet().getRange('F1').setValue(REP_CREATED);
    // Row 2
    // Row 3
    newSS.getActiveSheet().getRange('A3').setValue(COMPANY_NAME);
    newSS.getActiveSheet().getRange('B3').setValue(BILLING_ADDRESS);
    newSS.getActiveSheet().getRange('C3').setValue(NEW_CUSTOMER);
    newSS.getActiveSheet().getRange('D3').setValue(POINT_OF_CONTACT);
    newSS.getActiveSheet().getRange('E3').setValue(PO);
    newSS.getActiveSheet().getRange('F3').setValue(RGA);
    // Row 4
    // Row 5
    newSS.getActiveSheet().getRange('B5').setValue(ATTN);
    newSS.getActiveSheet().getRange('F5').setValue(TIME_STAMP);
    // Row 6
    // Row 7
    newSS.getActiveSheet().getRange('A7').setValue(SHIP_IMPORTANCE);
    newSS.getActiveSheet().getRange('B7').setValue(SHIPPING_ADDRESS);
    newSS.getActiveSheet().getRange('C7').setValue(SHIPPING_METHOD);
    newSS.getActiveSheet().getRange('D7').setValue(SHIPPING_CHARGE);
    newSS.getActiveSheet().getRange('E7').setValue(BLIND_SHIPPING);
    newSS.getActiveSheet().getRange('F7').setValue(CUSTOMER_SHIPPING_ACCOUNT);
    // Row 8
    newSS.getActiveSheet().getRange('B8').setValue(SHIPPING_TRACKING_EMAIL);
    // Row 9
    // Row 10
    // Row 11
    // Row 12 --- pump 1 ----
    newSS.getActiveSheet().getRange('A12').setValue(PUMP_1_ORDER_TYPE);
    newSS.getActiveSheet().getRange('B12').setValue(QUANTITY_OF_PUMP_1);
    newSS.getActiveSheet().getRange('C12').setValue(PUMP_1_MODEL);
    newSS.getActiveSheet().getRange('F12').setValue(PRICE_FOR_PUMP_1);
    // Row 13
    newSS.getActiveSheet().getRange('B13').setValue(NOTES_FOR_PUMP_1);
    // Row 14 --- pump 2 ----
    newSS.getActiveSheet().getRange('A14').setValue(PUMP_2_ORDER_TYPE);
    newSS.getActiveSheet().getRange('B14').setValue(QUANTITY_OF_PUMP_2);
    newSS.getActiveSheet().getRange('C14').setValue(PUMP_2_MODEL);
    newSS.getActiveSheet().getRange('F14').setValue(PRICE_FOR_PUMP_2);
    // Row 15
    newSS.getActiveSheet().getRange('B15').setValue(NOTES_FOR_PUMP_2);
    // Row 16 --- pump 3 ----
    newSS.getActiveSheet().getRange('A16').setValue(PUMP_3_ORDER_TYPE);
    newSS.getActiveSheet().getRange('B16').setValue(QUANTITY_OF_PUMP_3);
    newSS.getActiveSheet().getRange('C16').setValue(PUMP_3_MODEL);
    newSS.getActiveSheet().getRange('F16').setValue(PRICE_FOR_PUMP_3);
    // Row 17
    newSS.getActiveSheet().getRange('B17').setValue(NOTES_FOR_PUMP_3);
    // Row 18 --- pump 4 ----
    newSS.getActiveSheet().getRange('A18').setValue(PUMP_4_ORDER_TYPE);
    newSS.getActiveSheet().getRange('B18').setValue(QUANTITY_OF_PUMP_4);
    newSS.getActiveSheet().getRange('C18').setValue(PUMP_4_MODEL);
    newSS.getActiveSheet().getRange('F18').setValue(PRICE_FOR_PUMP_4);
    // Row 19
    newSS.getActiveSheet().getRange('B19').setValue(NOTES_FOR_PUMP_4);
    // Row 20 --- pump 5 ----
    newSS.getActiveSheet().getRange('A20').setValue(PUMP_5_ORDER_TYPE);
    newSS.getActiveSheet().getRange('B20').setValue(QUANTITY_OF_PUMP_5);
    newSS.getActiveSheet().getRange('C20').setValue(PUMP_5_MODEL);
    newSS.getActiveSheet().getRange('F20').setValue(PRICE_FOR_PUMP_5);
    // Row 21
    newSS.getActiveSheet().getRange('B21').setValue(NOTES_FOR_PUMP_5);
    // Row 22
    // Row 23
    // Row 24
    // Row 25 --- accessory 1 ----
    newSS.getActiveSheet().getRange('A25').setValue(ORDER_TYPE_FOR_ACCESSORY_1);
    newSS.getActiveSheet().getRange('B25').setValue(QUANTITY_FOR_ACCESSORY_1);
    newSS.getActiveSheet().getRange('C25').setValue(ACCESSORY_1);
    newSS.getActiveSheet().getRange('F25').setValue(PRICE_FOR_ACCESSORY_1);
    // Row 26
    newSS.getActiveSheet().getRange('B26').setValue(NOTES_FOR_PUMP_1);
    // Row 27 --- accessory 2 ----
    newSS.getActiveSheet().getRange('A27').setValue(ORDER_TYPE_FOR_ACCESSORY_2);
    newSS.getActiveSheet().getRange('B27').setValue(QUANTITY_FOR_ACCESSORY_2);
    newSS.getActiveSheet().getRange('C27').setValue(ACCESSORY_2);
    newSS.getActiveSheet().getRange('F27').setValue(PRICE_FOR_ACCESSORY_2);
    // Row 28
    newSS.getActiveSheet().getRange('B28').setValue(NOTES_FOR_PUMP_2);
    // Row 29 --- accessory 3 ----
    newSS.getActiveSheet().getRange('A29').setValue(ORDER_TYPE_FOR_ACCESSORY_3);
    newSS.getActiveSheet().getRange('B29').setValue(QUANTITY_FOR_ACCESSORY_3);
    newSS.getActiveSheet().getRange('C29').setValue(ACCESSORY_3);
    newSS.getActiveSheet().getRange('F29').setValue(PRICE_FOR_ACCESSORY_3);
    // Row 30
    newSS.getActiveSheet().getRange('B30').setValue(NOTES_FOR_PUMP_3);
    // Row 31 --- accessory 3 ----
    newSS.getActiveSheet().getRange('A31').setValue(ORDER_TYPE_FOR_ACCESSORY_4);
    newSS.getActiveSheet().getRange('B31').setValue(QUANTITY_FOR_ACCESSORY_4);
    newSS.getActiveSheet().getRange('C31').setValue(ACCESSORY_4);
    newSS.getActiveSheet().getRange('F31').setValue(PRICE_FOR_ACCESSORY_4);
    // Row 32
    newSS.getActiveSheet().getRange('B32').setValue(NOTES_FOR_PUMP_4)
    // Row 33 --- accessory 4 ----
    newSS.getActiveSheet().getRange('A33').setValue(ORDER_TYPE_FOR_ACCESSORY_5);
    newSS.getActiveSheet().getRange('B33').setValue(QUANTITY_FOR_ACCESSORY_5);
    newSS.getActiveSheet().getRange('C33').setValue(ACCESSORY_5);
    newSS.getActiveSheet().getRange('F33').setValue(PRICE_FOR_ACCESSORY_5);
    // Row 34
    newSS.getActiveSheet().getRange('B34').setValue(NOTES_FOR_PUMP_5);


    // Move to (jonathan's) My Drive/Sales Sheets/New Orders
    var originalFolder = DriveApp.getFileById(SOURCE_TEMPLATE).getParents().next();
    var newSSFile = DriveApp.getFileById(newSS.getId());
    originalFolder.addFile(newSSFile); // adds the new sheet to the folder
    var SALES_SHEET_URL = newSSFile.getUrl();

    DriveApp.getRootFolder().removeFile(newSSFile); // removes the old temp sheet


    ////////////////////////////////////////////////////////////////
    //////// Fills out page 2 of the order form/////////////////////
    ////////////////////////////////////////////////////////////////

    var biotech_info = newSS.getSheetByName('Biotech Info'); //replace with source Sheet tab name
    var quantityStart = 8; // forst row that holds the quantity number
    var mostItems = 0; // this is to get the largest number of rows for each order. can not find a max function so this crap is used
    if (mostItems < QUANTITY_OF_PUMP_1) {
        mostItems = QUANTITY_OF_PUMP_1;
    }
    if (mostItems < QUANTITY_OF_PUMP_2) {
        mostItems = QUANTITY_OF_PUMP_2;
    }
    if (mostItems < QUANTITY_OF_PUMP_3) {
        mostItems = QUANTITY_OF_PUMP_3;
    }

    if (mostItems < QUANTITY_OF_PUMP_4) {
        mostItems = QUANTITY_OF_PUMP_4;
    }

    if (mostItems < QUANTITY_OF_PUMP_5) {
        mostItems = QUANTITY_OF_PUMP_5;
    }
    if (mostItems < QUANTITY_FOR_ACCESSORY_1) {
        mostItems = QUANTITY_FOR_ACCESSORY_1;
    }
    if (mostItems < QUANTITY_FOR_ACCESSORY_2) {
        mostItems = QUANTITY_FOR_ACCESSORY_2;
    }
    if (mostItems < QUANTITY_FOR_ACCESSORY_3) {
        mostItems = QUANTITY_FOR_ACCESSORY_3;
    }
    if (mostItems < QUANTITY_FOR_ACCESSORY_4) {
        mostItems = QUANTITY_FOR_ACCESSORY_4;
    }
    if (mostItems < QUANTITY_FOR_ACCESSORY_5) {
        mostItems = QUANTITY_FOR_ACCESSORY_5;
    }
testWait();



    biotech_info.getRange('B1').setValue(COMPANY_NAME);
    biotech_info.getRange('B2').setValue(ORDER_NUMBER);
    biotech_info.getRange('B3').setValue(TIME_STAMP);
    // biotech_info.getRange('F1').setValue(requiredRows);

    // Batch 1
    biotech_info.getRange('A5').setValue("(" + QUANTITY_OF_PUMP_1 + ") " + PUMP_1_MODEL + " - " + PUMP_1_ORDER_TYPE); // pump quantity
    // biotech_info.getRange('B7').setValue(PUMP_1_MODEL); // pump model
    biotech_info.getRange('D7').setValue(QUANTITY_FOR_ACCESSORY_1); // accessory quantity
    biotech_info.getRange('E7').setValue(ACCESSORY_1); // accessory model


    for (var rowNum = 1; rowNum < 161; rowNum++) {
        if (rowNum <= QUANTITY_OF_PUMP_1) {
            biotech_info.getRange(quantityStart + rowNum - 1, 1).setValue(rowNum);
        } else {
            biotech_info.getRange(quantityStart + rowNum - 1, 1).setValue("-");
        }
    }


    // batch 2
    biotech_info.getRange('G5').setValue("(" + QUANTITY_OF_PUMP_2 + ") " + PUMP_2_MODEL + " - " + PUMP_2_ORDER_TYPE); // pump quantity
    //  biotech_info.getRange('I7').setValue(PUMP_2_MODEL); // pump model
    biotech_info.getRange('J7').setValue(QUANTITY_FOR_ACCESSORY_2); // accessory quantity
    biotech_info.getRange('K7').setValue(ACCESSORY_2); // accessory model

    for (var rowNum = 1; rowNum < 161; rowNum++) {
        if (rowNum <= QUANTITY_OF_PUMP_2) {
            biotech_info.getRange(quantityStart + rowNum - 1, 7).setValue(rowNum);
        } else {
            biotech_info.getRange(quantityStart + rowNum - 1, 7).setValue("-");
        }
    }

    // batch 3
    biotech_info.getRange('M5').setValue("(" + QUANTITY_OF_PUMP_3 + ") " + PUMP_3_MODEL + " - " + PUMP_3_ORDER_TYPE); // pump quantity
    //   biotech_info.getRange('P7').setValue(PUMP_3_MODEL); // pump model
    biotech_info.getRange('P7').setValue(QUANTITY_FOR_ACCESSORY_3); // accessory quantity
    biotech_info.getRange('Q7').setValue(ACCESSORY_3); // accessory model

    for (var rowNum = 1; rowNum < 161; rowNum++) {
        if (rowNum <= QUANTITY_OF_PUMP_3) {
            biotech_info.getRange(quantityStart + rowNum - 1, 13).setValue(rowNum);
        } else {
            biotech_info.getRange(quantityStart + rowNum - 1, 13).setValue("-");
        }
    }


    // batch 4
    biotech_info.getRange('S5').setValue("(" + QUANTITY_OF_PUMP_4 + ") " + PUMP_4_MODEL + " - " + PUMP_4_ORDER_TYPE); // pump quantity
    //    biotech_info.getRange('W7').setValue(PUMP_4_MODEL); // pump model
    biotech_info.getRange('V7').setValue(QUANTITY_FOR_ACCESSORY_4); // accessory quantity
    biotech_info.getRange('W7').setValue(ACCESSORY_4); // accessory model

    for (var rowNum = 1; rowNum < 161; rowNum++) {
        if (rowNum <= QUANTITY_OF_PUMP_4) {
            biotech_info.getRange(quantityStart + rowNum - 1, 19).setValue(rowNum);
        } else {
            biotech_info.getRange(quantityStart + rowNum - 1, 19).setValue("-");
        }
    }

    // batch 5
    biotech_info.getRange('Y7').setValue("(" + QUANTITY_OF_PUMP_5 + ") " + PUMP_5_MODEL + " - " + PUMP_5_ORDER_TYPE); // pump quantity
    //  biotech_info.getRange('AD7').setValue(PUMP_5_MODEL); // pump model
    biotech_info.getRange('AB7').setValue(QUANTITY_FOR_ACCESSORY_5); // accessory quantity
    biotech_info.getRange('AC7').setValue(ACCESSORY_5); // accessory model

    for (var rowNum = 1; rowNum < 161; rowNum++) {
        if (rowNum <= QUANTITY_OF_PUMP_5) {
            biotech_info.getRange(quantityStart + rowNum - 1, 25).setValue(rowNum);
        } else {
            biotech_info.getRange(quantityStart + rowNum - 1, 25).setValue("-");
        }
    }

testWait();
	var rowsToHide = parseInt(mostItems)+ parseInt(quantityStart); // adds the start row and the max accessories b/c it adds them as a string otherwise.
    // hides all the extra rows
    biotech_info.hideRows(rowsToHide, 200);




    // Hides all the col that are not being used
    if (requiredRows == 1) {
        biotech_info.hideColumns(7, 25);
    } else if (requiredRows == 2) {
        biotech_info.hideColumns(13, 19);
    } else if (requiredRows == 3) {
        biotech_info.hideColumns(19, 13);
    } else if (requiredRows == 4) {
        biotech_info.hideColumns(25, 7);
    }

    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^///////
    ////////////// End Create a new copy of the template ///////////////
    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^///////  



    ///////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV///////
    //////// Start Copies all relevent info to the Active Orders tab /////
    //////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV////////

    var tab = ss.getSheetByName('Active Orders'); //replace with source Sheet tab name
    var lastRow = tab.getLastRow() + 1; // returns integer — the last row of the sheet that contains content
    var firstOrder = true; // variable to hold if the loop is the first time


    // This is inserting the amount of rows needed, if any. Using requiredRows value. After the last row of data.
    if (requiredRows > 0) tab.insertRowsAfter(lastRow, requiredRows);

    var blackOutRow = tab.getLastRow() + 1; // returns the next row to write date to
    // loops through and fills the line between RGA with a black empty fill  
    for (var lines2fill = 1; lines2fill <= 14; lines2fill++) {
        tab.getRange(blackOutRow, lines2fill).setBackgroundRGB(216, 216, 216); // fills the first 10 cells with black fill
    }

    var afterLastRow = tab.getLastRow() + 2; // returns the next row to write date to add 1 for a blank line, add another 1 to skip over the blacked out line
    // loops through each available line of pumps/accessories ordered to add to the Biotech tab
    for (var x = 0; x < requiredRows; x++) {
        // checks if the accessory order type is the same as the pump order type
        if (e.namedValues[headers[17 + (x * 11)]].toString() == e.namedValues[headers[22 + (x * 11)]].toString()) { // its the same, so no change
            var ACCESSORY = e.namedValues[headers[24 + (x * 11)]].toString();
        } else { // they are not the same, adding order type to accessory label ie. (rental) pole clamp
            var ACCESSORY = "(" + e.namedValues[headers[22]].toString() + ") " + e.namedValues[headers[24 + (x * 11)]].toString();
        }
      var numLines = parseInt(requiredRows)+1;
        // creats an array of all the info to add to each line
        var newLine = [, ORDER_NUMBER,
            TIME_STAMP,
            SHIP_IMPORTANCE, // SHIP_IMPORTANCE
            COMPANY_NAME, // COMPANY_NAME
            e.namedValues[headers[17 + (x * 10)]].toString(), //ORDER TYPE for PUMP number X 
            e.namedValues[headers[18 + (x * 10)]].toString(), //QUANTITY_OF_PUMP X
            e.namedValues[headers[19 + (x * 10)]].toString(), //MODEL_OF_PUMP X
            BLIND_SHIPPING,
            e.namedValues[headers[23 + (x * 10)]].toString(), //accessory X quantity 
            ACCESSORY, "", "", PO + "/" + RGA, SALES_SHEET_URL, numLines
        ]; // accessory name with optional lable




        // loops through the array and writes each item to its respective cell 1 at a time    
        // if there are 7 columns that are user filled before the data, set lines2write + 7 
        for (var lines2write = 1; lines2write <= 15; lines2write++) {
            tab.getRange(afterLastRow, lines2write).setValue(newLine[lines2write]).setVerticalAlignment("middle").setHorizontalAlignment("center").setWrap(true); // inserts the data into the next line

        }


        if (x == 0) // checks if this is the first time it loops the data
        {
            // firstOrder = false; // first loop is done


            // Set the data validation for Tech cell to require 'TBD', 'BH','MC','RA','JH','ND','LC' for the last row and the 11th col (tech), with dropdown menu.
            var techCell = tab.getRange(afterLastRow, 11);
            var techRule = SpreadsheetApp.newDataValidation().requireValueInList(['BH', 'MC', 'RA', 'JH', 'ND', 'LC'], true).build();
            techCell.setDataValidation(techRule);

            // Set the data validation for Completed? cell to require 'No' ,'In progress' , 'Yes' sdfsdf  for the last row and the 11th col (tech), with dropdown menu.
            var completedCell = tab.getRange(afterLastRow, 12);
            completedCell.setBackground("RED"); // sets that cell to red
            completedCell.setValue('No');// sets default to no for the completed col
            var completedRule = SpreadsheetApp.newDataValidation().requireValueInList(['No', 'In progress', 'Completed'], true).build();
            completedCell.setDataValidation(completedRule);
          
          
          
          ////////////// This sets the color rules for the background fill on certain cells //////////////
          
          // This sets the order type color
          if ((PUMP_1_ORDER_TYPE.indexOf("Rental") > -1)||(PUMP_1_ORDER_TYPE.indexOf("RENTAL") > -1)||(PUMP_1_ORDER_TYPE.indexOf("Current") > -1)||(PUMP_1_ORDER_TYPE.indexOf("CURRENT") > -1))
          {
            var OrderTypeCell1 = tab.getRange(afterLastRow, 5);
            OrderTypeCell1.setBackgroundRGB(0,255,0); // sets that cell to red
          }
          else if ((PUMP_1_ORDER_TYPE.indexOf("Rush") > -1)||(PUMP_1_ORDER_TYPE.indexOf("Rush") > -1))
          {
            var OrderTypeCell1 = tab.getRange(afterLastRow, 5);
            OrderTypeCell1.setBackground("RED"); // sets that cell to red
          } 
          else if ((PUMP_1_ORDER_TYPE.indexOf("Today") > -1)||(PUMP_1_ORDER_TYPE.indexOf("TODAY") > -1)||(PUMP_1_ORDER_TYPE.indexOf("Same Day") > -1)||(PUMP_1_ORDER_TYPE.indexOf("SAME DAY") > -1))
          {
            var OrderTypeCell1 = tab.getRange(afterLastRow, 5);
            OrderTypeCell1.setBackground("YELLOW"); // sets that cell to red
          }
          else if ((PUMP_1_ORDER_TYPE.indexOf("Sale") > -1)||(PUMP_1_ORDER_TYPE.indexOf("SALE") > -1)||(PUMP_1_ORDER_TYPE.indexOf("PURCHASE") > -1)||(PUMP_1_ORDER_TYPE.indexOf("Purchase") > -1))
          {
            var OrderTypeCell1 = tab.getRange(afterLastRow, 5);
            OrderTypeCell1.setBackground("CYAN"); // sets that cell to red
          }
          else if ((PUMP_1_ORDER_TYPE.indexOf("NEXT FEW DAYS") > -1)||(PUMP_1_ORDER_TYPE.indexOf("Next Few Days") > -1))
          {
            var OrderTypeCell1 = tab.getRange(afterLastRow, 5);
            OrderTypeCell1.setBackground("MAGENTA"); // sets that cell to red
          }
          
          
          
          // This sets the priority type fill color
          
          if ((SHIP_IMPORTANCE.indexOf("Current") > -1)||(SHIP_IMPORTANCE.indexOf("CURRENT") > -1))
          {
            var priotityCell = tab.getRange(afterLastRow, 3);
            priotityCell.setBackgroundRGB(0,255,0); // sets that cell to red
          }
          else if ((SHIP_IMPORTANCE.indexOf("Rush") > -1)||(SHIP_IMPORTANCE.indexOf("Rush") > -1))
          {
            var priotityCell = tab.getRange(afterLastRow, 3);
            priotityCell.setBackground("RED"); // sets that cell to red
          } 
          else if ((SHIP_IMPORTANCE.indexOf("Today") > -1)||(SHIP_IMPORTANCE.indexOf("TODAY") > -1))
          {
            var priotityCell = tab.getRange(afterLastRow, 3);
            priotityCell.setBackground("YELLOW"); // sets that cell to red
          }
          else if ((SHIP_IMPORTANCE.indexOf("Same Day") > -1)||(SHIP_IMPORTANCE.indexOf("SAME DAY") > -1))
          {
            var priotityCell = tab.getRange(afterLastRow, 3);
            priotityCell.setBackground("ORANGE"); // sets that cell to red
          }
          else if ((SHIP_IMPORTANCE.indexOf("AS FINISHED") > -1)||(SHIP_IMPORTANCE.indexOf("As Finished") > -1))
          {
            var priotityCell = tab.getRange(afterLastRow, 3);
            priotityCell.setBackground("MAGENTA"); // sets that cell to red
          }
          
          
          //////////////////////////////// end of fill color ///////////////////////////

            afterLastRow = afterLastRow + 1; // adds an additional row 
        } else // if not the first row, grey out the tech and competed cells
        {
            tab.getRange(afterLastRow, 11).setBackgroundRGB(216, 216, 216); // fills the tech cell with grey
            tab.getRange(afterLastRow, 12).setBackgroundRGB(216, 216, 216); // fills the tech cell with grey
        }


    }




    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^///////
    //////// End Copies all relevent info to the Biotech tab ///////////
    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^/////// 


    ///////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV///////
    ////////////////////// Starts forming Email //////////////////////////
    //////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV////////

    // This is to send an email to the customer
    var subject = "New Order: " + ORDER_NUMBER + " for " + COMPANY_NAME;
    var message = "Adepto,\n\n";

    message += "You got a new order for " + COMPANY_NAME + "\n\n";
    message += "Here is the sales sheet for the order:\n";

    message += SALES_SHEET_URL;
    message += "\n\nIt has been added to the biomed queue on the TV.\n\n";
    message += "For now, you will wait until the biotech makes the order as done on the Active Orders sheet. \nAfterwards, the order will be moved to the Finished tab and you will get an email letting you know its done."
    message += "\n\n Thank you,\nYour Automated Ordering System"

    var email = "sales@adeptomed.com";
    MailApp.sendEmail(email, subject, message, {
        bcc: "jonathan@adeptomed.com",
        noReply: true
    }); // {bcc: "rgaform@adeptomed.com", noReply: true}


    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^///////
    //////// End Copies all relevent info to the Biotech tab ///////////
    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^/////// 



    ///////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV///////
    ////////////////////// Starts blinking new order /////////////////////
    //////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV////////



    for (var i = 0; i < 120; i++) {

        if (i % 2 == 0) // flash on
        {
            tab.getRange("A1").setBackground("RED"); // sets cell to red
            tab.getRange('A1').setValue("NEW ORDER: " + ORDER_NUMBER + "-ASSIGN TECH!"); // puts text as new order

            if (techCell.isBlank()) // stops the blinking if a tech is assigned and returns the status cell back to blank
            {
                techCell.setBackground("RED"); // sets cell to red
            } else {
                i = 1000; // basically exits the loop next round
                tab.getRange("A1").setBackground("WHITE"); // sets cell to white
                tab.getRange('A1').setValue("Link here Adeptomed.com/ActiveOrders"); // resets the cell text
            }
        } else // flash off
        {
            techCell.setBackground("WHITE"); // sets cell to white
            tab.getRange('A1').setValue("Link here Adeptomed.com/ActiveOrders"); // resets the cell text
            tab.getRange("A1").setBackground("WHITE"); // sets cell to red

        }
        SpreadsheetApp.flush();
        Utilities.sleep(500);


    }


    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^///////
    ///////////////////// End blinking new order ///////////////////////
    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^/////// 
}