// known issues:
// the notes section for the accessories on the sales sheet duplicates the notes for the pump.
// this is because i do not have a notes section on the oprder form.


function myFunction(e) {

    // Google Doc id from the document template
    // (Get ids from the URL)
    //var SOURCE_TEMPLATE = "1lP_vNLwlp2uEDg_VORWNBVrKiYbe__tgFrdN1UUkkqM"; // offical template
    var SOURCE_TEMPLATE = "1SwdmVWvZqQaS7Pp7C23BvH4Wr5Aot4MO--nSEMFHM5w"; // test template

    var FORM_DATA = "15ivw-rVW8U_3nBnQs6UOp5qIM9_MsFsIDifI_jqeVm4";


    ///////////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV/////////////
    /////////////////// Start getting Variables from the form //////////////////////
    ///////////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV/////////////
    var sp = PropertiesService.getScriptProperties();
    var sheet = SpreadsheetApp.getActiveSheet()
    var ss = SpreadsheetApp.openById(FORM_DATA); // this selects the spreadsheet by its ID#

    var s = SpreadsheetApp.getActiveSheet();
    var headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0]; // gets all the content from the form

    var BED_PM_NUMBER = sheet.getRange("A1").getValue(); // gets the next available order number
    e.range.offset(0, -1, 1, 1).setValue(BED_PM_NUMBER); // sets the first cell to the order number

    var TIME_STAMP = e.namedValues[headers[1]].toString(); // gets the DATE_CREATED
    var HOSPITAL = e.namedValues[headers[2]].toString(); // gets the REP_CREATED
    var MODEL = e.namedValues[headers[3]].toString(); // gets the company name
    var PM_DUE_DATE = e.namedValues[headers[4]].toString(); // gets the BILLING_ADDRESS
    var CM_PM = e.namedValues[headers[5]].toString(); // gets the NEW_CUSTOMER
    var OLD_TAG = e.namedValues[headers[6]].toString(); // gets the POINT_OF_CONTACT
    var NEW_TAG = e.namedValues[headers[7]].toString();
    var SN = e.namedValues[headers[8]].toString();
    var RFID = e.namedValues[headers[9]].toString();
    var HRS_WORKED = e.namedValues[headers[10]].toString();
    var PROBLEM_REPORTED = e.namedValues[headers[11]].toString();
    var WORK_PERFORMED = e.namedValues[headers[12]].toString();
    var TECH = e.namedValues[headers[13]].toString();
    var Q1 = e.namedValues[headers[14]].toString();
    var Q2 = e.namedValues[headers[15]].toString();
    var Q3 = e.namedValues[headers[16]].toString();
    var Q4 = e.namedValues[headers[17]].toString();
    var Q5 = e.namedValues[headers[18]].toString();
    var Q6 = e.namedValues[headers[19]].toString();
    var Q7 = e.namedValues[headers[20]].toString();
    var Q8 = e.namedValues[headers[21]].toString();
    var Q9 = e.namedValues[headers[22]].toString();
    var Q10 = e.namedValues[headers[23]].toString();
    var Q11 = e.namedValues[headers[24]].toString();
  var PARTS_USED = e.namedValues[headers[25]].toString();

  try{
  Logger.log("BED_PM_NUMBER: "+BED_PM_NUMBER);
  Logger.log("TIME_STAMP: "+TIME_STAMP);
  Logger.log("HOSPITAL: "+HOSPITAL);
  Logger.log("MODEL: "+MODEL);
  Logger.log("PM_DUE_DATE: "+PM_DUE_DATE);
  Logger.log("CM_PM: "+CM_PM);
  Logger.log("OLD_TAG: "+OLD_TAG);
  Logger.log("NEW_TAG: "+NEW_TAG);
  Logger.log("SN: "+SN);
  Logger.log("RFID: "+RFID);
  Logger.log("HRS_WORKED: "+HRS_WORKED);
  Logger.log("PROBLEM_REPORTED: "+PROBLEM_REPORTED);
  Logger.log("WORK_PERFORMED: "+WORK_PERFORMED);
  Logger.log("TECH: "+TECH);
  Logger.log("Q1: "+Q1);
  Logger.log("Q2: "+Q2);
  Logger.log("Q3: "+Q3);
  Logger.log("Q4: "+Q4);
  Logger.log("Q5: "+Q5);
  Logger.log("Q6: "+Q6);
  Logger.log("Q7: "+Q7);
  Logger.log("Q8: "+Q8);
  Logger.log("Q9: "+Q9);
  Logger.log("Q10: "+Q10);
  Logger.log("Q11: "+Q11);
  Logger.log("PARTS_USED: "+PARTS_USED);

  
    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^///////  
    ///////// End getting Variables from the form //////////////////////
    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^/////// 	  



    ///////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV///////
    //////// Start Create and fill a new copy of the template ////////////
    //////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV////////
    var spreadsheetTemplate = SpreadsheetApp.openById(SOURCE_TEMPLATE); // gets the template for the order sheet
  var REPORT_NAME = HOSPITAL+"_Bed_PM_REF: " + BED_PM_NUMBER ;
    Logger.log("REPORT_NAME: "+REPORT_NAME);
  var newSS = spreadsheetTemplate.copy(REPORT_NAME); // creates a title for the new sheet
    // This sets the vales from the form
    // Row 2
     newSS.getActiveSheet().getRange('G2').setValue(REPORT_NAME);
      // Row 3
      newSS.getActiveSheet().getRange('G3').setValue(HOSPITAL);
      newSS.getActiveSheet().getRange('J3').setValue(OLD_TAG);
      newSS.getActiveSheet().getRange('L3').setValue(HRS_WORKED);
        // Row 4
      newSS.getActiveSheet().getRange('G4').setValue(MODEL);
      newSS.getActiveSheet().getRange('J4').setValue(NEW_TAG);
      newSS.getActiveSheet().getRange('L4').setValue(TECH);
        // Row 5
      newSS.getActiveSheet().getRange('G5').setValue(PM_DUE_DATE);
      newSS.getActiveSheet().getRange('J5').setValue(SN);
        // Row 6
      newSS.getActiveSheet().getRange('G6').setValue(CM_PM);
      newSS.getActiveSheet().getRange('J6').setValue(RFID);
  
  // row 10
  newSS.getActiveSheet().getRange('B10').setValue(Q1);
  newSS.getActiveSheet().getRange('B11').setValue(Q2);
  newSS.getActiveSheet().getRange('B12').setValue(Q3);
  newSS.getActiveSheet().getRange('B13').setValue(Q4);
  newSS.getActiveSheet().getRange('B14').setValue(Q5);
  newSS.getActiveSheet().getRange('B15').setValue(Q6);
  newSS.getActiveSheet().getRange('B16').setValue(Q7);
  newSS.getActiveSheet().getRange('B17').setValue(Q8);
  newSS.getActiveSheet().getRange('B18').setValue(Q9);
  newSS.getActiveSheet().getRange('B19').setValue(Q10);
  // 10a
  // 10b
  newSS.getActiveSheet().getRange('B22').setValue(Q11);
  
  // notes
  newSS.getActiveSheet().getRange('B25').setValue("Work Performed: "+WORK_PERFORMED);
  newSS.getActiveSheet().getRange('B26').setValue("Problems: " + PROBLEM_REPORTED);
  
  // parts used
  newSS.getActiveSheet().getRange('H25').setValue(PARTS_USED);




    // Move to (jonathan's) My Drive/Sales Sheets/New Orders
    var originalFolder = DriveApp.getFileById(SOURCE_TEMPLATE).getParents().next();
    var newSSFile = DriveApp.getFileById(newSS.getId());
    originalFolder.addFile(newSSFile); // adds the new sheet to the folder
    var SALES_SHEET_URL = newSSFile.getUrl();

 //  DriveApp.getRootFolder().removeFile(newSSFile); // removes the old temp sheet
   DriveApp.getRootFolder().removeFile(newSSFile); // removes the old temp sheet
  Logger.log("Removed copy of report from the root: "+REPORT_NAME);



    ///////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV///////
    ////////////////////// Starts forming Email //////////////////////////
    //////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV////////
  }catch(errorMakingForm){
    Logger.log("error making form: "+errorMakingForm);
  }finally{
   try{
  testWait();
                  var pdf = DriveApp.getFileById(newSSFile.getId()).getAs('application/pdf').getBytes();

                var attach = {
                    fileName: REPORT_NAME,
                    content: pdf,
                    mimeType: 'application/pdf'
                };

  
  
    // This is to send an email to the customer
    var subject = REPORT_NAME;
    var message = "";
  Logger.log(": "+"");
 
      Logger.log("Trying to email form");
   // var email = "77uph@adeptomed.com";
  var printer_email = "adeptoinvoices1@hpeprint.com,bedteam1@adeptomed.com,jonathan@adeptomed.com";
    MailApp.sendEmail(printer_email, subject, message, {
       // bcc: "jonathan@adeptomed.com",
        noReply: true,
      attachments: [attach]
    }); // {bcc: "rgaform@adeptomed.com", noReply: true}
  }catch(e){
    Logger.log("failed to email form");
    mailMe(e+Logger.getLog(),"error in mailing PM form");
  }
  }

    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^///////
    //////// End Copies all relevent info to the Biotech tab ///////////
    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^/////// 



}