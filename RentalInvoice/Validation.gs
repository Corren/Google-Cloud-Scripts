function requireValidationOnAll() {
    var ss = SpreadsheetApp.getActive();
    var allNames = "";
//    var colors = ["#2952A3", "#4E5D6C", "#0D7813", "#5229A3", "#BE6D00", "#705770", "#7A367A", "#AB8B00", "#28754E", "#88880E"];
//    var c = 0;
//    var color = "";
    for (var n in ss.getSheets()) {
        var sheet = ss.getSheets()[n]; // look at every sheet in spreadsheet
        var name = sheet.getName(); //get name
        if (name != 'Index') { // exclude some names
 //           allNames += addTabColor(name, colors[c]) + "\n";
            allNames += addException(name) + "\n";
    //        c++;
    //        if (c == 10)
   //             c = 0;
        }
     //   testWait();
      
    }
    emailMe("Fixed the following:\n" + allNames, "Done fixing requirments. Rentals and Accessories log 2018");
}

function addException(tabName) {
    Logger.log("==== Starting addException function for tab [%s] =====", tabName);
    var RENTAL_DATA = "1NKVar1I2BhSJd5LdAT5g37w8syj9H22k9BV73wu0hgw";
    var rentalSheet = SpreadsheetApp.openById('1NKVar1I2BhSJd5LdAT5g37w8syj9H22k9BV73wu0hgw'); // this gets you the rental log spreadsheet 
    if (tabName == null) {
        var activeSheet = SpreadsheetApp.getActiveSheet();
    } else {
        var activeSheet = SpreadsheetApp.openById('1NKVar1I2BhSJd5LdAT5g37w8syj9H22k9BV73wu0hgw').getSheetByName(tabName);
    }
  
  Logger.log("== Tab Name: [%s] =====", activeSheet.getName());

  Logger.log("sent [%s] to fixMerge(tabName)",fixMerge(tabName));

  try{
    var fixCell = activeSheet.getRange('J3');
    fixCell.setValue("Accessory #1").setFontWeight("bold")
        .setBackground('#a64d79')
        .setFontSize(11)
        .setFontColor("black")
        .setWrap(true).setDataValidation(null);
   var repCell = activeSheet.getRange('J1'); // rental start
  var repValue = repCell.getValue();
  Logger.log("J1: ["+ repValue+"]");
  repCell.setValue("Rep: House")
        .setFontWeight("normal")
        .setBackground('#568b22')
        .setFontSize(11)
        .setFontColor("black")
        .setWrap(true);
   var rep1 = SpreadsheetApp.newDataValidation().requireValueInList(
        ['Rep: House',
         'Rep: Tommy',
         'Rep: Clint',
         'Rep: Scott'
        ], true).setAllowInvalid(true).setHelpText('Select the account owner from the list.').build();
  
    Logger.log("Rep cell added");
      repCell.setDataValidation(rep1);
    Logger.log("Added the drop down for rep");
  }catch(e)
  {
    emailMe(e,tabName);
  }
var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy")
    Logger.log("Date: " + date);

    var surchargeCell = activeSheet.getRange('K1'); // rental start



    var apply = SpreadsheetApp.newDataValidation().requireValueInList(
        ['Apply $1+/Mo',
            'No Apply $1+/Mo'
        ], true).setAllowInvalid(false).setHelpText('Select if this customer is eligible for the $1 per month increase in rental rate.').build();
    surchargeCell.setValue("Apply $1+/Mo")
        .setFontWeight("normal")
        .setBackground('#28754E')
        .setFontSize(11)
        .setFontColor("black")
        .setWrap(true)
        .setNote("Updated charge to [Apply $1/Mo] on " + date);
    Logger.log("Added default text to charge cell");


    //  SpreadsheetApp.getActiveSheet.getRange('K1').setFontColor("black")
    surchargeCell.setDataValidation(apply);
    Logger.log("Added the drop down from charge/no charge");

    var editCell = SpreadsheetApp.getActiveSheet().getRange('L1'); // rental start


    
    editCell.setValue("Charge updated on " + date).setFontWeight("normal")
        .setBackground('white')
        .setFontSize(11)
        .setFontColor("black")
        .setWrap(true);
    Logger.log("Added updated charge on cell");
  


  

    activeSheet.getRange("K1:L1").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.DASHED);
    Logger.log("Added cell border");


    Logger.log("==== Finished addException function for tab [%s] =====", tabName);
    return tabName;
}

function fixMerge(tabName)
{
  try{
      Logger.log("==== Starting fixMerge function for tab [%s] =====", tabName);
    var RENTAL_DATA = "1NKVar1I2BhSJd5LdAT5g37w8syj9H22k9BV73wu0hgw";
    var rentalSheet = SpreadsheetApp.openById(RENTAL_DATA); // this gets you the rental log spreadsheet 
    if (tabName == null) {
        var activeSheet = SpreadsheetApp.getActiveSheet();
    } else {
        var activeSheet = rentalSheet.getSheetByName(tabName);
    }
  
    var rangeOld = activeSheet.getRange("A1:L1");
    var rangeNew = activeSheet.getRange("A1:I1");
    rangeOld.breakApart();
    Logger.log("Unmerged cells A1:L1");

    rangeNew.mergeAcross();
    Logger.log("Merged cells A1:J1");
  }
  catch(e)
  {
    return ("Merge error: "+e);
  }
 Logger.log("==== Ending fixMerge function for tab [%s] =====", tabName);
}






















function requireValidation(tabName) {
    var RENTAL_DATA = "1NKVar1I2BhSJd5LdAT5g37w8syj9H22k9BV73wu0hgw";
    var rentalSheet = SpreadsheetApp.openById(RENTAL_DATA); // this gets you the rental log spreadsheet 
    if (tabName == null) {
        var activeSheet = SpreadsheetApp.getActiveSheet();
    } else {
        var activeSheet = rentalSheet.getSheetByName(tabName);
    }

    // activeSheet.getRange("F2:H2").breakApart()
    activeSheet.getRange('F2').setValue("Email address:"); // adds the email label
    // The code below merges cells C5:E5 into one cell
    //activeSheet.getRange("G2:H2").mergeAcross();  

    var colC = activeSheet.getRange('C4:C'); // rental start
    var colD = activeSheet.getRange('D4:D'); // rental end

    var colE = activeSheet.getRange('E4:E'); // monthly rate

    var colF = activeSheet.getRange('F4:F'); // total acc 

    var colH = activeSheet.getRange('H4:H'); // acc 1 start
    var colI = activeSheet.getRange('I4:I'); // acc 1 end

    var colK = activeSheet.getRange('K4:K'); // acc 2 end
    var colL = activeSheet.getRange('L4:L'); // acc 2 start 

    var dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText('Must be a valid date.').build();
    var rateRule = SpreadsheetApp.newDataValidation().requireNumberBetween(30, 150).setAllowInvalid(false).setHelpText('Must be a number between 30 and 150. Only enter the number. Do not enter $ or other text.').build();
    var accRule = SpreadsheetApp.newDataValidation().requireNumberBetween(0, 5).setAllowInvalid(false).setHelpText('Must be a number between 0 and 5. Only enter the number and other text.').build();


    activeSheet.setFrozenRows(3);

    colC.setDataValidation(dateRule);
    colD.setDataValidation(dateRule);
    colE.setDataValidation(rateRule);
    colF.setDataValidation(accRule);
    colH.setDataValidation(dateRule);
    colI.setDataValidation(dateRule);
    colK.setDataValidation(dateRule);
    colL.setDataValidation(dateRule);

    return tabName;
}




function requireNumberFormat(tabName) {
    var RENTAL_DATA = "1NKVar1I2BhSJd5LdAT5g37w8syj9H22k9BV73wu0hgw";
    var rentalSheet = SpreadsheetApp.openById(RENTAL_DATA); // this gets you the rental log spreadsheet 
    if (tabName == null) {
        var activeSheet = SpreadsheetApp.getActiveSheet();
    } else {
        var activeSheet = rentalSheet.getSheetByName(tabName);
    }

    var colE = activeSheet.getRange('E4:E'); // monthly rate

    var colF = activeSheet.getRange('F4:F'); // total acc 




    colE.setNumberFormat("0.00");
    colF.setNumberFormat("0");


    return tabName;
}




function freezeTop3Rows(tabName) {
    var RENTAL_DATA = "1NKVar1I2BhSJd5LdAT5g37w8syj9H22k9BV73wu0hgw";
    var rentalSheet = SpreadsheetApp.openById(RENTAL_DATA); // this gets you the rental log spreadsheet 
    if (tabName == null) {
        var activeSheet = SpreadsheetApp.getActiveSheet();
    } else {
        var activeSheet = rentalSheet.getSheetByName(tabName);
    }

    activeSheet.setFrozenRows(3);

    return tabName;
}




function addTabColor(tabName, color) {
    var RENTAL_DATA = "1NKVar1I2BhSJd5LdAT5g37w8syj9H22k9BV73wu0hgw";
    var rentalSheet = SpreadsheetApp.openById(RENTAL_DATA); // this gets you the rental log spreadsheet 
    if (tabName == null) {
        var activeSheet = SpreadsheetApp.getActiveSheet();
    } else {
        var activeSheet = rentalSheet.getSheetByName(tabName);
    }
      if (color == null) {
        var color = "white";
    } 
 // emailMe("color: "+color+"-sheet name: "+tabName,"");
    activeSheet.setTabColor(color);

    return tabName;
}




function pumpDropdown(tabName) {
    var RENTAL_DATA = "1NKVar1I2BhSJd5LdAT5g37w8syj9H22k9BV73wu0hgw";
    var rentalSheet = SpreadsheetApp.openById(RENTAL_DATA); // this gets you the rental log spreadsheet 
    if (tabName == null) {
        var activeSheet = SpreadsheetApp.getActiveSheet();
    } else {
        var activeSheet = rentalSheet.getSheetByName(tabName);
    }

    var colA = activeSheet.getRange('A4:A'); // rental start



    var pumpList = SpreadsheetApp.newDataValidation().requireValueInList(
            ['Acclaim',
            'Alaris 8100',
            'Alaris 8015 PCU',
            'Alaris 8100',
            'Alaris 8110 Syringe',
            'Alaris 8120 PCA',
            'Alaris 8220 Sp02',
            'Alaris 8300 ETC02',
            'Alaris Medsystem III',
            'Cadd 5800',
            'Cadd One 5100',
            'Cadd PCA 5800',
            'Cadd Plus 5400',
            'Cadd TPN 5700',
            'Baxter 150 XL',
            'Baxter 300 XL',
            'Baxter 6201',
            'Baxter 6201 Custom',
            'Baxter 6301',
            'Baxter AS40A',
            'Baxter AS50',
            'Baxter I Pump',
            'Baxter InfusOR',
            'Baxter PCA II',
            'Baxter Sigma Spectrum',
            'Cadd Legacy One 6400',
            'Cadd Legacy PCA 6300',
            'Cadd Legacy Plus 6500',
            'Cadd Prizm PCS II',
            'Cadd Prizm PCS II (Epi)',
            'Cadd Prizm VIP 6100',
            'Cadd Prizm VIP 6101',
            'Cadd Solis PCA 2110',
            'Cadd Solis VIP 2120',
            'CME Bodyguard 323',
            'Crib',
            'Curlin 4000',
            'Curlin 4000CMS',
            'Curlin 6000',
            'Curlin Painsmart IOD',
            'Curlin Painsmart IOD Epi',
            'Custom Baxter 6201',
            'Custom Baxter 6301',
            'Gemstar',
            'Gemstar (Yellow-Epi)',
            'Graseby 3400',
            'Graseby MS 16A',
            'Graseby MS 26',
            'Infusomat Space',
            'Kangaroo 224',
            'Kangaroo 324',
            'Kangaroo E Pump',
            'Kangaroo Joey',
            'Kangaroo Pet',
            'Lifecare PCA',
            'Lifecare PCA III',
            'Medfusion 2001',
            'Medfusion 2010',
            'Medfusion 2010i',
            'Medfusion 3010A',
            'Medfusion 3500',
            'Medfusion 4000',
            'Outlook 100',
            'Outlook 200',
            'Outlook 400ES',
            'Perfusor Basic',
            'Perfusor Space',
            'Plum A+ (11.6)',
            'Plum A+ (13.41)',
            'Plum A+ Triple',
            'Plum XL',
            'Plum XL Triple',
            'RMS Freedom 60',
            'Sabratek 3030',
            'Sigma 8000',
            'Vista Basic',
            'Walkmed 350',
            'Walkmed 350VL',
            'Zevex Infinity '], false).setAllowInvalid(true).setHelpText('Select an item from the list or add your own.').build(); 
  colA.setDataValidation(pumpList);



    }