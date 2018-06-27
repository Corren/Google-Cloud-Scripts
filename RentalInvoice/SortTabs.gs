function sortGoogleSheets() {
   Logger.log('========== Starting the "%s" function ==========', 'sortGoogleSheets');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabToSort = ss.getActiveSheet().getName();
 // var actualSheetName = SpreadsheetApp.getActiveSpreadsheet().getName();

   Logger.log('=========sorting %s tab===========',tabToSort);

  // Store all the worksheets in this array
  var sheetNameArray = [];
  var sheets = ss.getSheets();
  for (var i = 2; i < sheets.length; i++) {
    var name = sheets[i].getName();
//    if(name !== "Index" && name !== "Template" && name !== "Master Template"){ // exclude some names
   if(name !== tabToSort ){ // exclude some names
     Logger.log("adding "+name+" to list of tabs.");
    sheetNameArray.push(name);
   }
//    }
  }
  Logger.log('total item: '+sheets.length);
  for( var j = 0; j < sheets.length; j++ ) {
    Logger.log('compaing item '+j);
     Logger.log('========== comparing "%s" and '+sheetNameArray[j],tabToSort);
    if (tabToSort.localeCompare(sheetNameArray[j]) ==-1){
      Logger.log(sheetNameArray[j]+' then [[%s]] then ' + sheetNameArray[j+1], tabToSort);
       ss.moveActiveSheet(j +3);
        break;
    }else if (j == sheets.length - 1)
    {
        Logger.log(sheetNameArray[j]+' then %s  ', tabToSort);
       ss.moveActiveSheet(j + 3);
      break;
    }
  
    }
}


function testF()
{
 console.error("this is a console error from jg"); 
}