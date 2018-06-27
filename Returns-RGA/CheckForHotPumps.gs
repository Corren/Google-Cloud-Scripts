function tester()
{
  Logger.log("******* starting test ***********");
  var result = hotPumpCheck("CADD PRIZIM");
  Logger.log("Result: "+ result);
  Logger.log("******* ending test ***********");
}




function hotPumpCheck(model) {
  Logger.log("======== Starting hotPumpCheck Function ========");
    var sheet = SpreadsheetApp.openById("1dxt6c62NJrjiEOa_Ye5xAze14UQTgLBn26IunymDAQc").getSheetByName("Hot Pumps");
    var lastRow = sheet.getLastRow();

    var data = sheet.getRange(2, 1, lastRow, 1).getValues(); //getRange(starting Row, starting column, number of rows, number of columns)
  Logger.log("");

Logger.log(" ---- loading hot pump list ---- ");
    for(var i=0;i<(lastRow-1);i++)
    {//    .toLowerCase()
      Logger.log("Checking if ["+model.toLowerCase()+"] contains ["+(data[i][0]).toLowerCase()+"] - Hot Pump ["+Number(i+1)+" of "+Number(lastRow-1)+"]");

      if(model.toLowerCase().indexOf((data[i][0]).toLowerCase())>-1)
      {
        Logger.log(" >>>> MATCH - PUMP IS HOT!!!<<<<");
        Logger.log("======== Ending hotPumpCheck Function ========");
        Logger.log("");
        return true;
      }
    }
  Logger.log(" >>>> No match - Pump is not hot <<<<");
  Logger.log("======== Ending hotPumpCheck Function ========");
   Logger.log("");
  return false;
}
