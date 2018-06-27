function onEditCharge1MorePerMonth(event) {
    Logger.log("==== Starting onEditCharge function  =====");

    // assumes source data in sheet named Needed
    // target sheet of move to named Acquired
    // test column with yes/no is col 4 or D
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s = event.source.getActiveSheet();
    var r = event.source.getActiveRange();
    Logger.log("Row: " + r.getRow());
    Logger.log("Col: " + r.getColumn());
    Logger.log("Value: " + r.getValue());
    //  emailMe("row: "+r.getRow()+ "-col: "+r.getColumn(),"row: "+r.getRow()+ "-col: "+r.getColumn());
    //s.getName() !== "Index" && s.getName() !== "Template" && s.getName() !== "Master Template" &&
    if (s.getName() !== "Index" && s.getName() !== "Template" && s.getName() !== "Master Template" && r.getColumn() == 11 && r.getRow() == 1) {
        var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy")
        Logger.log("Date: " + date);
        if (r.getValue() == "Apply $1+/Mo") {
            SpreadsheetApp.getActiveSheet().getRange('K1').setBackground('#28754E');
        } else {
            SpreadsheetApp.getActiveSheet().getRange('K1').setBackground('#B1440E');
        }


        var editCell = SpreadsheetApp.getActiveSheet().getRange('L1'); // rental start
        editCell.setValue("Charge updated on " + date).setFontWeight("normal")
            .setBackground('white')
            .setFontSize(11)
            .setFontColor("black")
            .setWrap(true);

      var changeUser = event.user;
      var updatedValue = r.getValue();
      var lastValue = event.oldValue;
      SpreadsheetApp.getActiveSheet().getRange('K1').setNote("Last updated on ["+date+"] to ["+updatedValue+"] by ["+changeUser+"] --- /n The old value was set to ["+lastValue+"]"); // rental start


      oldValue

        // emailMe("updated date","updated date");
    }
    Logger.log("==== Finished onEditCharge function  =====");

}
