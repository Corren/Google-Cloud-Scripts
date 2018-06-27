// Run this code on a google sheet where col A is the sql model key and col D is the price to update
function getSQLcode()
{
  var activeSheet = SpreadsheetApp.getActiveSheet();
	var firstRowRental = 2;
	var lastRowRental = activeSheet.getLastRow();
	var message = "";
	var sub="SQL Code to update all prices";
	var pref="UPDATE \n  [ADEPTO].[dbo].[Equipment] \nset \n  [EstPurchaseCost] =1,\n  [PurchaseCost]=";
	var post ="\nwhere \n  ([DeptCharg2Key] = 220 \n  or [DeptCharg2Key] = 1322 \n  or [DeptCharg2Key] = 132 \n  or [DeptCharg2Key] = 2 \n  or [DeptCharg2Key] = 1)\n  and [PurchaseCost] = 0 \n  and [EquipmentStatKey] !=5\n  and [ModelKey] = ";

	for (var row = firstRowRental; row <= lastRowRental; row++) {
		var model = activeSheet.getRange(row, 1).getValue();
		var price = activeSheet.getRange(row, 4).getValue();
		message+=pref+price+post+model+"\n\n---------------------------------------------------\n\n\n";
	}
  emailMe(message, sub);
 }
 
 
// This just send me an email with the message
//    REQUIRED: string body - this is the message you want to send
function emailMe(body, sub) {
    if (sub == null) {
        var subject = "testing";
    } else {
        var subject = sub;
    }
    var message = body;
    var email = "jonathan@adeptomed.com";

    MailApp.sendEmail(email, subject, message)
}

 function onOpen() {
    var menuEntries = [{
            name: "Generate SQL code",
            functionName: "getSQLcode"
        }
    ];

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.addMenu("Jonathan's Magic", menuEntries);
}