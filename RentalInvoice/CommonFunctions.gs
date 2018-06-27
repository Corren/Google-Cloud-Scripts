function ToBottom() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mysheet = ss.getActiveSheet();
    var lastrow = mysheet.getLastRow();

    mysheet.setActiveCell(mysheet.getDataRange().offset(lastrow - 1, 0, 1, 1));
};


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


// This returns the number of days between 2 dates inclusive
function daysBetween(date1, date2) {
    date1 = new Date((date1));
    date2 = new Date((date2));
    var diff = (date2 - date1) / (1000 * 60 * 60 * 24);
    return (diff + 1);
}



//holds processing of next script till last one has completed
function testWait() {
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    SpreadsheetApp.flush();
    lock.releaseLock();
    // Utilities.sleep(200);
}

