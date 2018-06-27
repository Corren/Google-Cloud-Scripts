//holds processing of next script till last one has completed
function testWait() {
    var lock = LockService.getScriptLock();
    lock.waitLock(500);
    SpreadsheetApp.flush();
    lock.releaseLock();
    // Utilities.sleep(200);
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