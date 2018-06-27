function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{name:"ToBottom", functionName:"ToBottom"}];
  sheet.addMenu("Scripts", entries);
   ToBottom();
};

function ToBottom() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mysheet = ss.getActiveSheet();

  var lastrow = mysheet.getLastRow();
     
  mysheet.setActiveCell(mysheet.getDataRange().offset(lastrow-1, 0, 1, 1));
}; 