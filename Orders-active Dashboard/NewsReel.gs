function updateNews() {
    var SALES_ORDER_FORM = "1oUFi-zQm-XfvVwvHyZu_cArykys5KEo5GoGzKmu99fA";
    Logger.log('========== Starting the "%s" function ==========', 'updateNews');
    var scriptProperties = PropertiesService.getScriptProperties();
    //   scriptProperties.setProperty('newsNumber', 1);
    var activeSheet = SpreadsheetApp.openById(SALES_ORDER_FORM);
    var news = activeSheet.getSheetByName('News'); //replace with source Sheet tab name
    var lastNewsRow = activeSheet.getLastRow();
    var firstNewsRow = 3;
    var lastNewsRow = Number(scriptProperties.getProperty('lastNewsRow'));
    var story = "";
    var fontColor;
    var backgroundColor;
    var newsRange;
    if (lastNewsRow == null || lastNewsRow < 3 || lastNewsRow == 0 || news.getRange(lastNewsRow, 1).isBlank()) {
        lastNewsRow = firstNewsRow;
    }
    // testWait();
    newsRange = news.getRange(lastNewsRow, 1);
    story = newsRange.getValue();
    Logger.log('Set story to "%s"', story);
    fontColor = newsRange.getFontColor();
    Logger.log('Set fontColor to "%s"', fontColor);
    backgroundColor = newsRange.getBackground();
    Logger.log('Set backgroundColor to "%s"', backgroundColor);
    scriptProperties.setProperty('lastNewsRow', lastNewsRow + 1);
    var emailMessage = "row = ["+lastNewsRow+"]\n\nstory = [" +story+"]\n\n fontColor =["+fontColor+"]\n\n backgroundColor =["+backgroundColor+"]\n";
    Logger.log('Story Properties "%s"', emailMessage);
    // testWait();
    // emailMe(emailMessage,"properties of the news");
    updateNewsOnActiveOrders(story, fontColor, backgroundColor);
    Logger.log('========== Done with the "%s" function ==========', 'updateNews');

}

function updateNewsOnActiveOrders(story, fontColor, backgroundColor) {
    if (fontColor == null) {
        fontColor = "black";
    }

    if (backgroundColor == null) {
        backgroundColor = "white";
    }

    ///////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV///////
    ////////////////////// Starts blinking new order /////////////////////
    //////VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV////////
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tab = ss.getSheetByName('Active Orders'); //replace with source Sheet tab name
    for (var i = 0; i < 5; i++) {

        if (i % 2 == 0) // flash on
        {
            tab.getRange("L1").setValue(story).setBackground(backgroundColor).setFontColor(fontColor); // sets cell to red
        } else // flash off
        {
            tab.getRange("L1").setValue(story).setBackground("white").setFontColor("black"); // sets cell to red
        }
      //  testWait();
    }
    tab.getRange("L1").setValue(story).setBackground(backgroundColor).setFontColor(fontColor); // sets cell to red


    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^///////
    ///////////////////// End blinking new order ///////////////////////
    //////^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^/////// 


}