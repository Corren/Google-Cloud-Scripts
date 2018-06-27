function processAllInvoices(invoiceMonth, printPDF, emailReport) {
    try {
      var startTime = (new Date()).getTime();
        Logger.log('========== Starting the "%s" function ==========', 'processAllInvoices');

        var scriptProperties = PropertiesService.getScriptProperties();

        var tempMonth = scriptProperties.getProperty('month');
        if (invoiceMonth == "Last" || tempMonth == "Last") {
            var month = "Last";
            scriptProperties.setProperty('month', "Last");
        } else //if (invoiceMonth == "Current" || tempMonth == "Current") 
        {
            var month = "Current";
            scriptProperties.setProperty('month', "Current");
        }
        Logger.log('Invoice month set to "%s"', month);
     
        if (emailReport == true || Number(scriptProperties.getProperty('emailReport')) == 1) {
            var sendEmailReport = true;
            scriptProperties.setProperty('emailReport', 1);
        } else {
            var sendEmailReport = false;
        }
        Logger.log('Email report set to "%s"', sendEmailReport);

        if (printPDF == true || Number(scriptProperties.getProperty('printPDF')) == 1) {
            var print = true; // should the item be printed to the HP printer automatically
            scriptProperties.setProperty('printPDF', 1);
        } else {
            var print = false; // should the item be printed to the HP printer automatically
        }
        Logger.log('Print report set to "%s"', print);
        
        var ss = SpreadsheetApp.getActive();
        var toMS = 6 * 10000; // 1 minute = 6 × 10000 milliseconds = 60000 milliseconds
        var invoiceSummary = "";
        var tabNames = []; //

        Logger.log('Starting to get all tab names');
        for (var n in ss.getSheets()) {
            var sheet = ss.getSheets()[n]; // look at every sheet in spreadsheet
            var name = sheet.getName(); //get name
        //    if(name != "Index" && name != "Template" && name != "Master Template")
        //    {
                  tabNames.push(name);
        //    }
          
        }
        Logger.log('Finished getting all tab names');

        var firstTab = 5;
        var totalTabs = Number(scriptProperties.getProperty('totalTabs'));
        var lastTab = Number(scriptProperties.getProperty('lastTab'));
        var invoiceSummaryinvoiceTotal = scriptProperties.getProperty('invoiceSummary');
        if (totalTabs == null || totalTabs == 0) {
            totalTabs = Number(ss.getSheets().length);
            scriptProperties.setProperty('totalTabs', totalTabs);
            Logger.log('Set the totalTabs property to "%s"', totalTabs);
        }        

        if (lastTab == null || lastTab == 0) {
            lastTab = firstTab;
            scriptProperties.setProperty('lastTab', lastTab);
            Logger.log('Set the lastTab properties to "%s"', lastTab);
        }
        if (invoiceSummary == null) {
            invoiceSummary = "";
        }

        Logger.log('Starting to process all the tabs.');
        for (var currTab = lastTab; currTab <= totalTabs; currTab++) {
            Logger.log('Processing all tabs, current tab is %s',currTab);
            var currTime = (new Date()).getTime();
            var timeDiff = currTime - startTime;
            if (timeDiff >= (5 * toMS)) {
                Logger.log('The run time for this script has passed the 5 miniute limit at %s miniutes.', (timeDiff/toMS));
                scriptProperties.setProperty('lastTab', currTab);
                Logger.log('Set the lastTab property to %s in prep for the next trigger.', currTab);
                scriptProperties.setProperty('invoiceSummary', invoiceSummary);
                Logger.log('Set the invoiceSummary property to %s in prep for the next trigger.', invoiceSummary);
                //   emailMe(invoiceSummary,"End of trigger script. Finished "+currTab+" of "+totalTabs+" at "+currTime);

                // Deletes all triggers in the current project.
                var triggers = ScriptApp.getProjectTriggers();
                for (var i = 0; i < triggers.length; i++) {
                    Logger.log('Removed old trigger %s ', triggers[i]);
                    ScriptApp.deleteTrigger(triggers[i]);
                }

                ScriptApp.newTrigger("processAllInvoices").timeBased().at(new Date(currTime + (6 * toMS))).create();
                //   emailMe(invoiceSummary,"Added new trigger script. Finished "+currTab+" of "+totalTabs+" at "+currTime);
                Logger.log('========== Added new trigger to continue the process at %s ==========',new Date(currTime + (6 * toMS)));
                break;
            } else {
                Logger.log('-- Strating to process invoice for %s --',tabNames[currTab]);
                invoiceSummary += (tabNames[currTab] + " " + createInvoice(month, tabNames[currTab], print, sendEmailReport) + "\n");
              //  testWait();
                Logger.log('-- Finished processing invoice for %s and testWait() was ran --',tabNames[currTab]);
              scriptProperties.setProperty('lastTab', currTab);
            Logger.log('Set the lastTab property to "%s"', currTab);
            }
        
        }

    } catch (e) {
     // console.error("This is a consol error. JG");
      emailMe(e + "\n\nError Log:\n"+Logger.getLog()+"\n\n" + invoiceSummary, "Trigger script error");
        Logger.log('%s. While processing %s', JSON.stringify(e, null, 2), invoiceSummary);
       // throw e;
    }finally{
              if (currTab >= totalTabs) {
            // script should be done and all the tabs processes
          emailMe("Finished with report. \n\n --------------------- \n\n Logger info: \n"+Logger.getLog(),"Done with [" + currTab + " of " + totalTabs+"]")
            emailMe(invoiceSummary, "Totally done with trigger script. Finished " + currTab + " of " + totalTabs);
         ///   testWait();
          scriptProperties.deleteAllProperties(); // removes the saved properties
            Logger.log('-- Removing all script properties --');
            Logger.log('========== Finished with the trigger script ===========');
        } else {
            //  emailMe(invoiceSummary,"Exiting the trigger script. Finished "+currTab+" of "+totalTabs);
        }
    }
  
}

function clearSettings() {
    Logger.log('========== Starting the clearSettings function ===========');
    var scriptProperties = PropertiesService.getScriptProperties();
    // Deletes all triggers in the current project.
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        Logger.log('Removed old trigger %s ', triggers[i]);
        ScriptApp.deleteTrigger(triggers[i]);
    }
    scriptProperties.deleteAllProperties(); // removes the saved properties
    Logger.log('-- Removing all script properties --');
    Logger.log('========== Finished with the clearSettings function ===========');
}