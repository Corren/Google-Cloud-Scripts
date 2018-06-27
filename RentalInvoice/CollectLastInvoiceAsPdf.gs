function getNewestFileInFolder(rentalFolderID) {
  var arryFileDates,file,fileDate,files,folder,folders,
      newestDate,newestFileID,objFilesByDate;

  folders = DriveApp.getFileById(rentalFolderID).getFolders();  
  arryFileDates = [];
  objFilesByDate = {};
  
  var milsecPerDay =  24 * 60 * 60 * 1000;
  var invoiceDaysAgeThreshold = 15;

  while (folders.hasNext()) {
    folder = folders.next();

    files = folder.getFilesByType("application/vnd.google-apps.spreadsheet");
    fileDate = "";

    while (files.hasNext()) {
      file = files.next();
      Logger.log('xxxx: file data: ' + file.getLastUpdated());
      Logger.log('xxxx: file name: ' + file.getName());
      Logger.log('xxxx: mime type: ' + file.getMimeType())
      Logger.log(" ");

	  
	  
	  
	  
	     if (new Date() - file.getLastUpdated() > (invoiceDaysAgeThreshold*milsecPerDay)) {
			Logger.log('Passed: file not older than %s days.',invoiceDaysAgeThreshold);
			fileDate = file.getLastUpdated();
			objFilesByDate[fileDate] = file.getId(); //Create an object of file names by file ID
			arryFileDates.push(file.getLastUpdated());
		}
   
   

    }
    arryFileDates.sort(function(a,b){return b-a});

    Logger.log('arryFileDates: '+arryFileDates);

    newestDate = arryFileDates[0];
    Logger.log('Newest date is: ' + newestDate);

    newestFileID = objFilesByDate[newestDate];

    Logger.log('newestFile ID: ' + newestFileID);
    //return newestFile;
    var pdfName = SpreadsheetApp.openById(newestFileID).getName()+".pdf"; // this gets you the rental log spreadsheet 
	Logger.log('pdfName: ' + pdfName);
	
	// get the destination folder
	var destinationFolderID = lastMonthFolderID();
	var destinationFolder = DriveApp.getFolderById(destinationFolderID);
	Logger.log('Got destination folder');

	// checks if the file already exist in the destination folder
	var haBDs = destinationFolder.getFilesByName(pdfName);
		if(!haBDs.hasNext()){
		Logger.log(pdfName+'does not already exist in'+destinationFolder.getName());

    
	//Copy whole spreadsheet
	var lastSpreadsheetCopy = SpreadsheetApp.openById(DriveApp.getFileById(newestFileID).makeCopy(pdfName, folder))
	Logger.log('Made a copy of : ' + SpreadsheetApp.openById(newestFileID).getName());
	

	
	//save to pdf
	var theBlob = lastSpreadsheetCopy.getBlob().getAs('application/pdf').setName(pdfName);
	var newFile = destinationFolder.createFile(theBlob);

	//Delete the temporary sheet
	DriveApp.getFileById(destSpreadsheet.getId()).setTrashed(true);
 	}else
	{
		Logger.log(pdfName+'does not already exist in'+destinationFolder.getName());
	}
  }
}




function lastMonthFolderID()
{
	
	 Logger.log("======Starting lastMonthFolderID======");
	 
	 var LAST_INVOICE_IN_EACH_MONTH_FOLDER_ID = "1cCf1ffI9kGtq_j-97U6D0pf1TVgXtxc3";
        // gets the folder for the last month
        var par_fdr = DriveApp.getFolderById(LAST_INVOICE_IN_EACH_MONTH_FOLDER_ID);
		var lastMonthNameAndYear = getLastMonthNameAndYear();
        // This checks if the folder for the last month exist. if not, creates it
        try { // gets the folder for the last month if exist
            var newFdr = par_fdr.getFoldersByName(lastMonthNameAndYear).next();
			Logger.log("Found the folder for "+lastMonthNameAndYear);
        } catch (e) { // creates the folder for the last month if not exist
            var newFdr = par_fdr.createFolder(lastMonthNameAndYear);
			Logger.log("Creating the folder for "+lastMonthNameAndYear);
        }
		var lastMonthNameAndYearFolderID = newFdr.getId();
		Logger.log("Folder ID [%s}",lastMonthNameAndYearFolderID);
	
	
		 Logger.log("======Ending lastMonthFolderID======");
	return lastMonthNameAndYearFolderID;
	
}


function getLastMonthNameAndYear()
{
	Logger.log('======Starting getLastMonthNameAndYear ======' );
	var date = new Date();
    var currentYear = date.getYear();
    var currentDay = date.getDate();
    var currentMonth = date.getMonth();
	
	if (currentMonth == 0) // if you are billing dec in jan
        {
            var lastMonthYear = currentYear - 1;
            var lastMonth = 11; // months are from 0(jan) - 11(dec)
        } else {
            var lastMonthYear = currentYear;
            var lastMonth = date.getMonth() - 1;
        }

        // Billing date range logic
        // new Date(year, month, day, hours, minutes, seconds, milliseconds);
  
        var lastMonthDate = new Date(lastMonthYear, lastMonth, 1, 0, 0, 0, 0);

		var lastMonthNameAndYear = Utilities.formatDate(lastMonthDate, Session.getScriptTimeZone(), "MMM yyyy");
Logger.log('======Ending getLastMonthNameAndYear - returning: [%s] ======',lastMonthNameAndYear);
return lastMonthNameAndYear;
	
}






