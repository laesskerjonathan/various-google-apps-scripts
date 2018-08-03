function myFunction() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var parentFolders = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents();
  var parentFolder = parentFolders.next();
  var childFolderIterator = parentFolder.getFolders();
  while(childFolderIterator.hasNext()) {
    var possibleSpecsFolder = childFolderIterator.next();
    if(possibleSpecsFolder.getName() == "Specs") {
      break;
    }
  }

  var specsFolder = possibleSpecsFolder;
  
  var lastSpecsRow = sheet.getLastRow();

  //DeleteAllData
  var start, end;
  start = 1;
  end = sheet.getLastRow() + 10;
  Logger.log(start);
  Logger.log(end);
  sheet.deleteRows(start, end);

  
  var childFilesIterator = parentFolder.getFiles();
  while(childFilesIterator.hasNext()) {
    var possibleManualOverviewSheet = childFilesIterator.next();
    if(possibleManualOverviewSheet.getName() == "Manuelle Geräte Informationen") {
      break;
    }
  }

       //GetManualOverview
  var manualOverviewSheet = SpreadsheetApp.open(possibleManualOverviewSheet).getActiveSheet();
  var lastRow = manualOverviewSheet.getLastRow();
  
  var data = manualOverviewSheet.getRange(1, 1, lastRow, 5).getValues(); 
  Logger.log(data);
  
  //Write Header
  sheet.appendRow(['Computer Name', 'Owner', 'Donor', 'Remarks', 'Brand Name', 'Operating System', 'Processor', 'Memory Capacity', 'Total Memory', 'Drive Capacity', 'License Key']);
  sheet.getRange(1, 1, 1, 20).setFontWeight("bold");
  
  var specsFileIterator = specsFolder.getFiles();
  while(specsFileIterator.hasNext()) {
    var specsFile = specsFileIterator.next();
    
    var html = specsFile.getBlob().getDataAsString();
    
    var computerNameRegExp = new RegExp("<TR><TD><TD CLASS=di>Computer Name:<TD CLASS=di>(.*)", "gi");
    var computerName = computerNameRegExp.exec(html)[1];
    Logger.log("Computer Name: " + computerName);
    
    var brandNameRegExp = new RegExp("<TR><TD><TD CLASS=di>Computer Brand Name:<TD CLASS=di>(.*)", "gi"); 
    var brandName = brandNameRegExp.exec(html)[1];
    Logger.log("Brand Name: " + brandName);
    
    var operatingSystemRegExp = new RegExp("<TR><TD><TD>Operating System:<TD>(.*)", "gi"); 
    var operatingSystem = operatingSystemRegExp.exec(html)[1];
    Logger.log("Operating System: " + operatingSystem);
    
    var processorRegExp = new RegExp("<TR><TD><TD CLASS=di>Processor Name:<TD CLASS=di>(.*)", "gi"); 
    var processor = processorRegExp.exec(html)[1];
    Logger.log("Processor: " + processor);
    
    var memoryCapacityRegExp = new RegExp("<TR><TD><TD>Memory Capacity:<TD>(.*)", "gi"); 
    var memoryCapacity = memoryCapacityRegExp.exec(html)[1];
    Logger.log("Memory Capacity: " + memoryCapacity);      
    
    var totalMemoryRegExp = new RegExp("<TR><TD><TD>Total Memory Size:<TD>(.*)", "gi"); 
    var totalMemory = totalMemoryRegExp.exec(html)[1];
    Logger.log("Total Memory: " + totalMemory);
    
    var driveCapacityRegExp = new RegExp("<TR><TD><TD CLASS=di>Drive Capacity:<TD CLASS=di>.*? MBytes (.*)", "gi"); 
    var driveCapacity = driveCapacityRegExp.exec(html)[1];
    driveCapacity = driveCapacity.substring(1, driveCapacity.length-1);
    Logger.log("Drive Capacity: " + driveCapacity);
    
    var remarks = "";
    var donor = "";
    var licenseKey = "";
    var owner = "";

    for(var i = 0; i < lastRow; i++) {
      if(data[i][0].toLowerCase() === computerName.toLowerCase()){
        Logger.log("ManualName: " + data[i][0] + ", Automatic: " + computerName);
        remarks = data[i][1];
        owner = data[i][2];
        donor = data[i][3];
        licenseKey = data[i][4];
      }
    }
      
    sheet.appendRow([computerName, owner, donor, remarks, brandName, operatingSystem, processor, memoryCapacity, totalMemory, driveCapacity, licenseKey]);
  }
  sheet.sort(1);
}
