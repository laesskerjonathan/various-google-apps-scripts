function yearlyReport() {
  var d = new Date(),
      year = d.getYear();
  var statisticsSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var overviewSheet = statisticsSpreadSheet.getSheetByName("Overview");
  
  //DeleteAllData
  var start, end;
  start = 1;
  end = overviewSheet.getLastRow() + 10;
  Logger.log(start);
  Logger.log(end);
  overviewSheet.deleteRows(start, end);
  overviewSheet.appendRow(["Title", "Number of Events", "Total Duration", "Average Duration"]);
  
  var settingsSheet = statisticsSpreadSheet.getSheetByName("Settings");
  var settingsAmount = settingsSheet.getLastRow();
  var settingsData = settingsSheet.getRange(1, 1, settingsAmount, 5).getValues(); 

  var calendars = CalendarApp.getAllCalendars();
  for(var i = 0; i < calendars.length; i++) {
    
    for(var k = 0; k < settingsAmount; k++) {
      if(settingsData[k][0].toLowerCase() === calendars[i].getName().toLowerCase()){
        Logger.log("Calendar: " + settingsData[k][0]);
        
        var calendar = calendars[i];
        
        var startOfTheYear =  new Date(year,0,1);
        var now = new Date();
        
        var events = calendar.getEvents(startOfTheYear, now);
        var totalDuration = 0.0;
        for(var j = 0; j < events.length; j++) {
          var event = events[j];
          var duration = (event.getEndTime() - event.getStartTime()) / 3600000;
          Logger.log(event.getTitle() + ": " + duration);
          totalDuration += duration;
        }
        
        var numberOfWeeks = weeksBetween(startOfTheYear, now);
        var numberOfEvents = events.length;
        var averageDuration = totalDuration / numberOfWeeks;
        Logger.log('Number of events: ' + events.length);
        Logger.log('Total Duration: ' + totalDuration);
        overviewSheet.appendRow([calendar.getName(), numberOfEvents, totalDuration, averageDuration]);
    
      }
    }
  }
  
}

function weeksBetween(d1, d2) {
    return Math.round((d2 - d1) / (7 * 24 * 60 * 60 * 1000));
}

function getMondays() {
  var today = new Date(),
      year = today.getYear(),
      d = new Date(year,0,1),
      mondays = [];
  
  d.setDate(1);
  
  // Get the first Monday in the month
  while (d.getDay() !== 1) {
    d.setDate(d.getDate() + 1);
  }
  
  // Get all the other Mondays in the month
  while (d.getTime() <= today.getTime()) {
    var pushDate = new Date(d.getTime());
    mondays.push(pushDate);
    d.setDate(d.getDate() + 7);
  }
  return mondays;
}


function makeWeeklyReportForWholeYear(){
  var d = new Date(),
      year = d.getYear();
  var statisticsSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var mondays = getMondays();
  
  var settingsSheet = statisticsSpreadSheet.getSheetByName("Settings");
  var settingsAmount = settingsSheet.getLastRow();
  var settingsData = settingsSheet.getRange(1, 1, settingsAmount, 5).getValues(); 

  var calendars = CalendarApp.getAllCalendars();
  for(var i = 0; i < calendars.length; i++) {
    for(var k = 0; k < settingsAmount; k++) {
      if(settingsData[k][0].toLowerCase() === calendars[i].getName().toLowerCase()){
        //Get Sheet
        var currentSheet = statisticsSpreadSheet.getSheetByName(calendars[i].getName());
        Logger.log(typeof(currentSheet));
        if(typeof(currentSheet) === "undefined") {
          currentSheet = statisticsSpreadSheet.insertSheet(3);   
          currentSheet.setName(calendars[i].getName())
        }
        var start, end;
        start = 1;
        end = currentSheet.getLastRow() + 10;
        Logger.log(start);
        Logger.log(end);
        currentSheet.deleteRows(start, end);
        currentSheet.appendRow(["Week", "Number of Events", "Total Duration", "Percentage"]);
        
        var calendar = calendars[i];
        //Calculate One Week
        for(var z = 0; z < mondays.length - 1; z++) {
          
          var start = mondays[z], end = mondays[z + 1];
          var events = calendar.getEvents(start, end);
          var totalDuration = 0.0;
          for(var j = 0; j < events.length; j++) {
            var event = events[j];
            var duration = (event.getEndTime() - event.getStartTime()) / 3600000;
            Logger.log(event.getTitle() + ": " + duration);
            totalDuration += duration;
            
            var numberOfEvents = events.length;
            var percentage = totalDuration / 40 * 100;
          }
          Logger.log('Number of events: ' + events.length);
          Logger.log('Total Duration: ' + totalDuration);
          currentSheet.appendRow([mondays[z], numberOfEvents, totalDuration, percentage]);
        }
      }
    }
  }
}


function makeWeeklyReport(year, weekNumber){
  year = 2017;
  weekNumber = 1;
  var d = new Date();
  var statisticsSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var monday = getMondays()[weekNumber - 1];
  
  var settingsSpreadsheet = SpreadsheetApp.openById('1G2afXI8KatS2zkjPE6HCiB3zgI6eGhV0WZM-AiRpgH0');
  
  var calendarSettingsSheet = statisticsSpreadSheet.getSheetByName("Settings");
  var settingsAmount = calendarSettingsSheet.getLastRow();
  var settingsData = calendarSettingsSheet.getRange(1, 1, settingsAmount, 5).getValues(); 

  var calendars = CalendarApp.getAllCalendars();
  var spreadsheetIterator = DriveApp.getFolderById("1tE1XuelJzd3Yih-OQQ-XRfU1TrYPFTwj").getFilesByName("Weekly_" + year + "-" + weekNumber);
  var spreadsheet = null;
  if(spreadsheetIterator.hasNext()){
    var spreadsheetFile = spreadsheetIterator.next();
    spreadsheet = SpreadsheetApp.open(spreadsheetFile);
  } else {
    spreadsheet = SpreadsheetApp.create("Weekly_" + year + "-" + weekNumber);
    var file = DriveApp.getFileById(spreadsheet.getId());
    file.getParents().next().removeFile(file);
    DriveApp.getFolderById('1tE1XuelJzd3Yih-OQQ-XRfU1TrYPFTwj').addFile(file);
  }
  var logSheet = spreadsheet.getSheetByName("LogSheet");
  if(logSheet === null) {
    logSheet = spreadsheet.insertSheet(1);   
    logSheet.setName("LogSheet");
  }
  var start, end;
  start = 1;
  end = logSheet.getLastRow() + 10;
  logSheet.deleteRows(start, end);
        
  logSheet.appendRow(["StartTime", "EndTime", "Title", "ErrorTag"]);
  
  for(var i = 0; i < calendars.length; i++) {
    if(calendars[i].getName() === "Voluteering") {
    
      //LoadSettings
    var projectSettingsSheet = settingsSpreadsheet.getSheetByName(calendars[i].getName());
    var projectSettingsAmount = projectSettingsSheet.getLastRow();
    var projectSettingsData = projectSettingsSheet.getRange(1, 1, settingsAmount, 3).getValues(); 
    var settingsJson = {};
    for(var h = 1; h < projectSettingsData.length; h++){
      if(typeof(settingsJson[projectSettingsData[h][0]]) === "undefined") {
        settingsJson[projectSettingsData[h][0]] = {};
      }        
      settingsJson[projectSettingsData[h][0]][projectSettingsData[h][1]] = {"time" : 0};
    }
    Logger.log(settingsJson);
  
    for(var k = 0; k < settingsAmount; k++) {
      
      
      if(settingsData[k][0].toLowerCase() === calendars[i].getName().toLowerCase()){
        //Get Sheet
        
        var currentSheet = spreadsheet.getSheetByName(calendars[i].getName());
        if(currentSheet === null) {
          currentSheet = spreadsheet.insertSheet(3);   
          currentSheet.setName(calendars[i].getName())
        }
        var start, end;
        start = 1;
        end = currentSheet.getLastRow() + 10;
        currentSheet.deleteRows(start, end);
        currentSheet.appendRow(["Project", "Task", "Total Duration", "Percentage"]);
        
        var calendar = calendars[i];
        //Calculate One Week
        
        var start = monday, end = new Date();
        
        end.setDate(monday.getDate()+7);
        
        var events = calendar.getEvents(start, end);
        var totalDuration = 0.0;
        
        var hashtags = {};   
        for(var j = 0; j < events.length; j++) {
          var event = events[j];
                   
          var description = event.getDescription();
             
          var hashtagRegex = new RegExp("#([a-zA-Z0-1.]*)", "gi");
          var match;
          var error = true;
          var errorTag = null;
          while ((match = hashtagRegex.exec(description)) != null) {
            var project = match[1].split(".")[0], task = match[1].split(".")[1];
            
            if(typeof(settingsJson[project]) !== "undefined" 
               && typeof(settingsJson[project][task]) !== "undefined") {  
              var duration = (event.getEndTime() - event.getStartTime()) / 3600000;
              settingsJson[project][task].time += duration;
              error = false;
              break;
            } else {
              errorTag = match[1];
              Logger.log(errorTag);
            }
          }
          if(error) {
            logSheet.appendRow([event.getStartTime(), event.getEndTime(), event.getTitle(), errorTag])
          }
        }
        
        Logger.log(settingsJson);
      }
    }
  }
  }
}