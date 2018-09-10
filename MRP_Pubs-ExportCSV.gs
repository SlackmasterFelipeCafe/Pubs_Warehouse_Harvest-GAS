/*
 * script to export data in all sheets in the current spreadsheet as individual csv files
 * files will be named according to the name of the sheet
 * author: Michael Derazon https://gist.github.com/mderazon/9655893
 * contact pbrown@usgs.gov
*/

///function onOpen() {
  ///var ss = SpreadsheetApp.getActiveSpreadsheet();
  ///var csvMenuEntries = [{name: "export as csv files", functionName: "saveAsCSV"}];
  ///ss.addMenu("csv", csvMenuEntries);
///};

function saveAsCSV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  // create a folder from the name of the spreadsheet
  var folder = DriveApp.createFolder(ss.getName().toLowerCase().replace(/ /g,'_') + '_csv_' + new Date().getTime()); 
  
  for (var i = 0 ; i < sheets.length ; i++) {
    var sheet = sheets[i];
    // append ".csv" extension to the sheet name
    fileName = sheet.getName() + ".csv";
    // convert all available sheet data to csv format
    var csvFile = convertRangeToCsvFile_(fileName, sheet);
    // create a file in the Docs List with the given name and the csv data
    folder.createFile(fileName, csvFile);
  }
  // move folder to MrPOP and delete olde folder
  var targetFolderId = '1_ndYmDoQC8HyTfnU72Nq9MyU7NjayTy2';
  var sourceFolderId = folder.getId();
  moveFolderToFolder (sourceFolderId,targetFolderId)
  Browser.msgBox('Files are waiting in a folder named ' + folder.getName() + 'contained in MrPOP folder https://drive.google.com/drive/folders/1_ndYmDoQC8HyTfnU72Nq9MyU7NjayTy2');
}


function testMoveFolder (){
var targetFolderId = '1_ndYmDoQC8HyTfnU72Nq9MyU7NjayTy2';
  var sourceFolderId = '1WpCNVOzj0_y9YEE2UkTSYupYnwOIE8eu';
  moveFolderToFolder (sourceFolderId,targetFolderId)
}

function moveFolderToFolder(sourceFolderId, targetFolderId) {
  //added by Phil B. to move folder to MrPOP, contact: pbrown@usgs.gov
  var targetFolder = DriveApp.getFolderById(targetFolderId);
  var sourceFolder = DriveApp.getFolderById(sourceFolderId);
  var currentFolders = sourceFolder.getParents();
  while (currentFolders.hasNext()) {
    var currentFolder = currentFolders.next();
    currentFolder.removeFolder(sourceFolder);
  }
  targetFolder.addFolder(sourceFolder);
};

function convertRangeToCsvFile_(csvFileName, sheet) {
  // get available data range in the spreadsheet
  var activeRange = sheet.getDataRange();
  try {
    var data = activeRange.getValues();
    var csvFile = undefined;

    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }

        // join each row's columns
        // add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join(",") + "\r\n";
        }
        else {
          csv += data[row];
        }
      }
      csvFile = csv;
    }
    return csvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}
