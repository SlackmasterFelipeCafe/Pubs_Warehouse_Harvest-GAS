
function harvestPubs () {
/* contact Phil Brown pbrown@usgs.gov */
  //Define Sheets
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CGGSC_PubsQuery');//Change the sheet name as appropriate
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CMERSC_PubsQuery');//Change the sheet name as appropriate
  var sheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GMEG_PubsQuery');//Change the sheet name as appropriate
  var sheet4 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EMERSC_PubsQuery');//Change the sheet name as appropriate
  var sheet5 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NMIC_PubsQuery');//Change the sheet name as appropriate
  var sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AK_PubsQuery');//Change the sheet name as appropriate
  var sheet7 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GGGSC_PubsQuery');//Change the sheet name as appropriate
  //Clear Sheets
  sheet1.clear;
  sheet2.clear;
  sheet3.clear;
  sheet4.clear;
  sheet5.clear;
  sheet6.clear;
  sheet7.clear;
 //Assign harvest date
  sheet1.getRange(1, 1).setValue('Last USGS Publications Harvest Performed: ' + new Date());
  sheet2.getRange(1, 1).setValue('Last USGS Publications Harvest Performed: ' + new Date());
  sheet3.getRange(1, 1).setValue('Last USGS Publications Harvest Performed: ' + new Date());
  sheet4.getRange(1, 1).setValue('Last USGS Publications Harvest Performed: ' + new Date());
  sheet5.getRange(1, 1).setValue('Last USGS Publications Harvest Performed: ' + new Date());
  sheet6.getRange(1, 1).setValue('Last USGS Publications Harvest Performed: ' + new Date());
  sheet7.getRange(1, 1).setValue('Last USGS Publications Harvest Performed: ' + new Date());
 //Query Pubs for JSON files and load values into sheet
  sheet1.getRange(2, 1).setValue('=ImportJSON("https://pubs.er.usgs.gov/pubs-services/publication/?page_size=25&page_number=1&q=&contributingOffice=Crustal+Geophysics+and+Geochemistry+Science+Center&year=2018","","")' );
  sheet2.getRange(2, 1).setValue('=ImportJSON("https://pubs.er.usgs.gov/pubs-services/publication/?page_size=25&page_number=1&q=&contributingOffice=Central+Mineral+and+Environmental+Resources+Science+Center&year=2018","","")' );
  sheet3.getRange(2, 1).setValue('=ImportJSON("https://pubs.er.usgs.gov/pubs-services/publication/?page_size=25&page_number=1&q=&contributingOffice=Geology%2C+Minerals%2C+Energy%2C+and+Geophysics+Science+Center&year=2018","","")' );
  sheet4.getRange(2, 1).setValue('=ImportJSON("https://pubs.er.usgs.gov/pubs-services/publication/?page_size=25&page_number=1&q=&contributingOffice=Eastern+Mineral+and+Environmental+Resources+Science+Center&year=2018","","")' );
  sheet5.getRange(2, 1).setValue('=ImportJSON("https://pubs.er.usgs.gov/pubs-services/publication/?page_size=25&page_number=1&q=&contributingOffice=National+Minerals+Information+Center&year=2018","","")' );
  sheet6.getRange(2, 1).setValue('=ImportJSON("https://pubs.er.usgs.gov/pubs-services/publication/?page_size=25&page_number=1&q=&year=2018&contributingOffice=Alaska+Science+Center+Geology+Minerals","","")' );
  sheet7.getRange(2, 1).setValue('=ImportJSON("https://pubs.er.usgs.gov/pubs-services/publication/?page_size=25&page_number=1&q=&contributingOffice=Geology%2C+Geophysics%2C+and+Geochemistry+Science+Center&year=2018","","")' );
//Transpose Records for Tagging
  var QuerySheet = 'CGGSC_PubsQuery';
  var TagSheet = 'GGGSC_Tags';
  AllScienceCentersTransposePubsForTagging(QuerySheet,TagSheet);
  QuerySheet = 'CMERSC_PubsQuery';
  TagSheet = 'GGGSC_Tags';
  AllScienceCentersTransposePubsForTagging(QuerySheet,TagSheet);
  QuerySheet = 'EMERSC_PubsQuery';
  TagSheet = 'EMERSC_Tags';
  AllScienceCentersTransposePubsForTagging(QuerySheet,TagSheet);
  QuerySheet = 'GMEG_PubsQuery';
  TagSheet = 'GMEG_Tags';
  AllScienceCentersTransposePubsForTagging(QuerySheet,TagSheet);
  QuerySheet = 'NMIC_PubsQuery';
  TagSheet = 'NMIC_Tags';
  AllScienceCentersTransposePubsForTagging(QuerySheet,TagSheet);
  QuerySheet = 'AK_PubsQuery';
  TagSheet = 'AK_Tags';
  AllScienceCentersTransposePubsForTagging(QuerySheet,TagSheet);
  QuerySheet = 'GGGSC_PubsQuery';
  TagSheet = 'GGGSC_Tags';
  AllScienceCentersTransposePubsForTagging(QuerySheet,TagSheet);
}

function addCheckbox(SheetName,Row,Col) {// Blank Check Box must already be available in cell Row 1 Col 26
/* contact Phil Brown pbrown@usgs.gov */  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);//Change the sheet name as appropriate
  sheet.getRange(1,26).activate();//Blank Check Box MUST BE HERE ALREADY
  sheet.getRange(Row,Col).activate();
  sheet.getRange(1,26).copyTo(sheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  }

function isOdd(num) { return num % 2;}

function createSDCQueryURL() {
/* contact Phil Brown pbrown@usgs.gov */
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GGGSCDataReleases');//Change the sheet name as appropriate
var activeCell = sheet.getActiveCell(); //Detect the ActiveCell
var row = activeCell.getRow(); 

for (var i = 2; i < 96; i++) {
   var doiURL = sheet.getRange(i, 6).getValue();
   sheet.getRange(i, 14).setValue('https://usgs.ornl.gov/sdcsolr/core1/select?fq=web_url:"'+ doiURL + '"&q=*:*&start=0&rows=10&wt=json');
                             }
}

function addJSON () {
/* contact Phil Brown pbrown@usgs.gov */  
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SDC_Records');
  sheet2.clear();//clear worksheet before reloading data
  Logger.log('Sheet is cleared for newdata');
  sheet2.getRange(1, 1).setValue('Last SDC Report Harvest Performed: ' + new Date());
  sheet2.getRange(1, 1).setFontStyle('italic');
  sheet2.getRange(1, 1).setFontWeight('bold')
  for (var i = 2; i < 96; i++) {
    loadJSONvalues (i);
    Utilities.sleep(3000);
  }
  
  
}

function loadJSONvalues (i) {//i + 1 is the row where the values are printed
/* contact Phil Brown pbrown@usgs.gov */
var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GGGSCDataReleases');//Change the sheet name as appropriate
var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SDC_Records');//Change the sheet name as appropriate
var lastRowData = sheet2.getDataRange().getValues();
var lastRow = lastRowData.length;  
	
Logger.log('Last row is: ' + lastRow);
var citation = sheet1.getRange(i, 11).getValue();
var querySDCurl = String(sheet1.getRange(i, 14).getValue());
var res = querySDCurl.replace('"', "%22");
var res2 = res.replace('"', "%22");
sheet2.getRange(lastRow + 1, 1).setValue(citation);
sheet2.getRange(lastRow + 1, 1, 1, 51).setBackground('#FFFF00');
Logger.log('URL is: ' + querySDCurl);
sheet2.getRange(lastRow + 2, 1).setValue('=ImportJSON("' + res2 + '","","")' );



}


function testSetBackground (){
/* contact Phil Brown pbrown@usgs.gov */
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SDC_Records');//Change the sheet name as appropriate
  sheet2.getRange(1,1).setValue('Test');
  sheet2.getRange(1,1,1,51).setBackground('#FFFF00');
}



function AllScienceCentersTransposePubsForTagging(QuerySheet,TagSheet) { //QuerySheet, TagSheet
 /* contact Phil Brown pbrown@usgs.gov */
  //var QuerySheet = 'CGGSC_PubsQuery';
  //var TagSheet = 'CGGSC_Tags';  
  
  //So pubs warehouse keeps changing the column orders and locations for various scince centers
  //Decided to search for column headers to get column indexs rather than hardcoding these index values.
  //Will create functions to find 'Records Indexid' == 'recIndexID, 'Records Ipdsid' == 'recDOI', 'Records Usgscitation' == 'usgsCitation'
  //The elusive column values are then replaced by these variable index values ;)
  
  
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(QuerySheet);//Change the sheet name as appropriate
  var data1 = sheet1.getDataRange().getValues();
  var lastRow1 = sheet1.getLastRow();
  var lastCol1 = sheet1.getLastColumn();
  Logger.log('Last Row Sheet 1: ' + lastRow1);
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TagSheet);//Change the sheet name as appropriate
  
  returnRecIndexIDcolumn (QuerySheet,lastCol1);
  
  var indexRecIndexID = returnRecIndexIDcolumn(QuerySheet,lastCol1);
  var indexRecDOI = returnRecDOIcolumn(QuerySheet,lastCol1);
  var indexUsgsCitation = returnUsgsCitation(QuerySheet,lastCol1);
  
  Logger.log ('Variables Passed: ');
  indexRecIndexID
  indexRecDOI
  indexUsgsCitation
  
  
  Logger.log('Record Index ID: ' + recIndexID); 
 sheet2.getRange(1, 1).setValue(sheet1.getRange(1, 1).getValue());
  for (var k = 3; k < (lastRow1 + 1); k++) {
    var recExists = 0;
    var data2 = sheet2.getDataRange().getValues();
    var lastRow2 = data2.length;
    var recIndexID = sheet1.getRange(k, indexRecIndexID).getValue();
       for (var i = 3; i < (lastRow2 + 1); i++) {
         var rec2Check = sheet2.getRange(i, 9).getValue();
         Logger.log('Compare ' + recIndexID + ' and ' + rec2Check); 
         if (recIndexID == rec2Check) {
                        recExists = 1;
                         Logger.log('Record Exists');
                                       }
                                                }
         
    
       if (recExists == 1) {
          Logger.log('No Record Added End of Loop: ' + k);
                            }
       if (recExists == 0) {
         Logger.log('Adding New Record')
         sheet2.getRange(lastRow2 + 1, 11).setValue(new Date());
         addCheckbox(TagSheet,(lastRow2 + 1),1);
    
              for (var j = 3; j < 9; j++) {
                   addCheckbox(TagSheet,(lastRow2 + 1),j);
                                          }
  
             var recDOI = sheet1.getRange(k, indexRecDOI).getValue()
             sheet2.getRange(lastRow2 + 1, 9).setValue(recIndexID);
             sheet2.getRange(lastRow2 + 1, 12).setValue('https://pubs.er.usgs.gov/publication/' + recIndexID);
             sheet2.getRange(lastRow2 + 1, 10).setValue(recDOI);
             sheet2.getRange(lastRow2 + 1, 13).setValue('https://doi.org/' + recDOI);
             sheet2.getRange(lastRow2 + 1, 2).setValue(sheet1.getRange(k, indexUsgsCitation).getValue());
             if (isOdd(k) == 0) {
               sheet2.getRange(k,1,1,14).setBackground('#eaeaea');
             }   
             Logger.log('Record Added End of Loop: ' + k);
                            }
             
                                               }
         
                                             }
  
  function returnRecIndexIDcolumn (QuerySheet,lastCol1) {
    /* contact Phil Brown pbrown@usgs.gov */
    var colHeader = []
    var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(QuerySheet);
    for (var j = 1; j < lastCol1; j++)
     colHeader[j] = sheet1.getRange(2, j).getValue();
    {
     for (var j = 1; j < lastCol1; j++)
       if (colHeader[j] == 'Records Indexid') {
          var index = j;
    Logger.log('Leaving returnRecIndexIDcolumn');
    Logger.log('Last Column to search: ' + lastCol1);                        }
    Logger.log('RecIndexIDcolumn: ' + index);
    Logger.log('Array: ' + colHeader);

    return index;
  }
  }
  
  function returnRecDOIcolumn (QuerySheet,lastCol1) {
    /* contact Phil Brown pbrown@usgs.gov */
    //QuerySheet,lastCol1
    var colHeader = []
    var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(QuerySheet);
    for (var j = 1; j < lastCol1; j++)
     colHeader[j] = sheet1.getRange(2, j).getValue();
    {
     for (var j = 1; j < lastCol1; j++)
       if (colHeader[j] == 'Records Ipdsid') {
          var index = j;
                            }
    Logger.log('Leaving returnRecDOIcolumn');
    Logger.log('Last Column to search: ' + lastCol1);                        }
    Logger.log('Records Ipdsid: ' + index);
    Logger.log('Array: ' + colHeader);

    return index;
  }
  
  
  function returnUsgsCitation (QuerySheet,lastCol1) {
    /* contact Phil Brown pbrown@usgs.gov */
    var colHeader = []
    var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(QuerySheet);
    for (var j = 1; j < lastCol1; j++)
     colHeader[j] = sheet1.getRange(2, j).getValue();
    {
     for (var j = 1; j < lastCol1; j++)
       if (colHeader[j] === 'Records Usgscitation') {
          var index = j;
                            }
    Logger.log('returnUsgsCitation');
    Logger.log('Last Column to search: ' + lastCol1);                        }
    Logger.log('Records Usgscitation: ' + index);
    Logger.log('Array: ' + colHeader);

    return index;
 
    }
  
  
