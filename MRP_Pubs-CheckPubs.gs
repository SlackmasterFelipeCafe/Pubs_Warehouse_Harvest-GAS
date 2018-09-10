
function harvestPubs () {
/* contact Phil Brown pbrown@usgs.gov */
  //Define Sheets
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CGGSC_PubsQuery');//Change the sheet name as appropriate
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CMERSC_PubsQuery');//Change the sheet name as appropriate
  var sheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GMEG_PubsQuery');//Change the sheet name as appropriate
  var sheet4 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EMERSC_PubsQuery');//Change the sheet name as appropriate
  var sheet5 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NMIC_PubsQuery');//Change the sheet name as appropriate
  var sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AK_PubsQuery');//Change the sheet name as appropriate
  //Clear Sheets
  sheet1.clear;
  sheet2.clear;
  sheet3.clear;
  sheet4.clear;
  sheet5.clear;
  sheet6.clear;
 //Assign harvest date
  sheet1.getRange(1, 1).setValue('Last USGS Publications Harvest Performed: ' + new Date());
  sheet2.getRange(1, 1).setValue('Last USGS Publications Harvest Performed: ' + new Date());
  sheet3.getRange(1, 1).setValue('Last USGS Publications Harvest Performed: ' + new Date());
  sheet4.getRange(1, 1).setValue('Last USGS Publications Harvest Performed: ' + new Date());
  sheet5.getRange(1, 1).setValue('Last USGS Publications Harvest Performed: ' + new Date());
  sheet6.getRange(1, 1).setValue('Last USGS Publications Harvest Performed: ' + new Date());
 //Query Pubs for JSON files and load values into sheet
  sheet1.getRange(2, 1).setValue('=ImportJSON("https://pubs.er.usgs.gov/pubs-services/publication/?page_size=25&page_number=1&q=&contributingOffice=Crustal+Geophysics+and+Geochemistry+Science+Center&year=2018","","")' );
  sheet2.getRange(2, 1).setValue('=ImportJSON("https://pubs.er.usgs.gov/pubs-services/publication/?page_size=25&page_number=1&q=&contributingOffice=Central+Mineral+and+Environmental+Resources+Science+Center&year=2018","","")' );
  sheet3.getRange(2, 1).setValue('=ImportJSON("https://pubs.er.usgs.gov/pubs-services/publication/?page_size=25&page_number=1&q=&contributingOffice=Geology%2C+Minerals%2C+Energy%2C+and+Geophysics+Science+Center&year=2018","","")' );
  sheet4.getRange(2, 1).setValue('=ImportJSON("https://pubs.er.usgs.gov/pubs-services/publication/?page_size=25&page_number=1&q=&contributingOffice=Eastern+Mineral+and+Environmental+Resources+Science+Center&year=2018","","")' );
  sheet5.getRange(2, 1).setValue('=ImportJSON("https://pubs.er.usgs.gov/pubs-services/publication/?page_size=25&page_number=1&q=&contributingOffice=National+Minerals+Information+Center&year=2018","","")' );
  sheet6.getRange(2, 1).setValue('=ImportJSON("https://pubs.er.usgs.gov/pubs-services/publication/?page_size=25&page_number=1&q=&year=2018&contributingOffice=Alaska+Science+Center+Geology+Minerals","","")' );
//Transpose Records for Tagging
  var QuerySheet = 'CGGSC_PubsQuery';
  var TagSheet = 'CGGSC_Tags';
  transposePubsForTagging(QuerySheet,TagSheet);
  QuerySheet = 'CMERSC_PubsQuery';
  TagSheet = 'CMERSC_Tags';
  transposePubsForTagging(QuerySheet,TagSheet);
  QuerySheet = 'EMERSC_PubsQuery';
  TagSheet = 'EMERSC_Tags';
  transposePubsForTagging(QuerySheet,TagSheet);
  QuerySheet = 'GMEG_PubsQuery';
  TagSheet = 'GMEG_Tags';
  GMEGtransposePubsForTagging(QuerySheet,TagSheet);
  QuerySheet = 'NMIC_PubsQuery';
  TagSheet = 'NMIC_Tags';
  NMICtransposePubsForTagging(QuerySheet,TagSheet);
  QuerySheet = 'AK_PubsQuery';
  TagSheet = 'AK_Tags';
  transposePubsForTagging(QuerySheet,TagSheet);
}

function addCheckbox(SheetName,Row,Col) {// Blank Check Box must already be available in cell Row 1 Col 26
/* contact Phil Brown pbrown@usgs.gov */  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);//Change the sheet name as appropriate
  sheet.getRange(1,26).activate();//Blank Check Box MUST BE HERE ALREADY
  sheet.getRange(Row,Col).activate();
  sheet.getRange(1,26).copyTo(sheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  }

function isOdd(num) { return num % 2;}

function transposePubsForTagging(QuerySheet,TagSheet) { //QuerySheet, TagSheet
  //var QuerySheet = 'CGGSC_PubsQuery';
  //var TagSheet = 'CGGSC_Tags';  
  
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(QuerySheet);//Change the sheet name as appropriate
  var data1 = sheet1.getDataRange().getValues();
  var lastRow1 = data1.length;  
  Logger.log('Last Row Sheet 1: ' + lastRow1);
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TagSheet);//Change the sheet name as appropriate
                             
  
  Logger.log('Record Index ID: ' + recIndexID); 
 sheet2.getRange(1, 1).setValue(sheet1.getRange(1, 1).getValue());
  for (var k = 3; k < (lastRow1 + 1); k++) {
    var recExists = 0;
    var data2 = sheet2.getDataRange().getValues();
    var lastRow2 = data2.length;
    var recIndexID = sheet1.getRange(k, 9).getValue();
       for (var i = 3; i < (lastRow2 + 1); i++) {
         var rec2Check = sheet2.getRange(i, 10).getValue();
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
         sheet2.getRange(lastRow2 + 1, 12).setValue(new Date());
         addCheckbox(TagSheet,(lastRow2 + 1),1);
    
              for (var j = 3; j < 10; j++) {
                   addCheckbox(TagSheet,(lastRow2 + 1),j);
                                          }
  
             var recDOI = sheet1.getRange(k, 25).getValue()
             sheet2.getRange(lastRow2 + 1, 10).setValue(recIndexID);
             sheet2.getRange(lastRow2 + 1, 13).setValue('https://pubs.er.usgs.gov/publication/' + recIndexID);
             sheet2.getRange(lastRow2 + 1, 11).setValue(recDOI);
             sheet2.getRange(lastRow2 + 1, 14).setValue('https://doi.org/' + recDOI);
             sheet2.getRange(lastRow2 + 1, 2).setValue(sheet1.getRange(k, 26).getValue());
             if (isOdd(k) == 0) {
               sheet2.getRange(k,1,1,15).setBackground('#eaeaea');
             }   
   
             Logger.log('Record Added End of Loop: ' + k);
                            }
 
             
            
                                               }
     
    
                                             }


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
var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SDC_Records');//Change the sheet name as appropriate
  sheet2.getRange(1,1).setValue('Test');
  sheet2.getRange(1,1,1,51).setBackground('#FFFF00');
}


function GMEGtransposePubsForTagging(QuerySheet,TagSheet) { //QuerySheet, TagSheet
  //var QuerySheet = 'CGGSC_PubsQuery';
  //var TagSheet = 'CGGSC_Tags';  
  
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(QuerySheet);//Change the sheet name as appropriate
  var data1 = sheet1.getDataRange().getValues();
  var lastRow1 = data1.length;  
  Logger.log('Last Row Sheet 1: ' + lastRow1);
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TagSheet);//Change the sheet name as appropriate
                             
  
  Logger.log('Record Index ID: ' + recIndexID); 
 sheet2.getRange(1, 1).setValue(sheet1.getRange(1, 1).getValue());
  for (var k = 3; k < (lastRow1 + 1); k++) {
    var recExists = 0;
    var data2 = sheet2.getDataRange().getValues();
    var lastRow2 = data2.length;
    var recIndexID = sheet1.getRange(k, 9).getValue();
       for (var i = 3; i < (lastRow2 + 1); i++) {
         var rec2Check = sheet2.getRange(i, 10).getValue();
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
         sheet2.getRange(lastRow2 + 1, 12).setValue(new Date());
         addCheckbox(TagSheet,(lastRow2 + 1),1);
    
              for (var j = 3; j < 10; j++) {
                   addCheckbox(TagSheet,(lastRow2 + 1),j);
                                          }
  
             var recDOI = sheet1.getRange(k, 30).getValue()
             sheet2.getRange(lastRow2 + 1, 10).setValue(recIndexID);
             sheet2.getRange(lastRow2 + 1, 13).setValue('https://pubs.er.usgs.gov/publication/' + recIndexID);
             sheet2.getRange(lastRow2 + 1, 11).setValue(recDOI);
             sheet2.getRange(lastRow2 + 1, 14).setValue('https://doi.org/' + recDOI);
             sheet2.getRange(lastRow2 + 1, 2).setValue(sheet1.getRange(k, 31).getValue());
             if (isOdd(k) == 0) {
               sheet2.getRange(k,1,1,15).setBackground('#eaeaea');
             }   
             Logger.log('Record Added End of Loop: ' + k);
                            }
 
             
            
                                               }
     
    
                                             }

function NMICtransposePubsForTagging(QuerySheet,TagSheet) { //QuerySheet, TagSheet
  //var QuerySheet = 'CGGSC_PubsQuery';
  //var TagSheet = 'CGGSC_Tags';  
  
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(QuerySheet);//Change the sheet name as appropriate
  var data1 = sheet1.getDataRange().getValues();
  var lastRow1 = data1.length;  
  Logger.log('Last Row Sheet 1: ' + lastRow1);
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TagSheet);//Change the sheet name as appropriate
                             
  
  Logger.log('Record Index ID: ' + recIndexID); 
 sheet2.getRange(1, 1).setValue(sheet1.getRange(1, 1).getValue());
  for (var k = 3; k < (lastRow1 + 1); k++) {
    var recExists = 0;
    var data2 = sheet2.getDataRange().getValues();
    var lastRow2 = data2.length;
    var recIndexID = sheet1.getRange(k, 9).getValue();
       for (var i = 3; i < (lastRow2 + 1); i++) {
         var rec2Check = sheet2.getRange(i, 10).getValue();
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
         sheet2.getRange(lastRow2 + 1, 12).setValue(new Date());
         addCheckbox(TagSheet,(lastRow2 + 1),1);
    
              for (var j = 3; j < 10; j++) {
                   addCheckbox(TagSheet,(lastRow2 + 1),j);
                                          }
  
             var recDOI = sheet1.getRange(k, 25).getValue()
             sheet2.getRange(lastRow2 + 1, 10).setValue(recIndexID);
             sheet2.getRange(lastRow2 + 1, 13).setValue('https://pubs.er.usgs.gov/publication/' + recIndexID);
             sheet2.getRange(lastRow2 + 1, 11).setValue(recDOI);
             sheet2.getRange(lastRow2 + 1, 14).setValue('https://doi.org/' + recDOI);
             sheet2.getRange(lastRow2 + 1, 2).setValue(sheet1.getRange(k, 59).getValue());
             if (isOdd(k) == 0) {
               sheet2.getRange(k,1,1,15).setBackground('#eaeaea');
             }       
         Logger.log('Record Added End of Loop: ' + k);
                            }
 
             
            
                                               }
     
    
                                             }
