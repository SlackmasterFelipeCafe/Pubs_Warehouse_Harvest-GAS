# Pubs Warehouse Harvest GAS
GGGSC SDC HarvestGAS - Google Application Scripts being used by the GGGSC for tracking harvesting new pubblication records for various science centers daily.  These records can then be used to track science center publications and how they are tagged in the USGS CMS

US Geological Survey (USGS)

Geology, Geophysics, and Geochemistry Science Center (GGGSC)

Google Application Scripts (GAS)

Contact Phil Brown (pbrown@usgs.gov)

USGS Profile: https://www.usgs.gov/staff-profiles/philip-j-brown

ORCID: https://orcid.org/0000-0002-2415-7462

GitHub: https://github.com/pbrown-usgs


## MRP_Pubs-CheckPubs:

Google sheet functions that sets up queries of the Science Data Catalog for a list of Data Releases based upon DOI numbers as the Data Unique Identifier.  These include:

- **harvestPubs**, function that sets up new worksheets and calls query function for the Pubs Warehouse API

- **addCheckbox**, function that adds checkboxes to new worksheet rows and columns

- **isOdd**, function that determines if a number is even or odd

- **createSDCQueryURL**, function that creates the URL used for the Pubs Warehouse API query

- **addJSON**, function that calls functions that loads, parses JSON queries

- **loadJSONvalues**, function that displays appropriat JSON array values to Google Worksheets in a flatfile (table) format with an approipriate header line

- **testSetBackground**, function that tests how to set the background color of a cell in a Google Sheet worksheet

- **GMEGtransposePubsForTagging**, function that culls harvest results, transposes and displays for user tag tracking interaction on a seperate Google Sheet worksheet.  Used to transpose format unique to GMEG records only

- **NMICtransposePubsForTagging**, same as function above but used to transpose formating unique to NMIC records only

## MRP_Pubs-ImportJSON:

 ### ImportJSON by Trevor Lohrbeer (@FastFedora)

  Version:      1.1
  Project Page: http://blog.fastfedora.com/projects/import-json
  Copyright:    (c) 2012 by Trevor Lohrbeer
  License:      GNU General Public License, version 3 (GPL-3.0) 
                http://www.opensource.org/licenses/gpl-3.0.html
  ------------------------------------------------------------------------------------------------------------------------------------
  A library for importing JSON feeds into Google spreadsheets. Functions include:
     ImportJSON            For use by end users to import a JSON feed from a URL 
     ImportJSONAdvanced    For use by script developers to easily extend the functionality of this library
  Future enhancements may include:
   - Support for a real XPath like syntax similar to ImportXML for the query parameter
   - Support for OAuth authenticated APIs
  Or feel free to write these and add on to the library yourself!
  ------------------------------------------------------------------------------------------------------------------------------------
  Changelog:
  
  1.1    Added support for the noHeaders option
  1.0    Initial release
 *====================================================================================================================================*/
/**
 Imports a JSON feed and returns the results to be inserted into a Google Spreadsheet. The JSON feed is flattened to create 
 a two-dimensional array. The first row contains the headers, with each column header indicating the path to that data in 
 the JSON feed. The remaining rows contain the data. 
 
 By default, data gets transformed so it looks more like a normal data import. Specifically:
 
   - Data from parent JSON elements gets inherited to their child elements, so rows representing child elements contain the values 
      of the rows representing their parent elements.
   - Values longer than 256 characters get truncated.
   - Headers have slashes converted to spaces, common prefixes removed and the resulting text converted to title case. 
 
 To change this behavior, pass in one of these values in the options parameter:
 
     noInherit:     Don't inherit values from parent elements
     noTruncate:    Don't truncate values
     rawHeaders:    Don't prettify headers
     noHeaders:     Don't include headers, only the data
     debugLocation: Prepend each value with the row & column it belongs in
 
  For example:
 
    =ImportJSON("http://gdata.youtube.com/feeds/api/standardfeeds/most_popular?v=2&alt=json", "/feed/entry/title,/feed/entry/content",
                "noInherit,noTruncate,rawHeaders")
  
  @param {url} the URL to a public JSON feed
  @param {query} a comma-separated lists of paths to import. Any path starting with one of these paths gets imported.
  @param {options} a comma-separated list of options that alter processing of the data
 
  @return a two-dimensional array containing the data, with the first row containing headers
  @customfunction
 **/
 /**
  An advanced version of ImportJSON designed to be easily extended by a script. This version cannot be called from within a 
  spreadsheet.
 
  Imports a JSON feed and returns the results to be inserted into a Google Spreadsheet. The JSON feed is flattened to create 
  a two-dimensional array. The first row contains the headers, with each column header indicating the path to that data in 
  the JSON feed. The remaining rows contain the data. 
 
  Use the include and transformation functions to determine what to include in the import and how to transform the data after it is
  imported. 
 
  For example:
 
    =ImportJSON("http://gdata.youtube.com/feeds/api/standardfeeds/most_popular?v=2&alt=json", 
                "/feed/entry",
                 function (query, path) { return path.indexOf(query) == 0; },
                 function (data, row, column) { data[row][column] = data[row][column].toString().substr(0, 100); } )
 
  In this example, the import function checks to see if the path to the data being imported starts with the query. The transform 
  function takes the data and truncates it. For more robust versions of these functions, see the internal code of this library.
 
  @param {url}           the URL to a public JSON feed
  @param {query}         the query passed to the include function
  @param {options}       a comma-separated list of options that may alter processing of the data
  @param {includeFunc}   a function with the signature func(query, path, options) that returns true if the data element at the given path
                         should be included or false otherwise. 
  @param {transformFunc} a function with the signature func(data, row, column, options) where data is a 2-dimensional array of the data 
                         and row & column are the current row and column being processed. Any return value is ignored. Note that row 0 
                         contains the headers for the data, so test for row==0 to process headers only.
 
  @return a two-dimensional array containing the data, with the first row containing headers
 **/
 - **readRows**
 - **onOpen**, Function that adds Import JSON menu
 - **ImportJSON**
 - **ImportJSONAdvanced**
 - **parseJSONObject_**
 - **parseData_**
 - **parseHeaders_**
 - **transformData_**
 - **parseHeaders_**
 - **isObject_**
 - **isObjectArray_**
 - **includeXPath_**
 - **defaultTransform_**
 - **removeCommonPrefixes_**
 - **toTitleCase_**
 
 ## MRP_Pubs-ExportJSON:

Includes functions for exporting active sheet or all sheets as JSON object (also Python object syntax compatible).
Tweak the makePrettyJSON_ function to customize what kind of JSON to export.

var FORMAT_ONELINE   = 'One-line';
var FORMAT_MULTILINE = 'Multi-line';
var FORMAT_PRETTY    = 'Pretty';

var LANGUAGE_JS      = 'JavaScript';
var LANGUAGE_PYTHON  = 'Python';

var STRUCTURE_LIST = 'List';
var STRUCTURE_HASH = 'Hash (keyed by "id" column)';

/* Defaults for this particular spreadsheet, change as desired */
var DEFAULT_FORMAT = FORMAT_PRETTY;
var DEFAULT_LANGUAGE = LANGUAGE_JS;
var DEFAULT_STRUCTURE = STRUCTURE_LIST;

- **exportOptions**
- **makeLabel**
- **makeListBox**
- **makeButton**
- **makeTextBox**
- **exportAllSheets**
- **makeJSON_**
- **exportSheet**
- **getExportOptions**
- **makeJSON_**
- **displayText_**
- **getRowsData_**
- **getColumnsData_**
- **getObjects_**
- **normalizeHeaders_**
- **isCellEmpty_**

## MRP_Pubs-ExportCSV:
### ExportCSV by Michael Derazon
 #### - Script to export data in all sheets in the current spreadsheet as individual csv files
 #### - Files will be named according to the name of the sheet
 #### - Author: Michael Derazon https://gist.github.com/mderazon/9655893
 
 - **saveAsCSV**
 - **testMoveFolder**
 - **moveFolderToFolder**
 - **convertRangeToCsvFile_**
 
## MRP_Pubs-OnOpen:
Google sheet functions that are run when the Google sheet is opened
- **OnOpen**, Function that builds menu items for the worksheet

