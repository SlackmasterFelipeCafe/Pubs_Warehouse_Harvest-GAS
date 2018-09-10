# Pubs_Warehouse_Harvest-GAS-GAS
GGGSC SDC HarvestGAS - Google Application Scripts being used by the GGGSC for tracking harvesting new pubblication records for various science centers daily.  These records can then be used to track science center publications and how they are tagged in the USGS CMS

US Geological Survey (USGS)

Geology, Geophysics, and Geochemistry Science Center (GGGSC)

Data Managment Team (gs_gggsc_dm_team@usgs.gov)

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

- **GMEGtransposePubsForTagging**, function that culls harvest results, transposes and displays for user tag tracking interaction on a seperate Google Sheet worksheet

## isOdd:

Generic JSON Parser not being used by this project but needs to be tested more.  Functions include:

- **getJSON**
