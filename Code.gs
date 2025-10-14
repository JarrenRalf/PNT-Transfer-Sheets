/**
 * This function handles all of the on edit events of the spreadsheet, specifically looking for rows that need to be moved to different sheets,
 * barcodes that are scanned on the Item Scan sheet, searches that are made, and formatting issues that need to be fixed.
 * 
 * @param {Event Object} e : An instance of an event object that occurs when the spreadsheet is editted
 * @author Jarren Ralf
 */
function installedOnEdit(e)
{
  var spreadsheet = e.source;
  var sheet = spreadsheet.getActiveSheet(); // The active sheet that the onEdit event is occuring on
  var sheetName = sheet.getSheetName();

  try
  {
    switch (sheetName)
    {
      case "Shipped":
        receiveAll(e, spreadsheet, sheet); // Check if a user is trying to receive all of the items from one particular courrier
      case "Order":
      case "ItemsToRichmond":
      case "Received":
        itemScan(e, spreadsheet, sheet, sheetName); // Check if a barcode has been scanned
        moveRow(e, spreadsheet, sheet, sheetName);  // Check if the user is trying to move or add a row
        break;
      case "Item Search":
        search(e, spreadsheet, sheet); // Check if the user is searching for an item or trying to marry, unmarry or add a new item to the upc database
        break;
      case "Manual Counts":
      case "InfoCounts":
        warning(e, spreadsheet, sheet, sheetName); // Check if the user typed in the quantity in the wrong column
        break;
      case "Manual Scan":
      case "Manual Scan2":
        manualScan(e, spreadsheet, sheet) // Check if a barcode has been scanned
        break;
    }
  } 
  catch (err) 
  {
    var error = err['stack'];
    Logger.log(error)
    Browser.msgBox(error)
    throw new Error(error);
  }
}

/**
 * This places an add-on menu on the Parksville and Rupert spreadsheets.
 * 
 * @param {Event} e : The event object 
 */
function onOpen(e)
{
  if (e.source.getName().split(" ")[1] !== 'Richmond')
  {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Carrier Not Assigned')
      .addItem('Insert carrier not assigned banner', 'insertCarrierNotAssignedBanner')
      .addItem('Move selected items', 'moveSelectedItemsFromCarrierNotAssigned').addToUi();
    ui.createMenu('Email Trites')
      .addItem('Ask trites if they have stock', 'sendEmailToTrites').addToUi();
    ui.createMenu('PO')
      .addItem('Add items to Export 1', 'addToPurchaseOrderSpreadsheet_Export_1')
      .addItem('Add items to Export 2', 'addToPurchaseOrderSpreadsheet_Export_2').addToUi();
  }
}

/**
 * This function is run when an html web app is launched. In our case, when the modal dialog box is produced at 
 * the point a user has clicked the Download inFlow Pick List button inorder to produce the csv file.
 * 
 * @param {Event} e : The event object 
 * @return Returns the Html Output for the web app.
 */
function doGet(e)
{
  if (e.parameter['inFlowImport'] !== undefined) // The request parameter
  {
    const inFlowImportType = e.parameter['inFlowImport'];

    if (inFlowImportType === 'SalesOrder')
      return downloadInflowPickList()
    else if (inFlowImportType === 'StockLevels')
      return downloadInflowStockLevels()
    else if (inFlowImportType === 'Barcodes')
      return downloadInflowBarcodes()
  }
}

const inflow_conversions = {
  'WEB: 210/60x3-1/4"X100md X200FM Body #21 -  - Twisted Tarred Nylon - FOOT - 10010021FT': 1200,
  'WEB: 210/27x1-1/8"x200MDx105FMx235# -  - Twisted Tarred Nylon - POUND - 10100027': 235, 
  'WEB: 210/27x1-1/8"x100MDx 105FMS -  - Twisted Tarred Nylon - POUND - 101021027118': 226, 
  'WEB: 210/96 (6x16) x3"x100MDx50FMx230lbs -  - Cargo/Barrier - POUND - 10110096': 230, 
  'WEB: 210/224x3"x100MDxfoot ) #14x16 -  - Braided Tarred Nylon - FOOT - 10120495FOOT': 150, 
  'WEB: 210/96x3-5/8"x25MDx100 FMS Braid k -  - Braided Tarred Nylon - POUND - 10210096': 96,
  'WEB: 210/27x 2"x400MDx foot GOLF -  - Golf - FOOT - 10500027FT': 600, 
  'WEB: 210/144x3"x100MDx50Fx350#9x16 -  - Braided Tarred Nylon - POUND - 10120144': 330,
  'WEB: 210/15x1-1/4"x200MD x foot -  - Twisted Tarred Nylon - FOOT - 10000015FT': 35.7,
  '3MM Braided Knotted PE 4"x100MD 50F 285# -  - Soccer - POUND - 10500030': 285,
  'WEB: 210/20x1/2"x800MDx100FM RASCHEL -  - Raschel Knotless - POUND - 10722000': 255,
  'WEB: BRD KNOT 210/224 X 2-1/8" X foot -  - Web - Miscellaneous - FOOT - 10501416FT': 1, ////////////////////// 
  'WEB:#30 x 2" x 150MD BLACK HD ACRYLIC -  - Golf - POUND - 10503150': 1, //////////////////////
  'WEB: 210/128x2"x50MDx100FMx250LBS - North Pacific - Hockey/Lacrosse - POUND - 10500128': 250,
  'Black Cod Web 210/144 x 3in x 28md x 200 -  - Web - Miscellaneous - POUND - 10500144': 375, 
  'WEB: #36 x 3"x34MD BROWN HD ACRYLIC -  - Golf - POUND - 10500360': 1,  //////////////////////
  'WEB: #36 x 3" x 102md BROWN HD ACRYLIC -  - Golf - POUND - 10500361' : 1,  //////////////////////
  'WEB: PNT BLACKBIRD 15mm Sq x 2m deep -  - Golf - FOOT - 10501001FT': 328.084, 
  'WEB: #30 x 2"x50MD BLACK HD ACRYLIC COAT -  - Golf - POUND - 10503000': 300, 
  "VEXAR L36 WEB for CRAB CAGE  (100'/ROLL) -  - Golf - FOOT - 10503600": 100,
  'WEB: 210/10x1/2"x800MDx100FMx235# RACHL -  - Raschel Knotless - FOOT - 10710010FT': 600, 
  'Rachel Black We 210/9 X 3/8" X465MDX 900 -  - Raschel Knotless - POUND - 10782109038': 235,
  "BLACK RUBBER MATTING RIBBED    3 ' WIDE - ERIKS - Mats & Tables - FOOT - 24400000": 225,
  'MN3-7 Vexar Oyster Tube Netting Black -  - Golf - FOOT - 10501000': 1000,
  'MN8-25 Vexar Oyster Tube Netting Red -  - Golf - FOOT - 10501002': 2000,
  'NORPAC Dura Leadline 300 lbs/ 100 fm - NOVA BRAID - Gillnet Leadline - FATHOM - 17300008FM': 1, //////////////////////
  'ROPE, QUIK SPLICE POLYTRON 1/2" -  - Seine Beachline 12 Strand Quick Splice - FOOT - 18150001': 600,
  'GRADE 43 HIGH TEST GALV CHAIN 1/4" -  - Chain - FOOT - 26014025': 500,
  'GRADE 43 HIGH TEST GALV CHAIN 5/16" -  - Chain - FOOT - 26014516': 1,
  'CHAIN: PROOF COIL 1/4" Hot Dipped Galv - VANGUARD - Chain - FOOT - 21000001': 500,
  'CHAIN: PROOF COIL 3/8" Hot Dipped Galv - VANGUARD - Chain - FOOT - 21000003': 400, 
  'CHAIN: PROOF COIL 1/2" Hot Dipped Galv - VANGUARD - Chain - FOOT - 21000004': 200,
  'CHAIN: PROOF COIL 5/16" Hot Dipped Galv - VANGUARD - Chain - FOOT - 21000002': 1, //////////////////////////////
  'CHAIN: PROOF COIL 5/8" Hot Dipped Galv - VANGUARD - Chain - FOOT - 21000005' : 150
}

/**
 * This function allows Adrian to select items on the INVENTORY page and move them to the SearchData page which will cause 
 * them to now be available for search on Item Search sheet. This function effectively removes "No TS" for the current day.
 * 
 * @author Jarren Ralf
 */
function addItemsToSearchData()
{
  const NUM_COLS = 6;
  const spreadsheet = SpreadsheetApp.getActive();
  const searchDataSheet = spreadsheet.getSheetByName("SearchData");
  const startRow = searchDataSheet.getLastRow() + 1; // The bottom of the list
  var firstRows, row, lastRow, rows = [];
  [firstRows, numRows, row, lastRow] = copySelectedValues(searchDataSheet, startRow, NUM_COLS); // Move items to SearchData
  const totalNumRows = lastRow - row + 1;
  const numActiveRanges = firstRows.length;

  // Determine which rows "No TS" needs to be removed from
  for (var i = 0; i < numActiveRanges; i++)
    for (var j = 0; j < numRows[i]; j++)
      rows.push(firstRows[i] - row + j);

  var range = spreadsheet.getActiveSheet().getRange(row, 9, totalNumRows);
  var values = range.getValues();
  rows.map(row => values[row][0] = '');
  range.setValues(values);
}

/**
 * This function allows the user to select items on the Manual Counts page and move them to the UPC Database and Manually Added UPCs pages.
 * In turn, this will now allow the items to be searchable via a Manual Scan.
 * 
 * @author Jarren Ralf
 */
function addItemsToUpcData()
{
  const NUM_COLS = 4;
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();
  const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
  const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");

  if (sheet.getSheetName() === "Manual Scan")
  { 
    const ui = SpreadsheetApp.getUi();
    const barcodeInputRange = sheet.getRange(1, 1);
    const values = barcodeInputRange.getValue().split('\n');

    const response = ui.prompt('Manually Add UPCs', 'Please scan the barcode for:\n\n' + values[0] +'.', ui.ButtonSet.OK_CANCEL)
    {
      if (response.getSelectedButton() == ui.Button.OK)
      {
        const item = values[0].split(' - ');
        const upc = response.getResponseText();
        manAddedUPCsSheet.getRange(manAddedUPCsSheet.getLastRow() + 1, 1, 1, 5).setNumberFormat('@').setValues([[item[item.length - 1], upc, item[item.length - 2], values[0], '']]);
        upcDatabaseSheet.getRange(upcDatabaseSheet.getLastRow() + 1, 1, 1, NUM_COLS).setNumberFormat('@').setValues([[upc, item[item.length - 2], values[0], values[4]]]);
        barcodeInputRange.activate();
      }
    }
  }
  else if (sheet.getSheetName() === "Manual Counts" || sheet.getSheetName() === "Item Search")
  {
    const startRow = upcDatabaseSheet.getLastRow() + 1; // The bottom of the list
    copySelectedValues(upcDatabaseSheet, startRow, NUM_COLS, 'upc', false, [manAddedUPCsSheet]); // Move items to UPC Database and the Manually Added UPCs
    const row = sheet.getActiveRange().getRow();
    populateManualScan(spreadsheet, sheet, row);
  }
}

/**
 * This function allows the user to select items on the Manual Counts page and move them to the UPC Database and Manually Added UPCs pages.
 * In turn, this will now allow the items to be searchable via a Manual Scan. In this case, the item/s in question appear not to be found in the Adagio database.
 * 
 * @author Jarren Ralf
 */
function addItemsToUpcData_ItemsNotFound()
{
  const NUM_COLS = 4;
  const spreadsheet = SpreadsheetApp.getActive();
  const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
  const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");
  const inventorySheet = (isRichmondSpreadsheet(spreadsheet)) ? spreadsheet.getSheetByName("INVENTORY") : spreadsheet.getSheetByName("SearchData");
  const startRow = upcDatabaseSheet.getLastRow() + 1; // The bottom of the list

  copySelectedValues(upcDatabaseSheet, startRow, NUM_COLS, 'upc', false, [manAddedUPCsSheet, inventorySheet], true); // Move items to UPC Database, Manually Added UPCs, and INVENTORY sheets
  spreadsheet.getSheetByName("Manual Scan").getRange(1, 1).activate();
}

/**
 * This function adds a new item to the bottom of the inventory page as to make it now searchable on the Item Search sheet.
 * 
 * @author Jarren Ralf
 */
function addNewItem()
{
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Please type out the SKU *space*space* Description", ui.ButtonSet.OK_CANCEL)

  if (response.getSelectedButton() == ui.Button.OK)
  {
    const item = response.getResponseText().toUpperCase().split("  "); // Split at the double space mark "*space*space*"

    if (item[1] == undefined) // Only one word is typed into the text response box
      ui.alert("Invalid Input!", "Remember to type *space*space* inbetween the SKU and Description.",ui.ButtonSet.OK);
    else if (item[0].trim() == '') // Too many spaces at the front of the typed string
      ui.alert("Missing Data!", "The SKU is blank.",ui.ButtonSet.OK);
    else if (item[1].trim() == '') // Too many spaces at the end of the typed string
      ui.alert("Missing Data!", "The description is blank.",ui.ButtonSet.OK);
    else
    {
      const today = new Date();
      const createdDate = today.getDate() + '.' + (today.getMonth() + 1) + '.' + today.getFullYear();
      const spreadsheet = SpreadsheetApp.getActive();
      const sheet = spreadsheet.getActiveSheet();

      // Append the item to the bottom of list and take the user to the new item
      if (isRichmondSpreadsheet(spreadsheet))
        sheet.appendRow(["EACH", item[0] + ' - ' + item[1] + ' - - - EACH', '', '', '', '', '', item[0], createdDate]).getRange(sheet.getLastRow(), 8).activate();
      else
      {
        spreadsheet.getSheetByName("SearchData").appendRow(["EACH", item[0] + ' - ' + item[1] + ' - - - EACH', '', '', '', '', ''])
        sheet.appendRow(["EACH", item[0] + ' - ' + item[1] + ' - - - EACH', '', '', '', '', item[0], createdDate, '']).getRange(sheet.getLastRow(), 7).activate();
      }
    }
  }
}

/**
 * This function toggles between the Manual Scan's Add-One mode and regular scanning mode. The Add-One mode immediate updates the quantity of the scanned item by 1 
 * without specifying the quantity.
 * 
 * @author Jarren Ralf
 */
function addOneMode()
{
  const range = SpreadsheetApp.getActiveSheet().getRange(3, 7);

  if (range.isChecked())
  {
    range.uncheck();
    SpreadsheetApp.getActive().toast('Off','Add-One Mode');
  }
  else
  {
    range.check();
    SpreadsheetApp.getActive().toast('On','Add-One Mode');
  } 
}

/**
* This function moves all of the selected values on the item search page to the Manual Counts page on the Richmond, Parksville, and Prince Rupert spreadsheet.
*
* @author Jarren Ralf
*/
function addToAllManualCountsPages()
{
  const spreadsheets =  [SpreadsheetApp.openById('1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c'), // Rupert
                         SpreadsheetApp.openById('181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM'), // Parksville
                         SpreadsheetApp.getActive()] // Richmond
  const itemSearchSheet = spreadsheets[2].getSheetByName("Item Search")
  const activeRanges = itemSearchSheet.getActiveRangeList().getRanges(); // The selected ranges on the item search sheet
  const numActiveRanges = activeRanges.length;
  var itemValues = [[[]]], firstRows = [], lastRows = [], items, manualCountsSheet, finalRow, startRow, groupedItems;

  
  // Find the first row and last row in the the set of all active ranges
  for (var r = 0; r < numActiveRanges; r++)
  {
    firstRows[r] = activeRanges[r].getRow();
     lastRows[r] = activeRanges[r].getLastRow()
  }
  
  var     row = Math.min(...firstRows); // This is the smallest starting row number out of all active ranges
  var lastRow = Math.max( ...lastRows); // This is the largest     final row number out of all active ranges
  var finalDataRow = itemSearchSheet.getLastRow() + 1;
  var numHeaders = 3;

  if (row > numHeaders && lastRow <= finalDataRow) // If the user has not selected an item, alert them with an error message
  { 
    for (var r = 0; r < numActiveRanges; r++)
      itemValues[r] = itemSearchSheet.getSheetValues(firstRows[r], 2, lastRows[r] - firstRows[r] + 1, 5);
    
    var itemVals = [].concat.apply([], itemValues); // Concatenate all of the item values as a 2-D array
    var numItems = itemVals.length;

    const manualCountsSheets = spreadsheets.map((spreadsheet, store) => {

      switch (store)
      {
        case 0: // Rupert
          items = itemVals.map(u => [u[0], u[4]]);
          break;
        case 1: // Parkville
          items = itemVals.map(u => [u[0], u[3]]);
          break;
        case 2: // Richmond
          items = itemVals.map(u => [u[0], u[2]]);
          break;
      }

      manualCountsSheet = spreadsheet.getSheetByName("Manual Counts")
      finalRow = manualCountsSheet.getLastRow();
      startRow = (finalRow < 3) ? 4 : finalRow + 1;

      if (startRow > 4) // There are existing items on the Manual Counts page
      {
        // Retrieve the existing items and add them to the new items
        // groupedItems = groupHoochieTypes(manualCountsSheet.getSheetValues(4, 1, startRow - 4, manualCountsSheet.getMaxColumns()).concat(items.map(val => [...val, '', '', '', '', ''])), 0)
        // items.length = 0;

        // Object.keys(groupedItems).forEach(key => items.push(...sortHoochies(groupedItems[key], 0, key)));

        const items_ = manualCountsSheet.getSheetValues(4, 1, startRow - 4, manualCountsSheet.getMaxColumns()).concat(items.map(val => [...val, '', '', '', '', '']));

        manualCountsSheet.getRange(4, 1, items_.length, items_[0].length).setNumberFormat('@').setValues(items_); // Move the item values to the destination sheet
        applyFullRowFormatting(manualCountsSheet, 4, items_.length, 7); // Apply the proper formatting
      }
      else
      {
        // groupedItems = groupHoochieTypes(items, 0)
        // items.length = 0;

        // Object.keys(groupedItems).forEach(key => items.push(...sortHoochies(groupedItems[key], 0, key)));

        manualCountsSheet.getRange(startRow, 1, numItems, items[0].length).setNumberFormat('@').setValues(items); // Move the item values to the destination sheet
        applyFullRowFormatting(manualCountsSheet, startRow, numItems, 7); // Apply the proper formatting
      }

      switch (store)
      {
        case 0: // Rupert
          spreadsheet.toast('Items added to Prince Rupert Manual Counts')
          break;
        case 1: // Parkville
          spreadsheet.toast('Items added to Parksville Manual Counts')
          break;
      }

      return manualCountsSheet;
    })

    manualCountsSheets[2].getRange(startRow, 3).activate()
  }
  else
    SpreadsheetApp.getUi().alert('Please select an item from the list.');
}

/**
 * This function takes the user's selected items on the Item Search page of the Richmond spreadsheet and it places those items on the inFlowPick page.
 * 
 * @param {Number} qty : If an argument is passed to this function, it is the quantity that a user is entering on the Order page for the inFlow pick list
 * @author Jarren Ralf
 */
function addToInflowPickList(qty)
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = (!isRichmondSpreadsheet(spreadsheet)) ? SpreadsheetApp.openById('1fSkuXdmLEjsGMWVSmaqbO_344VNBxTVjdXFL1y0lyHk').getSheetByName("inFlowPick") : 
                                                                                                                    spreadsheet.getSheetByName("inFlowPick");
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const activeRanges = activeSheet.getActiveRangeList().getRanges(); // The selected ranges on the item search sheet
  const numActiveRanges = activeRanges.length;
  const firstRows = [], lastRows = [], itemValues = [];

  const inflowData = Utilities.parseCsv(DriveApp.getFilesByName("inFlow_StockLevels.csv").next().getBlob().getDataAsString())
    .filter(item => item[0].split(' - ').length > 4).map(descrip => descrip[0])

  if (activeSheet.getSheetName() === "Item Search")
  {
    // Find the first row and last row in the the set of all active ranges
    for (var r = 0; r < numActiveRanges; r++)
    {
       firstRows[r] = activeRanges[r].getRow();
        lastRows[r] = activeRanges[r].getLastRow();
      itemValues[r] = activeSheet.getSheetValues(firstRows[r], 2, lastRows[r] - firstRows[r] + 1, 6);
    }
    
    const     row = Math.min(...firstRows); // This is the smallest starting row number out of all active ranges
    const lastRow = Math.max( ...lastRows); // This is the largest     final row number out of all active ranges
    const itemVals = [].concat.apply([], itemValues).map(item => ['newRichmondPick', 'Richmond PNT', inflowData.find(description => description === item[0]), '', item[5]])
                                                    .filter(itemNotFound => itemNotFound[2] != null)

    if (row > 3 && lastRow <= activeSheet.getLastRow())
    {
      const numItems = itemVals.length;

      if (numItems !== 0)
        sheet.getRange(sheet.getLastRow() + 1, 1, numItems, 5).setValues(itemVals).offset(0, 3, numItems, 1).activate()
      else
        SpreadsheetApp.getUi().alert('Your current selection(s) can\'t be placed on an inFlow picklist due to ambiguity of the Adagio description(s).');
    }
    else
      SpreadsheetApp.getUi().alert('Please select an item from the list.');
  }
  else if (activeSheet.getSheetName() === "Suggested inFlowPick")
  {
    // Find the first row and last row in the the set of all active ranges
    for (var r = 0; r < numActiveRanges; r++)
    {
       firstRows[r] = activeRanges[r].getRow();
        lastRows[r] = activeRanges[r].getLastRow();
      itemValues[r] = activeSheet.getSheetValues(firstRows[r], 1, lastRows[r] - firstRows[r] + 1, 3);
    }
    
    const     row = Math.min(...firstRows); // This is the smallest starting row number out of all active ranges
    const lastRow = Math.max( ...lastRows); // This is the largest     final row number out of all active ranges
    const itemVals = [].concat.apply([], itemValues).map(item => ['newSuggestedPick', 'Richmond PNT', inflowData.find(description => description === item[2]), item[0], item[2]])
                                                    .filter(itemNotFound => itemNotFound[2] != null)

    if (row > 1 && lastRow <= activeSheet.getLastRow())
    {
      const numItems = itemVals.length;

      if (numItems !== 0)
        sheet.getRange(sheet.getLastRow() + 1, 1, numItems, 5).setValues(itemVals).offset(0, 3, numItems, 1).activate()
      else
        SpreadsheetApp.getUi().alert('Your current selection(s) can\'t be placed on an inFlow picklist due to ambiguity of the Adagio description(s).');
    }
    else
      SpreadsheetApp.getUi().alert('Please select an item from the list.');
  }
  else if (activeSheet.getSheetName() === "Order")
  {
    // Find the first row and last row in the the set of all active ranges
    for (var r = 0; r < numActiveRanges; r++)
    {
       firstRows[r] = activeRanges[r].getRow();
      itemValues[r] = activeSheet.getSheetValues(firstRows[r], 3, activeRanges[r].getLastRow() - firstRows[r] + 1, 7);
    }

    if (isParksvilleSpreadsheet(spreadsheet))
    {
      var inFlowOrderNumber = 'newParksvillePick';
      var inFlowCustomerName = 'Parksville PNT';
    }
    else
    {
      var inFlowOrderNumber = 'newRupertPick';
      var inFlowCustomerName = 'Rupert PNT';
    }

    const row = Math.min(...firstRows); // This is the smallest starting row number out of all active ranges
    const itemVals = [].concat.apply([], itemValues).map(item => [inFlowOrderNumber, inFlowCustomerName, 
                                                    inflowData.find(description => description === item[2]), (qty) ? qty : item[0], item[3].split('): ')[1]])
                                                    .filter(itemNotFound => itemNotFound[2] != null)
    
    if (row > 3)
    {
      const numItems = itemVals.length;

      if (numItems !== 0)
      {
        sheet.getRange(sheet.getLastRow() + 1, 1, numItems, 5).setValues(itemVals).offset(0, 3, numItems, 1).activate()
        spreadsheet.toast('Item(s) added to inFlow Pick List on the Richmond sheet')
      }
      else
        SpreadsheetApp.getUi().alert('Your current selection(s) can\'t be placed on an inFlow picklist due to ambiguity of the Adagio description(s).');
    }
    else
      SpreadsheetApp.getUi().alert('Please select an item from the list.');
  }
}

/**
 * This function moves selected items from the ItemsToRichmond sheet to the opposite stores Shipped page. This is in the event when we are shipping items from 
 * Parksville to Rupert (or vice versa), but they are shipping to Moncton street first.
 * 
 * @author Jarren Ralf
 */
function addToOppositeStoreShippedPage()
{
  const spreadsheet = SpreadsheetApp.getActive()
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const activeRanges = activeSheet.getActiveRangeList().getRanges(); // The selected ranges on the item search sheet
  const numActiveRanges = activeRanges.length;
  const firstRows = [], numRows = [], itemValues = [], backgroundColours = [], richTextValues = [];

  // Find the first row and last row in the the set of all active ranges
  for (var r = 0; r < numActiveRanges; r++)
  {
    firstRows.push(activeRanges[r].getRow());
    numRows.push(activeRanges[r].getLastRow() - firstRows[r] + 1);
    itemValues.push(activeSheet.getSheetValues(firstRows[r], 1, numRows[r], 6));
    backgroundColours.push(...activeSheet.getRange(firstRows[r], 1, numRows[r], 5).getBackgrounds().map(u => [u[0], 'white',  'white', 'white', 'white', u[4]]));
    richTextValues.push(...activeSheet.getRange(firstRows[r], 5, numRows[r]).getRichTextValues())
  }

  const row = Math.min(...firstRows); // This is the smallest starting row number out of all active ranges
  
  if (row > 3)
  {
    const timeZone = spreadsheet.getSpreadsheetTimeZone();
    const shippedDate = Utilities.formatDate(new Date(), timeZone, 'dd MMM yyyy');
    const itemVals = [].concat.apply([], itemValues).map(item => 
      [Utilities.formatDate(item[0], timeZone, 'dd MMM yyyy'), item[1], item[5], item[2], item[3], item[4], '', '', item[5], 'Carrier Not Assigned', shippedDate]
    )

    if (isParksvilleSpreadsheet(spreadsheet))
    {
      var targetSpreadsheet = SpreadsheetApp.openById('1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c');
      var branch_location = 'Rupert'
    }
    else
    {
      var targetSpreadsheet = SpreadsheetApp.openById('181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM');
      var branch_location = 'Parksville'
    }

    const shippedPage = targetSpreadsheet.getSheetByName("Shipped")
    const row = shippedPage.getLastRow() + 1;
    const numRows = itemVals.length;
    const numCols = 11;
    const range = shippedPage.getRange(row, 1, numRows, numCols);
    range.setValues(itemVals)
    applyFullRowFormatting(shippedPage, row, numRows, numCols)
    range.offset(0, 0, numRows, 6).setBackgrounds(backgroundColours).offset(0, 5, numRows, 1).setRichTextValues(richTextValues)
    spreadsheet.toast('Item(s) have been added to the ' + branch_location + ' Shipped page.')
  }
  else
    SpreadsheetApp.getUi().alert('Please select an item from the list.');
}

/**
 * This function take the items that are selected by the user and it adds them to the Purchase Order spreadsheet.
 * 
 * @param {Number} exportNum : The number referring to which export sheet to add the purchase order items to,
 * @author Jarren Ralf
 */
function addToPurchaseOrderSpreadsheet(exportNum)
{
  const activeRanges = SpreadsheetApp.getActiveRangeList().getRanges(); // The selected ranges on the item search sheet

  if (SpreadsheetApp.getActiveSheet().getSheetName() === "Order" && Math.min(...activeRanges.map(rng => rng.getRow())) > 3) // If the user has not selected an item, alert them with an error message
  { 
    const exportData = [].concat.apply([], activeRanges.map(rng => rng.offset(0, 3 - rng.getColumn(), rng.getNumRows(), 3).getValues())).map(row => ['R', row[2].split(' - ')[row[2].split(' - ').length - 1], row[0], row[2]]);
    const exportSheet = SpreadsheetApp.openById('1Fx_0LCt8_j1VeM6w0vCup_1GXjrHT5gBfUaojzvP46Y').getSheetByName('Export ' + exportNum);
    exportSheet.getRange(exportSheet.getLastRow() + 1, 1, exportData.length, 4).setNumberFormat('@').setHorizontalAlignment('left')
      .setFontColor('black').setFontFamily('Arial').setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('bold').setVerticalAlignment('middle').setBackground(null).setValues(exportData)

    SpreadsheetApp.getActive().toast('The selected items were added to the Purchase Order spreadsheet on the Export ' + exportNum + ' tab.')
  }
  else
    Browser.msgBox('Please select an item or items on the Order page.')
}

/**
 * This function take the items that are selected by the user and it adds them to the Purchase Order spreadsheet on the Export 1 page.
 * 
 * @author Jarren Ralf
 */
function addToPurchaseOrderSpreadsheet_Export_1()
{
  addToPurchaseOrderSpreadsheet(1)
}

/**
 * This function take the items that are selected by the user and it adds them to the Purchase Order spreadsheet on the Export 2 page.
 * 
 * @author Jarren Ralf
 */
function addToPurchaseOrderSpreadsheet_Export_2()
{
  addToPurchaseOrderSpreadsheet(2)
}

/**
 * Apply the proper formatting to the Order, Shipped, Received, ItemsToRichmond, Manual Counts, or InfoCounts page.
 *
 * @param {Sheet}   sheet  : The current sheet that needs a formatting adjustment
 * @param {Number}   row   : The row that needs formating
 * @param {Number} numRows : The number of rows that needs formatting
 * @param {Number} numCols : The number of columns that needs formatting
 * @author Jarren Ralf
 */
function applyFullRowFormatting(sheet, row, numRows, numCols)
{
  const SHEET_NAME = sheet.getSheetName();

  if (SHEET_NAME === "InfoCounts")
  {
    var numberFormats = [...Array(numRows)].map(e => ['@', '#.#', '0.#']);
    sheet.getRange(row, 1, numRows, numCols).setBorder(null, true, false, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK).setNumberFormats(numberFormats);
    sheet.getRange(row, 3, numRows         ).setBorder(null, true, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
                                            .setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  }
  else if (SHEET_NAME === "Manual Counts")
  {
    var numberFormats = [...Array(numRows)].map(e => ['@', '#.#', '0.#', '@', '#', '@', '@']);
    sheet.getRange(row, 1, numRows, numCols).setBorder(null, true, false, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK).setNumberFormats(numberFormats);
    sheet.getRange(row, 3, numRows         ).setBorder(null, true, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
                                            .setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
    sheet.getRange(row, 5, numRows,       2).setBorder(null, true, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID) 
                                            .setBorder(null, null, null, null, true, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK)
                                            .setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID)
  }
  else if (SHEET_NAME === "Trites Counts")
  {
    var numberFormats = [...Array(numRows)].map(e => ['@', '#.#', '#.#']);
    sheet.getRange(row, 1, numRows, 3).setBorder(null, true, false, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK).setNumberFormats(numberFormats);
  }
  else
  {
    const   BLUE = '#c9daf8', GREEN = '#d9ead3', YELLOW = '#fff2cc';
    const isItemsToRichmondPage = (SHEET_NAME === "ItemsToRichmond") ? true : false;

    if (isItemsToRichmondPage)
    {
      var      borderRng = sheet.getRange(row, 1, numRows, 8);
      var  shippedColRng = sheet.getRange(row, 6, numRows   );
      var thickBorderRng = sheet.getRange(row, 6, numRows, 3);
      var numberFormats = [...Array(numRows)].map(e => ['dd MMM yyyy', '@', '@', '@', '@', '#.#', '@', '@']);
      var horizontalAlignments = [...Array(numRows)].map(e => ['right', 'center', 'center', 'left', 'center', 'center', 'center', 'left']);
      var wrapStrategies = [...Array(numRows)].map(e => [...new Array(2).fill(SpreadsheetApp.WrapStrategy.OVERFLOW), ...new Array(3).fill(SpreadsheetApp.WrapStrategy.WRAP), 
          SpreadsheetApp.WrapStrategy.CLIP, SpreadsheetApp.WrapStrategy.WRAP, SpreadsheetApp.WrapStrategy.WRAP]);
    }
    else
    {
      var         borderRng = sheet.getRange(row, 1, numRows, numCols);
      var     shippedColRng = sheet.getRange(row, 9, numRows         );
      var    thickBorderRng = sheet.getRange(row, 9, numRows,       2);
      var numberFormats = [...Array(numRows)].map(e => ['dd MMM yyyy', '@', '#.#', '@', '@', '@', '#.#', '0.#', '#.#', '@', 'dd MMM yyyy']);
      var horizontalAlignments = [...Array(numRows)].map(e => ['right', ...new Array(3).fill('center'), 'left', ...new Array(6).fill('center')]);
      var wrapStrategies = [...Array(numRows)].map(e => [...new Array(3).fill(SpreadsheetApp.WrapStrategy.OVERFLOW), ...new Array(3).fill(SpreadsheetApp.WrapStrategy.WRAP), ...new Array(3).fill(SpreadsheetApp.WrapStrategy.CLIP), SpreadsheetApp.WrapStrategy.WRAP, SpreadsheetApp.WrapStrategy.CLIP]);
    }

    borderRng.setFontSize(10).setFontLine('none').setFontWeight('bold').setFontStyle('normal').setFontFamily('Arial').setFontColor('black')
                  .setNumberFormats(numberFormats).setHorizontalAlignments(horizontalAlignments).setWrapStrategies(wrapStrategies)
                  .setBorder(true, true, true, true,  false, true, 'black', SpreadsheetApp.BorderStyle.SOLID).setBackground('white');

    thickBorderRng.setBorder(null, true, null, true, false, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK).setBackground(GREEN);
    shippedColRng.setBackground(YELLOW);

     if (!isItemsToRichmondPage)
       sheet.getRange(row, 7, numRows, 2).setBorder(null,  true,  null,  null,  true,  null, 'black', SpreadsheetApp.BorderStyle.SOLID).setBackground(BLUE);
  }
}

/**
 * This function sets the formatting across every sheet in this spreadsheet if called with null parameters. When this function is called by the 
 * formatActiveSheet function (see below), then a singular sheet is formatted.
 * 
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @param   {Sheet[]}      sheets   : The active sheet in an array.
 * @author Jarren Ralf
 */
function applyFullSpreadsheetFormatting(spreadsheet, sheets)
{
  if (arguments.length === 0) // If no arguments are sent to the 
  {
    spreadsheet = SpreadsheetApp.getActive();
    sheets = spreadsheet.getSheets();
  }

  const Store_Name = spreadsheet.getName().split(" ")[1]; // Gets the store name from the name of the spreadsheet
  const STORE_NAME = Store_Name.toUpperCase();            // Makes the store name upper case
  const sheetNames = sheets.map(sheet => sheet.getSheetName());
  const numSheets = sheets.length;
  const RED = '#ea9999', GREEN = '#b6d7a8', YELLOW = '#ffd666'; // The colours of the order date highlighting
  const today = new Date();
  const      YEAR = today.getFullYear();
  const     MONTH = today.getMonth()
  const       DAY = today.getDate();
  const ONE_WEEK  = new Date(YEAR, MONTH, DAY -  7); // Used to highlight the order dates correctly
  const ONE_MONTH = new Date(YEAR, MONTH, DAY - 31);
  var numHeaders, rowStart, maxRow, lastRow, lastCol, numRows, dataRange, dataValues, descriptionWithHyperlinkRange, descriptionWithHyperlink, fontSizes, fontColours, 
    numberFormats, backgroundColours, horizontalAlignments, wrapStrategies, notesRange, noteBackgroundColours, richTextValues, headerNumberFormats, headerValues, 
    headerBackgroundColours, headerFontColours, headerFontWeights, headerFontSizes, headerHorizontalAlignments, headerFonts, columnWidths;

  for (var j = 0; j < numSheets; j++)
  {
    if(sheetNames[j] === "Order" || sheetNames[j] === "Shipped" || sheetNames[j] === "Received" )
    {
      numHeaders = 3;
      rowStart = numHeaders + 1;
      lastRow = sheets[j].getLastRow();
      lastCol = sheets[j].getMaxColumns();
      numRows = lastRow - numHeaders;
      headerRange = sheets[j].getRange(       1, 1, numHeaders, 10);
        dataRange = sheets[j].getRange(rowStart, 1, numRows,    11);
       dataValues = dataRange.getValues();

      // Set the column widths and the header's row heights
      sheets[j].setRowHeights(1, 2, 65).setRowHeightsForced(3, 1, 2).setFrozenRows(2);
      sheets[j].hideRows(3);
      for (var c = 0; c < lastCol; c++)
        sheets[j].setColumnWidth(c + 1, [90, 100, 50, 75, 650, 250, 40, 40, 75, 180, 125, 25, 50][c]);

      headerValues = [['Scan Here','','','','','','Current Stock','Actual Count','',''],
                      ['Order Date','Entered By:','Qty','UoM','Description','Notes','','','Shipped','Shipment Status'], 
                      ['', '', '', '', '', '', '', '', '', '']];
      headerBackgroundColours = [ '', [...new Array(8).fill('white'), '#fff2cc', '#d9ead3'], new Array(10).fill('white')];
      headerFontWeights = [['normal', ...new Array(9).fill('bold')], new Array(10).fill('bold'), new Array(10).fill('bold')]

      if (sheetNames[j] === "Order")
      {
        headerValues[0][3] = 'ITEMS ORDERED BY PNT ' + STORE_NAME;
        headerBackgroundColours[0] = ['white', ...new Array(9).fill('#5b95f9')];
        headerFontColours = [ ['black', ...new Array(9).fill('white')],  new Array(10).fill('black'), new Array(10).fill('black')];
        descriptionWithHyperlinkRange = sheets[j].getRange(1, 9);
        descriptionWithHyperlink = descriptionWithHyperlinkRange.getRichTextValue();
      }
      else if (sheetNames[j] === "Shipped")
      {
        headerValues[0][3] = 'SHIPPED ITEMS IN TRANSIT TO PNT ' + STORE_NAME;
        headerBackgroundColours[0] = ['white', ...new Array(9).fill('#ffd666')];
        headerFontColours = [...Array(numHeaders)].map(e => new Array(10).fill('black'));
      }
      else if (sheetNames[j] === "Received")
      {
        headerValues[0][3] = 'ITEMS RECEIVED INTO PNT ' + STORE_NAME;
        headerBackgroundColours[0] = ['white', ...new Array(9).fill('#8bc34a')];
        headerFontColours = [...Array(numHeaders)].map(e => new Array(10).fill('black'));
        descriptionWithHyperlinkRange = sheets[j].getRange(2, 5);
        descriptionWithHyperlink = descriptionWithHyperlinkRange.getRichTextValue();
      }

      // Prepare and set all of the headerRange values
      headerFontSizes = [[35, 10, 10, 27, ...new Array(6).fill(10)], [...new Array(8).fill(14), ...new Array(2).fill(12)], new Array(10).fill(10)];
          headerFonts = [['Libre Barcode 128', ...new Array(9).fill('Verdana')], new Array(10).fill('Arial'), new Array(10).fill('Arial')];
      headerRange.setWrap(true).setNumberFormat('@').setBackgrounds(headerBackgroundColours)
        .setFontLine('none').setFontWeights(headerFontWeights).setFontStyle('normal').setFontFamilies(headerFonts).setFontSizes(headerFontSizes).setFontColors(headerFontColours)
        .setVerticalAlignment('middle').setHorizontalAlignment('center').setValues(headerValues);

      if (sheetNames[j] === "Received")
      {
        descriptionWithHyperlinkRange.setRichTextValue(descriptionWithHyperlink);
        sheets[j].getRange(3, 10).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Back to Shipped']).build());
        sheets[j].getRange(3, 12).insertCheckboxes().check();
      }
      else
      {
        if (sheetNames[j] === "Order")
        {
          var col = 'B';
          descriptionWithHyperlinkRange.setRichTextValue(descriptionWithHyperlink.copy().setTextStyle(SpreadsheetApp.newTextStyle().setForegroundColor('white').build()).build()); // White
        }
        else
        {
          var col = 'D';
          sheets[j].getRange(3, 13).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Receive ALL']).build());
        }

        var dataValidationSheet = (sheets.length === 1) ? spreadsheet.getSheetByName("Data Validation") : sheets[sheetNames.indexOf("Data Validation")];
        sheets[j].getRange(3, 10).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(dataValidationSheet.getRange('$' + col + '$1:$' + col)).build());
      }

      // Prepare all of the dataRange values and formats
      fontSizes            = [...Array(numRows)].map(e => new Array(11).fill(10));
      fontColours          = [...Array(numRows)].map(e => new Array(11).fill('black'));
      numberFormats        = [...Array(numRows)].map(e => ['dd MMM yyyy', '@', '#.#', '@', '@', '@', '#.#', '0.#', '#.#', '@', 'dd MMM yyyy']);
      backgroundColours    = [...Array(numRows)].map(e => [null, 'white', 'white', 'white', 'white', null, '#c9daf8', '#c9daf8', '#fff2cc', '#d9ead3', 'white']);
      horizontalAlignments = [...Array(numRows)].map(e => ['right', 'center', 'center', 'center', 'left', 'center', 'center', 'center', 'center', 'center', 'center']);
      wrapStrategies       = [...Array(numRows)].map(e => [...new Array(3).fill(SpreadsheetApp.WrapStrategy.OVERFLOW), ...new Array(3).fill(SpreadsheetApp.WrapStrategy.WRAP), 
                                   ...new Array(3).fill(SpreadsheetApp.WrapStrategy.CLIP), SpreadsheetApp.WrapStrategy.WRAP, SpreadsheetApp.WrapStrategy.CLIP]);
      notesRange = sheets[j].getRange(rowStart, 6, numRows); // To preserve the background and text colours of the Notes, we must store that data first
      noteBackgroundColours = notesRange.getBackgrounds();
      richTextValues = notesRange.getRichTextValues();

      if (sheetNames[j] === "Shipped")
      {
        for (var i = 0; i < numRows; i++)
        {
          backgroundColours[i][0] = (dataValues[i][0] >= ONE_WEEK) ? GREEN : ( (dataValues[i][0] >= ONE_MONTH) ? YELLOW : RED ); // Highlight the dates correctly

          if (dataValues[i][10] === "via") // Locate the shipping carrier banner and apply the appropriate changes
          {
            fontSizes[i] = new Array(11).fill(14);
            numberFormats[i] = new Array(11).fill('@');
            backgroundColours[i] = new Array(11).fill('#6d9eeb');
            horizontalAlignments[i] = new Array(11).fill('left');
            fontColours[i] = [...new Array(10).fill('white'), '#6d9eeb'];
            sheets[j].getRange(i + 4, 1, 1, 10).merge();
            sheets[j].setRowHeight(i + 4, 40).getRange(i + 4, 1, 1, 11).setBorder(true,true,true,true,false,false);
          }
        }

        sheets[j].getRange(3, 12).setFormula('=ArrayFormula(if(K3:K="via",A3:A,""))');
        sheets[j].hideColumns(12, 2);
      }
      else
      {
        for (var i = 0; i < numRows; i++)
          backgroundColours[i][0] = (dataValues[i][0] >= ONE_WEEK) ? GREEN : ( (dataValues[i][0] >= ONE_MONTH) ? YELLOW : RED ); // Highlight the dates correctly
      }

      // Set all of the dataRange values and formats
      dataRange.setFontLine('none').setFontStyle('normal').setFontWeight('bold').setFontFamily('Arial').setFontSizes(fontSizes).setFontColors(fontColours)
        .setHorizontalAlignments(horizontalAlignments).setVerticalAlignment('middle').setWrapStrategies(wrapStrategies)
        .setNumberFormats(numberFormats).setBackgrounds(backgroundColours).setBorder(true, true, true, true, false, true);

      sheets[j].getRange(rowStart, 7, numRows, 2).setBorder(true, true, true, true, true, null);
      sheets[j].getRange(rowStart, 9, numRows, 2).setBorder(null, true, null, true, false, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);

      if (sheetNames[j] !== "Shipped")
        sheets[j].autoResizeRows(rowStart, numRows);
      
      notesRange.setBackgrounds(noteBackgroundColours).setRichTextValues(richTextValues);
    }
    else if (sheetNames[j] === "ItemsToRichmond")
    {
      numHeaders = 3;
      rowStart = numHeaders + 1;
      lastRow = sheets[j].getLastRow();
      lastCol = 8;
      numRows = lastRow - numHeaders;
      headerRange = sheets[j].getRange(       1, 1, numHeaders, lastCol + 1);
        dataRange = sheets[j].getRange(rowStart, 1,    numRows, lastCol);
       dataValues = dataRange.getValues(); 

      // Set the column widths and the header's row heights
      sheets[j].setRowHeights(1, 2, 65);
      for (var c = 0; c < lastCol + 1; c++)
        sheets[j].setColumnWidth(c + 1, [90, 100, 75, 700, 250, 75, 100, 125, 25][c]);

      // Prepare and set all of the headerRange values and formats
      headerValues = [  ['Scan Here', '', '', 'ITEMS BEING SHIPPED TO PNT RICHMOND', '', '', '', '', 'TRANSFERRED'],
                        ['Order Date','Entered By:','UoM','Description','Notes', 'Shipped','Carrier','Received By', ''], 
                        [1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0]];
      headerFonts = [['Libre Barcode 128', ...new Array(8).fill('Verdana')], new Array(9).fill('Arial'), new Array(9).fill('Arial')];
      headerFontSizes = [[35, 10, 10, 30, ...new Array(5).fill(10)], [...new Array(5).fill(14), ...new Array(4).fill(12)], new Array(9).fill(10)];
      headerFontColours = [['black', ...new Array(8).fill('white')],  new Array(9).fill('black'), new Array(9).fill('black')];
      headerFontWeights = [['normal', ...new Array(8).fill('bold')],  new Array(9).fill('bold'), new Array(9).fill('bold')];
      headerBackgroundColours = [['white', ...new Array(8).fill('#5b95f9')], [...new Array(5).fill('white'), '#fff2cc', '#d9ead3', '#d9ead3', ''], new Array(9).fill('white')];
      headerRange.setFontLine('none').setFontStyle('normal').setFontFamilies(headerFonts).setFontSizes(headerFontSizes).setFontWeights(headerFontWeights).setFontColors(headerFontColours)
        .setNumberFormat('@').setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true).setBackgrounds(headerBackgroundColours).setValues(headerValues);

      // Prepare all of the dataRange values and formats
      horizontalAlignments = [...Array(numRows)].map(e => ['right', 'center', 'center', 'left', 'center', 'center', 'center', 'left']);
      wrapStrategies = [...Array(numRows)].map(e => [...new Array(2).fill(SpreadsheetApp.WrapStrategy.OVERFLOW),  ...new Array(3).fill(SpreadsheetApp.WrapStrategy.WRAP), 
                                                                          SpreadsheetApp.WrapStrategy.CLIP,       ...new Array(2).fill(SpreadsheetApp.WrapStrategy.WRAP)]);
      backgroundColours = [...Array(numRows)].map(e => [null, 'white', 'white', 'white', null, '#fff2cc', '#d9ead3', '#d9ead3']);
      numberFormats = [...Array(numRows)].map(e => ['dd MMM yyyy', '@', '@', '@', '@', '#.#', '@', '@']);
      notesRange = sheets[j].getRange(rowStart, 5, numRows); // To preserve the background and text colours of the Notes, we must store that data first
      noteBackgroundColours = notesRange.getBackgrounds();
      richTextValues = notesRange.getRichTextValues();

      // Apply the correct highlighting for the dates
      for (var i = 0; i < numRows; i++)
        backgroundColours[i][0] = (dataValues[i][0] >= ONE_WEEK) ? GREEN : ( (dataValues[i][0] >= ONE_MONTH) ? YELLOW : RED );

      // Set all of the dataRange values and formats
      dataRange.setFontSize(10).setFontLine('none').setFontStyle('normal').setFontWeight('bold').setFontFamily('Arial')
        .setHorizontalAlignments(horizontalAlignments).setVerticalAlignment('middle').setWrapStrategies(wrapStrategies)
        .setNumberFormats(numberFormats).setBackgrounds(backgroundColours).setBorder(true, true, false, false, false, true);
      
      sheets[j].getRange(rowStart, 6, numRows, 3).setBorder(null, true, null, true, false, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
      notesRange.setBackgrounds(noteBackgroundColours).setRichTextValues(richTextValues);
      sheets[j].autoResizeRows(rowStart, numRows);
    }
    else if (sheetNames[j] === "Manual Counts" || sheetNames[j] === "InfoCounts")
    {
      numHeaders = 3;
      rowStart = numHeaders + 1;
       maxRow = sheets[j].getMaxRows();
      lastRow = sheets[j].getLastRow();
      lastCol = sheets[j].getMaxColumns();
      numRows = lastRow - numHeaders;
      headerRange = sheets[j].getRange(1, 1, numHeaders, lastCol);
      
      // Set the header's row heights and sheet's column widths
      for (var r = 0; r < numHeaders; r++)
        sheets[j].setRowHeightsForced(r + 1, 1, [45, 45, 2][r]);
      for (var c = 0; c < lastCol; c++)
        sheets[j].setColumnWidth(c + 1, [900, 80, 80, 500, 130, 85, 85][c]);

      numberFormats = [...Array(numHeaders)].map(e => new Array(lastCol).fill('@'));

      if (sheetNames[j] === "Manual Counts")
      {
        spreadsheet.setNamedRange('Completed_ManualCounts', sheets[j].getRange('B1'));
        spreadsheet.setNamedRange('Progress_ManualCounts', sheets[j].getRange('B3'));
        spreadsheet.setNamedRange('Remaining_ManualCounts', sheets[j].getRange('C1'));
        headerValues = [[  Store_Name + ' Manual Counts', '=COUNTA($C$4:$C)', '=COUNTA($A$4:$A)-Completed_ManualCounts', 'Scanning Data', '', 'Inflow Data', ''], 
                        ['Description - Vendor - Category - UoM - Sku#', 'Current Count', 'New Count', 'Running Sum', 'Last Scan Time (ms)', 'Location', 'Quantity'], 
                        ['', '=Completed_ManualCounts&\"/\"&(Completed_ManualCounts + Remaining_ManualCounts)', '', '', '', '', '']];
        sheets[j].hideColumns(4, 4);
        numberFormats[2][6] = '#';
      }
      else
      {
        spreadsheet.setNamedRange('Completed_InfoCounts', sheets[j].getRange('B1'));
        spreadsheet.setNamedRange('Progress_InfoCounts', sheets[j].getRange('B3'));
        spreadsheet.setNamedRange('Remaining_InfoCounts', sheets[j].getRange('C1'));
        headerValues = [['New ' + Store_Name + ' Counts', '=COUNTA($C$4:$C$' + lastRow + ')', '=' + numRows + '-Completed_InfoCounts'], 
                        ['Description - Vendor - Category - UoM - Sku#', 'Current Count', 'New Count'], 
                        ['', '=Completed_InfoCounts&\"/\"&(Completed_InfoCounts+Remaining_InfoCounts)', '']];
      }

      // Prepare and set all of the headerRange formatting
      headerFontSizes = [new Array(lastCol).fill(18), new Array(lastCol).fill(12), new Array(lastCol).fill(2)];
      headerFontColours = [['black', '#b7b7b7', ...new Array(lastCol - 2).fill('black')], new Array(lastCol).fill('black'), new Array(lastCol).fill('black')]
      headerHorizontalAlignments = [['right', ...new Array(lastCol - 1).fill('center')], ['right', ...new Array(lastCol - 1).fill('center')], ['right', ...new Array(lastCol - 1).fill('center')]];
      headerRange.setFontLine('none').setFontWeight('bold').setFontStyle('normal').setFontFamily('Verdana').setFontColors(headerFontColours).setFontSizes(headerFontSizes)
        .setWrap(true).setNumberFormats(numberFormats).setVerticalAlignment('middle').setHorizontalAlignments(headerHorizontalAlignments).setBackground('white')
        .setBorder(null, null, null, null, null, true).setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK).setValues(headerValues)

      sheets[j].hideRows(3);

      if (sheetNames[j] === "Manual Counts")
        sheets[j].getRange(3, lastCol).insertCheckboxes().uncheck() // This is the "Add-One" mode for the Manual Scanner
        
      if (numRows > 0)
      {
        // Prepare and set all of the dataRange values and formats
        dataRange = sheets[j].getRange(rowStart, 1, numRows, lastCol);
        fontColours = [...Array(numRows)].map(e => ['black', '#b7b7b7', ...new Array(lastCol - 2).fill('black')]);
        horizontalAlignments = [...Array(numRows)].map(e => ['right', ...new Array(lastCol - 1).fill('center')]);
        numberFormats = [...Array(numRows)].map(e => ['@', '#.#', '0.#', ...new Array(lastCol - 3).fill('@')]);
        wrapStrategies = [...Array(numRows)].map(e => [SpreadsheetApp.WrapStrategy.OVERFLOW, ...new Array(lastCol - 1).fill(SpreadsheetApp.WrapStrategy.CLIP)]);
        dataRange.setFontSize(10).setFontLine('none').setFontWeight('bold').setFontStyle('normal').setFontFamily('Verdana').setFontColors(fontColours)
          .setBackground('white').setNumberFormats(numberFormats).setHorizontalAlignments(horizontalAlignments).setVerticalAlignment('middle').setWrapStrategies(wrapStrategies)
          .setBorder(true, true, false, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);

        if (sheetNames[j] === "Manual Counts")
          sheets[j].getRange(rowStart, 5, numRows, 2).setBorder(null, true, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID) 
                                                     .setBorder(null, null, null, null, true, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK)
                                                     .setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID)
          
        sheets[j].getRange(1, 3, numHeaders + numRows).setBorder(null, true, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
                                                      .setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);

        if (maxRow > lastRow)
          sheets[j].deleteRows(lastRow + 1, maxRow - lastRow); // Delete the blank rows

        sheets[j].autoResizeRows(rowStart, numRows);
      }
      else if (maxRow >= 5)
        sheets[j].deleteRows(5, maxRow - 4) // Leave 1 blank row
    }
    else if (sheetNames[j] === "Item Search")
    {
      if (sheets.length > 1) recentlyCreatedItems(spreadsheet, sheets[j]); // If the full spreadsheet is being formatted, then put the recently created items on the search page
      numHeaders = 3;
      rowStart = numHeaders + 1;
      lastCol = sheets[j].getMaxColumns();
      const MAX_NUM_ITEMS = 500;
      numRows = MAX_NUM_ITEMS;
      maxRow = sheets[j].getMaxRows();
      headerRange = sheets[j].getRange(1, 1, numHeaders, lastCol);
        dataRange = sheets[j].getRange(rowStart, 1, numRows, lastCol);
      headerValues = headerRange.getValues();
      
      sheets[j].setRowHeight(1, 150);
      sheets[j].setRowHeight(2,  32);
      sheets[j].setRowHeight(3,  22);
      for (var c = 0; c < lastCol; c++)
        sheets[j].setColumnWidth(c + 1, [160, 725, 85, 60, 60, 60, 60, 180][c]);

      // Prepare and set all of the headerRange values and formats
      headerValues[1][3] = 'Current Stock In Each Location';
      headerValues[1][7] = 'Items last updated on';
      headerValues[2][3] = (isRichmondSpreadsheet(spreadsheet)) ? 'Rich' : ((isParksvilleSpreadsheet(spreadsheet)) ? 'Parks' : 'Rupert');
      headerValues[2][4] = (isRichmondSpreadsheet(spreadsheet)) ? 'Parks' : 'Rich';
      headerValues[2][5] = (isRichmondSpreadsheet(spreadsheet) || isParksvilleSpreadsheet(spreadsheet)) ? 'Rupert' : 'Parks';
      headerValues[2][6] = 'Trites';
      headerNumberFormats = [new Array(8).fill('@'), new Array(8).fill('@'), [...new Array(7).fill('@'), 'dd MMM yyyy']]
      headerFontSizes = [[16, 14, 14, ...new Array(5).fill(28)], [12, ...new Array(7).fill(11)], [12, ...new Array(7).fill(11)]];
      headerBackgroundColours = [['#4a86e8', 'white', 'white', ...new Array(5).fill('#4a86e8')], new Array(8).fill('4a86e8'), new Array(8).fill('4a86e8')];
      headerFontColours = [['white', 'black', 'black', ...new Array(5).fill('white')], new Array(8).fill('white'), new Array(8).fill('white')];
      headerRange.setFontFamily('Arial').setFontWeight('bold').setFontLine('none').setFontStyle('normal').setFontSizes(headerFontSizes).setFontColors(headerFontColours)
        .setBackgrounds(headerBackgroundColours).setNumberFormats(headerNumberFormats).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true)
        .setBorder(true, true, true, true, null, null).setValues(headerValues);

      // Format the search box
      sheets[j].getRange(1, 2, 1, 2).setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK)
        .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center").setVerticalAlignment("middle")
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).merge();

      fontSizes = [...Array(numRows)].map(e => [...new Array(lastCol - 1).fill(10), 12]);
      horizontalAlignments = [...Array(numRows)].map(e => ['center', 'left', ...new Array(lastCol - 2).fill('center')]);
      numberFormats = [...Array(numRows)].map(e => ['@', '@', 'dd MMM yyyy', '@', '@', '@', '@', '@']);
      dataRange.setFontFamily('Arial').setFontWeight('bold').setFontLine('none').setFontStyle('normal').setFontSizes(fontSizes).setNumberFormats(numberFormats)
        .setVerticalAlignment('middle').setHorizontalAlignments(horizontalAlignments).setWrap(true).setBorder(true, true, true, true, false, false);

      // Apply all of the different borders
      sheets[j].getRange(2, 4, 2, 5).setBorder(null, null, null, null, true, true, '#1155cc', SpreadsheetApp.BorderStyle.SOLID);
      sheets[j].getRange(2, 4, 2, 5).setBorder(true, true, null, null, null, null, '#1155cc', SpreadsheetApp.BorderStyle.SOLID_THICK);
      sheets[j].getRange(2, 4, 2, 4).setBorder(null, null, null, true, null, null, '#1155cc', SpreadsheetApp.BorderStyle.SOLID_THICK);
      sheets[j].getRange(1, 4).setFormula('=Remaining_InfoCounts&\" Items left to count on the InfoCounts page.\"')

      if (maxRow > MAX_NUM_ITEMS + 3)
        sheets[j].deleteRows(MAX_NUM_ITEMS + 4, maxRow - MAX_NUM_ITEMS - 3); // Delete the blank rows
    }
    else if (sheetNames[j] === "INVENTORY" || sheetNames[j] === "SearchData" || sheetNames[j] === "Recent")
    {
      numHeaders = (sheetNames[j] === "INVENTORY") ? (isRichmondSpreadsheet(spreadsheet) ? 7 : 9) : 1;
      rowStart = numHeaders + 1;
      maxRow = sheets[j].getMaxRows();
      lastRow = sheets[j].getLastRow();
      lastCol = sheets[j].getMaxColumns();
      numRows = lastRow - numHeaders;
      headerRange = sheets[j].getRange(       1, 1, numHeaders, lastCol);
        dataRange = sheets[j].getRange(rowStart, 1,    numRows, lastCol);
      columnWidths = isRichmondSpreadsheet(spreadsheet) ? [100, 675, 100, 100, 100, 120, 120, 100, 100] : [100, 675, 100, 100, 100, 100, 120, 100, 100];

      for (var c = 0; c < lastCol; c++)
        sheets[j].setColumnWidth(c + 1, columnWidths[c]);

      // Prepare and set all of the headerRange values and formats
      if (sheetNames[j] === "INVENTORY")
      {
        const date = sheets[j].getRange(4, 1).getValue().split(' on ')[1];
        const upcDatabaseLastUpdated = new Date(date);

        if (isRichmondSpreadsheet(spreadsheet))
        {
          sheets[j].setRowHeights(1, numHeaders - 1, 28);
          headerValues = headerRange.getValues();
          headerFontSizes = [ [30, 11, 13, 13, 13, 13, 20, 11, 12], 
                              [11, 11, 13, 13, 13, 13, 20, 11, 12],
                              [11, 11, 13, 13, 13, 13, 20, 11, 12],
                              [11, 11, 13, 13, 13, 13, 20, 11, 12],
                              [11, 11, 13, 13, 13, 13, 20, 11, 12],
                              [11, 11, 13, 13, 13, 13, 20, 11, 12],
                              new Array(lastCol).fill(8)];
          headerFontWeights = [ ['bold',   'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                new Array(lastCol).fill('bold')];
          headerFontColours = [ ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                new Array(lastCol).fill('black')];
          if (upcDatabaseLastUpdated <= ONE_WEEK) headerFontColours[1][3] = 'red';
          headerBackgroundColours = [ ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      new Array(lastCol).fill('#f0f0f0')];
          headerHorizontalAlignments = [['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        new Array(lastCol).fill('center')];
          sheets[j].getRange(3, 3, 3, 5).setFormulas([  ['=Remaining_InfoCounts&\" items on the infoCounts\npage that haven\'t been counted\"', '', '', '', '=Progress_InfoCounts'],
                                                        ['' , '', '', '', ''],
                                                        ['=Remaining_ManualCounts&\" items on the Manual Counts\npage that haven\'t been counted\"', '', '', '', '=Progress_ManualCounts']]);
        }
        else
        {
          sheets[j].setRowHeights(1, numHeaders - 1, 40);
          headerValues = headerRange.getValues();
          headerFontSizes = [ [30, 12, 12, 12, 12, 12, 22, 11, 12], 
                              [14, 12, 12, 12, 12, 12, 22, 11, 12],
                              [14, 12, 12, 12, 12, 12, 22, 11, 12],
                              [14, 12, 12, 12, 12, 12, 22, 11, 12],
                              [14, 12, 12, 12, 12, 12, 22, 11, 12],
                              [14, 12, 12, 12, 12, 12, 22, 11, 12],
                              [14, 12, 12, 12, 12, 12, 22, 11, 12],
                              [14, 12, 12, 12, 12, 12, 22, 11, 12],
                              new Array(lastCol).fill(8)];
          headerFontWeights = [ ['bold',   'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold', 'bold'],
                                new Array(lastCol).fill('bold')];
          headerFontColours = [ ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                ['black', 'black', 'black', 'black', 'black', 'black', 'black', '#666666', '#666666'],
                                new Array(lastCol).fill('black')];
          if (upcDatabaseLastUpdated <= ONE_WEEK) headerFontColours[3][0] = 'red';
          headerBackgroundColours = [ ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      ['#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', '#f0f0f0', 'white', 'white', 'white'],
                                      new Array(lastCol).fill('#f0f0f0')];
          headerHorizontalAlignments = [['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        ['right', 'right', 'right', 'right', 'right', 'right', 'center', 'center', 'center'],
                                        new Array(lastCol).fill('center')];
          sheets[j].getRange(3, 7, 4).setFormulas([ ['=COUNTIF(Received_Checkbox,FALSE)'],
                                                    ['=COUNTIF(ItemsToRichmond_Checkbox,FALSE)'],
                                                    ['=COUNTIF(Order_ActualCounts,">=0")'],
                                                    ['=COUNTIF(Shipped_ActualCounts,">=0")']]);
          sheets[j].getRange(7, 1, 2, 7).setFormulas([['=Remaining_InfoCounts&\" items on the infoCounts page that haven\'t been counted\"',      '', '', '', '', '', '=Progress_InfoCounts'],
                                                      ['=Remaining_ManualCounts&\" items on the Manual Counts page that haven\'t been counted\"', '', '', '', '', '', '=Progress_ManualCounts']]);
        }
      }
      else
      {
        headerFontWeights = [new Array(lastCol).fill('bold')]
        headerHorizontalAlignments = [new Array(lastCol).fill('center')]
        headerBackgroundColours = [new Array(lastCol).fill('#f0f0f0')];
        headerFontColours = [new Array(lastCol).fill('black')];
        headerFontSizes = [new Array(lastCol).fill(8)];
        sheets[j].hideSheet();
      }
      
      headerRange.setFontFamily('Arial').setFontWeights(headerFontWeights).setFontLine('none').setFontStyle('normal').setFontSizes(headerFontSizes).setFontColors(headerFontColours)
        .setBackgrounds(headerBackgroundColours).setNumberFormat('@').setHorizontalAlignments(headerHorizontalAlignments).setVerticalAlignment('middle');

      // Prepare and set all of the dataRange values and formats
      horizontalAlignments = [...Array(numRows)].map(e => ['center', 'right', ...new Array(lastCol - 2).fill('center')]);
      dataRange.setFontSize(8).setFontLine('none').setFontStyle('normal').setFontWeight('normal').setFontFamily('Arial').setBackground('white')
        .setHorizontalAlignments(horizontalAlignments).setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setNumberFormat('@');

      if (maxRow > lastRow)
        sheets[j].deleteRows(lastRow + 1, maxRow - lastRow); // Delete the blank rows

      sheets[j].setFrozenRows(numHeaders);
      sheets[j].autoResizeRows(numHeaders, numRows + 1);
    }
    else if (sheetNames[j] === "UPC Database" || sheetNames[j] === "Manually Added UPCs" || sheetNames[j] === "UPCs to Unmarry")
    {
      numHeaders = 1;
      rowStart = numHeaders + 1;
      maxRow = sheets[j].getMaxRows();
      lastRow = sheets[j].getLastRow();
      lastCol = sheets[j].getMaxColumns();
      numRows = lastRow - numHeaders;
      headerRange = sheets[j].getRange(1, 1, numHeaders, lastCol);

      if (sheetNames[j] === "UPC Database")
      {
        columnWidths = [125, 100, 600, 100]
        horizontalAlignments = [...Array(numRows)].map(e => ['left', 'center', 'left', 'center']);
      }
      else if (sheetNames[j] === "Manually Added UPCs")
      {
        columnWidths = [125, 125, 100, 600, 125]
        horizontalAlignments = [...Array(numRows)].map(e => ['left', 'left', 'center', 'left', 'center']);
        sheets[j].hideColumns(5);
      }
      else
      {
        columnWidths = [125, 600]
        horizontalAlignments = [...Array(numRows)].map(e => ['left', 'left']);
      }

      for (var c = 0; c < lastCol; c++)
        sheets[j].setColumnWidth(c + 1, columnWidths[c]);

      // Prepare and set all of the headerRange values and formats
      headerRange.setFontFamily('Arial').setFontWeight('bold').setFontLine('none').setFontStyle('normal').setFontSize(18).setFontColor('black')
        .setBackground('white').setNumberFormat('@').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);

      if (numRows > 0)
      {
        // Prepare and set all of the dataRange values and formats
        dataRange = sheets[j].getRange(rowStart, 1, numRows, lastCol).setFontSize(10).setFontLine('none').setFontStyle('normal').setFontWeight('normal').setFontFamily('Arial')
          .setBackground('white').setHorizontalAlignments(horizontalAlignments).setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setNumberFormat('@');

        if (maxRow > lastRow)
          sheets[j].deleteRows(lastRow + 1, maxRow - lastRow); // Delete the blank rows
      }

      sheets[j].hideSheet();
    }
    else if (sheetNames[j] === "Count Log")
    {
      numHeaders = 1;
      rowStart = numHeaders + 1;
       maxRow = sheets[j].getMaxRows();
      lastRow = sheets[j].getLastRow();
      lastCol = sheets[j].getMaxColumns();
      numRows = lastRow - numHeaders;
      headerRange = sheets[j].getRange(1, 1, numHeaders, lastCol);
      
      // Set the header's row heights and sheet's column widths
      sheets[j].setRowHeight(1, 30);
      for (var c = 0; c < lastCol; c++)
        sheets[j].setColumnWidth(c + 1, [150, 1000, 100, 100][c]);

      // Prepare and set all of the headerRange values and formats
      headerValues = [["SKU", "Description", "Sheet", "Date"]];
      headerRange.setFontLine('none').setFontWeight('bold').setFontStyle('normal').setFontFamily('Arial').setFontColor('black').setFontSize(14)
        .setWrap(true).setNumberFormat('@').setVerticalAlignment('middle').setHorizontalAlignment('center').setBackground('white').setBorder(false, false, false, false, false, false);

      // Prepare and set all of the dataRange values and formats
      dataRange = sheets[j].getRange(rowStart, 1, numRows, lastCol);
      numberFormats = [...Array(numRows)].map(e => [...new Array(lastCol - 1).fill('@'), 'dd MMM yyyy']);
      horizontalAlignments = [...Array(numRows)].map(e => ['left', 'left', 'center', 'right']);
      dataRange.setFontSize(10).setFontLine('none').setFontStyle('normal').setFontWeight('normal').setFontFamily('Arial').setFontColor('black')
        .setBackground('white').setNumberFormats(numberFormats).setHorizontalAlignments(horizontalAlignments).setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
        .setBorder(false, false, false, false, false, false);

      if (maxRow > lastRow)
        sheets[j].deleteRows(lastRow + 1, maxRow - lastRow); // Delete the blank rows

      sheets[j].autoResizeRows(rowStart, numRows).hideSheet();
    }
    else if (sheetNames[j] === "Data Validation")
      sheets[j].hideSheet().getDataRange().setFontSize(20).setFontLine('none').setFontStyle('normal').setFontWeight('normal').setFontFamily('Arial').setFontColor('black')
        .setBackground('white').setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    else if (sheetNames[j] === "Manual Scan" || sheetNames[j] === 'Item Scan')
      (sheetNames[j] === "Manual Scan") ? 
        sheets[j].getRange(1, 1, 1, 2).setNumberFormats([['@', '#.#']]).setFontSize(25).setFontLine('none').setFontStyle('none').setFontWeight('normal').setFontFamily('Arial')
          .setFontColor('black').setBackground('white').setVerticalAlignment('middle').setHorizontalAlignment('center').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP) :
        sheets[j].getRange(1, 1).setNumberFormat('@').setFontSize(25).setFontLine('none').setFontStyle('none').setFontWeight('normal').setFontFamily('Arial').setFontColor('black')
          .setBackground('white').setVerticalAlignment('middle').setHorizontalAlignment('center').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    else if (sheetNames[j] === "inFlowPick" || sheetNames[j] === "Suggested inFlowPick" || sheetNames[j] === "Moncton's inFlow Item Quantities")
    {
      numHeaders = 1;
      rowStart = numHeaders + 1;
       maxRow = sheets[j].getMaxRows();
      lastRow = sheets[j].getLastRow();
      lastCol = sheets[j].getMaxColumns();
      numRows = lastRow - numHeaders;
      headerRange = sheets[j].getRange(1, 1, numHeaders, lastCol);

      headerRange.setFontLine('none').setFontWeight('bold').setFontStyle('normal').setFontFamily('Arial').setFontColor('white').setFontSize(16)
        .setWrap(true).setNumberFormat('@').setVerticalAlignment('middle').setHorizontalAlignment('center').setBackground('#f1c232')

      if (sheetNames[j] === "Suggested inFlowPick")
      {
        var richTextValue = SpreadsheetApp.newRichTextValue().setText('Adagio Quantity (Trites + Moncton)')
          .setTextStyle(0, 16, SpreadsheetApp.newTextStyle().setFontSize(16).build())
          .setTextStyle(16, 34, SpreadsheetApp.newTextStyle().setFontSize(14).build())
          .build()
        headerRange.offset(0, 4, 1, 1).setRichTextValues([[richTextValue]])
      }

      // Prepare and set all of the dataRange values and formats
      dataRange = sheets[j].getRange(rowStart, 1, numRows, lastCol);
      dataRange.setFontSize(10).setFontLine('none').setFontStyle('normal').setFontFamily('Arial').setFontColor('black')
        .setBackground('white').setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    }
    else 
      sheets[j].hideSheet()
  }
}

/**
 * This function checks the order and shipped sheet for rows that should have been transfered to another page. It will automatically delete those rows if necessary and 
 * move them to a new sheet if necessary. This function will be run on a trigger in before work.
 * 
 * @author Jarren Ralf
 */
// function autoMoveRows()
// {
//   const numCols = 11;
//   const numHeaders = 3;
//   const NOTES_COL = 6;
//   const TRANSFERRED_COL = 12;
//   const STATUS_COL_INDEX = 9;
//   const rowStart = numHeaders + 1;
//   const spreadsheet = SpreadsheetApp.getActive();
//   const orderSheet = spreadsheet.getSheetByName("Order");
//   const shippedSheet = spreadsheet.getSheetByName("Shipped");
//   const receivedSheet = spreadsheet.getSheetByName("Received");
//   const orderRange = orderSheet.getRange(rowStart, 1, orderSheet.getLastRow() - numHeaders, numCols);
//   const shippedRange = shippedSheet.getRange(rowStart, 1, shippedSheet.getLastRow() - numHeaders, numCols);
//   const receivedRange = receivedSheet.getRange(rowStart, 1, receivedSheet.getLastRow() - numHeaders, numCols)
//   const  ordVals = orderRange.getValues();
//   const shipVals = shippedRange.getValues();
//   const recdVals = receivedRange.getValues();
//   const numItemsOnShipped = shipVals.length;
//   const notesFontShippedSheet = shippedSheet.getRange(rowStart, NOTES_COL, numItemsOnShipped).getRichTextValues();
//   const backgroundColoursShippedSheet = shippedRange.getBackgrounds();
//   const rowValuesShippedSheet = [], coloursShippedSheet = [], richTextValuesShippedSheet = [];
//   var onReceived = false;

//   // Th following convert the Date object to an integer so comparison operators, namely ===, will work appropriately when comparing dates 
//   const orderValues = ordVals.map(v => [(typeof v[ 0] === 'object') ? v[0].getTime() : v[0], v[1], v[2], v[3], v[4], 
//                         v[5], v[6], v[7], v[8], v[9], (typeof v[10] === 'object') ? v[10].getTime() : v[10]]);
//   const shippedValues = shipVals.map(v => [(typeof v[ 0] === 'object') ? v[0].getTime() : v[0], v[1], v[2], v[3], v[4], 
//                         v[5], v[6], v[7], v[8], v[9], (typeof v[10] === 'object') ? v[10].getTime() : v[10]]);
//   const receivedValues = recdVals.map(v => [(typeof v[ 0] === 'object') ? v[0].getTime() : v[0], v[1], v[2], v[3], v[4], 
//                         v[5], v[6], v[7], v[8], v[9], (typeof v[10] === 'object') ? v[10].getTime() : v[10]]);

//   for (var i = numItemsOnShipped - 1; i >= 0; i--) // Loop through all of the shipped rows (starting from the bottom)
//   {
//     if (shippedValues[i][STATUS_COL_INDEX] === 'Received') // Find the values in the STATUS column that say RECEIVED, and hence need to be deleted and possibly moved over
//     {
//       for (var j = 0; j < receivedValues.length; j++) // Loop through all of the received values
//       {
//         for (var k = 0; k < numCols; k++) // Loop through all of the columms
//         {
//           if (shippedValues[i][k] === receivedValues[j][k]) // Check that all of the row values match
//           {
//             if (k === numCols - 1) // All of the row values match
//               onReceived = true;
//           }
//           else 
//             break; // One of the row values didn't match, so move on the the next row
//         }
//       }

//       shippedSheet.deleteRow(i + rowStart); // Delete the original
//       if (!onReceived) // The item is not on the received page, so store the necessary values that are need to move it over
//       {
//         rowValuesShippedSheet.push(shipVals[i]);
//         coloursShippedSheet.push(backgroundColoursShippedSheet[i]);
//         richTextValuesShippedSheet.push(notesFontShippedSheet[i]);
//       }
//     }
//   }

//   const numRows = coloursShippedSheet.length

//   // Move the appropriate rows over to the received page
//   if (numRows !== 0)
//   {
//     receivedSheet.insertRowsAfter(numHeaders, numRows);
//     const destinationRange = receivedSheet.getRange(rowStart, 1, numRows, numCols);
//     destinationRange.setValues(rowValuesShippedSheet).setBackgrounds(coloursShippedSheet);
//     receivedSheet.getRange(rowStart, NOTES_COL, numRows).setRichTextValues(richTextValuesShippedSheet);
//     receivedSheet.getRange(rowStart, TRANSFERRED_COL, numRows).insertCheckboxes();
//   }

//   // Set up the values for the order sheet and recompute the values for the shipped page because rows may have been deleted
//   const numItemsOnOrder = orderValues.length;
//   const rowNums = [], rowValuesOrderSheet = [];
//   const shippedRangeUpdated = shippedSheet.getRange(rowStart, 1, shippedSheet.getLastRow() - numHeaders, numCols);
//   const shippedValuesUpdated = shippedRangeUpdated.getValues().map(v => [(typeof v[ 0] === 'object') ? v[0].getTime() : v[0], v[1], 
//                                 v[2], v[3], v[4], v[5], v[6], v[7], v[8], v[9], (typeof v[10] === 'object') ? v[10].getTime() : v[10]]);
//   var onShipped = false;

//   for (var i = numItemsOnOrder - 1; i >= 0; i--) // Loop through all of the order rows (starting from the bottom)
//   {
//     // Find the rows on the order page that need to be deleted and possibly moved over
//     if (orderValues[i][STATUS_COL_INDEX] !== '' && orderValues[i][STATUS_COL_INDEX] !== 'Back to Order' && orderValues[i][STATUS_COL_INDEX] !== 'B/O')
//     {
//       for (var j = 0; j < shippedValuesUpdated.length; j++) // Loop through all of the UPDATED shipped values
//       {
//         for (var k = 0; k < numCols; k++) // Loop through all of the columms
//         {
//           if (orderValues[i][k] === shippedValuesUpdated[j][k]) // Check that all of the row values match
//           {
//             if (k === numCols - 1) // All of the row values match
//               onShipped = true;
//           }
//           else
//             break;
//         }
//       }

//       if (!onShipped) // Store the row values and row numbers of the items that need to be moved to the shipped sheet
//       {
//         rowNums.push(i + rowStart);
//         rowValuesOrderSheet.push([ordVals[i]]);
//       }
//     }
//   }

//   orderSheet.activate(); // Set the order sheet as the active sheet, this is necessary for time-stamping

//   // Loop through all of the row numbers that need to be moved over and possibly deleted
//   for (var r = 0; r < rowNums.length; r++)
//   {
//     var shippedQty = rowValuesOrderSheet[r][0][8];
//     var orderedQty = rowValuesOrderSheet[r][0][2];
//     var value = rowValuesOrderSheet[r][0][STATUS_COL_INDEX];
//     var rowRange = orderSheet.getRange(rowNums[r], 1, 1, numCols);

//     if (value == "Item Not Available" || value == "Discontinued")
//       transferRow(orderSheet, shippedSheet, rowNums[r], rowValuesOrderSheet[r], numCols, true);
//     else // This means order and shipped quantities need to be checked
//     {
//       if (isNumber(shippedQty) && shippedQty > 0 && isNumber(orderedQty) && isNotBlank(orderedQty)) // If the shipped and order quantities are valid 
//       {
//         if (shippedQty >= orderedQty) // This is a complete shipment (No Back Orders)
//         {
//           if (value == "Carrier Not Assigned")
//             transferRow(orderSheet, shippedSheet, rowNums[r], rowValuesOrderSheet[r], numCols, true);
//           else
//           {
//             var dataValidation = spreadsheet.getSheetByName("Data Validation").getRange('B:C').getValues(); // These are all the data validation choices of carriers, etc.
            
//             for (var i = 0; i < dataValidation.length; i++)
//             {
//               if (value == dataValidation[i][0]) // The value selected matches th i-th data validation
//                 transferRow(orderSheet, shippedSheet, rowNums[r], rowValuesOrderSheet[r], numCols, true, dataValidation[i][1], dataValidation[i][0]);
//             }
//           }
//         }
//         else // Partial shipment, there some portion of the item will be on back order
//         {
//           if (value == "Carrier Not Assigned")
//           {
//             transferRow(orderSheet, shippedSheet, rowNums[r], rowValuesOrderSheet[r], numCols, false);
//             updateBO(rowRange, rowValuesOrderSheet[r]);
//           }
//           else
//           {
//             var dataValidation = spreadsheet.getSheetByName("Data Validation").getRange('B:C').getValues(); // These are all the data validation choices of carriers, etc.
            
//             for (var i = 0; i < dataValidation.length; i++)
//             {
//               if (value == dataValidation[i][0]) // The value selected matches th i-th data validation
//               {
//                 transferRow(orderSheet, shippedSheet, rowNums[r], rowValuesOrderSheet[r], numCols, false, dataValidation[i][1], dataValidation[i][0]);
//                 updateBO(rowRange, rowValuesOrderSheet[r]);
//               }
//             }
//           }
//         }
//       }
//     }
//   }
// }

/**
* Calculates Easter in the Gregorian/Western (Catholic and Protestant) calendar 
* based on the algorithm by Oudin (1940) from http://www.tondering.dk/claus/cal/easter.php
* @returns {array} [int month, int day]
*/
function calculateGoodFriday(year)
{
	var f = Math.floor,
		// Golden Number - 1
		G = year % 19,
		C = f(year / 100),
		// related to Epact
		H = (C - f(C / 4) - f((8 * C + 13)/25) + 19 * G + 15) % 30,
		// number of days from 21 March to the Paschal full moon
		I = H - f(H/28) * (1 - f(29/(H + 1)) * f((21-G)/11)),
		// weekday for the Paschal full moon
		J = (year + f(year / 4) + I + 2 - C + f(C / 4)) % 7,
		// number of days from 21 March to the Sunday on or before the Paschal full moon
		L = I - J,
		month = 3 + f((L + 40)/44),
		day = L + 28 - 31 * f(month / 4) - 2;
  
    // If the day is negative, make the appropriate changes to the values of month and day
    if (day < 0) 
    {
      month--;
      day = 31 + day
    }

	return [month - 1, day];
}

/**
 * This function clears the inFlow pick list.
 * 
 * @author Jarren Ralf
 */
function clearInflowPickList()
{
  const sheet = SpreadsheetApp.getActiveSheet();
  const numRows = sheet.getLastRow() - 2

  if (numRows > 0)
    SpreadsheetApp.getActiveSheet().getRange(3, 1, numRows, 5).clearContent()
}

/**
 * This function updates the inventory and search data from a csv file.
 * 
 * @author Jarren Ralf
 */
function clearInventory()
{
  const startTime = new Date().getTime();
  const spreadsheet = SpreadsheetApp.getActive();
  const itemSearchSheet = spreadsheet.getSheetByName("Item Search");
  const  inventorySheet = spreadsheet.getSheetByName("INVENTORY");
  const numRowsRange = (isRichmondSpreadsheet(spreadsheet)) ? inventorySheet.getRange(1, 7, 1, 3) : inventorySheet.getRange(1, 7, 2, 3);

  itemSearchSheet.getRange('B1').clearContent(); // Clear the search box on the Item Search page
  numRowsRange.clearContent(); // Clear the number of rows on the inventory page
  
  dateStamp(3, 8, 1, itemSearchSheet); // Place a dateStamp on the item Search page
  var itemNumber;
  
  const items = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString());
  const numRows = items.length;
  const sku = items[0].indexOf('Item #')
  const tritesQty = items[0].indexOf('Trites')
  const inflowData = Object.values(Utilities.parseCsv(DriveApp.getFilesByName("inFlow_StockLevels.csv").next().getBlob().getDataAsString()).reduce((sum, item) => {
    itemNumber = item[0].split(' - ').pop().toString().toUpperCase(); // Get the SKU number from the back of the description

    // If the item already has a sum associate with this item Number, then this skus is put away in multiple locations, and therefore add up the quantities from all locations
    if (sum[itemNumber]) // Inflow item might have a conversion factor associated with it so that the Price Unit in Adagio is consistent with the recorded quantity in inFlow
      sum[itemNumber][1] = (inflow_conversions.hasOwnProperty(itemNumber)) ? Number(sum[itemNumber][1]) + Number(item[4])*inflow_conversions[itemNumber] : Number(sum[itemNumber][1]) + Number(item[4]); 

    // Add the item to the new list if it contains the typical google sheets item format with "space - space" more than 4 times
    else if (item[0].split(' - ').length > 4) // Inflow item might have a conversion factor associated with it so that the Price Unit in Adagio is consistent with the recorded quantity in inFlow
      sum[itemNumber] = [itemNumber, (inflow_conversions.hasOwnProperty(itemNumber)) ? Number(item[4])*inflow_conversions[itemNumber] : Number(item[4])]; 

    return sum; // Return the total inventory quantity for the item number so far
  }, {}));
  
  var isInFlowItem;

  if (isRichmondSpreadsheet(spreadsheet))
  {
    const data = items.map(col => {
      isInFlowItem = inflowData.find(item => item[0].split(' - ').pop().toUpperCase() == col[sku].toString().toUpperCase()) // Inflow sku at back
      col[tritesQty] = (isInFlowItem) ? isInFlowItem[1] : ''; // Add Trites inventory values if they are found in inFlow

      return [col[0], col[1], null, col[2], col[3], col[4], col[5], col[sku], col[6]];
    }) 

    data[0][tritesQty + 1] = "Trites (inFlow)";
    inventorySheet.getRange('A8:I').clearContent();
    inventorySheet.getRange('A7:A').activate(); // This line activates the entire first column of the spreadsheet to verify the number of rows of the sheet
    inventorySheet.getRange(7, 1, numRows, data[0].length).setNumberFormat('@').setValues(data);
    numRowsRange.setValues([[numRows - 1, dateStamp(undefined, null, null, null, 'dd MMM HH:mm'), getRunTime(startTime)]]);

    spreadsheet.getSheetByName("Assembly").clearContents();
    spreadsheet.getSheetByName("UoM Conversion").clearContents();
  }
  else
  {
    const data = items.map(col => {
      isInFlowItem = inflowData.find(item => item[0].split(' - ').pop().toUpperCase() == col[sku].toString().toUpperCase()) // Inflow sku at back
      col[tritesQty] = (isInFlowItem) ? isInFlowItem[1] : ''; // Add Trites inventory values if they are found in inFlow

      return [col[0], col[1], col[2], col[3], col[4], col[5], col[sku], col[6], col[7]]
    }) 

    data[0][tritesQty] = "Trites (inFlow)";
    inventorySheet.getRange('A10:I').clearContent();
    inventorySheet.getRange('A9:A').activate(); // This line activates the entire first column of the spreadsheet to verify the number of rows of the sheet
    inventorySheet.getRange(9, 1, numRows, data[0].length).setNumberFormat('@').setValues(data);
    const date1 = dateStamp(undefined, null, null, null, 'dd MMM HH:mm');
    const runTime1 = getRunTime(startTime);
    numRowsRange.setValues([[numRows - 1, date1, runTime1],[null, null, null]]); // The number of active items from Adagio, including "No TS"
    
    const startTime2 = new Date().getTime();
    const searchData = data.filter(e => e[8] !== "No TS").map(f => [f[0], f[1], null, f[2], f[3], f[4], f[5]]); // Remove "No TS" items and keep units, descriptions and inventory
    const numItems = searchData.length;
    spreadsheet.getSheetByName("SearchData").clearContents().getRange(1, 1, numItems, searchData[0].length).setNumberFormat('@').setValues(searchData);
    numRowsRange.setValues([[numRows - 1, date1, runTime1],[numItems - 1, dateStamp(undefined, null, null, null, 'dd MMM HH:mm'), getRunTime(startTime2)]]);
  }
}

/**
 * This function clears all of the manual counts that have been completed, but leaves the ones that have a blank in the counts column.
 * 
 * @author Jarren Ralf
 */
function clearManualCounts()
{
  const startTime = new Date().getTime();
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName("Manual Counts");
  const numHeaders = 3;
  const numItems = sheet.getLastRow() - numHeaders;

  if (numItems > 0) // If there are items on the manual counts page
  {
    const numCols = sheet.getLastColumn();
    const rowStart = numHeaders + 1;
    const items = sheet.getSheetValues(rowStart, 1, numItems, numCols);
    const nonCountedItems = items.filter(count => count[2] === '' || count[0].split(' - ').pop() === 'MAKE_NEW_SKU'); // These are the items that have not been counted
    const numRemainingItems = nonCountedItems.length;

    if (numItems !== numRemainingItems) // If there are some items that have been counted, enter this code block
    {
      const numRows = sheet.getMaxRows() - numHeaders;
      sheet.getRange(rowStart, 1, numRows, numCols).clearContent();

      if (numRemainingItems !== 0) // There are some remaining items to count
      {
        sheet.getRange(rowStart, 1, numRemainingItems, numCols).setValues(nonCountedItems);
        sheet.deleteRows(numRemainingItems + rowStart, numRows - numRemainingItems);
      }
      else if (numRows - 1 !== 0) // There are no more items to count
        sheet.deleteRows(rowStart + 1, numRows - 1);
    }
  }

  if (isRichmondSpreadsheet(spreadsheet))
    spreadsheet.getSheetByName("INVENTORY").getRange(5, 3, 1, 7)
      .setValues([[ '=Remaining_ManualCounts&\" items on the Manual Counts page that haven\'t been counted\"', null, null, null, 
                    '=Progress_ManualCounts', dateStamp(undefined, null, null, null, 'dd MMM HH:mm'), getRunTime(startTime)]]);
  else
    spreadsheet.getSheetByName("INVENTORY").getRange(8, 1, 1, 9)
      .setValues([[ '=Remaining_ManualCounts&\" items on the Manual Counts page that haven\'t been counted\"', null, null, null, null, null, 
                    '=Progress_ManualCounts', dateStamp(undefined, null, null, null, 'dd MMM HH:mm'), getRunTime(startTime)]]);    
}

/**
 * This function sets a checkmark in the Transfered column of the Received page. The checkmark represents the status of 
 * whether the transfers have been completed in the Adagio system or not.
 * 
 * @author Jarren Ralf
 */
function completeReceived()
{ 
  const startTime = new Date().getTime();
  const START_ROW = 4;
  
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName("Received");
  const numRows = sheet.getLastRow() - START_ROW + 1;
  sheet.getRange(START_ROW, 12, numRows).setValue(true);
  spreadsheet.getSheetByName("INVENTORY").getRange(3, 7, 1, 3).setValues([['=COUNTIF(Received_Checkbox,FALSE)', dateStamp(undefined, null, null, null, 'dd MMM HH:mm'), getRunTime(startTime)]]);
}

/**
 * This function sets a checkmark in the Transfered column of the ItemsToRichmond page (given that hey have been "received" by one of the 
 * Richmond employees). The checkmark represents the status of whether the transfers have been completed in the Adagio system or not.
 * 
 * @author Jarren Ralf
 */
function completeToRichmond()
{
  const startTime = new Date().getTime();
  const RECEIVED_BY_COL = 0;
  const  TRANSFERED_COL = 1;
  const       START_ROW = 3;
  
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName("ItemsToRichmond");
  const numRows = sheet.getLastRow() - START_ROW + 1;
  const range = sheet.getRange(START_ROW, 8, numRows, 2); // Get the range off the last two columns
  var data = range.getValues();
  
  for (var i = 0; i < numRows; i++)
  {
    if (isNotBlank(data[i][RECEIVED_BY_COL])) // If the item has been received
      data[i][TRANSFERED_COL] = true;         // then set the transfer status to true
  }
  
  range.setValues(data); // Set the range with the updated values
  spreadsheet.getSheetByName("INVENTORY").getRange(4, 7, 1, 3)
    .setValues([['=COUNTIF(ItemsToRichmond_Checkbox,FALSE)', dateStamp(undefined, null, null, null, 'dd MMM HH:mm'), getRunTime(startTime)]]);
}

/**
 * This function concatenates the manually added UPCs with the imported list.
 * 
 * @author Jarren Ralf
 */
function concatManuallyAddedUPCs()
{
  var isInUpcDatabase, isInAdagioDatabase, additionalUPCs = [];
  const NUM_COLS = 5;
  const spreadsheet = SpreadsheetApp.getActive();
  const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
  const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");
  const inventorySheet = spreadsheet.getSheetByName("INVENTORY");
  const manuallyAddedUPCs = manAddedUPCsSheet.getSheetValues(1, 1, manAddedUPCsSheet.getLastRow(), NUM_COLS).filter(a => isNotBlank(a[0]));
  const currentUPCs = upcDatabaseSheet.getSheetValues(2, 1, upcDatabaseSheet.getLastRow() - 1, 1);

  manAddedUPCsSheet.clearContents().getRange(1, 1, manuallyAddedUPCs.length, NUM_COLS).setValues(manuallyAddedUPCs);
  additionalUPCs.shift(); // Remove the header
  var currentStock = 0; // Changes the index number for selecting the current stock from inventory data
  var transferData = inventorySheet.getRange('D8:H').getValues();
  var upc = 0; // The index of the sku

  const data = manuallyAddedUPCs.filter(v => {
    return currentUPCs.filter(u => {
      isInUpcDatabase = u[upc] == v[1]; // Match the UPC code
      if (!isInUpcDatabase) return isInUpcDatabase; // If the UPC isn't found in the UPC database, return false
      transferData.filter(w => {
        isInAdagioDatabase = w[4] == v[0];
        if (!isInAdagioDatabase) return isInAdagioDatabase;
        v[4] = w[currentStock];

        return isInAdagioDatabase;
      })
      return isInUpcDatabase;
    }).length != 0;
  })

  Logger.log(data.length)
  Logger.log(data)

  // header[1] = "Item Unit";
  // header[2] = "Adagio Description";
  // header[3] = "Current Stock";
  // const numRows = data.unshift(header); // Put the header back at the top of the database
  // upcDatabaseSheet.clearContents().getRange(1, 1, numRows, 4).setNumberFormat('@').setValues(data);
}

/**
 * This function moves the selected values on the sheet to the desired output page (Order, ItemsToRichmond, Manual Counts, and SearchData sheets).
 * 
 * @param {Sheet}   sheet   : The sheet that the selected items are being moved to.
 * @param {Number} startRow : The first row of the target sheet where the selected items will be moved to.
 * @param {Number}  numCols : The number of columns to grab from the item search page and move to the target sheet.
 * @param {Number}  qtyCol  : The column number of which the sheet in particular has the quantity value inputed.
 * @param {Boolean} isInfoCountsPage : A boolean that represents whether the user is on the infoCounts page or not.
 * @param {Sheet[]}  sheets : The additional sheets to post information to.
 * @param {Boolean} isNotFound : Whether the item value being copied to another sheet appears to exist in Adagio or not.
 * @return {[Number[], Number[], Number, Number]} The firstRows and numRows of the active ranges as well as the first and last row that the set of active ranges span.
 * @author Jarren Ralf
 */
function copySelectedValues(sheet, startRow, numCols, qtyCol, isInfoCountsPage, sheets, isNotFound)
{
  if (arguments.length !== 5) isInfoCountsPage = false;
  const isInventoryPage    = qtyCol == undefined;
  const isUpcPage          = qtyCol == 'upc';
  const isOrderPage        = qtyCol == 9;
  const isItemsToRichPage  = qtyCol == 6;
  const isManualCountsPage = qtyCol == 3 && !isInfoCountsPage;
  const  activeSheet = SpreadsheetApp.getActiveSheet();
  const activeRanges = activeSheet.getActiveRangeList().getRanges(); // The selected ranges on the item search sheet
  const numActiveRanges = activeRanges.length;

  var firstRows = [], lastRows = [], numRows = [];
  var itemValues = [[[]]];
  
  // Find the first row and last row in the the set of all active ranges
  for (var r = 0; r < numActiveRanges; r++)
  {
    firstRows[r] = activeRanges[r].getRow();
     lastRows[r] = activeRanges[r].getLastRow()
  }
  
  var     row = Math.min(...firstRows); // This is the smallest starting row number out of all active ranges
  var lastRow = Math.max( ...lastRows); // This is the largest     final row number out of all active ranges
  var finalDataRow = activeSheet.getLastRow() + 1;
  var numHeaders = (isInventoryPage) ? 9: 3;

  var col = (isManualCountsPage) ? 2 : 1; // Set the column of the item selection based on whether we are doing manual counts or item transfers
  var startCol = (isOrderPage) ? 4 : ( (isItemsToRichPage) ? 3 : 1 ); // Set the start column of the range destination based on whether we are doing manual counts or item transfers

  if (row > numHeaders && lastRow <= finalDataRow) // If the user has not selected an item, alert them with an error message
  { 
    for (var r = 0; r < numActiveRanges; r++)
    {
         numRows[r] = lastRows[r] - firstRows[r] + 1;
      itemValues[r] = activeSheet.getSheetValues(firstRows[r], col, numRows[r], numCols);
    }
    
    var itemVals = [].concat.apply([], itemValues); // Concatenate all of the item values as a 2-D array
    var numItems = itemVals.length;

    if (isInventoryPage) // Removing the "No TS" from items on the inventory page and moving them to the SearchData
    {
      itemVals.map(u => u.splice(2, 0, null)); // Add a null column to the items, where the Last Counted Date goes
      numCols++;
    }
    else if (isManualCountsPage) // Moving items from the search page to the manual counts page
    {
      itemVals.map(u => u.splice(1, 1)); // Remove the column that contains the last counted on date
      numCols--;
    }
    else if (isOrderPage) // Items that are being transfered from one location to another
      itemVals = itemVals.map(u => [u[0], u[1], (u[6] > 0) ? 'Trites Stock (As of Order Date): ' + u[6] : '', u[3]]); // Replace the column that has the last counted date with Trites Stock
    else if (isItemsToRichPage) // Items that are being transfered from one location to another
      itemVals.map(u => u.splice(2, 1, null)) // Replace the column that has the last counted date with a blank
    else if (isUpcPage)
    {
      const ui = SpreadsheetApp.getUi();
      var response, response2, item, itemJoined, upc, upcTemporaryValues = [], itemTemporaryValues = [];

      if (activeSheet.getSheetName() === "Manual Counts")
      {
        if (isNotFound)
        {
          itemVals = itemVals.map(() => {

            response = ui.prompt('Item Not Found', 'Please enter a new description:', ui.ButtonSet.OK_CANCEL)
            
            if (ui.Button.OK === response.getSelectedButton())
            {
              item = response.getResponseText().split(' - ');
              item[item.length - 1] = 'MAKE_NEW_SKU';
              itemJoined = item.join(' - ')
              response2 = ui.prompt('Item Not Found', 'Please scan the barcode for:\n\n' + itemJoined +'.', ui.ButtonSet.OK_CANCEL)

              if (ui.Button.OK === response2.getSelectedButton())
              {
                upc = response2.getResponseText();
                upcTemporaryValues.push(['MAKE_NEW_SKU', upc, item[item.length - 2], itemJoined])
                itemTemporaryValues.push([item[item.length - 2], itemJoined, '', ''])
                return [upc, item[item.length - 2], itemJoined, '']
              }
              else
              {
                itemTemporaryValues.push([null, null, null, null])
                upcTemporaryValues.push([null, null, null, null])
                return [null, null, null, null]
              }
            }
            else
            {
              itemTemporaryValues.push([null, null, null, null])
              upcTemporaryValues.push([null, null, null, null])
              return [null, null, null, null]
            }
          });

          sheets[1].getRange(sheets[1].getLastRow() + 1, 1, numItems, numCols).setNumberFormat('@').setValues(itemTemporaryValues);
        }
        else
        {
          itemVals = itemVals.map(u => {
            response = ui.prompt('Manually Add UPCs', 'Please scan the barcode for:\n\n' + u[0] +'.', ui.ButtonSet.OK_CANCEL)
            if (ui.Button.OK === response.getSelectedButton())
            {
              item = u[0].split(' - ');
              upc = response.getResponseText();
              upcTemporaryValues.push([item[item.length - 1], upc, item[item.length - 2], u[0]])
              return [upc, item[item.length - 2], u[0], u[1]]
            }
            else
            {
              upcTemporaryValues.push([null, null, null, null])
              return [null, null, null, null]
            }
          });
        }
      }
      else // Item Search Page
      {
        if (isNotFound)
        {
          itemVals = itemVals.map(() => {

            response = ui.prompt('Item Not Found', 'Please enter a new description:', ui.ButtonSet.OK_CANCEL)
            
            if (ui.Button.OK === response.getSelectedButton())
            {
              item = response.getResponseText().split(' - ');
              item[item.length - 1] = 'MAKE_NEW_SKU';
              itemJoined = item.join(' - ')
              response2 = ui.prompt('Item Not Found', 'Please scan the barcode for:\n\n' + itemJoined +'.', ui.ButtonSet.OK_CANCEL)

              if (ui.Button.OK === response2.getSelectedButton())
              {
                upc = response2.getResponseText();
                itemTemporaryValues.push([item[item.length - 2], itemJoined, '', ''])
                upcTemporaryValues.push(['MAKE_NEW_SKU', upc, item[4], itemJoined])
                return [upc, item[item.length - 2], itemJoined, '']
              }
              else
              {
                itemTemporaryValues.push([null, null, null, null])
                upcTemporaryValues.push([null, null, null, null])
                return [null, null, null, null]
              }
            }
            else
            {
              itemTemporaryValues.push([null, null, null, null])
              upcTemporaryValues.push([null, null, null, null])
              return [null, null, null, null]
            }
          });

          sheets[1].getRange(sheets[1].getLastRow() + 1, 1, numItems, numCols).setNumberFormat('@').setValues(itemTemporaryValues);
        }
        else
        {
          itemVals = itemVals.map(u => {
            response = ui.prompt('Manually Add UPCs', 'Please scan the barcode for:\n\n' + u[1] +'.', ui.ButtonSet.OK_CANCEL)
            if (ui.Button.OK === response.getSelectedButton())
            {
              upc = response.getResponseText();
              upcTemporaryValues.push([u[1].split(' - ').pop(), upc, u[0], u[1]])
              return [upc, u[0], u[1], u[3]]
            }
            else
            {
              upcTemporaryValues.push([null, null, null, null])
              return [null, null, null, null]
            }
          });
        }
      }

      sheets[0].getRange(sheets[0].getLastRow() + 1, 1, numItems, numCols).setNumberFormat('@').setValues(upcTemporaryValues);
    }

    if (sheet.getSheetName() !== "Manual Counts") 
      sheet.getRange(startRow, startCol, numItems, itemVals[0].length).setNumberFormat('@').setValues(itemVals); // Move the item values to the destination sheet
    else // Moving items to the Manual Counts page
    {
      if (startRow > 4) // There are existing items on the Manual Counts page
      {
        // Retrieve the existing items and add them to the new items
        // const groupedItems = groupHoochieTypes(sheet.getSheetValues(4, startCol, startRow - 4, sheet.getMaxColumns()).concat(itemVals.map(val => [...val, '', '', '', '', ''])), 0)
        // const items = [];

        // Object.keys(groupedItems).forEach(key => items.push(...sortHoochies(groupedItems[key], 0, key)));

        const items = sheet.getSheetValues(4, startCol, startRow - 4, sheet.getMaxColumns()).concat(itemVals.map(val => [...val, '', '', '', '', '']))
        sheet.getRange(4, startCol, items.length, items[0].length).setNumberFormat('@').setValues(items); // Move the item values to the destination sheet
      }
      else
      {
        // const groupedItems = groupHoochieTypes(itemVals, 0)
        // const items = [];

        // Object.keys(groupedItems).forEach(key => items.push(...sortHoochies(groupedItems[key], 0, key)));

        sheet.getRange(startRow, startCol, numItems, itemVals[0].length).setNumberFormat('@').setValues(itemVals); // Move the item values to the destination sheet
      }
    }

    if (!isInventoryPage && !isUpcPage) 
    {
      const nCols = (sheet.getSheetName() === "Manual Counts") ? 7 : 11;
      applyFullRowFormatting(sheet, startRow, numItems, nCols); // Apply the proper formatting
      
      if (sheet.getSheetName() !== "Manual Counts") 
        sheet.getRange(startRow, qtyCol).activate(); // Go to the quantity column on the destination sheet
      
      // If we are moving items onto the transfer pages, set the ordered date
      if (isOrderPage || isItemsToRichPage)
        dateStamp(startRow, 1, numItems); // Set the ordered date
    }
  }
  else
    SpreadsheetApp.getUi().alert('Please select an item from the list.');

  return [firstRows, numRows, row, lastRow];
}

/**
 * This function logs the latest inventory counts with a date, including SKUs, Descriptions, Vendors, Categories, 
 * UoMs, and Sheets that the inventory count was recorded on.
 * 
 * @author Jarren Ralf
 */
function countLog()
{
  const NUM_COLS = 4;
  const today = new Date();
  const yesterday = new Date(today.getFullYear(), today.getMonth(), today.getDay() - 1);
  const spreadsheet = SpreadsheetApp.getActive();
  const countLogPage = spreadsheet.getSheetByName("Count Log")
  var   countLogData = countLogPage.getSheetValues(2, 1, countLogPage.getLastRow() - 1, NUM_COLS);
  const recentCounts = countLogData.filter(c => c[3] > yesterday); // These are the counts that have been done since yesterday (helps with not adding duplicates to the list)
  const sheets = [spreadsheet.getSheetByName("Manual Counts"), spreadsheet.getSheetByName("InfoCounts")];
  if (!isRichmondSpreadsheet(spreadsheet)) sheets.push(spreadsheet.getSheetByName("Shipped"), spreadsheet.getSheetByName("Order"))
  countLogData = countLogData.concat(getPhysicalCounted_CountLog(sheets, today, recentCounts)).sort(sortByCountedDate); // All of the counts sourted by date
  const numRows = countLogData.length;
  const numberFormats = [...Array(numRows)].map(e => ['@', '@', '@', 'dd MMM yyyy']);
  countLogPage.getRange(2, 1, numRows, NUM_COLS).setNumberFormats(numberFormats).setValues(countLogData);
  spreadsheet.getSheetByName("INVENTORY").getRange(5, 1).setValue('The Count Log was last updated at ' + today.toLocaleTimeString() + ' on ' +  today.toDateString());
}

/**
 * This function logs the latest inventory counts with a date, including SKUs, Descriptions, Vendors, Categories, 
 * UoMs, and Sheets that the inventory count was recorded on.
 * 
 * @author Jarren Ralf
 */
function countsRemaining()
{
  const NUM_COLS = 4;
  const today = new Date();
  const yesterday = new Date(today.getFullYear(), today.getMonth(), today.getDay() - 1);
  const ONE_YEAR = new Date(today.getFullYear() - 1, today.getMonth(), today.getDay());
  const spreadsheet = SpreadsheetApp.getActive();
  const remainingCountsPage = spreadsheet.getSheetByName("Remaining Counts");
  const countLogPage = spreadsheet.getSheetByName("Count Log")
  var   countLogData = countLogPage.getSheetValues(2, 1, countLogPage.getLastRow() - 1, NUM_COLS);
  const recentCounts = countLogData.filter(c => c[3] > yesterday); // These are the counts that have been done since yesterday (helps with not adding duplicates to the list)
  const sheets = [spreadsheet.getSheetByName("Manual Counts"), spreadsheet.getSheetByName("InfoCounts")];
  var countsLeft = [], formats = [];

  if (isRichmondSpreadsheet(spreadsheet)) 
  {
    const inventorySheet = spreadsheet.getSheetByName("INVENTORY")
    const numItems = inventorySheet.getLastRow() - 7;
    const fullInventory = inventorySheet.getSheetValues(8, 2, numItems, 7);

    for (var i = 0; i < numItems; i++)
    {
      if (fullInventory[i][1] === '' || fullInventory[i][1] < ONE_YEAR)
      {
        countsLeft.push([fullInventory[i][6], fullInventory[i][0], fullInventory[i][1]]);
        formats.push(['@', '@', 'dd MMM yyy'])
      }
    }

    remainingCountsPage.getRange(2, 1, countsLeft.length, 3).setNumberFormats(formats).setValues(countsLeft)
  }
  else
  {
    const searchDataSheet = spreadsheet.getSheetByName("SearchData")
    const fullInventory = searchDataSheet.getSheetValues(2, 2, searchDataSheet.getLastRow() - 1, 2);

    for (var i = 0; i < numItems; i++)
    {
      if (fullInventory[i][1] === '' || fullInventory[i][1] < ONE_YEAR)
      {
        countsLeft.push([fullInventory[i][0].split(' - ').pop(), fullInventory[i][0], fullInventory[i][1]]);
        formats.push(['@', '@', 'dd MMM yyy'])
      }
    }

    remainingCountsPage.getRange(2, 1, countsLeft.length, 3).setNumberFormats(formats).setValues(countsLeft)

    sheets.push(spreadsheet.getSheetByName("Shipped"), spreadsheet.getSheetByName("Order"));
  }

  countLogData = countLogData.concat(getPhysicalCounted_CountLog(sheets, today, recentCounts)).sort(sortByCountedDate); // All of the counts sourted by date
  const numRows = countLogData.length;

  const numberFormats = [...Array(numRows)].map(e => ['@', '@', '@', 'dd MMM yyyy']);
  countLogPage.getRange(2, 1, numRows, NUM_COLS).setNumberFormats(numberFormats).setValues(countLogData);
  spreadsheet.getSheetByName("INVENTORY").getRange(5, 1).setValue('The Count Log was last updated at ' + today.toLocaleTimeString() + ' on ' +  today.toDateString());
}

/**
* This function creates a dateStamp and places it on the chosen row/s for the give column.
*
* @param {Number}     row      : The  row   number
* @param {Number}     col      : The column number
* @param {Number}   numRows    : *OPTIONAL* The number of rows
* @param {Sheet}     sheet     : *OPTIONAL* The destination sheet
* @param {String} customFormat : *OPTIONAL* The date / time format
* @return {Date}  timeNow : Returns the formated date dateStamp
* @author Jarren Ralf
*/
function dateStamp(row, col, numRows, sheet, customFormat)
{
  // If the function is sent only two arguments, namely the row and column, then set the dateStampRange appropriately
  var timeZone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();             // set timezone
  var dateStampFormat = (arguments.length === 5) ? customFormat : 'dd MMM yyyy';  // set dateStamp format
  var today = new Date();                                                         // Date object representing today's date
  var timeNow = Utilities.formatDate(today, timeZone, dateStampFormat);           // Set variable for current time string

  if (row !== undefined) // If the row value is defined, then print the timestamp in the appropriate place
  {
    if (arguments.length !== 4) sheet = SpreadsheetApp.getActiveSheet();
    var dateStampRange = (arguments.length == 2) ? sheet.getRange(row, col) : sheet.getRange(row, col, numRows); 
    (col === 1) ? dateStampRange.setBackground('#b6d7a8').setValue(timeNow) : dateStampRange.setValue(timeNow);
  }

  return timeNow;
}

/**
 * This function launches a modal dialog box which allows the user to click a download button, which will lead to 
 * a csv file of the export data being downloaded.
 * 
 * @param {String} importType : The type of information that will be imported into inFlow.
 * @author Jarren Ralf
 */
function downloadButton(importType)
{
  var htmlTemplate = HtmlService.createTemplateFromFile('DownloadButton')
  htmlTemplate.inFlowImportType = importType;
  var html = htmlTemplate.evaluate().setWidth(250).setHeight(75)
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Export');
}

/**
 * This function calls another function that will launch a modal dialog box which allows the user to click a download button, which will lead to 
 * a csv file of an inFlow Product Details containing barcodes to be downloaded, then imported into the inFlow inventory system.
 * 
 * @author Jarren Ralf
 */
function downloadButton_Barcodes()
{
  downloadButton('Barcodes')
}

/**
 * This function calls another function that will launch a modal dialog box which allows the user to click a download button, which will lead to 
 * a csv file of an inFlow Sales Order to be downloaded, then imported into the inFlow inventory system.
 * 
 * @author Jarren Ralf
 */
function downloadButton_SalesOrder()
{
  downloadButton('SalesOrder')
}

/**
 * This function calls another function that will launch a modal dialog box which allows the user to click a download button, which will lead to 
 * a csv file of inFlow Stock Levels for a particular set of items to be downloaded, then imported into the inFlow inventory system.
 * 
 * @author Jarren Ralf
 */
function downloadButton_StockLevels()
{
  downloadButton('StockLevels')
}

/**
 * This function takes the array of data on the Moncton's inFlow Item Quantities page and it creates a csv file that can be downloaded from the Browser.
 * 
 * @return Returns the csv text file that file be downloaded by the user.
 * @author Jarren Ralf
 */
function downloadInflowBarcodes()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName("Moncton's inFlow Item Quantities");
  const upcDatabase = spreadsheet.getSheetByName("UPC Database");
  const upcs = upcDatabase.getSheetValues(2, 1, upcDatabase.getLastRow() - 1, 3)
  const numRows = sheet.getLastRow() - 2;
  const data = sheet.getSheetValues(3, 1, numRows, 1).map(item => {
    item.push('');
    upcs.map(upc => {
      if (upc[2] === item[0])
        item[1] += upc[0] + ','
    })
    return item;
  })

  const numCols = data[0].length;

  for (var row = 0, csv = "Name,Barcode\r\n"; row < numRows; row++)
  {
    for (var col = 0; col < numCols; col++)
    {
      if (data[row][col].toString().indexOf(",") != - 1)
        data[row][col] = "\"" + data[row][col] + "\"";
    }

    csv += (row < numRows - 1) ? data[row].join(",") + "\r\n" : data[row];
  }

  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV).downloadAsFile('inFlow_ProductDetails.csv');
}

/**
 * This function takes the array of data on the inFlowPick page and it creates a csv file that can be downloaded from the Browser.
 * 
 * @return Returns the csv text file that file be downloaded by the user.
 * @author Jarren Ralf
 */
function downloadInflowPickList()
{
  const sheet = SpreadsheetApp.getActive().getSheetByName("inFlowPick");
  const numRows = sheet.getLastRow() - 2;
  const numCols = sheet.getLastColumn() - 1;
  const data = sheet.getSheetValues(3, 1, numRows, numCols)

  for (var row = 0, csv = "OrderNumber,Customer,ItemName,ItemQuantity\r\n"; row < numRows; row++)
  {
    for (var col = 0; col < numCols; col++)
    {
      if (data[row][col].toString().indexOf(",") != - 1)
        data[row][col] = "\"" + data[row][col] + "\"";
    }

    csv += (row < numRows - 1) ? data[row].join(",") + "\r\n" : data[row];
  }

  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV).downloadAsFile('inFlow_SalesOrder.csv');
}

/**
 * This function takes the array of data on the Manual Counts page and it creates a csv file that can be downloaded from the Browser.
 * 
 * @return Returns the csv text file that file be downloaded by the user.
 * @author Jarren Ralf
 */
function downloadInflowStockLevels()
{
  const sheet = SpreadsheetApp.getActive().getSheetByName("Manual Counts");
  const data = [];
  var loc, qty, i;

  sheet.getSheetValues(4, 1, sheet.getLastRow() - 3, sheet.getLastColumn()).map(item => {
    loc = item[5].split('\n')
    qty = item[6].split('\n')

    if (loc.length === qty.length) // Make sure there is a location for every quantity and vice versa
      for (i = 0; i < loc.length; i++) // Loop through the number of inflow locations
        if (isNotBlank(loc[i]) && isNotBlank(qty)) // Do not add the data to the csv file if either the location or the quantity is blank
          data.push([item[0], loc[i], qty[i]])

  })

  const numRows = data.length;
  const numCols = data[0].length;

  for (var row = 0, csv = "Item,Location,Quantity\r\n"; row < numRows; row++)
  {
    for (var col = 0; col < numCols; col++)
    {
      if (data[row][col].toString().indexOf(",") != - 1)
        data[row][col] = "\"" + data[row][col] + "\"";
    }

    csv += (row < numRows - 1) ? data[row].join(",") + "\r\n" : data[row];
  }

  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV).downloadAsFile('inFlow_StockLevels.csv');
}

/**
 * This function formats the active sheet only by calling the applyFullSpreadsheetFormatting function with the active sheet as a parameter.
 * 
 * @author Jarren Ralf
 */
function formatActiveSheet()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheetArray = [spreadsheet.getActiveSheet()];
  applyFullSpreadsheetFormatting(spreadsheet, sheetArray);
}

/**
 * This function generates a list of items in the inFlow inventory system that based on the corresponding Adagio inventory values, should be picked and 
 * brought to Moncton street.
 * 
 * @author Jarren Ralf
 */
function generateSuggestedInflowPick()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const suggestedValuesSheet = spreadsheet.getSheetByName("Moncton's inFlow Item Quantities");
  const suggestInflowPickSheet = spreadsheet.getSheetByName("Suggested inFlowPick");
  const numSuggestedItems = suggestedValuesSheet.getLastRow() - 1;
  const suggestedValues = suggestedValuesSheet.getSheetValues(2, 1, numSuggestedItems, 3);
  const inventorySheet = spreadsheet.getSheetByName("INVENTORY");

  Utilities.parseCsv(DriveApp.getFilesByName("inFlow_StockLevels.csv").next().getBlob().getDataAsString()).map(item =>{
    if (item[0].split(' - ').length > 4) // If there are more than 4 "space-dash-space" strings within the inFlow description, then that item is recognized in Adagio 
    {
      for (var i = 0; i < numSuggestedItems; i++)
        if (suggestedValues[i][0] == item[0]) // The ith item of the suggested inFlowPick page was found in the inFlow csv, therefore break the for loop
          break;

      if (i === numSuggestedItems)
        suggestedValues.push([item[0], '', '']) // If there is an item in inFlow but not on the suggested inFlowPick page, then add it
    }
  })

  if (numSuggestedItems > numSuggestedItems) // Items from the inFlow csv have been added to the suggested inFlowPick page
  {
    suggestedValues.sort((a, b) => a[0].localeCompare(b[0])); // Sort the items by the description
    suggestedValuesSheet.getRange(2, 1, numSuggestedItems, 3).setValues(suggestedValues)
  }
  
  const output = inventorySheet.getSheetValues(8, 2, inventorySheet.getLastRow() - 7, 6).map(e => {

    if (isNotBlank(e[5]) && Number(e[2]) >= Number(e[5])) // Trites Inventory Column is not blank and the Adagio inventory is greater than or equal to inFlow inventory 
    {
      for (var i = 0; i < numSuggestedItems; i++)
      {
        if (suggestedValues[i][0] == e[0]) // Match the SKUs of the suggestValues list and the available inFlow inventory
        {
          const monctonStock = Number(e[2] - e[5]); // The stock levels in moncton street (Adagio - inFlow)

          if (Number(e[2]) <= Number(suggestedValues[i][1])) // If Moncton plus Trites less than or equal to the suggested quantity, then bring back everything from Trites to Moncton
            return [e[0], e[5], e[5], monctonStock, e[2]] // Bring back ALL trties stock
          else if (monctonStock < Number(suggestedValues[i][1])) // Moncton stock is less than the suggest amount for Moncton
          {
            const orderQty = Number(suggestedValues[i][1] - monctonStock);

            if (suggestedValues[i][2]) // If we try and pick this item in multiples of 'n' items, such as picking bait jars by the case and hence as multiples of 100 pcs
            {
              if (orderQty > Number(suggestedValues[i][2])) // Order quantity is greater then the number of items that we want to bring this SKU back in mutiples of
              {
                const suggestedAmount = Math.floor(orderQty/Number(suggestedValues[i][2]))*Number(suggestedValues[i][2])

                // If the suggestedAmount is greater than the Trites inventory, then bring back all of the Trites inventory, otherwise bring back the suggestedAmount
                return (suggestedAmount >= Number(e[5])) ? [e[0], e[5], e[5], monctonStock, e[2]] : [e[0], suggestedAmount, e[5], monctonStock, e[2]]
              }
            }
            else // If the orderQty is greater than the Trites inventory, then bring back all of the Trites inventory, otherwise bring back the orderQty
              return (orderQty >= Number(e[5])) ? [e[0], e[5], e[5], monctonStock, e[2]] : [e[0], orderQty, e[5], monctonStock, e[2]]
          }
        }
      }
    }

    return false // Not an available item at Trites
  }).filter(f => f) // Remove the unavailable items

  const numItems = output.length;
  const range = suggestInflowPickSheet.getRange(2, 1, suggestInflowPickSheet.getMaxRows(), 5).clearContent()
  
  if (numItems > 0)
  {
    output.sort((a,b) => a[3] - b[3]) // Sort list by the quantity in Moncton street because if Moncton has 0, then those items are the most important to pick from Trites
    range.offset(0, 0, output.length, 5).setValues(output)
  }
}

/**
 * This function gets the time that the particular item was last counted and calculates how long it has been since now, then displays that info to the user.
 * 
 * @param {Number} lastScannedTime : The time that an item was last scanned on the Manual Scan page or inputed on the Manual Counts page, represented as a number in milliseconds (Epoche Time).
 * @returns {String} 
 * @author Jarren Ralf
 */
function getCountedSinceString(lastScannedTime)
{
  if (isNotBlank(lastScannedTime))
  {
    const countedSince = (new Date().getTime() - lastScannedTime)/(1000) // This is in seconds

    if (countedSince < 60) // Number of seconds in 1 minute
      return Math.floor(countedSince) + ' seconds ago'
    else if (countedSince < 3600) // Number of seconds in 1 hour
      return (Math.floor(countedSince/60) === 1) ? Math.floor(countedSince/60) +  ' minute ago' : Math.floor(countedSince/60) +  ' minutes ago'
    else if (countedSince < 86400) // Number of seconds in 24 hours
    {
      const numHours = Math.floor(countedSince/3600);
      const numMinutes = Math.floor((countedSince - numHours*3600)/60);

      return (numHours === 1) ? numHours + ' hour ' + ((numMinutes === 0) ? 'ago' : (numMinutes === 1) ? numMinutes +  ' minute ago' : numMinutes +  ' minutes ago') : 
        numHours + ' hours ' + ((numMinutes === 0) ? 'ago' : (numMinutes === 1) ? numMinutes +  ' minute ago' : numMinutes +  ' minutes ago');
    }
    else // Greater than 24 hours
    {
      const numDays = Math.floor(countedSince/86400);
      const numHours = Math.floor((countedSince - numDays*86400)/3600);

      return (numDays === 1) ? numDays + ' day ' + ((numHours === 0) ? 'ago' : (numHours === 1) ? numHours + ' hour ago' : numHours + ' hours ago') : 
        numDays + ' days ' + ((numHours === 0) ? 'ago' : (numHours === 1) ? numHours + ' hour ago' : numHours + ' hours ago');
    }
  }
  else
    return '1 second ago'
}

/**
 * This function gets the items that have a negative inventory and sets it on the infoCounts page.
 * 
 * @author Jarren Ralf
 */
function getCounts()
{
  const startTime = new Date().getTime();
  const spreadsheet = SpreadsheetApp.getActive();
  const infoCountsSheet = spreadsheet.getSheetByName("InfoCounts");

  if (isRichmondSpreadsheet(spreadsheet))
  {
    var searchDataSheet = spreadsheet.getSheetByName("INVENTORY")
    var numHeaders = 7;
    var numCols = 3 // This changes the column reference for the current stock of the current store
  }
  else
  {
    var searchDataSheet = spreadsheet.getSheetByName("SearchData")
    var numHeaders = 1;
    var numCols = (isParksvilleSpreadsheet(spreadsheet)) ? 4 : 5; // This changes the column reference for the current stock of the current store
  }

  const storeCountsIndex = numCols - 1;
  const data = searchDataSheet.getSheetValues(numHeaders + 1, 2, searchDataSheet.getLastRow() - numHeaders, numCols);
  const output = data.filter(e => e[storeCountsIndex] < 0).map(f => [f[0], f[storeCountsIndex], '']) // All of the items with negative inventory
  const numItems = output.length;
  infoCountsSheet.getRange('A4:C').clearContent();
  infoCountsSheet.getRange(1, 2, 1, 2).setFormulas([['=COUNTA($C$4:$C$' + (numItems + 3) + ')','=' + numItems + '-Completed_InfoCounts']]);

  if (numItems > 0)
  {
    infoCountsSheet.getRange(4, 1, numItems, 3).setValues(output);
    applyFullRowFormatting(infoCountsSheet, 4, numItems, 3);
  }
    
  if (isRichmondSpreadsheet(spreadsheet))
    searchDataSheet.getRange(3, 3, 1, 7)
      .setValues([[ '=Remaining_InfoCounts&\" items on the infoCounts page that haven\'t been counted\"', null, null, null, 
                    '=Progress_InfoCounts', dateStamp(undefined, null, null, null, 'dd MMM HH:mm'), getRunTime(startTime)]]);
  else
    spreadsheet.getSheetByName("INVENTORY").getRange(7, 1, 1, 9)
      .setValues([[ '=Remaining_InfoCounts&\" items on the infoCounts page that haven\'t been counted\"', null, null, null, null, null, 
                    '=Progress_InfoCounts', dateStamp(undefined, null, null, null, 'dd MMM HH:mm'), getRunTime(startTime)]]);
}

/**
* This function calculates the day that New Years Day, Canada Day, Remembrance Day, and Christmas Day, is observed on for the giving year and month. 
*
* @param  {Number}  year The given year
* @param  {Number} month The given month
* @return {Number}   day The day of the Holiday for the particular year and month
* @author Jarren Ralf
*/
function getDay(year, month)
{
  const JANUARY  =  0;
  const JULY     =  6;
  const NOVEMBER = 10;
  const DECEMBER = 11;
  const SUNDAY   =  0;
  const SATURDAY =  6;
  
  if (month == JANUARY || month == JULY || month == DECEMBER) // New Years Day or Canada Day or Christmas Day
  {
    var holiday = (month == DECEMBER) ? new Date(year, month, 25) : new Date(year, month);
    var dayOfWeek = holiday.getDay();
    var day = (dayOfWeek == SATURDAY) ? holiday.getDate() + 2 : ( (dayOfWeek == SUNDAY) ? holiday.getDate() + 1 : holiday.getDate() ); // Rolls over to the following Monday
  }
  else if (month == NOVEMBER) // Remembrance Day
  {
    var holiday = new Date(year, month, 11);
    var dayOfWeek = holiday.getDay();
    var day = (dayOfWeek == SATURDAY) ? holiday.getDate() - 1 : ( (dayOfWeek == SUNDAY) ? holiday.getDate() + 1 : holiday.getDate() ); // Rolls back to Friday, or over to Monday
  }
  
  return day;
}

/**
* Gets the last row number based on a selected column range values
*
* @param {Object[][]} range Takes a 2d array of a single column's values
* @returns {Number} The last row number with a value. 
*/
function getLastRowSpecial(range)
{
  var rowNum = 0;
  var blank = false;
  const numRows = range.length;
  
  for (var row = 0; row < numRows; row++)
  {
    if(range[row][0] === "" && !blank)
    {
      rowNum = row;
      blank = true;
    }
    else if (isNotBlank(range[row][0]))
      blank = false;
  }
  return rowNum;
}

/**
* This function calculates what the nth Monday in the given month is for the given year. This function is used for determining the holidays in a given year.
* Victoria Day is an exception to the rule since it is defined to be the preceding Monday before May 25th. The fourth Boolean parameter handles this scenario.
*
* @param  {Number}              n : The nth Monday the user wants to be calculated
* @param  {Number}          month : The given month
* @param  {Number}           year : The given year
* @param  {Boolean} isVictoriaDay : Whether it is Victoria Day or not
* @return {Number} The day of the month that the nth Monday is on (or that Victoria Day is on)
* @author Jarren Ralf
*/
function getMonday(n, month, year, isVictoriaDay)
{
  const NUM_DAYS_IN_WEEK = 7;
  var firstDayOfMonth = new Date(year, month).getDay();
  
  if (isVictoriaDay)
    n = (firstDayOfMonth % (NUM_DAYS_IN_WEEK - 1) < 2) ? 4 : 3; // Corresponds to the Monday preceding May 25th 
  
  return ((NUM_DAYS_IN_WEEK - firstDayOfMonth + 1) % NUM_DAYS_IN_WEEK) + NUM_DAYS_IN_WEEK*n - 6;
}

/**
* This function gets the physical inventory counts that have been recorded on a set of sheets.
*
* @param  {Sheet[]}    sheets : The sheets that the data is coming from
* @param  {Date}         DATE : The current formatted date
* @param  {String[][]} recentCounts : The recent counts including yesterday and today
* @return {String[][]} The list of SKUs, Descriptions, Vendors, Categories, UoMs, Quantities, Sheets and Dates
* @author Jarren Ralf
*/
function getPhysicalCounted_CountLog(sheets, DATE, recentCounts)
{
  const  DATA_START_ROW = 4;
  const  BACK_ORDER_COL = 5;
  const currentStockCol = 2;
  const numSheets = sheets.length;
  const numRecentCounts = recentCounts.length;
  var sku, numRows, sheetName, descripCol, numCols, quantityColIndex, data, countedItems = [];

  for (var s = 0; s < numSheets; s++)
  {
    numRows = sheets[s].getLastRow() - DATA_START_ROW + 1;
    sheetName = sheets[s].getSheetName();

    if (numRows > 0) // Check if there is any data
    {
      if (sheetName === "Order")
      {
        descripCol = 5;
        numCols = 6;
      }
      else if (sheetName === "Shipped")
      {
        descripCol = 5;
        numCols = 4;
      }
      else // InfoCounts or Manual Counts
      {
        descripCol = 1;
        numCols = 3;
      }

      quantityColIndex = (numCols === 3) ? 2 : 3;
      data = sheets[s].getSheetValues(DATA_START_ROW, descripCol, numRows, numCols);

      for (var i = 0; i < numRows; i++)
      { 
        // Check if the entry is a number, then if Order or Shipped sheet, then check if actual and current stock are different, and don't include Back Orders if on the Order sheet
        if (!(isNaN(parseInt(data[i][quantityColIndex]))) && ((descripCol === 1) || ((data[i][quantityColIndex] != data[i][currentStockCol]) && (numCols != 6 || data[i][BACK_ORDER_COL] != "B/O"))))
        {
          for (var j = 0; j < numRecentCounts; j++)
          {
            if (recentCounts[j][1] == data[i][0]) // The item hasn't been counted in the last 2 days
              break;
          }

          if (j === numRecentCounts) // The item was not found in the recent counts, therefore add it to the log
          {
            sku = data[i][0].split(' - ').pop(); 

            if (sku != data[i][0]) // Only log counted items that appear to be skus (based on the formatting of the string " - ")
              countedItems.push([sku, data[i][0], sheetName, DATE]);
          }
        }
      }
    }
  }

  return countedItems;
}

/**
* This function calculated and returns the runtime of a particular script.
*
* @param  {Number} startTime : The start time that the script began running at represented by a number in milliseconds
* @return {String}  runTime  : The runtime of the script represented by a number followed by the unit abbreviation for seconds.
* @author Jarren Ralf
*/
function getRunTime(startTime)
{
  return (new Date().getTime() - startTime)/1000 + ' s';
}

/**
 * This function take a list of selected items and it sorts the items with their corresponding hoochies and returns an object with the corresponding hoochie sku prefix
 * as the object key or else adds the non-hoochie item to be accessed with the key: "other".
 * 
 * @param {String[][]} hoochies     : The list of hoochies and any corresponding data.
 * @param {Number} descriptionIndex : The index number of the column that contains the google description.
 * @return {Object} Returns an object containing sets of multi-arrays with each
 * @author Jarren Ralf
 */
function groupHoochieTypes(hoochies, descriptionIndex)
{
  const groupedHoochies = {'16060005': [], '16010005': [], '16050005': [], '16020000': [], '16020010': [], 
                           '16060065': [], '16060010': [], '16070000': [], '16075300': [], '16070975': [], 
                           '16030000': [], '16060175': [], '16200030': [], '16200000': [], '16200025': [], 
                           '16200061': [], '16200065': [], '16200021': [], '16200022': [],    "other": []};
  
  hoochies.map(hoochie => {
    if (groupedHoochies[hoochie[descriptionIndex].split(" - ").pop().toString().substring(0, 8)] == null)
      groupedHoochies["other"].push(hoochie)
    else
      groupedHoochies[hoochie[descriptionIndex].split(" - ").pop().toString().substring(0, 8) || "other"].push(hoochie)
  })

  return groupedHoochies;
}

/**
* This function inserts the Carrier Not Assigned banner on the shipped sheet.
*
* @author Jarren Ralf
*/
function insertCarrierNotAssignedBanner()
{
  const BANNER_COL = 0;
  const STATUS_COL = 9;
  const sheet = SpreadsheetApp.getActive().getSheetByName("Shipped");
  const values = sheet.getDataRange().getValues();
  const LAST_COL = sheet.getLastColumn();
  const numRows = sheet.getLastRow() - 1;
  const bannerRow = [];

  conditional: if (true)
  {
    for (var i = numRows; i >= 4; i--)
    {
      if (values[i][BANNER_COL] === 'Carrier Not Assigned') // Carrier Not Assigned banner was found!
      {
        SpreadsheetApp.getUi().alert('Carrier Not Assigned banner already exists.')
        break conditional; // Break the conditional code block because we don't want to insert a Carrier Not Assigned banner
      } 
      else if (bannerRow.length === 0 && isNotBlank(values[i][STATUS_COL]) && values[i][STATUS_COL] !== 'Carrier Not Assigned' && values[i][STATUS_COL] !== 'Order From Distributor' && 
              values[i][STATUS_COL] !== 'Discontinued' && values[i][STATUS_COL] !== 'Item Not Available' && values[i][STATUS_COL] !== 'Back to Shipped')
        bannerRow.push(i + 1); // Determine which row the banner should go
    }

    sheet.insertRowsAfter(bannerRow[0], 1).setRowHeight(bannerRow[0] + 1, 40).getRange(bannerRow[0] + 1, 1, 1, LAST_COL).clearDataValidations()
      .setBackgrounds([[...new Array(LAST_COL - 1).fill('#6d9eeb'), 'white']]).setFontColors([[...new Array(LAST_COL - 2).fill('white'), '#6d9eeb', 'white']])
      .setFontFamily('Arial').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('left').setNumberFormat('@').setVerticalAlignment('middle')
      .setValues([['Carrier Not Assigned', ...new Array(LAST_COL - 3).fill(null), 'via', '']])
      .offset(0, 0, 1, LAST_COL - 1).setBorder(true, true, true, true, false, null)
      .offset(0, 0, 1, LAST_COL - 2).merge();
  }
}

/**
 * This function checks if a given value is precisely a non-blank string.
 * 
 * @param  {String}  value : A given string.
 * @return {Boolean} Returns a boolean based on whether an inputted string is not-blank or not.
 * @author Jarren Ralf
 */
function isNotBlank(value)
{
  return value !== '';
}

/**
 * This function searches the UPC Database for the upc value (the barcode that was scanned) and puts it on the current page.
 * If scanned on the Received page, then it takes the user to that item on the Shipped sheet in order for them to receive it.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}       sheet     : The active sheet.
 * @param    {String}     sheetName  : The name of the active sheet.
 * @author Jarren Ralf
 */
function itemScan(e, spreadsheet, sheet, sheetName)
{
  const range    = e.range;           // The range of the edited cell
  const rowStart = range.rowStart;    // The first row of the edited range
  const colStart = range.columnStart; // The first column of the edited range
  const colEnd   = range.columnEnd;   // The last column of the edited range
  const value    = range.getValue();  // The value of the edited cell

  if (rowStart == range.rowEnd && range.columnEnd && rowStart == 1 && colStart == 1 && (colEnd == 1 || colEnd == 3)) // Make sure that only the top left cell is being editted
  {
    if (colEnd == 3)
      range.merge()

    if (isNotBlank(value)) 
    {
      const upcDatabase = spreadsheet.getSheetByName("UPC Database").getDataRange().getValues();
      const upcCode = range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setFontFamily("Libre Barcode 128")
        .setVerticalAlignment("middle").setHorizontalAlignment("center").setFontColor("black").setFontSize(35)
        .setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID)
        .getValue();
      const numUpcs = upcDatabase.length - 1;
      
      const lastRow = sheet.getLastRow();
      const row = lastRow + 1; 

      loop: for (var i = numUpcs; i >= 1; i--)
      {
        if (upcDatabase[i][0] == upcCode)
        {
          if (sheetName === "Shipped")
          {
            sheet.getRange(row, 4, 1, 7).setValues([[upcDatabase[i][1], upcDatabase[i][2], null, upcDatabase[i][3], '', '', 'Carrier Not Assigned']])
            applyFullRowFormatting(sheet, row, 1, 11)
            sheet.getRange(row, 3).activate();
            dateStamp(row, 11);
            spreadsheet.toast('Item Added to Bottom of Page')
          }
          else if (sheetName === 'ItemsToRichmond')
          {
            sheet.getRange(row, 3, 1, 2).setValues([[upcDatabase[i][1], upcDatabase[i][2]]])
            applyFullRowFormatting(sheet, row, 1)
            sheet.getRange(row, 6).activate();
            spreadsheet.toast('Item Added to Bottom of Page')
          }
          else if (sheetName === "Order")
          {
            const numRows_OrderSheet = lastRow - 3;
            const orderPageValues = sheet.getSheetValues(4, 5, numRows_OrderSheet, 1);

            for (var j = 0; j < numRows_OrderSheet; j++)
            {
              if (orderPageValues[j][0] == upcDatabase[i][2])
              {
                sheet.getRange(j + 4, 3, 1, 4).activate();
                spreadsheet.toast('Item Already on Order Page')
                break loop;
              }
            }

            if (j === numRows_OrderSheet) // Item not found on order page
            {
              sheet.getRange(row, 4, 1, 4).setValues([[upcDatabase[i][1], upcDatabase[i][2], null, upcDatabase[i][3]]])
              applyFullRowFormatting(sheet, row, 1, 11)
              sheet.getRange(row, 3).activate();
              spreadsheet.toast('Item Added to Bottom of Page')
            }
          }
          else if (sheetName === "Received")
          {
            const shippedPage = spreadsheet.getSheetByName("Shipped");
            const numRows_ShippedSheet = lastRow - 3;
            const shippedPageValues = shippedPage.getSheetValues(4, 5, numRows_ShippedSheet, 1);

            for (var j = 0; j < numRows_ShippedSheet; j++)
            {
              if (shippedPageValues[j][0] == upcDatabase[i][2])
              {
                shippedPage.getRange(j + 4, 10).activate();
                spreadsheet.toast('Please Change Shipment Status to Rec\'d ...')
                break loop;
              }
            }

            spreadsheet.toast('Item Not Found On Shipped Page')
            break loop;
          }

          dateStamp(row, 1);
          break;
        }
      }

      if (i == 0)
      {
        range.setValue('Barcode Not Found');
        spreadsheet.toast('Barcode Not Found')
      }
      else
        range.setValue('Scan Here');
    }
    else // The user has clicked delete
      range.setValue('Scan Here');
  }
}

/**
 * This function checks if every value in the import multi-array is blank, which means that the user has
 * highlighted and deleted all of the data.
 * 
 * @param {Object[][]} values : The import data
 * @return {Boolean} Whether the import data is deleted or not
 * @author Jarren Ralf
 */
function isEveryValueBlank(values)
{
  return values.every(arr => arr.every(val => val == '') === true);
}

/**
* This function checks if the given input is a number or not.
*
* @param {Object} num The inputted argument, assumed to be a number.
* @return {Boolean} Returns a boolean reporting whether the input paramater is a number or not
* @author Jarren Ralf
*/
function isNumber(num)
{
  return !(isNaN(Number(num)));
}

/**
* This function checks if the current spreadsheet being used is the Parksville spreadsheet or not.
*
* @param {Spreadsheet} spreadsheet : The active spreadsheet.
* @return {Boolean} Returns a boolean reporting whether the second word of the spreadsheet name is "Parsville" or not.
* @author Jarren Ralf
*/
function isParksvilleSpreadsheet(spreadsheet)
{
  return spreadsheet.getName().split(" ")[1] === "Parksville";
}

/**
* This function checks if today's date is a stat holiday or not.
*
* @param {Date} today : Today's date
* @return {Boolean} Returns a true boolean if today is not a stat and false otherwise.
* @author Jarren Ralf
*/
function isNotStatHoliday(today)
{
  today = today.getTime();
  const JAN =  0, FEB =  1, MAY =  4, JUL =  6, AUG =  7, SEP =  8, OCT =  9, NOV = 10, DEC = 11;
  const YEAR = new Date().getFullYear(); // An integer corresponding to today's year
  const ONE_DAY = 24*60*60*1000;
  var MMM, DD;
  [MMM, DD] = calculateGoodFriday(YEAR);

  const statHolidays = [new Date(YEAR, JAN, getDay(YEAR, JAN)),          // New Year's Day
                        new Date(YEAR, FEB, getMonday(3, FEB, YEAR)),    // Family Day
                        new Date(YEAR, MMM, DD),                         // Good Friday
                        new Date(YEAR, MAY, getMonday(0, MAY, YEAR, 1)), // Victoria Day
                        new Date(YEAR, JUL, getDay(YEAR, JUL)),          // Canada Day
                        new Date(YEAR, AUG, getMonday(1, AUG, YEAR)),    // BC Day
                        new Date(YEAR, SEP, getMonday(1, SEP, YEAR)),    // Labour Day
                        new Date(YEAR, OCT, getMonday(2, OCT, YEAR)),    // Thanksgiving Day
                        new Date(YEAR, NOV, getDay(YEAR, NOV)),          // Remembrance Day
                        new Date(YEAR, DEC, getDay(YEAR, DEC))];         // Christmas Day

  const isStat = statHolidays.reduce((a, holiday) => {if (0 < today - holiday && today - holiday < ONE_DAY) return true})

  return !isStat;
}

/**
* This function checks if the current spreadsheet being used is the Parksville spreadsheet or not.
*
* @param {Spreadsheet} spreadsheet : The active spreadsheet.
* @return {Boolean} Returns a boolean reporting whether the second word of the spreadsheet name is "Parsville" or not.
* @author Jarren Ralf
*/
function isRichmondSpreadsheet(spreadsheet)
{
  return spreadsheet.getName().split(" ")[1] === "Richmond";
}

/**
 * This function returns true if the presented number is a UPC-A, false otherwise.
 * 
 * @param {Number} upcNumber : The UPC-A number
 * @returns Whether the given value is a UPC-A or not
 * @author Jarren Ralf
 */
function isUPC_A(upcNumber)
{
  for (var i = 0, sum = 0, upc = upcNumber.toString(); i < upc.length - 1; i++)
    sum += (i % 2 === 0) ? Number(upc[i])*3 : Number(upc[i])

  return upc.endsWith(Math.ceil(sum/10)*10 - sum)
}

/**
 * This function returns true if the presented number is a EAN_13, false otherwise.
 * 
 * @param {Number} upcNumber : The EAN_13 number
 * @returns Whether the given value is a EAN_13 or not
 * @author Jarren Ralf
 */
function isEAN_13(upcNumber)
{
  for (var i = 0, sum = 0, upc = upcNumber.toString(); i < upc.length - 1; i++)
    sum += (i % 2 === 0) ? Number(upc[i]) : Number(upc[i])*3

  return upc.endsWith(Math.ceil(sum/10)*10 - sum) && upc.length === 13
}

/**
 * This function is run on a trigger between 11 pm and 12 am everyday. The if statement subsequently only runs the countLog function on working days.
 * 
 * @author Jarren Ralf
 */
function logCountsOnWorkdays()
{
  const SUNDAY = 0;
  const today = new Date();
  const day = today.getDay();

  if (day !== SUNDAY && isNotStatHoliday(today))
    countLog();
}

/**
* This function moves all of the selected values on the item search page to the Manual Counts page
*
* @author Jarren Ralf
*/
function manualCounts()
{
  const QTY_COL = 3;
  const NUM_COLS = 3;
  
  var manualCountsSheet = SpreadsheetApp.getActive().getSheetByName("Manual Counts");
  var lastRow = manualCountsSheet.getLastRow();
  var startRow = (lastRow < 3) ? 4 : lastRow + 1;

  copySelectedValues(manualCountsSheet, startRow, NUM_COLS, QTY_COL);
}

/**
* This function moves all of the selected values on the info counts page to the Manual Counts page
*
* @author Jarren Ralf
*/
function manualCounts_FromInfoCounts()
{
  const QTY_COL = 3;
  const NUM_COLS = 3;
  
  var manualCountsSheet = SpreadsheetApp.getActive().getSheetByName("Manual Counts");
  var lastRow = manualCountsSheet.getLastRow();
  var startRow = (lastRow < 3) ? 4 : lastRow + 1;

  copySelectedValues(manualCountsSheet, startRow, NUM_COLS, QTY_COL, true);
}

/**
 * This function watches two cells and if the left one is edited then it searches the UPC Database for the upc value (the barcode that was scanned).
 * It then checks if the item is on the manual counts page and stores the relevant data in the left cell. If the right cell is edited, then the function
 * uses the data in the left cell and moves the item over to the manual counts page with the updated quantity.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf
 */
function manualScan(e, spreadsheet, sheet)
{
  const manualCountsPage = spreadsheet.getSheetByName("Manual Counts");
  const barcodeInputRange = e.range;
  const row = barcodeInputRange.rowStart;
  const col = barcodeInputRange.columnStart;

  if (row === barcodeInputRange.rowEnd && col === barcodeInputRange.columnEnd) // Single cell was edited
  {
    if (row !== 1) // This is scanning of a barcode or entering a quantity
    {
      if (manualCountsPage.getRange(3, 7).isChecked()) // Manual Scanner is in "Add-One" mode
      {
        const upcCode = barcodeInputRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP) // Wrap strategy for the cell
          .setFontFamily("Arial").setFontColor("black").setFontSize(25)                     // Set the font parameters
          .setVerticalAlignment("middle").setHorizontalAlignment("center")                  // Set the alignment parameters
          .getValue();
        
        if (isNotBlank(upcCode)) // The user may have hit the delete key
        {
          const upcString = upcCode.toString().toLowerCase().split(" ")
          const lastRow = manualCountsPage.getLastRow();
          const upcDatabase = spreadsheet.getSheetByName("UPC Database").getDataRange().getValues();

          if (upcString[0] == 'clear')
          {
            var item = e.oldValue;

            if (item === "UPC Code has been added to the database temporarily." || item === "UPC Code has been added to the unmarry list." || item === undefined)
              item = barcodeInputRange.offset(0, -1).getValue();

            item = item.split('\n');
            
            if (item[1].split(' ')[0] === 'will') // The item was not found on the manual counts page
              sheet.getRange(2, 1, 1, 2).setValues([['Item Not Found on Manual Counts page.', '']]);
            else
            {
              manualCountsPage.getRange(item[2], 3, 1, 5).setNumberFormat('@').setValues([['', '', '', '', '']])
              sheet.getRange(2, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Manual Counts page at line :\n' + item[2] 
                                                              + '\nCurrent Stock :\n' + item[4] 
                                                              + '\nCurrent Manual Count :\n\nCurrent Running Sum :\n',
                                                              '']]);
            }
          }
          else if (upcString[0] == 'undo')
          {
            var item = e.oldValue;

            if (item === "UPC Code has been added to the database temporarily." || item === "UPC Code has been added to the unmarry list." || item === undefined)
              item = barcodeInputRange.offset(0, -1).getValue();

            item = item.split('\n');

            if (item[1].split(' ')[0] === 'will') // The item was not found on the manual counts page
              sheet.getRange(2, 1, 1, 2).setValues([['Item Not Found on Manual Counts page.', '']]);
            else
            {
              var range = manualCountsPage.getRange(item[2], 3, 1, 5);
              var manualCountsValues = range.getValues()
              
              if (isNotBlank(manualCountsValues[0][1]))
              {
                var runningSumSplit = manualCountsValues[0][1].split(' ');

                if (runningSumSplit.length === 1)
                {
                  range.setNumberFormat('@').setValues([['', '', '', '', '']])
                  manualCountsValues[0][0] = ''
                  manualCountsValues[0][1] = ''
                  manualCountsValues[0][2] = ''
                  var countedSince = ''
                }
                else if (runningSumSplit[runningSumSplit.length - 2] === '+')
                {
                  manualCountsValues[0][0] -= Number(runningSumSplit[runningSumSplit.length - 1])
                  runningSumSplit.pop();
                  runningSumSplit.pop();
                  manualCountsValues[0][1] = runningSumSplit.join(' ')
                  manualCountsValues[0][2] = new Date().getTime()
                  var countedSince = getCountedSinceString(manualCountsValues[0][2])
                }
                else if (runningSumSplit[runningSumSplit.length - 2] === '-')
                {
                  manualCountsValues[0][0] += Number(runningSumSplit[runningSumSplit.length - 1])
                  runningSumSplit.pop();
                  runningSumSplit.pop();
                  manualCountsValues[0][1] = runningSumSplit.join(' ')
                  manualCountsValues[0][2] = new Date().getTime()
                  var countedSince = getCountedSinceString(manualCountsValues[0][2])
                }
              }

              manualCountsValues[0][2] = new Date().getTime()
              range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues(manualCountsValues)
              sheet.getRange(2, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Manual Counts page at line :\n' + (item[2]) 
                                                              + '\nCurrent Stock :\n' + item[4]
                                                              + '\nCurrent Manual Count :\n' + manualCountsValues[0][0] 
                                                              + '\nCurrent Running Sum :\n' + manualCountsValues[0][1]
                                                              + '\nLast Counted :\n' + countedSince,
                                                              '']]);
            }
          }
          else if (upcString[0] == 'uuu')
          {
            if (upcString[1] > 100000)
            {
              var item = e.oldValue;

              if (item === "UPC Code has been added to the database temporarily." || item === "UPC Code has been added to the unmarry list." || item === undefined)
                item = sheet.getRange(2, 1).getValue();

              item = item.split('\n');

              const unmarryUpcSheet = spreadsheet.getSheetByName("UPCs to Unmarry");
              unmarryUpcSheet.getRange(unmarryUpcSheet.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[upcString[1], item[0]]]);
              barcodeInputRange.setValue('UPC Code has been added to the unmarry list.')
              sheet.getRange(2, 1).activate();
            }
            else
              barcodeInputRange.setValue('Please enter a valid UPC Code to unmarry.')
          }
          else if (upcString[0] == 'mmm')
          {
            if (upcString[1] > 100000)
            {
              var item = e.oldValue;

              if (item === "UPC Code has been added to the database temporarily." || item === "UPC Code has been added to the unmarry list." || item === undefined)
                item = sheet.getRange(2, 1).getValue();

              item = item.split('\n');

              const marriedItem = item[0].split(' - ');
              const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
              const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");
              const sku = marriedItem.pop()
              const uom = marriedItem.pop()
              manAddedUPCsSheet.getRange(manAddedUPCsSheet.getLastRow() + 1, 1, 1, 4).setNumberFormat('@').setValues([[sku, upcString[1], uom, item[0]]]);
              upcDatabaseSheet.getRange(upcDatabaseSheet.getLastRow() + 1, 1, 1, 4).setNumberFormat('@').setValues([[upcString[1], uom, item[0], item[4]]]); 
              barcodeInputRange.setValue('UPC Code has been added to the database temporarily.')
              sheet.getRange(2, 1).activate();
            }
            else
              barcodeInputRange.setValue('Please enter a valid UPC Code to marry.')
          }
          else if (upcCode <= 100000) // In this case, variable name: upcCode is assumed to be the quantity
          {
            var item = e.oldValue;

            if (item === "UPC Code has been added to the database temporarily." || item === "UPC Code has been added to the unmarry list." || item === undefined)
              item = barcodeInputRange.offset(0, -1).getValue();

            item = item.split('\n');

            if (item[1].split(' ')[0] === 'will') // The item was not found on the manual counts page
            {
              manualCountsPage.getRange(item[2], 1, 1, 7).setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], item[4], upcCode, '\'' + upcCode, new Date().getTime(), '', '']]);
              applyFullRowFormatting(manualCountsPage, item[2], 1, 7)
              sheet.getRange(2, 1, 1, 2).setValues([[item[0]  + '\nwas added to the Manual Counts page at line :\n' + item[2] 
                                                              + '\nCurrent Stock :\n' + item[4] 
                                                              + '\nCurrent Manual Count :\n' + upcCode,
                                                              '']]);
            }
            else
            {
              const range = manualCountsPage.getRange(item[2], 3, 1, 3);
              const manualCountsValues = range.getValues()
              manualCountsValues[0][2] = new Date().getTime()
              manualCountsValues[0][1] = (isNotBlank(manualCountsValues[0][1])) ? ((Math.sign(upcCode) === 1 || Math.sign(upcCode) === 0)  ? 
                                                                                  String(manualCountsValues[0][1]) + ' \+ ' + String(   upcCode)  : 
                                                                                  String(manualCountsValues[0][1]) + ' \- ' + String(-1*upcCode)) :
                                                                                    ((isNotBlank(manualCountsValues[0][0])) ? 
                                                                                      String(manualCountsValues[0][0]) + ' \+ ' + String(upcCode) : 
                                                                                      String(upcCode));
              manualCountsValues[0][0] = Number(manualCountsValues[0][0]) + Number(upcCode);
              range.setNumberFormats([['#.#', '@', '#']]).setValues(manualCountsValues)
              sheet.getRange(2, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Manual Counts page at line :\n' + item[2] 
                                                              + '\nCurrent Stock :\n' + item[4] 
                                                              + '\nCurrent Manual Count :\n' + manualCountsValues[0][0] 
                                                              + '\nCurrent Running Sum :\n' + manualCountsValues[0][1]
                                                              + '\nLast Counted :\n' + getCountedSinceString(manualCountsValues[0][2]),
                                                              '']]);
            }
          }
          else // Add one
          {
            if (lastRow <= 3) // There are no items on the manual counts page
            {
              const numUpcs = upcDatabase.length - 1;

              for (var i = numUpcs; i >= 1; i--) // Loop through the UPC values
              {
                if (upcDatabase[i][0] == upcCode) // UPC found
                {
                  const row = lastRow + 1;
                  manualCountsPage.getRange(row, 1, 1, 5).setNumberFormats([['@', '@', '#.#', '@', '#']]).setValues([[upcDatabase[i][2], upcDatabase[i][3], 1, '\'' + String(1), new Date().getTime()]])
                  applyFullRowFormatting(manualCountsPage, row, 1, 7)
                  sheet.getRange(2, 1, 1, 2).setValues([[upcDatabase[i][2]  + '\nwas added to the Manual Counts page at line :\n' + row 
                                                                            + '\nCurrent Stock :\n' + upcDatabase[i][3]
                                                                            + '\nCurrent Manual Count :\n1',
                                                                            '']]);
                }
              }
            }
            else // There are existing items on the manual counts page
            {
              const row = lastRow + 1;
              const numRows = row - 3;
              const manualCountsValues = manualCountsPage.getSheetValues(4, 1, numRows, 5);
              const numUpcs = upcDatabase.length - 1;

              for (var i = numUpcs; i >= 1; i--) // Loop through the UPC values
              {
                if (upcDatabase[i][0] == upcCode)
                {
                  for (var j = 0; j < numRows; j++) // Loop through the manual counts page
                  {
                    if (manualCountsValues[j][0] === upcDatabase[i][2]) // The description matches
                    {
                      if (isNotBlank(manualCountsValues[j][4]))
                      {
                        const updatedCount = Number(manualCountsValues[j][2]) + 1;
                        const countedSince = getCountedSinceString(manualCountsValues[j][4])
                        const runningSum = (isNotBlank(manualCountsValues[j][3])) ? (String(manualCountsValues[j][3]) + ' \+ 1') : ((isNotBlank(manualCountsValues[j][2])) ? 
                                                                                                                                    String(manualCountsValues[j][2]) + ' \+ 1' : 
                                                                                                                                    String(1));
                        manualCountsPage.getRange(j + 4, 3, 1, 3).setNumberFormats([['#.#', '@', '#']]).setValues([[updatedCount, runningSum, new Date().getTime()]])
                        sheet.getRange(2, 1, 1, 2).setValues([[manualCountsValues[j][0] + '\nwas found on the Manual Counts page at line :\n' + (j + 4) 
                                                                                        + '\nCurrent Stock :\n' + manualCountsValues[j][1]
                                                                                        + '\nCurrent Manual Count :\n' + updatedCount 
                                                                                        + '\nCurrent Running Sum :\n' + runningSum
                                                                                        + '\nLast Counted :\n' + countedSince,
                                                                                        '']]);
                      }
                      else
                      {
                        manualCountsPage.getRange(j + 4, 3, 1, 3).setNumberFormats([['#.#', '@', '#']]).setValues([[1, '1', new Date().getTime()]])
                        sheet.getRange(2, 1, 1, 2).setValues([[manualCountsValues[j][0] + '\nwas found on the Manual Counts page at line :\n' + (j + 4) 
                                                                                        + '\nCurrent Stock :\n' + manualCountsValues[j][1]
                                                                                        + '\nCurrent Manual Count :\n1',
                                                                                        '']]);
                      }
                      break; // Item was found on the manual counts page, therefore stop searching
                    } 
                  }

                  if (j === numRows) // Item was not found on the manual counts page
                  {
                    manualCountsPage.getRange(row, 1, 1, 5).setNumberFormats([['@', '@', '#.#', '@', '#']])
                      .setValues([[upcDatabase[i][2], upcDatabase[i][3], 1, '\'' + String(1), new Date().getTime()]])
                    applyFullRowFormatting(manualCountsPage, row, 1, 7)
                    sheet.getRange(2, 1, 1, 2).setValues([[upcDatabase[i][2]  + '\nwas added to the Manual Counts page at line :\n' + row 
                                                                              + '\nCurrent Stock :\n' + upcDatabase[i][3]
                                                                              + '\nCurrent Manual Count :\n1',
                                                                              '']]);
                  }

                  break;
                }
              }
            }

            sheet.deleteColumn(2) // This line is added here so that a user who is using a tablet can continually scan without problems

            if (i === 0)
            {
              if (upcCode.toString().length > 25)
                sheet.getRange(2, 1, 1, 2).setNumberFormats([['@', '#.#']]).setValues([['Barcode is Not Found.', '']]);
              else
                sheet.getRange(2, 1, 1, 2).setNumberFormats([['@', '#.#']]).setValues([['Barcode:\n\n' + upcCode + '\n\n is NOT FOUND.', '']]);
            }
            else
              sheet.getRange(2, 2).setValue('').setNumberFormat('#.#').activate(); // Since the second column was deleted, the number format needs to be restored

            sheet.setColumnWidth(2, 350).getRange(2, 1).activate()
          }
        }
      }
      else
      {
        if (barcodeInputRange.columnEnd === 1) // Barcode is scanned
        {
          const upcCode = barcodeInputRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP) // Wrap strategy for the cell
            .setFontFamily("Arial").setFontColor("black").setFontSize(25)                     // Set the font parameters
            .setVerticalAlignment("middle").setHorizontalAlignment("center")                  // Set the alignment parameters
            .getValue();

          if (isNotBlank(upcCode)) // The user may have hit the delete key
          {
            const lastRow = manualCountsPage.getLastRow();
            const upcDatabase = spreadsheet.getSheetByName("UPC Database").getDataRange().getValues();
            const numUpcs = upcDatabase.length - 1;

            if (lastRow <= 3) // There are no items on the manual counts page
            {
              for (var i = numUpcs; i >= 1; i--) // Loop through the UPC values
              {
                if (upcDatabase[i][0] == upcCode) // UPC found
                {
                  barcodeInputRange.setValue(upcDatabase[i][2] + '\nwill be added to the Manual Counts page at line :\n' + 4 + '\nCurrent Stock :\n' + upcDatabase[i][3]);
                  break; // Item was found, therefore stop searching
                }
              }
            }
            else // There are existing items on the manual counts page
            {
              const row = lastRow + 1;
              const numRows = row - 3;
              const manualCountsValues = manualCountsPage.getSheetValues(4, 1, numRows, 5);

              for (var i = numUpcs; i >= 1; i--) // Loop through the UPC values
              {
                if (upcDatabase[i][0] == upcCode)
                {
                  for (var j = 0; j < numRows; j++) // Loop through the manual counts page
                  {
                    if (manualCountsValues[j][0] === upcDatabase[i][2]) // The description matches
                    {
                      const countedSince = getCountedSinceString(manualCountsValues[j][4])
                        
                      barcodeInputRange.setValue(upcDatabase[i][2]  + '\nwas found on the Manual Counts page at line :\n' + (j + 4) 
                                                                    + '\nCurrent Stock :\n' + upcDatabase[i][3] 
                                                                    + '\nCurrent Manual Count :\n' + manualCountsValues[j][2] 
                                                                    + '\nCurrent Running Sum :\n' + manualCountsValues[j][3]
                                                                    + '\nLast Counted :\n' + countedSince);
                      break; // Item was found on the manual counts page, therefore stop searching
                    }
                  }

                  if (j === manualCountsValues.length) // Item was not found on the manual counts page
                    barcodeInputRange.setValue(upcDatabase[i][2] + '\nwill be added to the Manual Counts page at line :\n' + row + '\nCurrent Stock :\n' + upcDatabase[i][3]);

                  break;
                }
              }
            }

            if (i === 0)
            {
              if (upcCode.toString().length > 25)
                sheet.getRange(2, 1, 1, 2).setValues([['Barcode is Not Found.', '']]);
              else
                sheet.getRange(2, 1, 1, 2).setValues([['Barcode:\n\n' + upcCode + '\n\n is NOT FOUND.', '']]);

              sheet.getRange(2, 1).activate()
            }
            else
              sheet.getRange(2, 2).setValue('').activate();
          }
        }
        else if (barcodeInputRange.columnStart !== 1) // Quantity is entered
        {
          const quantity = barcodeInputRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP) // Wrap strategy for the cell
            .setFontFamily("Arial").setFontColor("black").setFontSize(25)                      // Set the font parameters
            .setVerticalAlignment("middle").setHorizontalAlignment("center")                   // Set the alignment parameters
            .getValue();

          if (isNotBlank(quantity)) // The user may have hit the delete key
          {
            const item = sheet.getRange(2, 1).getValue().split('\n');    // The information from the left cell that is used to move the item to the manual counts page
            const quantity_String = quantity.toString().toLowerCase();
            const quantity_String_Split = quantity_String.split(' ');

            if (quantity_String === 'clear')
            {
              manualCountsPage.getRange(item[2], 3, 1, 5).setNumberFormat('@').setValues([['', '', '', '', '']])
              sheet.getRange(2, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Manual Counts page at line :\n' + item[2] 
                                                              + '\nCurrent Stock :\n' + item[4] 
                                                              + '\nCurrent Manual Count :\n\nCurrent Running Sum :\n',
                                                              '']]);
            }
            else if (quantity_String_Split[0] === 'uuu') // Unmarry upc
            {
              const upc = quantity_String_Split[1];

              if (upc > 100000)
              {
                const unmarryUpcSheet = spreadsheet.getSheetByName("UPCs to Unmarry");
                unmarryUpcSheet.getRange(unmarryUpcSheet.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[upc, item[0]]]);
                barcodeInputRange.setValue('UPC Code has been added to the unmarry list.')
                sheet.getRange(2, 1).activate();
              }
              else
                barcodeInputRange.setValue('Please enter a valid UPC Code to unmarry.')
            }
            else if (quantity_String_Split[0] === 'mmm') // Marry upc
            {
              const upc = quantity_String_Split[1];

              if (upc > 100000)
              {
                const marriedItem = item[0].split(' - ');
                const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
                const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");
                const sku = marriedItem.pop()
                const uom = marriedItem.pop()
                manAddedUPCsSheet.getRange(manAddedUPCsSheet.getLastRow() + 1, 1, 1, 4).setNumberFormat('@').setValues([[sku, upc, uom, item[0]]]);
                upcDatabaseSheet.getRange(upcDatabaseSheet.getLastRow() + 1, 1, 1, 4).setNumberFormat('@').setValues([[upc, uom, item[0], item[4]]]); 
                barcodeInputRange.setValue('UPC Code has been added to the database temporarily.')
                sheet.getRange(2, 1).activate();
              }
              else
                barcodeInputRange.setValue('Please enter a valid UPC Code to marry.')
            }
            else if (isNumber(quantity_String_Split[0]) && isNotBlank(quantity_String_Split[1]) && quantity_String_Split[1] != null)
            {
              if (item.length !== 1) // The cell to the left contains valid item information
              {
                quantity_String_Split[1] = quantity_String_Split[1].toUpperCase()

                if (item[1].split(' ')[0] === 'was') // The item was already on the manual counts page
                {
                  const range = manualCountsPage.getRange(item[2], 3, 1, 5);
                  const itemValues = range.getValues()
                  const updatedCount = Number(itemValues[0][0]) + Number(quantity_String_Split[0]);
                  const countedSince = getCountedSinceString(itemValues[0][2])
                  const runningSum = (isNotBlank(itemValues[0][1])) ? ((Math.sign(quantity_String_Split[0]) === 1 || Math.sign(quantity_String_Split[0]) === 0)  ? 
                                                                        String(itemValues[0][1]) + ' \+ ' + String(   quantity_String_Split[0])  : 
                                                                        String(itemValues[0][1]) + ' \- ' + String(-1*quantity_String_Split[0])) :
                                                                          ((isNotBlank(itemValues[0][0])) ? 
                                                                            String(itemValues[0][0]) + ' \+ ' + String(quantity_String_Split[0]) : 
                                                                            String(quantity_String_Split[0]));

                  if (isNotBlank(itemValues[0][3]) && isNotBlank(itemValues[0][4]))
                    range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                      itemValues[0][3] + '\n' + quantity_String_Split[1], itemValues[0][4] + '\n' + quantity_String_Split[0].toString()]]);
                  else if (isNotBlank(itemValues[0][3]))
                    range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                      itemValues[0][3] + '\n' + quantity_String_Split[1], quantity_String_Split[0].toString()]]);
                  else if (isNotBlank(itemValues[0][4]))
                    range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                      quantity_String_Split[1], itemValues[0][4] + '\n' + quantity_String_Split[0].toString()]]);
                  else
                    range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                      quantity_String_Split[1], quantity_String_Split[0].toString()]]);

                  sheet.getRange(2, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Manual Counts page at line :\n' + item[2] 
                                                                  + '\nCurrent Stock :\n' + item[4] 
                                                                  + '\nCurrent Manual Count :\n' + updatedCount 
                                                                  + '\nCurrent Running Sum :\n' + runningSum
                                                                  + '\nLast Counted :\n' + countedSince,
                                                                  '']]);
                }
                else
                {
                  const lastRow = manualCountsPage.getLastRow();
                  const row = lastRow + 1;
                  const range = manualCountsPage.getRange(row, 1, 1, 7)
                  const itemValues = range.getValues()

                  if (isNotBlank(itemValues[0][5]) && isNotBlank(itemValues[0][6]))
                    range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], item[4], quantity_String_Split[0], '\'' + String(quantity_String_Split[0]),
                      new Date().getTime(), itemValues[0][5] + '\n' + quantity_String_Split[1], itemValues[0][6] + '\n' + quantity_String_Split[0].toString()]]);
                  else if (isNotBlank(itemValues[0][5]))
                    range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], item[4], quantity_String_Split[0], '\'' + String(quantity_String_Split[0]),
                      new Date().getTime(), itemValues[0][5] + '\n' + quantity_String_Split[1], quantity_String_Split[0].toString()]]);
                  else if (isNotBlank(itemValues[0][6]))
                    range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], item[4], quantity_String_Split[0], '\'' + String(quantity_String_Split[0]),
                      new Date().getTime(), quantity_String_Split[1], itemValues[0][6] + '\n' + quantity_String_Split[0].toString()]]);
                  else
                    range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], item[4], quantity_String_Split[0], '\'' + String(quantity_String_Split[0]),
                      new Date().getTime(), quantity_String_Split[1], quantity_String_Split[0].toString()]]);

                  applyFullRowFormatting(manualCountsPage, row, 1, 7)
                  sheet.getRange(2, 1, 1, 2).setValues([[item[0]  + '\nwas added to the Manual Counts page at line :\n' + item[2] 
                                                                  + '\nCurrent Stock :\n' + item[4] 
                                                                  + '\nCurrent Manual Count :\n' + quantity_String_Split[0],
                                                                  '']]);
                }
              }
              else // The cell to the left does not contain the necessary item information to be able to move it to the manual counts page
                barcodeInputRange.setValue('Please scan your barcode in the left cell again.')

              sheet.getRange(2, 1).activate();
            }
            else if (isNumber(quantity_String_Split[1]))
            {
              if (item.length !== 1) // The cell to the left contains valid item information
              {
                quantity_String_Split[0] = quantity_String_Split[0].toUpperCase()

                if (item[1].split(' ')[0] === 'was') // The item was already on the manual counts page
                {
                  const range = manualCountsPage.getRange(item[2], 3, 1, 5);
                  const itemValues = range.getValues()
                  const updatedCount = Number(itemValues[0][0]) + Number(quantity_String_Split[1]);
                  const countedSince = getCountedSinceString(itemValues[0][2])
                  const runningSum = (isNotBlank(itemValues[0][1])) ? ((Math.sign(quantity_String_Split[1]) === 1 || Math.sign(quantity_String_Split[1]) === 0)  ? 
                                                                        String(itemValues[0][1]) + ' \+ ' + String(   quantity_String_Split[1])  : 
                                                                        String(itemValues[0][1]) + ' \- ' + String(-1*quantity_String_Split[1])) :
                                                                          ((isNotBlank(itemValues[0][0])) ? 
                                                                            String(itemValues[0][0]) + ' \+ ' + String(quantity_String_Split[1]) : 
                                                                            String(quantity_String_Split[1]));

                  if (isNotBlank(itemValues[0][3]) && isNotBlank(itemValues[0][4]))
                    range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                      itemValues[0][3] + '\n' + quantity_String_Split[0], itemValues[0][4] + '\n' + quantity_String_Split[1].toString()]]);
                  else if (isNotBlank(itemValues[0][3]))
                    range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                      itemValues[0][3] + '\n' + quantity_String_Split[0], quantity_String_Split[1].toString()]]);
                  else if (isNotBlank(itemValues[0][4]))
                    range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                      quantity_String_Split[0], itemValues[0][4] + '\n' + quantity_String_Split[1].toString()]]);
                  else
                    range.setNumberFormats([['#.#', '@', '#', '@', '@']]).setValues([[updatedCount, runningSum, new Date().getTime(), 
                      quantity_String_Split[0], quantity_String_Split[1].toString()]]);

                  sheet.getRange(2, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Manual Counts page at line :\n' + item[2] 
                                                                  + '\nCurrent Stock :\n' + item[4] 
                                                                  + '\nCurrent Manual Count :\n' + updatedCount 
                                                                  + '\nCurrent Running Sum :\n' + runningSum
                                                                  + '\nLast Counted :\n' + countedSince,
                                                                  '']]);
                }
                else
                {
                  const lastRow = manualCountsPage.getLastRow();
                  const row = lastRow + 1;
                  const range = manualCountsPage.getRange(row, 1, 1, 7)
                  const itemValues = range.getValues()

                  if (isNotBlank(itemValues[0][5]) && isNotBlank(itemValues[0][6]))
                    range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], item[4], quantity_String_Split[1], '\'' + String(quantity_String_Split[1]),
                      new Date().getTime(), itemValues[0][5] + '\n' + quantity_String_Split[0], itemValues[0][6] + '\n' + quantity_String_Split[1].toString()]]);
                  else if (isNotBlank(itemValues[0][5]))
                    range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], item[4], quantity_String_Split[1], '\'' + String(quantity_String_Split[1]),
                      new Date().getTime(), itemValues[0][5] + '\n' + quantity_String_Split[0], quantity_String_Split[1].toString()]]);
                  else if (isNotBlank(itemValues[0][6]))
                    range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], item[4], quantity_String_Split[1], '\'' + String(quantity_String_Split[1]),
                      new Date().getTime(), quantity_String_Split[0], itemValues[0][6] + '\n' + quantity_String_Split[1].toString()]]);
                  else
                    range.setNumberFormats([['@', '@', '#.#', '@', '#', '@', '@']]).setValues([[item[0], item[4], quantity_String_Split[1], '\'' + String(quantity_String_Split[1]),
                      new Date().getTime(), quantity_String_Split[0], quantity_String_Split[1].toString()]]);

                  applyFullRowFormatting(manualCountsPage, row, 1, 7)
                  sheet.getRange(2, 1, 1, 2).setValues([[item[0]  + '\nwas added to the Manual Counts page at line :\n' + item[2] 
                                                                  + '\nCurrent Stock :\n' + item[4] 
                                                                  + '\nCurrent Manual Count :\n' + quantity_String_Split[1],
                                                                  '']]);
                }
              }
              else // The cell to the left does not contain the necessary item information to be able to move it to the manual counts page
                barcodeInputRange.setValue('Please scan your barcode in the left cell again.')

              sheet.getRange(2, 1).activate();
            }
            else if (quantity <= 100000) // If false, Someone probably scanned a barcode in the quantity cell (not likely to have counted an inventory amount of 100 000)
            {
              if (item.length !== 1) // The cell to the left contains valid item information
              {
                if (item[1].split(' ')[0] === 'was') // The item was already on the manual counts page
                {
                  const range = manualCountsPage.getRange(item[2], 3, 1, 3);
                  const itemValues = range.getValues()
                  const updatedCount = Number(itemValues[0][0]) + quantity;
                  const countedSince = (isNotBlank(itemValues[0][2])) ? getCountedSinceString(itemValues[0][2]) : '1 second ago'
                  const runningSum = (isNotBlank(itemValues[0][1])) ? ((Math.sign(quantity) === 1 || Math.sign(quantity) === 0)  ? 
                                                                        String(itemValues[0][1]) + ' \+ ' + String(   quantity)  : 
                                                                        String(itemValues[0][1]) + ' \- ' + String(-1*quantity)) :
                                                                          ((isNotBlank(itemValues[0][0])) ? 
                                                                            String(itemValues[0][0]) + ' \+ ' + String(quantity) : 
                                                                            String(quantity));
                  range.setNumberFormats([['#.#', '@', '#']]).setValues([[updatedCount, runningSum, new Date().getTime()]])
                  sheet.getRange(2, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Manual Counts page at line :\n' + item[2] 
                                                                  + '\nCurrent Stock :\n' + item[4] 
                                                                  + '\nCurrent Manual Count :\n' + updatedCount 
                                                                  + '\nCurrent Running Sum :\n' + runningSum
                                                                  + '\nLast Counted :\n' + countedSince,
                                                                  '']]);
                }
                else
                {
                  const lastRow = manualCountsPage.getLastRow();
                  const row = lastRow + 1;
                  manualCountsPage.getRange(row, 1, 1, 5).setNumberFormats([['@', '@', '#.#', '@', '#']]).setValues([[item[0], item[4], quantity, '\'' + String(quantity), new Date().getTime()]])
                  applyFullRowFormatting(manualCountsPage, row, 1, 7)
                  sheet.getRange(2, 1, 1, 2).setValues([[item[0]  + '\nwas added to the Manual Counts page at line :\n' + item[2] 
                                                                  + '\nCurrent Stock :\n' + item[4] 
                                                                  + '\nCurrent Manual Count :\n' + quantity,
                                                                  '']]);
                }
              }
              else // The cell to the left does not contain the necessary item information to be able to move it to the manual counts page
                barcodeInputRange.setValue('Please scan your barcode in the left cell again.')

              sheet.getRange(2, 1).activate();
            }
            else 
              barcodeInputRange.setValue('Please enter a valid quantity.')
          }
        }
      }
    }
    else if (col !== 2) // An item is either being search for or an item from the search list is being selected
    {
      const value = barcodeInputRange.getValue()

      if (isNotBlank(value)) // The user has either selected on of the data validation options, or tried to search for a particular item
      {
        if (barcodeInputRange.getDataValidation().getCriteriaValues()[0].includes(value)) // A user has selected one of the items in the data validation
        {
          const item = value.split(' - Stock: ', 2)
          const lastRow = manualCountsPage.getLastRow();

          if (lastRow <= 3) // There are no items on the manual counts page
            barcodeInputRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([]).build())
              .offset( 1, 0, 1, 2).setValues([[item[0] + '\nwill be added to the Manual Counts page at line :\n' + 4 + '\nCurrent Stock :\n' + item[1], '']])
              .offset(-1, 1, 1, 1).setValue('1 result found');
          else // There are existing items on the manual counts page
          {
            const numRows = lastRow - 3;
            const manualCountsValues = manualCountsPage.getSheetValues(4, 1, lastRow - 3, 5);

            for (var j = 0; j < numRows; j++) // Loop through the manual counts page
            {
              if (manualCountsValues[j][0] === item[0]) // The description matches
              {
                const countedSince = getCountedSinceString(manualCountsValues[j][4])
                  
                barcodeInputRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([]).build())
                  .offset(1, 0, 1, 2).setValues([[item[0] + '\nwas found on the Manual Counts page at line :\n' + (j + 4) 
                                                          + '\nCurrent Stock :\n' + item[1] 
                                                          + '\nCurrent Manual Count :\n' + manualCountsValues[j][2] 
                                                          + '\nCurrent Running Sum :\n' + manualCountsValues[j][3]
                                                          + '\nLast Counted :\n' + countedSince, '']])
                  .offset(-1, 1, 1, 1).setValue('1 result found');
                break; // Item was found on the manual counts page, therefore stop searching
              }
            }

            if (j === numRows) // Item was not found on the manual counts page
              barcodeInputRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([]).build())
                .offset( 1, 0, 1, 2).setValues([[item[0] + '\nwill be added to the Manual Counts page at line :\n' + (lastRow + 1) + '\nCurrent Stock :\n' + item[1], '']])
                .offset(-1, 1, 1, 1).setValue('1 result found');
          }

          sheet.getRange(2, 2).activate();
        }
        else // A user has NOT selected one of the items in the data validation, therefore we assume that the user is doing a search
        {
          const inventorySheet = spreadsheet.getSheetByName("INVENTORY");
          const searchWords = barcodeInputRange.clearFormat()                                     // Clear the formatting of the range of the search box
            .setFontFamily("Arial").setFontColor("black").setFontWeight("normal").setFontSize(12) // Set the various font parameters
            .setHorizontalAlignment("center").setVerticalAlignment("middle")                      // Set the alignment
            .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)                                    // Set the wrap strategy
            .trimWhitespace()                                                                     // trim the whitespaces at the end of the string
            .getValue().toString().toLowerCase().split(/\s+/);                                    // Split the search string at whitespacecharacters into an array of search words
          
          if (isNotBlank(searchWords[0]))
          {
            barcodeInputRange.offset(0, 1).setValue('Searching...')
            SpreadsheetApp.flush()

            const numSearchWords = searchWords.length - 1;
            const searchResults = [];

            if (isRichmondSpreadsheet(spreadsheet)) 
            {
              var currentStock = 2; // Changes the index number for selecting the current stock from inventory data
              var numItems = inventorySheet.getLastRow() - 7;
              var transferData = inventorySheet.getSheetValues(8, 2, numItems, currentStock + 1)
            }
            else
            {
              var currentStock = (isParksvilleSpreadsheet(spreadsheet)) ? 2 : 3; // Changes the index number for selecting the current stock from inventory data
              var numItems = inventorySheet.getLastRow() - 9;
              var transferData = inventorySheet.getSheetValues(10, 2, numItems, currentStock + 1)   // Rupert
            }

            search: for (var i = 0; i < numItems; i++)
            {
              for (var j = 0; j <= numSearchWords; j++) // Loop through each word in the User's query
              {
                if (transferData[i][0].toString().toLowerCase().includes(searchWords[j])) // Does the i-th item description contain the j-th search word
                {
                  if (j === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                  {
                    searchResults.push(transferData[i][0] + ' - Stock: ' + transferData[i][currentStock]);

                    if (searchResults.length === 500)
                      break search;
                  }
                }
                else
                  break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
              }
            }

            const numResults = searchResults.length

            if (numResults === 1) // If only 1 result is found, populate the manual scan with that item
            {
              const item = searchResults[0].split(' - Stock: ', 2)
              const lastRow = manualCountsPage.getLastRow();

              if (lastRow <= 3) // There are no items on the manual counts page
                barcodeInputRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([]).build())
                  .offset( 1, 0, 1, 2).setValues([[item[0] + '\nwill be added to the Manual Counts page at line :\n' + 4 + '\nCurrent Stock :\n' + item[1], '']])
                  .offset(-1, 1, 1, 1).setValue('1 result found');
              else // There are existing items on the manual counts page
              {
                const numRows = lastRow - 3;
                const manualCountsValues = manualCountsPage.getSheetValues(4, 1, lastRow - 3, 5);

                for (var j = 0; j < numRows; j++) // Loop through the manual counts page
                {
                  if (manualCountsValues[j][0] === item[0]) // The description matches
                  {
                    const countedSince = getCountedSinceString(manualCountsValues[j][4])
                      
                    barcodeInputRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([]).build())
                      .offset(1, 0, 1, 2).setValues([[item[0] + '\nwas found on the Manual Counts page at line :\n' + (j + 4) 
                                                              + '\nCurrent Stock :\n' + item[1] 
                                                              + '\nCurrent Manual Count :\n' + manualCountsValues[j][2] 
                                                              + '\nCurrent Running Sum :\n' + manualCountsValues[j][3]
                                                              + '\nLast Counted :\n' + countedSince, '']])
                      .offset(-1, 1, 1, 1).setValue('1 result found');
                    break; // Item was found on the manual counts page, therefore stop searching
                  }
                }

                if (j === numRows) // Item was not found on the manual counts page
                  barcodeInputRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([]).build())
                    .offset(1, 0, 1, 2).setValues([[item[0] + '\nwill be added to the Manual Counts page at line :\n' + (lastRow + 1) + '\nCurrent Stock :\n' + item[1], '']])
                    .offset(-1, 1, 1, 1).setValue('1 result found');
              }

              sheet.getRange(2, 2).activate();
            }
            else if (numResults === 500)
              barcodeInputRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(searchResults).build())
                .offset(0, 1).setValue('Too many results, 500 items displayed')
            else if (numResults !== 0) // Items found based on the user's search words
              barcodeInputRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(searchResults).build())
                .offset(0, 1).setValue(numResults + ' results found')
            else if (numSearchWords === 0 && isNumber(Number(searchWords[0])))
            {
                const lastRow = manualCountsPage.getLastRow();
                const upcDatabase = spreadsheet.getSheetByName("UPC Database").getDataRange().getValues();
                const numUpcs = upcDatabase.length - 1;

                if (lastRow <= 3) // There are no items on the manual counts page
                {
                  for (var i = numUpcs; i >= 1; i--) // Loop through the UPC values
                  {
                    if (upcDatabase[i][0] == searchWords[0]) // UPC found
                    {
                      barcodeInputRange.offset(1, 0).setValue(upcDatabase[i][2] + '\nwill be added to the Manual Counts page at line :\n' + 4 + '\nCurrent Stock :\n' + upcDatabase[i][3]);
                      break; // Item was found, therefore stop searching
                    }
                  }
                }
                else // There are existing items on the manual counts page
                {
                  const row = lastRow + 1;
                  const numRows = row - 3
                  const manualCountsValues = manualCountsPage.getSheetValues(4, 1, numRows, 5);

                  for (var i = numUpcs; i >= 1; i--) // Loop through the UPC values
                  {
                    if (upcDatabase[i][0] == searchWords[0])
                    {
                      for (var j = 0; j < numRows; j++) // Loop through the manual counts page
                      {
                        if (manualCountsValues[j][0] === upcDatabase[i][2]) // The description matches
                        {
                          const countedSince = getCountedSinceString(manualCountsValues[j][4])
                            
                          barcodeInputRange.offset(0, 0, 2, 2).setValues([['', ''], [
                            upcDatabase[i][2]  + '\nwas found on the Manual Counts page at line :\n' + (j + 4) 
                                               + '\nCurrent Stock :\n' + upcDatabase[i][3] 
                                               + '\nCurrent Manual Count :\n' + manualCountsValues[j][2] 
                                               + '\nCurrent Running Sum :\n' + manualCountsValues[j][3]
                                               + '\nLast Counted :\n' + countedSince, '']]);
                          break; // Item was found on the manual counts page, therefore stop searching
                        }
                      }

                      if (j === numRows) // Item was not found on the manual counts page
                        barcodeInputRange.offset(0, 0, 2, 2).setValues([['', ''], [upcDatabase[i][2] + '\nwill be added to the Manual Counts page at line :\n' + row + '\nCurrent Stock :\n' + upcDatabase[i][3], '']]);

                      break;
                    }
                  }
                }

                if (i === 0)
                  barcodeInputRange.setDataValidation(SpreadsheetApp.newDataValidation()
                      .requireValueInList(['No item descriptions contain all of the following search words: ' + searchWords.join(" ")]).build())
                    .offset(0, 1).setValue('0 results found')
                else
                  barcodeInputRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['UPC Code Found in Database: ' + searchWords[0]]).build())
                    .offset(1, 1).activate()
            }
            else // No items found based on the user's search words 
              barcodeInputRange.setDataValidation(SpreadsheetApp.newDataValidation()
                  .requireValueInList(['No item descriptions contain all of the following search words: ' + searchWords.join(" ")]).build())
                .offset(0, 1).setValue('0 results found')
          }
          else if (isNotBlank(e.oldValue) && userHasPressedDelete(e.value)) // If the user deletes the data in the search box, then the recently created items are displayed
            barcodeInputRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([]).build())
              .offset(0, 1).setValue('Invalid Search');
          else
            barcodeInputRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([]).build())
              .offset(0, 1).setValue('Invalid Search');
        }
      }
      else
        barcodeInputRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([]).build())
          .offset(0, 1).setValue('Invalid Selection or Search')
    }
  }
}

/**
 * This function moves rows from one sheet to another and fixes formatting issues when rows are added by the user directly on the order page.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @param    {String}     sheetName  : The name of the sheet that is being edited
 * @author Jarren Ralf
 */
function moveRow(e, spreadsheet, sheet, sheetName)
{
  const  value = e.value;           // The value of the edited cell
  const  range = e.range;           // The range of the edited cell
  const row    = range.rowStart;    // The first row of the edited range
  const col    = range.columnStart; // The first column of the edited range
  const rowEnd = range.rowEnd;      // The last row in the edited range
  const colEnd = range.columnEnd;   // The last column in the edited range

  if (row == rowEnd && col == colEnd) // Make sure that only one cell is being edited
  {
    if (col == 10 && userHasNotPressedDelete(value)) // Status Column edited, excluding pressing the delete key
    {
      const numCols = 11;                                         // The number of columns in a row
      const rowRange = sheet.getRange(row, 1, 1, numCols);        // The full row range
      const rowValues = rowRange.getValues();                     // The full row values
      const orderedQty = rowValues[0][2];                         // The ordered quantity
      const shippedQty = rowValues[0][8];                         // The shipped quantity  
      const shippedSheet = spreadsheet.getSheetByName("Shipped"); // The shipped sheet

      if (sheetName == "Order") // An edit is occuring on the Order sheet
      {        
        if (value == "Item Not Available") // The cell is set to "Item Not Available"  
        {
          rowValues[0][8] = 'N/A';
          transferRow(sheet, shippedSheet, row, rowValues, numCols, true);
          sendEmailToBranchStore(value, row, rowValues, sheet, spreadsheet)
        }
        else if (value == "Discontinued") // The cell is set to "Discontinued"  
        {
          rowValues[0][8] = 'Discont';
          transferRow(sheet, shippedSheet, row, rowValues, numCols, true);
          sendEmailToBranchStore(value, row, rowValues, sheet, spreadsheet)
        }
        else if (value == "Order From Distributor") // The cell is set to "Order From Distributor"  
        {
          rowValues[0][8] = 'Reorder';
          transferRow(sheet, shippedSheet, row, rowValues, numCols, true);
          sendEmailToBranchStore(value, row, rowValues, sheet, spreadsheet)
        }
        else // This means order and shipped quantities need to be checked
        {
          if (isNumber(shippedQty) && shippedQty > 0) // If the shipped quantity is a positive number 
          {
            if (isNumber(orderedQty) && isNotBlank(orderedQty)) // Check if the order quantity is a valid number
            {
              if (shippedQty >= orderedQty) // This is a complete shipment (No Back Orders)
              {
                if (value == "Carrier Not Assigned")
                  transferRow(sheet, shippedSheet, row, rowValues, numCols, true);
                else
                {
                  const dataValidation = spreadsheet.getSheetByName("Data Validation").getRange('B:C').getValues(); // These are all the data validation choices of carriers, etc.
                  const numDataValidations = dataValidation.length
                  
                  for (var i = 0; i < numDataValidations; i++)
                    if (value == dataValidation[i][0]) // The value selected matches th i-th data validation
                      transferRow(sheet, shippedSheet, row, rowValues, numCols, true, dataValidation[i][1], dataValidation[i][0]);

                }
              }
              else // Partial shipment, there some portion of the item will be on back order
              {
                if (value == "Carrier Not Assigned")
                {
                  const richText = transferRow(sheet, shippedSheet, row, rowValues, numCols, false);
                  updateBO(rowRange,rowValues);
                  sheet.getRange(row, 6).setRichTextValue(richText);
                }
                else
                {
                  const dataValidation = spreadsheet.getSheetByName("Data Validation").getRange('B:C').getValues(); // These are all the data validation choices of carriers, etc.
                  const numDataValidations = dataValidation.length
                  
                  for (var i = 0; i < numDataValidations; i++)
                  {
                    if (value == dataValidation[i][0]) // The value selected matches th i-th data validation
                    {
                      const richText = transferRow(sheet, shippedSheet, row, rowValues, numCols, false, dataValidation[i][1], dataValidation[i][0]);
                      updateBO(rowRange,rowValues);
                      sheet.getRange(row, 6).setRichTextValue(richText);
                    }
                  }
                }
              }
            }
            else // The order quantity is invalid
            {
              Browser.msgBox('The ordered quantity is invalid.');
              rowValues[0][9] = e.oldValue;
              rowRange.setValues(rowValues);
            }
          }
          else // If the shipped quantity is BLANK or not a positive number
          {
            if (isNumber(orderedQty) && isNotBlank(orderedQty)) // Check if the order quantity is a valid number
            {
              const ui = SpreadsheetApp.getUi(); // Get the User Interface object
              var response = ui.prompt('Invalid Input in the Shipped Column!', 
                                      'The ordered quantity is ' + orderedQty + '. \nHow many are you shipping?.',
                                      ui.ButtonSet.OK_CANCEL);
          
              if (response.getSelectedButton() == ui.Button.OK ) // If the user clicks 'OK'
              {
                var userTypedResponse = response.getResponseText();
                
                if (isNumber(userTypedResponse)) // Check if the input is a number
                {
                  if (userTypedResponse >= orderedQty) // Complete shipment
                  {
                    if (value == "Carrier Not Assigned")
                    {
                      rowValues[0][8] = userTypedResponse;
                      transferRow(sheet, shippedSheet, row, rowValues, numCols, true);
                    }
                    else
                    {
                      const dataValidation = spreadsheet.getSheetByName("Data Validation").getRange('B:C').getValues(); // These are all the data validation choices of carriers, etc.
                      const numDataValidations = dataValidation.length
                      
                      for (var i = 0; i < numDataValidations; i++)
                      {
                        if (value == dataValidation[i][0]) // The value selected matches th i-th data validation
                        {
                          rowValues[0][8] = userTypedResponse;
                          transferRow(sheet, shippedSheet, row, rowValues, numCols, true, dataValidation[i][1], dataValidation[i][0]);
                        }
                      }
                    }
                  }
                  else if (userTypedResponse > 0) // Partial shipment
                  {
                    if (value == "Carrier Not Assigned")
                    {
                      rowValues[0][8] = userTypedResponse;
                      const richText = transferRow(sheet, shippedSheet, row, rowValues, numCols, false);
                      updateBO(rowRange,rowValues);
                      sheet.getRange(row, 6).setRichTextValue(richText);
                    }
                    else
                    {
                      const dataValidation = spreadsheet.getSheetByName("Data Validation").getRange('B:C').getValues(); // These are all the data validation choices of carriers, etc.
                      const numDataValidations = dataValidation.length
                      
                      for (var i = 0; i < numDataValidations; i++)
                      {
                        if (value == dataValidation[i][0]) // The value selected matches th i-th data validation
                        {
                          rowValues[0][8] = userTypedResponse;
                          const richText = transferRow(sheet, shippedSheet, row, rowValues, numCols, false, dataValidation[i][1], dataValidation[i][0]);
                          updateBO(rowRange,rowValues);
                          sheet.getRange(row, 6).setRichTextValue(richText);
                        }
                      }
                    }
                  }
                  else // The user has entered 0, or a negative number as the quantity
                  {
                    ui.alert('Invalid Response. Number must be positive.');
                    rowValues[0][9] = e.oldValue;
                    rowRange.setValues(rowValues);
                  }
                }
                else // The user's typed response was not a number, and hence invalid
                {
                  ui.alert('Invalid Response. User must enter a positive number.');
                  rowValues[0][9] = e.oldValue;
                  rowRange.setValues(rowValues);
                }
              }
              else // The user selected Cancel
              {
                rowValues[0][9] = e.oldValue;
                rowRange.setValues(rowValues);
              }
            }
            else // The shipped quantity and the order quantity are both invalid
            {
              Browser.msgBox('The ordered quantity is invalid.');
              rowValues[0][9] = e.oldValue;
              rowRange.setValues(rowValues);
            }
          }
        }
      }              
      else if (sheetName == "Shipped") // An edit is occuring on the Shipped sheet
      {
        const dataValidationSheet = spreadsheet.getSheetByName("Data Validation");
        const lastRow = getLastRowSpecial(dataValidationSheet.getRange('C:C').getValues())
        const isReceived = value.split(' ', 1)[0]

        if (isReceived == "Rec'd") // The cell is set to "Received"  
        {
          if (shippedQty != 0)
          {
            const dataValidation = dataValidationSheet.getSheetValues(1, 3, lastRow, 1); // These are all the data validation choices of carriers, etc.
            transferRow(sheet, spreadsheet.getSheetByName("Received"), row, rowValues, numCols, true, undefined, undefined, dataValidation, e);
          }
          else
          {
            rowValues[0][9] = e.oldValue;
            rowRange.setValues(rowValues)
            Browser.msgBox('You are unable to receive an item that has a shipped quantity of zero.')
          }
        }
        else if (value == "Back to Order")
        {
          const dataValidation = dataValidationSheet.getSheetValues(1, 3, lastRow, 1) // These are all the data validation choices of carriers, etc.
          rowValues[0][8] = '';
          transferRow(sheet, spreadsheet.getSheetByName("Order"), row, rowValues, numCols, true, undefined, undefined, dataValidation, e);
        }
        else if (value == "Carrier Not Assigned")
        {
          const dataValidation = dataValidationSheet.getSheetValues(1, 3, lastRow, 1); // These are all the data validation choices of carriers, etc.
          transferRow(sheet, sheet, row, rowValues, numCols, true, undefined, undefined, dataValidation, e);
        }
        else // A specific Carrier choice
        {
          const dataValidation = dataValidationSheet.getSheetValues(1, 3, lastRow, 2);
          const numDataValidations = dataValidation.length

          for (var i = 0; i < numDataValidations; i++)
            if (value == dataValidation[i][1]) // Find the carrier and place the the line at the correct row number
              transferRow(sheet, sheet, row, rowValues, numCols, true, dataValidation[i][0], dataValidation[i][1], dataValidation, e);
        }
      }
      else if (sheetName == "Received" && value == "Back to Shipped") // An edit is occuring on the Received sheet 
        transferRow(sheet, shippedSheet, row, rowValues, numCols, true);
    }
    else if (sheetName == "Shipped" && col == 11 && row > 3) // Shipped Date Column
    {
      var oldValue = e.oldValue;

      if (oldValue === 'via')
      {
        range.setValue(oldValue);
        Browser.msgBox('Please don\'t edit this cell.');
      }
    }
    else if (sheetName == "Order" && col == 9)
    {
      var qty = range.getValue().toString().toLowerCase().split(' ')

      if (qty.length === 2 && isNotBlank(qty[1]) && isNumber(qty[1]) && (qty[0] == 'tt' || qty[0] == 'ttt' || qty[0] == 'tit' || qty[0] == 'tits' || qty[0] == 'boob' || qty[0] == 'boobs'))
      {
        range.setValue(qty[1])
        addToInflowPickList(qty[1])
      }
    }
  }
  else if (value === undefined && col === 1) // A row might have been added by the user by clicking on the "Add" button on the botton of the sheet
  {
    const numRows = rowEnd - row + 1;
    const numCols = sheet.getLastColumn();

    if (colEnd === numCols) // The user added rows to the Order sheet
    {
      if (sheetName == "Order" || sheetName == "ItemsToRichmond")
      {
        applyFullRowFormatting(sheet, row, numRows, numCols);
        dateStamp(row, 1, numRows);
      }
      else
      {
        Browser.msgBox('Please don\'t create new rows on this Sheet.');
        sheet.deleteRows(row, numRows);
      }
    }
    else if (colEnd === 9) // The user added rows to the ItemsToRichmond page
    {
      applyFullRowFormatting(sheet, row, numRows);
      dateStamp(row, 1, numRows);
    }
  }
}

/**
 * This function moves the selected "Carrier Not Assigned" items to the chosen shipment.
 * 
 * @author Jarren Ralf
 */
function moveSelectedItemsFromCarrierNotAssigned()
{
  const ui = SpreadsheetApp.getUi();
  const shippedSheet = SpreadsheetApp.getActiveSheet()

  if (shippedSheet.getSheetName() !== "Shipped")
  {
    ui.alert('You must be on the shipped page to run this function.');
    SpreadsheetApp.getActive().getSheetByName("Shipped").activate();
  }
  else // The user is on the shipped sheet
  {
    const spreadsheet = SpreadsheetApp.getActive()
    const activeRanges = spreadsheet.getActiveRangeList().getRanges();
    const rangesToDelete = spreadsheet.getActiveRangeList().getRanges().reverse(); // Delete them from the bottom up, otherwise row numbers will change
    const numCols = shippedSheet.getLastColumn();
    const dataValidationSheet = spreadsheet.getSheetByName("Data Validation");
    const allCarriers = dataValidationSheet.getSheetValues(1, 1, dataValidationSheet.getLastRow(), 3)
    const currentShipmentList = [...new Set(allCarriers.map(carrier => carrier[0]).flat().filter(carrier => isNotBlank(carrier)))];

    const response = ui.prompt('Choose a shipment:\n\n' + currentShipmentList.map((shipment, i) => i.toString() + ': ' + shipment).join('\n') 
      + '\n\nMissing carrier? Get Adrian to format the shipped page and run this function again.');

    if (response.getSelectedButton() === ui.Button.OK)
    {
      const responseText = response.getResponseText();

      if (isNotBlank(responseText))
      {
        const index = Number(responseText);

        if (index > -1 && index < currentShipmentList.length) // Valid index selection
        {
          const chosenShipment = currentShipmentList[index];
          var shipmentRow = allCarriers.find(carrier => carrier[1] === chosenShipment)[2];
          var range, col, column_1, range_NumRows, items, backgroundColours_Dates, backgroundColours_Notes, richTextValues, notesRange, 
            values = [], dateColours = [], notesColours = [], notesRichText = [], isCarrierNotAssigned = true, isNewShipment = false;

          if (typeof shipmentRow === 'string') // We must create a new carrier line
          {
            isNewShipment = true;
            shipmentRow = Number(shipmentRow.replace(/^\D+/g,'')); // Convert the string to a number
            const cols = numCols - 1;
            shipmentRow++;
            shippedSheet.insertRowAfter(shipmentRow - 1).setRowHeight(shipmentRow, 40)
              .getRange(shipmentRow, 1, 1, cols)
                .setBackgrounds([new Array(cols).fill('#6d9eeb')])
                .setFontColors([[...new Array(cols - 1).fill('white'), '#6d9eeb']])
                .setFontSizes([new Array(cols).fill(14)])
                .setFontLine('none').setFontWeight('bold').setFontStyle('normal').setFontFamily('Arial')
                .setHorizontalAlignments([new Array(cols).fill('left')])
                .setWrapStrategies([new Array(cols).fill(SpreadsheetApp.WrapStrategy.OVERFLOW)])
                .setDataValidations([new Array(cols).fill(null)])
                .setBorder(true, true, true, true, null, null)
                .setValues([[chosenShipment, ...new Array(9).fill(null), 'via']])
              .offset(0,  0, 1, cols - 1).merge()
              .offset(0, 12, 1, 1).setDataValidation(shippedSheet.getRange(3, 13).getDataValidation())
          }

          shipmentRow++;
          
          while (activeRanges.length > 0) // Loop through the active ranges, leave as "activeRanges.length" because the length of the items array changes with is critical for exiting the while loop
          {
            range = (isNewShipment) ? activeRanges.pop().offset(1, 0) : activeRanges.pop();
            col = range.getColumn();
            column_1 = 1 - col;
            range_NumRows = range.getNumRows();
            items = range.offset(0, column_1, range_NumRows, numCols).getValues().map(carrier => {if (carrier[9] !== 'Carrier Not Assigned') isCarrierNotAssigned = false; carrier[9] = chosenShipment; return carrier});
            backgroundColours_Dates = range.offset(0, column_1, range_NumRows, 1).getBackgrounds();
            notesRange = range.offset(0, 6 - col, range_NumRows, 1);
            backgroundColours_Notes = notesRange.getBackgrounds();
            richTextValues = notesRange.getRichTextValues();
            
            while (items.length > 0) // Leave as "items.length" because the length of the items array changes with is critical for exiting the while loop
            {
              values.push(items.pop());
              dateColours.push(backgroundColours_Dates.pop());
              notesColours.push(backgroundColours_Notes.pop());
              notesRichText.push(richTextValues.pop());
            }

            if (shipmentRow <= 3)
              ui.alert('The following Carrier was not found:\n\n' + chosenShipment)
          }

          if (shipmentRow > 3)
          {
            if (isCarrierNotAssigned)
            {
              const numValues = values.length;

              shippedSheet.insertRowsBefore(shipmentRow, numValues).getRange(shipmentRow, 1, numValues, numCols).setValues(values.reverse());
              applyFullRowFormatting(shippedSheet, shipmentRow, numValues, numCols - 1); 
              shippedSheet.autoResizeRows(shipmentRow, numValues).getRange(shipmentRow, 1, numValues).setBackgrounds(dateColours.reverse())
                .offset(0,   5, numValues).setBackgrounds(notesColours.reverse()).setRichTextValues(notesRichText.reverse())
                .offset(0,   4, numValues).setDataValidation((isNewShipment) ? 
                  shippedSheet.getRange(3, 10).getDataValidation().copy().requireValueInRange(dataValidationSheet.getRange("$D$1:$D")).build() : 
                  shippedSheet.getRange(3, 10).getDataValidation())
                .offset(0,   3, numValues).setDataValidation(null)
                .offset(0, -12, numValues, numCols - 1).activate();

              rangesToDelete.map(range => {
                if (isNewShipment)
                {
                  range = range.offset(1, 0);
                  shippedSheet.deleteRows(numValues + range.getRow(), range.getNumRows()) 
                }
                else
                  shippedSheet.deleteRows(numValues + range.getRow(), range.getNumRows())
              });
            }
            else
              ui.alert('You may only select items with Shipment Status: Carrier Not Assigned.');
          }
        }
        else
          ui.alert('Your input: "' + index.toString() + '" was invalid.');
      }
      else
        ui.alert('Your input was invalid.');
    }
  }
}

/**
 * This function moves the user to the search box on the Item Search page
 * 
 * @author Jarren Ralf
 */
function moveToItemSearch()
{
  ss.getSheetByName("Item Search").getRange(1, 2).activate();
}

/**
 * This function moves the user to the Manual Counts page.
 * 
 * @author Jarren Ralf
 */
function moveToManualCounts()
{
  ss.getSheetByName("Manual Counts").activate();
}

/**
 * This function moves the user to the barcode input cell (left) on the Manual Scan page
 * 
 * @author Jarren Ralf
 */
function moveToManualScan()
{
  ss.getSheetByName("Manual Scan").getRange(1, 1).activate();
}

/**
 * This function moves the user to the UPC Database page.
 * 
 * @author Jarren Ralf
 */
function moveToUpcDatabse()
{
  ss.getSheetByName("UPC Database").activate();
}

/**
 * This function takes the information from the Item Search or Manual Counts page and the user's recently scanned barcode in the created date column and it 
 * populates the Manual Scan page with the relevant data need to update the count of the particular item.
 * 
 * @param {Spreadsheet}   ss          : The active spreadsheet.
 * @param {Sheet}        sheet        : The active sheet.
 * @param {Number}       rowNum       : The row number of the current item.
 * @param {String} newItemDescription : The new description that has been added
 * @author Jarren Ralf
 */
function populateManualScan(ss, sheet, rowNum, newItemDescription)
{
  const barcodeInputRange = ss.getSheetByName("Manual Scan").getRange(1, 1);
  const manualCountsPage = ss.getSheetByName("Manual Counts");
  const currentStock = (sheet.getSheetName() === "Item Search") ? 2 : 1;
  const lastRow = manualCountsPage.getLastRow();
  var itemValues = (sheet.getSheetName() === "Item Search") ? sheet.getSheetValues(rowNum, 2, 1, 3)[0] : sheet.getSheetValues(rowNum, 1, 1, 2)[0];

  if (newItemDescription != null)
  {
    itemValues[0] = newItemDescription;
    itemValues[currentStock] = '';
  }

  if (lastRow <= 3) // There are no items on the manual counts page
    barcodeInputRange.setValue(itemValues[0] + '\nwill be added to the Manual Counts page at line :\n' + 4 + '\nCurrent Stock :\n' + itemValues[currentStock]);
  else // There are existing items on the manual counts page
  {
    const row = lastRow + 1;
    const numRows = row - 4;
    const manualCountsValues = manualCountsPage.getSheetValues(4, 1, numRows, 4);

    for (var j = 0; j < numRows; j++) // Loop through the manual counts page
    {
      if (manualCountsValues[j][0] === itemValues[0]) // The description matches
      {
        barcodeInputRange.setValue(itemValues[0]  + '\nwas found on the Manual Counts page at line :\n' + (j + 4) 
                                                      + '\nCurrent Stock :\n' + itemValues[currentStock] 
                                                      + '\nCurrent Manual Count :\n' + manualCountsValues[j][2] 
                                                      + '\nCurrent Running Sum :\n' + manualCountsValues[j][3]);
        break; // Item was found on the manual counts page, therefore stop searching
      }
    }

    if (j === numRows) // Item was not found on the manual counts page
      barcodeInputRange.setValue(itemValues[0] + '\nwill be added to the Manual Counts page at line :\n' + row + '\nCurrent Stock :\n' + itemValues[currentStock]);
  }

  barcodeInputRange.offset(0, 1).activate();
}

/**
* This function replaces all instances of a number in the given range with an 'x'
*
* @param    {String}     sheetName  : The string that represents the sheet name
* @return {Spreadsheet} spreadsheet : The active spreadsheet
* @author Jarren Ralf
*/
function print_X(sheetName)
{
  const START_ROW = 4;
  const ACTUAL_COUNT_COL = 8;
  
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName(sheetName);
  const numRows = sheet.getLastRow() - START_ROW + 1;
  const actualCountsRange = sheet.getRange(START_ROW, ACTUAL_COUNT_COL, numRows);
  var actualCounts = actualCountsRange.getValues();
  
  for (var i = 0; i < numRows; i++)
    if (!(isNaN(parseInt(actualCounts[i])))) // If the entry is a number then replace enty with an 'x'
      actualCounts[i][0] = 'x';
  
  actualCountsRange.setValues(actualCounts); // Replace the values with the new array that contains the x's

  return spreadsheet;
}

/**
* This function run the print_X function for the Order sheet.
*
* @author Jarren Ralf
*/
function print_X_Order()
{ 
  const startTime = new Date().getTime()
  const spreadsheet = print_X("Order"); 
  spreadsheet.getSheetByName("INVENTORY").getRange(5, 7, 1, 3).setValues([['=COUNTIF(Order_ActualCounts,">=0")', dateStamp(undefined, null, null, null, 'dd MMM HH:mm'), getRunTime(startTime)]]);
}

/**
* This function run the print_X function for the Shipped sheet.
*
* @author Jarren Ralf
*/
function print_X_Shipped()
{
  const startTime = new Date().getTime()
  const spreadsheet = print_X("Shipped");
  spreadsheet.getSheetByName("INVENTORY").getRange(6, 7, 1, 3).setValues([['=COUNTIF(Shipped_ActualCounts,">=0")', dateStamp(undefined, null, null, null, 'dd MMM HH:mm'), getRunTime(startTime)]]);
}

/**
* This function moves all of the selected values on the item search page to the Order page.
*
* @author Jarren Ralf
*/
function richmondToStoreTransfers()
{
  const QTY_COL  = 9;
  const NUM_COLS = 7;
  
  var orderSheet = SpreadsheetApp.getActive().getSheetByName("Order");
  var lastRow = orderSheet.getLastRow();
  
  copySelectedValues(orderSheet, lastRow + 1, NUM_COLS, QTY_COL);
}

/**
 * This function moves all of the items under one carrier on the Shipped page to the Received page.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @param    {String}     sheetName  : The name of the sheet that is being edited
 * @author Jarren Ralf
 */
function receiveAll(e, spreadsheet, sheet)
{
  const  value = e.value;           // The value of the edited cell
  const  range = e.range;           // The range of the edited cell
  const row    = range.rowStart;    // The first row of the edited range
  const col    = range.columnStart; // The first column of the edited range
  const rowEnd = range.rowEnd;      // The last row in the edited range
  const colEnd = range.columnEnd;   // The last column in the edited range
  const numCols = 11; 
  const today = new Date()

  if (row == rowEnd && col == colEnd) // Make sure that only one cell is being edited
  {
    if (col == 13 && userHasNotPressedDelete(value) && value === 'Receive ALL') // Status Column edited, excluding pressing the delete key
    {
      var numRows = 0;
      const numPossibleShipments = sheet.getLastRow() - row + 1;
      const shipments = sheet.getSheetValues(row, col - 2, numPossibleShipments, 2);
      sheet.hideColumns(12, 2);

      if (shipments[0][0] !== 'via')
        Browser.msgBox('The word \'via\' is missing from the K column in row ' + row + '.\n\nCheck the other banners as well.')
      else if (shipments[0][1] === '')
      {
        sheet.getRange(3, col - 1).setFormula('=ArrayFormula(if(K3:K=\"via\",A3:A,\"\"))')
        Browser.msgBox('An important formula may have been missing on this sheet, which is now restored.\n\nPlease try receiving the items again.')
      }
      else if (shipments[0][1] === 'Carrier Not Assigned') // Don't delete the carrier not assigned banner
      {
        range.setValue('');
        numRows = shipments.length - 1;
        const receivedSheet = spreadsheet.getSheetByName("Received");
        const shippedItemsRange = sheet.getRange(row + 1, 1, numRows, numCols);
        const backgroundColours = shippedItemsRange.getBackgrounds();
        const richTextValues = sheet.getRange(row + 1, 6, numRows).getRichTextValues();
        const items = shippedItemsRange.getValues().map(r => {r[9] = 'Received'; r[10] = today; return r});
        receivedSheet.insertRowsAfter(3, numRows).setRowHeights(4, numRows, 21)
        applyFullRowFormatting(receivedSheet, 4, numRows, numCols)
        receivedSheet.getRange(4, 1, numRows, numCols).setBackgrounds(backgroundColours).setValues(items);
        receivedSheet.getRange(4, 6, numRows).setRichTextValues(richTextValues);
        sheet.deleteRows(row + 1, numRows);
      }
      else if (shipments[0][1].split(' - ', 1)[0] === 'Direct') // Check the checkboxes on the received page so that the inventory does not come from 100 stock
      {
        range.setValue('');

        for (var i = 1; i < numPossibleShipments; i++)
        {
          if (shipments[i][0] !== 'via' && shipments[i][1] === '')
            numRows++;
          else if (shipments[i][0] === 'via' && isNotBlank(shipments[i][1]))
            break;
        }

        if (numRows > 0)
        {
          const receivedSheet = spreadsheet.getSheetByName("Received");
          const shippedItemsRange = sheet.getRange(row + 1, 1, numRows, numCols);
          const backgroundColours = shippedItemsRange.getBackgrounds();
          const richTextValues = sheet.getRange(row + 1, 6, numRows).getRichTextValues();
          const items = shippedItemsRange.getValues().map(r => {r[9] = 'Received Direct'; r[10] = today; return r});
          receivedSheet.insertRowsAfter(3, numRows).setRowHeights(4, numRows, 21)
          applyFullRowFormatting(receivedSheet, 4, numRows, numCols)
          receivedSheet.getRange(4, 1, numRows, numCols).setBackgrounds(backgroundColours).setValues(items);
          receivedSheet.getRange(4, 6, numRows).setRichTextValues(richTextValues);
          receivedSheet.getRange(4, 12, numRows).insertCheckboxes().check();
          sheet.deleteRows(row, numRows + 1);
        }
      }
      else // Regular shipments
      {
        range.setValue('');

        for (var i = 1; i < numPossibleShipments; i++)
        {
          if (shipments[i][0] !== 'via' && shipments[i][1] === '')
            numRows++;
          else if (shipments[i][0] === 'via' && isNotBlank(shipments[i][1]))
            break;
        }

        if (numRows > 0)
        {
          const receivedSheet = spreadsheet.getSheetByName("Received");
          const shippedItemsRange = sheet.getRange(row + 1, 1, numRows, numCols);
          const backgroundColours = shippedItemsRange.getBackgrounds();
          const richTextValues = sheet.getRange(row + 1, 6, numRows).getRichTextValues();
          const items = shippedItemsRange.getValues().map(r => {r[9] = 'Received'; r[10] = today; return r});
          receivedSheet.insertRowsAfter(3, numRows).setRowHeights(4, numRows, 21)
          applyFullRowFormatting(receivedSheet, 4, numRows, numCols)
          receivedSheet.getRange(4, 1, numRows, numCols).setBackgrounds(backgroundColours).setValues(items);
          receivedSheet.getRange(4, 6, numRows).setRichTextValues(richTextValues);
          sheet.deleteRows(row, numRows + 1);
        }
      }
    }
  }
}

/**
* This function grabs the MAX_NUM_ITEMS most recently created items from the Recent page and displays them on the search page.
*
* @param {Spreadsheet}   spreadsheet   : The active spreadsheet
* @param    {Sheet}    itemSearchSheet : The active sheet
* @author Jarren Ralf
*/
function recentlyCreatedItems(spreadsheet, itemSearchSheet)
{
  const startTime = new Date().getTime();
  const MAX_NUM_ITEMS = 500;

  if (arguments.length !== 2)
  {
    spreadsheet = SpreadsheetApp.getActive();
    itemSearchSheet = spreadsheet.getActiveSheet();
  }

  const recentData = spreadsheet.getSheetByName("Recent").getSheetValues(2, 1, MAX_NUM_ITEMS, 6);

  if (isRichmondSpreadsheet(spreadsheet))
  {
    recentData.unshift( ["The last " + MAX_NUM_ITEMS + " created items are displayed.", null, null, '=Remaining_InfoCounts&\" Items left to count on the InfoCounts page.\"', null, null], 
                        [(new Date().getTime() - startTime)/1000 + ' s', null, 'Counted\nOn', 'Current Stock In Each Location', null, null],
                        [null, null, null, 'Rich', 'Parks', 'Rupert'])
    itemSearchSheet.getRange(1, 1, MAX_NUM_ITEMS + 3, 6).setValues(recentData);
  }
  else
  {
    const orderSheet = spreadsheet.getSheetByName("Order");
    const shippedSheet = spreadsheet.getSheetByName("Shipped")
    const numOrderedItems = orderSheet.getLastRow() - 3;
    const numShippedItems = shippedSheet.getLastRow() - 3;
    const orderedItems =   orderSheet.getSheetValues(4, 5, numOrderedItems, 1); // The items on the order sheet
    const shippedItems = shippedSheet.getSheetValues(4, 5, numShippedItems, 1); // The items on the shipped sheet
    const backgroundColours = [], fontColours = [];
    var isOnOrderPage, isOnShippedPage;

    for (var i = 0; i < MAX_NUM_ITEMS; i++)
    {
      for (var o = 0; o < numOrderedItems; o++) // Check if the item is on the order page
      {
        if (orderedItems[o][0] === recentData[i][1])
        {
          isOnOrderPage = true
          break;
        }
        isOnOrderPage = false;
      }
      for (var s = 0; s < numShippedItems; s++) // Check if the item is on the shipped page
      {
        if (shippedItems[s][0] === recentData[i][1])
        {
          isOnShippedPage = true
          break;
        }
        isOnShippedPage = false;
      }

      if (isOnShippedPage)
      {
        recentData[i].push(null, 'SHIPPED - On it\'s way');
        backgroundColours.push([...new Array(8).fill('#cc0000')]) // Highlight red
        fontColours.push([...new Array(8).fill('yellow')])        // Yellow font
      }
      else if (isOnOrderPage)
      {
        recentData[i].push(null, 'Already on OrderSheet');
        backgroundColours.push([...new Array(8).fill('yellow')]) // Highlight yellow
        fontColours.push([...new Array(8).fill('#cc0000')])      // Red font
      }
      else // The item is neither on the shipped nor the ordered page
      {
        recentData[i].push('', '');
        backgroundColours.push([...new Array(8).fill('white')])
        fontColours.push([...new Array(8).fill('black')])
      }
    }

    itemSearchSheet.getRange(4, 1, MAX_NUM_ITEMS, 8).setBackgrounds(backgroundColours).setFontColors(fontColours).setValues(recentData);
    itemSearchSheet.getRange(1, 1, 3, 3).setValues([["The last " + MAX_NUM_ITEMS + " created items are displayed.", null, null], [null, null, null], [(new Date().getTime() - startTime)/1000 + ' s', null, null]]);
  }
}

/**
 * This function first applies the standard formatting to the search box, then it seaches the SearchData page for the items in question.
 * It also highlights the items that are already on the shipped page and already on the order page.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf
 */
function search(e, spreadsheet, sheet)
{
  const startTime = new Date().getTime();
  const MAX_NUM_ITEMS = 500;
  const range = e.range;
  const row = range.rowStart;
  const col = range.columnStart;
  const colEnd = range.columnEnd;

  if (row == range.rowEnd) // Check and make sure only a single row is being edited
  {
    if (colEnd == null || colEnd == 3 || col == colEnd) // Check and make sure only a single column is being edited
    {
      if (row === 1 && col === 2) // Check if the search box is edited
      {
        spreadsheet.toast('Searching...')
        
        const searchesOrNot = sheet.getRange(1, 2, 1, 2).clearFormat()                                    // Clear the formatting of the range of the search box
          .setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK) // Set the border
          .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14)             // Set the various font parameters
          .setHorizontalAlignment("center").setVerticalAlignment("middle")                                // Set the alignment
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)                                              // Set the wrap strategy
          .merge().trimWhitespace()                                                                       // Merge and trim the whitespaces at the end of the string
          .getValue().toString().toLowerCase().split(' not ')                                             // Split the search string at the word 'not'

        const searches = searchesOrNot[0].split(' or ').map(words => words.split(/\s+/)) // Split the search values up by the word 'or' and split the results of that split by whitespace

        if (isRichmondSpreadsheet(spreadsheet))
        {
          if (isNotBlank(searches[0][0])) // If the value in the search box is NOT blank, then compute the search
          {
            const inventorySheet = spreadsheet.getSheetByName("INVENTORY");
            const numRows = inventorySheet.getLastRow() - 7;
            const data = inventorySheet.getSheetValues(8, 1, numRows, 7);
            const numSearches = searches.length; // The number searches
            const output = [];
            var numSearchWords, numWordsToNotInlude;

            if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
            {
              if (searches[0][0].toLowerCase() === 'trites')
              {
                if (numSearches === 1 && searches[0].length == 1)
                  output.push(...data.filter(item => item[6] > 0))
                else
                {
                  for (var i = 0; i < numRows; i++) // Loop through all of the descriptions from the search data
                  {
                    loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                    {
                      numSearchWords = searches[j].length - 1;

                      for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                      {
                        if (searches[j][k] === 'trites')
                          continue;

                        if (data[i][6] > 0 && data[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                        {
                          if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                          {
                            output.push(data[i]);
                            break loop;
                          }
                        }
                        else
                          break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                      }
                    }
                  }
                }
              }
              else if (searches[0][0].toLowerCase() === 'all' && searches[0][1].toLowerCase() === 'hoochies')
              {
                const hoochiePrefixes = ['16060005', '16010005', '16050005', '16020000', '16020010', '16060065', '16060010', '16070000', '16075300', '16070975',
                                         '16030000', '16060175', '16200030', '16200000', '16200025', '16200065', '16200021', '16200022', '16200061'];
                const numTypesOfHoochies = hoochiePrefixes.length;
                var hoochies = new Array(numTypesOfHoochies).fill('').map(() => []);

                for (var j = 0; j < numTypesOfHoochies; j++) // Loop through the number of searches
                {
                  for (var i = 0; i < numRows; i++) // Loop through all of the descriptions from the search data
                  {
                    if (data[i][7].toString().substring(0, 8) === hoochiePrefixes[j] && !data[i][1].toString().toLowerCase().includes('rig')) // Does the i-th sku contain begin with the j-th hoochie prefix 
                      hoochies[j].push(data[i]); // The description also does not contain the word "rig"
                  }

                  hoochies[j] = sortHoochies(hoochies[j], 1, hoochiePrefixes[j])
                }

                output.push(...[].concat(...hoochies));
              }
              else
              {
                for (var i = 0; i < numRows; i++) // Loop through all of the descriptions from the search data
                {
                  loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                  {
                    numSearchWords = searches[j].length - 1;

                    for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                    {
                      if (data[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          output.push(data[i]);
                          break loop;
                        }
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                    }
                  }
                }
              }
            }
            else // The word 'not' was found in the search string
            {
              var dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);

              if (searches[0][0].toLowerCase() === 'trites')
              {
                if (numSearches === 1 && searches[0].length == 1)
                  output.push(...data.filter(item => item[6] > 0))
                else
                {
                  for (var i = 0; i < numRows; i++) // Loop through all of the descriptions from the search data
                  {
                    loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                    {
                      numSearchWords = searches[j].length - 1;

                      for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                      {
                        if (searches[j][k] === 'trites')
                          continue;

                        if (data[i][6] > 0 && data[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                        {
                          if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                          {
                            numWordsToNotInlude = dontIncludeTheseWords.length;

                            for (var l = 0; l < numWordsToNotInlude; l++)
                            {
                              if (!data[i][1].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                              {
                                if (l === numWordsToNotInlude - 1)
                                {
                                  output.push(data[i]);
                                  break loop;
                                }
                              }
                              else
                                break;
                            }
                          }
                        }
                        else
                          break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                      }
                    }
                  }
                }
              }
              else
              {
                for (var i = 0; i < numRows; i++) // Loop through all of the descriptions from the search data
                {
                  loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                  {
                    numSearchWords = searches[j].length - 1;

                    for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                    {
                      if (data[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          numWordsToNotInlude = dontIncludeTheseWords.length;

                          for (var l = 0; l < numWordsToNotInlude; l++)
                          {
                            if (!data[i][1].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                            {
                              if (l === numWordsToNotInlude - 1)
                              {
                                output.push(data[i]);
                                break loop;
                              }
                            }
                            else
                              break;
                          }
                        }
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                    }
                  }
                }
              }
            }

            const numItems = output.length;

            if (numItems === 0) // No items were found
              sheet.getRange('B1').activate()                   // Take the user to the search box
                .offset(3, -1, MAX_NUM_ITEMS, 7).clearContent() // Clear the entire item display range
                .offset(-3, 0, 1, 1).setRichTextValue(          // Display message stating "No results found"
                  SpreadsheetApp.newRichTextValue().setText("No results found.\n\nPlease try again.").setTextStyle(0, 16, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build()).build())
                .offset(1, 0, 2, 1).setValue((new Date().getTime() - startTime)/1000 + " s"); // Function runtime
            else
            {
              if (numItems > MAX_NUM_ITEMS) // Over MAX_NUM_ITEMS items were found
              {
                output.splice(MAX_NUM_ITEMS); // Slice off all the entries after MAX_NUM_ITEMS

                const text = numItems + "\nresults found,\nonly\n" + MAX_NUM_ITEMS + " displayed.";
                const n = text.length; 

                sheet.getRange('B4').activate()                      // Move the user to the first result that was found in their search
                  .offset(0, -1, MAX_NUM_ITEMS, 7).setValues(output) // Display the max number of items
                .offset(-3, 0, 1, 1).setRichTextValue(               // Display message stating how many results were found and that the MAX number is being displayed
                  SpreadsheetApp.newRichTextValue().setText(text)
                    .setTextStyle(0, n - MAX_NUM_ITEMS.toString().length - 12, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('white').build())
                    .setTextStyle(n - MAX_NUM_ITEMS.toString().length - 11, n, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build()).build())
                .offset(1, 0, 2, 1).setValue((new Date().getTime() - startTime)/1000 + " s"); // Function runtime
              }
              else // Less than MAX_NUM_ITEMS items were found
                sheet.getRange('B4').activate()                   // Move the user to the first result that was found in their search
                  .offset(0, -1, MAX_NUM_ITEMS, 7).clearContent() // Clear the entire item display range
                  .offset(0, 0, numItems, 7).setValues(output)    // Display all of the items found
                  .offset(-3, 0, 1, 1).setValue((numItems !== 1) ? numItems + " results found." : numItems + " result found.") // Display message stating how many results were found
                  .offset(1, 0, 2, 1).setValue((new Date().getTime() - startTime)/1000 + " s"); // Function runtime
            }
          }
          else if (isNotBlank(e.oldValue) && userHasPressedDelete(e.value)) // If the user deletes the data in the search box, then the recently created items are displayed
            sheet.getRange('B4').activate() // Move the user to the first result that was found in their search
              .offset(0, -1, MAX_NUM_ITEMS, 7).setValues(spreadsheet.getSheetByName("Recent").getSheetValues(2, 1, MAX_NUM_ITEMS, 7)) // Display the max number of items
              .offset(-3, 0, 1, 1).setValue("The last " + MAX_NUM_ITEMS + " created items are displayed.") // Display message stating that the max number of items are being displayed
              .offset(1, 0, 2, 1).setValue((new Date().getTime() - startTime)/1000 + " s");                // Function runtime
          else
            sheet.getRange('B1').activate()                   // Take the user to the search box
              .offset(3, -1, MAX_NUM_ITEMS, 7).clearContent() // Clear the entire item display range
              .offset(-3, 0, 1, 1).setRichTextValue(          // Display message that tells the user that this was an invalid search
                SpreadsheetApp.newRichTextValue().setText("Invalid search.\n\nPlease try again.").setTextStyle(0, 14, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build()).build())
              .offset(1, 0, 2, 1).setValue((new Date().getTime() - startTime)/1000 + " s"); // Function runtime
        }
        else // Parksville or Rupert spreadsheet
        {
          if (isNotBlank(searches[0][0])) // If the value in the search box is NOT blank, then compute the search
          {
            const searchDataSheet = spreadsheet.getSheetByName("SearchData");
            const numDescriptions = searchDataSheet.getLastRow() - 1;
            const descriptions = searchDataSheet.getSheetValues(2, 2, numDescriptions, 1); // All the descriptions (ONLY) from the SearchData sheet
            const numSearches = searches.length; // The number searches
            const firstOutput = [], itemIndices = [];
            var numSearchWords, numWordsToNotInlude;

            if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
            {
              if (searches[0][0].toLowerCase() === 'trites')
              {
                if (numSearches === 1 && searches[0].length == 1)
                {
                  var isTrites;
                
                  searchDataSheet.getSheetValues(2, 2, numDescriptions, 6).filter((item, index) => {
                    isTrites = item[5] > 0;

                    if (isTrites)
                    {
                      firstOutput.push([item[0]]);
                      itemIndices.push(index)
                    }
                      
                    return isTrites
                  })
                }
                else // Specific search for inflow items
                {
                  const tritesData = searchDataSheet.getSheetValues(2, 2, numDescriptions, 6);
                  const numTritesData = tritesData.length;

                  for (var i = 0; i < numTritesData; i++) // Loop through all of the descriptions from the search data
                  {
                    loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                    {
                      numSearchWords = searches[j].length - 1;

                      for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                      {
                        if (searches[j][k] === 'trites')
                          continue;

                        if (tritesData[i][5] > 0 && tritesData[i][0].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                        {
                          if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                          {
                            firstOutput.push([tritesData[i][0]]);
                            itemIndices.push(i);
 
                            break loop;
                          }
                        }
                        else
                          break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                      }
                    }
                  }
                }
              }
              else
              {
                for (var i = 0; i < numDescriptions; i++) // Loop through all of the descriptions from the search data
                {
                  loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                  {
                    numSearchWords = searches[j].length - 1;

                    for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                    {
                      if (descriptions[i][0].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          firstOutput.push([descriptions[i][0]]);
                          itemIndices.push(i);

                          break loop;
                        }
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                    }
                  }
                }
              }
            }
            else // The word 'not' was found in the search string
            {
              var dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);

              if (searches[0][0].toLowerCase() === 'trites')
              {
                const tritesData = searchDataSheet.getSheetValues(2, 2, numDescriptions, 6);
                const numTritesData = tritesData.length;

                for (var i = 0; i < numTritesData; i++) // Loop through all of the descriptions from the search data
                {
                  loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                  {
                    numSearchWords = searches[j].length - 1;

                    for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                    {
                      if (searches[j][k] === 'trites')
                        continue;

                      if (tritesData[i][5] > 0 && tritesData[i][0].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          numWordsToNotInlude = dontIncludeTheseWords.length;

                          for (var l = 0; l < numWordsToNotInlude; l++)
                          {
                            if (!tritesData[i][0].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                            {
                              if (l === numWordsToNotInlude - 1)
                              {
                                firstOutput.push([tritesData[i][0]]);
                                itemIndices.push(i);

                                break loop;
                              }
                            }
                            else
                              break;
                          }
                        }
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                    }
                  }
                }
              }
              else
              {
                for (var i = 0; i < numDescriptions; i++) // Loop through all of the descriptions from the search data
                {
                  loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                  {
                    numSearchWords = searches[j].length - 1;

                    for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                    {
                      if (descriptions[i][0].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                      {
                        if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                        {
                          numWordsToNotInlude = dontIncludeTheseWords.length;

                          for (var l = 0; l < numWordsToNotInlude; l++)
                          {
                            if (!descriptions[i][0].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                            {
                              if (l === numWordsToNotInlude - 1)
                              {
                                firstOutput.push([descriptions[i][0]]);
                                itemIndices.push(i);

                                break loop;
                              }
                            }
                            else
                              break;
                          }
                        }
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                    }
                  }
                }
              }
            }

            const numItems = firstOutput.length;

            if (numItems === 0) // No items were found
              sheet.getRange('B1').activate()                                                 // Take the user to the search box
                .offset(3, -1, MAX_NUM_ITEMS, 8).setBackground('white').setFontColor('black') // Clear content and reset the text format
                .offset(-3, 0, 1, 1).setRichTextValue(                                        // Display message that tells the user that this was an invalid search
                  SpreadsheetApp.newRichTextValue().setText("No results found.\n\nPlease try again.").setTextStyle(0, 16, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build()).build()) 
                .offset(1, 0, 2, 1).setValues([[null], [(new Date().getTime() - startTime)/1000 + " s"]]); // Function runtime
            else
            {
              if (numItems > MAX_NUM_ITEMS) // Over MAX_NUM_ITEMS items were found
              {
                const text = numItems + "\nresults found,\nonly\n" + MAX_NUM_ITEMS + " displayed.";
                const n = text.length; 

                firstOutput.splice(MAX_NUM_ITEMS); // Slice off all the entires after MAX_NUM_ITEMS

                sheet.getRange('B4').activate()                                                                // Move the user to the top of the search items
                  .offset(0, -1, MAX_NUM_ITEMS, 8).clearContent().setBackground('white').setFontColor('black') // Clear the range
                  .offset(0, 1, MAX_NUM_ITEMS, 1).setValues(firstOutput)                                       // Display the first output (the descriptions of the items)
                  .offset(-3, -1, 1, 1).setRichTextValue(                                                      // Display message
                    SpreadsheetApp.newRichTextValue().setText(text)
                      .setTextStyle(0, n - MAX_NUM_ITEMS.toString().length - 12, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('white').build())
                      .setTextStyle(n - MAX_NUM_ITEMS.toString().length - 11, n, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build()).build())
                  .offset(1, 0, 2, 1).setValues([[(new Date().getTime() - startTime)/1000 + " s"], [null]]);   // Function runtime
              }
              else // Less than MAX_NUM_ITEMS items were found
                sheet.getRange('B4').activate()                                                                // Move the user to the top of the search items
                  .offset(0, -1, MAX_NUM_ITEMS, 8).clearContent().setBackground('white').setFontColor('black') // Clear the range
                  .offset(0, 1, numItems, 1).setValues(firstOutput)                                            // Display the first output (the descriptions of the items)
                  .offset(-3, -1, 1, 1).setValue((numItems !== 1) ? numItems + " results found." : numItems + " result found.") // Display message that tells the user that this was an invalid search
                  .offset(1, 0, 2, 1).setValues([[(new Date().getTime() - startTime)/1000 + " s"], [null]]);   // Function runtime

              const columnIndex = (isParksvilleSpreadsheet(spreadsheet)) ? [4, 3, 5, 6] : [5, 3, 4, 6]; // This makes sure the current stock reference on the Order sheet is correct
              const orderSheet = spreadsheet.getSheetByName("Order");
              const shippedSheet = spreadsheet.getSheetByName("Shipped");
              const numOrderedItems = orderSheet.getLastRow() - 3;
              const numShippedItems = shippedSheet.getLastRow() - 3;
              const orderedItems =   orderSheet.getSheetValues(4, 5, numOrderedItems, 1).map(u => u[0].split(' - ').pop()); // The items on the order sheet
              const shippedItems = shippedSheet.getSheetValues(4, 5, numShippedItems, 1).map(u => u[0].split(' - ').pop()); // The items on the shipped sheet
              const data = searchDataSheet.getSheetValues(2, 1, searchDataSheet.getLastRow() - 1, 7);
              const numItemIndices = itemIndices.length;
              const backgroundColours = [], fontColours = [], secondOutput = [];
              var isOnOrderPage, isOnShippedPage;

              for (var i = 0; i < numItemIndices; i++) // Loop through the indices of the found items
              {
                for (var o = 0; o < numOrderedItems; o++) // Check if the item is on the order page
                {
                  if (orderedItems[o] === data[itemIndices[i]][1].split(' - ').pop())
                  {
                    isOnOrderPage = true
                    break;
                  }
                  isOnOrderPage = false;
                }
                for (var s = 0; s < numShippedItems; s++) // Check if the item is on the shipped page
                {
                  if (shippedItems[s] === data[itemIndices[i]][1].split(' - ').pop())
                  {
                    isOnShippedPage = true
                    break;
                  }
                  isOnShippedPage = false;
                }

                if (isOnShippedPage)
                {
                  secondOutput.push([data[itemIndices[i]][0], data[itemIndices[i]][1], data[itemIndices[i]][2], ...columnIndex.map(col => data[itemIndices[i]][col]), 'SHIPPED - On it\'s way']);
                  backgroundColours.push([...new Array(8).fill('#cc0000')]) // Highlight red
                  fontColours.push([...new Array(8).fill('yellow')])        // Yellow font
                }
                else if (isOnOrderPage)
                {
                  secondOutput.push([data[itemIndices[i]][0], data[itemIndices[i]][1], data[itemIndices[i]][2], ...columnIndex.map(col => data[itemIndices[i]][col]), 'Already on OrderSheet']);
                  backgroundColours.push([...new Array(8).fill('yellow')]) // Highlight yellow
                  fontColours.push([...new Array(8).fill('#cc0000')])      // Red font
                }
                else // The item is neither on the shipped nor the ordered page
                {
                  secondOutput.push([data[itemIndices[i]][0], data[itemIndices[i]][1], data[itemIndices[i]][2], ...columnIndex.map(col => data[itemIndices[i]][col]), '']);
                  backgroundColours.push([...new Array(8).fill('white')])
                  fontColours.push([...new Array(8).fill('black')])
                }
              }

              if (numItems > MAX_NUM_ITEMS)
              {
                secondOutput.splice(MAX_NUM_ITEMS); // Slice off all the entries after MAX_NUM_ITEMS
                fontColours.splice(MAX_NUM_ITEMS);
                backgroundColours.splice(MAX_NUM_ITEMS);
                sheet.getRange(4, 1, MAX_NUM_ITEMS, 8).setBackgrounds(backgroundColours).setFontColors(fontColours).setValues(secondOutput).offset(-1, 0, 1, 1).setValue((new Date().getTime() - startTime)/1000 + " s")
              }
              else
                sheet.getRange(4, 1, numItems, 8).setBackgrounds(backgroundColours).setFontColors(fontColours).setValues(secondOutput).offset(-1, 0, 1, 1).setValue((new Date().getTime() - startTime)/1000 + " s")
            }
          }
          else if (isNotBlank(e.oldValue) && userHasPressedDelete(e.value)) // If the user deletes the data in the search box, then the recently created items are displayed
          {
            const orderSheet = spreadsheet.getSheetByName("Order");
            const shippedSheet = spreadsheet.getSheetByName("Shipped");
            const numOrderedItems = orderSheet.getLastRow() - 3;
            const numShippedItems = shippedSheet.getLastRow() - 3;
            const orderedItems =   orderSheet.getSheetValues(4, 5, numOrderedItems, 1); // The items on the order sheet
            const shippedItems = shippedSheet.getSheetValues(4, 5, numShippedItems, 1); // The items on the shipped sheet
            const recentData = spreadsheet.getSheetByName("Recent").getSheetValues(2, 1, MAX_NUM_ITEMS, 7); // These are the most recently created items
            const backgroundColours = [], fontColours = [];

            for (var i = 0; i < MAX_NUM_ITEMS; i++)
            {
              for (var o = 0; o < numOrderedItems; o++) // Check if the item is on the order page
              {
                if (orderedItems[o][0] === recentData[i][1])
                {
                  isOnOrderPage = true
                  break;
                }
                isOnOrderPage = false;
              }
              for (var s = 0; s < numShippedItems; s++) // Check if the item is on the shipped page
              {
                if (shippedItems[s][0] === recentData[i][1])
                {
                  isOnShippedPage = true
                  break;
                }
                isOnShippedPage = false;
              }

              if (isOnShippedPage)
              {
                recentData[i].push('SHIPPED - On it\'s way');
                backgroundColours.push([...new Array(8).fill('#cc0000')]) // Highlight red
                fontColours.push([...new Array(8).fill('yellow')])        // Yellow font
              }
              else if (isOnOrderPage)
              {
                recentData[i].push('Already on OrderSheet');
                backgroundColours.push([...new Array(8).fill('yellow')]) // Highlight yellow
                fontColours.push([...new Array(8).fill('#cc0000')])      // Red font
              }
              else // The item is neither on the shipped nor the ordered page
              {
                recentData[i].push('');
                backgroundColours.push([...new Array(8).fill('white')])
                fontColours.push([...new Array(8).fill('black')])
              }
            }

            sheet.getRange('B1').activate()                                                                // Take the user to the search box
              .offset(3, -1, MAX_NUM_ITEMS, 8).clearContent().setBackgrounds(backgroundColours).setFontColors(fontColours).setValues(recentData) // Display the most recently created items
              .offset(-3, 0, 1, 1).setValue("The last " + MAX_NUM_ITEMS + " created items are displayed.") // Display message that tells the user that this was an invalid search
              .offset(1, 0, 2, 1).setValues([[null], [(new Date().getTime() - startTime)/1000 + " s"]]);   // Function runtime
          }
          else
            sheet.getRange('B1').activate()                   // Take the user to the search box
              .offset(3, -1, MAX_NUM_ITEMS, 8).clearContent().setBackground('white').setFontColor('black') // Clear the entire item display range and set the default font colour and background colour
              .offset(-3, 0, 1, 1).setRichTextValue(          // Display message that tells the user that this was an invalid search
                SpreadsheetApp.newRichTextValue().setText("Invalid search.\n\nPlease try again.").setTextStyle(0, 14, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build()).build())
              .offset(1, 0, 2, 1).setValues([[null], [(new Date().getTime() - startTime)/1000 + " s"]]); // Function runtime
        }

        spreadsheet.toast('Searching Complete.')
      }
      else if (col === 3) // Check for the user trying to marry / unmarry upcs or add a new item
      {
        if (userHasNotPressedDelete(e.value))
        {
          const value = e.value.split(' ', 2);
          range.setValue(e.oldValue);
          range.setNumberFormat('dd MMM yyyy')

          if (value[0].toLowerCase() === 'mmm')
          {
            if (value[1] > 100000)
            {
              const item = sheet.getSheetValues(row, 1, 1, 4)[0]
              const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
              const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");
              manAddedUPCsSheet.getRange(manAddedUPCsSheet.getLastRow() + 1, 1, 1, 4).setNumberFormat('@').setValues([[item[1].split(' - ').pop(), value[1], item[0], item[1]]]);
              upcDatabaseSheet.getRange(upcDatabaseSheet.getLastRow() + 1, 1, 1, 4).setNumberFormat('@').setValues([[value[1], item[0], item[1], item[3]]]); 
              populateManualScan(spreadsheet, sheet, row)
            }
            else
              Browser.msgBox('Invalid UPC Code', 'Please type either mmm, uuu, aaa, or sss, followed by SPACE and the UPC Code.', Browser.Buttons.OK)
          }
          else if (value[0].toLowerCase() === 'uuu')
          {
            if (value[1] > 100000)
            {
              const item = sheet.getSheetValues(row, 2, 1, 1)[0][0];
              const unmarryUpcSheet = spreadsheet.getSheetByName("UPCs to Unmarry");
              unmarryUpcSheet.getRange(unmarryUpcSheet.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[value[1], item]]);
              spreadsheet.getSheetByName("Manual Scan").getRange(1, 1).activate()
            }
            else
              Browser.msgBox('Invalid UPC Code', 'Please type either mmm, uuu, aaa, or sss, followed by SPACE and the UPC Code.', Browser.Buttons.OK)
          }
          else if (value[0].toLowerCase() === 'aaa')
          {
            if (value[1] > 100000)
            {
              const item = sheet.getSheetValues(row, 1, 1, 2)[0]
              const newItem = item[1].split(' - ')
              newItem[newItem.length - 1] = 'MAKE_NEW_SKU'
              item[1] = newItem.join(' - ')
              const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
              const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");
              const inventorySheet = (isRichmondSpreadsheet(spreadsheet)) ? spreadsheet.getSheetByName("INVENTORY") : spreadsheet.getSheetByName("SearchData");
              manAddedUPCsSheet.getRange(manAddedUPCsSheet.getLastRow() + 1, 1, 1, 4).setNumberFormat('@').setValues([['MAKE_NEW_SKU', value[1], item[0], item[1]]]);
              upcDatabaseSheet.getRange(upcDatabaseSheet.getLastRow() + 1, 1, 1, 3).setNumberFormat('@').setValues([[value[1], item[0], item[1]]]); 
              inventorySheet.getRange(inventorySheet.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[item[0], item[1]]]); // Add the 'MAKE_NEW_SKU' item to the inventory sheet

              populateManualScan(spreadsheet, sheet, row, item[1])
              sheet.getRange(4, 1, MAX_NUM_ITEMS, 6).setValues(spreadsheet.getSheetByName("Recent").getSheetValues(2, 1, MAX_NUM_ITEMS, 6));
              sheet.getRange(1, 1, 1, 2).setValues([["The last " + MAX_NUM_ITEMS + " created items are displayed.", ""]]);
            }
            else
              Browser.msgBox('Invalid UPC Code', 'Please type either mmm, uuu, aaa, or sss, followed by SPACE and the UPC Code.', Browser.Buttons.OK)
          }
          else if (value[0].toLowerCase() === 'sss')
            populateManualScan(spreadsheet, sheet, row)
          else
            Browser.msgBox('Invalid Entry', 'Please begin the command with either mmm , uuu, or aaa.', Browser.Buttons.OK)
        }
        else // User has hit the delete key on one of the dates in the Counted On column
        {
          range.setValue(e.oldValue);
          range.setNumberFormat('dd MMM yyyy')
        }
      }
      else if (col !== 2) // If the user edits an cell other than the desciption column (they may need to edit the description column if they want to add barcodes etc.)
        range.setValue(e.oldValue); // If the user accidently edits information, put the original data back
    }
  }
  else if (row > 3) // multiple rows are being edited
  {
    spreadsheet.toast('Searching...')                                                                    
    const values = range.getValues().filter(blank => isNotBlank(blank[0]))

    if (values.length !== 0) // Don't run function if every value is blank, probably means the user pressed the delete key on a large selection
    {
      var data, someSKUsNotFound = false, someUPCsNotFound = false, skus, itemNumber;

      if (isRichmondSpreadsheet(spreadsheet))
      {
        var inventorySheet = spreadsheet.getSheetByName("INVENTORY");
        const numRows = inventorySheet.getLastRow() - 7;
        
        data = inventorySheet.getSheetValues(8, 1, numRows, 8);

        if (/^\d+$/.test(values[0][0]) && (isUPC_A(values[0][0]) || isEAN_13(values[0][0]))) // Check if the values pasted are UPC codes
        {
          const upcDatabaseSheet = spreadsheet.getSheetByName('UPC Database');
          const numRows_UPC = upcDatabaseSheet.getLastRow() - 1;
          const upcDatabase = upcDatabaseSheet.getSheetValues(2, 1, numRows_UPC, 3);

          skus = values.map(item => {

            for (var i = 0; i < numRows_UPC; i++)
            {
              if (upcDatabase[i][0] == item[0])
              {
                itemNumber = upcDatabase[i][2].split(" - ").pop().toString().toUpperCase();

                for (var j = 0; j < numRows; j++)
                {
                  if (data[j][7] == itemNumber)
                    return [data[j][0], data[j][1], data[j][2], data[j][3], data[j][4], data[j][5], data[j][6]];
                }
              }
            }

            someUPCsNotFound = true;

            return ['UPC Not Found:', item[0], '', '', '', '', '']
          });
        }
        else if (values[0][0].toString().includes(' - ')) // Strip the sku from the first part of the google description
        {
          skus = values.map(item => {
          
            for (var i = 0; i < numRows; i++)
            {
              // After checking the SKU, check the description (assuming it could be the vendor's product code
              if (data[i][7] == item[0].toString().split(' - ').pop().toUpperCase() || data[i][1].includes(item[0].toString()))
                return [data[i][0], data[i][1], data[i][2], data[i][3], data[i][4], data[i][5], data[i][6]]
            }

            someSKUsNotFound = true;

            return ['SKU Not Found:', item[0].toString().split(' - ').pop().toUpperCase(), '', '', '', '', '']
          });
        }
        else if (values[0][0].toString().includes('-')) // The SKU contains dashes because that's the convention from Adagio
        {
          skus = values.map(sku => (sku[0].substring(0,4) + sku[0].substring(5,9) + sku[0].substring(10)).trim()).map(item => {
          
            for (var i = 0; i < numRows; i++)
            {
              // After checking the SKU, check the description (assuming it could be the vendor's product code
              if (data[i][7] == item.toString().toUpperCase() || data[i][1].includes(item[0].toString()))
                return [data[i][0], data[i][1], data[i][2], data[i][3], data[i][4], data[i][5], data[i][6]]
            }

            someSKUsNotFound = true;

            return ['SKU Not Found:', item, '', '', '', '', '']
          });
        }
        else // The regular plain SKUs are being pasted
        {
          skus = values.map(item => {
          
            for (var i = 0; i < numRows; i++)
            {
              // After checking the SKU, check the description (assuming it could be the vendor's product code
              if (data[i][7] == item[0].toString().toUpperCase() || data[i][1].includes(item[0].toString()))
                return [data[i][0], data[i][1], data[i][2], data[i][3], data[i][4], data[i][5], data[i][6]]
            }

            someSKUsNotFound = true;

            return ['SKU Not Found:', item[0], '', '', '', '', '']
          });
        }

        if (someSKUsNotFound)
        {
          const skusNotFound = [];
          var isSkuFound;

          const skusFound = skus.filter(item => {
            isSkuFound = item[0] !== 'SKU Not Found:'

            if (!isSkuFound)
              skusNotFound.push(item)

            return isSkuFound;
          })

          const numSkusFound = skusFound.length;
          const numSkusNotFound = skusNotFound.length;
          const items = [].concat.apply([], [skusNotFound, skusFound]); // Concatenate all of the item values as a 2-D array
          const numItems = items.length

          sheet.getRange(1, 2, 1, 2).clearContent() // Clear the search bar
            .offset(3, -1, MAX_NUM_ITEMS, 7).clearContent().setBackground('white').setFontColor('black').setBorder(true, true, true, true, false, false)
            .offset(0, 0, numItems, 7)
              .setFontFamily('Arial').setFontWeight('bold').setFontSizes(new Array(numItems).fill([10, 10, 10, 10, 10, 10, 10]))
              .setHorizontalAlignments(new Array(numItems).fill(['center', 'left', 'center', 'center', 'center', 'center', 'center'])).setVerticalAlignment('middle')
              .setBackgrounds([].concat.apply([], [new Array(numSkusNotFound).fill(new Array(7).fill('#ffe599')), new Array(numSkusFound).fill(new Array(7).fill('white'))]))
              .setBorder(null, null, false, null, false, false).setValues(items)
            .offset((numSkusFound != 0) ? numSkusNotFound : 0, 0, (numSkusFound != 0) ? numSkusFound : numSkusNotFound, 7).activate()
            .offset((numSkusFound != 0) ? -1*numSkusNotFound - 3: -3, 0, 1, 1).setValue((numSkusFound !== 1) ? numSkusFound + " results found." : numSkusFound + " result found.")
            .offset(1, 0, 2, 1).setValue((new Date().getTime() - startTime)/1000 + " s"); // Function runtime
        }
        else if (someUPCsNotFound)
        {
          const upcsNotFound = [];
          var isUpcFound;

          const upcsFound = skus.filter(item => {
            isUpcFound = item[0] !== 'UPC Not Found:'

            if (!isUpcFound)
              upcsNotFound.push(item)

            return isUpcFound;
          })

          const numUpcsFound = upcsFound.length;
          const numUpcsNotFound = upcsNotFound.length;
          const items = [].concat.apply([], [upcsNotFound, upcsFound]); // Concatenate all of the item values as a 2-D array
          const numItems = items.length

          sheet.getRange(1, 2, 1, 2).clearContent() // Clear the search bar
            .offset(3, -1, MAX_NUM_ITEMS, 7).clearContent().setBackground('white').setFontColor('black').setBorder(true, true, true, true, false, false)
            .offset(0, 0, numItems, 7)
              .setFontFamily('Arial').setFontWeight('bold').setFontSizes(new Array(numItems).fill([10, 10, 10, 10, 10, 10, 10]))
              .setHorizontalAlignments(new Array(numItems).fill(['center', 'left', 'center', 'center', 'center', 'center', 'center'])).setVerticalAlignment('middle')
              .setBackgrounds([].concat.apply([], [new Array(numUpcsNotFound).fill(new Array(7).fill('#ffe599')), new Array(numUpcsFound).fill(new Array(7).fill('white'))]))
              .setBorder(null, null, false, null, false, false).setValues(items)
            .offset((numUpcsFound != 0) ? numUpcsNotFound : 0, 0, (numUpcsFound != 0) ? numUpcsFound : numUpcsNotFound, 7).activate()
            .offset((numUpcsFound != 0) ? -1*numUpcsNotFound - 3: -3, 0, 1, 1).setValue((numUpcsFound !== 1) ? numUpcsFound + " results found." : numUpcsFound + " result found.")
            .offset(1, 0, 2, 1).setValue((new Date().getTime() - startTime)/1000 + " s"); // Function runtime
        }
        else // All SKUs were succefully found
        {
          const numItems = skus.length

          sheet.getRange(1, 2, 1, 2).clearContent() // Clear the search bar
            .offset(3, -1, MAX_NUM_ITEMS, 7).clearContent().setBackground('white').setFontColor('black')
            .offset(0, 0, numItems, 7).setFontFamily('Arial').setFontWeight('bold')
              .setFontSizes(new Array(numItems).fill([10, 10, 10, 10, 10, 10, 10]))
              .setHorizontalAlignments(new Array(numItems).fill(['center', 'left', 'center', 'center', 'center', 'center', 'center']))
              .setVerticalAlignment('middle')
              .setBorder(null, null, false, null, false, false).setValues(skus).activate()
            .offset(-3, 0, 1, 1).setValue((numItems !== 1) ? numItems + " results found." : numItems + " result found.")
            .offset(1, 0, 2, 1).setValue((new Date().getTime() - startTime)/1000 + " s"); // Function runtime
        }
      }
      else // Parksville or Rupert
      {
        var inventorySheet = spreadsheet.getSheetByName("SearchData");
        const numRows = inventorySheet.getLastRow() - 1;
        data = inventorySheet.getSheetValues(2, 1, numRows, 7);
        var columnIndex = (isParksvilleSpreadsheet(spreadsheet)) ? [4, 3, 5, 6] : [5, 3, 4, 6];
        
        if (/^\d+$/.test(values[0][0]) && (isUPC_A(values[0][0]) || isEAN_13(values[0][0]))) // Check if the values pasted are UPC codes
        {
          const upcDatabaseSheet = spreadsheet.getSheetByName('UPC Database');
          const numRows_UPC = upcDatabaseSheet.getLastRow() - 1;
          const upcDatabase = upcDatabaseSheet.getSheetValues(2, 1, numRows_UPC, 3);

          skus = values.map(item => {

            for (var i = 0; i < numRows_UPC; i++)
            {
              if (upcDatabase[i][0] == item[0])
              {
                itemNumber = upcDatabase[i][2].split(" - ").pop().toString().toUpperCase();

                for (var j = 0; j < numRows; j++)
                {
                  if (data[j][7] == itemNumber)
                    return [data[j][0], data[j][1], data[j][2], data[j][3], data[j][4], data[j][5], data[j][6]];
                }
              }
            }

            someUPCsNotFound = true;

            return ['UPC Not Found:', item[0], '', '', '', '', '']
          });
        }
        else if (values[0][0].toString().includes(' - ')) // Strip the sku from the first part of the google description
        {
          skus = values.map(item => {
          
            for (var i = 0; i < numRows; i++)
            {
              // After checking the SKU, check the description (assuming it could be the vendor's product code
              if (data[i][1].toString().split(' - ').pop().toUpperCase() == item[0].toString().split(' - ').pop().toUpperCase() || data[i][1].includes(item[0].toString()))
                return [data[i][0], data[i][1], data[i][2], ...columnIndex.map(col => data[i][col]), '']
            }

            someSKUsNotFound = true;

            return ['SKU Not Found:', item[0].toString().split(' - ').pop().toUpperCase(), '', '', '', '', '', '']
          });
        }
        else if (values[0][0].toString().includes('-'))
        {
          skus = values.map(sku => (sku[0].substring(0,4) + sku[0].substring(5,9) + sku[0].substring(10)).trim()).map(item => {
          
            for (var i = 0; i < numRows; i++)
            {
              // After checking the SKU, check the description (assuming it could be the vendor's product code
              if (data[i][1].toString().split(' - ').pop().toUpperCase() == item.toString().toUpperCase() || data[i][1].includes(item[0].toString()))
                return [data[i][0], data[i][1], data[i][2],  ...columnIndex.map(col => data[i][col]), '']
            }

            someSKUsNotFound = true;

            return ['SKU Not Found:', item, '', '', '', '', '', '']
          });
        }
        else
        {
          skus = values.map(item => {
          
            for (var i = 0; i < numRows; i++)
            {
              // After checking the SKU, check the description (assuming it could be the vendor's product code
              if (data[i][1].toString().split(' - ').pop().toUpperCase() == item[0].toString().toUpperCase() || data[i][1].includes(item[0].toString()))
                return [data[i][0], data[i][1], data[i][2], ...columnIndex.map(col => data[i][col]), '']
            }

            someSKUsNotFound = true;

            return ['SKU Not Found:', item[0], '', '', '', '', '', '']
          });
        }

        if (someSKUsNotFound)
        {
          const skusNotFound = [];
          var isSkuFound;

          const skusFound = skus.filter(item => {
            isSkuFound = item[0] !== 'SKU Not Found:'

            if (!isSkuFound)
              skusNotFound.push(item)

            return isSkuFound;
          })

          const numSkusFound = skusFound.length;
          const orderSheet = spreadsheet.getSheetByName("Order");
          const shippedSheet = spreadsheet.getSheetByName("Shipped");
          const numOrderedItems = orderSheet.getLastRow() - 3;
          const numShippedItems = shippedSheet.getLastRow() - 3;
          const orderedItems =   orderSheet.getSheetValues(4, 5, numOrderedItems, 1).map(u => u[0].split(' - ').pop().toString().toUpperCase()); // The items on the order sheet
          const shippedItems = shippedSheet.getSheetValues(4, 5, numShippedItems, 1).map(u => u[0].split(' - ').pop().toString().toUpperCase()); // The items on the shipped sheet
          const backgroundColours = [], fontColours = [];
          var isOnOrderPage, isOnShippedPage;

          for (var i = 0; i < numSkusFound; i++) // Loop through the skus that were pasted
          {
            for (var o = 0; o < numOrderedItems; o++) // Check if the item is on the order page
            {
              if (orderedItems[o] === skusFound[i][1].split(' - ').pop().toString().toUpperCase())
              {
                isOnOrderPage = true
                break;
              }
              isOnOrderPage = false;
            }
            for (var s = 0; s < numShippedItems; s++) // Check if the item is on the shipped page
            {
              if (shippedItems[s] === skusFound[i][1].split(' - ').pop().toString().toUpperCase())
              {
                isOnShippedPage = true
                break;
              }
              isOnShippedPage = false;
            }

            if (isOnShippedPage)
            {
              skusFound[i][7] = 'SHIPPED - On it\'s way';
              backgroundColours.push([...new Array(8).fill('#cc0000')]) // Highlight red
              fontColours.push([...new Array(8).fill('yellow')])        // Yellow font
            }
            else if (isOnOrderPage)
            {
              skusFound[i][7] = 'Already on OrderSheet';
              backgroundColours.push([...new Array(8).fill('yellow')]) // Highlight yellow
              fontColours.push([...new Array(8).fill('#cc0000')])      // Red font
            }
            else // The item is neither on the shipped nor the ordered page
            {
              backgroundColours.push([...new Array(8).fill('white')])
              fontColours.push([...new Array(8).fill('black')])
            }
          }

          const numSkusNotFound = skusNotFound.length;
          const items = [].concat.apply([], [skusNotFound, skusFound]); // Concatenate all of the item values as a 2-D array
          const numItems = items.length
          fontColours.unshift(...new Array(numSkusNotFound).fill(new Array(8).fill('black')))

          sheet.getRange(1, 2, 1, 2).clearContent() // Clear the search bar
            .offset(3, -1, MAX_NUM_ITEMS, 8).clearContent().setBackground('white').setFontColor('black').setBorder(true, true, true, true, false, false)
            .offset(0, 0, numItems, 8)
              .setFontFamily('Arial').setFontWeight('bold')
                .setFontSizes(new Array(numItems).fill([10, 10, 10, 10, 10, 10, 10, 12]))
                .setHorizontalAlignments(new Array(numItems).fill(['center', 'left', 'center', 'center', 'center', 'center', 'center', 'center'])).setVerticalAlignment('middle')
                .setBackgrounds([].concat.apply([], [new Array(numSkusNotFound).fill(new Array(8).fill('#ffe599')), backgroundColours]))
                .setFontColors(fontColours).setBorder(null, null, false, null, false, false)
                .setValues(items)
            .offset((numSkusFound != 0) ? numSkusNotFound : 0, 0, (numSkusFound != 0) ? numSkusFound : numSkusNotFound, 8).activate()
            .offset((numSkusFound != 0) ? -1*numSkusNotFound - 3: -3, 0, 1, 1).setValue((numSkusFound !== 1) ? numSkusFound + " results found." : numSkusFound + " result found.")
            .offset(1, 0, 2, 1).setValues([[null], (new Date().getTime() - startTime)/1000 + " s"])
        }
        else if (someUPCsNotFound)
        {
          const upcsNotFound = [];
          var isUpcFound;

          const upcsFound = skus.filter(item => {
            isUpcFound = item[0] !== 'UPC Not Found:'

            if (!isUpcFound)
              upcsNotFound.push(item)

            return isUpcFound;
          })

          const numUpcsFound = upcsFound.length;
          const orderSheet = spreadsheet.getSheetByName("Order");
          const shippedSheet = spreadsheet.getSheetByName("Shipped");
          const numOrderedItems = orderSheet.getLastRow() - 3;
          const numShippedItems = shippedSheet.getLastRow() - 3;
          const orderedItems =   orderSheet.getSheetValues(4, 5, numOrderedItems, 1).map(u => u[0].split(' - ').pop().toString().toUpperCase()); // The items on the order sheet
          const shippedItems = shippedSheet.getSheetValues(4, 5, numShippedItems, 1).map(u => u[0].split(' - ').pop().toString().toUpperCase()); // The items on the shipped sheet
          const backgroundColours = [], fontColours = [];
          var isOnOrderPage, isOnShippedPage;

          for (var i = 0; i < numUpcsFound; i++) // Loop through the skus that were pasted
          {
            for (var o = 0; o < numOrderedItems; o++) // Check if the item is on the order page
            {
              if (orderedItems[o] === upcsFound[i][1].split(' - ').pop().toString().toUpperCase())
              {
                isOnOrderPage = true
                break;
              }
              isOnOrderPage = false;
            }
            for (var s = 0; s < numShippedItems; s++) // Check if the item is on the shipped page
            {
              if (shippedItems[s] === upcsFound[i][1].split(' - ').pop().toString().toUpperCase())
              {
                isOnShippedPage = true
                break;
              }
              isOnShippedPage = false;
            }

            if (isOnShippedPage)
            {
              upcsFound[i][7] = 'SHIPPED - On it\'s way';
              backgroundColours.push([...new Array(8).fill('#cc0000')]) // Highlight red
              fontColours.push([...new Array(8).fill('yellow')])        // Yellow font
            }
            else if (isOnOrderPage)
            {
              upcsFound[i][7] = 'Already on OrderSheet';
              backgroundColours.push([...new Array(8).fill('yellow')]) // Highlight yellow
              fontColours.push([...new Array(8).fill('#cc0000')])      // Red font
            }
            else // The item is neither on the shipped nor the ordered page
            {
              backgroundColours.push([...new Array(8).fill('white')])
              fontColours.push([...new Array(8).fill('black')])
            }
          }

          const numUpcsNotFound = upcsNotFound.length;
          const items = [].concat.apply([], [upcsNotFound, upcsFound]); // Concatenate all of the item values as a 2-D array
          const numItems = items.length
          fontColours.unshift(...new Array(numUpcsNotFound).fill(new Array(8).fill('black')))

          sheet.getRange(1, 2, 1, 2).clearContent() // Clear the search bar
            .offset(3, -1, MAX_NUM_ITEMS, 8).clearContent().setBackground('white').setFontColor('black').setBorder(true, true, true, true, false, false)
            .offset(0, 0, numItems, 8)
              .setFontFamily('Arial').setFontWeight('bold')
                .setFontSizes(new Array(numItems).fill([10, 10, 10, 10, 10, 10, 10, 12]))
                .setHorizontalAlignments(new Array(numItems).fill(['center', 'left', 'center', 'center', 'center', 'center', 'center', 'center'])).setVerticalAlignment('middle')
                .setBackgrounds([].concat.apply([], [new Array(numUpcsNotFound).fill(new Array(8).fill('#ffe599')), backgroundColours]))
                .setFontColors(fontColours).setBorder(null, null, false, null, false, false)
                .setValues(items)
            .offset((numUpcsFound != 0) ? numUpcsNotFound : 0, 0, (numUpcsFound != 0) ? numUpcsFound : numUpcsNotFound, 8).activate()
            .offset((numUpcsFound != 0) ? -1*numUpcsNotFound - 3: -3, 0, 1, 1).setValue((numUpcsFound !== 1) ? numUpcsFound + " results found." : numUpcsFound + " result found.")
            .offset(1, 0, 2, 1).setValues([[null], (new Date().getTime() - startTime)/1000 + " s"])
        }
        else // All SKUs were succefully found
        {
          const numItems = skus.length
          const orderSheet = spreadsheet.getSheetByName("Order");
          const shippedSheet = spreadsheet.getSheetByName("Shipped");
          const numOrderedItems = orderSheet.getLastRow() - 3;
          const numShippedItems = shippedSheet.getLastRow() - 3;
          const orderedItems =   orderSheet.getSheetValues(4, 5, numOrderedItems, 1).map(u => u[0].split(' - ').pop().toString().toUpperCase()); // The items on the order sheet
          const shippedItems = shippedSheet.getSheetValues(4, 5, numShippedItems, 1).map(u => u[0].split(' - ').pop().toString().toUpperCase()); // The items on the shipped sheet
          const backgroundColours = [], fontColours = [];
          var isOnOrderPage, isOnShippedPage;

          for (var i = 0; i < numItems; i++) // Loop through the skus that were pasted
          {
            for (var o = 0; o < numOrderedItems; o++) // Check if the item is on the order page
            {
              if (orderedItems[o] === skus[i][1].split(' - ').pop().toString().toUpperCase())
              {
                isOnOrderPage = true
                break;
              }
              isOnOrderPage = false;
            }
            for (var s = 0; s < numShippedItems; s++) // Check if the item is on the shipped page
            {
              if (shippedItems[s] === skus[i][1].split(' - ').pop().toString().toUpperCase())
              {
                isOnShippedPage = true
                break;
              }
              isOnShippedPage = false;
            }

            if (isOnShippedPage)
            {
              skus[i][7] = 'SHIPPED - On it\'s way';
              backgroundColours.push([...new Array(8).fill('#cc0000')]) // Highlight red
              fontColours.push([...new Array(8).fill('yellow')])        // Yellow font
            }
            else if (isOnOrderPage)
            {
              skus[i][7] = 'Already on OrderSheet';
              backgroundColours.push([...new Array(8).fill('yellow')]) // Highlight yellow
              fontColours.push([...new Array(8).fill('#cc0000')])      // Red font
            }
            else // The item is neither on the shipped nor the ordered page
            {
              backgroundColours.push([...new Array(8).fill('white')])
              fontColours.push([...new Array(8).fill('black')])
            }
          }

          sheet.getRange(1, 2, 1, 2).clearContent() // Clear the search bar
            .offset(3, -1, MAX_NUM_ITEMS, 8).clearContent().setBackground('white').setFontColor('black')
            .offset(0, 0, numItems, 8).setFontFamily('Arial').setFontWeight('bold')
              .setFontSizes(new Array(numItems).fill([10, 10, 10, 10, 10, 10, 10, 12]))
              .setHorizontalAlignments(new Array(numItems).fill(['center', 'left', 'center', 'center', 'center', 'center', 'center', 'center'])).setVerticalAlignment('middle')
              .setBackgrounds(backgroundColours).setFontColors(fontColours).setBorder(null, null, false, null, false, false).setValues(skus).activate()
            .offset(-3, 0, 1, 1).setValue((numItems !== 1) ? numItems + " results found." : numItems + " result found.")
            .offset( 1, 0, 2, 1).setValues([[null], [(new Date().getTime() - startTime)/1000 + " s"]])
        }
      }
    }

    spreadsheet.toast('Searching Complete.')
  }
}

/**
 * This function sends an email to the relevant people at the particular branch store when the status of an item on the Order page is changed to either
 * Order From Distributor or Discontinued or Item Not Available. 
 * 
 * @param {String}         status   : The status of the item.
 * @param {Number}          row     : The row on the order sheet that has had the status change.
 * @param {String[]}     rowValues  : All of the values from the row that had the status change.
 * @param {Sheet}          sheet    : The order sheet.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function sendEmailToBranchStore(status, row, rowValues, sheet, spreadsheet)
{
  const htmlTemplate = HtmlService.createTemplateFromFile('shipmentStatusChangeEmail');
  htmlTemplate.orderDateBackgroundColour = sheet.getRange(row, 1).getBackground();
  htmlTemplate.notesBackgroundColour = sheet.getRange(row, 6).getBackground();
  htmlTemplate.orderDate = Utilities.formatDate(rowValues[0][0], spreadsheet.getSpreadsheetTimeZone(), "dd MMM yyyy");
  htmlTemplate.enteredBy = rowValues[0][1];
  htmlTemplate.qty = rowValues[0][2];
  htmlTemplate.uom = rowValues[0][3];
  htmlTemplate.description = rowValues[0][4];
  htmlTemplate.notes = rowValues[0][5];
  htmlTemplate.currentStock = rowValues[0][6];
  htmlTemplate.actualStock = rowValues[0][7];
  htmlTemplate.shipmentStatus = status;
  htmlTemplate.url = spreadsheet.getUrl() + '#gid=' + spreadsheet.getSheetByName("Shipped").getSheetId();

  if (isParksvilleSpreadsheet(spreadsheet))
  {
    if (MailApp.getRemainingDailyQuota() > 5)
      MailApp.sendEmail({
        to: "eryn@pacificnetandtwine.com, lodgesales@pacificnetandtwine.com, noah@pacificnetandtwine.com, shane@pacificnetandtwine.com, pntparksville@gmail.com, parksville@pacificnetandtwine.com",
        subject: "Shipment Status Change on the Transfer Sheet: " + status,
        htmlBody: htmlTemplate.evaluate().getContent(),
      });
  }
  else if (MailApp.getRemainingDailyQuota() > 1)
  {
    MailApp.sendEmail({
      to: "pr@pacificnetandtwine.com, pntrupert@gmail.com",
      subject: "Shipment Status Change on the Transfer Sheet: " + status,
      htmlBody: htmlTemplate.evaluate().getContent(),
    });
  }
}

/**
 * This function sends an email to the relevant people at Trites and asks them if they have stock of the selected items. It also adds a timestamp to the Notes
 * column informing everyone that Trites has been contacted about the items.
 * 
 * @author Jarren Ralf
 */
function sendEmailToTrites()
{
  const activeRanges = SpreadsheetApp.getActiveRangeList().getRanges(); // The selected ranges on the item search sheet

  if (SpreadsheetApp.getActiveSheet().getSheetName() === "Order" && Math.min(...activeRanges.map(rng => rng.getRow())) > 3) // If the user has not selected an item, alert them with an error message
  { 
    const spreadsheet = SpreadsheetApp.getActive();
    const timeZone = spreadsheet.getSpreadsheetTimeZone();
    const emailTimestamp = "\n*Email Sent to Trites on " + Utilities.formatDate(new Date(), timeZone, "dd MMM yyyy")+"*";
    const emailTimestamp_TextStyle = SpreadsheetApp.newTextStyle().setBold(true).setFontFamily('Arial').setFontSize(10).setForegroundColor('#cc0000').setUnderline(true).build();
    var range, notesRange, richText_Notes, richText_Notes_Runs, fullText, fullTextLength, numRuns, backgroundColours = [], row = []; // 

    const itemValues = [].concat.apply([], activeRanges.map(rng => {
        row.push(rng.getRow()) // We will use this row number to hyperlink the user to this cell through the email url
        range = rng.offset(0, 1 - rng.getColumn(), rng.getNumRows(), 6);
        notesRange = rng.offset(0, 6 - rng.getColumn(), rng.getNumRows(), 1);
        backgroundColours.push(...range.getBackgrounds());

        richText_Notes = notesRange.getRichTextValues().map(note_RichText => {
          fullText = note_RichText[0].getText()
          fullTextLength = fullText.length;
          richText_Notes_Runs = note_RichText[0].getRuns().map(run => [run.getStartIndex(), run.getEndIndex(), run.getTextStyle()]);
          numRuns = richText_Notes_Runs.length;

          if (isNotBlank(fullText))
            for (var i = 0, richTextBuilder = SpreadsheetApp.newRichTextValue().setText(fullText + emailTimestamp); i < numRuns; i++)
              richTextBuilder.setTextStyle(richText_Notes_Runs[i][0], richText_Notes_Runs[i][1], richText_Notes_Runs[i][2])
          else 
            return [SpreadsheetApp.newRichTextValue().setText("*Email Sent to Trites on " + Utilities.formatDate(new Date(), timeZone, "dd MMM yyyy")+"*").setTextStyle(emailTimestamp_TextStyle).build()]
          
          return [richTextBuilder.setTextStyle(fullTextLength + 1, fullTextLength + emailTimestamp.length, emailTimestamp_TextStyle).build()];
        })

        notesRange.setRichTextValues(richText_Notes).setBackgrounds(notesRange.getBackgrounds());

        return range.getValues()
      })
    );

    const pntStoreLocation = (isParksvilleSpreadsheet(spreadsheet)) ? 'Parksville' : 'Prince Rupert'
    const gid = spreadsheet.getSheetByName('Order').getSheetId();
    const htmlTemplate = HtmlService.createTemplateFromFile('tritesStockCheckEmail')
    htmlTemplate.storeLocation = pntStoreLocation;
    htmlTemplate.transferSheetUrl = spreadsheet.getUrl() + '?gid=' + gid + '#gid=' + gid + '&range=I' + row.shift();
    const htmlOutput = htmlTemplate.evaluate();
    const numItems = itemValues.length;
    
    for (var i = 0; i < numItems; i++)
      htmlOutput.append(
        '<tr style="height: 20px">'+
        '<td class="s4" dir="ltr" style="background-color:' + backgroundColours[i][0] + '">' + 
          Utilities.formatDate(itemValues[i][0], timeZone, "dd MMM yyyy") + '</td>' +
        '<td class="s5" dir="ltr">' + itemValues[i][1] + '</td>'+
        '<td class="s5" dir="ltr">' + itemValues[i][2] + '</td>'+
        '<td class="s6" dir="ltr">' + itemValues[i][3] + '</td>'+
        '<td class="s7" dir="ltr">' + itemValues[i][4] + '</td>'+
        '<td class="s8" dir="ltr" style="background-color:' + backgroundColours[i][5] + '">' + itemValues[i][5] + '</td></tr>'
      )

    htmlOutput.append('</tbody></table></div>')

    if (MailApp.getRemainingDailyQuota() > 5)
      MailApp.sendEmail({
        to: "scottnakashima@hotmail.com, jarren@pacificnetandtwine.com",
        cc: "mark@pacificnetandtwine.com, warehouse@pacificnetandtwine.com, triteswarehouse@pacificnetandtwine.com",
        subject: pntStoreLocation + " store has ordered the following items. Do you have any of them at Trites?",
        htmlBody: htmlOutput.getContent(),
      });
  }
  else
    Browser.msgBox('Please select an item or items on the Order page.')
}

/**
* Sorts data by the categories while ignoring capitals and pushing blanks to the bottom of the list.
*
* @param  {String[]} a : The current array value to compare
* @param  {String[]} b : The next array value to compare
* @return {String[][]} The output data.
* @author Jarren Ralf
*/
function sortByCategories(a, b)
{
  return (a[9].toLowerCase() === b[9].toLowerCase()) ? 0 : (a[9] === '') ? 1 : (b[9] === '') ? -1 : (a[9].toLowerCase() < b[9].toLowerCase()) ? -1 : 1;
}

/**
* Sorts data by the created date of the product for the parksville and rupert spreadsheets.
*
* @param  {String[]} a : The current array value to compare
* @param  {String[]} b : The next array value to compare
* @return {String[][]} The output data.
* @author Jarren Ralf
*/
function sortByCreatedDate(a, b)
{
  return (a[7] === b[7]) ? 0 : (a[7] < b[7]) ? 1 : -1;
}

/**
* Sorts data by the created date of the product for the richmond spreadsheet.
*
* @param  {String[]} a : The current array value to compare
* @param  {String[]} b : The next array value to compare
* @return {String[][]} The output data.
* @author Jarren Ralf
*/
function sortByCreatedDate_Richmond(a, b)
{
  return (a[8] === b[8]) ? 0 : (a[8] < b[8]) ? 1 : -1;
}

/**
* Sorts data by the created date of the product
*
* @param  {String[]} a : The current array value to compare
* @param  {String[]} b : The next array value to compare
* @return {String[][]} The output data.
* @author Jarren Ralf
*/
function sortByCountedDate(a, b)
{
  return (a[3] === b[3]) ? 0 : (a[3] < b[3]) ? -1 : 1;
}

/**
 * This function takes a set of hoochies and it orders them alpha-numerically.
 * 
 * @param {String[][]} hoochies     : The list of hoochies and any corresponding data.
 * @param {Number} descriptionIndex : The index number of the column that contains the google description.
 * @param {String}  hoochiePrefix   : The prefix of the SKU number which denotes the type of hoochie.
 * @return {String[][]} Returns the list of hoochies sorted alpha-numerically.
 * @author Jarren Ralf
 */
function sortHoochies(hoochies, descriptionIndex, hoochiePrefix)
{
  const splitParameter = {'16060005': '4-1/4"',   '16010005': 'CUTTLEFISH', '16050005': 'NEEDLEFISH', '16020000': 'BAIT',  '16020010': 'SWIVL',
                          '16060065': '6-1/2"',     '16060010': '10"',        '16070000': '6"',       '16075300': '6.7"',  '16070975': '1/2"',
                          '16030000': 'MACKERAL',   '16060175': '1-3/4"',     '16200030': '4-1/4"',   '16200000': '4-3/4', '16200025': 'HOOCHIE',
                          '16200061': 'BAIT',       '16200065': 'BAIT',       '16200021': 'SARDINE',  '16200022': '2.5"'};
  var matchA, matchB;
  
  const values = hoochies.sort((a, b) =>
  {
    // Logger.log(a[descriptionIndex].split(" - ", 1)[0].split(splitParameter[hoochiePrefix], 2))
    // Logger.log(b[descriptionIndex].split(" - ", 1)[0].split(splitParameter[hoochiePrefix], 2))

    if (hoochiePrefix !== 'other') // Sort hoochies
    {
      // Regex to find the prefix letters, then number, then the suffix letters followed by the secondary number e.g. [M35WPS, M, 35, WPS, ] or [J66-1, J, 66, -, 1]
      matchA = a[descriptionIndex].split(" - ", 1)[0].split(splitParameter[hoochiePrefix], 2)[1].trim().match(/(\D*)([\d\.]*)(\D*)([\d\.]*)/);
      matchB = b[descriptionIndex].split(" - ", 1)[0].split(splitParameter[hoochiePrefix], 2)[1].trim().match(/(\D*)([\d\.]*)(\D*)([\d\.]*)/);

      if (parseInt(matchA[2]) === parseInt(matchB[2])) // If the numerals are the same
        if (matchA[1] === matchB[1]) // If the prefix letters are the same
          if (matchA[3] === matchB[3]) // If the suffix letters are the same
            return parseInt(matchA[4]) - parseInt(matchB[4]); // If the numerals and letters are all the same, sort by the secondary number if there is one
          else
            return matchA[3].localeCompare(matchB[3]); // If the numerals and prefix letters are the same, sort by the suffix letters
        else
          return matchA[1].localeCompare(matchB[1]); // If numerals are equal, sort alphabetically by the prefix letters
      else if (matchA[2] === '') // If the numerals are the blank
        return 1;
      else if (matchB[2] === '') // If the numerals are the blank
        return -1;
      else
        return parseInt(matchA[2]) - parseInt(matchB[2]); // Sort numerically
    }
    else // Sort non-hoochies
    {
      matchA = a[descriptionIndex].split(" - ");
      matchB = b[descriptionIndex].split(" - ");

      if (matchA[matchA.length - 4] === matchB[matchA.length - 4])   // If the vendors are the same
        if (matchA[matchA.length - 3] === matchB[matchA.length - 3]) // If the categories are the same
          return matchA[matchA.length - 1].localeCompare(matchB[matchA.length - 1]) // If the vendors and categories are the same, sort by the skus
        else
          return matchA[matchA.length - 3].localeCompare(matchB[matchA.length - 3]) // If the vendors are the same, sort by the categories
      else
        return matchA[matchA.length - 4].localeCompare(matchB[matchA.length - 4]) // Sort by the vendors
    }
  })

  return values;
}

/**
 * This function takes the selected items on the Manual Counts page and it sorts them. If they are hoochies, then it groups them together and sorts them alpha-numerically.
 * If they are non-hoochies, then it sorts by vendor first, then category, and then by SKU number
 * 
 * @author Jarren Ralf
 */
function sort_ManualCounts()
{
  const activeRange = SpreadsheetApp.getActiveRange();

  if (activeRange.getRow() > 3)
  {
    const range = activeRange.offset(0, 1 - activeRange.getColumn(), activeRange.getNumRows(), SpreadsheetApp.getActiveSheet().getMaxColumns());
    const groupedItems = groupHoochieTypes(range.getValues(), 0)
    const items = [];

    Object.keys(groupedItems).forEach(key => items.push(...sortHoochies(groupedItems[key], 0, key)));

    range.setValues(items)
  }
  else
    Browser.msgBox('Select the items that you want sorted.')
}

/**
* This function moves all of the selected values on the item search page to the ItemsToRichmond page.
*
* @author Jarren Ralf
*/
function storeToRichmondTransfers()
{
  const QTY_COL  = 6;
  const NUM_COLS = 3;
  
  var itemsToRichmondSheet = SpreadsheetApp.getActive().getSheetByName("ItemsToRichmond");
  var lastRow = itemsToRichmondSheet.getLastRow();
  
  copySelectedValues(itemsToRichmondSheet, lastRow + 1, NUM_COLS, QTY_COL);
}

/**
* This function transfers the entire chosen row from one sheet to another sheet.
*
* @param   {Sheet}     fromSheet  : The active sheet or the sheet the row is being moved FROM
* @param   {Sheet}       sheet    : The destination sheet or the sheet that the row is being moved TO
* @param   {Number}       row     : The row number
* @param {Object[][]}  rowValues  : The values of the row being moved
* @param   {Number}     numCols   : The number of columns
* @param   {Boolean} isRowDeleted : Whether the original row is deleted or not
* @param   {Number}  rowInsertNum : The row number that the item will be placed under
* @param   {String}     carrier   : The name of the carrier
* @param {Object[][]} carrierBannerRowNumbers : The line numbers of all the carriers
* @param   {Object}   eventObject : The event object generated from the onEdit trigger
* @return {RichTextValue} richText : The function returns the rich text value of the Notes cell
* @author Jarren Ralf
*/
function transferRow(fromSheet, sheet, row, rowValues, numCols, isRowDeleted, rowInsertNum, carrier, carrierBannerRowNumbers, eventObject)
{
  const rowBackgroundColours = fromSheet.getRange(row, 1, 1, numCols).getBackgrounds(); // So the order date and notes can keep the same highlight colour
  const richText = fromSheet.getRange(row, 6).getRichTextValue(); // So the notes can keep the same rich text
  const     sheetName =     sheet.getSheetName();
  const fromSheetName = fromSheet.getSheetName();
  
  if (fromSheetName !== sheetName) rowValues[0][10] = dateStamp(row, 11); // This represents when the item was moved to a different page

  if (sheetName === "Received") // Put all of the receivings at the top of the page
  {
    var destinationRow = 4;
    sheet.insertRowAfter(3);
    applyFullRowFormatting(sheet, destinationRow, 1, numCols); // Make sure all the formatting is correct
    
    if (eventObject.oldValue == undefined || eventObject.oldValue.split(' - ', 1)[0] != "Direct") // If the shipment is direct then make the Transfered column checked (won't show up on Adagio update page)
      sheet.getRange(destinationRow, 12).insertCheckboxes(); // Unchecked
    else
    {
      rowValues[0][9] = "Received Direct";
      sheet.getRange(destinationRow, 12).insertCheckboxes().check();
    }

    sheet.getRange(destinationRow, 1, 1, numCols).setBackgrounds(rowBackgroundColours).setValues(rowValues);
    sheet.autoResizeRows(destinationRow, 1);
  }
  else if (rowInsertNum === undefined) // Insert the row at the bottom of the list
  {
    var destinationRow = sheet.getLastRow() + 1;
    applyFullRowFormatting(sheet, destinationRow, 1, numCols); // Make sure all the formatting is correct
    sheet.getRange(destinationRow, 1, 1, numCols).setBackgrounds(rowBackgroundColours).setValues(rowValues);
    sheet.getRange(destinationRow, 13).setDataValidation(null)
    sheet.autoResizeRows(destinationRow, 1);
  }
  else if (typeof rowInsertNum === 'string') // We must create a new carrier line in addition to moving the row accross
  {
    rowInsertNum = Number(rowInsertNum.replace(/^\D+/g,'')); // Convert the string to a number
    sheet.insertRowsAfter(rowInsertNum, 2);
    var newCarrierRow = rowInsertNum + 1;
    var destinationRow = rowInsertNum + 2;
    applyFullRowFormatting(sheet, destinationRow, 1, numCols); // Make sure all the formatting is correct
    sheet.setRowHeight(newCarrierRow, 40);
    sheet.getRange(newCarrierRow, 1, 2, numCols)
      .setBackgrounds([new Array(numCols).fill('#6d9eeb'), ...rowBackgroundColours])
      .setFontColors([[...new Array(10).fill('white'), '#6d9eeb'], new Array(numCols).fill('black')])
      .setFontSizes([new Array(numCols).fill(14),new Array(numCols).fill(10)])
      .setFontLine('none').setFontWeight('bold').setFontStyle('normal').setFontFamily('Arial')
      .setHorizontalAlignments([new Array(numCols).fill('left'), ['right', ...new Array(3).fill('center'), 'left', ...new Array(6).fill('center')]])
      .setWrapStrategies([new Array(numCols).fill(SpreadsheetApp.WrapStrategy.OVERFLOW), [...new Array(3).fill(SpreadsheetApp.WrapStrategy.OVERFLOW), 
      ...new Array(3).fill(SpreadsheetApp.WrapStrategy.WRAP), ...new Array(3).fill(SpreadsheetApp.WrapStrategy.CLIP), SpreadsheetApp.WrapStrategy.WRAP, SpreadsheetApp.WrapStrategy.CLIP]])
      .setDataValidations([new Array(numCols).fill(null), [...new Array(9).fill(null), sheet.getRange(3, 10).getDataValidation()]])
      .setBorder(true,true,true,true,null,null)
      .setValues([[carrier, ...new Array(9).fill(null), 'via'], ...rowValues]);
    sheet.getRange(newCarrierRow, 13, 2).setDataValidations([[sheet.getRange(3, 13).getDataValidation()],[null]]);
    sheet.getRange(newCarrierRow, 1, 1, numCols - 1).merge();
    sheet.autoResizeRows(destinationRow, 1);
  }
  else // Move a row to the specified row
  {
    sheet.insertRowsAfter(rowInsertNum, 1);
    var destinationRow = rowInsertNum + 1;
    applyFullRowFormatting(sheet, destinationRow, 1, numCols); // Make sure all the formatting is correct
    sheet.getRange(destinationRow, 1, 1, numCols).setBackgrounds(rowBackgroundColours).setValues(rowValues);
    sheet.getRange(destinationRow, 13).setDataValidation(null);
    sheet.autoResizeRows(destinationRow, 1);
  }

  sheet.getRange(destinationRow,  6).setRichTextValue(richText);                                   // Keep the notes rich text the same
  sheet.getRange(destinationRow, 10).setDataValidation(sheet.getRange(3, 10).getDataValidation()); // Set the correct data validation
  
  if (isRowDeleted) 
  {
    if (fromSheetName === "Shipped") // If we are on the shipped page, we need to check if a carrier banner needs to be deleted
    {
      // First check if shipping banner needs to be delete
      const numCarrierRowBanners = carrierBannerRowNumbers.length - 1; // We don't care about the Back to Order data validation.
      const isCarrierBannerAboveRow = row - 1;
      const isCarrierBannerBelowRow = row + 1;
      const previousCarrier = eventObject.oldValue;
      var numRowsToDelete = 1;

      for (var i = 0, j = 0, doesPreviousCarrierBannerExist = 1; i < numCarrierRowBanners; i++)
      {
        if (carrierBannerRowNumbers[i][1] === previousCarrier && typeof carrierBannerRowNumbers[i][0] === 'string') // Check if the previous carrier exists and the row number is a string
          doesPreviousCarrierBannerExist--; 

        if (carrierBannerRowNumbers[i][0] === isCarrierBannerAboveRow || carrierBannerRowNumbers[i][0] === isCarrierBannerBelowRow)
        {
          if (j == 1) // If the above loop was entered twice, then there is a banner directly above and below the item, which means the shipping banner needs to be deleted
          {
            numRowsToDelete = 2;
            row--; // Subract 1 to therefore include the banner in the deleteRows function below
          }
          j++; // Increment the second counter 
        }
      }

      if (doesPreviousCarrierBannerExist)
      {
        if (row > rowInsertNum) // Moving the row up has an effect on what the row number is of the line that needs to be deleted
          (rowInsertNum + 2 === destinationRow) ? fromSheet.deleteRows(row + 2, numRowsToDelete) : fromSheet.deleteRows(row + 1, numRowsToDelete);
        else
          fromSheet.deleteRows(row, numRowsToDelete);
      }
      else
        fromSheet.deleteRows(row + 2, 1); // This means a user manually edited the carrier banner
    }
    else
      fromSheet.deleteRow(row); // Delete the row from the original sheet
  }

  return richText;
}

/**
* This is a function I found and modified to keep the last instance of an item in a muli-array based on the uniqueness of one of the values.
*
* @param      {Object[][]}    arr : The given array
* @param  {Callback Function} key : A function that chooses one of the elements of the object or array
* @return     {Object[][]}    The reduced array containing only unique items based on the key
*/
function uniqByKeepLast(arr, key) {
    return [...new Map(arr.map(x => [key(x), x])).values()]
}

/**
 * This function takes the item that was just scanned on the manual scan page and copies it to the list of UPCs to unmarry from the countersales data.
 * A user interface is launched that accepts the UPC value to unmarry
 * 
 * @author Jarren Ralf
 */
function unmarryUPC()
{
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActive()
  const item = spreadsheet.getActiveSheet().getSheetValues(1, 1, 1, 1)[0][0].toString().split('\n')
  const response = ui.prompt('Unmarry UPCs', 'Please scan the barcode for:\n\n' + item[0] +'.', ui.ButtonSet.OK_CANCEL)

  if (ui.Button.OK === response.getSelectedButton())
  {
    const unmarryUpcSheet = spreadsheet.getSheetByName("UPCs to Unmarry");
    unmarryUpcSheet.getRange(unmarryUpcSheet.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[response.getResponseText(), item[0]]]);
  }
}

/**
* This function updates the Back Order.
*
* @param   {Range}    rowRange  : The range of the entire row
* @param {Object[][]} rowValues : The values of the entire row
* @author Jarren Ralf
*/
function updateBO(rowRange, rowValues)
{ 
  rowValues[0][2] -= rowValues[0][8]; // Update the order quantity by subtracting off the amount shipped
  rowValues[0][8] = '';
  rowValues[0][9] = 'B/O';
  rowRange.setValues(rowValues);
}

/**
 * This function first updates the Recent sheet which contains the last MAX_NUM_ITEMS items that have been created in Adagio.
 * It also includes the data from the count log as to when each item was last counted.
 * 
 * @author Jarren Ralf
 */
function updateRecentlyCreatedItems()
{
  const MAX_NUM_ITEMS = 500;
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName("INVENTORY");
  const numHeaders = (isRichmondSpreadsheet(spreadsheet)) ? 7 : 9;
  const transferData = sheet.getSheetValues(numHeaders + 1, 1, sheet.getLastRow() - numHeaders, 9);
  const today = new Date();
  const countLog = spreadsheet.getSheetByName("Count Log")
  const countLogData = countLog.getSheetValues(2, 1, countLog.getLastRow(), countLog.getLastColumn());
  const mostRecentCounts = uniqByKeepLast(countLogData, sku => sku[0]);
  const numRecentCounts = mostRecentCounts.length;
  var d; // This variable is used in the following filter function and it represents the value in the Created Date column
  
  if (isRichmondSpreadsheet(spreadsheet))
  {
    var recentlyCreatedItems = transferData.map(val => {
      d = val[8].split('.');                           // Split the date at the "."
      val[8] = new Date(d[2],d[1] - 1,d[0]).getTime(); // Convert the date sting to a striong object for sorting purposes
      return val;
    }).sort(sortByCreatedDate_Richmond);

    recentlyCreatedItems.splice(MAX_NUM_ITEMS); // Keep the first MAX_NUM_ITEMS rows

    // Place the dates of the most recently counted items on the recently created items list
    for (var i = 0; i < MAX_NUM_ITEMS; i++)
    {
      for (var j = 0; j < numRecentCounts; j++)
      {
        if (recentlyCreatedItems[i][7] == mostRecentCounts[j][0])
          recentlyCreatedItems[i][2] = mostRecentCounts[j][3];
      }
    }

    recentlyCreatedItems = recentlyCreatedItems.map(f => f.slice(0, 7)); // Keep the price unit, item description, and inventory columns
  }
  else
  {
    var [currentStock, otherStoreStock] = (isParksvilleSpreadsheet(spreadsheet)) ? [3, 4] : [4, 3]; // Make sure the column references are correct for parksville and rupert

    var recentlyCreatedItems = transferData.filter(val => {
      d = val[7].split('.');                           // Split the date at the "."
      val[7] = new Date(d[2],d[1] - 1,d[0]).getTime(); // Convert the date sting to a striong object for sorting purposes
      return val[8] !== "No TS";
    }).sort(sortByCreatedDate);

    recentlyCreatedItems.splice(MAX_NUM_ITEMS); // Keep the first MAX_NUM_ITEMS rows
    
    // Place the dates of the most recently counted items on the recently created items list
    for (var i = 0; i < MAX_NUM_ITEMS; i++)
    {
      for (var j = 0; j < numRecentCounts; j++)
      {
        if (recentlyCreatedItems[i][6] == mostRecentCounts[j][0])
          recentlyCreatedItems[i][8] = mostRecentCounts[j][3];
      }
    }

    recentlyCreatedItems = recentlyCreatedItems.map(f => [...f.slice(0, 2), f[8], f[currentStock], f[2], f[otherStoreStock], f[5]]);
  }

  spreadsheet.getSheetByName("Recent").getRange(2, 1, MAX_NUM_ITEMS, 7).setNumberFormat('@').setValues(recentlyCreatedItems);
  sheet.getRange(3, 1).setValue('The most recently created item list was last updated at ' + today.toLocaleTimeString() + ' on ' +  today.toDateString())
}

/**
 * This function updates the search data with the date which particular items were last counted.
 * 
 * @author Jarren Ralf
 */
function updateSearchData()
{
  const today = new Date();
  const spreadsheet = SpreadsheetApp.getActive();
  const searchDataRng = (isRichmondSpreadsheet(spreadsheet)) ? spreadsheet.getSheetByName("INVENTORY").getRange('B7:C') : spreadsheet.getSheetByName("SearchData").getRange('B1:C');
  const searchData = searchDataRng.getValues();
  const numItems = searchData.length;
  const countLog = spreadsheet.getSheetByName("Count Log");
  const numOldCounts = countLog.getLastRow() - 1;
  const countLogRange = countLog.getRange(2, 1, numOldCounts, 4);
  const countLogData = countLogRange.getValues().sort(sortByCountedDate);
  const mostRecentCounts = uniqByKeepLast(countLogData, sku => sku[0]); // Remove duplicates
  const numNewCounts = mostRecentCounts.length;
  const numberFormats = [...Array(numItems)].map(e => ['@', 'dd MMM yyyy']);
  countLogRange.clearContent();
  countLog.getRange(2, 1, numNewCounts, 4).setValues(mostRecentCounts);
  searchData[0][1] = "Last Counted On";
  numberFormats[0][1] = '@';

  for (var i = 1; i < numItems; i++)
  {
    for (var j = 0; j < numNewCounts; j++)
    {
      if (searchData[i][0].split(' - ').pop().toString() == mostRecentCounts[j][0])
        searchData[i][1] = mostRecentCounts[j][3];
    }
  }
  
  searchDataRng.setNumberFormats(numberFormats).setValues(searchData);
  spreadsheet.getSheetByName("INVENTORY").getRange(6, 1).setValue('The Recent Counts were last updated at ' + today.toLocaleTimeString() + ' on ' +  today.toDateString());

  if (numOldCounts > numNewCounts)
    countLog.deleteRows(numNewCounts + 2, numOldCounts - numNewCounts); // Delete the blank rows
}

/**
 * This function first updates the UPC Database sheet by importing the csv with all of the UPC codes, and then adding the Adagio descriptions 
 * and current stock (from the INVENTORY page) to the same double array. The INVENTORY page is only time-stamped when the user clicks on the 'Update
 * UPC Database' button to run the function, otherwise when the function is run by the trigger, no date is written to the cell.
 * 
 * @param {Boolean} isButtonClicked : A boolean variable that represents whether the button was clicked or not (run manually or by a trigger otherwise)
 * @author Jarren Ralf
 */
function updateUPC_Database(isButtonClicked)
{
  const today = new Date();
  const spreadsheet = SpreadsheetApp.getActive();
  const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
  const inventorySheet = spreadsheet.getSheetByName("INVENTORY");
  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("BarcodeInput.csv").next().getBlob().getDataAsString());
  var header = csvData.shift();
  var isInAdagioDatabase; // Boolean variable that checks if the SKU from the upc database is in the adagio database

  if (isRichmondSpreadsheet(spreadsheet))
  {
    var currentStock = 2; // Changes the index number for selecting the current stock from inventory data
    var transferData = inventorySheet.getRange('B8:H').getValues();
    var sku = 6; // The index of the sku
  }
  else
  {
    var currentStock = (isParksvilleSpreadsheet(spreadsheet)) ? 2 : 3; // Changes the index number for selecting the current stock from inventory data
    var transferData = inventorySheet.getRange('B10:G').getValues();
    var sku = 5;// The index of the sku
  }

  // Replace the csvData with the Adagio descriptions and current stock values
  const data = csvData.filter(v => {
    return transferData.filter(u => {
      isInAdagioDatabase = u[sku] == v[1].toString().toUpperCase(); // Match the SKU
      if (!isInAdagioDatabase) return isInAdagioDatabase; // If the SKU isn't found in the Adagio database, return false
      v[1] = v[3];            // Move the Item Unit to column 2
      v[2] = u[0];            // Move the Item Unit to column 2
      v[3] = u[currentStock]; // Move the Current Stock to column 4
      return isInAdagioDatabase;
    }).length != 0; // Keep only the items in the UPC database that have found a matching sku in Adagio
  })

  header[1] = "Item Unit";
  header[2] = "Adagio Description";
  header[3] = "Current Stock";
  const numRows = data.unshift(header); // Put the header back at the top of the database
  upcDatabaseSheet.clearContents().getRange(1, 1, numRows, 4).setNumberFormat('@').setValues(data);

  if (isButtonClicked)
    inventorySheet.getRange(4, 1).setFontColor('black').setValue('The UPC Database was last updated at ' + today.toLocaleTimeString() + ' on ' +  today.toDateString());
}

/**
 * This function runs the updateUPC_Database function by sending the function a true boolean. This will ensure that the INVENTORY page gets time stamped.
 * 
 * @author Jarren Ralf
 */
function updateUPC_Database_ButtonClicked()
{
  updateUPC_Database(true);
}

/**
* This function checks if the user has pressed delete on a certain cell or not, returning false if they have.
*
* @param {String or Undefined} value : An inputed string or undefined
* @return {Boolean} Returns a boolean reporting whether the event object new value is not-undefined or not.
* @author Jarren Ralf
*/
function userHasNotPressedDelete(value)
{
  return value !== undefined;
}

/**
* This function checks if the user has pressed delete on a certain cell or not, returning true if they have.
*
* @param {String or Undefined} value : An inputed string or undefined
* @return {Boolean} Returns a boolean reporting whether the event object new value is undefined or not.
* @author Jarren Ralf
*/
function userHasPressedDelete(value)
{
  return value === undefined;
}

/**
 * This function checks if the user edits the item description or the Current Count column on the 
 * Manual Counts page. If they did, then a warning appears and reverses the changes that they made.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The active spreadsheet
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @param    {String}     sheetName  : The string that represents the name of the sheet
 * @author Jarren Ralf
 */
function warning(e, spreadsheet, sheet, sheetName)
{
  const range = e.range;
  const row = range.rowStart;
  const col = range.columnStart;

  if (row == range.rowEnd && col == range.columnEnd) // Single cell
  {
    if (col == 1)
    {
      if (!isRichmondSpreadsheet(spreadsheet))
      {
        (sheetName === "Manual Counts") ? // sheetName === 'TitesCounts'
          SpreadsheetApp.getUi().alert("Please don't attempt to change the items from the Manual Counts page.\n\nGo to the Item Search or Manual Scan page to add new products to this list.") :
          SpreadsheetApp.getUi().alert("Please don't attempt to change the items on the InfoCounts page.");

        range.setValue(e.oldValue); // Put the old value back in the cell
      }
    }
    else if (col == 2)
    {
      SpreadsheetApp.getUi().alert("Please don't change values in the Current Count column.\n\nType your updated inventory quantity in the New Count column.");
      range.setValue(e.oldValue); // Put the old value back in the cell
      if (userHasNotPressedDelete(e.value)) sheet.getRange(row, 3).setValue(e.value).activate(); // Move the count the user entered to the New Count column
    }
    else if (col == 3 && sheetName === "Manual Counts")
    {
      if (e.oldValue !== undefined) // Old value is NOT blank
      {
        if (userHasNotPressedDelete(e.value)) // New value is NOT blank
        {
          const valueSplit = e.value.toString().split(' ');

          if (isNumber(e.value))
          {
            if (isNumber(e.oldValue))
            {
              const difference  = e.value - e.oldValue;
              const newCountDataRange = sheet.getRange(row, 4, 1, 2);
              var runningSumValue = newCountDataRange.getValue().toString();

              if (runningSumValue === '')
                runningSumValue = Math.round(e.oldValue).toString();

              (difference > 0) ? 
                newCountDataRange.setValues([[runningSumValue.toString() + ' + ' + difference.toString(), new Date().getTime()]]) : 
                newCountDataRange.setValues([[runningSumValue.toString() + ' - ' + (-1*difference).toString(), new Date().getTime()]]);
            }
            else // Old value is not a number
            {
              const newCountDataRange = sheet.getRange(row, 4, 1, 2);
              var runningSumValue = newCountDataRange.getValue().toString();

              if (isNotBlank(runningSumValue))
                newCountDataRange.setValues([[runningSumValue + ' + ' + Math.round(e.value).toString(), new Date().getTime()]]);
              else
                newCountDataRange.setValues([[Math.round(e.value).toString(), new Date().getTime()]]);
            }
          }
          else if (valueSplit[0].toLowerCase() === 'a' || valueSplit[0].toLowerCase() === 'add') // The number is preceded by the letter 'a' and a space, in order to trigger an "add" operation
          {
            if (valueSplit.length === 3) // An add event with an inflow location
            { 
              const newCountDataRange = sheet.getRange(row, 3, 1, 5);
              var newCountValues = newCountDataRange.getValues()

              if (isNumber(valueSplit[1]))
              {
                newCountValues[0][0] = valueSplit[1]
                valueSplit[2] = valueSplit[2].toUpperCase()

                if (isNumber(newCountValues[0][0])) // New Count is a number
                {
                  if (isNumber(e.oldValue))
                  {
                    if (isNotBlank(newCountValues[0][1]))
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), newCountValues[0][1].toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), newCountValues[0][1].toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), newCountValues[0][1].toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[2], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), newCountValues[0][1].toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[2], parseInt(newCountValues[0][0]).toString()]]);
                    }
                    else
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), parseInt(e.oldValue).toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), parseInt(e.oldValue).toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), parseInt(e.oldValue).toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[2], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), parseInt(e.oldValue).toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[2], parseInt(newCountValues[0][0]).toString()]]);
                    }
                  }
                  else
                  {
                    if (isNotBlank(newCountValues[0][1]))
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[2], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[2], parseInt(newCountValues[0][0]).toString()]]);
                    }
                    else
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], newCountValues[0][4] + '\n' + newCountValues[0][0].toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], newCountValues[0][0].toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[2], newCountValues[0][4] + '\n' + newCountValues[0][0].toString()]]);
                      else
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[2], newCountValues[0][0].toString()]]);
                    }
                  }
                }
                else // New count is Not a number
                {
                  if (isNumber(e.oldValue))
                  {
                    if (isNotBlank(newCountValues[0][1])) // Running Sum is not blank
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], newCountValues[0][4] + '\n' + Math.round(e.oldValue).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], Math.round(e.oldValue).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), valueSplit[2], newCountValues[0][4] + '\n' + Math.round(e.oldValue).toString()]]);
                      else
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), valueSplit[2], Math.round(e.oldValue).toString()]]);
                    }
                    else
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], newCountValues[0][4] + '\n' + Math.round(e.oldValue).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[2], Math.round(e.oldValue).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), valueSplit[2], newCountValues[0][4] + '\n' + Math.round(e.oldValue).toString()]]);
                      else
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), valueSplit[2], Math.round(e.oldValue).toString()]]);
                    }
                  }

                  SpreadsheetApp.getUi().alert("The quantity you entered is not a number.")
                }
              }
              else if (isNumber(valueSplit[2]))
              {
                newCountValues[0][0] = valueSplit[2]
                valueSplit[1] = valueSplit[1].toUpperCase()

                if (isNumber(newCountValues[0][0])) // New Count is a number
                {
                  if (isNumber(e.oldValue))
                  {
                    if (isNotBlank(newCountValues[0][1]))
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), newCountValues[0][1].toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), newCountValues[0][1].toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), newCountValues[0][1].toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[1], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), newCountValues[0][1].toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[1], parseInt(newCountValues[0][0]).toString()]]);
                    }
                    else
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), parseInt(e.oldValue).toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), parseInt(e.oldValue).toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), parseInt(e.oldValue).toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[1], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else
                        newCountDataRange.setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), parseInt(e.oldValue).toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[1], parseInt(newCountValues[0][0]).toString()]]);
                    }
                  }
                  else
                  {
                    if (isNotBlank(newCountValues[0][1]))
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], parseInt(newCountValues[0][0]).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[1], newCountValues[0][4] + '\n' + parseInt(newCountValues[0][0]).toString()]]);
                      else
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[1], parseInt(newCountValues[0][0]).toString()]]);
                    }
                    else
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + newCountValues[0][0].toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][0].toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][0].toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[1], newCountValues[0][4] + '\n' + newCountValues[0][0].toString()]]);
                      else
                        newCountDataRange.setValues([[newCountValues[0][0], newCountValues[0][0].toString(), 
                          new Date().getTime(), valueSplit[1], newCountValues[0][0].toString()]]);
                    }
                  }
                }
                else // New count is Not a number
                {
                  if (isNumber(e.oldValue))
                  {
                    if (isNotBlank(newCountValues[0][1])) // Running Sum is not blank
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + Math.round(e.oldValue).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], Math.round(e.oldValue).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), valueSplit[1], newCountValues[0][4] + '\n' + Math.round(e.oldValue).toString()]]);
                      else
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), valueSplit[1], Math.round(e.oldValue).toString()]]);
                    }
                    else
                    {
                      if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + Math.round(e.oldValue).toString()]]);
                      else if (isNotBlank(newCountValues[0][3]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), newCountValues[0][3] + '\n' + valueSplit[1], Math.round(e.oldValue).toString()]]);
                      else if (isNotBlank(newCountValues[0][4]))
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), valueSplit[1], newCountValues[0][4] + '\n' + Math.round(e.oldValue).toString()]]);
                      else
                        newCountDataRange.setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), 
                          new Date().getTime(), valueSplit[1], Math.round(e.oldValue).toString()]]);
                    }
                  }

                  SpreadsheetApp.getUi().alert("The quantity you entered is not a number.")
                }
              }
              else
              {
                if (isNumber(e.oldValue))
                {
                  if (isNotBlank(newCountValues[0][1])) // Running Sum is not blank
                    newCountDataRange.setNumberFormat('@').setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), new Date().getTime(), 
                      newCountValues[0][3], newCountValues[0][4].toString()]])
                  else
                    newCountDataRange.setNumberFormat('@').setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), new Date().getTime(),
                      newCountValues[0][3], newCountValues[0][4].toString()]])
                }

                SpreadsheetApp.getUi().alert("The quantity you entered is not a number.")
              }
            }
            else if (valueSplit.length === 2) // Just an add event with NO inflow location assosiated to the inventory
            {
              const newCountDataRange = sheet.getRange(row, 3, 1, 3);
              var newCountValues = newCountDataRange.getValues()
              newCountValues[0][0] = valueSplit[1]

              if (isNumber(newCountValues[0][0])) // New Count is a number
              {
                if (isNumber(e.oldValue))
                {
                  if (isNotBlank(newCountValues[0][1]))
                    newCountDataRange.setNumberFormat('@').setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), 
                      newCountValues[0][1].toString() + ' + ' + newCountValues[0][0].toString(), new Date().getTime()]])
                  else
                    newCountDataRange.setNumberFormat('@').setValues([[parseInt(e.oldValue) + parseInt(newCountValues[0][0]), 
                      parseInt(e.oldValue).toString() + ' + ' + newCountValues[0][0].toString(), new Date().getTime()]])
                }
                else
                {
                  if (isNotBlank(newCountValues[0][1]))
                    newCountDataRange.setNumberFormat('@').setValues([[newCountValues[0][0], 
                      newCountValues[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountValues[0][0].toString(), new Date().getTime()]])
                  else
                    newCountDataRange.setNumberFormat('@').setValues([[newCountValues[0][0], newCountValues[0][0].toString(), new Date().getTime()]])
                }
              }
              else // New count is Not a number
              {
                if (isNumber(e.oldValue))
                {
                  if (isNotBlank(newCountValues[0][1])) // Running Sum is not blank
                    newCountDataRange.setNumberFormat('@').setValues([[Math.round(e.oldValue).toString(), newCountValues[0][1].toString() + ' + ' + NaN.toString(), new Date().getTime()]])
                  else
                    newCountDataRange.setNumberFormat('@').setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString(), new Date().getTime()]])
                }

                SpreadsheetApp.getUi().alert("The quantity you entered is not a number.")
              }
            }
          }
          else if (isNumber(valueSplit[0])) // The first split value is a number and the other is an inflow location
          {
            valueSplit[1] = valueSplit[1].toUpperCase()

            if (isNumber(e.oldValue))
            {
              const difference  = valueSplit[0] - e.oldValue;
              const newCountDataRange = sheet.getRange(row, 3, 1, 5);
              var newCountValues = newCountDataRange.getValues();

              if (newCountValues[0][1] === '')
                newCountValues[0][1] = Math.round(e.oldValue).toString();

              if (difference > 0)
              {
                if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1].toString() + ' + ' + difference.toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + difference.toString()]]);
                else if (isNotBlank(newCountValues[0][3]))
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1].toString() + ' + ' + difference.toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[1], difference.toString()]]);
                else if (isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1].toString() + ' + ' + difference.toString(), new Date().getTime(), 
                    valueSplit[1], newCountValues[0][4] + '\n' + difference.toString()]]);
                else
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1].toString() + ' + ' + difference.toString(), new Date().getTime(), 
                    valueSplit[1], difference.toString()]]);
              }
              else
              { 
                if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1].toString() + ' - ' + difference.toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + difference.toString()]]);
                else if (isNotBlank(newCountValues[0][3]))
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1].toString() + ' - ' + difference.toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[1], difference.toString()]]);
                else if (isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1].toString() + ' - ' + difference.toString(), new Date().getTime(), 
                    valueSplit[1], newCountValues[0][4] + '\n' + difference.toString()]]);
                else
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1].toString() + ' - ' + difference.toString(), new Date().getTime(), 
                    valueSplit[1], difference.toString()]]);
              }
            }
            else // Old value is not a number
            {
              const newCountDataRange = sheet.getRange(row, 3, 1, 5);
              var newCountValues = newCountDataRange.getValues()

              if (isNotBlank(newCountValues[0][1]))
              {
                if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1] + ' + ' + Math.round(valueSplit[0]).toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + valueSplit[0].toString()]]);
                else if (isNotBlank(newCountValues[0][3]))
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1] + ' + ' + Math.round(valueSplit[0]).toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[1], valueSplit[0].toString()]]);
                else if (isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1] + ' + ' + Math.round(valueSplit[0]).toString(), new Date().getTime(), 
                    valueSplit[1], newCountValues[0][4] + '\n' + valueSplit[0].toString()]]);
                else
                  newCountDataRange.setValues([[valueSplit[0], newCountValues[0][1] + ' + ' + Math.round(valueSplit[0]).toString(), new Date().getTime(), 
                    valueSplit[1], valueSplit[0].toString()]]);
              }
              else
              {
                if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[0], Math.round(valueSplit[0]).toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[1], newCountValues[0][4] + '\n' + valueSplit[0].toString()]]);
                else if (isNotBlank(newCountValues[0][3]))
                  newCountDataRange.setValues([[valueSplit[0], Math.round(valueSplit[0]).toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[1], valueSplit[0].toString()]]);
                else if (isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[0], Math.round(valueSplit[0]).toString(), new Date().getTime(), 
                    valueSplit[1], newCountValues[0][4] + '\n' + valueSplit[0].toString()]]);
                else
                  newCountDataRange.setValues([[valueSplit[0], Math.round(valueSplit[0]).toString(), new Date().getTime(), 
                    valueSplit[1], valueSplit[0].toString()]]);
              }
            }
          }
          else if (isNumber(valueSplit[1])) // The first split value is an inflow location and the other is a number
          {
            valueSplit[0] = valueSplit[0].toUpperCase()

            if (isNumber(e.oldValue))
            {
              const difference  = valueSplit[1] - e.oldValue;
              const newCountDataRange = sheet.getRange(row, 3, 1, 5);
              var newCountValues = newCountDataRange.getValues();

              if (newCountValues[0][1] === '')
                newCountValues[0][1] = Math.round(e.oldValue).toString();

              if (difference > 0)
              {
                if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1].toString() + ' + ' + difference.toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[0], newCountValues[0][4] + '\n' + difference.toString()]]);
                else if (isNotBlank(newCountValues[0][3]))
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1].toString() + ' + ' + difference.toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[0], difference.toString()]]);
                else if (isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1].toString() + ' + ' + difference.toString(), new Date().getTime(), 
                    valueSplit[0], newCountValues[0][4] + '\n' + difference.toString()]]);
                else
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1].toString() + ' + ' + difference.toString(), new Date().getTime(), 
                    valueSplit[0], difference.toString()]]);
              }
              else
              { 
                if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1].toString() + ' - ' + difference.toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[0], newCountValues[0][4] + '\n' + difference.toString()]]);
                else if (isNotBlank(newCountValues[0][3]))
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1].toString() + ' - ' + difference.toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[0], difference.toString()]]);
                else if (isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1].toString() + ' - ' + difference.toString(), new Date().getTime(), 
                    valueSplit[0], newCountValues[0][4] + '\n' + difference.toString()]]);
                else
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1].toString() + ' - ' + difference.toString(), new Date().getTime(), 
                    valueSplit[0], difference.toString()]]);
              }
            }
            else // Old value is not a number
            {
              const newCountDataRange = sheet.getRange(row, 3, 1, 5);
              var newCountValues = newCountDataRange.getValues()

              if (isNotBlank(newCountValues[0][1]))
              {
                if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1] + ' + ' + Math.round(valueSplit[1]).toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[0], newCountValues[0][4] + '\n' + valueSplit[1].toString()]]);
                else if (isNotBlank(newCountValues[0][3]))
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1] + ' + ' + Math.round(valueSplit[1]).toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[0], valueSplit[1].toString()]]);
                else if (isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1] + ' + ' + Math.round(valueSplit[1]).toString(), new Date().getTime(), 
                    valueSplit[0], newCountValues[0][4] + '\n' + valueSplit[1].toString()]]);
                else
                  newCountDataRange.setValues([[valueSplit[1], newCountValues[0][1] + ' + ' + Math.round(valueSplit[1]).toString(), new Date().getTime(), 
                    valueSplit[0], valueSplit[1].toString()]]);
              }
              else
              {
                if (isNotBlank(newCountValues[0][3]) && isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[1], Math.round(valueSplit[1]).toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[0], newCountValues[0][4] + '\n' + valueSplit[1].toString()]]);
                else if (isNotBlank(newCountValues[0][3]))
                  newCountDataRange.setValues([[valueSplit[1], Math.round(valueSplit[1]).toString(), new Date().getTime(), 
                    newCountValues[0][3] + '\n' + valueSplit[0], valueSplit[1].toString()]]);
                else if (isNotBlank(newCountValues[0][4]))
                  newCountDataRange.setValues([[valueSplit[1], Math.round(valueSplit[1]).toString(), new Date().getTime(), 
                    valueSplit[0], newCountValues[0][4] + '\n' + valueSplit[1].toString()]]);
                else
                  newCountDataRange.setValues([[valueSplit[1], Math.round(valueSplit[1]).toString(), new Date().getTime(), 
                    valueSplit[0], valueSplit[1].toString()]]);
              }
            }
          }
          else // New value is not a number
          {
            const runningSumRange = sheet.getRange(row, 4);
            const runningSumValue = runningSumRange.getValue().toString();

            if (isNumber(e.oldValue))
            {
              if (isNotBlank(runningSumValue))
                runningSumRange.setNumberFormat('@').setValue(runningSumValue + ' + ' + NaN.toString())
              else
                runningSumRange.setNumberFormat('@').setValue(Math.round(e.oldValue).toString())
            }

            SpreadsheetApp.getUi().alert("The quantity you entered is not a number.")
          }
        }
        else // New value IS blank
          sheet.getRange(row, 4, 1, 4).setValues([['', '', '', '']]); // Clear the running sum and last counted time
      }
      else
      {
        if (isNumber(e.value))
          sheet.getRange(row, 4, 1, 2).setNumberFormats([['@', '#']]).setValues([[e.value, new Date().getTime()]])
        else
        {
          const inflowData = e.value.split(' ');

          if (isNumber(inflowData[0]))
            sheet.getRange(row, 3, 1, 5).setNumberFormats([['#', '@', '#', '@', '#']]).setValues([[inflowData[0], inflowData[0], new Date().getTime(), inflowData[1].toUpperCase(), inflowData[0]]])
          else if (isNumber(inflowData[1]))
            sheet.getRange(row, 3, 1, 5).setNumberFormats([['#', '@', '#', '@', '#']]).setValues([[inflowData[1], inflowData[1], new Date().getTime(), inflowData[0].toUpperCase(), inflowData[1]]])
          else
            SpreadsheetApp.getUi().alert("The quantity you entered is not a number.");
        }
      }
    }
  }
}