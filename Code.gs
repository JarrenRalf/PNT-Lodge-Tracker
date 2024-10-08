/**
 * This function handles the import of an excel file by identifying the creation of a new sheet.
 * 
 * @param {Event Object} e : The event object.
 */
function onChange(e)
{
  if (e.changeType === 'INSERT_GRID')
  {
    var spreadsheet = e.source;
    var sheets = spreadsheet.getSheets();
    var info, numRows = 0, numCols = 1, maxRow = 2, maxCol = 3, isAdagioOE = 4, isBackOrderItems = 5;

    for (var sheet = 0; sheet < sheets.length; sheet++) // Loop through all of the sheets in this spreadsheet and find the new one
    {
      info = [
        sheets[sheet].getLastRow(),
        sheets[sheet].getLastColumn(),
        sheets[sheet].getMaxRows(),
        sheets[sheet].getMaxColumns(),
        sheets[sheet].getSheetValues(1, 1, 1, sheets[sheet].getLastColumn()).flat().includes('Created by User'),
        sheets[sheet].getSheetValues(1, 1, 1, sheets[sheet].getLastColumn()).flat().includes('Qty Original Ordered')
      ]

      // A new sheet is imported by File -> Import -> Insert new sheet(s) - The left disjunct is for a csv and the right disjunct is for an excel file
      if ((info[maxRow] - info[numRows] === 2 && info[maxCol] - info[numCols] === 2) || (info[maxRow] === 1000 && info[maxCol] === 26 && info[numRows] !== 0 && info[numCols] !== 0) || 
          ((info[maxRow] === info[numRows] && info[maxCol] === info[numCols]) && (info[isAdagioOE] || info[isBackOrderItems]))) 
      {
        spreadsheet.toast('Processing imported data...', '', 60)
        
        const values = sheets[sheet].getSheetValues(1, 1, info[numRows], info[numCols]); // This is the shopify order data
        const fileName = sheets[sheet].getSheetName();

        if (fileName.substring(0, 7) !== "Copy Of") // Don't delete the sheets that are duplicates
          spreadsheet.deleteSheet(sheets[sheet]) // Delete the new sheet that was created

        if (info[isAdagioOE])
          updateOrdersOnTracker(values, spreadsheet);
        else if (info[isBackOrderItems])
          updateItemsOnTracker(values, spreadsheet, fileName);

        break;
      }
    }
  }
}

/**
 * This function checks to see if a user is moving a row from one sheet to another.
 * 
 * @param {Event Object} e : The event object.
 */
function onEdit(e)
{
  try
  {
    moveRow(e)
  }
  catch (error)
  {
    Browser.msgBox(error)
  }
}

/**
 * This function allows the user to add items from the P/O or B/O page to the relevant transfer page.
 * 
 * @author Jarren Ralf
 */
function addItemsToTransferSheet()
{
  const activeSheet = SpreadsheetApp.getActiveSheet();
  var activeRanges = activeSheet.getActiveRangeList().getRanges(); // The selected ranges on the item search sheet
  var firstRows = [], lastRows = [], numRows = [], values, sku = [], qty = [], name = [], ordNum = [];
  
  // Find the first row and last row in the the set of all active ranges
  for (var r = 0; r < activeRanges.length; r++)
  {
    firstRows.push(activeRanges[r].getRow());
     lastRows.push(activeRanges[r].getLastRow())
      numRows.push(lastRows[r] - firstRows[r] + 1);
      values = activeSheet.getSheetValues(firstRows[r], 3, numRows[r], 9)
       sku.push(...values.map(v => v[3]))
       qty.push(...values.map(v => v[2]))
      name.push(...values.map(v => v[0]))
    ordNum.push(...values.map(v => v[8]))
  }

  var firstRow = Math.min(...firstRows); // This is the smallest starting row number out of all active ranges
  var skus = [].concat.apply([], sku); // Concatenate all of the item values as a 2-D array
  var numItems = skus.length;

  if (firstRow < 3)
    Browser.msgBox('Please select items from the list.')
  else if (numItems === 0)
    Browser.msgBox('Please select items from the list.')
  else
  {
    const spreadsheet = SpreadsheetApp.getActive()
    const sheetName = activeSheet.getSheetName();
    const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
    const itemNum = csvData[0].indexOf('Item #')
    const items = csvData.filter(item => skus.includes(item[itemNum]));
    const today = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'dd MMM yyyy');
    var row = 0, numRows = 0, sheet, itemValues, fromLocation, toLocation, url;

    if (qty.length !== items.length)
      Browser.msgBox('Contact AJ and tell him what SKUs you are trying to put on the transfer sheet. Let him know that ' + (qty.length - items.length) + ' of those SKUs can\'t be found in the inventory.csv. **Please note that if you see \"Comment Line\" in the SKU column that you can\'t select that line while performing this action.')
    else
    {
      var ui = SpreadsheetApp.getUi();

      var response = ui.prompt('Which PNT location are you shipping FROM?', 'Please type: \"rich", \"parks\", or \"pr\".', ui.ButtonSet.OK_CANCEL);

      // Process the user's response.
      if (response.getSelectedButton() == ui.Button.OK)
      {
        var textResponse = response.getResponseText().toUpperCase();

        if (textResponse == 'RICH')
        {
          fromLocation = 'Richmond';

          response = ui.prompt('Which PNT location are you shipping TO?', 'Please type: \"rich", \"parks\", or \"pr\".', ui.ButtonSet.OK_CANCEL);

          // Process the user's response.
          if (response.getSelectedButton() == ui.Button.OK)
          {
            textResponse = response.getResponseText().toUpperCase();

            if (textResponse == 'PARKS')
              toLocation = 'Parskville'
            else if (textResponse == 'PR')
              toLocation = 'Rupert'
            else
              ui.alert('Your typed response did not exactly match any of the location choices. Please Try again.')
          }
          else // The user has clicked on CLOSE or CANCEL
            return;
        }
        else if (textResponse == 'PARKS')
        {
          toLocation = 'Richmond';
          fromLocation = 'Parksville';
        }
        else if (textResponse == 'PR')
        {
          toLocation = 'Richmond';
          fromLocation = 'Rupert';
        }
        else
          ui.alert('Your typed response did not exactly match any of the location choices. Please Try again.')
      }
      else // The user has clicked on CLOSE or CANCEL
        return;

      switch (fromLocation)
      {
        case 'Richmond':

          switch (toLocation)
          {
            case 'Parskville':
              url = 'https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit?gid=1340095049#gid=1340095049'
              sheet = SpreadsheetApp.openByUrl(url).getSheetByName('Order')
              itemValues = items.map((v,idx) => [today, 'Lodge\nTracker', qty[idx], v[0], v[1], 'ATTN: Eryn (Lodge Items)\n' + name[idx] + '\nORD# ' + ordNum[idx], v[3], '']) 
              row = sheet.getLastRow() + 1;
              numRows = itemValues.length;
              sheet.getRange(row, 1, numRows, 8).setNumberFormat('@').setValues(itemValues)
              applyFullRowFormatting(sheet, row, numRows, false)
              break;
            case 'Rupert':
              url = 'https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit?gid=407280159#gid=407280159'
              sheet = SpreadsheetApp.openByUrl(url).getSheetByName('Order')
              itemValues = items.map((v,idx) => [today, 'Lodge\nTracker', qty[idx], v[0], v[1], 'ATTN: Doug (Lodge Items)\n' + name[idx] + '\nORD# ' + ordNum[idx], v[4], '']) 
              row = sheet.getLastRow() + 1;
              numRows = itemValues.length;
              sheet.getRange(row, 1, numRows, 8).setNumberFormat('@').setValues(itemValues)
              applyFullRowFormatting(sheet, row, numRows, false)
              break;
          }
          break;
        case 'Parksville':
          url = 'https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit?gid=269292771#gid=269292771'
          sheet = SpreadsheetApp.openByUrl(url).getSheetByName('ItemsToRichmond')
          itemValues = items.map((v,idx) => [today, 'Lodge\nTracker', v[0], v[1], 'ATTN: Scott (Lodge Items)\n' + name[idx] + '\nORD# ' + ordNum[idx], qty[idx]]) 
          row = sheet.getLastRow() + 1;
          numRows = itemValues.length;
          sheet.getRange(row, 1, numRows, 6).setNumberFormat('@').setValues(itemValues)
          applyFullRowFormatting(sheet, row, numRows, true)
          break;
        case 'Rupert':
          url = 'https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit?gid=1569594370#gid=1569594370'
          sheet = SpreadsheetApp.openByUrl(url).getSheetByName('ItemsToRichmond')
          itemValues = items.map((v,idx) => [today, 'Lodge\nTracker', v[0], v[1], 'ATTN: Scott (Lodge Items)\n' + name[idx] + '\nORD# ' + ordNum[idx], qty[idx]]) 
          row = sheet.getLastRow() + 1;
          numRows = itemValues.length;
          sheet.getRange(row, 1, numRows, 6).setNumberFormat('@').setValues(itemValues)
          applyFullRowFormatting(sheet, row, numRows, true)
          break;
      }

      if (sheetName == 'B/O' && fromLocation != undefined && toLocation != undefined)
      {
        activeSheet.getFilter().remove(); // Remove the filter
        activeSheet.getRange(2, 1, 1, activeSheet.getLastColumn()).createFilter(); // Create a filter in the header
        SpreadsheetApp.flush();

        const linkToTransferSheet = SpreadsheetApp.newRichTextValue().setText('Shipping from ' + fromLocation + ' to ' + toLocation).setLinkUrl(url + '&range=B' + row).build()
        activeRanges.map(rng => rng.offset(0, 12 - rng.getColumn(), rng.getNumRows(), 1).setRichTextValues(new Array(rng.getNumRows()).fill([linkToTransferSheet])));
      }

      if (sheetName == 'I/O' && fromLocation != undefined && toLocation != undefined)
      {
        activeSheet.getFilter().remove(); // Remove the filter
        activeSheet.getRange(2, 1, 1, activeSheet.getLastColumn()).createFilter(); // Create a filter in the header
        SpreadsheetApp.flush();

        const linkToTransferSheet = SpreadsheetApp.newRichTextValue().setText('Shipping from ' + fromLocation + ' to ' + toLocation).setLinkUrl(url + '&range=B' + row).build()
        activeRanges.map(rng => rng.offset(0, 12 - rng.getColumn(), rng.getNumRows(), 1).setRichTextValues(new Array(rng.getNumRows()).fill([linkToTransferSheet])));
      }
    }
  }
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
function applyFullRowFormatting(sheet, row, numRows, isItemsToRichmondPage)
{
  const BLUE = '#c9daf8', GREEN = '#d9ead3', YELLOW = '#fff2cc', GREEN_DATE = '#b6d7a8';

  if (isItemsToRichmondPage)
  {
    var      borderRng = sheet.getRange(row, 1, numRows, 8);
    var  shippedColRng = sheet.getRange(row, 6, numRows   );
    var thickBorderRng = sheet.getRange(row, 6, numRows, 3);
    var backgroundColours = [...Array(numRows)].map(_ => [GREEN_DATE, 'white', 'white', 'white', 'white', 'white', 'white', 'white'])
    var numberFormats = [...Array(numRows)].map(_ => ['dd MMM yyyy', '@', '@', '@', '@', '#.#', '@', '@']);
    var horizontalAlignments = [...Array(numRows)].map(_ => ['right', 'center', 'center', 'left', 'center', 'center', 'center', 'left']);
    var wrapStrategies = [...Array(numRows)].map(_ => [...new Array(2).fill(SpreadsheetApp.WrapStrategy.OVERFLOW), ...new Array(3).fill(SpreadsheetApp.WrapStrategy.WRAP), 
        SpreadsheetApp.WrapStrategy.CLIP, SpreadsheetApp.WrapStrategy.WRAP, SpreadsheetApp.WrapStrategy.WRAP]);
  }
  else
  {
    var         borderRng = sheet.getRange(row, 1, numRows, 11);
    var     shippedColRng = sheet.getRange(row, 9, numRows    );
    var    thickBorderRng = sheet.getRange(row, 9, numRows,  2);
    var backgroundColours = [...Array(numRows)].map(_ => [GREEN_DATE, 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white'])
    var numberFormats = [...Array(numRows)].map(e => ['dd MMM yyyy', '@', '#.#', '@', '@', '@', '#.#', '0.#', '#.#', '@', 'dd MMM yyyy']);
    var horizontalAlignments = [...Array(numRows)].map(e => ['right', ...new Array(3).fill('center'), 'left', ...new Array(6).fill('center')]);
    var wrapStrategies = [...Array(numRows)].map(e => [...new Array(3).fill(SpreadsheetApp.WrapStrategy.OVERFLOW), ...new Array(3).fill(SpreadsheetApp.WrapStrategy.WRAP),
      ...new Array(3).fill   (SpreadsheetApp.WrapStrategy.CLIP), SpreadsheetApp.WrapStrategy.WRAP, SpreadsheetApp.WrapStrategy.CLIP]);
  }

  borderRng.setFontSize(10).setFontLine('none').setFontWeight('bold').setFontStyle('normal').setFontFamily('Arial').setFontColor('black')
    .setNumberFormats(numberFormats).setHorizontalAlignments(horizontalAlignments).setWrapStrategies(wrapStrategies)
    .setBorder(true, true, true, true,  null, true, 'black', SpreadsheetApp.BorderStyle.SOLID).setBackgrounds(backgroundColours);

  thickBorderRng.setBorder(null, true, null, true, false, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK).setBackground(GREEN);
  shippedColRng.setBackground(YELLOW);

  if (!isItemsToRichmondPage)
    sheet.getRange(row, 7, numRows, 2).setBorder(null,  true,  null,  null,  true,  null, 'black', SpreadsheetApp.BorderStyle.SOLID).setBackground(BLUE);
}

/**
 * This function creates the onChange trigger that handles all of the imported files.
 * 
 * @author Jarren Ralf
 */
function createTriggers()
{
  ScriptApp.newTrigger('onChange').forSpreadsheet(SpreadsheetApp.getActive()).onChange().create();
}

/**
 * This function finds items with on the B/O tab that matches the given order number and deletes them.
 * 
 * @param {String || String[][]} orderNumber : The order number of the current order being updated on the ORDERS page.
 * @param     {Spreadsheet}      spreadsheet : The active spreadsheet
 * @author Jarren Ralf
 */
function deleteBackOrderedItems(orderNumber, spreadsheet)
{
  const boSheet = spreadsheet.getSheetByName('B/O');
  const ioSheet = spreadsheet.getSheetByName('I/O');
  boSheet.getFilter().remove(); // Remove the filter
  ioSheet.getFilter().remove();
        

  Logger.log('Array.isArray(orderNumber): ' + Array.isArray(orderNumber))

  if (Array.isArray(orderNumber)) // When importing new orders the argument passed to this function is an array with multiple order numbers
  {
    var orderNumbers, row, numRows;

    orderNumber.map(ordNum => {

      if (!isBlank(ordNum[2]))
      {
        orderNumbers = boSheet.getSheetValues(3, 11, boSheet.getLastRow() - 2, 1);
        row = orderNumbers.findIndex(ordNum => ordNum[0] == ordNum[2]);

        if (row !== -1)
        {
          numRows = orderNumbers.findLastIndex(ordNum => ordNum[0] == ordNum[2]) - row + 1;
          boSheet.deleteRows(row + 3, numRows);
          SpreadsheetApp.flush();
        }

        // Inital Orders Sheet
        orderNumbers = ioSheet.getSheetValues(3, 11, ioSheet.getLastRow() - 2, 1);
        row = orderNumbers.findIndex(ordNum => ordNum[0] == ordNum[2]);

        if (row !== -1)
        {
          numRows = orderNumbers.findLastIndex(ordNum => ordNum[0] == ordNum[2]) - row + 1;
          ioSheet.deleteRows(row + 3, numRows);
          SpreadsheetApp.flush();
        }
      }
    })
  }
  else
  {
    if (!isBlank(orderNumber)) // Order number is not blank on the Orders page
    {
      const orderNumbers_BO = boSheet.getSheetValues(3, 11, boSheet.getLastRow() - 2, 1);
      const orderNumbers_IO = ioSheet.getSheetValues(3, 11, ioSheet.getLastRow() - 2, 1);
      const row_BO = orderNumbers_BO.findIndex(ordNum => ordNum[0] == orderNumber);
      const row_IO = orderNumbers_IO.findIndex(ordNum => ordNum[0] == orderNumber);

      if (row_BO !== -1)
      {
        const numRows_BO = orderNumbers_BO.findLastIndex(ordNum => ordNum[0] == orderNumber) - row_BO + 1;
        boSheet.deleteRows(row_BO + 3, numRows_BO);
      }

      if (row_IO !== -1)
      {
        const numRows_IO = orderNumbers_IO.findLastIndex(ordNum => ordNum[0] == orderNumber) - row_IO + 1;
        ioSheet.deleteRows(row_IO + 3, numRows_IO);
      }
    }
  }  

  boSheet.getRange(2, 1, 1, boSheet.getLastColumn()).createFilter(); // Create a filter in the header
  ioSheet.getRange(2, 1, 1, ioSheet.getLastColumn()).createFilter();
}

/**
 * This function checks whether the given order number contains a back ordered items or not.
 * 
 * @param {String} order : The order number of the given order.
 * @param {String[]} backOrderNumbers : The list of order numbers that contains back orders.
 * @returns {Boolean} Returns true if there are back ordered items on the given order, or false if it is an initial order.
 * @author Jarren Ralf
 */
function doesOrderContainBOs(order, backOrderNumbers)
{
  return backOrderNumbers.includes(order)
}

/**
 * This function converts Yes or No response from Order Entry regarding the status of the order into Back Order or not.
 * 
 * @param {String} isOrderComplete : Yes or No depending on whether the order is complete.
 * @return {String} Returns what the back order status is.
 * @author Jarren Ralf
 */
function getBoStatus(isOrderComplete)
{
  return (isOrderComplete !== 'No') ? '' : 'BO';
}

/**
 * This function converts the location code from Order Entry to the name of the city.
 * 
 * @param {String} locationCode : The location code where the inventory is coming from.
 * @param {Object} months : An object containing all of the names of each month.
 * @return {String} Returns the name of the city.
 * @author Jarren Ralf
 */
function getDateString(date, months)
{
  const d_split = date.toString().split('-');
  
  return months[d_split[0]] + ' ' + d_split[1] + ', ' + d_split[2];
}

/**
 * This function takes the paired list of order numbers and the employee name that entered that order in adagio, as well as an order number to try and find
 * within that list, it searches and returns the name of the employee, or blank if not found.
 * 
 * @param {String[]} orderNumber : The given order number that is being searched for.
 * @param {String}     orders    : A list of the current orders and who entered them into Adagio.
 * @return {String} Returns the name of the person who entered the order in Adagio or blank.
 * @author Jarren Ralf
 */
function getEnteredByNameAndApprovalStatus(orderNumber, orders)
{
  const enteredBy = orders.find(ordNum => ordNum[1] == orderNumber)
  
  return (enteredBy != null) ? [enteredBy[0], enteredBy[2]] : '';
}

/**
 * This function converts the initials from Order Entry to the full first name of the employee.
 * 
 * @param {String} initials : The intials of the employee.
 * @return {String} Returns the full first name of the employee.
 * @author Jarren Ralf
 */
function getFullName(initials)
{
  switch (initials.trim())
  {
    case '':
      return '';
    case 'AJ':
      return 'Adrian';
    case 'BK':
      return 'Brent';
    case 'EG':
      return 'Eryn';
    case 'FN':
      return 'Frank';
    case 'GN':
      return 'Gary';
    case 'KN':
      return 'Kris';
    case 'KT':
      return 'Karen';
    case 'MW':
      return 'Mark';
    case 'SN':
      return 'Scott';
    case 'TW':
      return 'Jarren';
    default:
      return initials;
  }
}

/**
 * This function converts Yes or No response from Order Entry regarding the status of the order into Back Order or not.
 * 
 * @param {String} isOrderComplete : Yes or No depending on whether the order is complete.
 * @return {String} Returns what the back order status is.
 * @author Jarren Ralf
 */
function getInvoiceNumber(invNum, isCompletedOrders)
{
  return (invNum === ' ') ? '' : (isCompletedOrders) ? invNum : 'multiple';
}

/**
 * This function converts the location code from Order Entry to the name of the city.
 * 
 * @param {String} locationCode : The location code where the inventory is coming from.
 * @return {String} Returns the name of the city.
 * @author Jarren Ralf
 */
function getLocationName(locationCode)
{
  switch (locationCode)
  {
    case '200':
      return 'Parksville';
    case '300':
      return 'Rupert';
    case '100':
    case '400':
      return 'Trites';
  }
}

/**
 * This function checks if the tab of the imported excel sheet contains the Adagio Order Number, if not, it prompts the user to enter it. 
 * If there are any uynexpected inputs, the order number is left blank.
 * 
 * @param {String} ordNum : The tab name of the imported excel spreadsheet (assumed to be the order number)
 * @returns {String} Returns the order number if it has been determined to be correct, or blank otherwise.
 * @auther Jarren Ralf
 */
function getOrderNumber(ordNum)
{
  if (isNumber(ordNum) && ordNum.toString().length === 5)
    return ordNum;
  else
  {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('Enter the order number:',);
    const orderNumber = response.getResponseText().trim(); 

    return (response.getSelectedButton() !== ui.Button.OK) ? '' : (isNumber(orderNumber) && orderNumber.length === 5) ? orderNumber : '';
  }
}

/**
 * This function converts Yes or No response from Order Entry regarding the status of the order into Complete or Partial.
 * 
 * @param {String} isOrderComplete : Yes or No depending on whether the order is complete.
 * @return {String} Returns what the order status is.
 * @author Jarren Ralf
 */
function getOrderStatus(isOrderComplete, isCompletedOrders, invNum)
{
  return (isCompletedOrders) ? (isOrderComplete !== 'No') ? 'Completed' : 'Partial' : (invNum !== ' ') ? 'Partial Order' : '';
}

/**
 * This function takes the Adagio name of the Lodge or Guide, finds it in the given list and returns the proper typeset version of the name.
 * 
 * @param {String} name : The name of the lodge from Adagio
 * @param {String[][]} listOfNames : The list of names from Adagio with their typeset counterparts.
 * @param {Number}   colSelector   : The index of the columns to select. 1: For Imported Orders 2: For B/O items 
 * @return {String} Returns the proper typest name of the lodge or charter, if found.
 * @author Jarren Ralf
 */
function getProperTypesetName(name, listOfNames, colSelector)
{
  const properTypesetName = listOfNames.find(customer => customer[0] === name)
  
  return (properTypesetName != null) ? properTypesetName[colSelector] : name;
}

/**
 * This function checks if the given string is blank.
 * 
 * @param {String} str : The given string.
 * @returns {Boolean} Whether the given string is blank or not.
 * @author Jarren Ralf
 */
function isBlank(str)
{
  return str === '';
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
 * This function moves the selected row from the Lodge or Guide order page to the completed page.
 * 
 * @param {Event Object} e : The event object.
 */
function moveRow(e)
{
  const range = e.range;
  const row = range.rowStart;
  const col = range.columnStart;  

  if (row == range.rowEnd && col == range.columnEnd) // Only look at a single cell edit
  {
    const spreadsheet = e.source;
    const sheet = spreadsheet.getActiveSheet();
    const sheetNames = sheet.getSheetName().split(" ") // Split the sheet name, which will be used to distinguish between Logde and Guide page

    if (sheetNames[1] == "ORDERS") // An edit is happening on one of the Order pages
    {
      const numCols = sheet.getLastColumn()

      if (col == numCols) // Order Status is changing
      {
        const value = e.value; 
        const numCols = sheet.getLastColumn()

        if (value == 'Updated')
        {
          range.setValue('')
          sheet.getRange(row, 5).setValue('').offset(0, -4, 1, numCols).setBackground('#00ff00')
        }
        else if (value == 'Picking')
        {
          const ui = SpreadsheetApp.getUi()
          ui.alert('Order NOT Approved', 'You have started picking an order that may not be approved by the customer yet.\n\nYou may want to check with ' + 
            sheet.getRange(row, 2).getValue() + ' before picking any items.', ui.ButtonSet.OK)
        }
        else
        {
          const rowValues = sheet.getSheetValues(row, 1, 1, numCols)[0]; // Entire row values
          const timeZone = spreadsheet.getSpreadsheetTimeZone(); // Set the timezone

          rowValues[0] = Utilities.formatDate(rowValues[0], timeZone, 'MMM dd, yyyy'); // Set the format of the order date
          rowValues.push(Utilities.formatDate(     new Date(), timeZone, 'MMM dd, yyyy')); // Set the current time for the completion date

          if (value == "Completed") // The order status is being set to complete 
          {
            rowValues[4] = ''; // Clear the Back Order column
            spreadsheet.getSheetByName(sheetNames[0] +  " COMPLETED").appendRow(rowValues) // Move the row of values to the completed page
            sheet.deleteRow(row); // Delete the row from the order page
            deleteBackOrderedItems(rowValues[2], spreadsheet);
          }
          else if (value == "Cancelled") // The order status is being set to cancelled 
          { 
            spreadsheet.getSheetByName("CANCELLED").appendRow(rowValues) // Move the row of values to the cancelled page
            sheet.deleteRow(row); // Delete the row from the order page
            deleteBackOrderedItems(rowValues[2], spreadsheet);
          }
          else if (value == "Partial") // The order status is being set to partial
          {
            rowValues[4] = 'BO'; // Set the value in the back order column to 'BO'
            spreadsheet.getSheetByName(sheetNames[0] +  " COMPLETED").appendRow(rowValues); // Move the row of values to the completed page
            sheet.getRange(row, 12, 1, 4).setValues([['multiple', '', '',  'Partial Order']]); // Clear the invoice values, and set the status
            deleteBackOrderedItems(rowValues[2], spreadsheet);
          }
        }
      }
      else if (col == 5) // Adding a Printed By name
      {
        if (range.getValue() !== '')
        {
          if (!range.offset(0, -1).isChecked())
          {
            const ui = SpreadsheetApp.getUi()
            ui.alert('Order NOT Approved', 'You have printed an order that may not be approved by the customer yet.\n\nYou may want to check with ' + 
              range.offset(0, -3).getValue() + ' before picking any items.', ui.ButtonSet.OK)
          }
           
          sheet.getRange(row, 1, 1, numCols).setBackground((sheet.getRange(row, 2).getBackground() === 'white') ? 'white' : '#e8f0fe');
        }
      }
    }
  }
}

/**
 * This function removes the dashes from the SKU number.
 * 
 * @param {String} sku : The sku number with 2 dashes in it.
 * @retruns {String} Returns the sku number without dahses in it.
 * @author Jarren Ralf
 */
function removeDashesFromSku(sku)
{
  return sku.toString().substring(0, 4) + sku.toString().substring(5, 9) + sku.toString().substring(10);
}

/**
 * This function handles the import of an order entry order that may contain back ordered items.
 * 
 * @param {String[][]}     items    : A list of items on the order that was imported.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @param {String}        ordNum    : The order number that is being imported.
 * @author Jarren Ralf
 */
function updateItemsOnTracker(items, spreadsheet, ordNum)
{
  items.pop(); // Remove the "Total" or final line

  // Get all the indexes of the relevant headers
  const headerOE = items.shift();
  const dateIdx = headerOE.indexOf('Date');
  const customerNumIdx = headerOE.indexOf('Cust #');
  const originalOrderedQtyIdx = headerOE.indexOf('Qty Original Ordered');
  const orderedQtyIdx = headerOE.indexOf('Qty Ordered'); 
  const backOrderQtyIdx = headerOE.indexOf('Backorder'); 
  const skuIdx = headerOE.indexOf('Item');
  const descriptionIdx = headerOE.indexOf('Description');
  const unitPriceIdx = headerOE.indexOf('Unit Price');
  const locationIdx = headerOE.indexOf('Loc');
  const isItemCompleteIdx = headerOE.indexOf('Complete?');
  const months = {'01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun', '07': 'Jul', '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'};
  const orderNumber = getOrderNumber(ordNum);

  const lodgeCustomerSheet = spreadsheet.getSheetByName('Lodge Customer List');
  const charterGuideCustomerSheet = spreadsheet.getSheetByName('Charter & Guide Customer List');
  const lodgeOrdersSheet = spreadsheet.getSheetByName('LODGE ORDERS');
  const partialOrdersSheet = spreadsheet.getSheetByName('Partial Orders');
  const enteredByNamesAndApprovalStatus = lodgeOrdersSheet.getSheetValues(3, 2, lodgeOrdersSheet.getLastRow() - 2, 3);
  const customerNames = lodgeCustomerSheet.getSheetValues(3, 1, lodgeCustomerSheet.getLastRow() - 2, 3).concat(charterGuideCustomerSheet.getSheetValues(3, 1, charterGuideCustomerSheet.getLastRow() - 2, 3))
  const orderNumbers_BO = partialOrdersSheet.getSheetValues(2, 1, partialOrdersSheet.getRange(partialOrdersSheet.getLastRow(), 1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow() - 1, 1).flat()

  const orderDate = getDateString(items[0][dateIdx], months);
  const enteredByAndApproval = getEnteredByNameAndApprovalStatus(orderNumber, enteredByNamesAndApprovalStatus);
  const customerName = getProperTypesetName(items[0][customerNumIdx], customerNames, 2);
  const locationName = getLocationName(items[0][locationIdx])

  if (doesOrderContainBOs(orderNumber, orderNumbers_BO))
  {
    var newItems = items.filter(item => item[isItemCompleteIdx] ).filter(item => item[backOrderQtyIdx]).map(item => {
        return [orderDate, enteredByAndApproval[0], customerName, item[originalOrderedQtyIdx], item[backOrderQtyIdx], 
          removeDashesFromSku(item[skuIdx]), item[descriptionIdx], item[unitPriceIdx], Number(item[backOrderQtyIdx])*Number(item[unitPriceIdx]), locationName , orderNumber, '', '', ''] // Back Ordered Items
    });

    var itemSheet = spreadsheet.getSheetByName('B/O'); 
  }
  else 
  {
    var newItems = items.map(item => {
        return [orderDate, enteredByAndApproval[0], customerName, enteredByAndApproval[1], item[orderedQtyIdx], 
          removeDashesFromSku(item[skuIdx]), item[descriptionIdx], item[unitPriceIdx], Number(item[orderedQtyIdx])*Number(item[unitPriceIdx]), locationName , orderNumber, '', '', ''] // Back Ordered Items
    });

    var itemSheet = spreadsheet.getSheetByName('I/O'); 
  }

  const numRows = itemSheet.getLastRow() - 2;
  const numNewItems = newItems.length;
  var numItemsRemoved = numNewItems;
  itemSheet.getFilter().remove(); // Remove the filter

  if (numRows > 0)
  {
    
    const ordNum = itemSheet.getSheetValues(2, 1, 1, 14).flat().indexOf('Order #');
    var currentItems = itemSheet.getSheetValues(3, 1, numRows, itemSheet.getLastColumn()).filter(item => isBlank(item[ordNum]) || item[ordNum] !== orderNumber);
    var numCurrentItems = currentItems.length;
    itemSheet.getRange(3, 1, numCurrentItems, currentItems[0].length).setValues(currentItems);

    if (numRows > numCurrentItems)
    {
      numItemsRemoved = numRows - numCurrentItems;
      itemSheet.deleteRows(numCurrentItems + 3, numItemsRemoved);
    }
  }

  Logger.log('Order Number: ' + orderNumber)

  if (numNewItems > 0)
  {
    const numCols = newItems[0].length;

    if (doesOrderContainBOs(orderNumber, orderNumbers_BO))
    {
      if (numRows > 0)
        itemSheet.getRange(numCurrentItems + 3, 1, numNewItems, numCols)
            .setNumberFormats(new Array(numNewItems).fill(['MMM dd, yyyy', '@', '@','#', '#', '@', '@', '$#,##0.00', '$#,##0.00', '@', '@', '@', '@', '@'])).setValues(newItems)
          .offset(-1*numCurrentItems, 0, numCurrentItems + numNewItems, numCols).sort([{column: 1, ascending: true}]);
      else
        itemSheet.getRange(3, 1, numNewItems, numCols).setNumberFormats(new Array(numNewItems).fill(['MMM dd, yyyy', '@', '@', '#', '#', '@', '@', '$#,##0.00', '$#,##0.00', '@', '@', '@', '@', '@']))
          .setValues(newItems)

      Logger.log('The following new Back Ordered items were added to the B/O tab:')
      Logger.log(newItems)

      spreadsheet.toast(numNewItems + ' Added ' + (numItemsRemoved - numNewItems) + ' Removed', 'B/O Items Imported', 60)
    }
    else
    {
      if (numRows > 0)
        itemSheet.getRange(numCurrentItems + 3, 1, numNewItems, numCols)
            .setNumberFormats(new Array(numNewItems).fill(['MMM dd, yyyy', '@', '@','#', '#', '@', '@', '$#,##0.00', '$#,##0.00', '@', '@', '@', '@', '@'])).setValues(newItems)
          .offset(-1*numCurrentItems, 0, numCurrentItems + numNewItems, numCols).sort([{column: 1, ascending: true}]);
      else
        itemSheet.getRange(3, 1, numNewItems, numCols).setNumberFormats(new Array(numNewItems).fill(['MMM dd, yyyy', '@', '@', '#', '#', '@', '@', '$#,##0.00', '$#,##0.00', '@', '@', '@', '@', '@']))
          .setValues(newItems)

      Logger.log('The following new Ordered items were added to the I/O tab:')
      Logger.log(newItems)

      spreadsheet.toast(numNewItems + ' Added ' + (numItemsRemoved - numNewItems) + ' Removed', 'I/O Items Imported', 60)
    }
  }
  else
    spreadsheet.toast('ORD# ' + orderNumber + ' may be in the process of being shipped.', '**NO B/O or I/O Items Imported**', 60)

  itemSheet.getRange(2, 1, 1, itemSheet.getLastColumn()).createFilter(); // Create a filter in the header
}

/**
 * This function handles the import of the list of current and completed orders into the spreadsheet.
 * 
 * @param {String[][]} allOrders : All of the current orders from Adagio.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function updateOrdersOnTracker(allOrders, spreadsheet)
{
  allOrders.pop(); // Remove the "Total" or final line

  // Get all the indexes of the relevant headers
  const headerOE = allOrders.shift();
  var totalIdx = headerOE.indexOf('Total Dollar Value');

  if (totalIdx === -1)
    totalIdx = headerOE.indexOf('Amount'); // The non history invoices are titled with Amount and not Total Dollar Value

  const isCompletedOrders = headerOE[totalIdx];
  const custNumIdx = headerOE.indexOf('Customer');
  const dateIdx = headerOE.indexOf('Created Date');
  const orderNumIdx = isCompletedOrders && headerOE.indexOf('Order') || headerOE.indexOf('Order #');
  const invoiceNumIdx = isCompletedOrders && headerOE.indexOf('Invoice') || headerOE.indexOf('Inv #'); 
  const locationIdx = headerOE.indexOf('Loc');
  const customerNameIdx = headerOE.indexOf('Name');
  const employeeNameIdx = headerOE.indexOf('Created by User');
  const isOrderCompleteIdx = headerOE.indexOf('Order Complete?');
  const invoiceDateIdx = (headerOE.indexOf('Inv Date') !== -1) ? headerOE.indexOf('Inv Date') : headerOE.indexOf('OE Invoice Date');
  const invoicedByIdx = headerOE.indexOf('OE Invoice Initials');
  const months = {'01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun', '07': 'Jul', '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'};
  
  const lodgeCustomerSheet = spreadsheet.getSheetByName('Lodge Customer List');
  const charterGuideCustomerSheet = spreadsheet.getSheetByName('Charter & Guide Customer List');
  const lodgeCustomerNumbers = lodgeCustomerSheet.getSheetValues(3, 1, lodgeCustomerSheet.getLastRow() - 2, 1).flat();
  const charterGuideCustomerNumbers = charterGuideCustomerSheet.getSheetValues(3, 1, charterGuideCustomerSheet.getLastRow() - 2, 1).flat();
  const lodgeCustomerNames = lodgeCustomerSheet.getSheetValues(3, 2, lodgeCustomerSheet.getLastRow() - 2, 2);
  const charterGuideCustomerNames = charterGuideCustomerSheet.getSheetValues(3, 2, charterGuideCustomerSheet.getLastRow() - 2, 2);

  const lodgeOrdersSheet = spreadsheet.getSheetByName('LODGE ORDERS');
  const charterGuideOrdersSheet = spreadsheet.getSheetByName('GUIDE ORDERS');
  const lodgeCompletedSheet = spreadsheet.getSheetByName('LODGE COMPLETED');
  const charterGuideCompletedSheet = spreadsheet.getSheetByName('GUIDE COMPLETED');
  
  const numLodgeOrders = lodgeOrdersSheet.getRange(lodgeOrdersSheet.getLastRow(), 7).getNextDataCell(SpreadsheetApp.Direction.UP).getRow() - 2 || lodgeOrdersSheet.getLastRow() - 2;
  const numCharterGuideOrders = charterGuideOrdersSheet.getRange(charterGuideOrdersSheet.getLastRow(), 7).getNextDataCell(SpreadsheetApp.Direction.UP).getRow() - 2 || charterGuideOrdersSheet.getLastRow() - 2;
  const numCompletedLodgeOrders = lodgeCompletedSheet.getLastRow() - 2;
  const numCompletedCharterGuideOrders = charterGuideCompletedSheet.getLastRow() - 2;

  const lodgeOrders = (numLodgeOrders > 0) ? lodgeOrdersSheet.getSheetValues(3, 3, numLodgeOrders, 1).flat().map(ordNum => ordNum.toString()) : [];
  const charterGuideOrders = (numCharterGuideOrders > 0) ? charterGuideOrdersSheet.getSheetValues(3, 3, numCharterGuideOrders, 1).flat().map(ordNum => ordNum.toString()) : [];
  const lodgeCompleted = (numCompletedLodgeOrders > 0) ? lodgeCompletedSheet.getSheetValues(3, 12, numCompletedLodgeOrders, 1).flat().map(ordNum => ordNum.toString()) : [];
  const charterGuideCompleted = (numCompletedCharterGuideOrders > 0) ? charterGuideCompletedSheet.getSheetValues(3, 12, numCompletedCharterGuideOrders, 1).flat().map(ordNum => ordNum.toString()) : [];

  const newLodgeOrders = (isCompletedOrders) ? // If true, then the import is a set of invoiced and completed orders
    allOrders.filter(order => lodgeCustomerNumbers.includes(order[custNumIdx]) && order[dateIdx].substring(6) === '2024' && !lodgeCompleted.includes(order[invoiceNumIdx].toString().trim())).map(order => {
      return [getDateString(order[dateIdx], months), getFullName(order[employeeNameIdx]), order[orderNumIdx], 'TRUE', '', getBoStatus(order[isOrderCompleteIdx]), getProperTypesetName(order[customerNameIdx], lodgeCustomerNames, 1), getLocationName(order[locationIdx]), '', '', 'This order was automatically imported', order[invoiceNumIdx], '$' + order[totalIdx], getFullName(order[invoicedByIdx]), getOrderStatus(order[isOrderCompleteIdx], isCompletedOrders), getDateString(order[invoiceDateIdx], months)] // Lodge Completed
    }) :
    allOrders.filter(order => lodgeCustomerNumbers.includes(order[custNumIdx]) && order[dateIdx].substring(6) === '2024' && order[isOrderCompleteIdx] === 'No' && !lodgeOrders.includes(order[orderNumIdx].toString().trim())).map(order => {
      return [getDateString(order[dateIdx], months), getFullName(order[employeeNameIdx]), order[orderNumIdx], '', '', '', getProperTypesetName(order[customerNameIdx], lodgeCustomerNames, 1), getLocationName(order[locationIdx]), '', '', 'This order was automatically imported', getInvoiceNumber(order[invoiceNumIdx], isCompletedOrders), '', '', getOrderStatus(order[isOrderCompleteIdx], isCompletedOrders, order[invoiceNumIdx])] // Lodge Orders
  });

  const newCharterGuideOrders = (isCompletedOrders) ?  // If true, then the import is a set of invoiced and completed orders
    allOrders.filter(order => charterGuideCustomerNumbers.includes(order[custNumIdx]) && order[dateIdx].substring(6) === '2024' && !charterGuideCompleted.includes(order[invoiceNumIdx].toString().trim())).map(order => { 
      return [getDateString(order[dateIdx], months), getFullName(order[employeeNameIdx]), order[orderNumIdx], 'TRUE', '', getBoStatus(order[isOrderCompleteIdx]), getProperTypesetName(order[customerNameIdx], charterGuideCustomerNames, 1), getLocationName(order[locationIdx]), '', '', 'This order was automatically imported', order[invoiceNumIdx], '$' + order[totalIdx], getFullName(order[invoicedByIdx]), getOrderStatus(order[isOrderCompleteIdx], isCompletedOrders), getDateString(order[invoiceDateIdx], months)] // Charter & Guide Completed
    }) :
    allOrders.filter(order => charterGuideCustomerNumbers.includes(order[custNumIdx]) && order[dateIdx].substring(6) === '2024' && order[isOrderCompleteIdx] === 'No' && !charterGuideOrders.includes(order[orderNumIdx].toString().trim())).map(order => {
      return [getDateString(order[dateIdx], months), getFullName(order[employeeNameIdx]), order[orderNumIdx], '', '', '', getProperTypesetName(order[customerNameIdx], charterGuideCustomerNames, 1), getLocationName(order[locationIdx]), '', '', 'This order was automatically imported', getInvoiceNumber(order[invoiceNumIdx], isCompletedOrders), '', '', getOrderStatus(order[isOrderCompleteIdx], isCompletedOrders, order[invoiceNumIdx])] // Charter & Guide Orders
  });

  const numNewLodgeOrder = newLodgeOrders.length;
  const numNewCharterGuideOrder = newCharterGuideOrders.length;

  if (numNewLodgeOrder > 0)
  {
    var numCols = newLodgeOrders[0].length;

    if (isCompletedOrders)
      lodgeCompletedSheet.getRange(numCompletedLodgeOrders + 3, 1, numNewLodgeOrder, numCols)
          .setNumberFormats(new Array(numNewLodgeOrder).fill(['MMM dd, yyyy', '@', '@', '#', '@', '@', '@', '@', '@', '@', '@', '@', '$#,##0.00', '@', '@', 'MMM dd, yyyy'])).setValues(newLodgeOrders)
        .offset(-1*numCompletedLodgeOrders, 0, numCompletedLodgeOrders + numNewLodgeOrder, numCols).sort([{column: 16, ascending: true}, {column: 1, ascending: true}]);
    else
      lodgeOrdersSheet.getRange(numLodgeOrders + 3, 1, numNewLodgeOrder, numCols)
          .setNumberFormats(new Array(numNewLodgeOrder).fill(['MMM dd, yyyy', '@', '@', '#', '@', '@', '@', '@', '@', '@', '@', '@', '$#,##0.00', '@', '@'])).setValues(newLodgeOrders)
        .offset(-1*numLodgeOrders, 0, numLodgeOrders + numNewLodgeOrder, numCols).sort([{column: 1, ascending: true}]);

    Logger.log('The following new Lodge orders were added to the tracker:')
    Logger.log(newLodgeOrders)

    deleteBackOrderedItems(newLodgeOrders, spreadsheet);
  }

  if (numNewCharterGuideOrder > 0)
  {
    var numCols = newCharterGuideOrders[0].length;

    if (isCompletedOrders)
      charterGuideCompletedSheet.getRange(numCompletedCharterGuideOrders + 3, 1, numNewCharterGuideOrder, numCols)
          .setNumberFormats(new Array(numNewCharterGuideOrder).fill(['MMM dd, yyyy', '@', '@', '#', '@', '@', '@', '@', '@', '@', '@', '@', '$#,##0.00', '@', '@', 'MMM dd, yyyy'])).setValues(newCharterGuideOrders)
        .offset(-1*numCompletedCharterGuideOrders, 0, numCompletedCharterGuideOrders + numNewCharterGuideOrder, numCols).sort([{column: 16, ascending: true}, {column: 1, ascending: true}]);
    else
      charterGuideOrdersSheet.getRange(numCharterGuideOrders + 3, 1, numNewCharterGuideOrder, numCols)
          .setNumberFormats(new Array(numNewCharterGuideOrder).fill(['MMM dd, yyyy', '@', '@', '#', '@', '@', '@', '@', '@', '@', '@', '@', '$#,##0.00', '@', '@'])).setValues(newCharterGuideOrders)
        .offset(-1*numCharterGuideOrders, 0, numCharterGuideOrders + numNewCharterGuideOrder, numCols).sort([{column: 1, ascending: true}]);

    Logger.log('The following new Charter and Guide orders were added to the tracker:')
    Logger.log(newCharterGuideOrders)
  }

  // Orders that are fully completed may need to be removed from the Lodge Orders and Guide Orders page
  if (isCompletedOrders)
  {
    var isLodgeOrderComplete, isCharterGuideOrderComplete;
    SpreadsheetApp.flush();
    const completedLodgeOrders = lodgeCompletedSheet.getSheetValues(3, 3, lodgeCompletedSheet.getLastRow() - 2, 13)
      .filter(ord => ord[12] === 'Completed')
      .map(ord => ord[0]).flat()
      .filter(ordNum => ordNum !== ''); 

    Logger.log('The following Lodge Orders were removed because they were found to be fully completed as per the invoice history:')
    const currentLodgeOrders = lodgeOrdersSheet.getSheetValues(3, 1, numLodgeOrders, 15)
      .filter(currentOrd => {

        isLodgeOrderComplete = completedLodgeOrders.includes(currentOrd[2]);

        if (isLodgeOrderComplete)
          Logger.log(currentOrd);
        
        return !isLodgeOrderComplete;
      });

    var numCurrentLodgeOrders = currentLodgeOrders.length;

    if (numCurrentLodgeOrders < numLodgeOrders)
      lodgeOrdersSheet.getRange(3, 1, numLodgeOrders, 15).clearContent().offset(0, 0, numCurrentLodgeOrders, 15).setValues(currentLodgeOrders);

    const completedCharterGuideOrders = charterGuideCompletedSheet.getSheetValues(3, 3, charterGuideCompletedSheet.getLastRow() - 2, 13)
      .filter(ord => ord[14] === 'Completed')
      .map(ord => ord[2]).flat()
      .filter(ordNum => ordNum !== '');

    Logger.log('The following Guide Orders were removed because they were found to be fully completed as per the invoice history:')
    const currentCharterGuideOrders = charterGuideOrdersSheet.getSheetValues(3, 1, numCharterGuideOrders, 15)
      .filter(currentOrd => {

        isCharterGuideOrderComplete = completedCharterGuideOrders.includes(currentOrd[2]);

        if (isCharterGuideOrderComplete)
          Logger.log(currentOrd);

        return !isCharterGuideOrderComplete;
      });

    var numCurrentCharterGuideOrders = currentCharterGuideOrders.length;

    if (numCurrentCharterGuideOrders < numCharterGuideOrders)
      charterGuideOrdersSheet.getRange(3, 1, numCharterGuideOrders, 15).clearContent().offset(0, 0, numCurrentCharterGuideOrders, 15).setValues(currentCharterGuideOrders);
  }
  else
  {
    var numCurrentLodgeOrders = numLodgeOrders;
    var numCurrentCharterGuideOrders = numCharterGuideOrders;
  }

  spreadsheet.toast('LODGE: ' + numNewLodgeOrder + ' Added\n ' + (numLodgeOrders - numCurrentLodgeOrders) + ' Removed GUIDE: ' + numNewCharterGuideOrder + ' Added ' + (numCharterGuideOrders - numCurrentCharterGuideOrders) + ' Removed', 'Orders Imported', 60)
}