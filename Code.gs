/**
 * This function....
 * 
 * @param {Event Object} e : The event object from an installed onChange trigger.
 */
function onChange(e)
{
  try
  {
    processImportedData(e)
  }
  catch (error)
  {
    Logger.log(error['stack'])
    Browser.msgBox(error['stack'])
  }
}

/**
 * This function...
 * 
 * @param {Event Object} e : The event object from an installed onEdit trigger.
 */
function installedOnEdit(e)
{
  const spreadsheet = e.source;
  const sheet = spreadsheet.getActiveSheet();
  const sheetName = sheet.getSheetName();

  try
  {
    if (sheetName === 'B/O')
      updateBoHyperLink(e, sheet, spreadsheet)
    else if (sheetName === 'P/O')
      updatePoSheet(e, sheet, spreadsheet)
    else
      moveRow(e, sheet, sheetName, spreadsheet)  
  }
  catch (error)
  {
    Logger.log(error['stack'])
    Browser.msgBox(error['stack'])
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
  var firstRows = [], lastRows = [], numRows = [], values, sku = [], qty = [], name = [];
  
  // Find the first row and last row in the the set of all active ranges
  for (var r = 0; r < activeRanges.length; r++)
  {
    firstRows.push(activeRanges[r].getRow());
     lastRows.push(activeRanges[r].getLastRow())
      numRows.push(lastRows[r] - firstRows[r] + 1);
      values = activeSheet.getSheetValues(firstRows[r], 2, numRows[r], 4)[0]
     sku.push(values[3])
     qty.push(values[2])
    name.push(values[0])
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
    const items = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString()).filter(item => skus.includes(item[6]));
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
              url = 'https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit#gid=1340095049'
              sheet = SpreadsheetApp.openByUrl(url).getSheetByName('Order')
              itemValues = items.map((v,idx) => [today, 'Lodge\nTracker', qty[idx], v[0], v[1], 'ATTN: Nate & Deryk (Lodge Items)\n' + name[idx], v[3], '']) 
              row = sheet.getLastRow() + 1;
              numRows = itemValues.length;
              sheet.getRange(row, 1, numRows, 8).setNumberFormat('@').setValues(itemValues)
              applyFullRowFormatting(sheet, row, numRows, false)
              break;
            case 'Rupert':
              url = 'https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM/edit#gid=407280159'
              sheet = SpreadsheetApp.openByUrl(url).getSheetByName('Order')
              itemValues = items.map((v,idx) => [today, 'Lodge\nTracker', qty[idx], v[0], v[1], 'ATTN: Sonya (Lodge Items)\n' + name[idx], v[4], '']) 
              row = sheet.getLastRow() + 1;
              numRows = itemValues.length;
              sheet.getRange(row, 1, numRows, 8).setNumberFormat('@').setValues(itemValues)
              applyFullRowFormatting(sheet, row, numRows, false)
              break;
          }
          break;
        case 'Parksville':
          url = 'https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit#gid=269292771'
          sheet = SpreadsheetApp.openByUrl(url).getSheetByName('ItemsToRichmond')
          itemValues = items.map((v,idx) => [today, 'Lodge\nTracker', v[0], v[1], 'ATTN: Scott (Lodge Items)\n' + name[idx], qty[idx]]) 
          row = sheet.getLastRow() + 1;
          numRows = itemValues.length;
          sheet.getRange(row, 1, numRows, 6).setNumberFormat('@').setValues(itemValues)
          applyFullRowFormatting(sheet, row, numRows, true)
          break;
        case 'Rupert':
          url = 'https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM/edit#gid=1569594370'
          sheet = SpreadsheetApp.openByUrl(url).getSheetByName('ItemsToRichmond')
          itemValues = items.map((v,idx) => [today, 'Lodge\nTracker', v[0], v[1], 'ATTN: Scott (Lodge Items)\n' + name[idx], qty[idx]]) 
          row = sheet.getLastRow() + 1;
          numRows = itemValues.length;
          sheet.getRange(row, 1, numRows, 6).setNumberFormat('@').setValues(itemValues)
          applyFullRowFormatting(sheet, row, numRows, true)
          break;
      }

      if (sheetName == 'B/O' && fromLocation != undefined && toLocation != undefined)
        activeRanges.map(rng => activeSheet.getRange(rng.getRow(), 10, rng.getNumRows(), 1).setRichTextValue(
          SpreadsheetApp
            .newRichTextValue()
            .setText('Shipping from ' + fromLocation + ' to ' + toLocation)
            .setLinkUrl(url + '&range=B' + row)
            .build()))
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
 * This function clears the order status from a particular line when a user presses the close button on
 * the modeless dialogue box for uploading back ordered items.
 * 
 * @param {} :
 * @author Jarren Ralf
 */
function clearOrderStatus(orderNumber)
{
  const spreadsheet = SpreadsheetApp.getActive()
  const lodgeOrderSheet = spreadsheet.getSheetByName('LODGE ORDERS')
  const lodgeOrderNumbers = lodgeOrderSheet.getSheetValues(3, 3, lodgeOrderSheet.getLastRow() - 2, lodgeOrderSheet.getLastColumn() - 2)

  // Look for the order on the Lodge Orders page inorder to create a hyperlink
  for (var i = 0; i < lodgeOrderNumbers.length; i++)
  {
    if (lodgeOrderNumbers[i][0] == orderNumber.toString())
    {
      lodgeOrderSheet.getRange(i + 3, 14).clearContent()
      break;
    }
  }

  if (lodgeOrderNumbers.length === i)
  {
    const guideOrderSheet = spreadsheet.getSheetByName('GUIDE ORDERS')
    const guideOrderNumbers = guideOrderSheet.getSheetValues(3, 3, guideOrderSheet.getLastRow() - 2, guideOrderSheet.getLastColumn() - 2)

    // Look for the order on the Guide Orders page inorder to create a hyperlink
    for (var i = 0; i < guideOrderNumbers.length; i++)
    {
      if (guideOrderNumbers[i][0] == orderNumber.toString())
      {
        guideOrderSheet.getRange(i + 3, 14).clearContent()
        break;
      }
    }
  }
}

/**
 * This function clears the order status from a particular line when a user presses the close button on
 * the modeless dialogue box for updating purchase order items.
 * 
 * @author Jarren Ralf
 */
function clearPoStatus()
{
  const spreadsheet = SpreadsheetApp.getActive()
  const purchaseOrderSheet = spreadsheet.getSheetByName('P/O')
  purchaseOrderSheet.getRange(3, 12, purchaseOrderSheet.getLastRow() - 2, 1).clearContent()
}

/**
 * This function ...
 * 
 * @param {Range} notes : 
 * @param {Range} skus : 
 * @return {Number} The number of items that are not ordered
 * @author Jarren Ralf
 */
function COUNT_ITEMS_NOT_ORDERED(notes, skus)
{
  for (var i = 0,counter = 0; i < notes.length; i++)
    if (skus[i][0] !== '' && skus[i][0] !== 'Comment Line:' && notes[i][0] === '' && notes[i][1] === '' && notes[i][2] === '' && notes[i][3] === '')
      counter++;
  return counter
}

/**
 * This function retrieves the back ordered items that were just for the current order that the user is 
 * trying to update the back ordered inventory for.
 * 
 * @return {Object[][]} The item information for a back order, including qty, sku, and description.
 * @author Jarren Ralf
 */
function getBackOrderedItems()
{
  return SpreadsheetApp.getActive().getSheetByName('BO Items').getDataRange().getValues()
}

/**
 * This function...
 * 
 * @param {String} : 
 * @param {Spreadsheet} : 
 * @return {String} : The name of the customer (if found) otherwise, returns the customer number.
 * @author Jarren Ralf
 */
function getCustomerName(customerNumber, spreadsheet)
{
  const sheet = spreadsheet.getSheetByName('CUSTOMERS');
  const customers = sheet.getSheetValues(1, 1, sheet.getLastRow(), 2);

  for (var i = 0; i < customers.length; i++)
  {
    if (customerNumber === customers[i][0])
      return customers[i][1];
  }

  sheet.showSheet()
  spreadsheet.toast('Please add the customer name and number to the CUSTOMERS page', 'CUSTOMERS', 60)

  return customers[i][0]
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
 * This function checks if a particular header at the provided index number is present in the data or missing.
 * 
 * @param {Number} index : The index number of the header.
 * @return {Boolean} Returns true if the index for the header equals -1 (Not Found) or false if it is greater than or equal to zero.
 * @author Jarren Ralf
 */
function isHeaderMissing(index)
{
  return index === -1;
}

/**
 * This function retrieves the remaining purchase order items that were just for the current order that the user is 
 * trying to update the purchase order inventory for.
 * 
 * @return {Object[][]} The item information for a purchase order, including qty, sku, and description.
 * @author Jarren Ralf
 */
function getRemainingPurchaseOrderItems()
{
  return SpreadsheetApp.getActive().getSheetByName('PO Items').getDataRange().getValues()
}

/**
 * This function...
 * 
 * @param {e}
 * @param {Sheet} sheet : The active sheet.
 * @param {String} sheetName : The name of the active sheet.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function moveRow(e, sheet, sheetName, spreadsheet)
{
  const range = e.range;
  const row = range.rowStart;
  const col = range.columnStart;  

  if (row == range.rowEnd && col == range.columnEnd) // Only look at a single cell edit
  {
    const sheetNames = sheetName.split(" ") // Split the sheet name, which will be used to distinguish between Logde and Guide page

    if (sheetNames[1] == "ORDERS" && col == 14) // An edit is happening on one of the Order pages in column 14, the order status column
    {
      const value = e.value; 
      const rowValues = sheet.getSheetValues(row, 1, 1, 14); // Entire row values
      const timeZone = spreadsheet.getSpreadsheetTimeZone(); // Set the timezone

      rowValues[0][0] = Utilities.formatDate(rowValues[0][0], timeZone, 'MMM dd, yyyy'); // Set the format of the order date
      rowValues[0].push(   Utilities.formatDate(     new Date(), timeZone, 'MMM dd, yyyy')); // Set the current time for the completion date

      if (value == "Completed") // The order status is being set to complete 
      {
        rowValues[0][4] = ''; // Clear the Back Order column
        spreadsheet.getSheetByName(sheetNames[0] +  " COMPLETED").appendRow(rowValues[0]) // Move the row of values to the completed page
        sheet.deleteRow(row); // Delete the row from the order page

        if (rowValues[0][4] === 'BO')
        {
          if (rowValues[0][2] !== '')
          {
            var backOrderSheet = spreadsheet.getSheetByName('B/O')
            const backOrders = backOrderSheet.getRange(3, 1, backOrderSheet.getLastRow() - 2, 13).clearContent().getValues();
            const remainingOrders = backOrders.filter(v => v[8] !== rowValues[0][2]);

            if (remainingOrders !== 0)
              backOrderSheet.getRange(3, 1, remainingOrders.length, 13).setValues(remainingOrders)

            if (remainingOrders === backOrders.length)
            {
              backOrderSheet.activate()
              Browser.msgBox('There were no back order items on the B/O sheet pertaining to Order # ' + rowValues[0][2] + 
                '. Please review the back ordered items and delete any lines that are associated with this order.')
            }
          }
          else
          {
            backOrderSheet.activate()
            Browser.msgBox('Since the Order # is blank, you\'ll have to delete the corresponding back ordered items manually.')
          }
        }
      }
      else if (value == "Cancelled") // The order status is being set to cancelled 
      { 
        spreadsheet.getSheetByName("CANCELLED").appendRow(rowValues[0]) // Move the row of values to the cancelled page
        sheet.deleteRow(row); // Delete the row from the order page
      }
      else if (value == "Partial") // The order status is being set to partial
      {
        if (rowValues[0][4] !== 'BO')
        {
          partiallyCompleteOrder(row, rowValues, sheet, sheetNames, spreadsheet)

          var html = HtmlService.createHtmlOutputFromFile('backOrderPrompt.html')
            .setWidth(849) 
            .setHeight(630);
          SpreadsheetApp.getUi() 
            .showModelessDialog(html, "Download Back Order Data from Adagio Order Entry");
        }
        else // The partially completed line contained a back order
        {
          const orderNumber = sheet.getSheetValues(row, 3, 1, 1)[0][0];

          if (orderNumber !== '')
          {
            const backOrderSheet = spreadsheet.getSheetByName('B/O')
            const lastRow = backOrderSheet.getLastRow();

            if (lastRow !== 2)
            {
              const itemsOnBackOrder = backOrderSheet.getSheetValues(3, 4, lastRow - 2, 6)
                .filter(u => u[5] == orderNumber && u[1] !== 'Comment Line:' && u[0] != 0).map(v => [v[0], v[1], v[2], orderNumber]);

              if (itemsOnBackOrder.length !== 0)
              {
                spreadsheet.getSheetByName('BO Items').clearContents().getRange(1, 1, itemsOnBackOrder.length, 4).setValues(itemsOnBackOrder)

                var html = HtmlService.createHtmlOutputFromFile('partialBackOrderPrompt.html')
                  .setWidth(750) 
                  .setHeight(500);
                SpreadsheetApp.getUi() 
                  .showModelessDialog(html, "Update the Back Order Quantities");
              }
              else
              {
                const ui = SpreadsheetApp.getUi();

                var response = ui.alert('Ord # ' + orderNumber + ' Missing From Back Order Page', 'Is this order from Adagio OrderEntry?\n\nClick Yes to watch a video that explains how to upload the current back ordered information first before partially completing this order.\n\nClick No to complete this order without providing the neccessary data to track our back ordered items. NOT RECOMMENDED.', ui.ButtonSet.YES_NO_CANCEL)
                
                if (response == ui.Button.YES)
                {
                  range.setValue('')
                  var html = HtmlService.createHtmlOutputFromFile('backOrderPrompt.html')
                    .setWidth(849) 
                    .setHeight(630);
                  ui.showModelessDialog(html, "Download Back Order Data from Adagio Order Entry");
                }
                else if (response == ui.Button.NO)
                  partiallyCompleteOrder(row, rowValues, sheet, sheetNames, spreadsheet)
                else // The user clicked cancel or close
                  range.setValue('')
              }
            }
            else
            {
              const ui = SpreadsheetApp.getUi();

              var response = ui.alert('Data Missing from Back Order Page', 'Is this order, # ' + orderNumber + ', from Adagio OrderEntry?\n\nClick Yes to watch a video that explains how to upload the current back ordered information first before partially completing this order.\n\nClick No to complete this order without providing the neccessary data to track our back ordered items. NOT RECOMMENDED.', ui.ButtonSet.YES_NO_CANCEL)
              
              if (response == ui.Button.YES)
              {
                range.setValue('')
                var html = HtmlService.createHtmlOutputFromFile('backOrderPrompt.html')
                  .setWidth(849) 
                  .setHeight(630);
                ui.showModelessDialog(html, "Download Back Order Data from Adagio Order Entry");
              }
              else if (response == ui.Button.NO)
                partiallyCompleteOrder(row, rowValues, sheet, sheetNames, spreadsheet)
              else // The user clicked cancel or close
                range.setValue('')
            }
          }
          else
          {
            const ui = SpreadsheetApp.getUi();

            var response = ui.alert('Order # Missing from Order Page', 'Do you have an Adagio order # that you can input?\n\nClick Yes and input the number in the appropriate cell, then redo the process of partially completing this order.\n\nClick No to complete this order without the neccessary information required to update the back ordered items. NOT RECOMMENDED.', ui.ButtonSet.YES_NO_CANCEL)
            
            if (response == ui.Button.YES)
              range.setValue('').offset(0, -11).activate()
            else if (response == ui.Button.NO)
              partiallyCompleteOrder(row, rowValues, sheet, sheetNames, spreadsheet)
            else // The user clicked cancel or close
              range.setValue('')
          }
        }
      }
    }
  }
}

/**
 * This function...
 * 
 * @param {} : This var...
 * @param {} : This var...
 * @param {} : This var...
 * @param {} : This var...
 * @param {} : This var...
 * @author Jarren Ralf
 */
function partiallyCompleteOrder(row, rowValues, sheet, sheetNames, spreadsheet)
{
  rowValues[0][4] = 'BO'; // Set the value in the back order column to 'BO'
  spreadsheet.getSheetByName(sheetNames[0] +  " COMPLETED").appendRow(rowValues[0]); // Move the row of values to the completed page
  rowValues[0][10] = 'multiple'; // Set the invoice numbers to multiple
  rowValues[0][11] = ''; // Set the Invoice Value to blank
  rowValues[0][12] = ''; // Set the Invoiced By to blank
  rowValues[0][13] = ''; // Set the Order Status to blank
  rowValues[0].pop() // Remove the completed date (which shows up on the completed orders page)
  sheet.getRange(row, 1, 1, 14).setValues(rowValues); // Clear the invoice values, and set the status
}

/**
 * This function...
 * 
 * @author Jarren Ralf
 */
function partialOrderDataRetreivalInstructions()
{
  var html = HtmlService.createHtmlOutputFromFile('backOrderPrompt.html')
    .setWidth(1400)
    .setHeight(600);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showModelessDialog(html, "Download Back Order Data from Adagio Order Entry");
}

/**
 * This function....
 * 
 * @param {Event Object} : The event object on an spreadsheet edit.
 * @author Jarren Ralf
 */
function processImportedData(e)
{
  
  if (e.changeType === 'INSERT_GRID')
  {
    var spreadsheet = e.source;
    var sheets = spreadsheet.getSheets();
    var info, numRows = 0, numCols = 1, maxRow = 2, maxCol = 3;

    for (var sheet = 0; sheet < sheets.length; sheet++) // Loop through all of the sheets in this spreadsheet and find the new one
    {
      info = [
        sheets[sheet].getLastRow(),
        sheets[sheet].getLastColumn(),
        sheets[sheet].getMaxRows(),
        sheets[sheet].getMaxColumns()
      ]

      // A new sheet is imported by File -> Import -> Insert new sheet(s) - The left disjunct is for a csv and the right disjunct is for an excel file
      if ((info[maxRow] - info[numRows] === 2 && info[maxCol] - info[numCols] === 2) || 
          (info[maxRow] === 1000 && info[maxCol] === 26 && info[numRows] !== 0 && info[numCols] !== 0)) 
      {
        const values = sheets[sheet].getSheetValues(1, 1, info[numRows], info[numCols]); // This is the order entry data
        var orderNumber = sheets[sheet].getSheetName()
        updateBO_Or_PO(values, spreadsheet, orderNumber)

        if (sheets[sheet].getSheetName().substring(0, 7) !== "Copy Of") // Don't delete the sheets that are duplicates
          spreadsheet.deleteSheet(sheets[sheet]) // Delete the new sheet that was created

        break;
      }
    }

    // Try and find the file created and delete it
    var file1 = DriveApp.getFilesByName(orderNumber + '.xlsx')
    var file2 = DriveApp.getFilesByName("Book1.xlsx")

    if (file1.hasNext())
      file1.next().setTrashed(true)

    if (file2.hasNext())
      file2.next().setTrashed(true)
  }
}

/**
 * This function updates the back ordered quantites of a particular order. It is run from the html ui that is launched when a "Partial" order status is selected when 
 * there is an existing back order.
 * 
 * @param {Number[]} qtys : The new shipping quantities for the back ordered items of the particular order in question.
 * @author Jarren Ralf
 */
function updateBackOrderedItems(qtys)
{
  const spreadsheet = SpreadsheetApp.getActive();
  const backOrderedItems = spreadsheet.getSheetByName('BO Items').getDataRange().getValues()
  const numItems = backOrderedItems.length;
  const backOrderSheet = spreadsheet.getSheetByName('B/O');
  const numCols = backOrderSheet.getLastColumn()
  const all_Items_Range = backOrderSheet.getRange(3, 1, backOrderSheet.getLastRow() - 2, numCols)
  const all_Items = all_Items_Range.getValues();
  const orderNumber = backOrderedItems[0][3]

  for (var i = 0; i < all_Items.length; i++)
  {
    if (all_Items[i][8] == orderNumber)
    {
      var j = 0;
      var numCompleted = 0;

      while (j < numItems)
      {
        if ( all_Items[i][8] == backOrderedItems[j][3]  // Order Number
          && all_Items[i][4] != 'Comment Line:'         // SKU says Comment Line and therefore ignore
          && all_Items[i][3] == backOrderedItems[j][0]  // Back Order Quantity
          && all_Items[i][4] == backOrderedItems[j][1]  // SKU
          && all_Items[i][5] == backOrderedItems[j][2]) // Descriptions
        {
          if (all_Items[i][3] <= qtys[j])
          {
            all_Items[i] = 'Complete';
            numCompleted++;
          }
          else
            all_Items[i][3] -= qtys[j]

          j++;
        }

        i++;
      }

      break;
    }
  }

  if (numCompleted !== numItems)
  {
    var boItems = all_Items.filter(item => item != 'Complete')
    var numRows = boItems.length
    var numberFormats = new Array(numRows).fill(['mmm dd, yyyy', '@', '#', '#', '@', '@', '$#,##0.00', '@', '@', '@', '@', '@', '@'])
    var range = all_Items_Range.clearContent().offset(0, 0, numRows, numCols).setNumberFormats(numberFormats).setValues(boItems).activate();
    const lodgeOrderSheet = spreadsheet.getSheetByName('LODGE ORDERS')
    const lodgeOrderNumbers = lodgeOrderSheet.getSheetValues(3, 1, lodgeOrderSheet.getLastRow() - 2, lodgeOrderSheet.getLastColumn())

    // Look for the order on the Lodge Orders page inorder to create a hyperlink
    for (var i = 0; i < lodgeOrderNumbers.length; i++)
    {
      if (lodgeOrderNumbers[i][2] == orderNumber)
      { 
        const timeZone = spreadsheet.getSpreadsheetTimeZone(); // Set the timezone

        lodgeOrderNumbers[i][13] = 'Partial'; // Set the Order Status to partial
        lodgeOrderNumbers[i][0] = Utilities.formatDate(lodgeOrderNumbers[i][0], timeZone, 'MMM dd, yyyy'); // Set the format of the order date
        lodgeOrderNumbers[i].push(Utilities.formatDate(             new Date(), timeZone, 'MMM dd, yyyy')); // Set the current time for the completion date
        spreadsheet.getSheetByName("LODGE COMPLETED").appendRow(lodgeOrderNumbers[i]); // Move the row of values to the completed page
        lodgeOrderNumbers[i][10] = 'multiple'; // Set the invoice numbers to multiple
        lodgeOrderNumbers[i][11] = ''; // Set the Invoice Value to blank
        lodgeOrderNumbers[i][12] = ''; // Set the Invoiced By to blank
        lodgeOrderNumbers[i][13] = ''; // Set the Order Status to blank
        lodgeOrderNumbers[i].pop() // Remove the completed date (which shows up on the completed orders page)

        lodgeOrderSheet.getRange(i + 3, 1, 1, 14).setValues([lodgeOrderNumbers[i]]); // Clear the invoice values, and set the status

        lodgeOrderSheet.getRange(i + 3, 5).setRichTextValue(
          SpreadsheetApp.newRichTextValue()
            .setText('BO')
            .setLinkUrl('https://docs.google.com/spreadsheets/d/1n1EDjZM_fQs3FNKv6dpYl2yt03-pr0_YCSSML6o6rsw/edit#gid=' + backOrderSheet.getSheetId() + '&range=' + range.getA1Notation())
            .setTextStyle(
              SpreadsheetApp.newTextStyle()
                .setForegroundColor('#1155cc')
                .setUnderline(true)
                .build())
            .build())

        break;
      }
    }

    if (lodgeOrderNumbers.length === i)
    {
      const guideOrderSheet = spreadsheet.getSheetByName('GUIDE ORDERS')
      const guideOrderNumbers = guideOrderSheet.getSheetValues(3, 1, guideOrderSheet.getLastRow() - 2, guideOrderSheet.getLastColumn())

      // Look for the order on the Guide Orders page inorder to create a hyperlink
      for (var i = 0; i < guideOrderNumbers.length; i++)
      {
        if (guideOrderNumbers[i][2] == orderNumber)
        {
          const timeZone = spreadsheet.getSpreadsheetTimeZone(); // Set the timezone

          guideOrderNumbers[i][13] = 'Partial'; // Set the Order Status to partial
          guideOrderNumbers[i][0] = Utilities.formatDate(guideOrderNumbers[i][0], timeZone, 'MMM dd, yyyy'); // Set the format of the order date
          guideOrderNumbers[i].push(Utilities.formatDate(             new Date(), timeZone, 'MMM dd, yyyy')); // Set the current time for the completion date
          spreadsheet.getSheetByName("GUIDE COMPLETED").appendRow(guideOrderNumbers[i]); // Move the row of values to the completed page
          guideOrderNumbers[i][10] = 'multiple'; // Set the invoice numbers to multiple
          guideOrderNumbers[i][11] = ''; // Set the Invoice Value to blank
          guideOrderNumbers[i][12] = ''; // Set the Invoiced By to blank
          guideOrderNumbers[i][13] = ''; // Set the Order Status to blank
          guideOrderNumbers[i].pop() // Remove the completed date (which shows up on the completed orders page)

          guideOrderSheet.getRange(i + 3, 1, 1, 14).setValues([guideOrderNumbers[i]]); // Clear the invoice values, and set the status

          guideOrderSheet.getRange(i + 3, 5).setRichTextValue(
            SpreadsheetApp.newRichTextValue()
              .setText('BO')
              .setLinkUrl('https://docs.google.com/spreadsheets/d/1n1EDjZM_fQs3FNKv6dpYl2yt03-pr0_YCSSML6o6rsw/edit#gid=' + backOrderSheet.getSheetId() + '&range=' + range.getA1Notation())
              .setTextStyle(
                SpreadsheetApp.newTextStyle()
                  .setForegroundColor('#1155cc')
                  .setUnderline(true)
                  .build())
              .build())

          break;
        }
      }
    }
  }
  else
  {
    const lodgeOrderSheet = spreadsheet.getSheetByName('LODGE ORDERS')
    const lodgeOrderNumbers = lodgeOrderSheet.getSheetValues(3, 1, lodgeOrderSheet.getLastRow() - 2, lodgeOrderSheet.getLastColumn())

    // Look for the order on the Lodge Orders page inorder to delete the appropriate row
    for (var i = 0; i < lodgeOrderNumbers.length; i++)
    {
      if (lodgeOrderNumbers[i][2] == orderNumber)
      {
        const timeZone = spreadsheet.getSpreadsheetTimeZone(); // Set the timezone

        lodgeOrderNumbers[i][13] = 'Completed'; // Set the Order Status to Completed
        lodgeOrderNumbers[i][4] = ''; // Clear the Back Order column
        lodgeOrderNumbers[i][0] = Utilities.formatDate(lodgeOrderNumbers[i][0], timeZone, 'MMM dd, yyyy'); // Set the format of the order date
        lodgeOrderNumbers[i].push(Utilities.formatDate(             new Date(), timeZone, 'MMM dd, yyyy')); // Set the current time for the completion date
        spreadsheet.getSheetByName("LODGE COMPLETED").appendRow(lodgeOrderNumbers[i]); // Move the row of values to the completed page
        
        lodgeOrderSheet.deleteRow(i + 3); // Delete the row from the order page
        break;
      }
    }

    if (lodgeOrderNumbers.length === i)
    {
      const guideOrderSheet = spreadsheet.getSheetByName('GUIDE ORDERS')
      const guideOrderNumbers = guideOrderSheet.getSheetValues(3, 1, guideOrderSheet.getLastRow() - 2, guideOrderSheet.getLastColumn())

      // Look for the order on the Guide Orders page inorder to create a hyperlink
      for (var i = 0; i < guideOrderNumbers.length; i++)
      {
        if (guideOrderNumbers[i][2] == orderNumber)
        {
          const timeZone = spreadsheet.getSpreadsheetTimeZone(); // Set the timezone

          guideOrderNumbers[i][13] = 'Completed'; // Set the Order Status to Completed
          guideOrderNumbers[i][4] = ''; // Clear the Back Order column
          guideOrderNumbers[i][0] = Utilities.formatDate(guideOrderNumbers[i][0], timeZone, 'MMM dd, yyyy'); // Set the format of the order date
          guideOrderNumbers[i].push(Utilities.formatDate(             new Date(), timeZone, 'MMM dd, yyyy')); // Set the current time for the completion date
          spreadsheet.getSheetByName("GUIDE COMPLETED").appendRow(guideOrderNumbers[i]); // Move the row of values to the completed page

          guideOrderSheet.deleteRow(i + 3); // Delete the row from the order page
          break;
        }
      }
    }

    var boItems = all_Items.filter(order => order != 'Complete' && order[8] != orderNumber)
    var rng = all_Items_Range.clearContent()
  
    if (boItems.length !== 0)
      rng.offset(0, 0, boItems.length, numCols).setValues(boItems);
  }
}

/**
 * This function updates the back ordered quantites of a particular order.
 * 
 * @param {Number[]} qtys : The new shipping quantities for the back ordered items of the particular order in question.
 * @author Jarren Ralf
 */
function updatePoItems(qtys)
{
  const spreadsheet = SpreadsheetApp.getActive();
  const purchaseOrderItems = spreadsheet.getSheetByName('PO Items').getDataRange().getValues()
  const numItems = purchaseOrderItems.length;
  const purchaseOrderSheet = spreadsheet.getSheetByName('P/O');
  const numCols = purchaseOrderSheet.getLastColumn()
  const all_Items_Range = purchaseOrderSheet.getRange(3, 1, purchaseOrderSheet.getLastRow() - 2, numCols)
  const all_Items = all_Items_Range.getValues();
  const poNumber = purchaseOrderItems[0][3]

  for (var i = 0; i < all_Items.length; i++)
  {
    if (all_Items[i][7] == poNumber)
    {
      var j = 0;
      var numCompleted = 0;

      while (j < numItems)
      {
        if ( all_Items[i][7] == purchaseOrderItems[j][3]  // Order Number
          && all_Items[i][4] != ''                        // SKU blank and therefore ignore
          && all_Items[i][3] == purchaseOrderItems[j][0]  // Back Order Quantity
          && all_Items[i][4] == purchaseOrderItems[j][1]  // SKU
          && all_Items[i][5] == purchaseOrderItems[j][2]) // Descriptions
        {
          if (all_Items[i][3] <= qtys[j])
          {
            all_Items[i] = 'Complete';
            numCompleted++;
          }
          else
            all_Items[i][3] -= qtys[j]

          j++;
        }

        i++;
      }

      break;
    }
  }

  if (numCompleted !== numItems)
  {
    var poItems = all_Items.filter(item => item != 'Complete')
    var numRows = poItems.length
    var numberFormats = new Array(numRows).fill(['mmm dd, yyyy', '@', '#', '#', '@', '@', '@', '@', '@', '@', '@', '@'])
    all_Items_Range.clearContent().offset(0, 0, numRows, numCols).setNumberFormats(numberFormats).setValues(poItems).activate();
  }
  else
  {
    var poItems = all_Items.filter(order => order != 'Complete' && order[7] != poNumber)
    var numRows = poItems.length
    var rng = all_Items_Range.clearContent()
  
    if (numRows !== 0)
      rng.offset(0, 0, numRows, numCols).setValues(poItems);
  }

  const backOrderSheet = spreadsheet.getSheetByName('B/O')
  const poNumbers = backOrderSheet.getSheetValues(3, 12, backOrderSheet.getLastRow() - 2, 2)
  const nRows = poNumbers.length;

  for (var i = 0; i < nRows; i++)
  {
    if (poNumbers[i][0] == poNumber)
    { 
      poNumbers[i][1] = 'Arrived: ' + Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'MMM dd, yyyy')

      backOrderSheet.getRange(i + 3, 12).setRichTextValue( // Remove the hyperlink
        SpreadsheetApp.newRichTextValue()
          .setText(poNumber)
          .setLinkUrl(null)
          .setTextStyle(SpreadsheetApp.newTextStyle().setForegroundColor('black').setUnderline(false).build())
          .build())
    }
  }

  const receivedItems = poNumbers.map(p => [p[1]]);
  backOrderSheet.getRange(3, 13, nRows).setValues(receivedItems)
}

/**
 * This function....
 * 
 * @param {Object[][]} data : This..
 * @param {Spreadsheet} spreadsheet : This..
 * @param {Number} orderNumber : This 
 * @author Jarren Ralf
 */
function updateBO_Or_PO(data, spreadsheet, orderNumber)
{
  data.pop() // Remove the data summary row from the bottom of the data
  const header = data.shift();

  const VENDOR_NAME = header.indexOf('Vendor name');

  if (VENDOR_NAME === -1) // This must be an order from Adagio OrderEntry and therefore a back order
  {
    const DATE = header.indexOf('Date')
    const CUSTOMER_NUM = header.indexOf('Cust #')
    const SHIPPING_LOCATION = header.indexOf('Loc')
    const SKU = header.indexOf('Item')
    const DESCRIPTION = header.indexOf('Description')
    const ORDERED_QTY = header.indexOf('Qty Original Ordered')
    const BACK_ORDER_QTY = header.indexOf('Backorder')
    const PRICE = header.indexOf('Unit Price')
    const LOCATION_NAMES = {'100': 'Trites', '200': 'Parksville', '300': 'Rupert', '400': 'Trites'}

    const customerName = isHeaderMissing(CUSTOMER_NUM) ? '' : getCustomerName(data[0][CUSTOMER_NUM], spreadsheet);
    const orderDate = isHeaderMissing(DATE) ? '' : data[0][DATE];
    const shippingLocation = isHeaderMissing(SHIPPING_LOCATION) ? '' : LOCATION_NAMES[data[0][SHIPPING_LOCATION]]

    var backOrders = data.map(item => [
      orderDate,
      customerName,
      isHeaderMissing(ORDERED_QTY) ? '' : item[ORDERED_QTY],
      isHeaderMissing(BACK_ORDER_QTY) ? '' : item[BACK_ORDER_QTY],
      isHeaderMissing(SKU) ? '' : (item[SKU] !== 'Comment') ? item[SKU].toString().substring(0, 4) + 
                                                              item[SKU].toString().substring(5, 9) + 
                                                              item[SKU].toString().substring(10) :
                                                              'Comment Line:',
      isHeaderMissing(DESCRIPTION) ? '' : item[DESCRIPTION],
      isHeaderMissing(PRICE) ? '' : item[PRICE],
      shippingLocation
    ])

    const numItems = backOrders.length;
    const backOrderSheet = spreadsheet.getSheetByName('B/O');

    if (orderNumber !== 'Sheet1') // The user changed the tab name in the excel document to be the Adagio order number
    {
      backOrders = backOrders.map(row => {row.push(orderNumber); return row;})
      var numberFormats = new Array(numItems).fill(['mmm dd, yyyy', '@', '#', '#', '@', '@', '$#,##0.00', '@', '@'])

      const range = backOrderSheet.getRange(backOrderSheet.getLastRow() + 1, 1, numItems, backOrders[0].length).setNumberFormats(numberFormats).setValues(backOrders)
        .offset(0,0, numItems, backOrderSheet.getLastColumn()).activate()
      const lodgeOrderSheet = spreadsheet.getSheetByName('LODGE ORDERS')
      const lodgeOrderNumbers = lodgeOrderSheet.getSheetValues(3, 3, lodgeOrderSheet.getLastRow() - 2, 1)

      // Look for the order on the Lodge Orders page inorder to create a hyperlink
      for (var i = 0; i < lodgeOrderNumbers.length; i++)
      {
        if (lodgeOrderNumbers[i][0] === orderNumber)
        {
          lodgeOrderSheet.getRange(i + 3, 5).setRichTextValue(
            SpreadsheetApp.newRichTextValue()
              .setText('BO')
              .setLinkUrl('https://docs.google.com/spreadsheets/d/1n1EDjZM_fQs3FNKv6dpYl2yt03-pr0_YCSSML6o6rsw/edit#gid=' + backOrderSheet.getSheetId() + '&range=' + range.getA1Notation())
              .setTextStyle(
                SpreadsheetApp.newTextStyle()
                  .setForegroundColor('#1155cc')
                  .setUnderline(true)
                  .build())
              .build())

          return;
        }
      }

      const guideOrderSheet = spreadsheet.getSheetByName('GUIDE ORDERS')
      const guideOrderNumbers = guideOrderSheet.getSheetValues(3, 3, guideOrderSheet.getLastRow() - 2, 1)

      // Look for the order on the Guide Orders page inorder to create a hyperlink
      for (var i = 0; i < guideOrderNumbers.length; i++)
      {
        if (guideOrderNumbers[i][0] === orderNumber)
        {
          guideOrderSheet.getRange(i + 3, 5).setRichTextValue(
            SpreadsheetApp.newRichTextValue()
              .setText('BO')
              .setLinkUrl('https://docs.google.com/spreadsheets/d/1n1EDjZM_fQs3FNKv6dpYl2yt03-pr0_YCSSML6o6rsw/edit#gid=' + backOrderSheet.getSheetId() + '&range=' + range.getA1Notation())
              .setTextStyle(
                SpreadsheetApp.newTextStyle()
                  .setForegroundColor('#1155cc')
                  .setUnderline(true)
                  .build())
              .build())

          return;
        }
      }
    }
    else
    {
      var numberFormats = new Array(numItems).fill(['mmm dd, yyyy', '@', '#', '#', '@', '@', '$#,##0.00', '@'])
      backOrderSheet.getRange(backOrderSheet.getLastRow() + 1, 1, numItems, backOrders[0].length).setNumberFormats(numberFormats).setValues(backOrders).activate()
    }
  }
  else // This is a purchase order
  {
    
    const ORDERED_QTY = header.indexOf('Qty Originally Ordered')
    const BACK_ORDER_QTY = header.indexOf('Backordered')
    const SKU = header.indexOf('Item#')
    const DESCRIPTION = header.indexOf('Description')
    const SHIPPING_LOCATION = header.indexOf('Location')
    const PO_NUM = header.indexOf('Doc #')
    const LOCATION_NAMES = {'100': 'Trites', '200': 'Parksville', '300': 'Rupert', '400': 'Trites'}

    const orderDate = new Date();
    const poNumber = isHeaderMissing(PO_NUM) ? '' : data[0][PO_NUM];
    const shippingLocation = isHeaderMissing(SHIPPING_LOCATION) ? '' : LOCATION_NAMES[data[0][SHIPPING_LOCATION]]

    var purchaseOrders = data.map(item => [
      orderDate,
      item[VENDOR_NAME],
      isHeaderMissing(ORDERED_QTY) ? '' : item[ORDERED_QTY],
      isHeaderMissing(BACK_ORDER_QTY) ? '' : item[BACK_ORDER_QTY],
      isHeaderMissing(SKU) ? '' : item[SKU].toString().substring(0, 4) + 
                                            item[SKU].toString().substring(5, 9) + 
                                            item[SKU].toString().substring(10),
      isHeaderMissing(DESCRIPTION) ? '' : item[DESCRIPTION],
      shippingLocation,
      poNumber
    ])

    const numItems = purchaseOrders.length;
    const purchaseOrderSheet = spreadsheet.getSheetByName('P/O');

    var numberFormats = new Array(numItems).fill(['mmm dd, yyyy', '@', '#', '#', '@', '@', '@', '@'])
    var range = purchaseOrderSheet.getRange(purchaseOrderSheet.getLastRow() + 1, 1, numItems, purchaseOrders[0].length).setNumberFormats(numberFormats).setValues(purchaseOrders)
      .offset(0,0, numItems, purchaseOrderSheet.getLastColumn()).activate()
    const backOrderSheet = spreadsheet.getSheetByName('B/O')
    const backOrders = backOrderSheet.getSheetValues(3, 12, backOrderSheet.getLastRow() - 2, 1)

    // Look for the order on the B/O page inorder to create a hyperlink
    for (var i = 0; i < backOrders.length; i++)
    {
      if (backOrders[i][0] === poNumber)
      {
        backOrderSheet.getRange(i + 3, 12).setRichTextValue(
          SpreadsheetApp.newRichTextValue()
            .setText(poNumber)
            .setLinkUrl('https://docs.google.com/spreadsheets/d/1n1EDjZM_fQs3FNKv6dpYl2yt03-pr0_YCSSML6o6rsw/edit#gid=' + purchaseOrderSheet.getSheetId() + '&range=' + range.getA1Notation())
            .setTextStyle(
              SpreadsheetApp.newTextStyle()
                .setForegroundColor('#1155cc')
                .setUnderline(true)
                .build())
            .build())
      }
    }
  }
}

/**
 * This function detects when a user updates the B/O page in the order # column for the purposes of updating the hyperlink on the order pages.
 * 
 * @param {Event Object} e : The even object of an onEdit triggered event.
 * @param {Sheet} backOrderSheet : The sheet of the back order page.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet. 
 * @author Jarren Ralf
 */
function updateBoHyperLink(e, backOrderSheet, spreadsheet)
{
  const range = e.range;
  const values = range.getValues()
  const row = range.rowStart;
  const rowEnd = range.rowEnd;
  const col = range.columnStart;
  const colEnd = range.columnEnd;
  const ORDER_NUM_COL = 9;
  const PO_NUM_COL = 12;

  if (row >= 3 && col == colEnd) // There is an edit to a single column
  {
    if (!isEveryValueBlank(values)) // This is someone deleting several cells
    {
      if (col === ORDER_NUM_COL) // There is an edit in the order number column
      {
        if (row == rowEnd) // This is a paste of a single value into the cell
        {
          const orderNumber = values[0][0];
          const rng = range.offset(0, -8, rowEnd - row + 1, colEnd + 4);
          const lodgeOrderSheet = spreadsheet.getSheetByName('LODGE ORDERS')
          const lodgeOrderNumbers = lodgeOrderSheet.getSheetValues(3, 3, lodgeOrderSheet.getLastRow() - 2, 1)

          // Look for the order on the Lodge Orders page inorder to create a hyperlink
          for (var i = 0; i < lodgeOrderNumbers.length; i++)
          {
            if (lodgeOrderNumbers[i][0] == orderNumber)
            { 
              lodgeOrderSheet.getRange(i + 3, 5).setRichTextValue(
                SpreadsheetApp.newRichTextValue()
                  .setText('BO')
                  .setLinkUrl('https://docs.google.com/spreadsheets/d/1n1EDjZM_fQs3FNKv6dpYl2yt03-pr0_YCSSML6o6rsw/edit#gid=' + backOrderSheet.getSheetId() + '&range=' + rng.getA1Notation())
                  .setTextStyle(
                    SpreadsheetApp.newTextStyle()
                      .setForegroundColor('#1155cc')
                      .setUnderline(true)
                      .build())
                  .build())

              break;
            }
          }

          if (lodgeOrderNumbers.length === i)
          {
            const guideOrderSheet = spreadsheet.getSheetByName('GUIDE ORDERS')
            const guideOrderNumbers = guideOrderSheet.getSheetValues(3, 3, guideOrderSheet.getLastRow() - 2, 1)

            // Look for the order on the Guide Orders page inorder to create a hyperlink
            for (var i = 0; i < guideOrderNumbers.length; i++)
            {
              if (guideOrderNumbers[i][0] == orderNumber)
              {
                guideOrderSheet.getRange(i + 3, 5).setRichTextValue(
                  SpreadsheetApp.newRichTextValue()
                    .setText('BO')
                    .setLinkUrl('https://docs.google.com/spreadsheets/d/1n1EDjZM_fQs3FNKv6dpYl2yt03-pr0_YCSSML6o6rsw/edit#gid=' + backOrderSheet.getSheetId() + '&range=' + rng.getA1Notation())
                    .setTextStyle(
                      SpreadsheetApp.newTextStyle()
                        .setForegroundColor('#1155cc')
                        .setUnderline(true)
                        .build())
                    .build())

                break;
              }
            }
          }
        }
        else // This is presumed to be a drag down event i.e. an edit of multiple rows
        {
          const orderNumber = range.offset(-1, 0).getValue();
          const rng = range.setValue(orderNumber).offset(-1, -8, rowEnd - row + 2, colEnd + 4);

          const lodgeOrderSheet = spreadsheet.getSheetByName('LODGE ORDERS')
          const lodgeOrderNumbers = lodgeOrderSheet.getSheetValues(3, 3, lodgeOrderSheet.getLastRow() - 2, 1)

          // Look for the order on the Lodge Orders page inorder to create a hyperlink
          for (var i = 0; i < lodgeOrderNumbers.length; i++)
          {
            if (lodgeOrderNumbers[i][0] == orderNumber)
            { 
              lodgeOrderSheet.getRange(i + 3, 5).setRichTextValue(
                SpreadsheetApp.newRichTextValue()
                  .setText('BO')
                  .setLinkUrl('https://docs.google.com/spreadsheets/d/1n1EDjZM_fQs3FNKv6dpYl2yt03-pr0_YCSSML6o6rsw/edit#gid=' + backOrderSheet.getSheetId() + '&range=' + rng.getA1Notation())
                  .setTextStyle(
                    SpreadsheetApp.newTextStyle()
                      .setForegroundColor('#1155cc')
                      .setUnderline(true)
                      .build())
                  .build())

              break;
            }
          }

          if (lodgeOrderNumbers.length === i)
          {
            const guideOrderSheet = spreadsheet.getSheetByName('GUIDE ORDERS')
            const guideOrderNumbers = guideOrderSheet.getSheetValues(3, 3, guideOrderSheet.getLastRow() - 2, 1)

            // Look for the order on the Guide Orders page inorder to create a hyperlink
            for (var i = 0; i < guideOrderNumbers.length; i++)
            {
              if (guideOrderNumbers[i][0] == orderNumber)
              {
                guideOrderSheet.getRange(i + 3, 5).setRichTextValue(
                  SpreadsheetApp.newRichTextValue()
                    .setText('BO')
                    .setLinkUrl('https://docs.google.com/spreadsheets/d/1n1EDjZM_fQs3FNKv6dpYl2yt03-pr0_YCSSML6o6rsw/edit#gid=' + backOrderSheet.getSheetId() + '&range=' + rng.getA1Notation())
                    .setTextStyle(
                      SpreadsheetApp.newTextStyle()
                        .setForegroundColor('#1155cc')
                        .setUnderline(true)
                        .build())
                    .build())

                break;
              }
            }
          }
        }
      }
      else if (col === PO_NUM_COL) // There is an edit in the po number column
      {
        var poNumber = (row == rowEnd) ? values[0][0] : range.offset(-1, 0).getValue();

        const purchaseOrderSheet = spreadsheet.getSheetByName('P/O')
        const purchaseOrderNumbers = purchaseOrderSheet.getSheetValues(3, 8, purchaseOrderSheet.getLastRow() - 2, 1)
        var numRows = 0, startRow = 0;

        // Look for the po on the P/O page inorder to create a hyperlink
        for (var i = 0; i < purchaseOrderNumbers.length; i++)
        {
          if (purchaseOrderNumbers[i][0] == poNumber)
          {
            if (numRows === 0)
              startRow = i + 3;

            numRows++;
          }
          else if (numRows !== 0)
            break;
        }

        if (numRows !== 0)
        {
          range.setValue(poNumber)
          range.setRichTextValue(
            SpreadsheetApp.newRichTextValue()
              .setText(poNumber)
              .setLinkUrl('https://docs.google.com/spreadsheets/d/1n1EDjZM_fQs3FNKv6dpYl2yt03-pr0_YCSSML6o6rsw/edit#gid=' + purchaseOrderSheet.getSheetId() + '&range=A' + startRow + ':L' + (startRow + numRows - 1))
              .setTextStyle(
                SpreadsheetApp.newTextStyle()
                  .setForegroundColor('#1155cc')
                  .setUnderline(true)
                  .build())
              .build())
        }        
      }
    }
  }
}

/**
 * This function detects when a user updates the 
 * 
 * @param {Event Object} e : The even object of an onEdit triggered event.
 * @param {Sheet} purchaseOrderSheet : The sheet of the purchase order page.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet. 
 * @author Jarren Ralf
 */
function updatePoSheet(e, purchaseOrderSheet, spreadsheet)
{
  const range = e.range;
  const values = range.getValues()
  const row = range.rowStart;
  const rowEnd = range.rowEnd;
  const col = range.columnStart;
  const colEnd = range.columnEnd;

  if (col == colEnd) // Single Column is being edited
  {
    if (row == rowEnd) // Single Cell is being edited
    {
      if (row > 2 && col == 12)
      {
        if (values[0][0] === 'Every Item Received Fully')
        {
          const rng = purchaseOrderSheet.getRange(3, 1, purchaseOrderSheet.getLastRow() - 2, purchaseOrderSheet.getLastColumn())
          const data = rng.getValues();
          const poNumber = data[row - 3][7];
          const remainingItems = data.filter(po => po[7] != poNumber)
          const numItemsRemaining = remainingItems.length;
          rng.clearContent()

          if (numItemsRemaining !== 0)
            rng.offset(0, 0, numItemsRemaining, remainingItems[0].length).setValues(remainingItems)

          const backOrderSheet = spreadsheet.getSheetByName('B/O')
          const poNumbers = backOrderSheet.getSheetValues(3, 12, backOrderSheet.getLastRow() - 2, 2)
          const numRows = poNumbers.length;

          for (var i = 0; i < poNumbers.length; i++)
          {
            if (poNumbers[i][0] == poNumber)
            {
              const timeZone = spreadsheet.getSpreadsheetTimeZone(); // Set the timezone

              poNumbers[i][1] = 'Arrived: ' + Utilities.formatDate(new Date(), timeZone, 'MMM dd, yyyy')

              backOrderSheet.getRange(i + 3, 12).setRichTextValue( // Remove the hyperlink
                SpreadsheetApp.newRichTextValue()
                  .setText(poNumber)
                  .setLinkUrl(null)
                  .setTextStyle(SpreadsheetApp.newTextStyle().setForegroundColor('black').setUnderline(false).build())
                  .build())
            }
          }

          const receivedItems = poNumbers.map(p => [p[1]]);
          backOrderSheet.getRange(3, 13, numRows).setValues(receivedItems)
        }
        else if (values[0][0] === 'Some Items Partially Received')
        {
          const poNumber = purchaseOrderSheet.getSheetValues(row, 8, 1, 1)[0][0];

          if (poNumber !== '')
          {
            const remainingItemsOnPurchaseOrder = purchaseOrderSheet.getSheetValues(3, 4, purchaseOrderSheet.getLastRow() - 2, 5)
              .filter(u => u[4] == poNumber && u[0] != 0).map(v => [v[0], v[1], v[2], poNumber]);

            spreadsheet.getSheetByName('PO Items').clearContents().getRange(1, 1, remainingItemsOnPurchaseOrder.length, 4).setValues(remainingItemsOnPurchaseOrder)

            var html = HtmlService.createHtmlOutputFromFile('partialPurchaseOrderPrompt.html')
              .setWidth(750) 
              .setHeight(500);
            SpreadsheetApp.getUi() 
              .showModelessDialog(html, "Update the Purchase Order Quantities");
          }
          else
          {
            const ui = SpreadsheetApp.getUi();

            var response = ui.alert('Purchase Order # Missing', 'Is this purchase order only the one item on this row?\n\nClick Yes to remove just this line only.\n\nClick Cancel if you know the PO number so you can add it to all the appropriate rows in column H of this sheet. Then repeat the receiving process.', ui.ButtonSet.OK_CANCEL)
            
            if (response == ui.Button.YES)
              purchaseOrderSheet.deleteRow(row);
            else // The user clicked cancel or close
              range.setValue('')
          }
        }
      }
    }
    else // Multiple Rows in a single column are being edited
    {
      ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    }
  }
}