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
    var info, numRows = 0, numCols = 1, maxRow = 2, maxCol = 3, isAdagioOE = 4, isAdagioPO = 5, 
      isAdagioPO_Receipts = 6, isReceivedItems = 7, isBackOrderItems = 8, isPurchaseOrderItems = 9, 
      isInvoicedItems = 10, isCreditedItems = 11, nRows = 0, nCols = 0;

    for (var sheet = 0; sheet < sheets.length; sheet++) // Loop through all of the sheets in this spreadsheet and find the new one
    {
      if (sheets[sheet].getType() == SpreadsheetApp.SheetType.GRID) // Ignore the chart
      {
        nRows = sheets[sheet].getLastRow();
        nCols = sheets[sheet].getLastColumn();

        info = [
          nRows,
          nCols,
          sheets[sheet].getMaxRows(),
          sheets[sheet].getMaxColumns(),
          (nRows > 0 && nCols > 0) ? sheets[sheet].getSheetValues(1, 1, 1, nCols).flat().includes('Order Complete?')        : false, // There is a sheet with no rows and no columns
          (nRows > 0 && nCols > 0) ? sheets[sheet].getSheetValues(1, 1, 1, nCols).flat().includes('Automatic Style Code')   : false,
          (nRows > 0 && nCols > 0) ? sheets[sheet].getSheetValues(1, 1, 1, nCols).flat().includes('Receipt Date')           : false, 
          (nRows > 0 && nCols > 0) ? sheets[sheet].getSheetValues(1, 1, 1, nCols).flat().includes('Rcpt #')                 : false,
          (nRows > 0 && nCols > 0) ? sheets[sheet].getSheetValues(1, 1, 1, nCols).flat().includes('Qty Original Ordered')   : false, 
          (nRows > 0 && nCols > 0) ? sheets[sheet].getSheetValues(1, 1, 1, nCols).flat().includes('Qty Originally Ordered') : false,
          (nRows > 0 && nCols > 0) ? sheets[sheet].getSheetValues(1, 1, 1, nCols).flat().includes('Ordered')                : false,
          (nRows > 0 && nCols > 0) ? sheets[sheet].getSheetValues(1, 1, 1, nCols).flat().includes('Return') || 
                                     sheets[sheet].getSheetValues(1, 1, 1, nCols).flat().includes('Returned')               : false
        ]

        // A new sheet is imported by File -> Import -> Insert new sheet(s) - The left disjunct is for a csv and the right disjunct is for an excel file
        if ((info[maxRow] - info[numRows] === 2 && info[maxCol] - info[numCols] === 2) || (info[maxRow] === 1000 && info[maxCol] === 26 && info[numRows] !== 0 && info[numCols] !== 0) || 
            ((info[maxRow] === info[numRows] && (info[maxCol] === info[numCols] || info[maxCol] == 26)) && 
            (info[isAdagioOE] || info[isAdagioPO] || info[isAdagioPO_Receipts] || info[isBackOrderItems] || info[isPurchaseOrderItems] || info[isReceivedItems] || info[isInvoicedItems]))) 
        {
          spreadsheet.toast('Processing imported data...', '', 60)
          
          const values = sheets[sheet].getSheetValues(1, 1, info[numRows], info[numCols]); // This is the shopify order data
          const fileName = sheets[sheet].getSheetName();

          if (fileName.substring(0, 7) !== "Copy Of") // Don't delete the sheets that are duplicates
            spreadsheet.deleteSheet(sheets[sheet]) // Delete the new sheet that was created

          if (info[isAdagioOE])
            updateOrdersOnTracker(values, spreadsheet);
          else if (info[isReceivedItems])
            updateReceivedItemsOnTracker(values, spreadsheet);
          else if (info[isAdagioPO])
            updatePurchaseOrdersOnTracker(values, spreadsheet);
          else if (info[isAdagioPO_Receipts])
            updatePoReceiptsOnTracker(values, spreadsheet);
          else if (info[isBackOrderItems])
            updateItemsOnTracker(values, spreadsheet, fileName);
          else if (info[isPurchaseOrderItems])
            updatePoItemsOnTracker(values, spreadsheet);
          else if (info[isInvoicedItems])
            updateInvoicedItemsOnTracker(values, spreadsheet, fileName, false);
          else if (info[isCreditedItems])
            updateInvoicedItemsOnTracker(values, spreadsheet, fileName, true);
          
          break;
        }
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
  const spreadsheet = e.source;
  const sheet = spreadsheet.getActiveSheet();
  const sheetName = sheet.getSheetName();

  try
  {
    if (sheetName === 'Lead Cost & Pricing' || sheetName === 'Bait Cost & Pricing')
      managePriceChange(e, sheetName, spreadsheet)
    else if (sheetName.split(" ").pop() === 'ORDERS')
      moveRow(e, sheet, spreadsheet)
    else if (sheetName === 'Item Management (Jarren Only ;)')
      manageDocumentNumbers(e, sheet, spreadsheet)
  }
  catch (err)
  {
    var error = err['stack'] 
    Logger.log(error)
    Browser.msgBox(error)
  }
}

/**
 * This function checks to see if a user is moving a row from one sheet to another.
 * 
 * @param {Event Object} e : The event object.
 */
function onOpen(e)
{
  const spreadsheet = e.source;
  const newSpreadsheetUrl = spreadsheet.getSheetByName('New Tracker').getSheetValues(1, 1, 1, 1)[0][0];
  const areTriggersCreated = spreadsheet.getSheetByName('Triggers').getRange(1, 1).isChecked();
  const currentTransferSheetYear = spreadsheet.getSheetByName('LODGE ORDERS').getSheetValues(1, 1, 1, 1)[0][0].split(" ").shift();

  if (isBlank(newSpreadsheetUrl) && (new Date().getFullYear() + 1).toString() === currentTransferSheetYear && !areTriggersCreated)
    SpreadsheetApp.getUi().createMenu('Create Triggers').addItem('Create Triggers', 'triggers_CreateAll').addToUi();
}

/**
 * This function checks to see if a user is moving a row from one sheet to another.
 * 
 * @param {Event Object} e : The event object.
 */
function installedOnOpen(e)
{
  const today = new Date();
  const year = (today.getMonth() > 7) ? (today.getFullYear() + 1).toString() : today.getFullYear().toString();
  const spreadsheet = e.source;
  const lodgeOrdersSheet = spreadsheet.getSheetByName('LODGE ORDERS');
  const newSpreadsheetUrl = spreadsheet.getSheetByName('New Tracker').getSheetValues(1, 1, 1, 1)[0][0];
  const areTriggersCreated = spreadsheet.getSheetByName('Triggers').getRange(1, 1).isChecked();
  const currentTransferSheetYear = lodgeOrdersSheet.getSheetValues(1, 1, 1, 1)[0][0].split(" ").shift();
  var guideOrdersSheet, lodgeCompletedSheet, guideCompletedSheet;

  if (!isBlank(newSpreadsheetUrl))
  {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Add selected rows to new Lodge Tracker').addItem('Add selected rows to new Lodge Tracker', 'addSelectedRowsToNewLodgeTracker').addToUi();
    ui.showModalDialog(HtmlService.createHtmlOutput('<p><a href="' + newSpreadsheetUrl + '" target="_blank">' + year + ' Lodge Order Tracking 2.0</a></p>').setWidth(250).setHeight(50), 'New Lodge Tracker');
  }
  else if (year !== currentTransferSheetYear) // If it is september or later and the current spreadsheet is not "next years" spreadsheet
    if (today.getMonth() > 7) SpreadsheetApp.getUi().createMenu('Create ' + year + ' Lodge Tracker').addItem('Create ' + year + ' Lodge Tracker', 'createNewLodgeTracker').addToUi();
  else if (!areTriggersCreated)
    SpreadsheetApp.getUi().createMenu('Create Triggers').addItem('Create Triggers', 'triggers_CreateAll').addToUi();
  // else
  //   SpreadsheetApp.getUi().createMenu('PNT Menu').addItem('Check Approval of Selected Orders', 'sendEmails_CheckApprovalOfSelectedOrders').addToUi();

  [guideOrdersSheet, lodgeCompletedSheet, guideCompletedSheet] = setItemLinks(lodgeOrdersSheet, spreadsheet)
  setTransferSheetLinks(spreadsheet, lodgeOrdersSheet, guideOrdersSheet, lodgeCompletedSheet, guideCompletedSheet, spreadsheet.getSheetByName('B/O'),  spreadsheet.getSheetByName('I/O'))
}

/**
 * This function allows the user to add items from the I/O or B/O page to the relevant transfer page.
 * 
 * @author Jarren Ralf
 */
function addItemsToTransferSheet()
{
  const activeSheet = SpreadsheetApp.getActiveSheet();
  var activeRanges = activeSheet.getActiveRangeList().getRanges(); // The selected ranges
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

          response = ui.prompt('Which PNT location are you shipping TO?', 'Please type: \"parks\", or \"pr\".', ui.ButtonSet.OK_CANCEL);

          // Process the user's response.
          if (response.getSelectedButton() == ui.Button.OK)
          {
            textResponse = response.getResponseText().toUpperCase();

            if (textResponse == 'PARKS')
              toLocation = 'Parksville'
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
            case 'Parksville':
              url = 'https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit?gid=1340095049#gid=1340095049'
              sheet = SpreadsheetApp.openByUrl(url).getSheetByName('Order')
              itemValues = items.map((v,idx) => [today, 'Lodge\nTracker', qty[idx], v[0], v[1], 'ATTN: Jesse (Lodge Items)\n' + name[idx] + '\nORD# ' + ordNum[idx], v[3], '']) 
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

      if (fromLocation != undefined && toLocation != undefined && (sheetName == 'B/O' || sheetName == 'I/O'))
      {
        activeSheet?.getFilter()?.remove(); // Remove the filter
        activeSheet.getRange(2, 1, activeSheet.getLastRow() - 2, activeSheet.getLastColumn()).createFilter().sort(11, true); // Create a filter in the header and sort by the order number
        SpreadsheetApp.flush();
        const linkToTransferSheetText = "\nShipping from " + fromLocation + " to " + toLocation;
        const urlToTransferSheet = url + '&range=B' + row;
        var range, richText, fullText, fullTextLength, richText_Runs, numRuns = 0;

        activeRanges.map(rng => {
          range = rng.offset(0, 12 - rng.getColumn(), rng.getNumRows(), 1);
          richText = range.getRichTextValues().map(richTextVals => {
            fullText = richTextVals[0].getText();
            fullTextLength = fullText.length;
            richText_Runs = richTextVals[0].getRuns().map(run => [run.getStartIndex(), run.getEndIndex(), run.getTextStyle()]);
            numRuns = richText_Runs.length

            if (isNotBlank(fullText))
              for (var i = 0, richTextBuilder = SpreadsheetApp.newRichTextValue().setText(fullText + linkToTransferSheetText); i < numRuns; i++)
                richTextBuilder.setTextStyle(richText_Runs[i][0], richText_Runs[i][1], richText_Runs[i][2]);
            else
              return [SpreadsheetApp.newRichTextValue().setText('Shipping from ' + fromLocation + ' to ' + toLocation).setLinkUrl(urlToTransferSheet).build()];

            return [richTextBuilder.setLinkUrl(fullTextLength + 1, fullTextLength + linkToTransferSheetText.length, urlToTransferSheet).build()];
          })
          
          range.setRichTextValues(richText).setBackgrounds(range.getBackgrounds());
        });

        spreadsheet.toast('Click on the link and and make the necessary changes to the Transfer sheet', 'Item(s) Added to Transfer Sheet', -1)
      }
    }
  }
}

/**
 * This function takes all of the selected rows on the current lodge tracker sheet and it transfers them to the new one.
 * 
 * @author Jarren Ralf
 */
function addSelectedRowsToNewLodgeTracker()
{
  const currentSpreadsheet = SpreadsheetApp.getActive();
  const activeSheet = currentSpreadsheet.getActiveSheet();
  const sheetName = activeSheet.getSheetName()

  if (sheetName !== 'LODGE ORDERS' && sheetName !== 'GUIDE ORDERS')
    SpreadsheetApp.getUi().alert('You may only select orders from the LODGE ORDERS or GUIDE ORDERS sheet.');
  else
  {
    const numCols = activeSheet.getLastColumn();
    const firstRows = [], lastRows = [], numRows = [];
    
    const itemValues = currentSpreadsheet.getActiveRangeList().getRanges().map((activeRange, r) => {
      firstRows.push(activeRange.getRow())
      lastRows.push(activeRange.getLastRow())
        numRows.push(lastRows[r] - firstRows[r] + 1)
      return activeSheet.getSheetValues(firstRows[r], 1, numRows[r], numCols)
    })

    if (Math.min(...firstRows) > 2 && Math.max( ...lastRows) <= activeSheet.getLastRow()) // If the user has not selected an item, alert them with an error message
    {   
      const itemVals = [].concat.apply([], itemValues); // Concatenate all of the item values as a 2-D array
      const numOrders = itemVals.length;
      var url = currentSpreadsheet.getSheetByName('New Tracker').getSheetValues(1, 1, 1, 1)[0][0];
      const destinationSheet = SpreadsheetApp.openByUrl(url).getSheetByName(activeSheet.getSheetName());
      const lastRow = destinationSheet.getLastRow();

      destinationSheet.getRange((lastRow > 2) ? lastRow + 1 : 3, 1, numOrders, numCols)
        .setNumberFormats(new Array(numOrders).fill(['MMM dd, yyyy', '@', '@', '#', '@', '@', '@', '@', '@', '@', 'MMM dd, yyyy', '@', '$#,##0.00', '@', '@'])).setValues(itemVals);

      Logger.log('The following rows where delete:')
      firstRows.sort((a,b) => b - a).map((row, r) => {Logger.log('row: ' + row); Logger.log('numRows[r]: ' + numRows[r]); activeSheet.deleteRows(row, numRows[r]);}); // Delete the rows that were moved over to the new tracker
      const sheetId = destinationSheet.getSheetId()
      url += '?gid=' + sheetId + '#gid=' + sheetId;
      SpreadsheetApp.flush()

      SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput('<p><a href="' + url + '" target="_blank">' + (new Date().getFullYear() + 1) + ' Lodge Order Tracking 2.0</a></p>')
        .setWidth(250).setHeight(50), 'New Lodge Tracker');
    }
    else
      SpreadsheetApp.getUi().alert('Please select an order from the list.');
  }
}

/**
 * This function allows the user to add order from the LODGE ORDERS or GUIDE ORDERS page to the relevant transfer sheet.
 * 
 * @author Jarren Ralf
 */
function addOrdersToTransferSheet()
{
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const firstRows = [], lastRows = [], numRows = [], values = [];
    
  SpreadsheetApp.getActive().getActiveRangeList().getRanges().map((activeRange, r) => {
    firstRows.push(activeRange.getRow())
     lastRows.push(activeRange.getLastRow())
      numRows.push(lastRows[r] - firstRows[r] + 1)
       values.push(...activeSheet.getSheetValues(firstRows[r], 3, numRows[r], 9))
  })

  if (Math.min(...firstRows) < 3)
    Browser.msgBox('Please select items from the list.')
  else
  {
    const spreadsheet = SpreadsheetApp.getActive()
    const sheetName = activeSheet.getSheetName();
    const today = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'dd MMM yyyy');
    var row = 0, numOrders = 0, sheet, itemValues, fromLocation, toLocation, url;

    var ui = SpreadsheetApp.getUi();

    var response = ui.prompt('Which PNT location are you shipping FROM?', 'Please type: \"rich", \"parks\", or \"pr\".', ui.ButtonSet.OK_CANCEL);

    // Process the user's response.
    if (response.getSelectedButton() == ui.Button.OK)
    {
      var textResponse = response.getResponseText().toUpperCase();

      if (textResponse == 'RICH')
      {
        fromLocation = 'Richmond';

        response = ui.prompt('Which PNT location are you shipping TO?', 'Please type: \"parks\", or \"pr\".', ui.ButtonSet.OK_CANCEL);

        // Process the user's response.
        if (response.getSelectedButton() == ui.Button.OK)
        {
          textResponse = response.getResponseText().toUpperCase();

          if (textResponse == 'PARKS')
            toLocation = 'Parksville'
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
          case 'Parksville':
            url = 'https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit?gid=1340095049#gid=1340095049'
            sheet = SpreadsheetApp.openByUrl(url).getSheetByName('Order')
            itemValues = values.map(value => 
              [today, 'Lodge\nTracker', '', '', 'Order# ' + value[0] + ' for ' + value[3] + ' - ' + ((isBlank(value[8]) || value[8] === 'multiple') ? 'NOT INVOICED' : 'Inv# ' + value[8]), 'ATTN: Jesse (Lodge Order)']) 
            row = sheet.getLastRow() + 1;
            numOrders = itemValues.length;
            sheet.getRange(row, 1, numOrders, 6).setNumberFormat('@').setValues(itemValues)
            applyFullRowFormatting(sheet, row, numOrders, false)
            break;
          case 'Rupert':
            url = 'https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit?gid=407280159#gid=407280159'
            sheet = SpreadsheetApp.openByUrl(url).getSheetByName('Order')
            itemValues = values.map(value => 
              [today, 'Lodge\nTracker', '', '', 'Order# ' + value[0] + ' for ' + value[3] + ' - ' + ((isBlank(value[8]) || value[8] === 'multiple') ? 'NOT INVOICED' : 'Inv# ' + value[8]), 'ATTN: Doug (Lodge Order)'])
            row = sheet.getLastRow() + 1;
            numOrders = itemValues.length;
            sheet.getRange(row, 1, numOrders, 6).setNumberFormat('@').setValues(itemValues)
            applyFullRowFormatting(sheet, row, numOrders, false)
            break;
        }
        break;
      case 'Parksville':
        url = 'https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit?gid=269292771#gid=269292771'
        sheet = SpreadsheetApp.openByUrl(url).getSheetByName('ItemsToRichmond')
        itemValues = values.map(value => 
          [today, 'Lodge\nTracker', '', 'Order# ' + value[0] + ' for ' + value[3] + ' - ' + ((isBlank(value[8]) || value[8] === 'multiple') ? 'NOT INVOICED' : 'Inv# ' + value[8]), 'ATTN: Scott (Lodge Order)']
        ) 
        row = sheet.getLastRow() + 1;
        numOrders = itemValues.length;
        sheet.getRange(row, 1, numOrders, 5).setNumberFormat('@').setValues(itemValues)
        applyFullRowFormatting(sheet, row, numOrders, true)
        break;
      case 'Rupert':
        url = 'https://docs.google.com/spreadsheets/d/1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c/edit?gid=1569594370#gid=1569594370'
        sheet = SpreadsheetApp.openByUrl(url).getSheetByName('ItemsToRichmond')
        itemValues = values.map(value => 
          [today, 'Lodge\nTracker', '', 'Order# ' + value[0] + ' for ' + value[3] + ' - ' + ((isBlank(value[8]) || value[8] === 'multiple') ? 'NOT INVOICED' : 'Inv# ' + value[8]), 'ATTN: Scott (Lodge Items)']
        ) 
        row = sheet.getLastRow() + 1;
        numOrders = itemValues.length;
        sheet.getRange(row, 1, numOrders, 5).setNumberFormat('@').setValues(itemValues)
        applyFullRowFormatting(sheet, row, numOrders, true)
        break;
    }

    if (fromLocation != undefined && toLocation != undefined && (sheetName == 'LODGE ORDERS' || sheetName == 'GUIDE ORDERS' || sheetName == 'LODGE COMPLETED' || sheetName == 'GUIDE COMPLETED'))
    {
      SpreadsheetApp.flush();
      const linkToTransferSheetText = '\nShipping from ' + fromLocation + ' to ' + toLocation;
      const urlToTransferSheet = url + '&range=B' + row;
      var rng, richText_Notes, fullText, fullTextLength, richText_Notes_Runs, numRuns = 0;

      numRows.map((nRows, r) => {
        rng = activeSheet.getRange(firstRows[r], 10, nRows, 1);
        richText_Notes = rng.getRichTextValues().map(note_RichText => {
          fullText = note_RichText[0].getText()
          fullTextLength = fullText.length;
          richText_Notes_Runs = note_RichText[0].getRuns().map(run => [run.getStartIndex(), run.getEndIndex(), run.getTextStyle()]);
          numRuns = richText_Notes_Runs.length;

          if (!isBlank(fullText))
            for (var i = 0, richTextBuilder = SpreadsheetApp.newRichTextValue().setText(fullText + linkToTransferSheetText); i < numRuns; i++)
              richTextBuilder.setTextStyle(richText_Notes_Runs[i][0], richText_Notes_Runs[i][1], richText_Notes_Runs[i][2])
          else 
            return [SpreadsheetApp.newRichTextValue().setText('Shipping from ' + fromLocation + ' to ' + toLocation).setLinkUrl(urlToTransferSheet).build()]
          
          return [richTextBuilder.setLinkUrl(fullTextLength + 1, fullTextLength + linkToTransferSheetText.length, urlToTransferSheet).build()];
        });

        rng.setRichTextValues(richText_Notes).setBackgrounds(rng.getBackgrounds());
      });

      spreadsheet.toast('Click on the link and and make the necessary changes to the Transfer sheet', 'Order(s) Added to Transfer Sheet', -1)
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
 * This function takes the given string, splits it at the chosen delimiter, and capitalizes the first letter of each perceived word.
 * 
 * @param {String} str : The given string
 * @param {String} delimiter : The delimiter that determines where to split the given string
 * @return {String} The output string with proper case
 * @author Jarren Ralf
 */
function capitalizeSubstrings(str, delimiter)
{
  var numLetters;
  var words = str.toString().split(delimiter); // Split the string at the chosen location/s based on the delimiter

  for (var word = 0, string = ''; word < words.length; word++) // Loop through all of the words in order to build the new string
  {
    numLetters = words[word].length;

    if (numLetters == 0) // The "word" is a blank string (a sentence contained 2 spaces)
      continue; // Skip this iterate
    else if (numLetters == 1) // Single character word
    {
      if (words[word][0] !== words[word][0].toUpperCase()) // If the single letter is not capitalized
        words[word] = words[word][0].toUpperCase(); // Then capitalize it
    }
    else if (numLetters == 2 && words[word].toUpperCase() === 'PO') // So that PO Box is displayed correctly
      words[word] = words[word].toUpperCase();
    else
    {
      /* If the first letter is not upper case or the second letter is not lower case, then
       * capitalize the first letter and make the rest of the word lower case.
       */
      if (words[word][0] !== words[word][0].toUpperCase() || words[word][1] !== words[word][1].toLowerCase())
        words[word] = words[word][0].toUpperCase() + words[word].substring(1).toLowerCase();
    }

    string += words[word] + delimiter; // Add a blank space at the end
  }

  string = string.slice(0, -1); // Remove the last space

  return string;
}

/**
 * This function creates a new Lodge Tracker for the next season.
 * 
 * @author Jarren Ralf
 */
function createNewLodgeTracker()
{
  if (Session.getActiveUser().getEmail() !== 'jarrencralf@gmail.com') 
    Browser.msgBox('Please ask Jarren to create the new spreadsheet so the full functionality of the Lodge Transfer sheet is preserved.')
  else
  {
    const year = new Date().getFullYear() + 1
    const currentSpreadsheet = SpreadsheetApp.getActive();
    const newSpreadsheet = currentSpreadsheet.copy(year + ' Lodge Order Tracking 2.0')
    const url = newSpreadsheet.getUrl();
    currentSpreadsheet.getSheetByName('New Tracker').getRange(1, 1).setValue(url);
    newSpreadsheet.getSheetByName('Triggers').getRange(1, 1).uncheck();
    newSpreadsheet.addEditors(currentSpreadsheet.getEditors().map(editor => editor.getEmail())).getSheetByName('New Tracker').clear() // Add edditors and remove old url
    const lodgeOrdersSheet = newSpreadsheet.getSheetByName('LODGE ORDERS')
    const guideOrdersSheet = newSpreadsheet.getSheetByName('GUIDE ORDERS')
    const lodgeCompletedSheet = newSpreadsheet.getSheetByName('LODGE COMPLETED')
    const guideCompletedSheet = newSpreadsheet.getSheetByName('GUIDE COMPLETED')
    const cancelledSheet = newSpreadsheet.getSheetByName('CANCELLED')
    const boSheet = newSpreadsheet.getSheetByName('B/O')
    const ioSheet = newSpreadsheet.getSheetByName('I/O')
    const poSheet = newSpreadsheet.getSheetByName('P/O')
    const numCols = boSheet.getLastColumn();
    
    lodgeOrdersSheet.getRange(1, 1).setValue(year + ' Lodge Orders').offset(2, 0, lodgeOrdersSheet.getMaxRows() - 2, lodgeOrdersSheet.getLastColumn()).clearContent()
    guideOrdersSheet.getRange(1, 1).setValue(year + ' Guide Orders').offset(2, 0, guideOrdersSheet.getMaxRows() - 2, guideOrdersSheet.getLastColumn()).clearContent()
    lodgeCompletedSheet.getRange(1, 1).setValue(year + ' Completed Lodge Orders').offset(2, 0, lodgeCompletedSheet.getMaxRows() - 2, lodgeCompletedSheet.getLastColumn()).clearContent()
    guideCompletedSheet.getRange(1, 1).setValue(year + ' Completed Guide Orders').offset(2, 0, guideCompletedSheet.getMaxRows() - 2, guideCompletedSheet.getLastColumn()).clearContent()
    cancelledSheet.getRange(1, 1).setValue(year + ' Cancelled Orders').offset(2, 0, cancelledSheet.getMaxRows() - 2, cancelledSheet.getLastColumn()).clearContent()
    boSheet.getRange(1, 1).setValue(year + ' Back Orders').offset(2, 0, boSheet.getMaxRows() - 2, numCols).clearContent()
    ioSheet.getRange(1, 1).setValue(year + '  Initial Items Ordered').offset(2, 0, ioSheet.getMaxRows() - 2, numCols).clearContent()
    poSheet.getRange(1, 1).setValue(year + ' Purchase Orders').offset(2, 0, poSheet.getMaxRows() - 2, poSheet.getLastColumn()).clearContent()
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput('<p><a href="' + url + '" target="_blank">' + year + ' Lodge Order Tracking 2.0</a></p>').setWidth(250).setHeight(50), 'New Lodge Tracker');
  }
}

/**
 * This function finds items with on the B/O tab that matches the given order number and deletes them.
 * 
 * @param {String || String[][]} orderNumber       : The order number of the current order being updated on the ORDERS page.
 * @param {Spreadsheet}          spreadsheet       : The active spreadsheet
 * @param {String[][]} listOfOrderCompletionStatus : A list of order numbers paired with either "Yes" for fully complete and "No" otherwise
 * @author Jarren Ralf
 */
function deleteBackOrderedItems(orderNumber, spreadsheet, listOfOrderCompletionStatus)
{
  const templateSheet = spreadsheet.getSheetByName('Reupload:');
  const boSheet = spreadsheet.getSheetByName('B/O');
  const ioSheet = spreadsheet.getSheetByName('I/O');
  const boSheet_NumRows = boSheet.getLastRow() - 2;
  const ioSheet_NumRows = ioSheet.getLastRow() - 2;
  const numCols = boSheet.getLastColumn();
  const boItems = (boSheet_NumRows > 0) ? boSheet.getSheetValues(3, 1, boSheet_NumRows, numCols) : null;
  const ioItems = (ioSheet_NumRows > 0) ? ioSheet.getSheetValues(3, 1, ioSheet_NumRows, numCols) : null;
  
  boSheet?.getFilter()?.remove(); // Remove the filter
  ioSheet?.getFilter()?.remove();
  boSheet.getRange(2, 1, boSheet_NumRows + 1, numCols).createFilter().sort(11, true)
  ioSheet.getRange(2, 1, ioSheet_NumRows + 1, numCols).createFilter().sort(11, true)
  SpreadsheetApp.flush();

  if (Array.isArray(orderNumber)) // When importing new orders the argument passed to this function is an array with multiple order numbers
  {
    var isLodgeOrderComplete, itemsOnOrder = [], row, numRows;

    if (listOfOrderCompletionStatus)
    {
      orderNumber.map(ordNum => {

        if (!isBlank(ordNum[2]))
        {
          isLodgeOrderComplete = listOfOrderCompletionStatus.find(partialOrd => partialOrd[0] == ordNum[2])

          if (isLodgeOrderComplete && isLodgeOrderComplete[1] === 'No')
          {
            if (ioItems)
            {
              row = ioItems.findIndex(ordNum_IO => ordNum_IO[10] == ordNum[2]);

              if (row !== -1)
              {
                numRows = Number(ioItems.findLastIndex(ordNum_IO => ordNum_IO[10] == ordNum[2])) - row + 1;
                itemsOnOrder = ioItems.filter(item => item[10] == ordNum[2]).map(item => [item[5], item[11], item[12], ''])

                if (itemsOnOrder.some(row => isNotBlank(row[1]) || isNotBlank(row[2]) || isNotBlank(row[3])))
                  spreadsheet.insertSheet('Reupload:' + ordNum[2], {template: templateSheet}).hideSheet()
                    .getRange(2, 1, itemsOnOrder.length, 4).setNumberFormat('@').setValues(itemsOnOrder)

                Logger.log('Found ORD# ' + ordNum[2] + ' on the I/O sheet. Will delete ' + numRows + ' row(s) starting at row ' + (row + 3));

                ioSheet.deleteRows(row + 3, numRows);
                SpreadsheetApp.flush();
              }
              else
              {
                row = boItems.findIndex(ordNum_BO => ordNum_BO[10] == ordNum[2]);

                if (row !== -1)
                {
                  numRows = boItems.findLastIndex(ordNum_BO => ordNum_BO[10] == ordNum[2]) - row + 1;
                  itemsOnOrder = boItems.filter(item => item[10] == ordNum[2]).map(item => [item[5], item[11], item[12], ''])

                  if (itemsOnOrder.some(row => isNotBlank(row[1]) || isNotBlank(row[2]) || isNotBlank(row[3])))
                    spreadsheet.insertSheet('Reupload:' + ordNum[2], {template: templateSheet}).hideSheet()
                      .getRange(2, 1, itemsOnOrder.length, 4).setNumberFormat('@').setValues(itemsOnOrder)

                  Logger.log('Found ORD# ' + ordNum[2] + ' on the B/O sheet. Will delete ' + numRows + ' row(s) starting at row ' + (row + 3));

                  boSheet.deleteRows(row + 3, numRows);
                  SpreadsheetApp.flush();
                }
              }
            }
            else if (boItems)
            {
              row = boItems.findIndex(ordNum_BO => ordNum_BO[10] == ordNum[2]);

              if (row !== -1)
              {
                numRows = boItems.findLastIndex(ordNum_BO => ordNum_BO[10] == ordNum[2]) - row + 1;
                itemsOnOrder = boItems.filter(item => item[10] == ordNum[2]).map(item => [item[5], item[11], item[12], ''])

                if (itemsOnOrder.some(row => isNotBlank(row[1]) || isNotBlank(row[2]) || isNotBlank(row[3])))
                  spreadsheet.insertSheet('Reupload:' + ordNum[2], {template: templateSheet}).hideSheet()
                    .getRange(2, 1, itemsOnOrder.length, 4).setNumberFormat('@').setValues(itemsOnOrder)

                Logger.log('Found ORD# ' + ordNum[2] + ' on the B/O sheet. Will delete ' + numRows + ' row(s) starting at row ' + (row + 3));

                boSheet.deleteRows(row + 3, numRows);
                SpreadsheetApp.flush();
              }
            }

            itemsOnOrder.length = 0;
          }
        }
      })
    }
    else
    {
      orderNumber.map(ordNum => {

        if (!isBlank(ordNum[2]))
        {
          // Back Orders Sheet
          if (boSheet_NumRows > 0)
          {
            orderNumbers = boSheet.getSheetValues(3, 11, boSheet_NumRows, 1);
            row = orderNumbers.findIndex(ordNum_BO => ordNum_BO[0] == ordNum[2]);

            if (row !== -1)
            {
              Logger.log('Found ORD# ' + ordNum[2] + ' on the B/O sheet. Will delete ' + numRows + ' row(s) starting at row ' + (row + 3));
              numRows = orderNumbers.findLastIndex(ordNum_BO => ordNum_BO[0] == ordNum[2]) - row + 1;
              boSheet.deleteRows(row + 3, numRows);
              SpreadsheetApp.flush();
            }
          }

          // Inital Orders Sheet
          if (ioSheet_NumRows > 0)
          {
            orderNumbers = ioSheet.getSheetValues(3, 11, ioSheet_NumRows, 1);
            row = orderNumbers.findIndex(ordNum_IO => ordNum_IO[0] == ordNum[2]);

            if (row !== -1)
            {
              Logger.log('Found ORD# ' + ordNum[2] + ' on the I/O sheet. Will delete ' + numRows + ' row(s) starting at row ' + (row + 3));
              numRows = orderNumbers.findLastIndex(ordNum_IO => ordNum_IO[0] == ordNum[2]) - row + 1;
              ioSheet.deleteRows(row + 3, numRows);
              SpreadsheetApp.flush();
            }
          }
        }
      })
    }
  }
  else
  {
    if (!isBlank(orderNumber)) // Order number is not blank on the Orders page
    {
      if (boSheet_NumRows > 0)
      {
        const orderNumbers_BO = boSheet.getSheetValues(3, 11, boSheet_NumRows, 1);
        const row_BO = orderNumbers_BO.findIndex(ordNum_BO => ordNum_BO[0] == orderNumber);

        if (row_BO !== -1)
        {
          Logger.log('Found ORD# ' + ordNum[2] + ' on the B/O sheet. Will delete ' + numRows + ' row(s) starting at row ' + (row + 3));
          const numRows_BO = orderNumbers_BO.findLastIndex(ordNum_BO => ordNum_BO[0] == orderNumber) - row_BO + 1;
          boSheet.deleteRows(row_BO + 3, numRows_BO);
        }
      }
      
      if (ioSheet_NumRows > 0)
      {
        const orderNumbers_IO = ioSheet.getSheetValues(3, 11, ioSheet_NumRows, 1);
        const row_IO = orderNumbers_IO.findIndex(ordNum_IO => ordNum_IO[0] == orderNumber);

        if (row_IO !== -1)
        {
          Logger.log('Found ORD# ' + ordNum[2] + ' on the I/O sheet. Will delete ' + numRows + ' row(s) starting at row ' + (row + 3));
          const numRows_IO = orderNumbers_IO.findLastIndex(ordNum_IO => ordNum_IO[0] == orderNumber) - row_IO + 1;
          ioSheet.deleteRows(row_IO + 3, numRows_IO);
        }
      }
    }
  }  

  SpreadsheetApp.flush()
  boSheet?.getFilter()?.remove(); // Remove the filter
  ioSheet?.getFilter()?.remove();
  const boSheet_NumRows_Updated = boSheet.getLastRow() - 1;
  const ioSheet_NumRows_Updated = ioSheet.getLastRow() - 1;

  if (boSheet_NumRows_Updated > 0)
    boSheet.getRange(2, 1, boSheet_NumRows_Updated, numCols).createFilter().sort(11, true); // Create a filter in the header and sort by the order number

  if (ioSheet_NumRows_Updated > 0)
    ioSheet.getRange(2, 1, ioSheet_NumRows_Updated, numCols).createFilter().sort(11, true); // Create a filter in the header and sort by the order number;
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
 * This function emails me if there has been a change in the lead cost or frozen bait cost so that I can update our system with the new information.
 * 
 * @author Jarren Ralf
 */
function emailCostChangeOfLeadOrFrozenBait()
{
  const SATURDAY = 6;
  const spreadsheet = SpreadsheetApp.getActive();
  const url = spreadsheet.getUrl();
  const leadSheet = spreadsheet.getSheetByName('Lead Cost & Pricing');
  const baitSheet = spreadsheet.getSheetByName('Bait Cost & Pricing');
  const hasLeadCostsChanged_OnThisSS = leadSheet.getSheetValues(3, leadSheet.getMaxColumns(), leadSheet.getLastRow() - 2, 1).some(recentChanges => recentChanges[0] === 'Yes');
  const hasBaitCostsChanged_OnThisSS = baitSheet.getSheetValues(3, baitSheet.getMaxColumns(), baitSheet.getLastRow() - 2, 1).some(recentChanges => recentChanges[0] === 'Yes');
  const hasLeadCostsChanged_InAdagio = leadSheet.getSheetValues(3, 15, leadSheet.getLastRow() - 2, 1).some(recentChanges => !isBlank(recentChanges[0]));

  if (new Date().getDay() !== SATURDAY) // Don't send the emails on Saturday (Sunday is fine because I can address them Monday morning)
  {
    if (hasLeadCostsChanged_OnThisSS || hasLeadCostsChanged_InAdagio)
      sendEmail(url + '?gid=' + leadSheet.getSheetId(), "Lead Cost & Pricing")

    if (hasBaitCostsChanged_OnThisSS)
      sendEmail(url + '?gid=' + baitSheet.getSheetId(), "Bait Cost & Pricing")
  }
}

/**
 * This function is passed the completed sheets and it sets the hyperlinks from those sheets to the Inv'd page.
 * 
 * @param {Sheet}  invdSheet : The sheet with invoiced items on it
 * @param {Sheet[]} sheets   : An array of sheets, assumed to be the two completed sheets
 * @author Jarren Ralf 
 */
function establishItemLinks_INVD(invdSheet, ...sheets)
{
  invdSheet?.getFilter()?.remove(); // Remove the filter
  SpreadsheetApp.flush();

  const invdSheet_NumRows = invdSheet.getLastRow() - 1
  invdSheet.getRange(2, 1, invdSheet_NumRows, invdSheet.getLastColumn()).createFilter().sort(12, true); // Create a filter in the header and sort by the invoice number
  SpreadsheetApp.flush();

  const invoiceAndCreditNumbers_Invd = (invdSheet_NumRows > 1) ? invdSheet.getSheetValues(3, 12, invdSheet_NumRows - 1, 2) : null;
  const invdSheetId = invdSheet.getSheetId()
  var invoiceNumber, notes, isCreditNumInNotes, creditNumber, startIndex, endIndex, row_invd, row_cred, numRows, range, notesAndinvoiceNumbers;

  sheets.map(sheet => {

    numRows = sheet.getLastRow() - 2;

    if (numRows > 0)
    {
      range = sheet.getRange(3, 10, numRows, 2).setNumberFormat('@');
      SpreadsheetApp.flush()
      
      notesAndinvoiceNumbers = range.getRichTextValues().map(rowVals => {
        invoiceNumber = rowVals[1].getText();
        notes = rowVals[0].getText();
        isCreditNumInNotes = notes.match(/\d{5}/); // match 5-digit number
        row_invd = (invoiceAndCreditNumbers_Invd) ? invoiceAndCreditNumbers_Invd.findIndex(inv => inv[0] == invoiceNumber && isBlank(inv[1])) + 3 : -1;
        row_cred = -1, creditNumber = '';

        if (isCreditNumInNotes)
        {
          creditNumber = isCreditNumInNotes[0];
          startIndex = notes.indexOf(creditNumber);
          endIndex = startIndex + creditNumber.length;
          row_cred = (invoiceAndCreditNumbers_Invd) ? invoiceAndCreditNumbers_Invd.findIndex(cred => cred[1] == creditNumber) + 3 : -1;
        }

        return (row_invd > 2 && row_cred > 2) ? 
            [rowVals[0].copy().setLinkUrl(startIndex, endIndex, '#gid=' + invdSheetId + '&range=A' + row_cred + ':M' + (invoiceAndCreditNumbers_Invd.findLastIndex(cred => cred[1] == creditNumber) + 3)).build(),
             rowVals[1].copy().setLinkUrl('#gid=' + invdSheetId + '&range=A' + row_invd + ':M' + (invoiceAndCreditNumbers_Invd.findLastIndex(inv  => inv[0] == invoiceNumber && isBlank(inv[1])) + 3)).build()] : 
          (row_invd > 2) ? 
            [rowVals[0],
             rowVals[1].copy().setLinkUrl('#gid=' + invdSheetId + '&range=A' + row_invd + ':M' + (invoiceAndCreditNumbers_Invd.findLastIndex(inv  => inv[0] == invoiceNumber && isBlank(inv[1])) + 3)).build()] : 
          (row_cred > 2) ? 
            [rowVals[0].copy().setLinkUrl(startIndex, endIndex, '#gid=' + invdSheetId + '&range=A' + row_cred + ':M' + (invoiceAndCreditNumbers_Invd.findLastIndex(cred => cred[1] == creditNumber) + 3)).build(),
             rowVals[1]] :
          rowVals;
      })

      range.setRichTextValues(notesAndinvoiceNumbers);
    }
  })
}

/**
 * This function is passed the order sheets and it sets the hyperlinks from those sheets to the I/O and B/O pages.
 * 
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @param {Sheet[]}       sheets    : An array of sheets, assumed to be the two order sheets
 * @return {Sheet[]} Returns the bo, io, po, and invd sheets.
 * @author Jarren Ralf 
 */
function establishItemLinks_IO_BO(spreadsheet, ...sheets)
{
  const   boSheet = spreadsheet.getSheetByName('B/O')
  const   ioSheet = spreadsheet.getSheetByName('I/O')
  const   poSheet = spreadsheet.getSheetByName('P/O')
  const invdSheet = spreadsheet.getSheetByName("Inv'd")
  boSheet?.getFilter()?.remove(); // Remove the filter
  ioSheet?.getFilter()?.remove(); // Remove the filter
  SpreadsheetApp.flush();

  const   boSheet_NumRowsPlusHeader = boSheet.getLastRow() - 1;
  const   ioSheet_NumRowsPlusHeader = ioSheet.getLastRow() - 1;
  const   boSheet_NumRows =   boSheet_NumRowsPlusHeader - 1;
  const   ioSheet_NumRows =   ioSheet_NumRowsPlusHeader - 1;
  const   poSheet_NumRows =   poSheet.getLastRow() - 2;
  const invdSheet_NumRows = invdSheet.getLastRow() - 2;
  const numCols = boSheet.getLastColumn()
    boSheet.getRange(2, 1, boSheet_NumRowsPlusHeader, numCols).createFilter().sort(11, true); // Create a filter in the header and sort by the order number
    ioSheet.getRange(2, 1, ioSheet_NumRowsPlusHeader, numCols).createFilter().sort(11, true); // Create a filter in the header and sort by the order number
  SpreadsheetApp.flush();
  
  const orderNumbersAndSku_BO   = (  boSheet_NumRows > 0) ?   boSheet.getSheetValues(3, 6,   boSheet_NumRows, 6) : null;
  const orderNumbersAndSku_IO   = (  ioSheet_NumRows > 0) ?   ioSheet.getSheetValues(3, 6,   ioSheet_NumRows, 6) : null;
  const orderNumbersAndSku_INVD = (invdSheet_NumRows > 0) ? invdSheet.getSheetValues(3, 6, invdSheet_NumRows, 6) : null;
  const   boSheetId =   boSheet.getSheetId();
  const   ioSheetId =   ioSheet.getSheetId();
  const invdSheetId = invdSheet.getSheetId();
  var numRows, range, orderNumbers, orderNumber, row_io, row_bo, row_invd, orderNumbersInNotes, richTextBuilder, startIndex, endIndex, notes;

  sheets.map(sheet => {

    numRows = sheet.getLastRow() - 2;

    if (numRows > 0)
    {
      range = sheet.getRange(3, 3, numRows, 1)

      orderNumbers = range.getRichTextValues().map(ordNum => {
        orderNumber = ordNum[0].getText();
        row_io = (orderNumbersAndSku_IO) ? orderNumbersAndSku_IO.findIndex(ord => ord[5] == orderNumber) + 3 : -1;

        if (row_io > 2)
          return [ordNum[0].copy().setLinkUrl('#gid=' + ioSheetId + '&range=A' + row_io + ':M' + (orderNumbersAndSku_IO.findLastIndex(ord => ord[5] == orderNumber) + 3)).build()]      
        else if (orderNumbersAndSku_BO) // Make sure there are back order items on the list
        {
          row_bo = orderNumbersAndSku_BO.findIndex(ord => ord[5] == orderNumber) + 3;
        
          return (row_bo > 2) ? [ordNum[0].copy().setLinkUrl('#gid=' + boSheetId + '&range=A' + row_bo + ':M' + (orderNumbersAndSku_BO.findLastIndex(ord => ord[5] == orderNumber) + 3)).build()] : ordNum;
        }
        else 
          return ordNum;
      })

      range.setRichTextValues(orderNumbers);
    }
  })

  if (poSheet_NumRows > 0)
  {
    const skus = poSheet.getSheetValues(3, 5, poSheet_NumRows, 1)
    range = poSheet.getRange(3, 11, poSheet_NumRows, 1)

    orderNumbers = range.getRichTextValues().map((noteValues, sku) => {

      notes = noteValues[0].getText();
      orderNumbersInNotes = [...notes.matchAll(/\b\d{5}\b/g)];

      if (orderNumbersInNotes.length > 0) // If there are order numbers in the notes
      {
        richTextBuilder = noteValues[0].copy();

        orderNumbersInNotes.map(ordNum => {
          orderNumber = ordNum[0];
          startIndex = notes.indexOf(orderNumber);
          endIndex = startIndex + orderNumber.length;
          row_bo   = (orderNumbersAndSku_BO)   ?   orderNumbersAndSku_BO.findIndex(ord => ord[5] == orderNumber && ord[0] == skus[sku][0]) + 3 : -1;
          row_io   = (orderNumbersAndSku_IO)   ?   orderNumbersAndSku_IO.findIndex(ord => ord[5] == orderNumber && ord[0] == skus[sku][0]) + 3 : -1;
          row_invd = (orderNumbersAndSku_INVD) ? orderNumbersAndSku_INVD.findIndex(ord => ord[5] == orderNumber && ord[0] == skus[sku][0]) + 3 : -1;

          if (row_bo > 2)
            richTextBuilder.setLinkUrl(startIndex, endIndex, '#gid=' +   boSheetId + '&range=A' + row_bo   + ':M' + row_bo)
          else if (row_io > 2)
            richTextBuilder.setLinkUrl(startIndex, endIndex, '#gid=' +   ioSheetId + '&range=A' + row_io   + ':M' + row_io)
          else if (row_invd > 2)
            richTextBuilder.setLinkUrl(startIndex, endIndex, '#gid=' + invdSheetId + '&range=A' + row_invd + ':M' + row_invd)
              .setTextStyle(startIndex, endIndex, SpreadsheetApp.newTextStyle().setForegroundColor('#b45f06').setUnderline(true).build())
        })

        return [richTextBuilder.build()];
      }
      else // No order numbers in the notes
        return noteValues;
    })

    range.setRichTextValues(orderNumbers); 
  }

  return [boSheet, ioSheet, poSheet, invdSheet]
}

/**
 * This function is passed the B/O and I/O sheets and it sets the hyperlinks from those sheets to P/O sheet.
 * 
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @param {Sheet}         poSheet   : The sheet with purchase order items on it
 * @param {Sheet[]}       sheets    : An array of sheets, assumed to be the two item sheet, B/O and I/O
 * @author Jarren Ralf 
 */
function establishItemLinks_PO(spreadsheet, poSheet, ...sheets)
{
  const recdSheet = spreadsheet.getSheetByName("Rec'd")
    poSheet?.getFilter()?.remove(); // Remove the filter
  recdSheet?.getFilter()?.remove(); // Remove the filter
  SpreadsheetApp.flush();

  const   poSheet_NumRowsPlusHeader =   poSheet.getLastRow() - 1;
  const recdSheet_NumRowsPlusHeader = recdSheet.getLastRow() - 1;
  const   poSheet_NumRows =   poSheet_NumRowsPlusHeader - 1;
  const recdSheet_NumRows = recdSheet_NumRowsPlusHeader - 1;
    poSheet.getRange(2, 1,   poSheet_NumRowsPlusHeader,   poSheet.getLastColumn()).createFilter().sort(10, true); // Create a filter in the header and sort by the purchase order number
  recdSheet.getRange(2, 1, recdSheet_NumRowsPlusHeader, recdSheet.getLastColumn()).createFilter().sort(12, true); // Create a filter in the header and sort by the receipt number
  SpreadsheetApp.flush();

  const purchaseOrderNumbersAndSku_PO = (  poSheet_NumRows > 0) ?   poSheet.getSheetValues(3, 5,   poSheet_NumRows, 6) : null;
  const     receiptNumbersAndSku_RECD = (recdSheet_NumRows > 0) ? recdSheet.getSheetValues(3, 6, recdSheet_NumRows, 7) : null;
  const   poSheetId =   poSheet.getSheetId()
  const recdSheetId = recdSheet.getSheetId()
  var numRows, skus, notesRange, purchaseOrderNumbers, notes, isPoNumInNotes, poNumber, startIndex, endIndex, row_PO, isReceiptNumInNotes, receiptNumber, row_RECD, idx_RECD, endOfPoNum;

  sheets.map(sheet => {

    numRows = sheet.getLastRow() - 2;

    if (numRows > 0)
    {
      skus = sheet.getSheetValues(3,  6, numRows, 1)
      notesRange = sheet.getRange(3, 12, numRows, 1)

      purchaseOrderNumbers = notesRange.getRichTextValues().map((noteValues, sku) => {

        notes = noteValues[0].getText();
        isPoNumInNotes = notes.match(/PO0\d{5}/); // match 5-digit number
        poNumber = '', row_PO = -1, row_RECD = -1;
        
        if (isPoNumInNotes)
        {
          poNumber = isPoNumInNotes[0];
          startIndex = notes.indexOf(poNumber);
          endIndex = startIndex + poNumber.length;
          row_PO   = (purchaseOrderNumbersAndSku_PO) ? purchaseOrderNumbersAndSku_PO.findIndex(po => po[5] == poNumber && po[0] == skus[sku][0]) + 3 : -1;
          isReceiptNumInNotes = notes.match(/RC0\d{5}/); // match 5-digit number

          if (isReceiptNumInNotes) // There is a receipt number in the notes, make sure the hyperlink is pointed to the correct row on the Rec'd page
          {
            receiptNumber = isReceiptNumInNotes[0];
            startIndex = notes.indexOf(receiptNumber);
            endIndex = startIndex + receiptNumber.length;
            row_RECD = (receiptNumbersAndSku_RECD) ? receiptNumbersAndSku_RECD.findIndex(rct => rct[6] == receiptNumber && rct[0] == skus[sku][0]) + 3 : -1;
          }
          else if (row_PO <= 2) // There is no receipt number and the PO number is not found on the P/O sheet, therefore check if the item is on the Rec'd sheet
          {
            idx_RECD = (receiptNumbersAndSku_RECD) ? receiptNumbersAndSku_RECD.findIndex(rct => rct[5] == poNumber && rct[0] == skus[sku][0]) : -1;

            if (idx_RECD !== -1) // The item was found on the Rec'd page using the PO number, therefore add the RC number to the notes and hyperlink it
            {
              receiptNumber = receiptNumbersAndSku_RECD[idx_RECD][6];
              endOfPoNum = startIndex + poNumber.length;
              endIndex = endOfPoNum + receiptNumber.length + 1;
              row_RECD = idx_RECD + 3;
              
              return [noteValues[0].copy()
                        .setText(notes.slice(0, endOfPoNum) + (" " + receiptNumber) + notes.slice(endOfPoNum))
                        .setLinkUrl(endOfPoNum + 1, endIndex, '#gid=' + recdSheetId + '&range=A' + row_RECD + ':L' + row_RECD).build()];
            }
          }
        }

        return (row_PO > 2) ? [noteValues[0].copy().setLinkUrl(startIndex, endIndex, '#gid=' +   poSheetId + '&range=A' + row_PO   + ':L' + row_PO  ).build()] : 
             (row_RECD > 2) ? [noteValues[0].copy().setLinkUrl(startIndex, endIndex, '#gid=' + recdSheetId + '&range=A' + row_RECD + ':L' + row_RECD).build()] : noteValues;
      })

      notesRange.setRichTextValues(purchaseOrderNumbers);
    }
  })
}

/**
 * This function accesses the invoice data from the Lodge, Charter, and Guide Data spreadsheet and extrats the last eight years of information. 
 * With that info if produces a chart of cummulative sales week-to-week for each year.
 * 
 * @author Jarren Ralf
 */
function getChartData()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const   currentYear = new Date().getFullYear();
  const      lastYear =   currentYear - 1;
  const   twoYearsAgo =      lastYear - 1;
  const threeYearsAgo =   twoYearsAgo - 1;
  const  fourYearsAgo = threeYearsAgo - 1;
  const  fiveYearsAgo =  fourYearsAgo - 1;
  const   sixYearsAgo =  fiveYearsAgo - 1;
  const sevenYearsAgo =   sixYearsAgo - 1; 
  const invoiceDataSheet = SpreadsheetApp.openById('1xKw4GAtNbAsTEodCDmCMbPCbXUlK9OHv0rt5gYzqx9c').getSheetByName('All Data');
  const dashboard = spreadsheet.getSheetByName('Dashboard')
  const ioAmount = Number(dashboard.getSheetValues(14, 9, 1, 1)[0][0]);
  const boAmount = Number(dashboard.getSheetValues(12, 9, 1, 1)[0][0]) + ioAmount;
  const millisecondsInWeek = 7 * 24 * 60 * 60 * 1000;
  var chartData = [], lastWeek = 0, currentSales = 0;

  const firstDayOfYear = {
      [currentYear]: new Date(  currentYear, 0, 1).getTime(),
         [lastYear]: new Date(     lastYear, 0, 1).getTime(),
      [twoYearsAgo]: new Date(  twoYearsAgo, 0, 1).getTime(),
    [threeYearsAgo]: new Date(threeYearsAgo, 0, 1).getTime(),
     [fourYearsAgo]: new Date( fourYearsAgo, 0, 1).getTime(),
     [fiveYearsAgo]: new Date( fiveYearsAgo, 0, 1).getTime(),
      [sixYearsAgo]: new Date(  sixYearsAgo, 0, 1).getTime(),
    [sevenYearsAgo]: new Date(sevenYearsAgo, 0, 1).getTime()
  }

  for (var i = 0; i < 53; i++)
    chartData.push(new Array(8).fill(0))

  // Gather the chart data from the invoice data
  invoiceDataSheet.getSheetValues(2, 3, invoiceDataSheet.getLastRow() - 1, 6).filter(date =>  date[0].getFullYear() >= sevenYearsAgo)
    .map(amount => chartData[Math.ceil((amount[0] - firstDayOfYear[amount[0].getFullYear()]) / millisecondsInWeek)][currentYear - amount[0].getFullYear()] += Number(amount[5]))

  // Convert the data into cumulative data
  chartData = chartData.map((amount, week, cumulativeData) => {

    if (week > 0)
    {
      lastWeek = week - 1;

      for (var i = 0; i < 8; i++)
        amount[i] += (amount[i] == 0 && i == 0) ? 0 : Number(cumulativeData[lastWeek][i]);
    }

    currentSales = (currentSales < amount[0]) ? amount[0] : currentSales;

    return ["Week " + (week + 1), ...amount]
  }).map(potentialSales => {potentialSales.push(currentSales + ioAmount, currentSales + boAmount); return potentialSales})

  const numRows = chartData.unshift(["", currentYear, lastYear, twoYearsAgo, threeYearsAgo, fourYearsAgo, fiveYearsAgo, sixYearsAgo, sevenYearsAgo, "I/O", "B/O"]);
  spreadsheet.getSheetByName('Chart Data').getRange(1, 1, numRows, chartData[0].length).setValues(chartData)
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
  
  return (enteredBy) ? [enteredBy[0], enteredBy[2]] : '';
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
  if (initials)
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
      case 'SG':
        return 'Shane';
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
      case 'NV':
        return 'Nathan';
      default:
        return initials;
    }
  }
  
  return '';
}

/**
 * This function converts Yes or No response from Order Entry regarding the status of the order into Back Order or not.
 * 
 * @param {String}     invNum      : The invoice number, if there is one.
 * @param {String} isOrderComplete : Yes or No depending on whether the order is complete.
 * @return {String} Returns what the back order status is.
 * @author Jarren Ralf
 */
function getInvoiceNumber(invNum, isOrderComplete)
{
  return (invNum === ' ') ? '' : (isOrderComplete === 'Yes') ? invNum : 'multiple';
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
 * @param {String}     ordNum    : The tab name of the imported excel spreadsheet (assumed to be the order number)
 * @param {Boolean} isInvoiceNum : Whether the order being imported is an invoice or not.
 * @param {Boolean} isCreditNum  : Whether the order being imported is a credit or not.
 * @returns {String} Returns the order number if it has been determined to be correct, or blank otherwise.
 * @auther Jarren Ralf
 */
function getOrderNumber(ordNum, isInvoiceNum, isCreditNum)
{
  if (isNumber(ordNum) && ordNum.toString().length === 5)
    return ordNum;
  else
  {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt((isInvoiceNum) ? (isCreditNum) ? 'Enter the credit number:' : 'Enter the invoice number:' : 'Enter the order number:',);
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
function getOrderStatus(isOrderComplete, isInvoicedOrders, invNum)
{
  return (isInvoicedOrders) ? (isOrderComplete !== 'No') ? 'Completed' : 'Partial' : (invNum !== ' ') ? 'Partial Order' : '';
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
  
  return (properTypesetName) ? properTypesetName[colSelector] : name;
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
 * This function checks if a given value is precisely a non-blank string.
 * 
 * @param  {String}  str : A given string.
 * @return {Boolean} Returns a boolean based on whether an inputted string is not-blank or not.
 * @author Jarren Ralf
 */
function isNotBlank(str)
{
  return str !== '';
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
 * This function detects when the user presses the delete key on a po or receipt number and it moves that item to the Non-Lodge column.
 * 
 * @param {Event Object} e : The event object.
 * @param {Sheet} sheet : The active sheet.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function manageDocumentNumbers(e, sheet, spreadsheet)
{
  const range = e.range;
  const row = range.rowStart;
  const col = range.columnStart; 
  const colEnd = range.columnEnd; 

  if (row > 1 && (col === 7 || col === 12) && (colEnd === 7 || colEnd === 12) && row === range.rowEnd)
  {
    const deletedDocumentNumber = e.oldValue;
    const lastRow = sheet.getLastRow() + 1
    const numRows = lastRow - row;
    const rowOffSet = sheet.getSheetValues(row, col + 2, numRows, 1).findIndex(docNum => isBlank(docNum[0]));

    if (rowOffSet === -1) // If there are no more rows at the bottom of the page, then add one and set the values to blanks initially
      sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).setValue('').setHorizontalAlignment('center')
        .offset(0, 0, 1, 2).setBorder(null, true, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK)
        .offset(0, 3, 1, 2).setBorder(null, true, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK)
        .offset(0, 3, 1, 3).setBorder(null, true, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK)
        .offset(0, 4, 1, 4).setBorder(null, true, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK)
        .offset(0, 5, 1, 3).setBorder(null, true, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK)
        .offset(0, 4, 1, 4).setBorder(null, true, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);

    if (col !== 12) // POs
    {
      const poNumbersRange = sheet.getRange(row, col, numRows, 1)
      const poNumbers = poNumbersRange.getValues().filter(poNum => !isBlank(poNum[0]))
      const numPos = poNumbers.length;

      if (numPos > 0)
        poNumbersRange.clearContent() // Necessary for making sure the po number at the bottom is not duplicated
          .offset(0, 0, numPos, 1).setValues(poNumbers)  // Shift the po numbers up 1
          .offset((rowOffSet !== -1) ? rowOffSet : numRows, 2, 1, 1).setValue(deletedDocumentNumber)
      else
        range.offset((rowOffSet !== -1) ? rowOffSet : numRows, 2, 1, 1).setValue(deletedDocumentNumber)
      
      spreadsheet.toast('Added to bottom of Non-Lodge PO #s', deletedDocumentNumber)
    }
    else // Receipts
    {
      const poAndReceiptNumbersRange = sheet.getRange(row, col - 1, numRows, 2)
      const poAndReceiptNumbers = poAndReceiptNumbersRange.getValues().filter(rctNum => !isBlank(rctNum[1]))
      const numReceipts = poAndReceiptNumbers.length

      if (numReceipts > 0)
        poAndReceiptNumbersRange.clearContent() // Necessary for making sure the receipt number at the bottom is not duplicated
          .offset(0, 0, poAndReceiptNumbers.length, 2).setValues(poAndReceiptNumbers) // Shift the po and receipt numbers up 1
          .offset((rowOffSet !== -1) ? rowOffSet : numRows, 3, 1, 1).setValue(deletedDocumentNumber) // Place the deleted document number at the bottom of the list
      else
        range.offset(0, -1).clearContent().offset((rowOffSet !== -1) ? rowOffSet : numRows, 3, 1, 1).setValue(deletedDocumentNumber) // Place the deleted document number at the bottom of the list
      
      spreadsheet.toast('Added to bottom of Non-Lodge Rct #s', deletedDocumentNumber)
    }
  }
}

/**
 * This function manages the price changes on the Lead and Bait Cost & Pricing sheets.
 * 
 * @param {Event Object} e : The event object.
 * @param {String} sheetName : The name of the sheet that was editted.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 */
function managePriceChange(e, sheetName, spreadsheet)
{
  const range = e.range;
  const row = range.rowStart;
  const col = range.columnStart;  

  if (row == range.rowEnd && col == range.columnEnd && row > 2) // Only look at a single cell edit
  {
    if (col === 8 || col === 9 || col === 10) // Cost is changing
    {
      const isLeadPricingSheet = sheetName !== 'Bait Cost & Pricing';
      const formattedDate = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(),"dd MMM yyyy")

      if (isLeadPricingSheet)
      {
        if (range.offset(0, 16 - col).getValue() != (Number(e.value) + Number(range.offset(0, 11 - col).getValue())).toFixed(2)) // If new cost is different that previous cost, display new cost and trigger the update reminder email to send
          range.offset(0, 7 - col).setValue(formattedDate).offset(0, 10).uncheck().offset(0, 11).setValue('Yes');
        else
          range.offset(0, 7 - col).setValue(formattedDate);
      }
      else
        range.offset(0, 7 - col).setValue(formattedDate).offset(0, 3).uncheck().offset(0, 9).setValue('Yes');
    }
    else if (range.isChecked())
    {
      if (sheetName === 'Bait Cost & Pricing')
      {
        if (col === 10)
          range.offset(0, 9).setValue('');
      }
      else if (sheetName === 'Lead Cost & Pricing' && col === 17)
        range.offset(0, 11).setValue('');
    }
  }
}

/**
 * This function moves the selected row from the Lodge or Guide order page to the completed page.
 * 
 * @param {Event Object} e : The event object.
 * @param {Sheet} sheet : The active sheet.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 */
function moveRow(e, sheet, spreadsheet)
{
  const range = e.range;
  const row = range.rowStart;
  const col = range.columnStart;  

  if (row == range.rowEnd && col == range.columnEnd) // Only look at a single cell edit
  {
    const sheetNames = sheet.getSheetName().split(" ") // Split the sheet name, which will be used to distinguish between Logde and Guide page

    if (sheetNames.pop() == "ORDERS") // An edit is happening on one of the Order pages
    {
      const numCols = sheet.getLastColumn()

      if (row > 2) // Not a header
      {
        if (col == numCols) // Order Status is changing
        {
          const value = e.value; 
          const numCols = sheet.getLastColumn()

          if (value == 'Updated')
            range.setValue('').offset(0, 5 - col).setValue('').offset(0, -4, 1, numCols).setBackground('#00ff00');
          else if (value == 'Picking')
          {
            const rowValues = sheet.getSheetValues(row, 1, 1, numCols)[0]; // Entire row values

            if (!rowValues[3]) // Order is not approved
            {
              const ui = SpreadsheetApp.getUi()
              ui.alert('Order NOT Approved', 'You have started picking an order that may not be approved by the customer yet.\n\nYou may want to check with ' + 
                rowValues[1] + ' before picking any items.', ui.ButtonSet.OK)
            }
          }
          else
          {
            const rowValues = sheet.getSheetValues(row, 1, 1, numCols)[0]; // Entire row values
            const timeZone = spreadsheet.getSpreadsheetTimeZone(); // Set the timezone

            rowValues[0] = Utilities.formatDate(rowValues[0], timeZone, 'MMM dd, yyyy'); // Set the format of the order date
            rowValues.push(Utilities.formatDate(     new Date(), timeZone, 'MMM dd, yyyy')); // Set the current time for the completion date

            if (value == "Completed") // The order status is being set to complete 
            {
              rowValues[3] = true;
              spreadsheet.getSheetByName(sheetNames.pop() +  " COMPLETED").appendRow(rowValues) // Move the row of values to the completed page
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
              rowValues[3] = true;
              spreadsheet.getSheetByName(sheetNames.pop() +  " COMPLETED").appendRow(rowValues); // Move the row of values to the completed page
              sheet.getRange(row, 11, 1, 4).setValues([['multiple', '', '',  'Partial Order']]).offset(0, -7, 1, 1).check(); // Clear the invoice values, and set the status
              deleteBackOrderedItems(rowValues[2], spreadsheet);
            }
          }
        }
        else if (col == 5) // Adding a Printed By name
        {
          if (isNotBlank(range.getValue()))
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
        else if (col == 4) // The approval of an order has changed
        {
          if (isBlank(range.offset(0, 7).getValue())) // Is this order an inital order?
          {
            const approval = range.isChecked();
            const approvedOrderNumber = range.offset(0, -1).getValue();
            const ioSheet = spreadsheet.getSheetByName('I/O')
            ioSheet.getSheetValues(3, 11, ioSheet.getLastRow() - 2, 1).map((ordNum, o) => (ordNum[0] == approvedOrderNumber) ? ioSheet.getRange(o + 3, 4).setValue(approval) : null);
          }
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
 * This function sends an email to me whenever there are costs in Access that need to be updated.
 * 
 * @author Jarren
 */
function sendEmail(url, sheetName)
{
  MailApp.sendEmail({
    to: "jarren@pacificnetandtwine.com",
    subject: "Some Costs in Access Need Updating",
    htmlBody: '<a href="' + url + '">' + sheetName + ' has changed.</a>'
  });
}

/**
 * This function takes the users selection on the Orders sheet and it sents an email to the appropriate employees asking if there order/s are approved or not.
 * 
 * @author Jarren Ralf
 */
function sendAnEmailToSelectedPeopleAskingIfOrderIsApproved()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const activeSheet = spreadsheet.getActiveSheet();
  const sheetName = activeSheet.getSheetName().split(" ");
  const numCols = activeSheet.getLastColumn() - 4; // Only include up to the Notes column
  const timeZone = spreadsheet.getSpreadsheetTimeZone()
  // const emails = {   'Brent': 'brent@pacificnetandtwine.com', 
  //                  'Derrick': 'dmizuyabu@pacificnetandtwine.com',
  //                    'Deryk': 'deryk@pacificnetandtwine.com',
  //                   'Jarren': 'jarren@pacificnetandtwine.com, scottnakashima@hotmail.com, deryk@pacificnetandtwine.com',
  //                     'Kris': 'kris@pacificnetandtwine.com',
  //                   'Nathan': 'nathan@pacificnetandtwine.com',
  //                     'Noah': 'noah@pacificnetandtwine.com',
  //                 'Scarlett': '',
  //                    'Scott': 'scott@pacificnetandtwine.com',
  //                    'Endis': 'triteswarehouse@pacificnetandtwine.com',
  //                     'Eryn': 'eryn@pacificnetandtwine.com'}
  const emails = {   'Brent': 'lb_blitz_allstar@hotmail.com', 
                   'Derrick': 'lb_blitz_allstar@hotmail.com',
                     'Deryk': 'lb_blitz_allstar@hotmail.com',
                    'Jarren': 'lb_blitz_allstar@hotmail.com',
                      'Kris': 'lb_blitz_allstar@hotmail.com',
                    'Nathan': 'lb_blitz_allstar@hotmail.com',
                      'Noah': 'lb_blitz_allstar@hotmail.com',
                  'Scarlett': '',
                     'Scott': 'lb_blitz_allstar@hotmail.com',
                     'Endis': 'lb_blitz_allstar@hotmail.com',
                      'Eryn': 'lb_blitz_allstar@hotmail.com'}
  var idx = -1, emailRecipients = [], emailBodies = [], backgroundColours = [], col, numRows;

  if (sheetName.pop() === 'ORDERS')
  {
    spreadsheet.getActiveRangeList().getRanges().map(rng => {

      col = 1 - rng.getColumn();
      numRows = rng.getNumRows();

      rng.offset(0, col, numRows, numCols).getValues().map(order => {
        if (order[3]) // If this order is already approved, then ignore it
          return null;
        else // Order is not approved
        {
          if (emailRecipients.includes(order[1])) // If there is more than 1 order for a particular employee to determine the approval status of
          {
            idx = emailRecipients.indexOf(order[1])
            emailBodies[idx].push(order);
            backgroundColours[idx].push(rng.offset(0, col, numRows, numCols).getBackgrounds())
          }
          else // This is the first instance of an employee who has an unapproved order
          {
            idx = emailRecipients.push(order[1]) - 1;
            emailBodies[idx] = [order];
            backgroundColours[idx] = [rng.offset(0, col, numRows, numCols).getBackgrounds()];
          }
        }
      })
    })

    var htmlTemplate, htmlOutput;

    Logger.log('emailRecipients:')
    Logger.log(emailRecipients)
    Logger.log('emailBodies:')
    Logger.log(emailBodies)
  
    emailRecipients.map((recipient, r) => {

      htmlTemplate = HtmlService.createTemplateFromFile('getApprovalStatusEmail')
      htmlTemplate.customerType = toProper(sheetName[0])
      htmlOutput = htmlTemplate.evaluate();

      for (var i = 0; i < emailBodies[r].length; i++)
        htmlOutput.append(
          '<tr style="height: 20px">'+
          '<td class="s5" dir="ltr" style="background-color:' + backgroundColours[r][i][0] + '">' + 
            Utilities.formatDate(emailBodies[r][i][0], timeZone, "dd MMM yyyy") + '</td>' +
          '<td class="s6" dir="ltr">' + emailBodies[r][i][1] + '</td>'+
          '<td class="s7" dir="ltr">' + emailBodies[r][i][2] + '</td>'+
          '<td class="s8" dir="ltr">' + emailBodies[r][i][3] + '</td>'+
          '<td class="s6" dir="ltr">' + emailBodies[r][i][4] + '</td>'+
          '<td class="s6" dir="ltr">' + emailBodies[r][i][5] + '</td>'+
          '<td class="s6" dir="ltr">' + emailBodies[r][i][6] + '</td>'+
          '<td class="s6" dir="ltr">' + emailBodies[r][i][7] + '</td>'+
          '<td class="s6" dir="ltr">' + emailBodies[r][i][8] + '</td>'+
          '<td class="s6" dir="ltr">' + emailBodies[r][i][9] + '</td></tr>'
        )

      htmlOutput.append('</tbody></table></div>')

      Logger.log(htmlOutput)

      MailApp.sendEmail({
        to: emails[recipient],
        //cc: "jarren@pacificnetandtwine.com, scott@pacificnetandtwine.com, deryk@pacificnetandtwine.com",
        replyTo: "jarren@pacificnetandtwine.com, scott@pacificnetandtwine.com, deryk@pacificnetandtwine.com",
        subject: "Are the following orders APPROVED?",
        htmlBody: htmlOutput.getContent(),
      })
    })
  }
  else
    Browser.msgBox('You must be on an order page to run this function.')
}

/**
 * This function sets the column widths of 4 of the sheets on this spreadsheet, namely the Order and Completed pages.
 * 
 * @author Jarren Ralf
 */
function setColumnWidths()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const lodgeOrdersSheet = spreadsheet.getSheetByName('LODGE ORDERS');
  const guideOrdersSheet = spreadsheet.getSheetByName('GUIDE ORDERS');
  const lodgeCompletedSheet = spreadsheet.getSheetByName('LODGE COMPLETED');
  const guideCompletedSheet = spreadsheet.getSheetByName('GUIDE COMPLETED');
  const cancelledSheet = spreadsheet.getSheetByName('CANCELLED');
  const widths = [80, 80, 68, 54, 80, 61, 200, 84, 150, 150, 500, 63, 79, 70, 104];
  const numCols = widths.length;

  for (var c = 1; c < numCols; c++)
  {
       lodgeOrdersSheet.setColumnWidth(c, widths[c]);
       guideOrdersSheet.setColumnWidth(c, widths[c]);
    lodgeCompletedSheet.setColumnWidth(c, widths[c]);
    guideCompletedSheet.setColumnWidth(c, widths[c]);
         cancelledSheet.setColumnWidth(c, widths[c]);
  }

  const lastColumnWidth = widths.shift()
  lodgeCompletedSheet.setColumnWidth(c, lastColumnWidth);
  guideCompletedSheet.setColumnWidth(c, lastColumnWidth);
       cancelledSheet.setColumnWidth(c, lastColumnWidth);
}

/**
 * This function takes all of the order numbers on the LODGE ORDERS and GUIDE ORDERS sheets and it hyperlinks them to the corresponding
 * set of items that are either on the BO sheet or the IO sheet. In addition, this function takes all of the invoice numbers on the LODGE COMPLETED and GUIDE COMPLETED sheets
 * and it hyperlinks them so that they link to the corresponding set of items that are on the Inv'd sheet. This function also takes all of the PO numbers on the I/O and B/O pages
 * and it hyperlinks them so that they link to the corresponding set of items that are on the PO sheet.
 * 
 * @param {Sheet}  lodgeOrdersSheet : The LODGE ORDERS sheet.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @return {Sheet[]} Returns the Guide Orders, Lodge Completed, and Guide Completed sheets.
 * @author Jarren Ralf 
 */
function setItemLinks(lodgeOrdersSheet, spreadsheet)
{
  spreadsheet.toast('Order and Invoice # hyperlinks being established...', '', -1)
  const guideOrdersSheet = spreadsheet.getSheetByName('GUIDE ORDERS');
  const lodgeCompletedSheet = spreadsheet.getSheetByName('LODGE COMPLETED');
  const guideCompletedSheet = spreadsheet.getSheetByName('GUIDE COMPLETED');

  [boSheet, ioSheet, poSheet, invdSheet] = establishItemLinks_IO_BO(spreadsheet, lodgeOrdersSheet, guideOrdersSheet)
  establishItemLinks_INVD(invdSheet, lodgeCompletedSheet, guideCompletedSheet)
  establishItemLinks_PO(spreadsheet, poSheet, ioSheet, boSheet)
  spreadsheet.toast('Transfer sheet hyperlinks being established...', 'Order and Invoice # hyperlinks completed.', -1)

  return [guideOrdersSheet, lodgeCompletedSheet, guideCompletedSheet]
}

/**
 * This function takes all of the hyper links to the transfer sheets and it updates the urls for the transfer sheets.
 * 
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @param {Sheet[]} sheets : An array of sheets, assumed to be the two order sheets and completed sheets.
 * @author Jarren Ralf 
 */
function setTransferSheetLinks(spreadsheet, ...sheets)
{
  var numRows, notesRange, orderNumbers, invoiceNumbers, note, noteLength = 0, numLinksToParksille = 1, numLinksToRupert = 1, numLinksFromParksille = 1, 
    numLinksFromRupert = 1, rangeLink = [], richTextBuilder, noteSplit, locations, ss, url, itemsToRichmondSheet, orderSheet, shippedSheet, receivedSheet, gid;

  sheets.map(sheet => {

    numRows = sheet.getLastRow() - 2;

    if (numRows > 0)
    {
      if (sheet.getSheetName().charAt(1) !== '/') // Lodge Orders, Guide Order, Lodge Completed, or Guide Completed
      {
        notesRange = sheet.getRange(3, 10, numRows, 1)
        orderNumbers = sheet.getSheetValues(3, 3, numRows, 1)
        invoiceNumbers = sheet.getSheetValues(3, 11, numRows, 1)
        
        notes = notesRange.getRichTextValues().map((richText, ordNum) => {
          note = richText[0].getText();
          
          if (note.includes("Shipping from"))
          {
            numLinksToParksille = 1, numLinksToRupert = 1, numLinksFromParksille = 1, numLinksFromRupert = 1, rangeLink.length = 0, noteLength = 0;
            richTextBuilder = SpreadsheetApp.newRichTextValue().setText(note);
            noteSplit = note.split('\n');

            richText = noteSplit.map(run => {

              if (run.includes("Shipping from "))
              {
                locations = run.split("Shipping from ").pop().split(' to ');

                switch (locations[0]) // From Location
                {
                  case "Richmond": 

                    switch (locations[1]) // To Location
                    {
                      case "Parksville":
                        ss = SpreadsheetApp.openById('181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM')

                        receivedSheet = ss.getSheetByName('Received');
                        url = ss.getUrl();
                        gid = receivedSheet.getSheetId()

                        receivedSheet.getSheetValues(4, 5, receivedSheet.getLastRow() - 3, 1)
                          .map((description, row) => (description[0].includes('Inv# ' + invoiceNumbers[ordNum][0])) ? rangeLink.push(url + '?gid=' + gid + '#gid=' + gid + '&range=E' + (row + 4)) : null)

                        if (rangeLink.length > 0) // Atleast 1 invoice number was found on the Received sheet
                        {
                          for (var i = 1; i < numLinksToParksille; i++)
                            rangeLink.pop()

                          richTextBuilder.setLinkUrl(noteLength, noteLength + run.length, rangeLink.pop())
                        }
                        else
                        {
                          receivedSheet.getSheetValues(4, 5, receivedSheet.getLastRow() - 3, 1)
                            .map((description, row) => (description[0].includes('Order# ' + orderNumbers[ordNum][0])) ? rangeLink.push(url + '?gid=' + gid + '#gid=' + gid + '&range=E' + (row + 4)) : null)

                          if (rangeLink.length > 0) // Atleast 1 order number was found on the Received sheet
                          {
                            for (var i = 1; i < numLinksToParksille; i++)
                              rangeLink.pop()

                            richTextBuilder.setLinkUrl(noteLength, noteLength + run.length, rangeLink.pop())
                          }
                          else
                          {
                            shippedSheet = ss.getSheetByName('Shipped');
                            gid = shippedSheet.getSheetId()

                            shippedSheet.getSheetValues(4, 5, shippedSheet.getLastRow() - 3, 1)
                              .map((description, row) => (description[0].includes('Inv# ' + invoiceNumbers[ordNum][0])) ? rangeLink.push(url + '?gid=' + gid + '#gid=' + gid + '&range=E' + (row + 4)) : null)

                            if (rangeLink.length > 0) // Atleast 1 invoice number was found on the Shipped sheet
                            {
                              for (var i = 1; i < numLinksToParksille; i++)
                                rangeLink.pop()

                              richTextBuilder.setLinkUrl(noteLength, noteLength + run.length, rangeLink.pop())
                            }
                            else
                            {
                              shippedSheet.getSheetValues(4, 5, shippedSheet.getLastRow() - 3, 1)
                                .map((description, row) => (description[0].includes('Order# ' + orderNumbers[ordNum][0])) ? rangeLink.push(url + '?gid=' + gid + '#gid=' + gid + '&range=E' + (row + 4)) : null)

                              if (rangeLink.length > 0) // Atleast 1 order number was found on the Shipped sheet
                              {
                                for (var i = 1; i < numLinksToParksille; i++)
                                  rangeLink.pop()

                                richTextBuilder.setLinkUrl(noteLength, noteLength + run.length, rangeLink.pop())
                              }
                              else
                              {
                                orderSheet = ss.getSheetByName('Order');
                                gid = orderSheet.getSheetId()

                                orderSheet.getSheetValues(4, 5, orderSheet.getLastRow() - 3, 1)
                                  .map((description, row) => (description[0].includes('Inv# ' + invoiceNumbers[ordNum][0])) ? rangeLink.push(url + '?gid=' + gid + '#gid=' + gid + '&range=B' + (row + 4)) : null)

                                if (rangeLink.length > 0) // Atleast 1 invoice number was found on the Order sheet
                                {
                                  for (var i = 1; i < numLinksToParksille; i++)
                                    rangeLink.pop()

                                  richTextBuilder.setLinkUrl(noteLength, noteLength + run.length, rangeLink.pop())
                                }
                                else
                                {
                                  orderSheet.getSheetValues(4, 5, orderSheet.getLastRow() - 3, 1)
                                    .map((description, row) => (description[0].includes('Order# ' + invoiceNumbers[ordNum][0])) ? rangeLink.push(url + '?gid=' + gid + '#gid=' + gid + '&range=B' + (row + 4)) : null)

                                  if (rangeLink.length > 0) // Atleast 1 order number was found on the Order sheet
                                  {
                                    for (var i = 1; i < numLinksToParksille; i++)
                                      rangeLink.pop()

                                    richTextBuilder.setLinkUrl(noteLength, noteLength + run.length, rangeLink.pop())
                                  } 
                                }
                              }
                            }
                          }
                        }

                        numLinksToParksille++;
                        break;
                      case "Rupert":
                        ss = SpreadsheetApp.openById('1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c')

                        receivedSheet = ss.getSheetByName('Received');
                        url = ss.getUrl();
                        gid = receivedSheet.getSheetId()

                        receivedSheet.getSheetValues(4, 5, receivedSheet.getLastRow() - 3, 1)
                          .map((description, row) => (description[0].includes('Inv# ' + invoiceNumbers[ordNum][0])) ? rangeLink.push(url + '?gid=' + gid + '#gid=' + gid + '&range=E' + (row + 4)) : null)

                        if (rangeLink.length > 0) // Atleast 1 invoice number was found on the Received sheet
                        {
                          for (var i = 1; i < numLinksToRupert; i++)
                            rangeLink.pop()

                          richTextBuilder.setLinkUrl(noteLength, noteLength + run.length, rangeLink.pop())
                        }
                        else
                        {
                          receivedSheet.getSheetValues(4, 5, receivedSheet.getLastRow() - 3, 1)
                            .map((description, row) => (description[0].includes('Order# ' + orderNumbers[ordNum][0])) ? rangeLink.push(url + '?gid=' + gid + '#gid=' + gid + '&range=E' + (row + 4)) : null)

                          if (rangeLink.length > 0) // Atleast 1 order number was found on the Received sheet
                          {
                            for (var i = 1; i < numLinksToRupert; i++)
                              rangeLink.pop()

                            richTextBuilder.setLinkUrl(noteLength, noteLength + run.length, rangeLink.pop())
                          }
                          else
                          {
                            shippedSheet = ss.getSheetByName('Shipped');
                            gid = shippedSheet.getSheetId()

                            shippedSheet.getSheetValues(4, 5, shippedSheet.getLastRow() - 3, 1)
                              .map((description, row) => (description[0].includes('Inv# ' + invoiceNumbers[ordNum][0])) ? rangeLink.push(url + '?gid=' + gid + '#gid=' + gid + '&range=E' + (row + 4)) : null)

                            if (rangeLink.length > 0) // Atleast 1 invoice number was found on the Shipped sheet
                            {
                              for (var i = 1; i < numLinksToRupert; i++)
                                rangeLink.pop()

                              richTextBuilder.setLinkUrl(noteLength, noteLength + run.length, rangeLink.pop())
                            }
                            else
                            {
                              shippedSheet.getSheetValues(4, 5, shippedSheet.getLastRow() - 3, 1)
                                .map((description, row) => (description[0].includes('Order# ' + orderNumbers[ordNum][0])) ? rangeLink.push(url + '?gid=' + gid + '#gid=' + gid + '&range=E' + (row + 4)) : null)

                              if (rangeLink.length > 0) // Atleast 1 order number was found on the Shipped sheet
                              {
                                for (var i = 1; i < numLinksToRupert; i++)
                                  rangeLink.pop()

                                richTextBuilder.setLinkUrl(noteLength, noteLength + run.length, rangeLink.pop())
                              }
                              else
                              {
                                orderSheet = ss.getSheetByName('Order');
                                gid = orderSheet.getSheetId()

                                orderSheet.getSheetValues(4, 5, orderSheet.getLastRow() - 3, 1)
                                  .map((description, row) => (description[0].includes('Inv# ' + invoiceNumbers[ordNum][0])) ? rangeLink.push(url + '?gid=' + gid + '#gid=' + gid + '&range=B' + (row + 4)) : null)

                                if (rangeLink.length > 0) // Atleast 1 invoice number was found on the Order sheet
                                {
                                  for (var i = 1; i < numLinksToRupert; i++)
                                    rangeLink.pop()

                                  richTextBuilder.setLinkUrl(noteLength, noteLength + run.length, rangeLink.pop())
                                }
                                else
                                {
                                  orderSheet.getSheetValues(4, 5, orderSheet.getLastRow() - 3, 1)
                                    .map((description, row) => (description[0].includes('Order# ' + invoiceNumbers[ordNum][0])) ? rangeLink.push(url + '?gid=' + gid + '#gid=' + gid + '&range=B' + (row + 4)) : null)

                                  if (rangeLink.length > 0) // Atleast 1 order number was found on the Order sheet
                                  {
                                    for (var i = 1; i < numLinksToRupert; i++)
                                      rangeLink.pop()

                                    richTextBuilder.setLinkUrl(noteLength, noteLength + run.length, rangeLink.pop())
                                  } 
                                }
                              }
                            }
                          }
                        }

                        numLinksToRupert++;
                        break;
                    }
                    break;
                  case "Parksville":
                    ss = SpreadsheetApp.openById('181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM')
                    itemsToRichmondSheet = ss.getSheetByName('ItemsToRichmond');
                    gid = itemsToRichmondSheet.getSheetId();

                    itemsToRichmondSheet.getSheetValues(4, 4, itemsToRichmondSheet.getLastRow() - 3, 1)
                      .map((description, row) => (description[0].includes('Inv# ' + invoiceNumbers[ordNum][0])) ? rangeLink.push(url + '?gid=' + gid + '#gid=' + gid + '&range=B' + (row + 4)) : null)

                    if (rangeLink.length > 0) // Atleast 1 invoice number was found on the Transfer sheet
                    {
                      for (var i = 1; i < numLinksFromParksille; i++)
                        rangeLink.pop()

                      richTextBuilder.setLinkUrl(noteLength, noteLength + run.length, url + '?gid=' + gid + '#gid=' + gid + rangeLink.pop())
                    }
                    else 
                    {
                      itemsToRichmondSheet.getSheetValues(4, 4, itemsToRichmondSheet.getLastRow() - 3, 1)
                        .map((description, row) => (description[0].includes('Order# ' + orderNumbers[ordNum][0])) ? rangeLink.push(url + '?gid=' + gid + '#gid=' + gid + '&range=B' + (row + 4)) : null)

                      if (rangeLink.length > 0) // Atleast 1 order number was found on the Transfer sheet
                      {
                        for (var i = 1; i < numLinksFromParksille; i++)
                          rangeLink.pop()

                        richTextBuilder.setLinkUrl(noteLength, noteLength + run.length, url + '?gid=' + gid + '#gid=' + gid + rangeLink.pop())
                      }
                    }

                    numLinksFromParksille++;
                    break;
                  case "Rupert":
                    ss = SpreadsheetApp.openById('1IEJfA5x7sf54HBMpCz3TAosJup4TrjXdUOqm4KK3t9c')
                    itemsToRichmondSheet = ss.getSheetByName('ItemsToRichmond');
                    gid = itemsToRichmondSheet.getSheetId();

                    itemsToRichmondSheet.getSheetValues(4, 4, itemsToRichmondSheet.getLastRow() - 3, 1)
                      .map((description, row) => (description[0].includes('Inv# ' + invoiceNumbers[ordNum][0])) ? rangeLink.push(url + '?gid=' + gid + '#gid=' + gid + '&range=B' + (row + 4)) : null)

                    if (rangeLink.length > 0) // Atleast 1 invoice number was found on the Transfer sheet
                    {
                      for (var i = 1; i < numLinksFromRupert; i++)
                        rangeLink.pop()

                      richTextBuilder.setLinkUrl(noteLength, noteLength + run.length, url + '?gid=' + gid + '#gid=' + gid + rangeLink.pop())
                    }
                    else 
                    {
                      itemsToRichmondSheet.getSheetValues(4, 4, itemsToRichmondSheet.getLastRow() - 3, 1)
                        .map((description, row) => (description[0].includes('Order# ' + orderNumbers[ordNum][0])) ? rangeLink.push(url + '?gid=' + gid + '#gid=' + gid + '&range=B' + (row + 4)) : null)

                      if (rangeLink.length > 0) // Atleast 1 order number was found on the Transfer sheet
                      {
                        for (var i = 1; i < numLinksFromRupert; i++)
                          rangeLink.pop()

                        richTextBuilder.setLinkUrl(noteLength, noteLength + run.length, url + '?gid=' + gid + '#gid=' + gid + rangeLink.pop())
                      }
                    }

                    numLinksFromRupert++;
                    break;
                }
              }

              noteLength += run.length + 1;
            })
            
            return [richTextBuilder.build()];
          }

          return richText
        })

        notesRange.setRichTextValues(notes);
      }
      else // B/O or I/O
      {
        Logger.log('Need to create an updater Script for I/O and B/O')
      }
    }
  })

  spreadsheet.toast('', 'Order #, Invoice #, and Transfer sheet hyperlinks completed.', -1)
}

/**
 * This function takes the given string and makes sure that each word in the string has a capitalized 
 * first letter followed by lower case.
 * 
 * @param {String} str : The given string
 * @return {String} The output string with proper case
 * @author Jarren Ralf
 */
function toProper(str)
{
  return capitalizeSubstrings(capitalizeSubstrings(str, '-'), ' ');
}

/**
 * This function creates all of the triggers for this spreadsheet.
 * 
 * @author Jarren Ralf
 */
function triggers_CreateAll()
{
  const spreadsheet = SpreadsheetApp.getActive()
  ScriptApp.newTrigger('onChange').forSpreadsheet(spreadsheet). onChange().create();
  ScriptApp.newTrigger('installedOnOpen').forSpreadsheet(spreadsheet).onOpen().create();
  ScriptApp.newTrigger('getChartData').timeBased().atHour(9).everyDays(1).create();
  ScriptApp.newTrigger('setColumnWidths').timeBased().atHour(7).everyDays(1).create();
  ScriptApp.newTrigger('updatedPntReceivingSpreadsheet').timeBased().atHour(20).everyDays(1).create(); 
  ScriptApp.newTrigger('updatePriceAndCostOfLeadAndFrozenBait').timeBased().atHour(7).everyDays(1).create();
  ScriptApp.newTrigger('emailCostChangeOfLeadOrFrozenBait').timeBased().atHour(15).everyDays(1).create();
  spreadsheet.getSheetByName('Triggers').getRange(1, 1).check();
}

/**
 * This function deletes all of the triggers for this spreadsheet.
 * 
 * @author Jarren Ralf
 */
function triggers_DeleteAll()
{
  ScriptApp.getProjectTriggers().map(trigger => ScriptApp.deleteTrigger(trigger));
  SpreadsheetApp.getActive().getSheetByName('Triggers').getRange(1, 1).uncheck();
}

/**
 * This function handles the import of an Invoice (from Adagio OrderEntry) that contains items that have already been billed and presumably shipped out.
 * 
 * @param {String[][]}     items    : A list of items on the invoice that was imported.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @param {String}        invNum    : The invoice number that is being imported.
 * @author Jarren Ralf
 */
function updateInvoicedItemsOnTracker(items, spreadsheet, invNum, isCreditedItems)
{
  items.pop(); // Remove the "Total" or final line
  
  // Get all the indexes of the relevant headers
  const headerOE = items.shift();
  const isInvoicedHistory = items[0][headerOE.indexOf('Cust #')] === ' ';
  const returnedQtyIdx = headerOE.indexOf('Quantity');
  const orderedQtyIdx = headerOE.indexOf('Ordered');
  const shippedQtyIdx = headerOE.indexOf('Shipped'); 
  const backOrderQtyIdx = headerOE.indexOf('Backorder'); 
  const skuIdx = headerOE.indexOf('Item');
  const descriptionIdx = headerOE.indexOf('Description');
  const unitPriceIdx = isInvoicedHistory && headerOE.indexOf('Unit Price') || 
                       isCreditedItems   && headerOE.indexOf('Unit Price') || headerOE.indexOf('Display Price');
  const extendedunitPriceIdx = isInvoicedHistory && headerOE.indexOf('Extension') ||
                               isCreditedItems   && headerOE.indexOf('Extension') || headerOE.indexOf('Display Ext Price');
  const locationName = getLocationName(items[0][headerOE.indexOf('Loc')]);
  const completedOrdersSheet_Lodge = spreadsheet.getSheetByName('LODGE COMPLETED')
  const completedOrdersSheet_Guide = spreadsheet.getSheetByName('GUIDE COMPLETED')
  const numCols_CompletedSheet = completedOrdersSheet_Lodge.getLastColumn();
  const numCompletedLodgeOrders = completedOrdersSheet_Lodge.getLastRow() - 2;
  const numCompletedGuideOrders = completedOrdersSheet_Guide.getLastRow() - 2;

  const invoicedItemSheet = spreadsheet.getSheetByName("Inv'd").activate(); 
  const numCurrentItems = invoicedItemSheet.getLastRow() - 2;
  invoicedItemSheet?.getFilter()?.remove(); // Remove the filter

  if (isCreditedItems)
  {
    const creditNumber = getOrderNumber(invNum, true, isCreditedItems);
    Logger.log('Credit Number: ' + creditNumber)

    const completedOrder = (numCompletedGuideOrders === 0) ? 
      completedOrdersSheet_Lodge.getSheetValues(3, 1, numCompletedLodgeOrders, numCols_CompletedSheet)
        .find(credtNum => credtNum[9].includes(creditNumber))
      : (numCompletedGuideOrders === 0) ? 
        completedOrdersSheet_Guide.getSheetValues(3, 1, numCompletedGuideOrders, numCols_CompletedSheet)
          .find(credtNum => credtNum[9].includes(creditNumber))
        : completedOrdersSheet_Lodge.getSheetValues(3, 1, numCompletedLodgeOrders, numCols_CompletedSheet)
          .concat(completedOrdersSheet_Guide.getSheetValues(3, 1, numCompletedGuideOrders, numCols_CompletedSheet))
            .find(credtNum => credtNum[9].includes(creditNumber))
    
    const invoiceDate   = completedOrder.pop();
    const customerName  = completedOrder[ 5];
    const orderNumber   = completedOrder[ 2];
    const invoiceNumber = completedOrder[10];

    var newItems = items.filter(returnedQty => returnedQty[returnedQtyIdx] != 0).map(item => 
      [invoiceDate, customerName, '', -1*item[returnedQtyIdx], '', 
      removeDashesFromSku(item[skuIdx]), item[descriptionIdx], item[unitPriceIdx], item[extendedunitPriceIdx], 
      locationName , orderNumber, invoiceNumber, creditNumber])
  }
  else
  {
    const invoiceNumber = getOrderNumber(invNum, true);
    Logger.log('Invoice Number: ' + invoiceNumber)

    const completedOrder = (numCompletedGuideOrders === 0) ? 
      completedOrdersSheet_Lodge.getSheetValues(3, 1, numCompletedLodgeOrders, numCols_CompletedSheet)
        .find(invNum => invNum[10] == invoiceNumber) 
      : (numCompletedGuideOrders === 0) ? 
        completedOrdersSheet_Guide.getSheetValues(3, 1, numCompletedGuideOrders, numCols_CompletedSheet)
          .find(invNum => invNum[10] == invoiceNumber) 
        : completedOrdersSheet_Lodge.getSheetValues(3, 1, numCompletedLodgeOrders, numCols_CompletedSheet)
          .concat(completedOrdersSheet_Guide.getSheetValues(3, 1, numCompletedGuideOrders, numCols_CompletedSheet))
            .find(invNum => invNum[10] == invoiceNumber)
    
    const invoiceDate  = completedOrder.pop();
    const customerName = completedOrder[5];
    const orderNumber  = completedOrder[2];

    var newItems = items.filter(shippedQty => shippedQty[shippedQtyIdx] != 0).map(item =>  // Remove the items that weren't shipped
      [invoiceDate, customerName, item[orderedQtyIdx], item[shippedQtyIdx], item[backOrderQtyIdx], 
      removeDashesFromSku(item[skuIdx]), item[descriptionIdx], item[unitPriceIdx], item[extendedunitPriceIdx], 
      locationName , orderNumber, invoiceNumber, ''])
  }

  const numNewItems = newItems.length;
  const numCols = newItems[0].length;
  const invoiceNumCol = invoicedItemSheet.getLastColumn() - 1;

  if (numCurrentItems > 0)
    invoicedItemSheet.getRange(numCurrentItems + 3, 1, numNewItems, numCols)
        .setNumberFormats(new Array(numNewItems).fill(['MMM dd, yyyy', '@', '#', '#', '#', '@', '@', '$#,##0.00', '$#,##0.00', '@', '@', '@', '@'])).setValues(newItems)
      .offset(-1*numCurrentItems, 0, numCurrentItems + numNewItems, numCols).sort([{column: invoiceNumCol, ascending: true}]);
  else
    invoicedItemSheet.getRange(3, 1, numNewItems, numCols).setNumberFormats(new Array(numNewItems).fill(['MMM dd, yyyy', '@', '#', '#', '#', '@', '@', '$#,##0.00', '$#,##0.00', '@', '@', '@', '@'])).setValues(newItems)

  Logger.log("The following new invoiced items were added to the Inv'd tab:")
  Logger.log(newItems)

  spreadsheet.toast(numNewItems + ' Added ', "Inv'd Items Imported", 60)

  SpreadsheetApp.flush()
  invoicedItemSheet.getRange(2, 1, invoicedItemSheet.getLastRow() - 1, invoiceNumCol + 1).createFilter(); // Create a filter in the header
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
  const shippedQtyIdx = headerOE.indexOf('Qty Shipped'); 
  const backOrderQtyIdx = headerOE.indexOf('Backorder'); 
  const skuIdx = headerOE.indexOf('Item');
  const descriptionIdx = headerOE.indexOf('Description');
  const unitPriceIdx = headerOE.indexOf('Unit Price');
  const locationIdx = headerOE.indexOf('Loc');
  const isItemCompleteIdx = headerOE.indexOf('Complete?');
  const months = {'01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun', '07': 'Jul', '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'};
  const orderNumber = getOrderNumber(ordNum, false);

  const lodgeCustomerSheet = spreadsheet.getSheetByName('Lodge Customer List');
  const charterGuideCustomerSheet = spreadsheet.getSheetByName('Charter & Guide Customer List');
  const lodgeOrdersSheet = spreadsheet.getSheetByName('LODGE ORDERS');
  const guideOrdersSheet = spreadsheet.getSheetByName('GUIDE ORDERS');
  const partialOrdersSheet = spreadsheet.getSheetByName('Item Management (Jarren Only ;)');

  const numLodgeOrders = lodgeOrdersSheet.getLastRow() - 2;
  const numGuideOrders = guideOrdersSheet.getLastRow() - 2;

  const enteredByNamesAndApprovalStatus = (numGuideOrders === 0) ? lodgeOrdersSheet.getSheetValues(3, 2, numLodgeOrders, 3) : 
                                          (numLodgeOrders === 0) ? guideOrdersSheet.getSheetValues(3, 2, numGuideOrders, 3) : 
                                          lodgeOrdersSheet.getSheetValues(3, 2, numLodgeOrders, 3).concat(guideOrdersSheet.getSheetValues(3, 2, numGuideOrders, 3));

  const customerNames = lodgeCustomerSheet.getSheetValues(3, 1, lodgeCustomerSheet.getLastRow() - 2, 3).concat(charterGuideCustomerSheet.getSheetValues(3, 1, charterGuideCustomerSheet.getLastRow() - 2, 3))
  const orderNumbers_BO = partialOrdersSheet.getSheetValues(2, 1, partialOrdersSheet.getRange(partialOrdersSheet.getLastRow(), 1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow() - 1, 1).flat()

  const orderDate = getDateString(items[0][dateIdx], months);
  const enteredByAndApproval = getEnteredByNameAndApprovalStatus(orderNumber, enteredByNamesAndApprovalStatus);
  const customerName = getProperTypesetName(items[0][customerNumIdx], customerNames, 2);
  const locationName = getLocationName(items[0][locationIdx])
  const reuploadSheet = spreadsheet.getSheetByName('Reupload:' + orderNumber);
  const itemSheet = (doesOrderContainBOs(orderNumber, orderNumbers_BO)) ? spreadsheet.getSheetByName('B/O').activate() : spreadsheet.getSheetByName('I/O').activate();
  const numRows = itemSheet.getLastRow() - 2;
  var noteValues, qty;
  itemSheet?.getFilter()?.remove(); // Remove the filter

  if (reuploadSheet)
  {
    const reuploadNotes = reuploadSheet.getSheetValues(2, 1, reuploadSheet.getLastRow() - 1, 5);
    const ordNum = itemSheet.getSheetValues(2, 1, 1, 14).flat().indexOf('Order #');
    var currentItems = itemSheet.getSheetValues(3, 1, numRows, itemSheet.getLastColumn()).filter(item => isBlank(item[ordNum]) || !isCurrentOrder)

    var newItems = (doesOrderContainBOs(orderNumber, orderNumbers_BO)) ? 
                      items.filter(item => item[isItemCompleteIdx])
                        .filter(item => item[backOrderQtyIdx] || item[shippedQtyIdx] || item[skuIdx] === 'Comment')
                        .map(item => {

                          noteValues = reuploadNotes.find(sku => sku[0] == removeDashesFromSku(item[skuIdx])) 
                          qty = Number(item[backOrderQtyIdx]) + Number(item[shippedQtyIdx]);

                          return (noteValues) ? 
                              [orderDate, enteredByAndApproval[0], customerName, item[originalOrderedQtyIdx], qty, removeDashesFromSku(item[skuIdx]), 
                              item[descriptionIdx], item[unitPriceIdx], qty*Number(item[unitPriceIdx]), locationName , orderNumber, 
                              noteValues[1], noteValues[2]] : 
                            [orderDate, enteredByAndApproval[0], customerName, item[originalOrderedQtyIdx], qty, removeDashesFromSku(item[skuIdx]), 
                            item[descriptionIdx], item[unitPriceIdx], qty*Number(item[unitPriceIdx]), locationName , orderNumber, 
                            '', '']}) : 
      items.map(item => { 

        noteValues = reuploadNotes.find(sku => sku[0] == removeDashesFromSku(item[skuIdx])) 

        return (noteValues) ? 
            [orderDate, enteredByAndApproval[0], customerName, enteredByAndApproval[1], item[orderedQtyIdx], removeDashesFromSku(item[skuIdx]), 
            item[descriptionIdx], item[unitPriceIdx], Number(item[orderedQtyIdx])*Number(item[unitPriceIdx]), locationName , orderNumber, 
            noteValues[1], noteValues[2]] :
          [orderDate, enteredByAndApproval[0], customerName, enteredByAndApproval[1], item[orderedQtyIdx], removeDashesFromSku(item[skuIdx]), 
          item[descriptionIdx], item[unitPriceIdx], Number(item[orderedQtyIdx])*Number(item[unitPriceIdx]), locationName , orderNumber, 
          '', '']})

    spreadsheet.deleteSheet(reuploadSheet)
  }
  else
  {
    var isCurrentOrder, deletedItemsFromCurrentOrder = [];

    if (numRows > 0)
    { 
      const ordNum = itemSheet.getSheetValues(2, 1, 1, 14).flat().indexOf('Order #');

      var currentItems = itemSheet.getSheetValues(3, 1, numRows, itemSheet.getLastColumn()).filter(item => {

        isCurrentOrder = item[ordNum] == orderNumber;

        if (isCurrentOrder)
          deletedItemsFromCurrentOrder.push([item[5], item[11], item[12]])

        return isBlank(item[ordNum]) || !isCurrentOrder;
      });
    }

    if (deletedItemsFromCurrentOrder.length !== 0)
    {
      var newItems = (doesOrderContainBOs(orderNumber, orderNumbers_BO)) ? 
                        items.filter(item => item[isItemCompleteIdx])
                             .filter(item => item[backOrderQtyIdx] || item[shippedQtyIdx] || item[skuIdx] === 'Comment')
                             .map(item => {

                                noteValues = deletedItemsFromCurrentOrder.find(sku => sku[0] == removeDashesFromSku(item[skuIdx])) 
                                qty = Number(item[backOrderQtyIdx]) + Number(item[shippedQtyIdx]);

                                return (noteValues) ? 
                                    [orderDate, enteredByAndApproval[0], customerName, item[originalOrderedQtyIdx], qty, removeDashesFromSku(item[skuIdx]), 
                                    item[descriptionIdx], item[unitPriceIdx], qty*Number(item[unitPriceIdx]), locationName , orderNumber, 
                                    noteValues[1], noteValues[2]] : 
                                  [orderDate, enteredByAndApproval[0], customerName, item[originalOrderedQtyIdx], qty, removeDashesFromSku(item[skuIdx]), 
                                  item[descriptionIdx], item[unitPriceIdx], qty*Number(item[unitPriceIdx]), locationName , orderNumber, 
                                  '', '']}) : 
        items.map(item => { 

          noteValues = deletedItemsFromCurrentOrder.find(sku => sku[0] == removeDashesFromSku(item[skuIdx])) 

          return (noteValues) ? 
              [orderDate, enteredByAndApproval[0], customerName, enteredByAndApproval[1], item[orderedQtyIdx], removeDashesFromSku(item[skuIdx]), 
              item[descriptionIdx], item[unitPriceIdx], Number(item[orderedQtyIdx])*Number(item[unitPriceIdx]), locationName , orderNumber, 
              noteValues[1], noteValues[2]] :
            [orderDate, enteredByAndApproval[0], customerName, enteredByAndApproval[1], item[orderedQtyIdx], removeDashesFromSku(item[skuIdx]), 
            item[descriptionIdx], item[unitPriceIdx], Number(item[orderedQtyIdx])*Number(item[unitPriceIdx]), locationName , orderNumber, 
            '', '']})
    }
    else
    {
      var newItems = (doesOrderContainBOs(orderNumber, orderNumbers_BO)) ? 
                        items.filter(item => item[isItemCompleteIdx])
                             .filter(item => item[backOrderQtyIdx] || item[shippedQtyIdx] || item[skuIdx] === 'Comment')
                             .map(item => 
                                [orderDate, enteredByAndApproval[0], customerName, item[originalOrderedQtyIdx], 
                                Number(item[backOrderQtyIdx]) + Number(item[shippedQtyIdx]), removeDashesFromSku(item[skuIdx]), item[descriptionIdx], 
                                item[unitPriceIdx], Number(Number(item[backOrderQtyIdx]) + Number(item[shippedQtyIdx]))*Number(item[unitPriceIdx]), 
                                locationName , orderNumber, '', '']) : 
      items.map(item => 
        [orderDate, enteredByAndApproval[0], customerName, enteredByAndApproval[1], item[orderedQtyIdx], removeDashesFromSku(item[skuIdx]), 
        item[descriptionIdx], item[unitPriceIdx], Number(item[orderedQtyIdx])*Number(item[unitPriceIdx]), locationName , orderNumber, 
        '', ''])
    }
  }

  const numNewItems = newItems.length;
  var numItemsRemoved = numNewItems;

  if (numRows > 0)
  { 
    var numCurrentItems = currentItems.length;
    itemSheet.getRange(3, 1, numCurrentItems, currentItems[0].length).setValues(currentItems);

    if (numRows > numCurrentItems)
    {
      numItemsRemoved = numRows - numCurrentItems;
      itemSheet.deleteRows(numCurrentItems + 3, numItemsRemoved);
    }
  }

  if (numNewItems > 0)
  {
    const numCols = newItems[0].length;

    if (numRows > 0)
      itemSheet.getRange(numCurrentItems + 3, 1, numNewItems, numCols)
          .setNumberFormats(new Array(numNewItems).fill(['MMM dd, yyyy', '@', '@','#', '#', '@', '@', '$#,##0.00', '$#,##0.00', '@', '@', '@', '@'])).setValues(newItems)
        .offset(-1*numCurrentItems, 0, numCurrentItems + numNewItems, numCols).sort([{column: 11, ascending: true}]);
    else
      itemSheet.getRange(3, 1, numNewItems, numCols).setNumberFormats(new Array(numNewItems).fill(['MMM dd, yyyy', '@', '@', '#', '#', '@', '@', '$#,##0.00', '$#,##0.00', '@', '@', '@', '@']))
        .setValues(newItems)

    if (doesOrderContainBOs(orderNumber, orderNumbers_BO))
    {
      Logger.log('The following new Back Ordered items were added to the B/O tab:')
      Logger.log(newItems)

      spreadsheet.toast(numNewItems + ' Added ' + (numItemsRemoved - numNewItems) + ' Removed', 'B/O Items Imported', 60)
    }
    else
    {
      Logger.log('The following new Ordered items were added to the I/O tab:')
      Logger.log(newItems)

      spreadsheet.toast(numNewItems + ' Added ' + (numItemsRemoved - numNewItems) + ' Removed', 'I/O Items Imported', 60)
    }
  }
  else
    spreadsheet.toast('ORD# ' + orderNumber + ' may be in the process of being shipped.', '**NO B/O or I/O Items Imported**', 60)

  SpreadsheetApp.flush()
  itemSheet.getRange(2, 1, itemSheet.getLastRow() - 1, itemSheet.getLastColumn()).createFilter(); // Create a filter in the header
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
  const totalIdx = headerOE.indexOf('Amount');
  const custNumIdx = headerOE.indexOf('Customer');
  const dateIdx = headerOE.indexOf('Created Date');
  const invoicedByIdx = headerOE.indexOf('OE Invoice Initials');
  const creditNumIdx = headerOE.indexOf('Credit');
  const isInvoicedOrders = invoicedByIdx !== -1;
  const isCreditedOrders = creditNumIdx !== -1;
  const   orderNumIdx = (isInvoicedOrders || isCreditedOrders) && headerOE.indexOf('Order')   || headerOE.indexOf('Order #');
  const invoiceNumIdx = (isInvoicedOrders || isCreditedOrders) && headerOE.indexOf('Invoice') || headerOE.indexOf('Inv #');  
  const locationIdx = headerOE.indexOf('Loc');
  const customerNameIdx = headerOE.indexOf('Name');
  const employeeNameIdx = headerOE.indexOf('Created by User');
  const orderValueIdx = headerOE.indexOf('Total Order Value');
  const isOrderCompleteIdx = headerOE.indexOf('Order Complete?');
  const invoiceDateIdx = (headerOE.indexOf('Inv Date') !== -1) ? headerOE.indexOf('Inv Date') : headerOE.indexOf('OE Invoice Date');
  
  const creditedByIdx = headerOE.indexOf('OE Credit Note Initials');
  const creditDateIdx = (headerOE.indexOf('OE Credit Note Date') !== -1) ? headerOE.indexOf('OE Credit Note Date') : headerOE.indexOf('Credited');
  const months = {'01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun', '07': 'Jul', '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'};
  var notesContainCreditNumber;
  
  const lodgeCustomerSheet = spreadsheet.getSheetByName('Lodge Customer List');
  const charterGuideCustomerSheet = spreadsheet.getSheetByName('Charter & Guide Customer List');
  const lodgeCustomerNumbers = lodgeCustomerSheet.getSheetValues(3, 1, lodgeCustomerSheet.getLastRow() - 2, 1).flat();
  const charterGuideCustomerNumbers = charterGuideCustomerSheet.getSheetValues(3, 1, charterGuideCustomerSheet.getLastRow() - 2, 1).flat();
  const lodgeCustomerNames = lodgeCustomerSheet.getSheetValues(3, 2, lodgeCustomerSheet.getLastRow() - 2, 2);
  const charterGuideCustomerNames = charterGuideCustomerSheet.getSheetValues(3, 2, charterGuideCustomerSheet.getLastRow() - 2, 2);

  const lodgeOrdersSheet = spreadsheet.getSheetByName('LODGE ORDERS').activate();
  const charterGuideOrdersSheet = spreadsheet.getSheetByName('GUIDE ORDERS');
  const lodgeCompletedSheet = spreadsheet.getSheetByName('LODGE COMPLETED');
  const charterGuideCompletedSheet = spreadsheet.getSheetByName('GUIDE COMPLETED');

  const possibleNumRows_Lodge = lodgeOrdersSheet.getRange(lodgeOrdersSheet.getLastRow(), 6).getNextDataCell(SpreadsheetApp.Direction.UP).getRow() - 2;
  const possibleNumRows_Charters = charterGuideOrdersSheet.getRange(charterGuideOrdersSheet.getLastRow(), 6).getNextDataCell(SpreadsheetApp.Direction.UP).getRow() - 2;
  const numLodgeOrders = (possibleNumRows_Lodge > 0) ? possibleNumRows_Lodge : lodgeOrdersSheet.getLastRow() - 2;
  const numCharterGuideOrders = (possibleNumRows_Charters > 0) ? possibleNumRows_Charters : charterGuideOrdersSheet.getLastRow() - 2;
  const numCompletedLodgeOrders = lodgeCompletedSheet.getLastRow() - 2;
  const numCompletedCharterGuideOrders = charterGuideCompletedSheet.getLastRow() - 2;

  const lodgeOrders = (numLodgeOrders > 0) ? lodgeOrdersSheet.getSheetValues(3, 3, numLodgeOrders, 1).flat().map(ordNum => ordNum.toString()) : [];
  const charterGuideOrders = (numCharterGuideOrders > 0) ? charterGuideOrdersSheet.getSheetValues(3, 3, numCharterGuideOrders, 1).flat().map(ordNum => ordNum.toString()) : [];
  const lodgeCompleted = (numCompletedLodgeOrders > 0) ? lodgeCompletedSheet.getSheetValues(3, 11, numCompletedLodgeOrders, 1).flat().map(ordNum => ordNum.toString()) : [];
  const charterGuideCompleted = (numCompletedCharterGuideOrders > 0) ? charterGuideCompletedSheet.getSheetValues(3, 11, numCompletedCharterGuideOrders, 1).flat().map(ordNum => ordNum.toString()) : [];

  const lodgeCredited = (numCompletedLodgeOrders > 0) ? lodgeCompletedSheet.getSheetValues(3, 10, numCompletedLodgeOrders, 1).flat().map(creditNum => {
    notesContainCreditNumber = creditNum.toString().split('Credit # ');

    return (notesContainCreditNumber.length > 1) ? notesContainCreditNumber.pop().split('\n').shift().toString() : false;
  }).filter(creditNum => creditNum) : [];

  const charterGuideCredited = (numCompletedCharterGuideOrders > 0) ? charterGuideCompletedSheet.getSheetValues(3, 10, numCompletedCharterGuideOrders, 1).flat().map(creditNum => {
    notesContainCreditNumber = creditNum.toString().split('Credit # ');

    return (notesContainCreditNumber.length > 1) ? notesContainCreditNumber.pop().split('\n').shift().toString() : false;
  }).filter(creditNum => creditNum) : [];

  const today = new Date();
  const currentYear = today.getFullYear().toString()
  const lastYear = (Number(currentYear) - 1).toString()
  const lodgeSheetYear = lodgeOrdersSheet.getSheetValues(1, 1, 1, 1)[0][0].split(' ').shift();
  const todayFormattedDate = Utilities.formatDate(today, spreadsheet.getSpreadsheetTimeZone(), 'MM-dd-yyyy')

  if (lodgeSheetYear == (Number(currentYear) + 1).toString()) // Is this next years lodge sheet?
    var includeLastYearsFinalQuarterOrders = true;

  if (lodgeSheetYear == currentYear) // Is this the current lodge sheet?
    var isCurrentLodgeSeasonYear = true;

  const newLodgeOrders = 
    (isInvoicedOrders) ?       // If true, then the import is a set of invoiced orders
      allOrders.filter(order => lodgeCustomerNumbers.includes(order[custNumIdx]) && 
        ((includeLastYearsFinalQuarterOrders && order[dateIdx].substring(6) == lastYear &&
          (order[dateIdx].substring(0, 2) == '09' || order[dateIdx].substring(0, 2) == '10' || order[dateIdx].substring(0, 2) == '11' || order[dateIdx].substring(0, 2) == '12')) 
          || (isCurrentLodgeSeasonYear && order[dateIdx].substring(6) == currentYear))
        && !lodgeCompleted.includes(order[invoiceNumIdx].toString().trim())).map(order => {
        return [getDateString(order[dateIdx].toString(), months), getFullName(order[employeeNameIdx]), order[orderNumIdx], 'TRUE', '', getProperTypesetName(order[customerNameIdx], lodgeCustomerNames, 1), getLocationName(order[locationIdx]), '', '', 'This order was automatically imported', order[invoiceNumIdx], '$' + order[totalIdx], getFullName(order[invoicedByIdx]), getOrderStatus(order[isOrderCompleteIdx], isInvoicedOrders), getDateString((order[invoiceDateIdx].toString() == ' ') ? todayFormattedDate : order[invoiceDateIdx].toString(), months), order[isOrderCompleteIdx]] // Lodge Completed
      }) :
    (isCreditedOrders) ? // If true, then the import is a set of credits
      allOrders.filter(order => lodgeCustomerNumbers.includes(order[custNumIdx]) && 
        ((includeLastYearsFinalQuarterOrders && order[dateIdx].substring(6) == lastYear &&
          (order[dateIdx].substring(0, 2) == '09' || order[dateIdx].substring(0, 2) == '10' || order[dateIdx].substring(0, 2) == '11' || order[dateIdx].substring(0, 2) == '12')) 
          || (isCurrentLodgeSeasonYear && order[dateIdx].substring(6) == currentYear))
        && !lodgeCredited.includes(order[creditNumIdx].toString().trim())).map(order => {
        return [getDateString(order[dateIdx].toString(), months), getFullName(order[employeeNameIdx]), order[orderNumIdx], 'TRUE', '', getProperTypesetName(order[customerNameIdx], lodgeCustomerNames, 1), getLocationName(order[locationIdx]), '', '', 'Credit # ' + order[creditNumIdx] + '\nThis credit was automatically imported', order[invoiceNumIdx], '$' + -1*Number(order[totalIdx]), getFullName(order[creditedByIdx]), 'Credited', getDateString(order[creditDateIdx].toString(), months), 'No'] // Lodge Completed
      }) : 
    allOrders.filter(order => lodgeCustomerNumbers.includes(order[custNumIdx]) && 
      ((includeLastYearsFinalQuarterOrders && order[dateIdx].substring(6) == lastYear &&
        (order[dateIdx].substring(0, 2) == '09' || order[dateIdx].substring(0, 2) == '10' || order[dateIdx].substring(0, 2) == '11' || order[dateIdx].substring(0, 2) == '12'))
        || (isCurrentLodgeSeasonYear && order[dateIdx].substring(6) == currentYear))
      && order[isOrderCompleteIdx] == 'No' && !lodgeOrders.includes(order[orderNumIdx].toString().trim())).map(order => { 
      return [getDateString(order[dateIdx].toString(), months), getFullName(order[employeeNameIdx]), order[orderNumIdx], '', '', getProperTypesetName(order[customerNameIdx], lodgeCustomerNames, 1), getLocationName(order[locationIdx]), '', '', 'This order was automatically imported', getInvoiceNumber(order[invoiceNumIdx], order[isOrderCompleteIdx]), '', '', getOrderStatus(order[isOrderCompleteIdx], isInvoicedOrders, order[invoiceNumIdx])] // Lodge Orders
  });

  const newCharterGuideOrders = 
    (isInvoicedOrders) ?       // If true, then the import is a set of invoiced orders
      allOrders.filter(order => charterGuideCustomerNumbers.includes(order[custNumIdx]) &&
        ((includeLastYearsFinalQuarterOrders && order[dateIdx].substring(6) == lastYear &&
          (order[dateIdx].substring(0, 2) == '09' || order[dateIdx].substring(0, 2) == '10' || order[dateIdx].substring(0, 2) == '11' || order[dateIdx].substring(0, 2) == '12'))
          || (isCurrentLodgeSeasonYear && order[dateIdx].substring(6) == currentYear))
        && !charterGuideCompleted.includes(order[invoiceNumIdx].toString().trim())).map(order => { 
        return [getDateString(order[dateIdx].toString(), months), getFullName(order[employeeNameIdx]), order[orderNumIdx], 'TRUE', '', getProperTypesetName(order[customerNameIdx], charterGuideCustomerNames, 1), getLocationName(order[locationIdx]), '', '', 'This order was automatically imported', order[invoiceNumIdx], '$' + order[totalIdx], getFullName(order[invoicedByIdx]), getOrderStatus(order[isOrderCompleteIdx], isInvoicedOrders), getDateString((order[invoiceDateIdx].toString() == ' ') ? todayFormattedDate : order[invoiceDateIdx].toString(), months), order[isOrderCompleteIdx]] // Charter & Guide Completed
      }) :
    (isCreditedOrders) ? // If true, then the import is a set of credits
      allOrders.filter(order => charterGuideCustomerNumbers.includes(order[custNumIdx]) &&
        ((includeLastYearsFinalQuarterOrders && order[dateIdx].substring(6) == lastYear &&
          (order[dateIdx].substring(0, 2) == '09' || order[dateIdx].substring(0, 2) == '10' || order[dateIdx].substring(0, 2) == '11' || order[dateIdx].substring(0, 2) == '12'))
          || (isCurrentLodgeSeasonYear && order[dateIdx].substring(6) == currentYear))
        && !charterGuideCredited.includes(order[creditNumIdx].toString().trim())).map(order => { 
        return [getDateString(order[dateIdx].toString(), months), getFullName(order[employeeNameIdx]), order[orderNumIdx], 'TRUE', '', getProperTypesetName(order[customerNameIdx], charterGuideCustomerNames, 1), getLocationName(order[locationIdx]), '', '', 'Credit # ' + order[creditNumIdx] + '\nThis credit was automatically imported', order[invoiceNumIdx], '$' + -1*Number(order[totalIdx]), getFullName(order[creditedByIdx]), 'Credited', getDateString(order[creditDateIdx].toString(), months), 'No'] // Charter & Guide Completed
      }) : 
    allOrders.filter(order => charterGuideCustomerNumbers.includes(order[custNumIdx]) &&
      ((includeLastYearsFinalQuarterOrders && order[dateIdx].substring(6) == lastYear &&
        (order[dateIdx].substring(0, 2) == '09' || order[dateIdx].substring(0, 2) == '10' || order[dateIdx].substring(0, 2) == '11' || order[dateIdx].substring(0, 2) == '12'))
        || (isCurrentLodgeSeasonYear && order[dateIdx].substring(6) == currentYear))
      && order[isOrderCompleteIdx] == 'No' && !charterGuideOrders.includes(order[orderNumIdx].toString().trim())).map(order => {
      return [getDateString(order[dateIdx].toString(), months), getFullName(order[employeeNameIdx]), order[orderNumIdx], '', '', getProperTypesetName(order[customerNameIdx], charterGuideCustomerNames, 1), getLocationName(order[locationIdx]), '', '', 'This order was automatically imported', getInvoiceNumber(order[invoiceNumIdx], order[isOrderCompleteIdx]), '', '', getOrderStatus(order[isOrderCompleteIdx], isInvoicedOrders, order[invoiceNumIdx])] // Charter & Guide Orders
  });

  const numNewLodgeOrder = newLodgeOrders.length;
  const numNewCharterGuideOrder = newCharterGuideOrders.length;

  if (numNewLodgeOrder > 0)
  {
    var numCols_Lodge = newLodgeOrders[0].length;

    if (isInvoicedOrders || isCreditedOrders)
    {
      var lodgePartiallyCompleteOrders = newLodgeOrders.map(ord => [ord[2], ord.pop()])
      numCols_Lodge--; // The "Order Complete?" staus is removed from the end of the array

      lodgeCompletedSheet.activate().getRange(numCompletedLodgeOrders + 3, 1, numNewLodgeOrder, numCols_Lodge)
          .setNumberFormats(new Array(numNewLodgeOrder).fill(['MMM dd, yyyy', '@', '@', '#', '@', '@', '@', '@', '@', '@', '@', '$#,##0.00', '@', '@', 'MMM dd, yyyy'])).setValues(newLodgeOrders)
        .offset(-1*numCompletedLodgeOrders, 0, numCompletedLodgeOrders + numNewLodgeOrder, numCols_Lodge)
          .sort([{column: 15, ascending: true}, {column: 11, ascending: true}, {column: 1, ascending: true}]);
    }
    else
      lodgeOrdersSheet.activate().getRange(numLodgeOrders + 3, 1, numNewLodgeOrder, numCols_Lodge)
          .setNumberFormats(new Array(numNewLodgeOrder).fill(['MMM dd, yyyy', '@', '@', '#', '@', '@', '@', '@', '@', '@', '@', '$#,##0.00', '@', '@']))
          .setFontColor('black').setFontLine('none').setValues(newLodgeOrders)
        .offset(-1*numLodgeOrders, 0, numLodgeOrders + numNewLodgeOrder, numCols_Lodge)
          .sort([{column: 1, ascending: true}, {column: 3, ascending: true}]);

    Logger.log('The following new Lodge orders were added to the tracker:')
    Logger.log(newLodgeOrders)

    deleteBackOrderedItems(newLodgeOrders, spreadsheet, lodgePartiallyCompleteOrders);
  }
  else
    var lodgePartiallyCompleteOrders = [];

  if (numNewCharterGuideOrder > 0)
  {
    var numCols_CharterGuide = newCharterGuideOrders[0].length;

    if (isInvoicedOrders || isCreditedOrders)
    {
      var charterGuidePartiallyCompleteOrders = newCharterGuideOrders.map(ord => [ord[2], ord.pop()])
      numCols_CharterGuide--; // The "Order Complete?" staus is removed from the end of the array
      
      charterGuideCompletedSheet.getRange(numCompletedCharterGuideOrders + 3, 1, numNewCharterGuideOrder, numCols_CharterGuide)
          .setNumberFormats(new Array(numNewCharterGuideOrder).fill(['MMM dd, yyyy', '@', '@', '#', '@', '@', '@', '@', '@', '@', '@', '$#,##0.00', '@', '@', 'MMM dd, yyyy'])).setValues(newCharterGuideOrders)
        .offset(-1*numCompletedCharterGuideOrders, 0, numCompletedCharterGuideOrders + numNewCharterGuideOrder, numCols_CharterGuide)
          .sort([{column: 15, ascending: true}, {column: 11, ascending: true}, {column: 1, ascending: true}]);
    }
    else
      charterGuideOrdersSheet.getRange(numCharterGuideOrders + 3, 1, numNewCharterGuideOrder, numCols_CharterGuide)
          .setNumberFormats(new Array(numNewCharterGuideOrder).fill(['MMM dd, yyyy', '@', '@', '#', '@', '@', '@', '@', '@', '@', '@', '$#,##0.00', '@', '@'])).setValues(newCharterGuideOrders)
        .offset(-1*numCharterGuideOrders, 0, numCharterGuideOrders + numNewCharterGuideOrder, numCols_CharterGuide)
          .sort([{column: 1, ascending: true}, {column: 3, ascending: true}]);

    Logger.log('The following new Charter and Guide orders were added to the tracker:')
    Logger.log(newCharterGuideOrders)

    deleteBackOrderedItems(newCharterGuideOrders, spreadsheet, charterGuidePartiallyCompleteOrders);
  }
  else
    var charterGuidePartiallyCompleteOrders = []

  // Orders that are fully completed may need to be removed from the Lodge Orders and Guide Orders page
  if (isInvoicedOrders || isCreditedOrders)
  {
    var isLodgeOrderComplete, isCharterGuideOrderComplete;
    SpreadsheetApp.flush();

    if (lodgeCompletedSheet.getLastRow() > 2)
    {
      const completedLodgeOrderNumbers = lodgeCompletedSheet.getSheetValues(3, 3, lodgeCompletedSheet.getLastRow() - 2, 12)
        .filter(ord => ord[11] === 'Completed')
        .map(ord => ord[0].toString()).flat()
        .filter(ordNum => isNotBlank(ordNum) && ordNum !== 'No Order'); 

      Logger.log('The following Lodge Orders were removed because they were found to be fully completed as per the invoice history:')
      const currentLodgeOrders = lodgeOrdersSheet.getSheetValues(3, 1, numLodgeOrders, 14).map(currentOrd => {

        isLodgeOrderComplete = lodgePartiallyCompleteOrders.find(partialOrd => partialOrd[0] == currentOrd[2])

        if (isLodgeOrderComplete && isLodgeOrderComplete[1] === 'No')
        {
          currentOrd[10] = 'multiple';
          currentOrd[13] = 'Partial Order';
        }

        return currentOrd;

        }).filter(currentOrd => {

          isLodgeOrderComplete = completedLodgeOrderNumbers.includes(currentOrd[2]); 

          if (isLodgeOrderComplete)
            Logger.log(currentOrd);
          
          return !isLodgeOrderComplete;
      });

      var numCurrentLodgeOrders = currentLodgeOrders.length;

      if (numCurrentLodgeOrders < numLodgeOrders)
        lodgeOrdersSheet.getRange(3, 1, numLodgeOrders, 14).clearContent().offset(0, 0, numCurrentLodgeOrders, 14).setValues(currentLodgeOrders);
      else if (numCurrentLodgeOrders === numLodgeOrders)
        lodgeOrdersSheet.getRange(3, 1, numLodgeOrders, 14).setValues(currentLodgeOrders);
    }
    else
      var numCurrentLodgeOrders = numLodgeOrders;

    if (charterGuideCompletedSheet.getLastRow() > 2)
    {
      const completedCharterGuideOrderNumbers = charterGuideCompletedSheet.getSheetValues(3, 3, charterGuideCompletedSheet.getLastRow() - 2, 12)
        .filter(ord => ord[11] === 'Completed')
        .map(ord => ord[0].toString()).flat()
        .filter(ordNum => isNotBlank(ordNum) && ordNum !== 'No Order');

      Logger.log('The following Guide Orders were removed because they were found to be fully completed as per the invoice history:')
      const currentCharterGuideOrders = charterGuideOrdersSheet.getSheetValues(3, 1, numCharterGuideOrders, 14).map(currentOrd => {
        
        isCharterGuideOrderComplete = charterGuidePartiallyCompleteOrders.find(partialOrd => partialOrd[0] == currentOrd[2])

        if (isCharterGuideOrderComplete && isCharterGuideOrderComplete[1] === 'No')
        {
          currentOrd[10] = 'multiple';
          currentOrd[13] = 'Partial Order';
        }

        return currentOrd;

        }).filter(currentOrd => {

          isCharterGuideOrderComplete = completedCharterGuideOrderNumbers.includes(currentOrd[2]); // Invoice # and Order Status must both be blank

          if (isCharterGuideOrderComplete)
            Logger.log(currentOrd);

          return !isCharterGuideOrderComplete;
        });

      var numCurrentCharterGuideOrders = currentCharterGuideOrders.length;

      if (numCurrentCharterGuideOrders < numCharterGuideOrders)
        charterGuideOrdersSheet.getRange(3, 1, numCharterGuideOrders, 14).clearContent().offset(0, 0, numCurrentCharterGuideOrders, 14).setValues(currentCharterGuideOrders);
      else if (numCurrentCharterGuideOrders === numCharterGuideOrders)
        charterGuideOrdersSheet.getRange(3, 1, numCharterGuideOrders, 14).setValues(currentCharterGuideOrders);
    }
    else
      var numCurrentCharterGuideOrders = numCharterGuideOrders;
  }
  else // Cancelled Orders (if ANY)
  {
    var isLodgeOrderCancelled, isCharterGuideOrderCancelled, cancelledOrders = [];
    const templateSheet = spreadsheet.getSheetByName('Reupload:');
    const boSheet = spreadsheet.getSheetByName('B/O')
    const ioSheet = spreadsheet.getSheetByName('I/O')
    const numCols = boSheet.getLastColumn()
    const numBoItems = boSheet.getLastRow() - 2;
    const numIoItems = ioSheet.getLastRow() - 2;
    var boItems = (numBoItems > 0) ? boSheet.getSheetValues(3, 1, numBoItems, numCols) : null;
    var ioItems = (numIoItems > 0) ? ioSheet.getSheetValues(3, 1, numIoItems, numCols) : null;
    var itemsOnOrder = [], reuploadSheet, isThisOrdNumberRemovedFromItemsSheet, numIOsRemoved = 0, numBOsRemoved = 0;

    SpreadsheetApp.flush();

    Logger.log('The following orders need to be reuploaded because the Total Order Value has changed which means items may have been added or removed from the order:')

    // Return a list of the current order numbers but while compiling that list, check if any orders have changed and the B/O or I/O sheets need to have their items updated
    const currentOrderNumbers = allOrders.map(ord => {

      if (lodgeOrders.includes(ord[orderNumIdx].toString().trim()) || charterGuideOrders.includes(ord[orderNumIdx].toString().trim())) // Check for current orders
      {
        if (ioItems)
        {
          itemsOnOrder = ioItems.filter(ordNum => ordNum[10] == ord[orderNumIdx]);

          if (itemsOnOrder.length !== 0)
          {
            if (Math.round((itemsOnOrder.map(amount => Number(amount[8])).reduce((total, amount) => total + amount, 0) + Number.EPSILON)*100)/100 % Number(ord[orderValueIdx]))
            {
              Logger.log(Number(ord[orderNumIdx]))
              reuploadSheet = spreadsheet.insertSheet('Reupload:' + ord[orderNumIdx], {template: templateSheet}).hideSheet();

              ioItems = ioItems.filter(ordNum => {

                isThisOrdNumberRemovedFromItemsSheet = ordNum[10] !== ord[orderNumIdx];

                if (!isThisOrdNumberRemovedFromItemsSheet)
                  reuploadSheet.appendRow([ordNum[5], ordNum[11], ordNum[12], '', ordNum[13]]);

                return isThisOrdNumberRemovedFromItemsSheet
              }); // Remove the items from the I/O page

              numIOsRemoved++;
            }
          }
          else
          {
            itemsOnOrder = boItems.filter(ordNum => ordNum[10] == ord[orderNumIdx]);

            if (itemsOnOrder.length !== 0 && Math.round((itemsOnOrder.map(amount => Number(amount[8])).reduce((total, amount) => total + amount, 0) + Number.EPSILON)*100)/100 % Number(ord[orderValueIdx]))
            {
              Logger.log(Number(ord[orderNumIdx]))
              reuploadSheet = spreadsheet.insertSheet('Reupload:' + ord[orderNumIdx], {template: templateSheet}).hideSheet();

              boItems = boItems.filter(ordNum => {

                isThisOrdNumberRemovedFromItemsSheet = ordNum[10] !== ord[orderNumIdx];

                if (!isThisOrdNumberRemovedFromItemsSheet)
                  reuploadSheet.appendRow([ordNum[5], ordNum[11], ordNum[12], '', ordNum[13]]);

                return isThisOrdNumberRemovedFromItemsSheet
              }); // Remove the items from the B/O page

              numBOsRemoved++;
            }
          }
        }
        else if (boItems)
        {
          itemsOnOrder = boItems.filter(ordNum => ordNum[10] == ord[orderNumIdx]);

          if (itemsOnOrder.length !== 0 && Math.round((itemsOnOrder.map(amount => Number(amount[8])).reduce((total, amount) => total + amount, 0) + Number.EPSILON)*100)/100 % Number(ord[orderValueIdx]))
          {
            Logger.log(Number(ord[orderValueIdx]))
            reuploadSheet = spreadsheet.insertSheet('Reupload:' + ord[orderNumIdx], {template: templateSheet}).hideSheet();

            boItems = boItems.filter(ordNum => {

              isThisOrdNumberRemovedFromItemsSheet = ordNum[10] !== ord[orderNumIdx];

              if (!isThisOrdNumberRemovedFromItemsSheet)
                reuploadSheet.appendRow([ordNum[5], ordNum[11], ordNum[12], '', ordNum[13]]);

              return isThisOrdNumberRemovedFromItemsSheet
            }); // Remove the items from the B/O page

            numBOsRemoved++;
          }
        }

        itemsOnOrder.length = 0;
      }
  
      return ord[orderNumIdx]
    }).flat().filter(ordNum => isNotBlank(ordNum)); 

    const currentLodgeOrders = lodgeOrdersSheet.getSheetValues(3, 1, numLodgeOrders, 14)
      .filter(currentOrd => {

        isLodgeOrderCancelled = !isBlank(currentOrd[2]) && !currentOrderNumbers.includes(currentOrd[2]);

        if (isLodgeOrderCancelled)
        {
          currentOrd.push(today) // Set the cancelled date as today
          currentOrd[9] = (isBlank(currentOrd[9])) ? 'This order was automatically cancelled' : 'This order was automatically cancelled\n' + currentOrd[9];
          cancelledOrders.push(currentOrd)
        }
          
        return !isLodgeOrderCancelled;
    });

    const currentCharterGuideOrders = charterGuideOrdersSheet.getSheetValues(3, 1, numCharterGuideOrders, 14)
      .filter(currentOrd => {

        isCharterGuideOrderCancelled = !isBlank(currentOrd[2]) && !currentOrderNumbers.includes(currentOrd[2]);

        if (isCharterGuideOrderCancelled)
        {
          currentOrd.push(today) // Set the cancelled date as today
          currentOrd[9] = (isBlank(currentOrd[9])) ? 'This order was automatically cancelled' : 'This order was automatically cancelled\n' + currentOrd[9];
          cancelledOrders.push(currentOrd)
        }
          
        return !isCharterGuideOrderCancelled;
    });

    var numCancelledOrders = cancelledOrders.length;
    var numCurrentLodgeOrders = currentLodgeOrders.length;
    var numCurrentCharterGuideOrders = currentCharterGuideOrders.length;

    if (numCancelledOrders > 0)
    {
      const cancelledSheet = spreadsheet.getSheetByName('CANCELLED')
      cancelledSheet.getRange(cancelledSheet.getLastRow() + 1, 1, numCancelledOrders, 15)
        .setNumberFormats(new Array(numCancelledOrders).fill(['MMM dd, yyyy', '@', '@', '#', '@', '@', '@', '@', '@', '@', '@', '$#,##0.00', '@', '@', 'MMM dd, yyyy']))
        .setValues(cancelledOrders)
      Logger.log('The following orders were removed from the tracker and placed on the CANCELLED page because they were NOT found in OrderEntry:')
      Logger.log(cancelledOrders)

      deleteBackOrderedItems(cancelledOrders, spreadsheet);
    }

    if (numCurrentLodgeOrders < numLodgeOrders)
      lodgeOrdersSheet.getRange(3, 1, numLodgeOrders, 14).clearContent().offset(0, 0, numCurrentLodgeOrders, 14).setValues(currentLodgeOrders);

    if (numCurrentCharterGuideOrders < numCharterGuideOrders)
      charterGuideOrdersSheet.getRange(3, 1, numCharterGuideOrders, 14).clearContent().offset(0, 0, numCurrentCharterGuideOrders, 14).setValues(currentCharterGuideOrders);

    if (numIOsRemoved > 0)
    {
      const numCurrentIoItems = ioItems.length

      if (numCurrentIoItems < numIoItems && numCurrentIoItems > 0)
      {
        ioSheet.getRange(3, 1, numCurrentIoItems, numCols).setValues(ioItems)
        ioSheet.deleteRows(numCurrentIoItems + 3, numIoItems - numCurrentIoItems);
      }
    }

    if (numBOsRemoved > 0)
    {
      const numCurrentBoItems = boItems.length

      if (numCurrentBoItems < numBoItems && numCurrentBoItems > 0)
      {
        boSheet.getRange(3, 1, numCurrentBoItems, numCols).setValues(boItems)
        boSheet.deleteRows(numCurrentBoItems + 3, numBoItems - numCurrentBoItems);
      }
    }
  }

  spreadsheet.toast('LODGE: ' + numNewLodgeOrder + ' Added ' + (numLodgeOrders - numCurrentLodgeOrders) + ' Removed GUIDE: ' + numNewCharterGuideOrder + ' Added ' + (numCharterGuideOrders - numCurrentCharterGuideOrders) + ' Removed', 'Orders Imported', 60)
}

/**
 * This function handles the import of a purchase order that contains items that the lodge has ordered.
 * 
 * @param {String[][]}     items    : A list of items on the purchase order that was imported.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function updatePoItemsOnTracker(items, spreadsheet)
{
  items.pop(); // Remove the "Total" or final line

  // Get all the indexes of the relevant headers
  const headerOE = items.shift();
  const originalOrderedQtyIdx = headerOE.indexOf('Qty Originally Ordered');
  const backOrderQtyIdx = headerOE.indexOf('Backordered'); 
  const skuIdx = headerOE.indexOf('Item#');
  const descriptionIdx = headerOE.indexOf('Description');
  const unitCostIdx = headerOE.indexOf('Unit Cost');
  const extendedUnitCostIdx = headerOE.indexOf('Extended Order Cost');
  const months = {'01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun', '07': 'Jul', '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'};
  const orderDate = getDateString(items[0][headerOE.indexOf('Rate Date')], months);
  const locationName = getLocationName(items[0][headerOE.indexOf('Location')]);
  const vendorName = items[0][headerOE.indexOf('Vendor name')];
  const purchaseOrderNumber = items[0][headerOE.indexOf('Doc #')];
  const reuploadSheet = spreadsheet.getSheetByName('Reupload:' + purchaseOrderNumber);

  if (reuploadSheet)
  {
    const reuploadNotes = reuploadSheet.getSheetValues(2, 1, reuploadSheet.getLastRow() - 1, 5);
    var noteValues;

    var newItems = items.map(item => {

      noteValues = reuploadNotes.find(sku => sku[0] == removeDashesFromSku(item[skuIdx]));

      return (noteValues) ? 
          [orderDate, vendorName, item[originalOrderedQtyIdx], item[backOrderQtyIdx], removeDashesFromSku(item[skuIdx]), 
            item[descriptionIdx], item[unitCostIdx], item[extendedUnitCostIdx], locationName, purchaseOrderNumber, 
            noteValues[1], noteValues[3]] : 
        [orderDate, vendorName, item[originalOrderedQtyIdx], item[backOrderQtyIdx], removeDashesFromSku(item[skuIdx]), 
          item[descriptionIdx], item[unitCostIdx], item[extendedUnitCostIdx], locationName, purchaseOrderNumber, '', '']

    }).filter(item => item[3] !== 0 || item[4] === ' '); // Remove items that have already been received, as well as keep comments / notes

    spreadsheet.deleteSheet(reuploadSheet)
  }
  else
  {
    var newItems = items.map(item => [orderDate, vendorName, item[originalOrderedQtyIdx], item[backOrderQtyIdx], 
      removeDashesFromSku(item[skuIdx]), item[descriptionIdx], item[unitCostIdx], item[extendedUnitCostIdx], locationName , purchaseOrderNumber, '', '']
    ).filter(item => item[3] !== 0 || item[4] === ' '); // Remove items that have already been received, as well as keep comments / notes
  }

  const poItemSheet = spreadsheet.getSheetByName('P/O').activate(); 
  const numRows = poItemSheet.getLastRow() - 2;
  const numCols = poItemSheet.getLastColumn();
  const numNewItems = newItems.length;
  var numItemsRemoved = numNewItems;
  poItemSheet?.getFilter()?.remove(); // Remove the filter

  if (numRows > 0)
  {
    const poNum = poItemSheet.getSheetValues(2, 1, 1, numCols).flat().indexOf('Purchase Order #');
    var currentItems = poItemSheet.getSheetValues(3, 1, numRows, numCols).filter(item => item[poNum] !== purchaseOrderNumber);
    var numCurrentItems = currentItems.length;
    poItemSheet.getRange(3, 1, numCurrentItems, currentItems[0].length).setValues(currentItems);

    if (numRows > numCurrentItems)
    {
      numItemsRemoved = numRows - numCurrentItems;
      poItemSheet.deleteRows(numCurrentItems + 3, numItemsRemoved);
    }
  }

  Logger.log('Purchase Order Number: ' + purchaseOrderNumber)

  if (numNewItems > 0)
  {
    if (numRows > 0)
      poItemSheet.getRange(numCurrentItems + 3, 1, numNewItems, numCols)
          .setNumberFormats(new Array(numNewItems).fill(['MMM dd, yyyy', '@', '#','#', '@', '@', '$#,##0.00', '$#,##0.00', '@', '@', '@', '@'])).setValues(newItems)
        .offset(-1*numCurrentItems, 0, numCurrentItems + numNewItems, numCols).sort([{column: 10, ascending: true}]);
    else
      poItemSheet.getRange(3, 1, numNewItems, numCols).setNumberFormats(new Array(numNewItems).fill(['MMM dd, yyyy', '@', '#','#', '@', '@', '$#,##0.00', '$#,##0.00', '@', '@', '@', '@']))
        .setValues(newItems)

    Logger.log('The following new Ordered items were added to the P/O tab:')
    Logger.log(newItems)

    spreadsheet.toast(numNewItems + ' Added ' + (numItemsRemoved - numNewItems) + ' Removed', 'P/O Items Imported', 60)
  }
  else
    spreadsheet.toast(purchaseOrderNumber + ' may be in the process of being received.', '**NO Items Imported**', 60)

  SpreadsheetApp.flush()
  poItemSheet.getRange(2, 1, poItemSheet.getLastRow() - 1, poItemSheet.getLastColumn()).createFilter(); // Create a filter in the header
}

/**
 * This function will be run on a trigger daily and it will update the PNT Receiving spreadsheet with the relevant data that it finds on this spreadsheet.
 * 
 * @author Jarren Ralf
 */
function updatedPntReceivingSpreadsheet()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const poItemsSheet = spreadsheet.getSheetByName('P/O') 
  const recdItemsSheet = spreadsheet.getSheetByName("Rec'd")
  const numCols_PoSheet = poItemsSheet.getLastColumn();
  const numCols_RecdSheet = recdItemsSheet.getLastColumn();
  const header_PoSheet = poItemsSheet.getSheetValues(2, 1, 1, numCols_PoSheet)[0];
  const header_RecdSheet = recdItemsSheet.getSheetValues(2, 1, 1, numCols_RecdSheet)[0];
  const orderDateIdx_PoSheet = header_PoSheet.indexOf('Order Date')
  const    vendorIdx_PoSheet = header_PoSheet.indexOf('Vendor')
  const  poNumberIdx_PoSheet = header_PoSheet.indexOf('Purchase Order #')
  const    receiptDateIdx_RecdSheet = header_RecdSheet.indexOf('Receipt Date')
  const         vendorIdx_RecdSheet = header_RecdSheet.indexOf('Vendor')
  const       poNumberIdx_RecdSheet = header_RecdSheet.indexOf('Purchase Order #')
  const  receiptNumberIdx_RecdSheet = header_RecdSheet.indexOf('Receipt #')
  const tritesPackingSlipsSheet = SpreadsheetApp.openById('1qzyTAmtVIfOCxuhv0KzBkHpnWXFqE79DpnJSAqnZyvk').getSheetByName('Trites Packing slips')
  const lastRow = tritesPackingSlipsSheet.getLastRow();
  const numRows = lastRow - 2;
  const tritesPackingSlips_ReceiptNumbers = tritesPackingSlipsSheet.getSheetValues(3, 4, numRows, 1).flat()
  const tritesPackingSlips_PoNumbersNotReceived = tritesPackingSlipsSheet.getSheetValues(3, 3, numRows, 2)
    .filter(poNum => !isBlank(poNum[0]) && isBlank(poNum[1])).map(poNum => poNum[0]) // PO number is not blank while the receipt number is blank

  const posAndReceipts = poItemsSheet.getSheetValues(3, 1, poItemsSheet.getLastRow() - 2, numCols_PoSheet)                   // All PO items
      .filter((row, index, arr) => arr.findIndex(row2 => row2[poNumberIdx_PoSheet] === row[poNumberIdx_PoSheet]) >= index)   // Keep the unique PO numbers
      .map(newPos => [newPos[orderDateIdx_PoSheet], newPos[vendorIdx_PoSheet], newPos[poNumberIdx_PoSheet], '', '', '', '']) // Map to the correct format
    .concat(recdItemsSheet.getSheetValues(3, 1, recdItemsSheet.getLastRow() - 2, numCols_RecdSheet)                                                                  // All Received items
      .filter((row, index, arr) => arr.findIndex(row2 => row2[receiptNumberIdx_RecdSheet] === row[receiptNumberIdx_RecdSheet]) >= index)                             // Keep the unique Receipt numbers
      .map(newPos => [newPos[receiptDateIdx_RecdSheet], newPos[vendorIdx_RecdSheet], newPos[poNumberIdx_RecdSheet], newPos[receiptNumberIdx_RecdSheet], '', '', '']) // Map to the correct format
    ).filter(rctNum => 
      !tritesPackingSlips_ReceiptNumbers.includes(rctNum[3]) ||                                                   // Remove receipts that are already on the list
      (isBlank(rctNum[3]) && !isBlank(rctNum[2]) && !tritesPackingSlips_PoNumbersNotReceived.includes(rctNum[2])) // Remove pos that are already on the list
  )

  const numNewPosAndReceipts = posAndReceipts.length;

  if (numNewPosAndReceipts > 0)
    tritesPackingSlipsSheet.getRange(lastRow + 1, 1, numNewPosAndReceipts, 7).setValues(posAndReceipts).offset(-1*numRows - 1, 0, numRows + numNewPosAndReceipts + 1, 7).sort([{column: 1, ascending: false}]);
}

/**
 * This function handles the import of the list of receipts into the spreadsheet.
 * 
 * @param {String[][]} allReceipts : All of the current receipts from Adagio.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function updatePoReceiptsOnTracker(allReceipts, spreadsheet)
{
  allReceipts.pop(); // Remove the "Total" or final line

  // Get all the indexes of the relevant headers
  const headerOE = allReceipts.shift();
  const numReceipts = allReceipts.length;
  const orderDateIdx = headerOE.indexOf('Order Date');
  const poNumberIdx = headerOE.indexOf('Original Doc');
  const receiptNumberIdx = headerOE.indexOf('Document');
  const itemManagementSheet = spreadsheet.getSheetByName('Item Management (Jarren Only ;)')
  const itemManagement_NumRows = itemManagementSheet.getLastRow() - 1;
  const itemManagement_Receipt = itemManagementSheet.getSheetValues(2, 12, itemManagement_NumRows, 1).filter(u => !isBlank(u[0])).flat();
  const itemManagement_NonLodgeReceipt = itemManagementSheet.getSheetValues(2, 14, itemManagement_NumRows, 1).filter(v => !isBlank(v[0])).flat();
  const itemManagement_ReceiptsWithPos = itemManagementSheet.getSheetValues(2, 11, itemManagement_NumRows, 2).filter(u => !isBlank(u[1]))
  const currentYear = new Date().getFullYear().toString();
  const lastYear = new Date().getFullYear().toString();
  const lodgeSheetYear = spreadsheet.getSheetByName('LODGE ORDERS').getSheetValues(1, 1, 1, 1)[0][0].split(' ').shift();
  var numReceiptsAdded = 0;

  if (lodgeSheetYear == (new Date().getFullYear() + 1).toString()) // Is this next years lodge sheet?
    var includeLastYearsFinalQuarterOrders = true;

  if (lodgeSheetYear == currentYear) // Is this next years lodge sheet?
    var isCurrentLodgeSeasonYear = true;

  for (var i = 0; i < numReceipts; i++)
  {
    // Make sure the Receipt is for this year
    if (((includeLastYearsFinalQuarterOrders && allReceipts[i][orderDateIdx].toString().substring(6) == lastYear &&
      (allReceipts[i][orderDateIdx].toString().substring(0, 2) == '09' || allReceipts[i][orderDateIdx].toString().substring(0, 2) == '10' || allReceipts[i][orderDateIdx].toString().substring(0, 2) == '11' || allReceipts[i][orderDateIdx].toString().substring(0, 2) == '12')) 
      || (isCurrentLodgeSeasonYear && allReceipts[i][orderDateIdx].toString().substring(6) == currentYear)))
    {
      if (!itemManagement_Receipt.includes(allReceipts[i][receiptNumberIdx]) && !itemManagement_NonLodgeReceipt.includes(allReceipts[i][receiptNumberIdx])) // This PO is not in either item managment PO list
      {
        itemManagement_ReceiptsWithPos.push([allReceipts[i][poNumberIdx], allReceipts[i][receiptNumberIdx]]) // Add the PO number to the item management po list
        Logger.log('Add this Receipt to Item Management List: ' + allReceipts[i][receiptNumberIdx])
        numReceiptsAdded++;
      }
    }
  }

  if (numReceiptsAdded > 0)
    itemManagementSheet.getRange(2, 11, itemManagement_ReceiptsWithPos.length, 2).setValues(itemManagement_ReceiptsWithPos.sort((a, b) => (a[1] < b[1]) ? -1 : (a[1] > b[1]) ? 1 : 0)).activate()

  Logger.log('numReceiptsAdded: ' + numReceiptsAdded)

  spreadsheet.toast(numReceiptsAdded + ' Added ', 'Receipts Imported', 60)
}

/**
 * Update the discount structure and cost for lead and bait on this spreadsheet.
 * 
 * @author Jarren Ralf
 */
function updatePriceAndCostOfLeadAndFrozenBait()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const today = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'dd MMM yyyy')
  const leadSheet = spreadsheet.getSheetByName('Lead Cost & Pricing');
  const baitSheet = spreadsheet.getSheetByName('Bait Cost & Pricing');
  const numLeadItems = leadSheet.getLastRow() - 2;
  const numBaitItems = baitSheet.getLastRow() - 2;
  const lastColumn_LeadSheet = leadSheet.getMaxColumns();
  const lastColumn_BaitSheet = baitSheet.getMaxColumns();
  const leadSheetRange = leadSheet.getRange(3, 1, numLeadItems, lastColumn_LeadSheet);
  const baitSheetRange = baitSheet.getRange(3, 1, numBaitItems, lastColumn_BaitSheet);
  const formats_leadSheet = ['@', '@', '@', '@', '@', '@', 'dd MMM yyyy', '$0.00', '$0.00', '$0.00', '$0.00', '$0.00', '$0.00', '$0.00', '$0.00', '$0.00', '#', '#%', '$0.00', '#%', '$0.00', '#%', '$0.00', '#%', '$0.00', '#%', '$0.00', '@'];
  const formats_baitSheet = ['@', '@', '@', '@', '@', '@', 'dd MMM yyyy', '$0.00', '$0.00', '#', '#%', '$0.00', '#%', '$0.00', '#%', '$0.00', '#%', '$0.00', '@'];
  const costData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString());
  const discountSS = SpreadsheetApp.openById('1gXQ7uKEYPtyvFGZVmlcbaY6n6QicPBhnCBxk-xqwcFs');
  const discountSheet = discountSS.getSheetByName('Discount Percentages');
  const discounts = discountSheet.getSheetValues(2, 11, discountSheet.getLastRow() - 1, 5)
  const header = costData.shift();
  const itemNumber_InventoryCsv = header.indexOf('Item #')
  const cost = header.indexOf('Cost')
  var itemValues, discountValues, googleDescription, category, vendor;

  const leadItems = leadSheetRange.getValues().map(item => {
    itemValues = costData.find(sku => sku[itemNumber_InventoryCsv].toString().toUpperCase() == item[0])
    discountValues = discounts.find(description => description[0].split(' - ').pop().toString().toUpperCase() == item[0].toString().toUpperCase())
    item[10] = ''; // Freight per pc ***Clear the cost and price fields that are calculated by formulas
    item[11] = ''; // Freight to PR
    item[12] = ''; // Alberni + Freight
    item[13] = ''; // Jason + Freight
    item[14] = ''; // NEW
    item[20] = ''; // Guide Price
    item[22] = ''; // Lodge Price
    item[24] = ''; // Wholesale Price
    item[26] = ''; // Early Booking Price

    item[25] = .23; // Early Booking Percent

    if (discountValues)
    {
      item[18] = Number(discountValues[1]);     // Base Price
      item[19] = Number(discountValues[2])/100; // Guide Percent
      item[21] = Number(discountValues[3])/100; // Lodge Percent
      item[23] = Number(discountValues[4])/100; // Wholesale Percent
    }

    if (itemValues)
    {
      googleDescription = itemValues[1].split(' - ')
      googleDescription.pop() // SKU
      googleDescription.pop() // UoM
      category = googleDescription.pop()
      vendor = googleDescription.pop()

      if (vendor !== item[3])
      {
        item[2] = vendor;
        leadSheet.showColumns(3, 2)
      }

      if (category !== item[5])
      {
        item[4] = category;
        leadSheet.showColumns(5, 2);
      }

      item[15] = itemValues[cost];                         // Adagio Cost
      item[17] = Number(item[18])/Number(itemValues[cost]) // Markup %
    }

    return item
  })

  leadSheet.hideColumns(lastColumn_LeadSheet);
  leadSheetRange.setNumberFormats(new Array(numLeadItems).fill(formats_leadSheet)).setValues(leadItems)
    .offset(-2, 1, 1, 1).setValue('Description\n\n[Updated At: ' + new Date().toLocaleTimeString() + ' on ' + today + ']\n[Prices Updated At: ' + discountSS.getSheetValues(2, 2, 1, 1)[0][0].split(' at ')[1] + ']')

  const baitItems = baitSheetRange.getValues().map((item, i) => {
    itemValues = costData.find(sku => sku[itemNumber_InventoryCsv].toString().toUpperCase() == item[0])
    discountValues = discounts.find(description => description[0].split(' - ').pop().toString().toUpperCase() == item[0].toString().toUpperCase())

    if (discountValues)
    {
      item[11] =  Number(discountValues[1]);                                                   // Base Price
      item[12] =  Number(discountValues[2])/100;                                               // Guide Percent
      item[13] = (Number(discountValues[1])*(100 - Number(discountValues[2]))/100).toFixed(2); // Guide Price
      item[14] =  Number(discountValues[3])/100;                                               // Lodge Percent
      item[15] = (Number(discountValues[1])*(100 - Number(discountValues[3]))/100).toFixed(2); // Lodge Price
      item[16] =  Number(discountValues[4])/100;                                               // Wholesale Percent
      item[17] = (Number(discountValues[1])*(100 - Number(discountValues[4]))/100).toFixed(2); // Wholesale Price
    }

    if (itemValues)
    {
      googleDescription = itemValues[1].split(' - ')
      googleDescription.pop() // SKU
      googleDescription.pop() // UoM
      category = googleDescription.pop()
      vendor = googleDescription.pop()

      if (vendor !== item[3])
      {
        item[2] = vendor;
        baitSheet.showColumns(3, 2)
      }

      if (category !== item[5])
      {
        item[4] = category;
        baitSheet.showColumns(5, 2);
      }

      item[ 8] = itemValues[cost]; // Adagio Cost
      item[10] = Number(item[11])/Number(itemValues[cost]) // Markup %
    }

    return item
  })

  baitSheet.hideColumns(lastColumn_BaitSheet)
  baitSheetRange.setNumberFormats(new Array(numBaitItems).fill(formats_baitSheet)).setValues(baitItems)
    .offset(-2, 1, 1, 1).setValue('Description\n\n[Updated At: ' + new Date().toLocaleTimeString() + ' on ' + today + ']')
}

/**
 * This function handles the import of the list of purchase orders into the spreadsheet.
 * 
 * @param {String[][]} allPurchaseOrders : All of the current purchase orders from Adagio.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function updatePurchaseOrdersOnTracker(allPurchaseOrders, spreadsheet)
{
  allPurchaseOrders.pop(); // Remove the "Total" or final line

  // Get all the indexes of the relevant headers
  const headerOE = allPurchaseOrders.shift();
  const numPOs = allPurchaseOrders.length;
  const dateIdx = headerOE.indexOf('Order Date');
  const poNumberIdx = headerOE.indexOf('Document');
  const poStatusIdx = headerOE.indexOf('Automatic Style Code');
  const totalValueIdx = headerOE.indexOf('Total Value');
  const poItemsSheet = spreadsheet.getSheetByName('P/O')
  const numItemsOnPos = poItemsSheet.getLastRow() - 2;
  const numCols_PoItems = poItemsSheet.getLastColumn();

  if (numItemsOnPos > 0)
    var poItems = poItemsSheet.getSheetValues(3, 1, numItemsOnPos, numCols_PoItems);
    
  const itemManagementSheet = spreadsheet.getSheetByName('Item Management (Jarren Only ;)')
  const itemManagement_NumRows = itemManagementSheet.getLastRow() - 1;
  const itemManagement_Po_Range = itemManagementSheet.getRange(2, 7, itemManagement_NumRows, 1);
  const itemManagement_Po = itemManagement_Po_Range.getValues().filter(u => !isBlank(u[0])).flat();
  const itemManagement_NonLodgePo_Range = itemManagementSheet.getRange(2, 9, itemManagement_NumRows, 1);
  const itemManagement_NonLodgePo = itemManagement_NonLodgePo_Range.getValues().filter(v => !isBlank(v[0])).flat();
  const currentYear = new Date().getFullYear().toString();
  const lastYear = new Date().getFullYear().toString();
  const lodgeSheetYear = spreadsheet.getSheetByName('LODGE ORDERS').getSheetValues(1, 1, 1, 1)[0][0].split(' ').shift();
  const templateSheet = spreadsheet.getSheetByName('Reupload:');
  var numPOsAdded = 0, numPOsRemoved = 0, itemManagement_Po_Idx = -1, itemManagement_NonLodgePo_Idx = -1, isThisPoNumberRemovedFromItemsSheet, reuploadSheet;

  if (lodgeSheetYear == (new Date().getFullYear() + 1).toString()) // Is this next years lodge sheet?
    var includeLastYearsFinalQuarterOrders = true;

  if (lodgeSheetYear == currentYear) // Is this next years lodge sheet?
    var isCurrentLodgeSeasonYear = true;

  for (var i = 0; i < numPOs; i++)
  {
    // Make sure the PO is for this year
    if (((includeLastYearsFinalQuarterOrders && allPurchaseOrders[i][dateIdx].toString().substring(6).toString() == lastYear &&
      (allPurchaseOrders[i][dateIdx].toString().substring(0, 2) == '09' || allPurchaseOrders[i][dateIdx].toString().substring(0, 2) == '10' || allPurchaseOrders[i][dateIdx].toString().substring(0, 2) == '11' || allPurchaseOrders[i][dateIdx].toString().substring(0, 2) == '12')) 
      || (isCurrentLodgeSeasonYear && allPurchaseOrders[i][dateIdx].toString().substring(6).toString() == currentYear)))
    {
      Logger.log('PO Number: ' + allPurchaseOrders[i][poNumberIdx])
      Logger.log('PO Status: ' + allPurchaseOrders[i][poStatusIdx])

      if (allPurchaseOrders[i][poStatusIdx] !== 'PO Completed')
      {
        if (!itemManagement_Po.includes(allPurchaseOrders[i][poNumberIdx]) && !itemManagement_NonLodgePo.includes(allPurchaseOrders[i][poNumberIdx])) // This PO is not in either item managment PO list
        {
          Logger.log('Add this PO to Item Management List: ' + allPurchaseOrders[i][poNumberIdx])
          itemManagement_Po.push(allPurchaseOrders[i][poNumberIdx]) // Add the PO number to the item management po list
          numPOsAdded++;
        }
        else if (allPurchaseOrders[i][poStatusIdx]  === 'PO Part Received' && allPurchaseOrders[i][totalValueIdx] != 0 && numItemsOnPos > 0 &&
          !(Math.round((poItems.filter(poNum => poNum[9] == allPurchaseOrders[i][poNumberIdx]).map(amount => Number(amount[7])).reduce((total, amount) => total + amount, 0) + Number.EPSILON)*100)/100 % 
          Number(allPurchaseOrders[i][totalValueIdx]) === 0)) 
        {
          Logger.log('This PO is partially received: ' + allPurchaseOrders[i][poNumberIdx] + '. The items on this order need to be imported again.')

          reuploadSheet = spreadsheet.insertSheet('Reupload:' + allPurchaseOrders[i][poNumberIdx], {template: templateSheet}).hideSheet();

          poItems = poItems.filter(poNum => {

            isThisPoNumberRemovedFromItemsSheet = poNum[9] !== allPurchaseOrders[i][poNumberIdx];

            if (!isThisPoNumberRemovedFromItemsSheet)
              reuploadSheet.appendRow([poNum[4], poNum[10], '', poNum[11], poNum[12]]);

            return isThisPoNumberRemovedFromItemsSheet
          }); // Remove the items from the P/O page

          numPOsRemoved++;
        }
      }
      else // PO is complete
      {
        Logger.log('This PO is complete.')
        // Remove all lines that match this PO number from the P/O sheet
        itemManagement_Po_Idx = itemManagement_Po.findIndex(poNum => poNum == allPurchaseOrders[i][poNumberIdx])
        itemManagement_NonLodgePo_Idx = itemManagement_NonLodgePo.findIndex(poNum => poNum == allPurchaseOrders[i][poNumberIdx])

        if (itemManagement_Po_Idx !== -1)
        {
          itemManagement_Po[itemManagement_Po_Idx] = false;
          Logger.log('Remove this PO from Item Management List: ' + allPurchaseOrders[i][poNumberIdx])
          numPOsRemoved++;
        }

        if (itemManagement_NonLodgePo_Idx !== -1)
        {
          itemManagement_NonLodgePo[itemManagement_NonLodgePo_Idx] = false;
          Logger.log('Remove this PO from Item Management List: ' + allPurchaseOrders[i][poNumberIdx])
          numPOsRemoved++;
        }

        if (numItemsOnPos > 0)
          poItems = poItems.filter(poNum => poNum[9] !== allPurchaseOrders[i][poNumberIdx]); // Remove the items from the P/O page
      }
      Logger.log('-------------------------------------------')
    }
  }

  var numPoItemsRemoved = 0;

  if (numPOsAdded !== 0 || numPOsRemoved !== 0)
  {
    const itemManagement_Po_Updated = itemManagement_Po.filter(u => u).sort().map(v => [v]);
    itemManagement_Po_Range.clearContent().offset(0, 0, itemManagement_Po_Updated.length).setValues(itemManagement_Po_Updated);

    const itemManagement_NonLodgePo_Updated = itemManagement_NonLodgePo.filter(u => u).sort().map(v => [v]);
    itemManagement_NonLodgePo_Range.clearContent().offset(0, 0, itemManagement_NonLodgePo_Updated.length).setValues(itemManagement_NonLodgePo_Updated);

    if (numItemsOnPos > 0)
    {
      const numCurrentItems = poItems.length

      if (numCurrentItems < numItemsOnPos)
      {
        numPoItemsRemoved = numItemsOnPos - numCurrentItems;
        poItemsSheet.getRange(3, 1, numCurrentItems, numCols_PoItems).setValues(poItems)
        poItemsSheet.deleteRows(numCurrentItems + 3, numPoItemsRemoved);
      }
    }
  }

  Logger.log('numPOsAdded: ' + numPOsAdded)
  Logger.log('numPOsRemoved: ' + numPOsRemoved)
  Logger.log('numPoItemsRemoved: ' + numPoItemsRemoved)

  itemManagementSheet.getRange('G2').activate();
  spreadsheet.toast(numPOsAdded + ' Added ' + numPOsRemoved + ' Removed   ' + numPoItemsRemoved + ' Items Removed from P/O sheet', 'POs Imported', 60)
}

/**
 * This function handles the import of a Receipt (from Adagio PurchaseOrder) that contains items that the lodge has ordered and received.
 * 
 * @param {String[][]}     items    : A list of items on the receipt that was imported.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function updateReceivedItemsOnTracker(items, spreadsheet)
{
  items.pop(); // Remove the "Total" or final line

  // Get all the indexes of the relevant headers
  const headerOE = items.shift();
  const originalOrderedQtyIdx = headerOE.indexOf('Qty Originally Ordered');
  const receivedQtyIdx = headerOE.indexOf('Received'); 
  const backOrderQtyIdx = headerOE.indexOf('Backordered'); 
  const skuIdx = headerOE.indexOf('Item #');
  const descriptionIdx = headerOE.indexOf('Description');
  const unitCostIdx = headerOE.indexOf('Unit Cost');
  const extendedUnitCostIdx = headerOE.indexOf('Ext Cost');
  const months = {'01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun', '07': 'Jul', '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'};
  const vendorName = items[0][headerOE.indexOf('Vendor name')];
  const receiptDate = getDateString(items[0][headerOE.indexOf('Expected Date')], months);
  const locationName = getLocationName(items[0][headerOE.indexOf('Location')]);
  const itemManagementSheet = spreadsheet.getSheetByName('Item Management (Jarren Only ;)')
  const itemManagement_ReceiptsWithPos = itemManagementSheet.getSheetValues(2, 11, itemManagementSheet.getLastRow() - 1, 2).filter(u => !isBlank(u[1]))
  const receiptNumber = items[0][headerOE.indexOf('Rcpt #')];
  const purchaseOrderNumber = itemManagement_ReceiptsWithPos.find(rct => rct[1] == receiptNumber)[0];
  const recdItemSheet = spreadsheet.getSheetByName("Rec'd").activate(); 
  const numCurrentItems = recdItemSheet.getLastRow() - 2;
  recdItemSheet?.getFilter()?.remove(); // Remove the filter

  Logger.log('Receipt Number: ' + receiptNumber)

  const newItems = items.map(item => [receiptDate, vendorName, item[originalOrderedQtyIdx], item[receivedQtyIdx], item[backOrderQtyIdx], 
    removeDashesFromSku(item[skuIdx]), item[descriptionIdx], item[unitCostIdx], item[extendedUnitCostIdx], locationName , purchaseOrderNumber, receiptNumber])

  const numNewItems = newItems.length;
  const numCols = newItems[0].length;
  const receiptNumCol = recdItemSheet.getLastColumn();

  if (numCurrentItems > 0)
    recdItemSheet.getRange(numCurrentItems + 3, 1, numNewItems, numCols)
        .setNumberFormats(new Array(numNewItems).fill(['MMM dd, yyyy', '@', '#', '#', '#', '@', '@', '$#,##0.00', '$#,##0.00', '@', '@', '@'])).setValues(newItems)
      .offset(-1*numCurrentItems, 0, numCurrentItems + numNewItems, numCols).sort([{column: receiptNumCol, ascending: true}]);
  else
    recdItemSheet.getRange(3, 1, numNewItems, numCols).setNumberFormats(new Array(numNewItems).fill(['MMM dd, yyyy', '@', '#', '#', '#', '@', '@', '$#,##0.00', '$#,##0.00', '@', '@', '@'])).setValues(newItems)

  Logger.log("The following new received items were added to the Rec'd tab:")
  Logger.log(newItems)

  spreadsheet.toast(numNewItems + ' Added ', "Rec'd Items Imported", 60)

  SpreadsheetApp.flush()
  recdItemSheet.getRange(2, 1, recdItemSheet.getLastRow() - 1, receiptNumCol).createFilter(); // Create a filter in the header
}