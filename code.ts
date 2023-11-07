const dataArray = [];
const userProperties = PropertiesService.getUserProperties();



/**
 * Creates a custom menu in the Google Sheets UI for accessing Amazon Seller API functions.
 *
 * This function is triggered when the Google Sheet is opened and creates a custom menu
 * in the UI, allowing users to access the "Fetch Inventory" function.
 *
 * @param {Object} e - The event object (not used in this function).
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("Amazon Seller API")
    .addItem("Fetch Inventory", "fetchInventory")
    .addItem("Reset Inventory Pagination", "resetPagination")
    .addToUi();
}

/**
 * Reset the API pagination to start it from 1 again
 */
function resetPagination() {
  userProperties.deleteProperty('nextToken');
  userProperties.deleteProperty('dialogShownFlag');
  showAlert({
    title: 'Data pagination restarted',
    message: 'Next time you make a reuqest it will start from the first inventory item'
  });
}

/**
 * Fetches inventory data from the Amazon Seller API and inserts it into a Google Sheet.
 *
 * This function constructs the API URL using provided parameters, fetches data from the Amazon Seller API,
 * and inserts the inventory data into a Google Sheet in the specified format.
 *
 * @param {boolean} details - Indicates whether to include additional details in the API request.
 * @param {string} granularityType - The type of granularity for the API request (e.g., "Marketplace").
 * @param {string} granularityId - The unique identifier for the granularity (e.g., "").
 * @param {string} marketplaceIds - The marketplace IDs for the API request (e.g., "").
 */
function fetchInventory(nextToken = '') {
  const storedToken = userProperties.getProperty('nextToken') || '';

  const dialogShownFlag = userProperties.getProperty('dialogShownFlag');

  if (!storedToken && !dialogShownFlag) {
    // Display a dialog box with a message and "Yes" and "No" buttons. The user can also close the
    // dialog by clicking the close button in its title bar.
    const ui = SpreadsheetApp.getUi();
    var userResponse = ui.alert('Are you sure you want to continue? You are about to start fetching from the first inventory item', ui.ButtonSet.YES_NO);

    // Process the user's userResponse.
    if (userResponse == ui.Button.NO) {
      return false;
    } else {
      // Set a flag to indicate that the dialog has been shown.
      userProperties.setProperty('dialogShownFlag', 'true');
    }
  }

  var details = true;
  var granularityType = 'Marketplace';
  var granularityId = '';
  var marketplaceIds = '';

  const apiUrl = `https://sellingpartnerapi-eu.amazon.com/fba/inventory/v1/summaries?details=${details}&granularityType=${granularityType}&granularityId=${granularityId}&marketplaceIds=${marketplaceIds}&nextToken=${nextToken || storedToken}`;

  const accessToken = getAuthenticationToken();

  if (!accessToken) {
    showAlert({
      title: 'Authentication Error',
      message: 'Unable to access api token'
    });
    return;
  }

  const headers = {
    'Accept': 'application/json',
    'x-amz-access-token': accessToken
  };

  const options = {
    'method': 'get',
    'headers': headers
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseData = response.getContentText();
    const responseJson = JSON.parse(responseData);
    const nextTokenData = responseJson?.pagination?.nextToken || null;
    const inventorySummeries = responseJson?.payload?.inventorySummaries || null;

    dataArray.push(...inventorySummeries);

    if (nextTokenData && dataArray?.length < 1000) {
      fetchInventory(nextTokenData);

      return;
    }

    insertData(dataArray);

    if (nextTokenData) {
      userProperties.setProperty('nextToken', nextTokenData);
      showAlert({
        title: 'More data to fetch',
        message: 'There are more data you can fetch. Please re-run the script again to get the rest of the remaining data'
      });
    } else {
      userProperties.deleteProperty('nextToken');
      userProperties.deleteProperty('dialogShownFlag');
      showAlert({
        title: 'Fetch Complete',
        message: "There are no more data to fetch. You've completed fetching all data"
      });
    }
  } catch (error) {
    console.log(error);
  }
}

/**
 * Organizes inventory data into a format suitable for insertion into a Google Sheet.
 *
 * @param {Array} data - The inventory data to be organized.
 * @returns {Array} - The organized data in the specified format.
 */
function getAuthenticationToken() {
  const apiUrl = 'https://api.amazon.com/auth/o2/token';
  const grantType = 'refresh_token';
  const refreshToken = '';
  const clientId = '';
  const clientSecret = '';

  const payload = {
    grant_type: grantType,
    refresh_token: refreshToken,
    client_id: clientId,
    client_secret: clientSecret,
  };

  const options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: payload
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseData = response.getContentText();
    const responseJson = JSON.parse(responseData);

    return responseJson?.access_token || null;
  } catch (error) {
    Logger.log(error);
  }
}

/**
 * Displays an alert dialog with the specified title and message.
 *
 * This function creates a simple alert dialog with the provided title and message
 * and presents it to the user for informational purposes.
 *
 * @param {Object} args - An object containing the title and message for the alert.
 * @param {string} args.title - The title of the alert dialog.
 * @param {string} args.message - The message content of the alert dialog.
 */
function showAlert(args) {
  let { title, message } = args;
  let ui = SpreadsheetApp.getUi();
  ui.alert(title, message, ui.ButtonSet.OK);
}

/**
 * Inserts organized inventory data into a Google Sheet in a specified format.
 *
 * This function takes the organized data and inserts it into a Google Sheet with the specified
 * column order (A to Q). If the provided data is not an array or is empty, it displays an error message.
 *
 * @param {Array} data - The organized inventory data to be inserted into the Google Sheet.
 */
function insertData(data) {

  if (!Array.isArray(data) || !data?.length) {
    showAlert({
      title: 'Data not found',
      message: 'There is no data in the api response'
    });
    return;
  }

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('inventory');

  // Define the column order (A to Q)
  const columns = [
    "Product Name", // Column A
    "asin", // Column B
    "fnSku", // Column C
    "sellerSku", // Column D
    "fulfillableQuantity", // Column E
    "inboundWorkingQuantity", // Column F
    "inboundShippedQuantity", // Column G
    "inboundReceivingQuantity", // Column H
    "totalReservedQuantity", // Column I
    "pendingCustomerOrderQuantity", // Column J
    "pendingTransshipmentQuantity", // Column K
    "fcProcessingQuantity", // Column L
    "reservedFutureSupplyQuantity", // Column M
    "futureSupplyBuyableQuantity", // Column N
    "totalQuantity" // Column O
  ];

  const sortedData = organizeData(data);

  const getSheetRange = sheet.getRange(sheet.getLastRow() + 1, 1, sortedData.length, sortedData[0].length);

  getSheetRange.setValues(sortedData);
}

function organizeData(data) {
  let organizedData = [];

  data.forEach(summary => {
    organizedData.push([
      summary?.productName || '', // Column A: Product Name
      summary?.asin || '', // Column B: asin
      summary?.fnSku || '', // Column C: fnSku
      summary?.sellerSku || '', // Column D: sellerSku
      (summary?.inventoryDetails?.fulfillableQuantity === 0) ? 0 : summary?.inventoryDetails?.fulfillableQuantity, // Column E: fulfillableQuantity

      (summary?.inventoryDetails?.inboundWorkingQuantity === 0) ? 0 : summary?.inventoryDetails?.inboundWorkingQuantity, // Column F: inboundWorkingQuantity

      (summary?.inventoryDetails?.inboundShippedQuantity === 0) ? 0 : summary?.inventoryDetails?.inboundShippedQuantity, // Column G: inboundShippedQuantity

      (summary?.inventoryDetails?.inboundReceivingQuantity === 0) ? 0 : summary?.inventoryDetails?.inboundReceivingQuantity, // Column H: inboundReceivingQuantity

      (summary?.inventoryDetails?.reservedQuantity?.totalReservedQuantity === 0) ? 0 : summary?.inventoryDetails?.reservedQuantity?.totalReservedQuantity, // Column I: totalReservedQuantity

      (summary?.inventoryDetails?.reservedQuantity?.pendingCustomerOrderQuantity === 0) ? 0 : summary?.inventoryDetails?.reservedQuantity?.pendingCustomerOrderQuantity, // Column J: pendingCustomerOrderQuantity

      (summary?.inventoryDetails?.reservedQuantity?.pendingTransshipmentQuantity === 0) ? 0 : summary?.inventoryDetails?.reservedQuantity?.pendingTransshipmentQuantity, // Column K: pendingTransshipmentQuantity

      (summary?.inventoryDetails?.reservedQuantity?.fcProcessingQuantity === 0) ? 0 : summary?.inventoryDetails?.reservedQuantity?.fcProcessingQuantity, // Column L: fcProcessingQuantity

      (summary?.inventoryDetails?.futureSupplyQuantity?.reservedFutureSupplyQuantity === 0) ? 0 : summary?.inventoryDetails?.futureSupplyQuantity?.reservedFutureSupplyQuantity, // Column M: reservedFutureSupplyQuantity

      (summary?.inventoryDetails?.futureSupplyQuantity?.futureSupplyBuyableQuantity === 0) ? 0 : summary?.inventoryDetails?.futureSupplyQuantity?.futureSupplyBuyableQuantity, // Column N: futureSupplyBuyableQuantity

      (summary?.totalQuantity === 0) ? 0 : summary?.totalQuantity, // Column O: totalQuantity
    ]);
  });

  return organizedData;
}