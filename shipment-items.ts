let shipmentItemAccessToken = ''

function fetchShipmentItems(nextToken = '') {
  const storedToken = userProperties.getProperty('nextToken') || '';

  const queryType = 'DATE_RANGE';
  const marketplaceId = 'A1PA6795UKMFR9'
  const lastUpdatedAfter = '2015-09-18T22:51:57.926Z';
  const lastUpdatedBefore = new Date().toISOString();

  const apiUrl = `https://sellingpartnerapi-eu.amazon.com/fba/inbound/v0/shipmentItems?LastUpdatedAfter=${lastUpdatedAfter}&LastUpdatedBefore=${lastUpdatedBefore}&QueryType=${queryType}&MarketplaceId=${marketplaceId}&NextToken=${encodeURIComponent(nextToken || storedToken)}`;

  if (!shipmentItemAccessToken) {
    shipmentItemAccessToken = getAuthenticationToken();
  }

  const headers = {
    'Accept': 'application/json',
    'x-amz-access-token': shipmentItemAccessToken
  };

  const options = {
    'method': 'get',
    'headers': headers,
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseData = response.getContentText();
    const responseJson = JSON.parse(responseData);
    // const nextTokenData = responseJson?.payload?.NextToken || null;
    const shipmentItemsData = responseJson?.payload?.ItemData || null;

    if (Array.isArray(shipmentItemsData) && shipmentItemsData.length > 0) {
      dataArray.push(...shipmentItemsData);
    }


    // if (nextTokenData && dataArray?.length < 5000) {
    //   fetchShipmentItems(nextTokenData);

    //   return;
    // }

    insertShipmentItems(dataArray);

    // showAlert({
    //     title: 'Fetch Complete',
    //     message: "There are no more shipment items to fetch. You've completed fetching all shipment items"
    // });

    // if (nextTokenData) {
    //   userProperties.setProperty('nextToken', nextTokenData);
    // } else {
    //   userProperties.deleteProperty('nextToken');
    //   showAlert({
    //     title: 'Fetch Complete',
    //     message: "There are no more shipment items to fetch. You've completed fetching all shipment items"
    //   });
    // }
  } catch (error) {
    console.log(error);
  }
}

function insertShipmentItems(data) {
  if (!Array.isArray(data) || !data?.length) {
    showAlert({
      title: 'Data not found',
      message: 'There is no data in the api response'
    });
    return;
  }

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('shipments_item');

  const sortedData = organizeShipmentItems(data);

  const getSheetRange = sheet.getRange(sheet.getLastRow() + 1, 1, sortedData.length, sortedData[0].length);

  getSheetRange.setValues(sortedData);
}


function organizeShipmentItems(data) {
  let organizedData = [];

  data.forEach(item => {
    organizedData.push([
      item?.ShipmentId || '', // Column A: Shipment Id
      item?.SellerSKU || '', // Column B: Seller SKU
      item?.FulfillmentNetworkSKU || '', // Column C: Fulfillment Network SKU
      item?.QuantityShipped || 0, // Column D: Quantity Shipped
      item?.QuantityReceived || 0, // Column E: Quantity Received
      item?.QuantityInCase || 0, // Column F: Quantity In Case
      item?.PrepDetailsList[0]?.PrepInstruction || '', // Column G: Prep Instruction
      item?.PrepDetailsList[0]?.PrepOwner || '', // Column H: Prep Owner
    ]);
  });

  return organizedData;
}