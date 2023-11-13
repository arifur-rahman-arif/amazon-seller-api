let accessToken = ''

function fetchShipments(nextToken = '') {
  const storedToken = userProperties.getProperty('nextToken') || '';

  const shipmentStatusList = 'CLOSED,CHECKED_IN,WORKING,READY_TO_SHIP,SHIPPED,RECEIVING,CANCELLED,DELETED,CLOSED,ERROR,IN_TRANSIT,DELIVERED,CHECKED_IN';
  const queryType = 'SHIPMENT';
  const marketplaceId = 'MarketplaceId'


  const apiUrl = `https://sellingpartnerapi-eu.amazon.com/fba/inbound/v0/shipments?ShipmentStatusList=${shipmentStatusList}&QueryType=${queryType}&MarketplaceId=${marketplaceId}&NextToken=${encodeURIComponent(nextToken || storedToken)}`;

  if (!accessToken) {
    accessToken = getAuthenticationToken();
  }

  const headers = {
    'Accept': 'application/json',
    'x-amz-access-token': accessToken
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
    const nextTokenData = responseJson?.payload?.NextToken || null;
    const shipmentsData = responseJson?.payload?.ShipmentData || null;

    if (Array.isArray(shipmentsData) && shipmentsData.length > 0) {
      dataArray.push(...shipmentsData);
    } else {
      insertShipments(dataArray);

      userProperties.deleteProperty('nextToken');
      // userProperties.deleteProperty('dialogShownFlag');
      showAlert({
        title: 'Fetch Complete',
        message: "There are no more data to fetch. You've completed fetching all data"
      });

      return;
    }


    if (nextTokenData && dataArray?.length < 5000) {
      fetchShipments(nextTokenData);

      return;
    }

    insertShipments(dataArray);

    if (nextTokenData) {
      userProperties.setProperty('nextToken', nextTokenData);
      // showAlert({
      //   title: 'More data to fetch',
      //   message: 'There are more data you can fetch. Please re-run the script again to get the rest of the remaining data'
      // });
    } else {
      userProperties.deleteProperty('nextToken');
      // userProperties.deleteProperty('dialogShownFlag');
      showAlert({
        title: 'Fetch Complete',
        message: "There are no more data to fetch. You've completed fetching all data"
      });
    }
  } catch (error) {
    console.log(error);
  }
}

function insertShipments(data) {
  if (!Array.isArray(data) || !data?.length) {
    showAlert({
      title: 'Data not found',
      message: 'There is no data in the api response'
    });
    return;
  }

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('shipments');

  const sortedData = organizeShipmentData(data);

  const getSheetRange = sheet.getRange(sheet.getLastRow() + 1, 1, sortedData.length, sortedData[0].length);

  getSheetRange.setValues(sortedData);
}


function organizeShipmentData(data) {
  let organizedData = [];

  data.forEach(shipment => {
    organizedData.push([
      shipment?.ShipmentId || '', // Column A: Shipment Id
      shipment?.ShipmentName || '', // Column B: Shipment Name
      shipment?.ShipFromAddress?.Name || '', // Column C: Ship From Address (Name)
      shipment?.ShipFromAddress?.AddressLine1 || '', // Column D: Ship From Address (Address Line1)
      shipment?.ShipFromAddress?.City || '', // Column E: Ship From Address (City)
      shipment?.ShipFromAddress?.CountryCode || '', // Column F: Ship From Address (CountryCode)
      shipment?.ShipFromAddress?.PostalCode || '', // Column G: Ship From Address (Postal Code)
      shipment?.DestinationFulfillmentCenterId || '', // Column H: Destination Fulfillment Center Id
      shipment?.ShipmentStatus || '', // Column I: Shipment Status
      shipment?.LabelPrepType || '', // Column J: Label Prep Type
      shipment?.AreCasesRequired ? 'Yes' : 'No', // Column K: Are Cases Required
    ]);
  });

  return organizedData;
}