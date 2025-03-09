// TODO: COPY API KEY PASTE YOUR API KEY
// Guide https://help.take.app/en/article/google-sheets-integration-ytpy3p/
const API_KEY="cm80f9lvz000ijy03uf0xa5iu";

function importOrders() {
  // Get saved API key from script properties
  const properties = PropertiesService.getScriptProperties();
  
  if (!API_KEY) {
    return;
  }
  
  const url = `https://take.app/api/platform/orders?api_key=${API_KEY}`;
  try {
    var response = UrlFetchApp.fetch(url);
    var orders = JSON.parse(response.getContentText());

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Check if the sheet has headers; if not, add them
    var headers = [
      'Link', 'Created At', 'Order', 'Status', 'Payment Status', 'Fulfillment Status', 
      'Remark', 'Internal Note', 'Line Item Name', 'Line Item Price', 'Line Item Quantity', 'Line Item Options',
      'Total', 'Customer Name', 'Customer Phone', 'Customer Email',
      'Delivery Name', 'Delivery Fee', 'Delivery Date', 'Delivery Time', 'Delivery Slot',
      'Delivery Address', 'Delivery Address 2', 'Delivery City', 'Delivery State', 'Delivery Zip'
    ];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
    }

    // Get existing order IDs
    var existingIds = sheet.getRange(1, 1, sheet.getLastRow()).getValues().flat();

    // Add parsed rows
    orders.forEach(order => {
      const link = `https://take.app/orders/${order.id}`; // Link to order
      if (!existingIds.includes(link)) {
        // Process line items
        const lineItems = order.items || [];
        
        // Create rows for line items
        lineItems.forEach((item, index) => {
          const options = item.options || [];
          const itemOptions = options.map(opt => {
            const answers = opt.answers.join(', ');
            return `${opt.text}: ${answers}`;
          }).join('; ');

          var row = [
            link,
            new Date(order.createdAt).toLocaleString(), // Format date
            order.number,
            index === 0 ? order.status.replace('ORDER_STATUS_', '') : '', // Remove prefix
            index === 0 ? order.paymentStatus.replace('PAYMENT_STATUS_', '') : '', // Remove prefix
            index === 0 ? order.fulfillmentStatus : '',
            index === 0 ? order.remark : '',
            index === 0 ? order.internalNote : '',
            item.name,
            (item.price / 100).toFixed(2),
            item.quantity,
            itemOptions,
            index === 0 ? (order.totalAmount / 100).toFixed(2) : '',
            index === 0 ? order.customerName : '',
            index === 0 ? order.customerPhone : '',
            index === 0 ? order.customerEmail : '',

            index === 0 ? order.orderService?.name : '',
            index === 0 ? (order.orderService?.price / 100).toFixed(2) : '',
            index === 0 ? (order.orderService?.scheduleDate ? new Date(order.orderService.scheduleDate).toLocaleString() : '') : '',
            index === 0 ? (order.orderService?.scheduleTime ? new Date(order.orderService.scheduleTime).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '') : '', // Format time in hours
            index === 0 ? (order.orderService?.scheduleTimeslotStart ? `${new Date(order.orderService.scheduleTimeslotStart).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })} ~ ${new Date(order.orderService.scheduleTimeslotEnd).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}` : '') : '', // Format timeslot

            index === 0 ? (order.orderService?.location?.address || '') : '',
            index === 0 ? (order.orderService?.location?.address2 || '') : '',
            index === 0 ? (order.orderService?.location?.city || '') : '',
            index === 0 ? (order.orderService?.location?.state || '') : '',
            index === 0 ? (order.orderService?.location?.zip || '') : ''
          ];
          sheet.appendRow(row);
        });
      }
    });

    Logger.log('Orders imported successfully.');
  } catch (error) {
    Logger.log('Error fetching orders: ' + error.message);
  }
}

/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Import Orders', 'importOrders')
      .addToUi();
}
