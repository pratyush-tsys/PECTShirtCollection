function doGet() {
  return HtmlService.createHtmlOutputFromFile('Page');
}

// Checks if the EmployeeID exists in the sheet and returns details including BusNumber and RoomNumber
function checkEmployeeId(employeeId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';

  // Get header row to find column indices for BusNumber and RoomNumber
  const headers = data[0];
  const busIdx = headers.indexOf('BusNumber');
  const roomIdx = headers.indexOf('RoomNumber');
  const ticketIdx = headers.indexOf('TicketNumber');

  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] + '').toLowerCase() === (employeeId + '').toLowerCase()) {
      const collected = data[i][2];
      const time = data[i][3];
      const timeStr = time
        ? Utilities.formatDate(new Date(time), tz, 'dd-MMM-yyyy hh:mm a')
        : '';
      return {
        row: i + 1,
        name: data[i][1],
        collected: collected,
        timeStr: timeStr,
        busNumber: busIdx !== -1 ? data[i][busIdx] : '',
        roomNumber: roomIdx !== -1 ? data[i][roomIdx] : '',
        employeeId: employeeId,
        ticketNumber: ticketIdx !== -1 ? data[i][ticketIdx] : ''
      };
    }
  }
  return null;
}

function updateStatus(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const current = sheet.getRange(row, 3).getValue();
  if (current === 'Yes') {
    return "Already marked ✅ You cannot update again.";
  }
  sheet.getRange(row, 3).setValue('Yes');                // Shirt Collected
  sheet.getRange(row, 4).setValue(new Date());           // Timestamp
  return 'Updated Successfully ✅';
}
