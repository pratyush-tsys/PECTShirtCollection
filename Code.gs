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
  const busSPOCIdx = headers.indexOf('BusSPOCNumber');
  const nameIdx = 1; // Name column index (assuming it's column B)

  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] + '').toLowerCase() === (employeeId + '').toLowerCase()) {
      const collected = data[i][2];
      const time = data[i][3];
      const timeStr = time
        ? Utilities.formatDate(new Date(time), tz, 'dd-MMM-yyyy hh:mm a')
        : '';
      
      // Get room partners - find all employees with same room number excluding current employee
      const currentRoom = roomIdx !== -1 ? data[i][roomIdx] : '';
      const currentName = data[i][nameIdx];
      const roomPartners = [];
      
      if (currentRoom && currentRoom !== '') {
        for (let j = 1; j < data.length; j++) {
          if (j !== i && // Not the current employee
              roomIdx !== -1 && 
              data[j][roomIdx] === currentRoom && // Same room number
              data[j][nameIdx] && // Has a name
              data[j][nameIdx] !== '') { // Name is not empty
            roomPartners.push(data[j][nameIdx]);
          }
        }
      }
      
      return {
        row: i + 1,
        name: data[i][1],
        collected: collected,
        timeStr: timeStr,
        busNumber: busIdx !== -1 ? data[i][busIdx] : '',
        roomNumber: roomIdx !== -1 ? data[i][roomIdx] : '',
        employeeId: employeeId,
        busSPOCNumber: busSPOCIdx !== -1 ? data[i][busSPOCIdx] : '',
        roomPartners: roomPartners
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
