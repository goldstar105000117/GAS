function updateDailyDash() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leadSheet = ss.getSheetByName('Lead Entered');
  const dashSheet = ss.getSheetByName('Daily Dash');

  // Get all data from Lead Entered tab
  const leadData = leadSheet.getDataRange().getValues();
  const headers = leadData[0];

  // Find the column index for "Date Of Meeting"
  const dateColIndex = headers.indexOf('Date Of Meeting');

  if (dateColIndex === -1) {
    SpreadsheetApp.getUi().alert('Error: "Date Of Meeting" column not found in Lead Entered sheet');
    return;
  }

  // Prepare output data with headers
  const outputData = [headers];

  // Loop through data rows (skip header row)
  for (let i = 1; i < leadData.length; i++) {
    const row = leadData[i];
    const dateValue = row[dateColIndex];

    // Check if date exists and is valid
    if (dateValue && dateValue !== '') {
      outputData.push(row);
    }
  }

  // Clear existing data in Daily Dash (except headers if you want to keep formatting)
  dashSheet.clear();

  // Write data to Daily Dash
  if (outputData.length > 0) {
    dashSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
  }

  SpreadsheetApp.getUi().alert('Daily Dash updated successfully with ' + (outputData.length - 1) + ' records!');
}

// Optional: Add a menu to run the script easily
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Daily Dash Tools')
    .addItem('Update Daily Dash', 'updateDailyDash')
    .addToUi();
}