function updateDailyDash() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Lead Entered");
  const targetSheet = ss.getSheetByName("Daily Dash");

  // Adjust column index for "Date Of Meeting"
  const dateColIndex = 1;  // Column A

  // Get all data from source sheet
  const data = sourceSheet.getDataRange().getValues();

  // Skip header row and collect all unique dates
  const uniqueDates = new Set();

  for (let i = 1; i < data.length; i++) {
    const dateValue = data[i][dateColIndex - 1];

    // Skip empty dates
    if (!dateValue) continue;

    try {
      // Convert to date and format consistently
      const date = new Date(dateValue);
      const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "M/d/yyyy");
      uniqueDates.add(dateStr);
    } catch (e) {
      // Skip invalid dates
      continue;
    }
  }

  // Convert Set to Array and sort by date
  const sortedDates = Array.from(uniqueDates).sort((a, b) => {
    return new Date(a) - new Date(b);
  });

  // Clear target sheet
  targetSheet.clear();

  // Create two header rows
  const headerRow1 = ["Date", "No Shows", "No Shows", "Showed Calls", "Showed Calls", "Total", "Total", "Total"];
  const headerRow2 = ["", "Responded to 1st Text", "Responded to Morning Text", "Responded to 1st Text", "Responded to Morning Text", "Total Calls for the Day", "Rescheduled", "Canceled"];

  targetSheet.appendRow(headerRow1);
  targetSheet.appendRow(headerRow2);

  // Merge parent headers
  targetSheet.getRange(1, 1, 2, 1).merge(); // Merge "Date" across 2 rows
  targetSheet.getRange(1, 2, 1, 2).merge(); // No Shows
  targetSheet.getRange(1, 4, 1, 2).merge(); // Showed Calls
  targetSheet.getRange(1, 6, 1, 3).merge(); // Total

  // Center-align and style headers
  targetSheet.getRange(1, 1, 2, 8).setHorizontalAlignment("center").setVerticalAlignment("middle");
  targetSheet.getRange(1, 1, 2, 8).setFontWeight("bold");

  // Set column widths
  const columnWidths = [100, 150, 180, 150, 180, 150, 110, 100];
  columnWidths.forEach((width, i) => {
    targetSheet.setColumnWidth(i + 1, width);
  });

  // Write data rows starting from row 3
  // All values are 0 for now
  sortedDates.forEach(date => {
    targetSheet.appendRow([
      date,
      0,  // No Shows - Responded to 1st Text
      0,  // No Shows - Responded to Morning Text
      0,  // Showed Calls - Responded to 1st Text
      0,  // Showed Calls - Responded to Morning Text
      0,  // Total Calls for the Day
      0,  // Rescheduled
      0   // Canceled
    ]);
  });

  if (sortedDates.length > 0) {
    targetSheet.getRange(1, 1, sortedDates.length + 2, 8).setBorder(true, true, true, true, true, true);
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Daily Dash Tools')
    .addItem('Update Daily Dash', 'updateDailyDash')
    .addToUi();
}