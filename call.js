function updateDailyDash() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Lead Entered");
  const targetSheet = ss.getSheetByName("Daily Dash");

  // Adjust column index for "Date Of Meeting"
  const dateColIndex = 3;  // Example: column C

  // Get data
  const data = sourceSheet.getDataRange().getValues();

  // Map of date -> counts
  const results = {};

  data.forEach(row => {
    const dateValue = row[dateColIndex - 1];
    if (!dateValue) return;

    const dateStr = Utilities.formatDate(new Date(dateValue), Session.getScriptTimeZone(), "M/d/yyyy");

    if (!results[dateStr]) {
      results[dateStr] = {
        noShows: {
          respondedFirstText: 0,
          respondedMorningText: 0
        },
        showedCalls: {
          respondedFirstText: 0,
          respondedMorningText: 0
        },
        total: {
          totalCalls: 0,
          rescheduled: 0,
          canceled: 0
        }
      };
    }

    // Placeholder: increment total calls for now
    results[dateStr].total.totalCalls++;
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

  // Center-align first row (parent headers)
  targetSheet.getRange(1, 1, 2, 8).setHorizontalAlignment("center").setVerticalAlignment("middle");
  targetSheet.getRange(1, 1, 1, 8).setHorizontalAlignment("center").setVerticalAlignment("middle");
  targetSheet.getRange(2, 1, 1, 8).setHorizontalAlignment("center").setVerticalAlignment("middle");

  const columnWidths = [100, 150, 180, 150, 180, 150, 110, 100];
  columnWidths.forEach((width, i) => {
    targetSheet.setColumnWidth(i + 1, width);
  });

  // Write data rows starting from row 3
  Object.keys(results).sort((a, b) => new Date(a) - new Date(b)).forEach(date => {
    const r = results[date];
    targetSheet.appendRow([
      date,
      r.noShows.respondedFirstText,
      r.noShows.respondedMorningText,
      r.showedCalls.respondedFirstText,
      r.showedCalls.respondedMorningText,
      r.total.totalCalls,
      r.total.rescheduled,
      r.total.canceled
    ]);
  });
}

// Optional: Add a menu to run the script easily
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Daily Dash Tools')
    .addItem('Update Daily Dash', 'updateDailyDash')
    .addToUi();
}