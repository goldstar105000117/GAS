function updateDailyDash() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Lead Entered");
  const targetSheet = ss.getSheetByName("Daily Dash");

  // Column indices (0-based)
  const dateColIndex = 0;  // Column A - Date Of Meeting
  const columnF_Index = 5;  // Column F - Responded to 1st Text
  const columnQ_Index = 16; // Column Q - Responded to Morning Text
  const columnV_Index = 21; // Column V - Status

  // Get all data from source sheet
  const data = sourceSheet.getDataRange().getValues();

  // Object to store counts by date
  const dateStats = {};

  // Process data starting from row 2 (skip header)
  for (let i = 1; i < data.length; i++) {
    const dateValue = data[i][dateColIndex];
    
    // Skip empty dates
    if (!dateValue) continue;

    try {
      // Convert to date and format consistently
      const date = new Date(dateValue);
      const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "M/d/yyyy");
      
      // Initialize date stats if not exists
      if (!dateStats[dateStr]) {
        dateStats[dateStr] = {
          noShowRespondedFirst: 0,
          noShowRespondedMorning: 0,
          showedRespondedFirst: 0,
          showedRespondedMorning: 0
        };
      }

      // Get column values
      const columnF = String(data[i][columnF_Index]).toLowerCase().trim();
      const columnQ = String(data[i][columnQ_Index]).toLowerCase().trim();
      const columnV = String(data[i][columnV_Index]).replace(/\s+/g, '').toLowerCase();

      // Check if status is a "showed" status (contains won-, lost-, marker-, or followup)
      const isShowedStatus = columnV.includes("won-") || 
                            columnV.includes("lost-") || 
                            columnV.includes("marker-") || 
                            columnV.includes("followup");

      // Check for No Show - Responded to 1st Text
      // Column F is "yes" AND Column V is "noshow"
      if (columnF === "yes" && columnV === "noshow") {
        dateStats[dateStr].noShowRespondedFirst++;
      }

      // Check for No Show - Responded to Morning Text
      // Column Q is "yes" AND Column V is "noshow"
      if (columnQ === "yes" && columnV === "noshow") {
        dateStats[dateStr].noShowRespondedMorning++;
      }

      // Check for Showed Calls - Responded to 1st Text
      // Column F is "yes" AND Column V contains won-, lost-, marker-, or followup
      if (columnF === "yes" && isShowedStatus) {
        dateStats[dateStr].showedRespondedFirst++;
      }

      // Check for Showed Calls - Responded to Morning Text
      // Column Q is "yes" AND Column V contains won-, lost-, marker-, or followup
      if (columnQ === "yes" && isShowedStatus) {
        dateStats[dateStr].showedRespondedMorning++;
      }

    } catch (e) {
      // Skip invalid dates
      continue;
    }
  }

  // Sort dates
  const sortedDates = Object.keys(dateStats).sort((a, b) => {
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
  sortedDates.forEach(date => {
    const stats = dateStats[date];
    targetSheet.appendRow([
      date,
      stats.noShowRespondedFirst,      // No Shows - Responded to 1st Text
      stats.noShowRespondedMorning,    // No Shows - Responded to Morning Text
      stats.showedRespondedFirst,      // Showed Calls - Responded to 1st Text
      stats.showedRespondedMorning,    // Showed Calls - Responded to Morning Text
      0,  // Total Calls for the Day
      0,  // Rescheduled
      0   // Canceled
    ]);
  });

  // Add borders
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