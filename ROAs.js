function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Actions')
    .addItem('Generate Clipping Stats', 'generateClippingStats')
    .addToUi();
}

function generateClippingStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Clipping Stats");

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet "Clipping Stats" not found. Please create it first.');
    return;
  }

  // Generate headers
  const headers = [
    "Month",
    "Total Clipping Spend",
    "Total Clipping Views",
    "New Low Ticket Revenue",
    "New High Ticket Revenue",
    "ROAs"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#4285f4");
  headerRange.setFontColor("#ffffff");
  headerRange.setHorizontalAlignment("center");

  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  // Get the external spreadsheet with low ticket revenue data
  const externalSpreadsheetId = '1fZwddCvFHV1e3yuP_BjEctZHlE5Q6GC5yBrBGDHjT2U';
  const externalSs = SpreadsheetApp.openById(externalSpreadsheetId);
  const allDataSheet = externalSs.getSheetByName('ALL DATA VIEW');

  if (!allDataSheet) {
    SpreadsheetApp.getUi().alert('External sheet "ALL DATA VIEW" not found.');
    return;
  }

  // Get all data from columns A (Date) and E (Revenue)
  const lastRow = allDataSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No data found in external sheet.');
    return;
  }

  const dates = allDataSheet.getRange(2, 1, lastRow - 1, 1).getValues(); // Column A
  const revenues = allDataSheet.getRange(2, 5, lastRow - 1, 1).getValues(); // Column E

  // Group revenue by month/year
  const monthlyRevenue = {};

  for (let i = 0; i < dates.length; i++) {
    const date = new Date(dates[i][0]);
    if (isNaN(date.getTime())) continue; // Skip invalid dates

    const month = date.getMonth();
    const year = date.getFullYear();
    const monthKey = `${year}-${String(month + 1).padStart(2, '0')}`; // Format: "2025-01"

    if (!monthlyRevenue[monthKey]) {
      monthlyRevenue[monthKey] = 0;
    }

    const revenue = parseFloat(revenues[i][0]) || 0;
    monthlyRevenue[monthKey] += revenue;
  }

  // Sort months and prepare data rows
  const sortedMonths = Object.keys(monthlyRevenue).sort();
  const dataRows = [];

  for (const monthKey of sortedMonths) {
    const [year, month] = monthKey.split('-');
    const date = new Date(year, month - 1, 1);
    const monthName = date.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });

    dataRows.push([
      monthName,                    // Month
      '',                           // Total Clipping Spend (to be filled manually)
      '',                           // Total Clipping Views (to be filled manually)
      monthlyRevenue[monthKey],     // New Low Ticket Revenue
      '',                           // New High Ticket Revenue (to be filled manually)
      ''                            // ROAs (to be calculated)
    ]);
  }

  // Write data to sheet
  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, 6).setValues(dataRows);

    // Format revenue column as currency
    sheet.getRange(2, 4, dataRows.length, 1).setNumberFormat('$#,##0.00');
  }

  SpreadsheetApp.getUi().alert(`Clipping Stats generated successfully!\n\nFound ${dataRows.length} months of data.`);
}