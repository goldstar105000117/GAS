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

  // Fetch Low Ticket Revenue data
  const lowTicketSpreadsheetId = '1fZwddCvFHV1e3yuP_BjEctZHlE5Q6GC5yBrBGDHjT2U';
  const lowTicketSheet = SpreadsheetApp.openById(lowTicketSpreadsheetId);
  const lowTicketTab = lowTicketSheet.getSheetByName('ALL DATA VIEW');
  
  if (!lowTicketTab) {
    SpreadsheetApp.getUi().alert('External sheet "ALL DATA VIEW" not found.');
    return;
  }

  // Fetch High Ticket Revenue data
  const highTicketSpreadsheetId = '10tbO1W5qC3X7vY_EbNp6nMf6E7BsvrrFrgvpL15HBLw';
  const highTicketSheet = SpreadsheetApp.openById(highTicketSpreadsheetId);
  const highTicketTab = highTicketSheet.getSheetByName('Post Call Reports');
  
  if (!highTicketTab) {
    SpreadsheetApp.getUi().alert('External sheet "Post Call Reports" not found.');
    return;
  }

  // Get Low Ticket data (columns A and E)
  const lowTicketLastRow = lowTicketTab.getLastRow();
  if (lowTicketLastRow < 2) {
    SpreadsheetApp.getUi().alert('No data found in Low Ticket sheet.');
    return;
  }
  
  const lowTicketDates = lowTicketTab.getRange(2, 1, lowTicketLastRow - 1, 1).getValues();
  const lowTicketRevenues = lowTicketTab.getRange(2, 5, lowTicketLastRow - 1, 1).getValues();

  // Get High Ticket data (columns A, I, and L)
  const highTicketLastRow = highTicketTab.getLastRow();
  if (highTicketLastRow < 2) {
    SpreadsheetApp.getUi().alert('No data found in High Ticket sheet.');
    return;
  }
  
  const highTicketValues = highTicketTab.getRange(2, 1, highTicketLastRow - 1, 12).getValues();

  // Valid statuses for High Ticket Revenue (normalized)
  const validStatuses = [
    'won-innercircle(paymentplan)',
    'won-lifetime(paymentplan)',
    'won-innercircle(cashpif)',
    'won-lifetime(cashpif)',
    'won-innercircle(financed)',
    'won-lifetime(financed)',
    'marker-innercircle',
    'marker-lifetime'
  ];

  // Group Low Ticket revenue by month/year
  const lowTicketMonthlyRevenue = {};
  
  for (let i = 0; i < lowTicketDates.length; i++) {
    const date = new Date(lowTicketDates[i][0]);
    if (isNaN(date.getTime())) continue;
    
    const month = date.getMonth();
    const year = date.getFullYear();
    const monthKey = `${year}-${String(month + 1).padStart(2, '0')}`;
    
    if (!lowTicketMonthlyRevenue[monthKey]) {
      lowTicketMonthlyRevenue[monthKey] = 0;
    }
    
    const revenue = parseFloat(lowTicketRevenues[i][0]) || 0;
    lowTicketMonthlyRevenue[monthKey] += revenue;
  }

  // Group High Ticket revenue by month/year
  const highTicketMonthlyRevenue = {};
  
  for (let i = 0; i < highTicketValues.length; i++) {
    const date = new Date(highTicketValues[i][0]); // Column A
    const status = cleanStatusString(highTicketValues[i][8]); // Column I
    const cashCollected = parseFloat(highTicketValues[i][11]) || 0; // Column L
    
    if (isNaN(date.getTime()) || !validStatuses.includes(status)) continue;
    
    const month = date.getMonth();
    const year = date.getFullYear();
    const monthKey = `${year}-${String(month + 1).padStart(2, '0')}`;
    
    if (!highTicketMonthlyRevenue[monthKey]) {
      highTicketMonthlyRevenue[monthKey] = 0;
    }
    
    highTicketMonthlyRevenue[monthKey] += cashCollected;
  }

  // Get all unique months from both datasets
  const allMonthKeys = new Set([
    ...Object.keys(lowTicketMonthlyRevenue),
    ...Object.keys(highTicketMonthlyRevenue)
  ]);
  
  const sortedMonths = Array.from(allMonthKeys).sort();
  const dataRows = [];
  
  for (const monthKey of sortedMonths) {
    const [year, month] = monthKey.split('-');
    const date = new Date(year, month - 1, 1);
    const monthName = date.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    
    const lowTicketRev = lowTicketMonthlyRevenue[monthKey] || 0;
    const highTicketRev = highTicketMonthlyRevenue[monthKey] || 0;
    
    dataRows.push([
      monthName,                    // Month
      '',                           // Total Clipping Spend (to be filled manually)
      '',                           // Total Clipping Views (to be filled manually)
      lowTicketRev,                 // New Low Ticket Revenue
      highTicketRev,                // New High Ticket Revenue
      ''                            // ROAs (to be calculated)
    ]);
  }

  // Write data to sheet
  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, 6).setValues(dataRows);
    
    // Format revenue columns as currency
    sheet.getRange(2, 4, dataRows.length, 1).setNumberFormat('$#,##0.00'); // Low Ticket
    sheet.getRange(2, 5, dataRows.length, 1).setNumberFormat('$#,##0.00'); // High Ticket
  }

  SpreadsheetApp.getUi().alert(`Clipping Stats generated successfully!\n\nFound ${dataRows.length} months of data.`);
}

// Helper function to clean and normalize status strings
function cleanStatusString(value) {
  if (!value) return '';
  return value.toString()
    .toLowerCase()
    .trim()
    .replace(/\s+/g, '');
}