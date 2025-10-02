const CONFIG = {
  DATA_SHEET_ID: '1LRaCtgmYXSNKSfi4pP-ftkdY_9PDHf41AsSc3gGBR5c',
  DATA_TAB_NAME: 'Helper_long',
  PICKUP_SHEET_ID: '1iGKSUlWNZ-A0jI40DRxW97ucutCd9OwFr4IuK7omDsU',
  PICKUP_TAB_NAME: 'Meaningful Call Logs',
  SETS_SHEET_ID: '10tbO1W5qC3X7vY_EbNp6nMf6E7BsvrrFrgvpL15HBLw',
  SETS_TAB_NAME: 'Post Call Reports',
  SETTERS: [
    'Alan Kamuramira',
    'Angel Flores',
    'Brandon Parker',
    'Christopher Wallace',
    'Ethen Kaswell',
    'Jayson Pablo',
    'Liam Burgoyne',
    'Noah Sluder'
  ],
  PERIODS: {
    last7: 7,
    last14: 14,
    last30: 30
  },

  THRESHOLDS: {
    DCC_RATE: 30,    // DCC% >= 30% = green
    SHOW_RATE: 65,   // Show% >= 65% = green
    CLOSE_RATE: 30   // Close% >= 30% = green
  },

  // Add this new configuration for Low Ticket Buys
  LOW_TICKET_SHEET_ID: '1fZwddCvFHV1e3yuP_BjEctZHlE5Q6GC5yBrBGDHjT2U',
  LOW_TICKET_TAB_NAME: 'OVERALL PERFORMANCE',
  PARTIAL_TRIAGE_SHEET_ID: '1LRaCtgmYXSNKSfi4pP-ftkdY_9PDHf41AsSc3gGBR5c',
  PARTIAL_TRIAGE_TAB_NAME: 'Partial Submits',
  BOOKED_TRIAGE_SHEET_ID: '1LRaCtgmYXSNKSfi4pP-ftkdY_9PDHf41AsSc3gGBR5c',
  BOOKED_TRIAGE_TAB_NAME: 'Calendly Booked Calls',
};

const CLOSER_CONFIG = {
  CLOSERS: [
    'Alex El-H',
    'Easton Tomak',
    'Carmen Pinto'
  ]
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Setter Study')
    .addItem('Generate Report by Days', 'generateReportAsync')
    .addItem('Generate Report by Setter', 'generateSetterStudyBySetter')
    .addItem('Generate Closer Study by Days', 'generateCloserStudyByDays')
    .addItem('Generate Closer Study by Closer', 'generateCloserStudyByCloser')
    .addItem('Generate Vortex Data', 'generateVortexData')
    .addItem('Generate Vortex Data Weekly', 'generateVortexDataWeekly')
    .addItem('Generate Vortex Data Monthly', 'generateVortexDataMonthly')
    .addToUi();
}

/**
 * Get September 5th, 2025 as baseline date in New York timezone
 */
function getBaselineDate() {
  return new Date('2025-09-05T00:00:00');
}

/**
 * Main function to generate the setter study report (BY DAYS)
 */
function generateReportAsync() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    showLoadingMessage(ui);

    let sheet = getOrCreateSheet(spreadsheet, 'Setter Study by Days');
    sheet.clear();

    // Fetch all data in parallel for better performance
    const [rawData, pickupData, setsData, outcomesData] = fetchAllDataAsync();

    // Process data for all periods (now using September 5th baseline for 30-day)
    const processedData = {
      dials: processDialsData(rawData),
      pickups: processPickupsData(pickupData),
      sets: processSetsData(setsData),
      outcomes: processOutcomesData(outcomesData)
    };

    // Generate report
    generateReport(sheet, processedData);

    showSuccessMessage(ui);

  } catch (error) {
    showErrorMessage(ui, error);
    console.error('Report generation failed:', error);
  }
}

/**
 * Generate Setter Study report grouped by Setter (instead of by time periods)
 */
function generateSetterStudyBySetter() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    showLoadingMessage(ui);

    let sheet = getOrCreateSheet(spreadsheet, 'Setter Study by Setter');
    sheet.clear();

    // Fetch all data in parallel for better performance
    const [rawData, pickupData, setsData, outcomesData] = fetchAllDataAsync();

    // Process data for all periods (now using September 5th baseline for 30-day)
    const processedData = {
      dials: processDialsData(rawData),
      pickups: processPickupsData(pickupData),
      sets: processSetsData(setsData),
      outcomes: processOutcomesData(outcomesData)
    };

    // Process data for totals (now starting from September 5th)
    const allTimeData = {
      dials: processAllTimeDialsData(rawData),
      pickups: processAllTimePickupsData(pickupData),
      sets: processAllTimeSetsData(setsData),
      outcomes: processAllTimeOutcomesData(outcomesData)
    };

    // Generate report grouped by setter
    generateSetterGroupedReport(sheet, processedData, allTimeData);

    showSuccessMessage(ui);

  } catch (error) {
    showErrorMessage(ui, error);
    console.error('Setter Study by Setter generation failed:', error);
  }
}

function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    // Sheet doesn't exist, create it
    sheet = spreadsheet.insertSheet(sheetName);
  }

  return sheet;
}

/**
 * Creates an hourly trigger to update both report types automatically
 * Run this function once to set up the automation
 */
function setupHourlyTrigger() {
  try {
    // Delete any existing triggers first to avoid duplicates
    deleteAllTriggers();

    // Create hourly trigger
    ScriptApp.newTrigger('updateReportByDays')
      .timeBased()
      .everyHours(1)
      .create();

    ScriptApp.newTrigger('updateReportBySetter')
      .timeBased()
      .everyHours(1)
      .create();

    ScriptApp.newTrigger('updateCloserReportByDays')
      .timeBased()
      .everyHours(1)
      .create();

    ScriptApp.newTrigger('updateCloserReportByCloser')
      .timeBased()
      .everyHours(1)
      .create();

    ScriptApp.newTrigger('updateVortexData')
      .timeBased()
      .everyHours(1)
      .create();

    ScriptApp.newTrigger('updateVortexDataWeekly')
      .timeBased()
      .everyHours(1)
      .create();

    ScriptApp.newTrigger('updateVortexDataMonthly')
      .timeBased()
      .everyHours(1)
      .create();

    console.log('Hourly trigger created successfully');

    // Show confirmation to user if run manually
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Trigger Setup', 'Hourly trigger has been created successfully! Both reports will now update every hour automatically.', ui.ButtonSet.OK);
    } catch (e) {
      // Ignore UI errors if running from trigger
    }

  } catch (error) {
    console.error('Failed to create hourly trigger:', error);
    throw error;
  }
}

/**
 * Update the "By Days" report on the first sheet
 */
function updateReportByDays() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // Create or get the "By Days" sheet
    let sheet = spreadsheet.getSheetByName('Setter Study by Days');
    if (!sheet) {
      sheet = spreadsheet.insertSheet('Setter Study by Days', 0);
    } else {
      spreadsheet.setActiveSheet(sheet);
    }

    console.log('Updating "By Days" report with smart September 5th baseline...');

    // Clear existing data
    sheet.clear();

    // Fetch and process data
    const [rawData, pickupData, setsData, outcomesData] = fetchAllDataAsync();
    const processedData = {
      dials: processDialsData(rawData),
      pickups: processPickupsData(pickupData),
      sets: processSetsData(setsData),
      outcomes: processOutcomesData(outcomesData)
    };

    // Generate report
    generateReport(sheet, processedData);

    console.log('"By Days" report updated successfully with smart September 5th baseline');

  } catch (error) {
    console.error('Failed to update "By Days" report:', error);
    throw error;
  }
}

/**
 * Update the "By Setter" report on the second sheet
 */
function updateReportBySetter() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // Create or get the "By Setter" sheet
    let sheet = spreadsheet.getSheetByName('Setter Study by Setter');
    if (!sheet) {
      sheet = spreadsheet.insertSheet('Setter Study by Setter', 1);
    } else {
      spreadsheet.setActiveSheet(sheet);
    }

    console.log('Updating "By Setter" report with smart September 5th baseline...');

    // Clear existing data
    sheet.clear();

    // Fetch and process data
    const [rawData, pickupData, setsData, outcomesData] = fetchAllDataAsync();
    const processedData = {
      dials: processDialsData(rawData),
      pickups: processPickupsData(pickupData),
      sets: processSetsData(setsData),
      outcomes: processOutcomesData(outcomesData)
    };

    const allTimeData = {
      dials: processAllTimeDialsData(rawData),
      pickups: processAllTimePickupsData(pickupData),
      sets: processAllTimeSetsData(setsData),
      outcomes: processAllTimeOutcomesData(outcomesData)
    };

    // Generate setter-grouped report
    generateSetterGroupedReport(sheet, processedData, allTimeData);

    console.log('"By Setter" report updated successfully with smart September 5th baseline');

  } catch (error) {
    console.error('Failed to update "By Setter" report:', error);
    throw error;
  }
}

/**
 * Delete all existing triggers to avoid duplicates
 */
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runHourlyUpdates') {
      ScriptApp.deleteTrigger(trigger);
      console.log('Deleted existing trigger');
    }
  });
}

/**
 * Remove the hourly trigger (run this if you want to stop automatic updates)
 */
function removeHourlyTrigger() {
  try {
    deleteAllTriggers();
    console.log('All hourly triggers removed');

    // Show confirmation to user
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Trigger Removed', 'Hourly trigger has been removed. Automatic updates are now disabled.', ui.ButtonSet.OK);
    } catch (e) {
      // Ignore UI errors
    }

  } catch (error) {
    console.error('Failed to remove triggers:', error);
    throw error;
  }
}

/**
 * Fetch data from all sheets in parallel
 */
function fetchAllDataAsync() {
  const promises = [
    () => fetchSheetData(CONFIG.DATA_SHEET_ID, CONFIG.DATA_TAB_NAME, parseDialsData),
    () => fetchSheetData(CONFIG.PICKUP_SHEET_ID, CONFIG.PICKUP_TAB_NAME, parsePickupData),
    () => fetchSheetData(CONFIG.SETS_SHEET_ID, CONFIG.SETS_TAB_NAME, parseSetsData),
    () => fetchSheetData(CONFIG.SETS_SHEET_ID, CONFIG.SETS_TAB_NAME, parseOutcomesData)
  ];

  // Execute all data fetching operations
  return promises.map(fn => fn());
}

/**
 * Generic sheet data fetcher with error handling
 */
function fetchSheetData(sheetId, tabName, parser) {
  try {
    const sheet = SpreadsheetApp.openById(sheetId);
    const tab = sheet.getSheetByName(tabName);

    if (!tab) {
      throw new Error(`Tab "${tabName}" not found`);
    }

    const values = tab.getDataRange().getValues();
    if (values.length <= 1) {
      throw new Error(`No data found in ${tabName}`);
    }

    return parser(values);

  } catch (error) {
    throw new Error(`Failed to fetch data from ${tabName}: ${error.message}`);
  }
}

/**
 * Parse dials data from Helper_long sheet
 * Expected columns: Date, Setter First Name, Number of Dials
 */
function parseDialsData(values) {
  const [headers, ...data] = values;

  // Find column indices with flexible matching
  const dateIndex = findColumnIndex(headers, ['date']);
  const setterIndex = findColumnIndex(headers, ['setter first name', 'setter', 'first name']);
  const dialsIndex = findColumnIndex(headers, ['number of dials', 'dials']);

  console.log(`Dials - Date: ${dateIndex}, Setter: ${setterIndex}, Dials: ${dialsIndex}`);

  return data
    .map(row => ({
      date: parseDate(row[dateIndex]),
      setter: cleanString(row[setterIndex]),
      dials: parseInt(row[dialsIndex]) || 0
    }))
    .filter(row => row.date && row.setter && row.dials >= 0);
}

/**
 * Parse pickup data from Meaningful Call Logs sheet
 * Expected columns: date, time_est, setter
 * Logic: Count rows per setter (each row = 1 pickup)
 */
function parsePickupData(values) {
  const [headers, ...data] = values;

  const dateIndex = findColumnIndex(headers, ['date']);
  const setterIndex = findColumnIndex(headers, ['setter']);

  console.log(`Pickups - Date: ${dateIndex}, Setter: ${setterIndex}`);

  return data
    .map(row => ({
      date: parseDate(row[dateIndex]),
      setter: cleanString(row[setterIndex])
    }))
    .filter(row => row.date && row.setter);
}

/**
 * Parse sets data from Post Call Reports sheet
 * Expected columns: Date Submitted, Setter, What's the new Lead Status
 * Filter: Status = "Discovery Call Complete"
 */
function parseSetsData(values) {
  const [headers, ...data] = values;

  const dateIndex = findColumnIndex(headers, ['date submitted']);
  const setterIndex = findColumnIndex(headers, ['setter']);
  const statusIndex = findColumnIndex(headers, ["what's the new lead status"]);

  console.log(`Sets - Date: ${dateIndex}, Setter: ${setterIndex}, Status: ${statusIndex}`);

  return data
    .map(row => ({
      date: parseDate(row[dateIndex]),
      setter: cleanString(row[setterIndex]),
      status: cleanStatusString(row[statusIndex])
    }))
    .filter(row =>
      row.date &&
      row.setter &&
      row.status === 'discoverycallcomplete'
    );
}

/**
 * Parse outcomes data (Won/Lost) from Post Call Reports sheet
 * Expected columns: Date Submitted, Setter, What's the new Lead Status
 * Filter: Status includes "won" or "lost"
 */
function parseOutcomesData(values) {
  const [headers, ...data] = values;

  const dateIndex = findColumnIndex(headers, ['date submitted']);
  const setterIndex = findColumnIndex(headers, ['setter']);
  const statusIndex = findColumnIndex(headers, ["what's the new lead status", 'lead status', 'status']);

  console.log(`Outcomes - Date: ${dateIndex}, Setter: ${setterIndex}, Status: ${statusIndex}`);

  return data
    .map(row => ({
      date: parseDate(row[dateIndex]),
      setter: cleanString(row[setterIndex]),
      status: cleanStatusString(row[statusIndex])
    }))
    .filter(row => {
      if (!row.date || !row.setter || !row.status) return false;

      // Include both won and lost statuses
      return row.status.includes('won') || row.status.includes('lost');
    });
}

/**
 * Find column index with flexible matching
 */
function findColumnIndex(headers, keywords) {
  for (const keyword of keywords) {
    const index = headers.findIndex(h =>
      cleanString(h).includes(keyword.toLowerCase())
    );

    if (index !== -1) {
      return index;
    }
  }

  throw new Error(`Column not found. Looking for: ${keywords.join(' or ')}`);
}

/**
 * Clean and normalize string values
 */
function cleanString(value) {
  if (!value) return '';
  return value.toString().toLowerCase().trim();
}

function cleanStatusString(value) {
  if (!value) return '';
  return value.toString()
    .toLowerCase()
    .trim()
    .replace(/\s+/g, '');
}

/**
 * Enhanced date parser with better error handling
 * Handles formats like 7/14/2025 and 9/15/2025
 */
function parseDate(dateValue) {
  if (!dateValue) return null;
  if (dateValue instanceof Date) {
    // If it's already a Date object, normalize it to NY timezone
    const dateString = Utilities.formatDate(dateValue, 'America/New_York', 'yyyy-MM-dd');
    return new Date(dateString + 'T00:00:00');
  }

  if (typeof dateValue === 'string') {
    const parts = dateValue.split('/');
    if (parts.length === 3) {
      const [month, day, year] = parts.map(p => parseInt(p));
      // Create date and normalize to NY timezone
      const tempDate = new Date(year, month - 1, day);
      const dateString = Utilities.formatDate(tempDate, 'America/New_York', 'yyyy-MM-dd');
      return new Date(dateString + 'T00:00:00');
    }
  }

  const parsed = new Date(dateValue);
  if (isNaN(parsed.getTime())) return null;

  // Normalize parsed date to NY timezone
  const dateString = Utilities.formatDate(parsed, 'America/New_York', 'yyyy-MM-dd');
  return new Date(dateString + 'T00:00:00');
}

/**
 * Match setter names with flexibility for partial matches
 */
function matchSetter(dataValue, configSetter) {
  const cleanData = cleanString(dataValue);
  const cleanConfig = cleanString(configSetter);

  // Extract first name from config setter for matching
  const firstName = cleanConfig.split(' ')[0];

  return cleanData.includes(firstName) || cleanData.includes(cleanConfig);
}

/**
 * Modified generic data processor for different time periods using NY timezone
 * Uses September 5th baseline for 30-day period only if it's more recent than normal 30-day lookback
 */
function processModifiedDataByPeriod(rawData, aggregator) {
  const today = getCurrentDateNY();
  const baselineDate = getBaselineDate();
  const result = {};

  Object.entries(CONFIG.PERIODS).forEach(([period, days]) => {
    let startDate;

    if (period === 'last30') {
      // Compare September 5th with normal 30-day lookback, use whichever is more recent
      const normalThirtyDaysAgo = getDateNDaysAgoNY(days);
      startDate = baselineDate > normalThirtyDaysAgo ? baselineDate : normalThirtyDaysAgo;
    } else {
      startDate = getDateNDaysAgoNY(days);
    }

    result[period] = {};

    CONFIG.SETTERS.forEach(setter => {
      result[period][setter] = aggregator(rawData, setter, startDate, today);
    });
  });

  return result;
}

/**
 * Process dials data for all time periods
 */
function processDialsData(rawData) {
  return processModifiedDataByPeriod(rawData, (records, setter, startDate, endDate) =>
    records
      .filter(r => matchSetter(r.setter, setter) && r.date >= startDate && r.date <= endDate)
      .reduce((sum, r) => sum + r.dials, 0)
  );
}

/**
 * Process pickups data - count rows per setter
 */
function processPickupsData(rawData) {
  return processModifiedDataByPeriod(rawData, (records, setter, startDate, endDate) =>
    records.filter(r =>
      matchSetter(r.setter, setter) &&
      r.date >= startDate &&
      r.date <= endDate
    ).length
  );
}

/**
 * Process sets data - count Discovery Call Complete rows per setter
 */
function processSetsData(rawData) {
  return processModifiedDataByPeriod(rawData, (records, setter, startDate, endDate) =>
    records.filter(r =>
      matchSetter(r.setter, setter) &&
      r.date >= startDate &&
      r.date <= endDate
    ).length
  );
}

/**
 * Process outcomes data (Won/Lost) for all time periods
 */
function processOutcomesData(rawData) {
  return processModifiedDataByPeriod(rawData, (records, setter, startDate, endDate) => {
    const filteredRecords = records.filter(r =>
      matchSetter(r.setter, setter) &&
      r.date >= startDate &&
      r.date <= endDate
    );

    const won = filteredRecords.filter(r => r.status.includes('won')).length;
    const lost = filteredRecords.filter(r => r.status.includes('lost')).length;

    return { won, lost, total: won + lost };
  });
}

/**
 * Get current date in New York timezone
 */
function getCurrentDateNY() {
  const now = new Date();
  // Convert to NY timezone using Utilities.formatDate
  const nyDateString = Utilities.formatDate(now, 'America/New_York', 'yyyy-MM-dd');
  const nyDate = new Date(nyDateString + 'T00:00:00');
  return nyDate;
}

/**
 * Get date N days ago in New York timezone
 */
function getDateNDaysAgoNY(days) {
  const today = getCurrentDateNY();
  const pastDate = new Date(today.getTime() - (days * 24 * 60 * 60 * 1000));
  return pastDate;
}

/**
 * Process data for all time (no date filtering) - for total calculations
 */
function processAllTimeData(rawData, aggregator) {
  const today = getCurrentDateNY();
  const baselineDate = getBaselineDate();
  const result = {};

  CONFIG.SETTERS.forEach(setter => {
    // For totals, always use September 5th as the baseline (this gives us consistent totals from that date)
    result[setter] = aggregator(rawData, setter, baselineDate, today);
  });

  return result;
}

/**
 * Process dials data for all time (no date filtering)
 */
function processAllTimeDialsData(rawData) {
  return processAllTimeData(rawData, (records, setter, startDate, endDate) =>
    records
      .filter(r => matchSetter(r.setter, setter) && r.date >= startDate && r.date <= endDate)
      .reduce((sum, r) => sum + r.dials, 0)
  );
}

/**
 * Process pickups data for all time (no date filtering)
 */
function processAllTimePickupsData(rawData) {
  return processAllTimeData(rawData, (records, setter, startDate, endDate) =>
    records.filter(r =>
      matchSetter(r.setter, setter) &&
      r.date >= startDate &&
      r.date <= endDate
    ).length
  );
}

/**
 * Process sets data for all time (no date filtering)
 */
function processAllTimeSetsData(rawData) {
  return processAllTimeData(rawData, (records, setter, startDate, endDate) =>
    records.filter(r =>
      matchSetter(r.setter, setter) &&
      r.date >= startDate &&
      r.date <= endDate
    ).length
  );
}

/**
 * Process outcomes data for all time (no date filtering)
 */
function processAllTimeOutcomesData(rawData) {
  return processAllTimeData(rawData, (records, setter, startDate, endDate) => {
    const filteredRecords = records.filter(r =>
      matchSetter(r.setter, setter) &&
      r.date >= startDate &&
      r.date <= endDate
    );

    const won = filteredRecords.filter(r => r.status.includes('won')).length;
    const lost = filteredRecords.filter(r => r.status.includes('lost')).length;

    return { won, lost, total: won + lost };
  });
}

/**
 * Generate the complete report (BY DAYS)
 */
function generateReport(sheet, data) {
  setupHeaders(sheet);

  const periods = ['last7', 'last14', 'last30'];
  const periodLabels = ['Last 7 Days', 'Last 14 Days', 'Last 30 Days'];
  let currentRow = 3;

  periods.forEach((period, index) => {
    populatePeriodData(sheet, currentRow, periodLabels[index], period, data, CONFIG.SETTERS.length + 1);
    currentRow += CONFIG.SETTERS.length + 1;
  });

  formatSheet(sheet);
}

/**
 * Generate the report grouped by setter
 */
function generateSetterGroupedReport(sheet, data, allTimeData) {
  setupSetterGroupedHeaders(sheet);

  let currentRow = 3;
  const periods = ['last7', 'last14', 'last30'];
  const periodLabels = ['Last 7', 'Last 14', 'Last 30'];

  // Process each setter
  CONFIG.SETTERS.forEach((setter, setterIndex) => {
    populateSetterGroupedData(sheet, currentRow, setter, periods, periodLabels, data, allTimeData);
    currentRow += periods.length + 1 + 1;
  });

  // Add team totals section
  populateTeamTotalsSection(sheet, currentRow, periods, periodLabels, data, allTimeData);

  formatSetterGroupedSheet(sheet);
}

/**
 * Setup report headers (BY DAYS)
 */
function setupHeaders(sheet) {
  // Main header
  const mainHeader = sheet.getRange('A1:I1');
  mainHeader.merge()
    .setValue('Setter Study Report')
    .setBackground('#4a90e2')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(14);

  sheet.setColumnWidth(1, 120);  // Period
  sheet.setColumnWidth(2, 120);  // Setter
  sheet.setColumnWidth(3, 90);   // Dials
  sheet.setColumnWidth(4, 100);  // Pickups
  sheet.setColumnWidth(5, 80);   // Sets
  sheet.setColumnWidth(6, 110);  // Pickup Rate
  sheet.setColumnWidth(7, 80);   // DCC%
  sheet.setColumnWidth(8, 80);   // Show%
  sheet.setColumnWidth(9, 80);   // Close%

  // Column headers
  const headers = ['Period', 'Setter', 'Dials', 'Pickups', 'Sets', 'Pickup Rate', 'DCC%', 'Show%', 'Close%'];
  const headerRange = sheet.getRange(2, 1, 1, headers.length);

  headerRange.setValues([headers])
    .setBackground('#d9e2f3')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

/**
 * Setup headers for setter-grouped report
 */
function setupSetterGroupedHeaders(sheet) {
  // Main header
  const mainHeader = sheet.getRange('A1:I1');
  mainHeader.merge()
    .setValue('Setter Study by Setter')
    .setBackground('#4a90e2')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(14);

  // Set column widths first (before setting headers)
  sheet.setColumnWidth(1, 150);  // Setter
  sheet.setColumnWidth(2, 80);   // Days
  sheet.setColumnWidth(3, 90);   // Dials
  sheet.setColumnWidth(4, 100);  // Pickups
  sheet.setColumnWidth(5, 80);   // Sets
  sheet.setColumnWidth(6, 110);  // Pickup Rate
  sheet.setColumnWidth(7, 80);   // DCC%
  sheet.setColumnWidth(8, 80);   // Show%
  sheet.setColumnWidth(9, 80);   // Close%

  // Column headers
  const headers = ['Setter', 'Days', 'Dials', 'Pickups', 'Sets', 'Pickup Rate', 'DCC%', 'Show%', 'Close%'];
  const headerRange = sheet.getRange(2, 1, 1, headers.length);

  headerRange.setValues([headers])
    .setBackground('#d9e2f3')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

/**
 * Helper function to parse percentage value from string
 */
function parsePercentage(percentageString) {
  if (!percentageString || typeof percentageString !== 'string') return 0;
  return parseInt(percentageString.replace('%', '')) || 0;
}

/**
 * Helper function to apply conditional formatting to performance metrics
 */
function applyConditionalFormatting(sheet, row, dccValue, showValue, closeValue) {
  const dccCol = 7;   // Column G (DCC%)
  const showCol = 8;  // Column H (Show%)
  const closeCol = 9; // Column I (Close%)

  // Apply green background if metrics meet thresholds
  if (parsePercentage(dccValue) >= CONFIG.THRESHOLDS.DCC_RATE) {
    sheet.getRange(row, dccCol).setBackground('#07fc03'); // Light green
  }

  if (parsePercentage(showValue) >= CONFIG.THRESHOLDS.SHOW_RATE) {
    sheet.getRange(row, showCol).setBackground('#07fc03'); // Light green
  }

  if (parsePercentage(closeValue) >= CONFIG.THRESHOLDS.CLOSE_RATE) {
    sheet.getRange(row, closeCol).setBackground('#07fc03'); // Light green
  }
}

/**
 * Populate data for a specific time period (BY DAYS)
 */
function populatePeriodData(sheet, startRow, periodLabel, period, data, length) {
  // Period label
  sheet.getRange(startRow, 1, length, 1)
    .merge()
    .setValue(periodLabel)
    .setBackground('#9fc5e8')
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  const periodData = CONFIG.SETTERS.map(setter => {
    const dials = data.dials[period][setter] || 0;
    const pickups = data.pickups[period][setter] || 0;
    const sets = data.sets[period][setter] || 0;
    const outcomes = data.outcomes[period][setter] || { won: 0, lost: 0, total: 0 };

    const pickupRate = dials > 0 ? `${Math.round((pickups / dials) * 100)}%` : '0%';
    const dccRate = pickups > 0 ? `${Math.round((sets / pickups) * 100)}%` : '0%';
    const showRate = sets > 0 ? `${Math.round((outcomes.total / sets) * 100)}%` : '0%';
    const closeRate = outcomes.total > 0 ? `${Math.round((outcomes.won / outcomes.total) * 100)}%` : '0%';

    return [
      setter,
      dials,
      pickups,
      sets,
      pickupRate,
      dccRate,
      showRate,
      closeRate
    ];
  });

  // Calculate team totals
  const teamTotals = periodData.reduce((totals, row) => ({
    dials: totals.dials + row[1],
    pickups: totals.pickups + row[2],
    sets: totals.sets + row[3]
  }), { dials: 0, pickups: 0, sets: 0 });

  // Calculate team outcomes totals
  const teamOutcomes = CONFIG.SETTERS.reduce((totals, setter) => {
    const outcomes = data.outcomes[period][setter] || { won: 0, lost: 0, total: 0 };
    return {
      won: totals.won + outcomes.won,
      lost: totals.lost + outcomes.lost,
      total: totals.total + outcomes.total
    };
  }, { won: 0, lost: 0, total: 0 });

  // Team metrics using same formulas
  const teamPickupRate = teamTotals.dials > 0 ? `${Math.round((teamTotals.pickups / teamTotals.dials) * 100)}%` : '0%';
  const teamDccRate = teamTotals.pickups > 0 ? `${Math.round((teamTotals.sets / teamTotals.pickups) * 100)}%` : '0%';
  const teamShowRate = teamTotals.sets > 0 ? `${Math.round((teamOutcomes.total / teamTotals.sets) * 100)}%` : '0%';
  const teamCloseRate = teamOutcomes.total > 0 ? `${Math.round((teamOutcomes.won / teamOutcomes.total) * 100)}%` : '0%';

  periodData.push([
    'TEAM TOTAL',
    teamTotals.dials,
    teamTotals.pickups,
    teamTotals.sets,
    teamPickupRate,
    teamDccRate,
    teamShowRate,
    teamCloseRate
  ]);

  // Write data to sheet and apply conditional formatting
  periodData.forEach((rowData, index) => {
    const row = startRow + index;
    sheet.getRange(row, 2, 1, rowData.length).setValues([rowData]);

    // Apply conditional formatting for performance metrics
    const dccValue = rowData[5];   // DCC%
    const showValue = rowData[6];  // Show%
    const closeValue = rowData[7]; // Close%
    applyConditionalFormatting(sheet, row, dccValue, showValue, closeValue);

    // Format team row
    if (rowData[0] === 'TEAM TOTAL') {
      sheet.getRange(row, 2, 1, rowData.length)
        .setBackground('#f1f3f4')
        .setFontWeight('bold');

      // Reapply conditional formatting for team row (override the gray background for performance cells)
      applyConditionalFormatting(sheet, row, dccValue, showValue, closeValue);
    }
  });
}

/**
 * Populate data for a specific setter across all time periods
 */
function populateSetterGroupedData(sheet, startRow, setter, periods, periodLabels, data, allTimeData) {
  // Setter name header - spans 4 rows (3 periods + 1 total)
  sheet.getRange(startRow, 1, 4, 1)
    .merge()
    .setValue(setter)
    .setBackground('#9fc5e8')
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  // Data for each period
  const setterPeriodData = periods.map((period, index) => {
    const dials = data.dials[period][setter] || 0;
    const pickups = data.pickups[period][setter] || 0;
    const sets = data.sets[period][setter] || 0;
    const outcomes = data.outcomes[period][setter] || { won: 0, lost: 0, total: 0 };

    // Calculate metrics
    const pickupRate = dials > 0 ? `${Math.round((pickups / dials) * 100)}%` : '0%';
    const dccRate = pickups > 0 ? `${Math.round((sets / pickups) * 100)}%` : '0%';
    const showRate = sets > 0 ? `${Math.round((outcomes.total / sets) * 100)}%` : '0%';
    const closeRate = outcomes.total > 0 ? `${Math.round((outcomes.won / outcomes.total) * 100)}%` : '0%';

    return [
      periodLabels[index],
      dials,
      pickups,
      sets,
      pickupRate,
      dccRate,
      showRate,
      closeRate
    ];
  });

  // Use all-time data for totals instead of summing periods
  const setterTotals = {
    dials: allTimeData.dials[setter] || 0,
    pickups: allTimeData.pickups[setter] || 0,
    sets: allTimeData.sets[setter] || 0,
    won: (allTimeData.outcomes[setter] || { won: 0 }).won,
    lost: (allTimeData.outcomes[setter] || { lost: 0 }).lost,
    total: (allTimeData.outcomes[setter] || { total: 0 }).total
  };

  // Calculate total metrics for this setter using all-time data
  const totalPickupRate = setterTotals.dials > 0 ? `${Math.round((setterTotals.pickups / setterTotals.dials) * 100)}%` : '0%';
  const totalDccRate = setterTotals.pickups > 0 ? `${Math.round((setterTotals.sets / setterTotals.pickups) * 100)}%` : '0%';
  const totalShowRate = setterTotals.sets > 0 ? `${Math.round((setterTotals.total / setterTotals.sets) * 100)}%` : '0%';
  const totalCloseRate = setterTotals.total > 0 ? `${Math.round((setterTotals.won / setterTotals.total) * 100)}%` : '0%';

  setterPeriodData.push([
    'Total',
    setterTotals.dials,
    setterTotals.pickups,
    setterTotals.sets,
    totalPickupRate,
    totalDccRate,
    totalShowRate,
    totalCloseRate
  ]);

  // Write data to sheet and apply conditional formatting
  setterPeriodData.forEach((rowData, index) => {
    const row = startRow + index;
    sheet.getRange(row, 2, 1, rowData.length).setValues([rowData]);

    // Apply conditional formatting for performance metrics
    const dccValue = rowData[5];   // DCC%
    const showValue = rowData[6];  // Show%
    const closeValue = rowData[7]; // Close%
    applyConditionalFormatting(sheet, row, dccValue, showValue, closeValue);

    // Format total row
    if (rowData[0] === 'Total') {
      sheet.getRange(row, 2, 1, rowData.length)
        .setBackground('#f1f3f4')
        .setFontWeight('bold');

      // Reapply conditional formatting for total row (override the gray background for performance cells)
      applyConditionalFormatting(sheet, row, dccValue, showValue, closeValue);
    }
  });
}

/**
 * Populate team totals section
 */
function populateTeamTotalsSection(sheet, startRow, periods, periodLabels, data, allTimeData) {
  // Team header
  sheet.getRange(startRow, 1, 4, 1)
    .merge()
    .setValue('Team')
    .setBackground('#34a853')
    .setFontColor('white')
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  // Team data for each period
  const teamPeriodData = periods.map((period, index) => {
    // Calculate team totals for this period
    const teamTotals = CONFIG.SETTERS.reduce((totals, setter) => {
      const dials = data.dials[period][setter] || 0;
      const pickups = data.pickups[period][setter] || 0;
      const sets = data.sets[period][setter] || 0;
      const outcomes = data.outcomes[period][setter] || { won: 0, lost: 0, total: 0 };

      return {
        dials: totals.dials + dials,
        pickups: totals.pickups + pickups,
        sets: totals.sets + sets,
        won: totals.won + outcomes.won,
        lost: totals.lost + outcomes.lost,
        total: totals.total + outcomes.total
      };
    }, { dials: 0, pickups: 0, sets: 0, won: 0, lost: 0, total: 0 });

    // Calculate team metrics for this period
    const pickupRate = teamTotals.dials > 0 ? `${Math.round((teamTotals.pickups / teamTotals.dials) * 100)}%` : '0%';
    const dccRate = teamTotals.pickups > 0 ? `${Math.round((teamTotals.sets / teamTotals.pickups) * 100)}%` : '0%';
    const showRate = teamTotals.sets > 0 ? `${Math.round((teamTotals.total / teamTotals.sets) * 100)}%` : '0%';
    const closeRate = teamTotals.total > 0 ? `${Math.round((teamTotals.won / teamTotals.total) * 100)}%` : '0%';

    return [
      periodLabels[index],
      teamTotals.dials,
      teamTotals.pickups,
      teamTotals.sets,
      pickupRate,
      dccRate,
      showRate,
      closeRate
    ];
  });

  // Calculate overall team totals using all-time data
  const overallTeamTotals = CONFIG.SETTERS.reduce((grandTotals, setter) => {
    const dials = allTimeData.dials[setter] || 0;
    const pickups = allTimeData.pickups[setter] || 0;
    const sets = allTimeData.sets[setter] || 0;
    const outcomes = allTimeData.outcomes[setter] || { won: 0, lost: 0, total: 0 };

    return {
      dials: grandTotals.dials + dials,
      pickups: grandTotals.pickups + pickups,
      sets: grandTotals.sets + sets,
      won: grandTotals.won + outcomes.won,
      lost: grandTotals.lost + outcomes.lost,
      total: grandTotals.total + outcomes.total
    };
  }, { dials: 0, pickups: 0, sets: 0, won: 0, lost: 0, total: 0 });

  // Calculate overall team metrics using all-time data
  const overallPickupRate = overallTeamTotals.dials > 0 ? `${Math.round((overallTeamTotals.pickups / overallTeamTotals.dials) * 100)}%` : '0%';
  const overallDccRate = overallTeamTotals.pickups > 0 ? `${Math.round((overallTeamTotals.sets / overallTeamTotals.pickups) * 100)}%` : '0%';
  const overallShowRate = overallTeamTotals.sets > 0 ? `${Math.round((overallTeamTotals.total / overallTeamTotals.sets) * 100)}%` : '0%';
  const overallCloseRate = overallTeamTotals.total > 0 ? `${Math.round((overallTeamTotals.won / overallTeamTotals.total) * 100)}%` : '0%';

  teamPeriodData.push([
    'Total',
    overallTeamTotals.dials,
    overallTeamTotals.pickups,
    overallTeamTotals.sets,
    overallPickupRate,
    overallDccRate,
    overallShowRate,
    overallCloseRate
  ]);

  // Write team data to sheet and apply conditional formatting
  teamPeriodData.forEach((rowData, index) => {
    const row = startRow + index;
    sheet.getRange(row, 2, 1, rowData.length).setValues([rowData]);

    // Apply conditional formatting for performance metrics
    const dccValue = rowData[5];   // DCC%
    const showValue = rowData[6];  // Show%
    const closeValue = rowData[7]; // Close%
    applyConditionalFormatting(sheet, row, dccValue, showValue, closeValue);

    // Format total row
    if (rowData[0] === 'Total') {
      sheet.getRange(row, 2, 1, rowData.length)
        .setBackground('#f1f3f4')
        .setFontWeight('bold');

      // Reapply conditional formatting for total row (override the gray background for performance cells)
      applyConditionalFormatting(sheet, row, dccValue, showValue, closeValue);
    }
  });
}

/**
 * Format the sheet for better readability (BY DAYS)
 */
function formatSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Auto-resize columns
  sheet.autoResizeColumns(1, lastCol);

  // Set borders
  sheet.getRange(1, 1, lastRow, lastCol)
    .setBorder(true, true, true, true, true, true);

  // Align columns
  sheet.getRange(3, 4, lastRow - 2, 3).setHorizontalAlignment('right');
  sheet.getRange(3, 7, lastRow - 2, 4).setHorizontalAlignment('center');

  // Freeze header rows
  sheet.setFrozenRows(2);
}

/**
 * Format the setter-grouped sheet for better readability
 */
function formatSetterGroupedSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Auto-resize columns
  sheet.autoResizeColumns(1, lastCol);

  // Set borders
  sheet.getRange(1, 1, lastRow, lastCol)
    .setBorder(true, true, true, true, true, true);

  // Align columns
  sheet.getRange(3, 4, lastRow - 2, 3).setHorizontalAlignment('right');
  sheet.getRange(3, 7, lastRow - 2, 4).setHorizontalAlignment('center');

  // Freeze header rows
  sheet.setFrozenRows(2);
}

/**
 * Show loading message to user
 */
function showLoadingMessage(ui) {
  ui.alert('Generating Report', 'Please wait while the Setter Study report is being generated...', ui.ButtonSet.OK);
}

/**
 * Show success message to user
 */
function showSuccessMessage(ui) {
  ui.alert('Report Generated', 'Setter Study report has been successfully generated!', ui.ButtonSet.OK);
}

/**
 * Show error message to user
 */
function showErrorMessage(ui, error) {
  ui.alert('Error', `Failed to generate report: ${error.toString()}`, ui.ButtonSet.OK);
}

/**
 * Main function to generate the Closer Study report (BY DAYS)
 */
function generateCloserStudyByDays() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    showLoadingMessage(ui);

    let sheet = getOrCreateSheet(spreadsheet, 'Closer Study by Days');
    sheet.clear();

    // Fetch closer data from the Post Call Reports sheet
    const closerData = fetchCloserData();

    // Process data for all periods
    const processedData = {
      callsBooked: processCloserCallsBookedData(closerData),
      showRate: processCloserShowRateData(closerData),
      offers: processCloserOffersData(closerData),
      deposits: processCloserDepositsData(closerData),
      followUpPayments: processCloserFollowUpPaymentsData(closerData),
      callsTaken: processCloserCallsTakenData(closerData),
      setsClosed: processCloserSetsClosedData(closerData),
      cashCollected: processCloserCashCollectedData(closerData),
      revenue: processCloserRevenueData(closerData)
    };

    // Generate closer report
    generateCloserReport(sheet, processedData);

    showSuccessMessage(ui);

  } catch (error) {
    showErrorMessage(ui, error);
    console.error('Closer Study report generation failed:', error);
  }
}

/**
 * Fetch closer data from Post Call Reports sheet
 */
function fetchCloserData() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SETS_SHEET_ID);
    const tab = sheet.getSheetByName(CONFIG.SETS_TAB_NAME);

    if (!tab) {
      throw new Error(`Tab "${CONFIG.SETS_TAB_NAME}" not found`);
    }

    const values = tab.getDataRange().getValues();
    if (values.length <= 1) {
      throw new Error(`No data found in ${CONFIG.SETS_TAB_NAME}`);
    }

    return parseCloserData(values);

  } catch (error) {
    throw new Error(`Failed to fetch closer data: ${error.message}`);
  }
}

/**
 * Parse closer data from Post Call Reports sheet
 * Expected columns: Closing Call Time (Formatted), Closer, What's the new Lead Status
 */
function parseCloserData(values) {
  const [headers, ...data] = values;

  const closingCallTimeIndex = findColumnIndex(headers, ['closing call time (formatted)']);
  const dateSubmittedIndex = findColumnIndex(headers, ['date submitted']);
  const closerIndex = findColumnIndex(headers, ['closer']);
  const statusIndex = findColumnIndex(headers, ["what's the new lead status"]);
  const cashCollectedIndex = findColumnIndex(headers, ['cash collected']);
  const revenueIndex = findColumnIndex(headers, ['revenue generated']);
  const didMakeOfferIndex = findColumnIndex(headers, ['did you make a offer', 'did you make an offer']);

  console.log(`Closer Data - Closing Call Time: ${closingCallTimeIndex}, Date Submitted: ${dateSubmittedIndex}, Closer: ${closerIndex}, Status: ${statusIndex}, Cash Collected: ${cashCollectedIndex}, Revenue: ${revenueIndex}, Did Make Offer: ${didMakeOfferIndex}`);

  return data
    .map(row => ({
      closingCallTime: parseCloserDate(row[closingCallTimeIndex]),
      dateSubmitted: parseCloserDate(row[dateSubmittedIndex]),
      closer: cleanString(row[closerIndex]),
      status: cleanStatusString(row[statusIndex]),
      cashCollected: parseFloat(row[cashCollectedIndex]) || 0,
      revenue: parseFloat(row[revenueIndex]) || 0,
      didMakeOffer: cleanString(row[didMakeOfferIndex])
    }))
    .filter(row => row.closer && row.status);
}

/**
 * Match closer names with flexibility for partial matches
 */
function matchCloser(dataValue, configCloser) {
  const cleanData = cleanString(dataValue);
  const cleanConfig = cleanString(configCloser);

  // Handle "Alex El-H" specifically
  if (cleanConfig.includes('alex el-h')) {
    return cleanData.includes('alex') && (cleanData.includes('el-h') || cleanData.includes('elh'));
  }

  // Extract first name for other closers
  const firstName = cleanConfig.split(' ')[0];
  return cleanData.includes(firstName) || cleanData.includes(cleanConfig);
}

/**
 * Enhanced date parser for Closer data with multiple formats
 * Handles formats like: 07/01/2025, 7/3/2025, 07/09/2025 11:00 AM, 06/30/2025 12:00 PM
 */
function parseCloserDate(dateValue) {
  if (!dateValue) return null;
  if (dateValue instanceof Date) {
    // If it's already a Date object, normalize it to NY timezone
    const dateString = Utilities.formatDate(dateValue, 'America/New_York', 'yyyy-MM-dd');
    return new Date(dateString + 'T00:00:00');
  }

  if (typeof dateValue === 'string') {
    // Remove time portion if present (everything after space)
    const dateOnly = dateValue.trim().split(' ')[0];

    const parts = dateOnly.split('/');
    if (parts.length === 3) {
      const [month, day, year] = parts.map(p => parseInt(p.trim()));
      // Validate the parsed values
      if (!isNaN(month) && !isNaN(day) && !isNaN(year) &&
        month >= 1 && month <= 12 &&
        day >= 1 && day <= 31 &&
        year >= 1900) {
        // Create date and normalize to NY timezone
        const tempDate = new Date(year, month - 1, day);
        const dateString = Utilities.formatDate(tempDate, 'America/New_York', 'yyyy-MM-dd');
        return new Date(dateString + 'T00:00:00');
      }
    }
  }

  const parsed = new Date(dateValue);
  if (isNaN(parsed.getTime())) return null;

  // Normalize parsed date to NY timezone
  const dateString = Utilities.formatDate(parsed, 'America/New_York', 'yyyy-MM-dd');
  return new Date(dateString + 'T00:00:00');
}

/**
 * Process Calls Booked data - count rows with "discovery call complete" using Closing Call Time (Formatted)
 */
function processCloserCallsBookedData(rawData) {
  return processCloserDataByPeriod(rawData, (records, closer, startDate, endDate) =>
    records.filter(r =>
      matchCloser(r.closer, closer) &&
      r.closingCallTime &&
      r.closingCallTime >= startDate &&
      r.closingCallTime <= endDate &&
      r.status === 'discoverycallcomplete'
    ).length
  );
}

/**
 * Process Show Rate data - count rows with any outcome status (won/lost/marker) using Date Submitted
 */
function processCloserShowRateData(rawData) {
  const outcomeStatuses = [
    'won-innercircle(paymentplan)',
    'won-lifetime(paymentplan)',
    'followuppayment',
    'won-innercircle(cashpif)',
    'won-lifetime(cashpif)',
    'won-innercircle(financed)',
    'won-lifetime(financed)',
    'marker-innercircle',
    'marker-lifetime',
    'lost-innercircle(badlead)',
    'lost-innercircle(qualifiedlead)',
    'lost-innercircle(fianancing)',
    'lost-lifetime(badlead)',
    'lost-lifetime(qualifiedlead)',
    'lost-lifetime(fianancing)'
  ];

  return processCloserDataByPeriod(rawData, (records, closer, startDate, endDate) =>
    records.filter(r =>
      matchCloser(r.closer, closer) &&
      r.dateSubmitted &&
      r.dateSubmitted >= startDate &&
      r.dateSubmitted <= endDate &&
      outcomeStatuses.includes(r.status)
    ).length
  );
}

/**
 * Process Offers data - count rows with won/marker status using Date Submitted
 */
function processCloserOffersData(rawData) {
  return processCloserDataByPeriod(rawData, (records, closer, startDate, endDate) =>
    records.filter(r =>
      matchCloser(r.closer, closer) &&
      r.dateSubmitted &&
      r.dateSubmitted >= startDate &&
      r.dateSubmitted <= endDate &&
      r.didMakeOffer === 'yes'
    ).length
  );
}

/**
 * Process Deposits data - count rows with won status using Date Submitted
 */
function processCloserDepositsData(rawData) {
  const depositStatuses = [
    'marker-innercircle',
    'marker-lifetime'
  ];

  return processCloserDataByPeriod(rawData, (records, closer, startDate, endDate) =>
    records
      .filter(r =>
        matchCloser(r.closer, closer) &&
        r.dateSubmitted &&
        r.dateSubmitted >= startDate &&
        r.dateSubmitted <= endDate &&
        depositStatuses.includes(r.status)
      )
      .reduce((sum, r) => sum + r.cashCollected, 0)
  );
}

/**
 * Process Follow Up Payments data - count rows with "followuppayment" status using Date Submitted
 */
function processCloserFollowUpPaymentsData(rawData) {
  return processCloserDataByPeriod(rawData, (records, closer, startDate, endDate) =>
    records
      .filter(r =>
        matchCloser(r.closer, closer) &&
        r.dateSubmitted &&
        r.dateSubmitted >= startDate &&
        r.dateSubmitted <= endDate &&
        r.status === 'followuppayment'
      )
      .reduce((sum, r) => sum + r.cashCollected, 0)
  );
}

/**
 * Process Calls Taken data - count rows with "won" or "lost" status per closer
 */
function processCloserCallsTakenData(rawData) {
  const validStatuses = [
    'won-innercircle(paymentplan)',
    'won-lifetime(paymentplan)',
    'followuppayment',
    'won-innercircle(cashpif)',
    'won-lifetime(cashpif)',
    'won-innercircle(financed)',
    'won-lifetime(financed)',
    'marker-innercircle',
    'marker-lifetime',
    'lost-innercircle(badlead)',
    'lost-innercircle(qualifiedlead)',
    'lost-innercircle(fianancing)',
    'lost-lifetime(badlead)',
    'lost-lifetime(qualifiedlead)',
    'lost-lifetime(fianancing)'
  ];

  return processCloserDataByPeriod(rawData, (records, closer, startDate, endDate) =>
    records.filter(r =>
      matchCloser(r.closer, closer) &&
      r.dateSubmitted &&
      r.dateSubmitted >= startDate &&
      r.dateSubmitted <= endDate &&
      validStatuses.includes(r.status)
    ).length
  );
}

/**
 * Process Cash Collected data - sum cash from specific WON status rows using Date Submitted
 */
function processCloserCashCollectedData(rawData) {
  const wonStatuses = [
    'won-innercircle(paymentplan)',
    'won-lifetime(paymentplan)',
    'won-innercircle(cashpif)',
    'won-lifetime(cashpif)',
    'won-innercircle(financed)',
    'won-lifetime(financed)'
  ];

  return processCloserDataByPeriod(rawData, (records, closer, startDate, endDate) =>
    records
      .filter(r =>
        matchCloser(r.closer, closer) &&
        r.dateSubmitted &&
        r.dateSubmitted >= startDate &&
        r.dateSubmitted <= endDate &&
        wonStatuses.includes(r.status)
      )
      .reduce((sum, r) => sum + r.cashCollected, 0)
  );
}

/**
 * Process Revenue data - sum revenue from specific WON status rows using Date Submitted
 */
function processCloserRevenueData(rawData) {
  const wonStatuses = [
    'won-innercircle(paymentplan)',
    'won-lifetime(paymentplan)',
    'won-innercircle(cashpif)',
    'won-lifetime(cashpif)',
    'won-innercircle(financed)',
    'won-lifetime(financed)'
  ];

  return processCloserDataByPeriod(rawData, (records, closer, startDate, endDate) =>
    records
      .filter(r =>
        matchCloser(r.closer, closer) &&
        r.dateSubmitted &&
        r.dateSubmitted >= startDate &&
        r.dateSubmitted <= endDate &&
        wonStatuses.includes(r.status)
      )
      .reduce((sum, r) => sum + r.revenue, 0)
  );
}

/**
 * Process Sets Closed data - count "won" status rows per closer
 */
function processCloserSetsClosedData(rawData) {
  const wonStatuses = [
    'won-innercircle(paymentplan)',
    'won-lifetime(paymentplan)',
    'won-innercircle(cashpif)',
    'won-lifetime(cashpif)',
    'won-innercircle(financed)',
    'won-lifetime(financed)'
  ];

  return processCloserDataByPeriod(rawData, (records, closer, startDate, endDate) =>
    records.filter(r =>
      matchCloser(r.closer, closer) &&
      r.dateSubmitted &&
      r.dateSubmitted >= startDate &&
      r.dateSubmitted <= endDate &&
      wonStatuses.includes(r.status)
    ).length
  );
}

/**
 * Generic data processor for closers across different time periods
 */
function processCloserDataByPeriod(rawData, aggregator) {
  const today = getCurrentDateNY();
  const result = {};

  Object.entries(CONFIG.PERIODS).forEach(([period, days]) => {
    const startDate = getDateNDaysAgoNY(days);
    result[period] = {};

    CLOSER_CONFIG.CLOSERS.forEach(closer => {
      result[period][closer] = aggregator(rawData, closer, startDate, today);
    });
  });

  return result;
}

/**
 * Generate the complete Closer Study report (BY DAYS)
 */
function generateCloserReport(sheet, data) {
  setupCloserHeaders(sheet);

  const periods = ['last7', 'last14', 'last30'];
  const periodLabels = ['Last 7 Days', 'Last 14 Days', 'Last 30 Days'];
  let currentRow = 3;

  periods.forEach((period, index) => {
    populateCloserPeriodData(sheet, currentRow, periodLabels[index], period, data, CLOSER_CONFIG.CLOSERS.length + 1);
    currentRow += CLOSER_CONFIG.CLOSERS.length + 1;
  });

  formatCloserSheet(sheet);
}

/**
 * Setup Closer Study report headers
 */
function setupCloserHeaders(sheet) {
  // Main header
  const mainHeader = sheet.getRange('A1:P1');
  mainHeader.merge()
    .setValue('Closer Study Report')
    .setBackground('#4a90e2')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(14);

  // Column headers
  const headers = [
    'Period', 'Closer', 'Calls Booked', 'Show Rate', 'Offers', 'Offer (%)',
    'Deposits', 'Follow up Payments', 'Calls Taken', 'Sets Closed', 'Close Rate',
    'Collected ($)', 'Collected Per Call', 'Revenue ($)', 'AOV', 'Collected %'
  ];
  const headerRange = sheet.getRange(2, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground('#d9e2f3')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

/**
 * Helper function to apply conditional formatting to closer performance metrics
 */
function applyCloserConditionalFormatting(sheet, row, closeRateValue, collectedPercentValue) {
  const closeRateCol = 11;
  const collectedPercentCol = 16;

  // Apply green background if close rate >= 35%
  if (parsePercentage(closeRateValue) >= 35) {
    sheet.getRange(row, closeRateCol).setBackground('#07fc03');
  }

  // Apply green background if collected % >= 75%
  if (parsePercentage(collectedPercentValue) >= 75) {
    sheet.getRange(row, collectedPercentCol).setBackground('#07fc03');
  }
}

/**
 * Populate data for a specific time period (Closer Study)
 */
function populateCloserPeriodData(sheet, startRow, periodLabel, period, data, length) {
  // Period label
  sheet.getRange(startRow, 1, length, 1)
    .merge()
    .setValue(periodLabel)
    .setBackground('#9fc5e8')
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  const periodData = CLOSER_CONFIG.CLOSERS.map(closer => {
    const callsBooked = data.callsBooked[period][closer] || 0;
    const callsTaken = data.callsTaken[period][closer] || 0;
    const offers = data.offers[period][closer] || 0;
    const deposits = data.deposits[period][closer] || 0;
    const followUpPayments = data.followUpPayments[period][closer] || 0;
    const setsClosed = data.setsClosed[period][closer] || 0;
    const cashCollected = data.cashCollected[period][closer] || 0;
    const revenue = data.revenue[period][closer] || 0;

    // Calculate metrics
    const showRate = callsBooked > 0 ? `${Math.round((callsTaken / callsBooked) * 100)}%` : '0%';
    const offerRate = callsTaken > 0 ? `${Math.round((offers / callsTaken) * 100)}%` : '0%';
    const closeRate = callsTaken > 0 ? `${Math.round((setsClosed / callsTaken) * 100)}%` : '0%';
    const collectedPerCall = callsTaken > 0 ? `$${Math.round(cashCollected / callsTaken)}` : '$0';
    const aov = setsClosed > 0 ? `$${Math.round(cashCollected / setsClosed)}` : '$0';
    const collectedPercent = revenue > 0 ? `${Math.round((cashCollected / revenue) * 100)}%` : '0%';

    return [
      closer,
      callsBooked,
      showRate,
      offers,
      offerRate,
      `$${deposits}`,
      `$${followUpPayments}`,
      callsTaken,
      setsClosed,
      closeRate,
      `$${cashCollected}`,
      collectedPerCall,
      `$${revenue}`,
      aov,
      collectedPercent
    ];
  });

  // Calculate team totals
  const teamTotals = CLOSER_CONFIG.CLOSERS.reduce((totals, closer) => {
    const callsBooked = data.callsBooked[period][closer] || 0;
    const callsTaken = data.callsTaken[period][closer] || 0;
    const offers = data.offers[period][closer] || 0;
    const deposits = data.deposits[period][closer] || 0;
    const followUpPayments = data.followUpPayments[period][closer] || 0;
    const setsClosed = data.setsClosed[period][closer] || 0;
    const cashCollected = data.cashCollected[period][closer] || 0;
    const revenue = data.revenue[period][closer] || 0;

    return {
      callsBooked: totals.callsBooked + callsBooked,
      callsTaken: totals.callsTaken + callsTaken,
      offers: totals.offers + offers,
      deposits: totals.deposits + deposits,
      followUpPayments: totals.followUpPayments + followUpPayments,
      setsClosed: totals.setsClosed + setsClosed,
      cashCollected: totals.cashCollected + cashCollected,
      revenue: totals.revenue + revenue
    };
  }, { callsBooked: 0, callsTaken: 0, offers: 0, deposits: 0, followUpPayments: 0, setsClosed: 0, cashCollected: 0, revenue: 0 });

  // Team metrics
  const teamShowRate = teamTotals.callsBooked > 0 ? `${Math.round((teamTotals.callsTaken / teamTotals.callsBooked) * 100)}%` : '0%';
  const teamOfferRate = teamTotals.callsTaken > 0 ? `${Math.round((teamTotals.offers / teamTotals.callsTaken) * 100)}%` : '0%';
  const teamCloseRate = teamTotals.callsTaken > 0 ? `${Math.round((teamTotals.setsClosed / teamTotals.callsTaken) * 100)}%` : '0%';
  const teamCollectedPerCall = teamTotals.callsTaken > 0 ? `$${Math.round(teamTotals.cashCollected / teamTotals.callsTaken)}` : '$0';
  const teamAOV = teamTotals.setsClosed > 0 ? `$${Math.round(teamTotals.cashCollected / teamTotals.setsClosed)}` : '$0';
  const teamCollectedPercent = teamTotals.revenue > 0 ? `${Math.round((teamTotals.cashCollected / teamTotals.revenue) * 100)}%` : '0%';

  periodData.push([
    'TEAM TOTAL',
    teamTotals.callsBooked,
    teamShowRate,
    teamTotals.offers,
    teamOfferRate,
    `$${teamTotals.deposits}`,
    `$${teamTotals.followUpPayments}`,
    teamTotals.callsTaken,
    teamTotals.setsClosed,
    teamCloseRate,
    `$${teamTotals.cashCollected}`,
    teamCollectedPerCall,
    `$${teamTotals.revenue}`,
    teamAOV,
    teamCollectedPercent
  ]);

  // Write data to sheet and apply conditional formatting
  periodData.forEach((rowData, index) => {
    const row = startRow + index;
    sheet.getRange(row, 2, 1, rowData.length).setValues([rowData]);

    // Apply conditional formatting for close rate
    const closeRateValue = rowData[9]; // Close Rate
    const collectedPercentValue = rowData[14]; // Collected %
    applyCloserConditionalFormatting(sheet, row, closeRateValue, collectedPercentValue);

    // Format team row
    if (rowData[0] === 'TEAM TOTAL') {
      sheet.getRange(row, 2, 1, rowData.length)
        .setBackground('#f1f3f4')
        .setFontWeight('bold');

      // Reapply conditional formatting for team row
      applyCloserConditionalFormatting(sheet, row, closeRateValue);
    }
  });
}

/**
 * Format the Closer Study sheet for better readability
 */
function formatCloserSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Auto-resize columns
  sheet.autoResizeColumns(1, lastCol);

  // Set borders
  sheet.getRange(1, 1, lastRow, lastCol)
    .setBorder(true, true, true, true, true, true);

  // Align columns
  sheet.getRange(3, 3, lastRow - 2, 5).setHorizontalAlignment('right'); // Numbers
  sheet.getRange(3, 8, lastRow - 2, 8).setHorizontalAlignment('center'); // Percentages and currency

  // Freeze header rows
  sheet.setFrozenRows(2);
}

/**
 * Update the "Closer Study by Days" report
 */
function updateCloserReportByDays() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName('Closer Study by Days');
    if (!sheet) {
      sheet = spreadsheet.insertSheet('Closer Study by Days', 2);
    } else {
      spreadsheet.setActiveSheet(sheet);
    }

    console.log('Updating "Closer Study by Days" report...');

    // Clear existing data
    sheet.clear();

    // Fetch and process data
    const closerData = fetchCloserData();
    const processedData = {
      callsBooked: processCloserCallsBookedData(closerData),
      showRate: processCloserShowRateData(closerData),
      offers: processCloserOffersData(closerData),
      deposits: processCloserDepositsData(closerData),
      followUpPayments: processCloserFollowUpPaymentsData(closerData),
      callsTaken: processCloserCallsTakenData(closerData),
      setsClosed: processCloserSetsClosedData(closerData),
      cashCollected: processCloserCashCollectedData(closerData),
      revenue: processCloserRevenueData(closerData)
    };

    // Generate report
    generateCloserReport(sheet, processedData);

    console.log('"Closer Study by Days" report updated successfully');

  } catch (error) {
    console.error('Failed to update "Closer Study by Days" report:', error);
    throw error;
  }
}

/**
 * Generate Closer Study report grouped by Closer (instead of by time periods)
 */
function generateCloserStudyByCloser() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    showLoadingMessage(ui);

    let sheet = getOrCreateSheet(spreadsheet, 'Closer Study by Closer');
    sheet.clear();

    // Fetch closer data
    const closerData = fetchCloserData();

    // Process data for all periods
    const processedData = {
      callsBooked: processCloserCallsBookedData(closerData),
      showRate: processCloserShowRateData(closerData),
      offers: processCloserOffersData(closerData),
      deposits: processCloserDepositsData(closerData),
      followUpPayments: processCloserFollowUpPaymentsData(closerData),
      callsTaken: processCloserCallsTakenData(closerData),
      setsClosed: processCloserSetsClosedData(closerData),
      cashCollected: processCloserCashCollectedData(closerData),
      revenue: processCloserRevenueData(closerData)
    };

    // Process data for all time (for totals)
    const allTimeData = {
      callsBooked: processAllTimeCloserCallsBookedData(closerData),
      showRate: processAllTimeCloserShowRateData(closerData),
      offers: processAllTimeCloserOffersData(closerData),
      deposits: processAllTimeCloserDepositsData(closerData),
      followUpPayments: processAllTimeCloserFollowUpPaymentsData(closerData),
      callsTaken: processAllTimeCloserCallsTakenData(closerData),
      setsClosed: processAllTimeCloserSetsClosedData(closerData),
      cashCollected: processAllTimeCloserCashCollectedData(closerData),
      revenue: processAllTimeCloserRevenueData(closerData)
    };

    // Generate report grouped by closer
    generateCloserGroupedReport(sheet, processedData, allTimeData);

    showSuccessMessage(ui);

  } catch (error) {
    showErrorMessage(ui, error);
    console.error('Closer Study by Closer generation failed:', error);
  }
}

/**
 * Process Calls Booked data for all time (no date filtering)
 */
function processAllTimeCloserCallsBookedData(rawData) {
  return processAllTimeCloserData(rawData, (records, closer) =>
    records.filter(r =>
      matchCloser(r.closer, closer) &&
      r.status === 'discoverycallcomplete'
    ).length
  );
}

/**
 * Process Show Rate data for all time (no date filtering)
 */
function processAllTimeCloserShowRateData(rawData) {
  const outcomeStatuses = [
    'won-innercircle(paymentplan)',
    'won-lifetime(paymentplan)',
    'followuppayment',
    'won-innercircle(cashpif)',
    'won-lifetime(cashpif)',
    'won-innercircle(financed)',
    'won-lifetime(financed)',
    'marker-innercircle',
    'marker-lifetime',
    'lost-innercircle(badlead)',
    'lost-innercircle(qualifiedlead)',
    'lost-innercircle(fianancing)',
    'lost-lifetime(badlead)',
    'lost-lifetime(qualifiedlead)',
    'lost-lifetime(fianancing)'
  ];

  return processAllTimeCloserData(rawData, (records, closer) =>
    records.filter(r =>
      matchCloser(r.closer, closer) &&
      outcomeStatuses.includes(r.status)
    ).length
  );
}

/**
 * Process Offers data for all time (no date filtering)
 */
function processAllTimeCloserOffersData(rawData) {
  return processAllTimeCloserData(rawData, (records, closer) =>
    records.filter(r =>
      matchCloser(r.closer, closer) &&
      r.didMakeOffer === 'yes'
    ).length
  );
}

/**
 * Process Deposits data for all time (no date filtering)
 */
function processAllTimeCloserDepositsData(rawData) {
  const depositStatuses = [
    'marker-innercircle',
    'marker-lifetime'
  ];

  return processAllTimeCloserData(rawData, (records, closer) =>
    records
      .filter(r =>
        matchCloser(r.closer, closer) &&
        depositStatuses.includes(r.status)
      )
      .reduce((sum, r) => sum + r.cashCollected, 0)
  );
}

/**
 * Process Follow Up Payments data for all time (no date filtering)
 */
function processAllTimeCloserFollowUpPaymentsData(rawData) {
  return processAllTimeCloserData(rawData, (records, closer) =>
    records
      .filter(r =>
        matchCloser(r.closer, closer) &&
        r.status === 'followuppayment'
      )
      .reduce((sum, r) => sum + r.cashCollected, 0)
  );
}

/**
 * Process Calls Taken data for all time (no date filtering)
 */
function processAllTimeCloserCallsTakenData(rawData) {
  const validStatuses = [
    'won-innercircle(paymentplan)',
    'won-lifetime(paymentplan)',
    'followuppayment',
    'won-innercircle(cashpif)',
    'won-lifetime(cashpif)',
    'won-innercircle(financed)',
    'won-lifetime(financed)',
    'marker-innercircle',
    'marker-lifetime',
    'lost-innercircle(badlead)',
    'lost-innercircle(qualifiedlead)',
    'lost-innercircle(fianancing)',
    'lost-lifetime(badlead)',
    'lost-lifetime(qualifiedlead)',
    'lost-lifetime(fianancing)'
  ];

  return processAllTimeCloserData(rawData, (records, closer) =>
    records.filter(r =>
      matchCloser(r.closer, closer) &&
      validStatuses.includes(r.status)
    ).length
  );
}

/**
 * Process Sets Closed data for all time (no date filtering)
 */
function processAllTimeCloserSetsClosedData(rawData) {
  const wonStatuses = [
    'won-innercircle(paymentplan)',
    'won-lifetime(paymentplan)',
    'won-innercircle(cashpif)',
    'won-lifetime(cashpif)',
    'won-innercircle(financed)',
    'won-lifetime(financed)'
  ];

  return processAllTimeCloserData(rawData, (records, closer) =>
    records.filter(r =>
      matchCloser(r.closer, closer) &&
      wonStatuses.includes(r.status)
    ).length
  );
}

/**
 * Process Cash Collected data for all time (no date filtering)
 */
function processAllTimeCloserCashCollectedData(rawData) {
  const wonStatuses = [
    'won-innercircle(paymentplan)',
    'won-lifetime(paymentplan)',
    'won-innercircle(cashpif)',
    'won-lifetime(cashpif)',
    'won-innercircle(financed)',
    'won-lifetime(financed)'
  ];

  return processAllTimeCloserData(rawData, (records, closer) =>
    records
      .filter(r =>
        matchCloser(r.closer, closer) &&
        wonStatuses.includes(r.status)
      )
      .reduce((sum, r) => sum + r.cashCollected, 0)
  );
}

/**
 * Process Revenue data for all time (no date filtering)
 */
function processAllTimeCloserRevenueData(rawData) {
  const wonStatuses = [
    'won-innercircle(paymentplan)',
    'won-lifetime(paymentplan)',
    'won-innercircle(cashpif)',
    'won-lifetime(cashpif)',
    'won-innercircle(financed)',
    'won-lifetime(financed)'
  ];

  return processAllTimeCloserData(rawData, (records, closer) =>
    records
      .filter(r =>
        matchCloser(r.closer, closer) &&
        wonStatuses.includes(r.status)
      )
      .reduce((sum, r) => sum + r.revenue, 0)
  );
}

/**
 * Process closer data for all time (no date filtering) - for total calculations
 */
function processAllTimeCloserData(rawData, aggregator) {
  const result = {};

  CLOSER_CONFIG.CLOSERS.forEach(closer => {
    result[closer] = aggregator(rawData, closer);
  });

  return result;
}

/**
 * Generate the report grouped by closer
 */
function generateCloserGroupedReport(sheet, data, allTimeData) {
  setupCloserGroupedHeaders(sheet);

  let currentRow = 3;
  const periods = ['last7', 'last14', 'last30'];
  const periodLabels = ['Last 7', 'Last 14', 'Last 30'];

  // Process each closer
  CLOSER_CONFIG.CLOSERS.forEach((closer, closerIndex) => {
    populateCloserGroupedData(sheet, currentRow, closer, periods, periodLabels, data, allTimeData);
    currentRow += periods.length + 1 + 1;
  });

  // Add team totals section
  populateCloserTeamTotalsSection(sheet, currentRow, periods, periodLabels, data, allTimeData);

  formatCloserGroupedSheet(sheet);
}

/**
 * Setup headers for closer-grouped report
 */
function setupCloserGroupedHeaders(sheet) {
  // Main header
  const mainHeader = sheet.getRange('A1:P1');
  mainHeader.merge()
    .setValue('Closer Study by Closer')
    .setBackground('#4a90e2')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(14);

  // Column headers
  const headers = [
    'Closer', 'Days', 'Calls Booked', 'Show Rate', 'Offers', 'Offer (%)',
    'Deposits', 'Follow up Payments', 'Calls Taken', 'Sets Closed', 'Close Rate',
    'Collected ($)', 'Collected Per Call', 'Revenue ($)', 'AOV', 'Collected %'
  ];
  const headerRange = sheet.getRange(2, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground('#d9e2f3')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

/**
 * Populate data for a specific closer across all time periods
 */
function populateCloserGroupedData(sheet, startRow, closer, periods, periodLabels, data, allTimeData) {
  // Closer name header - spans 4 rows (3 periods + 1 total)
  sheet.getRange(startRow, 1, 4, 1)
    .merge()
    .setValue(closer)
    .setBackground('#9fc5e8')
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  // Data for each period
  const closerPeriodData = periods.map((period, index) => {
    const callsBooked = data.callsBooked[period][closer] || 0;
    const callsTaken = data.callsTaken[period][closer] || 0;
    const offers = data.offers[period][closer] || 0;
    const deposits = data.deposits[period][closer] || 0;
    const followUpPayments = data.followUpPayments[period][closer] || 0;
    const setsClosed = data.setsClosed[period][closer] || 0;
    const cashCollected = data.cashCollected[period][closer] || 0;
    const revenue = data.revenue[period][closer] || 0;

    // Calculate metrics
    const showRate = callsBooked > 0 ? `${Math.round((callsTaken / callsBooked) * 100)}%` : '0%';
    const offerRate = callsTaken > 0 ? `${Math.round((offers / callsTaken) * 100)}%` : '0%';
    const closeRate = callsTaken > 0 ? `${Math.round((setsClosed / callsTaken) * 100)}%` : '0%';
    const collectedPerCall = callsTaken > 0 ? `$${Math.round(cashCollected / callsTaken)}` : '$0';
    const aov = setsClosed > 0 ? `$${Math.round(cashCollected / setsClosed)}` : '$0';
    const collectedPercent = revenue > 0 ? `${Math.round((cashCollected / revenue) * 100)}%` : '0%';

    return [
      periodLabels[index],
      callsBooked,
      showRate,
      offers,
      offerRate,
      `$${deposits}`,
      `$${followUpPayments}`,
      callsTaken,
      setsClosed,
      closeRate,
      `$${cashCollected}`,
      collectedPerCall,
      `$${revenue}`,
      aov,
      collectedPercent
    ];
  });

  // Use all-time data for totals instead of summing periods
  const closerTotals = {
    callsBooked: allTimeData.callsBooked[closer] || 0,
    callsTaken: allTimeData.callsTaken[closer] || 0,
    offers: allTimeData.offers[closer] || 0,
    deposits: allTimeData.deposits[closer] || 0,
    followUpPayments: allTimeData.followUpPayments[closer] || 0,
    setsClosed: allTimeData.setsClosed[closer] || 0,
    cashCollected: allTimeData.cashCollected[closer] || 0,
    revenue: allTimeData.revenue[closer] || 0
  };

  // Calculate total metrics for this closer using all-time data
  const totalShowRate = closerTotals.callsBooked > 0 ? `${Math.round((closerTotals.callsTaken / closerTotals.callsBooked) * 100)}%` : '0%';
  const totalOfferRate = closerTotals.callsTaken > 0 ? `${Math.round((closerTotals.offers / closerTotals.callsTaken) * 100)}%` : '0%';
  const totalCloseRate = closerTotals.callsTaken > 0 ? `${Math.round((closerTotals.setsClosed / closerTotals.callsTaken) * 100)}%` : '0%';
  const totalCollectedPerCall = closerTotals.callsTaken > 0 ? `$${Math.round(closerTotals.cashCollected / closerTotals.callsTaken)}` : '$0';
  const totalAOV = closerTotals.setsClosed > 0 ? `$${Math.round(closerTotals.cashCollected / closerTotals.setsClosed)}` : '$0';
  const totalCollectedPercent = closerTotals.revenue > 0 ? `${Math.round((closerTotals.cashCollected / closerTotals.revenue) * 100)}%` : '0%';

  closerPeriodData.push([
    'Total',
    closerTotals.callsBooked,
    totalShowRate,
    closerTotals.offers,
    totalOfferRate,
    `$${closerTotals.deposits}`,
    `$${closerTotals.followUpPayments}`,
    closerTotals.callsTaken,
    closerTotals.setsClosed,
    totalCloseRate,
    `$${closerTotals.cashCollected}`,
    totalCollectedPerCall,
    `$${closerTotals.revenue}`,
    totalAOV,
    totalCollectedPercent
  ]);

  // Write data to sheet and apply conditional formatting
  closerPeriodData.forEach((rowData, index) => {
    const row = startRow + index;
    sheet.getRange(row, 2, 1, rowData.length).setValues([rowData]);

    // Apply conditional formatting for close rate
    const closeRateValue = rowData[9]; // Close Rate
    const collectedPercentValue = rowData[14]; // Collected %
    applyCloserConditionalFormatting(sheet, row, closeRateValue, collectedPercentValue);

    // Format total row
    if (rowData[0] === 'Total') {
      sheet.getRange(row, 2, 1, rowData.length)
        .setBackground('#f1f3f4')
        .setFontWeight('bold');

      // Reapply conditional formatting for total row (override the gray background for performance cells)
      applyCloserConditionalFormatting(sheet, row, closeRateValue);
    }
  });
}

/**
 * Populate team totals section for closer-grouped report
 */
function populateCloserTeamTotalsSection(sheet, startRow, periods, periodLabels, data, allTimeData) {
  // Team header
  sheet.getRange(startRow, 1, 4, 1)
    .merge()
    .setValue('Team')
    .setBackground('#34a853')
    .setFontColor('white')
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  // Team data for each period
  const teamPeriodData = periods.map((period, index) => {
    // Calculate team totals for this period
    const teamTotals = CLOSER_CONFIG.CLOSERS.reduce((totals, closer) => {
      const callsBooked = data.callsBooked[period][closer] || 0;
      const callsTaken = data.callsTaken[period][closer] || 0;
      const offers = data.offers[period][closer] || 0;
      const deposits = data.deposits[period][closer] || 0;
      const followUpPayments = data.followUpPayments[period][closer] || 0;
      const setsClosed = data.setsClosed[period][closer] || 0;
      const cashCollected = data.cashCollected[period][closer] || 0;
      const revenue = data.revenue[period][closer] || 0;

      return {
        callsBooked: totals.callsBooked + callsBooked,
        callsTaken: totals.callsTaken + callsTaken,
        offers: totals.offers + offers,
        deposits: totals.deposits + deposits,
        followUpPayments: totals.followUpPayments + followUpPayments,
        setsClosed: totals.setsClosed + setsClosed,
        cashCollected: totals.cashCollected + cashCollected,
        revenue: totals.revenue + revenue
      };
    }, { callsBooked: 0, callsTaken: 0, offers: 0, deposits: 0, followUpPayments: 0, setsClosed: 0, cashCollected: 0, revenue: 0 });

    // Calculate team metrics for this period
    const showRate = teamTotals.callsBooked > 0 ? `${Math.round((teamTotals.callsTaken / teamTotals.callsBooked) * 100)}%` : '0%';
    const offerRate = teamTotals.callsTaken > 0 ? `${Math.round((teamTotals.offers / teamTotals.callsTaken) * 100)}%` : '0%';
    const closeRate = teamTotals.callsTaken > 0 ? `${Math.round((teamTotals.setsClosed / teamTotals.callsTaken) * 100)}%` : '0%';
    const collectedPerCall = teamTotals.callsTaken > 0 ? `$${Math.round(teamTotals.cashCollected / teamTotals.callsTaken)}` : '$0';
    const aov = teamTotals.setsClosed > 0 ? `$${Math.round(teamTotals.revenue / teamTotals.setsClosed)}` : '$0';
    const collectedPercent = teamTotals.revenue > 0 ? `${Math.round((teamTotals.cashCollected / teamTotals.revenue) * 100)}%` : '0%';

    return [
      periodLabels[index],
      teamTotals.callsBooked,
      showRate,
      teamTotals.offers,
      offerRate,
      `$${teamTotals.deposits}`,
      `$${teamTotals.followUpPayments}`,
      teamTotals.callsTaken,
      teamTotals.setsClosed,
      closeRate,
      `$${teamTotals.cashCollected}`,
      collectedPerCall,
      `$${teamTotals.revenue}`,
      aov,
      collectedPercent
    ];
  });

  // Calculate overall team totals using all-time data
  const overallTeamTotals = CLOSER_CONFIG.CLOSERS.reduce((grandTotals, closer) => {
    const callsBooked = allTimeData.callsBooked[closer] || 0;
    const callsTaken = allTimeData.callsTaken[closer] || 0;
    const offers = allTimeData.offers[closer] || 0;
    const deposits = allTimeData.deposits[closer] || 0;
    const followUpPayments = allTimeData.followUpPayments[closer] || 0;
    const setsClosed = allTimeData.setsClosed[closer] || 0;
    const cashCollected = allTimeData.cashCollected[closer] || 0;
    const revenue = allTimeData.revenue[closer] || 0;

    return {
      callsBooked: grandTotals.callsBooked + callsBooked,
      callsTaken: grandTotals.callsTaken + callsTaken,
      offers: grandTotals.offers + offers,
      deposits: grandTotals.deposits + deposits,
      followUpPayments: grandTotals.followUpPayments + followUpPayments,
      setsClosed: grandTotals.setsClosed + setsClosed,
      cashCollected: grandTotals.cashCollected + cashCollected,
      revenue: grandTotals.revenue + revenue
    };
  }, { callsBooked: 0, callsTaken: 0, offers: 0, deposits: 0, followUpPayments: 0, setsClosed: 0, cashCollected: 0, revenue: 0 });

  // Calculate overall team metrics using all-time data
  const overallShowRate = overallTeamTotals.callsBooked > 0 ? `${Math.round((overallTeamTotals.callsTaken / overallTeamTotals.callsBooked) * 100)}%` : '0%';
  const overallOfferRate = overallTeamTotals.callsTaken > 0 ? `${Math.round((overallTeamTotals.offers / overallTeamTotals.callsTaken) * 100)}%` : '0%';
  const overallCloseRate = overallTeamTotals.callsTaken > 0 ? `${Math.round((overallTeamTotals.setsClosed / overallTeamTotals.callsTaken) * 100)}%` : '0%';
  const overallCollectedPerCall = overallTeamTotals.callsTaken > 0 ? `$${Math.round(overallTeamTotals.cashCollected / overallTeamTotals.callsTaken)}` : '$0';
  const overallAOV = overallTeamTotals.setsClosed > 0 ? `$${Math.round(overallTeamTotals.revenue / overallTeamTotals.setsClosed)}` : '$0';
  const overallCollectedPercent = overallTeamTotals.revenue > 0 ? `${Math.round((overallTeamTotals.cashCollected / overallTeamTotals.revenue) * 100)}%` : '0%';

  teamPeriodData.push([
    'Total',
    overallTeamTotals.callsBooked,
    overallShowRate,
    overallTeamTotals.offers,
    overallOfferRate,
    `$${overallTeamTotals.deposits}`,
    `$${overallTeamTotals.followUpPayments}`,
    overallTeamTotals.callsTaken,
    overallTeamTotals.setsClosed,
    overallCloseRate,
    `$${overallTeamTotals.cashCollected}`,
    overallCollectedPerCall,
    `$${overallTeamTotals.revenue}`,
    overallAOV,
    overallCollectedPercent
  ]);

  // Write team data to sheet and apply conditional formatting
  teamPeriodData.forEach((rowData, index) => {
    const row = startRow + index;
    sheet.getRange(row, 2, 1, rowData.length).setValues([rowData]);

    // Apply conditional formatting for close rate
    const closeRateValue = rowData[9]; // Close Rate
    const collectedPercentValue = rowData[14]; // Collected %
    applyCloserConditionalFormatting(sheet, row, closeRateValue, collectedPercentValue);

    // Format total row
    if (rowData[0] === 'Total') {
      sheet.getRange(row, 2, 1, rowData.length)
        .setBackground('#f1f3f4')
        .setFontWeight('bold');

      // Reapply conditional formatting for total row (override the gray background for performance cells)
      applyCloserConditionalFormatting(sheet, row, closeRateValue);
    }
  });
}

/**
 * Format the closer-grouped sheet for better readability
 */
function formatCloserGroupedSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Auto-resize columns
  sheet.autoResizeColumns(1, lastCol);

  // Set borders
  sheet.getRange(1, 1, lastRow, lastCol)
    .setBorder(true, true, true, true, true, true);

  // Align columns
  sheet.getRange(3, 4, lastRow - 2, 4).setHorizontalAlignment('right'); // Numbers
  sheet.getRange(3, 8, lastRow - 2, 1).setHorizontalAlignment('center'); // Percentages

  // Freeze header rows
  sheet.setFrozenRows(2);
}

/**
 * Update the "Closer Study by Closer" report
 */
function updateCloserReportByCloser() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName('Closer Study by Closer');
    if (!sheet) {
      sheet = spreadsheet.insertSheet('Closer Study by Closer', 3);
    } else {
      spreadsheet.setActiveSheet(sheet);
    }

    console.log('Updating "Closer Study by Closer" report...');

    // Clear existing data
    sheet.clear();

    // Fetch and process data
    const closerData = fetchCloserData();
    const processedData = {
      callsBooked: processCloserCallsBookedData(closerData),
      showRate: processCloserShowRateData(closerData),
      offers: processCloserOffersData(closerData),
      deposits: processCloserDepositsData(closerData),
      followUpPayments: processCloserFollowUpPaymentsData(closerData),
      callsTaken: processCloserCallsTakenData(closerData),
      setsClosed: processCloserSetsClosedData(closerData),
      cashCollected: processCloserCashCollectedData(closerData),
      revenue: processCloserRevenueData(closerData)
    };

    const allTimeData = {
      callsBooked: processAllTimeCloserCallsBookedData(closerData),
      showRate: processAllTimeCloserShowRateData(closerData),
      offers: processAllTimeCloserOffersData(closerData),
      deposits: processAllTimeCloserDepositsData(closerData),
      followUpPayments: processAllTimeCloserFollowUpPaymentsData(closerData),
      callsTaken: processAllTimeCloserCallsTakenData(closerData),
      setsClosed: processAllTimeCloserSetsClosedData(closerData),
      cashCollected: processAllTimeCloserCashCollectedData(closerData),
      revenue: processAllTimeCloserRevenueData(closerData)
    };

    // Generate closer-grouped report
    generateCloserGroupedReport(sheet, processedData, allTimeData);

    console.log('"Closer Study by Closer" report updated successfully');

  } catch (error) {
    console.error('Failed to update "Closer Study by Closer" report:', error);
    throw error;
  }
}

/**
 * Main function to generate the Vortex Study report (BY DAYS)
 */
function generateVortexData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    showLoadingMessage(ui);

    let sheet = getOrCreateSheet(spreadsheet, 'Vortex Data');
    sheet.clear();

    // Generate daily data from September 1st to today with real Low Ticket Buys data
    const dailyData = generateDailyVortexData();

    // Generate vortex report with daily data
    generateVortexReport(sheet, dailyData);

    showSuccessMessage(ui);

  } catch (error) {
    showErrorMessage(ui, error);
    console.error('Vortex Study report generation failed:', error);
  }
}

/**
 * Generate daily Vortex data from September 1st to today
*/
function generateDailyVortexData() {
  const today = getCurrentDateNY();
  const startDate = new Date('2025-09-01T00:00:00');
  const dailyData = [];

  // Fetch all data sources
  const lowTicketData = fetchLowTicketBuysData();
  const partialTriageData = fetchPartialTriageData();
  const bookedTriageData = fetchBookedTriageData();
  const bookedClosingCallData = fetchBookedClosingCallData();
  const shownCallsData = fetchShownCallsData();
  const closesData = fetchClosesData();
  const newCashData = fetchNewCashData();
  const fuCashData = fetchFUCashData();

  // Create a map for quick lookup by date for Low Ticket Buys
  const discrodJoinedMap = new Map();
  const lowTicketMap = new Map();
  lowTicketData.forEach(record => {
    const month = record.date.getMonth() + 1;
    const day = record.date.getDate();
    const year = record.date.getFullYear();
    const dateKey = `${month}/${day}/${year}`;
    lowTicketMap.set(dateKey, record.lowTicket);
    discrodJoinedMap.set(dateKey, record.totalInvites);
  });

  // Create a map for quick lookup by date for Partial Triage (count occurrences)
  const partialTriageMap = new Map();
  partialTriageData.forEach(record => {
    const month = record.date.getMonth() + 1;
    const day = record.date.getDate();
    const year = record.date.getFullYear();
    const dateKey = `${month}/${day}/${year}`;

    // Count occurrences for each date
    const currentCount = partialTriageMap.get(dateKey) || 0;
    partialTriageMap.set(dateKey, currentCount + 1);
  });

  // Create a map for quick lookup by date for Booked Triage (count occurrences)
  const bookedTriageMap = new Map();
  bookedTriageData.forEach(record => {
    const month = record.date.getMonth() + 1;
    const day = record.date.getDate();
    const year = record.date.getFullYear();
    const dateKey = `${month}/${day}/${year}`;

    // Count occurrences for each date
    const currentCount = bookedTriageMap.get(dateKey) || 0;
    bookedTriageMap.set(dateKey, currentCount + 1);
  });

  // Create a map for quick lookup by date for Booked Closing Call (count occurrences)
  const bookedClosingCallMap = new Map();
  bookedClosingCallData.forEach(record => {
    const month = record.date.getMonth() + 1;
    const day = record.date.getDate();
    const year = record.date.getFullYear();
    const dateKey = `${month}/${day}/${year}`;

    // Count occurrences for each date
    const currentCount = bookedClosingCallMap.get(dateKey) || 0;
    bookedClosingCallMap.set(dateKey, currentCount + 1);
  });

  // Create a map for quick lookup by date for Shown Calls (count occurrences)
  const shownCallsMap = new Map();
  shownCallsData.forEach(record => {
    const month = record.date.getMonth() + 1;
    const day = record.date.getDate();
    const year = record.date.getFullYear();
    const dateKey = `${month}/${day}/${year}`;

    // Count occurrences for each date
    const currentCount = shownCallsMap.get(dateKey) || 0;
    shownCallsMap.set(dateKey, currentCount + 1);
  });

  // Create a map for quick lookup by date for Closes (count occurrences)
  const closesMap = new Map();
  closesData.forEach(record => {
    const month = record.date.getMonth() + 1;
    const day = record.date.getDate();
    const year = record.date.getFullYear();
    const dateKey = `${month}/${day}/${year}`;

    // Count occurrences for each date
    const currentCount = closesMap.get(dateKey) || 0;
    closesMap.set(dateKey, currentCount + 1);
  });

  // Create a map for quick lookup by date for New Cash (sum cash amounts)
  const newCashMap = new Map();
  newCashData.forEach(record => {
    const month = record.date.getMonth() + 1;
    const day = record.date.getDate();
    const year = record.date.getFullYear();
    const dateKey = `${month}/${day}/${year}`;

    // Sum cash collected for each date
    const currentAmount = newCashMap.get(dateKey) || 0;
    newCashMap.set(dateKey, currentAmount + record.cashCollected);
  });

  // Create a map for quick lookup by date for FU Cash (sum cash amounts)
  const fuCashMap = new Map();
  fuCashData.forEach(record => {
    const month = record.date.getMonth() + 1;
    const day = record.date.getDate();
    const year = record.date.getFullYear();
    const dateKey = `${month}/${day}/${year}`;

    // Sum cash collected for each date
    const currentAmount = fuCashMap.get(dateKey) || 0;
    fuCashMap.set(dateKey, currentAmount + record.cashCollected);
  });

  // Generate data for each day from Sept 1st to today
  for (let date = new Date(startDate); date <= today; date.setDate(date.getDate() + 1)) {
    const currentDate = new Date(date);

    // Create date key in M/D/YYYY format to match source data
    const month = currentDate.getMonth() + 1;
    const day = currentDate.getDate();
    const year = currentDate.getFullYear();
    const dateKey = `${month}/${day}/${year}`;

    // Get data for this date, default to 0 if not found
    const discordJoins = discrodJoinedMap.get(dateKey) || 0;
    const lowTicketBuys = lowTicketMap.get(dateKey) || 0;
    const partialTriage = partialTriageMap.get(dateKey) || 0;
    const bookedTriage = bookedTriageMap.get(dateKey) || 0;
    const totalTriageCalls = partialTriage + bookedTriage;
    const bookedClosingCall = bookedClosingCallMap.get(dateKey) || 0;
    const shownCalls = shownCallsMap.get(dateKey) || 0;
    const closes = closesMap.get(dateKey) || 0;
    const newCash = newCashMap.get(dateKey) || 0;
    const fuCash = fuCashMap.get(dateKey) || 0;

    dailyData.push({
      date: currentDate,
      discordJoins: discordJoins, // Only placeholder remaining
      lowTicketBuys: lowTicketBuys, // Real data
      bookedTriage: bookedTriage, // Now using real data
      partialTriage: partialTriage, // Real data
      totalTriageCalls: totalTriageCalls, // Total Data
      bookedClosingCall: bookedClosingCall, // Real data
      shownCalls: shownCalls, // Real data
      closes: closes, // Real data
      newCash: newCash, // Real data
      fuCash: fuCash // Real data
    });
  }

  return dailyData;
}

/**
 * Generate the complete Vortex Study report with daily data
 */
function generateVortexReport(sheet, dailyData) {
  setupVortexHeaders(sheet);

  let currentRow = 3;
  let previousDayData = null;

  // Populate each day's data
  dailyData.forEach((dayData, index) => {
    populateVortexDailyData(sheet, currentRow, dayData, previousDayData);
    previousDayData = dayData; // Set current day as previous for next iteration
    currentRow += 1; // Move to next row
  });

  formatVortexSheet(sheet);
}

/**
 * Setup Vortex Study report headers
 */
function setupVortexHeaders(sheet) {
  // Main header
  const mainHeader = sheet.getRange('A1:L1');
  mainHeader.merge()
    .setValue('Vortex Study Report')
    .setBackground('#4a90e2')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(14);

  // Column headers
  const headers = [
    'Date', 'Discord Joins', 'Low Ticket Buys', 'Conversion Rate', 'Booked Triage', 'Partial Triage', 'Total Triage Calls',
    'Booked Closing Call', 'Shown Calls', 'Closes', 'New Cash', 'FU Cash'
  ];
  const headerRange = sheet.getRange(2, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground('#d9e2f3')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

/**
 * Populate data for a specific day (Vortex Study)
 */
function populateVortexDailyData(sheet, row, dayData, previousDayData = null) {
  // Format date as M/D/YYYY to match your screenshot
  const formattedDate = Utilities.formatDate(dayData.date, 'America/New_York', 'M/d/yyyy');

  const discordJoins = dayData.discordJoins || 0;
  const lowTicketBuys = dayData.lowTicketBuys || 0;
  const bookedTriage = dayData.bookedTriage || 0;
  const partialTriage = dayData.partialTriage || 0;
  const totalTriageCalls = dayData.totalTriageCalls || 0;
  const bookedClosingCall = dayData.bookedClosingCall || 0;
  const shownCalls = dayData.shownCalls || 0;
  const closes = dayData.closes || 0;
  const newCash = dayData.newCash || 0;
  const fuCash = dayData.fuCash || 0;
  const conversionRate = discordJoins > 0 ? `${Math.round((lowTicketBuys / discordJoins) * 100)}%` : '0%';
  const formattedDiscordJoins = discordJoins > 0 ? discordJoins.toLocaleString() : discordJoins;
  const formattedtotalTriageCalls = totalTriageCalls > 0 ? totalTriageCalls.toLocaleString() : totalTriageCalls;

  const rowData = [
    formattedDate,
    formattedDiscordJoins,
    lowTicketBuys,
    conversionRate,
    bookedTriage,
    partialTriage,
    formattedtotalTriageCalls,
    bookedClosingCall,
    shownCalls,
    closes,
    newCash > 0 ? `$${newCash}` : newCash,
    fuCash > 0 ? `$${fuCash}` : fuCash
  ];

  // Write data to sheet
  sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);

  // Apply conditional formatting if we have previous day data
  if (previousDayData) {
    const prevConversionRate = previousDayData.discordJoins > 0 ? (previousDayData.lowTicketBuys / previousDayData.discordJoins) * 100 : 0;
    const currConversionRate = discordJoins > 0 ? (lowTicketBuys / discordJoins) * 100 : 0;
    
    const currentValues = [discordJoins, lowTicketBuys, currConversionRate, bookedTriage, partialTriage, totalTriageCalls, bookedClosingCall, shownCalls, closes, newCash, fuCash];
    const previousValues = [
      previousDayData.discordJoins || 0,
      previousDayData.lowTicketBuys || 0,
      prevConversionRate,
      previousDayData.bookedTriage || 0,
      previousDayData.partialTriage || 0,
      previousDayData.totalTriageCalls || 0,
      previousDayData.bookedClosingCall || 0,
      previousDayData.shownCalls || 0,
      previousDayData.closes || 0,
      previousDayData.newCash || 0,
      previousDayData.fuCash || 0
    ];

    // Apply formatting to columns 2-12 (skip date column)
    for (let col = 2; col <= 12; col++) {
      const currentValue = currentValues[col - 2];
      const previousValue = previousValues[col - 2];

      let backgroundColor = null;

      if (currentValue > previousValue) {
        backgroundColor = '#07fc03'; // Green
      } else if (currentValue < previousValue) {
        backgroundColor = '#ff0000'; // Red
      } else if (currentValue === previousValue) {
        backgroundColor = '#ffff00'; // Yellow
      }

      if (backgroundColor) {
        sheet.getRange(row, col).setBackground(backgroundColor);
      }
    }
  }
}

/**
 * Format the Vortex Study sheet for better readability
 */
function formatVortexSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Set specific column widths for better visibility
  sheet.setColumnWidth(1, 100);  // Date column
  sheet.setColumnWidth(2, 120);  // Discord Joins
  sheet.setColumnWidth(3, 130);  // Low Ticket Buys
  sheet.setColumnWidth(4, 130);  // Conversion Rate
  sheet.setColumnWidth(5, 120);  // Booked Triage
  sheet.setColumnWidth(6, 120);  // Partial Triage
  sheet.setColumnWidth(7, 140);  // Total Triage Calls
  sheet.setColumnWidth(8, 150);  // Booked Closing Call
  sheet.setColumnWidth(9, 110);  // Shown Calls
  sheet.setColumnWidth(10, 80);   // Closes
  sheet.setColumnWidth(11, 100);  // New Cash
  sheet.setColumnWidth(12, 100); // FU Cash

  sheet.getRange(1, 1, lastRow, lastCol)
    .setBorder(true, true, true, true, true, true);

  sheet.getRange(3, 1, lastRow - 2, lastCol).setHorizontalAlignment('center');

  sheet.setFrozenRows(2);
}

/**
 * Fetch Low Ticket Buys data from OVERALL PERFORMANCE sheet
 */
function fetchLowTicketBuysData() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.LOW_TICKET_SHEET_ID);
    const tab = sheet.getSheetByName(CONFIG.LOW_TICKET_TAB_NAME);

    if (!tab) {
      throw new Error(`Tab "${CONFIG.LOW_TICKET_TAB_NAME}" not found`);
    }

    const values = tab.getDataRange().getValues();
    if (values.length <= 1) {
      throw new Error(`No data found in ${CONFIG.LOW_TICKET_TAB_NAME}`);
    }

    const sheetTimeZone = sheet.getSpreadsheetTimeZone();
    console.log('Sheet timezone:', sheetTimeZone);

    // Convert Date objects back to string format
    const formattedValues = values.map(row =>
      row.map(cell => {
        if (cell instanceof Date) {
          return Utilities.formatDate(cell, sheetTimeZone, 'M/d/yyyy');
        }
        return cell;
      })
    );

    return parseLowTicketBuysData(formattedValues);

  } catch (error) {
    throw new Error(`Failed to fetch Low Ticket Buys data: ${error.message}`);
  }
}

/**
 * Parse Low Ticket Buys data from OVERALL PERFORMANCE sheet
 * Expected columns: Date, Total Invites
*/
function parseLowTicketBuysData(values) {
  const [headers, ...data] = values;

  // Find column indices with flexible matching
  const dateIndex = findColumnIndex(headers, ['date']);
  const TotalIndex = findColumnIndex(headers, ['total invites']);
  const totalInvitesIndex = findColumnIndex(headers, ['total invites', 'invites']);
  const lowticketIndex = findColumnIndex(headers, ['sales on that day']);

  console.log(`Low Ticket Buys - Date: ${dateIndex}, Total Invites: ${totalInvitesIndex}`);

  return data
    .map(row => {
      const dateValue = row[dateIndex];
      const TotalValue = row[TotalIndex];
      console.log(dateValue, TotalValue)
      let parsedDate = null;

      // Handle the specific date format from your sheet
      if (dateValue instanceof Date) {
        parsedDate = new Date(dateValue);
      } else if (typeof dateValue === 'string') {
        // Parse MM/DD/YYYY format directly
        const parts = dateValue.split('/');
        if (parts.length === 3) {
          const month = parseInt(parts[0]);
          const day = parseInt(parts[1]);
          const year = parseInt(parts[2]);
          parsedDate = new Date(year, month - 1, day);
        }
      }

      return {
        date: parsedDate,
        totalInvites: parseInt(row[totalInvitesIndex]) || 0,
        lowTicket: parseInt(row[lowticketIndex]) || 0
      };
    })
    .filter(row => row.date && !isNaN(row.date.getTime()) && row.totalInvites >= 0);
}

/**
 * Fetch Partial Triage data from Partial Submits sheet
 */
function fetchPartialTriageData() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.PARTIAL_TRIAGE_SHEET_ID);
    const tab = sheet.getSheetByName(CONFIG.PARTIAL_TRIAGE_TAB_NAME);

    if (!tab) {
      throw new Error(`Tab "${CONFIG.PARTIAL_TRIAGE_TAB_NAME}" not found`);
    }

    const values = tab.getDataRange().getValues();
    if (values.length <= 1) {
      throw new Error(`No data found in ${CONFIG.PARTIAL_TRIAGE_TAB_NAME}`);
    }

    const sheetTimeZone = sheet.getSpreadsheetTimeZone();
    console.log('Partial Triage sheet timezone:', sheetTimeZone);

    // Convert Date objects back to string format if needed
    const formattedValues = values.map(row =>
      row.map(cell => {
        if (cell instanceof Date) {
          return Utilities.formatDate(cell, sheetTimeZone, 'M/d/yyyy');
        }
        return cell;
      })
    );

    return parsePartialTriageData(formattedValues);

  } catch (error) {
    throw new Error(`Failed to fetch Partial Triage data: ${error.message}`);
  }
}

/**
 * Parse Partial Triage data from Partial Submits sheet
 * Expected columns: Date
 * Logic: Count rows per date (each row = 1 partial triage)
 */
function parsePartialTriageData(values) {
  const [headers, ...data] = values;

  // Find column indices with flexible matching
  const dateIndex = findColumnIndex(headers, ['date']);

  console.log(`Partial Triage - Date: ${dateIndex}`);

  return data
    .map(row => {
      const dateValue = row[dateIndex];
      let parsedDate = null;

      // Handle the specific date format from your sheet (8/2/2025, 8/15/2025)
      if (dateValue instanceof Date) {
        parsedDate = new Date(dateValue);
      } else if (typeof dateValue === 'string') {
        // Parse M/D/YYYY format directly
        const parts = dateValue.split('/');
        if (parts.length === 3) {
          const month = parseInt(parts[0]);
          const day = parseInt(parts[1]);
          const year = parseInt(parts[2]);
          parsedDate = new Date(year, month - 1, day);
        }
      }

      return {
        date: parsedDate
      };
    })
    .filter(row => row.date && !isNaN(row.date.getTime()));
}

/**
 * Fetch Booked Closing Call data from Post Call Reports sheet
 * This reuses the existing SETS_SHEET_ID and SETS_TAB_NAME from CONFIG
 */
function fetchBookedClosingCallData() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SETS_SHEET_ID);
    const tab = sheet.getSheetByName(CONFIG.SETS_TAB_NAME);

    if (!tab) {
      throw new Error(`Tab "${CONFIG.SETS_TAB_NAME}" not found`);
    }

    const values = tab.getDataRange().getValues();
    if (values.length <= 1) {
      throw new Error(`No data found in ${CONFIG.SETS_TAB_NAME}`);
    }

    return parseBookedClosingCallData(values);

  } catch (error) {
    throw new Error(`Failed to fetch Booked Closing Call data: ${error.message}`);
  }
}

/**
 * Parse Booked Closing Call data from Post Call Reports sheet
 * Expected columns: Date Submitted, What's the new Lead Status
 * Filter: Status = "Discovery Call Complete"
 * Logic: Count rows per date where status = "Discovery Call Complete"
 */
function parseBookedClosingCallData(values) {
  const [headers, ...data] = values;

  // Find column indices with flexible matching (reusing existing logic)
  const dateIndex = findColumnIndex(headers, ['date submitted']);
  const statusIndex = findColumnIndex(headers, ["what's the new lead status"]);

  console.log(`Booked Closing Call - Date: ${dateIndex}, Status: ${statusIndex}`);

  return data
    .map(row => ({
      date: parseDate(row[dateIndex]), // Reuse existing parseDate function
      status: cleanStatusString(row[statusIndex]) // Reuse existing cleanStatusString function
    }))
    .filter(row =>
      row.date &&
      row.status === 'discoverycallcomplete' // Filter for Discovery Call Complete status
    );
}

/**
 * Fetch Shown Calls data from Post Call Reports sheet
 * This reuses the existing SETS_SHEET_ID and SETS_TAB_NAME from CONFIG
 */
function fetchShownCallsData() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SETS_SHEET_ID);
    const tab = sheet.getSheetByName(CONFIG.SETS_TAB_NAME);

    if (!tab) {
      throw new Error(`Tab "${CONFIG.SETS_TAB_NAME}" not found`);
    }

    const values = tab.getDataRange().getValues();
    if (values.length <= 1) {
      throw new Error(`No data found in ${CONFIG.SETS_TAB_NAME}`);
    }

    return parseShownCallsData(values);

  } catch (error) {
    throw new Error(`Failed to fetch Shown Calls data: ${error.message}`);
  }
}

/**
 * Parse Shown Calls data from Post Call Reports sheet
 * Expected columns: Date Submitted, What's the new Lead Status
 * Filter: Status matches any of the outcome statuses (WON/MARKER/LOST)
 * Logic: Count rows per date where status matches outcome statuses
 */
function parseShownCallsData(values) {
  const [headers, ...data] = values;

  // Find column indices with flexible matching (reusing existing logic)
  const dateIndex = findColumnIndex(headers, ['date submitted']);
  const statusIndex = findColumnIndex(headers, ["what's the new lead status"]);

  console.log(`Shown Calls - Date: ${dateIndex}, Status: ${statusIndex}`);

  // Define the valid outcome statuses (normalized to match cleanStatusString format)
  const outcomeStatuses = [
    'won-innercircle(paymentplan)',
    'won-lifetime(paymentplan)',
    'won-innercircle(cashpif)',
    'won-lifetime(cashpif)',
    'won-innercircle(financed)',
    'won-lifetime(financed)',
    'marker-innercircle',
    'marker-lifetime',
    'lost-innercircle(badlead)',
    'lost-innercircle(qualifiedlead)',
    'lost-innercircle(fianancing)',
    'lost-lifetime(badlead)',
    'lost-lifetime(qualifiedlead)',
    'lost-lifetime(fianancing)'
  ];

  return data
    .map(row => ({
      date: parseDate(row[dateIndex]), // Reuse existing parseDate function
      status: cleanStatusString(row[statusIndex]) // Reuse existing cleanStatusString function
    }))
    .filter(row =>
      row.date &&
      row.status &&
      outcomeStatuses.includes(row.status) // Filter for outcome statuses
    );
}

/**
 * Fetch Closes data from Post Call Reports sheet
 * This reuses the existing SETS_SHEET_ID and SETS_TAB_NAME from CONFIG
 */
function fetchClosesData() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SETS_SHEET_ID);
    const tab = sheet.getSheetByName(CONFIG.SETS_TAB_NAME);

    if (!tab) {
      throw new Error(`Tab "${CONFIG.SETS_TAB_NAME}" not found`);
    }

    const values = tab.getDataRange().getValues();
    if (values.length <= 1) {
      throw new Error(`No data found in ${CONFIG.SETS_TAB_NAME}`);
    }

    return parseClosesData(values);

  } catch (error) {
    throw new Error(`Failed to fetch Closes data: ${error.message}`);
  }
}

/**
 * Parse Closes data from Post Call Reports sheet
 * Expected columns: Date Submitted, What's the new Lead Status
 * Filter: Status matches WON statuses only
 * Logic: Count rows per date where status is a WON status
 */
function parseClosesData(values) {
  const [headers, ...data] = values;

  // Find column indices with flexible matching (reusing existing logic)
  const dateIndex = findColumnIndex(headers, ['date submitted']);
  const statusIndex = findColumnIndex(headers, ["what's the new lead status"]);

  console.log(`Closes - Date: ${dateIndex}, Status: ${statusIndex}`);

  // Define the valid WON statuses only (normalized to match cleanStatusString format)
  const wonStatuses = [
    'won-innercircle(paymentplan)',
    'won-lifetime(paymentplan)',
    'won-innercircle(cashpif)',
    'won-lifetime(cashpif)',
    'won-innercircle(financed)',
    'won-lifetime(financed)'
  ];

  return data
    .map(row => ({
      date: parseDate(row[dateIndex]), // Reuse existing parseDate function
      status: cleanStatusString(row[statusIndex]) // Reuse existing cleanStatusString function
    }))
    .filter(row =>
      row.date &&
      row.status &&
      wonStatuses.includes(row.status) // Filter for WON statuses only
    );
}

/**
 * Fetch New Cash data from Post Call Reports sheet
 * This reuses the existing SETS_SHEET_ID and SETS_TAB_NAME from CONFIG
 */
function fetchNewCashData() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SETS_SHEET_ID);
    const tab = sheet.getSheetByName(CONFIG.SETS_TAB_NAME);

    if (!tab) {
      throw new Error(`Tab "${CONFIG.SETS_TAB_NAME}" not found`);
    }

    const values = tab.getDataRange().getValues();
    if (values.length <= 1) {
      throw new Error(`No data found in ${CONFIG.SETS_TAB_NAME}`);
    }

    return parseNewCashData(values);

  } catch (error) {
    throw new Error(`Failed to fetch New Cash data: ${error.message}`);
  }
}

/**
 * Parse New Cash data from Post Call Reports sheet
 * Expected columns: Date Submitted, What's the new Lead Status, Cash Collected
 * Filter: Status matches WON and MARKER statuses
 * Logic: Sum Cash Collected per date where status matches qualifying statuses
 */
function parseNewCashData(values) {
  const [headers, ...data] = values;

  // Find column indices with flexible matching (reusing existing logic)
  const dateIndex = findColumnIndex(headers, ['date submitted']);
  const statusIndex = findColumnIndex(headers, ["what's the new lead status"]);
  const cashCollectedIndex = findColumnIndex(headers, ['cash collected']);

  console.log(`New Cash - Date: ${dateIndex}, Status: ${statusIndex}, Cash Collected: ${cashCollectedIndex}`);

  // Define the valid statuses for New Cash (WON and MARKER statuses)
  const newCashStatuses = [
    'won-innercircle(paymentplan)',
    'won-lifetime(paymentplan)',
    'won-innercircle(cashpif)',
    'won-lifetime(cashpif)',
    'won-innercircle(financed)',
    'won-lifetime(financed)',
    'marker-innercircle',
    'marker-lifetime'
  ];

  return data
    .map(row => ({
      date: parseDate(row[dateIndex]), // Reuse existing parseDate function
      status: cleanStatusString(row[statusIndex]), // Reuse existing cleanStatusString function
      cashCollected: parseFloat(row[cashCollectedIndex]) || 0 // Parse cash amount, default to 0
    }))
    .filter(row =>
      row.date &&
      row.status &&
      newCashStatuses.includes(row.status) && // Filter for qualifying statuses
      row.cashCollected >= 0 // Ensure valid cash amount
    );
}

/**
 * Fetch FU Cash data from Post Call Reports sheet
 * This reuses the existing SETS_SHEET_ID and SETS_TAB_NAME from CONFIG
 */
function fetchFUCashData() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SETS_SHEET_ID);
    const tab = sheet.getSheetByName(CONFIG.SETS_TAB_NAME);

    if (!tab) {
      throw new Error(`Tab "${CONFIG.SETS_TAB_NAME}" not found`);
    }

    const values = tab.getDataRange().getValues();
    if (values.length <= 1) {
      throw new Error(`No data found in ${CONFIG.SETS_TAB_NAME}`);
    }

    return parseFUCashData(values);

  } catch (error) {
    throw new Error(`Failed to fetch FU Cash data: ${error.message}`);
  }
}

/**
 * Parse FU Cash data from Post Call Reports sheet
 * Expected columns: Date Submitted, What's the new Lead Status, Cash Collected
 * Filter: Status = "FOLLOW UP PAYMENT"
 * Logic: Sum Cash Collected per date where status is FOLLOW UP PAYMENT
 */
function parseFUCashData(values) {
  const [headers, ...data] = values;

  // Find column indices with flexible matching (reusing existing logic)
  const dateIndex = findColumnIndex(headers, ['date submitted']);
  const statusIndex = findColumnIndex(headers, ["what's the new lead status"]);
  const cashCollectedIndex = findColumnIndex(headers, ['cash collected']);

  console.log(`FU Cash - Date: ${dateIndex}, Status: ${statusIndex}, Cash Collected: ${cashCollectedIndex}`);

  return data
    .map(row => ({
      date: parseDate(row[dateIndex]), // Reuse existing parseDate function
      status: cleanStatusString(row[statusIndex]), // Reuse existing cleanStatusString function
      cashCollected: parseFloat(row[cashCollectedIndex]) || 0 // Parse cash amount, default to 0
    }))
    .filter(row =>
      row.date &&
      row.status === 'followuppayment' && // Filter for FOLLOW UP PAYMENT status only
      row.cashCollected >= 0 // Ensure valid cash amount
    );
}

/**
 * Fetch Booked Triage data from Calendly Booked Calls sheet
 */
function fetchBookedTriageData() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.BOOKED_TRIAGE_SHEET_ID);
    const tab = sheet.getSheetByName(CONFIG.BOOKED_TRIAGE_TAB_NAME);

    if (!tab) {
      throw new Error(`Tab "${CONFIG.BOOKED_TRIAGE_TAB_NAME}" not found`);
    }

    const values = tab.getDataRange().getValues();
    if (values.length <= 1) {
      throw new Error(`No data found in ${CONFIG.BOOKED_TRIAGE_TAB_NAME}`);
    }

    return parseBookedTriageData(values);

  } catch (error) {
    throw new Error(`Failed to fetch Booked Triage data: ${error.message}`);
  }
}

/**
 * Parse Booked Triage data from Calendly Booked Calls sheet
 * Expected columns: Date (UTC format), Call Type
 * Filter: Call Type = "eMoney Inner Circle Discovery Call"
 * Logic: Count rows per date where call type matches
 */
function parseBookedTriageData(values) {
  const [headers, ...data] = values;

  // Find column indices with flexible matching
  const dateIndex = findColumnIndex(headers, ['date']);
  const callTypeIndex = findColumnIndex(headers, ['call type']);

  console.log(`Booked Triage - Date: ${dateIndex}, Call Type: ${callTypeIndex}`);

  return data
    .map(row => {
      const dateValue = row[dateIndex];
      const callType = cleanString(row[callTypeIndex]);

      // Convert UTC date to Eastern time and extract date only
      const easternDate = convertUTCToEasternDate(dateValue);

      return {
        date: easternDate,
        callType: callType
      };
    })
    .filter(row =>
      row.date &&
      row.callType.includes('emoney') &&
      row.callType.includes('inner') &&
      row.callType.includes('circle') &&
      row.callType.includes('discovery') &&
      row.callType.includes('call')
    );
}

/**
 * Convert UTC date string to Eastern timezone date
 * Handles format: 2025-06-26T21:10:00.000-04:00
 */
function convertUTCToEasternDate(dateValue) {
  if (!dateValue) return null;

  try {
    let parsedDate;

    if (dateValue instanceof Date) {
      parsedDate = dateValue;
    } else if (typeof dateValue === 'string') {
      // Handle ISO 8601 format with timezone offset
      parsedDate = new Date(dateValue);
    } else {
      return null;
    }

    if (isNaN(parsedDate.getTime())) {
      return null;
    }

    // Convert to Eastern timezone and get date only
    const easternDateString = Utilities.formatDate(parsedDate, 'America/New_York', 'yyyy-MM-dd');
    return new Date(easternDateString + 'T00:00:00');

  } catch (error) {
    console.error('Error parsing UTC date:', dateValue, error);
    return null;
  }
}

/**
 * Update the "Vortex Data" report
 */
function updateVortexData() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // Create or get the "Vortex Data" sheet
    let sheet = spreadsheet.getSheetByName('Vortex Data');
    if (!sheet) {
      sheet = spreadsheet.insertSheet('Vortex Data', 4);
    } else {
      spreadsheet.setActiveSheet(sheet);
    }

    console.log('Updating "Vortex Data" report...');

    // Clear existing data
    sheet.clear();

    // Generate daily data from September 1st to today with real data
    const dailyData = generateDailyVortexData();

    // Generate vortex report with daily data
    generateVortexReport(sheet, dailyData);

    console.log('"Vortex Data" report updated successfully');

  } catch (error) {
    console.error('Failed to update "Vortex Data" report:', error);
    throw error;
  }
}

/**
 * Main function to generate the Vortex Study Monthly report
*/
function generateVortexDataMonthly() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    showLoadingMessage(ui);

    let sheet = getOrCreateSheet(spreadsheet, 'Vortex Data Monthly');
    sheet.clear();

    // Generate monthly data from September 2025 to current month
    const monthlyData = generateMonthlyVortexData();

    // Generate vortex monthly report
    generateVortexMonthlyReport(sheet, monthlyData);

    showSuccessMessage(ui);

  } catch (error) {
    showErrorMessage(ui, error);
    console.error('Vortex Study Monthly report generation failed:', error);
  }
}

/**
 * Generate monthly Vortex data from September 2025 to current month
*/
function generateMonthlyVortexData() {
  const today = getCurrentDateNY();
  const startDate = new Date('2025-09-01T00:00:00');
  const monthlyData = [];

  // Fetch all data sources (reusing existing functions)
  const lowTicketData = fetchLowTicketBuysData();
  const partialTriageData = fetchPartialTriageData();
  const bookedTriageData = fetchBookedTriageData();
  const bookedClosingCallData = fetchBookedClosingCallData();
  const shownCallsData = fetchShownCallsData();
  const closesData = fetchClosesData();
  const newCashData = fetchNewCashData();
  const fuCashData = fetchFUCashData();

  // Create maps for aggregating data by month
  const discordJoinedMap = createMonthlyDataMap(lowTicketData, 'totalInvites');
  const lowTicketMap = createMonthlyDataMap(lowTicketData, 'lowTicket');
  const partialTriageMap = createMonthlyCountMap(partialTriageData);
  const bookedTriageMap = createMonthlyCountMap(bookedTriageData);
  const bookedClosingCallMap = createMonthlyCountMap(bookedClosingCallData);
  const shownCallsMap = createMonthlyCountMap(shownCallsData);
  const closesMap = createMonthlyCountMap(closesData);
  const newCashMap = createMonthlySumMap(newCashData, 'cashCollected');
  const fuCashMap = createMonthlySumMap(fuCashData, 'cashCollected');

  // Generate data for each month from Sept 2025 to current month
  const currentDate = new Date(startDate);
  while (currentDate <= today) {
    const year = currentDate.getFullYear();
    const month = currentDate.getMonth() + 1;
    const monthKey = `${year}-${month.toString().padStart(2, '0')}`;
    
    // Get data for this month, default to 0 if not found
    const discordJoins = discordJoinedMap.get(monthKey) || 0;
    const lowTicketBuys = lowTicketMap.get(monthKey) || 0;
    const partialTriage = partialTriageMap.get(monthKey) || 0;
    const bookedTriage = bookedTriageMap.get(monthKey) || 0;
    const totalTriageCalls = partialTriage + bookedTriage;
    const bookedClosingCall = bookedClosingCallMap.get(monthKey) || 0;
    const shownCalls = shownCallsMap.get(monthKey) || 0;
    const closes = closesMap.get(monthKey) || 0;
    const newCash = newCashMap.get(monthKey) || 0;
    const fuCash = fuCashMap.get(monthKey) || 0;

    monthlyData.push({
      date: new Date(currentDate),
      discordJoins: discordJoins,
      lowTicketBuys: lowTicketBuys,
      bookedTriage: bookedTriage,
      partialTriage: partialTriage,
      totalTriageCalls: totalTriageCalls,
      bookedClosingCall: bookedClosingCall,
      shownCalls: shownCalls,
      closes: closes,
      newCash: newCash,
      fuCash: fuCash
    });

    // Move to next month
    currentDate.setMonth(currentDate.getMonth() + 1);
  }

  return monthlyData;
}

/**
 * Create monthly data map for summing values by month
 */
function createMonthlyDataMap(data, valueField) {
  const map = new Map();
  
  data.forEach(record => {
    if (record.date) {
      const year = record.date.getFullYear();
      const month = record.date.getMonth() + 1;
      const monthKey = `${year}-${month.toString().padStart(2, '0')}`;
      
      const currentValue = map.get(monthKey) || 0;
      map.set(monthKey, currentValue + (record[valueField] || 0));
    }
  });
  
  return map;
}

/**
 * Create monthly count map for counting occurrences by month
 */
function createMonthlyCountMap(data) {
  const map = new Map();
  
  data.forEach(record => {
    if (record.date) {
      const year = record.date.getFullYear();
      const month = record.date.getMonth() + 1;
      const monthKey = `${year}-${month.toString().padStart(2, '0')}`;
      
      const currentCount = map.get(monthKey) || 0;
      map.set(monthKey, currentCount + 1);
    }
  });
  
  return map;
}

/**
 * Create monthly sum map for summing cash values by month
 */
function createMonthlySumMap(data, valueField) {
  const map = new Map();
  
  data.forEach(record => {
    if (record.date) {
      const year = record.date.getFullYear();
      const month = record.date.getMonth() + 1;
      const monthKey = `${year}-${month.toString().padStart(2, '0')}`;
      
      const currentValue = map.get(monthKey) || 0;
      map.set(monthKey, currentValue + (record[valueField] || 0));
    }
  });
  
  return map;
}

/**
 * Generate the complete Vortex Study Monthly report
 */
function generateVortexMonthlyReport(sheet, monthlyData) {
  setupVortexMonthlyHeaders(sheet);

  let currentRow = 3;
  let previousMonthData = null;

  // Populate each month's data
  monthlyData.forEach((monthData, index) => {
    populateVortexMonthlyData(sheet, currentRow, monthData, previousMonthData);
    previousMonthData = monthData;
    currentRow += 1;
  });

  formatVortexMonthlySheet(sheet);
}

/**
 * Setup Vortex Study Monthly report headers
*/
function setupVortexMonthlyHeaders(sheet) {
  const mainHeader = sheet.getRange('A1:L1');
  mainHeader.merge()
    .setValue('Vortex Study Report - Monthly View')
    .setBackground('#4a90e2')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(14);

  // Column headers
  const headers = [
    'Date', 'Discord Joins', 'Low Ticket Buys', 'Conversion Rate', 'Booked Triage', 'Partial Triage',
    'Total Triage Calls', 'Booked Closing Call', 'Shown Calls', 'Closes', 'New Cash', 'FU Cash'
  ];
  const headerRange = sheet.getRange(2, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground('#d9e2f3')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

/**
 * Populate data for a specific month (Vortex Study Monthly)
 */
function populateVortexMonthlyData(sheet, row, monthData, previousMonthData = null) {
  // Format date as "Month YYYY" (e.g., "September 2025")
  const formattedDate = Utilities.formatDate(monthData.date, 'America/New_York', 'MMMM yyyy');

  const discordJoins = monthData.discordJoins || 0;
  const lowTicketBuys = monthData.lowTicketBuys || 0;
  const bookedTriage = monthData.bookedTriage || 0;
  const partialTriage = monthData.partialTriage || 0;
  const totalTriageCalls = monthData.totalTriageCalls || 0;
  const bookedClosingCall = monthData.bookedClosingCall || 0;
  const shownCalls = monthData.shownCalls || 0;
  const closes = monthData.closes || 0;
  const newCash = monthData.newCash || 0;
  const fuCash = monthData.fuCash || 0;
  const conversionRate = discordJoins > 0 ? `${Math.round((lowTicketBuys / discordJoins) * 100)}%` : '0%';

  const formattedDiscordJoins = discordJoins > 0 ? discordJoins.toLocaleString() : discordJoins;
  const formattedlowTicketBuys = lowTicketBuys > 0 ? lowTicketBuys.toLocaleString() : lowTicketBuys;
  const formattedbookedTriage = bookedTriage > 0 ? bookedTriage.toLocaleString() : bookedTriage;
  const formattedpartialTriage = partialTriage > 0 ? partialTriage.toLocaleString() : partialTriage;
  const formattedtotalTriageCalls = totalTriageCalls > 0 ? totalTriageCalls.toLocaleString() : totalTriageCalls;

  const rowData = [
    formattedDate,
    formattedDiscordJoins,
    formattedlowTicketBuys,
    conversionRate,
    formattedbookedTriage,
    formattedpartialTriage,
    formattedtotalTriageCalls,
    bookedClosingCall,
    shownCalls,
    closes,
    newCash > 0 ? `$${newCash.toLocaleString()}` : newCash,
    fuCash > 0 ? `$${fuCash.toLocaleString()}` : fuCash
  ];

  // Write data to sheet
  sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);

  if (previousMonthData) {
    const prevConversionRate = previousMonthData.discordJoins > 0 ? (previousMonthData.lowTicketBuys / previousMonthData.discordJoins) * 100 : 0;
    const currConversionRate = discordJoins > 0 ? (lowTicketBuys / discordJoins) * 100 : 0;
    
    const currentValues = [discordJoins, lowTicketBuys, currConversionRate, bookedTriage, partialTriage, totalTriageCalls, bookedClosingCall, shownCalls, closes, newCash, fuCash];
    const previousValues = [
      previousMonthData.discordJoins || 0,
      previousMonthData.lowTicketBuys || 0,
      prevConversionRate,
      previousMonthData.bookedTriage || 0,
      previousMonthData.partialTriage || 0,
      previousMonthData.totalTriageCalls || 0,
      previousMonthData.bookedClosingCall || 0,
      previousMonthData.shownCalls || 0,
      previousMonthData.closes || 0,
      previousMonthData.newCash || 0,
      previousMonthData.fuCash || 0
    ];

    // Apply formatting to columns 2-12 (skip date column)
    for (let col = 2; col <= 12; col++) {
      const currentValue = currentValues[col - 2];
      const previousValue = previousValues[col - 2];

      let backgroundColor = null;

      if (currentValue > previousValue) {
        backgroundColor = '#07fc03'; // Green
      } else if (currentValue < previousValue) {
        backgroundColor = '#ff0000'; // Red
      } else if (currentValue === previousValue) {
        backgroundColor = '#ffff00'; // Yellow
      }

      if (backgroundColor) {
        sheet.getRange(row, col).setBackground(backgroundColor);
      }
    }
  }
}

/**
 * Format the Vortex Study Monthly sheet for better readability
 */
function formatVortexMonthlySheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Set specific column widths for better visibility
  sheet.setColumnWidth(1, 120);  // Date column
  sheet.setColumnWidth(2, 120);  // Discord Joins
  sheet.setColumnWidth(3, 130);  // Low Ticket Buys
  sheet.setColumnWidth(4, 130);  // Conversion Rate
  sheet.setColumnWidth(5, 120);  // Booked Triage
  sheet.setColumnWidth(6, 120);  // Partial Triage
  sheet.setColumnWidth(7, 140);  // Total Triage Calls
  sheet.setColumnWidth(8, 150);  // Booked Closing Call
  sheet.setColumnWidth(9, 110);  // Shown Calls
  sheet.setColumnWidth(10, 80);   // Closes
  sheet.setColumnWidth(11, 120); // New Cash
  sheet.setColumnWidth(12, 100); // FU Cash

  sheet.getRange(1, 1, lastRow, lastCol)
    .setBorder(true, true, true, true, true, true);

  sheet.getRange(3, 1, lastRow - 2, lastCol).setHorizontalAlignment('center');

  sheet.setFrozenRows(2);
}

/**
 * Update the "Vortex Data Monthly" report
*/
function updateVortexDataMonthly() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // Create or get the "Vortex Data Monthly" sheet
    let sheet = spreadsheet.getSheetByName('Vortex Data Monthly');
    if (!sheet) {
      sheet = spreadsheet.insertSheet('Vortex Data Monthly', 5);
    } else {
      spreadsheet.setActiveSheet(sheet);
    }

    console.log('Updating "Vortex Data Monthly" report...');

    // Clear existing data
    sheet.clear();

    // Generate monthly data from September 2025 to current month
    const monthlyData = generateMonthlyVortexData();

    // Generate vortex monthly report
    generateVortexMonthlyReport(sheet, monthlyData);

    console.log('"Vortex Data Monthly" report updated successfully');

  } catch (error) {
    console.error('Failed to update "Vortex Data Monthly" report:', error);
    throw error;
  }
}

/**
 * Main function to generate the Vortex Study Weekly report
 */
function generateVortexDataWeekly() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    showLoadingMessage(ui);

    let sheet = getOrCreateSheet(spreadsheet, 'Vortex Data Weekly');
    sheet.clear();

    // Generate weekly data from first Friday-Thursday period to current week
    const weeklyData = generateWeeklyVortexData();

    // Generate vortex weekly report
    generateVortexWeeklyReport(sheet, weeklyData);

    showSuccessMessage(ui);

  } catch (error) {
    showErrorMessage(ui, error);
    console.error('Vortex Study Weekly report generation failed:', error);
  }
}

/**
 * Generate weekly Vortex data from September 2025 onwards (Friday to Thursday)
 */
function generateWeeklyVortexData() {
  const today = getCurrentDateNY();
  const weeklyData = [];

  // Fetch all data sources (reusing existing functions)
  const lowTicketData = fetchLowTicketBuysData();
  const partialTriageData = fetchPartialTriageData();
  const bookedTriageData = fetchBookedTriageData();
  const bookedClosingCallData = fetchBookedClosingCallData();
  const shownCallsData = fetchShownCallsData();
  const closesData = fetchClosesData();
  const newCashData = fetchNewCashData();
  const fuCashData = fetchFUCashData();

  // Create maps for aggregating data by week
  const discordJoinedMap = createWeeklyDataMap(lowTicketData, 'totalInvites');
  const lowTicketMap = createWeeklyDataMap(lowTicketData, 'lowTicket');
  const partialTriageMap = createWeeklyCountMap(partialTriageData);
  const bookedTriageMap = createWeeklyCountMap(bookedTriageData);
  const bookedClosingCallMap = createWeeklyCountMap(bookedClosingCallData);
  const shownCallsMap = createWeeklyCountMap(shownCallsData);
  const closesMap = createWeeklyCountMap(closesData);
  const newCashMap = createWeeklySumMap(newCashData, 'cashCollected');
  const fuCashMap = createWeeklySumMap(fuCashData, 'cashCollected');

  // Find first Friday on or after September 1st, 2025
  const startDate = findFirstFriday(new Date('2025-08-29T00:00:00'));
  
  // Generate data for each week (Friday to Thursday) until today
  let weekStart = new Date(startDate);
  
  while (weekStart <= today) {
    let weekEnd = new Date(weekStart);
    weekEnd.setDate(weekEnd.getDate() + 6); // Thursday of the same week
    
    const weekKey = formatWeekKey(weekStart);
    
    // Get data for this week, default to 0 if not found
    const discordJoins = discordJoinedMap.get(weekKey) || 0;
    const lowTicketBuys = lowTicketMap.get(weekKey) || 0;
    const partialTriage = partialTriageMap.get(weekKey) || 0;
    const bookedTriage = bookedTriageMap.get(weekKey) || 0;
    const totalTriageCalls = partialTriage + bookedTriage;
    const bookedClosingCall = bookedClosingCallMap.get(weekKey) || 0;
    const shownCalls = shownCallsMap.get(weekKey) || 0;
    const closes = closesMap.get(weekKey) || 0;
    const newCash = newCashMap.get(weekKey) || 0;
    const fuCash = fuCashMap.get(weekKey) || 0;

    weeklyData.push({
      weekStart: new Date(weekStart),
      weekEnd: new Date(weekEnd),
      discordJoins: discordJoins,
      lowTicketBuys: lowTicketBuys,
      bookedTriage: bookedTriage,
      partialTriage: partialTriage,
      totalTriageCalls: totalTriageCalls,
      bookedClosingCall: bookedClosingCall,
      shownCalls: shownCalls,
      closes: closes,
      newCash: newCash,
      fuCash: fuCash
    });

    // Move to next Friday
    weekStart.setDate(weekStart.getDate() + 7);
  }

  return weeklyData;
}

/**
 * Find the first Friday on or after the given date
 */
function findFirstFriday(date) {
  const dayOfWeek = date.getDay(); // 0 = Sunday, 5 = Friday
  const daysUntilFriday = (5 - dayOfWeek + 7) % 7;
  const firstFriday = new Date(date);
  firstFriday.setDate(date.getDate() + daysUntilFriday);
  return firstFriday;
}

/**
 * Format week key for Friday start date (YYYY-MM-DD format)
 */
function formatWeekKey(fridayDate) {
  const year = fridayDate.getFullYear();
  const month = (fridayDate.getMonth() + 1).toString().padStart(2, '0');
  const day = fridayDate.getDate().toString().padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Get the Friday start date for any given date
 */
function getWeekStartFriday(date) {
  const dayOfWeek = date.getDay(); // 0 = Sunday, 5 = Friday
  let daysFromFriday;
  
  if (dayOfWeek >= 5) { // Friday or Saturday
    daysFromFriday = dayOfWeek - 5;
  } else { // Sunday through Thursday
    daysFromFriday = dayOfWeek + 2; // Days since last Friday
  }
  
  const weekStart = new Date(date);
  weekStart.setDate(date.getDate() - daysFromFriday);
  return weekStart;
}

/**
 * Create weekly data map for summing values by week
 */
function createWeeklyDataMap(data, valueField) {
  const map = new Map();
  
  data.forEach(record => {
    if (record.date) {
      const weekStart = getWeekStartFriday(record.date);
      const weekKey = formatWeekKey(weekStart);
      
      const currentValue = map.get(weekKey) || 0;
      map.set(weekKey, currentValue + (record[valueField] || 0));
    }
  });
  
  return map;
}

/**
 * Create weekly count map for counting occurrences by week
 */
function createWeeklyCountMap(data) {
  const map = new Map();
  
  data.forEach(record => {
    if (record.date) {
      const weekStart = getWeekStartFriday(record.date);
      const weekKey = formatWeekKey(weekStart);
      
      const currentCount = map.get(weekKey) || 0;
      map.set(weekKey, currentCount + 1);
    }
  });
  
  return map;
}

/**
 * Create weekly sum map for summing cash values by week
 */
function createWeeklySumMap(data, valueField) {
  const map = new Map();
  
  data.forEach(record => {
    if (record.date) {
      const weekStart = getWeekStartFriday(record.date);
      const weekKey = formatWeekKey(weekStart);
      
      const currentValue = map.get(weekKey) || 0;
      map.set(weekKey, currentValue + (record[valueField] || 0));
    }
  });
  
  return map;
}

/**
 * Generate the complete Vortex Study Weekly report
 */
function generateVortexWeeklyReport(sheet, weeklyData) {
  setupVortexWeeklyHeaders(sheet);

  let currentRow = 3;
  let previousWeekData = null;

  // Populate each week's data
  weeklyData.forEach((weekData, index) => {
    populateVortexWeeklyData(sheet, currentRow, weekData, previousWeekData);
    previousWeekData = weekData;
    currentRow += 1;
  });

  formatVortexWeeklySheet(sheet);
}

/**
 * Setup Vortex Study Weekly report headers
 */
function setupVortexWeeklyHeaders(sheet) {
  // Main header
  const mainHeader = sheet.getRange('A1:L1');
  mainHeader.merge()
    .setValue('Vortex Study Report - Weekly View')
    .setBackground('#4a90e2')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(14);

  // Column headers
  const headers = [
    'Date', 'Discord Joins', 'Low Ticket Buys', 'Conversion Rate', 'Booked Triage', 'Partial Triage',
    'Total Triage Calls', 'Booked Closing Call', 'Shown Calls', 'Closes', 'New Cash', 'FU Cash'
  ];
  const headerRange = sheet.getRange(2, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground('#d9e2f3')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

/**
 * Populate data for a specific week (Vortex Study Weekly)
 */
function populateVortexWeeklyData(sheet, row, weekData, previousWeekData = null) {
  // Format date range as "M/D/YYYY - M/D/YYYY" (e.g., "8/29/2025 - 9/4/2025")
  const startFormatted = Utilities.formatDate(weekData.weekStart, 'America/New_York', 'M/d/yyyy');
  const endFormatted = Utilities.formatDate(weekData.weekEnd, 'America/New_York', 'M/d/yyyy');
  const formattedDate = `${startFormatted} - ${endFormatted}`;

  const discordJoins = weekData.discordJoins || 0;
  const lowTicketBuys = weekData.lowTicketBuys || 0;
  const bookedTriage = weekData.bookedTriage || 0;
  const partialTriage = weekData.partialTriage || 0;
  const totalTriageCalls = weekData.totalTriageCalls || 0;
  const bookedClosingCall = weekData.bookedClosingCall || 0;
  const shownCalls = weekData.shownCalls || 0;
  const closes = weekData.closes || 0;
  const newCash = weekData.newCash || 0;
  const fuCash = weekData.fuCash || 0;
  const conversionRate = discordJoins > 0 ? `${Math.round((lowTicketBuys / discordJoins) * 100)}%` : '0%';

  const formattedDiscordJoins = discordJoins > 0 ? discordJoins.toLocaleString() : discordJoins;
  const formattedlowTicketBuys = lowTicketBuys > 0 ? lowTicketBuys.toLocaleString() : lowTicketBuys;
  const formattedbookedTriage = bookedTriage > 0 ? bookedTriage.toLocaleString() : bookedTriage;
  const formattedpartialTriage = partialTriage > 0 ? partialTriage.toLocaleString() : partialTriage;
  const formattedtotalTriageCalls = totalTriageCalls > 0 ? totalTriageCalls.toLocaleString() : totalTriageCalls;

  const rowData = [
    formattedDate,
    formattedDiscordJoins,
    formattedlowTicketBuys,
    conversionRate,
    formattedbookedTriage,
    formattedpartialTriage,
    formattedtotalTriageCalls,
    bookedClosingCall,
    shownCalls,
    closes,
    newCash > 0 ? `$${newCash.toLocaleString()}` : newCash,
    fuCash > 0 ? `$${fuCash.toLocaleString()}` : fuCash
  ];

  // Write data to sheet
  sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);

  if (previousWeekData) {
    const prevConversionRate = previousWeekData.discordJoins > 0 ? (previousWeekData.lowTicketBuys / previousWeekData.discordJoins) * 100 : 0;
    const currConversionRate = discordJoins > 0 ? (lowTicketBuys / discordJoins) * 100 : 0;
    
    const currentValues = [discordJoins, lowTicketBuys, currConversionRate, bookedTriage, partialTriage, totalTriageCalls, bookedClosingCall, shownCalls, closes, newCash, fuCash];
    const previousValues = [
      previousWeekData.discordJoins || 0,
      previousWeekData.lowTicketBuys || 0,
      prevConversionRate,
      previousWeekData.bookedTriage || 0,
      previousWeekData.partialTriage || 0,
      previousWeekData.totalTriageCalls || 0,
      previousWeekData.bookedClosingCall || 0,
      previousWeekData.shownCalls || 0,
      previousWeekData.closes || 0,
      previousWeekData.newCash || 0,
      previousWeekData.fuCash || 0
    ];

    // Apply formatting to columns 2-12 (skip date column)
    for (let col = 2; col <= 12; col++) {
      const currentValue = currentValues[col - 2];
      const previousValue = previousValues[col - 2];

      let backgroundColor = null;

      if (currentValue > previousValue) {
        backgroundColor = '#07fc03'; // Green
      } else if (currentValue < previousValue) {
        backgroundColor = '#ff0000'; // Red
      } else if (currentValue === previousValue) {
        backgroundColor = '#ffff00'; // Yellow
      }

      if (backgroundColor) {
        sheet.getRange(row, col).setBackground(backgroundColor);
      }
    }
  }
}

/**
 * Format the Vortex Study Weekly sheet for better readability
 */
function formatVortexWeeklySheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Set specific column widths for better visibility
  sheet.setColumnWidth(1, 160);  // Date column (wider for date range)
  sheet.setColumnWidth(2, 120);  // Discord Joins
  sheet.setColumnWidth(3, 130);  // Low Ticket Buys
  sheet.setColumnWidth(4, 130);  // Conversion Rate
  sheet.setColumnWidth(5, 120);  // Booked Triage
  sheet.setColumnWidth(6, 120);  // Partial Triage
  sheet.setColumnWidth(7, 140);  // Total Triage Calls
  sheet.setColumnWidth(8, 150);  // Booked Closing Call
  sheet.setColumnWidth(9, 110);  // Shown Calls
  sheet.setColumnWidth(10, 80);   // Closes
  sheet.setColumnWidth(11, 120); // New Cash
  sheet.setColumnWidth(12, 100); // FU Cash

  sheet.getRange(1, 1, lastRow, lastCol)
    .setBorder(true, true, true, true, true, true);

  sheet.getRange(3, 1, lastRow - 2, lastCol).setHorizontalAlignment('center');

  sheet.setFrozenRows(2);
}

/**
 * Update the "Vortex Data Weekly" report
 */
function updateVortexDataWeekly() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // Create or get the "Vortex Data Weekly" sheet
    let sheet = spreadsheet.getSheetByName('Vortex Data Weekly');
    if (!sheet) {
      sheet = spreadsheet.insertSheet('Vortex Data Weekly', 5);
    } else {
      spreadsheet.setActiveSheet(sheet);
    }

    console.log('Updating "Vortex Data Weekly" report...');

    // Clear existing data
    sheet.clear();

    // Generate weekly data from first Friday to current week
    const weeklyData = generateWeeklyVortexData();

    // Generate vortex weekly report
    generateVortexWeeklyReport(sheet, weeklyData);

    console.log('"Vortex Data Weekly" report updated successfully');

  } catch (error) {
    console.error('Failed to update "Vortex Data Weekly" report:', error);
    throw error;
  }
}