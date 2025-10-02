function updateDailyDash() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Lead Entered");
  const targetSheet = ss.getSheetByName("Daily Dash");

  const dateColIndex = 0;  // Column A - Date Of Meeting
  const columnF_Index = 5;  // Column F - Responded to 1st Text
  const columnQ_Index = 16; // Column Q - Responded to Morning Text
  const columnV_Index = 21; // Column V - Status

  const data = sourceSheet.getDataRange().getValues();
  const dateStats = {};

  for (let i = 1; i < data.length; i++) {
    const dateValue = data[i][dateColIndex];

    if (!dateValue) continue;

    try {
      const date = new Date(dateValue);
      const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "M/d/yyyy");

      if (!dateStats[dateStr]) {
        dateStats[dateStr] = {
          noShowRespondedFirst: 0,
          noShowRespondedMorning: 0,
          showedRespondedFirst: 0,
          showedRespondedMorning: 0,
          totalCalls: 0,
          rescheduled: 0,
          canceled: 0
        };
      }

      const columnF = String(data[i][columnF_Index]).toLowerCase().trim();
      const columnQ = String(data[i][columnQ_Index]).toLowerCase().trim();
      const columnV = String(data[i][columnV_Index]).replace(/\s+/g, '').toLowerCase();

      const isShowedStatus = columnV.includes("won-") ||
        columnV.includes("lost-") ||
        columnV.includes("marker-") ||
        columnV.includes("followup");

      if (columnF === "yes" && columnV === "noshow") {
        dateStats[dateStr].noShowRespondedFirst++;
      }

      if (columnQ === "yes" && columnV === "noshow") {
        dateStats[dateStr].noShowRespondedMorning++;
      }

      if (columnF === "yes" && isShowedStatus) {
        dateStats[dateStr].showedRespondedFirst++;
      }

      if (columnQ === "yes" && isShowedStatus) {
        dateStats[dateStr].showedRespondedMorning++;
      }

      dateStats[dateStr].totalCalls++;

      if (columnV.includes("rescheduled")) {
        dateStats[dateStr].rescheduled++;
      }

      if (columnV.includes("cancelled") || columnV.includes("canceled")) {
        dateStats[dateStr].canceled++;
      }

    } catch (e) {
      continue;
    }
  }

  const sortedDates = Object.keys(dateStats).sort((a, b) => {
    return new Date(a) - new Date(b);
  });

  targetSheet.clear();

  const headerRow1 = ["Date", "No Shows", "No Shows", "Showed Calls", "Showed Calls", "Total", "Total", "Total"];
  const headerRow2 = ["", "Responded to 1st Text", "Responded to Morning Text", "Responded to 1st Text", "Responded to Morning Text", "Total Calls for the Day", "Rescheduled", "Canceled"];

  targetSheet.appendRow(headerRow1);
  targetSheet.appendRow(headerRow2);

  // Merge parent headers
  targetSheet.getRange(1, 1, 2, 1).merge(); // Merge "Date" across 2 rows
  targetSheet.getRange(1, 2, 1, 2).merge(); // No Shows
  targetSheet.getRange(1, 4, 1, 2).merge(); // Showed Calls
  targetSheet.getRange(1, 6, 1, 3).merge(); // Total

  targetSheet.getRange(1, 1, 2, 8).setHorizontalAlignment("center").setVerticalAlignment("middle");
  targetSheet.getRange(1, 1, 2, 8).setFontWeight("bold");

  const columnWidths = [100, 150, 190, 150, 190, 170, 110, 100];
  columnWidths.forEach((width, i) => {
    targetSheet.setColumnWidth(i + 1, width);
  });

  let previousStats = null;

  sortedDates.forEach((date, index) => {
    const stats = dateStats[date];
    const rowData = [
      date,
      stats.noShowRespondedFirst,
      stats.noShowRespondedMorning,
      stats.showedRespondedFirst,
      stats.showedRespondedMorning,
      stats.totalCalls,
      stats.rescheduled,
      stats.canceled
    ];

    targetSheet.appendRow(rowData);

    if (index > 0 && previousStats) {
      const currentRow = index + 3;

      for (let col = 2; col <= 8; col++) {
        const currentValue = rowData[col - 1];
        const previousValue = col === 2 ? previousStats.noShowRespondedFirst :
          col === 3 ? previousStats.noShowRespondedMorning :
            col === 4 ? previousStats.showedRespondedFirst :
              col === 5 ? previousStats.showedRespondedMorning :
                col === 6 ? previousStats.totalCalls :
                  col === 7 ? previousStats.rescheduled :
                    previousStats.canceled;

        let color;
        if (currentValue > previousValue) {
          color = '#90EE90'; // Light green
        } else if (currentValue === previousValue) {
          color = '#FFFF99'; // Light yellow
        } else {
          color = '#FFB6C1'; // Light red
        }

        targetSheet.getRange(currentRow, col).setBackground(color);
      }
    } else if (index === 0) {
      const currentRow = 3;
      for (let col = 2; col <= 8; col++) {
        targetSheet.getRange(currentRow, col).setBackground('#FFFFFF');
      }
    }

    previousStats = stats;
  });

  if (sortedDates.length > 0) {
    targetSheet.getRange(1, 1, sortedDates.length + 2, 8).setBorder(true, true, true, true, true, true);
  }
}

function updateWeeklyDash() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Lead Entered");
  const targetSheet = ss.getSheetByName("Weekly Dash");

  const dateColIndex = 0;  // Column A - Date Of Meeting
  const columnF_Index = 5;  // Column F - Responded to 1st Text
  const columnQ_Index = 16; // Column Q - Responded to Morning Text
  const columnV_Index = 21; // Column V - Status

  const data = sourceSheet.getDataRange().getValues();
  const weekStats = {};

  for (let i = 1; i < data.length; i++) {
    const dateValue = data[i][dateColIndex];

    if (!dateValue) continue;

    try {
      const date = new Date(dateValue);

      const weekStart = new Date(date);
      const dayOfWeek = weekStart.getDay();
      weekStart.setDate(weekStart.getDate() - dayOfWeek);
      weekStart.setHours(0, 0, 0, 0);

      const weekEnd = new Date(weekStart);
      weekEnd.setDate(weekEnd.getDate() + 6);

      const weekStartStr = Utilities.formatDate(weekStart, Session.getScriptTimeZone(), "M/d/yyyy");
      const weekEndStr = Utilities.formatDate(weekEnd, Session.getScriptTimeZone(), "M/d/yyyy");
      const weekKey = weekStartStr + " - " + weekEndStr;

      if (!weekStats[weekKey]) {
        weekStats[weekKey] = {
          weekStart: weekStart,
          noShowRespondedFirst: 0,
          noShowRespondedMorning: 0,
          showedRespondedFirst: 0,
          showedRespondedMorning: 0,
          totalCalls: 0,
          rescheduled: 0,
          canceled: 0
        };
      }

      const columnF = String(data[i][columnF_Index]).toLowerCase().trim();
      const columnQ = String(data[i][columnQ_Index]).toLowerCase().trim();
      const columnV = String(data[i][columnV_Index]).replace(/\s+/g, '').toLowerCase();

      const isShowedStatus = columnV.includes("won-") ||
        columnV.includes("lost-") ||
        columnV.includes("marker-") ||
        columnV.includes("followup");

      if (columnF === "yes" && columnV === "noshow") {
        weekStats[weekKey].noShowRespondedFirst++;
      }

      if (columnQ === "yes" && columnV === "noshow") {
        weekStats[weekKey].noShowRespondedMorning++;
      }

      if (columnF === "yes" && isShowedStatus) {
        weekStats[weekKey].showedRespondedFirst++;
      }

      if (columnQ === "yes" && isShowedStatus) {
        weekStats[weekKey].showedRespondedMorning++;
      }

      weekStats[weekKey].totalCalls++;

      if (columnV.includes("rescheduled")) {
        weekStats[weekKey].rescheduled++;
      }

      if (columnV.includes("cancelled") || columnV.includes("canceled")) {
        weekStats[weekKey].canceled++;
      }

    } catch (e) {
      continue;
    }
  }

  const sortedWeeks = Object.keys(weekStats).sort((a, b) => {
    return weekStats[a].weekStart - weekStats[b].weekStart;
  });

  targetSheet.clear();

  const headerRow1 = ["Week", "No Shows", "No Shows", "Showed Calls", "Showed Calls", "Total", "Total", "Total"];
  const headerRow2 = ["", "Responded to 1st Text", "Responded to Morning Text", "Responded to 1st Text", "Responded to Morning Text", "Total Calls for the Week", "Rescheduled", "Canceled"];

  targetSheet.appendRow(headerRow1);
  targetSheet.appendRow(headerRow2);

  targetSheet.getRange(1, 1, 2, 1).merge(); // Merge "Week" across 2 rows
  targetSheet.getRange(1, 2, 1, 2).merge(); // No Shows
  targetSheet.getRange(1, 4, 1, 2).merge(); // Showed Calls
  targetSheet.getRange(1, 6, 1, 3).merge(); // Total

  targetSheet.getRange(1, 1, 2, 8).setHorizontalAlignment("center").setVerticalAlignment("middle");
  targetSheet.getRange(1, 1, 2, 8).setFontWeight("bold");

  const columnWidths = [200, 150, 190, 150, 190, 170, 110, 100];
  columnWidths.forEach((width, i) => {
    targetSheet.setColumnWidth(i + 1, width);
  });

  let previousStats = null;

  sortedWeeks.forEach((weekKey, index) => {
    const stats = weekStats[weekKey];
    const rowData = [
      weekKey,
      stats.noShowRespondedFirst,
      stats.noShowRespondedMorning,
      stats.showedRespondedFirst,
      stats.showedRespondedMorning,
      stats.totalCalls,
      stats.rescheduled,
      stats.canceled
    ];

    targetSheet.appendRow(rowData);

    if (index > 0 && previousStats) {
      const currentRow = index + 3;

      for (let col = 2; col <= 8; col++) {
        const currentValue = rowData[col - 1];
        const previousValue = col === 2 ? previousStats.noShowRespondedFirst :
          col === 3 ? previousStats.noShowRespondedMorning :
            col === 4 ? previousStats.showedRespondedFirst :
              col === 5 ? previousStats.showedRespondedMorning :
                col === 6 ? previousStats.totalCalls :
                  col === 7 ? previousStats.rescheduled :
                    previousStats.canceled;

        let color;
        if (currentValue > previousValue) {
          color = '#90EE90'; // Light green
        } else if (currentValue === previousValue) {
          color = '#FFFF99'; // Light yellow
        } else {
          color = '#FFB6C1'; // Light red
        }

        targetSheet.getRange(currentRow, col).setBackground(color);
      }
    } else if (index === 0) {
      const currentRow = 3;
      for (let col = 2; col <= 8; col++) {
        targetSheet.getRange(currentRow, col).setBackground('#FFFFFF');
      }
    }

    previousStats = stats;
  });

  if (sortedWeeks.length > 0) {
    targetSheet.getRange(1, 1, sortedWeeks.length + 2, 8).setBorder(true, true, true, true, true, true);
  }
}

function updateMonthlyDash() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Lead Entered");
  const targetSheet = ss.getSheetByName("Monthly Dash");

  const dateColIndex = 0;  // Column A - Date Of Meeting
  const columnF_Index = 5;  // Column F - Responded to 1st Text
  const columnQ_Index = 16; // Column Q - Responded to Morning Text
  const columnV_Index = 21; // Column V - Status

  const data = sourceSheet.getDataRange().getValues();
  const monthStats = {};

  for (let i = 1; i < data.length; i++) {
    const dateValue = data[i][dateColIndex];

    if (!dateValue) continue;

    try {
      const date = new Date(dateValue);

      const monthKey = Utilities.formatDate(date, Session.getScriptTimeZone(), "MMMM yyyy");

      const monthStart = new Date(date.getFullYear(), date.getMonth(), 1);

      if (!monthStats[monthKey]) {
        monthStats[monthKey] = {
          monthStart: monthStart,
          noShowRespondedFirst: 0,
          noShowRespondedMorning: 0,
          showedRespondedFirst: 0,
          showedRespondedMorning: 0,
          totalCalls: 0,
          rescheduled: 0,
          canceled: 0
        };
      }

      const columnF = String(data[i][columnF_Index]).toLowerCase().trim();
      const columnQ = String(data[i][columnQ_Index]).toLowerCase().trim();
      const columnV = String(data[i][columnV_Index]).replace(/\s+/g, '').toLowerCase();

      const isShowedStatus = columnV.includes("won-") ||
        columnV.includes("lost-") ||
        columnV.includes("marker-") ||
        columnV.includes("followup");

      if (columnF === "yes" && columnV === "noshow") {
        monthStats[monthKey].noShowRespondedFirst++;
      }

      if (columnQ === "yes" && columnV === "noshow") {
        monthStats[monthKey].noShowRespondedMorning++;
      }

      if (columnF === "yes" && isShowedStatus) {
        monthStats[monthKey].showedRespondedFirst++;
      }

      if (columnQ === "yes" && isShowedStatus) {
        monthStats[monthKey].showedRespondedMorning++;
      }

      monthStats[monthKey].totalCalls++;

      if (columnV.includes("rescheduled")) {
        monthStats[monthKey].rescheduled++;
      }

      if (columnV.includes("cancelled") || columnV.includes("canceled")) {
        monthStats[monthKey].canceled++;
      }

    } catch (e) {
      continue;
    }
  }

  const sortedMonths = Object.keys(monthStats).sort((a, b) => {
    return monthStats[a].monthStart - monthStats[b].monthStart;
  });

  targetSheet.clear();

  const headerRow1 = ["Month", "No Shows", "No Shows", "Showed Calls", "Showed Calls", "Total", "Total", "Total"];
  const headerRow2 = ["", "Responded to 1st Text", "Responded to Morning Text", "Responded to 1st Text", "Responded to Morning Text", "Total Calls for the Month", "Rescheduled", "Canceled"];

  targetSheet.appendRow(headerRow1);
  targetSheet.appendRow(headerRow2);

  targetSheet.getRange(1, 1, 2, 1).merge(); // Merge "Month" across 2 rows
  targetSheet.getRange(1, 2, 1, 2).merge(); // No Shows
  targetSheet.getRange(1, 4, 1, 2).merge(); // Showed Calls
  targetSheet.getRange(1, 6, 1, 3).merge(); // Total

  targetSheet.getRange(1, 1, 2, 8).setHorizontalAlignment("center").setVerticalAlignment("middle");
  targetSheet.getRange(1, 1, 2, 8).setFontWeight("bold");

  const columnWidths = [150, 150, 190, 150, 190, 170, 110, 100];
  columnWidths.forEach((width, i) => {
    targetSheet.setColumnWidth(i + 1, width);
  });

  let previousStats = null;

  sortedMonths.forEach((monthKey, index) => {
    const stats = monthStats[monthKey];
    const rowData = [
      monthKey,
      stats.noShowRespondedFirst,
      stats.noShowRespondedMorning,
      stats.showedRespondedFirst,
      stats.showedRespondedMorning,
      stats.totalCalls,
      stats.rescheduled,
      stats.canceled
    ];

    targetSheet.appendRow(rowData);

    if (index > 0 && previousStats) {
      const currentRow = index + 3;

      for (let col = 2; col <= 8; col++) {
        const currentValue = rowData[col - 1];
        const previousValue = col === 2 ? previousStats.noShowRespondedFirst :
          col === 3 ? previousStats.noShowRespondedMorning :
            col === 4 ? previousStats.showedRespondedFirst :
              col === 5 ? previousStats.showedRespondedMorning :
                col === 6 ? previousStats.totalCalls :
                  col === 7 ? previousStats.rescheduled :
                    previousStats.canceled;

        let color;
        if (currentValue > previousValue) {
          color = '#90EE90'; // Light green
        } else if (currentValue === previousValue) {
          color = '#FFFF99'; // Light yellow
        } else {
          color = '#FFB6C1'; // Light red
        }

        targetSheet.getRange(currentRow, col).setBackground(color);
      }
    } else if (index === 0) {
      const currentRow = 3;
      for (let col = 2; col <= 8; col++) {
        targetSheet.getRange(currentRow, col).setBackground('#FFFFFF');
      }
    }

    previousStats = stats;
  });

  if (sortedMonths.length > 0) {
    targetSheet.getRange(1, 1, sortedMonths.length + 2, 8).setBorder(true, true, true, true, true, true);
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Daily Dash Tools')
    .addItem('Update Daily Dash', 'updateDailyDash')
    .addItem('Update Weekly Dash', 'updateWeeklyDash')
    .addItem('Update Monthly Dash', 'updateMonthlyDash')
    .addToUi();
}