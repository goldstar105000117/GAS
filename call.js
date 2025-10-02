function updateDailyDash() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Lead Entered");
  const targetSheet = ss.getSheetByName("Daily Dash");

  const dateColIndex = 0;  // Column A - Date Of Meeting
  const columnF_Index = 5;  // Column F - Responded to 1st Text
  const columnK_Index = 10; // Column K
  const columnN_Index = 13; // Column N
  const columnQ_Index = 16; // Column Q - Responded to Morning Text
  const columnT_Index = 19; // Column T
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
          totalRespondedMessage: 0,
          totalCalls: 0,
          rescheduled: 0,
          canceled: 0
        };
      }

      const columnF = String(data[i][columnF_Index]).toLowerCase().trim();
      const columnK = String(data[i][columnK_Index]).toLowerCase().trim();
      const columnN = String(data[i][columnN_Index]).toLowerCase().trim();
      const columnQ = String(data[i][columnQ_Index]).toLowerCase().trim();
      const columnT = String(data[i][columnT_Index]).toLowerCase().trim();
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

      // Total Responded Message logic
      const hasYesResponse = columnF === "yes" || columnK === "yes" || columnN === "yes" || columnQ === "yes" || columnT === "yes";
      const isCountableStatus = columnV.includes("won-") || columnV.includes("lost-") || columnV.includes("marker-") || columnV.includes("followup") || columnV === "noshow";
      
      if (hasYesResponse && isCountableStatus) {
        dateStats[dateStr].totalRespondedMessage++;
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

  const headerRow1 = ["Date", "No Shows", "No Shows", "No Shows", "No Shows", "Showed Calls", "Showed Calls", "Showed Calls", "Showed Calls", "Total", "Total", "Total", "Total"];
  const headerRow2 = ["", "Responded to 1st Text", "Responded to 1st Text", "Responded to Morning Text", "Responded to Morning Text", "Responded to 1st Text", "Responded to 1st Text", "Responded to Morning Text", "Responded to Morning Text", "Total Responded Message", "Total Calls for the Day", "Rescheduled", "Canceled"];
  const headerRow3 = ["", "Total Number", "Percentage", "Total Number", "Percentage", "Total Number", "Percentage", "Total Number", "Percentage", "", "", "", ""];

  targetSheet.appendRow(headerRow1);
  targetSheet.appendRow(headerRow2);
  targetSheet.appendRow(headerRow3);

  targetSheet.getRange(1, 1, 3, 1).merge(); // Merge "Date" across 3 rows
  targetSheet.getRange(1, 2, 1, 4).merge(); // No Shows (spans 4 columns)
  targetSheet.getRange(1, 6, 1, 4).merge(); // Showed Calls (spans 4 columns)
  targetSheet.getRange(1, 10, 1, 4).merge(); // Total (spans 4 columns)

  targetSheet.getRange(2, 2, 1, 2).merge(); // No Shows - Responded to 1st Text
  targetSheet.getRange(2, 4, 1, 2).merge(); // No Shows - Responded to Morning Text
  targetSheet.getRange(2, 6, 1, 2).merge(); // Showed Calls - Responded to 1st Text
  targetSheet.getRange(2, 8, 1, 2).merge(); // Showed Calls - Responded to Morning Text
  targetSheet.getRange(2, 10, 2, 1).merge(); // Total Responded Message
  targetSheet.getRange(2, 11, 2, 1).merge(); // Total Calls for the Day
  targetSheet.getRange(2, 12, 2, 1).merge(); // Rescheduled
  targetSheet.getRange(2, 13, 2, 1).merge(); // Canceled

  targetSheet.getRange(1, 1, 3, 13).setHorizontalAlignment("center").setVerticalAlignment("middle");
  targetSheet.getRange(1, 1, 3, 13).setFontWeight("bold");

  const columnWidths = [100, 90, 90, 90, 90, 90, 90, 90, 90, 190, 160, 100, 100];
  columnWidths.forEach((width, i) => {
    targetSheet.setColumnWidth(i + 1, width);
  });

  let previousStats = null;

  sortedDates.forEach((date, index) => {
    const stats = dateStats[date];

    const noShowFirstPct = stats.totalCalls > 0 ? (stats.noShowRespondedFirst / stats.totalCalls * 100).toFixed(1) + "%" : "0%";
    const noShowMorningPct = stats.totalCalls > 0 ? (stats.noShowRespondedMorning / stats.totalCalls * 100).toFixed(1) + "%" : "0%";
    const showedFirstPct = stats.totalCalls > 0 ? (stats.showedRespondedFirst / stats.totalCalls * 100).toFixed(1) + "%" : "0%";
    const showedMorningPct = stats.totalCalls > 0 ? (stats.showedRespondedMorning / stats.totalCalls * 100).toFixed(1) + "%" : "0%";

    const rowData = [
      date,
      stats.noShowRespondedFirst,
      noShowFirstPct,
      stats.noShowRespondedMorning,
      noShowMorningPct,
      stats.showedRespondedFirst,
      showedFirstPct,
      stats.showedRespondedMorning,
      showedMorningPct,
      stats.totalRespondedMessage,
      stats.totalCalls,
      stats.rescheduled,
      stats.canceled
    ];

    targetSheet.appendRow(rowData);

    if (index > 0 && previousStats) {
      const currentRow = index + 4;

      const columnsToColor = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13];

      columnsToColor.forEach(col => {
        const currentValue = col === 2 ? stats.noShowRespondedFirst :
          col === 3 ? noShowFirstPct :
            col === 4 ? stats.noShowRespondedMorning :
              col === 5 ? noShowMorningPct :
                col === 6 ? stats.showedRespondedFirst :
                  col === 7 ? showedFirstPct :
                    col === 8 ? stats.showedRespondedMorning :
                      col === 9 ? showedMorningPct :
                        col === 10 ? stats.totalRespondedMessage :
                          col === 11 ? stats.totalCalls :
                            col === 12 ? stats.rescheduled :
                              stats.canceled;

        const previousValue = col === 2 ? previousStats.noShowRespondedFirst :
          col === 3 ? previousStats.noShowFirstPct :
            col === 4 ? previousStats.noShowRespondedMorning :
              col === 5 ? previousStats.noShowMorningPct :
                col === 6 ? previousStats.showedRespondedFirst :
                  col === 7 ? previousStats.showedFirstPct :
                    col === 8 ? previousStats.showedRespondedMorning :
                      col === 9 ? previousStats.showedMorningPct :
                        col === 10 ? previousStats.totalRespondedMessage :
                          col === 11 ? previousStats.totalCalls :
                            col === 12 ? previousStats.rescheduled :
                              previousStats.canceled;

        let color;

        if (String(currentValue).includes("%") || String(previousValue).includes("%")) {
          const formattedCurrent = parseFloat(String(currentValue).replace("%", ""));
          const formattedPrevious = parseFloat(String(previousValue).replace("%", ""));

          if (formattedCurrent > formattedPrevious) {
            color = '#90EE90'; // Light green
          } else if (formattedCurrent === formattedPrevious) {
            color = '#FFFF99'; // Light yellow
          } else {
            color = '#FFB6C1'; // Light red
          }
        } else {
          if (currentValue > previousValue) {
            color = '#90EE90'; // Light green
          } else if (currentValue === previousValue) {
            color = '#FFFF99'; // Light yellow
          } else {
            color = '#FFB6C1'; // Light red
          }
        }

        targetSheet.getRange(currentRow, col).setBackground(color);
      });
    } else if (index === 0) {
      const currentRow = 4;
      for (let col = 2; col <= 13; col++) {
        targetSheet.getRange(currentRow, col).setBackground('#FFFFFF');
      }
    }

    previousStats = { ...stats, noShowFirstPct, noShowMorningPct, showedFirstPct, showedMorningPct };
  });

  if (sortedDates.length > 0) {
    targetSheet.getRange(1, 1, sortedDates.length + 3, 13).setBorder(true, true, true, true, true, true);
  }
}

function updateWeeklyDash() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Lead Entered");
  const targetSheet = ss.getSheetByName("Weekly Dash");

  const dateColIndex = 0;  // Column A - Date Of Meeting
  const columnF_Index = 5;  // Column F - Responded to 1st Text
  const columnK_Index = 10; // Column K
  const columnN_Index = 13; // Column N
  const columnQ_Index = 16; // Column Q - Responded to Morning Text
  const columnT_Index = 19; // Column T
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
          totalRespondedMessage: 0,
          totalCalls: 0,
          rescheduled: 0,
          canceled: 0
        };
      }

      const columnF = String(data[i][columnF_Index]).toLowerCase().trim();
      const columnK = String(data[i][columnK_Index]).toLowerCase().trim();
      const columnN = String(data[i][columnN_Index]).toLowerCase().trim();
      const columnQ = String(data[i][columnQ_Index]).toLowerCase().trim();
      const columnT = String(data[i][columnT_Index]).toLowerCase().trim();
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

      // Total Responded Message logic
      const hasYesResponse = columnF === "yes" || columnK === "yes" || columnN === "yes" || columnQ === "yes" || columnT === "yes";
      const isCountableStatus = columnV.includes("won-") || columnV.includes("lost-") || columnV.includes("marker-") || columnV.includes("followup") || columnV === "noshow";
      
      if (hasYesResponse && isCountableStatus) {
        weekStats[weekKey].totalRespondedMessage++;
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

  const headerRow1 = ["Week", "No Shows", "No Shows", "No Shows", "No Shows", "Showed Calls", "Showed Calls", "Showed Calls", "Showed Calls", "Total", "Total", "Total", "Total"];
  const headerRow2 = ["", "Responded to 1st Text", "Responded to 1st Text", "Responded to Morning Text", "Responded to Morning Text", "Responded to 1st Text", "Responded to 1st Text", "Responded to Morning Text", "Responded to Morning Text", "Total Responded Message", "Total Calls for the Week", "Rescheduled", "Canceled"];
  const headerRow3 = ["", "Total Number", "Percentage", "Total Number", "Percentage", "Total Number", "Percentage", "Total Number", "Percentage", "", "", "", ""];

  targetSheet.appendRow(headerRow1);
  targetSheet.appendRow(headerRow2);
  targetSheet.appendRow(headerRow3);

  targetSheet.getRange(1, 1, 3, 1).merge();
  targetSheet.getRange(1, 2, 1, 4).merge();
  targetSheet.getRange(1, 6, 1, 4).merge();
  targetSheet.getRange(1, 10, 1, 4).merge();

  targetSheet.getRange(2, 2, 1, 2).merge();
  targetSheet.getRange(2, 4, 1, 2).merge();
  targetSheet.getRange(2, 6, 1, 2).merge();
  targetSheet.getRange(2, 8, 1, 2).merge();
  targetSheet.getRange(2, 10, 2, 1).merge();
  targetSheet.getRange(2, 11, 2, 1).merge();
  targetSheet.getRange(2, 12, 2, 1).merge();
  targetSheet.getRange(2, 13, 2, 1).merge();

  targetSheet.getRange(1, 1, 3, 13).setHorizontalAlignment("center").setVerticalAlignment("middle");
  targetSheet.getRange(1, 1, 3, 13).setFontWeight("bold");

  const columnWidths = [160, 90, 90, 90, 90, 90, 90, 90, 90, 190, 160, 100, 100];
  columnWidths.forEach((width, i) => {
    targetSheet.setColumnWidth(i + 1, width);
  });

  let previousStats = null;

  sortedWeeks.forEach((weekKey, index) => {
    const stats = weekStats[weekKey];

    const noShowFirstPct = stats.totalCalls > 0 ? (stats.noShowRespondedFirst / stats.totalCalls * 100).toFixed(1) + "%" : "0%";
    const noShowMorningPct = stats.totalCalls > 0 ? (stats.noShowRespondedMorning / stats.totalCalls * 100).toFixed(1) + "%" : "0%";
    const showedFirstPct = stats.totalCalls > 0 ? (stats.showedRespondedFirst / stats.totalCalls * 100).toFixed(1) + "%" : "0%";
    const showedMorningPct = stats.totalCalls > 0 ? (stats.showedRespondedMorning / stats.totalCalls * 100).toFixed(1) + "%" : "0%";

    const rowData = [
      weekKey,
      stats.noShowRespondedFirst,
      noShowFirstPct,
      stats.noShowRespondedMorning,
      noShowMorningPct,
      stats.showedRespondedFirst,
      showedFirstPct,
      stats.showedRespondedMorning,
      showedMorningPct,
      stats.totalRespondedMessage,
      stats.totalCalls,
      stats.rescheduled,
      stats.canceled
    ];

    targetSheet.appendRow(rowData);

    if (index > 0 && previousStats) {
      const currentRow = index + 4;
      const columnsToColor = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13];

      columnsToColor.forEach(col => {
        const currentValue = col === 2 ? stats.noShowRespondedFirst :
          col === 3 ? noShowFirstPct :
            col === 4 ? stats.noShowRespondedMorning :
              col === 5 ? noShowMorningPct :
                col === 6 ? stats.showedRespondedFirst :
                  col === 7 ? showedFirstPct :
                    col === 8 ? stats.showedRespondedMorning :
                      col === 9 ? showedMorningPct :
                        col === 10 ? stats.totalRespondedMessage :
                          col === 11 ? stats.totalCalls :
                            col === 12 ? stats.rescheduled :
                              stats.canceled;

        const previousValue = col === 2 ? previousStats.noShowRespondedFirst :
          col === 3 ? previousStats.noShowFirstPct :
            col === 4 ? previousStats.noShowRespondedMorning :
              col === 5 ? previousStats.noShowMorningPct :
                col === 6 ? previousStats.showedRespondedFirst :
                  col === 7 ? previousStats.showedFirstPct :
                    col === 8 ? previousStats.showedRespondedMorning :
                      col === 9 ? previousStats.showedMorningPct :
                        col === 10 ? previousStats.totalRespondedMessage :
                          col === 11 ? previousStats.totalCalls :
                            col === 12 ? previousStats.rescheduled :
                              previousStats.canceled;

        let color;

        if (String(currentValue).includes("%") || String(previousValue).includes("%")) {
          const formattedCurrent = parseFloat(String(currentValue).replace("%", ""));
          const formattedPrevious = parseFloat(String(previousValue).replace("%", ""));

          if (formattedCurrent > formattedPrevious) {
            color = '#90EE90'; // Light green
          } else if (formattedCurrent === formattedPrevious) {
            color = '#FFFF99'; // Light yellow
          } else {
            color = '#FFB6C1'; // Light red
          }
        } else {
          if (currentValue > previousValue) {
            color = '#90EE90'; // Light green
          } else if (currentValue === previousValue) {
            color = '#FFFF99'; // Light yellow
          } else {
            color = '#FFB6C1'; // Light red
          }
        }

        targetSheet.getRange(currentRow, col).setBackground(color);
      });
    } else if (index === 0) {
      const currentRow = 4;
      for (let col = 2; col <= 13; col++) {
        targetSheet.getRange(currentRow, col).setBackground('#FFFFFF');
      }
    }

    previousStats = { ...stats, noShowFirstPct, noShowMorningPct, showedFirstPct, showedMorningPct };
  });

  if (sortedWeeks.length > 0) {
    targetSheet.getRange(1, 1, sortedWeeks.length + 3, 13).setBorder(true, true, true, true, true, true);
  }
}

function updateMonthlyDash() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Lead Entered");
  const targetSheet = ss.getSheetByName("Monthly Dash");

  const dateColIndex = 0;  // Column A - Date Of Meeting
  const columnF_Index = 5;  // Column F - Responded to 1st Text
  const columnK_Index = 10; // Column K
  const columnN_Index = 13; // Column N
  const columnQ_Index = 16; // Column Q - Responded to Morning Text
  const columnT_Index = 19; // Column T
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
          totalRespondedMessage: 0,
          totalCalls: 0,
          rescheduled: 0,
          canceled: 0
        };
      }

      const columnF = String(data[i][columnF_Index]).toLowerCase().trim();
      const columnK = String(data[i][columnK_Index]).toLowerCase().trim();
      const columnN = String(data[i][columnN_Index]).toLowerCase().trim();
      const columnQ = String(data[i][columnQ_Index]).toLowerCase().trim();
      const columnT = String(data[i][columnT_Index]).toLowerCase().trim();
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

      // Total Responded Message logic
      const hasYesResponse = columnF === "yes" || columnK === "yes" || columnN === "yes" || columnQ === "yes" || columnT === "yes";
      const isCountableStatus = columnV.includes("won-") || columnV.includes("lost-") || columnV.includes("marker-") || columnV.includes("followup") || columnV === "noshow";
      
      if (hasYesResponse && isCountableStatus) {
        monthStats[monthKey].totalRespondedMessage++;
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

  const headerRow1 = ["Month", "No Shows", "No Shows", "No Shows", "No Shows", "Showed Calls", "Showed Calls", "Showed Calls", "Showed Calls", "Total", "Total", "Total", "Total"];
  const headerRow2 = ["", "Responded to 1st Text", "Responded to 1st Text", "Responded to Morning Text", "Responded to Morning Text", "Responded to 1st Text", "Responded to 1st Text", "Responded to Morning Text", "Responded to Morning Text", "Total Responded Message", "Total Calls for the Month", "Rescheduled", "Canceled"];
  const headerRow3 = ["", "Total Number", "Percentage", "Total Number", "Percentage", "Total Number", "Percentage", "Total Number", "Percentage", "", "", "", ""];

  targetSheet.appendRow(headerRow1);
  targetSheet.appendRow(headerRow2);
  targetSheet.appendRow(headerRow3);

  targetSheet.getRange(1, 1, 3, 1).merge();
  targetSheet.getRange(1, 2, 1, 4).merge();
  targetSheet.getRange(1, 6, 1, 4).merge();
  targetSheet.getRange(1, 10, 1, 4).merge();

  targetSheet.getRange(2, 2, 1, 2).merge();
  targetSheet.getRange(2, 4, 1, 2).merge();
  targetSheet.getRange(2, 6, 1, 2).merge();
  targetSheet.getRange(2, 8, 1, 2).merge();
  targetSheet.getRange(2, 10, 2, 1).merge();
  targetSheet.getRange(2, 11, 2, 1).merge();
  targetSheet.getRange(2, 12, 2, 1).merge();
  targetSheet.getRange(2, 13, 2, 1).merge();

  targetSheet.getRange(1, 1, 3, 13).setHorizontalAlignment("center").setVerticalAlignment("middle");
  targetSheet.getRange(1, 1, 3, 13).setFontWeight("bold");

  const columnWidths = [120, 90, 90, 90, 90, 90, 90, 90, 90, 190, 160, 100, 100];
  columnWidths.forEach((width, i) => {
    targetSheet.setColumnWidth(i + 1, width);
  });

  let previousStats = null;

  sortedMonths.forEach((monthKey, index) => {
    const stats = monthStats[monthKey];

    const noShowFirstPct = stats.totalCalls > 0 ? (stats.noShowRespondedFirst / stats.totalCalls * 100).toFixed(1) + "%" : "0%";
    const noShowMorningPct = stats.totalCalls > 0 ? (stats.noShowRespondedMorning / stats.totalCalls * 100).toFixed(1) + "%" : "0%";
    const showedFirstPct = stats.totalCalls > 0 ? (stats.showedRespondedFirst / stats.totalCalls * 100).toFixed(1) + "%" : "0%";
    const showedMorningPct = stats.totalCalls > 0 ? (stats.showedRespondedMorning / stats.totalCalls * 100).toFixed(1) + "%" : "0%";

    const rowData = [
      monthKey,
      stats.noShowRespondedFirst,
      noShowFirstPct,
      stats.noShowRespondedMorning,
      noShowMorningPct,
      stats.showedRespondedFirst,
      showedFirstPct,
      stats.showedRespondedMorning,
      showedMorningPct,
      stats.totalRespondedMessage,
      stats.totalCalls,
      stats.rescheduled,
      stats.canceled
    ];

    targetSheet.appendRow(rowData);

    if (index > 0 && previousStats) {
      const currentRow = index + 4;
      const columnsToColor = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13];

      columnsToColor.forEach(col => {
        const currentValue = col === 2 ? stats.noShowRespondedFirst :
          col === 3 ? noShowFirstPct :
            col === 4 ? stats.noShowRespondedMorning :
              col === 5 ? noShowMorningPct :
                col === 6 ? stats.showedRespondedFirst :
                  col === 7 ? showedFirstPct :
                    col === 8 ? stats.showedRespondedMorning :
                      col === 9 ? showedMorningPct :
                        col === 10 ? stats.totalRespondedMessage :
                          col === 11 ? stats.totalCalls :
                            col === 12 ? stats.rescheduled :
                              stats.canceled;

        const previousValue = col === 2 ? previousStats.noShowRespondedFirst :
          col === 3 ? previousStats.noShowFirstPct :
            col === 4 ? previousStats.noShowRespondedMorning :
              col === 5 ? previousStats.noShowMorningPct :
                col === 6 ? previousStats.showedRespondedFirst :
                  col === 7 ? previousStats.showedFirstPct :
                    col === 8 ? previousStats.showedRespondedMorning :
                      col === 9 ? previousStats.showedMorningPct :
                        col === 10 ? previousStats.totalRespondedMessage :
                          col === 11 ? previousStats.totalCalls :
                            col === 12 ? previousStats.rescheduled :
                              previousStats.canceled;

        let color;

        if (String(currentValue).includes("%") || String(previousValue).includes("%")) {
          const formattedCurrent = parseFloat(String(currentValue).replace("%", ""));
          const formattedPrevious = parseFloat(String(previousValue).replace("%", ""));

          if (formattedCurrent > formattedPrevious) {
            color = '#90EE90'; // Light green
          } else if (formattedCurrent === formattedPrevious) {
            color = '#FFFF99'; // Light yellow
          } else {
            color = '#FFB6C1'; // Light red
          }
        } else {
          if (currentValue > previousValue) {
            color = '#90EE90'; // Light green
          } else if (currentValue === previousValue) {
            color = '#FFFF99'; // Light yellow
          } else {
            color = '#FFB6C1'; // Light red
          }
        }

        targetSheet.getRange(currentRow, col).setBackground(color);
      });
    } else if (index === 0) {
      const currentRow = 4;
      for (let col = 2; col <= 13; col++) {
        targetSheet.getRange(currentRow, col).setBackground('#FFFFFF');
      }
    }

    previousStats = { ...stats, noShowFirstPct, noShowMorningPct, showedFirstPct, showedMorningPct };
  });

  if (sortedMonths.length > 0) {
    targetSheet.getRange(1, 1, sortedMonths.length + 3, 13).setBorder(true, true, true, true, true, true);
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