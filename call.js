function updateDailyDash() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Lead Entered");
  const targetSheet = ss.getSheetByName("Daily Dash");

  const dateColIndex = 0;
  const columnF_Index = 5;
  const columnK_Index = 10;
  const columnN_Index = 13;
  const columnQ_Index = 16;
  const columnT_Index = 19;
  const columnV_Index = 21;

  const data = sourceSheet.getDataRange().getValues();
  const dateStats = {};

  // Process data
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

      const isShowedStatus = columnV.includes("won-") || columnV.includes("lost-") ||
        columnV.includes("marker-") || columnV.includes("followup");

      if (columnF === "yes" && columnV === "noshow") dateStats[dateStr].noShowRespondedFirst++;
      if (columnQ === "yes" && columnV === "noshow") dateStats[dateStr].noShowRespondedMorning++;
      if (columnF === "yes" && isShowedStatus) dateStats[dateStr].showedRespondedFirst++;
      if (columnQ === "yes" && isShowedStatus) dateStats[dateStr].showedRespondedMorning++;

      const hasYesResponse = columnF === "yes" || columnK === "yes" || columnN === "yes" ||
        columnQ === "yes" || columnT === "yes";
      const isCountableStatus = columnV.includes("won-") || columnV.includes("lost-") ||
        columnV.includes("marker-") || columnV.includes("followup") ||
        columnV === "noshow";

      if (hasYesResponse && isCountableStatus) dateStats[dateStr].totalRespondedMessage++;

      dateStats[dateStr].totalCalls++;

      if (columnV.includes("rescheduled")) dateStats[dateStr].rescheduled++;
      if (columnV.includes("cancelled") || columnV.includes("canceled")) dateStats[dateStr].canceled++;

    } catch (e) {
      continue;
    }
  }

  const sortedDates = Object.keys(dateStats).sort((a, b) => new Date(a) - new Date(b));

  // Prepare all data at once
  const allRows = [];
  const backgroundColors = [];

  // Headers
  const headerRow1 = ["Date", "No Shows", "No Shows", "No Shows", "No Shows", "Showed Calls", "Showed Calls", "Showed Calls", "Showed Calls", "Total", "Total", "Total", "Total"];
  const headerRow2 = ["", "Responded to 1st Text", "Responded to 1st Text", "Responded to Morning Text", "Responded to Morning Text", "Responded to 1st Text", "Responded to 1st Text", "Responded to Morning Text", "Responded to Morning Text", "Total Responded Message", "Total Calls for the Day", "Rescheduled", "Canceled"];
  const headerRow3 = ["", "Total Number", "Percentage", "Total Number", "Percentage", "Total Number", "Percentage", "Total Number", "Percentage", "", "", "", ""];

  allRows.push(headerRow1, headerRow2, headerRow3);
  backgroundColors.push(
    Array(13).fill('#FFFFFF'),
    Array(13).fill('#FFFFFF'),
    Array(13).fill('#FFFFFF')
  );

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

    allRows.push(rowData);

    // Calculate colors for this row
    const rowColors = ['#FFFFFF']; // Date column

    if (index === 0) {
      rowColors.push(...Array(12).fill('#FFFFFF'));
    } else {
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
            color = '#90EE90';
          } else if (formattedCurrent === formattedPrevious) {
            color = '#FFFF99';
          } else {
            color = '#FFB6C1';
          }
        } else {
          if (currentValue > previousValue) {
            color = '#90EE90';
          } else if (currentValue === previousValue) {
            color = '#FFFF99';
          } else {
            color = '#FFB6C1';
          }
        }

        rowColors.push(color);
      });
    }

    backgroundColors.push(rowColors);

    previousStats = { ...stats, noShowFirstPct, noShowMorningPct, showedFirstPct, showedMorningPct };
  });

  targetSheet.clear();

  if (allRows.length > 0) {
    targetSheet.getRange(1, 1, allRows.length, 13).setValues(allRows);
    targetSheet.getRange(1, 1, allRows.length, 13).setBackgrounds(backgroundColors);
  }

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

  targetSheet.getRange(1, 1, 3, 13).setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontWeight("bold");

  const columnWidths = [120, 90, 90, 90, 90, 90, 90, 90, 90, 190, 160, 100, 100];
  columnWidths.forEach((width, i) => {
    targetSheet.setColumnWidth(i + 1, width);
  });

  if (allRows.length > 0) {
    targetSheet.getRange(1, 1, allRows.length, 13).setBorder(true, true, true, true, true, true);
  }
}

function updateWeeklyDash() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Lead Entered");
  const targetSheet = ss.getSheetByName("Weekly Dash");

  const dateColIndex = 0;
  const columnF_Index = 5;
  const columnK_Index = 10;
  const columnN_Index = 13;
  const columnQ_Index = 16;
  const columnT_Index = 19;
  const columnV_Index = 21;

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

      const isShowedStatus = columnV.includes("won-") || columnV.includes("lost-") ||
        columnV.includes("marker-") || columnV.includes("followup");

      if (columnF === "yes" && columnV === "noshow") weekStats[weekKey].noShowRespondedFirst++;
      if (columnQ === "yes" && columnV === "noshow") weekStats[weekKey].noShowRespondedMorning++;
      if (columnF === "yes" && isShowedStatus) weekStats[weekKey].showedRespondedFirst++;
      if (columnQ === "yes" && isShowedStatus) weekStats[weekKey].showedRespondedMorning++;

      const hasYesResponse = columnF === "yes" || columnK === "yes" || columnN === "yes" ||
        columnQ === "yes" || columnT === "yes";
      const isCountableStatus = columnV.includes("won-") || columnV.includes("lost-") ||
        columnV.includes("marker-") || columnV.includes("followup") ||
        columnV === "noshow";

      if (hasYesResponse && isCountableStatus) weekStats[weekKey].totalRespondedMessage++;

      weekStats[weekKey].totalCalls++;

      if (columnV.includes("rescheduled")) weekStats[weekKey].rescheduled++;
      if (columnV.includes("cancelled") || columnV.includes("canceled")) weekStats[weekKey].canceled++;

    } catch (e) {
      continue;
    }
  }

  const sortedWeeks = Object.keys(weekStats).sort((a, b) => weekStats[a].weekStart - weekStats[b].weekStart);

  const allRows = [];
  const backgroundColors = [];

  const headerRow1 = ["Week", "No Shows", "No Shows", "No Shows", "No Shows", "Showed Calls", "Showed Calls", "Showed Calls", "Showed Calls", "Total", "Total", "Total", "Total"];
  const headerRow2 = ["", "Responded to 1st Text", "Responded to 1st Text", "Responded to Morning Text", "Responded to Morning Text", "Responded to 1st Text", "Responded to 1st Text", "Responded to Morning Text", "Responded to Morning Text", "Total Responded Message", "Total Calls for the Week", "Rescheduled", "Canceled"];
  const headerRow3 = ["", "Total Number", "Percentage", "Total Number", "Percentage", "Total Number", "Percentage", "Total Number", "Percentage", "", "", "", ""];

  allRows.push(headerRow1, headerRow2, headerRow3);
  backgroundColors.push(
    Array(13).fill('#FFFFFF'),
    Array(13).fill('#FFFFFF'),
    Array(13).fill('#FFFFFF')
  );

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

    allRows.push(rowData);

    const rowColors = ['#FFFFFF'];

    if (index === 0) {
      rowColors.push(...Array(12).fill('#FFFFFF'));
    } else {
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
            color = '#90EE90';
          } else if (formattedCurrent === formattedPrevious) {
            color = '#FFFF99';
          } else {
            color = '#FFB6C1';
          }
        } else {
          if (currentValue > previousValue) {
            color = '#90EE90';
          } else if (currentValue === previousValue) {
            color = '#FFFF99';
          } else {
            color = '#FFB6C1';
          }
        }

        rowColors.push(color);
      });
    }

    backgroundColors.push(rowColors);

    previousStats = { ...stats, noShowFirstPct, noShowMorningPct, showedFirstPct, showedMorningPct };
  });

  targetSheet.clear();

  if (allRows.length > 0) {
    targetSheet.getRange(1, 1, allRows.length, 13).setValues(allRows);
    targetSheet.getRange(1, 1, allRows.length, 13).setBackgrounds(backgroundColors);
  }

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

  targetSheet.getRange(1, 1, 3, 13).setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontWeight("bold");

  const columnWidths = [160, 90, 90, 90, 90, 90, 90, 90, 90, 190, 160, 100, 100];
  columnWidths.forEach((width, i) => {
    targetSheet.setColumnWidth(i + 1, width);
  });

  if (allRows.length > 0) {
    targetSheet.getRange(1, 1, allRows.length, 13).setBorder(true, true, true, true, true, true);
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

      const isShowedStatus = columnV.includes("won-") || columnV.includes("lost-") ||
        columnV.includes("marker-") || columnV.includes("followup");

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
      const hasYesResponse = columnF === "yes" || columnK === "yes" || columnN === "yes" ||
        columnQ === "yes" || columnT === "yes";
      const isCountableStatus = columnV.includes("won-") || columnV.includes("lost-") ||
        columnV.includes("marker-") || columnV.includes("followup") ||
        columnV === "noshow";

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

  // Sort months chronologically
  const sortedMonths = Object.keys(monthStats).sort((a, b) => {
    return monthStats[a].monthStart - monthStats[b].monthStart;
  });

  // Prepare all rows and colors in memory
  const allRows = [];
  const backgroundColors = [];

  // Create header rows
  const headerRow1 = ["Month", "No Shows", "No Shows", "No Shows", "No Shows", "Showed Calls", "Showed Calls", "Showed Calls", "Showed Calls", "Total", "Total", "Total", "Total"];
  const headerRow2 = ["", "Responded to 1st Text", "Responded to 1st Text", "Responded to Morning Text", "Responded to Morning Text", "Responded to 1st Text", "Responded to 1st Text", "Responded to Morning Text", "Responded to Morning Text", "Total Responded Message", "Total Calls for the Month", "Rescheduled", "Canceled"];
  const headerRow3 = ["", "Total Number", "Percentage", "Total Number", "Percentage", "Total Number", "Percentage", "Total Number", "Percentage", "", "", "", ""];

  allRows.push(headerRow1, headerRow2, headerRow3);
  backgroundColors.push(
    Array(13).fill('#FFFFFF'),
    Array(13).fill('#FFFFFF'),
    Array(13).fill('#FFFFFF')
  );

  let previousStats = null;

  // Process each month
  sortedMonths.forEach((monthKey, index) => {
    const stats = monthStats[monthKey];

    // Calculate percentages
    const noShowFirstPct = stats.totalCalls > 0 ? (stats.noShowRespondedFirst / stats.totalCalls * 100).toFixed(1) + "%" : "0%";
    const noShowMorningPct = stats.totalCalls > 0 ? (stats.noShowRespondedMorning / stats.totalCalls * 100).toFixed(1) + "%" : "0%";
    const showedFirstPct = stats.totalCalls > 0 ? (stats.showedRespondedFirst / stats.totalCalls * 100).toFixed(1) + "%" : "0%";
    const showedMorningPct = stats.totalCalls > 0 ? (stats.showedRespondedMorning / stats.totalCalls * 100).toFixed(1) + "%" : "0%";

    // Build row data
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

    allRows.push(rowData);

    // Calculate colors for this row
    const rowColors = ['#FFFFFF']; // Month column always white

    if (index === 0) {
      // First row - all white
      rowColors.push(...Array(12).fill('#FFFFFF'));
    } else {
      // Compare with previous month
      const columnsToColor = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13];

      columnsToColor.forEach(col => {
        // Get current value
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

        // Get previous value
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

        // Compare values and assign color
        if (String(currentValue).includes("%") || String(previousValue).includes("%")) {
          const formattedCurrent = parseFloat(String(currentValue).replace("%", ""));
          const formattedPrevious = parseFloat(String(previousValue).replace("%", ""));

          if (formattedCurrent > formattedPrevious) {
            color = '#90EE90'; // Light green - improvement
          } else if (formattedCurrent === formattedPrevious) {
            color = '#FFFF99'; // Light yellow - no change
          } else {
            color = '#FFB6C1'; // Light red - decline
          }
        } else {
          if (currentValue > previousValue) {
            color = '#90EE90'; // Light green - improvement
          } else if (currentValue === previousValue) {
            color = '#FFFF99'; // Light yellow - no change
          } else {
            color = '#FFB6C1'; // Light red - decline
          }
        }

        rowColors.push(color);
      });
    }

    backgroundColors.push(rowColors);

    // Store current stats for next iteration
    previousStats = {
      ...stats,
      noShowFirstPct,
      noShowMorningPct,
      showedFirstPct,
      showedMorningPct
    };
  });

  // Clear sheet and write all data at once
  targetSheet.clear();

  if (allRows.length > 0) {
    targetSheet.getRange(1, 1, allRows.length, 13).setValues(allRows);
    targetSheet.getRange(1, 1, allRows.length, 13).setBackgrounds(backgroundColors);
  }

  // Apply header merging
  targetSheet.getRange(1, 1, 3, 1).merge();  // Merge "Month" across 3 rows
  targetSheet.getRange(1, 2, 1, 4).merge();  // No Shows (spans 4 columns)
  targetSheet.getRange(1, 6, 1, 4).merge();  // Showed Calls (spans 4 columns)
  targetSheet.getRange(1, 10, 1, 4).merge(); // Total (spans 4 columns)

  targetSheet.getRange(2, 2, 1, 2).merge();  // No Shows - Responded to 1st Text
  targetSheet.getRange(2, 4, 1, 2).merge();  // No Shows - Responded to Morning Text
  targetSheet.getRange(2, 6, 1, 2).merge();  // Showed Calls - Responded to 1st Text
  targetSheet.getRange(2, 8, 1, 2).merge();  // Showed Calls - Responded to Morning Text
  targetSheet.getRange(2, 10, 2, 1).merge(); // Total Responded Message
  targetSheet.getRange(2, 11, 2, 1).merge(); // Total Calls for the Month
  targetSheet.getRange(2, 12, 2, 1).merge(); // Rescheduled
  targetSheet.getRange(2, 13, 2, 1).merge(); // Canceled

  // Apply header formatting (chained for efficiency)
  targetSheet.getRange(1, 1, 3, 13)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontWeight("bold");

  // Set column widths
  const columnWidths = [120, 90, 90, 90, 90, 90, 90, 90, 90, 190, 160, 100, 100];
  columnWidths.forEach((width, i) => {
    targetSheet.setColumnWidth(i + 1, width);
  });

  // Apply borders to all data
  if (allRows.length > 0) {
    targetSheet.getRange(1, 1, allRows.length, 13).setBorder(true, true, true, true, true, true);
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