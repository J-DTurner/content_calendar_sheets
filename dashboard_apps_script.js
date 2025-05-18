/**
 * Dashboard Automation Script for Social Media Content Calendar
 * 
 * This script provides functions for dashboard interactivity, data updates,
 * filtering, and visualization enhancements.
 */

// Dashboard configuration
const DASHBOARD_CONFIG = {
  SHEET_NAME: 'Dashboard',
  DATE_RANGE_FILTER_CELL: 'B43',
  CHANNEL_FILTER_CELL: 'D43',
  PILLAR_FILTER_CELL: 'F43',
  TEAM_FILTER_CELL: 'H43',
  CURRENT_WEEK_CELL: 'B24',
  WEEK_START_CELL: 'B25',
  WEEK_END_CELL: 'B26',
  REFRESH_TIMESTAMP_CELL: 'I2'
};

/**
 * Updates the dashboard when it's opened
 * This is called by the onOpen trigger
 */
function updateDashboardOnOpen() {
  try {
    // Get the dashboard sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboard = ss.getSheetByName(DASHBOARD_CONFIG.SHEET_NAME);
    
    if (!dashboard) return;
    
    // Update the current week
    const currentWeek = getCurrentWeekNumber();
    dashboard.getRange(DASHBOARD_CONFIG.CURRENT_WEEK_CELL).setValue(currentWeek);
    
    // Update the week start and end dates
    updateWeekDates(dashboard);
    
    // Update refresh timestamp
    dashboard.getRange(DASHBOARD_CONFIG.REFRESH_TIMESTAMP_CELL)
      .setValue('Last updated: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
    
    // Refresh pivot data if it exists
    refreshPivotData();
    
  } catch (e) {
    console.error('Error updating dashboard on open:', e);
  }
}

/**
 * Gets the current ISO week number
 */
function getCurrentWeekNumber() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'w');
}

/**
 * Updates the week start and end dates based on the current week number
 */
function updateWeekDates(dashboard) {
  const weekNumber = dashboard.getRange(DASHBOARD_CONFIG.CURRENT_WEEK_CELL).getValue();
  const year = new Date().getFullYear();
  
  // Calculate the start date of the week
  // This uses the ISO week definition (week 1 is the week with the first Thursday of the year)
  const startDate = getWeekStartDate(weekNumber, year);
  dashboard.getRange(DASHBOARD_CONFIG.WEEK_START_CELL).setValue(startDate);
  
  // Calculate the end date (start date + 6 days)
  const endDate = new Date(startDate);
  endDate.setDate(startDate.getDate() + 6);
  dashboard.getRange(DASHBOARD_CONFIG.WEEK_END_CELL).setValue(endDate);
}

/**
 * Gets the start date of a specified week number
 */
function getWeekStartDate(weekNumber, year) {
  // January 4th is always in week 1 in ISO weeks
  const jan4 = new Date(year, 0, 4);
  // Find the Monday of week 1
  const week1Start = new Date(jan4);
  week1Start.setDate(jan4.getDate() - jan4.getDay() + 1);
  if (jan4.getDay() === 0) week1Start.setDate(week1Start.getDate() - 7);
  
  // Add the required number of weeks
  const weekStart = new Date(week1Start);
  weekStart.setDate(week1Start.getDate() + (weekNumber - 1) * 7);
  
  return weekStart;
}

/**
 * Refreshes the dashboard data manually
 * This is called by the refresh button
 */
function refreshDashboard() {
  // Get the dashboard sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(DASHBOARD_CONFIG.SHEET_NAME);
  
  if (!dashboard) {
    SpreadsheetApp.getUi().alert('Dashboard sheet not found.');
    return;
  }
  
  // Update the refresh timestamp
  dashboard.getRange(DASHBOARD_CONFIG.REFRESH_TIMESTAMP_CELL)
    .setValue('Last updated: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
  
  // Refresh pivot data if it exists
  refreshPivotData();
  
  // Force recalculation of dashboard formulas
  dashboard.getRange('A1').setValue(dashboard.getRange('A1').getValue());
  
  // Show confirmation
  SpreadsheetApp.getUi().alert('Dashboard refreshed!');
}

/**
 * Refreshes the pivot data source for dashboard charts
 */
function refreshPivotData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const analyticsSheet = ss.getSheetByName('Analytics');
  
  // Skip if Analytics sheet doesn't exist
  if (!analyticsSheet) return;
  
  // Check for pivot tables in the Analytics sheet
  try {
    // Get all pivot tables in the sheet
    const pivotTables = analyticsSheet.getPivotTables();
    
    // Refresh each pivot table
    for (let i = 0; i < pivotTables.length; i++) {
      pivotTables[i].refresh();
    }
  } catch (e) {
    // Older versions of Google Sheets may not support this method
    // In that case, we'll use a workaround to force refresh
    
    // Toggle a cell value to force recalculation
    const toggleCell = analyticsSheet.getRange('A1');
    const currentValue = toggleCell.getValue();
    toggleCell.setValue('Refreshing...');
    SpreadsheetApp.flush();
    toggleCell.setValue(currentValue);
  }
}

/**
 * Resets all dashboard filters to default values
 */
function resetDashboardFilters() {
  // Get the dashboard sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(DASHBOARD_CONFIG.SHEET_NAME);
  
  if (!dashboard) return;
  
  // Reset all filters to "All"
  dashboard.getRange(DASHBOARD_CONFIG.DATE_RANGE_FILTER_CELL).setValue('All Time');
  dashboard.getRange(DASHBOARD_CONFIG.CHANNEL_FILTER_CELL).setValue('All');
  dashboard.getRange(DASHBOARD_CONFIG.PILLAR_FILTER_CELL).setValue('All');
  dashboard.getRange(DASHBOARD_CONFIG.TEAM_FILTER_CELL).setValue('All');
  
  // Refresh data with cleared filters
  refreshDashboard();
  
  // Show confirmation
  SpreadsheetApp.getUi().alert('Filters have been reset.');
}

/**
 * Navigates to the previous week
 */
function previousWeek() {
  // Get the dashboard sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(DASHBOARD_CONFIG.SHEET_NAME);
  
  if (!dashboard) return;
  
  // Get current week and decrement
  const currentWeek = dashboard.getRange(DASHBOARD_CONFIG.CURRENT_WEEK_CELL).getValue();
  let newWeek = parseInt(currentWeek) - 1;
  
  // Handle year transition
  if (newWeek < 1) {
    // Go to last week of previous year
    const prevYear = new Date().getFullYear() - 1;
    const lastWeekOfYear = getISOWeeksInYear(prevYear);
    newWeek = lastWeekOfYear;
  }
  
  // Update week number
  dashboard.getRange(DASHBOARD_CONFIG.CURRENT_WEEK_CELL).setValue(newWeek);
  
  // Update dates
  updateWeekDates(dashboard);
  
  // Update weekly data
  updateWeeklyData(dashboard, newWeek);
}

/**
 * Navigates to the next week
 */
function nextWeek() {
  // Get the dashboard sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(DASHBOARD_CONFIG.SHEET_NAME);
  
  if (!dashboard) return;
  
  // Get current week and increment
  const currentWeek = dashboard.getRange(DASHBOARD_CONFIG.CURRENT_WEEK_CELL).getValue();
  let newWeek = parseInt(currentWeek) + 1;
  
  // Handle year transition
  const weeksInYear = getISOWeeksInYear(new Date().getFullYear());
  if (newWeek > weeksInYear) {
    // Go to first week of next year
    newWeek = 1;
  }
  
  // Update week number
  dashboard.getRange(DASHBOARD_CONFIG.CURRENT_WEEK_CELL).setValue(newWeek);
  
  // Update dates
  updateWeekDates(dashboard);
  
  // Update weekly data
  updateWeeklyData(dashboard, newWeek);
}

/**
 * Navigates to the current week
 */
function currentWeek() {
  // Get the dashboard sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(DASHBOARD_CONFIG.SHEET_NAME);
  
  if (!dashboard) return;
  
  // Get current week number
  const currentWeek = getCurrentWeekNumber();
  
  // Update week number
  dashboard.getRange(DASHBOARD_CONFIG.CURRENT_WEEK_CELL).setValue(currentWeek);
  
  // Update dates
  updateWeekDates(dashboard);
  
  // Update weekly data
  updateWeeklyData(dashboard, currentWeek);
}

/**
 * Gets the number of ISO weeks in a year
 */
function getISOWeeksInYear(year) {
  // A year has 53 weeks if it starts on a Thursday or is a leap year that starts on a Wednesday
  const jan1 = new Date(year, 0, 1);
  const dec31 = new Date(year, 11, 31);
  
  // Check if Jan 1 is a Thursday or Dec 31 is a Thursday
  return (jan1.getDay() === 4 || dec31.getDay() === 4) ? 53 : 52;
}

/**
 * Updates the weekly data display in the dashboard
 */
function updateWeeklyData(dashboard, weekNumber) {
  // Implementation depends on the specific design of the weekly view
  // This is a placeholder for the function
  
  // Example: Update a named range that's used in chart data
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get the content calendar sheet
    const calendarSheet = ss.getSheetByName('Content Calendar');
    if (!calendarSheet) return;
    
    // Get data for the selected week
    const weekData = getWeekData(calendarSheet, weekNumber);
    
    // Update the dashboard display with the week data
    // For example, populate a range with the data:
    const weekDataRange = dashboard.getRange('D20:H24'); // Adjust range as needed
    
    if (weekData.length > 0) {
      // Truncate to maximum 5 rows for display
      const displayData = weekData.slice(0, 5);
      weekDataRange.setValues(displayData);
    } else {
      // Clear the range if no data
      weekDataRange.clearContent();
    }
  } catch (e) {
    console.error('Error updating weekly data:', e);
  }
}

/**
 * Gets content data for a specific week
 */
function getWeekData(calendarSheet, weekNumber) {
  // Get all data from the content calendar
  const dataRange = calendarSheet.getDataRange();
  const data = dataRange.getValues();
  
  // Extract headers and data rows
  const headers = data[1]; // Assuming row 2 contains headers
  const rows = data.slice(2); // Skip header rows
  
  // Find index of relevant columns
  const weekIndex = headers.indexOf('Week');
  const dateIndex = headers.indexOf('Date');
  const channelIndex = headers.indexOf('Channel');
  const contentIndex = headers.indexOf('Content/Idea');
  const pillarIndex = headers.indexOf('Content Pillar');
  
  // Filter data for the specified week
  const weekData = rows.filter(row => row[weekIndex] == weekNumber)
    .map(row => [
      formatDate(row[dateIndex]), // Format date
      row[channelIndex], // Channel
      truncateText(row[contentIndex], 30), // Truncated content
      row[pillarIndex], // Content pillar
      '' // Additional column if needed
    ])
    .sort((a, b) => {
      // Sort by date
      const dateA = new Date(a[0]);
      const dateB = new Date(b[0]);
      return dateA - dateB;
    });
  
  return weekData;
}

/**
 * Formats a date for display
 */
function formatDate(date) {
  if (!(date instanceof Date)) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd');
}

/**
 * Truncates text to a specified length
 */
function truncateText(text, maxLength) {
  if (!text) return '';
  if (text.length <= maxLength) return text;
  return text.substring(0, maxLength) + '...';
}

/**
 * Creates a dashboard chart for the content status distribution
 * This can be used to programmatically create or update charts
 */
function createStatusDistributionChart() {
  // Get the spreadsheet and sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(DASHBOARD_CONFIG.SHEET_NAME);
  
  if (!dashboard) return;
  
  // Define the data range for status distribution
  const dataRange = dashboard.getRange('A5:C9'); // Adjust range as needed
  
  // Create a new chart
  const chart = dashboard.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(dataRange)
    .setPosition(5, 1, 0, 0) // row, column, offsetX, offsetY
    .setOption('title', 'Content by Status')
    .setOption('pieSliceText', 'percentage')
    .setOption('legend', { position: 'right' })
    .setOption('colors', ['#D0E0FF', '#FFFFD0', '#D0FFD0', '#FFE0C0', '#E0D0FF'])
    .setOption('is3D', false)
    .setOption('pieHole', 0)
    .build();
  
  // Add the chart to the dashboard
  dashboard.insertChart(chart);
}

/**
 * Creates a dashboard chart for the channel distribution
 */
function createChannelDistributionChart() {
  // Get the spreadsheet and sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(DASHBOARD_CONFIG.SHEET_NAME);
  
  if (!dashboard) return;
  
  // Define the data range for channel distribution
  const dataRange = dashboard.getRange('K5:P8'); // Adjust range as needed
  
  // Create a new chart
  const chart = dashboard.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(dataRange)
    .setPosition(5, 4, 0, 0) // row, column, offsetX, offsetY
    .setOption('title', 'Content by Channel')
    .setOption('isStacked', 'true')
    .setOption('legend', { position: 'bottom' })
    .setOption('hAxis', { 
      title: 'Number of Content Items',
      minValue: 0
    })
    .setOption('vAxis', { 
      title: 'Channel'
    })
    .setOption('colors', ['#D0E0FF', '#FFFFD0', '#D0FFD0', '#FFE0C0', '#E0D0FF'])
    .build();
  
  // Add the chart to the dashboard
  dashboard.insertChart(chart);
}

/**
 * Updates a dynamic dashboard widget with performance metrics
 * This illustrates how to update dashboard widgets with current data
 */
function updatePerformanceMetricsWidget() {
  // Get the spreadsheet and sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(DASHBOARD_CONFIG.SHEET_NAME);
  const calendarSheet = ss.getSheetByName('Content Calendar');
  
  if (!dashboard || !calendarSheet) return;
  
  // Calculate metrics
  const metrics = calculatePerformanceMetrics(calendarSheet);
  
  // Update dashboard cells with calculated metrics
  dashboard.getRange('B2').setValue(metrics.totalActive);
  dashboard.getRange('B3').setValue(metrics.publishedThisMonth);
  dashboard.getRange('D2').setValue(metrics.scheduledNextWeek);
  dashboard.getRange('D3').setValue(metrics.needsAttention);
  
  // Calculate completion ratio
  if (metrics.totalActive > 0) {
    const completionRatio = metrics.publishedThisMonth / metrics.totalActive;
    dashboard.getRange('F3').setValue(completionRatio);
  } else {
    dashboard.getRange('F3').setValue(0);
  }
}

/**
 * Calculates performance metrics from calendar data
 */
function calculatePerformanceMetrics(calendarSheet) {
  // Get all data
  const dataRange = calendarSheet.getDataRange();
  const data = dataRange.getValues();
  
  // Extract headers and data rows
  const headers = data[1]; // Assuming row 2 contains headers
  const rows = data.slice(2); // Skip header rows
  
  // Find column indexes
  const dateIndex = headers.indexOf('Date');
  const statusIndex = headers.indexOf('Status');
  const weekIndex = headers.indexOf('Week');
  
  // Current date information
  const today = new Date();
  const currentWeek = getCurrentWeekNumber();
  const nextWeek = parseInt(currentWeek) + 1;
  const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  const monthEnd = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  
  // Calculate metrics
  let totalActive = 0;
  let publishedThisMonth = 0;
  let scheduledNextWeek = 0;
  let needsAttention = 0;
  
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const date = row[dateIndex];
    const status = row[statusIndex];
    const week = row[weekIndex];
    
    // Skip empty rows
    if (!status) continue;
    
    // Count total active items
    totalActive++;
    
    // Count published this month
    if (status === 'Schedule' && date instanceof Date &&
        date >= monthStart && date <= monthEnd) {
      publishedThisMonth++;
    }
    
    // Count scheduled for next week
    if (week == nextWeek) {
      scheduledNextWeek++;
    }
    
    // Count items needing attention
    if (date instanceof Date && date < today && status !== 'Schedule') {
      needsAttention++;
    }
  }
  
  return {
    totalActive: totalActive,
    publishedThisMonth: publishedThisMonth,
    scheduledNextWeek: scheduledNextWeek,
    needsAttention: needsAttention
  };
}

/**
 * Creates a PDF export of the dashboard
 */
function exportDashboardToPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(DASHBOARD_CONFIG.SHEET_NAME);
  
  if (!dashboard) {
    SpreadsheetApp.getUi().alert('Dashboard sheet not found.');
    return;
  }
  
  // Create PDF blob
  const url = 'https://docs.google.com/spreadsheets/d/' + 
              ss.getId() + 
              '/export?format=pdf&gid=' + 
              dashboard.getSheetId() + 
              '&portrait=false&size=letter&fitw=true';
  
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });
  
  const blob = response.getBlob().setName('Content Calendar Dashboard.pdf');
  
  // Create a dialog with download link
  const html = HtmlService.createHtmlOutput(
    '<p>Your dashboard export is ready.</p>' +
    '<p><a href="' + url + '" target="_blank">Click here to download</a></p>' +
    '<p>Or you can close this dialog and check your Downloads folder.</p>'
  )
  .setWidth(300)
  .setHeight(200);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard Export Ready');
  
  return blob;
}

/**
 * Emails the dashboard as a PDF to specified recipients
 */
function emailDashboardReport() {
  // Get settings
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  
  if (!settingsSheet) {
    SpreadsheetApp.getUi().alert('Settings sheet not found.');
    return;
  }
  
  // Get email recipient from settings
  const emailAddress = settingsSheet.getRange('B7').getValue();
  
  if (!emailAddress) {
    SpreadsheetApp.getUi().alert('No email address found in Settings sheet (B7).');
    return;
  }
  
  // Create PDF export
  const pdfBlob = exportDashboardToPDF();
  
  // Send email with PDF attachment
  MailApp.sendEmail({
    to: emailAddress,
    subject: 'Content Calendar Dashboard Report - ' + 
             Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    body: 'Please find attached the latest content calendar dashboard report.',
    attachments: [pdfBlob]
  });
  
  SpreadsheetApp.getUi().alert('Dashboard report sent to ' + emailAddress);
}

/**
 * Jumps to a specific view of the content calendar
 */
function jumpToCalendarView(viewType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarSheet = ss.getSheetByName('Content Calendar');
  
  if (!calendarSheet) {
    SpreadsheetApp.getUi().alert('Content Calendar sheet not found.');
    return;
  }
  
  // Activate the calendar sheet
  calendarSheet.activate();
  
  // Get current date information
  const today = new Date();
  const currentWeek = getCurrentWeekNumber();
  
  // Clear any existing filters
  let existingFilters = calendarSheet.getFilter();
  if (existingFilters) {
    existingFilters.remove();
  }
  
  // Create a filter for the content data
  const lastRow = Math.max(calendarSheet.getLastRow(), 3);
  const lastCol = Math.max(calendarSheet.getLastColumn(), 14);
  const range = calendarSheet.getRange(2, 1, lastRow - 1, lastCol);
  const filter = range.createFilter();
  
  // Apply different filter criteria based on view type
  switch (viewType) {
    case 'today':
      // Filter for today's content
      filter.setColumnFilterCriteria(2, SpreadsheetApp.newFilterCriteria()
        .whenDateEqualTo(today)
        .build());
      break;
      
    case 'thisWeek':
      // Filter for this week's content
      filter.setColumnFilterCriteria(3, SpreadsheetApp.newFilterCriteria()
        .whenNumberEqualTo(currentWeek)
        .build());
      break;
      
    case 'nextWeek':
      // Filter for next week's content
      filter.setColumnFilterCriteria(3, SpreadsheetApp.newFilterCriteria()
        .whenNumberEqualTo(parseInt(currentWeek) + 1)
        .build());
      break;
      
    case 'needsAttention':
      // Filter for items that need attention (past due and not scheduled)
      filter.setColumnFilterCriteria(2, SpreadsheetApp.newFilterCriteria()
        .whenDateBefore(today)
        .build());
      filter.setColumnFilterCriteria(4, SpreadsheetApp.newFilterCriteria()
        .whenTextDoesNotContain('Schedule')
        .build());
      break;
      
    default:
      // No filtering
      filter.remove();
      break;
  }
}

/**
 * Updates the dashboard menu with dashboard-specific items
 */
function updateDashboardMenu() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.createMenu('Dashboard')
      .addItem('Refresh Data', 'refreshDashboard')
      .addSeparator()
      .addItem('Go to Current Week', 'currentWeek')
      .addItem('Go to Previous Week', 'previousWeek')
      .addItem('Go to Next Week', 'nextWeek')
      .addSeparator()
      .addItem('Reset All Filters', 'resetDashboardFilters')
      .addSeparator()
      .addItem('Export Dashboard to PDF', 'exportDashboardToPDF')
      .addItem('Email Dashboard Report', 'emailDashboardReport')
      .addToUi();
  } catch (e) {
    // Menu might already exist, no need to handle error
    console.error('Error creating dashboard menu:', e);
  }
}