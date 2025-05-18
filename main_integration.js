/**
 * Main Integration Script for Social Media Content Calendar
 * 
 * This script integrates all the individual functionality modules:
 * - Calendar core functionality
 * - Status tracking
 * - Drive integration
 * - Date and week utilities
 * - Workflow automation
 */

/**
 * Legacy menu creation function - NOT triggered automatically anymore.
 * 
 * This function has been renamed from onOpen_Legacy to createLegacyMenu 
 * to avoid confusing Apps Script about which onOpen is the real entry point.
 * It can be called manually if needed but is NOT an onOpen trigger handler.
 */
function createLegacyMenu() {
  try {
    // Create the main menu
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Legacy Content Calendar')
      .addItem('Create New Content Item', 'createNewContentItem')
      .addSeparator()
      .addSubMenu(ui.createMenu('Status')
        .addItem('Advance to Next Status', 'advanceStatus')
        .addItem('Revert to Previous Status', 'revertStatus')
        .addItem('View Status History', 'showStatusHistory'))
      .addSubMenu(ui.createMenu('Drive')
        .addItem('Link Asset from Drive', 'openFilePicker')
        .addItem('Create New Asset', 'createNewAsset')
        .addItem('Preview Selected Asset', 'previewAsset')
        .addItem('Open Assets Folder', 'openAssetsFolder'))
      .addSubMenu(ui.createMenu('Dates')
        .addItem('Auto-Populate Dates', 'autoPopulateDates')
        .addItem('Jump to Week...', 'jumpToWeek')
        .addItem('View Current Week', 'viewCurrentWeek')
        .addItem('View Next Week', 'viewNextWeek'))
      .addSubMenu(ui.createMenu('Workflow')
        .addItem('Run Content Automation...', 'runContentAutomation')
        .addItem('Generate Content Template', 'generateContentTemplate')
        .addItem('Auto-Assign Team Member', 'autoAssignTeamMember')
        .addItem('Calculate Due Dates', 'calculateDueDates')
        .addItem('Create Workflow Calendar Event', 'createWorkflowEvent'))
      .addSeparator()
      .addItem('Refresh Dashboard', 'refreshDashboard')
      .addItem('Update Data Validation', 'updateDataValidation')
      .addSeparator()
      .addItem('Archive Old Content', 'archiveOldContent')
      .addSeparator()
      .addItem('Help & Instructions', 'showHelp')
      .addToUi();
      
    // Attach keyboard shortcuts if supported
    // Note: This requires advanced permissions and may not work in all environments
    try {
      const doc = SpreadsheetApp.getActive();
      doc.setShortcut('Ctrl+Shift+N', 'createNewContentItem');
      doc.setShortcut('Ctrl+Shift+A', 'advanceStatus');
      doc.setShortcut('Ctrl+Shift+R', 'revertStatus');
      doc.setShortcut('Ctrl+Shift+L', 'openFilePicker');
    } catch (e) {
      // Shortcuts not supported, continue without them
      Logger.log('Shortcuts not supported: ' + e.toString());
    }
    
    // Check if the calendar is set up
    checkCalendarSetup();
    
    return true;
  } catch (e) {
    // Log error but don't interrupt the process
    Logger.log('Error in createLegacyMenu: ' + e.toString());
    return false;
  }
}

/**
 * Processes edits made to the spreadsheet
 * This function has been renamed from onEdit to handleContentEdit
 * to avoid conflicts with the trigger function in trigger_functions.js
 */
function handleContentEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    
    // Only process edits in the main calendar sheet
    if (sheetName !== 'Content Calendar') return;
    
    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();
    
    // Skip header rows
    if (row < 3) return;
    
    // Update Last Modified timestamp
    updateLastModified(sheet, row);
    
    // Track status changes if status column was edited
    if (col === 4) { // Status column
      // Get old and new status values
      const oldStatus = e.oldValue || '';
      const newStatus = range.getValue() || '';
      
      // Track the status change
      trackStatusChange(sheet, row, oldStatus, newStatus);
    }
    
    // Update week number if date was changed
    if (col === 2) { // Date column
      updateWeekNumber(sheet, row);
    }
    
  } catch (e) {
    // Log error but don't interrupt the process
    console.error('Error in onEdit: ' + e.toString());
  }
}

/**
 * Checks if the calendar is set up and prompts for setup if needed
 */
function checkCalendarSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if required sheets exist
  const requiredSheets = ['Content Calendar', 'Dashboard', 'Lists', 'Settings'];
  let missingSheets = [];
  
  requiredSheets.forEach(function(sheetName) {
    if (!ss.getSheetByName(sheetName)) {
      missingSheets.push(sheetName);
    }
  });
  
  // If any sheets are missing, prompt for setup
  if (missingSheets.length > 0) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Setup Required',
      'This spreadsheet is not set up as a content calendar. ' + 
      'Missing sheets: ' + missingSheets.join(', ') + '. ' +
      'Would you like to set up the content calendar now?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      setupContentCalendar();
    }
  }
}

/**
 * Refreshes the dashboard data and charts
 */
function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName('Dashboard');
  
  if (!dashboardSheet) {
    SpreadsheetApp.getUi().alert('Dashboard sheet not found.');
    return;
  }
  
  // Implementation depends on dashboard design
  // For this example, we'll assume the dashboard has formula-based calculations
  // that will update automatically when we refresh the sheet
  
  // Force recalculation of all formulas
  dashboardSheet.getRange('A1').setValue(dashboardSheet.getRange('A1').getValue());
  
  // Update any charts
  // (Charts typically update automatically with their source data)
  
  // Show confirmation
  SpreadsheetApp.getUi().alert('Dashboard refreshed!');
}

/**
 * Shows help and instructions
 */
function showHelp() {
  try {
    // Create HTML content with help information
    const htmlContent = `
      <h2>Content Calendar Help</h2>
      <h3>Basic Usage</h3>
      <ul>
        <li><strong>Create New Content:</strong> Use the "Create New Content Item" menu option</li>
        <li><strong>Set Publication Date:</strong> Enter the date in the Date column</li>
        <li><strong>Update Status:</strong> Select from dropdown as content progresses</li>
        <li><strong>Link Assets:</strong> Use "Link Google Drive Asset" to attach files</li>
      </ul>
      <h3>Keyboard Shortcuts</h3>
      <ul>
        <li>Press <strong>Ctrl+Shift+N</strong> to create a new content item</li>
        <li>Press <strong>Ctrl+Shift+A</strong> to advance status</li>
        <li>Press <strong>Ctrl+Shift+R</strong> to revert status</li>
        <li>Press <strong>Ctrl+Shift+L</strong> to link a Drive asset</li>
      </ul>
      <h3>Tips</h3>
      <ul>
        <li>Use the Dashboard for a quick overview of your content</li>
        <li>Filter by Week number to focus on specific timeframes</li>
        <li>Archive old content to keep the calendar manageable</li>
        <li>Use templates to speed up content creation</li>
      </ul>
      <h3>Support</h3>
      <p>For additional help or to report issues, please see the documentation or contact the administrator.</p>
    `;
    
    // Use createSafeHtmlOutput if available, otherwise create HTML with IFRAME mode manually
    let htmlOutput;
    if (typeof createSafeHtmlOutput === 'function') {
      // Use the helper function from main_menu.js if available
      htmlOutput = createSafeHtmlOutput('HelpDocumentation', 500, 400);
      if (!htmlOutput) {
        throw new Error("Failed to create HTML output from HelpDocumentation file");
      }
    } else {
      // If helper function not available, create HTML output with IFRAME mode manually
      htmlOutput = HtmlService.createHtmlOutput(htmlContent)
        .setWidth(500)
        .setHeight(400)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    }
    
    // Show the help dialog
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Content Calendar Help');
    Logger.log("Help documentation displayed successfully");
  } catch (e) {
    Logger.log(`Error showing help: ${e.message}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Could not display help. Please try again later.",
      "Help Error",
      7
    );
  }
}

/**
 * Runs daily automation tasks
 * This function should be set up as a time-driven trigger to run daily
 */
function dailyAutomation() {
  try {
    // Check for overdue items
    checkOverdueItems();
    
    // Send status reminders
    sendStatusReminders();
    
    // Update dashboard stats
    refreshDashboard();
    
  } catch (e) {
    // Log error but don't interrupt the process
    console.error('Error in dailyAutomation: ' + e.toString());
  }
}

/**
 * Checks for overdue items and sends notifications
 */
function checkOverdueItems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Content Calendar');
  
  // Skip if sheet doesn't exist
  if (!sheet) return;
  
  // Get today's date
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  // Find overdue items (date is in the past but status is not "Schedule")
  const lastRow = Math.max(sheet.getLastRow(), 3);
  const dateRange = sheet.getRange(3, 2, lastRow - 2, 1); // Column B (Date)
  const statusRange = sheet.getRange(3, 4, lastRow - 2, 1); // Column D (Status)
  
  const dates = dateRange.getValues();
  const statuses = statusRange.getValues();
  
  let overdueItems = [];
  
  for (let i = 0; i < dates.length; i++) {
    const date = dates[i][0];
    const status = statuses[i][0];
    
    // Skip if no date or not a date
    if (!date || !(date instanceof Date)) continue;
    
    // Set time to midnight for comparison
    const itemDate = new Date(date);
    itemDate.setHours(0, 0, 0, 0);
    
    // Check if overdue (date is before today and status is not "Schedule")
    if (itemDate < today && status !== 'Schedule') {
      overdueItems.push({
        row: i + 3, // +3 because we started at row 3
        date: itemDate,
        status: status
      });
    }
  }
  
  // Send notifications if enabled
  if (overdueItems.length > 0) {
    // Get settings
    const settingsSheet = ss.getSheetByName('Settings');
    const notifyEmail = settingsSheet.getRange('B7').getValue(); // Assuming notification email is stored here
    
    if (notifyEmail) {
      // Compose the email
      const subject = `[Content Calendar] ${overdueItems.length} Overdue Items`;
      let body = `The following items in the content calendar are overdue:\n\n`;
      
      overdueItems.forEach(function(item) {
        const contentId = sheet.getRange(item.row, 1).getValue();
        const contentTitle = sheet.getRange(item.row, 6).getValue() || '[No title]';
        const formattedDate = Utilities.formatDate(item.date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        
        body += `- ${contentId}: "${contentTitle}" (Due: ${formattedDate}, Status: ${item.status})\n`;
      });
      
      body += `\nAccess the content calendar: ${ss.getUrl()}`;
      
      // Send the email
      try {
        MailApp.sendEmail(notifyEmail, subject, body);
      } catch (e) {
        console.error('Failed to send overdue items notification:', e);
      }
    }
    
    // Apply conditional formatting to overdue items
    // This should be done via the onOpen function to ensure it's always applied
    // but we'll refresh it here just in case
  }
}

/**
 * Sends reminders for items that need attention
 */
function sendStatusReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Content Calendar');
  
  // Skip if sheet doesn't exist
  if (!sheet) return;
  
  // Get settings
  const settingsSheet = ss.getSheetByName('Settings');
  const reminderEnabled = settingsSheet.getRange('B12').getValue(); // Assuming reminder setting is stored here
  
  // Skip if reminders are disabled
  if (!reminderEnabled) return;
  
  // Get today's date
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  // Add 3 days to today for upcoming items
  const upcomingDate = new Date(today);
  upcomingDate.setDate(upcomingDate.getDate() + 3);
  
  // Find items needing attention (due within 3 days and status is not final)
  const lastRow = Math.max(sheet.getLastRow(), 3);
  const dateRange = sheet.getRange(3, 2, lastRow - 2, 1); // Column B (Date)
  const statusRange = sheet.getRange(3, 4, lastRow - 2, 1); // Column D (Status)
  const assignedRange = sheet.getRange(3, 10, lastRow - 2, 1); // Column J (Assigned To)
  
  const dates = dateRange.getValues();
  const statuses = statusRange.getValues();
  const assigned = assignedRange.getValues();
  
  let upcomingItems = [];
  
  for (let i = 0; i < dates.length; i++) {
    const date = dates[i][0];
    const status = statuses[i][0];
    const assignedTo = assigned[i][0];
    
    // Skip if no date or not a date
    if (!date || !(date instanceof Date)) continue;
    
    // Set time to midnight for comparison
    const itemDate = new Date(date);
    itemDate.setHours(0, 0, 0, 0);
    
    // Check if upcoming (due within 3 days and status is not final)
    if (itemDate <= upcomingDate && itemDate >= today && status !== 'Schedule') {
      upcomingItems.push({
        row: i + 3, // +3 because we started at row 3
        date: itemDate,
        status: status,
        assignedTo: assignedTo
      });
    }
  }
  
  // Send notifications for upcoming items
  if (upcomingItems.length > 0) {
    // Get notification email
    const notifyEmail = settingsSheet.getRange('B7').getValue(); // Assuming notification email is stored here
    
    if (notifyEmail) {
      // Compose the email
      const subject = `[Content Calendar] ${upcomingItems.length} Upcoming Items`;
      let body = `The following items in the content calendar are coming up soon:\n\n`;
      
      upcomingItems.forEach(function(item) {
        const contentId = sheet.getRange(item.row, 1).getValue();
        const contentTitle = sheet.getRange(item.row, 6).getValue() || '[No title]';
        const formattedDate = Utilities.formatDate(item.date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        
        body += `- ${contentId}: "${contentTitle}" (Due: ${formattedDate}, Status: ${item.status}, Assigned to: ${item.assignedTo})\n`;
      });
      
      body += `\nAccess the content calendar: ${ss.getUrl()}`;
      
      // Send the email
      try {
        MailApp.sendEmail(notifyEmail, subject, body);
      } catch (e) {
        console.error('Failed to send upcoming items notification:', e);
      }
    }
  }
}