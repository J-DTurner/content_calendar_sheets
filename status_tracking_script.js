/**
 * Status Tracking System for Social Media Content Calendar
 * 
 * This script implements comprehensive status tracking functionality including:
 * - Automatic status change logging
 * - Status history tracking
 * - Status change notifications
 * - Status-based workflow automation
 */

// Status tracking configuration
const STATUS_CONFIG = {
  COLUMN: 4,                // Column D
  CHANGED_COLUMN: 14,       // Column N
  HISTORY_SHEET: 'Status History',
  NOTIFICATION_ENABLED: true,
  WORKFLOW_AUTOMATION: true
};

// Status values and their workflow order
const STATUS_VALUES = [
  'Planned',
  'Copywriting Complete',
  'Creative Completed',
  'Ready for Review',
  'Schedule'
];

/**
 * Tracks status changes when a cell in the status column is edited
 * This function is called by the onEdit trigger
 */
function trackStatusChange(sheet, row, oldStatus, newStatus) {
  // Skip if no actual change
  if (oldStatus === newStatus) return;
  
  // Update the status changed timestamp
  const now = new Date();
  sheet.getRange(row, STATUS_CONFIG.CHANGED_COLUMN).setValue(now);
  
  // Log the status change in the history sheet
  logStatusChange(sheet.getName(), row, oldStatus, newStatus, now);
  
  // Send notification if enabled
  if (STATUS_CONFIG.NOTIFICATION_ENABLED) {
    notifyStatusChange(sheet, row, oldStatus, newStatus);
  }
  
  // Apply workflow automation if enabled
  if (STATUS_CONFIG.WORKFLOW_AUTOMATION) {
    applyWorkflowRules(sheet, row, oldStatus, newStatus);
  }
}

/**
 * Logs a status change to the history sheet
 */
function logStatusChange(sheetName, row, oldStatus, newStatus, timestamp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create the history sheet if it doesn't exist
  let historySheet = ss.getSheetByName(STATUS_CONFIG.HISTORY_SHEET);
  if (!historySheet) {
    historySheet = ss.insertSheet(STATUS_CONFIG.HISTORY_SHEET);
    setupHistorySheet(historySheet);
  }
  
  // Get the content ID and title for reference
  const contentSheet = ss.getSheetByName(sheetName);
  const contentId = contentSheet.getRange(row, 1).getValue();
  const contentTitle = contentSheet.getRange(row, 6).getValue();
  const contentDate = contentSheet.getRange(row, 2).getValue();
  
  // Get user email
  const userEmail = Session.getActiveUser().getEmail();
  
  // Log the change to history
  const lastRow = Math.max(historySheet.getLastRow(), 1);
  historySheet.getRange(lastRow + 1, 1, 1, 7).setValues([[
    timestamp,
    contentId,
    contentTitle,
    contentDate,
    oldStatus,
    newStatus,
    userEmail
  ]]);
}

/**
 * Sets up the history sheet with headers and formatting
 */
function setupHistorySheet(sheet) {
  // Set headers
  const headers = [
    'Timestamp', 
    'Content ID', 
    'Content Title',
    'Publication Date',
    'Previous Status', 
    'New Status', 
    'Changed By'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4285F4')
    .setFontColor('white')
    .setFontWeight('bold');
  
  // Set column widths
  sheet.setColumnWidth(1, 180);  // Timestamp
  sheet.setColumnWidth(2, 100);  // Content ID
  sheet.setColumnWidth(3, 300);  // Content Title
  sheet.setColumnWidth(4, 120);  // Publication Date
  sheet.setColumnWidth(5, 150);  // Previous Status
  sheet.setColumnWidth(6, 150);  // New Status
  sheet.setColumnWidth(7, 200);  // Changed By
  
  // Freeze the header row
  sheet.setFrozenRows(1);
  
  // Set timestamp format
  sheet.getRange(2, 1, 999, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');
  
  // Set date format
  sheet.getRange(2, 4, 999, 1).setNumberFormat('yyyy-MM-dd');
}

/**
 * Sends notifications for status changes
 */
function notifyStatusChange(sheet, row, oldStatus, newStatus) {
  // Get content information for the notification
  const contentId = sheet.getRange(row, 1).getValue();
  const contentTitle = sheet.getRange(row, 6).getValue();
  const channel = sheet.getRange(row, 5).getValue();
  const assignedTo = sheet.getRange(row, 10).getValue();
  
  // Get settings
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  
  // Determine who should receive the notification
  let recipientEmail = '';
  
  // Try to find the email in a team members sheet or use a default
  try {
    const teamSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Team Members');
    if (teamSheet) {
      // Find the row with the assigned team member
      const teamData = teamSheet.getDataRange().getValues();
      for (let i = 1; i < teamData.length; i++) {
        if (teamData[i][0] === assignedTo) {
          recipientEmail = teamData[i][1]; // Assuming email is in column B
          break;
        }
      }
    }
  } catch (e) {
    // If no email found, use the default notification email
    recipientEmail = settingsSheet.getRange('B7').getValue(); // Assuming default email is stored here
  }
  
  // If no recipient found, exit
  if (!recipientEmail) return;
  
  // Compose the email
  const subject = `[Content Calendar] Status Changed: ${contentTitle}`;
  const body = `
    Content item ${contentId} has changed status:
    
    Title: ${contentTitle}
    Channel: ${channel}
    Previous Status: ${oldStatus}
    New Status: ${newStatus}
    
    View the content calendar for more details.
  `;
  
  // Send the email
  try {
    MailApp.sendEmail(recipientEmail, subject, body);
  } catch (e) {
    // Log error but don't interrupt the process
    console.error('Failed to send status change notification:', e);
  }
}

/**
 * Applies workflow rules based on status changes
 */
function applyWorkflowRules(sheet, row, oldStatus, newStatus) {
  // Get status indices for comparison (to know the direction of change)
  const oldIndex = STATUS_VALUES.indexOf(oldStatus);
  const newIndex = STATUS_VALUES.indexOf(newStatus);
  
  // Skip if invalid status values
  if (oldIndex === -1 || newIndex === -1) return;
  
  // Get content information
  const contentId = sheet.getRange(row, 1).getValue();
  const contentTitle = sheet.getRange(row, 6).getValue();
  
  // Moving to "Copywriting Complete"
  if (newStatus === 'Copywriting Complete') {
    // Set default assigned person for the next stage (if not already set)
    const assignedCell = sheet.getRange(row, 10);
    if (!assignedCell.getValue()) {
      // Get designer from settings
      const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
      const defaultDesigner = settingsSheet.getRange('B6').getValue(); // Assume default designer is stored here
      if (defaultDesigner) {
        assignedCell.setValue(defaultDesigner);
      }
    }
  }
  
  // Moving to "Ready for Review"
  if (newStatus === 'Ready for Review') {
    // Change assigned person to the reviewer
    const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
    const defaultReviewer = settingsSheet.getRange('B5').getValue(); // Assume default reviewer is stored here
    if (defaultReviewer) {
      sheet.getRange(row, 10).setValue(defaultReviewer);
    }
  }
  
  // Moving to "Schedule"
  if (newStatus === 'Schedule') {
    // Create calendar event if date is set
    const publishDate = sheet.getRange(row, 2).getValue();
    if (publishDate instanceof Date) {
      createCalendarEvent(contentId, contentTitle, publishDate, sheet.getRange(row, 5).getValue());
    }
  }
}

/**
 * Creates a calendar event for scheduled content
 */
function createCalendarEvent(contentId, contentTitle, publishDate, channel) {
  // Get settings
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  const calendarId = settingsSheet.getRange('B8').getValue(); // Assuming calendar ID is stored here
  
  // If no calendar ID, use the default calendar
  let calendar;
  try {
    if (calendarId) {
      calendar = CalendarApp.getCalendarById(calendarId);
    } else {
      calendar = CalendarApp.getDefaultCalendar();
    }
  } catch (e) {
    // If error accessing calendar, exit silently
    console.error('Failed to access calendar:', e);
    return;
  }
  
  // Set the event duration (default to 30 minutes)
  const endDate = new Date(publishDate.getTime() + 30 * 60000);
  
  // Create the event
  try {
    const event = calendar.createEvent(
      `[${channel}] ${contentTitle}`,
      publishDate,
      endDate,
      {
        description: `Content ID: ${contentId}\nChannel: ${channel}\nStatus: Schedule`,
        location: 'Content Calendar'
      }
    );
    
    // Get the event URL to store in the spreadsheet
    const eventId = event.getId();
    const eventUrl = 'https://www.google.com/calendar/event?eid=' + Utilities.base64Encode(eventId);
    
    // Store the event URL in a notes or dedicated column if available
    // For this example, we'll add it to the notes column
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Content Calendar');
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // Find the row with the matching content ID
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === contentId) {
        // Add event URL to the notes column (column K, index 10)
        const currentNotes = values[i][10] || '';
        const newNotes = currentNotes + (currentNotes ? '\n' : '') + 'Calendar Event: ' + eventUrl;
        sheet.getRange(i + 1, 11).setValue(newNotes);
        break;
      }
    }
  } catch (e) {
    console.error('Failed to create calendar event:', e);
  }
}

/**
 * Gets the status history for a specific content item
 * This can be called from a custom menu to show status history
 */
function showStatusHistory() {
  // Get the active sheet and selected row
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = SpreadsheetApp.getActiveRange().getRow();
  
  // Skip if not in the content calendar sheet or header row
  if (sheet.getName() !== 'Content Calendar' || row < 3) {
    SpreadsheetApp.getUi().alert('Please select a content item in the Content Calendar sheet.');
    return;
  }
  
  // Get the content ID
  const contentId = sheet.getRange(row, 1).getValue();
  
  // Get the history sheet
  const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_CONFIG.HISTORY_SHEET);
  if (!historySheet) {
    SpreadsheetApp.getUi().alert('No status history found.');
    return;
  }
  
  // Get all history data
  const historyData = historySheet.getDataRange().getValues();
  if (historyData.length <= 1) {
    SpreadsheetApp.getUi().alert('No status history found for this item.');
    return;
  }
  
  // Filter history for this content ID
  let itemHistory = [];
  for (let i = 1; i < historyData.length; i++) {
    if (historyData[i][1] === contentId) {
      itemHistory.push(historyData[i]);
    }
  }
  
  // If no history found
  if (itemHistory.length === 0) {
    SpreadsheetApp.getUi().alert('No status history found for this item.');
    return;
  }
  
  // Build HTML to display the history
  let html = '<h2>Status History for ' + contentId + '</h2>';
  html += '<table style="border-collapse: collapse; width: 100%;">';
  html += '<tr style="background-color: #f2f2f2; font-weight: bold;">';
  html += '<th style="border: 1px solid #ddd; padding: 8px;">Timestamp</th>';
  html += '<th style="border: 1px solid #ddd; padding: 8px;">Previous Status</th>';
  html += '<th style="border: 1px solid #ddd; padding: 8px;">New Status</th>';
  html += '<th style="border: 1px solid #ddd; padding: 8px;">Changed By</th>';
  html += '</tr>';
  
  // Add each history entry
  itemHistory.forEach(function(entry) {
    const timestamp = entry[0];
    const previousStatus = entry[4];
    const newStatus = entry[5];
    const changedBy = entry[6];
    
    // Format the timestamp
    let formattedTimestamp;
    if (timestamp instanceof Date) {
      formattedTimestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    } else {
      formattedTimestamp = timestamp;
    }
    
    html += '<tr>';
    html += '<td style="border: 1px solid #ddd; padding: 8px;">' + formattedTimestamp + '</td>';
    html += '<td style="border: 1px solid #ddd; padding: 8px;">' + previousStatus + '</td>';
    html += '<td style="border: 1px solid #ddd; padding: 8px;">' + newStatus + '</td>';
    html += '<td style="border: 1px solid #ddd; padding: 8px;">' + changedBy + '</td>';
    html += '</tr>';
  });
  
  html += '</table>';
  
  // Show the history in a modal dialog
  const htmlOutput = HtmlService
    .createHtmlOutput(html)
    .setWidth(600)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Status History');
}

/**
 * Validates a status change based on workflow rules
 * Returns true if the status change is valid, false otherwise
 */
function validateStatusChange(oldStatus, newStatus) {
  // Get indices in the workflow
  const oldIndex = STATUS_VALUES.indexOf(oldStatus);
  const newIndex = STATUS_VALUES.indexOf(newStatus);
  
  // Invalid status values
  if (oldIndex === -1 || newIndex === -1) return false;
  
  // Only allow moving one step forward or backward in the workflow
  return Math.abs(newIndex - oldIndex) <= 1;
}

/**
 * Returns the next status in the workflow
 */
function getNextStatus(currentStatus) {
  const currentIndex = STATUS_VALUES.indexOf(currentStatus);
  if (currentIndex === -1 || currentIndex === STATUS_VALUES.length - 1) return currentStatus;
  return STATUS_VALUES[currentIndex + 1];
}

/**
 * Returns the previous status in the workflow
 */
function getPreviousStatus(currentStatus) {
  const currentIndex = STATUS_VALUES.indexOf(currentStatus);
  if (currentIndex <= 0) return currentStatus;
  return STATUS_VALUES[currentIndex - 1];
}

/**
 * Advances the status of the selected content item to the next status in the workflow
 * This can be called from a custom menu or button
 */
function advanceStatus() {
  // Get the active sheet and selected row
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = SpreadsheetApp.getActiveRange().getRow();
  
  // Skip if not in the content calendar sheet or header row
  if (sheet.getName() !== 'Content Calendar' || row < 3) {
    SpreadsheetApp.getUi().alert('Please select a content item in the Content Calendar sheet.');
    return;
  }
  
  // Get the current status
  const statusCell = sheet.getRange(row, STATUS_CONFIG.COLUMN);
  const currentStatus = statusCell.getValue();
  
  // Get the next status
  const nextStatus = getNextStatus(currentStatus);
  
  // If already at the last status, show a message
  if (nextStatus === currentStatus) {
    SpreadsheetApp.getUi().alert('This item is already at the final status.');
    return;
  }
  
  // Set the new status
  const oldStatus = currentStatus;
  statusCell.setValue(nextStatus);
  
  // Track the status change (will be called by onEdit trigger, but let's call it directly to be safe)
  trackStatusChange(sheet, row, oldStatus, nextStatus);
  
  // Show confirmation
  SpreadsheetApp.getUi().alert('Status advanced to: ' + nextStatus);
}

/**
 * Moves the status of the selected content item back to the previous status in the workflow
 * This can be called from a custom menu or button
 */
function revertStatus() {
  // Get the active sheet and selected row
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = SpreadsheetApp.getActiveRange().getRow();
  
  // Skip if not in the content calendar sheet or header row
  if (sheet.getName() !== 'Content Calendar' || row < 3) {
    SpreadsheetApp.getUi().alert('Please select a content item in the Content Calendar sheet.');
    return;
  }
  
  // Get the current status
  const statusCell = sheet.getRange(row, STATUS_CONFIG.COLUMN);
  const currentStatus = statusCell.getValue();
  
  // Get the previous status
  const previousStatus = getPreviousStatus(currentStatus);
  
  // If already at the first status, show a message
  if (previousStatus === currentStatus) {
    SpreadsheetApp.getUi().alert('This item is already at the initial status.');
    return;
  }
  
  // Set the new status
  const oldStatus = currentStatus;
  statusCell.setValue(previousStatus);
  
  // Track the status change (will be called by onEdit trigger, but let's call it directly to be safe)
  trackStatusChange(sheet, row, oldStatus, previousStatus);
  
  // Show confirmation
  SpreadsheetApp.getUi().alert('Status reverted to: ' + previousStatus);
}

/**
 * Updates the custom menu with status-related commands
 */
function updateStatusMenu() {
  // Check if the menu already exists
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.createMenu('Status Actions')
      .addItem('Advance to Next Status', 'advanceStatus')
      .addItem('Revert to Previous Status', 'revertStatus')
      .addSeparator()
      .addItem('View Status History', 'showStatusHistory')
      .addToUi();
  } catch (e) {
    // Menu might already exist, no need to handle error
  }
}