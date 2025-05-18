/**
 * Email Notification System for Social Media Content Calendar
 * 
 * This script implements comprehensive email notifications for status changes,
 * deadline alerts, and other important events in the content calendar workflow.
 */

// Notification system configuration
const NOTIFICATION_CONFIG = {
  ENABLED: true,
  SETTINGS_EMAIL_CELL: 'B7',  // Cell containing notification email in Settings sheet
  TEAM_EMAIL_COLUMN: 2,       // Column B in Team Members sheet
  EMAIL_LOGS_SHEET: 'Email Logs',
  GROUP_NOTIFICATIONS: true,   // Group multiple notifications into a single email
  NOTIFICATION_TYPES: {
    STATUS_CHANGE: true,
    APPROACHING_DEADLINE: true,
    OVERDUE_ITEMS: true,
    CONTENT_ASSIGNMENT: true,
    WEEKLY_SUMMARY: true
  },
  EMAIL_TEMPLATES: {
    STATUS_CHANGE: {
      SUBJECT: "[Content Calendar] Status Changed: {contentId}",
      BODY: "Content item {contentId} has changed status:\n\n" +
            "Title: {contentTitle}\n" +
            "Channel: {channel}\n" +
            "Previous Status: {oldStatus}\n" +
            "New Status: {newStatus}\n" +
            "Changed by: {user}\n" +
            "Date: {date}\n\n" +
            "View the content calendar: {spreadsheetUrl}"
    },
    APPROACHING_DEADLINE: {
      SUBJECT: "[Content Calendar] Approaching Deadline: {contentId}",
      BODY: "The following content item is approaching its deadline:\n\n" +
            "Content ID: {contentId}\n" +
            "Title: {contentTitle}\n" +
            "Current Status: {status}\n" +
            "Publication Date: {publicationDate}\n" +
            "Days Remaining: {daysRemaining}\n" +
            "Assigned To: {assignedTo}\n\n" +
            "View the content calendar: {spreadsheetUrl}"
    },
    OVERDUE_ITEMS: {
      SUBJECT: "[Content Calendar] Overdue Content Items",
      BODY: "The following content items are overdue:\n\n" +
            "{overdueList}\n\n" +
            "View the content calendar: {spreadsheetUrl}"
    },
    CONTENT_ASSIGNMENT: {
      SUBJECT: "[Content Calendar] New Content Assigned: {contentId}",
      BODY: "You have been assigned a new content item:\n\n" +
            "Content ID: {contentId}\n" +
            "Title: {contentTitle}\n" +
            "Channel: {channel}\n" +
            "Current Status: {status}\n" +
            "Publication Date: {publicationDate}\n\n" +
            "View the content calendar: {spreadsheetUrl}"
    },
    WEEKLY_SUMMARY: {
      SUBJECT: "[Content Calendar] Weekly Summary - Week {weekNumber}",
      BODY: "Here's your content calendar summary for Week {weekNumber}:\n\n" +
            "Total Active Items: {totalActive}\n" +
            "Items Published: {published}\n" +
            "Items in Progress: {inProgress}\n" +
            "Upcoming Deadlines: {upcoming}\n\n" +
            "Status Breakdown:\n{statusBreakdown}\n\n" +
            "View the content calendar: {spreadsheetUrl}"
    }
  }
};

// Global notification queue to collect notifications for batching
let notificationQueue = [];

/**
 * Sends a status change notification
 * @param {object} contentInfo Information about the content item
 * @param {string} oldStatus Previous status
 * @param {string} newStatus New status
 */
function sendStatusChangeNotification(contentInfo, oldStatus, newStatus) {
  // Skip if notifications are disabled
  if (!NOTIFICATION_CONFIG.ENABLED || !NOTIFICATION_CONFIG.NOTIFICATION_TYPES.STATUS_CHANGE) {
    return;
  }
  
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get recipient(s)
  let recipients = [];
  
  // Always include the primary notification email
  const settingsSheet = ss.getSheetByName('Settings');
  const primaryEmail = settingsSheet.getRange(NOTIFICATION_CONFIG.SETTINGS_EMAIL_CELL).getValue();
  if (primaryEmail) {
    recipients.push(primaryEmail);
  }
  
  // Include the assigned team member's email if available
  if (contentInfo.assignedTo) {
    const teamEmail = getTeamMemberEmail(contentInfo.assignedTo);
    if (teamEmail && !recipients.includes(teamEmail)) {
      recipients.push(teamEmail);
    }
  }
  
  // If no recipients, exit
  if (recipients.length === 0) {
    console.log('No recipients found for status change notification');
    return;
  }
  
  // Replace template variables
  const subject = replaceTemplateVariables(
    NOTIFICATION_CONFIG.EMAIL_TEMPLATES.STATUS_CHANGE.SUBJECT,
    contentInfo,
    { oldStatus: oldStatus, newStatus: newStatus }
  );
  
  const body = replaceTemplateVariables(
    NOTIFICATION_CONFIG.EMAIL_TEMPLATES.STATUS_CHANGE.BODY,
    contentInfo,
    { 
      oldStatus: oldStatus, 
      newStatus: newStatus,
      user: Session.getActiveUser().getEmail(),
      date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      spreadsheetUrl: ss.getUrl()
    }
  );
  
  // Send the email or add to queue
  if (NOTIFICATION_CONFIG.GROUP_NOTIFICATIONS) {
    // Add to notification queue for batching
    notificationQueue.push({
      type: 'STATUS_CHANGE',
      recipients: recipients,
      subject: subject,
      body: body,
      contentInfo: contentInfo
    });
  } else {
    // Send immediately
    try {
      MailApp.sendEmail({
        to: recipients.join(','),
        subject: subject,
        body: body
      });
      
      // Log the notification
      logEmailNotification('STATUS_CHANGE', recipients.join(','), contentInfo.contentId);
    } catch (e) {
      console.error('Failed to send status change notification:', e);
    }
  }
}

/**
 * Sends an approaching deadline notification
 * @param {object} contentInfo Information about the content item
 * @param {number} daysRemaining Days remaining until deadline
 */
function sendApproachingDeadlineNotification(contentInfo, daysRemaining) {
  // Skip if notifications are disabled
  if (!NOTIFICATION_CONFIG.ENABLED || !NOTIFICATION_CONFIG.NOTIFICATION_TYPES.APPROACHING_DEADLINE) {
    return;
  }
  
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get recipient(s)
  let recipients = [];
  
  // Always include the primary notification email
  const settingsSheet = ss.getSheetByName('Settings');
  const primaryEmail = settingsSheet.getRange(NOTIFICATION_CONFIG.SETTINGS_EMAIL_CELL).getValue();
  if (primaryEmail) {
    recipients.push(primaryEmail);
  }
  
  // Include the assigned team member's email if available
  if (contentInfo.assignedTo) {
    const teamEmail = getTeamMemberEmail(contentInfo.assignedTo);
    if (teamEmail && !recipients.includes(teamEmail)) {
      recipients.push(teamEmail);
    }
  }
  
  // If no recipients, exit
  if (recipients.length === 0) {
    console.log('No recipients found for deadline notification');
    return;
  }
  
  // Replace template variables
  const subject = replaceTemplateVariables(
    NOTIFICATION_CONFIG.EMAIL_TEMPLATES.APPROACHING_DEADLINE.SUBJECT,
    contentInfo,
    { daysRemaining: daysRemaining }
  );
  
  const body = replaceTemplateVariables(
    NOTIFICATION_CONFIG.EMAIL_TEMPLATES.APPROACHING_DEADLINE.BODY,
    contentInfo,
    { 
      daysRemaining: daysRemaining,
      spreadsheetUrl: ss.getUrl()
    }
  );
  
  // Send the email or add to queue
  if (NOTIFICATION_CONFIG.GROUP_NOTIFICATIONS) {
    // Add to notification queue for batching
    notificationQueue.push({
      type: 'APPROACHING_DEADLINE',
      recipients: recipients,
      subject: subject,
      body: body,
      contentInfo: contentInfo
    });
  } else {
    // Send immediately
    try {
      MailApp.sendEmail({
        to: recipients.join(','),
        subject: subject,
        body: body
      });
      
      // Log the notification
      logEmailNotification('APPROACHING_DEADLINE', recipients.join(','), contentInfo.contentId);
    } catch (e) {
      console.error('Failed to send deadline notification:', e);
    }
  }
}

/**
 * Sends an overdue items notification
 * @param {Array} overdueItems List of overdue content items
 */
function sendOverdueItemsNotification(overdueItems) {
  // Skip if notifications are disabled or no overdue items
  if (!NOTIFICATION_CONFIG.ENABLED || 
      !NOTIFICATION_CONFIG.NOTIFICATION_TYPES.OVERDUE_ITEMS ||
      !overdueItems || overdueItems.length === 0) {
    return;
  }
  
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get primary notification email
  const settingsSheet = ss.getSheetByName('Settings');
  const primaryEmail = settingsSheet.getRange(NOTIFICATION_CONFIG.SETTINGS_EMAIL_CELL).getValue();
  
  // If no recipient, exit
  if (!primaryEmail) {
    console.log('No recipient found for overdue items notification');
    return;
  }
  
  // Build the list of overdue items
  let overdueList = '';
  
  for (let i = 0; i < overdueItems.length; i++) {
    const item = overdueItems[i];
    overdueList += `- ${item.contentId}: "${item.contentTitle}" (Due: ${item.formattedDate}, Status: ${item.status}, Assigned To: ${item.assignedTo})\n`;
  }
  
  // Replace template variables
  const subject = NOTIFICATION_CONFIG.EMAIL_TEMPLATES.OVERDUE_ITEMS.SUBJECT
    .replace('{count}', overdueItems.length);
  
  const body = NOTIFICATION_CONFIG.EMAIL_TEMPLATES.OVERDUE_ITEMS.BODY
    .replace('{overdueList}', overdueList)
    .replace('{spreadsheetUrl}', ss.getUrl());
  
  // Send the email or add to queue
  if (NOTIFICATION_CONFIG.GROUP_NOTIFICATIONS) {
    // Add to notification queue for batching
    notificationQueue.push({
      type: 'OVERDUE_ITEMS',
      recipients: [primaryEmail],
      subject: subject,
      body: body
    });
  } else {
    // Send immediately
    try {
      MailApp.sendEmail({
        to: primaryEmail,
        subject: subject,
        body: body
      });
      
      // Log the notification
      logEmailNotification('OVERDUE_ITEMS', primaryEmail, 'Multiple');
    } catch (e) {
      console.error('Failed to send overdue items notification:', e);
    }
  }
}

/**
 * Sends a content assignment notification
 * @param {object} contentInfo Information about the content item
 * @param {string} assignedTo Team member assigned to the content
 */
function sendContentAssignmentNotification(contentInfo, assignedTo) {
  // Skip if notifications are disabled
  if (!NOTIFICATION_CONFIG.ENABLED || !NOTIFICATION_CONFIG.NOTIFICATION_TYPES.CONTENT_ASSIGNMENT) {
    return;
  }
  
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get recipient from team member email
  const teamEmail = getTeamMemberEmail(assignedTo);
  
  // If no recipient, exit
  if (!teamEmail) {
    console.log('No recipient found for content assignment notification');
    return;
  }
  
  // Replace template variables
  const subject = replaceTemplateVariables(
    NOTIFICATION_CONFIG.EMAIL_TEMPLATES.CONTENT_ASSIGNMENT.SUBJECT,
    contentInfo
  );
  
  const body = replaceTemplateVariables(
    NOTIFICATION_CONFIG.EMAIL_TEMPLATES.CONTENT_ASSIGNMENT.BODY,
    contentInfo,
    { spreadsheetUrl: ss.getUrl() }
  );
  
  // Send the email or add to queue
  if (NOTIFICATION_CONFIG.GROUP_NOTIFICATIONS) {
    // Add to notification queue for batching
    notificationQueue.push({
      type: 'CONTENT_ASSIGNMENT',
      recipients: [teamEmail],
      subject: subject,
      body: body,
      contentInfo: contentInfo
    });
  } else {
    // Send immediately
    try {
      MailApp.sendEmail({
        to: teamEmail,
        subject: subject,
        body: body
      });
      
      // Log the notification
      logEmailNotification('CONTENT_ASSIGNMENT', teamEmail, contentInfo.contentId);
    } catch (e) {
      console.error('Failed to send content assignment notification:', e);
    }
  }
}

/**
 * Sends a weekly summary notification
 */
function sendWeeklySummaryNotification() {
  // Skip if notifications are disabled
  if (!NOTIFICATION_CONFIG.ENABLED || !NOTIFICATION_CONFIG.NOTIFICATION_TYPES.WEEKLY_SUMMARY) {
    return;
  }
  
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get primary notification email
  const settingsSheet = ss.getSheetByName('Settings');
  const primaryEmail = settingsSheet.getRange(NOTIFICATION_CONFIG.SETTINGS_EMAIL_CELL).getValue();
  
  // If no recipient, exit
  if (!primaryEmail) {
    console.log('No recipient found for weekly summary notification');
    return;
  }
  
  // Get current week number
  const currentWeek = getCurrentWeekNumber();
  
  // Get content calendar data
  const calendarSheet = ss.getSheetByName('Content Calendar');
  const data = getContentCalendarData(calendarSheet);
  
  // Calculate summary statistics
  const summaryStats = calculateWeeklySummaryStats(data, currentWeek);
  
  // Build status breakdown text
  let statusBreakdown = '';
  const statusValues = ['Planned', 'Copywriting Complete', 'Creative Completed', 'Ready for Review', 'Schedule'];
  
  for (const status of statusValues) {
    const count = summaryStats.statusCounts[status] || 0;
    statusBreakdown += `- ${status}: ${count}\n`;
  }
  
  // Replace template variables
  const subject = NOTIFICATION_CONFIG.EMAIL_TEMPLATES.WEEKLY_SUMMARY.SUBJECT
    .replace('{weekNumber}', currentWeek);
  
  const body = NOTIFICATION_CONFIG.EMAIL_TEMPLATES.WEEKLY_SUMMARY.BODY
    .replace('{weekNumber}', currentWeek)
    .replace('{totalActive}', summaryStats.totalActive)
    .replace('{published}', summaryStats.published)
    .replace('{inProgress}', summaryStats.inProgress)
    .replace('{upcoming}', summaryStats.upcoming)
    .replace('{statusBreakdown}', statusBreakdown)
    .replace('{spreadsheetUrl}', ss.getUrl());
  
  // Send the email
  try {
    MailApp.sendEmail({
      to: primaryEmail,
      subject: subject,
      body: body
    });
    
    // Log the notification
    logEmailNotification('WEEKLY_SUMMARY', primaryEmail, 'Week ' + currentWeek);
  } catch (e) {
    console.error('Failed to send weekly summary notification:', e);
  }
}

/**
 * Processes the notification queue and sends batched emails
 */
function processNotificationQueue() {
  // Skip if queue is empty
  if (notificationQueue.length === 0) {
    return;
  }
  
  // Group notifications by recipient
  const emailsByRecipient = {};
  
  for (const notification of notificationQueue) {
    for (const recipient of notification.recipients) {
      if (!emailsByRecipient[recipient]) {
        emailsByRecipient[recipient] = [];
      }
      
      emailsByRecipient[recipient].push({
        type: notification.type,
        subject: notification.subject,
        body: notification.body,
        contentInfo: notification.contentInfo
      });
    }
  }
  
  // Send batched emails to each recipient
  for (const recipient in emailsByRecipient) {
    const notifications = emailsByRecipient[recipient];
    
    // Skip if no notifications
    if (notifications.length === 0) {
      continue;
    }
    
    // If only one notification, send it directly
    if (notifications.length === 1) {
      try {
        MailApp.sendEmail({
          to: recipient,
          subject: notifications[0].subject,
          body: notifications[0].body
        });
        
        // Log the notification
        const contentId = notifications[0].contentInfo ? notifications[0].contentInfo.contentId : 'Multiple';
        logEmailNotification(notifications[0].type, recipient, contentId);
      } catch (e) {
        console.error('Failed to send notification:', e);
      }
      continue;
    }
    
    // For multiple notifications, batch them
    const batchSubject = `[Content Calendar] ${notifications.length} Notifications`;
    let batchBody = `You have ${notifications.length} notifications from the content calendar:\n\n`;
    
    for (let i = 0; i < notifications.length; i++) {
      const notification = notifications[i];
      batchBody += `--- NOTIFICATION ${i + 1} ---\n`;
      batchBody += `Subject: ${notification.subject}\n\n`;
      batchBody += `${notification.body}\n\n`;
      batchBody += `--------------------\n\n`;
    }
    
    // Send the batched email
    try {
      MailApp.sendEmail({
        to: recipient,
        subject: batchSubject,
        body: batchBody
      });
      
      // Log the notification
      logEmailNotification('BATCH', recipient, 'Multiple');
    } catch (e) {
      console.error('Failed to send batched notification:', e);
    }
  }
  
  // Clear the notification queue
  notificationQueue = [];
}

/**
 * Schedules a daily check for upcoming deadlines and overdue items
 */
function scheduleDeadlineNotifications() {
  // Delete any existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkDeadlinesAndSendNotifications') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create a new daily trigger
  ScriptApp.newTrigger('checkDeadlinesAndSendNotifications')
    .timeBased()
    .everyDays(1)
    .atHour(9)  // 9 AM
    .create();
}

/**
 * Schedules a weekly summary notification
 */
function scheduleWeeklySummary() {
  // Delete any existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendWeeklySummaryNotification') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create a new weekly trigger (Monday at 8 AM)
  ScriptApp.newTrigger('sendWeeklySummaryNotification')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
}

/**
 * Checks for approaching deadlines and overdue items, and sends notifications
 */
function checkDeadlinesAndSendNotifications() {
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarSheet = ss.getSheetByName('Content Calendar');
  
  if (!calendarSheet) {
    console.log('Content Calendar sheet not found');
    return;
  }
  
  // Get calendar data
  const data = getContentCalendarData(calendarSheet);
  
  // Check for approaching deadlines (3 days away)
  const approachingDeadlines = findApproachingDeadlines(data, 3);
  
  // Send notifications for approaching deadlines
  for (const item of approachingDeadlines) {
    sendApproachingDeadlineNotification(item.contentInfo, item.daysRemaining);
  }
  
  // Check for overdue items
  const overdueItems = findOverdueItems(data);
  
  // Send notification for overdue items
  if (overdueItems.length > 0) {
    sendOverdueItemsNotification(overdueItems);
  }
  
  // Process the notification queue
  processNotificationQueue();
}

/**
 * Finds content items with approaching deadlines
 * @param {Array} data Calendar data
 * @param {number} daysThreshold Number of days to consider "approaching"
 * @return {Array} List of items with approaching deadlines
 */
function findApproachingDeadlines(data, daysThreshold) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const approaching = [];
  
  for (const row of data) {
    // Skip if not a valid date or already in "Schedule" status
    if (!(row.date instanceof Date) || row.status === 'Schedule') {
      continue;
    }
    
    // Calculate days until deadline
    const deadline = new Date(row.date);
    deadline.setHours(0, 0, 0, 0);
    
    const timeDiff = deadline.getTime() - today.getTime();
    const daysRemaining = Math.ceil(timeDiff / (1000 * 3600 * 24));
    
    // Check if approaching deadline
    if (daysRemaining > 0 && daysRemaining <= daysThreshold) {
      approaching.push({
        contentInfo: {
          contentId: row.id,
          contentTitle: row.content,
          status: row.status,
          channel: row.channel,
          publicationDate: formatDate(row.date),
          assignedTo: row.assignedTo
        },
        daysRemaining: daysRemaining
      });
    }
  }
  
  return approaching;
}

/**
 * Finds overdue content items
 * @param {Array} data Calendar data
 * @return {Array} List of overdue items
 */
function findOverdueItems(data) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const overdue = [];
  
  for (const row of data) {
    // Skip if not a valid date or already in "Schedule" status
    if (!(row.date instanceof Date) || row.status === 'Schedule') {
      continue;
    }
    
    // Check if overdue
    const deadline = new Date(row.date);
    deadline.setHours(0, 0, 0, 0);
    
    if (deadline < today) {
      overdue.push({
        contentId: row.id,
        contentTitle: row.content,
        status: row.status,
        formattedDate: formatDate(row.date),
        assignedTo: row.assignedTo
      });
    }
  }
  
  return overdue;
}

/**
 * Gets the content calendar data as an array of objects
 * @param {Sheet} sheet The content calendar sheet
 * @return {Array} Data as array of objects
 */
function getContentCalendarData(sheet) {
  // Get all data
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // Extract headers (assuming row 2 contains headers)
  const headers = values[1];
  
  // Find column indexes
  const idIndex = headers.indexOf('ID');
  const dateIndex = headers.indexOf('Date');
  const statusIndex = headers.indexOf('Status');
  const channelIndex = headers.indexOf('Channel');
  const contentIndex = headers.indexOf('Content/Idea');
  const assignedIndex = headers.indexOf('Assigned To');
  
  // Skip if any required column is missing
  if (idIndex === -1 || dateIndex === -1 || statusIndex === -1 || 
      channelIndex === -1 || contentIndex === -1) {
    console.log('Missing required columns in Content Calendar');
    return [];
  }
  
  // Convert data to array of objects
  const data = [];
  
  // Start from row 3 (index 2) to skip headers
  for (let i = 2; i < values.length; i++) {
    const row = values[i];
    
    // Skip if ID is empty (likely an empty row)
    if (!row[idIndex]) {
      continue;
    }
    
    data.push({
      id: row[idIndex],
      date: row[dateIndex],
      status: row[statusIndex],
      channel: row[channelIndex],
      content: row[contentIndex],
      assignedTo: assignedIndex !== -1 ? row[assignedIndex] : ''
    });
  }
  
  return data;
}

/**
 * Calculates weekly summary statistics
 * @param {Array} data Calendar data
 * @param {number} weekNumber Current week number
 * @return {object} Summary statistics
 */
function calculateWeeklySummaryStats(data, weekNumber) {
  // Initialize counters
  let totalActive = 0;
  let published = 0;
  let inProgress = 0;
  let upcoming = 0;
  const statusCounts = {};
  
  // Get week start and end dates
  const weekDates = getWeekDates(weekNumber);
  
  // Calculate stats
  for (const row of data) {
    // Skip if not a valid row
    if (!row.id) {
      continue;
    }
    
    // Count total active items
    totalActive++;
    
    // Count published items
    if (row.status === 'Schedule') {
      published++;
    } else {
      inProgress++;
    }
    
    // Count by status
    if (!statusCounts[row.status]) {
      statusCounts[row.status] = 0;
    }
    statusCounts[row.status]++;
    
    // Count upcoming items (in current week)
    if (row.date instanceof Date && 
        row.date >= weekDates.startDate && 
        row.date <= weekDates.endDate) {
      upcoming++;
    }
  }
  
  return {
    totalActive: totalActive,
    published: published,
    inProgress: inProgress,
    upcoming: upcoming,
    statusCounts: statusCounts
  };
}

/**
 * Gets the start and end dates for a given week number
 * @param {number} weekNumber Week number
 * @return {object} Object with startDate and endDate
 */
function getWeekDates(weekNumber) {
  const year = new Date().getFullYear();
  
  // January 4th is always in week 1 in ISO weeks
  const jan4 = new Date(year, 0, 4);
  
  // Find the Monday of week 1
  const week1Start = new Date(jan4);
  week1Start.setDate(jan4.getDate() - jan4.getDay() + 1);
  if (jan4.getDay() === 0) week1Start.setDate(week1Start.getDate() - 7);
  
  // Calculate start date of the specified week
  const startDate = new Date(week1Start);
  startDate.setDate(week1Start.getDate() + (weekNumber - 1) * 7);
  
  // Calculate end date (start date + 6 days)
  const endDate = new Date(startDate);
  endDate.setDate(startDate.getDate() + 6);
  
  return { startDate, endDate };
}

/**
 * Gets the current ISO week number
 * @return {number} Current week number
 */
function getCurrentWeekNumber() {
  return parseInt(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'w'));
}

/**
 * Gets the email address for a team member
 * @param {string} teamMember Team member name
 * @return {string} Email address or null if not found
 */
function getTeamMemberEmail(teamMember) {
  // Skip if no team member name
  if (!teamMember) {
    return null;
  }
  
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if Team Members sheet exists
  let teamSheet = ss.getSheetByName('Team Members');
  
  // If not, try to find in Settings sheet
  if (!teamSheet) {
    const settingsSheet = ss.getSheetByName('Settings');
    if (!settingsSheet) {
      return null;
    }
    
    // Return the primary notification email as fallback
    return settingsSheet.getRange(NOTIFICATION_CONFIG.SETTINGS_EMAIL_CELL).getValue();
  }
  
  // Search for team member in Team Members sheet
  const teamData = teamSheet.getDataRange().getValues();
  
  // Skip header row
  for (let i = 1; i < teamData.length; i++) {
    if (teamData[i][0] === teamMember) {
      return teamData[i][NOTIFICATION_CONFIG.TEAM_EMAIL_COLUMN - 1];
    }
  }
  
  return null;
}

/**
 * Logs an email notification to the Email Logs sheet
 * @param {string} type Notification type
 * @param {string} recipient Email recipient(s)
 * @param {string} contentId Content ID or descriptor
 */
function logEmailNotification(type, recipient, contentId) {
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if Email Logs sheet exists, create if not
  let logsSheet = ss.getSheetByName(NOTIFICATION_CONFIG.EMAIL_LOGS_SHEET);
  if (!logsSheet) {
    logsSheet = ss.insertSheet(NOTIFICATION_CONFIG.EMAIL_LOGS_SHEET);
    
    // Set up headers
    logsSheet.getRange(1, 1, 1, 5).setValues([
      ['Timestamp', 'Type', 'Recipient', 'Content ID', 'Status']
    ]);
    
    // Format headers
    logsSheet.getRange(1, 1, 1, 5)
      .setBackground('#4285F4')
      .setFontColor('white')
      .setFontWeight('bold');
    
    // Set column widths
    logsSheet.setColumnWidth(1, 180); // Timestamp
    logsSheet.setColumnWidth(2, 150); // Type
    logsSheet.setColumnWidth(3, 250); // Recipient
    logsSheet.setColumnWidth(4, 120); // Content ID
    logsSheet.setColumnWidth(5, 100); // Status
  }
  
  // Add log entry
  const timestamp = new Date();
  const lastRow = Math.max(logsSheet.getLastRow(), 1);
  
  logsSheet.getRange(lastRow + 1, 1, 1, 5).setValues([
    [timestamp, type, recipient, contentId, 'Sent']
  ]);
  
  // Format timestamp
  logsSheet.getRange(lastRow + 1, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');
}

/**
 * Replaces template variables in a string
 * @param {string} template Template string with {variable} placeholders
 * @param {object} contentInfo Content information object
 * @param {object} additionalVars Additional variables to replace
 * @return {string} Formatted string with replaced variables
 */
function replaceTemplateVariables(template, contentInfo, additionalVars = {}) {
  let result = template;
  
  // Replace content info variables
  if (contentInfo) {
    for (const key in contentInfo) {
      result = result.replace(new RegExp(`\\{${key}\\}`, 'g'), contentInfo[key] || '');
    }
  }
  
  // Replace additional variables
  for (const key in additionalVars) {
    result = result.replace(new RegExp(`\\{${key}\\}`, 'g'), additionalVars[key] || '');
  }
  
  return result;
}

/**
 * Formats a date as a string
 * @param {Date} date Date to format
 * @return {string} Formatted date
 */
function formatDate(date) {
  if (!(date instanceof Date)) {
    return '';
  }
  
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Sets up the notification system
 */
function setupNotificationSystem() {
  // Schedule deadline notifications
  scheduleDeadlineNotifications();
  
  // Schedule weekly summary
  scheduleWeeklySummary();
  
  // Create Email Logs sheet if it doesn't exist
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logsSheet = ss.getSheetByName(NOTIFICATION_CONFIG.EMAIL_LOGS_SHEET);
  
  if (!logsSheet) {
    logsSheet = ss.insertSheet(NOTIFICATION_CONFIG.EMAIL_LOGS_SHEET);
    
    // Set up headers
    logsSheet.getRange(1, 1, 1, 5).setValues([
      ['Timestamp', 'Type', 'Recipient', 'Content ID', 'Status']
    ]);
    
    // Format headers
    logsSheet.getRange(1, 1, 1, 5)
      .setBackground('#4285F4')
      .setFontColor('white')
      .setFontWeight('bold');
    
    // Set column widths
    logsSheet.setColumnWidth(1, 180); // Timestamp
    logsSheet.setColumnWidth(2, 150); // Type
    logsSheet.setColumnWidth(3, 250); // Recipient
    logsSheet.setColumnWidth(4, 120); // Content ID
    logsSheet.setColumnWidth(5, 100); // Status
  }
  
  // Check if Team Members sheet exists, create if not
  let teamSheet = ss.getSheetByName('Team Members');
  
  if (!teamSheet) {
    teamSheet = ss.insertSheet('Team Members');
    
    // Set up headers
    teamSheet.getRange(1, 1, 1, 3).setValues([
      ['Name', 'Email', 'Role']
    ]);
    
    // Format headers
    teamSheet.getRange(1, 1, 1, 3)
      .setBackground('#4285F4')
      .setFontColor('white')
      .setFontWeight('bold');
    
    // Set column widths
    teamSheet.setColumnWidth(1, 150); // Name
    teamSheet.setColumnWidth(2, 250); // Email
    teamSheet.setColumnWidth(3, 150); // Role
    
    // Add sample data from Lists sheet
    const listsSheet = ss.getSheetByName('Lists');
    if (listsSheet) {
      // Get team member names from Lists sheet (column E)
      const teamMemberRange = listsSheet.getRange('E2:E' + listsSheet.getLastRow());
      const teamMembers = teamMemberRange.getValues();
      
      // Add each team member to Team Members sheet
      let rowIndex = 2;
      for (let i = 0; i < teamMembers.length; i++) {
        if (teamMembers[i][0]) {
          teamSheet.getRange(rowIndex, 1, 1, 3).setValues([
            [teamMembers[i][0], '', '']
          ]);
          rowIndex++;
        }
      }
    }
  }
  
  // Show confirmation
  SpreadsheetApp.getUi().alert('Notification system set up successfully. Please add team member email addresses in the Team Members sheet.');
}

/**
 * Tests the notification system by sending a test email
 */
function testNotificationSystem() {
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the settings sheet
  const settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    SpreadsheetApp.getUi().alert('Settings sheet not found.');
    return;
  }
  
  // Get notification email
  const email = settingsSheet.getRange(NOTIFICATION_CONFIG.SETTINGS_EMAIL_CELL).getValue();
  if (!email) {
    SpreadsheetApp.getUi().alert('No notification email found in Settings sheet (cell B7).');
    return;
  }
  
  // Send test email
  try {
    MailApp.sendEmail({
      to: email,
      subject: '[Content Calendar] Test Notification',
      body: 'This is a test notification from your Content Calendar.\n\n' +
            'If you received this email, your notification system is working correctly.\n\n' +
            'Spreadsheet: ' + ss.getName() + '\n' +
            'Time: ' + new Date().toString()
    });
    
    // Log the notification
    logEmailNotification('TEST', email, 'Test');
    
    SpreadsheetApp.getUi().alert('Test email sent successfully to ' + email);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Failed to send test email: ' + e.toString());
  }
}

/**
 * Creates a menu for notification settings
 */
function createNotificationMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Notifications')
    .addItem('Set Up Notification System', 'setupNotificationSystem')
    .addItem('Send Test Notification', 'testNotificationSystem')
    .addItem('Schedule Deadline Notifications', 'scheduleDeadlineNotifications')
    .addItem('Schedule Weekly Summary', 'scheduleWeeklySummary')
    .addItem('Send Weekly Summary Now', 'sendWeeklySummaryNotification')
    .addToUi();
}