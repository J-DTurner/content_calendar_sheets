/**
 * Workflow Automation for Social Media Content Calendar
 * 
 * This script implements automated workflows for content creation:
 * - Content template generation
 * - Status change automation
 * - Due date calculations
 * - Team assignments
 * - Reminders and notifications
 */

// Workflow automation configuration
const WORKFLOW_CONFIG = {
  STATUS_COLUMN: 4,         // Column D
  CHANNEL_COLUMN: 5,        // Column E
  CONTENT_COLUMN: 6,        // Column F
  FORMAT_COLUMN: 8,         // Column H
  ASSIGNED_COLUMN: 10,      // Column J
  DATE_COLUMN: 2,           // Column B
  TEMPLATE_FOLDER_SETTING: 'B10', // Cell containing template folder ID in Settings
  ENABLE_NOTIFICATIONS: true
};

/**
 * Generates content templates based on channel and format
 */
function generateContentTemplate() {
  // Get the active sheet and selected cell
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = SpreadsheetApp.getActiveRange();
  const row = range.getRow();
  
  // Skip if not in the content calendar or in header rows
  if (sheet.getName() !== 'Content Calendar' || row < 3) {
    SpreadsheetApp.getUi().alert('Please select a content item in the Content Calendar sheet.');
    return;
  }
  
  // Get content details
  const channel = sheet.getRange(row, WORKFLOW_CONFIG.CHANNEL_COLUMN).getValue();
  const format = sheet.getRange(row, WORKFLOW_CONFIG.FORMAT_COLUMN).getValue();
  const content = sheet.getRange(row, WORKFLOW_CONFIG.CONTENT_COLUMN).getValue();
  
  // Check if we already have content
  if (content && content.trim() !== '') {
    // Ask for confirmation before overwriting
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Content Already Exists',
      'This cell already has content. Do you want to replace it with a template?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
  }
  
  // Generate template based on channel and format
  let template = '';
  
  if (!channel && !format) {
    // No channel or format specified, show selection dialog
    template = showTemplateSelectionDialog();
    if (!template) return; // User canceled
  } else {
    template = getTemplateForChannelAndFormat(channel, format);
  }
  
  // Apply the template
  sheet.getRange(row, WORKFLOW_CONFIG.CONTENT_COLUMN).setValue(template);
  
  // Update status if it's empty
  const statusCell = sheet.getRange(row, WORKFLOW_CONFIG.STATUS_COLUMN);
  if (statusCell.getValue() === '') {
    statusCell.setValue('Planned');
  }
  
  // Update the last modified timestamp if that column exists
  try {
    sheet.getRange(row, 13).setValue(new Date()); // Column M (Last Modified)
  } catch (e) {
    // Column might not exist, ignore
  }
}

/**
 * Shows a dialog for selecting a template
 */
function showTemplateSelectionDialog() {
  const ui = SpreadsheetApp.getUi();
  
  // Get available channels from the Lists sheet
  const listsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Lists');
  const channels = listsSheet.getRange('B2:B' + listsSheet.getLastRow()).getValues()
    .filter(row => row[0] !== '')
    .map(row => row[0]);
    
  // Get available formats from the Lists sheet
  const formats = listsSheet.getRange('D2:D' + listsSheet.getLastRow()).getValues()
    .filter(row => row[0] !== '')
    .map(row => row[0]);
  
  // Show channel selection
  let channelPrompt = 'Select a channel:\n\n';
  channels.forEach((channel, index) => {
    channelPrompt += `${index + 1}. ${channel}\n`;
  });
  
  const channelResponse = ui.prompt(
    'Template Selection',
    channelPrompt,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (channelResponse.getSelectedButton() !== ui.Button.OK) {
    return null;
  }
  
  // Parse channel selection
  let selectedChannel;
  try {
    const channelIndex = parseInt(channelResponse.getResponseText().trim()) - 1;
    if (isNaN(channelIndex) || channelIndex < 0 || channelIndex >= channels.length) {
      throw new Error('Invalid selection');
    }
    selectedChannel = channels[channelIndex];
  } catch (error) {
    ui.alert('Invalid selection. Please try again.');
    return null;
  }
  
  // Show format selection
  let formatPrompt = 'Select a format:\n\n';
  formats.forEach((format, index) => {
    formatPrompt += `${index + 1}. ${format}\n`;
  });
  
  const formatResponse = ui.prompt(
    'Template Selection',
    formatPrompt,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (formatResponse.getSelectedButton() !== ui.Button.OK) {
    return null;
  }
  
  // Parse format selection
  let selectedFormat;
  try {
    const formatIndex = parseInt(formatResponse.getResponseText().trim()) - 1;
    if (isNaN(formatIndex) || formatIndex < 0 || formatIndex >= formats.length) {
      throw new Error('Invalid selection');
    }
    selectedFormat = formats[formatIndex];
  } catch (error) {
    ui.alert('Invalid selection. Please try again.');
    return null;
  }
  
  // Get template for selected channel and format
  return getTemplateForChannelAndFormat(selectedChannel, selectedFormat);
}

/**
 * Gets a template for a specific channel and format
 */
function getTemplateForChannelAndFormat(channel, format) {
  // Default templates by channel and format
  const templates = {
    'Twitter': {
      'Text Post': 'Main point: \nKey message (under 280 chars): \nHashtags: \nMention: ',
      'Image': 'Main point: \nImage description: \nCaption (under 280 chars): \nHashtags: ',
      'Video': 'Video topic: \nKey message: \nCaption: \nHashtags: \nVideo length: ',
      'Thread': 'Main topic: \n\nTweet 1: \n\nTweet 2: \n\nTweet 3: \n\nCall to action: ',
      'Poll': 'Poll question: \n\nOption 1: \nOption 2: \nOption 3: \nOption 4: \n\nDuration: '
    },
    'YouTube': {
      'Video': 'Video title: \nDescription: \n\nIntro (0:00-0:30): \nMain points: \n- Point 1 (0:30-2:00): \n- Point 2 (2:00-4:00): \n- Point 3 (4:00-6:00): \nConclusion (6:00-7:00): \n\nTags: \nCategory: ',
      'Short': 'Short title: \nHook (first 3 seconds): \nMain concept: \nCall to action: \nCaption: \nHashtags: ',
      'Live': 'Stream title: \nScheduled date/time: \nDescription: \nTopics to cover: \n- Topic 1: \n- Topic 2: \n- Topic 3: \nQ&A prompts: '
    },
    'Telegram': {
      'Text Post': 'Title: \n\nMain content: \n\nKey points: \n- \n- \n- \n\nCall to action: ',
      'Image': 'Title: \n\nImage description: \n\nCaption: \n\nAdditional notes: ',
      'Video': 'Title: \n\nVideo description: \n\nCaption: \n\nKey timestamps: \n- 0:00 - \n- 0:00 - ',
      'Poll': 'Poll question: \n\nOptions: \n- \n- \n- \n- \n\nContext/introduction: '
    }
  };
  
  // Check if we have a template for this combination
  if (templates[channel] && templates[channel][format]) {
    return templates[channel][format];
  }
  
  // Check if we have a template for just the channel
  if (templates[channel] && templates[channel]['Default']) {
    return templates[channel]['Default'];
  }
  
  // Use default format templates if available
  const defaultTemplates = {
    'Text Post': 'Title: \nMain message: \nKey points: \n- \n- \n- \nCall to action: ',
    'Image': 'Caption: \nImage description: \nKey message: \nHashtags: ',
    'Video': 'Title: \nMain topic: \nKey points: \n- \n- \n- \nCall to action: ',
    'Poll': 'Question: \nOptions: \n- \n- \n- \n- \nContext: ',
    'Default': 'Content title: \nMain message: \nKey points: \n- \n- \n- \nCall to action: '
  };
  
  if (defaultTemplates[format]) {
    return defaultTemplates[format];
  }
  
  // Return a very generic template
  return 'Title: \nMain content: \nKey points: \n- \n- \n- \nCall to action: ';
}

/**
 * Automatically assigns team members based on content type
 */
function autoAssignTeamMember() {
  // Get the active sheet and selected cell
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = SpreadsheetApp.getActiveRange();
  const row = range.getRow();
  
  // Skip if not in the content calendar or in header rows
  if (sheet.getName() !== 'Content Calendar' || row < 3) {
    SpreadsheetApp.getUi().alert('Please select a content item in the Content Calendar sheet.');
    return;
  }
  
  // Get content details
  const channel = sheet.getRange(row, WORKFLOW_CONFIG.CHANNEL_COLUMN).getValue();
  const format = sheet.getRange(row, WORKFLOW_CONFIG.FORMAT_COLUMN).getValue();
  const status = sheet.getRange(row, WORKFLOW_CONFIG.STATUS_COLUMN).getValue();
  
  // Get settings
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  
  // Define team assignments based on channel, format, and status
  let assignedMember = '';
  
  // Determine team member by status first
  if (status === 'Planned') {
    // Default content strategist/writer
    assignedMember = settingsSheet.getRange('B4').getValue(); // Content Manager
  } else if (status === 'Copywriting Complete') {
    // Designer for creative work
    assignedMember = settingsSheet.getRange('B6').getValue(); // Designer
  } else if (status === 'Creative Completed') {
    // Video editor if video content
    if (format === 'Video' || format === 'Short') {
      assignedMember = settingsSheet.getRange('B11').getValue(); // Video Editor
    } else {
      // Otherwise, go to reviewer
      assignedMember = settingsSheet.getRange('B5').getValue(); // Reviewer
    }
  } else if (status === 'Ready for Review') {
    // Social media manager for scheduling
    assignedMember = settingsSheet.getRange('B9').getValue(); // Social Media Manager
  }
  
  // Override based on channel+format if no assignment yet
  if (!assignedMember && channel && format) {
    if (channel === 'YouTube' && (format === 'Video' || format === 'Short' || format === 'Live')) {
      assignedMember = settingsSheet.getRange('B11').getValue(); // Video Editor
    } else if (format === 'Image' || format === 'Infographic') {
      assignedMember = settingsSheet.getRange('B6').getValue(); // Designer
    } else {
      assignedMember = settingsSheet.getRange('B4').getValue(); // Content Manager
    }
  }
  
  // If still no assignment, use default
  if (!assignedMember) {
    assignedMember = settingsSheet.getRange('B4').getValue(); // Content Manager as default
  }
  
  // Set the assigned team member
  sheet.getRange(row, WORKFLOW_CONFIG.ASSIGNED_COLUMN).setValue(assignedMember);
  
  // Update the last modified timestamp if that column exists
  try {
    sheet.getRange(row, 13).setValue(new Date()); // Column M (Last Modified)
  } catch (e) {
    // Column might not exist, ignore
  }
  
  // Show confirmation
  SpreadsheetApp.getActiveSpreadsheet().toast(`Assigned to: ${assignedMember}`, 'Team Member Assigned');
}

/**
 * Calculates and sets due dates based on publication date
 */
function calculateDueDates() {
  // Get the active sheet and selected rows
  const sheet = SpreadsheetApp.getActiveSheet();
  const selectedRange = SpreadsheetApp.getActiveRange();
  const startRow = selectedRange.getRow();
  const numRows = selectedRange.getNumRows();
  
  // Skip if not in the content calendar
  if (sheet.getName() !== 'Content Calendar') {
    SpreadsheetApp.getUi().alert('Please select content items in the Content Calendar sheet.');
    return;
  }
  
  // Skip header rows
  if (startRow < 3) {
    SpreadsheetApp.getUi().alert('Please select only content rows (not headers).');
    return;
  }
  
  // Check if we have a due date column
  const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  let dueDateColumn = -1;
  
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === 'Due Date') {
      dueDateColumn = i + 1;
      break;
    }
  }
  
  // If no Due Date column exists, ask if we should add one
  if (dueDateColumn === -1) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Due Date Column',
      'No "Due Date" column found. Would you like to add one?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      // Find the last column and add a new one
      const lastColumn = sheet.getLastColumn();
      dueDateColumn = lastColumn + 1;
      
      // Add header
      sheet.getRange(2, dueDateColumn).setValue('Due Date');
      
      // Format the column
      sheet.getRange(3, dueDateColumn, sheet.getLastRow() - 2).setNumberFormat('yyyy-mm-dd');
    } else {
      return;
    }
  }
  
  // Process selected rows
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    
    // Get publication date
    const pubDate = sheet.getRange(row, WORKFLOW_CONFIG.DATE_COLUMN).getValue();
    
    // Skip if no publication date
    if (!pubDate || !(pubDate instanceof Date)) {
      continue;
    }
    
    // Get current status
    const status = sheet.getRange(row, WORKFLOW_CONFIG.STATUS_COLUMN).getValue();
    
    // Calculate due date based on status
    let dueDate = calculateDueDateForStatus(pubDate, status);
    
    // Set the due date
    sheet.getRange(row, dueDateColumn).setValue(dueDate);
  }
  
  // Show confirmation
  SpreadsheetApp.getActiveSpreadsheet().toast('Due dates have been calculated and set.', 'Due Dates Updated');
}

/**
 * Calculates the due date based on publication date and status
 */
function calculateDueDateForStatus(pubDate, status) {
  // Clone the publication date
  const dueDate = new Date(pubDate);
  
  // Default lead times (in days before publication)
  const leadTimes = {
    'Planned': 14, // Two weeks before
    'Copywriting Complete': 10, // 10 days before
    'Creative Completed': 7, // One week before
    'Ready for Review': 3, // Three days before
    'Schedule': 1 // One day before
  };
  
  // Get the lead time for this status
  let leadTime = leadTimes[status] || 14; // Default to two weeks if status not found
  
  // Subtract the lead time from the publication date
  dueDate.setDate(dueDate.getDate() - leadTime);
  
  return dueDate;
}

/**
 * Creates a Google Calendar event for content creation workflow
 */
function createWorkflowEvent() {
  // Get the active sheet and selected cell
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = SpreadsheetApp.getActiveRange();
  const row = range.getRow();
  
  // Skip if not in the content calendar or in header rows
  if (sheet.getName() !== 'Content Calendar' || row < 3) {
    SpreadsheetApp.getUi().alert('Please select a content item in the Content Calendar sheet.');
    return;
  }
  
  // Get content details
  const contentId = sheet.getRange(row, 1).getValue();
  const pubDate = sheet.getRange(row, WORKFLOW_CONFIG.DATE_COLUMN).getValue();
  const channel = sheet.getRange(row, WORKFLOW_CONFIG.CHANNEL_COLUMN).getValue();
  const content = sheet.getRange(row, WORKFLOW_CONFIG.CONTENT_COLUMN).getValue();
  const assignedTo = sheet.getRange(row, WORKFLOW_CONFIG.ASSIGNED_COLUMN).getValue();
  
  // Skip if no publication date
  if (!pubDate || !(pubDate instanceof Date)) {
    SpreadsheetApp.getUi().alert('Please set a publication date first.');
    return;
  }
  
  // Get the calendar ID from settings
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  const calendarId = settingsSheet.getRange('B8').getValue();
  
  // Get calendar
  let calendar;
  try {
    if (calendarId) {
      calendar = CalendarApp.getCalendarById(calendarId);
    } else {
      calendar = CalendarApp.getDefaultCalendar();
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('Could not access the calendar. Please check the calendar ID in Settings.');
    return;
  }
  
  // Check if we have a due date column
  const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  let dueDateColumn = -1;
  
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === 'Due Date') {
      dueDateColumn = i + 1;
      break;
    }
  }
  
  // Get due date if available
  let dueDate;
  if (dueDateColumn !== -1) {
    dueDate = sheet.getRange(row, dueDateColumn).getValue();
  }
  
  // If no due date, calculate it
  if (!dueDate || !(dueDate instanceof Date)) {
    const status = sheet.getRange(row, WORKFLOW_CONFIG.STATUS_COLUMN).getValue();
    dueDate = calculateDueDateForStatus(pubDate, status);
  }
  
  // Create a title for the event
  const title = `[${channel}] Content: ${contentId}`;
  
  // Create event details
  const description = `Content: ${content ? content.substring(0, 100) : 'N/A'}\n\n` +
    `Channel: ${channel}\n` +
    `Assigned to: ${assignedTo}\n` +
    `Publication date: ${Utilities.formatDate(pubDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')}\n\n` +
    `Link to content calendar: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
  
  // Create a 1-hour event on the due date
  const startTime = new Date(dueDate);
  startTime.setHours(10, 0, 0, 0); // Set to 10:00 AM
  
  const endTime = new Date(startTime);
  endTime.setHours(endTime.getHours() + 1); // 1 hour duration
  
  // Create the event
  try {
    const event = calendar.createEvent(title, startTime, endTime, {
      description: description,
      location: 'Content Calendar'
    });
    
    // Store event ID in the spreadsheet if we have a column for it
    let eventIdColumn = -1;
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === 'Event ID') {
        eventIdColumn = i + 1;
        break;
      }
    }
    
    if (eventIdColumn !== -1) {
      sheet.getRange(row, eventIdColumn).setValue(event.getId());
    }
    
    // Show confirmation with the event link
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Event Created',
      `Calendar event created for ${Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')}.\n\n` +
      `Event: ${title}`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error creating event: ' + error.toString());
  }
}

/**
 * Runs automation for selected content item
 */
function runContentAutomation() {
  // Get the active sheet and selected cell
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = SpreadsheetApp.getActiveRange();
  const row = range.getRow();
  
  // Skip if not in the content calendar or in header rows
  if (sheet.getName() !== 'Content Calendar' || row < 3) {
    SpreadsheetApp.getUi().alert('Please select a content item in the Content Calendar sheet.');
    return;
  }
  
  // Show options dialog
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Content Automation',
    'Select an automation to run:\n\n' +
    '1. Generate Content Template\n' +
    '2. Auto-Assign Team Member\n' +
    '3. Calculate Due Date\n' +
    '4. Create Calendar Event\n' +
    '5. Advance to Next Status\n' +
    '6. Run All Automations',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  // Parse the selection
  const selection = response.getResponseText().trim();
  
  switch (selection) {
    case '1':
      generateContentTemplate();
      break;
    case '2':
      autoAssignTeamMember();
      break;
    case '3':
      calculateDueDates();
      break;
    case '4':
      createWorkflowEvent();
      break;
    case '5':
      // Call the advanceStatus function from the status tracking script
      if (typeof advanceStatus === 'function') {
        advanceStatus();
      } else {
        ui.alert('Status advancement function not available. Make sure the status tracking script is loaded.');
      }
      break;
    case '6':
      // Run all automations
      generateContentTemplate();
      autoAssignTeamMember();
      calculateDueDates();
      createWorkflowEvent();
      // Advance status if function is available
      if (typeof advanceStatus === 'function') {
        advanceStatus();
      }
      break;
    default:
      ui.alert('Invalid selection. Please try again.');
      break;
  }
}

/**
 * Updates the custom menu with workflow automation commands
 */
function updateWorkflowMenu() {
  // Check if the menu already exists
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.createMenu('Workflow')
      .addItem('Run Content Automation...', 'runContentAutomation')
      .addSeparator()
      .addItem('Generate Content Template', 'generateContentTemplate')
      .addItem('Auto-Assign Team Member', 'autoAssignTeamMember')
      .addItem('Calculate Due Dates', 'calculateDueDates')
      .addItem('Create Workflow Calendar Event', 'createWorkflowEvent')
      .addToUi();
  } catch (e) {
    // Menu might already exist, no need to handle error
  }
}