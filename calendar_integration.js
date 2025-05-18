/**
 * Calendar Integration for Social Media Content Calendar
 * 
 * This script provides Google Calendar integration features:
 * - Creating calendar events for content due dates
 * - Syncing content changes with calendar events
 * - Importing calendar events as content items
 * - Event notifications and reminders
 */

// Calendar integration configuration
const CALENDAR_CONFIG = {
  SETTINGS_CALENDAR_ID_CELL: 'B8', // Cell containing calendar ID in Settings sheet
  DATE_COLUMN: 2,                  // Column B in Content Calendar
  STATUS_COLUMN: 4,                // Column D in Content Calendar
  CHANNEL_COLUMN: 5,               // Column E in Content Calendar
  CONTENT_COLUMN: 6,               // Column F in Content Calendar
  ID_COLUMN: 1,                    // Column A in Content Calendar
  EVENT_ID_COLUMN: 15,             // Column O in Content Calendar (for storing event IDs)
  DEFAULT_EVENT_DURATION: 30,      // Default event duration in minutes
  AUTO_CREATE_EVENTS: true,        // Automatically create events when status changes to "Schedule"
  UPDATE_EVENTS_ON_CHANGE: true,   // Update calendar events when content changes
  EVENT_COLOR_MAPPING: {           // Calendar event color IDs by channel
    'Twitter': '1',   // Blue
    'YouTube': '11',  // Red
    'Telegram': '9'   // Dark Blue
  }
};

/**
 * Creates a calendar event for a content item
 * @param {number} row Row number in the content calendar
 * @param {boolean} showConfirmation Whether to show a confirmation dialog
 * @return {string} Event ID if successful, null otherwise
 */
function createCalendarEvent(row, showConfirmation = true) {
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarSheet = ss.getSheetByName('Content Calendar');
  
  if (!calendarSheet) {
    if (showConfirmation) {
      SpreadsheetApp.getUi().alert('Content Calendar sheet not found.');
    }
    return null;
  }
  
  // Get content information
  const contentId = calendarSheet.getRange(row, CALENDAR_CONFIG.ID_COLUMN).getValue();
  const pubDate = calendarSheet.getRange(row, CALENDAR_CONFIG.DATE_COLUMN).getValue();
  const status = calendarSheet.getRange(row, CALENDAR_CONFIG.STATUS_COLUMN).getValue();
  const channel = calendarSheet.getRange(row, CALENDAR_CONFIG.CHANNEL_COLUMN).getValue();
  const content = calendarSheet.getRange(row, CALENDAR_CONFIG.CONTENT_COLUMN).getValue();
  
  // Validate data
  if (!contentId || !pubDate || !(pubDate instanceof Date)) {
    if (showConfirmation) {
      SpreadsheetApp.getUi().alert('Please ensure content ID and publication date are set.');
    }
    return null;
  }
  
  // Get the Google Calendar
  const calendar = getContentCalendar();
  if (!calendar) {
    if (showConfirmation) {
      SpreadsheetApp.getUi().alert('Could not access Google Calendar. Please check calendar ID in Settings sheet.');
    }
    return null;
  }
  
  // Check if an event ID already exists for this content
  let existingEventId = '';
  try {
    existingEventId = calendarSheet.getRange(row, CALENDAR_CONFIG.EVENT_ID_COLUMN).getValue();
  } catch (e) {
    // Column might not exist, will be created later
  }
  
  // If event ID exists, update the event instead of creating a new one
  if (existingEventId) {
    return updateCalendarEvent(row, existingEventId, showConfirmation);
  }
  
  // Set the event start time (default to noon on the publication date)
  const startTime = new Date(pubDate);
  startTime.setHours(12, 0, 0, 0);
  
  // Set the event end time
  const endTime = new Date(startTime);
  endTime.setMinutes(startTime.getMinutes() + CALENDAR_CONFIG.DEFAULT_EVENT_DURATION);
  
  // Create event title based on channel and content
  let title = `[${channel}] `;
  if (content) {
    // Truncate content if too long
    title += content.length > 50 ? content.substring(0, 47) + '...' : content;
  } else {
    title += contentId;
  }
  
  // Create event description
  const description = 
    `Content ID: ${contentId}\n` +
    `Channel: ${channel}\n` +
    `Status: ${status}\n\n` +
    `Content/Idea: ${content || 'N/A'}\n\n` +
    `View in Content Calendar: ${ss.getUrl()}`;
  
  // Create the event
  try {
    const event = calendar.createEvent(title, startTime, endTime, {
      description: description,
      location: 'Content Calendar'
    });
    
    // Set event color based on channel
    if (channel && CALENDAR_CONFIG.EVENT_COLOR_MAPPING[channel]) {
      event.setColor(CALENDAR_CONFIG.EVENT_COLOR_MAPPING[channel]);
    }
    
    // Store the event ID in the content calendar
    ensureEventIdColumnExists(calendarSheet);
    calendarSheet.getRange(row, CALENDAR_CONFIG.EVENT_ID_COLUMN).setValue(event.getId());
    
    if (showConfirmation) {
      SpreadsheetApp.getUi().alert('Calendar event created successfully!');
    }
    
    return event.getId();
  } catch (e) {
    if (showConfirmation) {
      SpreadsheetApp.getUi().alert('Failed to create calendar event: ' + e.toString());
    }
    console.error('Failed to create calendar event:', e);
    return null;
  }
}

/**
 * Updates an existing calendar event for a content item
 * @param {number} row Row number in the content calendar
 * @param {string} eventId ID of the existing calendar event
 * @param {boolean} showConfirmation Whether to show a confirmation dialog
 * @return {string} Event ID if successful, null otherwise
 */
function updateCalendarEvent(row, eventId, showConfirmation = true) {
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarSheet = ss.getSheetByName('Content Calendar');
  
  if (!calendarSheet || !eventId) {
    return null;
  }
  
  // Get content information
  const contentId = calendarSheet.getRange(row, CALENDAR_CONFIG.ID_COLUMN).getValue();
  const pubDate = calendarSheet.getRange(row, CALENDAR_CONFIG.DATE_COLUMN).getValue();
  const status = calendarSheet.getRange(row, CALENDAR_CONFIG.STATUS_COLUMN).getValue();
  const channel = calendarSheet.getRange(row, CALENDAR_CONFIG.CHANNEL_COLUMN).getValue();
  const content = calendarSheet.getRange(row, CALENDAR_CONFIG.CONTENT_COLUMN).getValue();
  
  // Validate data
  if (!contentId || !pubDate || !(pubDate instanceof Date)) {
    return null;
  }
  
  // Get the Google Calendar
  const calendar = getContentCalendar();
  if (!calendar) {
    return null;
  }
  
  // Get the existing event
  try {
    const event = calendar.getEventById(eventId);
    
    // If event doesn't exist, create a new one
    if (!event) {
      return createCalendarEvent(row, showConfirmation);
    }
    
    // Set the event start time (default to noon on the publication date)
    const startTime = new Date(pubDate);
    startTime.setHours(12, 0, 0, 0);
    
    // Set the event end time
    const endTime = new Date(startTime);
    endTime.setMinutes(startTime.getMinutes() + CALENDAR_CONFIG.DEFAULT_EVENT_DURATION);
    
    // Create event title based on channel and content
    let title = `[${channel}] `;
    if (content) {
      // Truncate content if too long
      title += content.length > 50 ? content.substring(0, 47) + '...' : content;
    } else {
      title += contentId;
    }
    
    // Create event description
    const description = 
      `Content ID: ${contentId}\n` +
      `Channel: ${channel}\n` +
      `Status: ${status}\n\n` +
      `Content/Idea: ${content || 'N/A'}\n\n` +
      `View in Content Calendar: ${ss.getUrl()}`;
    
    // Update the event
    event.setTitle(title);
    event.setDescription(description);
    event.setTime(startTime, endTime);
    
    // Set event color based on channel
    if (channel && CALENDAR_CONFIG.EVENT_COLOR_MAPPING[channel]) {
      event.setColor(CALENDAR_CONFIG.EVENT_COLOR_MAPPING[channel]);
    }
    
    if (showConfirmation) {
      SpreadsheetApp.getUi().alert('Calendar event updated successfully!');
    }
    
    return eventId;
  } catch (e) {
    if (showConfirmation) {
      SpreadsheetApp.getUi().alert('Failed to update calendar event: ' + e.toString());
    }
    console.error('Failed to update calendar event:', e);
    return null;
  }
}

/**
 * Deletes a calendar event for a content item
 * @param {number} row Row number in the content calendar
 * @return {boolean} True if successful, false otherwise
 */
function deleteCalendarEvent(row) {
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarSheet = ss.getSheetByName('Content Calendar');
  
  if (!calendarSheet) {
    return false;
  }
  
  // Get event ID
  let eventId = '';
  try {
    eventId = calendarSheet.getRange(row, CALENDAR_CONFIG.EVENT_ID_COLUMN).getValue();
  } catch (e) {
    // Column might not exist
    return false;
  }
  
  // If no event ID, nothing to delete
  if (!eventId) {
    return false;
  }
  
  // Get the Google Calendar
  const calendar = getContentCalendar();
  if (!calendar) {
    return false;
  }
  
  // Delete the event
  try {
    const event = calendar.getEventById(eventId);
    if (event) {
      event.deleteEvent();
      // Clear the event ID
      calendarSheet.getRange(row, CALENDAR_CONFIG.EVENT_ID_COLUMN).setValue('');
      return true;
    }
  } catch (e) {
    console.error('Failed to delete calendar event:', e);
  }
  
  return false;
}

/**
 * Syncs all content items with scheduled status to Google Calendar
 */
function syncCalendarEvents() {
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarSheet = ss.getSheetByName('Content Calendar');
  
  if (!calendarSheet) {
    SpreadsheetApp.getUi().alert('Content Calendar sheet not found.');
    return;
  }
  
  // Ensure event ID column exists
  ensureEventIdColumnExists(calendarSheet);
  
  // Get all data
  const dataRange = calendarSheet.getDataRange();
  const values = dataRange.getValues();
  
  // Skip if less than 3 rows (header rows only)
  if (values.length < 3) {
    SpreadsheetApp.getUi().alert('No content items found.');
    return;
  }
  
  // Process each row (starting from row 3, index 2)
  let createdCount = 0;
  let updatedCount = 0;
  let errorCount = 0;
  
  for (let i = 2; i < values.length; i++) {
    const row = values[i];
    const status = row[CALENDAR_CONFIG.STATUS_COLUMN - 1]; // Adjust for 0-based array
    
    // Only process items with "Schedule" status
    if (status === 'Schedule') {
      const eventId = row[CALENDAR_CONFIG.EVENT_ID_COLUMN - 1]; // Adjust for 0-based array
      
      if (eventId) {
        // Update existing event
        if (updateCalendarEvent(i + 1, eventId, false)) {
          updatedCount++;
        } else {
          errorCount++;
        }
      } else {
        // Create new event
        if (createCalendarEvent(i + 1, false)) {
          createdCount++;
        } else {
          errorCount++;
        }
      }
    }
  }
  
  // Show summary
  SpreadsheetApp.getUi().alert(
    `Calendar sync completed:\n\n` +
    `- Created: ${createdCount} events\n` +
    `- Updated: ${updatedCount} events\n` +
    `- Errors: ${errorCount} events`
  );
}

/**
 * Automatically creates or updates calendar events when content status changes to "Schedule"
 * This is called by the onEdit trigger
 * @param {object} e Edit event object
 */
function handleStatusChangeForCalendar(e) {
  // Skip if auto-create is disabled
  if (!CALENDAR_CONFIG.AUTO_CREATE_EVENTS) {
    return;
  }
  
  // Get the edited sheet and range
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // Only process edits in the Content Calendar sheet
  if (sheet.getName() !== 'Content Calendar') {
    return;
  }
  
  // Only process edits to the status column
  if (range.getColumn() !== CALENDAR_CONFIG.STATUS_COLUMN) {
    return;
  }
  
  // Skip header rows
  if (range.getRow() < 3) {
    return;
  }
  
  // Get the new status value
  const newStatus = range.getValue();
  
  // Only create events for "Schedule" status
  if (newStatus === 'Schedule') {
    createCalendarEvent(range.getRow(), false);
  }
}

/**
 * Updates calendar events when content details change
 * @param {object} e Edit event object
 */
function handleContentChangeForCalendar(e) {
  // Skip if update on change is disabled
  if (!CALENDAR_CONFIG.UPDATE_EVENTS_ON_CHANGE) {
    return;
  }
  
  // Get the edited sheet and range
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // Only process edits in the Content Calendar sheet
  if (sheet.getName() !== 'Content Calendar') {
    return;
  }
  
  // Only process edits to relevant columns
  const col = range.getColumn();
  if (col !== CALENDAR_CONFIG.DATE_COLUMN && 
      col !== CALENDAR_CONFIG.CHANNEL_COLUMN && 
      col !== CALENDAR_CONFIG.CONTENT_COLUMN) {
    return;
  }
  
  // Skip header rows
  if (range.getRow() < 3) {
    return;
  }
  
  // Check if this row has an event ID and status is "Schedule"
  try {
    const status = sheet.getRange(range.getRow(), CALENDAR_CONFIG.STATUS_COLUMN).getValue();
    if (status !== 'Schedule') {
      return;
    }
    
    const eventId = sheet.getRange(range.getRow(), CALENDAR_CONFIG.EVENT_ID_COLUMN).getValue();
    if (eventId) {
      updateCalendarEvent(range.getRow(), eventId, false);
    }
  } catch (e) {
    // Column might not exist or other error
    console.error('Error updating calendar event:', e);
  }
}

/**
 * Imports events from Google Calendar as content items
 */
function importCalendarEvents() {
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarSheet = ss.getSheetByName('Content Calendar');
  
  if (!calendarSheet) {
    SpreadsheetApp.getUi().alert('Content Calendar sheet not found.');
    return;
  }
  
  // Ensure event ID column exists
  ensureEventIdColumnExists(calendarSheet);
  
  // Get the Google Calendar
  const calendar = getContentCalendar();
  if (!calendar) {
    SpreadsheetApp.getUi().alert('Could not access Google Calendar. Please check calendar ID in Settings sheet.');
    return;
  }
  
  // Prompt for date range
  const ui = SpreadsheetApp.getUi();
  
  // Get start date
  const startDateResponse = ui.prompt(
    'Import Calendar Events',
    'Enter start date (YYYY-MM-DD):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (startDateResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  // Parse start date
  let startDate;
  try {
    startDate = new Date(startDateResponse.getResponseText().trim());
    if (isNaN(startDate.getTime())) {
      throw new Error('Invalid date format');
    }
  } catch (e) {
    ui.alert('Invalid start date. Please use YYYY-MM-DD format.');
    return;
  }
  
  // Get end date
  const endDateResponse = ui.prompt(
    'Import Calendar Events',
    'Enter end date (YYYY-MM-DD):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (endDateResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  // Parse end date
  let endDate;
  try {
    endDate = new Date(endDateResponse.getResponseText().trim());
    if (isNaN(endDate.getTime())) {
      throw new Error('Invalid date format');
    }
    
    // Set to end of day
    endDate.setHours(23, 59, 59);
  } catch (e) {
    ui.alert('Invalid end date. Please use YYYY-MM-DD format.');
    return;
  }
  
  // Validate date range
  if (endDate < startDate) {
    ui.alert('End date must be after start date.');
    return;
  }
  
  // Get events in the date range
  const events = calendar.getEvents(startDate, endDate);
  
  // If no events found
  if (events.length === 0) {
    ui.alert('No events found in the specified date range.');
    return;
  }
  
  // Find the first empty row
  let firstEmptyRow = 3;
  while (calendarSheet.getRange(firstEmptyRow, CALENDAR_CONFIG.ID_COLUMN).getValue() !== '') {
    firstEmptyRow++;
  }
  
  // Process events and add to content calendar
  let importedCount = 0;
  let skippedCount = 0;
  
  for (const event of events) {
    // Check if event already exists in the calendar
    const eventId = event.getId();
    let found = false;
    
    try {
      const eventIdColumn = calendarSheet.getRange(3, CALENDAR_CONFIG.EVENT_ID_COLUMN, 
                                                   firstEmptyRow - 3, 1).getValues();
      
      for (let i = 0; i < eventIdColumn.length; i++) {
        if (eventIdColumn[i][0] === eventId) {
          found = true;
          break;
        }
      }
    } catch (e) {
      // Column might not exist
    }
    
    // Skip if already imported
    if (found) {
      skippedCount++;
      continue;
    }
    
    // Parse event information
    const title = event.getTitle();
    let channel = 'Twitter'; // Default channel
    let content = title;
    
    // Try to extract channel from title format [Channel] Content
    const channelMatch = title.match(/^\[(.*?)\](.*)/);
    if (channelMatch) {
      channel = channelMatch[1].trim();
      content = channelMatch[2].trim();
    }
    
    // Get publication date from event start time
    const pubDate = event.getStartTime();
    
    // Create a new content row
    const newRow = firstEmptyRow + importedCount;
    
    // Generate content ID
    const contentId = generateContentId(newRow);
    
    // Set row data
    calendarSheet.getRange(newRow, CALENDAR_CONFIG.ID_COLUMN).setValue(contentId);
    calendarSheet.getRange(newRow, CALENDAR_CONFIG.DATE_COLUMN).setValue(pubDate);
    calendarSheet.getRange(newRow, CALENDAR_CONFIG.STATUS_COLUMN).setValue('Schedule');
    calendarSheet.getRange(newRow, CALENDAR_CONFIG.CHANNEL_COLUMN).setValue(channel);
    calendarSheet.getRange(newRow, CALENDAR_CONFIG.CONTENT_COLUMN).setValue(content);
    calendarSheet.getRange(newRow, CALENDAR_CONFIG.EVENT_ID_COLUMN).setValue(eventId);
    
    // Set week number
    try {
      // Find the week column (usually column C)
      const weekColumn = CALENDAR_CONFIG.DATE_COLUMN + 1;
      calendarSheet.getRange(newRow, weekColumn).setFormula(`=WEEKNUM(B${newRow},2)`);
    } catch (e) {
      // Week column might not exist
    }
    
    // Set created date if that column exists
    try {
      const createdColumn = 12; // Column L
      calendarSheet.getRange(newRow, createdColumn).setValue(new Date());
    } catch (e) {
      // Column might not exist
    }
    
    importedCount++;
  }
  
  // Show summary
  ui.alert(
    `Calendar import completed:\n\n` +
    `- Imported: ${importedCount} events\n` +
    `- Skipped (already imported): ${skippedCount} events`
  );
}

/**
 * Gets the Google Calendar for content
 * @return {Calendar} Google Calendar object or null if not found
 */
function getContentCalendar() {
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  
  if (!settingsSheet) {
    return null;
  }
  
  // Get calendar ID from settings
  const calendarId = settingsSheet.getRange(CALENDAR_CONFIG.SETTINGS_CALENDAR_ID_CELL).getValue();
  
  // If no calendar ID, use default calendar
  let calendar;
  try {
    if (calendarId) {
      calendar = CalendarApp.getCalendarById(calendarId);
    } else {
      calendar = CalendarApp.getDefaultCalendar();
    }
  } catch (e) {
    console.error('Error accessing calendar:', e);
    return null;
  }
  
  return calendar;
}

/**
 * Ensures that the Event ID column exists in the content calendar
 * @param {Sheet} sheet Content Calendar sheet
 */
function ensureEventIdColumnExists(sheet) {
  try {
    // Check if the column already exists by checking the header
    const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // If the last column is less than the event ID column, add columns
    if (sheet.getLastColumn() < CALENDAR_CONFIG.EVENT_ID_COLUMN) {
      const columnsToAdd = CALENDAR_CONFIG.EVENT_ID_COLUMN - sheet.getLastColumn();
      sheet.insertColumnsAfter(sheet.getLastColumn(), columnsToAdd);
    }
    
    // Check if the header exists
    if (headers.length < CALENDAR_CONFIG.EVENT_ID_COLUMN || 
        headers[CALENDAR_CONFIG.EVENT_ID_COLUMN - 1] !== 'Event ID') {
      // Set the header
      sheet.getRange(2, CALENDAR_CONFIG.EVENT_ID_COLUMN).setValue('Event ID');
    }
  } catch (e) {
    console.error('Error ensuring Event ID column exists:', e);
  }
}

/**
 * Generates a content ID for a new row
 * @param {number} row Row number
 * @return {string} Generated content ID
 */
function generateContentId(row) {
  return 'CONT-' + padNumber(row - 2, 3);
}

/**
 * Pads a number with leading zeros
 * @param {number} number Number to pad
 * @param {number} length Desired length
 * @return {string} Padded number as string
 */
function padNumber(number, length) {
  let str = '' + number;
  while (str.length < length) {
    str = '0' + str;
  }
  return str;
}

/**
 * Creates a link to open the calendar in a new tab
 */
function openCalendarInNewTab() {
  // Get the Google Calendar
  const calendar = getContentCalendar();
  if (!calendar) {
    SpreadsheetApp.getUi().alert('Could not access Google Calendar. Please check calendar ID in Settings sheet.');
    return;
  }
  
  // Get the calendar URL
  let calendarUrl = 'https://calendar.google.com/calendar/';
  
  // Try to get calendar ID
  try {
    const calendarId = calendar.getId();
    if (calendarId !== CalendarApp.getDefaultCalendar().getId()) {
      calendarUrl += 'r?cid=' + encodeURIComponent(calendarId);
    }
  } catch (e) {
    // Use default URL if error
  }
  
  // Create HTML to open URL in new tab
  const htmlOutput = HtmlService
    .createHtmlOutput(`
      <script>
        window.open('${calendarUrl}', '_blank');
        google.script.host.close();
      </script>
    `)
    .setWidth(1)
    .setHeight(1);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Opening Calendar...');
}

/**
 * Creates a calendar menu
 */
function createCalendarMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Calendar')
    .addItem('Create Event for Selected Content', 'createEventForSelectedContent')
    .addItem('Update Event for Selected Content', 'updateEventForSelectedContent')
    .addItem('Delete Event for Selected Content', 'deleteEventForSelectedContent')
    .addSeparator()
    .addItem('Sync All Scheduled Content to Calendar', 'syncCalendarEvents')
    .addItem('Import Events from Calendar', 'importCalendarEvents')
    .addSeparator()
    .addItem('Open Calendar in New Tab', 'openCalendarInNewTab')
    .addToUi();
}

/**
 * Creates a calendar event for the selected content
 */
function createEventForSelectedContent() {
  // Get the active sheet and selected cell
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = SpreadsheetApp.getActiveRange();
  
  // Only process in the Content Calendar sheet
  if (sheet.getName() !== 'Content Calendar') {
    SpreadsheetApp.getUi().alert('Please select a content item in the Content Calendar sheet.');
    return;
  }
  
  // Skip header rows
  if (range.getRow() < 3) {
    SpreadsheetApp.getUi().alert('Please select a content item row (not headers).');
    return;
  }
  
  // Create event for the selected row
  createCalendarEvent(range.getRow(), true);
}

/**
 * Updates a calendar event for the selected content
 */
function updateEventForSelectedContent() {
  // Get the active sheet and selected cell
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = SpreadsheetApp.getActiveRange();
  
  // Only process in the Content Calendar sheet
  if (sheet.getName() !== 'Content Calendar') {
    SpreadsheetApp.getUi().alert('Please select a content item in the Content Calendar sheet.');
    return;
  }
  
  // Skip header rows
  if (range.getRow() < 3) {
    SpreadsheetApp.getUi().alert('Please select a content item row (not headers).');
    return;
  }
  
  // Ensure event ID column exists
  ensureEventIdColumnExists(sheet);
  
  // Get event ID
  const eventId = sheet.getRange(range.getRow(), CALENDAR_CONFIG.EVENT_ID_COLUMN).getValue();
  
  if (!eventId) {
    SpreadsheetApp.getUi().alert('No calendar event found for this content item. Please create one first.');
    return;
  }
  
  // Update event for the selected row
  updateCalendarEvent(range.getRow(), eventId, true);
}

/**
 * Deletes a calendar event for the selected content
 */
function deleteEventForSelectedContent() {
  // Get the active sheet and selected cell
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = SpreadsheetApp.getActiveRange();
  
  // Only process in the Content Calendar sheet
  if (sheet.getName() !== 'Content Calendar') {
    SpreadsheetApp.getUi().alert('Please select a content item in the Content Calendar sheet.');
    return;
  }
  
  // Skip header rows
  if (range.getRow() < 3) {
    SpreadsheetApp.getUi().alert('Please select a content item row (not headers).');
    return;
  }
  
  // Try to delete the event
  if (deleteCalendarEvent(range.getRow())) {
    SpreadsheetApp.getUi().alert('Calendar event deleted successfully.');
  } else {
    SpreadsheetApp.getUi().alert('No calendar event found for this content item or error deleting event.');
  }
}