/**
 * Date and Week Number Utilities for Social Media Content Calendar
 * 
 * This script provides utilities for date handling and week number calculations:
 * - Automated date population
 * - Week number calculations and formatting
 * - Date range functions
 * - Recurring content generation
 */

// Date and Week Configuration
const DATE_CONFIG = {
  DATE_COLUMN: 2,           // Column B
  WEEK_COLUMN: 3,           // Column C
  DATE_FORMAT: 'yyyy-MM-dd',
  WEEK_FORMAT: 'Week WW, YYYY',
  ISO_WEEK_TYPE: 2          // ISO standard (type 2)
};

/**
 * Auto-populates dates for a specified time range
 */
function autoPopulateDates() {
  const ui = SpreadsheetApp.getUi();
  
  // Prompt for start date
  const startDateResponse = ui.prompt(
    'Auto-Populate Dates',
    'Enter start date (YYYY-MM-DD):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (startDateResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  // Parse start date
  let startDate;
  try {
    const startDateStr = startDateResponse.getResponseText().trim();
    startDate = new Date(startDateStr);
    
    if (isNaN(startDate.getTime())) {
      throw new Error('Invalid date format');
    }
  } catch (error) {
    ui.alert('Invalid start date. Please use YYYY-MM-DD format.');
    return;
  }
  
  // Prompt for end date
  const endDateResponse = ui.prompt(
    'Auto-Populate Dates',
    'Enter end date (YYYY-MM-DD):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (endDateResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  // Parse end date
  let endDate;
  try {
    const endDateStr = endDateResponse.getResponseText().trim();
    endDate = new Date(endDateStr);
    
    if (isNaN(endDate.getTime())) {
      throw new Error('Invalid date format');
    }
  } catch (error) {
    ui.alert('Invalid end date. Please use YYYY-MM-DD format.');
    return;
  }
  
  // Validate date range
  if (endDate < startDate) {
    ui.alert('End date must be after start date.');
    return;
  }
  
  // Prompt for frequency
  const frequencyResponse = ui.prompt(
    'Auto-Populate Dates',
    'Enter frequency:\n\n' +
    '1. Daily\n' +
    '2. Weekly\n' +
    '3. Monthly\n' +
    '4. Custom (days)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (frequencyResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  // Parse frequency
  let frequencyDays = 1;
  const frequencyOption = frequencyResponse.getResponseText().trim();
  
  switch (frequencyOption) {
    case '1':
      frequencyDays = 1;
      break;
    case '2':
      frequencyDays = 7;
      break;
    case '3':
      // For monthly, we'll use a special approach
      frequencyDays = 30;
      break;
    case '4':
      // Custom frequency
      const customResponse = ui.prompt(
        'Custom Frequency',
        'Enter number of days between dates:',
        ui.ButtonSet.OK_CANCEL
      );
      
      if (customResponse.getSelectedButton() !== ui.Button.OK) {
        return;
      }
      
      try {
        frequencyDays = parseInt(customResponse.getResponseText().trim());
        if (isNaN(frequencyDays) || frequencyDays < 1) {
          throw new Error('Invalid number');
        }
      } catch (error) {
        ui.alert('Invalid frequency. Please enter a positive number.');
        return;
      }
      break;
    default:
      ui.alert('Invalid option. Please select 1, 2, 3, or 4.');
      return;
  }
  
  // Get the calendar sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Content Calendar');
  
  // Generate the dates
  let dates = [];
  if (frequencyOption === '3') {
    // Monthly frequency (same day each month)
    dates = generateMonthlyDates(startDate, endDate);
  } else {
    // Regular interval
    dates = generateDatesBetween(startDate, endDate, frequencyDays);
  }
  
  // Check if we have any dates
  if (dates.length === 0) {
    ui.alert('No dates to add within the specified range.');
    return;
  }
  
  // Confirm with the user
  const confirmResponse = ui.alert(
    'Confirm Date Creation',
    `This will create ${dates.length} new rows with dates from ${formatDate(dates[0])} to ${formatDate(dates[dates.length - 1])}.`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (confirmResponse !== ui.Button.OK) {
    return;
  }
  
  // Find the first empty row
  let firstEmptyRow = 3; // Start after headers
  while (sheet.getRange(firstEmptyRow, DATE_CONFIG.DATE_COLUMN).getValue() !== '') {
    firstEmptyRow++;
  }
  
  // Populate the dates
  for (let i = 0; i < dates.length; i++) {
    const rowIndex = firstEmptyRow + i;
    
    // Set the date
    sheet.getRange(rowIndex, DATE_CONFIG.DATE_COLUMN).setValue(dates[i]);
    
    // Generate ID if needed (assuming column A has a formula)
    if (sheet.getRange(rowIndex, 1).getValue() === '') {
      // Copy formula from row above if it exists
      const prevRowFormula = sheet.getRange(rowIndex - 1, 1).getFormula();
      if (prevRowFormula) {
        sheet.getRange(rowIndex, 1).setFormula(prevRowFormula);
      } else {
        // Generate a default ID
        sheet.getRange(rowIndex, 1).setValue('CONT-' + padNumber(rowIndex - 2, 3));
      }
    }
    
    // Set default status (assuming column D is Status)
    if (sheet.getRange(rowIndex, 4).getValue() === '') {
      sheet.getRange(rowIndex, 4).setValue('Planned');
    }
    
    // Add timestamp if needed (assuming column L is Created Date)
    if (sheet.getRange(rowIndex, 12).getValue() === '') {
      sheet.getRange(rowIndex, 12).setValue(new Date());
    }
  }
  
  // Update week numbers (they should auto-update via formula, but let's make sure)
  updateWeekNumbers(sheet, firstEmptyRow, firstEmptyRow + dates.length - 1);
  
  // Select the range of dates added
  sheet.getRange(firstEmptyRow, DATE_CONFIG.DATE_COLUMN, dates.length).activate();
  
  // Show confirmation
  ui.alert(`Successfully added ${dates.length} dates to the calendar.`);
}

/**
 * Generates dates between start and end with a specified interval
 */
function generateDatesBetween(startDate, endDate, intervalDays) {
  const dates = [];
  let currentDate = new Date(startDate);
  
  while (currentDate <= endDate) {
    dates.push(new Date(currentDate));
    currentDate.setDate(currentDate.getDate() + intervalDays);
  }
  
  return dates;
}

/**
 * Generates monthly dates (same day each month)
 */
function generateMonthlyDates(startDate, endDate) {
  const dates = [];
  let currentDate = new Date(startDate);
  const dayOfMonth = startDate.getDate();
  
  while (currentDate <= endDate) {
    dates.push(new Date(currentDate));
    
    // Move to next month
    let nextMonth = currentDate.getMonth() + 1;
    let nextYear = currentDate.getFullYear();
    
    if (nextMonth > 11) {
      nextMonth = 0;
      nextYear++;
    }
    
    // Create the next date
    currentDate = new Date(nextYear, nextMonth, 1);
    
    // Set the day of month, adjusting for months with fewer days
    const lastDayOfMonth = new Date(nextYear, nextMonth + 1, 0).getDate();
    currentDate.setDate(Math.min(dayOfMonth, lastDayOfMonth));
  }
  
  return dates;
}

/**
 * Updates week numbers for a range of rows
 */
function updateWeekNumbers(sheet, startRow, endRow) {
  for (let row = startRow; row <= endRow; row++) {
    const dateCell = sheet.getRange(row, DATE_CONFIG.DATE_COLUMN);
    const date = dateCell.getValue();
    
    if (date instanceof Date) {
      const weekNumber = getISOWeekNumber(date);
      const weekCell = sheet.getRange(row, DATE_CONFIG.WEEK_COLUMN);
      
      // Check if cell has a formula
      if (weekCell.getFormula()) {
        // Skip if it has a formula (it will update automatically)
        continue;
      }
      
      weekCell.setValue(weekNumber);
    }
  }
}

/**
 * Gets the ISO week number for a date
 */
function getISOWeekNumber(date) {
  // Use the WEEKNUM formula equivalent with type 2 (ISO)
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'w');
}

/**
 * Formats a date according to the configured format
 */
function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), DATE_CONFIG.DATE_FORMAT);
}

/**
 * Pads a number with leading zeros
 */
function padNumber(number, length) {
  let str = '' + number;
  while (str.length < length) {
    str = '0' + str;
  }
  return str;
}

/**
 * Gets the week dates for a given week number and year
 */
function getWeekDates(weekNumber, year) {
  // Find the first day of the year
  const firstDayOfYear = new Date(year, 0, 1);
  
  // Get the day of week of the first day (0 = Sunday, 1 = Monday, etc.)
  const firstDayOfWeek = firstDayOfYear.getDay();
  
  // Calculate the date of the first day of the first week
  // In ISO 8601, week 1 is the week with the first Thursday of the year
  let daysToAdd = 1 - firstDayOfWeek;
  if (firstDayOfWeek === 0) daysToAdd = -6; // Sunday needs to go back to previous Monday
  else if (firstDayOfWeek > 1) daysToAdd = 9 - firstDayOfWeek; // Other days go to next Monday
  
  // Start date of week 1
  const firstWeekStart = new Date(year, 0, 1 + daysToAdd);
  
  // Calculate the start date of the requested week
  const weekStart = new Date(firstWeekStart);
  weekStart.setDate(firstWeekStart.getDate() + (weekNumber - 1) * 7);
  
  // Calculate the end date (start date + 6 days)
  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekStart.getDate() + 6);
  
  return {
    startDate: weekStart,
    endDate: weekEnd
  };
}

/**
 * Generates a formatted week string
 */
function formatWeekString(weekNumber, year) {
  return DATE_CONFIG.WEEK_FORMAT
    .replace('WW', padNumber(weekNumber, 2))
    .replace('YYYY', year);
}

/**
 * Gets the current week number
 */
function getCurrentWeekNumber() {
  const today = new Date();
  return parseInt(Utilities.formatDate(today, Session.getScriptTimeZone(), 'w'));
}

/**
 * Gets the current year
 */
function getCurrentYear() {
  return new Date().getFullYear();
}

/**
 * Jumps to a specific week in the calendar
 */
function jumpToWeek() {
  const ui = SpreadsheetApp.getUi();
  
  // Get the current week number as default
  const currentWeek = getCurrentWeekNumber();
  const currentYear = getCurrentYear();
  
  // Prompt for week number
  const weekResponse = ui.prompt(
    'Jump to Week',
    `Enter week number (1-53, current week is ${currentWeek}):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (weekResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  // Parse week number
  let weekNumber;
  try {
    weekNumber = parseInt(weekResponse.getResponseText().trim());
    if (isNaN(weekNumber) || weekNumber < 1 || weekNumber > 53) {
      throw new Error('Invalid week number');
    }
  } catch (error) {
    ui.alert('Invalid week number. Please enter a number between 1 and 53.');
    return;
  }
  
  // Prompt for year (default to current year)
  const yearResponse = ui.prompt(
    'Jump to Week',
    `Enter year (default: ${currentYear}):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (yearResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  // Parse year
  let year = currentYear;
  if (yearResponse.getResponseText().trim() !== '') {
    try {
      year = parseInt(yearResponse.getResponseText().trim());
      if (isNaN(year) || year < 2000 || year > 2100) {
        throw new Error('Invalid year');
      }
    } catch (error) {
      ui.alert('Invalid year. Please enter a valid year.');
      return;
    }
  }
  
  // Get the calendar sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Content Calendar');
  
  // Find all rows with the specified week number
  const weekColumn = sheet.getRange(3, DATE_CONFIG.WEEK_COLUMN, sheet.getLastRow() - 2, 1);
  const weekValues = weekColumn.getValues();
  
  let matchingRows = [];
  for (let i = 0; i < weekValues.length; i++) {
    const rowWeekNumber = weekValues[i][0];
    
    // Skip empty cells
    if (rowWeekNumber === '') continue;
    
    // Check for a match
    if (rowWeekNumber == weekNumber) { // Intentional loose equality
      // Get the date to check the year
      const date = sheet.getRange(i + 3, DATE_CONFIG.DATE_COLUMN).getValue();
      
      if (date instanceof Date && date.getFullYear() === year) {
        matchingRows.push(i + 3); // Add row index (i+3 because we started at row 3)
      }
    }
  }
  
  // If no matches, ask if the user wants to create entries for this week
  if (matchingRows.length === 0) {
    const createResponse = ui.alert(
      'No Content Found',
      `No content found for Week ${weekNumber}, ${year}. Would you like to create entries for this week?`,
      ui.ButtonSet.YES_NO
    );
    
    if (createResponse === ui.Button.YES) {
      // Get the week start and end dates
      const weekDates = getWeekDates(weekNumber, year);
      
      // Create entries for this week
      createEntriesForWeek(weekDates.startDate, weekDates.endDate);
    }
    
    return;
  }
  
  // Select the first matching row
  sheet.getRange(matchingRows[0], DATE_CONFIG.DATE_COLUMN).activate();
  
  // Optional: Add filter for this week
  const createFilterResponse = ui.alert(
    'Filter View',
    `Found ${matchingRows.length} entries for Week ${weekNumber}, ${year}. Would you like to create a filter view for this week?`,
    ui.ButtonSet.YES_NO
  );
  
  if (createFilterResponse === ui.Button.YES) {
    // Clear any existing filters
    let existingFilters = sheet.getFilter();
    if (existingFilters) {
      existingFilters.remove();
    }
    
    // Create a filter
    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    const filter = range.createFilter();
    
    // Set filter criteria for the week column
    filter.setColumnFilterCriteria(DATE_CONFIG.WEEK_COLUMN, 
      SpreadsheetApp.newFilterCriteria().whenNumberEqualTo(weekNumber).build());
  }
}

/**
 * Creates entries for a specific week
 */
function createEntriesForWeek(startDate, endDate) {
  const ui = SpreadsheetApp.getUi();
  
  // Ask how many entries to create
  const countResponse = ui.prompt(
    'Create Week Entries',
    'How many content items would you like to create for this week?',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (countResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  // Parse count
  let count;
  try {
    count = parseInt(countResponse.getResponseText().trim());
    if (isNaN(count) || count < 1) {
      throw new Error('Invalid count');
    }
  } catch (error) {
    ui.alert('Invalid number. Please enter a positive number.');
    return;
  }
  
  // Get the calendar sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Content Calendar');
  
  // Find the first empty row
  let firstEmptyRow = 3; // Start after headers
  while (sheet.getRange(firstEmptyRow, DATE_CONFIG.DATE_COLUMN).getValue() !== '') {
    firstEmptyRow++;
  }
  
  // Calculate date interval for distributing entries evenly across the week
  const daysDiff = Math.floor((endDate - startDate) / (1000 * 60 * 60 * 24));
  const interval = count > 1 ? daysDiff / (count - 1) : 0;
  
  // Create the entries
  for (let i = 0; i < count; i++) {
    const rowIndex = firstEmptyRow + i;
    
    // Calculate date for this entry
    const entryDate = new Date(startDate);
    if (count > 1) {
      entryDate.setDate(startDate.getDate() + Math.round(i * interval));
    }
    
    // Set the date
    sheet.getRange(rowIndex, DATE_CONFIG.DATE_COLUMN).setValue(entryDate);
    
    // Generate ID if needed (assuming column A has a formula)
    if (sheet.getRange(rowIndex, 1).getValue() === '') {
      // Copy formula from row above if it exists
      const prevRowFormula = sheet.getRange(rowIndex - 1, 1).getFormula();
      if (prevRowFormula) {
        sheet.getRange(rowIndex, 1).setFormula(prevRowFormula);
      } else {
        // Generate a default ID
        sheet.getRange(rowIndex, 1).setValue('CONT-' + padNumber(rowIndex - 2, 3));
      }
    }
    
    // Set default status (assuming column D is Status)
    if (sheet.getRange(rowIndex, 4).getValue() === '') {
      sheet.getRange(rowIndex, 4).setValue('Planned');
    }
    
    // Add timestamp if needed (assuming column L is Created Date)
    if (sheet.getRange(rowIndex, 12).getValue() === '') {
      sheet.getRange(rowIndex, 12).setValue(new Date());
    }
  }
  
  // Update week numbers
  updateWeekNumbers(sheet, firstEmptyRow, firstEmptyRow + count - 1);
  
  // Select the range of dates added
  sheet.getRange(firstEmptyRow, DATE_CONFIG.DATE_COLUMN, count).activate();
  
  // Show confirmation
  ui.alert(`Successfully added ${count} entries for the selected week.`);
}

/**
 * Adds date-related commands to the menu
 */
function updateDateMenu() {
  // Check if the menu already exists
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.createMenu('Date Tools')
      .addItem('Auto-Populate Dates', 'autoPopulateDates')
      .addItem('Jump to Week...', 'jumpToWeek')
      .addItem('View Current Week', 'viewCurrentWeek')
      .addItem('View Next Week', 'viewNextWeek')
      .addToUi();
  } catch (e) {
    // Menu might already exist, no need to handle error
  }
}

/**
 * Views the current week's content
 */
function viewCurrentWeek() {
  const currentWeek = getCurrentWeekNumber();
  const currentYear = getCurrentYear();
  
  filterByWeek(currentWeek, currentYear);
}

/**
 * Views the next week's content
 */
function viewNextWeek() {
  let nextWeek = getCurrentWeekNumber() + 1;
  let year = getCurrentYear();
  
  // Handle year transition
  if (nextWeek > 52) {
    nextWeek = 1;
    year++;
  }
  
  filterByWeek(nextWeek, year);
}

/**
 * Filters the content calendar by week
 */
function filterByWeek(weekNumber, year) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Content Calendar');
  
  // Clear any existing filters
  let existingFilters = sheet.getFilter();
  if (existingFilters) {
    existingFilters.remove();
  }
  
  // Create a filter
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const filter = range.createFilter();
  
  // Set filter criteria for the week column
  filter.setColumnFilterCriteria(DATE_CONFIG.WEEK_COLUMN, 
    SpreadsheetApp.newFilterCriteria().whenNumberEqualTo(weekNumber).build());
  
  // Find the first matching row
  const weekColumn = sheet.getRange(3, DATE_CONFIG.WEEK_COLUMN, sheet.getLastRow() - 2, 1);
  const weekValues = weekColumn.getValues();
  
  for (let i = 0; i < weekValues.length; i++) {
    const rowWeekNumber = weekValues[i][0];
    
    // Skip empty cells
    if (rowWeekNumber === '') continue;
    
    // Check for a match
    if (rowWeekNumber == weekNumber) { // Intentional loose equality
      // Get the date to check the year
      const date = sheet.getRange(i + 3, DATE_CONFIG.DATE_COLUMN).getValue();
      
      if (date instanceof Date && date.getFullYear() === year) {
        // Select the first matching row
        sheet.getRange(i + 3, DATE_CONFIG.DATE_COLUMN).activate();
        break;
      }
    }
  }
  
  // Show week info
  SpreadsheetApp.getActiveSpreadsheet().toast(`Showing content for Week ${weekNumber}, ${year}`, 'Week Filter Applied');
}