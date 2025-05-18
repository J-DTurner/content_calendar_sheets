/**
 * Social Media Content Calendar – Google Apps Script Functions
 *
 * This script provides automation features for the content calendar, including:
 * • Auto-updating timestamps
 * • Status-change tracking (basic, extended in status_tracking_script.js)
 * • Google Drive integration (basic, extended in drive_integration_script.js)
 * • Custom menu creation (now mostly consolidated in main_menu.js)
 * • Data-validation maintenance (now mostly consolidated in main_menu.js)
 */

// ───────────────────────────────────────────────────────────────────────────────
// Global constants (can be centralized if used by many scripts)
// ───────────────────────────────────────────────────────────────────────────────

// const SHEET_NAMES = { ... } // Assumed to be in main_menu.js or a central config.js
// const COLUMNS = { ... } // Assumed to be in main_menu.js or a central config.js

// ───────────────────────────────────────────────────────────────────────────────
// Event Triggers (onEdit is primary here)
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Processes edits made to the spreadsheet.
 * This is a trigger function called by Google Sheets.
 * Focuses on actions specific to the Content Calendar sheet.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The event object.
 */
function onEditSocialCalendar(e) { // Renamed to avoid conflict if a global onEdit exists
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  // Define expected sheet name - ideally use a global constant
  const CONTENT_CALENDAR_SHEET_NAME = 'Content Calendar';
  if (sheetName !== CONTENT_CALENDAR_SHEET_NAME) return;

  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  // Define expected column numbers - ideally use global constants
  const DATE_COLUMN = 2;
  const STATUS_COLUMN = 4;
  const MODIFIED_COLUMN = 13; // Assuming 'Updated' is column M
  const STATUS_CHANGED_COLUMN = 14; // Assuming 'Status Changed' is column N

  if (row < 3) return; // Skip header rows (assuming data starts row 3)

  // Update "Last Modified" timestamp
  try {
    sheet.getRange(row, MODIFIED_COLUMN).setValue(new Date());
  } catch (err) {
    Logger.log(`Error updating 'Last Modified' in onEditSocialCalendar for row ${row}: ${err}`);
  }

  // If Status column was edited
  if (col === STATUS_COLUMN) {
    try {
      sheet.getRange(row, STATUS_CHANGED_COLUMN).setValue(new Date());
      // More comprehensive status tracking (history, notifications) should be in status_tracking_script.js
      // and called by a central onEdit or its own trigger if necessary.
      // For now, this just updates the timestamp.
      if (typeof trackStatusChange === 'function') {
           // Ensure this specific onEdit doesn't duplicate calls if trackStatusChange is also in a global onEdit
           // trackStatusChange(sheet, row, e.oldValue, e.value);
      }
    } catch (err) {
      Logger.log(`Error updating 'Status Changed' in onEditSocialCalendar for row ${row}: ${err}`);
    }
  }

  // If Date column was edited, update Week Number
  if (col === DATE_COLUMN) {
    try {
      updateWeekNumberForRow(sheet, row); // Specific row update
    } catch (err) {
      Logger.log(`Error updating week number in onEditSocialCalendar for row ${row}: ${err}`);
    }
  }

  // Call calendar integration hooks if they exist and are relevant
  if (typeof handleStatusChangeForCalendar === 'function' && col === STATUS_COLUMN) {
      handleStatusChangeForCalendar(e);
  }
  if (typeof handleContentChangeForCalendar === 'function' && (col === DATE_COLUMN || col === 5 /*Channel*/ || col === 6 /*Content*/)) {
      handleContentChangeForCalendar(e);
  }

}

// ───────────────────────────────────────────────────────────────────────────────
// Timestamp and Week Number Helpers
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Updates the week number for a specific row based on its date.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 * @param {number} row The row number to update.
 */
function updateWeekNumberForRow(sheet, row) {
  const DATE_COLUMN = 2; // Date is in Column B
  const WEEK_COLUMN = 3; // Week is in Column C

  const dateCell = sheet.getRange(row, DATE_COLUMN);
  const dateValue = dateCell.getValue();
  const weekCell = sheet.getRange(row, WEEK_COLUMN);

  if (dateValue instanceof Date && !isNaN(dateValue.valueOf())) {
    // Set formula for dynamic week calculation
    weekCell.setFormula(`=IF(${dateCell.getA1Notation()}<>"", WEEKNUM(${dateCell.getA1Notation()},2), "")`);
  } else {
    // If date is cleared or invalid, clear the week cell
    weekCell.clearContent();
  }
}

/**
 * Utility to get the last row with content in a sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to check.
 * @param {number} [checkColumn=1] The column to check for data (default: Column A 'ID').
 * @param {number} [headerRows=2] The number of header rows to skip.
 * @return {number} The last row number containing data.
 */
function getLastContentRow(sheet, checkColumn = 1, headerRows = 2) {
  const lastRowData = sheet.getLastRow();
  if (lastRowData <= headerRows) return headerRows; // Only headers or empty

  const columnValues = sheet.getRange(headerRows + 1, checkColumn, lastRowData - headerRows, 1).getValues();
  for (let i = columnValues.length - 1; i >= 0; i--) {
    if (columnValues[i][0] !== "") {
      return headerRows + i + 1; // Actual row number
    }
  }
  return headerRows; // No data found below headers
}

// ───────────────────────────────────────────────────────────────────────────────
// Content Item Creation
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Creates a new content item row in the Content Calendar sheet.
 */
function createNewContentItem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarSheetName = 'Content Calendar';
  const sheet = ss.getSheetByName(calendarSheetName);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet "${calendarSheetName}" not found. Please run initialization.`);
    return;
  }

  // Column definitions (ensure these match global constants or are passed)
  const COLS = { ID: 1, DATE: 2, WEEK: 3, STATUS: 4, CHANNEL: 5, CONTENT: 6, LINK: 7, PILLAR: 8, FORMAT: 9, ASSIGNED: 10, NOTES: 11, CREATED: 12, UPDATED: 13, STATUS_CHANGED: 14, EVENT_ID: 15 };
  const HEADER_ROWS = 2;

  const newRow = getLastContentRow(sheet, COLS.ID, HEADER_ROWS) + 1;
  const now = new Date();
  const tomorrow = new Date(now.getTime() + (24 * 60 * 60 * 1000));

  // Set default values
  sheet.getRange(newRow, COLS.DATE).setValue(tomorrow).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(newRow, COLS.STATUS).setValue('Planned'); // Default status
  sheet.getRange(newRow, COLS.CREATED).setValue(now).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange(newRow, COLS.UPDATED).setValue(now).setNumberFormat('yyyy-mm-dd hh:mm:ss');

  // Set formulas for ID and Week
  const dateCellA1 = sheet.getRange(newRow, COLS.DATE).getA1Notation();
  try {
    sheet.getRange(newRow, COLS.ID).setFormula(`=IF(${dateCellA1}<>"", "CONT-" & TEXT(${newRow}-${HEADER_ROWS},"000"), "")`);
    sheet.getRange(newRow, COLS.WEEK).setFormula(`=IF(${dateCellA1}<>"", WEEKNUM(${dateCellA1},2), "")`);
  } catch (formulaError) {
    Logger.log(`Error setting formulas for new row ${newRow}: ${formulaError}`);
    // Fallback to static value if formula fails (e.g., if row is too high for simple TEXT(ROW()) logic)
    sheet.getRange(newRow, COLS.ID).setValue(`CONT-${String(newRow - HEADER_ROWS).padStart(3, '0')}`);
    if (sheet.getRange(newRow, COLS.DATE).getValue() instanceof Date) {
        sheet.getRange(newRow, COLS.WEEK).setValue(sheet.getRange(newRow, COLS.DATE).getValue().getWeekNumber());
    }
  }

  sheet.getRange(newRow, COLS.CONTENT).activate(); // Focus on Content/Idea cell for user
  SpreadsheetApp.getActiveSpreadsheet().toast(`New content item added in row ${newRow}. Fill in the details.`, "New Item", 7);
}

// ───────────────────────────────────────────────────────────────────────────────
// Data Validation & Formatting (called from main_menu during init or refresh)
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Updates data validation rules. Typically called from initializeContentCalendar or a manual refresh.
 */
function updateDataValidation() {
    Logger.log("Attempting to update Data Validation from social_media_content_calendar.js...");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const calendarSheet = ss.getSheetByName('Content Calendar');
    if (calendarSheet && typeof setupDataValidation === 'function') { // Check if setupDataValidation is globally available
        setupDataValidation(calendarSheet); // Assumes setupDataValidation is defined (e.g. in main_menu.js)
        SpreadsheetApp.getActiveSpreadsheet().toast('Data validation rules updated.', "Setup", 5);
    } else if (!calendarSheet) {
         SpreadsheetApp.getUi().alert('Content Calendar sheet not found for data validation update.');
    } else {
        Logger.log("setupDataValidation function not found. Cannot update data validation rules.");
        SpreadsheetApp.getUi().alert("Data validation setup function is missing.");
    }
}

/**
 * Updates status colors. Typically called from initializeContentCalendar or a manual refresh.
 */
function updateStatusColors() {
   Logger.log("Attempting to update Status Colors from social_media_content_calendar.js...");
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const sheet = ss.getSheetByName('Content Calendar');
   if(sheet && typeof setupConditionalFormatting === 'function') { // Check if setupConditionalFormatting is globally available
        setupConditionalFormatting(sheet); // Assumes setupConditionalFormatting is defined (e.g. in main_menu.js)
        SpreadsheetApp.getActiveSpreadsheet().toast('Status colors and conditional formatting updated.', "Setup", 5);
   } else if (!sheet) {
        SpreadsheetApp.getUi().alert('Content Calendar sheet not found for color update.');
   } else {
       Logger.log("setupConditionalFormatting function not found. Cannot update status colors.");
       SpreadsheetApp.getUi().alert("Conditional formatting setup function is missing.");
   }
}

// ───────────────────────────────────────────────────────────────────────────────
// Archiving
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Archives old content items based on a date threshold.
 */
function archiveOldContent() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings'); // Expected name
  const calendarSheet = ss.getSheetByName('Content Calendar');
  const archivesSheet = ss.getSheetByName('Archives');

  let archiveDays = 30; // Default
  if (settingsSheet) {
      const settingsData = settingsSheet.getRange("A1:B" + settingsSheet.getLastRow()).getValues();
      const thresholdRow = settingsData.find(row => row[0] === "Archive Days Threshold:");
      if (thresholdRow && !isNaN(parseInt(thresholdRow[1])) && parseInt(thresholdRow[1]) > 0) {
          archiveDays = parseInt(thresholdRow[1]);
      } else {
          Logger.log("Could not find/parse 'Archive Days Threshold:' in Settings or value is invalid. Using default 30 days.");
      }
  } else {
       Logger.log(`Sheet "Settings" not found. Using default 30 days for archive threshold.`);
  }

  const response = ui.alert(
      'Archive Old Content',
      `This will move content items with a 'Date' older than ${archiveDays} days to the 'Archives' sheet. This action cannot be easily undone. Continue?`,
      ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) {
      SpreadsheetApp.getActiveSpreadsheet().toast("Archive cancelled by user.", "Archive Status", 5);
      return;
  }

   if (!calendarSheet || !archivesSheet) {
        ui.alert(`Required sheets ("Content Calendar", "Archives") not found. Please run initialization.`);
        return;
   }

  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - archiveDays);
  cutoffDate.setHours(0, 0, 0, 0); // Compare against the start of the day

  const HEADER_ROWS = 2;
  const DATE_COLUMN_INDEX = 2; // Column B for Date
  const startRow = HEADER_ROWS + 1;
  const lastRowCalendar = calendarSheet.getLastRow();

   if (lastRowCalendar < startRow) {
       ui.alert('No content items found in the calendar to check for archiving.');
       return;
   }

  const numDataRows = lastRowCalendar - HEADER_ROWS;
  const calendarDataRange = calendarSheet.getRange(startRow, 1, numDataRows, calendarSheet.getLastColumn());
  const allCalendarData = calendarDataRange.getValues();
  const datesInData = calendarSheet.getRange(startRow, DATE_COLUMN_INDEX, numDataRows, 1).getValues();

  const rowsToArchive = []; // Stores actual data rows to be archived
  const rowIndicesToDelete = []; // Stores original row numbers for deletion

  for (let i = 0; i < datesInData.length; i++) {
    const itemDate = datesInData[i][0];
    if (itemDate instanceof Date && !isNaN(itemDate.valueOf()) && itemDate < cutoffDate) {
        rowsToArchive.push(allCalendarData[i]); // Store the entire row's data
        rowIndicesToDelete.push(startRow + i); // Store the actual sheet row number
    }
  }

  if (rowsToArchive.length === 0){
      ui.alert(`No content items found older than ${archiveDays} days to archive.`);
      return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(`Archiving ${rowsToArchive.length} items... Please wait.`, "Archiving", -1);

  try {
    // Append to Archives sheet
    // Ensure Archives sheet has headers matching Calendar sheet
    if (archivesSheet.getLastRow() === 0 || archivesSheet.getRange(1,1).getValue() === "") { // If archives is empty, copy headers
        const calendarHeaders = calendarSheet.getRange(1, 1, HEADER_ROWS, calendarSheet.getLastColumn()).getValues();
        archivesSheet.getRange(1, 1, HEADER_ROWS, calendarHeaders[0].length).setValues(calendarHeaders);
        // Copy formatting of headers as well
        calendarSheet.getRange(1,1,HEADER_ROWS, calendarSheet.getLastColumn()).copyFormatToRange(archivesSheet, 1, calendarHeaders[0].length, 1, HEADER_ROWS);
        archivesSheet.setFrozenRows(HEADER_ROWS);
    }

    archivesSheet.getRange(archivesSheet.getLastRow() + 1, 1, rowsToArchive.length, rowsToArchive[0].length)
                 .setValues(rowsToArchive);

    // Delete rows from Content Calendar (in reverse order to avoid index shifts)
    rowIndicesToDelete.reverse().forEach(rowIndex => {
        calendarSheet.deleteRow(rowIndex);
    });

    SpreadsheetApp.getActiveSpreadsheet().toast(`${rowsToArchive.length} item(s) archived successfully.`, "Archive Complete", 10);
    ui.alert(`${rowsToArchive.length} content item(s) archived.`);
  } catch (archiveError) {
      Logger.log(`Error during archiving: ${archiveError.toString()}`);
      SpreadsheetApp.getActiveSpreadsheet().toast(`Archiving failed. Check logs.`, "Archive Error", 10);
      ui.alert(`An error occurred during archiving: ${archiveError.toString()}`);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// Week Number Prototype (ensure it's defined once globally)
// ───────────────────────────────────────────────────────────────────────────────
// This is now defined in main_menu.js to ensure it's available globally.
// If Date.prototype.getWeekNumber is defined in multiple files, it might cause issues or be overwritten.
// It's best to have such prototype extensions in one central place, like main_menu.js or a dedicated utilities.js file.