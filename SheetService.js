/**
 * SheetService.gs
 * Provides utility functions for working with Google Sheets data.
 */

/**
 * Finds all rows in a sheet where a specific column contains the given value.
 * 
 * @param {Sheet} sheet - The Google Sheet to search in
 * @param {number} columnIndex - The column index to check (0-based)
 * @param {string|number} value - The value to search for
 * @return {Array} Array of row data (as arrays) that contain the value
 * @private
 */
function findRowByValue_(sheet, columnIndex, value) {
  // Get all data from the sheet
  var data = sheet.getDataRange().getValues();
  var matchingRows = [];
  
  // Search for rows with matching value in the specified column
  for (var i = 0; i < data.length; i++) {
    if (data[i][columnIndex] == value) {
      matchingRows.push(data[i]);
    }
  }
  
  return matchingRows;
}

/**
 * Finds the row number (1-based) of the first row where a specific column contains the given value.
 * 
 * @param {Sheet} sheet - The Google Sheet to search in
 * @param {number} columnIndex - The column index to check (0-based)
 * @param {string|number} value - The value to search for
 * @return {number} The row number (1-based) or -1 if not found
 * @private
 */
function findRowNumberByValue_(sheet, columnIndex, value) {
  // Get all data from the sheet
  var data = sheet.getDataRange().getValues();
  
  // Search for the first row with matching value in the specified column
  for (var i = 0; i < data.length; i++) {
    if (data[i][columnIndex] == value) {
      return i + 1; // Convert to 1-based row number
    }
  }
  
  return -1; // Not found
}

/**
 * Updates an existing row if found based on key column, or appends a new row if not found.
 * 
 * @param {Sheet} sheet - The Google Sheet to update
 * @param {Array} rowData - The full row data to update or append
 * @param {number} keyColumnIndex - The column index used to identify the row (0-based)
 * @return {boolean} True if row was updated, false if a new row was appended
 * @private
 */
function updateOrAppendRow_(sheet, rowData, keyColumnIndex) {
  // Get the key value from the row data
  var keyValue = rowData[keyColumnIndex];
  
  // Find the row number with this key value
  var rowNumber = findRowNumberByValue_(sheet, keyColumnIndex, keyValue);
  
  if (rowNumber > 0) {
    // Update existing row
    var range = sheet.getRange(rowNumber, 1, 1, rowData.length);
    range.setValues([rowData]);
    return true;
  } else {
    // Append new row
    sheet.appendRow(rowData);
    return false;
  }
}

/**
 * Example usage:
 *
 * function testSheetFunctions() {
 *   var ss = SpreadsheetApp.getActiveSpreadsheet();
 *   var sheet = ss.getSheetByName("Assets");
 *   
 *   // Find rows with ID "asset123"
 *   var rows = findRowByValue_(sheet, 0, "asset123");
 *   Logger.log("Found rows: " + JSON.stringify(rows));
 *   
 *   // Find row number of first row with ID "asset123"
 *   var rowNum = findRowNumberByValue_(sheet, 0, "asset123");
 *   Logger.log("Found at row: " + rowNum);
 *   
 *   // Update or append a row
 *   var newData = ["asset123", "New Asset", "https://example.com/image.jpg", new Date()];
 *   var wasUpdated = updateOrAppendRow_(sheet, newData, 0);
 *   Logger.log("Row was updated: " + wasUpdated);
 * }
 */