// Configuration object and constants
var CONFIG = {};
var CONTENT_SHEET_NAME;
var ASSETS_SHEET_NAME = 'Assets';
var ASSET_ACTION_COLUMN_NAME;
var ROW_ID_COLUMN_NAME;
var ASSET_ACTION_COL_IDX;
var ROW_ID_COL_IDX;

const SETTINGS_CONFIG = {
  SETTINGS_SHEET: 'Settings',
  CONTENT_SHEET_CELL: 'B39',
  ASSET_ACTION_COLUMN_CELL: 'B40',
  ROW_ID_COLUMN_CELL: 'B41'
};

/**
 * Automatically runs when the spreadsheet is opened.
 * Creates the Asset Management menu with options.
 */
function onOpen() {
  // Get the UI service
  var ui = SpreadsheetApp.getUi();
  
  // Create a menu
  var menu = ui.createMenu('Asset Management');
  
  // Add menu items
  menu.addItem('1. Initialize/Refresh Asset Column', 'initializeOrRefreshAssetColumn');
  menu.addItem('2. Perform Action on Selected Asset Cell', 'performActionOnSelectedAssetCell');
  menu.addItem('3. (Admin) Load Configuration', 'loadConfigAndDisplay');
  
  // Add the menu to the UI
  menu.addToUi();
  
  // Load configuration
  loadConfig();
}

/**
 * Initializes asset management functionality.
 * This should be called from the main onOpen handler.
 * 
 * @return {boolean} True if initialization succeeded.
 */
function initializeAssetManagement() {
  // Load configuration
  return loadConfig();
}

/**
 * Loads configuration settings from the Settings sheet.
 * @return {boolean} True if configuration loaded successfully, false otherwise.
 */
function loadConfig() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = ss.getSheetByName(SETTINGS_CONFIG.SETTINGS_SHEET);

    if (!settingsSheet) {
      Logger.log('Settings sheet not found when loading config.');
      if (typeof initializeMainSheets === 'function') {
        initializeMainSheets(ss, false, true, false); // attempt to create Settings sheet
        settingsSheet = ss.getSheetByName(SETTINGS_CONFIG.SETTINGS_SHEET);
      }
      if (!settingsSheet) {
        SpreadsheetApp.getUi().alert('Error: Settings sheet not found. Please run setup to create it.');
        return false;
      }
    }

    if (typeof setupSettings === 'function') {
      try { setupSettings(settingsSheet, true); } catch(_) {}
    }

    CONFIG = {};
    CONTENT_SHEET_NAME = settingsSheet.getRange(SETTINGS_CONFIG.CONTENT_SHEET_CELL).getValue() || 'Content Calendar';
    ASSET_ACTION_COLUMN_NAME = settingsSheet.getRange(SETTINGS_CONFIG.ASSET_ACTION_COLUMN_CELL).getValue() || 'Asset Action';
    ROW_ID_COLUMN_NAME = settingsSheet.getRange(SETTINGS_CONFIG.ROW_ID_COLUMN_CELL).getValue() || 'ID';

    CONFIG.ContentSheetName = CONTENT_SHEET_NAME;
    CONFIG.AssetActionColumnName = ASSET_ACTION_COLUMN_NAME;
    CONFIG.RowIdColumnName = ROW_ID_COLUMN_NAME;
    
    // Validate essential settings
    if (!CONTENT_SHEET_NAME) {
      Logger.log('Error: ContentSheetName not found in Settings.');
      SpreadsheetApp.getUi().alert('Error: ContentSheetName not found in Settings sheet.');
      return false;
    }

    if (!ASSET_ACTION_COLUMN_NAME) {
      Logger.log('Error: AssetActionColumnName not found in Settings.');
      SpreadsheetApp.getUi().alert('Error: AssetActionColumnName not found in Settings sheet.');
      return false;
    }
    
    // Find column indices for important columns
    var contentSheet = ss.getSheetByName(CONTENT_SHEET_NAME);
    if (contentSheet) {
      getHeaderIndexes(contentSheet);
    }

    // Ensure Assets sheet exists whenever configuration is loaded
    ensureAssetsSheetExists();

    return true;
  } catch (e) {
    Logger.log('Error loading configuration: ' + e.toString());
    SpreadsheetApp.getUi().alert('Error loading configuration: ' + e.toString());
    return false;
  }
}

/**
 * Loads configuration and displays the loaded values in an alert.
 * @return {boolean} True if configuration loaded successfully, false otherwise.
 */
function loadConfigAndDisplay() {
  var result = loadConfig();
  if (result) {
    var configMessage = 'Configuration loaded successfully:\n';
    configMessage += 'Content Sheet: ' + CONTENT_SHEET_NAME + '\n';
    configMessage += 'Assets Sheet: ' + ASSETS_SHEET_NAME + '\n';
    configMessage += 'Asset Action Column: ' + ASSET_ACTION_COLUMN_NAME + '\n';
    configMessage += 'Row ID Column: ' + ROW_ID_COLUMN_NAME + '\n';
    
    // Get asset folder ID from centralized location
    var assetFolderId = "Not configured";
    if (typeof getPrimaryDriveAssetsFolderId === 'function') {
      var folderId = getPrimaryDriveAssetsFolderId();
      if (folderId) {
        assetFolderId = folderId;
      }
    }
    configMessage += 'Asset Folder ID: ' + assetFolderId + ' (from Settings B18)';
    
    SpreadsheetApp.getUi().alert(configMessage);
  }
  return result;
}

/**
 * Checks if the user has granted necessary permissions.
 * If not, calling this function via google.script.run will trigger the auth dialog.
 * @return {object} A status object indicating success if authorization is in place.
 */
function checkAuthorizationAndTriggerPrompt() {
  try {
    // This call ensures that basic script execution scopes are checked.
    // If any scopes from appsscript.json are not yet authorized,
    // Google Apps Script will show the authorization dialog automatically.
    Logger.log('User performing authorization check: ' + Session.getEffectiveUser().getEmail());

    // If the script reaches here, authorization is considered successful
    // (either pre-existing or just granted through the dialog).
    return { status: "success", message: "Authorization check passed." };
  } catch (e) {
    // This catch block might handle unexpected errors during the check itself,
    // though auth dialogs are typically handled by GAS before returning to google.script.run.
    Logger.log('Error during authorization check: ' + e.toString());
    // It's often better to let google.script.run's onFailure handler catch auth issues.
    // However, returning an error structure can be an alternative.
    // For forcing the standard auth flow, direct errors might not be needed here.
    // Let's rely on the successful execution path.
    // throw e; // Or re-throw if specific handling is needed.
    // For this flow, we primarily care about the success path post-GAS-auth-dialog.
    // If an error occurs here AFTER auth, it's a different issue.
     return { status: "error", message: "An unexpected error occurred during authorization check: " + e.toString() };
  }
}

/**
 * Initializes core services, specifically the Google Drive connection.
 * @return {object} A status object indicating success or failure of initialization.
 */
function initializeCoreServices() {
  try {
    // Placeholder for Google Drive initialization logic.
    // This interaction requires the Drive scope (e.g., https://www.googleapis.com/auth/drive.readonly or https://www.googleapis.com/auth/drive)
    // to be present in appsscript.json.

    // Example: Attempt a simple Drive operation.
    // Replace 'YOUR_TARGET_FOLDER_ID' with an actual Folder ID for testing,
    // or implement logic to retrieve/receive it.
    // For this example, we'll just check if we can access the root Drive folder as a basic test.
    var rootFolder = DriveApp.getRootFolder();
    Logger.log('[INIT-INFO]: Successfully accessed Drive. Root folder name: ' + rootFolder.getName());

    //
    // TODO: Implement actual Google Drive connection logic here.
    // e.g., DriveApp.getFolderById('YOUR_FOLDER_ID_HERE');
    // Perform checks, pull necessary data, etc.
    //

    return { success: true, message: "Core services initialized successfully. Google Drive access confirmed." };

  } catch (e) {
    Logger.log('[INIT-ERROR]: Error during core service initialization: ' + e.toString());
    
    // Categorize the error for better user guidance
    let errorType = "UNKNOWN";
    let errorMessage = "Failed to initialize core services: " + e.toString();
    
    if (errorMessage.includes("Access denied") || errorMessage.includes("Permission")) {
      errorType = "PERMISSION";
      errorMessage = "Permission denied for Drive access. Please ensure you have proper access rights.";
    } else if (errorMessage.includes("network") || errorMessage.includes("timeout")) {
      errorType = "NETWORK";
      errorMessage = "Network issue detected while accessing Drive. Please check your internet connection.";
    } else if (errorMessage.includes("not found")) {
      errorType = "NOT_FOUND";
      errorMessage = "Requested resource was not found in Drive. Please verify the resource exists.";
    }
    
    return { success: false, error: errorMessage, errorType: errorType };
  }
}

/**
 * Gets the header indexes for important columns in the sheet.
 * 
 * @param {Sheet} sheet - The sheet to analyze
 * @throws {Error} If the asset action column is not found
 */
function getHeaderIndexes(sheet) {
  if (!sheet) {
    throw new Error('Sheet is null or undefined');
  }
  
  // Read the second row (column headers)
  var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Find the asset action column index
  ASSET_ACTION_COL_IDX = headers.indexOf(ASSET_ACTION_COLUMN_NAME) + 1;
  
  // Throw an error if the asset action column is not found
  if (ASSET_ACTION_COL_IDX <= 0) {
    Logger.log('Error: Asset action column "' + ASSET_ACTION_COLUMN_NAME + '" not found in sheet headers.');
    throw new Error('Asset action column "' + ASSET_ACTION_COLUMN_NAME + '" not found in sheet headers.');
  }
  
  // Find the row ID column index if configured
  if (ROW_ID_COLUMN_NAME) {
    ROW_ID_COL_IDX = headers.indexOf(ROW_ID_COLUMN_NAME) + 1;
    
    // Set to 0 and warn if the row ID column is not found
    if (ROW_ID_COL_IDX <= 0) {
      ROW_ID_COL_IDX = 0;
      Logger.log('Warning: Row ID column "' + ROW_ID_COLUMN_NAME + '" not found in sheet headers.');
    }
  } else {
    ROW_ID_COL_IDX = 0;
  }
  
  return {
    assetActionColIdx: ASSET_ACTION_COL_IDX,
    rowIdColIdx: ROW_ID_COL_IDX
  };
}

/**
 * Gets a unique identifier for a row.
 * 
 * @param {Sheet} contentSheet - The sheet containing the row
 * @param {number} rowNum - The row number (1-based index)
 * @return {string} The row identifier
 */
function getRowIdentifier(contentSheet, rowNum) {
  if (!contentSheet) {
    return 'Row_' + rowNum;
  }
  
  // Use ROW_ID_COL_IDX if set
  if (ROW_ID_COL_IDX > 0) {
    try {
      // Get the value from the specified column
      var idCell = contentSheet.getRange(rowNum, ROW_ID_COL_IDX);
      var idValue = idCell.getValue();
      
      // If the cell has a value, use it; otherwise, fall back to Row_rowNum
      if (idValue && idValue.toString().trim() !== '') {
        return idValue.toString();
      }
    } catch (e) {
      Logger.log('Error getting row identifier: ' + e.toString());
      // Continue to fallback
    }
  }
  
  // Fallback: Use Row_rowNum
  return 'Row_' + rowNum;
}

/**
 * Initializes or refreshes the asset column in the content sheet.
 * This function is called from the menu "1. Initialize/Refresh Asset Column".
 * It sets up the asset action cells based on whether assets exist for each row.
 */
function initializeOrRefreshAssetColumn() {
  // Load configuration
  var configLoaded = loadConfig();
  if (!configLoaded) {
    SpreadsheetApp.getUi().alert('Error: Configuration failed to load. Cannot initialize asset column.');
    return;
  }
  
  // Get active spreadsheet and sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var contentSheet = ss.getSheetByName(CONTENT_SHEET_NAME);
  var assetsSheet = ss.getSheetByName(ASSETS_SHEET_NAME);
  
  // Check if content sheet exists
  if (!contentSheet) {
    SpreadsheetApp.getUi().alert('Error: Content sheet "' + CONTENT_SHEET_NAME + '" not found.');
    return;
  }
  
  try {
    // Get header indexes from the content sheet
    getHeaderIndexes(contentSheet);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
    return;
  }
  
  // Check if there is data in the content sheet
  var lastRow = contentSheet.getLastRow();
  if (lastRow < 3) {
    SpreadsheetApp.getUi().alert('No data rows found in content sheet.');
    return;
  }
  
  // Create an asset map for quick lookup of row identifier to file ID
  var assetMap = new Map();
  
  // Read data from the Assets sheet if it exists
  if (assetsSheet) {
    var assetLastRow = assetsSheet.getLastRow();
    if (assetLastRow > 1) {  // If there's data beyond the header row
      var assetData = assetsSheet.getRange(2, 1, assetLastRow - 1, 3).getValues();  // Get Project ID, File ID, File Name columns
      
      // Populate the asset map: Project/Row ID -> {fileId, fileName}
      for (var i = 0; i < assetData.length; i++) {
        if (assetData[i][0] && assetData[i][1]) {  // If Project ID and File ID are not empty
          assetMap.set(assetData[i][0].toString(), {
            fileId: assetData[i][1].toString(),
            fileName: assetData[i][2] || "Asset"
          });
        }
      }
    }
  }
  
  // Show a processing dialog
  var html = HtmlService.createHtmlOutput('<p>Processing asset column...</p>')
      .setWidth(250)
      .setHeight(100);
  var processingDialog = SpreadsheetApp.getUi().showModalDialog(html, 'Processing');
  
  // Create a rich text builder for multiple cells
  var richTextBuilder = SpreadsheetApp.newRichTextValue();
  
  // Prepare a range of cells for batch update, starting at row 3 (after headers)
  var assetCells = [];
  
  // Iterate through data rows, skipping header rows 1 and 2
  for (var rowNum = 3; rowNum <= lastRow; rowNum++) {
    // Get row identifier for the current row
    var rowIdentifier = getRowIdentifier(contentSheet, rowNum);
    
    // Get the cell in the asset action column
    var actionCell = contentSheet.getRange(rowNum, ASSET_ACTION_COL_IDX);
    
    // Check if this row has an associated asset
    if (assetMap.has(rowIdentifier)) {
      var assetInfo = assetMap.get(rowIdentifier);
      var fileId = assetInfo.fileId;
      var fileName = assetInfo.fileName;
      
      // Create a rich text value with a "View Asset" link
      var richTextValue = SpreadsheetApp.newRichTextValue()
          .setText("View Asset")
          .setLinkUrl("https://drive.google.com/file/d/" + fileId + "/view")
          .build();
      
      actionCell.setRichTextValue(richTextValue);
    } else {
      // Create a plain text "Link Asset" value (we'll handle the click via onSelectionChange)
      actionCell.setValue("Link Asset");
    }
    
    // Flush changes periodically (every 20 rows) to improve performance
    if (rowNum % 20 === 0) {
      SpreadsheetApp.flush();
    }
  }
  
  // Ensure all changes are applied
  SpreadsheetApp.flush();
  
  // Make sure the onSelectionChange trigger is installed
  ensureSelectionChangeTrigger();
  
  // Close the processing dialog with a completion message
  var html = HtmlService.createHtmlOutput('<p>Asset column has been refreshed.</p><p>Remember to refresh the page if this is the first setup to activate the selection trigger.</p>')
      .setWidth(300)
      .setHeight(120);
  var completionDialog = SpreadsheetApp.getUi().showModalDialog(html, 'Update Complete!');
}

/**
 * Ensures the onSelectionChange trigger is installed.
 * This is critical for making asset cells clickable.
 */
function ensureSelectionChangeTrigger() {
  // Check if the trigger already exists
  var allTriggers = ScriptApp.getProjectTriggers();
  var triggerExists = false;
  
  for (var i = 0; i < allTriggers.length; i++) {
    var trigger = allTriggers[i];
    if (trigger.getEventType() === ScriptApp.EventType.ON_SELECTION_CHANGE && 
        trigger.getHandlerFunction() === 'onSelectionChange') {
      triggerExists = true;
      break;
    }
  }
  
  // Create the trigger if it doesn't exist
  if (!triggerExists) {
    try {
      ScriptApp.newTrigger('onSelectionChange')
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onSelectionChange()
        .create();
      Logger.log('onSelectionChange trigger created successfully.');
    } catch (e) {
      Logger.log('Error creating onSelectionChange trigger: ' + e.toString());
      // Don't alert the user, as this requires authorization that might not be present yet
    }
  }
}

/**
 * Performs an action on the selected asset cell.
 * This function is called from the menu "2. Perform Action on Selected Asset Cell".
 * It checks the cell value and takes appropriate action based on the value.
 */
function performActionOnSelectedAssetCell() {
  // Load configuration
  var configLoaded = loadConfig();
  if (!configLoaded) {
    SpreadsheetApp.getUi().alert('Error: Configuration failed to load. Cannot perform action on asset cell.');
    return;
  }
  
  // Get active cell and active sheet
  var activeCell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();
  var activeSheet = activeCell.getSheet();
  
  // Validate active sheet is the content sheet
  if (activeSheet.getName() !== CONTENT_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('Please select a cell in the "' + CONTENT_SHEET_NAME + '" sheet to perform this action.');
    return;
  }
  
  try {
    // Get header indexes from the active sheet
    getHeaderIndexes(activeSheet);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
    return;
  }
  
  // Validate selected cell is in the asset action column
  var selectedColumn = activeCell.getColumn();
  if (selectedColumn !== ASSET_ACTION_COL_IDX) {
    SpreadsheetApp.getUi().alert('Please select a cell in the "' + ASSET_ACTION_COLUMN_NAME + '" column to perform this action.');
    return;
  }
  
  // Get the cell value, row number, and row identifier
  var cellValue = activeCell.getValue();
  var rowNum = activeCell.getRow();
  var rowIdentifier = getRowIdentifier(activeSheet, rowNum);
  
  // Branch logic based on cell value
  if (cellValue === "Link Asset") {
    // Call the function to show link asset dialog
    Logger.log('Showing link asset dialog for row: ' + rowNum + ', identifier: ' + rowIdentifier);
    showAssignAssetDialog(rowIdentifier, rowNum);
  } else if (cellValue === "View Asset") {
    // Call the function to view asset for row
    Logger.log('Viewing asset for row: ' + rowNum + ', identifier: ' + rowIdentifier);
    viewAssetForRow(rowIdentifier);
  } else {
    // No action defined for this cell value
    SpreadsheetApp.getUi().alert('No action defined for cell value "' + cellValue + '". Expected "Link Asset" or "View Asset".');
  }
}

/**
 * Shows the assign asset dialog for the specified row.
 * 
 * @param {string} rowIdentifier - The unique identifier for the row
 * @param {number} rowNumForDisplay - The row number to display in the dialog
 */
function showAssignAssetDialog(rowIdentifier, rowNumForDisplay) {
  // Create HTML template from AssignAssetDialog.html
  var template = HtmlService.createTemplateFromFile('AssignAssetDialog');
  
  // Pass rowIdentifier and rowNumForDisplay to the template
  template.rowIdentifier = rowIdentifier;
  template.rowNumForDisplay = rowNumForDisplay;
  
  // Evaluate the template to HTML
  var html = template.evaluate()
                     .setWidth(400)
                     .setHeight(500)
                     .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
  // Show as a modal dialog with appropriate title
  SpreadsheetApp.getUi().showModalDialog(html, 'Link Asset for Row ' + rowNumForDisplay);
}

/**
 * Gets the primary asset folder ID from the centralized configuration in Settings sheet (cell B18).
 * This function now uses getPrimaryDriveAssetsFolderId() from api_integrations.js instead 
 * of retrieving from the Config sheet.
 * @private
 * @return {string|null} The primary asset folder ID or null if not configured
 */
function getAssetFolderId_() {
  // Call the new central function from api_integrations.js
  if (typeof getPrimaryDriveAssetsFolderId === 'function') {
    const folderId = getPrimaryDriveAssetsFolderId();
    if (folderId) {
      return folderId;
    } else {
      Logger.log("getAssetFolderId_ in Code.js: Primary Drive Assets Folder ID not configured in Settings (B18) or error retrieving it.");
      // SpreadsheetApp.getUi().alert("Primary Drive Assets Folder ID is not set. Please configure it in the Settings sheet (cell B18) via the Integrations Modal."); // Avoid UI alerts in backend helpers
      return null;
    }
  } else {
    Logger.log("getAssetFolderId_ in Code.js: Critical error - getPrimaryDriveAssetsFolderId function is not available. Make sure api_integrations.js is loaded.");
    // SpreadsheetApp.getUi().alert("Critical error: Drive configuration function missing.");
    return null;
  }
}

/**
 * Updates an existing row or appends a new row to a sheet based on a key column match.
 * @private
 * 
 * @param {Sheet} sheet - The sheet to update
 * @param {Array} rowData - The row data to write
 * @param {number} keyColumnIndex - The index of the column (0-based) to use as a key for matching
 * @return {boolean} True if an existing row was updated, false if a new row was appended
 */
function updateOrAppendRow_(sheet, rowData, keyColumnIndex) {
  if (!sheet || !rowData || keyColumnIndex === undefined) {
    Logger.log('Error: Invalid parameters for updateOrAppendRow_');
    return false;
  }
  
  var keyValue = rowData[keyColumnIndex];
  if (!keyValue) {
    Logger.log('Error: Key value at index ' + keyColumnIndex + ' is empty');
    return false;
  }
  
  // Get all data from the sheet
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {  // If only header row or no data
    // Append the row
    sheet.appendRow(rowData);
    return false;  // No update, new row appended
  }
  
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  
  // Look for matching row based on key column
  for (var i = 0; i < data.length; i++) {
    if (data[i][keyColumnIndex] && data[i][keyColumnIndex].toString() === keyValue.toString()) {
      // Found a match, update the row
      sheet.getRange(i + 2, 1, 1, rowData.length).setValues([rowData]);
      return true;  // Existing row updated
    }
  }
  
  // No match found, append the row
  sheet.appendRow(rowData);
  return false;  // No update, new row appended
}

/**
 * Links an asset file to a row in the content sheet.
 * 
 * @param {string} rowIdentifier - The unique identifier for the row
 * @param {string} fileId - The ID of the file to link
 * @param {string} fileName - The name of the file
 * @return {Object} Success status and details of the linking
 */
function linkAssetToRow(rowIdentifier, fileId, fileName) {
  try {
    // Set up the sheet data record
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var assetsSheet = ensureAssetsSheetExists();
    
    // Prepare the row data
    var rowData = [rowIdentifier, fileId, fileName, new Date()];
    
    // We need to find the Project ID column index in the Assets sheet
    // For simplicity, assuming it's the first column (index 0)
    var projectIdColIdx = 0;
    
    // Update or append the row in the Assets sheet
    var updated = updateOrAppendRow_(assetsSheet, rowData, projectIdColIdx);
    
    // Update the asset action cell in the content sheet
    updateAssetActionCell(rowIdentifier, fileId, fileName);
    
    // Return success with details
    return {
      success: true,
      fileId: fileId,
      fileName: fileName,
      rowIdentifier: rowIdentifier,
      updated: updated // true if an existing row was updated, false if a new row was appended
    };
  } catch (e) {
    Logger.log("Error linking asset to row: " + e.toString());
    return {
      success: false,
      error: "Failed to link asset to row: " + e.toString()
    };
  }
}

/**
 * Updates the asset action cell in the content sheet for a specified row.
 * Changes the cell from "Link Asset" to "View Asset" with a link to the Google Drive file.
 * 
 * @param {string} rowIdentifier - The unique identifier for the row
 * @param {string} fileId - The ID of the file to link
 * @param {string} fileName - The name of the file for display
 * @return {boolean} True if the cell was updated successfully
 */
function updateAssetActionCell(rowIdentifier, fileId, fileName) {
  try {
    if (!loadConfig()) return false;
    
    // Get the content sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var contentSheet = ss.getSheetByName(CONTENT_SHEET_NAME);
    if (!contentSheet) return false;
    
    try {
      getHeaderIndexes(contentSheet);
    } catch (e) {
      Logger.log("Error getting header indexes: " + e.toString());
      return false;
    }
    
    // Find the row with the matching identifier
    var rowNum = 0;
    var lastRow = contentSheet.getLastRow();
    
    // If we have a row ID column, use it to find the row
    if (ROW_ID_COL_IDX > 0) {
      var idValues = contentSheet.getRange(3, ROW_ID_COL_IDX, lastRow - 2, 1).getValues();
      for (var i = 0; i < idValues.length; i++) {
        if (idValues[i][0] && idValues[i][0].toString() === rowIdentifier.toString()) {
          rowNum = i + 3; // +3 because we start at row 3 (after headers)
          break;
        }
      }
    }
    
    // If we found the row, update the asset action cell
    if (rowNum > 0) {
      var actionCell = contentSheet.getRange(rowNum, ASSET_ACTION_COL_IDX);
      
      // Create a rich text value with a "View Asset" link
      var richTextValue = SpreadsheetApp.newRichTextValue()
          .setText("View Asset")
          .setLinkUrl("https://drive.google.com/file/d/" + fileId + "/view")
          .build();
      
      actionCell.setRichTextValue(richTextValue);
      return true;
    }
    
    return false;
  } catch (e) {
    Logger.log("Error updating asset action cell: " + e.toString());
    return false;
  }
}

/**
 * Uploads a file to Google Drive from a base64 data string.
 * Uses the primary asset folder ID from the Settings sheet (B18).
 * 
 * @param {string} fileDataString - The base64 data string of the file (including mime type)
 * @param {string} fileName - The name to give the uploaded file
 * @param {string} rowIdentifier - The unique identifier for the row to link the file to
 * @return {Object} Success status and details of the uploaded file
 */
function uploadFileToDrive(fileDataString, fileName, rowIdentifier) {
  try {
    // Load configuration
    var configLoaded = loadConfig();
    if (!configLoaded) {
      return {
        success: false,
        message: "Configuration failed to load. Cannot upload file."
      };
    }
    
    // Get the primary asset folder ID from the centralized settings
    var assetFolderId = getAssetFolderId_();
    if (!assetFolderId) {
      return {
        success: false,
        message: "Asset folder ID not configured."
      };
    }
    
    // Get the folder
    var folder;
    try {
      folder = DriveApp.getFolderById(assetFolderId);
    } catch (folderError) {
      return {
        success: false,
        message: "Asset folder not found or not accessible: " + folderError.toString()
      };
    }
    
    // Parse the file data string (format: data:mimeType;base64,base64Data)
    var matches = fileDataString.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/);
    if (!matches || matches.length !== 3) {
      return {
        success: false,
        message: "Invalid file data format."
      };
    }
    
    // Extract mime type and base64 data
    var mimeType = matches[1];
    var base64Data = matches[2];
    
    // Decode the base64 data
    var decodedData;
    try {
      decodedData = Utilities.base64Decode(base64Data);
    } catch (decodeError) {
      return {
        success: false,
        message: "Failed to decode file data: " + decodeError.toString()
      };
    }
    
    // Create a blob from the decoded data
    var blob = Utilities.newBlob(decodedData, mimeType, fileName);
    
    // Upload the file to the folder
    var newFile;
    try {
      newFile = folder.createFile(blob);
    } catch (uploadError) {
      return {
        success: false,
        message: "Failed to upload file to Drive: " + uploadError.toString()
      };
    }
    
    // Get the new file's ID and name
    var fileId = newFile.getId();
    var newFileName = newFile.getName();
    
    // Link the asset to the row
    var linkResult = linkAssetToRow(rowIdentifier, fileId, newFileName);
    if (!linkResult.success) {
      return {
        success: false,
        message: "File uploaded but failed to link to row: " + linkResult.error
      };
    }
    
    // Return success
    return {
      success: true,
      message: "File uploaded and linked successfully.",
      fileId: fileId,
      fileName: newFileName,
      fileUrl: newFile.getUrl()
    };
  } catch (e) {
    Logger.log("Error uploading file to Drive: " + e.toString());
    return {
      success: false,
      message: "Failed to upload file: " + e.toString()
    };
  }
}

/**
 * Associates an existing file in the asset folder with a project.
 * This function is callable from the frontend via google.script.run.
 * 
 * @param {string} projectId - The project ID to associate the file with
 * @param {string} fileIdToAssociate - The ID of the file to associate
 * @param {string} fileNameToAssociate - The name of the file to associate
 * @return {Object} Success status and association details
 */
function associateExistingAsset(projectId, fileIdToAssociate, fileNameToAssociate) {
  try {
    // Validate inputs
    if (!projectId) {
      return {
        success: false,
        error: "No project ID provided"
      };
    }
    
    if (!fileIdToAssociate) {
      return {
        success: false,
        error: "No file ID provided"
      };
    }
    
    if (!fileNameToAssociate) {
      return {
        success: false,
        error: "No file name provided"
      };
    }
    
    // Verify the file exists
    try {
      var file = DriveApp.getFileById(fileIdToAssociate);
      // If the file doesn't exist, an error will be thrown above
    } catch (fileError) {
      return {
        success: false,
        error: "File not found or not accessible: " + fileError.toString()
      };
    }
    
    // Set up the sheet data record
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var assetsSheet = ensureAssetsSheetExists();
    
    // Prepare the row data
    var rowData = [projectId, fileIdToAssociate, fileNameToAssociate, new Date()];
    
    // Update or append the row in the Assets sheet
    var projectIdColIdx = 0;  // Assuming Project ID is in first column
    var updated = updateOrAppendRow_(assetsSheet, rowData, projectIdColIdx);
    
    // Update the asset action cell in the content sheet
    updateAssetActionCell(projectId, fileIdToAssociate, fileNameToAssociate);
    
    // Return success with details
    return {
      success: true,
      fileId: fileIdToAssociate,
      fileName: fileNameToAssociate,
      projectId: projectId,
      updated: updated, // true if an existing row was updated, false if a new row was appended
      fileUrl: file.getUrl()
    };
  } catch (e) {
    // An error occurred, log it and return error information
    Logger.log("Error associating existing asset: " + e.toString());
    return {
      success: false,
      error: "Failed to associate existing asset: " + e.toString()
    };
  }
}

/**
 * Public wrapper for getting asset details from a project.
 * This function is callable from the frontend via google.script.run.
 * 
 * @param {string} projectId - The ID of the project to get assets for
 * @return {Object} Asset details object or null if no assets found
 */
function getAssetDetails(projectId) {
  try {
    // Call the private implementation function to get the asset details
    var assetDetails = getAssetDetailsForProject_(projectId);
    
    if (assetDetails) {
      // If we found asset details, return them with a success status
      return {
        success: true,
        assetDetails: assetDetails
      };
    } else {
      // No asset found for this project, return a null result with a message
      return {
        success: true,
        assetDetails: null,
        message: "No asset found for project ID: " + projectId
      };
    }
  } catch (e) {
    // An error occurred, log it and return error information
    Logger.log("Error in getAssetDetails: " + e.toString());
    return {
      success: false,
      error: "Failed to retrieve asset details: " + e.toString()
    };
  }
}

/**
 * Lists all files in the assets folder.
 * This function is callable from the frontend via google.script.run.
 * Uses the primary asset folder ID from the Settings sheet (B18).
 * 
 * @return {Object} Success status and array of file information
 */
function listAssetFolderFiles() {
  try {
    // Get the primary asset folder ID from the centralized settings
    var assetFolderId = getAssetFolderId_();
    if (!assetFolderId) {
      return {
        success: false,
        error: "Asset folder ID not configured"
      };
    }
    
    // Get the folder and its files
    var folder = DriveApp.getFolderById(assetFolderId);
    var files = folder.getFiles();
    
    // Create an array to hold file information
    var fileList = [];
    
    // Loop through the files and add their details to the array
    while (files.hasNext()) {
      var file = files.next();
      fileList.push({
        id: file.getId(),
        name: file.getName(),
        type: file.getMimeType(),
        url: file.getUrl(),
        dateCreated: file.getDateCreated(),
        lastUpdated: file.getLastUpdated(),
        size: file.getSize(),
        thumbnailUrl: file.getThumbnail() ? file.getThumbnail().getDataAsString() : null
      });
    }
    
    // Return the file list with success status
    return {
      success: true,
      files: fileList,
      count: fileList.length,
      folderUrl: folder.getUrl()
    };
  } catch (e) {
    // An error occurred, log it and return error information
    Logger.log("Error listing asset folder files: " + e.toString());
    return {
      success: false,
      error: "Failed to list asset folder files: " + e.toString()
    };
  }
}

/**
 * Uploads a file to the asset folder and associates it with a project.
 * This function is callable from the frontend via google.script.run.
 * Uses the primary asset folder ID from the Settings sheet (B18).
 * 
 * @param {Object} fileObject - The file object (blob) to upload
 * @param {string} projectId - The project ID to associate the file with
 * @return {Object} Success status and uploaded file details
 */
function uploadAndAssociateAsset(fileObject, projectId) {
  try {
    // Validate inputs
    if (!fileObject) {
      return {
        success: false,
        error: "No file provided"
      };
    }
    
    if (!projectId) {
      return {
        success: false,
        error: "No project ID provided"
      };
    }
    
    // Get the primary asset folder ID from the centralized settings
    var assetFolderId = getAssetFolderId_();
    if (!assetFolderId) {
      return {
        success: false,
        error: "Asset folder ID not configured"
      };
    }
    
    // Get the folder
    var folder = DriveApp.getFolderById(assetFolderId);
    
    // Upload the file to the folder
    var uploadedFile = folder.createFile(fileObject);
    var fileId = uploadedFile.getId();
    var fileName = uploadedFile.getName();
    
    // Set up the sheet data record
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var assetsSheet = ensureAssetsSheetExists();
    
    // Prepare the row data
    var rowData = [projectId, fileId, fileName, new Date()];
    
    // Update or append the row in the Assets sheet
    var projectIdColIdx = 0;  // Assuming Project ID is in first column
    var updated = updateOrAppendRow_(assetsSheet, rowData, projectIdColIdx);
    
    // Return success with details
    return {
      success: true,
      fileId: fileId,
      fileName: fileName,
      projectId: projectId,
      updated: updated, // true if an existing row was updated, false if a new row was appended
      fileUrl: uploadedFile.getUrl()
    };
  } catch (e) {
    // An error occurred, log it and return error information
    Logger.log("Error uploading and associating asset: " + e.toString());
    return {
      success: false,
      error: "Failed to upload and associate asset: " + e.toString()
    };
  }
}

/**
 * Ensures that the Assets sheet exists with proper headers.
 * @return {Sheet} The Assets sheet instance.
 */
function ensureAssetsSheetExists() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var assetsSheet = ss.getSheetByName(ASSETS_SHEET_NAME);
  if (!assetsSheet) {
    assetsSheet = ss.insertSheet(ASSETS_SHEET_NAME);
    assetsSheet.appendRow(["Project ID", "File ID", "File Name", "Upload Date"]);
  }
  return assetsSheet;
}

/**
 * Triggered when the user changes selection in the spreadsheet.
 * Opens the asset dialog when selecting a cell in the asset column.
 *
 * @param {GoogleAppsScript.Events.SheetsOnSelectionChange} e Event object.
 */
function onSelectionChange(e) {
  try {
    var sheet = e.range.getSheet();
    if (sheet.getName() !== CONTENT_SHEET_NAME) return;
    
    // Load configuration if needed
    if (!ASSET_ACTION_COL_IDX) {
      if (!loadConfig()) return;
      try {
        getHeaderIndexes(sheet);
      } catch (err) {
        return;
      }
    }
    
    var row = e.range.getRow();
    if (row < 3) return; // skip headers
    
    var column = e.range.getColumn();
    if (column !== ASSET_ACTION_COL_IDX) return;
    
    // Get the cell value
    var cell = sheet.getRange(row, column);
    var cellValue = cell.getValue();
    var rowIdentifier = getRowIdentifier(sheet, row);
    
    if (cellValue === "Link Asset") {
      Logger.log("Showing asset dialog for row " + row + " with ID: " + rowIdentifier);
      showAssignAssetDialog(rowIdentifier, row);
    } else if (cellValue === "View Asset") {
      Logger.log("Viewing asset for row " + row + " with ID: " + rowIdentifier);
      viewAssetForRow(rowIdentifier);
    }
  } catch (err) {
    Logger.log("Error in onSelectionChange: " + err.toString());
  }
}

/**
 * Views the asset associated with a specific row identifier.
 * Opens the Google Drive file in a new browser tab.
 * 
 * @param {string} rowIdentifier - The unique identifier for the row
 */
function viewAssetForRow(rowIdentifier) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var assetsSheet = ss.getSheetByName(ASSETS_SHEET_NAME);
    
    if (!assetsSheet) {
      SpreadsheetApp.getUi().alert("Assets sheet not found.");
      return;
    }
    
    // Search for the asset record
    var assetData = assetsSheet.getDataRange().getValues();
    var fileId = null;
    
    // Start from row 1 (which is the header) to get the correct index
    for (var i = 1; i < assetData.length; i++) {
      if (assetData[i][0] && assetData[i][0].toString() === rowIdentifier.toString()) {
        fileId = assetData[i][1];
        break;
      }
    }
    
    if (!fileId) {
      SpreadsheetApp.getUi().alert("No asset found for this row. Try linking an asset first.");
      return;
    }
    
    // Create a dialog to open the file
    var html = HtmlService.createHtmlOutput(
      '<script>' +
      'window.open("https://drive.google.com/file/d/' + fileId + '/view", "_blank");' +
      'google.script.host.close();' +
      '</script>'
    )
    .setWidth(1)
    .setHeight(1);
    
    SpreadsheetApp.getUi().showModalDialog(html, "Opening asset...");
    
  } catch (e) {
    Logger.log("Error viewing asset: " + e.toString());
    SpreadsheetApp.getUi().alert("Error viewing asset: " + e.toString());
  }
}
