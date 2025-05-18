/**
 * AssetService.gs
 * Provides functions for working with Assets in the sheet and Google Drive.
 */

// Constants for working with the Assets sheet
var ASSETS_SHEET_NAME = "Assets";
var PROJECT_ID_COL_ASSETS = 0;  // Column A - Project ID that the asset is associated with
var FILE_ID_COL_ASSETS = 1;     // Column B - Drive File ID of the asset
var FILE_NAME_COL_ASSETS = 2;   // Column C - Name of the asset file

/**
 * Retrieves asset details for a specific project from the Assets sheet.
 * 
 * @param {string} projectId - The project ID to find assets for
 * @return {Object} An object containing asset details or null if not found
 * @private
 */
function getAssetDetailsForProject_(projectId) {
  try {
    // Get the spreadsheet and Assets sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var assetsSheet = ss.getSheetByName(ASSETS_SHEET_NAME);
    
    // Check if the Assets sheet exists
    if (!assetsSheet) {
      Logger.log("Assets sheet not found. No asset details available.");
      return null;
    }
    
    // Find any rows matching the projectId
    var matchingRows = findRowByValue_(assetsSheet, PROJECT_ID_COL_ASSETS, projectId);
    
    // If no matching rows found, return null
    if (matchingRows.length === 0) {
      Logger.log("No assets found for project: " + projectId);
      return null;
    }
    
    // Take the first matching row (assuming one asset per project)
    var assetRow = matchingRows[0];
    
    // Extract and return the asset details
    return {
      projectId: assetRow[PROJECT_ID_COL_ASSETS],
      fileId: assetRow[FILE_ID_COL_ASSETS],
      fileName: assetRow[FILE_NAME_COL_ASSETS]
    };
  } catch (e) {
    Logger.log("Error retrieving asset details: " + e.toString());
    return null;
  }
}

/**
 * Test function for getAssetDetailsForProject_
 */
function testGetAssetDetails() {
  // Test with a project ID that should have an asset
  var projectWithAsset = "project123"; // Replace with an actual test project ID
  var assetDetails = getAssetDetailsForProject_(projectWithAsset);
  Logger.log("Asset details for " + projectWithAsset + ": " + JSON.stringify(assetDetails));
  
  // Test with a project ID that shouldn't have an asset
  var projectWithoutAsset = "nonexistent-project"; // A project ID that shouldn't exist
  var noAssetDetails = getAssetDetailsForProject_(projectWithoutAsset);
  Logger.log("Asset details for " + projectWithoutAsset + ": " + JSON.stringify(noAssetDetails));
}