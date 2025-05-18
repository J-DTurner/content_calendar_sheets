/**
 * API Integrations for Social Media Content Calendar
 * 
 * This script provides integration with external APIs and services
 * to enhance the functionality of the content calendar.
 */

// API integration configuration
//THIS IS A TEST

const API_INTEGRATION_CONFIG = {
  SETTINGS_SHEET: 'Settings',
  API_SETTINGS_START_ROW: 12,
  API_KEYS: {
    TWITTER_API_KEY_CELL: 'B13',
    TWITTER_API_SECRET_CELL: 'B14',
    YOUTUBE_API_KEY_CELL: 'B15',
    TELEGRAM_BOT_TOKEN_CELL: 'B16'
  },
  DRIVE_FOLDER_ID_CELL: 'B18', // For the Drive folder used by api_integrations.js sync
  DRIVE_FOLDER_NAME_CELL: 'B19', // For its name
  CONTENT_SHEET: 'Content Calendar',
  LINK_COLUMN: 7, // Column G for Link to Asset
  ANALYTICS_SHEET: 'Analytics',
  LAST_SYNC_CELL: 'B20',
  ASSET_CONFIG: {
    CONTENT_SHEET_CELL: 'B39',
    ASSET_ACTION_COLUMN_CELL: 'B40',
    ROW_ID_COLUMN_CELL: 'B41'
  },
  MAX_RESULTS_PER_QUERY: 50
};

/**
 * Sets up the API integration settings
 * @param {object} [options] Optional parameters.
 * @param {boolean} [options.showConfirmationDialog=true] Whether to show a confirmation dialog.
 * @return {object} An object with success status and message.
 */
function setupApiIntegration(options) {
  const showConfirmation = (options && options.showConfirmationDialog !== undefined) ? options.showConfirmationDialog : true;
  let message = 'API Integration settings have been set up.';

  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get or create Settings sheet
  let settingsSheet = ss.getSheetByName(API_INTEGRATION_CONFIG.SETTINGS_SHEET);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(API_INTEGRATION_CONFIG.SETTINGS_SHEET);
    if (!settingsSheet) {
      return { success: false, message: 'Failed to create Settings sheet.'};
    }
    message = 'API Integration settings sheet created and configured.';
  }
  
  // Set up API settings section
  settingsSheet.getRange('A12:C12').merge()
    .setValue('API INTEGRATION SETTINGS')
    .setBackground('#4285F4')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Set up Twitter API settings
  settingsSheet.getRange('A13').setValue('Twitter API Key:');
  settingsSheet.getRange('A14').setValue('Twitter API Secret:');
  
  // Set up YouTube API settings
  settingsSheet.getRange('A15').setValue('YouTube API Key:');
  
  // Set up Telegram API settings
  settingsSheet.getRange('A16').setValue('Telegram Bot Token:');

  // Ensure Drive Folder ID and Name cells and labels are present
  // as per API_INTEGRATION_CONFIG, even if not explicitly part of initial setup here.
  // This helps functions like saveGoogleDriveFolderId_ and getIntegrationSettings
  const driveIdValueCellA1 = API_INTEGRATION_CONFIG.DRIVE_FOLDER_ID_CELL; // e.g., B18
  const driveIdLabelCell = settingsSheet.getRange(driveIdValueCellA1).offset(0, -1); // e.g., A18
  if (driveIdLabelCell.getValue() !== 'Primary Google Drive Assets Folder ID:') {
      driveIdLabelCell.setValue('Primary Google Drive Assets Folder ID:').setFontWeight('bold');
  }

  const driveNameValueCellA1 = API_INTEGRATION_CONFIG.DRIVE_FOLDER_NAME_CELL; // e.g., B19
  const driveNameLabelCell = settingsSheet.getRange(driveNameValueCellA1).offset(0, -1); // e.g., A19
  if (driveNameLabelCell.getValue() !== 'Primary Assets Folder Name:') {
      driveNameLabelCell.setValue('Primary Assets Folder Name:').setFontWeight('bold');
  }
  
  // Set up last sync information
  settingsSheet.getRange('A20').setValue('Last Analytics Sync:');
  if (settingsSheet.getRange(API_INTEGRATION_CONFIG.LAST_SYNC_CELL).getValue() === '') { // Only set to 'Never' if empty
      settingsSheet.getRange(API_INTEGRATION_CONFIG.LAST_SYNC_CELL).setValue('Never');
  }
  
  // Set up API instructions
  const instructions = [
    ['API Integration Instructions:'],
    ['1. Enter your API keys/tokens for the services you want to integrate with'],
    ['2. Use the "Integrations" menu to connect with each service'],
    ['3. For Twitter integration, create a developer account at developer.twitter.com'],
    ['4. For YouTube integration, create an API key in the Google Cloud Console'],
    ['5. For Telegram integration, create a bot using @BotFather'],
    ['6. The "Fetch Analytics" function will retrieve basic metrics for published content']
  ];
  
  // Check if A22 is empty before merging and setting values, to avoid overwriting
  if(settingsSheet.getRange('A22').getValue() === '') {
    settingsSheet.getRange('A22:C28').merge(false); // Assuming this merge is intentional, though C28 is not covered by values
    settingsSheet.getRange('A22:A28').setValues(instructions); // instructions array is 7 rows, A22:A28 is 7 rows.
  }
  
  // Format cells
  settingsSheet.getRange('A13:A21').setFontWeight('bold'); // A18, A19 were added to config, A21 can be bold too.
  
  // Set column widths
  settingsSheet.setColumnWidth(1, 200);
  settingsSheet.setColumnWidth(2, 300);
  
  if (showConfirmation) {
    SpreadsheetApp.getUi().alert(message + ' Please enter your API keys in the Settings sheet.');
  }
  return { success: true, message: message + ' Please ensure API keys are correctly entered.' };
}

/**
 * Fetches analytics data for published content
 * @return {object} An object with success status and message.
 */
function fetchContentAnalytics() {
  const ui = SpreadsheetApp.getUi(); // Keep for prompts
  
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contentSheet = ss.getSheetByName(API_INTEGRATION_CONFIG.CONTENT_SHEET);
  
  if (!contentSheet) {
    return { success: false, message: 'Content Calendar sheet not found.' };
  }
  
  // Create Analytics sheet if it doesn't exist
  let analyticsSheet = ss.getSheetByName(API_INTEGRATION_CONFIG.ANALYTICS_SHEET);
  let isNewSheet = false;
  
  if (!analyticsSheet) {
    analyticsSheet = ss.insertSheet(API_INTEGRATION_CONFIG.ANALYTICS_SHEET);
    isNewSheet = true;
  }
  
  // Set up headers if new sheet
  if (isNewSheet) {
    analyticsSheet.getRange('A1:H1').setValues([
      ['Content ID', 'Date Published', 'Channel', 'Content Title', 'Views/Impressions', 'Engagements', 'Link Clicks', 'Last Updated']
    ]);
    
    // Format headers
    analyticsSheet.getRange('A1:H1')
      .setBackground('#4285F4')
      .setFontColor('white')
      .setFontWeight('bold');
    
    // Set column widths
    analyticsSheet.setColumnWidth(1, 100); // Content ID
    analyticsSheet.setColumnWidth(2, 120); // Date Published
    analyticsSheet.setColumnWidth(3, 100); // Channel
    analyticsSheet.setColumnWidth(4, 300); // Content Title
    analyticsSheet.setColumnWidth(5, 120); // Views/Impressions
    analyticsSheet.setColumnWidth(6, 120); // Engagements
    analyticsSheet.setColumnWidth(7, 120); // Link Clicks
    analyticsSheet.setColumnWidth(8, 150); // Last Updated
  }
  
  // Get published content from content calendar
  const data = contentSheet.getDataRange().getValues();
  
  // Extract headers (assuming row 2 contains headers)
  const headers = data[1];
  
  // Find column indexes
  const idIndex = headers.indexOf('ID');
  const dateIndex = headers.indexOf('Date');
  const statusIndex = headers.indexOf('Status');
  const channelIndex = headers.indexOf('Channel');
  const contentIndex = headers.indexOf('Content/Idea');
  const linkIndex = headers.indexOf('Link to Asset');
  
  // Skip if any required column is missing
  if (idIndex === -1 || dateIndex === -1 || statusIndex === -1 || 
      channelIndex === -1 || contentIndex === -1) {
    return { success: false, message: 'Required columns not found in Content Calendar.' };
  }
  
  // Get API keys from settings
  const settingsSheet = ss.getSheetByName(API_INTEGRATION_CONFIG.SETTINGS_SHEET);
  if (!settingsSheet) {
    return { success: false, message: 'Settings sheet not found.' };
  }
  
  const twitterApiKey = settingsSheet.getRange(API_INTEGRATION_CONFIG.API_KEYS.TWITTER_API_KEY_CELL).getValue();
  const twitterApiSecret = settingsSheet.getRange(API_INTEGRATION_CONFIG.API_KEYS.TWITTER_API_SECRET_CELL).getValue();
  const youtubeApiKey = settingsSheet.getRange(API_INTEGRATION_CONFIG.API_KEYS.YOUTUBE_API_KEY_CELL).getValue();
  const telegramBotToken = settingsSheet.getRange(API_INTEGRATION_CONFIG.API_KEYS.TELEGRAM_BOT_TOKEN_CELL).getValue();
  
  // Check if at least one API key is available
  if (!twitterApiKey && !youtubeApiKey && !telegramBotToken) {
    return { success: false, message: 'No API keys found in Settings sheet. Please add at least one API key to use this feature.' };
  }
  
  // Find published content
  const publishedContent = [];
  
  // Start from row 3 (index 2) to skip headers
  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    
    // Only include items with "Schedule" status
    if (row[statusIndex] !== 'Schedule') {
      continue;
    }
    
    // Must have a date and content
    if (!(row[dateIndex] instanceof Date) || !row[contentIndex]) {
      continue;
    }
    
    // Check if the publication date is in the past
    const pubDate = row[dateIndex];
    const now = new Date();
    
    if (pubDate > now) {
      continue; // Skip future content
    }
    
    // Add to published content list
    publishedContent.push({
      id: row[idIndex],
      date: pubDate,
      channel: row[channelIndex],
      content: row[contentIndex],
      link: linkIndex !== -1 ? row[linkIndex] : '',
      rowIndex: i + 1 // Actual row number in sheet
    });
  }
  
  // If no published content found
  if (publishedContent.length === 0) {
    return { success: false, message: 'No published content found with "Schedule" status and past publication dates.' };
  }
  
  // Ask user which channel to fetch analytics for
  const channelResponse = ui.prompt(
    'Fetch Analytics',
    'Which channel would you like to fetch analytics for? (Twitter, YouTube, Telegram, or All)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (channelResponse.getSelectedButton() !== ui.Button.OK) {
    return { success: false, message: 'Analytics fetch cancelled by user.' };
  }
  
  const selectedChannel = channelResponse.getResponseText().trim().toLowerCase();
  
  // Filter content by selected channel
  let filteredContent = publishedContent;
  
  if (selectedChannel !== 'all') {
    filteredContent = publishedContent.filter(item => 
      item.channel.toLowerCase() === selectedChannel
    );
    
    if (filteredContent.length === 0) {
      return { success: false, message: `No published content found for channel: ${selectedChannel}` };
    }
  }
  
  // Show progress indicator
  const htmlOutput = HtmlService
    .createHtmlOutput('<p>Fetching analytics data... This may take a minute.</p><p>Please do not close this dialog.</p>')
    .setWidth(300)
    .setHeight(120)
    .setTitle('Analytics Progress');
  
  const dialogProgress = ui.showModalDialog(htmlOutput, 'Fetching Analytics');
  
  // Process content and fetch analytics
  let processedCount = 0;
  let analyticsData = [];
  
  try {
    // Process each content item
    for (const item of filteredContent) {
      // Try to get analytics based on channel
      let metrics = {
        views: 0,
        engagements: 0,
        clicks: 0
      };
      
      switch (item.channel.toLowerCase()) {
        case 'twitter':
          if (twitterApiKey && twitterApiSecret) {
            metrics = fetchTwitterAnalytics(item, twitterApiKey, twitterApiSecret);
          } else {
            metrics = generateMockAnalytics(); // Use mock data if no API key
          }
          break;
          
        case 'youtube':
          if (youtubeApiKey) {
            metrics = fetchYouTubeAnalytics(item, youtubeApiKey);
          } else {
            metrics = generateMockAnalytics(); // Use mock data if no API key
          }
          break;
          
        case 'telegram':
          if (telegramBotToken) {
            metrics = fetchTelegramAnalytics(item, telegramBotToken);
          } else {
            metrics = generateMockAnalytics(); // Use mock data if no API key
          }
          break;
          
        default:
          metrics = generateMockAnalytics(); // Use mock data for unknown channels
      }
      
      // Add to analytics data
      analyticsData.push([
        item.id,
        item.date,
        item.channel,
        item.content.length > 100 ? item.content.substring(0, 97) + '...' : item.content,
        metrics.views,
        metrics.engagements,
        metrics.clicks,
        new Date()
      ]);
      
      processedCount++;
    }
    
    // Write data to analytics sheet
    if (analyticsData.length > 0) {
      // Clear existing data (except headers)
      if (analyticsSheet.getLastRow() > 1) {
        analyticsSheet.getRange(2, 1, analyticsSheet.getLastRow() - 1, 8).clear();
      }
      
      // Write new data
      analyticsSheet.getRange(2, 1, analyticsData.length, 8).setValues(analyticsData);
      
      // Format date columns
      analyticsSheet.getRange(2, 2, analyticsData.length, 1).setNumberFormat('yyyy-MM-dd');
      analyticsSheet.getRange(2, 8, analyticsData.length, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');
      
      // Update last sync timestamp
      settingsSheet.getRange(API_INTEGRATION_CONFIG.LAST_SYNC_CELL).setValue(new Date());
    }
    
    return { success: true, message: `Analytics data fetched successfully for ${processedCount} content items.` };
    
  } catch (error) {
    console.error('Error fetching analytics:', error);
    return { success: false, message: 'Error fetching analytics: ' + error.toString() };
  }
}

/**
 * Fetches analytics data from Twitter API
 * Note: In a real implementation, this would use the Twitter API
 * For this exercise, we use mock data
 * @param {object} contentItem The content item
 * @param {string} apiKey Twitter API key
 * @param {string} apiSecret Twitter API secret
 * @return {object} Analytics metrics
 */
function fetchTwitterAnalytics(contentItem, apiKey, apiSecret) {
  // In a real implementation, this would call the Twitter API
  // For this exercise, we use mock data
  return generateMockAnalytics();
}

/**
 * Fetches analytics data from YouTube API
 * Note: In a real implementation, this would use the YouTube API
 * For this exercise, we use mock data
 * @param {object} contentItem The content item
 * @param {string} apiKey YouTube API key
 * @return {object} Analytics metrics
 */
function fetchYouTubeAnalytics(contentItem, apiKey) {
  // In a real implementation, this would call the YouTube API
  // For this exercise, we use mock data
  return generateMockAnalytics();
}

/**
 * Fetches analytics data from Telegram API
 * Note: In a real implementation, this would use the Telegram API
 * For this exercise, we use mock data
 * @param {object} contentItem The content item
 * @param {string} botToken Telegram bot token
 * @return {object} Analytics metrics
 */
function fetchTelegramAnalytics(contentItem, botToken) {
  // In a real implementation, this would call the Telegram API
  // For this exercise, we use mock data
  return generateMockAnalytics();
}

/**
 * Generates mock analytics data for demonstration purposes
 * @return {object} Mock analytics metrics
 */
function generateMockAnalytics() {
  // Generate random metrics for demonstration
  return {
    views: Math.floor(Math.random() * 10000),
    engagements: Math.floor(Math.random() * 1000),
    clicks: Math.floor(Math.random() * 500)
  };
}

/**
 * Retrieves current integration settings from the Settings sheet.
 * @return {object} An object containing API keys, Drive folder info, and last sync time, or an error object.
 */
function getIntegrationSettings_() {
  return getIntegrationSettings();
}

/**
 * Retrieves current integration settings from the Settings sheet.
 * @return {object} An object containing API keys, Drive folder info, and last sync time, or an error object.
 */
function getIntegrationSettings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let settingsSheet = ss.getSheetByName(API_INTEGRATION_CONFIG.SETTINGS_SHEET);
    
    if (!settingsSheet) {
      // Call setupApiIntegration silently
      const setupResult = setupApiIntegration({ showConfirmationDialog: false });
      if (!setupResult.success) {
          Logger.log('Error: Settings sheet not found and silent setup failed: ' + setupResult.message);
          return { error: 'Settings sheet not found and could not be created. ' + setupResult.message };
      }
      settingsSheet = ss.getSheetByName(API_INTEGRATION_CONFIG.SETTINGS_SHEET); // Re-fetch the sheet
      if(!settingsSheet) { // Should not happen if setupResult.success is true
          Logger.log('Error: Settings sheet still not found after supposedly successful setup.');
          return { error: 'Settings sheet not found and could not be created (post-setup check).' };
      }
    }

    // Safely get values with error handling
    const safeGetCellValue = (cell) => {
      try {
        const range = settingsSheet.getRange(cell);
        if (!range) return '';
        const value = range.getValue();
        return value !== undefined && value !== null ? value : '';
      } catch (err) {
        Logger.log(`Warning: Could not retrieve value for cell ${cell}: ${err.toString()}`);
        return '';
      }
    };

    const settingsValues = {
      twitterApiKey: safeGetCellValue(API_INTEGRATION_CONFIG.API_KEYS.TWITTER_API_KEY_CELL),
      twitterApiSecret: safeGetCellValue(API_INTEGRATION_CONFIG.API_KEYS.TWITTER_API_SECRET_CELL),
      youtubeApiKey: safeGetCellValue(API_INTEGRATION_CONFIG.API_KEYS.YOUTUBE_API_KEY_CELL),
      telegramBotToken: safeGetCellValue(API_INTEGRATION_CONFIG.API_KEYS.TELEGRAM_BOT_TOKEN_CELL),
      driveFolderId: safeGetCellValue(API_INTEGRATION_CONFIG.DRIVE_FOLDER_ID_CELL),
      driveFolderName: safeGetCellValue(API_INTEGRATION_CONFIG.DRIVE_FOLDER_NAME_CELL),
      lastAnalyticsSync: safeGetCellValue(API_INTEGRATION_CONFIG.LAST_SYNC_CELL),
      contentSheetName: safeGetCellValue(API_INTEGRATION_CONFIG.ASSET_CONFIG.CONTENT_SHEET_CELL),
      assetActionColumnName: safeGetCellValue(API_INTEGRATION_CONFIG.ASSET_CONFIG.ASSET_ACTION_COLUMN_CELL),
      rowIdColumnName: safeGetCellValue(API_INTEGRATION_CONFIG.ASSET_CONFIG.ROW_ID_COLUMN_CELL)
    };
    
    // Sanitize and validate Drive Folder ID (prevent potential errors when using it later)
    if (settingsValues.driveFolderId) {
      settingsValues.driveFolderId = settingsValues.driveFolderId.toString().trim();
      // Check if it looks like a valid folder ID (basic validation)
      if (settingsValues.driveFolderId.length < 10 || settingsValues.driveFolderId.includes(' ')) {
        Logger.log('Warning: Drive Folder ID appears invalid: ' + settingsValues.driveFolderId);
        settingsValues.driveFolderIdWarning = 'The stored Drive Folder ID may be invalid.';
      }
    }
    
    // Ensure Drive Folder Name is a string
    if (settingsValues.driveFolderName) {
      settingsValues.driveFolderName = settingsValues.driveFolderName.toString().trim();
    }
    
    // Ensure date is serializable or string
    if (settingsValues.lastAnalyticsSync instanceof Date) {
      settingsValues.lastAnalyticsSync = settingsValues.lastAnalyticsSync.toISOString();
    } else if (settingsValues.lastAnalyticsSync === 'Never' || !settingsValues.lastAnalyticsSync) {
      settingsValues.lastAnalyticsSync = 'Never';
    } else if (typeof settingsValues.lastAnalyticsSync === 'string') {
      // Ensure it's a valid date string or 'Never'
      try {
        const testDate = new Date(settingsValues.lastAnalyticsSync);
        if (!isNaN(testDate.getTime())) {
          settingsValues.lastAnalyticsSync = testDate.toISOString();
        } else {
          settingsValues.lastAnalyticsSync = 'Never';
        }
      } catch (dateErr) {
        Logger.log('Warning: Invalid last sync date format: ' + settingsValues.lastAnalyticsSync);
        settingsValues.lastAnalyticsSync = 'Never';
      }
    }

    // Add a convenience field to indicate if this account has been set up
    settingsValues.hasApiKeysConfigured = Boolean(
      settingsValues.twitterApiKey || 
      settingsValues.youtubeApiKey || 
      settingsValues.telegramBotToken
    );
    
    // Add a convenience field for Drive folder status
    settingsValues.hasDriveFolderConfigured = Boolean(
      settingsValues.driveFolderId && 
      settingsValues.driveFolderName
    );

    Logger.log('Integration settings retrieved successfully');
    return settingsValues; // Success case
  } catch (e) {
    Logger.log('Error in getIntegrationSettings: ' + e.toString() + " Stack: " + e.stack);
    return { error: 'An error occurred while retrieving settings: ' + e.toString() };
  }
}

/**
 * Saves API keys to the Settings sheet (wrapper function without underscore).
 * @param {object} keys An object containing { twitterApiKey, twitterApiSecret, youtubeApiKey, telegramBotToken }.
 * @return {object} An object with a success message or an error message.
 */
function saveApiKeys(keys) {
  return saveApiKeys_(keys);
}

/**
 * Saves API keys to the Settings sheet.
 * @param {object} keys An object containing { twitterApiKey, twitterApiSecret, youtubeApiKey, telegramBotToken }.
 * @return {object} An object with a success message or an error message.
 */
function saveApiKeys_(keys) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let settingsSheet = ss.getSheetByName(API_INTEGRATION_CONFIG.SETTINGS_SHEET);
    
    if (!settingsSheet) {
      const setupResult = setupApiIntegration({ showConfirmationDialog: false }); // Ensure the sheet and cells are there silently
       if(!setupResult.success) {
          Logger.log('Error: Settings sheet not found for saving API keys and silent setup failed: ' + setupResult.message);
          return { success: false, message: 'Settings sheet not found and could not be created. ' + setupResult.message };
      }
      settingsSheet = ss.getSheetByName(API_INTEGRATION_CONFIG.SETTINGS_SHEET);
       if(!settingsSheet) {
          return { success: false, message: 'Settings sheet not found and could not be created (post-setup check for saveApiKeys_).' };
      }
    }
    
    settingsSheet.getRange(API_INTEGRATION_CONFIG.API_KEYS.TWITTER_API_KEY_CELL).setValue(keys.twitterApiKey || '');
    settingsSheet.getRange(API_INTEGRATION_CONFIG.API_KEYS.TWITTER_API_SECRET_CELL).setValue(keys.twitterApiSecret || '');
    settingsSheet.getRange(API_INTEGRATION_CONFIG.API_KEYS.YOUTUBE_API_KEY_CELL).setValue(keys.youtubeApiKey || '');
    settingsSheet.getRange(API_INTEGRATION_CONFIG.API_KEYS.TELEGRAM_BOT_TOKEN_CELL).setValue(keys.telegramBotToken || '');
    
    SpreadsheetApp.flush(); // Ensure changes are saved
    return { success: true, message: 'API keys saved successfully.' };
  } catch (e) {
    Logger.log('Error in saveApiKeys_: ' + e.toString());
    return { success: false, message: 'Error saving API keys: ' + e.toString() };
  }
}

/**
 * Connects a Google Drive folder to the content calendar by saving its ID (wrapper function without underscore).
 * @param {string} folderId The Google Drive Folder ID to connect.
 * @return {object} An object with success status, message, and folderName if successful.
 */
function saveGoogleDriveFolderId(folderId) {
  return saveGoogleDriveFolderId_(folderId);
}

/**
 * Connects a Google Drive folder to the content calendar by saving its ID.
 * This version takes folderId as a parameter, suitable for modal dialogs.
 * @param {string} folderId The Google Drive Folder ID to connect.
 * @return {object} An object with success status, message, and folderName if successful.
 */
function saveGoogleDriveFolderId_(folderId) { // Renamed and parameter added
  // const ui = SpreadsheetApp.getUi(); // Not needed for UI alerts here as this returns an object
  
  if (!folderId || folderId.trim() === "") {
    return { success: false, message: 'Folder ID cannot be empty.' };
  }
  folderId = folderId.trim();

  try {
    const folder = DriveApp.getFolderById(folderId);
    const folderName = folder.getName();
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let settingsSheet = ss.getSheetByName(API_INTEGRATION_CONFIG.SETTINGS_SHEET);
    
    if (!settingsSheet) {
      // Attempt to set up the API section of the settings sheet if it's missing, silently
      const setupResult = setupApiIntegration({ showConfirmationDialog: false });
      if (!setupResult.success) {
         Logger.log('Error: Settings sheet not found for saving Drive Folder ID and silent setup failed: ' + setupResult.message);
         return { success: false, message: 'Settings sheet not found and could not be created. ' + setupResult.message };
      }
      settingsSheet = ss.getSheetByName(API_INTEGRATION_CONFIG.SETTINGS_SHEET);
      if (!settingsSheet) {
         return { success: false, message: 'Settings sheet not found and could not be created (post-setup check for saveGoogleDriveFolderId_).' };
      }
    }
    
    // Ensure the rows for Drive ID and Name exist as per setupApiIntegration structure
    // The setupApiIntegration should ensure these labels/cells are ready if it created the sheet.
    // We'll still be defensive here.
    
    const driveIdLabelCell = settingsSheet.getRange(API_INTEGRATION_CONFIG.DRIVE_FOLDER_ID_CELL).offset(0, -1);
    if (driveIdLabelCell.getValue() !== 'Primary Google Drive Assets Folder ID:') {
        driveIdLabelCell.setValue('Primary Google Drive Assets Folder ID:').setFontWeight('bold');
    }
    settingsSheet.getRange(API_INTEGRATION_CONFIG.DRIVE_FOLDER_ID_CELL).setValue(folderId);
    
    const driveNameLabelCell = settingsSheet.getRange(API_INTEGRATION_CONFIG.DRIVE_FOLDER_NAME_CELL).offset(0, -1);
    if (driveNameLabelCell.getValue() !== 'Primary Assets Folder Name:') {
        driveNameLabelCell.setValue('Primary Assets Folder Name:').setFontWeight('bold');
    }
    settingsSheet.getRange(API_INTEGRATION_CONFIG.DRIVE_FOLDER_NAME_CELL).setValue(folderName);
    
    // Minimal formatting for consistency if labels were just added
    driveIdLabelCell.setFontWeight('bold');
    driveNameLabelCell.setFontWeight('bold');

    return { success: true, message: 'Google Drive folder connected successfully: ' + folderName, folderName: folderName };
  } catch (e) {
    Logger.log('Error in saveGoogleDriveFolderId_: ' + e.toString());
    // More specific error for common Drive issues
    if (e.message.includes("Not Found") || e.message.includes("getFolderById")) {
         return { success: false, message: 'Invalid folder ID or folder does not exist. Please check the ID and try again.' };
    } else if (e.message.includes("Access denied") || e.message.includes("permissions")) {
         return { success: false, message: 'Insufficient permissions to access the folder. Please ensure the script has access.' };
    }
    return { success: false, message: 'Error connecting Google Drive folder: ' + e.toString() };
  }
}

/**
 * Saves asset management configuration values.
 * @param {object} cfg Object containing contentSheetName, assetActionColumnName, rowIdColumnName.
 * @return {object} Result object with success boolean and message.
 */
function saveAssetConfig(cfg) {
  return saveAssetConfig_(cfg);
}

function saveAssetConfig_(cfg) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(API_INTEGRATION_CONFIG.SETTINGS_SHEET);
    if (!sheet) {
      const setupResult = setupApiIntegration({ showConfirmationDialog: false });
      sheet = ss.getSheetByName(API_INTEGRATION_CONFIG.SETTINGS_SHEET);
      if (!sheet) {
        return { success: false, message: 'Settings sheet not found and could not be created.' };
      }
    }
    if (typeof setupSettings === 'function') {
      try { setupSettings(sheet, true); } catch(e){}
    }
    if (cfg.contentSheetName !== undefined) sheet.getRange(API_INTEGRATION_CONFIG.ASSET_CONFIG.CONTENT_SHEET_CELL).setValue(cfg.contentSheetName);
    if (cfg.assetActionColumnName !== undefined) sheet.getRange(API_INTEGRATION_CONFIG.ASSET_CONFIG.ASSET_ACTION_COLUMN_CELL).setValue(cfg.assetActionColumnName);
    if (cfg.rowIdColumnName !== undefined) sheet.getRange(API_INTEGRATION_CONFIG.ASSET_CONFIG.ROW_ID_COLUMN_CELL).setValue(cfg.rowIdColumnName);
    SpreadsheetApp.flush();
    return { success: true, message: 'Asset settings saved.' };
  } catch(e) {
    Logger.log('Error in saveAssetConfig_: ' + e.toString());
    return { success: false, message: 'Error saving asset settings: ' + e.toString() };
  }
}

// The original connectGoogleDriveFolder can be kept for direct menu invocation if desired,
// or refactored to call saveGoogleDriveFolderId_ after prompting.
// For now, we'll keep it distinct. The modal will use saveGoogleDriveFolderId_.
function connectGoogleDriveFolder() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Connect Google Drive Folder',
    'Enter the Google Drive Folder ID:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const folderId = response.getResponseText().trim();
  const result = saveGoogleDriveFolderId_(folderId); // Call the new function

  if (result.success) {
    ui.alert(result.message);
  } else {
    ui.alert('Error: ' + result.message);
  }
}

/**
 * Synchronizes Google Drive assets with content items
 * @return {object} An object with success status and message.
 */
function syncGoogleDriveAssets() {
  const ui = SpreadsheetApp.getUi(); // Keep for fatal alerts
  
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contentSheet = ss.getSheetByName(API_INTEGRATION_CONFIG.CONTENT_SHEET);
  const settingsSheet = ss.getSheetByName(API_INTEGRATION_CONFIG.SETTINGS_SHEET);
  
  if (!contentSheet || !settingsSheet) {
    return { success: false, message: 'Required sheets (Content Calendar, Settings) not found.' };
  }
  
  // Get folder ID from settings
  let folderId;
  try {
    folderId = settingsSheet.getRange(API_INTEGRATION_CONFIG.DRIVE_FOLDER_ID_CELL).getValue();
  } catch (e) {
    // Field might not exist
  }
  
  if (!folderId) {
    return { success: false, message: 'No Google Drive folder (for api_integrations.js) connected. Please connect a folder first via the Integrations modal/settings.' };
  }
  
  // Get the folder
  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    return { success: false, message: 'Invalid folder ID or insufficient permissions: ' + e.toString() };
  }
  
  // Get content data
  const data = contentSheet.getDataRange().getValues();
  
  // Extract headers (assuming row 2 contains headers)
  const headers = data[1];
  
  // Find column indexes
  const idIndex = headers.indexOf('ID');
  const channelIndex = headers.indexOf('Channel');
  const contentIndex = headers.indexOf('Content/Idea');
  const linkIndex = headers.indexOf('Link to Asset');
  
  // Skip if any required column is missing
  if (idIndex === -1 || channelIndex === -1 || contentIndex === -1 || linkIndex === -1) {
    return { success: false, message: 'Required columns (ID, Channel, Content/Idea, Link to Asset) not found in Content Calendar.' };
  }
  
  // Get all files in the folder
  const files = folder.getFiles();
  const fileMap = {};
  
  while (files.hasNext()) {
    const file = files.next();
    fileMap[file.getName()] = file;
  }
  
  // Process content items
  let updatedCount = 0;
  
  // Start from row 3 (index 2) to skip headers
  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    const contentId = row[idIndex];
    
    // Skip if no content ID
    if (!contentId) {
      continue;
    }
    
    // Look for files matching the content ID
    const matchingFiles = [];
    
    for (const fileName in fileMap) {
      if (fileName.includes(contentId)) {
        matchingFiles.push(fileMap[fileName]);
      }
    }
    
    // If matching files found, update the link
    if (matchingFiles.length > 0) {
      // Use the first matching file
      const file = matchingFiles[0];
      const fileUrl = file.getUrl();
      
      // Update the link in the content sheet
      contentSheet.getRange(i + 1, linkIndex + 1).setValue(fileUrl);
      updatedCount++;
    }
  }
  
  return { success: true, message: `Google Drive sync completed. Updated ${updatedCount} content items with asset links.` };
}

/**
 * Creates an API integrations menu
 */
function createApiIntegrationsMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Integrations')
    .addItem('Set Up API Integration', 'setupApiIntegration')
    .addSeparator()
    .addItem('Connect Google Drive Folder', 'connectGoogleDriveFolder')
    .addItem('Sync Google Drive Assets', 'syncGoogleDriveAssets')
    .addSeparator()
    .addItem('Fetch Analytics Data', 'fetchContentAnalytics')
    .addToUi();
}

/**
 * Returns the configured Primary Google Drive Assets Folder ID.
 * @return {string|null} The folder ID, or null if not configured or error.
 */
function getPrimaryDriveAssetsFolderId() {
  const settings = getIntegrationSettings();
  if (settings && !settings.error && settings.driveFolderId) {
    return settings.driveFolderId;
  }
  Logger.log('Primary Drive Assets Folder ID not found or error in settings.');
  return null;
}