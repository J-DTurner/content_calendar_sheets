/**
 * Google Drive Integration for Social Media Content Calendar
 *
 * This script provides advanced integration with Google Drive, including:
 * – File browser/picker for asset selection
 * – Automated folder creation for weekly content organization
 * – Direct file creation from the calendar
 * – Asset preview functionality
 */

/* ──────────────────────────────────────────────────────────
 * Configuration
 * ────────────────────────────────────────────────────────── */
const DRIVE_CONFIG = {
  // ASSETS_FOLDER_ID_SETTING_LABEL: 'Google Drive Assets Folder ID:', // No longer used
  ASSET_LINK_COLUMN: 7,             // Column G (Link to Asset in Content Calendar) - ADJUSTED from 9 to match example main_menu.js headers
  WEEK_COLUMN: 3,                   // Column C (Week in Content Calendar)
  DATE_COLUMN: 2,                   // Column B (Date in Content Calendar)
  CHANNEL_COLUMN: 5,                // Column E (Channel in Content Calendar)
  CONTENT_COLUMN: 6,                // Column F (Content/Idea in Content Calendar)
  CREATE_WEEKLY_FOLDERS: true       // Setting to enable/disable weekly subfolders
};

/* ──────────────────────────────────────────────────────────
 * File-picker entry point
 * ────────────────────────────────────────────────────────── */
function openFilePicker() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = SpreadsheetApp.getActiveRange();

  const calendarSheetName = 'Content Calendar'; // Use a constant or global

  if (sheet.getName() !== calendarSheetName ||
      range.getColumn() !== DRIVE_CONFIG.ASSET_LINK_COLUMN) {
    ui.alert(`Please select a cell in the “Link to Asset” column (Column ${String.fromCharCode(64 + DRIVE_CONFIG.ASSET_LINK_COLUMN)}) in the "${calendarSheetName}" sheet.`);
    return;
  }

  // Get the primary asset folder ID from the centralized configuration
  const assetsFolderId = getAssetsFolderId();
  if (!assetsFolderId) {
    ui.alert('Primary Assets Folder ID not configured. Please set it in the "Settings" sheet (cell B18) via the Integrations Modal.');
    return;
  }

  const row = range.getRow();
  if (row < 3) { // Assuming headers in row 1 and 2
      ui.alert("Please select a content row, not a header row.");
      return;
  }
  const weekNumber  = sheet.getRange(row, DRIVE_CONFIG.WEEK_COLUMN).getValue();
  const channel     = sheet.getRange(row, DRIVE_CONFIG.CHANNEL_COLUMN).getValue();
  const contentDateValue = sheet.getRange(row, DRIVE_CONFIG.DATE_COLUMN).getValue();
  const contentDate = (contentDateValue instanceof Date && !isNaN(contentDateValue)) ? contentDateValue : null;


  let targetFolder;
  try {
      const rootAssetsFolder = DriveApp.getFolderById(assetsFolderId);
      if (DRIVE_CONFIG.CREATE_WEEKLY_FOLDERS && weekNumber && contentDate) {
          targetFolder = getOrCreateWeekFolder(rootAssetsFolder, weekNumber, contentDate.getFullYear(), channel);
      } else {
          targetFolder = rootAssetsFolder;
      }
  } catch (e) {
      Logger.log(`Error accessing folder: ${e.toString()}`);
      ui.alert(`Error accessing Google Drive folder. Check Folder ID and permissions: ${e.message}`);
      return;
  }


  showCustomFilePicker(targetFolder.getId(), row);
}

/* ──────────────────────────────────────────────────────────
 * Custom picker UI
 * ────────────────────────────────────────────────────────── */
function showCustomFilePicker(folderId, row) {
  let pickerHtml = `
    <html>
      <head>
        <base target="_top">
        <link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">
        <style>
          body { font-family: Arial, sans-serif; margin: 10px; }
          .container { padding: 10px; }
          .file-list { max-height: 300px; overflow-y: auto; border: 1px solid #ccc; margin-bottom: 10px; }
          .file-item { padding: 8px; cursor: pointer; border-bottom: 1px solid #eee; }
          .file-item:hover { background-color: #f0f0f0; }
          .file-item.selected { background-color: #e0e8ff; font-weight: bold; }
          .buttons { margin-top: 15px; text-align: right; }
          .loader { border: 5px solid #f3f3f3; border-top: 5px solid #3498db; border-radius: 50%;
                    width: 30px; height: 30px; animation: spin 1.5s linear infinite; margin: 20px auto; }
          @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        </style>
      </head>
      <body>
        <div class="container">
          <h4>Select a File or Create New</h4>
          <div id="picker-content"><div class="loader"></div><p style="text-align:center;">Loading files...</p></div>
        </div>
        <script>
          let selectedFileElement = null;
          let selectedFile = null;

          window.onload = function() {
            google.script.run
              .withSuccessHandler(displayFiles)
              .withFailureHandler(showError)
              .getFilesInFolder("${folderId}");
          };

          function displayFiles(files) {
            const contentDiv = document.getElementById('picker-content');
            if (!files || files.length === 0) {
              contentDiv.innerHTML = "<p>No files found in this folder.</p>";
            } else {
              let html = '<div class="file-list">';
              files.forEach((file, index) => {
                html += \`<div class="file-item" id="file-\${index}" data-id="\${file.id}" data-url="\${file.url}" data-name="\${file.name}" onclick="selectFile(this, \${index})">\${file.name} (\${formatMimeType(file.type)})</div>\`;
              });
              html += '</div>';
              contentDiv.innerHTML = html;
            }
            contentDiv.innerHTML += \`
              <div class="buttons">
                <button class="action" onclick="confirmSelection()" id="selectBtn" disabled>Select File</button>
                <button onclick="createNewFileInPicker()">Create New File</button>
                <button onclick="google.script.host.close()">Cancel</button>
              </div>\`;
          }

          function formatMimeType(mimeType) {
            if (mimeType.includes('image')) return 'Image';
            if (mimeType.includes('video')) return 'Video';
            if (mimeType.includes('audio')) return 'Audio';
            if (mimeType.includes('pdf')) return 'PDF';
            if (mimeType.includes('vnd.google-apps.document')) return 'Doc';
            if (mimeType.includes('vnd.google-apps.spreadsheet')) return 'Sheet';
            if (mimeType.includes('vnd.google-apps.presentation')) return 'Slide';
            if (mimeType.includes('vnd.google-apps.drawing')) return 'Drawing';
            return 'File';
          }

          function selectFile(element, index) {
            if (selectedFileElement) {
              selectedFileElement.classList.remove('selected');
            }
            selectedFileElement = element;
            selectedFileElement.classList.add('selected');
            selectedFile = {
              id: element.dataset.id,
              url: element.dataset.url,
              name: element.dataset.name
            };
            document.getElementById('selectBtn').disabled = false;
          }

          function confirmSelection() {
            if (!selectedFile) {
              // Optionally show a message if no file is selected
              showError("No file selected.");
              return;
            }
            
            // Update UI to indicate processing
            const contentDiv = document.getElementById('picker-content');
            contentDiv.innerHTML = '<div class=\"loader\"></div><p style=\"text-align:center;\">Linking file and updating records...</p>';

            // Call the new backend function
            google.script.run
              .withSuccessHandler((response) => {
                // The backend function handleAssetSelectionAndLinking now returns an object {success: boolean, error?: string, message?: string}
                hideLoading(); // Ensure loading indicator is hidden if you add one to the modal UI later
                if (response && response.success) {
                  // Close the modal dialog on success
                  google.script.host.close();
                  // Optionally show a toast in the spreadsheet on success
                  if (response.message) {
                      google.script.run.withSuccessHandler(function(){}).toast(response.message, 'Asset Link Success', 5);
                  }
                } else {
                  // Show error message if backend reported failure
                  showError(response.error || "Failed to link asset. Please check logs or try again.");
                  // Consider re-enabling buttons or providing a retry option here in a more advanced UI
                }
              })
              .withFailureHandler(showError) // showError will handle JavaScript errors or if the Apps Script function call fails entirely
              .handleAssetSelectionAndLinking(${row}, selectedFile.id, selectedFile.url, selectedFile.name); // Pass rowNum, fileId, fileUrl, fileName
          }

          function createNewFileInPicker() {
            document.getElementById('picker-content').innerHTML = '<div class=\"loader\"></div><p style=\"text-align:center;\">Creating file...</p>';
            google.script.run
              .withSuccessHandler(() => google.script.host.close()) // Or re-fetch files
              .withFailureHandler(showError)
              .createNewDriveFile("${folderId}", ${row});
          }

          // Add a simple showError function if not already present
          function showError(error) {
            const contentDiv = document.getElementById('picker-content');
            let errorMessage = "An unknown error occurred.";
            if (typeof error === 'string') {
                errorMessage = error;
            } else if (error && error.message) { // Standard JS Error object
                errorMessage = error.message;
            } else if (error && error.error) { // Custom {error: "message"} object from backend
                errorMessage = error.error;
            }
            contentDiv.innerHTML = \`<p style="color:red;">Error: \${errorMessage}</p><button onclick="google.script.host.close()">Close</button>\`;
            console.error("Picker Error:", error);
          }
          
          // Add basic loading/status functions if not already present for the picker HTML
           function showLoading(message) {
                const contentDiv = document.getElementById('picker-content');
                contentDiv.innerHTML = \`<div class="loader"></div><p style="text-align:center;">\${message || 'Processing...'}</p>\`;
           }

           function hideLoading() {
               // This simple picker doesn't have a persistent loading indicator outside the contentDiv,
               // so hideLoading might just re-render the buttons if needed, or do nothing if showError replaces content.
               // For this specific HTML structure, showError replaces the content, so hideLoading might not be strictly necessary
               // after a call completes. However, it's good practice to have.
               // If the UI were more complex with separate loading elements, this would hide them.
               console.log("hideLoading called - specific picker HTML structure handles progress display within contentDiv.");
           }
        </script>
      </body>
    </html>`;
  const htmlOutput = HtmlService.createHtmlOutput(pickerHtml).setWidth(600).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select Asset from Google Drive');
}


/* ──────────────────────────────────────────────────────────
 * Drive helper functions
 * ────────────────────────────────────────────────────────── */
function getFilesInFolder(folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const filesIterator  = folder.getFiles();
    const fileList   = [];
    while (filesIterator.hasNext()) {
      const file = filesIterator.next();
      fileList.push({
        id:   file.getId(),
        name: file.getName(),
        url:  file.getUrl(),
        type: file.getMimeType()
      });
    }
    // Also list subfolders for navigation (optional enhancement)
    // const foldersIterator = folder.getFolders();
    // while (foldersIterator.hasNext()) { ... }
    return fileList.sort((a,b) => a.name.localeCompare(b.name)); // Sort files
  } catch (e) {
    Logger.log(`Error getting files in folder ${folderId}: ${e.toString()}`);
    // Throw error to be caught by withFailureHandler in client-side JS
    throw new Error(`Failed to get files: ${e.message}`);
  }
}

function setFileLink(row, fileUrl, fileName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Content Calendar');
  if (sheet) {
    sheet.getRange(row, DRIVE_CONFIG.ASSET_LINK_COLUMN)
         .setFormula(`=HYPERLINK("${fileUrl}","${fileName.replace(/"/g, '""')}")`); // Escape double quotes in name
    SpreadsheetApp.getActiveSpreadsheet().toast(`Linked: ${fileName}`);
  }
}

function createNewDriveFile(folderId, row) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Content Calendar');
  if (!sheet) { ui.alert("Content Calendar sheet not found."); return; }

  const channel     = sheet.getRange(row, DRIVE_CONFIG.CHANNEL_COLUMN).getValue() || "General";
  const contentText = sheet.getRange(row, DRIVE_CONFIG.CONTENT_COLUMN).getValue();
  const contentDateValue = sheet.getRange(row, DRIVE_CONFIG.DATE_COLUMN).getValue();
  const contentDate = (contentDateValue instanceof Date && !isNaN(contentDateValue)) ? contentDateValue : new Date();

  const dateString  = Utilities.formatDate(contentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const safeContentText = contentText ? contentText.substring(0, 50).replace(/[^\w\s-]/g, '') : 'New Asset';
  const baseFileName = `${dateString} ${channel} - ${safeContentText}`;

  const fileType = showFileTypeDialog();
  if (!fileType) { ui.alert("File creation cancelled."); return; } // User cancelled

  let fileName = baseFileName;
  if (fileType === 'document') fileName += ".gdoc"; // Just an example, actual extension not needed for GSuite files
  else if (fileType === 'spreadsheet') fileName += ".gsheet";
  else if (fileType === 'presentation') fileName += ".gslides";

  try {
    const folder = DriveApp.getFolderById(folderId);
    let newFile;

    switch (fileType) {
      case 'document':
        newFile = DocumentApp.create(fileName);
        break;
      case 'spreadsheet':
        newFile = SpreadsheetApp.create(fileName);
        break;
      case 'presentation':
        newFile = SlidesApp.create(fileName);
        break;
      case 'drawing':
        // Google Drawings cannot be directly created and moved atomically by Apps Script in the same way.
        // Create a placeholder (e.g. a Doc) and instruct user, or use advanced Drive API if absolutely needed.
        const placeholder = DocumentApp.create(`DRAWING_PLACEHOLDER_FOR_${fileName}`);
        DriveApp.getFileById(placeholder.getId()).moveTo(folder);
        setFileLink(row, placeholder.getUrl(), `DRAWING: ${fileName} (Replace with actual Drawing)`);
        ui.alert(`Placeholder document created for Google Drawing: "${fileName}". Please create the actual Drawing in the folder and update the link.`);
        return; // Exit early for drawing
      default:
        ui.alert("Unsupported file type for creation.");
        return;
    }

    const fileId = newFile.getId();
    const createdFile = DriveApp.getFileById(fileId);
    createdFile.moveTo(folder); // Move the newly created file to the target folder

    setFileLink(row, createdFile.getUrl(), createdFile.getName());
    SpreadsheetApp.getActiveSpreadsheet().toast(`File created and linked: ${createdFile.getName()}`);

  } catch (err) {
    Logger.log(`Error creating file: ${err.toString()}`);
    ui.alert(`Error creating file: ${err.message}`);
  }
}


function showFileTypeDialog() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Create New File',
    'Choose file type:\n1. Google Doc (document)\n2. Google Sheet (spreadsheet)\n3. Google Slides (presentation)\n4. Google Drawing (placeholder)',
    ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() == ui.Button.OK) {
    const choice = result.getResponseText().trim();
    switch (choice) {
      case '1': return 'document';
      case '2': return 'spreadsheet';
      case '3': return 'presentation';
      case '4': return 'drawing';
      default:
        ui.alert('Invalid selection. Please enter a number from 1 to 4.');
        return null; // Or recall showFileTypeDialog()
    }
  }
  return null; // User cancelled
}

/* ──────────────────────────────────────────────────────────
 * Folder-creation helpers
 * ────────────────────────────────────────────────────────── */
function getOrCreateWeekFolder(rootAssetsFolder, week, year, channel) {
  const weekFolderName  = `Week ${String(week).padStart(2, '0')} - ${year}`;
  let weekFolder;

  const weekFolders = rootAssetsFolder.getFoldersByName(weekFolderName);
  if (weekFolders.hasNext()) {
    weekFolder = weekFolders.next();
  } else {
    weekFolder = rootAssetsFolder.createFolder(weekFolderName);
    Logger.log(`Created week folder: ${weekFolderName}`);
  }

  if (!channel || typeof channel !== 'string' || channel.trim() === '') {
    return weekFolder; // No channel subfolder needed
  }

  const safeChannelName = channel.replace(/[^\w\s-]/g, '_'); // Sanitize channel name for folder
  let channelFolder;
  const channelFolders = weekFolder.getFoldersByName(safeChannelName);
  if (channelFolders.hasNext()) {
    channelFolder = channelFolders.next();
  } else {
    channelFolder = weekFolder.createFolder(safeChannelName);
    Logger.log(`Created channel subfolder: ${safeChannelName} in ${weekFolderName}`);
  }
  return channelFolder;
}

/**
 * Gets the primary asset folder ID from the centralized configuration in Settings sheet (cell B18).
 * This function now uses getPrimaryDriveAssetsFolderId() from api_integrations.js instead
 * of using the local DRIVE_CONFIG settings.
 * @return {string|null} The primary asset folder ID or null if not configured
 */
function getAssetsFolderId() {
  // Call the new central function from api_integrations.js
  if (typeof getPrimaryDriveAssetsFolderId === 'function') {
    const folderId = getPrimaryDriveAssetsFolderId();
    if (folderId) {
      return folderId;
    } else {
      Logger.log("getAssetsFolderId in drive_integration_script.js: Primary Drive Assets Folder ID not configured in Settings (B18) or error retrieving it.");
      // SpreadsheetApp.getUi().alert("Primary Drive Assets Folder ID is not set. Please configure it in the Settings sheet (cell B18) via the Integrations Modal."); // Alerting here might be too noisy if called indirectly.
      return null;
    }
  } else {
    Logger.log("getAssetsFolderId in drive_integration_script.js: Critical error - getPrimaryDriveAssetsFolderId function is not available. Make sure api_integrations.js is loaded.");
    // SpreadsheetApp.getUi().alert("Critical error: Drive configuration function missing.");
    return null;
  }
}


/* ──────────────────────────────────────────────────────────
 * Quick-access menu
 * ────────────────────────────────────────────────────────── */
function updateDriveMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    // Check if menu already exists to avoid duplication if onOpen runs multiple times
    const menus = ui.getMenu().getItems();
    let driveMenuExists = false;
    menus.forEach(menu => {
      if (menu.getCaption() === '📎 Drive Tools') {
        driveMenuExists = true;
      }
    });

    if (!driveMenuExists) {
      ui.createMenu('📎 Drive Tools')
        .addItem('🔗 Link Asset from Drive',      'openFilePicker')
        .addItem('➕ Create New Asset in Drive',  'createNewAssetFromMenu') // Wrapper for clarity
        .addItem('🖼️ Preview Selected Asset',     'previewAsset')
        .addSeparator()
        .addItem('📂 Open Root Assets Folder',   'openAssetsFolderInDrive') // Wrapper
        .addToUi();
    }
  } catch (e) {
    Logger.log("Error creating Drive Tools menu: " + e.toString());
  }
}

function createNewAssetFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = SpreadsheetApp.getActiveRange();
  const calendarSheetName = 'Content Calendar';

  if (sheet.getName() !== calendarSheetName ||
      range.getColumn() !== DRIVE_CONFIG.ASSET_LINK_COLUMN) {
    ui.alert(`Please select a cell in the “Link to Asset” column (Column ${String.fromCharCode(64 + DRIVE_CONFIG.ASSET_LINK_COLUMN)}) in the "${calendarSheetName}" sheet to indicate where the new asset link should go.`);
    return;
  }
   const row = range.getRow();
   if (row < 3) { // Assuming headers in row 1 and 2
       ui.alert("Please select a content row, not a header row.");
       return;
   }

  // Get the primary asset folder ID from the centralized configuration
  const assetsFolderId = getAssetsFolderId();
  if (!assetsFolderId) {
    ui.alert('Primary Assets Folder ID not configured. Please set it in the "Settings" sheet (cell B18) via the Integrations Modal.');
    return;
  }

  const weekNumber  = sheet.getRange(row, DRIVE_CONFIG.WEEK_COLUMN).getValue();
  const channel     = sheet.getRange(row, DRIVE_CONFIG.CHANNEL_COLUMN).getValue();
  const contentDateValue = sheet.getRange(row, DRIVE_CONFIG.DATE_COLUMN).getValue();
  const contentDate = (contentDateValue instanceof Date && !isNaN(contentDateValue)) ? contentDateValue : null;

  let targetFolder;
  try {
      const rootAssetsFolder = DriveApp.getFolderById(assetsFolderId);
       if (DRIVE_CONFIG.CREATE_WEEKLY_FOLDERS && weekNumber && contentDate) {
          targetFolder = getOrCreateWeekFolder(rootAssetsFolder, weekNumber, contentDate.getFullYear(), channel);
      } else {
          targetFolder = rootAssetsFolder;
      }
  } catch(e) {
      Logger.log(`Error accessing base asset folder for new asset: ${e.toString()}`);
      ui.alert(`Error getting target folder: ${e.message}`);
      return;
  }

  createNewDriveFile(targetFolder.getId(), row);
}

function openAssetsFolderInDrive() {
  // Get the primary asset folder ID from the centralized configuration
  const assetsFolderId = getAssetsFolderId();
  if (!assetsFolderId) {
    SpreadsheetApp.getUi().alert('Primary Assets Folder ID not configured. Please set it in the "Settings" sheet (cell B18) via the Integrations Modal.');
    return;
  }
  const folderUrl = `https://drive.google.com/drive/folders/${assetsFolderId}`;
  const html = HtmlService.createHtmlOutput(
      `<script>window.open("${folderUrl}", "_blank"); google.script.host.close();</script>`)
    .setWidth(100).setHeight(1); // Minimal dialog
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Assets Folder...');
}


/* ──────────────────────────────────────────────────────────
 * Asset preview dialog
 * ────────────────────────────────────────────────────────── */
function previewAsset() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = SpreadsheetApp.getActiveRange();
  const calendarSheetName = 'Content Calendar';

  if (sheet.getName() !== calendarSheetName ||
      range.getColumn() !== DRIVE_CONFIG.ASSET_LINK_COLUMN) {
    ui.alert(`Please select a cell with an asset link in the "Link to Asset" column (Column ${String.fromCharCode(64 + DRIVE_CONFIG.ASSET_LINK_COLUMN)}) of the "${calendarSheetName}" sheet.`);
    return;
  }

  const cellValue = range.getFormula() || range.getValue(); // Check formula first for HYPERLINK
  let url = "";

  if (cellValue.toString().toUpperCase().startsWith("=HYPERLINK(")) {
    try {
      const matches = cellValue.match(/HYPERLINK\("([^"]+)"/i);
      if (matches && matches[1]) {
        url = matches[1];
      }
    } catch(e) { Logger.log("Error parsing hyperlink formula: " + e); }
  }
  if (!url && range.getValue().toString().startsWith("http")) {
      url = range.getValue().toString();
  }


  if (!url || !url.toString().trim().startsWith("http")) {
    ui.alert('No valid asset link found in the selected cell.');
    return;
  }

  let fileId = null;
  const patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]+)/,       // Standard Drive file URL
    /open\?id=([a-zA-Z0-9_-]+)/,          // Another common pattern
    /\/document\/d\/([a-zA-Z0-9_-]+)/,    // Google Docs
    /\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/,// Google Sheets
    /\/presentation\/d\/([a-zA-Z0-9_-]+)/, // Google Slides
    /\/drawings\/d\/([a-zA-Z0-9_-]+)/     // Google Drawings
  ];
  for (const re of patterns) {
    const m = url.match(re);
    if (m && m[1]) { fileId = m[1]; break; }
  }
  if (!fileId) {
    ui.alert('Could not determine the file ID from the URL. Preview might not work for this link type.');
    // Option: Show a simple iframe with the URL itself for non-standard links
    const basicPreviewHtml = `<html><body><iframe src="${url}" width="100%" height="95%" frameborder="0"></iframe></body></html>`;
    const basicOutput = HtmlService.createHtmlOutput(basicPreviewHtml).setWidth(800).setHeight(600).setTitle('Asset Preview');
    ui.showModalDialog(basicOutput, 'Asset Preview');
    return;
  }

  try {
    const file = DriveApp.getFileById(fileId);
    const mimeType = file.getMimeType();
    const fileName = file.getName();
    const fileSize = formatBytes(file.getSize());
    const lastUpdated = Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    let previewHtml = `
      <html>
        <head>
          <base target="_top">
          <style>
            body { margin: 0; font-family: Arial, sans-serif; display: flex; flex-direction: column; height: 100vh; }
            .file-info { padding: 10px; background: #f8f9fa; border-bottom: 1px solid #ddd; font-size: 0.9em; }
            .file-info strong { display: block; margin-bottom: 5px; font-size: 1.1em; }
            .preview-container { flex-grow: 1; overflow: auto; text-align: center; padding: 10px; }
            iframe { width: 100%; height: 100%; border: none; }
            img, video, audio { max-width: 100%; max-height: calc(100vh - 100px); display: block; margin: auto; }
            .unsupported { padding: 20px; text-align: center; }
            .unsupported a { padding:10px 15px; background-color:#4CAF50; color:white; text-decoration:none; border-radius:4px; }
          </style>
        </head>
        <body>
          <div class="file-info">
            <strong>${fileName}</strong>
            Type: ${mimeType}<br>
            Size: ${fileSize}<br>
            Last Updated: ${lastUpdated}
          </div>
          <div class="preview-container">`;

    // Determine preview type based on MIME type
    if (mimeType.startsWith('image/')) {
      previewHtml += `<img src="https://drive.google.com/uc?id=${fileId}&export=view" alt="${fileName}">`;
    } else if (mimeType === 'application/pdf' ||
               mimeType === 'application/vnd.google-apps.document' ||
               mimeType === 'application/vnd.google-apps.spreadsheet' ||
               mimeType === 'application/vnd.google-apps.presentation' ||
               mimeType === 'application/vnd.google-apps.drawing') {
      previewHtml += `<iframe src="https://drive.google.com/file/d/${fileId}/preview"></iframe>`;
    } else if (mimeType.startsWith('video/')) {
      previewHtml += `<video controls><source src="https://drive.google.com/uc?export=download&id=${fileId}" type="${mimeType}"></video>`;
    } else if (mimeType.startsWith('audio/')) {
      previewHtml += `<audio controls><source src="https://drive.google.com/uc?export=download&id=${fileId}" type="${mimeType}"></audio>`;
    } else {
      previewHtml += `<div class="unsupported"><p>Preview not available for this file type (${mimeType}).</p><p><a href="${url}" target="_blank">Open File in New Tab</a></p></div>`;
    }

    previewHtml += `</div></body></html>`;
    const htmlOutput = HtmlService.createHtmlOutput(previewHtml).setWidth(800).setHeight(650).setTitle('Asset Preview: ' + fileName);
    ui.showModalDialog(htmlOutput, 'Asset Preview');

  } catch (err) {
    Logger.log(`Error previewing file ${fileId}: ${err.toString()}`);
    ui.alert(`Error previewing file: ${err.message}. The file might not exist or you may not have permission.`);
  }
}

/* ──────────────────────────────────────────────────────────
 * Utility – format bytes
 * ────────────────────────────────────────────────────────── */
function formatBytes(bytes, decimals = 2) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const dm = decimals < 0 ? 0 : decimals;
  const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

/**
 * Handles the selection of a file from the picker, updates the Content Calendar,
 * and records the asset link in the Assets sheet.
 * @param {number} rowNum The row number in the Content Calendar sheet.
 * @param {string} fileId The Google Drive ID of the selected file.
 * @param {string} fileUrl The URL of the selected file.
 * @param {string} fileName The name of the selected file.
 * @return {object} An object with success status and an optional error message.
 */
function handleAssetSelectionAndLinking(rowNum, fileId, fileUrl, fileName) {
  try {
    Logger.log(`Handling asset selection: Row ${rowNum}, File ID ${fileId}, Name ${fileName}`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const contentSheet = ss.getSheetByName('Content Calendar');

    if (!contentSheet) {
      Logger.log('Content Calendar sheet not found.');
      return { success: false, error: 'Content Calendar sheet not found.' };
    }

    // 1. Set the hyperlink in the Content Calendar sheet
    // setFileLink is already defined in drive_integration_script.js
    setFileLink(rowNum, fileUrl, fileName);

    // 2. Get the Content ID from the Content Calendar sheet (assuming ID is in Column A)
    // Ensure this column index is correct based on your sheet structure.
    // If your ID column is not A (1), update this constant.
    const ID_COLUMN_CONTENT_CALENDAR = 1; // Column A
    const contentCalendarId = contentSheet.getRange(rowNum, ID_COLUMN_CONTENT_CALENDAR).getValue();

    if (!contentCalendarId) {
      Logger.log(`No Content ID found in Content Calendar sheet at row ${rowNum}, column ${ID_COLUMN_CONTENT_CALENDAR}.`);
      // Still consider partial success as the link was set. User might fix ID later.
      SpreadsheetApp.getActiveSpreadsheet().toast(`Asset linked in sheet, but Content ID for row ${rowNum} is missing. Asset sheet not updated.`, "Warning", 7);
      return { success: true, message: `Asset linked in sheet. Warning: Content ID for row ${rowNum} is missing, so 'Assets' sheet was not updated.` };
    }

    // 3. Record the asset link in the "Assets" sheet using Code.js's function
    // Assumes linkAssetToRow is globally available (defined in Code.js and included)
    if (typeof linkAssetToRow === "function") {
      // linkAssetToRow expects a string rowIdentifier, fileId, fileName
      const linkResult = linkAssetToRow(contentCalendarId.toString(), fileId, fileName);
      if (!linkResult.success) {
        Logger.log(`Failed to link asset in Assets sheet: ${linkResult.error}`);
        // Return false as the full process (linking *and* recording) failed
        return { success: false, error: `Asset linked in calendar, but failed to update Assets sheet: ${linkResult.error}` };
      }
      Logger.log(`Asset ${fileId} linked to Content ID ${contentCalendarId} in Assets sheet.`);
    } else {
      Logger.log("linkAssetToRow function from Code.js is not available.");
      // Return false as a critical part of the process failed
      return { success: false, error: "Asset linked in calendar, but 'Assets' sheet update function (linkAssetToRow) is missing." };
    }

    // If both linking in calendar and linking in Assets sheet were successful
    return { success: true, message: 'Asset linked and recorded successfully.' };

  } catch (e) {
    Logger.log(`Error in handleAssetSelectionAndLinking: ${e.toString()}\n${e.stack}`);
    return { success: false, error: `An unexpected error occurred during asset linking: ${e.toString()}` };
  }
}