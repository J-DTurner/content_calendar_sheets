/**
 * Google Apps Script Trigger Functions
 * 
 * This file contains functions that are specifically meant to be
 * triggered by Google Apps Script events such as onOpen, onEdit, etc.
 * 
 * The onOpen function here is the primary entry point when the 
 * spreadsheet is opened, and it delegates to other functions as needed.
 */

/**
 * Simple onOpen trigger that creates the menu and handles initial auth prompt.
 * This is the ONLY function that should be named "onOpen" in the entire project.
 */
function onOpen(e) {
  try {
    Logger.log("onOpen trigger called. AuthMode: " + (e ? e.authMode : 'N/A'));
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Always create a basic menu
    const menu = ui.createMenu('📊 Content Calendar')
      .addItem('🚀 Show App Panel', 'showAppPanel')
      .addSeparator()
      .addSubMenu(ui.createMenu('🖼️ Asset Management')
        .addItem('Initialize/Refresh Asset Column', 'initializeOrRefreshAssetColumn')
        .addItem('Manage Selected Asset Cell', 'performActionOnSelectedAssetCell')
        .addItem('Show Asset Configuration', 'loadConfigAndDisplay'))
      .addSeparator()
      .addItem('⚙️ Initial Setup / Re-authorize', 'userInitiatedFullSetup');
    
    menu.addToUi();
    Logger.log("Basic menu created.");

    const userProperties = PropertiesService.getUserProperties();
    const authCheckCompleted = userProperties.getProperty('AUTH_AND_CONFIG_CHECKS_COMPLETED') === 'true';

    if (e && e.authMode === ScriptApp.AuthMode.FULL && authCheckCompleted) {
      Logger.log("AuthMode is FULL and auth check previously completed. Proceeding with regular onOpen updates.");
      // Proceed with normal "silent" onOpen updates if already fully authorized and setup marked complete
      if (typeof handleOpenUpdates === 'function') {
        handleOpenUpdates(e, ss, ui); // This function should do its own checks for configured services
      }
      // Show sidebar if not already shown by handleOpenUpdates or if it's preferred here
      try { showAppPanel(); } catch (err) { Logger.log("Error showing app panel in onOpen full auth path: " + err); }

    } else if (e && e.authMode !== ScriptApp.AuthMode.NONE) { // Allow LIMITED mode to show prompt, as FULL might not be set yet
      Logger.log("AuthMode is not FULL or auth check not completed. Displaying initial auth prompt modal.");
      try {
        const htmlOutput = HtmlService.createHtmlOutputFromFile('InitialAuthPrompt.html')
          .setWidth(500)
          .setHeight(380)
          .setSandboxMode(HtmlService.SandboxMode.IFRAME);
        ui.showModalDialog(htmlOutput, 'Content Calendar Setup Required');
      } catch (modalError) {
        Logger.log("Error displaying InitialAuthPrompt.html modal: " + modalError.toString());
        ui.alert("Setup Required", "This script requires authorization to function fully. Please use the 'Initial Setup / Re-authorize' menu option.", ui.ButtonSet.OK);
      }
      // Do not proceed with handleOpenUpdates here; it will be handled by the modal's flow.
      // Show App Panel can be attempted after modal or by user action
      try { showAppPanel(); } catch (err) { Logger.log("Error showing app panel after modal decision: " + err); }


    } else {
      // AuthMode is NONE or event object 'e' is missing (e.g., manual run from editor)
      // This path is less common for typical sheet opens by users.
      // Menu item "Initial Setup / Re-authorize" is the primary way to recover.
      Logger.log("AuthMode is NONE or event object 'e' is missing. Full initialization deferred to user action via menu.");
      ss.toast("Full setup required. Use '📊 Content Calendar > Initial Setup / Re-authorize'.", "Setup Needed", 10);
      try { showAppPanel(); } catch (err) { Logger.log("Error showing app panel in AuthMode.NONE path: " + err); }
    }

  } catch (error) {
    Logger.log("CRITICAL ERROR in onOpen: " + error.toString() + "\nStack: " + error.stack);
    try {
      SpreadsheetApp.getUi()
        .createMenu('🚨 Emergency Menu')
        .addItem('⚙️ Initial Setup / Re-authorize', 'userInitiatedFullSetup')
        .addToUi();
    } catch (emergencyMenuError) {
      Logger.log("FATAL: Even emergency menu creation failed: " + emergencyMenuError.toString());
    }
  }
}

/**
 * Simple onEdit trigger that handles spreadsheet edits.
 * This should be the ONLY function named "onEdit" in the project.
 */
function onEdit(e) {
  Logger.log("onEdit trigger called");
  
  try {
    // Call the main onEdit function from main_integration.js if it exists
    if (typeof handleContentEdit === 'function') {
      handleContentEdit(e);
    }
  } 
  catch (error) {
    Logger.log("ERROR in onEdit: " + error.toString());
  }
}