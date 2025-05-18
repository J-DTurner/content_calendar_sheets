/**
 * OnOpen Update Functions for Social Media Content Calendar
 * 
 * This script provides functions for handling onOpen updates and authorizations.
 * It is meant to be used as a helper module for main_menu.js and not executed directly.
 */

/**
 * Handles authorization and configuration checks when the spreadsheet is opened.
 * This is called by initializeApp (formerly onOpen) in main_menu.js.
 * 
 * @param {GoogleAppsScript.Base.EventObject} e The event object passed to the onOpen trigger.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @param {GoogleAppsScript.Base.Ui} ui The UI object.
 * @return {boolean} True if initialization succeeded, false if it failed or was cancelled.
 */
function handleOpenUpdates(e, ss, ui) {
  try {
    Logger.log("Handling onOpen updates and authorization checks");
    
    // Check if we have full access or need to prompt for authorization
    let needsPrompt = false;
    
    try {
      // Try to access PropertiesService which requires authorization
      const userProperties = PropertiesService.getUserProperties();
      const authCheckPropertyKey = 'AUTH_AND_CONFIG_CHECKS_COMPLETED';
      const authCheckCompleted = userProperties.getProperty(authCheckPropertyKey);
      
      // Need to prompt if the auth check hasn't been completed
      needsPrompt = (authCheckCompleted !== 'true');
      
      // Check auth mode from event object if available
      if (e && (e.authMode === ScriptApp.AuthMode.NONE || e.authMode === ScriptApp.AuthMode.LIMITED)) {
        Logger.log(`Auth mode is ${e.authMode}, which requires authorization prompt`);
        needsPrompt = true;
      }
      
      // Handle based on whether we need to prompt
      if (needsPrompt) {
        // Check if DriveApp is accessible without a prompt
        let preAuthCheckOk = false;
        try {
          DriveApp.getRootFolder().getName(); 
          preAuthCheckOk = true;
          Logger.log("Pre-authorization check successful (DriveApp accessible).");
        } catch (err) {
          Logger.log("Pre-authorization check failed (DriveApp not accessible): " + err.message);
        }
    
        if (!preAuthCheckOk) {
          // Trigger full authorization prompt
          if (typeof promptAndTriggerApiAuthorizations === 'function') {
            const generalAuthSuccessful = promptAndTriggerApiAuthorizations(ui, ss);
            if (generalAuthSuccessful) { 
              userProperties.setProperty(authCheckPropertyKey, 'true');
              Logger.log(`${authCheckPropertyKey} property set to true.`);
            }
          } else {
            ss.toast('Authorization service not found. Some features may be unavailable.', 'Auth Error', 7);
          }
        } else {
          // DriveApp works but we haven't marked auth as complete
          Logger.log("Pre-authorization check passed; general permissions seem OK. Marking as completed.");
          userProperties.setProperty(authCheckPropertyKey, 'true');
          
          // Check configured services
          if (typeof _getconfiguredServiceAccessMessages === 'function') {
            const configMessages = _getconfiguredServiceAccessMessages(ss);
            // Display results silently in toast only if there are issues or successes
            if (configMessages.some(m => m.startsWith("✅"))) {
              ss.toast("Configured services accessible.", "Access Check Complete", 7);
            } else if (configMessages.some(m => m.startsWith("❌"))) {
              ss.toast("Some configured services inaccessible. See App Panel.", "Access Check Warning", 7);
            }
          }
        }
      } else {
        // Auth check already completed, just do a silent check
        Logger.log("Auth check completed previously. Doing silent configuration check.");
        
        if (typeof _getconfiguredServiceAccessMessages === 'function') {
          const configMessages = _getconfiguredServiceAccessMessages(ss);
          // Only show toast if there are issues
          if (configMessages.some(m => m.startsWith("❌"))) {
            ss.toast("Some configured services inaccessible. See App Panel.", "Access Check Warning", 7);
          }
        }

        // Initialize Asset Management functionality if available
        try {
          if (typeof initializeAssetManagement === 'function') {
            Logger.log("Initializing Asset Management functionality");
            initializeAssetManagement();
            
            // Automatically initialize/refresh the asset column
            if (typeof initializeOrRefreshAssetColumn === 'function') {
              Logger.log("Initializing/refreshing asset column");
              initializeOrRefreshAssetColumn();
            }
          }
        } catch (assetError) {
          Logger.log("Error initializing Asset Management: " + assetError.toString());
          ss.toast("Asset Management could not be fully initialized", "Warning", 5);
        }
      }
      
      return true;
    } catch (authError) {
      Logger.log("Error in authorization check: " + authError.toString());
      ss.toast("Error during setup checks. Some features may be unavailable.", "Setup Error", 7);
      return false;
    }
  } catch (e) {
    Logger.log("Critical error in handleOpenUpdates: " + e.toString());
    return false;
  }
}