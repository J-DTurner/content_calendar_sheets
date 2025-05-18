/**
 * Main Menu and Integration for Social Media Content Calendar
 * 
 * This script creates the main menu and integrates all the various
 * functionality modules for the content calendar.
 */

/**
 * Creates main menu and shows app panel option.
 * Also prompts for authorization checks if not previously completed.
 * 
 * This function has been renamed from onOpen to initializeApp to avoid conflicts.
 * It is now called by the trigger_functions.js onOpen function.
 */
function initializeApp(e) {
  try {
    // First and most important task: Create the menu
    // This should always succeed regardless of authorization status
    const ui = SpreadsheetApp.getUi();
    Logger.log("onOpen: Creating main menu");
    ui.createMenu('🚀 Content App')
      .addItem('✨ Show App Panel', 'showAppPanel')
      .addSeparator()
      .addItem('⚙️ Setup/Verify Calendar', 'initializeContentCalendar')
      .addItem('🔑 Authorize Services', 'triggerFullAuthAndConfigCheckFromPanel')
      .addToUi();
    Logger.log("onOpen: Menu created successfully");
    
    // After menu is created, try to do the rest of initialization
    fullInitialization(e);
  } catch (menuError) {
    // If even creating the menu fails, log the error
    Logger.log(`CRITICAL ERROR: Failed to create menu - ${menuError.message}\n${menuError.stack}`);
    // Try to create a minimal emergency menu if the main one failed
    try {
      const ui = SpreadsheetApp.getUi();
      ui.createMenu('🚨 Emergency Menu')
        .addItem('🔑 Authorize Services', 'triggerFullAuthAndConfigCheckFromPanel')
        .addToUi();
    } catch (e) {
      // At this point, we can't do anything more
      Logger.log("FATAL: Even emergency menu creation failed");
    }
  }
}

/**
 * Continues initialization tasks after the menu has been created.
 * This function is separate from initializeApp to ensure the menu always appears
 * even if other initialization tasks fail.
 */
function fullInitialization(e) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Try to show the app panel
  try {
    Logger.log("onOpen: Attempting to show app panel");
    // Only show sidebar if we're not in AuthMode.NONE or LIMITED
    const canShowUI = !e || !(e.authMode === ScriptApp.AuthMode.NONE || e.authMode === ScriptApp.AuthMode.LIMITED);
    if (canShowUI) {
      showAppPanel();
      Logger.log("onOpen: App panel shown successfully");
    } else {
      Logger.log("onOpen: Skipping sidebar due to auth restrictions");
    }
  } catch (sidebarError) {
    Logger.log(`onOpen: Error showing app panel - ${sidebarError.message}\n${sidebarError.stack}`);
    // Don't show alert, just log the error
  }

  // 2. Call other onOpen hooks like dashboard updates
  try {
    Logger.log("onOpen: Checking for dashboard update function");
    if (typeof updateDashboardOnOpen === 'function') {
      updateDashboardOnOpen();
      Logger.log("onOpen: Dashboard updated successfully");
    }
  } catch (dashboardError) {
    Logger.log(`onOpen: Error updating dashboard - ${dashboardError.message}\n${dashboardError.stack}`);
  }

  // 3. Authorization and Configuration Checks (only if we have full permissions)
  try {
    if (!e || e.authMode === ScriptApp.AuthMode.FULL) {
      Logger.log("onOpen: Running authorization checks");
      // This might fail if user hasn't authorized the script yet
      const userProperties = PropertiesService.getUserProperties();
      const authCheckPropertyKey = 'AUTH_AND_CONFIG_CHECKS_COMPLETED';
      const authCheckCompleted = userProperties.getProperty(authCheckPropertyKey);
      
      // Only show auth prompt if necessary and user hasn't already completed it
      if (authCheckCompleted !== 'true') {
        Logger.log("onOpen: Auth check not completed, showing prompt");
        // Show auth prompt silently without blocking
        ss.toast('Click "Authorize Services" in the menu to set up all features', 'Setup Required', 10);
      } else {
        // Run a quick check for configuration issues
        try {
          const configMessages = _getconfiguredServiceAccessMessages(ss);
          const configErrors = configMessages.filter(m => m.startsWith("❌"));
          if (configErrors.length > 0) {
            ss.toast('Configuration issues found. Check App Panel.', 'Config Warning', 10);
          }
        } catch (configError) {
          Logger.log(`onOpen: Error checking configurations - ${configError.message}`);
        }
      }
    } else {
      Logger.log(`onOpen: Skipping auth checks due to auth mode: ${e ? e.authMode : 'undefined'}`);
    }
  } catch (authError) {
    Logger.log(`onOpen: Error during auth checks - ${authError.message}\n${authError.stack}`);
    // Don't show alert, script will prompt for authorization when needed
  }
  
  Logger.log("onOpen: Initialization completed");
}

/**
 * Shows the main application sidebar.
 */
function showAppPanel() {
  try {
    const html = createSafeHtmlOutput('AppSidebar', 300, 800);
    if (!html) {
      throw new Error("Failed to create HTML output");
    }
    
    html.setTitle('Content Calendar App');
    SpreadsheetApp.getUi().showSidebar(html);
    Logger.log("App Panel sidebar shown.");
  } catch (err) {
    Logger.log("Error showing App Panel: " + err.message);
    // Show a toast instead of an alert
    SpreadsheetApp.getActiveSpreadsheet().toast("Could not load the App Panel. Try the menu options instead.", "App Panel Error", 7);
  }
}

// Function to recreate specific top menus if sidebar buttons trigger this
// This is an example if you want to dynamically add back old menus
function recreateSpecificTopMenu(menuName) {
    SpreadsheetApp.getUi().alert(`The functions for "${menuName}" are typically accessed via the App Panel. If you need the old top menu for this, it would be rebuilt here.`);
    // Example:
    // if (menuName === 'Integrations' && typeof createApiIntegrationsMenu === 'function') {
    //   createApiIntegrationsMenu();
    //   SpreadsheetApp.getActiveSpreadsheet().toast(menuName + " menu added to top bar.", "Menu Added", 5);
    // } // Add more for other menus
}

/**
 * Creates and shows the Integrations Management modal dialog.
 */
function openIntegrationsDialog() {
  try {
    const htmlOutput = createSafeHtmlOutput('IntegrationsModal', 700, 550);
    if (!htmlOutput) {
      throw new Error("Failed to create HTML output");
    }
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Manage Integrations');
    Logger.log("Integrations Modal opened successfully");
  } catch (e) {
    Logger.log('Error opening Integrations Modal: ' + e.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast('Could not open Integrations dialog. Try again later.', 'Dialog Error', 7);
  }
}


/**
 * Triggers benign API access calls to prompt for user authorization.
 * Can be called from menu or during initialization.
 */
/**
 * Triggers benign API access calls and returns authorization status.
 * @return {{success: boolean, errors: string[]}} Object with success status and error messages.
 */
function _triggerBenignApiAccess() {
    // const ui = SpreadsheetApp.getUi(); // No longer alerts directly
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let allAuthorized = true;
    let errors = [];

    Logger.log("Attempting to trigger API authorizations...");
    ss.toast('Checking service authorizations...', 'Authorization', 5);

    try {
        DriveApp.getRootFolder().getName();
        Logger.log("DriveApp access successful.");
    } catch (e) {
        Logger.log("DriveApp authorization needed or error: " + e.message);
        errors.push("Google Drive: " + e.message.replace("Exception: ", ""));
        allAuthorized = false;
    }

    try {
        CalendarApp.getDefaultCalendar().getName();
        Logger.log("CalendarApp access successful.");
    } catch (e) {
        Logger.log("CalendarApp authorization needed or error: " + e.message);
        errors.push("Google Calendar: " + e.message.replace("Exception: ", ""));
        allAuthorized = false;
    }

    try {
        UrlFetchApp.fetch('https://www.google.com');
        Logger.log("UrlFetchApp access successful.");
    } catch (e) {
        Logger.log("UrlFetchApp authorization needed or error: " + e.message);
        errors.push("External URL Fetch: " + e.message.replace("Exception: ", ""));
        allAuthorized = false;
    }

    try {
        MailApp.getRemainingDailyQuota();
        Logger.log("MailApp access successful.");
    } catch (e) {
        Logger.log("MailApp authorization needed or error: " + e.message);
        errors.push("MailApp (Notifications): " + e.message.replace("Exception: ", ""));
        allAuthorized = false;
    }

    if (allAuthorized) {
        ss.toast('General services seem authorized.', 'Authorization Check', 5);
    } else {
        ss.toast('Some general service authorizations may be pending.', 'Authorization Warning', 7);
    }
    return { success: allAuthorized, errors: errors };
}

/**
 * Checks access to specifically configured services, like the Google Drive folder.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet object.
 * @return {string[]} An array of messages describing access status.
 */
function _getconfiguredServiceAccessMessages(ss) {
  Logger.log("Checking access to configured services...");
  // ss.toast('Verifying configured service access...', 'Access Check', 5); // Toast moved to caller

  let messages = [];

  // Check Google Drive Folder from api_integrations.js settings
  let apiIntegrationSettings;
  try {
    if (typeof getIntegrationSettings === 'function') {
      apiIntegrationSettings = getIntegrationSettings(); // From api_integrations.js
    } else {
      Logger.log("getIntegrationSettings function not found. Skipping Drive folder check for api_integrations.js.");
      messages.push("ℹ️ Could not check api_integrations.js Drive Folder: getIntegrationSettings function is missing.");
      apiIntegrationSettings = { error: "getIntegrationSettings function not found." };
    }
  } catch (e) {
    Logger.log(`Error calling getIntegrationSettings: ${e.toString()}`);
    messages.push(`❌ Error retrieving API integration settings: ${e.message}`);
    apiIntegrationSettings = { error: e.message };
  }

  if (apiIntegrationSettings && !apiIntegrationSettings.error && apiIntegrationSettings.driveFolderId) {
    const folderIdToCheck = apiIntegrationSettings.driveFolderId;
    Logger.log(`Attempting to access configured Drive Folder (for api_integrations.js): ID ${folderIdToCheck}`);
    try {
      const folder = DriveApp.getFolderById(folderIdToCheck);
      const folderName = folder.getName();
      Logger.log(`Successfully accessed Drive Folder (api_integrations.js): "${folderName}" (ID: ${folderIdToCheck})`);
      messages.push(`✅ Successfully accessed configured Drive folder (for api_integrations.js sync): "${folderName}".`);
    } catch (e) {
      Logger.log(`Failed to access Drive Folder (api_integrations.js) ID ${folderIdToCheck}: ${e.toString()}`);
      messages.push(`❌ Error accessing configured Drive folder (for api_integrations.js sync, ID: ${folderIdToCheck}): ${e.message.replace("Exception: ", "")}. Please verify ID and script permissions.`);
    }
  } else if (apiIntegrationSettings && !apiIntegrationSettings.error && !apiIntegrationSettings.driveFolderId) {
    Logger.log("No Drive Folder ID (for api_integrations.js sync) configured in Settings. Skipping specific check.");
    messages.push("ℹ️ No Google Drive folder (for api_integrations.js sync) is configured in Settings for a specific access check.");
  } else if (apiIntegrationSettings && apiIntegrationSettings.error) {
     Logger.log(`Skipping Drive folder check due to error in getIntegrationSettings: ${apiIntegrationSettings.error}`);
     messages.push(`ℹ️ Could not check Drive folder status (api_integrations.js sync) due to: ${apiIntegrationSettings.error}`);
  }

  // --- Placeholder for checking other configured services in the future ---

  Logger.log("Finished checking access to configured services.");
  return messages;
}

/**
 * Prompts user and then triggers API authorization and configuration checks.
 * Presents a consolidated result, showing a retry modal if issues are found.
 * @param {GoogleAppsScript.Base.Ui} ui The Spreadsheet UI object.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet object.
 * @return {boolean} True if all checks passed, false otherwise or if cancelled.
 */
/**
 * Helper function to safely create HTML output for UI elements
 * This ensures consistent settings and error handling for all HTML dialogs
 * @param {string} htmlFileName The HTML file name without .html extension
 * @param {number} width Width of the dialog in pixels
 * @param {number} height Height of the dialog in pixels
 * @return {HtmlOutput} Configured HTML output object or null if error
 */
function createSafeHtmlOutput(htmlFileName, width = 600, height = 400) {
  try {
    return HtmlService.createHtmlOutputFromFile(htmlFileName)
      .setWidth(width)
      .setHeight(height)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME); // Always use IFRAME mode to prevent CSP issues
  } catch (e) {
    Logger.log(`Error creating HTML output for ${htmlFileName}: ${e.message}`);
    return null;
  }
}

function promptAndTriggerApiAuthorizations(ui, ss) {
    const response = ui.alert(
        'Service Authorization & Configuration Check',
        'The content calendar uses several Google services (Drive, Calendar, Mail, etc.) and may need to access external URLs for integrations. '+
        'You will likely be prompted by Google to authorize these permissions for the script to function correctly.\n\n'+
        'This step will also attempt to verify access to any specifically configured services (like a Google Drive folder if its ID is set in Settings).\n\n'+
        'This is typically a one-time process for your account with this script for general permissions.\n\nClick "OK" to proceed with authorization and configuration checks.',
        ui.ButtonSet.OK_CANCEL
    );

    if (response === ui.Button.OK) {
        Logger.log("User agreed to proceed with initial authorization and configuration checks.");
        ss.toast('Initiating authorization & config checks...', 'Setup Progress', -1);

        const benignAccessResult = _triggerBenignApiAccess(); // Returns {success, errors}
        const configuredAccessMessages = _getconfiguredServiceAccessMessages(ss); // Returns string[]

        const configuredServiceErrors = configuredAccessMessages.filter(m => m.startsWith("❌"));
        const overallSuccess = benignAccessResult.success && configuredServiceErrors.length === 0;

        if (overallSuccess) {
            ui.alert('Authorization & Configuration Complete', 'All necessary Google services appear authorized and configured services were accessed successfully.', ui.ButtonSet.OK);
            ss.toast('Authorization & Configuration Complete.', 'Setup Success', 7);
            PropertiesService.getUserProperties().setProperty('AUTH_AND_CONFIG_CHECKS_COMPLETED', 'true');
            Logger.log("AUTH_AND_CONFIG_CHECKS_COMPLETED property set to true after successful initial check.");
            return true;
        } else {
            Logger.log("Initial authorization/configuration check found issues. Displaying retry modal.");
            ss.toast('Authorization or configuration incomplete. Please review.', 'Setup Action Required', 10);
            // Pass the raw error arrays to the modal display function
            showAuthRetryModal_(benignAccessResult.errors, configuredAccessMessages);
            // Even if modal is shown, onOpen flow for auth is 'done' from its perspective.
            // The user property 'AUTH_AND_CONFIG_CHECKS_COMPLETED' should only be set on full success.
            return false; // Indicate that the initial guided process did not fully complete.
        }
    } else {
        Logger.log("User cancelled initial authorization & configuration check step.");
        ss.toast('Authorization & configuration check skipped by user.', 'Setup Notice', 7);
        ui.alert('Setup Notice', 'Authorization and configuration check step was skipped. Some features requiring Google service integration or specific configurations may not work until authorized/verified. You can run checks later via the App Panel.', ui.ButtonSet.OK);
        return false;
    }
}

/**
 * Wrapper function to be called from the App Panel to initiate
 * the full authorization and configured service check flow.
 * @return {boolean} The result of the general authorization attempt.
 */
function triggerFullAuthAndConfigCheckFromPanel() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const success = promptAndTriggerApiAuthorizations(ui, ss);
  
  // If general auth was successful after this user-initiated check,
  // update the property so onOpen doesn't prompt again.
  if (success) {
      PropertiesService.getUserProperties().setProperty('AUTH_AND_CONFIG_CHECKS_COMPLETED', 'true');
      Logger.log("AUTH_AND_CONFIG_CHECKS_COMPLETED property set to true via panel action.");
  }
  return success;
}

/**
 * Displays an HTML modal with authorization/configuration error messages and a retry button.
 * @param {string[]} generalAuthErrors Array of strings for general authorization errors.
 * @param {string[]} configuredServiceAccessMessages Array of strings for configured service access messages.
 */
function showAuthRetryModal_(generalAuthErrors, configuredServiceAccessMessages) {
  const ui = SpreadsheetApp.getUi();
  let htmlErrorOutput = "";

  if (generalAuthErrors && generalAuthErrors.length > 0) {
    htmlErrorOutput += '<div class="error-section">';
    htmlErrorOutput += '<div class="error-title">General Service Authorization Issues:</div>';
    htmlErrorOutput += '<div class="error-details">- ' + generalAuthErrors.join('\n- ') + '</div>';
    htmlErrorOutput += '</div>';
  }

  const configuredErrors = configuredServiceAccessMessages.filter(m => m.startsWith("❌"));
  const configuredInfo = configuredServiceAccessMessages.filter(m => m.startsWith("ℹ️"));
  // Success messages for configured services are not typically shown here if general auth failed,
  // but can be included if needed for comprehensive feedback even on partial failure.

  if (configuredErrors.length > 0) {
    htmlErrorOutput += '<div class="error-section">';
    htmlErrorOutput += '<div class="error-title">Configured Service Access Issues:</div>';
    htmlErrorOutput += '<div class="error-details">' + configuredErrors.join('\n').replace(/❌ /g, '- ') + '</div>';
    htmlErrorOutput += '</div>';
  }
  
  if (configuredInfo.length > 0 && configuredErrors.length === 0 && (!generalAuthErrors || generalAuthErrors.length === 0)) {
    // If no actual errors, but info messages, show them.
    htmlErrorOutput += '<div class="error-section">';
    htmlErrorOutput += '<div class="error-title">Configuration Notes:</div>';
    htmlErrorOutput += '<div class="error-details">' + configuredInfo.join('\n').replace(/ℹ️ /g, '- ') + '</div>';
    htmlErrorOutput += '</div>';
  }

  if (htmlErrorOutput === "") { // Should not happen if this function is called, but as a fallback
    htmlErrorOutput = "<p>An unspecified authorization or configuration step was not completed. Please try again.</p>";
  }
  
  try {
    // Create template and pass data to it
    const template = HtmlService.createTemplateFromFile('AuthRetryModal');
    template.errorMessagesHtml = htmlErrorOutput;
    
    // Evaluate the template and set IFRAME mode for security
    const evaluatedHtml = template.evaluate()
      .setWidth(600)
      .setHeight(450)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      
    ui.showModalDialog(evaluatedHtml, "Authorization & Configuration Issues");
  } catch (e) {
    Logger.log(`Error showing auth retry modal: ${e.message}`);
    // Fallback to simple toast if modal fails
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Authorization issues found. Please use the 'Authorize Services' menu option.",
      "Authorization Warning",
      10
    );
  }
}

/**
 * Called from the AuthRetryModal to re-attempt the authorization and configuration process.
 * @return {{success: boolean, errorMessagesHtml: string}} Result of the process.
 */
function rerunAuthorizationProcessFromModal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Modal requested re-run of authorization and configuration process.");

  ss.toast('Re-checking service authorizations...', 'Authorization Retry', 7);
  const benignAccessResult = _triggerBenignApiAccess(); // Returns {success, errors}

  ss.toast('Re-verifying configured service access...', 'Access Check Retry', 7);
  const configuredAccessMessages = _getconfiguredServiceAccessMessages(ss); // Returns string[]

  const overallSuccess = benignAccessResult.success && !configuredAccessMessages.some(m => m.startsWith("❌"));

  if (overallSuccess) {
    Logger.log("Re-run of authorization and configuration successful.");
    PropertiesService.getUserProperties().setProperty('AUTH_AND_CONFIG_CHECKS_COMPLETED', 'true');
    ss.toast('Authorization & Configuration Re-check Successful!', 'Success', 7);
    return { success: true, errorMessagesHtml: "<p style='color:green;'>All services authorized and configured successfully!</p>" };
  } else {
    Logger.log("Re-run of authorization and configuration still shows issues.");
    ss.toast('Authorization & Configuration Re-check Incomplete.', 'Warning', 7);
    
    let htmlErrorOutput = "";
    if (!benignAccessResult.success && benignAccessResult.errors.length > 0) {
      htmlErrorOutput += '<div class="error-section">';
      htmlErrorOutput += '<div class="error-title">General Service Authorization Issues:</div>';
      htmlErrorOutput += '<div class="error-details">- ' + benignAccessResult.errors.join('\n- ') + '</div>';
      htmlErrorOutput += '</div>';
    }
    const configuredErrors = configuredAccessMessages.filter(m => m.startsWith("❌"));
    if (configuredErrors.length > 0) {
      htmlErrorOutput += '<div class="error-section">';
      htmlErrorOutput += '<div class="error-title">Configured Service Access Issues:</div>';
      htmlErrorOutput += '<div class="error-details">' + configuredErrors.join('\n').replace(/❌ /g, '- ') + '</div>';
      htmlErrorOutput += '</div>';
    }
     if (htmlErrorOutput === "") {
         htmlErrorOutput = "<p>Retry completed, but some items may still require attention. Please check logs if issues persist.</p>";
     }
    return { success: false, errorMessagesHtml: htmlErrorOutput };
  }
}

/**
 * Initializes the content calendar with all required components.
 * @param {object} [options] Optional settings for initialization.
 * @param {boolean} [options.forceContentCalendar=false] Force re-setup of Content Calendar sheet.
 * @param {boolean} [options.forceSettings=false] Force re-setup of Settings sheet.
 * @param {boolean} [options.forceLists=false] Force re-setup of Lists sheet.
 * @param {boolean} [options.skipFeatures=false] Skip setting up feature-specific sheets/integrations.
 * @param {string[]} [options.specificFeatures=null] Array of specific features to set up.
 * @param {boolean} [options.skipAuth=false] Skip the API authorization step.
 */
function initializeContentCalendar(options = {}) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const defaults = {
    forceContentCalendar: false,
    forceSettings: false,
    forceLists: false,
    skipFeatures: false,
    specificFeatures: null,
    skipAuth: false 
  };
  const config = { ...defaults, ...options };

  let confirmationMessage = 'This will set up/verify necessary sheets and features for the content calendar. ';
  if (config.forceContentCalendar || config.forceSettings || config.forceLists) {
    confirmationMessage += 'Some sheets might be reset to their default state. ';
  }
  confirmationMessage += 'Continue?';

  // Only show confirmation if called from menu (i.e., options is empty or default)
  let calledFromMenu = Object.keys(options).length === 0 || 
                       (Object.keys(options).length === 1 && options.hasOwnProperty('skipAuth') && !options.skipAuth); // Heuristic for menu call

  if (calledFromMenu) {
      const response = ui.alert('Initialize Content Calendar', confirmationMessage, ui.ButtonSet.YES_NO);
      if (response !== ui.Button.YES) {
        Logger.log('Initialization cancelled by user.');
        ss.toast('Initialization cancelled.', 'Setup', 5);
        return;
      }
  }

  try {
    Logger.log(`Starting Initialization with options: ${JSON.stringify(config)}`);
    ss.toast('Initializing content calendar...', 'Setup Progress', -1);

    const totalSteps = 11; 
    let currentStep = 0;

    Logger.log('Initializing main sheets...');
    initializeMainSheets(ss, config.forceContentCalendar, config.forceSettings, config.forceLists);
    currentStep++;
    ss.toast('Main sheets configured.', `Setup Progress (${Math.round((currentStep/totalSteps)*100)}%)`, 5);

    if (!config.skipAuth) {
        Logger.log('Initiating API Authorizations as part of initializeContentCalendar...');
        // Calling promptAndTriggerApiAuthorizations will show UI prompts (alerts, modals for retry)
        // and will return true if basic auths passed, false otherwise or if cancelled.
        const authCallSuccessful = promptAndTriggerApiAuthorizations(ui, ss); 
        currentStep++;
        if (authCallSuccessful) {
             ss.toast('API Authorization checks completed.', `Setup Progress (${Math.round((currentStep/totalSteps)*100)}%)`, 5);
        } else {
             ss.toast('API Authorization checks were not fully completed or were cancelled.', `Setup Progress (${Math.round((currentStep/totalSteps)*100)}%)`, 5);
             // Optionally, you might want to stop further initialization if auth failed critically here.
             // For now, it will continue, and individual feature setups might fail if they need specific ungranted perms.
        }
    } else {
        Logger.log("Skipping explicit API Authorization step in initializeContentCalendar based on options (likely handled externally).");
        currentStep++; // Still count it as a "step" for progress calculation
    }

    const featuresToSetup = {
      'Search': typeof setupSearchSheet === 'function' ? setupSearchSheet : () => Logger.log('Skipping Search: setupSearchSheet not defined'),
      'API_Keys_Sheet': typeof setupApiIntegration === 'function' ? setupApiIntegration : () => Logger.log('Skipping API Keys Sheet: setupApiIntegration not defined'),
      'Notifications': typeof setupNotificationSystem === 'function' ? setupNotificationSystem : () => Logger.log('Skipping Notifications: setupNotificationSystem not defined'),
      'Templates': () => initializeTemplates(ss),
      'Dashboard': () => initializeDashboard(ss),
      'Analytics': () => initializeAnalyticsSheet(ss),
      'Archives': () => initializeArchivesSheet(ss)
    };

    if (!config.skipFeatures) {
      for (const featureName in featuresToSetup) {
        if (config.specificFeatures === null || config.specificFeatures.includes(featureName)) {
          Logger.log(`Setting up ${featureName}...`);
          try {
            featuresToSetup[featureName]();
            currentStep++;
            ss.toast(`${featureName} module set up.`, `Setup Progress (${Math.round((currentStep/totalSteps)*100)}%)`, 5);
          } catch (featureError) {
            Logger.log(`Error setting up ${featureName}: ${featureError.toString()}\n${featureError.stack}`);
            ss.toast(`Error with ${featureName}. Check logs.`, 'Setup Warning', 7);
             if (featureError instanceof TypeError && featureError.message.includes("not a function")) {
               Logger.log(`Potential missing function dependency for ${featureName}.`);
             }
            currentStep++;
          }
        } else {
             Logger.log(`Skipping ${featureName} based on specificFeatures option.`);
             currentStep++;
        }
      }
    } else {
      Logger.log('Skipping feature setups as per options.');
      currentStep += Object.keys(featuresToSetup).length; 
    }

    Logger.log('Updating Data Validation...');
    updateDataValidation();
    currentStep++;
    ss.toast('Data validation rules updated.', `Setup Progress (${Math.min(100, Math.round((currentStep/totalSteps)*100))}%)`, 5);

    Logger.log('Generating initial Week Numbers...');
    generateWeekNumbers();
    currentStep++;
    ss.toast('Week numbers generated.', `Setup Progress (${Math.min(100, Math.round((currentStep/totalSteps)*100))}%)`, 5);

    ss.toast('Initialization Complete!', 'Setup Success', 10);
    Logger.log('Initialization Complete.');
    if (calledFromMenu) {
        ui.alert('Content Calendar initialized successfully!');
    }

  } catch (error) {
    Logger.log(`Critical error during initialization: ${error.toString()}\n${error.stack}`);
    ss.toast('Initialization failed. Check Logs.', 'Error', 10);
     if (calledFromMenu) {
        if (error instanceof TypeError && error.message.includes("is not a function")) { 
             ui.alert(`Error initializing content calendar: A script tried to use an unsupported function (${error.message}). Please ensure all script files are up to date or check logs for details.`);
        } else {
            ui.alert(`Error initializing content calendar: ${error.toString()}. Please check the script logs for details.`);
        }
    } else {
        throw error;
    }
  }
}

/**
 * Initializes the main sheets for the content calendar: Content Calendar, Lists, Settings.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet object.
 * @param {boolean} forceContentCalendar If true, clears and resets the Content Calendar sheet.
 * @param {boolean} forceSettings If true, clears and resets the Settings sheet.
 * @param {boolean} forceLists If true, clears and resets the Lists sheet.
 */
function initializeMainSheets(ss, forceContentCalendar = false, forceSettings = false, forceLists = false) {
  const sheetNames = {
    calendar: 'Content Calendar',
    lists: 'Lists',
    settings: 'Settings'
  };

  let calendarSheet = ss.getSheetByName(sheetNames.calendar);
  if (!calendarSheet) {
    calendarSheet = ss.insertSheet(sheetNames.calendar, 0); // Insert as first sheet
    Logger.log(`Created sheet: ${sheetNames.calendar}`);
    setupCalendarHeaders(calendarSheet);
    setupConditionalFormatting(calendarSheet);
  } else if (forceContentCalendar) {
    Logger.log(`Forcing re-initialization of sheet: ${sheetNames.calendar}`);
    setupCalendarHeaders(calendarSheet); 
    setupConditionalFormatting(calendarSheet); 
  } else {
    Logger.log(`Sheet found: ${sheetNames.calendar}. Verifying structure.`);
    setupCalendarHeaders(calendarSheet); // Ensures headers are correct
    setupConditionalFormatting(calendarSheet); // Re-applies/verifies rules
  }

  let listsSheet = ss.getSheetByName(sheetNames.lists);
  if (!listsSheet) {
    listsSheet = ss.insertSheet(sheetNames.lists); // Insert at next available position
    Logger.log(`Created sheet: ${sheetNames.lists}`);
    setupDropdownLists(listsSheet);
  } else if (forceLists) {
    Logger.log(`Forcing re-initialization of sheet: ${sheetNames.lists}`);
    // listsSheet.clear(); // Clearing might be too destructive if user customized lists
    setupDropdownLists(listsSheet); // This will repopulate if headers are wrong or sheet is empty
  } else {
    Logger.log(`Sheet found: ${sheetNames.lists}. Verifying structure.`);
    setupDropdownLists(listsSheet); // Verifies headers and populates if empty
  }

  let settingsSheet = ss.getSheetByName(sheetNames.settings);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(sheetNames.settings);
    Logger.log(`Created sheet: ${sheetNames.settings}`);
    setupSettings(settingsSheet, false); // New sheet, use defaults
  } else if (forceSettings) {
    Logger.log(`Forcing re-initialization of sheet: ${sheetNames.settings}`);
    setupSettings(settingsSheet, false); // Force mode, reset to defaults
  } else {
    Logger.log(`Sheet found: ${sheetNames.settings}. Preserving existing settings where possible.`);
    setupSettings(settingsSheet, true); // Preserve existing valid settings
  }

  Logger.log('Applying data validation to Content Calendar sheet.');
  setupDataValidation(calendarSheet); // Must run after Lists sheet is ready
}

/**
 * Sets up the headers, formatting, and basic formulas for the Content Calendar sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Content Calendar sheet object.
 */
function setupCalendarHeaders(sheet) {
  const headerRangeToClear = sheet.getRange('A1:P2'); 
  headerRangeToClear.clearContent();                  
  headerRangeToClear.clearFormat();
  headerRangeToClear.clearDataValidations();
  // Conditional format rules on header range are generally not needed if direct formatting is applied.
  // If there were specific conditional rules ONLY for A1:P2, this would be more complex.
  // sheet.clearConditionalFormatRules(); // This clears for the whole sheet, too broad here.
  sheet.setFrozenRows(0); 

  sheet.getRange('A1:P1').merge().setValue('SOCIAL MEDIA CONTENT CALENDAR')
    .setBackground('#4285F4').setFontColor('white').setFontWeight('bold')
    .setHorizontalAlignment('center').setFontSize(14).setVerticalAlignment('middle');
  sheet.setRowHeight(1, 35);

  const headers = [
    'ID', 'Date', 'Week', 'Status', 'Channel', 'Content/Idea', 'Link to Asset',
    'Content Pillar', 'Content Format', 'Assigned To', 'Notes', 
    'Created', 'Updated', 'Status Changed', 'Event ID', 'Extra Col P' // Ensure 16 headers if P is last
  ];
  const actualHeaders = headers.slice(0, 16); 

  sheet.getRange(2, 1, 1, actualHeaders.length).setValues([actualHeaders])
    .setBackground('#EEEEEE').setFontWeight('bold').setVerticalAlignment('middle');
  sheet.setRowHeight(2, 25);

  const COLS = {
    ID: 1, DATE: 2, WEEK: 3, STATUS: 4, CHANNEL: 5, CONTENT: 6, LINK: 7,
    PILLAR: 8, FORMAT: 9, ASSIGNED: 10, NOTES: 11, CREATED: 12,
    UPDATED: 13, STATUS_CHANGED: 14, EVENT_ID: 15, COL_P: 16 
  };

  const widths = [80,100,60,150,100,350,200,150,150,120,250,150,150,150,200,100];
  widths.forEach((w, i) => sheet.setColumnWidth(i+1, w));

  sheet.setFrozenRows(2);

  const lastRow = Math.max(sheet.getMaxRows(), 1000); 
  try { 
    sheet.getRange(3, COLS.DATE, lastRow - 2, 1).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(3, COLS.CREATED, lastRow - 2, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange(3, COLS.UPDATED, lastRow - 2, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange(3, COLS.STATUS_CHANGED, lastRow - 2, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  } catch (e) {
    Logger.log("Error setting number formats: " + e);
  }

  try {
    const idFormulaRange = sheet.getRange(3, COLS.ID);
    if (idFormulaRange.isBlank() || !idFormulaRange.getFormula()) { // Set only if blank or no formula
        idFormulaRange.setFormula('=IF(B3<>"", "CONT-" & TEXT(ROW(A3)-2,"000"), "")');
        if(sheet.getLastRow() < 13 && sheet.getLastRow() >=3 ) { // Copy down a few rows if sheet is newish
             idFormulaRange.copyTo(sheet.getRange(4, COLS.ID, Math.min(10, sheet.getMaxRows() - 3) ));
        }
    }
    const weekFormulaRange = sheet.getRange(3, COLS.WEEK);
    if (weekFormulaRange.isBlank() || !weekFormulaRange.getFormula()) {
        weekFormulaRange.setFormula('=IF(B3<>"", WEEKNUM(B3,2), "")');
        if(sheet.getLastRow() < 13 && sheet.getLastRow() >=3) {
             weekFormulaRange.copyTo(sheet.getRange(4, COLS.WEEK, Math.min(10, sheet.getMaxRows() - 3)));
        }
    }
  } catch (formulaError) {
    Logger.log("Error setting initial formulas: " + formulaError);
  }
  Logger.log('Content Calendar headers and basic formatting set.');
}


/**
 * Sets up data validation rules for the Content Calendar sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Content Calendar sheet object.
 */
function setupDataValidation(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listsSheet = ss.getSheetByName('Lists');
  if (!listsSheet) {
    Logger.log('Lists sheet not found for Data Validation setup. Cannot proceed.');
    SpreadsheetApp.getUi().alert("Error: 'Lists' sheet is missing. Data validation cannot be set up.");
    return;
  }

  const getListRange = (colLetter) => {
     try {
        const colNum = listsSheet.getRange(colLetter + "1").getColumn();
        const lastRowWithData = listsSheet.getRange(colNum, 1, listsSheetgetMaxRows()).getValues().filter(String).length;
        if (lastRowWithData === 0 && colLetter === 'A') return null; // If status list is empty, special handling might be needed
        if (lastRowWithData < 1 && colLetter !== 'A') return null; // For other lists, if header is missing, it's an issue.
                                                                // Assuming header is row 1.
        return listsSheet.getRange(`${colLetter}2:${colLetter}${lastRowWithData}`);
     } catch (e) {
         Logger.log(`Error getting list range for column ${colLetter}: ${e}`);
         return null; 
     }
  };

  const lastContentRow = Math.max(sheet.getMaxRows(), 1000); 
  const COLS = { STATUS: 4, CHANNEL: 5, PILLAR: 8, FORMAT: 9, ASSIGNED: 10 };

  const applyValidation = (col, rangeSource, allowInvalid, helpText) => {
    if (rangeSource) {
        const rule = SpreadsheetApp.newDataValidation()
                      .requireValueInRange(rangeSource, true)
                      .setAllowInvalid(allowInvalid)
                      .setHelpText(helpText)
                      .build();
        sheet.getRange(3, col, lastContentRow - 2, 1).setDataValidation(rule);
    } else {
        Logger.log(`Skipping validation for column ${col} as source list is empty/invalid.`);
    }
  };

  try { applyValidation(COLS.STATUS, getListRange('A'), false, 'Select a status from the list.'); } 
  catch (e) { Logger.log("Error setting Status validation: " + e); }
  try { applyValidation(COLS.CHANNEL, getListRange('B'), false, 'Select a channel.'); }
  catch (e) { Logger.log("Error setting Channel validation: " + e); }
  try { applyValidation(COLS.PILLAR, getListRange('C'), true, 'Select a content pillar or enter a new one.'); }
  catch (e) { Logger.log("Error setting Pillar validation: " + e); }
  try { applyValidation(COLS.FORMAT, getListRange('D'), true, 'Select a content format or enter a new one.'); }
  catch (e) { Logger.log("Error setting Format validation: " + e); }
  try { applyValidation(COLS.ASSIGNED, getListRange('E'), true, 'Select a team member or enter a name.'); }
  catch (e) { Logger.log("Error setting Assigned To validation: " + e); }

  Logger.log('Data validation rules applied/updated to Content Calendar.');
}

/**
 * Sets up conditional formatting rules for the Content Calendar sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Content Calendar sheet object.
 */
function setupConditionalFormatting(sheet) {
  const lastRow = Math.max(sheet.getMaxRows(), 1000);
  sheet.clearConditionalFormatRules(); // Clear all existing rules on the sheet first
  Logger.log("Cleared all existing conditional format rules from Content Calendar sheet.");
  
  const newRules = []; 

  const COLS = { DATE: 2, STATUS: 4 };

  // --- Status Colors ---
  const statusRange = sheet.getRange(3, COLS.STATUS, lastRow - 2, 1); // Apply to column D from row 3
  const statusColors = {
    'Planned': '#FFF2CC', 
    'Copywriting Complete': '#D9EAD3', 
    'Creative Completed': '#CFE2F3', 
    'Ready for Review': '#FCE5CD', 
    'Schedule': '#D9D2E9' 
  };

  for (const status in statusColors) {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(status)
      .setBackground(statusColors[status])
      .setRanges([statusRange]) // Apply to the status column itself
      .build();
    newRules.push(rule);
  }

  // --- Date Highlighting ---
  const dateRange = sheet.getRange(3, COLS.DATE, lastRow - 2, 1); // Column B from row 3
  const dateRangeFirstCellA1 = sheet.getRange(3, COLS.DATE).getA1Notation(); // e.g., B3
  const statusRangeFirstCellA1 = sheet.getRange(3, COLS.STATUS).getA1Notation(); // e.g., D3

  const todayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenDateEqualTo(SpreadsheetApp.RelativeDate.TODAY)
    .setBackground('#FFD966') 
    .setBold(true)
    .setRanges([dateRange]) // Apply to date column
    .build();
  newRules.push(todayRule);

  const overdueFormula = `=AND(${dateRangeFirstCellA1}<TODAY(), ${dateRangeFirstCellA1}<>"", ${statusRangeFirstCellA1}<>"Schedule")`;
  const overdueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(overdueFormula)
    .setBackground('#F4CCCC') 
    .setRanges([dateRange]) // Apply format only to the date cell
    .build();
  newRules.push(overdueRule);

  sheet.setConditionalFormatRules(newRules);
  Logger.log(`Applied ${newRules.length} new conditional formatting rules.`);
}


/**
 * Sets up the Lists sheet with headers and default values for dropdowns.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Lists sheet object.
 */
function setupDropdownLists(sheet) {
  const listHeaders = ['Status', 'Channel', 'Content Pillar', 'Content Format', 'Team Members'];
  const headerRange = sheet.getRange(1, 1, 1, listHeaders.length);
  const currentHeaders = headerRange.getValues()[0];
  let headersMatch = listHeaders.every((h, i) => h === currentHeaders[i]);

  if (!headersMatch) {
      headerRange.setValues([listHeaders])
        .setBackground('#4285F4').setFontColor('white').setFontWeight('bold').setVerticalAlignment('middle');
      sheet.setRowHeight(1, 25);
      Logger.log('Lists sheet headers set/corrected.');
  } else {
      // Ensure formatting even if headers are present
      headerRange.setBackground('#4285F4').setFontColor('white').setFontWeight('bold').setVerticalAlignment('middle');
      sheet.setRowHeight(1, 25);
  }

  const populateListIfEmpty = (col, defaultValues) => {
      const firstDataCell = sheet.getRange(2, col);
      if (firstDataCell.isBlank()) {
          sheet.getRange(2, col, defaultValues.length, 1).setValues(defaultValues.map(v => [v]));
          Logger.log(`Default values added for column ${col} in Lists sheet.`);
      }
  };

  populateListIfEmpty(1, ['Planned', 'Copywriting Complete', 'Creative Completed', 'Ready for Review', 'Schedule']);
  populateListIfEmpty(2, ['Twitter', 'YouTube', 'Telegram']);
  populateListIfEmpty(3, ['Educational', 'Promotional', 'Community', 'Industry News', 'Product Updates', 'Case Studies', 'Tutorials', 'Thought Leadership', 'Entertainment', 'Behind the Scenes']);
  populateListIfEmpty(4, ['Text Post', 'Image', 'Video', 'Poll', 'Thread', 'Live', 'Story', 'Short', 'Article', 'Infographic', 'Carousel']);
  populateListIfEmpty(5, ['Content Manager', 'Copywriter', 'Designer', 'Video Editor', 'Social Media Manager', 'Review Team', 'Marketing Director', 'Community Manager', 'SEO Specialist', 'Guest Author']);

  const colWidths = [150, 100, 150, 150, 150];
  colWidths.forEach((w,i) => { if(sheet.getColumnWidth(i+1) !== w) sheet.setColumnWidth(i+1, w);});

  if (sheet.getFrozenRows() !== 1) {
      sheet.setFrozenRows(1);
  }
  Logger.log('Lists sheet populated/verified.');
}


/**
 * Sets up the Settings sheet with default structure and placeholders.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Settings sheet object.
 * @param {boolean} [preserveExisting=false] If true, existing values will not be overwritten.
 */
function setupSettings(sheet, preserveExisting = false) {
  Logger.log(`Setting up Settings sheet with preserveExisting=${preserveExisting}`);

  const titleRange = sheet.getRange('A1:C1'); // Assuming C is last relevant column for title
  if (!titleRange.isPartOfMerge() || titleRange.getValue() !== 'CONTENT CALENDAR SETTINGS') {
    // Ensure C1 is not part of another merge before attempting to merge A1:C1
    try { sheet.getRange('C1').breakApart(); } catch(e){}
    try { titleRange.breakApart(); } catch(e){} // Break current merge if it exists
    
    titleRange.merge().setValue('CONTENT CALENDAR SETTINGS')
      .setBackground('#4285F4').setFontColor('white').setFontWeight('bold')
      .setHorizontalAlignment('center').setFontSize(14).setVerticalAlignment('middle');
    sheet.setRowHeight(1, 35);
  }

  // Define all settings and their target cells for consistency
  const settingsStructure = [
    { section: 'GENERAL SETTINGS', range: 'A3:C3', items: [
        { row: 4, label: 'Company/Brand Name:', key: 'brandName', targetCell: 'B4', defaultValue: 'Your Company Name' },
        { row: 5, label: 'Calendar Start Date:', key: 'startDate', targetCell: 'B5', defaultValue: new Date(), format: 'yyyy-mm-dd' },
        { row: 6, label: 'Calendar End Date:', key: 'endDate', targetCell: 'B6', defaultValue: (() => { const d = new Date(); d.setFullYear(d.getFullYear() + 1); return d; })(), format: 'yyyy-mm-dd' },
      ]
    },
    { section: 'NOTIFICATIONS', range: 'A8:C8', items: [
        { row: 9, label: 'Primary Notification Email:', key: 'notifyEmail', targetCell: 'B9', defaultValue: Session.getActiveUser().getEmail() }, // NOTIFICATION_CONFIG.SETTINGS_EMAIL_CELL in email_notification_system.js is B7 - needs alignment
        { row: 10, label: 'Send Status Change Alerts:', key: 'notifyStatus', targetCell: 'B10', defaultValue: true, type: 'checkbox' },
        { row: 11, label: 'Send Deadline Alerts (Days Before):', key: 'deadlineAlertDays', targetCell: 'B11', defaultValue: 3 },
        { row: 12, label: 'Send Overdue Alerts:', key: 'notifyOverdue', targetCell: 'B12', defaultValue: true, type: 'checkbox' },
        { row: 13, label: 'Send Weekly Summary:', key: 'notifyWeekly', targetCell: 'B13', defaultValue: true, type: 'checkbox' }
      ]
    },
     { section: 'INTEGRATIONS', range: 'A15:C15', items: [
        // Note: Ensure targetCell values match consts in other scripts for consistency
        { row: 16, label: 'Google Calendar ID (Optional):', key: 'calendarId', targetCell: 'B8', defaultValue: '' }, // CALENDAR_CONFIG.SETTINGS_CALENDAR_ID_CELL ('B8')
        // Row 17 (old driveFolderIdPicker) is REMOVED
        
        { row: 18, label: 'Twitter API Key:', key: 'twitterApiKeyCell', targetCell: 'B13', sourceFile: 'api_integrations.js', defaultValue: '' }, // This label's value should go to B13
        { row: 19, label: 'Twitter API Secret:', key: 'twitterApiSecretCell', targetCell: 'B14', sourceFile: 'api_integrations.js', defaultValue: '' },
        { row: 20, label: 'YouTube API Key:', key: 'youtubeApiKeyCell', targetCell: 'B15', sourceFile: 'api_integrations.js', defaultValue: '' },
        { row: 21, label: 'Telegram Bot Token:', key: 'telegramBotTokenCell', targetCell: 'B16', sourceFile: 'api_integrations.js', defaultValue: '' },
        // Rows related to B18-B20 (Drive for API + Sync Time) are managed by setupApiIntegration
        { row: 24, label: 'Last Analytics Sync:', key: 'lastAnalyticsSyncCell', targetCell: 'B20', sourceFile: 'api_integrations.js', defaultValue: 'Never', readOnly: true }
      ]
    },
    { section: 'WORKFLOW DEFAULTS', range: 'A26:C26', items: [
        { row: 27, label: 'Default Assigned (Planned):', key: 'defaultPlanned', targetCell: 'B27', defaultValue: 'Content Manager' },
        { row: 28, label: 'Default Assigned (Copy Done):', key: 'defaultCopyDone', targetCell: 'B28', defaultValue: 'Designer' },
        { row: 29, label: 'Default Assigned (Creative Done):', key: 'defaultCreativeDone', targetCell: 'B29', defaultValue: 'Review Team' },
        { row: 30, label: 'Default Assigned (Review Done):', key: 'defaultReviewDone', targetCell: 'B30', defaultValue: 'Social Media Manager' },
      ]
    },
     { section: 'API INSTRUCTIONS', range: 'A32:C32', isInstructions: true, text: [
        ['API Integration Instructions:'], 
        ['1. Enter your API keys/tokens for the services you want to integrate with in the fields above (rows 18-22).'], 
        ['2. For "Google Drive Folder ID (Assets)", this is for the `api_integrations.js` GDrive sync. The "Google Drive Assets Folder ID" (row 17) is for `drive_integration_script.js` picker.'],
        ['3. Use the "Integrations" menu (if available via App Panel) or specific setup functions for detailed connections.'],
      ]
    },
  ];

  const existingSettings = {};
  if (preserveExisting && sheet.getLastRow() > 1) {
    settingsStructure.forEach(group => {
      if (group.items) {
        group.items.forEach(item => {
          if (!item.readOnly) {
            const cellAddress = item.targetCell || `B${item.row}`; // Fallback, but targetCell should be defined
            try {
              const value = sheet.getRange(cellAddress).getValue();
              if (value !== undefined && value !== null && value !== "") {
                existingSettings[item.key] = value;
              }
            } catch (e) {
              Logger.log(`Could not read existing setting for ${item.key} from ${cellAddress}: ${e}`);
            }
          }
        });
      }
    });
    Logger.log(`Found ${Object.keys(existingSettings).length} settings to preserve by reading target cells.`);
  }

  settingsStructure.forEach(group => {
    const sectionRange = sheet.getRange(group.range);
     if (!sectionRange.isPartOfMerge() || sectionRange.getValue() !== group.section) {
         try{sectionRange.breakApart();}catch(e){}
         sectionRange.merge().setValue(group.section)
           .setBackground('#EEEEEE').setFontWeight('bold').setHorizontalAlignment('left');
     }

    if (group.isInstructions && group.text) {
         const startRow = parseInt(group.range.match(/\d+/)[0]) + 1; 
         const instructionRange = sheet.getRange(startRow, 1, group.text.length, 1);
         instructionRange.setValues(group.text);
         sheet.getRange(startRow, 1, group.text.length, sheet.getRange(group.range).getNumColumns()).mergeAcross(); 
         Logger.log(`Wrote instructions for ${group.section}`);
    } else if (group.items) {
        group.items.forEach(item => {
          const labelCell = sheet.getRange(item.row, 1); 
          const valueCell = sheet.getRange(item.targetCell); // Must use targetCell

          if (labelCell.getValue() !== item.label) {
              labelCell.setValue(item.label).setFontWeight('bold').setHorizontalAlignment('right');
          }

          let valueToSet = item.defaultValue;
          if (preserveExisting && existingSettings.hasOwnProperty(item.key)) {
              const preserved = existingSettings[item.key];
              if (item.type === 'checkbox') {
                 valueToSet = (preserved === true || preserved?.toString().toLowerCase() === 'true');
              } else if (item.format === 'yyyy-mm-dd' && !(preserved instanceof Date)) {
                  try {
                      const parsedDate = new Date(preserved);
                      valueToSet = isNaN(parsedDate.valueOf()) ? item.defaultValue : parsedDate;
                  } catch (_) { valueToSet = item.defaultValue; }
              } else {
                 valueToSet = preserved;
              }
          }
          
          // Set value only if not preserving, or if it's readOnly, or if blank, or if value actually changed
          const currentValue = valueCell.getValue();
          let needsUpdate = !preserveExisting || item.readOnly || valueCell.isBlank();
          if (preserveExisting && !item.readOnly && !valueCell.isBlank()) {
              if (item.type === 'checkbox') {
                  needsUpdate = valueCell.isChecked() !== valueToSet; // For checkboxes
              } else if (valueToSet instanceof Date && currentValue instanceof Date) {
                  needsUpdate = valueToSet.getTime() !== currentValue.getTime();
              } else {
                  needsUpdate = currentValue?.toString() !== valueToSet?.toString();
              }
          }

          if (needsUpdate) {
              if (item.type === 'checkbox') {
                 // Check if it's not already a checkbox
                 const dataValidation = valueCell.getDataValidation();
                 if (!dataValidation || dataValidation.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
                    // Create checkbox validation
                    const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
                    valueCell.setDataValidation(rule);
                 }
                 valueCell.setValue(valueToSet); // TRUE/FALSE for checkbox
              } else {
                 if (valueCell.getDataValidation()?.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
                     valueCell.setDataValidation(null); // Remove checkbox validation
                 }
                 valueCell.setValue(valueToSet);
                 if (item.format) {
                     valueCell.setNumberFormat(item.format);
                 }
              }
              Logger.log(`Setting value for ${item.key} in ${item.targetCell} to: ${valueToSet}`);
          }

           if(item.readOnly) {
              valueCell.setFontStyle('italic').setFontColor('#666666');
           } else if (item.type !== 'checkbox') { 
               valueCell.setFontStyle('normal').setFontColor(null); 
           }
        });
    }
  });

  // Ensure column widths are set after all content
  const finalColWidths = {1: 250, 2: 350, 3: 50}; // Col A, B, C
  for (const col in finalColWidths) {
      if(sheet.getColumnWidth(parseInt(col)) !== finalColWidths[col]) {
          sheet.setColumnWidth(parseInt(col), finalColWidths[col]);
      }
  }
  Logger.log(`Settings sheet configured.`);
}


/**
 * Initializes the templates system by creating the sheet if needed.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet object.
 */
function initializeTemplates(ss) {
  const sheetName = 'Content Templates';
  let templatesSheet = ss.getSheetByName(sheetName);
  if (!templatesSheet) {
    templatesSheet = ss.insertSheet(sheetName);
    Logger.log(`Created sheet: ${sheetName}`);
    setupTemplatesSheet(templatesSheet); 
  } else {
    Logger.log(`Sheet found: ${sheetName}. Verifying headers.`);
    setupTemplatesSheet(templatesSheet); 
  }
}

/**
 * Sets up the Content Templates sheet with headers and examples.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Content Templates sheet object.
 */
function setupTemplatesSheet(sheet) {
  sheet.setFrozenRows(0); 

  const headers = ['Channel', 'Format', 'Template'];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
   if (headerRange.getValues()[0].join('') !== headers.join('')) {
       headerRange.setValues([headers])
         .setBackground('#4285F4').setFontColor('white').setFontWeight('bold').setVerticalAlignment('middle');
       sheet.setRowHeight(1, 25);
       Logger.log("Templates sheet headers set/corrected.");
   } else {
        headerRange.setBackground('#4285F4').setFontColor('white').setFontWeight('bold').setVerticalAlignment('middle');
        sheet.setRowHeight(1, 25);
   }

  if (sheet.getColumnWidth(1) !== 150) sheet.setColumnWidth(1, 150); 
  if (sheet.getColumnWidth(2) !== 150) sheet.setColumnWidth(2, 150); 
  if (sheet.getColumnWidth(3) !== 500) sheet.setColumnWidth(3, 500); 

  const exampleTemplates = [
    ['Twitter', 'Text Post', 'Main point: \nKey message (under 280 chars): \nHashtags: \nMention: '],
    ['Twitter', 'Image', 'Main point: \nImage description: \nCaption (under 280 chars): \nHashtags: '],
    ['YouTube', 'Video', 'Video title: \nDescription: \n\nIntro (0:00-0:30): \nMain points: \n- Point 1 (0:30-2:00): \n- Point 2 (2:00-4:00): \n- Point 3 (4:00-6:00): \nConclusion (6:00-7:00): \n\nTags: \nCategory: '],
    ['Telegram', 'Text Post', 'Title: \n\nMain content: \n\nKey points: \n- \n- \n- \n\nCall to action: ']
  ];
  if (sheet.getLastRow() < 2 && exampleTemplates.length > 0) {
      sheet.getRange(2, 1, exampleTemplates.length, 3).setValues(exampleTemplates).setWrap(true);
      Logger.log("Example templates added as Templates sheet was empty.");
  }

  if (sheet.getFrozenRows() !== 1) {
    sheet.setFrozenRows(1);
  }
  Logger.log('Content Templates sheet set up/verified.');
}

/**
 * Initializes the dashboard sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet object.
 */
function initializeDashboard(ss) {
  const sheetName = 'Dashboard';
  let dashboardSheet = ss.getSheetByName(sheetName);
  if (!dashboardSheet) {
    dashboardSheet = ss.insertSheet(sheetName);
    Logger.log(`Created sheet: ${sheetName}`);
    setupDashboardLayout(dashboardSheet, ss);  // Pass ss
  } else {
    Logger.log(`Sheet found: ${sheetName}. Verifying layout.`);
    setupDashboardLayout(dashboardSheet, ss); // Pass ss
  }
}

/**
 * Builds the WHERE clause for dashboard QUERY formulas based on filter selections.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboardSheet The dashboard sheet object.
 * @param {string} dateCol The letter of the date column in 'Content Calendar' (e.g., 'B').
 * @param {string} channelCol The letter of the channel column (e.g., 'E').
 * @param {string} statusCol The letter of the status column (e.g., 'D').
 * @return {string} The constructed WHERE clause string.
 */
function buildDashboardQueryWhereClause(dashboardSheet, dateCol, channelCol, statusCol) {
    let conditions = ["A IS NOT NULL"]; // Start with a base condition that's always true for data rows
    const today = new Date();
    const timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

    // Date Range Filter (B3)
    const dateRangeFilter = dashboardSheet.getRange('B3').getValue();
    switch (dateRangeFilter) {
        case 'This Week': // ISO week: Monday to Sunday
            const currentDay = today.getDay(); // Sunday = 0, Monday = 1, ...
            const daysToMonday = currentDay === 0 ? -6 : 1 - currentDay;
            const weekStart = new Date(today.setDate(today.getDate() + daysToMonday));
            const weekEnd = new Date(weekStart);
            weekEnd.setDate(weekStart.getDate() + 6);
            conditions.push(`${dateCol} >= date '${Utilities.formatDate(weekStart, timezone, "yyyy-MM-dd")}'`);
            conditions.push(`${dateCol} <= date '${Utilities.formatDate(weekEnd, timezone, "yyyy-MM-dd")}'`);
            break;
        case 'Last 7 Days':
            const sevenDaysAgo = new Date(today.getTime() - 6 * 24 * 60 * 60 * 1000);
            conditions.push(`${dateCol} >= date '${Utilities.formatDate(sevenDaysAgo, timezone, "yyyy-MM-dd")}'`);
            conditions.push(`${dateCol} <= date '${Utilities.formatDate(new Date(), timezone, "yyyy-MM-dd")}'`); // Use new Date() for current day, as today might have been modified
            break;
        case 'Next 7 Days':
            const sevenDaysHence = new Date(today.getTime() + 6 * 24 * 60 * 60 * 1000);
            conditions.push(`${dateCol} >= date '${Utilities.formatDate(new Date(), timezone, "yyyy-MM-dd")}'`);
            conditions.push(`${dateCol} <= date '${Utilities.formatDate(sevenDaysHence, timezone, "yyyy-MM-dd")}'`);
            break;
        case 'This Month':
            const firstDayThisMonth = new Date(today.getFullYear(), today.getMonth(), 1);
            const lastDayThisMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);
            conditions.push(`${dateCol} >= date '${Utilities.formatDate(firstDayThisMonth, timezone, "yyyy-MM-dd")}'`);
            conditions.push(`${dateCol} <= date '${Utilities.formatDate(lastDayThisMonth, timezone, "yyyy-MM-dd")}'`);
            break;
        case 'Last Month':
            const firstDayLastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
            const lastDayLastMonth = new Date(today.getFullYear(), today.getMonth(), 0);
            conditions.push(`${dateCol} >= date '${Utilities.formatDate(firstDayLastMonth, timezone, "yyyy-MM-dd")}'`);
            conditions.push(`${dateCol} <= date '${Utilities.formatDate(lastDayLastMonth, timezone, "yyyy-MM-dd")}'`);
            break;
    }

    // Channel Filter (D3)
    const channelFilter = dashboardSheet.getRange('D3').getValue();
    if (channelFilter && channelFilter !== 'All Channels') {
        conditions.push(`${channelCol} = '${channelFilter.replace(/'/g, "''")}'`); // Escape single quotes
    }

    // Status Filter (F3)
    const statusFilter = dashboardSheet.getRange('F3').getValue();
    if (statusFilter && statusFilter !== 'All Statuses') {
        conditions.push(`${statusCol} = '${statusFilter.replace(/'/g, "''")}'`); // Escape single quotes
    }
    
    return conditions.length > 1 ? "WHERE " + conditions.join(" AND ") : "WHERE " + conditions[0]; // Always include "A IS NOT NULL"
}


/**
 * Sets up the basic layout and formulas for the Dashboard sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Dashboard sheet object.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet object.
 */
function setupDashboardLayout(sheet, ss) { 
    const calendarSheetName = 'Content Calendar';
    const listsSheetName = 'Lists'; 
    const analyticsSheetName = 'Analytics';

    sheet.clearContents().clearFormats();
    sheet.getCharts().forEach(chart => sheet.removeChart(chart));
    sheet.clearConditionalFormatRules(); 
    Logger.log("Dashboard sheet cleared for new layout.");

    sheet.getRange('A1:H1').merge().setValue('CONTENT CALENDAR DASHBOARD')
        .setBackground('#2a56c6').setFontColor('white').setFontWeight('bold') 
        .setHorizontalAlignment('center').setFontSize(18).setVerticalAlignment('middle');
    sheet.setRowHeight(1, 45);

    sheet.getRange('A2').setValue('Last Updated:').setFontWeight('bold').setHorizontalAlignment('right');
    sheet.getRange('B2').setValue('Never').setFontStyle('italic'); 
    sheet.getRange('G2:H2').merge().setValue('🔄 REFRESH DASHBOARD') 
        .setBackground('#4CAF50').setFontColor('white').setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,null,null, '#388E3C', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.setRowHeight(2, 30);
    
    sheet.getRange('A3:H3').setBackground('#f0f0f0'); 
    sheet.setRowHeight(3, 30);

    sheet.getRange('A3').setValue('Date Range:').setFontWeight('bold').setHorizontalAlignment('right').setVerticalAlignment('middle');
    const dateRangeOptions = ['All Time', 'This Week', 'Last 7 Days', 'Next 7 Days', 'This Month', 'Last Month'];
    const dateRangeRule = SpreadsheetApp.newDataValidation().requireValueInList(dateRangeOptions, true).setAllowInvalid(false).build();
    sheet.getRange('B3').setDataValidation(dateRangeRule).setValue('All Time').setVerticalAlignment('middle'); 

    sheet.getRange('C3').setValue('Channel:').setFontWeight('bold').setHorizontalAlignment('right').setVerticalAlignment('middle');
    const listsSheet = ss.getSheetByName(listsSheetName); 
    if (listsSheet) {
        const channelList = listsSheet.getRange('B2:B' + listsSheet.getLastRow()).getValues().filter(row => row[0] !== "").map(r => r[0]);
        const channelRule = SpreadsheetApp.newDataValidation().requireValueInList(['All Channels'].concat(channelList), true).setAllowInvalid(false).build();
        sheet.getRange('D3').setDataValidation(channelRule).setValue('All Channels').setVerticalAlignment('middle');
    } else {
        sheet.getRange('D3').setValue('All Channels'); 
        Logger.log("Lists sheet not found for Channel filter dropdown on Dashboard.");
    }

    sheet.getRange('E3').setValue('Status:').setFontWeight('bold').setHorizontalAlignment('right').setVerticalAlignment('middle');
     if (listsSheet) {
        const statusList = listsSheet.getRange('A2:A' + listsSheet.getLastRow()).getValues().filter(row => row[0] !== "").map(r => r[0]);
        const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['All Statuses'].concat(statusList), true).setAllowInvalid(false).build();
        sheet.getRange('F3').setDataValidation(statusRule).setValue('All Statuses').setVerticalAlignment('middle');
    } else {
        sheet.getRange('F3').setValue('All Statuses'); 
        Logger.log("Lists sheet not found for Status filter dropdown on Dashboard.");
    }
    sheet.getRange('B3, D3, F3').setBackground('white'); 

    sheet.setRowHeight(4, 10);

    sheet.getRange('A5:H5').merge().setValue('📊 KEY PERFORMANCE INDICATORS')
        .setBackground('#e9edf3').setFontColor('#333333').setFontWeight('bold') 
        .setHorizontalAlignment('center').setFontSize(14).setVerticalAlignment('middle');
    sheet.setRowHeight(5, 35);

    const kpiBaseRow = 6;
    
    const kpiValues = [
        { label: 'Total Content (Filtered)', formulaBase: `=COUNTA(IFERROR(QUERY('${calendarSheetName}'!A3:A, "SELECT A %WHERE_CLAUSE% ",0)))`, whereDate: 'A', whereChannel: '', whereStatus: '', col: 'B' },
        { label: 'Published (Filtered)', formulaBase: `=COUNTA(IFERROR(QUERY('${calendarSheetName}'!A3:D, "SELECT A %WHERE_CLAUSE_AND_STATUS_SCHEDULED%",0)))`, whereDate: 'B', whereChannel: 'E', whereStatus: 'D', col: 'D' },
        { label: 'Avg. Engagement (All)', formulaBase: `=IFERROR(AVERAGE('${analyticsSheetName}'!F2:F),"N/A")`, col: 'F' }, // Not easily filterable by dashboard controls without complex queries
        { label: 'Upcoming (Next 7 Days, Not Scheduled)', formulaBase: `=COUNTA(IFERROR(QUERY('${calendarSheetName}'!A3:D, "SELECT A WHERE B >= TODAY() AND B <= TODAY()+7 AND D <> 'Schedule' %AND_WHERE_CLAUSE%",0)))`, whereDate: 'B', whereChannel: 'E', whereStatus: 'D', col: 'H' }
    ];
        
    kpiValues.forEach((kpi, index) => {
        const targetLabelCell = sheet.getRange(kpiBaseRow, String.fromCharCode(65 + index * 2)); 
        const targetValueCell = sheet.getRange(kpiBaseRow + 1, String.fromCharCode(65 + index * 2));
        
        let kpiWhereClause = buildDashboardQueryWhereClause(sheet, kpi.whereDate, kpi.whereChannel, kpi.whereStatus);
        // Adjust whereClause for specific KPI needs
        let formula = kpi.formulaBase;
        if (kpi.label === 'Published (Filtered)') {
            let tempWhere = kpiWhereClause.includes("WHERE") ? kpiWhereClause + ` AND D = 'Schedule'` : `WHERE D = 'Schedule'`;
            formula = kpi.formulaBase.replace('%WHERE_CLAUSE_AND_STATUS_SCHEDULED%', tempWhere);
        } else if (kpi.label === 'Upcoming (Next 7 Days, Not Scheduled)') {
            let tempWhere = kpiWhereClause.includes("WHERE A IS NOT NULL") ? kpiWhereClause.replace("WHERE A IS NOT NULL", "") : kpiWhereClause.replace("WHERE ", "AND ");
            tempWhere = tempWhere.startsWith(" AND ") ? tempWhere : (tempWhere ? " AND " + tempWhere : ""); // ensure AND if other conditions exist
            formula = kpi.formulaBase.replace('%AND_WHERE_CLAUSE%', tempWhere);
        } else {
            formula = kpi.formulaBase.replace('%WHERE_CLAUSE%', kpiWhereClause);
        }

        targetLabelCell.mergeAcross().setValue(kpi.label).setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setBorder(true,true,true,true,null,null,'#cccccc', SpreadsheetApp.BorderStyle.SOLID);
        targetValueCell.mergeAcross().setFormula(formula.replace(/"/g, '""')) // Double escape quotes for QUERY
                       .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setNumberFormat("0").setBorder(null,true,true,true,null,null,'#cccccc', SpreadsheetApp.BorderStyle.SOLID);
        sheet.setRowHeight(kpiBaseRow, 25);
        sheet.setRowHeight(kpiBaseRow+1, 40);
    });

    sheet.setRowHeight(kpiBaseRow + 2, 10); 

    sheet.getRange('A11:H11').merge().setValue('📈 VISUALIZATIONS')
        .setBackground('#e9edf3').setFontColor('#333333').setFontWeight('bold')
        .setHorizontalAlignment('center').setFontSize(14).setVerticalAlignment('middle');
    sheet.setRowHeight(11, 35);

    sheet.getRange('A12:D12').merge().setValue('Content Status Distribution').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange('A13:D25').setBackground('#f8f9fa').setBorder(true,true,true,true,null,null,'#cccccc', SpreadsheetApp.BorderStyle.SOLID_THICK); 

    sheet.getRange('E12:H12').merge().setValue('Content by Channel').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange('E13:H25').setBackground('#f8f9fa').setBorder(true,true,true,true,null,null,'#cccccc', SpreadsheetApp.BorderStyle.SOLID_THICK); 
    
    sheet.setRowHeight(26, 10);

    sheet.getRange('A27:H27').merge().setValue('📝 CONTENT DETAILS & TASKS')
        .setBackground('#e9edf3').setFontColor('#333333').setFontWeight('bold')
        .setHorizontalAlignment('center').setFontSize(14).setVerticalAlignment('middle');
    sheet.setRowHeight(27, 35);

    const upcomingRowStart = 28;
    sheet.getRange(upcomingRowStart, 1, 1, 8).merge().setValue('Upcoming Content (Filtered by Dashboard Controls)').setFontWeight('bold');
    const tableHeaders = ['ID', 'Date', 'Channel', 'Content', 'Status', 'Assigned To', 'Pillar', 'Link'];
    sheet.getRange(upcomingRowStart + 1, 1, 1, tableHeaders.length).setValues([tableHeaders])
        .setBackground('#d9d9d9').setFontWeight('bold');
    
    const whereClauseUpcoming = buildDashboardQueryWhereClause(sheet, 'B', 'E', 'D'); 
    const upcomingQueryFull = `SELECT A, B, E, F, D, J, H, G ${whereClauseUpcoming} ORDER BY B ASC LIMIT 10`; 
    const upcomingQueryFormula = `=IFERROR(QUERY('${calendarSheetName}'!A3:L, "${upcomingQueryFull.replace(/"/g, '""')}", 0), "No upcoming content matching filters.")`;
    sheet.getRange(upcomingRowStart + 2, 1).setFormula(upcomingQueryFormula);
    sheet.getRange(upcomingRowStart + 1, 1, 12, tableHeaders.length).setBorder(true,true,true,true,true,true,'#cccccc', SpreadsheetApp.BorderStyle.SOLID);

    const overdueRowStart = upcomingRowStart + 13;
    sheet.getRange(overdueRowStart, 1, 1, 8).merge().setValue('Overdue / Needs Attention (Filtered by Dashboard Controls)').setFontWeight('bold');
    sheet.getRange(overdueRowStart + 1, 1, 1, tableHeaders.length).setValues([tableHeaders]) 
        .setBackground('#d9d9d9').setFontWeight('bold');

    let whereClauseOverdue = buildDashboardQueryWhereClause(sheet, 'B', 'E', 'D');
    const todayFormatted = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    if (whereClauseOverdue) {
        whereClauseOverdue += ` AND B < date '${todayFormatted}' AND D <> 'Schedule'`;
    } else {
        whereClauseOverdue = `WHERE B < date '${todayFormatted}' AND D <> 'Schedule'`;
    }
    const overdueQueryFull = `SELECT A, B, E, F, D, J, H, G ${whereClauseOverdue} ORDER BY B ASC LIMIT 10`;
    const overdueQueryFormula = `=IFERROR(QUERY('${calendarSheetName}'!A3:L, "${overdueQueryFull.replace(/"/g, '""')}", 0), "No overdue items matching filters.")`;
    sheet.getRange(overdueRowStart + 2, 1).setFormula(overdueQueryFormula);
    sheet.getRange(overdueRowStart + 1, 1, 12, tableHeaders.length).setBorder(true,true,true,true,true,true,'#cccccc', SpreadsheetApp.BorderStyle.SOLID);

    const colWidths = [80, 90, 100, 200, 120, 120, 100, 100]; // Adjusted for 8 columns
    colWidths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
    
    sheet.setFrozenRows(3);

    // Chart Helper Ranges (e.g., starting in K1)
    const chartDataStartCol = "K";
    sheet.getRange(`${chartDataStartCol}1`).setValue("Status (Chart Data)").setFontWeight("bold");
    const statusesForChart = listsSheet ? listsSheet.getRange('A2:A' + listsSheet.getLastRow()).getValues().filter(r => r[0]).map(r => r[0]) : ['Planned', 'Schedule']; 
    let statusChartFormulas = [["Status", "Count"]];
    statusesForChart.forEach(status => {
        let cond = buildDashboardQueryWhereClause(sheet, 'B', 'E', 'D'); // Get base filters
        cond = cond.replace("WHERE A IS NOT NULL AND ", "WHERE ").replace("WHERE A IS NOT NULL", ""); // Clean base condition
        cond = cond ? cond + ` AND D = '${status.replace(/'/g, "''")}'` : `WHERE D = '${status.replace(/'/g, "''")}'`;
        statusChartFormulas.push([status, `=IFERROR(ROWS(QUERY('${calendarSheetName}'!A3:D, "SELECT A ${cond.replace(/"/g, '""')}",0)),0)`]);
    });
    sheet.getRange(`${chartDataStartCol}2`).offset(0,0,statusChartFormulas.length, 2).setValues(statusChartFormulas);
    
    const channelChartStartCol = "M";
    sheet.getRange(`${channelChartStartCol}1`).setValue("Channel (Chart Data)").setFontWeight("bold");
    const channelsForChart = listsSheet ? listsSheet.getRange('B2:B' + listsSheet.getLastRow()).getValues().filter(r=>r[0]).map(r => r[0]) : ['Twitter', 'YouTube']; 
    let channelChartFormulas = [["Channel", "Count"]];
    channelsForChart.forEach(channel => {
        let cond = buildDashboardQueryWhereClause(sheet, 'B', 'E', 'D');
        cond = cond.replace("WHERE A IS NOT NULL AND ", "WHERE ").replace("WHERE A IS NOT NULL", "");
        cond = cond ? cond + ` AND E = '${channel.replace(/'/g, "''")}'` : `WHERE E = '${channel.replace(/'/g, "''")}'`;
        channelChartFormulas.push([channel, `=IFERROR(ROWS(QUERY('${calendarSheetName}'!A3:E, "SELECT A ${cond.replace(/"/g, '""')}",0)),0)`]);
    });
    sheet.getRange(`${channelChartStartCol}2`).offset(0,0,channelChartFormulas.length, 2).setValues(channelChartFormulas);
    
    addDashboardCharts(sheet, 0); 
    sheet.hideColumns(sheet.getRange(chartDataStartCol+"1").getColumn(), 4); // Hide K, L, M, N (4 columns)

    Logger.log('New dashboard layout set up with dynamic query placeholders.');
}

/**
 * Adds charts to the dashboard.
 * This version attempts to use the dynamic helper ranges.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Dashboard sheet object.
 */
function addDashboardCharts(sheet, _metricsRowIgnored) { 
  const ui = SpreadsheetApp.getUi();

  try {
    // Status Chart
    const statusChartDataHeaderCell = sheet.getRange('K1'); // Helper data for status chart
    const statusLastRow = statusChartDataHeaderCell.getDataRegion(SpreadsheetApp.Dimension.ROWS).getLastRow();
    const statusChartDataSource = sheet.getRange('K1:L' + statusLastRow);

    if (statusLastRow > 1) { 
        let statusChart = sheet.newChart()
          .setChartType(Charts.ChartType.PIE)
          .addRange(statusChartDataSource)
          .setOption('title', 'Content Status Distribution (Filtered)')
          .setOption('pieSliceText', 'percentage') 
          .setOption('legend', { position: 'right' })
          .setOption('width', sheet.getColumnWidth(1)*3.8) 
          .setOption('height', sheet.getRowHeight(13)*11)  
          .setPosition(13, 1, 5, 5) 
          .build();
        sheet.insertChart(statusChart);
    } else {
        Logger.log("Status chart data source (K:L) is empty or header only. Skipping chart.");
    }
  } catch(e) {
    Logger.log("Error creating status chart: " + e.message);
  }

  try {
    // Channel Chart
    const channelChartDataHeaderCell = sheet.getRange('M1'); // Helper data for channel chart
    const channelLastRow = channelChartDataHeaderCell.getDataRegion(SpreadsheetApp.Dimension.ROWS).getLastRow();
    const channelChartDataSource = sheet.getRange('M1:N' + channelLastRow);

    if (channelLastRow > 1) {
        let channelChart = sheet.newChart()
          .setChartType(Charts.ChartType.COLUMN)
          .addRange(channelChartDataSource)
          .setOption('title', 'Content by Channel (Filtered)')
          .setOption('legend', { position: 'none' })
          .setOption('hAxis', { title: 'Channel' })
          .setOption('vAxis', { title: 'Count', minValue: 0 })
          .setOption('width', sheet.getColumnWidth(5)*3.8) 
          .setOption('height', sheet.getRowHeight(13)*11)   
          .setPosition(13, 5, 5, 5) 
          .build();
        sheet.insertChart(channelChart);
    } else {
        Logger.log("Channel chart data source (M:N) is empty or header only. Skipping chart.");
    }
  } catch(e) {
    Logger.log("Error creating channel chart: " + e.message);
  }
  Logger.log('Dashboard charts generation attempted.');
}


/**
 * Applies final formatting touches to the Dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Dashboard sheet object.
 */
function formatDashboard(sheet) {
  Logger.log('Dashboard final formatting applied/verified.');
}


/**
 * Initializes the Analytics sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet object.
 */
function initializeAnalyticsSheet(ss) {
  const sheetName = 'Analytics';
  let analyticsSheet = ss.getSheetByName(sheetName);
  if (!analyticsSheet) {
    analyticsSheet = ss.insertSheet(sheetName);
    Logger.log(`Created sheet: ${sheetName}`);
    setupAnalyticsSheet(analyticsSheet); 
  } else {
    Logger.log(`Sheet found: ${sheetName}. Verifying structure.`);
    setupAnalyticsSheet(analyticsSheet); 
  }
}

/**
 * Sets up the Analytics sheet headers.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Analytics sheet object.
 */
function setupAnalyticsSheet(sheet) {
    sheet.setFrozenRows(0); 

    const headers = [
      'Content ID', 'Date Published', 'Channel', 'Content Title', 'Views/Impressions', 'Engagements', 'Link Clicks', 'Last Updated'
    ];
    const headerRange = sheet.getRange(1, 1, 1, headers.length);

     if (headerRange.getValues()[0].join('') !== headers.join('')) {
        headerRange.setValues([headers])
          .setBackground('#4285F4')
          .setFontColor('white')
          .setFontWeight('bold')
          .setVerticalAlignment('middle');
        sheet.setRowHeight(1, 25);
         Logger.log("Analytics sheet headers set/corrected.");
     }

    if (sheet.getColumnWidth(1) !== 100) sheet.setColumnWidth(1, 100); 
    if (sheet.getColumnWidth(2) !== 120) sheet.setColumnWidth(2, 120); 
    if (sheet.getColumnWidth(3) !== 100) sheet.setColumnWidth(3, 100); 
    if (sheet.getColumnWidth(4) !== 300) sheet.setColumnWidth(4, 300); 
    if (sheet.getColumnWidth(5) !== 120) sheet.setColumnWidth(5, 120); 
    if (sheet.getColumnWidth(6) !== 120) sheet.setColumnWidth(6, 120); 
    if (sheet.getColumnWidth(7) !== 120) sheet.setColumnWidth(7, 120); 
    if (sheet.getColumnWidth(8) !== 150) sheet.setColumnWidth(8, 150); 

    const lastRow = Math.max(sheet.getMaxRows(), 1000);
    const datePubRange = sheet.getRange(2, 2, lastRow -1);
    const lastUpdRange = sheet.getRange(2, 8, lastRow -1);

    if (datePubRange.getNumberFormat() !== 'yyyy-mm-dd') datePubRange.setNumberFormat('yyyy-mm-dd');
    if (lastUpdRange.getNumberFormat() !== 'yyyy-mm-dd hh:mm:ss') lastUpdRange.setNumberFormat('yyyy-mm-dd hh:mm:ss');

    if (sheet.getFrozenRows() !== 1) sheet.setFrozenRows(1); 
    Logger.log('Analytics sheet verified/updated.');
}


/**
 * Initializes the Archives sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet object.
 */
function initializeArchivesSheet(ss) {
  const sheetName = 'Archives';
  const calendarSheetName = 'Content Calendar';
  let archivesSheet = ss.getSheetByName(sheetName);
  let calendarSheet = ss.getSheetByName(calendarSheetName);

  if (!calendarSheet) {
      Logger.log(`Cannot initialize Archives: ${calendarSheetName} sheet not found.`);
      ss.toast(`Error: ${calendarSheetName} sheet missing. Cannot setup Archives.`, "Setup Error", 10);
      return; 
  }

  if (!archivesSheet) {
    archivesSheet = ss.insertSheet(sheetName);
    Logger.log(`Created sheet: ${sheetName}`);
    setupArchivesSheetHeaders(archivesSheet, calendarSheet); 
  } else {
    Logger.log(`Sheet found: ${sheetName}. Verifying structure.`);
    setupArchivesSheetHeaders(archivesSheet, calendarSheet); 
  }
}

/**
 * Sets up the Archives sheet headers based on the Content Calendar sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Archives sheet object.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} calendarSheet The Content Calendar sheet object.
 */
function setupArchivesSheetHeaders(sheet, calendarSheet) {
    sheet.setFrozenRows(0); 

    let headers = [];
    let calendarLastCol = 0;
    try {
        calendarLastCol = calendarSheet.getLastColumn();
        const headerSourceRange = calendarSheet.getRange(2, 1, 1, calendarLastCol);
        headers = headerSourceRange.getValues();
        if (!headers || headers.length === 0 || headers[0].length === 0) {
            throw new Error("Could not read headers from Calendar sheet.");
        }
    } catch(e) {
        Logger.log(`Error getting headers from ${calendarSheet.getName()}: ${e}. Using default archive headers.`);
         headers = [['ID', 'Date', 'Week', 'Status', 'Channel', 'Content/Idea', 'Link to Asset', 'Content Pillar', 'Content Format', 'Assigned To', 'Notes', 'Created', 'Updated', 'Status Changed', 'Event ID', 'Archived Date']];
         calendarLastCol = headers[0].length -1; 
    }

    const archiveHeaderRange = sheet.getRange(1, 1, 1, headers[0].length);
    if (archiveHeaderRange.getValues()[0].join('') !== headers[0].join('')) {
        archiveHeaderRange.setValues(headers)
          .setBackground('#9E9E9E') 
          .setFontColor('white')
          .setFontWeight('bold')
          .setVerticalAlignment('middle');
        sheet.setRowHeight(1, 25);
        Logger.log("Archives sheet headers set/corrected.");
    }

    for (let i = 1; i <= calendarLastCol; i++) {
      try {
         const calWidth = calendarSheet.getColumnWidth(i);
         if (sheet.getColumnWidth(i) !== calWidth) {
             sheet.setColumnWidth(i, calWidth);
         }
      } catch(e) {
        Logger.log(`Could not set width for column ${i} in Archives: ${e}`);
      }
    }
     if (headers[0].length > calendarLastCol && headers[0][headers[0].length - 1] === 'Archived Date') {
         if (sheet.getColumnWidth(headers[0].length) !== 150) sheet.setColumnWidth(headers[0].length, 150);
     }

    const lastRow = Math.max(sheet.getMaxRows(), 1000);
    const COLS_ARC = { DATE: 2, CREATED: 12, UPDATED: 13, STATUS_CHANGED: 14 }; 

    try {
       const dateRangeArc = sheet.getRange(2, COLS_ARC.DATE, lastRow - 1);
       if (dateRangeArc.getNumberFormat() !== 'yyyy-mm-dd') dateRangeArc.setNumberFormat('yyyy-mm-dd');

       const createdRangeArc = sheet.getRange(2, COLS_ARC.CREATED, lastRow - 1);
       if (createdRangeArc.getNumberFormat() !== 'yyyy-mm-dd hh:mm:ss') createdRangeArc.setNumberFormat('yyyy-mm-dd hh:mm:ss');

       const updatedRangeArc = sheet.getRange(2, COLS_ARC.UPDATED, lastRow - 1);
       if (updatedRangeArc.getNumberFormat() !== 'yyyy-mm-dd hh:mm:ss') updatedRangeArc.setNumberFormat('yyyy-mm-dd hh:mm:ss');

       const statusChangedRangeArc = sheet.getRange(2, COLS_ARC.STATUS_CHANGED, lastRow - 1);
        if (statusChangedRangeArc.getNumberFormat() !== 'yyyy-mm-dd hh:mm:ss') statusChangedRangeArc.setNumberFormat('yyyy-mm-dd hh:mm:ss');

         if (headers[0].length > calendarLastCol && headers[0][headers[0].length - 1] === 'Archived Date') {
             const archivedDateRange = sheet.getRange(2, headers[0].length, lastRow - 1);
              if (archivedDateRange.getNumberFormat() !== 'yyyy-mm-dd hh:mm:ss') archivedDateRange.setNumberFormat('yyyy-mm-dd hh:mm:ss');
         }

    } catch (e) {
        Logger.log("Error setting number formats in Archives: " + e);
    }

    if(sheet.getFrozenRows() !== 1) sheet.setFrozenRows(1); 
    Logger.log('Archives sheet verified/updated.');
}


/**
 * Generates week numbers using formulas for all rows with dates in the Content Calendar.
 */
function generateWeekNumbers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Content Calendar'); 
  if (!sheet) {
    Logger.log('Content Calendar sheet not found for generating week numbers.');
     ss.toast("Cannot generate week numbers: 'Content Calendar' sheet missing.", "Error", 5);
    return;
  }

  const COLS = { DATE: 2, WEEK: 3 }; 
  const startRow = 3; 
  const lastRow = sheet.getLastRow();

  if (lastRow < startRow) {
    Logger.log('No data rows found to generate week numbers.');
    return; 
  }

  const numRows = lastRow - startRow + 1;
  const dateRange = sheet.getRange(startRow, COLS.DATE, numRows, 1);
  const weekRange = sheet.getRange(startRow, COLS.WEEK, numRows, 1);
  const dates = dateRange.getValues();
  const existingFormulas = weekRange.getFormulasR1C1(); 
  const formulasToSet = [];
  let formulasChanged = false;

  for (let i = 0; i < numRows; i++) {
    const dateCellA1 = sheet.getRange(startRow + i, COLS.DATE).getA1Notation();
    const expectedFormula = `=IF(${dateCellA1}<>"",WEEKNUM(${dateCellA1},2),"")`;
    const currentFormula = existingFormulas[i][0];

    if (dates[i][0] instanceof Date && !isNaN(dates[i][0])) {
      if (currentFormula !== expectedFormula) {
        formulasToSet.push([expectedFormula]);
        formulasChanged = true;
      } else {
        formulasToSet.push([currentFormula]); 
      }
    } else {
      if (currentFormula !== "") {
        formulasToSet.push(['']);
        formulasChanged = true;
      } else {
        formulasToSet.push(['']); 
      }
    }
  }

  if (formulasChanged) {
    weekRange.setFormulas(formulasToSet);
    Logger.log(`Week number formulas applied/verified for ${numRows} rows.`);
  } else {
      Logger.log("Week number formulas already up-to-date.");
  }
}


/**
 * Updates status colors based on conditional formatting
 */
function updateStatusColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Content Calendar');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Content Calendar sheet not found.');
    return;
  }

  setupConditionalFormatting(sheet);
  SpreadsheetApp.getUi().alert('Status colors updated successfully!');
}

/**
 * Updates data validation rules in the Content Calendar sheet based on current Lists.
 */
function updateDataValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarSheet = ss.getSheetByName('Content Calendar');

  if (!calendarSheet) {
    SpreadsheetApp.getUi().alert('Content Calendar sheet not found.');
    return;
  }
  setupDataValidation(calendarSheet);
  ss.toast('Data validation rules updated.', 'Update Complete', 5);
}

/**
 * Shows help and documentation
 */
function showHelpDocumentation() {
  const html = HtmlService.createHtmlOutputFromFile('HelpDocumentation') // Assuming HelpDocumentation.html exists
    .setWidth(800)
    .setHeight(600)
    .setTitle('Content Calendar Help & Documentation');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Content Calendar Help & Documentation');
}
// Create HelpDocumentation.html with the content from the previous version.

/**
 * Generates a content report
 */
function generateContentReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contentSheet = ss.getSheetByName('Content Calendar');
  const ui = SpreadsheetApp.getUi(); 

  if (!contentSheet) {
    ui.alert('Content Calendar sheet not found.');
    return;
  }

  let reportSheet = ss.getSheetByName('Content Report');
  if (!reportSheet) {
    reportSheet = ss.insertSheet('Content Report');
    Logger.log("Created 'Content Report' sheet.");
  } else {
    reportSheet.clear(); 
    Logger.log("Cleared existing 'Content Report' sheet.");
  }

   ss.setActiveSheet(reportSheet);

  reportSheet.getRange('A1:G1').merge().setValue('CONTENT CALENDAR REPORT - ' +
                                                Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'))
    .setBackground('#4285F4').setFontColor('white').setFontWeight('bold')
    .setHorizontalAlignment('center').setFontSize(14);
  reportSheet.setRowHeight(1, 35);

  const dataRange = contentSheet.getDataRange();
  const data = dataRange.getValues();

  if (data.length <= 2) {
    reportSheet.getRange('A3').setValue('No content data found in Content Calendar sheet.');
    Logger.log("No content data found for report.");
    return;
  }

  const headers = data[1];

   const getColIndex = (headerName) => headers.indexOf(headerName);
   const idIndex = getColIndex('ID');
   const dateIndex = getColIndex('Date');
   const weekIndex = getColIndex('Week');
   const statusIndex = getColIndex('Status');
   const channelIndex = getColIndex('Channel');
   const contentIndex = getColIndex('Content/Idea');
   const assignedIndex = getColIndex('Assigned To');

   if ([idIndex, dateIndex, weekIndex, statusIndex, channelIndex, contentIndex, assignedIndex].includes(-1)) {
       Logger.log("Error: One or more required columns (ID, Date, Week, Status, Channel, Content/Idea, Assigned To) not found in Content Calendar headers.");
       ui.alert("Error generating report: Required columns missing in Content Calendar sheet.");
       return;
   }

  let totalItems = 0;
  let plannedItems = 0;
  let completedItems = 0;
  let currentWeekItems = 0;
  let nextWeekItems = 0;
  const channelCounts = {}; 

  const currentWeek = new Date().getWeekNumber ? new Date().getWeekNumber() : parseInt(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'w'));
  Logger.log("Current Week for Report: " + currentWeek);

  for (let i = 2; i < data.length; i++) { 
    const row = data[i];

    if (!row[idIndex]) {
      continue;
    }
    totalItems++;

    const status = row[statusIndex];
    if (status === 'Planned') {
      plannedItems++;
    } else if (status === 'Schedule') {
      completedItems++;
    }

     const weekNum = parseInt(row[weekIndex]); 
     if (!isNaN(weekNum)) {
        if (weekNum === currentWeek) {
          currentWeekItems++;
        } else if (weekNum === currentWeek + 1) {
          nextWeekItems++;
        }
     }

     const channel = row[channelIndex];
     if (channel) {
         channelCounts[channel] = (channelCounts[channel] || 0) + 1;
     }
  }
   Logger.log(`Report Stats: Total=${totalItems}, Planned=${plannedItems}, Completed=${completedItems}, CurrentWk=${currentWeekItems}, NextWk=${nextWeekItems}, Channels=${JSON.stringify(channelCounts)}`);

  reportSheet.getRange('A3:G3').merge().setValue('SUMMARY STATISTICS')
    .setBackground('#EEEEEE').setFontWeight('bold').setHorizontalAlignment('center');

  reportSheet.getRange('A4:B7').setValues([
      ['Total Content Items:', totalItems],
      ['Planned Items:', plannedItems],
      ['Completed Items:', completedItems],
      ['Completion Rate:', totalItems > 0 ? `${Math.round((completedItems / totalItems) * 100)}%` : '0%']
  ]);
  reportSheet.getRange('A4:A7').setFontWeight('bold');

  reportSheet.getRange('D4:E5').setValues([
    [`Current Week (Week ${currentWeek}):`, currentWeekItems],
    [`Next Week (Week ${currentWeek + 1}):`, nextWeekItems]
  ]);
   reportSheet.getRange('D4:D5').setFontWeight('bold');

  reportSheet.getRange('A8:G8').merge().setValue('CHANNEL BREAKDOWN')
    .setBackground('#EEEEEE').setFontWeight('bold').setHorizontalAlignment('center');

  let channelRow = 9;
  for (const channel in channelCounts) {
      reportSheet.getRange(channelRow, 1).setValue(`${channel}:`).setFontWeight('bold');
      reportSheet.getRange(channelRow, 2).setValue(channelCounts[channel]);
      channelRow++;
  }

  const reportHeaders = ['ID', 'Date', 'Status', 'Channel', 'Content', 'Assigned To', 'Week'];
  let currentRow = channelRow + 1; 

  reportSheet.getRange(currentRow, 1, 1, reportHeaders.length).merge().setValue(`CURRENT WEEK CONTENT (WEEK ${currentWeek})`)
    .setBackground('#EEEEEE').setFontWeight('bold').setHorizontalAlignment('center');
  currentRow++;
  reportSheet.getRange(currentRow, 1, 1, reportHeaders.length).setValues([reportHeaders])
    .setBackground('#D9D9D9').setFontWeight('bold');
  currentRow++;

  let currentWeekData = [];
  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    const weekNum = parseInt(row[weekIndex]);
    if (row[idIndex] && !isNaN(weekNum) && weekNum === currentWeek) {
      currentWeekData.push([
        row[idIndex], row[dateIndex], row[statusIndex], row[channelIndex], row[contentIndex], row[assignedIndex], row[weekIndex]
      ]);
    }
  }

  if (currentWeekData.length > 0) {
    reportSheet.getRange(currentRow, 1, currentWeekData.length, reportHeaders.length).setValues(currentWeekData);
    reportSheet.getRange(currentRow, 2, currentWeekData.length, 1).setNumberFormat('yyyy-mm-dd');
    currentRow += currentWeekData.length;
  } else {
    reportSheet.getRange(currentRow, 1).setValue('No content scheduled for current week.');
    currentRow++;
  }

   currentRow++; 
   reportSheet.getRange(currentRow, 1, 1, reportHeaders.length).merge().setValue(`NEXT WEEK CONTENT (WEEK ${currentWeek + 1})`)
     .setBackground('#EEEEEE').setFontWeight('bold').setHorizontalAlignment('center');
   currentRow++;
   reportSheet.getRange(currentRow, 1, 1, reportHeaders.length).setValues([reportHeaders])
     .setBackground('#D9D9D9').setFontWeight('bold');
   currentRow++;

  let nextWeekData = [];
  for (let i = 2; i < data.length; i++) {
    const row = data[i];
     const weekNum = parseInt(row[weekIndex]);
    if (row[idIndex] && !isNaN(weekNum) && weekNum === currentWeek + 1) {
      nextWeekData.push([
        row[idIndex], row[dateIndex], row[statusIndex], row[channelIndex], row[contentIndex], row[assignedIndex], row[weekIndex]
      ]);
    }
  }

  if (nextWeekData.length > 0) {
    reportSheet.getRange(currentRow, 1, nextWeekData.length, reportHeaders.length).setValues(nextWeekData);
     reportSheet.getRange(currentRow, 2, nextWeekData.length, 1).setNumberFormat('yyyy-mm-dd');
    currentRow += nextWeekData.length;
  } else {
    reportSheet.getRange(currentRow, 1).setValue('No content scheduled for next week.');
    currentRow++;
  }

  const colWidths = [150, 100, 120, 120, 300, 150, 80]; 
  for(let i = 0; i < reportHeaders.length; i++) {
      if (reportSheet.getColumnWidth(i+1) !== colWidths[i]) {
          reportSheet.setColumnWidth(i+1, colWidths[i]);
      }
  }

  ui.alert('Content report generated successfully!');
   Logger.log('Content report generated.');
}


/**
 * Navigates to the dashboard
 */
function navigateToDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('Dashboard');
  
  if (!dashboard) {
    SpreadsheetApp.getUi().alert('Dashboard sheet not found. Please run the initialization function first.');
    return;
  }
  
  dashboard.activate();
}

if (typeof Date.prototype.getWeekNumber !== 'function') {
    Date.prototype.getWeekNumber = function() {
      const target = new Date(this.valueOf());
      const dayNum = (this.getDay() + 6) % 7;
      target.setDate(target.getDate() - dayNum + 3);
      const firstThursday = target.valueOf();
      target.setMonth(0, 1); 
      const firstDayOfYear = target.getDay(); 
      target.setDate(1 + ((4 - firstDayOfYear + 7) % 7)); 
       if (firstThursday < target.valueOf()) {
           return parseInt(Utilities.formatDate(new Date(this.valueOf()), Session.getScriptTimeZone(), 'w'));
       } else {
          return 1 + Math.ceil((firstThursday - target.valueOf()) / 604800000); 
       }
    };
}

function addNewContentItem() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Content Calendar');
    if (!sheet) {
        SpreadsheetApp.getUi().alert("Content Calendar sheet not found.");
        return;
    }

    const startRow = 3; 
    let newRow = sheet.getLastRow() + 1;

     const ids = sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, 1).getValues();
     let firstEmptyRowIndex = ids.findIndex(id => !id[0]); 
     newRow = (firstEmptyRowIndex === -1) ? sheet.getLastRow() + 1 : startRow + firstEmptyRowIndex;

    const now = new Date();
    const tomorrow = new Date(now.getTime() + 24 * 60 * 60 * 1000);

    const COLS = { ID: 1, DATE: 2, WEEK: 3, STATUS: 4, CONTENT: 6, CREATED: 12, UPDATED: 13 };

    sheet.getRange(newRow, COLS.DATE).setValue(tomorrow).setNumberFormat('yyyy-mm-dd'); 
    sheet.getRange(newRow, COLS.STATUS).setValue('Planned');
    sheet.getRange(newRow, COLS.CREATED).setValue(now).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange(newRow, COLS.UPDATED).setValue(now).setNumberFormat('yyyy-mm-dd hh:mm:ss');

    const idCell = sheet.getRange(newRow, COLS.ID);
    const weekCell = sheet.getRange(newRow, COLS.WEEK);
    const dateCellA1 = sheet.getRange(newRow, COLS.DATE).getA1Notation();
    const idFormula = `=IF(${dateCellA1}<>"", "CONT-" & TEXT(ROW(A${newRow})-2,"000"), "")`;
    const weekFormula = `=IF(${dateCellA1}<>"", WEEKNUM(${dateCellA1},2), "")`;

    if (!idCell.getFormula()) idCell.setFormula(idFormula);
    if (!weekCell.getFormula()) weekCell.setFormula(weekFormula);

    sheet.getRange(newRow, COLS.CONTENT).activate(); 
    Logger.log(`Added new content item placeholder row at ${newRow}`);
}

function batchAddContent() { SpreadsheetApp.getUi().alert("Function 'batchAddContent' not implemented yet."); }
function importContentFromCsv() { SpreadsheetApp.getUi().alert("Function 'importContentFromCsv' not implemented yet."); }

/**
 * Called from the InitialAuthPrompt.html modal to begin the authorization
 * and full setup process.
 * @return {{success: boolean, message?: string, error?: string}} Result object for the modal.
 */
function proceedWithAuthorizationAndSetupFromModal() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let result = { success: false, error: "Unknown error during setup." };

  Logger.log("Proceeding with Authorization and Setup from Modal...");
  ss.toast('Starting authorization & setup...', 'Setup Progress', -1); //Indefinite toast

  try {
    // Step 1: Trigger general API authorizations
    const authSuccess = promptAndTriggerApiAuthorizations(ui, ss); // This already sets AUTH_AND_CONFIG_CHECKS_COMPLETED on success

    if (authSuccess) {
      ss.toast('Core authorizations successful. Proceeding with full initialization...', 'Setup Progress', 7);
      Logger.log("Core authorizations successful. Proceeding with full initialization.");

      // Step 2: Run the full content calendar initialization
      // Pass skipAuth=true because promptAndTriggerApiAuthorizations already handled it.
      initializeContentCalendar({ skipAuth: true }); // initializeContentCalendar shows its own toasts/alerts for completion or errors

      // If initializeContentCalendar completes without throwing major errors, assume success for the modal.
      // initializeContentCalendar logs its own errors.
      result = { success: true, message: "Setup and initialization process completed. You can now close this dialog if it hasn't closed automatically." };
      ss.toast('Full setup process completed!', 'Setup Complete', 10);
      Logger.log("Full setup process completed via modal.");

    } else {
      // promptAndTriggerApiAuthorizations itself handles showing the AuthRetryModal if there were issues.
      // If it returns false, it means user cancelled or there were immediate, unrecoverable auth issues
      // not handled by AuthRetryModal, or AuthRetryModal was shown and this is the fallout.
      Logger.log("Core authorization step did not fully succeed or was cancelled.");
      result = { success: false, message: "Authorization was not fully completed or was cancelled. Some features may not work. You can try again via the 'Content Calendar > Initial Setup / Re-authorize' menu." };
      ss.toast('Authorization incomplete. Full setup deferred.', 'Setup Warning', 10);
    }
  } catch (e) {
    Logger.log(`Error in proceedWithAuthorizationAndSetupFromModal: ${e.toString()}\n${e.stack}`);
    result = { success: false, error: `Setup failed: ${e.message}. Check script logs for details.` };
    ss.toast(`Critical error during setup: ${e.message}`, 'Setup Error', 10);
  }
  
  return result; // Return status to the modal
}

/**
 * Wrapper function for the menu item "Initial Setup / Re-authorize".
 * This allows users to manually trigger the full setup and authorization process.
 */
function userInitiatedFullSetup() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("User initiated full setup via menu.");
  
  // Call the same function that the modal would, which handles both auth and initialization.
  // The return value here is mostly for logging or if we wanted to give a final UI alert from this function.
  const setupResult = proceedWithAuthorizationAndSetupFromModal();

  if (setupResult.success) {
    ui.alert("Setup Complete", "The Content Calendar has been set up and authorized successfully.", ui.ButtonSet.OK);
  } else {
    // Errors/warnings are typically handled by proceedWithAuthorizationAndSetupFromModal or its sub-functions with toasts/modals.
    // This alert is a fallback if the modal flow didn't provide sufficient feedback or was closed prematurely.
    let finalMessage = "The setup process encountered issues or was not fully completed. ";
    if (setupResult.message) finalMessage += setupResult.message;
    else if (setupResult.error) finalMessage += setupResult.error;
    finalMessage += " Please check any messages or logs, and try again if necessary.";
    ui.alert("Setup Incomplete", finalMessage, ui.ButtonSet.OK);
  }
}