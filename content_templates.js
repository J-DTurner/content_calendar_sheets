/**
 * Content Template System for Social Media Content Calendar
 * 
 * This script provides functionality for creating, managing, and using
 * content templates for different social media channels and formats.
 */

// Templates configuration
const TEMPLATES_CONFIG = {
  TEMPLATES_SHEET: 'Content Templates',
  CONTENT_COLUMN: 6,        // Column F in Content Calendar
  CHANNEL_COLUMN: 5,        // Column E in Content Calendar
  FORMAT_COLUMN: 8,         // Column H in Content Calendar
  DEFAULT_TEMPLATES: {
    // Templates for Twitter
    'Twitter': {
      'Text Post': 'Main point: \nKey message (under 280 chars): \nHashtags: \nMention: ',
      'Image': 'Main point: \nImage description: \nCaption (under 280 chars): \nHashtags: ',
      'Video': 'Video topic: \nKey message: \nCaption: \nHashtags: \nVideo length: ',
      'Thread': 'Main topic: \n\nTweet 1: \n\nTweet 2: \n\nTweet 3: \n\nCall to action: ',
      'Poll': 'Poll question: \n\nOption 1: \nOption 2: \nOption 3: \nOption 4: \n\nDuration: '
    },
    // Templates for YouTube
    'YouTube': {
      'Video': 'Video title: \nDescription: \n\nIntro (0:00-0:30): \nMain points: \n- Point 1 (0:30-2:00): \n- Point 2 (2:00-4:00): \n- Point 3 (4:00-6:00): \nConclusion (6:00-7:00): \n\nTags: \nCategory: ',
      'Short': 'Short title: \nHook (first 3 seconds): \nMain concept: \nCall to action: \nCaption: \nHashtags: ',
      'Live': 'Stream title: \nScheduled date/time: \nDescription: \nTopics to cover: \n- Topic 1: \n- Topic 2: \n- Topic 3: \nQ&A prompts: '
    },
    // Templates for Telegram
    'Telegram': {
      'Text Post': 'Title: \n\nMain content: \n\nKey points: \n- \n- \n- \n\nCall to action: ',
      'Image': 'Title: \n\nImage description: \n\nCaption: \n\nAdditional notes: ',
      'Video': 'Title: \n\nVideo description: \n\nCaption: \n\nKey timestamps: \n- 0:00 - \n- 0:00 - ',
      'Poll': 'Poll question: \n\nOptions: \n- \n- \n- \n- \n\nContext/introduction: '
    },
    // Default templates for any channel
    'Default': {
      'Text Post': 'Title: \nMain message: \nKey points: \n- \n- \n- \nCall to action: ',
      'Image': 'Caption: \nImage description: \nKey message: \nHashtags: ',
      'Video': 'Title: \nMain topic: \nKey points: \n- \n- \n- \nCall to action: ',
      'Poll': 'Question: \nOptions: \n- \n- \n- \n- \nContext: ',
      'Default': 'Content title: \nMain message: \nKey points: \n- \n- \n- \nCall to action: '
    }
  }
};

/**
 * Applies a content template to the selected cell
 */
function applyContentTemplate() {
  // Get the active sheet and selected cell
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = SpreadsheetApp.getActiveRange();
  
  // Only process in the Content Calendar sheet
  if (sheet.getName() !== 'Content Calendar') {
    SpreadsheetApp.getUi().alert('Please select a cell in the Content Calendar sheet.');
    return;
  }
  
  // Check if selected cell is in the Content column
  if (range.getColumn() !== TEMPLATES_CONFIG.CONTENT_COLUMN) {
    SpreadsheetApp.getUi().alert('Please select a cell in the Content/Idea column.');
    return;
  }
  
  // Skip header rows
  if (range.getRow() < 3) {
    SpreadsheetApp.getUi().alert('Please select a content row (not headers).');
    return;
  }
  
  // Get channel and format for the selected row
  const channel = sheet.getRange(range.getRow(), TEMPLATES_CONFIG.CHANNEL_COLUMN).getValue();
  const format = sheet.getRange(range.getRow(), TEMPLATES_CONFIG.FORMAT_COLUMN).getValue();
  
  // Get current content
  const currentContent = range.getValue();
  
  // If content already exists, ask for confirmation before overwriting
  if (currentContent && currentContent.trim() !== '') {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Content Already Exists',
      'This cell already has content. Do you want to replace it with a template?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
  }
  
  // Get the template
  let template;
  
  if (channel && format) {
    // Try to get a template for this specific channel and format
    template = getTemplateForChannelAndFormat(channel, format);
  } else {
    // Show template selection dialog
    template = showTemplateSelectionDialog();
  }
  
  // Apply the template if one was found
  if (template) {
    range.setValue(template);
  }
}

/**
 * Gets a template for a specific channel and format
 * @param {string} channel The channel (Twitter, YouTube, Telegram)
 * @param {string} format The content format
 * @return {string} The template text or null if not found
 */
function getTemplateForChannelAndFormat(channel, format) {
  // Try to get from Templates sheet first
  const template = getTemplateFromSheet(channel, format);
  if (template) {
    return template;
  }
  
  // Fall back to default templates
  return getDefaultTemplate(channel, format);
}

/**
 * Gets a template from the Templates sheet
 * @param {string} channel The channel
 * @param {string} format The content format
 * @return {string} The template text or null if not found
 */
function getTemplateFromSheet(channel, format) {
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if Templates sheet exists
  const templatesSheet = ss.getSheetByName(TEMPLATES_CONFIG.TEMPLATES_SHEET);
  if (!templatesSheet) {
    return null;
  }
  
  // Get all template data
  const dataRange = templatesSheet.getDataRange();
  const values = dataRange.getValues();
  
  // Find header row (first row)
  const headers = values[0];
  
  // Find column indexes
  const channelIndex = headers.indexOf('Channel');
  const formatIndex = headers.indexOf('Format');
  const templateIndex = headers.indexOf('Template');
  
  // Skip if any required column is missing
  if (channelIndex === -1 || formatIndex === -1 || templateIndex === -1) {
    return null;
  }
  
  // Search for matching template
  for (let i = 1; i < values.length; i++) {
    const templateChannel = values[i][channelIndex];
    const templateFormat = values[i][formatIndex];
    
    if (templateChannel === channel && templateFormat === format) {
      return values[i][templateIndex];
    }
  }
  
  return null;
}

/**
 * Gets a default template from the built-in templates
 * @param {string} channel The channel
 * @param {string} format The content format
 * @return {string} The template text or a generic template if not found
 */
function getDefaultTemplate(channel, format) {
  // Check if we have templates for this channel
  if (TEMPLATES_CONFIG.DEFAULT_TEMPLATES[channel]) {
    // Check if we have a template for this format
    if (TEMPLATES_CONFIG.DEFAULT_TEMPLATES[channel][format]) {
      return TEMPLATES_CONFIG.DEFAULT_TEMPLATES[channel][format];
    }
    
    // Try the default format for this channel
    if (TEMPLATES_CONFIG.DEFAULT_TEMPLATES[channel]['Default']) {
      return TEMPLATES_CONFIG.DEFAULT_TEMPLATES[channel]['Default'];
    }
  }
  
  // Try the default channel, specific format
  if (TEMPLATES_CONFIG.DEFAULT_TEMPLATES['Default'] && 
      TEMPLATES_CONFIG.DEFAULT_TEMPLATES['Default'][format]) {
    return TEMPLATES_CONFIG.DEFAULT_TEMPLATES['Default'][format];
  }
  
  // Fall back to most generic template
  return TEMPLATES_CONFIG.DEFAULT_TEMPLATES['Default']['Default'];
}

/**
 * Shows a dialog for selecting a template
 * @return {string} The selected template or null if canceled
 */
function showTemplateSelectionDialog() {
  const ui = SpreadsheetApp.getUi();
  
  // First, show channel selection
  const channelResponse = ui.prompt(
    'Select Channel',
    'Enter channel number:\n\n' +
    '1. Twitter\n' +
    '2. YouTube\n' +
    '3. Telegram\n',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (channelResponse.getSelectedButton() !== ui.Button.OK) {
    return null;
  }
  
  // Parse channel selection
  let channel;
  switch (channelResponse.getResponseText().trim()) {
    case '1':
      channel = 'Twitter';
      break;
    case '2':
      channel = 'YouTube';
      break;
    case '3':
      channel = 'Telegram';
      break;
    default:
      ui.alert('Invalid channel selection.');
      return null;
  }
  
  // Format selection depends on channel
  let formatPrompt = 'Enter format number:\n\n';
  
  if (channel === 'Twitter') {
    formatPrompt += '1. Text Post\n2. Image\n3. Video\n4. Thread\n5. Poll';
  } else if (channel === 'YouTube') {
    formatPrompt += '1. Video\n2. Short\n3. Live';
  } else { // Telegram
    formatPrompt += '1. Text Post\n2. Image\n3. Video\n4. Poll';
  }
  
  const formatResponse = ui.prompt(
    'Select Format',
    formatPrompt,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (formatResponse.getSelectedButton() !== ui.Button.OK) {
    return null;
  }
  
  // Parse format selection
  let format;
  const formatNum = formatResponse.getResponseText().trim();
  
  if (channel === 'Twitter') {
    switch (formatNum) {
      case '1': format = 'Text Post'; break;
      case '2': format = 'Image'; break;
      case '3': format = 'Video'; break;
      case '4': format = 'Thread'; break;
      case '5': format = 'Poll'; break;
      default: ui.alert('Invalid format selection.'); return null;
    }
  } else if (channel === 'YouTube') {
    switch (formatNum) {
      case '1': format = 'Video'; break;
      case '2': format = 'Short'; break;
      case '3': format = 'Live'; break;
      default: ui.alert('Invalid format selection.'); return null;
    }
  } else { // Telegram
    switch (formatNum) {
      case '1': format = 'Text Post'; break;
      case '2': format = 'Image'; break;
      case '3': format = 'Video'; break;
      case '4': format = 'Poll'; break;
      default: ui.alert('Invalid format selection.'); return null;
    }
  }
  
  // Get the template for the selected channel and format
  return getTemplateForChannelAndFormat(channel, format);
}

/**
 * Creates or updates a custom template
 */
function manageCustomTemplate() {
  const ui = SpreadsheetApp.getUi();
  
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Ensure Templates sheet exists
  let templatesSheet = ss.getSheetByName(TEMPLATES_CONFIG.TEMPLATES_SHEET);
  if (!templatesSheet) {
    // Create the sheet
    templatesSheet = ss.insertSheet(TEMPLATES_CONFIG.TEMPLATES_SHEET);
    
    // Set up headers
    templatesSheet.getRange(1, 1, 1, 3).setValues([
      ['Channel', 'Format', 'Template']
    ]);
    
    // Format headers
    templatesSheet.getRange(1, 1, 1, 3)
      .setBackground('#4285F4')
      .setFontColor('white')
      .setFontWeight('bold');
    
    // Set column widths
    templatesSheet.setColumnWidth(1, 150); // Channel
    templatesSheet.setColumnWidth(2, 150); // Format
    templatesSheet.setColumnWidth(3, 500); // Template
  }
  
  // Show channel selection
  const channelResponse = ui.prompt(
    'Template Channel',
    'Enter channel name (Twitter, YouTube, Telegram, or other):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (channelResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const channel = channelResponse.getResponseText().trim();
  if (!channel) {
    ui.alert('Channel name cannot be empty.');
    return;
  }
  
  // Show format selection
  const formatResponse = ui.prompt(
    'Template Format',
    'Enter format name (Text Post, Image, Video, etc.):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (formatResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const format = formatResponse.getResponseText().trim();
  if (!format) {
    ui.alert('Format name cannot be empty.');
    return;
  }
  
  // Check if template already exists
  const existing = getTemplateFromSheet(channel, format);
  
  // Set up template text prompt
  let templatePrompt = 'Enter the template text:';
  if (existing) {
    templatePrompt = 'Edit the existing template:\n\n' + existing;
  }
  
  // Show template text input
  const templateResponse = ui.prompt(
    'Template Content',
    templatePrompt,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (templateResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const template = templateResponse.getResponseText();
  if (!template) {
    ui.alert('Template content cannot be empty.');
    return;
  }
  
  // Save the template
  saveTemplate(templatesSheet, channel, format, template);
  
  ui.alert('Template saved successfully!');
}

/**
 * Saves a template to the Templates sheet
 * @param {Sheet} sheet The templates sheet
 * @param {string} channel The channel
 * @param {string} format The content format
 * @param {string} template The template text
 */
function saveTemplate(sheet, channel, format, template) {
  // Get all template data
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // Find header row (first row)
  const headers = values[0];
  
  // Find column indexes
  const channelIndex = headers.indexOf('Channel');
  const formatIndex = headers.indexOf('Format');
  
  // Check if this template already exists
  let existingRow = -1;
  
  for (let i = 1; i < values.length; i++) {
    if (values[i][channelIndex] === channel && values[i][formatIndex] === format) {
      existingRow = i + 1; // +1 because array is 0-based but sheet is 1-based
      break;
    }
  }
  
  if (existingRow > 0) {
    // Update existing template
    sheet.getRange(existingRow, 3).setValue(template);
  } else {
    // Add new template
    sheet.appendRow([channel, format, template]);
  }
}

/**
 * Deletes a custom template
 */
function deleteCustomTemplate() {
  const ui = SpreadsheetApp.getUi();
  
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if Templates sheet exists
  const templatesSheet = ss.getSheetByName(TEMPLATES_CONFIG.TEMPLATES_SHEET);
  if (!templatesSheet) {
    ui.alert('No custom templates found.');
    return;
  }
  
  // Get all template data
  const dataRange = templatesSheet.getDataRange();
  const values = dataRange.getValues();
  
  // Skip if only headers exist
  if (values.length <= 1) {
    ui.alert('No custom templates found.');
    return;
  }
  
  // Build list of templates
  let templateList = '';
  for (let i = 1; i < values.length; i++) {
    templateList += `${i}. ${values[i][0]} - ${values[i][1]}\n`;
  }
  
  // Show template selection
  const response = ui.prompt(
    'Delete Template',
    'Enter the number of the template to delete:\n\n' + templateList,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  // Parse template selection
  const templateIndex = parseInt(response.getResponseText().trim());
  if (isNaN(templateIndex) || templateIndex < 1 || templateIndex >= values.length) {
    ui.alert('Invalid template selection.');
    return;
  }
  
  // Confirm deletion
  const confirmResponse = ui.alert(
    'Confirm Deletion',
    `Are you sure you want to delete the template for ${values[templateIndex][0]} - ${values[templateIndex][1]}?`,
    ui.ButtonSet.YES_NO
  );
  
  if (confirmResponse !== ui.Button.YES) {
    return;
  }
  
  // Delete the template
  templatesSheet.deleteRow(templateIndex + 1); // +1 because array is 0-based but sheet is 1-based
  
  ui.alert('Template deleted successfully!');
}

/**
 * Shows all available templates
 */
function showAllTemplates() {
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get custom templates
  let customTemplates = [];
  const templatesSheet = ss.getSheetByName(TEMPLATES_CONFIG.TEMPLATES_SHEET);
  
  if (templatesSheet) {
    const dataRange = templatesSheet.getDataRange();
    const values = dataRange.getValues();
    
    // Skip header row
    for (let i = 1; i < values.length; i++) {
      customTemplates.push({
        channel: values[i][0],
        format: values[i][1],
        template: values[i][2],
        isCustom: true
      });
    }
  }
  
  // Get default templates
  let defaultTemplates = [];
  
  for (const channel in TEMPLATES_CONFIG.DEFAULT_TEMPLATES) {
    if (channel === 'Default') continue; // Skip Default
    
    for (const format in TEMPLATES_CONFIG.DEFAULT_TEMPLATES[channel]) {
      if (format === 'Default') continue; // Skip Default
      
      defaultTemplates.push({
        channel: channel,
        format: format,
        template: TEMPLATES_CONFIG.DEFAULT_TEMPLATES[channel][format],
        isCustom: false
      });
    }
  }
  
  // Combine and sort templates
  const allTemplates = [...customTemplates, ...defaultTemplates].sort((a, b) => {
    if (a.channel !== b.channel) {
      return a.channel.localeCompare(b.channel);
    }
    return a.format.localeCompare(b.format);
  });
  
  // Build HTML to display templates
  let html = '<h2>Available Content Templates</h2>';
  
  if (customTemplates.length > 0) {
    html += '<h3>Custom Templates</h3>';
    html += '<table style="width:100%; border-collapse: collapse;">';
    html += '<tr style="background-color: #f3f3f3;">';
    html += '<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Channel</th>';
    html += '<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Format</th>';
    html += '<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Template</th>';
    html += '</tr>';
    
    for (const template of customTemplates) {
      html += '<tr>';
      html += `<td style="border: 1px solid #ddd; padding: 8px;">${template.channel}</td>`;
      html += `<td style="border: 1px solid #ddd; padding: 8px;">${template.format}</td>`;
      html += `<td style="border: 1px solid #ddd; padding: 8px; white-space: pre-wrap;">${template.template.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</td>`;
      html += '</tr>';
    }
    
    html += '</table>';
  }
  
  html += '<h3>Default Templates</h3>';
  html += '<table style="width:100%; border-collapse: collapse;">';
  html += '<tr style="background-color: #f3f3f3;">';
  html += '<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Channel</th>';
  html += '<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Format</th>';
  html += '<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Template</th>';
  html += '</tr>';
  
  for (const template of defaultTemplates) {
    html += '<tr>';
    html += `<td style="border: 1px solid #ddd; padding: 8px;">${template.channel}</td>`;
    html += `<td style="border: 1px solid #ddd; padding: 8px;">${template.format}</td>`;
    html += `<td style="border: 1px solid #ddd; padding: 8px; white-space: pre-wrap;">${template.template.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</td>`;
    html += '</tr>';
  }
  
  html += '</table>';
  
  // Show the HTML
  const htmlOutput = HtmlService
    .createHtmlOutput(html)
    .setWidth(800)
    .setHeight(600)
    .setTitle('Content Templates');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Content Templates');
}

/**
 * Generates a content template based on the provided parameters
 * @param {string} channel The channel
 * @param {string} format The content format
 * @param {object} params Additional parameters for template customization
 * @return {string} The generated template
 */
function generateTemplateWithParams(channel, format, params = {}) {
  // Get the base template
  let template = getTemplateForChannelAndFormat(channel, format);
  
  // If no template found, use a generic one
  if (!template) {
    template = TEMPLATES_CONFIG.DEFAULT_TEMPLATES['Default']['Default'];
  }
  
  // Replace placeholders with parameter values
  for (const key in params) {
    const placeholder = `{${key}}`;
    template = template.replace(new RegExp(placeholder, 'g'), params[key]);
  }
  
  return template;
}

/**
 * Applies templates to multiple content items at once based on channel and format
 */
function batchApplyTemplates() {
  const ui = SpreadsheetApp.getUi();
  
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarSheet = ss.getSheetByName('Content Calendar');
  
  if (!calendarSheet) {
    ui.alert('Content Calendar sheet not found.');
    return;
  }
  
  // Prompt for filter options
  const response = ui.prompt(
    'Batch Apply Templates',
    'Do you want to apply templates to:\n\n' +
    '1. All empty content cells\n' +
    '2. All content cells (including those with existing content)\n' +
    '3. Only selected rows\n' +
    '4. Only specific channel and format\n',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const option = response.getResponseText().trim();
  
  // Get all data
  const dataRange = calendarSheet.getDataRange();
  const values = dataRange.getValues();
  
  // Skip if less than 3 rows (header rows only)
  if (values.length < 3) {
    ui.alert('No content items found.');
    return;
  }
  
  // Process rows based on selected option
  let processedCount = 0;
  let skippedCount = 0;
  
  switch (option) {
    case '1': // All empty content cells
      for (let i = 2; i < values.length; i++) { // Start from row 3 (index 2)
        const row = values[i];
        
        // Skip if no channel or format or if content already exists
        if (!row[TEMPLATES_CONFIG.CHANNEL_COLUMN - 1] || 
            !row[TEMPLATES_CONFIG.FORMAT_COLUMN - 1] || 
            (row[TEMPLATES_CONFIG.CONTENT_COLUMN - 1] && row[TEMPLATES_CONFIG.CONTENT_COLUMN - 1].toString().trim() !== '')) {
          skippedCount++;
          continue;
        }
        
        // Get template and apply it
        const template = getTemplateForChannelAndFormat(
          row[TEMPLATES_CONFIG.CHANNEL_COLUMN - 1],
          row[TEMPLATES_CONFIG.FORMAT_COLUMN - 1]
        );
        
        if (template) {
          calendarSheet.getRange(i + 1, TEMPLATES_CONFIG.CONTENT_COLUMN).setValue(template);
          processedCount++;
        } else {
          skippedCount++;
        }
      }
      break;
      
    case '2': // All content cells
      for (let i = 2; i < values.length; i++) { // Start from row 3 (index 2)
        const row = values[i];
        
        // Skip if no channel or format
        if (!row[TEMPLATES_CONFIG.CHANNEL_COLUMN - 1] || 
            !row[TEMPLATES_CONFIG.FORMAT_COLUMN - 1]) {
          skippedCount++;
          continue;
        }
        
        // Get template and apply it
        const template = getTemplateForChannelAndFormat(
          row[TEMPLATES_CONFIG.CHANNEL_COLUMN - 1],
          row[TEMPLATES_CONFIG.FORMAT_COLUMN - 1]
        );
        
        if (template) {
          calendarSheet.getRange(i + 1, TEMPLATES_CONFIG.CONTENT_COLUMN).setValue(template);
          processedCount++;
        } else {
          skippedCount++;
        }
      }
      break;
      
    case '3': // Only selected rows
      // Get selected rows
      const selectedRanges = calendarSheet.getActiveRangeList();
      
      if (!selectedRanges) {
        ui.alert('No rows selected.');
        return;
      }
      
      // Process each selected range
      const ranges = selectedRanges.getRanges();
      
      for (const range of ranges) {
        const startRow = range.getRow();
        const numRows = range.getNumRows();
        
        // Skip header rows
        if (startRow < 3) {
          continue;
        }
        
        // Process each row in the range
        for (let i = 0; i < numRows; i++) {
          const rowIndex = startRow + i;
          
          // Get channel and format
          const channel = calendarSheet.getRange(rowIndex, TEMPLATES_CONFIG.CHANNEL_COLUMN).getValue();
          const format = calendarSheet.getRange(rowIndex, TEMPLATES_CONFIG.FORMAT_COLUMN).getValue();
          
          // Skip if no channel or format
          if (!channel || !format) {
            skippedCount++;
            continue;
          }
          
          // Get template and apply it
          const template = getTemplateForChannelAndFormat(channel, format);
          
          if (template) {
            calendarSheet.getRange(rowIndex, TEMPLATES_CONFIG.CONTENT_COLUMN).setValue(template);
            processedCount++;
          } else {
            skippedCount++;
          }
        }
      }
      break;
      
    case '4': // Only specific channel and format
      // Prompt for channel
      const channelResponse = ui.prompt(
        'Select Channel',
        'Enter channel name:',
        ui.ButtonSet.OK_CANCEL
      );
      
      if (channelResponse.getSelectedButton() !== ui.Button.OK) {
        return;
      }
      
      const targetChannel = channelResponse.getResponseText().trim();
      
      // Prompt for format
      const formatResponse = ui.prompt(
        'Select Format',
        'Enter format name:',
        ui.ButtonSet.OK_CANCEL
      );
      
      if (formatResponse.getSelectedButton() !== ui.Button.OK) {
        return;
      }
      
      const targetFormat = formatResponse.getResponseText().trim();
      
      // Get template for the specified channel and format
      const template = getTemplateForChannelAndFormat(targetChannel, targetFormat);
      
      if (!template) {
        ui.alert('No template found for the specified channel and format.');
        return;
      }
      
      // Apply to matching rows
      for (let i = 2; i < values.length; i++) { // Start from row 3 (index 2)
        const row = values[i];
        
        if (row[TEMPLATES_CONFIG.CHANNEL_COLUMN - 1] === targetChannel && 
            row[TEMPLATES_CONFIG.FORMAT_COLUMN - 1] === targetFormat) {
          calendarSheet.getRange(i + 1, TEMPLATES_CONFIG.CONTENT_COLUMN).setValue(template);
          processedCount++;
        } else {
          skippedCount++;
        }
      }
      break;
      
    default:
      ui.alert('Invalid option selected.');
      return;
  }
  
  // Show summary
  ui.alert(
    `Template application completed:\n\n` +
    `- Applied: ${processedCount} items\n` +
    `- Skipped: ${skippedCount} items`
  );
}

/**
 * Creates a templates menu
 */
function createTemplatesMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Templates')
    .addItem('Apply Template to Selected Cell', 'applyContentTemplate')
    .addItem('Batch Apply Templates', 'batchApplyTemplates')
    .addSeparator()
    .addItem('Create/Edit Template', 'manageCustomTemplate')
    .addItem('Delete Template', 'deleteCustomTemplate')
    .addItem('View All Templates', 'showAllTemplates')
    .addToUi();
}