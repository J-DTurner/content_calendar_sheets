/**
 * Search and Filter Functionality for Social Media Content Calendar
 * 
 * This script provides advanced search and filtering capabilities
 * for the content calendar, enabling users to quickly find and filter
 * content based on various criteria.
 */

// Search and filter configuration
const SEARCH_CONFIG = {
  SEARCH_SHEET: 'Search',
  CONTENT_SHEET: 'Content Calendar',
  SEARCH_RESULTS_RANGE: 'A5:G50',
  FILTER_CRITERIA_CELL: 'B2',
  SEARCH_TEXT_CELL: 'B3',
  TOTAL_RESULTS_CELL: 'G2',
  MAX_RESULTS: 45,
  SEARCH_FIELDS: [
    { name: 'ID', column: 1 },
    { name: 'Date', column: 2 },
    { name: 'Week', column: 3 },
    { name: 'Status', column: 4 },
    { name: 'Channel', column: 5 },
    { name: 'Content', column: 6 },
    { name: 'Content Pillar', column: 7 },
    { name: 'Content Format', column: 8 },
    { name: 'Assigned To', column: 11 },
    { name: 'Notes', column: 13 }
  ],
  DATE_COLUMNS: [2, 9, 10] // Columns containing dates (Date, Created, Updated)
};

/**
 * Sets up or refreshes the Search sheet
 */
function setupSearchSheet() {
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if Search sheet exists, create if not
  let searchSheet = ss.getSheetByName(SEARCH_CONFIG.SEARCH_SHEET);
  if (!searchSheet) {
    searchSheet = ss.insertSheet(SEARCH_CONFIG.SEARCH_SHEET);
  } else {
    // Clear existing content
    searchSheet.clear();
  }
  
  // Set up search interface
  searchSheet.getRange('A1:G1').merge().setValue('CONTENT CALENDAR SEARCH').setFontWeight('bold')
    .setHorizontalAlignment('center').setBackground('#4285F4').setFontColor('white').setFontSize(14);
  
  searchSheet.getRange('A2').setValue('Filter By:').setFontWeight('bold');
  searchSheet.getRange('A3').setValue('Search Text:').setFontWeight('bold');
  
  // Create dropdown for filter criteria
  const criteria = ['Any Field', 'ID', 'Date', 'Week', 'Status', 'Channel', 'Content', 
                   'Content Pillar', 'Content Format', 'Assigned To', 'Notes'];
  
  const filterRule = SpreadsheetApp.newDataValidation().requireValueInList(criteria, true).build();
  searchSheet.getRange(SEARCH_CONFIG.FILTER_CRITERIA_CELL).setDataValidation(filterRule)
    .setValue('Any Field');
  
  // Add buttons
  searchSheet.getRange('D2').setValue('Search').setFontWeight('bold').setBackground('#4CAF50')
    .setFontColor('white').setHorizontalAlignment('center');
  
  searchSheet.getRange('D3').setValue('Clear Results').setFontWeight('bold').setBackground('#F44336')
    .setFontColor('white').setHorizontalAlignment('center');
  
  searchSheet.getRange('E2:F2').merge().setValue('Total Results:').setFontWeight('bold')
    .setHorizontalAlignment('right');
  
  searchSheet.getRange(SEARCH_CONFIG.TOTAL_RESULTS_CELL).setValue('0');
  
  // Set up results headers (row 4)
  const headers = ['ID', 'Date', 'Week', 'Status', 'Channel', 'Content', 'Actions'];
  searchSheet.getRange(4, 1, 1, headers.length).setValues([headers])
    .setBackground('#EEEEEE').setFontWeight('bold');
  
  // Format columns
  searchSheet.setColumnWidth(1, 100); // ID
  searchSheet.setColumnWidth(2, 100); // Date
  searchSheet.setColumnWidth(3, 80);  // Week
  searchSheet.setColumnWidth(4, 150); // Status
  searchSheet.setColumnWidth(5, 100); // Channel
  searchSheet.setColumnWidth(6, 300); // Content
  searchSheet.setColumnWidth(7, 120); // Actions
  
  // Add search instructions
  const instructions = [
    ['Search Instructions:'],
    ['• Enter search text in the "Search Text" field'],
    ['• Select a specific field to search in from the "Filter By" dropdown'],
    ['• Click "Search" to find matching content items'],
    ['• Click "Clear Results" to reset the search'],
    ['• Click on "Go to Row" in the Actions column to navigate to a specific content item']
  ];
  
  searchSheet.getRange(6 + SEARCH_CONFIG.MAX_RESULTS, 1, instructions.length, 1)
    .setValues(instructions);
  
  // Set up triggers for buttons
  const searchClick = ScriptApp.newTrigger('searchContentCalendar')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  
  const clearClick = ScriptApp.newTrigger('clearSearchResults')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  
  SpreadsheetApp.getUi().alert('Search sheet set up successfully!');
}

/**
 * Searches the content calendar based on filter criteria
 * Triggered by clicking the Search button
 * @param {object} e The onEdit event object
 */
function searchContentCalendar(e) {
  // Exit if this isn't the Search button click
  if (!e) return;
  if (e.range.getA1Notation() !== 'D2' || e.source.getActiveSheet().getName() !== SEARCH_CONFIG.SEARCH_SHEET) return;
  
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const searchSheet = ss.getSheetByName(SEARCH_CONFIG.SEARCH_SHEET);
  const contentSheet = ss.getSheetByName(SEARCH_CONFIG.CONTENT_SHEET);
  
  if (!searchSheet || !contentSheet) {
    SpreadsheetApp.getUi().alert('Required sheets not found. Please check that both Search and Content Calendar sheets exist.');
    return;
  }
  
  // Get search parameters
  const filterCriteria = searchSheet.getRange(SEARCH_CONFIG.FILTER_CRITERIA_CELL).getValue();
  const searchText = searchSheet.getRange(SEARCH_CONFIG.SEARCH_TEXT_CELL).getValue().toString().toLowerCase();
  
  // Exit if search text is empty
  if (!searchText) {
    SpreadsheetApp.getUi().alert('Please enter search text.');
    return;
  }
  
  // Clear previous results
  clearSearchResults();
  
  // Get all data from content calendar (excluding header rows)
  const data = contentSheet.getRange(3, 1, contentSheet.getLastRow() - 2, contentSheet.getLastColumn()).getValues();
  
  // Find the column to search based on filter criteria
  let searchColumn = -1;
  if (filterCriteria !== 'Any Field') {
    for (const field of SEARCH_CONFIG.SEARCH_FIELDS) {
      if (field.name === filterCriteria) {
        searchColumn = field.column - 1; // Adjust for 0-based array
        break;
      }
    }
  }
  
  // Search for matches
  const matches = [];
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    let isMatch = false;
    
    if (searchColumn === -1) {
      // Search in all fields
      for (const field of SEARCH_CONFIG.SEARCH_FIELDS) {
        const colIndex = field.column - 1; // Adjust for 0-based array
        
        if (colIndex >= row.length) continue; // Skip if column doesn't exist
        
        let fieldValue = row[colIndex];
        
        // Handle date fields
        if (SEARCH_CONFIG.DATE_COLUMNS.includes(field.column) && fieldValue instanceof Date) {
          fieldValue = Utilities.formatDate(fieldValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
          fieldValue = fieldValue ? fieldValue.toString().toLowerCase() : '';
        }
        
        if (fieldValue.includes(searchText)) {
          isMatch = true;
          break;
        }
      }
    } else {
      // Search in specific field
      let fieldValue = row[searchColumn];
      
      // Handle date fields
      if (SEARCH_CONFIG.DATE_COLUMNS.includes(searchColumn + 1) && fieldValue instanceof Date) {
        fieldValue = Utilities.formatDate(fieldValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        fieldValue = fieldValue ? fieldValue.toString().toLowerCase() : '';
      }
      
      if (fieldValue.includes(searchText)) {
        isMatch = true;
      }
    }
    
    if (isMatch) {
      matches.push({
        id: row[0],
        date: row[1],
        week: row[2],
        status: row[3],
        channel: row[4],
        content: row[5],
        rowIndex: i + 3 // Adjust for header rows
      });
      
      // Limit results to maximum
      if (matches.length >= SEARCH_CONFIG.MAX_RESULTS) {
        break;
      }
    }
  }
  
  // Display results
  if (matches.length === 0) {
    searchSheet.getRange(SEARCH_CONFIG.TOTAL_RESULTS_CELL).setValue('0');
    SpreadsheetApp.getUi().alert('No matching content found.');
    return;
  }
  
  // Update total results
  searchSheet.getRange(SEARCH_CONFIG.TOTAL_RESULTS_CELL).setValue(matches.length);
  
  // Prepare results data
  const resultsData = [];
  
  for (let i = 0; i < matches.length; i++) {
    const match = matches[i];
    
    // Format date
    let formattedDate = '';
    if (match.date instanceof Date) {
      formattedDate = Utilities.formatDate(match.date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    
    // Truncate content if too long
    let content = match.content || '';
    if (content.length > 50) {
      content = content.substring(0, 47) + '...';
    }
    
    // Add result row
    resultsData.push([
      match.id,
      formattedDate,
      match.week,
      match.status,
      match.channel,
      content,
      'Go to Row ' + match.rowIndex
    ]);
  }
  
  // Write results to sheet
  searchSheet.getRange(5, 1, resultsData.length, 7).setValues(resultsData);
  
  // Add hyperlinks to "Go to Row" cells
  for (let i = 0; i < resultsData.length; i++) {
    searchSheet.getRange(5 + i, 7).setFontColor('blue').setFontWeight('bold')
      .setTextStyle(SpreadsheetApp.newTextStyle().setUnderline(true).build());
  }
}

/**
 * Clears search results
 * Triggered by clicking the Clear Results button
 * @param {object} e The onEdit event object
 */
function clearSearchResults(e) {
  // Check if function was called directly or via event
  let calledDirectly = !e;
  
  // Exit if this isn't the Clear Results button click and not called directly
  if (!calledDirectly && (e.range.getA1Notation() !== 'D3' || e.source.getActiveSheet().getName() !== SEARCH_CONFIG.SEARCH_SHEET)) {
    return;
  }
  
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const searchSheet = ss.getSheetByName(SEARCH_CONFIG.SEARCH_SHEET);
  
  if (!searchSheet) return;
  
  // Clear results area
  searchSheet.getRange(SEARCH_CONFIG.SEARCH_RESULTS_RANGE).clearContent();
  
  // Reset total results
  searchSheet.getRange(SEARCH_CONFIG.TOTAL_RESULTS_CELL).setValue('0');
  
  // Only clear search text if explicitly clicked (not when called programmatically)
  if (!calledDirectly) {
    searchSheet.getRange(SEARCH_CONFIG.SEARCH_TEXT_CELL).clearContent();
  }
}

/**
 * Navigates to a specific row in the content calendar
 * Triggered by clicking a "Go to Row" cell
 * @param {object} e The onEdit event object
 */
function navigateToContentRow(e) {
  // Exit if not in the search sheet
  if (!e || e.source.getActiveSheet().getName() !== SEARCH_CONFIG.SEARCH_SHEET) return;
  
  // Check if clicked cell is in the Actions column
  if (e.range.getColumn() !== 7 || e.range.getRow() < 5) return;
  
  // Get the cell value
  const cellValue = e.range.getValue();
  
  // Check if it's a "Go to Row" cell
  if (!cellValue || !cellValue.toString().startsWith('Go to Row ')) return;
  
  // Extract row number
  const rowNumber = parseInt(cellValue.toString().replace('Go to Row ', ''));
  
  if (isNaN(rowNumber)) return;
  
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contentSheet = ss.getSheetByName(SEARCH_CONFIG.CONTENT_SHEET);
  
  if (!contentSheet) {
    SpreadsheetApp.getUi().alert('Content Calendar sheet not found.');
    return;
  }
  
  // Activate content sheet and scroll to row
  contentSheet.activate();
  contentSheet.setActiveRange(contentSheet.getRange(rowNumber, 1));
}

/**
 * Creates a custom filter for the content calendar
 */
function createCustomFilter() {
  // Get UI
  const ui = SpreadsheetApp.getUi();
  
  // Get the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarSheet = ss.getSheetByName(SEARCH_CONFIG.CONTENT_SHEET);
  
  if (!calendarSheet) {
    ui.alert('Content Calendar sheet not found.');
    return;
  }
  
  // Prompt for filter options
  const response = ui.prompt(
    'Create Custom Filter',
    'Enter filter criteria (field:value,field:value):\n\n' +
    'Example: Status:Planned,Channel:Twitter\n\n' +
    'Available fields: Week, Status, Channel, Content Pillar, Content Format, Assigned To',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const filterText = response.getResponseText().trim();
  
  // Parse filter criteria
  const filterCriteria = {};
  
  const pairs = filterText.split(',');
  for (const pair of pairs) {
    const [field, value] = pair.split(':').map(item => item.trim());
    if (field && value) {
      filterCriteria[field] = value;
    }
  }
  
  // Check if any valid criteria were provided
  if (Object.keys(filterCriteria).length === 0) {
    ui.alert('No valid filter criteria provided.');
    return;
  }
  
  // Map field names to column indexes
  const fieldMap = {
    'Week': 3,
    'Status': 4,
    'Channel': 5,
    'Content': 6,
    'Content Pillar': 7,
    'Content Format': 8,
    'Assigned To': 11
  };
  
  // Create a new sheet for filtered results
  let filterSheet = ss.getSheetByName('Filtered Results');
  if (filterSheet) {
    // If sheet already exists, clear it
    filterSheet.clear();
  } else {
    // Create new sheet
    filterSheet = ss.insertSheet('Filtered Results');
  }
  
  // Copy headers from content calendar
  const headers = calendarSheet.getRange(1, 1, 2, calendarSheet.getLastColumn()).getValues();
  filterSheet.getRange(1, 1, 2, headers[0].length).setValues(headers);
  
  // Format headers same as content calendar
  const headerFormatSource = calendarSheet.getRange(1, 1, 2, headers[0].length);
  const headerFormatDest = filterSheet.getRange(1, 1, 2, headers[0].length);
  
  headerFormatSource.copyFormatToRange(filterSheet, 1, headers[0].length, 1, 2);
  
  // Get all data from content calendar (excluding header rows)
  const data = calendarSheet.getRange(3, 1, calendarSheet.getLastRow() - 2, calendarSheet.getLastColumn()).getValues();
  
  // Filter data
  const filteredData = data.filter(row => {
    for (const field in filterCriteria) {
      const colIndex = fieldMap[field] - a1;
      
      if (colIndex === undefined || colIndex >= row.length) {
        continue;
      }
      
      const cellValue = row[colIndex];
      let stringValue = cellValue !== null && cellValue !== undefined ? cellValue.toString() : '';
      
      // For numerical comparisons
      if (typeof cellValue === 'number' && filterCriteria[field].match(/^[<>]=?\d+$/)) {
        const operator = filterCriteria[field].match(/^([<>]=?)/)[1];
        const compareValue = parseFloat(filterCriteria[field].replace(operator, ''));
        
        switch (operator) {
          case '<': return cellValue < compareValue;
          case '<=': return cellValue <= compareValue;
          case '>': return cellValue > compareValue;
          case '>=': return cellValue >= compareValue;
          default: return false;
        }
      }
      
      // For date comparisons
      if (cellValue instanceof Date && filterCriteria[field].match(/^\d{4}-\d{2}-\d{2}$/)) {
        const compareDate = new Date(filterCriteria[field]);
        return cellValue.toDateString() === compareDate.toDateString();
      }
      
      // For exact match
      if (stringValue !== filterCriteria[field]) {
        return false;
      }
    }
    
    return true;
  });
  
  // Display results
  if (filteredData.length === 0) {
    ui.alert('No content items match the filter criteria.');
    
    // Add a no results message
    filterSheet.getRange(3, 1).setValue('No matching results found.');
    return;
  }
  
  // Write filtered data to sheet
  filterSheet.getRange(3, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
  
  // Copy formatting from content calendar
  const formatSource = calendarSheet.getRange(3, 1, 1, data[0].length);
  
  for (let i = 0; i < filteredData.length; i++) {
    const formatDest = filterSheet.getRange(3 + i, 1, 1, data[0].length);
    formatSource.copyFormatToRange(filterSheet, 1, data[0].length, 3 + i, 3 + i);
  }
  
  // Adjust column widths to match content calendar
  for (let i = 1; i <= data[0].length; i++) {
    filterSheet.setColumnWidth(i, calendarSheet.getColumnWidth(i));
  }
  
  // Add filter details at the top
  filterSheet.insertRowBefore(1);
  
  let filterDescription = 'Custom Filter: ';
  for (const field in filterCriteria) {
    filterDescription += `${field} = ${filterCriteria[field]}, `;
  }
  filterDescription = filterDescription.slice(0, -2); // Remove trailing comma and space
  
  filterSheet.getRange(1, 1, 1, 5).merge().setValue(filterDescription)
    .setBackground('#EEEEEE').setFontWeight('bold');
  
  // Activate the filter sheet
  filterSheet.activate();
  
  // Show results message
  ui.alert(`Filter applied successfully. Found ${filteredData.length} matching items.`);
}

/**
 * Creates a search menu
 */
function createSearchMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Search & Filter')
    .addItem('Set Up Search Interface', 'setupSearchSheet')
    .addItem('Create Custom Filter', 'createCustomFilter')
    .addToUi();
}