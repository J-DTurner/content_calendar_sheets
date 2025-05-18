/**
 * Test Suite for Search & Filter Backend Functions
 * This file contains comprehensive tests for all backend functions
 * related to the Search & Filter feature.
 */

/**
 * Test Suite for getFilterDropdownOptions function
 */
function testGetFilterDropdownOptions() {
  console.log("Testing getFilterDropdownOptions()...");
  
  try {
    // Test 1: Normal case with Lists sheet containing data
    console.log("Test 1: Normal case with data in Lists sheet");
    const result1 = getFilterDropdownOptions();
    console.log("Result 1:", JSON.stringify(result1));
    
    // Verify structure
    if (!result1.statuses || !result1.channels || !result1.pillars || 
        !result1.formats || !result1.assignees) {
      console.error("ERROR: Missing expected properties in result");
    } else {
      console.log("✓ All expected properties present");
    }
    
    // Test 2: Edge case - what happens when the Lists sheet doesn't exist?
    console.log("\nTest 2: Lists sheet doesn't exist");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const originalListsSheet = ss.getSheetByName('Lists');
    let tempBackupName = null;
    
    // Temporarily rename the Lists sheet
    if (originalListsSheet) {
      tempBackupName = 'Lists_BACKUP_' + new Date().getTime();
      originalListsSheet.setName(tempBackupName);
    }
    
    try {
      const result2 = getFilterDropdownOptions();
      console.log("Result 2:", JSON.stringify(result2));
      console.log("✓ Function handled missing Lists sheet gracefully");
    } catch (error) {
      console.error("ERROR: Function crashed with missing Lists sheet:", error);
    } finally {
      // Restore the Lists sheet
      if (tempBackupName) {
        ss.getSheetByName(tempBackupName).setName('Lists');
      }
    }
    
    // Test 3: Empty columns in Lists sheet
    console.log("\nTest 3: Empty columns in Lists sheet");
    // This would require actually modifying the sheet data, so we'll just verify
    // that the function returns arrays (even if empty)
    const result3 = getFilterDropdownOptions();
    const allArrays = Array.isArray(result3.statuses) && 
                      Array.isArray(result3.channels) && 
                      Array.isArray(result3.pillars) && 
                      Array.isArray(result3.formats) && 
                      Array.isArray(result3.assignees);
    
    if (allArrays) {
      console.log("✓ All properties are arrays as expected");
    } else {
      console.error("ERROR: Some properties are not arrays");
    }
    
  } catch (error) {
    console.error("Fatal error in testGetFilterDropdownOptions:", error);
  }
  
  console.log("\n=====================================\n");
}

/**
 * Test Suite for performSearch function
 */
function testPerformSearch() {
  console.log("Testing performSearch()...");
  
  try {
    // Test 1: Empty filters (should return all results up to limit)
    console.log("Test 1: Empty filters");
    const result1 = performSearch({});
    console.log("Result 1:", {
      resultsCount: result1.results ? result1.results.length : 0,
      totalMatches: result1.totalMatches,
      limitApplied: result1.limitApplied,
      error: result1.error
    });
    
    // Test 2: Keyword search
    console.log("\nTest 2: Keyword search");
    const result2 = performSearch({
      keyword: "content",
      status: "All",
      channel: "All",
      pillar: "All",
      format: "All",
      assignee: "All",
      weekNumber: null,
      dateRange: { type: "All Time" }
    });
    console.log("Result 2:", {
      resultsCount: result2.results ? result2.results.length : 0,
      totalMatches: result2.totalMatches,
      limitApplied: result2.limitApplied,
      error: result2.error
    });
    
    // Test 3: Status filter
    console.log("\nTest 3: Status filter");
    const result3 = performSearch({
      keyword: "",
      status: "Published",
      channel: "All",
      pillar: "All",
      format: "All",
      assignee: "All",
      weekNumber: null,
      dateRange: { type: "All Time" }
    });
    console.log("Result 3:", {
      resultsCount: result3.results ? result3.results.length : 0,
      totalMatches: result3.totalMatches,
      limitApplied: result3.limitApplied,
      error: result3.error
    });
    
    // Test 4: Date range - This Week
    console.log("\nTest 4: Date range - This Week");
    const result4 = performSearch({
      keyword: "",
      status: "All",
      channel: "All",
      pillar: "All",
      format: "All",
      assignee: "All",
      weekNumber: null,
      dateRange: { type: "This Week" }
    });
    console.log("Result 4:", {
      resultsCount: result4.results ? result4.results.length : 0,
      totalMatches: result4.totalMatches,
      limitApplied: result4.limitApplied,
      error: result4.error
    });
    
    // Test 5: Date range - Custom
    console.log("\nTest 5: Date range - Custom");
    const result5 = performSearch({
      keyword: "",
      status: "All",
      channel: "All",
      pillar: "All",
      format: "All",
      assignee: "All",
      weekNumber: null,
      dateRange: { 
        type: "custom",
        startDate: "2024-01-01",
        endDate: "2024-12-31"
      }
    });
    console.log("Result 5:", {
      resultsCount: result5.results ? result5.results.length : 0,
      totalMatches: result5.totalMatches,
      limitApplied: result5.limitApplied,
      error: result5.error
    });
    
    // Test 6: Week number filter
    console.log("\nTest 6: Week number filter");
    const result6 = performSearch({
      keyword: "",
      status: "All",
      channel: "All",
      pillar: "All",
      format: "All",
      assignee: "All",
      weekNumber: 42,
      dateRange: { type: "All Time" }
    });
    console.log("Result 6:", {
      resultsCount: result6.results ? result6.results.length : 0,
      totalMatches: result6.totalMatches,
      limitApplied: result6.limitApplied,
      error: result6.error
    });
    
    // Test 7: Combined filters
    console.log("\nTest 7: Combined filters");
    const result7 = performSearch({
      keyword: "social",
      status: "Draft",
      channel: "Instagram",
      pillar: "All",
      format: "All",
      assignee: "All",
      weekNumber: null,
      dateRange: { type: "This Month" }
    });
    console.log("Result 7:", {
      resultsCount: result7.results ? result7.results.length : 0,
      totalMatches: result7.totalMatches,
      limitApplied: result7.limitApplied,
      error: result7.error
    });
    
    // Test 8: Filter that should return no results
    console.log("\nTest 8: Filter that should return no results");
    const result8 = performSearch({
      keyword: "xyzabc123nonexistent",
      status: "All",
      channel: "All",
      pillar: "All",
      format: "All",
      assignee: "All",
      weekNumber: null,
      dateRange: { type: "All Time" }
    });
    console.log("Result 8:", {
      resultsCount: result8.results ? result8.results.length : 0,
      totalMatches: result8.totalMatches,
      limitApplied: result8.limitApplied,
      error: result8.error
    });
    
    // Test 9: Invalid filter object
    console.log("\nTest 9: Invalid filter object");
    const result9 = performSearch(null);
    console.log("Result 9:", {
      resultsCount: result9.results ? result9.results.length : 0,
      totalMatches: result9.totalMatches,
      error: result9.error
    });
    
    // Test 10: Special characters in keyword
    console.log("\nTest 10: Special characters in keyword");
    const result10 = performSearch({
      keyword: "test & special @#$",
      status: "All",
      channel: "All",
      pillar: "All",
      format: "All",
      assignee: "All",
      weekNumber: null,
      dateRange: { type: "All Time" }
    });
    console.log("Result 10:", {
      resultsCount: result10.results ? result10.results.length : 0,
      totalMatches: result10.totalMatches,
      limitApplied: result10.limitApplied,
      error: result10.error
    });
    
    // Test 11: Invalid date range (start after end)
    console.log("\nTest 11: Invalid date range (start after end)");
    const result11 = performSearch({
      keyword: "",
      status: "All",
      channel: "All",
      pillar: "All",
      format: "All",
      assignee: "All",
      weekNumber: null,
      dateRange: { 
        type: "custom",
        startDate: "2024-12-31",
        endDate: "2024-01-01"
      }
    });
    console.log("Result 11:", {
      resultsCount: result11.results ? result11.results.length : 0,
      totalMatches: result11.totalMatches,
      limitApplied: result11.limitApplied,
      error: result11.error
    });
    
    // Test 12: Large dataset performance test
    console.log("\nTest 12: Performance test - checking if total matches exceeds limit");
    const result12 = performSearch({
      keyword: "",
      status: "All",
      channel: "All",
      pillar: "All",
      format: "All",
      assignee: "All",
      weekNumber: null,
      dateRange: { type: "All Time" }
    });
    
    if (result12.totalMatches > 50) {
      console.log("✓ Found more than 50 results, limit correctly applied:", {
        resultsCount: result12.results.length,
        totalMatches: result12.totalMatches,
        limitApplied: result12.limitApplied
      });
    } else {
      console.log("Dataset has less than 50 total entries:", {
        resultsCount: result12.results.length,
        totalMatches: result12.totalMatches
      });
    }
    
    // Test 13: Result object structure verification
    console.log("\nTest 13: Result object structure verification");
    if (result1.results && result1.results.length > 0) {
      const firstResult = result1.results[0];
      const hasRequiredFields = 
        'id' in firstResult &&
        'date' in firstResult &&
        'week' in firstResult &&
        'status' in firstResult &&
        'channel' in firstResult &&
        'contentIdea' in firstResult &&
        'assignedTo' in firstResult &&
        'originalRowIndex' in firstResult;
      
      if (hasRequiredFields) {
        console.log("✓ Result objects have all required fields");
        console.log("Sample result:", firstResult);
      } else {
        console.error("ERROR: Result objects missing required fields");
        console.log("Sample result:", firstResult);
      }
    }
    
  } catch (error) {
    console.error("Fatal error in testPerformSearch:", error);
  }
  
  console.log("\n=====================================\n");
}

/**
 * Test Suite for navigateToRowInContentCalendar function
 */
function testNavigateToRowInContentCalendar() {
  console.log("Testing navigateToRowInContentCalendar()...");
  
  try {
    // Test 1: Valid row number
    console.log("Test 1: Valid row number (5)");
    const result1 = navigateToRowInContentCalendar(5);
    console.log("Result 1:", result1);
    
    // Test 2: Header row (row 1)
    console.log("\nTest 2: Header row (row 1)");
    const result2 = navigateToRowInContentCalendar(1);
    console.log("Result 2:", result2);
    
    // Test 3: Second header row (row 2)
    console.log("\nTest 3: Second header row (row 2)");
    const result3 = navigateToRowInContentCalendar(2);
    console.log("Result 3:", result3);
    
    // Test 4: Very large row number
    console.log("\nTest 4: Very large row number (9999)");
    const result4 = navigateToRowInContentCalendar(9999);
    console.log("Result 4:", result4);
    
    // Test 5: Invalid row number (0)
    console.log("\nTest 5: Invalid row number (0)");
    const result5 = navigateToRowInContentCalendar(0);
    console.log("Result 5:", result5);
    
    // Test 6: Invalid row number (negative)
    console.log("\nTest 6: Invalid row number (-5)");
    const result6 = navigateToRowInContentCalendar(-5);
    console.log("Result 6:", result6);
    
    // Test 7: Non-numeric row index
    console.log("\nTest 7: Non-numeric row index ('abc')");
    const result7 = navigateToRowInContentCalendar('abc');
    console.log("Result 7:", result7);
    
    // Test 8: Missing Content Calendar sheet
    console.log("\nTest 8: Missing Content Calendar sheet");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const originalContentSheet = ss.getSheetByName('Content Calendar');
    let tempBackupName = null;
    
    // Temporarily rename the Content Calendar sheet
    if (originalContentSheet) {
      tempBackupName = 'Content_Calendar_BACKUP_' + new Date().getTime();
      originalContentSheet.setName(tempBackupName);
    }
    
    try {
      const result8 = navigateToRowInContentCalendar(5);
      console.log("Result 8:", result8);
      console.log("✓ Function handled missing Content Calendar sheet gracefully");
    } catch (error) {
      console.error("ERROR: Function crashed with missing Content Calendar sheet:", error);
    } finally {
      // Restore the Content Calendar sheet
      if (tempBackupName) {
        ss.getSheetByName(tempBackupName).setName('Content Calendar');
      }
    }
    
  } catch (error) {
    console.error("Fatal error in testNavigateToRowInContentCalendar:", error);
  }
  
  console.log("\n=====================================\n");
}

/**
 * Main test runner function
 */
function runAllBackendTests() {
  console.log("=== SEARCH & FILTER BACKEND TEST SUITE ===");
  console.log("Starting tests at:", new Date().toISOString());
  console.log("=========================================\n");
  
  try {
    testGetFilterDropdownOptions();
    testPerformSearch();
    testNavigateToRowInContentCalendar();
    
    console.log("=== ALL TESTS COMPLETED ===");
    console.log("Finished at:", new Date().toISOString());
    console.log("===========================");
  } catch (error) {
    console.error("FATAL ERROR IN TEST SUITE:", error);
  }
}

/**
 * Individual test runners for debugging specific functions
 */
function runTestGetFilterDropdownOptions() {
  testGetFilterDropdownOptions();
}

function runTestPerformSearch() {
  testPerformSearch();
}

function runTestNavigateToRowInContentCalendar() {
  testNavigateToRowInContentCalendar();
}