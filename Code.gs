function matchAndCopyRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("Inputs"); // Changed from "Sheet1" to "Inputs"
  var targetSheet = ss.getSheetByName("Matches");

  if (!sourceSheet) {
    throw new Error("Sheet 'Inputs' not found. Please ensure the sheet is named correctly.");
  }

  var repAccountsColumnName = "rep_accounts";
  var companyNameColumnName = "company_names";

  var sourceData = sourceSheet.getDataRange().getValues();
  var headerRowIndex = 2; // Headers are in row 3 (index 2)
  var headerRow = sourceData[headerRowIndex];

  var repAccountsIndex = headerRow.indexOf(repAccountsColumnName);
  var companyNameIndex = headerRow.indexOf(companyNameColumnName);

  if (repAccountsIndex === -1 || companyNameIndex === -1) {
    throw new Error("One or both specified columns not found. Please check the column names.");
  }

  // Get rep accounts (excluding header row)
  var repAccounts = sourceData.slice(headerRowIndex + 1).map(row => row[repAccountsIndex]).filter(String);
  
  function cleanAndTokenize(name) {
    if (typeof name !== 'string') {
      Logger.log("Warning: Non-string value encountered: " + name);
      return [];
    }
    return name.toString().toLowerCase().replace(/[^a-z0-9\s]/g, '').split(/\s+/).filter(word => word.length > 1);
  }

  // Preprocess rep accounts
  var processedRepAccounts = repAccounts.map(cleanAndTokenize);

  // Common abbreviations and their full forms
  var abbreviations = {
    'ins': 'insurance',
    'corp': 'corporation',
    'inc': 'incorporated'
  };

  function expandAbbreviations(tokens) {
    return tokens.map(token => abbreviations[token] || token);
  }

  function calculateMatchScore(repTokens, companyTokens) {
    repTokens = expandAbbreviations(repTokens);
    companyTokens = expandAbbreviations(companyTokens);

    var matchedTokens = repTokens.filter(token => companyTokens.includes(token));
    var repMatchRatio = repTokens.length > 0 ? matchedTokens.length / repTokens.length : 0;
    var companyMatchRatio = companyTokens.length > 0 ? matchedTokens.length / companyTokens.length : 0;
    
    // Boost score for partial name matches (e.g., "Aon" matching "Aon Corporation")
    if (repTokens[0] === companyTokens[0] && repTokens.length > companyTokens.length) {
      repMatchRatio = Math.max(repMatchRatio, 0.8);
    }

    return Math.max(repMatchRatio, companyMatchRatio); // Use the higher of the two ratios
  }

  var matchedRows = [];
  var additionalColumns = ["Match Type", "Accuracy Score", "Matched Rep Account"];
  
  // Create the header row
  var newHeaderRow = headerRow.filter((_, i) => i !== repAccountsIndex).concat(additionalColumns);
  matchedRows.push(newHeaderRow);

  sourceData.forEach(function(row, index) {
    if (index <= headerRowIndex) return; // Skip header rows
    
    var companyName = row[companyNameIndex];
    if (!companyName) return; // Skip empty company names
    
    var companyTokens = cleanAndTokenize(companyName);
    
    var bestMatch = {score: 0, repAccount: ""};
    processedRepAccounts.forEach((repTokens, i) => {
      var matchScore = calculateMatchScore(repTokens, companyTokens);
      if (matchScore > bestMatch.score) {
        bestMatch.score = matchScore;
        bestMatch.repAccount = repAccounts[i];
      }
    });
    
    if (bestMatch.score >= 0.7) { // Adjusted threshold for partial matches
      var matchType = bestMatch.score === 1 ? "Exact" : "Partial";
      // Create a new row without the rep_accounts column
      var newRow = row.filter((_, i) => i !== repAccountsIndex);
      matchedRows.push(newRow.concat([matchType, bestMatch.score, bestMatch.repAccount]));
      Logger.log(matchType + " match found: " + bestMatch.repAccount + " ~ " + companyName + " (Match score: " + bestMatch.score + ")");
    }
  });

  // Sort matched rows by accuracy score (descending)
  matchedRows.sort((a, b) => b[b.length - 2] - a[a.length - 2]);

  Logger.log("Total matches found: " + (matchedRows.length - 1));

  // Clear the target sheet and paste new data
  targetSheet.clear();
  if (matchedRows.length > 1) {
    targetSheet.getRange(1, 1, matchedRows.length, matchedRows[0].length).setValues(matchedRows);
    targetSheet.getRange(1, 1, matchedRows.length, matchedRows[0].length).setNumberFormat("@"); // Set all cells to text format
    targetSheet.getRange(2, matchedRows[0].length - 2, matchedRows.length - 1, 1).setNumberFormat("0.00"); // Set accuracy score to 2 decimal places
    SpreadsheetApp.getUi().alert((matchedRows.length - 1) + " matching rows have been copied to the Matches sheet.");
  } else {
    SpreadsheetApp.getUi().alert("No matches found. Please check your data.");
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Match and Copy Rows', 'matchAndCopyRows')
    .addToUi();
}
