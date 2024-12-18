function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Account Matcher')
    .addItem('Search Accounts', 'searchAccounts')
    .addToUi();
}

// Clean company names by removing unwanted suffixes and designations
function cleanCompanyName(name) {
  if (!name) return '';
  
  // Convert to string and trim
  name = String(name).trim();
  
  // Remove common suffixes and designations
  const patterns = [
    /\s*-\s*Parent/i,
    /\s+(Corp|Inc|LLC)(\.|\b)/i,
    /\s+(APAC|EMU|Dev)\b/i,
    /\s+(Copilot|AE)\b/i,
    /\s+Team\b/i,
    /\s+Test\b/i
  ];
  
  patterns.forEach(pattern => {
    name = name.replace(pattern, '');
  });
  
  return name.trim();
}

// Find matching accounts for an account owner
function findMatchingAccounts(accountOwner) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accountsSheet = ss.getSheetByName('All_Accounts');
  const accountsData = accountsSheet.getDataRange().getValues();
  
  // Get all accounts for the account owner
  let accounts = accountsData
    .filter(row => row[1] === accountOwner) // Column B (Account_Owner)
    .map(row => cleanCompanyName(row[2])) // Column C (Account_Name)
    .filter(name => name); // Remove empty values
  
  // Remove duplicates (case-insensitive)
  accounts = [...new Set(accounts.map(name => name.toLowerCase()))]
    .map(name => name.charAt(0).toUpperCase() + name.slice(1));
  
  return accounts;
}

// Find matching leads for the given accounts
function findMatchingLeads(accounts) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leadsSheet = ss.getSheetByName('Paste_Leads_Here');
  const leadsData = leadsSheet.getDataRange().getValues();
  const headers = leadsData.shift(); // Remove headers
  
  // Find leads where company matches any account (including partial matches)
  return leadsData.filter(lead => {
    const companyName = cleanCompanyName(lead[3]); // Column D (COMPANY)
    return accounts.some(account => 
      companyName.toLowerCase().includes(account.toLowerCase()) ||
      account.toLowerCase().includes(companyName.toLowerCase())
    );
  });
}

// Update the Search sheet with matching leads
function updateSearchSheet(matchingLeads, accountOwner) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const searchSheet = ss.getSheetByName('Search');
  
  // Clear existing results
  const lastRow = Math.max(searchSheet.getLastRow(), 2);
  searchSheet.getRange(3, 4, lastRow - 2, 7).clearContent();
  
  // If no matching leads, exit
  if (!matchingLeads.length) return;
  
  // Prepare data for Search sheet
  const searchData = matchingLeads.map(lead => [
    accountOwner,      // Matched_Rep (Column D)
    lead[1],          // Last_Name (Column E)
    lead[2],          // First_Name (Column F)
    lead[0],          // Email (Column G)
    lead[3],          // Company_Name (Column H)
    lead[4],          // Job_Title (Column I)
    lead[5]           // Country (Column J)
  ]);
  
  // Update Search sheet
  searchSheet.getRange(3, 4, searchData.length, 7).setValues(searchData);
}

// Main search function triggered by UI
function searchAccounts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const searchSheet = ss.getSheetByName('Search');
  const accountOwner = searchSheet.getRange('B3').getValue();
  
  if (!accountOwner) {
    searchSheet.getRange(3, 4, searchSheet.getLastRow() - 2, 7).clearContent();
    return;
  }
  
  const matchingAccounts = findMatchingAccounts(accountOwner);
  const matchingLeads = findMatchingLeads(matchingAccounts);
  updateSearchSheet(matchingLeads, accountOwner);
}

// Modified onEdit trigger
function onEdit(e) {
  // Check if edit was in Search sheet B3 for original functionality
  if (e.range.getA1Notation() === 'B3' && e.source.getActiveSheet().getName() === 'Search') {
    searchAccounts();
    return;
  }
  
  // Check if edit was in SortList sheet A3 (checkbox)
  if (e.range.getA1Notation() === 'A3' && e.source.getActiveSheet().getName() === 'SortList') {
    try {
      const isChecked = e.value === 'TRUE';
      handleSortListCheckbox(isChecked);
    } catch (error) {
      Logger.log('Error in onEdit: ' + error.toString());
    }
  }
}

// Modified handleSortListCheckbox function
function handleSortListCheckbox(isChecked) {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  if (isChecked) {
    // Save current state and reorganize
    try {
      saveOriginalState(sheet);
      reorganizeColumns(sheet);
    } catch (error) {
      Logger.log('Error when checking: ' + error.toString());
    }
  } else {
    // Restore original state
    try {
      restoreOriginalState(sheet);
    } catch (error) {
      Logger.log('Error when unchecking: ' + error.toString());
    }
  }
}

// Modified saveOriginalState function
function saveOriginalState(sheet) {
  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  
  // Check if there's data to save
  if (lastColumn < 3 || lastRow < 1) return;
  
  const range = sheet.getRange(1, 3, lastRow, lastColumn - 2);
  const values = range.getValues();
  
  // Save as a cache using PropertiesService
  PropertiesService.getScriptProperties().setProperties({
    'originalData': JSON.stringify(values),
    'lastColumn': lastColumn.toString(),
    'lastRow': lastRow.toString()
  });
}

// Modified reorganizeColumns function
function reorganizeColumns(sheet) {
  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  
  // Check if there's data to process
  if (lastColumn < 3 || lastRow < 1) return;
  
  const values = sheet.getRange(1, 3, lastRow, lastColumn - 2).getValues();
  const headers = values[0];
  const data = values.slice(1);
  
  // Define desired header order
  const desiredHeaders = ['Email', 'Last Name', 'First Name', 'Company', 'Job Title', 'Country', 'State'];
  
  // Create mapping of current indices to desired indices
  const headerMap = new Map();
  headers.forEach((header, index) => {
    const desiredIndex = desiredHeaders.findIndex(
      h => h.toLowerCase() === header.toString().toLowerCase()
    );
    if (desiredIndex !== -1) {
      headerMap.set(desiredIndex, index);
    }
  });
  
  // Create new arrays for reorganized data
  const newHeaders = [];
  const newData = Array(data.length).fill().map(() => []);
  
  // Populate new arrays in desired order
  desiredHeaders.forEach((header, desiredIndex) => {
    const currentIndex = headerMap.get(desiredIndex);
    if (currentIndex !== undefined) {
      newHeaders.push(headers[currentIndex]);
      data.forEach((row, rowIndex) => {
        newData[rowIndex].push(row[currentIndex]);
      });
    }
  });
  
  // Clear the existing content
  if (lastRow > 1) {
    sheet.getRange(1, 3, lastRow, lastColumn - 2).clearContent();
  }
  
  // Write the reorganized data if we have any
  if (newHeaders.length > 0) {
    sheet.getRange(1, 3, 1, newHeaders.length).setValues([newHeaders]);
    if (newData.length > 0) {
      sheet.getRange(2, 3, newData.length, newHeaders.length).setValues(newData);
    }
  }
}

// Modified restoreOriginalState function
function restoreOriginalState(sheet) {
  const props = PropertiesService.getScriptProperties();
  const originalDataStr = props.getProperty('originalData');
  const lastColumn = parseInt(props.getProperty('lastColumn'));
  const lastRow = parseInt(props.getProperty('lastRow'));
  
  // Check if we have data to restore
  if (!originalDataStr || !lastColumn || !lastRow) return;
  
  try {
    const originalData = JSON.parse(originalDataStr);
    
    // Clear existing content
    const currentLastRow = Math.max(sheet.getLastRow(), lastRow);
    const currentLastCol = Math.max(sheet.getLastColumn(), lastColumn);
    if (currentLastRow > 1 && currentLastCol > 2) {
      sheet.getRange(1, 3, currentLastRow, currentLastCol - 2).clearContent();
    }
    
    // Restore the original data
    if (originalData.length > 0) {
      sheet.getRange(1, 3, originalData.length, originalData[0].length).setValues(originalData);
    }
  } catch (error) {
    Logger.log('Error restoring data: ' + error.toString());
  }
}

// Optional: Add trigger to run search when B3 changes
function createOnEditTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}
