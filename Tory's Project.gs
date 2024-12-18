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

// Optional: Add trigger to run search when B3 changes
function createOnEditTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}

function onEdit(e) {
  if (e.range.getA1Notation() === 'B3' && e.source.getActiveSheet().getName() === 'Search') {
    searchAccounts();
  }
}
