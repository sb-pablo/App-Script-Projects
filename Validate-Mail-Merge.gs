function validateEmailsAdvanced() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  const headerCell = sheet.getRange(1, 1).getValue().toString().trim();
  let headerWarning = "";
  if (headerCell !== "Email" && headerCell !== "email") {
    headerWarning = "Warning: Header cell A1 contains '" + headerCell + "' instead of 'Email'. This might cause Mail Merge issues.\n\n";
  }
  
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
  const emailData = dataRange.getValues();
  
  let validCount = 0;
  let invalidCount = 0;
  let emptyCount = 0;
  let warningCount = 0;
  let duplicateCount = 0;
  let invalidEmails = [];
  let warningEmails = [];
  let duplicateEmails = [];
  
  const emailsFound = {};
  
  dataRange.setBackground(null);
  
  const strictEmailRegex = /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$/;
  
  const multipleEmailCheck = /[,;]/;
  
  const invisibleCharCheck = /[\u0000-\u001F\u007F-\u009F\u00A0\u2000-\u200F\u2028-\u202F\u205F-\u206F\uFEFF]/;
  
  for (let i = 0; i < emailData.length; i++) {
    const rawValue = emailData[i][0];
    const email = (rawValue || "").toString().trim();
    const currentRow = i + 2;
    
    if (email === "") {
      sheet.getRange(currentRow, 1).setBackground("#FFFF00");
      emptyCount++;
      invalidEmails.push(`Row ${currentRow}: Empty email`);
    } else if (multipleEmailCheck.test(email)) {
      sheet.getRange(currentRow, 1).setBackground("#FFA500");
      warningCount++;
      warningEmails.push(`Row ${currentRow}: Possible multiple emails (${email})`);
    } else if (invisibleCharCheck.test(email)) {
      sheet.getRange(currentRow, 1).setBackground("#FFA500");
      warningCount++;
      warningEmails.push(`Row ${currentRow}: Contains invisible characters (${email})`);
    } else if (!strictEmailRegex.test(email)) {
      sheet.getRange(currentRow, 1).setBackground("#FF0000");
      invalidCount++;
      invalidEmails.push(`Row ${currentRow}: Invalid format (${email})`);
    } else if (email !== email.toLowerCase()) {
      sheet.getRange(currentRow, 1).setBackground("#E6E6FA");
      warningCount++;
      warningEmails.push(`Row ${currentRow}: Contains uppercase (${email})`);
      
      const lowerEmail = email.toLowerCase();
      if (emailsFound[lowerEmail]) {
        sheet.getRange(currentRow, 1).setBackground("#FF00FF");
        duplicateCount++;
        duplicateEmails.push(`Row ${currentRow}: Duplicate of row ${emailsFound[lowerEmail].row} (${email})`);
      } else {
        validCount++;
        emailsFound[lowerEmail] = { row: currentRow, value: email };
      }
    } else {
      if (emailsFound[email]) {
        sheet.getRange(currentRow, 1).setBackground("#FF00FF");
        duplicateCount++;
        duplicateEmails.push(`Row ${currentRow}: Duplicate of row ${emailsFound[email].row} (${email})`);
      } else {
        validCount++;
        emailsFound[email] = { row: currentRow, value: email };
      }
    }
  }
  
  let lastRow = 0;
  for (let i = emailData.length - 1; i >= 0; i--) {
    const email = emailData[i][0].toString().trim();
    if (email !== "") {
      lastRow = i + 2;
      break;
    }
  }
  
  if (lastRow < sheet.getLastRow() && lastRow > 0) {
    warningEmails.push(`Warning: There appear to be blank rows after your data. Last email is in row ${lastRow}, but sheet extends to row ${sheet.getLastRow()}`);
    warningCount++;
  }
  
  let message = headerWarning + 
                "Validation Summary:\n" +
                `Total emails checked: ${emailData.length}\n` +
                `Valid unique emails: ${validCount}\n` +
                `Invalid emails: ${invalidCount}\n` +
                `Duplicate emails: ${duplicateCount}\n` +
                `Warnings: ${warningCount}\n` +
                `Empty cells: ${emptyCount}\n\n`;
  
  if (invalidEmails.length > 0) {
    message += "Invalid emails found:\n" + invalidEmails.join("\n") + "\n\n";
  }
  
  if (duplicateEmails.length > 0) {
    message += "Duplicate emails found:\n" + duplicateEmails.join("\n") + "\n\n";
  }
  
  if (warningEmails.length > 0) {
    message += "Warnings:\n" + warningEmails.join("\n");
  }
  
  if (invalidEmails.length > 0 || warningEmails.length > 0 || duplicateEmails.length > 0) {
    SpreadsheetApp.getUi().alert("Email Validation Issues", message, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    message += "All emails appear valid and ready for Gmail Mail Merge!";
    SpreadsheetApp.getUi().alert("Email Validation Successful", message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
  
  console.log(message);
  
  return {
    valid: validCount,
    invalid: invalidCount,
    warnings: warningCount,
    empty: emptyCount,
    duplicates: duplicateCount,
    issues: invalidEmails,
    warningItems: warningEmails,
    duplicateItems: duplicateEmails
  };
}

function cleanEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  const response = SpreadsheetApp.getUi().alert(
    "Create Backup?",
    "Do you want to create a backup of your email column before cleaning?",
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  
  if (response == SpreadsheetApp.getUi().Button.YES) {
    const dataRange = sheet.getRange(1, 1, sheet.getLastRow(), 1);
    const targetRange = sheet.getRange(1, 2, sheet.getLastRow(), 1);
    dataRange.copyTo(targetRange);
    sheet.getRange(1, 2).setValue("Email Backup");
  }
  
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
  const emailData = dataRange.getValues();
  
  let changedCount = 0;
  let cleanedEmails = [];
  
  for (let i = 0; i < emailData.length; i++) {
    const rawValue = emailData[i][0];
    const originalEmail = (rawValue || "").toString();
    const currentRow = i + 2;
    
    if (originalEmail.trim() === "") continue;
    
    let cleanedEmail = originalEmail.trim().toLowerCase();
    cleanedEmail = cleanedEmail.replace(/[\u0000-\u001F\u007F-\u009F\u00A0\u2000-\u200F\u2028-\u202F\u205F-\u206F\uFEFF]/g, "");
    
    if (/[,;]/.test(cleanedEmail)) {
      const firstEmail = cleanedEmail.split(/[,;]/)[0].trim();
      cleanedEmail = firstEmail;
    }
    
    if (cleanedEmail !== originalEmail) {
      sheet.getRange(currentRow, 1).setValue(cleanedEmail);
      changedCount++;
      cleanedEmails.push(`Row ${currentRow}: "${originalEmail}" → "${cleanedEmail}"`);
    }
  }
  
  const headerCell = sheet.getRange(1, 1).getValue().toString().trim();
  if (headerCell !== "Email") {
    sheet.getRange(1, 1).setValue("Email");
    changedCount++;
    cleanedEmails.push(`Row 1: Header changed from "${headerCell}" to "Email"`);
  }
  
  let message = "";
  if (changedCount > 0) {
    message = `Cleaned ${changedCount} email addresses:\n\n` + cleanedEmails.join("\n");
    SpreadsheetApp.getUi().alert("Email Cleaning Complete", message, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    message = "No emails needed cleaning. All emails are already in the correct format.";
    SpreadsheetApp.getUi().alert("Email Cleaning Complete", message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
  
  console.log(message);
  return {changed: changedCount, details: cleanedEmails};
}

function removeDuplicateEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  const response = SpreadsheetApp.getUi().alert(
    "Create Backup?",
    "Do you want to create a backup of your email column before removing duplicates?",
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  
  if (response == SpreadsheetApp.getUi().Button.YES) {
    const dataRange = sheet.getRange(1, 1, sheet.getLastRow(), 1);
    const targetRange = sheet.getRange(1, 2, sheet.getLastRow(), 1);
    dataRange.copyTo(targetRange);
    sheet.getRange(1, 2).setValue("Email Backup");
  }
  
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
  const emailData = dataRange.getValues();
  
  const emailsFound = {};
  const rowsToDelete = [];
  let duplicateCount = 0;
  let duplicateDetails = [];
  
  for (let i = 0; i < emailData.length; i++) {
    const rawValue = emailData[i][0];
    const email = (rawValue || "").toString().trim().toLowerCase();
    const currentRow = i + 2;
    
    if (email === "") continue;
    
    if (emailsFound[email]) {
      rowsToDelete.push(currentRow);
      duplicateCount++;
      duplicateDetails.push(`Row ${currentRow}: Duplicate of row ${emailsFound[email].row} (${email})`);
    } else {
      emailsFound[email] = { row: currentRow, value: email };
    }
  }
  
  rowsToDelete.sort((a, b) => b - a);
  
  for (const row of rowsToDelete) {
    sheet.deleteRow(row);
  }
  
  let message = "";
  if (duplicateCount > 0) {
    message = `Removed ${duplicateCount} duplicate emails:\n\n` + duplicateDetails.join("\n");
    SpreadsheetApp.getUi().alert("Duplicates Removed", message, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    message = "No duplicate emails were found.";
    SpreadsheetApp.getUi().alert("Duplicates Check Complete", message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
  
  console.log(message);
  return {removed: duplicateCount, details: duplicateDetails};
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge Tools')
    .addItem('Validate Emails (Advanced)', 'validateEmailsAdvanced')
    .addItem('Clean Emails', 'cleanEmails')
    .addItem('Remove Duplicate Emails', 'removeDuplicateEmails')
    .addToUi();
}
