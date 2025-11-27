/**
 * ------------------------------------------------------------------
 * BACKEND SCRIPT (Code.gs)
 * ------------------------------------------------------------------
 * This script handles the communication between your Dashboard (Frontend)
 * and your Google Sheet (Database) + Gmail (Email Service).
 * * INSTRUCTIONS:
 * 1. Paste this into your Google Apps Script editor.
 * 2. Click "Deploy" > "New Deployment".
 * 3. Select type "Web App".
 * 4. Execute as: "Me".
 * 5. Who has access: "Anyone" (or "Anyone within organization").
 * 6. Copy the Web App URL if you were using it externally, otherwise just refresh your test deployment.
 */

// 1. SERVE THE DASHBOARD HTML
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Workload Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- NEW: AUTHORIZE FUNCTION ---
// Run this function ONCE manually from the editor to fix the permission error.
function authorize() {
  var email = Session.getActiveUser().getEmail();
  console.log("Authorizing Gmail for: " + email);
  // These lines force Google to ask for permissions
  GmailApp.getInboxThreads(0, 1); 
  SpreadsheetApp.getActiveSpreadsheet();
  return "Authorization Successful! You can now deploy the app.";
}

// 2. API: GET RECORDS (Fetch data from Sheets)
function getRecords(type) {
  var sheetName = type === 'user' ? "User_Workload_Data" : "System_Extract_Data";
  var sheet = getOrCreateSheet(sheetName);
  var data = sheet.getDataRange().getValues();
  
  // If sheet is empty or only has headers, return empty array
  if (data.length <= 1) return [];

  var headers = data[0];
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var key = toCamelCase(headers[j]);
      
      // Handle Date Objects from Sheet to prevent JSON errors
      if (row[j] instanceof Date) {
        // Format as YYYY-MM-DD for consistency
        try {
          // Adjust for timezone offset to prevent "yesterday" errors
          var d = row[j];
          var offset = d.getTimezoneOffset() * 60000;
          obj[key] = new Date(d.getTime() - offset).toISOString().split('T')[0];
        } catch(e) {
          obj[key] = String(row[j]);
        }
      } else {
        obj[key] = row[j];
      }
    }
    result.push(obj);
  }
  return result;
}

// 3. API: ADD RECORD (Save data to Sheets)
function addRecord(type, record) {
  try {
    var sheetName = type === 'user' ? "User_Workload_Data" : "System_Extract_Data";
    var sheet = getOrCreateSheet(sheetName);
    
    if (type === 'user') {
      // Define Columns for User/Supervisor Data
      // Added "Known Name" after Employee Name
      var headers = [
        "Date", 
        "User AD", 
        "Employee Name",
        "Known Name",  // <--- NEW COLUMN
        "Role", 
        "Work Status", 
        "Notes", 
        "Reject Accounts", 
        "Completed Accounts", 
        "Completed CIF Count", 
        "Total Accounts", 
        "Modifications", 
        "Uploaded Accounts", 
        "Timestamp"
      ];
      setupHeaders(sheet, headers);
      
      // Add the row
      sheet.appendRow([
        record.date,                    // Date selected in form
        record.userAD,
        record.employeeName,
        record.knownName || '',         // <--- SAVE KNOWN NAME
        record.role,
        record.workStatus || 'Full Day',
        record.notes || '',
        record.rejectAccounts,
        record.completedAccounts,
        record.completedCIFCount,
        record.totalAccounts,
        record.modifications || 0,
        record.uploadedAccounts || 0,
        new Date().toISOString()        // System Timestamp
      ]);
      
    } else {
      // Define Columns for System Extract Data
      var headers = [
        "Date", 
        "Carried Forward", 
        "Extracted ACs", 
        "Auto Rejects", 
        "User Rejects", 
        "Completed ACs", 
        "Completed CIFs", 
        "New Extracted ACs", 
        "Amendments Extracted ACs", 
        "Timestamp"
      ];
      setupHeaders(sheet, headers);
      
      // Add the row
      sheet.appendRow([
        record.date,
        record.carriedForward,
        record.extractedACs,
        record.autoRejects,
        record.userRejects,
        record.completedACs,
        record.completedCIFs,
        record.newExtractedACs,
        record.amendmentsExtractedACs,
        new Date().toISOString()
      ]);
    }
    
    return { status: "success" };
    
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

// 4. API: SEND EMAIL (Sends via Gmail)
function sendEmail(details) {
  try {
    if (!details.to) throw new Error("Recipient email is required.");
    
    // Configure Email Options
    var options = {
      cc: details.cc || ''
    };
    
    // Inject HTML body if provided (Crucial for modern look)
    if (details.htmlBody) {
      options.htmlBody = details.htmlBody;
    }

    // Send the email
    GmailApp.sendEmail(
      details.to, 
      details.subject, 
      details.body || 'Please view this email in a client that supports HTML.', 
      options
    );
    
    return { status: "success" };
    
  } catch (e) {
    console.error("Email Error: " + e.toString()); 
    throw new Error(e.toString());
  }
}

// --- HELPER FUNCTIONS ---

// Gets a sheet or creates it if it doesn't exist
function getOrCreateSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

// Sets up header row style if the sheet is empty
function setupHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    // Style: Bold text, light gray background, borders
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight("bold")
         .setBackground("#e0e0e0")
         .setBorder(true, true, true, true, true, true);
    // Freeze top row
    sheet.setFrozenRows(1);
  }
}

// Helper to convert "User AD" -> "userAD" for JSON
function toCamelCase(str) {
  return str.replace(/(?:^\w|[A-Z]|\b\w)/g, function(word, index) {
    return index === 0 ? word.toLowerCase() : word.toUpperCase();
  }).replace(/\s+/g, '');
}