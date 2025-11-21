/**
 * Exports email contents from a specified Gmail search query to the active Google Sheet.
 * It uses the Message ID to prevent duplicate entries on subsequent runs. Sowparnika Nair, GAIN
 */
function getEmailsToSheet() {
  // ----------------------------------------------------
  // 1. CONFIGURATION: Customize your search query here.
  // This example searches for all emails labeled "Invoices".
  // Replace 'label:Invoices' with your desired query (e.g., 'from:sales@company.com subject:Order').
  const SEARCH_QUERY = "to:info@gain-ai.org";
  // ----------------------------------------------------

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Use GmailApp.search() to find all threads matching your query.
  // The numbers (0, 100) are for starting index and max results. Max is 500.
  const threads = GmailApp.search(SEARCH_QUERY, 0, 500);
  
  let messages = [];
  threads.forEach(thread => {
    // Get all individual messages within each thread.
    thread.getMessages().forEach(message => {
      messages.push(message);
    });
  });

  const lastRow = sheet.getLastRow();
  
  // 2. Initialize Headers if the sheet is empty
  if (lastRow === 0 || sheet.getRange(1, 1).getValue() === "") {
    sheet.appendRow(["Message ID", "Date", "Sender", "Subject", "Body"]);
  }
  
  // 3. Get existing IDs to prevent duplicates
  let existingIds = [];
  if (lastRow > 1) {
    // Column A is where we store the unique Message ID
    existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  }
  
  const dataToAppend = [];
  
  // 4. Loop through messages and prepare the data
  messages.forEach(message => {
    const messageId = message.getId(); 
    
    // Check if the message ID is NOT already in the sheet
    if (!existingIds.includes(messageId)) {
      const date = message.getDate();
      const sender = message.getFrom();
      const subject = message.getSubject();
      const body = message.getPlainBody(); // Use getPlainBody() for clean text
      
      dataToAppend.push([messageId, date, sender, subject, body]);
    }
  });

  // 5. Append all new data to the sheet at once for efficiency
  if (dataToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, dataToAppend.length, dataToAppend[0].length).setValues(dataToAppend);
    SpreadsheetApp.getUi().alert(`Successfully imported ${dataToAppend.length} new email(s).`);
  } else {
    SpreadsheetApp.getUi().alert("No new emails found matching your query.");
  }
}
