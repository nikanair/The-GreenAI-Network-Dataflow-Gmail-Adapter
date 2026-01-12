/**
 * Exports email contents from a specified Gmail search query to the active Google Sheet.
 * It uses the Message ID to prevent duplicate entries on subsequent runs. Sowparnika Nair, GAIN
 */
function getEmailsToSheet() {
  const SEARCH_QUERY = "to:info@gain-ai.org";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  const threads = GmailApp.search(SEARCH_QUERY, 0, 500);
  
  let messages = [];
  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      messages.push(message);
    });
  });

  const lastRow = sheet.getLastRow();
  
  if (lastRow === 0 || sheet.getRange(1, 1).getValue() === "") {
    sheet.appendRow(["Message ID", "Date", "Sender", "Subject", "Body"]);
  }
  
  let existingIds = [];
  if (lastRow > 1) {
    existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  }
  
  const dataToAppend = [];
  
  messages.forEach(message => {
    const messageId = message.getId(); 
    
    if (!existingIds.includes(messageId)) {
      const date = message.getDate();
      const sender = message.getFrom();
      const subject = message.getSubject();
      const body = message.getPlainBody(); 
      
      dataToAppend.push([messageId, date, sender, subject, body]);
    }
  });

  if (dataToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, dataToAppend.length, dataToAppend[0].length).setValues(dataToAppend);
    SpreadsheetApp.getUi().alert(`Successfully imported ${dataToAppend.length} new email(s).`);
  } else {
    SpreadsheetApp.getUi().alert("No new emails found matching your query.");
  }
}
