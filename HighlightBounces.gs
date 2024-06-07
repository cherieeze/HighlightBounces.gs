// Function to handle bounced emails
function handleBouncedEmails() {
  // Step 1: Get the bounced emails from the Google Sheets document
  var bouncedEmailsSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var bouncedEmailsRange = bouncedEmailsSheet.getDataRange();
  var bouncedEmailsData = bouncedEmailsRange.getValues();
  
  // Step 2: Get the original contact spreadsheet
  var originalContactSpreadsheetId = "19W6fsERfPohKofr_y8-ZAmUV75FmV_MdH6RzxOjv5rc"; // ID of the original contact spreadsheet
  var originalContactSpreadsheet = SpreadsheetApp.openById(originalContactSpreadsheetId);
  var sheets = originalContactSpreadsheet.getSheets();
  
  // Step 3: Loop through each sheet in the original contact spreadsheet
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var lastRow = sheet.getLastRow();
    var emailColumnValues = sheet.getRange("F2:F" + lastRow).getValues(); // Assuming email addresses are in column F
    
    // Step 4: Loop through bounced emails and identify bounced contacts
    for (var i = 0; i < bouncedEmailsData.length; i++) {
      var bouncedEmail = bouncedEmailsData[i][0]; // Assuming the email addresses are in the first column
      
      // Search for bounced email in the current sheet
      for (var j = 0; j < emailColumnValues.length; j++) {
        if (emailColumnValues[j][0] == bouncedEmail) {
          // Step 5: Highlight the row corresponding to the bounced contact
          sheet.getRange("A" + (j + 2) + ":Z" + (j + 2)).setBackground("red"); // Assuming contact information starts from row 2
          
          // // Step 6: Clear the content of the email cell for the bounced contact
          // sheet.getRange("F" + (j + 2)).clearContent(); // Assuming email addresses are in column F
          break;
        }
      }
    }
  }
}
