// specific Sheet
function protectSheetWithExceptions() {
 // Replace 'YourSheetName' with the name of your specific sheet
 var sheetName = 'IncomeExpenses';
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
 // Protect the entire sheet
 var protection = sheet.protect().setDescription('Protected Sheet');
  
 // Ensure the current user is an editor before removing others
 var me = Session.getEffectiveUser();
 protection.addEditor(me);
 protection.removeEditors(protection.getEditors());
 
 // if you want to add range
 // Set unprotected cells
 var unprotectedRanges = [];
 var cell1 = sheet.getRange("C14"); // Example cell 1
 var cell2 = sheet.getRange("F12"); // Example cell 2
 unprotectedRanges.push(cell1);
 unprotectedRanges.push(cell2);
  
 // Apply the unprotected cells to the protection
 protection.setUnprotectedRanges(unprotectedRanges);
}
