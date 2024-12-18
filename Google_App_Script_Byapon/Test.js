function protectSheetWithExceptions6_ExpenceDB() {
 var sheetName = 'ExtendDB';
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
 // Protect the entire sheet
 var protection = sheet.protect().setDescription('Protected Sheet');
  
 // Ensure the current user is an editor before removing others
 var me = Session.getEffectiveUser();
 protection.addEditor(me);
 protection.removeEditors(protection.getEditors());
}