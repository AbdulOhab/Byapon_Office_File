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
  
 // Set unprotected cells
 var unprotectedRanges = [];

 //Income
 var cell_1 = sheet.getRange("C6"); 
 var cell_2 = sheet.getRange("C8"); 
 var cell_3 = sheet.getRange("C10"); 
 var cell_4 = sheet.getRange("C12"); 
 var cell_5 = sheet.getRange("C14"); 
 var cell_6 = sheet.getRange("C16"); 
 var cell_7 = sheet.getRange("F6"); 
 var cell_8 = sheet.getRange("F8"); 
 var cell_9 = sheet.getRange("F10"); 
 var cell_10 = sheet.getRange("F12"); 
 var cell_11 = sheet.getRange("F14"); 
 var cell_12 = sheet.getRange("F16"); 
 var cell_13 = sheet.getRange("C18:F18");

 //  
 var cell_14 = sheet.getRange("I6"); 
 var cell_15 = sheet.getRange("I8"); 
 var cell_16 = sheet.getRange("I10"); 
 var cell_17 = sheet.getRange("I12"); 
 var cell_18 = sheet.getRange("L6"); 
 var cell_19 = sheet.getRange("L8");
 var cell_20 = sheet.getRange("L10");
 var cell_21 = sheet.getRange("L12");
 var cell_22 = sheet.getRange("I14:L14");
 unprotectedRanges.push(cell_1);
 unprotectedRanges.push(cell_2);
 unprotectedRanges.push(cell_3);
 unprotectedRanges.push(cell_4);
 unprotectedRanges.push(cell_5);
 unprotectedRanges.push(cell_6);
 unprotectedRanges.push(cell_7);
 unprotectedRanges.push(cell_8);
 unprotectedRanges.push(cell_9);
 unprotectedRanges.push(cell_10);

 unprotectedRanges.push(cell_11);
 unprotectedRanges.push(cell_12);
 unprotectedRanges.push(cell_13);
 unprotectedRanges.push(cell_14);
 unprotectedRanges.push(cell_15);
 unprotectedRanges.push(cell_16);
 unprotectedRanges.push(cell_17);
 unprotectedRanges.push(cell_18);
 unprotectedRanges.push(cell_19);
 unprotectedRanges.push(cell_20);
 unprotectedRanges.push(cell_21);
 unprotectedRanges.push(cell_22);
  
 // Apply the unprotected cells to the protection
 protection.setUnprotectedRanges(unprotectedRanges);
}


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

  // Set unprotected cells
  var unprotectedRanges = [];

  // Income
  var cells = [
    "C6", "C8", "C10", "C12", "C14", "C18:F18",
    "F6", "F8", "F10", "F12", "F14", "F16",
    "I6", "I8", "I10", "I12",
    "L6", "L8", "L10", "L12",
    "I14:L14"
    ];

  cells.forEach(function(cell) {
    unprotectedRanges.push(sheet.getRange(cell));
  });

  // Apply the unprotected cells to the protection
  protection.setUnprotectedRanges(unprotectedRanges);
}
