// This is for Income Form

function isNumeric(n) {
  return !isNaN(parseFloat(n)) && isFinite(n);
}
// function to valid to entry made by user 
function validEntry(){
 var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 var myActiveSheet =  mySpreadsheet.getSheetByName("IncomeExpenses"); // this working spreadsheet

 var uI = SpreadsheetApp.getUi(); // to create the instance of the user interface ot show the alert.

// Input field default color
 myActiveSheet.getRange("C6").setBackground("#FFFFFF"); // Date*
 myActiveSheet.getRange("C8").setBackground("#FFFFFF"); // Executor
 myActiveSheet.getRange("C10").setBackground("#FFFFFF"); //Income_Source*
 myActiveSheet.getRange("C12").setBackground("#FFFFFF"); //Type
 myActiveSheet.getRange("C16").setBackground("#FFFFFF"); //Money Receipt*
 myActiveSheet.getRange("C18:F18").setBackground("#FFFFFF");  //Comment

 myActiveSheet.getRange("F6").setBackground("#FFFFFF");  //Elements*
 myActiveSheet.getRange("F8").setBackground("#FFFFFF");  //Issue
 myActiveSheet.getRange("F10").setBackground("#FFFFFF"); //Quantity
 myActiveSheet.getRange("F12").setBackground("#FFFFFF"); //Taka
 myActiveSheet.getRange("F14").setBackground("#FFFFFF"); //Payment Method
 myActiveSheet.getRange("F16").setBackground("#FFFFFF"); //Discount

 //validateion Date
 if(myActiveSheet.getRange("C6").isBlank() == true )
 {
  uI.alert("Please Enter Date");
  myActiveSheet.getRange("C6").setBackground("#FF0000");
  return false;
  }

 //validateion Income Source
//  if(myActiveSheet.getRange("C10").isBlank() == true )
//  { 
//   uI.alert("Please Enter Income Source");
//   myActiveSheet.getRange("C10").setBackground("#FF0000");
//   return false;
//   }

 //validateion Money Receipt
 if(myActiveSheet.getRange("C16").isBlank() == true )
 {
  uI.alert("Please Enter Money Receipt Details");
  myActiveSheet.getRange("C16").setBackground("#FF0000");
  return false;
 }

 //validateion Selling Elements
//  if(myActiveSheet.getRange("F6").isBlank() == true )
//  {
//   uI.alert("Please Enter Elements Type details");
//   myActiveSheet.getRange("F6").setBackground("#FF0000");
//   return false;
//  }

  //validateion selling Quantity
//  var value = myActiveSheet.getRange("F10").getValue();
//  if(myActiveSheet.getRange("F10").isBlank() || !isNumeric(value)) {
//   uI.alert("Please Enter a Numeric Quantity");
//   myActiveSheet.getRange("F10").setBackground("#FF0000");
//   return false;
//   }
  
  return true;
}

// Function to generate unique employee IDs
function generateId() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("IncomeDB");
  var lastRow = sheet.getLastRow();
  var lastId = sheet.getRange(lastRow, 1).getValue();
  var newId = lastId ? parseInt(lastId) + 1 : 1;
  return newId;
}

//Function to submit the data to Database
function submitDeta(){
  //declear a variable and set the reference of active google sheet
 var mySpreadsheet2 = SpreadsheetApp.getActiveSpreadsheet();
 var myActiveSheet =  mySpreadsheet2.getSheetByName("IncomeExpenses"); 
 var incomeDB =  mySpreadsheet2.getSheetByName("IncomeDB");

 var uI2 = SpreadsheetApp.getUi(); //show the alert.
 
 var response = uI2.alert("Submit","Do you want to Submit?", uI2.ButtonSet.YES_NO); 

 // checkeing usder response

 if(response == uI2.Button.NO){
  return;
 }

 if(validEntry() == true){

  var blankRow = incomeDB.getLastRow() + 1;
  var newEmployeeID = generateId(); // Generate new employee ID

  incomeDB.getRange(blankRow,1).setValue(newEmployeeID); // ID
  incomeDB.getRange(blankRow,2).setValue(myActiveSheet.getRange("C6").getValue()); // Date*
  incomeDB.getRange(blankRow,3).setValue(myActiveSheet.getRange("C8").getValue()); // Executor
  incomeDB.getRange(blankRow,4).setValue(myActiveSheet.getRange("C10").getValue()); //Income_Source*
  incomeDB.getRange(blankRow,5).setValue(myActiveSheet.getRange("C12").getValue()); //Selling type
  incomeDB.getRange(blankRow,6).setValue(myActiveSheet.getRange("C16").getValue()); //Money Receipt*

  
  incomeDB.getRange(blankRow,7).setValue(myActiveSheet.getRange("F6").getValue()); //Elements*
  incomeDB.getRange(blankRow,8).setValue(myActiveSheet.getRange("F8").getValue()); //Issue
  incomeDB.getRange(blankRow,9).setValue(myActiveSheet.getRange("F10").getValue()); //Quantity
  incomeDB.getRange(blankRow,10).setValue(myActiveSheet.getRange("F12").getValue()); //Taka*
  incomeDB.getRange(blankRow,11).setValue(myActiveSheet.getRange("F14").getValue()); //Payment Method
  incomeDB.getRange(blankRow,12).setValue(myActiveSheet.getRange("F16").getValue()); //Discount
  incomeDB.getRange(blankRow,13).setValue(myActiveSheet.getRange("C18:F18").getValue()); //Comment
  



  // Code to update the date and time
  var currentDate = new Date();
  var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"); // Format the date as yyyy-MM-dd HH:mm
  incomeDB.getRange(blankRow, 14).setValue(formattedDate); 

  // Submitted by who
  incomeDB.getRange(blankRow,15).setValue(Session.getActiveUser().getEmail());

  uI2.alert("Submit Successfully");
 
  //Active Cell Data Input as gerbage Deta 
  var activeCellValue = myActiveSheet.getActiveCell().getValue();
  incomeDB.getRange(blankRow, 17).setValue(activeCellValue);

  myActiveSheet.getRange("C6").clearContent();
  myActiveSheet.getRange("C8").clearContent();
  myActiveSheet.getRange("C10").clearContent();
  myActiveSheet.getRange("C12").clearContent();
  myActiveSheet.getRange("C16").clearContent();

  myActiveSheet.getRange("F6").clearContent();
  myActiveSheet.getRange("F8").clearContent();
  myActiveSheet.getRange("F10").clearContent();
  myActiveSheet.getRange("F12").clearContent();
  myActiveSheet.getRange("F14").clearContent();
  myActiveSheet.getRange("F16").clearContent();

  // myActiveSheet.getRange("C18:F18").clearContent();
  myActiveSheet.getRange("C18").clearContent();
  myActiveSheet.getRange("D18").clearContent();
  myActiveSheet.getRange("E18").clearContent();
  myActiveSheet.getRange("F18").clearContent();
  
  // Clear the content of the selected cell
  var activeCell = myActiveSheet.getActiveCell();
  if (activeCell) {
    activeCell.clearContent();
  }

  //Assign new date in c6 cell
  myActiveSheet.getRange("C6").setValue(new Date());

  myActiveSheet.getRange("C6").setBackground("#FFFFFF"); // Date*
  myActiveSheet.getRange("C10").setBackground("#FFFFFF"); //Income_Source*
  myActiveSheet.getRange("C16").setBackground("#FFFFFF"); //Money Receipt*
  myActiveSheet.getRange("F6").setBackground("#FFFFFF");  //Elements*
  }
}


//Function to submit the data to Database
function clearData(){
  //declear a variable and set the reference of active google sheet
  var mySpreadsheet2 = SpreadsheetApp.getActiveSpreadsheet();
  var myActiveSheet =  mySpreadsheet2.getSheetByName("IncomeExpenses"); 
  
  var uI = SpreadsheetApp.getUi();
  myActiveSheet.getRange("C6").clearContent();
  myActiveSheet.getRange("C8").clearContent();
  myActiveSheet.getRange("C10").clearContent();
  myActiveSheet.getRange("C12").clearContent();
  myActiveSheet.getRange("C16").clearContent();

  
  myActiveSheet.getRange("F6").clearContent();
  myActiveSheet.getRange("F8").clearContent();
  myActiveSheet.getRange("F10").clearContent();
  myActiveSheet.getRange("F12").clearContent();
  myActiveSheet.getRange("F14").clearContent();
  myActiveSheet.getRange("F16").clearContent();

  // myActiveSheet.getRange("C18:F18").clearContent();
  myActiveSheet.getRange("C18").clearContent();
  myActiveSheet.getRange("D18").clearContent();
  myActiveSheet.getRange("E18").clearContent();
  myActiveSheet.getRange("F18").clearContent();

    // Clear the content of the selected cell
  var activeCell = myActiveSheet.getActiveCell();
  if (activeCell) {
    activeCell.clearContent();
  }

  //Assign new date in c6 cell
  myActiveSheet.getRange("C6").setValue(new Date());

  myActiveSheet.getRange("C6").setBackground("#FFFFFF"); // Date*
  myActiveSheet.getRange("C10").setBackground("#FFFFFF"); //Income_Source*
  myActiveSheet.getRange("C16").setBackground("#FFFFFF"); //Money Receipt*
  myActiveSheet.getRange("F6").setBackground("#FFFFFF");  //Elements*

  uI.alert("Clear Successfully");

}

// This is for Expence form.
// function to valid to entry made by user 
function validEntry2(){
 var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 var myActiveSheet =  mySpreadsheet.getSheetByName("IncomeExpenses"); // this working spreadsheet

 var uI = SpreadsheetApp.getUi(); // to create the instance of the user interface ot show the alert.

// Input field default color
 myActiveSheet.getRange("I6").setBackground("#FFFFFF"); // Date*
 myActiveSheet.getRange("I8").setBackground("#FFFFFF"); // Executor
 myActiveSheet.getRange("I10").setBackground("#FFFFFF"); //Permission
 myActiveSheet.getRange("I12").setBackground("#FFFFFF"); //Vouture

 myActiveSheet.getRange("L6").setBackground("#FFFFFF");  //Spending*
 myActiveSheet.getRange("L8").setBackground("#FFFFFF");  //Details
 myActiveSheet.getRange("L10").setBackground("#FFFFFF"); //Medium
 myActiveSheet.getRange("L12").setBackground("#FFFFFF"); //Quantity 
 myActiveSheet.getRange("I14:L14").setBackground("#FFFFFF"); //Comment

 //validateion Date
 if(myActiveSheet.getRange("I6").isBlank() == true )
 {
  uI.alert("Please Enter Date");
  myActiveSheet.getRange("I6").setBackground("#FF0000");
  return false;
  }

 //validateion Spending Money
 if(myActiveSheet.getRange("L6").isBlank() == true )
 { 
  uI.alert("Please Enter Sector About Spending Money");
  myActiveSheet.getRange("L6").setBackground("#FF0000");
  return false;
  }

 //validateion Details
 if(myActiveSheet.getRange("L8").isBlank() == true )
 {
  uI.alert("Please Enter Details About Spending Money");
  myActiveSheet.getRange("L8").setBackground("#FF0000");
  return false;
 }

 //validateion Quantity 

 if(myActiveSheet.getRange("L12").isBlank() == true )
 {
  uI.alert("Please Enter Quantity of money");
  myActiveSheet.getRange("L12").setBackground("#FF0000");
  return false;
 }
  
  return true;
}

// Function to generate unique employee IDs
function generateId2() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ExpenceDB");
  var lastRow = sheet.getLastRow();
  var lastId = sheet.getRange(lastRow, 1).getValue();
  var newId = lastId ? parseInt(lastId) + 1 : 1;
  return newId;
}

//Function to submit the data to Database
function submitDeta2(){
  //declear a variable and set the reference of active google sheet
 var mySpreadsheet2 = SpreadsheetApp.getActiveSpreadsheet();
 var myActiveSheet =  mySpreadsheet2.getSheetByName("IncomeExpenses"); 
 var incomeDB =  mySpreadsheet2.getSheetByName("ExpenceDB");

 var uI2 = SpreadsheetApp.getUi(); //show the alert.
 
 var response = uI2.alert("Submit","Do you want to Submit?", uI2.ButtonSet.YES_NO); 

 // checkeing usder response

 if(response == uI2.Button.NO){
  return;
 }

 if(validEntry2() == true){

  var blankRow = incomeDB.getLastRow() + 1;
  var newEmployeeID = generateId2(); // Generate new employee ID

  incomeDB.getRange(blankRow, 1).setValue(newEmployeeID); // ID
  incomeDB.getRange(blankRow,2).setValue(myActiveSheet.getRange("I6").getValue()); // Date*
  incomeDB.getRange(blankRow,3).setValue(myActiveSheet.getRange("I8").getValue()); // Executor
  incomeDB.getRange(blankRow,4).setValue(myActiveSheet.getRange("I10").getValue()); //Permission
  incomeDB.getRange(blankRow,5).setValue(myActiveSheet.getRange("I12").getValue()); //Vouture


  incomeDB.getRange(blankRow,6).setValue(myActiveSheet.getRange("L6").getValue()); //Spending*
  incomeDB.getRange(blankRow,7).setValue(myActiveSheet.getRange("L8").getValue()); //Details
  incomeDB.getRange(blankRow,8).setValue(myActiveSheet.getRange("L10").getValue()); //Medium
  incomeDB.getRange(blankRow,9).setValue(myActiveSheet.getRange("L12").getValue()); //Quantity 
  incomeDB.getRange(blankRow,10).setValue(myActiveSheet.getRange("I14:L14").getValue()); //Comment


  
  // Code to update the date and time
  var currentDate = new Date(); // Get current date and time
  var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"); // Format the date as yyyy-MM-dd HH:mm
  // Set the formatted date to the desired cell 
  incomeDB.getRange(blankRow, 11).setValue(formattedDate); 

  // Submitted by who
  incomeDB.getRange(blankRow,12).setValue(Session.getActiveUser().getEmail());

  uI2.alert(' "Submit Successfully" ' + myActiveSheet.getRange("C6").getValue() + '""' );

  //Active Cell Data Input as gerbage Deta 
  var activeCellValue = myActiveSheet.getActiveCell().getValue();
  incomeDB.getRange(blankRow, 13).setValue(activeCellValue);


  myActiveSheet.getRange("I6").clearContent();
  myActiveSheet.getRange("I8").clearContent();
  myActiveSheet.getRange("I10").clearContent();
  myActiveSheet.getRange("I12").clearContent();

  myActiveSheet.getRange("L6").clearContent();
  myActiveSheet.getRange("L8").clearContent();
  myActiveSheet.getRange("L10").clearContent();
  myActiveSheet.getRange("L12").clearContent();

  //This is for Comment
  myActiveSheet.getRange("I14").clearContent();
  myActiveSheet.getRange("J14").clearContent();
  myActiveSheet.getRange("K14").clearContent();
  myActiveSheet.getRange("L14").clearContent();
  
  // Clear the content of the selected cell
  var activeCell = myActiveSheet.getActiveCell();
  if (activeCell) {
    activeCell.clearContent();
  }

  //Assign new date in c6 cell
  myActiveSheet.getRange("I6").setValue(new Date());

  myActiveSheet.getRange("I6").setBackground("#FFFFFF"); // Date*
  myActiveSheet.getRange("L6").setBackground("#FFFFFF"); //Spending*
  myActiveSheet.getRange("L8").setBackground("#FFFFFF"); //Details*
  myActiveSheet.getRange("L12").setBackground("#FFFFFF");  //Quantity*
  }
}


//Function to submit the data to Database
function clearData2(){
  //declear a variable and set the reference of active google sheet
  var mySpreadsheet2 = SpreadsheetApp.getActiveSpreadsheet();
  var myActiveSheet =  mySpreadsheet2.getSheetByName("IncomeExpenses"); 
  
  var uI = SpreadsheetApp.getUi();
  myActiveSheet.getRange("I6").clearContent();
  myActiveSheet.getRange("I8").clearContent();
  myActiveSheet.getRange("I10").clearContent();
  myActiveSheet.getRange("I12").clearContent();

  myActiveSheet.getRange("L6").clearContent();
  myActiveSheet.getRange("L8").clearContent();
  myActiveSheet.getRange("L10").clearContent();
  myActiveSheet.getRange("L12").clearContent();

  //This is for Comment
  myActiveSheet.getRange("I14").clearContent();
  myActiveSheet.getRange("J14").clearContent();
  myActiveSheet.getRange("K14").clearContent();
  myActiveSheet.getRange("L14").clearContent();

    // Clear the content of the selected cell
  var activeCell = myActiveSheet.getActiveCell();
  if (activeCell) {
    activeCell.clearContent();
  }

  //Assign new date in c6 cell
  myActiveSheet.getRange("I6").setValue(new Date());

  myActiveSheet.getRange("I6").setBackground("#FFFFFF"); // Date*
  myActiveSheet.getRange("L6").setBackground("#FFFFFF"); //Spending*
  myActiveSheet.getRange("L8").setBackground("#FFFFFF"); //Details*
  myActiveSheet.getRange("L12").setBackground("#FFFFFF");  //Quantity*

  uI.alert("Clear Successfully");

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

 //Income
 var cell_1 = sheet.getRange("C6"); 
 var cell_2 = sheet.getRange("C8"); 
 var cell_3 = sheet.getRange("C10"); 
 var cell_4 = sheet.getRange("C12"); 
 var cell_6 = sheet.getRange("C16"); 
 var cell_7 = sheet.getRange("F6"); 
 var cell_8 = sheet.getRange("F8"); 
 var cell_9 = sheet.getRange("F10"); 
 var cell_10 = sheet.getRange("F12"); 
 var cell_11 = sheet.getRange("F14"); 
 var cell_12 = sheet.getRange("F16"); 
 var cell_13 = sheet.getRange("C18:F18");

 //Expence
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

//Marketing Form

// function to valid to entry made by user 
function validEntry3(){
 var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 var myActiveSheet =  mySpreadsheet.getSheetByName("Marketing"); // this working spreadsheet
 
 var uI = SpreadsheetApp.getUi(); //Show the alert.
 
 // Input field default color
 myActiveSheet.getRange("C4").setBackground("#FFFFFF"); // Date*
 myActiveSheet.getRange("C6").setBackground("#FFFFFF"); // Marketer*
 myActiveSheet.getRange("C8").setBackground("#FFFFFF"); // Marketing Area
 myActiveSheet.getRange("C10").setBackground("#FFFFFF"); // Type*
 myActiveSheet.getRange("C12").setBackground("#FFFFFF"); // Librery Name
 myActiveSheet.getRange("C14").setBackground("#FFFFFF"); // Adderss
 myActiveSheet.getRange("C17").setBackground("#FFFFFF"); // provided Issue number
 myActiveSheet.getRange("C19").setBackground("#FFFFFF");  // Circulation Quantity
 myActiveSheet.getRange("C22").setBackground("#FFFFFF");  // Issue Back
 myActiveSheet.getRange("C24").setBackground("#FFFFFF");  // Issue Back Quantity
 
 myActiveSheet.getRange("F6").setBackground("#FFFFFF");  // Withdrawal Issue Number
 myActiveSheet.getRange("F8").setBackground("#FFFFFF");  // Withdrawal Taka
 myActiveSheet.getRange("F10").setBackground("#FFFFFF"); // Due
 myActiveSheet.getRange("F14").setBackground("#FFFFFF"); // Materials 1
 myActiveSheet.getRange("F17").setBackground("#FFFFFF"); // Materials 2
 myActiveSheet.getRange("F19").setBackground("#FFFFFF"); // Materials 3
 myActiveSheet.getRange("F22").setBackground("#FFFFFF"); // Details 
 myActiveSheet.getRange("F24").setBackground("#FFFFFF"); // Comment

 //validateion Date
 if(myActiveSheet.getRange("C4").isBlank() == true )
 {
  uI.alert("Please Enter Date");
  myActiveSheet.getRange("C4").setBackground("#FF0000");
  return false;
  }

 //validateion Marketer Name
 if(myActiveSheet.getRange("C6").isBlank() == true )
 { 
  uI.alert("Please Enter Marketer Name");
  myActiveSheet.getRange("C6").setBackground("#FF0000");
  return false;
  }

 //validateion Marckting Type
 if(myActiveSheet.getRange("C10").isBlank() == true )
 {
  uI.alert("Please Enter Marckting Type");
  myActiveSheet.getRange("C10").setBackground("#FF0000");
  return false;
 }

 //validateion Librery Name
 if(myActiveSheet.getRange("C12").isBlank() == true )
 {
  uI.alert("Please Enter Visiting Librery Name");
  myActiveSheet.getRange("C12").setBackground("#FF0000");
  return false;
 }

  return true;
}

// Function to generate unique employee IDs
function generateId() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MarketingDB");
  var lastRow = sheet.getLastRow();
  var lastId = sheet.getRange(lastRow, 1).getValue();
  var newId = lastId ? parseInt(lastId) + 1 : 1;
  return newId;
}

//Function to submit the data to Database
function submitDeta3(){
  //declear a variable and set the reference of active google sheet
 var mySpreadsheet2 = SpreadsheetApp.getActiveSpreadsheet();
 var myActiveSheet =  mySpreadsheet2.getSheetByName("Marketing"); 
 var incomeDB =  mySpreadsheet2.getSheetByName("MarketingDB");

 var uI2 = SpreadsheetApp.getUi(); //show the alert.
 
 var response = uI2.alert("Submit","Do you want to Submit?", uI2.ButtonSet.YES_NO); 

 // checkeing usder response

 if(response == uI2.Button.NO){
  return;
 }

 if(validEntry3() == true){

  var blankRow = incomeDB.getLastRow() + 1;
  var newEmployeeID = generateId(); // Generate new employee ID

  incomeDB.getRange(blankRow, 1).setValue(newEmployeeID); // ID
  incomeDB.getRange(blankRow,2).setValue(myActiveSheet.getRange("C4").getValue());
  incomeDB.getRange(blankRow,3).setValue(myActiveSheet.getRange("C6").getValue());
  incomeDB.getRange(blankRow,4).setValue(myActiveSheet.getRange("C8").getValue());
  incomeDB.getRange(blankRow,5).setValue(myActiveSheet.getRange("C10").getValue());
  incomeDB.getRange(blankRow,6).setValue(myActiveSheet.getRange("C12").getValue());
  incomeDB.getRange(blankRow,7).setValue(myActiveSheet.getRange("C14").getValue());
  incomeDB.getRange(blankRow,8).setValue(myActiveSheet.getRange("C17").getValue());
  incomeDB.getRange(blankRow,9).setValue(myActiveSheet.getRange("C19").getValue());
  incomeDB.getRange(blankRow,10).setValue(myActiveSheet.getRange("C22").getValue()); 
  incomeDB.getRange(blankRow,11).setValue(myActiveSheet.getRange("C24").getValue());


  incomeDB.getRange(blankRow,12).setValue(myActiveSheet.getRange("F6").getValue());
  incomeDB.getRange(blankRow,13).setValue(myActiveSheet.getRange("F8").getValue());
  incomeDB.getRange(blankRow,14).setValue(myActiveSheet.getRange("F10").getValue());
  incomeDB.getRange(blankRow,15).setValue(myActiveSheet.getRange("F14").getValue());
  incomeDB.getRange(blankRow,16).setValue(myActiveSheet.getRange("F17").getValue());
  incomeDB.getRange(blankRow,17).setValue(myActiveSheet.getRange("F19").getValue());
  incomeDB.getRange(blankRow,18).setValue(myActiveSheet.getRange("F22").getValue());
  incomeDB.getRange(blankRow,19).setValue(myActiveSheet.getRange("F24").getValue());
  
  // Code to update the date and time
  var currentDate = new Date(); // Get current date and time
  var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"); // Format the date as yyyy-MM-dd HH:mm
  // Set the formatted date to the desired cell 
  incomeDB.getRange(blankRow, 20).setValue(formattedDate); 

  // Submitted by who
  incomeDB.getRange(blankRow,21).setValue(Session.getActiveUser().getEmail());

  uI2.alert(' "Submit Successfully" ' + myActiveSheet.getRange("C4").getValue() + '""' );
 
  //Active Cell Data Input as gerbage Deta 
  var activeCellValue = myActiveSheet.getActiveCell().getValue();
  incomeDB.getRange(blankRow, 22).setValue(activeCellValue);

  myActiveSheet.getRange("C4").clearContent();
  myActiveSheet.getRange("C6").clearContent();
  myActiveSheet.getRange("C8").clearContent();
  myActiveSheet.getRange("C10").clearContent();
  myActiveSheet.getRange("C12").clearContent();
  myActiveSheet.getRange("C14").clearContent();
  myActiveSheet.getRange("C17").clearContent();
  myActiveSheet.getRange("C19").clearContent();
  myActiveSheet.getRange("C22").clearContent();
  myActiveSheet.getRange("C24").clearContent();

  myActiveSheet.getRange("F6").clearContent();
  myActiveSheet.getRange("F8").clearContent();
  myActiveSheet.getRange("F10").clearContent();
  myActiveSheet.getRange("F14").clearContent();
  myActiveSheet.getRange("F17").clearContent();
  myActiveSheet.getRange("F19").clearContent();
  myActiveSheet.getRange("F22").clearContent();
  myActiveSheet.getRange("F24").clearContent();
  
  // Clear the content of the selected cell
  var activeCell = myActiveSheet.getActiveCell();
  if (activeCell) {
    activeCell.clearContent();
  }

  //Assign new date in c6 cell
  myActiveSheet.getRange("C4").setValue(new Date());

  myActiveSheet.getRange("C4").setBackground("#FFFFFF");
  myActiveSheet.getRange("C6").setBackground("#FFFFFF"); 
  myActiveSheet.getRange("C10").setBackground("#FFFFFF");
  myActiveSheet.getRange("C12").setBackground("#FFFFFF");
  }
}


//Function to submit the data to Database
function clearData3(){
  //declear a variable and set the reference of active google sheet
  var mySpreadsheet2 = SpreadsheetApp.getActiveSpreadsheet();
  var myActiveSheet =  mySpreadsheet2.getSheetByName("Marketing"); 
  
  var uI = SpreadsheetApp.getUi();
  myActiveSheet.getRange("C4").clearContent();
  myActiveSheet.getRange("C6").clearContent();
  myActiveSheet.getRange("C8").clearContent();
  myActiveSheet.getRange("C10").clearContent();
  myActiveSheet.getRange("C12").clearContent();
  myActiveSheet.getRange("C14").clearContent();
  myActiveSheet.getRange("C17").clearContent();
  myActiveSheet.getRange("C19").clearContent();
  myActiveSheet.getRange("C22").clearContent();
  myActiveSheet.getRange("C24").clearContent();

  myActiveSheet.getRange("F6").clearContent();
  myActiveSheet.getRange("F8").clearContent();
  myActiveSheet.getRange("F10").clearContent();
  myActiveSheet.getRange("F14").clearContent();
  myActiveSheet.getRange("F17").clearContent();
  myActiveSheet.getRange("F19").clearContent();
  myActiveSheet.getRange("F22").clearContent();
  myActiveSheet.getRange("F24").clearContent();

    // Clear the content of the selected cell
  var activeCell = myActiveSheet.getActiveCell();
  if (activeCell) {
    activeCell.clearContent();
  }

  //Assign new date in C4 cell
  myActiveSheet.getRange("C4").setValue(new Date());

  myActiveSheet.getRange("C4").setBackground("#FFFFFF");
  myActiveSheet.getRange("C6").setBackground("#FFFFFF"); 
  myActiveSheet.getRange("C10").setBackground("#FFFFFF");
  myActiveSheet.getRange("C12").setBackground("#FFFFFF");

  uI.alert("Clear Successfully");

}





