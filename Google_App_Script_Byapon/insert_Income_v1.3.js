// -------------------------------------------- this is for Income Form

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
 myActiveSheet.getRange("C14").setBackground("#FFFFFF"); //Pourpus
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
 if(myActiveSheet.getRange("C10").isBlank() == true )
 { 
  uI.alert("Please Enter Income Source");
  myActiveSheet.getRange("C10").setBackground("#FF0000");
  return false;
  }

 //validateion Money Receipt
 if(myActiveSheet.getRange("C16").isBlank() == true )
 {
  uI.alert("Please Enter Money Receipt Details");
  myActiveSheet.getRange("C16").setBackground("#FF0000");
  return false;
 }

 //validateion Selling Elements
 if(myActiveSheet.getRange("F6").isBlank() == true )
 {
  uI.alert("Please Enter Elements Type details");
  myActiveSheet.getRange("F6").setBackground("#FF0000");
  return false;
 }

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

  incomeDB.getRange(blankRow, 1).setValue(newEmployeeID); // ID
  incomeDB.getRange(blankRow,2).setValue(myActiveSheet.getRange("C6").getValue()); // Date*
  incomeDB.getRange(blankRow,3).setValue(myActiveSheet.getRange("C8").getValue()); // Executor
  incomeDB.getRange(blankRow,4).setValue(myActiveSheet.getRange("C10").getValue()); //Income_Source*
  incomeDB.getRange(blankRow,5).setValue(myActiveSheet.getRange("C12").getValue()); //Selling type
  incomeDB.getRange(blankRow,6).setValue(myActiveSheet.getRange("C14").getValue()); //Pourpus
  incomeDB.getRange(blankRow,7).setValue(myActiveSheet.getRange("C16").getValue()); //Money Receipt*

  
  incomeDB.getRange(blankRow,8).setValue(myActiveSheet.getRange("F6").getValue()); //Elements*
  incomeDB.getRange(blankRow,9).setValue(myActiveSheet.getRange("F8").getValue()); //Issue
  incomeDB.getRange(blankRow,10).setValue(myActiveSheet.getRange("F10").getValue()); //Quantity
  incomeDB.getRange(blankRow,11).setValue(myActiveSheet.getRange("F12").getValue()); //Taka*
  incomeDB.getRange(blankRow,12).setValue(myActiveSheet.getRange("F14").getValue()); //Payment Method
  incomeDB.getRange(blankRow,13).setValue(myActiveSheet.getRange("F16").getValue()); //Discount
  incomeDB.getRange(blankRow,14).setValue(myActiveSheet.getRange("C18:F18").getValue()); //Comment
  



  // Code to update the date and time
  var currentDate = new Date(); // Get current date and time
  var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"); // Format the date as yyyy-MM-dd HH:mm
  incomeDB.getRange(blankRow, 15).setValue(formattedDate); // Set the formatted date to the desired cell 

  // Submitted by who
  incomeDB.getRange(blankRow,16).setValue(Session.getActiveUser().getEmail());

  uI2.alert(' "Submit Successfully" ' + myActiveSheet.getRange("C6").getValue() + '""' );

  myActiveSheet.getRange("C6").clearContent();
  myActiveSheet.getRange("C8").clearContent();
  myActiveSheet.getRange("C10").clearContent();
  myActiveSheet.getRange("C12").clearContent();
  myActiveSheet.getRange("C14").clearContent();
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
  myActiveSheet.getRange("C14").clearContent();
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

  uI.alert("Clear Successfully");

}
