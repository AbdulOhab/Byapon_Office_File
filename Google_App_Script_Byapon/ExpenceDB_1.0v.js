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