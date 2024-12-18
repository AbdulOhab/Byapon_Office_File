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
  incomeDB.getRange(blankRow,2).setValue(myActiveSheet.getRange("C6").getValue());
  incomeDB.getRange(blankRow,3).setValue(myActiveSheet.getRange("C8").getValue());
  incomeDB.getRange(blankRow,4).setValue(myActiveSheet.getRange("C10").getValue());
  incomeDB.getRange(blankRow,5).setValue(myActiveSheet.getRange("C12").getValue());
  incomeDB.getRange(blankRow,6).setValue(myActiveSheet.getRange("C14").getValue());
  incomeDB.getRange(blankRow,7).setValue(myActiveSheet.getRange("C16").getValue());
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
