//Expence Form

// Function to validate entry made by user
function validEntry2() {
  var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var myActiveSheet = mySpreadsheet.getSheetByName("IncomeExpenses"); // this working spreadsheet
  var uI = SpreadsheetApp.getUi(); // to create the instance of the user interface to show the alert.

  // Input field default color
  var ranges = ["I6", "I8", "I10", "I12", "L6", "L8", "L10", "L12", "I14:L14"]
  ranges.forEach(function(range) {
    myActiveSheet.getRange(range).setBackground("#FFFFFF");
  });

 //validateion Date
 if(myActiveSheet.getRange("I6").isBlank() == true ){
  uI.alert("Please Enter Date");
  myActiveSheet.getRange("I6").setBackground("#FF0000");
  return false;
 }

 //validateion Spending Money
 if(myActiveSheet.getRange("L6").isBlank() == true ){ 
  uI.alert("Please Enter Sector About Spending Money");
  myActiveSheet.getRange("L6").setBackground("#FF0000");
  return false;
 }

 //validateion Details
 if(myActiveSheet.getRange("L8").isBlank() == true ){
  uI.alert("Please Enter Details About Spending Money");
  myActiveSheet.getRange("L8").setBackground("#FF0000");
  return false;
 }

 //validateion Quantity 
 if(myActiveSheet.getRange("L12").isBlank() == true ){
  uI.alert("Please Enter Quantity of money");
  myActiveSheet.getRange("L12").setBackground("#FF0000");
  return false;
 }

 return true;
}

// Function to submit the data to Database
function submitData2() {
  var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var myActiveSheet = mySpreadsheet.getSheetByName("IncomeExpenses");
  var incomeDB = mySpreadsheet.getSheetByName("ExpenceDB");
  var uI = SpreadsheetApp.getUi(); // show the alert.

  var response = uI.alert("Submit", "Do you want to Submit?", uI.ButtonSet.YES_NO);

  // checking user response
  if (response == uI.Button.NO) {
    return;
  }

  if (validEntry2()) {
    var blankRow = incomeDB.getLastRow() + 1;

    // Prepare a 2D array to hold the data
    var data = [
      [ 
        myActiveSheet.getRange("I6").getValue(), 
        myActiveSheet.getRange("I8").getValue(), 
        myActiveSheet.getRange("I10").getValue(),
        myActiveSheet.getRange("I12").getValue(),
        myActiveSheet.getRange("L6").getValue(),
        myActiveSheet.getRange("L8").getValue(),
        myActiveSheet.getRange("L10").getValue(),
        myActiveSheet.getRange("L12").getValue(),
        myActiveSheet.getRange("I14:L14").getValue()
      ],
    ];

    // Write the data to the spreadsheet in one operation
    incomeDB.getRange(blankRow, 1, data.length, data[0].length).setValues(data);

    // Code to update the date and time
    var currentDate = new Date();
    var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"); // Format the date as yyyy-MM-dd HH:mm
    incomeDB.getRange(blankRow, 10).setValue(formattedDate);

    // Submitted by who
    incomeDB.getRange(blankRow, 11).setValue(Session.getActiveUser().getEmail());

    //Active Cell Data Input as gerbage Deta 
    var activeCellValue = myActiveSheet.getActiveCell().getValue();
    incomeDB.getRange(blankRow, 12).setValue(activeCellValue);

    uI.alert("Submit Successfully");


    // Clear the content of multiple cells at once
    var rangesToClear = ["I6", "I8", "I10", "I12", "L6", "L8", "L10", "L12", "I14", "J14", "K14","L14"];
    rangesToClear.forEach(function(range) {
      myActiveSheet.getRange(range).clearContent();
    });

    // Clear the active cell
    myActiveSheet.getActiveCell().clearContent();

    // Assign new date in C6 cell
    myActiveSheet.getRange("I6").setValue(new Date());

    // Set background color for input fields
    var ranges = ["I6", "L6", "L8", "L12"];
    ranges.forEach(function(range) {
      myActiveSheet.getRange(range).setBackground("#FFFFFF");
    });
  }
}


// Function to clear data
function clearData2() {
  // Declare a variable and set the reference of the active Google sheet
  var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var myActiveSheet = mySpreadsheet.getSheetByName("IncomeExpenses");

  // Clear the active cell
  myActiveSheet.getActiveCell().clearContent();

  var uI = SpreadsheetApp.getUi();

  // Clear the content of multiple cells at once
  var rangesToClear = ["I6", "I8", "I10", "I12", "L6", "L8", "L10", "L12", "I14", "J14", "K14","L14"];
  rangesToClear.forEach(function(range) {
    myActiveSheet.getRange(range).clearContent();
  });

  // Clear the active cell
  myActiveSheet.getActiveCell().clearContent();

  // Assign new date in C6 cell
  myActiveSheet.getRange("I6").setValue(new Date());

  // Set background color for input fields
  var ranges = ["I6", "L6", "L8", "L12"];
  ranges.forEach(function(range) {
    myActiveSheet.getRange(range).setBackground("#FFFFFF");
  });

  uI.alert("Clear Successfully");
}
