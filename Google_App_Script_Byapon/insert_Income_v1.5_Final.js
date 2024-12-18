//Income Form

//Numaric Function
function isNumeric(n) {
  return !isNaN(parseFloat(n)) && isFinite(n);
}

// Function to validate entry made by user
function validEntry() {
  var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var myActiveSheet = mySpreadsheet.getSheetByName("IncomeExpenses"); // this working spreadsheet
  var uI = SpreadsheetApp.getUi(); // to create the instance of the user interface to show the alert.

  // Input field default color
  var ranges = ["C6", "C8", "C10", "C12", "C14", "C18:F18", "F6", "F8", "F10", "F12", "F14", "F16"];
  ranges.forEach(function(range) {
    myActiveSheet.getRange(range).setBackground("#FFFFFF");
  });

  // validate Date
  if (myActiveSheet.getRange("C6").isBlank()) {
    uI.alert("Please Enter Date");
    myActiveSheet.getRange("C6").setBackground("#FF0000");
    return false;
  }

  // validate Money Receipt
  if (myActiveSheet.getRange("C14").isBlank()) {
    uI.alert("Please Enter Money Receipt Details");
    myActiveSheet.getRange("C14").setBackground("#FF0000");
    return false;
  }

  return true;
}

// Function to submit the data to Database
function submitData() {
  var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var myActiveSheet = mySpreadsheet.getSheetByName("IncomeExpenses");
  var incomeDB = mySpreadsheet.getSheetByName("IncomeDB");
  var uI = SpreadsheetApp.getUi(); // show the alert.

  var response = uI.alert("Submit", "Do you want to Submit?", uI.ButtonSet.YES_NO);

  // checking user response
  if (response == uI.Button.NO) {
    return;
  }

  if (validEntry()) {
    var blankRow = incomeDB.getLastRow() + 1;

    // Prepare a 2D array to hold the data
    var data = [
      [ 
      	myActiveSheet.getRange("C6").getValue(), 
      	myActiveSheet.getRange("C8").getValue(), 
      	myActiveSheet.getRange("C10").getValue(),
      	myActiveSheet.getRange("C12").getValue(),
      	myActiveSheet.getRange("C14").getValue(),
      	myActiveSheet.getRange("F6").getValue(),
      	myActiveSheet.getRange("F8").getValue(),
      	myActiveSheet.getRange("F10").getValue(),
      	myActiveSheet.getRange("F12").getValue(),
      	myActiveSheet.getRange("F14").getValue(),
      	myActiveSheet.getRange("F16").getValue(),
      	myActiveSheet.getRange("C18:F18").getValue()
      ],
    ];

    // Write the data to the spreadsheet in one operation
    incomeDB.getRange(blankRow, 1, data.length, data[0].length).setValues(data);

    // Code to update the date and time
    var currentDate = new Date();
    var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"); // Format the date as yyyy-MM-dd HH:mm
    incomeDB.getRange(blankRow, 13).setValue(formattedDate);

    // Submitted by who
    incomeDB.getRange(blankRow, 14).setValue(Session.getActiveUser().getEmail());

    //Active Cell Data Input as gerbage Deta 
    var activeCellValue = myActiveSheet.getActiveCell().getValue();
    incomeDB.getRange(blankRow, 15).setValue(activeCellValue);

    uI.alert("Submit Successfully");


    // Clear the content of multiple cells at once
    var rangesToClear = ["C6", "C8", "C10", "C12", "C14", "F6", "F8", "F10", "F12", "F14", "F16", "C18", "D18", "E18", "F18"];
    rangesToClear.forEach(function(range) {
    	myActiveSheet.getRange(range).clearContent();
    });

    // Clear the active cell
    myActiveSheet.getActiveCell().clearContent();

    // Assign new date in C6 cell
    myActiveSheet.getRange("C6").setValue(new Date());

    // Set background color for input fields
    var ranges = ["C6", "C10", "C14", "F6"];
    ranges.forEach(function(range) {
      myActiveSheet.getRange(range).setBackground("#FFFFFF");
    });
  }
}

// Function to clear data
function clearData() {
  // Declare a variable and set the reference of the active Google sheet
  var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var myActiveSheet = mySpreadsheet.getSheetByName("IncomeExpenses");

  var uI = SpreadsheetApp.getUi();

  // Clear the content of multiple cells at once
  var rangesToClear = ["C6", "C8", "C10", "C12", "C14", "F6", "F8", "F10", "F12", "F14", "F16", "C18", "D18", "E18", "F18"];
  rangesToClear.forEach(function(range) {
    myActiveSheet.getRange(range).clearContent();
  });

  // Clear the active cell
  myActiveSheet.getActiveCell().clearContent();

  // Assign new date in C6 cell
  myActiveSheet.getRange("C6").setValue(new Date());

  // Set background color for input fields
  var ranges = ["C6", "C10", "C14", "F6"];
  ranges.forEach(function(range) {
    myActiveSheet.getRange(range).setBackground("#FFFFFF");
  });

  uI.alert("Clear Successfully");
}
