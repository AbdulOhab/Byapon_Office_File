//Marketing Form

// Function to validate entry made by user
function validEntry3() {
  var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var myActiveSheet = mySpreadsheet.getSheetByName("Marketing"); // this working spreadsheet
  var uI = SpreadsheetApp.getUi(); // to create the instance of the user interface to show the alert.

  // Input field default color
  var ranges = ["C4", "C6", "C8", "C10", "C12", "C16", "C18", "C22", "C24",
                 "F6", "F8", "F10", "F12", "F14", "F16"];
  ranges.forEach(function(range) {
    myActiveSheet.getRange(range).setBackground("#FFFFFF");
  });

  // Validation: Date
  if (myActiveSheet.getRange("C4").isBlank() == true) {
    uI.alert("Please Enter Date");
    myActiveSheet.getRange("C4").setBackground("#FF0000");
    return false;
  }

  // Validation: Marketer Name
  if (myActiveSheet.getRange("C6").isBlank() == true) {
    uI.alert("Please Enter Marketer Name");
    myActiveSheet.getRange("C6").setBackground("#FF0000");
    return false;
  }

  // Validation: Marketing Type
  if (myActiveSheet.getRange("C10").isBlank() == true) {
    uI.alert("Please Enter Marketing Type");
    myActiveSheet.getRange("C10").setBackground("#FF0000");
    return false;
  }

  // Validation: Library Name
  if (myActiveSheet.getRange("C12").isBlank() == true) {
    uI.alert("Please Enter Visiting Library Name");
    myActiveSheet.getRange("C12").setBackground("#FF0000");
    return false;
  }

  return true;
}

// Function to submit the data to Database
function submitData3() {
  var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var myActiveSheet = mySpreadsheet.getSheetByName("IncomeExpenses");
  var incomeDB = mySpreadsheet.getSheetByName("MarketingDB");
  var uI = SpreadsheetApp.getUi(); // show the alert.

  var response = uI.alert("Submit", "Do you want to Submit?", uI.ButtonSet.YES_NO);

  // checking user response
  if (response == uI.Button.NO) {
    return;
  }

  if (validEntry3()) {
    var blankRow = incomeDB.getLastRow() + 1;

    // Prepare a 2D array to hold the data
    var data = [
      [ 
      	myActiveSheet.getRange("C4").getValue(),
        myActiveSheet.getRange("C6").getValue(), 
      	myActiveSheet.getRange("C8").getValue(), 
      	myActiveSheet.getRange("C10").getValue(),
      	myActiveSheet.getRange("C12").getValue(),
      	myActiveSheet.getRange("C16").getValue(),
        myActiveSheet.getRange("C18").getValue(), 
        myActiveSheet.getRange("C22").getValue(),
        myActiveSheet.getRange("C24").getValue(),
      	myActiveSheet.getRange("F6").getValue(),
      	myActiveSheet.getRange("F8").getValue(),
      	myActiveSheet.getRange("F10").getValue(),
      	myActiveSheet.getRange("F14").getValue(),
      	myActiveSheet.getRange("F16").getValue()
      ],
    ];

    // Write the data to the spreadsheet in one operation
    incomeDB.getRange(blankRow, 1, data.length, data[0].length).setValues(data);

    // Code to update the date and time
    var currentDate = new Date();
    var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"); // Format the date as yyyy-MM-dd HH:mm
    incomeDB.getRange(blankRow, 15).setValue(formattedDate);

    // Submitted by who
    incomeDB.getRange(blankRow, 16).setValue(Session.getActiveUser().getEmail());

    //Active Cell Data Input as gerbage Deta 
    var activeCellValue = myActiveSheet.getActiveCell().getValue();
    incomeDB.getRange(blankRow, 17).setValue(activeCellValue);

    uI.alert("Submit Successfully");


    // Clear the content of multiple cells at once
    var rangesToClear = [ "C4", "C6", "C8", "C10", "C12", "C14", "C16", "C18", "C22", "C24",
                        "F6", "F8", "F10", "F12", "F14", "F16"];

    rangesToClear.forEach(function(range) {
    	myActiveSheet.getRange(range).clearContent();
    });

    // Clear the active cell
    myActiveSheet.getActiveCell().clearContent();

    // Assign new date in C4 cell
    myActiveSheet.getRange("C4").setValue(new Date());

    // Set background color for input fields
    var ranges = ["C4", "C6", "C8", "F10"];
    ranges.forEach(function(range) {
      myActiveSheet.getRange(range).setBackground("#FFFFFF");
    });
  }
}

// Function to clear data
function clearData3() {
  // Declare a variable and set the reference of the active Google sheet
  var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var myActiveSheet = mySpreadsheet.getSheetByName("Marketing");

  var uI = SpreadsheetApp.getUi();

  // Clear the content of multiple cells at once
  var rangesToClear = [ "C4", "C6", "C8", "C10", "C12", "C14", "C16", "C18", "C22", "C24",
                        "F6", "F8", "F10", "F12", "F14", "F16"];

  rangesToClear.forEach(function(range) {
    myActiveSheet.getRange(range).clearContent();
  });

  // Clear the active cell
  myActiveSheet.getActiveCell().clearContent();

  // Assign new date in C4 cell
  myActiveSheet.getRange("C4").setValue(new Date());

  // Set background color for input fields
  var ranges = ["C4", "C6", "C8", "F10"];
  ranges.forEach(function(range) {
    myActiveSheet.getRange(range).setBackground("#FFFFFF");
  });

  uI.alert("Clear Successfully");
}

function protectSheetWithExceptions() {
  // Replace 'YourSheetName' with the name of your specific sheet
  var sheetName = 'Marketing';
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
  var cells = [ "C4", "C6", "C8", "C10", "C12", "C14", "C16", "C18", "C22", "C24",
                "F6", "F8", "F10", "F12", "F14", "F16"];

  cells.forEach(function(cell) {
    unprotectedRanges.push(sheet.getRange(cell));
  });

  // Apply the unprotected cells to the protection
  protection.setUnprotectedRanges(unprotectedRanges);
}