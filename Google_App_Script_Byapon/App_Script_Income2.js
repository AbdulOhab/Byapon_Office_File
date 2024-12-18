// Income2
function validEntry_income2() {
  var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var myActiveSheet = mySpreadsheet.getSheetByName("Income2"); // this working spreadsheet
  var uI = SpreadsheetApp.getUi(); // to create the instance of the user interface to show the alert.

  // Input field default color
  // var ranges = ["C5", "C7", "C9", "C11", "C13",
  //   "C18", "C20", "C22", "C24", "C26", "C28",
  //   "F5", "F7", "F9", "F11", "F13", "F15" ];

  var ranges = [
  "C5", "C7", "C9", "C11", "C13",
  "C18", "C20", "C22", "C24", "C26", "C28",
  "F5", "F7", "F9", "F11", "F13", "F15",
  "F18", "F20", "F22", "F24", "F26", "F28",
  "I5", "I7", "I9", "I11", "I13", "I15",
  "I18", "I20", "I22", "I24", "I26", "I28",
  "C30:I30"];

  ranges.forEach(function (range) {
    myActiveSheet.getRange(range).setBackground("#FFFFFF");
  });

  // validate Date
  if (myActiveSheet.getRange("C5").isBlank()) {
    uI.alert("Please Enter Date");
    myActiveSheet.getRange("C5").setBackground("#FF0000");
    return false;
  }

  // validate Money Receipt
  if (myActiveSheet.getRange("C13").isBlank()) {
    uI.alert("Please Enter Money Receipt Details");
    myActiveSheet.getRange("C13").setBackground("#FF0000");
    return false;
  }

  return true;
}

// Function to submit the data to Database
function submitData_income2() {
  var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var myActiveSheet = mySpreadsheet.getSheetByName("Income2");
  var incomeDB = mySpreadsheet.getSheetByName("IncomeDB");
  var uI = SpreadsheetApp.getUi(); // show the alert.

  var response = uI.alert(
    "Submit",
    "Do you want to Submit?",
    uI.ButtonSet.YES_NO
  );

  // checking user response
  if (response == uI.Button.NO) {
    return;
  }

  if (validEntry_income2()) {
    var blankRow = incomeDB.getLastRow() + 1;

    var C5value = myActiveSheet.getRange("C5").getValue();
    var C7value = myActiveSheet.getRange("C7").getValue();
    var C9value = myActiveSheet.getRange("C9").getValue();
    var C11value = myActiveSheet.getRange("C11").getValue();
    var C13value = myActiveSheet.getRange("C13").getValue();
    var AddComments = myActiveSheet.getRange("C30:I30").getValue();

    var First_Submission = [
      [
        C5value, C7value, C9value, C11value, C13value,
        myActiveSheet.getRange("C18").getValue(),
        myActiveSheet.getRange("C20").getValue(),
        myActiveSheet.getRange("C22").getValue(),
        myActiveSheet.getRange("C24").getValue(),
        myActiveSheet.getRange("C26").getValue(),
        myActiveSheet.getRange("C28").getValue(),
      ],
    ];

    if (!myActiveSheet.getRange("C18").isBlank()) {
      // Write the data to the spreadsheet in one operation
      incomeDB.getRange(blankRow, 1, First_Submission.length, First_Submission[0].length).setValues(First_Submission);

      incomeDB.getRange(blankRow,12).setValue(AddComments);

      // Code to update the date and time | It is Timestamp
      var currentDate = new Date();
      var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"); // Format the date as yyyy-MM-dd HH:mm
      incomeDB.getRange(blankRow, 13).setValue(formattedDate);

      // Submitted by who
      incomeDB.getRange(blankRow, 14).setValue(Session.getActiveUser().getEmail());

      //Active Cell Data Input as gerbage Deta
      var activeCellValue = myActiveSheet.getActiveCell().getValue();
      incomeDB.getRange(blankRow, 15).setValue(activeCellValue);
    }

    var blankRow2 = incomeDB.getLastRow() + 1;

    var Second_Submission = [
      [
        C5value, C7value, C9value, C11value, C13value,
        myActiveSheet.getRange("F5").getValue(),
        myActiveSheet.getRange("F7").getValue(),
        myActiveSheet.getRange("F9").getValue(),
        myActiveSheet.getRange("F11").getValue(),
        myActiveSheet.getRange("F13").getValue(),
        myActiveSheet.getRange("F15").getValue(),
      ],
    ];

    if (!myActiveSheet.getRange("F5").isBlank()) {
      // Write the data to the spreadsheet in one operation
      incomeDB.getRange(blankRow2, 1, Second_Submission.length, Second_Submission[0].length).setValues(Second_Submission);

      // Code to update the date and time | It is Timestamp
      var currentDate = new Date();
      var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"); // Format the date as yyyy-MM-dd HH:mm
      incomeDB.getRange(blankRow2, 13).setValue(formattedDate);

      // Submitted by who
      incomeDB.getRange(blankRow2, 14).setValue(Session.getActiveUser().getEmail());

      //Active Cell Data Input as gerbage Deta
      var activeCellValue = myActiveSheet.getActiveCell().getValue();
      incomeDB.getRange(blankRow2, 15).setValue(activeCellValue);
    }

    var blankRow3 = incomeDB.getLastRow() + 1;

    var Submission3 = [
      [
        C5value, C7value, C9value, C11value, C13value,
        myActiveSheet.getRange("F18").getValue(),
        myActiveSheet.getRange("F20").getValue(),
        myActiveSheet.getRange("F22").getValue(),
        myActiveSheet.getRange("F24").getValue(),
        myActiveSheet.getRange("F26").getValue(),
        myActiveSheet.getRange("F28").getValue(),
      ],
    ];

    if (!myActiveSheet.getRange("F18").isBlank()) {
      // Write the data to the spreadsheet in one operation
      incomeDB.getRange(blankRow3, 1, Submission3.length, Submission3[0].length).setValues(Submission3);

      // Code to update the date and time | It is Timestamp
      var currentDate = new Date();
      var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"); // Format the date as yyyy-MM-dd HH:mm
      incomeDB.getRange(blankRow3, 13).setValue(formattedDate);

      // Submitted by who
      incomeDB.getRange(blankRow3, 14).setValue(Session.getActiveUser().getEmail());

      //Active Cell Data Input as gerbage Deta
      var activeCellValue = myActiveSheet.getActiveCell().getValue();
      incomeDB.getRange(blankRow3, 15).setValue(activeCellValue);
    }

    var blankRow4 = incomeDB.getLastRow() + 1;

    var Submission4 = [
      [
        C5value, C7value, C9value, C11value, C13value,
        myActiveSheet.getRange("I5").getValue(),
        myActiveSheet.getRange("I7").getValue(),
        myActiveSheet.getRange("I9").getValue(),
        myActiveSheet.getRange("I11").getValue(),
        myActiveSheet.getRange("I13").getValue(),
        myActiveSheet.getRange("I15").getValue(),
      ],
    ];

    if (!myActiveSheet.getRange("I5").isBlank()) {
      // Write the data to the spreadsheet in one operation
      incomeDB.getRange(blankRow4, 1, Submission4.length, Submission4[0].length).setValues(Submission4);

      // Code to update the date and time | It is Timestamp
      var currentDate = new Date();
      var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"); // Format the date as yyyy-MM-dd HH:mm
      incomeDB.getRange(blankRow4, 13).setValue(formattedDate);

      // Submitted by who
      incomeDB.getRange(blankRow4, 14).setValue(Session.getActiveUser().getEmail());

      //Active Cell Data Input as gerbage Deta
      var activeCellValue = myActiveSheet.getActiveCell().getValue();
      incomeDB.getRange(blankRow4, 15).setValue(activeCellValue);
    }

    var blankRow5 = incomeDB.getLastRow() + 1;

    var Submission5 = [
      [
        C5value, C7value, C9value, C11value, C13value,
        myActiveSheet.getRange("I18").getValue(),
        myActiveSheet.getRange("I20").getValue(),
        myActiveSheet.getRange("I22").getValue(),
        myActiveSheet.getRange("I24").getValue(),
        myActiveSheet.getRange("I26").getValue(),
        myActiveSheet.getRange("I28").getValue(),
      ],
    ];

    if (!myActiveSheet.getRange("I18").isBlank()) {
      // Write the data to the spreadsheet in one operation
      incomeDB.getRange(blankRow5, 1, Submission5.length, Submission5[0].length).setValues(Submission5);

      incomeDB.getRange(blankRow5,12).setValue(AddComments);
      // Code to update the date and time | It is Timestamp
      var currentDate = new Date();
      var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"); // Format the date as yyyy-MM-dd HH:mm
      incomeDB.getRange(blankRow5, 13).setValue(formattedDate);

      // Submitted by who
      incomeDB.getRange(blankRow5, 14).setValue(Session.getActiveUser().getEmail());

      //Active Cell Data Input as gerbage Deta
      var activeCellValue = myActiveSheet.getActiveCell().getValue();
      incomeDB.getRange(blankRow5, 15).setValue(activeCellValue);
    }


    uI.alert("Submit Successfully");

    // Clear the content of multiple cells at once
    var rangesToClear = [
      "C5", "C7", "C9", "C11", "C13",
      "C18", "C20", "C22", "C24", "C26", "C28",
      "F5", "F7", "F9", "F11", "F13", "F15",
      "F18", "F20", "F22", "F24", "F26", "F28",
      "I5", "I7", "I9", "I11", "I13", "I15",
      "I18", "I20", "I22", "I24", "I26", "I28", "C30:I30"];

    rangesToClear.forEach(function(range) {
    	myActiveSheet.getRange(range).clearContent();
    });

    // Clear the active cell
    myActiveSheet.getActiveCell().clearContent();

    // Assign new date in C6 cell
    myActiveSheet.getRange("C5").setValue(new Date());

    // Set background color for input fields
    var ranges = ["C5", "C13"];
    ranges.forEach(function(range) {
      myActiveSheet.getRange(range).setBackground("#FFFFFF");
    });
  }
}

// Function to clear data
function clearData_income2() {
  // Declare a variable and set the reference of the active Google sheet
  var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var myActiveSheet = mySpreadsheet.getSheetByName("Income2");

  var uI = SpreadsheetApp.getUi();

  // Clear the content of multiple cells at once
  var rangesToClear = [
    "C5", "C7", "C9", "C11", "C13",
    "C18", "C20", "C22", "C24", "C26", "C28",
    "F5", "F7", "F9", "F11", "F13", "F15",
    "F18", "F20", "F22", "F24", "F26", "F28",
    "I5", "I7", "I9", "I11", "I13", "I15",
    "I18", "I20", "I22", "I24", "I26", "I28","C30:I30"];

  rangesToClear.forEach(function(range) {
    myActiveSheet.getRange(range).clearContent();
  });

  // Clear the active cell
  myActiveSheet.getActiveCell().clearContent();

  // Assign new date in C6 cell
  myActiveSheet.getRange("C5").setValue(new Date());

  // Set background color for input fields
  var ranges = ["C5", "C13"];
  ranges.forEach(function(range) {
    myActiveSheet.getRange(range).setBackground("#FFFFFF");
  });

  uI.alert("Clear Successfully");
}

//lock MarketingSheet
function protectSheetWithExceptions_income2() {
  var sheetName = "Income2";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  // Protect the entire sheet
  var protection = sheet.protect().setDescription("Protected Sheet");

  // Ensure the current user is an editor before removing others
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors());

  // Set unprotected cells
  var unprotectedRanges = [];

  // Income
  var ClearCellAll = [
    "C5", "C7", "C9", "C11", "C13",
    "C18", "C20", "C22", "C24", "C26", "C28",
    "F5", "F7", "F9", "F11", "F13", "F15",
    "F18", "F20", "F22", "F24", "F26", "F28",
    "I5", "I7", "I9", "I11", "I13", "I15",
    "I18", "I20", "I22", "I24", "I26", "I28", "C30:I30"];

  ClearCellAll.forEach(function (cell) {
    unprotectedRanges.push(sheet.getRange(cell));
  });

  // Apply the unprotected cells to the protection
  protection.setUnprotectedRanges(unprotectedRanges);
}

