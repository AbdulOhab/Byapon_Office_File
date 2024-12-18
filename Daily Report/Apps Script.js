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

  // validate Source Of income
  if (myActiveSheet.getRange("C10").isBlank()) {
    uI.alert("Please Enter Source Of income");
    myActiveSheet.getRange("C10").setBackground("#FF0000");
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

  var uI = SpreadsheetApp.getUi();

  // Clear the active cell
  myActiveSheet.getActiveCell().clearContent();

  // Clear the content of multiple cells at once
  var rangesToClear = ["I6", "I8", "I10", "I12", "L6", "L8", "L10", "L12", "I14", "J14", "K14","L14"];
  rangesToClear.forEach(function(range) {
    myActiveSheet.getRange(range).clearContent();
  });

  // Assign new date in C6 cell
  myActiveSheet.getRange("I6").setValue(new Date());

  // Set background color for input fields
  var ranges = ["I6", "L6", "L8", "L12"];
  ranges.forEach(function(range) {
    myActiveSheet.getRange(range).setBackground("#FFFFFF");
  });

  uI.alert("Clear Successfully");
}

//IncomeExpenses Sheet protection.
function protectSheetWithExceptions_IncomeExpenses_Sheet() {
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

  // Add additional editors
  protection.addEditors(['byapon@gmail.com']); 
}


//Marketing Form

// Function to validate entry made by user
function validEntry3() {
  var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var myActiveSheet = mySpreadsheet.getSheetByName("Marketing"); // this working spreadsheet
  var uI = SpreadsheetApp.getUi(); // to create the instance of the user interface to show the alert.

  // Input field default color
  var ranges = ["C4", "C6", "C8", "C10", "C12", "C16", "C18", "C22", "C24", "F6", "F8", "F10", "F12", "F14", "F16", "F18"];
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
  if (myActiveSheet.getRange("C8").isBlank() == true) {
    uI.alert("Please Enter Marketing Type");
    myActiveSheet.getRange("C8").setBackground("#FF0000");
    return false;
  }

  // Validation: Library Name
  if (myActiveSheet.getRange("C10").isBlank() == true) {
    uI.alert("Please Enter Visiting Library Name");
    myActiveSheet.getRange("C10").setBackground("#FF0000");
    return false;
  }

  return true;
}

// Function to submit the data to Database
function submitData3() {
  var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var myActiveSheet = mySpreadsheet.getSheetByName("Marketing");
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
        myActiveSheet.getRange("F12").getValue(),
      	myActiveSheet.getRange("F14").getValue(),
      	myActiveSheet.getRange("F16").getValue(),
        myActiveSheet.getRange("F18").getValue()
      ],
    ];

    // Write the data to the spreadsheet in one operation
    incomeDB.getRange(blankRow, 1, data.length, data[0].length).setValues(data);

    // Code to update the date and time
    var currentDate = new Date();
    var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"); // Format the date as yyyy-MM-dd HH:mm
    incomeDB.getRange(blankRow, 17).setValue(formattedDate);

    // Submitted by who
    incomeDB.getRange(blankRow, 18).setValue(Session.getActiveUser().getEmail());

    //Active Cell Data Input as gerbage Deta 
    var activeCellValue = myActiveSheet.getActiveCell().getValue();
    incomeDB.getRange(blankRow, 19).setValue(activeCellValue);

    uI.alert("Submit Successfully");


    // Clear the content of multiple cells at once
    var rangesToClear = [ "C4", "C6", "C8", "C10", "C12", "C14", "C16", "C18", "C22", "C24","F6", "F8", "F10", "F12", "F14", "F16", "F18"];

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
  var rangesToClear = [ "C4", "C6", "C8", "C10", "C12", "C14", "C16", "C18", "C22", "C24","F6", "F8", "F10", "F12", "F14", "F16", "F18"];

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


//lock MarketingSheet
function protectSheetWithExceptions_Marketing() {
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
  var cells = [ "C4", "C6", "C8", "C10", "C12", "C14", "C16", "C18", "C22", "C24", "F6", "F8", "F10", "F12", "F14", "F16", "F18"];

  cells.forEach(function(cell) {
    unprotectedRanges.push(sheet.getRange(cell));
  });

  // Apply the unprotected cells to the protection
  protection.setUnprotectedRanges(unprotectedRanges);

  // Add additional editors
  protection.addEditors(['byapon@gmail.com']); 
}

//lock Marckting DB
// function protectSheetWithExceptions3_MarketingDB() {
//  var sheetName = 'MarketingDB';
//  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
//  // Protect the entire sheet
//  var protection = sheet.protect().setDescription('Protected Sheet');
  
//  // Ensure the current user is an editor before removing others
//  var me = Session.getEffectiveUser();
//  protection.addEditor(me);
//  protection.removeEditors(protection.getEditors());
// }

//Lock IncomdeDB
// function protectSheetWithExceptions4_IncomeDB() {
//  var sheetName = 'IncomeDB';
//  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
//  // Protect the entire sheet
//  var protection = sheet.protect().setDescription('Protected Sheet');
  
//  // Ensure the current user is an editor before removing others
//  var me = Session.getEffectiveUser();
//  protection.addEditor(me);
//  protection.removeEditors(protection.getEditors());
// }

//lock ExpenceDB
// function protectSheetWithExceptions6_ExpenceDB() {
//  var sheetName = 'ExpenceDB';
//  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
//  // Protect the entire sheet
//  var protection = sheet.protect().setDescription('Protected Sheet');
  
//  // Ensure the current user is an editor before removing others
//  var me = Session.getEffectiveUser();
//  protection.addEditor(me);
//  protection.removeEditors(protection.getEditors());
//  // Add additional editors
//  protection.addEditors(['byapon@gmail.com']);

// }
// The Second Income Page2
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

  // validate Source Of income
  if (myActiveSheet.getRange("C9").isBlank()) {
    uI.alert("Please Enter Source Of income");
    myActiveSheet.getRange("C9").setBackground("#FF0000");
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

// lock MarketingSheet
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

  // Add additional editors
  protection.addEditors(['byapon@gmail.com']); 
}

//lock ExtendDB
function protectSheetWithExceptions6_ExtendDB() {
 var sheetName = 'ExtendDB';
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
 // Protect the entire sheet
 var protection = sheet.protect().setDescription('Protected Sheet');
  
 // Ensure the current user is an editor before removing others
 var me = Session.getEffectiveUser();
 protection.addEditor(me);
 protection.removeEditors(protection.getEditors());
 // Add additional editors
 protection.addEditors(['byapon@gmail.com']);

}

function protectSheetWithExceptions5_bikkoroy() {
 var sheetName = 'বিক্রয়';
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
 // Protect the entire sheet
 var protection = sheet.protect().setDescription('Protected Sheet');
  
 // Ensure the current user is an editor before removing others
 var me = Session.getEffectiveUser();
 protection.addEditor(me);
 protection.removeEditors(protection.getEditors());
}

// IF you want to unprotect the sheet use this  script 
function unprotectSheet() {
  var sheetName = "IncomeExpenses";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  // Get all protections in the sheet
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);

  // Remove all protections
  protections.forEach(function(protection) {
    protection.remove();
  });
} 

