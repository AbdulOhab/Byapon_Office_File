// -------------------------------------------- this is for Income Form
// function to valid to entry made by user 
function validEntry(){
 var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 var myActiveSheet =  mySpreadsheet.getSheetByName("IncomeExpenses"); // this working spreadsheet

 var uI = SpreadsheetApp.getUi(); // to create the instance of the user interface ot show the alert.

// Input field defoult color
 myActiveSheet.getRange("C6").setBackground("#FFFFFF");
 myActiveSheet.getRange("C8").setBackground("#FFFFFF");
 myActiveSheet.getRange("C10").setBackground("#FFFFFF");
 myActiveSheet.getRange("C12").setBackground("#FFFFFF");
 myActiveSheet.getRange("C14").setBackground("#FFFFFF");
 myActiveSheet.getRange("C16").setBackground("#FFFFFF");
 myActiveSheet.getRange("F6").setBackground("#FFFFFF");
 myActiveSheet.getRange("F8").setBackground("#FFFFFF");
 myActiveSheet.getRange("F10").setBackground("#FFFFFF");
 myActiveSheet.getRange("F12").setBackground("#FFFFFF");
 myActiveSheet.getRange("F14").setBackground("#FFFFFF"); 
 myActiveSheet.getRange("F16").setBackground("#FFFFFF");
 myActiveSheet.getRange("C18").setBackground("#FFFFFF");
 myActiveSheet.getRange("C27").setBackground("#FFFFFF");

 //validateion date
if(myActiveSheet.getRange("C6").isBlank() == true )
{
  uI.alert("Please Enter Date");
  myActiveSheet.getRange("C6").setBackground("#FF0000");
  return false;
}
//validateion Employee Name
if(myActiveSheet.getRange("C8").isBlank() == true )
{
  uI.alert("Please Enter Employee Name");
  myActiveSheet.getRange("C8").setBackground("#FF0000");
  return false;
}

//validateion Income Source
if(myActiveSheet.getRange("C10").isBlank() == true )
{
  uI.alert("please Enter Income source");
  myActiveSheet.getRange("C10").setBackground("#FF0000");
  return false;
}

//validateion Selling type
if(myActiveSheet.getRange("C12").isBlank() == true )
{
  uI.alert("please Enter Selling Type");
  myActiveSheet.getRange("C12").setBackground("#FF0000");
  return false;
}

return true;

}

//Function to submit the deta to detabse 
function submitDeta(){
  //declear a variable and set the reference of active google sheet
 var mySpreadsheet2 = SpreadsheetApp.getActiveSpreadsheet();
 var myActiveSheet =  mySpreadsheet2.getSheetByName("IncomeExpenses"); 
 var incomeDB =  mySpreadsheet2.getSheetByName("IncomeDB");

 var uI2 = SpreadsheetApp.getUi(); //show the alert.
 
 var response = uI2.alert("Submit","Do you want to submit deta?", uI2.ButtonSet.YES_NO); 

 // checkeing usder response

 if(response == uI2.Button.NO){
  return;
 }

 if(validEntry() == true){

  var blankRow = incomeDB.getLastRow() + 1;

  incomeDB.getRange(blankRow,1).setValue(myActiveSheet.getRange("C6").getValue()); //date
  incomeDB.getRange(blankRow,2).setValue(myActiveSheet.getRange("C8").getValue()); //Employee Name
  incomeDB.getRange(blankRow,3).setValue(myActiveSheet.getRange("C10").getValue()); //Income Source
  incomeDB.getRange(blankRow,4).setValue(myActiveSheet.getRange("C12").getValue()); //Selling type

  // Code to update the date and time
  var currentDate = new Date(); // Get current date and time
  var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"); // Format the date as yyyy-MM-dd HH:mm
  incomeDB.getRange(blankRow, 5).setValue(formattedDate); // Set the formatted date to the desired cell 
  //incomeDB.getRange(blankRow,5).setValue(new Date().setNumberFormat("yyyy-mm-dd h:mm"));

  // Submitted by who
    incomeDB.getRange(blankRow,6).setValue(Session.getActiveUser().getEmail());

  uI2.alert(' "New data seve - emp #" ' + myActiveSheet.getRange("C6").getValue() + '""' );

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
  myActiveSheet.getRange("C18").clearContent();
  myActiveSheet.getRange("C27").clearContent();
}

}
























