//this is the ID of the calendar where you want the event to show up, this is found on the calendar settings page of your/kyowa calendar 
let calendarId = "YOUR_SECRET_CALENDAR_ID_GOES_HERE";

//below are the column ids that represents the values used in the spreadsheet (these start with 1 I believe)
//start date - make sure you also check the time box in the Google form 
let startDtId = 5;

//end date - make sure you check the time box in the Google form
let endDtId = 7;

//title of the event
let titleId = 2;

//description
let descId = 8;

//location
let locId = 9;

//timestamp of the form submission
let formTimeStampId = 1;


function getLatestAndSubmitToCalendar() {
  let sheet = SpreadsheetApp.getActiveSheet();
  let rows = sheet.getDataRange();
  let numRows = rows.getNumRows();
  let values = rows.getValues();
  let lr = rows.getLastRow();

  //date/time is entered here but the pattern is sheet.getRange(lr,VARIABLE let,1,1)getValue();
  let startDt = sheet.getRange(lr, startDtId, 1, 1).getValue();
  let endDt = sheet.getRange(lr, endDtId, 1, 1).getValue();

  let userName = sheet.getRange(lr, userId, 1, 1).getValue();

  let subOn =
    "Submitted on: " + sheet.getRange(lr, formTimeStampId, 1, 1).getValue();
  let desc =
    sheet.getRange(lr, descId, 1, 1).getValue() +
    "\n" +
    subOn +
    "\n" +
    "Added by: " +
    userName;
  let title = sheet.getRange(lr, titleId, 1, 1).getValue();
  let loc = sheet.getRange(lr, locId, 1, 1).getValue();
  createEvent(calendarId, title, startDt, endDt, desc, loc);
}

function createEvent(calendarId, title, startDt, endDt, desc, loc) {
  let cal = CalendarApp.getCalendarById(calendarId);
  let start = new Date(startDt);
  let end = new Date(endDt);
  let loc = loc;

  let event = cal.createEvent(title, start, end, {
    description: desc,
    location: loc,
  });
}
