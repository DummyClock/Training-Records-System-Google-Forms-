// Google Calendar Id, needed to create events for a specific calendar
const calendarID = 'de5d185b5d89096524a6155561d0f7a13be4f5793e4738bc35552e8589ff5e9a@group.calendar.google.com'
// Link to Google Colab
const colab_URL = 'https://colab.research.google.com/drive/1jWeEa4l0CUo58sxe9RBUCs1DE2FXjjG2?usp=sharing'

/**
 * Automatically sorts new rows added into Spreadsheet via Google Form submissions. The Spreadsheet consists of three different classifications for 
 * Sheets: Request Types, Preliminary Training Types, Reinforcement Types. The way a Sheet is sorted depends on the type, dictated by the name, of 
 * said Sheet which are listed below.
 * 
 * Different Names for Sheets
 *    - Request Types: "Requests" or "Retraining_Requests"
 *    - Preliminary Training Types: "Initial_Training" or "Retraining"
 *    - Reinforcement Types: Everything Else besides "Skill_Board" or "List_0f_Trainers"
 */
function autoSort() {

  // Get current active sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const ws = ss.getActiveSheet()
  const range = ws.getDataRange().offset(1, 0, ws.getLastRow() - 1)
  const name = ws.getName()
  
  // Verify Trainer Email before doing anything!
  var newestEntry = ws.getDataRange().getValues()[ws.getDataRange().getValues().length-1]

  if(name == "Requests" || name =="Retraining_Requests") //if it's for a request; enter
  {
    requestHandler(ss, range);
  }
  else if((name == 'Initial_Training' || name == 'Retraining')) //else if the ws is Initial_Training or Retraining; enter
  {
    if(verifyTrainer(newestEntry[1])) 
    {
      initialHandler(ss, range);
    }
    else 
    {   // Delete if email doesn't have access to the sheet
    ws.deleteRow(ws.getLastRow()) 
    }
  }
  else  if(!(name == 'Skill_Board' || name == 'List_of_Trainers') )  // else if for a Reinforcement Report; enter
  {
    if(verifyTrainer(newestEntry[1])) 
    {
      reinforcementRepotHandler(ss, range, newestEntry);
    }
    else 
    {   // Delete if email doesn't have access to the sheet
      ws.deleteRow(ws.getLastRow()) 
    }
  } 
}



/**
 * A function that verfies if a trainer has permission to the training spreadsheet. If they do, then they are "allowed" to use the training form
 * links to fill out a report. If they don't have access to the spreadsheet, their submission results will be deleted. This function checks the last
 * column of a sheet named "List_of_Trainers"
 * 
 * @param {string} email_addr This is the email address that will be checked for verification.
 * 
 *  Warning: It's column sensitive, thus any future adjustments to the position of columns in the sheets needs to be handled with care and 
 *  attention. Moving columns may break the code for Preliminary Training Types and Reinforcement Types. 
 *      - Column B: Stores Trainer Emails that will be used for verifyication (Not neccessary for Requests)
 */
function verifyTrainer(email_addr)
{
  var sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var file = DriveApp.getFileById(sheetId);
  // var owner = file.getOwner();
  // var editors = file.getEditors();    // Only returns if the user running this script has editing permissions; otherwise null + error
  // var viewers = file.getViewers();    // Only returns if the user running this script has editing permissions; otherwise null + error

  var trainer_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("List_of_Trainers")
  users = trainer_sheet.getRange("B:B").getValues()

  // Logger.log(trainer_sheet.getRange("B:B").getValues()[1].toString())
  // Logger.log(users[i])

  if(email_addr===file.getOwner()) {
      return true;
  }

  for(var i=0; i<users.length; i++)
  {
    if(email_addr===users[i].toString())
      return true;
  }
  Logger.log(email_addr + "'s response was removed. " + i + " email(s) in total checked.");
  return false;
}

/**
 * Takes an entry/row of data and starts checking the values of each columns. It will generate a grade that will be saved in the 5th column of the
 * entry. This function calculates score based on the number of cells with the value "Yes", "No", "1", "2", "3", "4" or "5"; anything other value 
 * will be ignored by this function and will have no impact on the grading system.
 * 
 * @param {Object[][]} entry An array of data that will be read; this array should come from a row orginating from a spreadsheet.
 * @param {Sheet} ws The Google Sheet that the "entry" parameter comes from. This will be the Sheet that gets updated with the new grade.
 * @return {int} Returns a grade based on the values assessed from the entry parameter.
 * 
 * Warning: It's column sensitive, thus any future adjustments to the position of columns in the sheets needs to be handled with care and attention. 
 *      Moving columns may break the code for Reinforcement Types.
 *      - Column E: Used to store Grades. It should be manually inserted.
 */
function gradeReport(entry,ws)
{
  var score = 0
  var questions = 0

  //Loop through the row: archaic grading criteria (warning  this  can  easily be abused if a text based question is given "Yes" as a response)
  for(var i = 0; i < entry.length; i++)
  {
    if(entry[i] === "Yes")
    {
      score = score +  5
      questions = questions + 1
    }
    else if(entry[i] === "No")
    {
      questions = questions + 1
    }
    else if(1 <= parseInt(entry[i]) && 5 >= parseInt(entry[i]))
    {
      score = score + parseInt(entry[i])
      questions = questions + 1
    }
  }

  //Update Grade in the SpreadSheet
  var grade = (score/(questions*5)) * 100
  ws.getRange(ws.getLastRow(), 5).setValue(grade)
  return grade
}

/**
 * A function that sorts a sheet by a single, specified column (higest to lowest)
 * @param {} range The range of cells you wish to sort. 
 *    (Note: Ideally, you'd want to sort by all the cells. So you'd use "something like ws.getDataRange().offset(1, 0, ws.getLastRow() - 1)"
 * @param {int} col The column number you wish to sort the sheet by. 
 *    (Note: The count starts at 1, not 0. Therefore, Column A can be referenced by '1')
 */
function sortByOneColumn(range, col)
{
  range.sort([{column: col, ascending: false}])
}

/**
 * A function that sorts a sheet by a single, specified column
 * @param {range} range The range of cells you wish to sort. 
 *    (Note: Ideally, you'd want to sort by all the cells. So you'd use "something like ws.getDataRange().offset(1, 0, ws.getLastRow() - 1)"
 * @param {int} col1 The column number you wish to sort the sheet by. 
 *    (Note: The count starts at 1, not 0. Therefore, Column A can be referenced by '1')
 * @param {int} col2 The column number you wish to sort the sheet by. 
 *    (Note: The count starts at 1, not 0. Therefore, Column A can be referenced by '1')
 */
function sortByTwoColumns(range, col1, col2)
{
  range.sort([{column: col1, ascending: true}, {column: col2, ascending: false}])
}

/**
 * A function that attempts to send a plain-text email.
 * 
 * @param {string} email The desired reciepent (email address) to recieve the email.
 * @param {string} subject The subject line.
 * @param {string} body The body of the email.
 */
function emailResults(email, subject, body)
{
  try{
    GmailApp.sendEmail(email, subject, body);
  }
  catch(e)
  {
    Logger.log("Invalid address (" + email + "). Unable to send Report via email.")
  }
}

/**
 * Sends an emails with HTML elements.
 * 
 * @param {string} email The desired reciepent (email address) to recieve the email.
 * @param {string} subject The subject line.
 * @param {string} html_body The HTMl String for the email.
 */
function emailResults_HTML(email, subject, html_body)
{
  try{
    GmailApp.sendEmail(email, subject, '', {htmlBody: html_body});
  }
  catch(e)
  {
    Logger.log("Invalid address (" + email + "). Unable to send Report via email.")
  }
}

/**
 * A helper function that formats report headers and report data into a readable format that can be sent as an email body message.
 * @param {string[][]} headers The report headers that describe the data
 * @param {string[][]} data The report information that matches with the headers
 */
function emailBodyFormat(headers, data)
{
  var body = "This an automated message relaying the details of your latest report. If you have any questions regarding your scores, please refer to a lead or trainer. Likewise, if you're confused about your scores or want advice, please refer to a lead or trainer. \n\n"

  for(var i = 0; i < headers.length; i++)
  {
    body = body + headers[i] + ": " + data[i] + "\n\n"
  }

  return body
}

/**
 * A helper function that formats report headers and report data into a format that can be sent as an HTML body message.
 * @param {string[][]} headers The report headers that describe the data
 * @param {string[][]} data The report information that matches with the headers
 * @param string pos The position Trained on
 */
function emailBodyFormat_HTML(headers, data, pos)
{
  // Set the Style
  var style = "<head><style>table, th, td {border: 1px solid black; border-collapse: collapse; } th, td {padding: 5px; text-align: left;} td {width: 200px; text-align: right} </style></head>";

  // Start composing the HTML file
  var warning = "This an automated message relaying the details of your latest report. If you have any questions regarding your scores, please refer to a lead or trainer. Likewise, if you're confused about your scores or want advice, please refer to a lead or trainer. <br><br>"
  var body = "<!DOCTYPE html><html lang='en-US'><body><h1>Your " + pos + " Report</h1><h6>Warning: Do not reply to email!<br>" + warning + "</h6><ul>";
  
  // Start adding "Context/Identification Data (Names, etc.)"
  var i;
  for(i = 0; i < 2; i++)
  {
    body = body + "<li>" + headers[i] + ": " + data[i] + "</li>";
  }
  body = body + "</ul><br><br>"

  //for(; i < 5; i++)
  //{
    //body = body + "<h3>" + headers[i] + ": " + data[i] + "</h3>";
  //}

  // Build the HTML Table Here
  body = body + "<table><tr><th><div>" + headers[i] + "</div></th><td><div>" + data[i] + "</div></td></tr>";
  i = i + 2;  // Skip Trainee Email
  for(; i < headers.length; i++) 
  {
    body = body + "<tr><th>" + headers[i] + "</th><td>" + data[i] + "</td></tr>";
  }

  return body + "</table></body></html>";
}

/**
 * Helper Function to check off Requests and Retraining Requests that have been fullfilled. 
 * Only works for sheets where column D is being changed. Only meant to be used when column D is for boolean values.
 * @param {string} empName The name of the employee.
 * @param {string} position The position the employee was trained on.
 * @param {Sheet} target_sheet The sheet that will have the fullfilled checkbox ticked off; The Sheet that needs to be updated
 * 
 * Warning: It's column sensitive, thus any future adjustments to the position of columns in the sheets needs to be handled with care and attention. 
 * Moving columns may break the code for Request Types.
 *    - Column D: Cells in this column should have checkboxes.
 */
function proof_of_fulfilledRequest(empName, position, target_sheet)
{
  // Store the rows of the sheet into a variable
  //target_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Retraining_Requests")
  target_sheet_range = target_sheet.getDataRange().offset(1, 0, target_sheet.getLastRow() - 1)
  sortByTwoColumns(target_sheet_range, 4, 3)  // Sort rows before iterating through
  var cells = target_sheet_range.getValues()

  for(var i = 0; i < cells.length; i++) // iterate through each row in the spreadsheet
  {
    if(cells[i][3] == false && cells[i][1] == empName && cells[i][2] == position) // if the employee name and position match the ones in a row, enter
    {
      target_sheet.getRange(i+2, 4).setValue(true)    // When found, check off
      Logger.log("Help please")
      break
    }
  }
}

/**
 * A Function used to color in cells for the Skill_Board sheet. It's meant to be called everytime a retraining form is filled
 * or when an initial training form is filled. It works with the helper function, setSkillBoard(); 
 * this function is in charge of finding the row to update by searching in the first column for employeeName
 * @param {string} emp The name of the employee.
 * @param {string} position The name of the position the employee was trained on.
 * @param {Sheet} target_sheet The sheet that will be modified 
 * @param {string} color The color that will be filled in
 */
function updateSkillBoard(emp, position, target_sheet, color)
{
  var i;
  var found = false;
  target_sheet_range = target_sheet.getDataRange().offset(1, 0, target_sheet.getLastRow() - 1)
  var cells = target_sheet_range.getValues()

  for(i = 0; i < cells.length; i++)
  {
    if(cells[i][0] === emp) // if an instance of $employeeName is already inside the sheet
    {
      setSkillBoard(position, target_sheet, i+1, color)  // +1 to account for offset
      found = true;
      break
    }
  }

  //if employeeName wasn't found (meaning it wasn't in the  sheet already), add a new entry + call sortByOneColumn()
  if(!found)
  {
    valuesToAppend = [emp]
    target_sheet.getRange(target_sheet.getLastRow() + 1, 1, valuesToAppend.length).setValue(emp)
    setSkillBoard(position, target_sheet, i+1, color)
  }

  //Sort
  target_sheet.getDataRange().offset(1, 0, target_sheet.getLastRow() - 1).sort([{column: 1, ascending: true}])
}

/** 
 * Helper function for updateSkillBoard(); in charge of actually changing the color AND add dates for Reinforcement Reports
 * @param {string} employeeName The name of the employee.
 * @param {string} position The name of the position the employee was trained on.
 * @param {Sheet} target_sheet The sheet that will be modified 
 * @param {string} color The color that will be filled in
 * @param {bool} initialTraining_completed A boolean. Set to true if an "Initial_Training" form was done, false if a "Retraining" form was done
*/
function setSkillBoard(position, target_sheet, target_row, color)
{
  var cells = target_sheet.getRange(1,1,1, target_sheet.getLastColumn()).getValues();
  for(var j = 1; j <= 9; j++)
  {
    if(cells[0][j] === position) // look through header until position is found; then add dates (if neccessary) and change the color
    {
      var target_cell = target_sheet.getRange(target_row+1, j+1)
      var target_cell_color = target_cell.getBackground()
      Logger.log(target_cell.getBackground())
      // If the previous color is "blue" or "green" or "white"
      if(target_cell_color == "#0000ff" || target_cell_color == "#9900FF" || target_cell_color == "#ffffff") {
        Logger.log("Inside if statement")
        target_cell.setValue(getRelativeDate(14,0).toISOString().slice(0,9))   // Save a date: 14 days from today
        target_cell.setFontColor('white')
        target_cell.setFontStyle("bold")
        target_cell.setHorizontalAlignment("center")
      }
      else {
        target_cell.setValue("")
      }

      target_cell.setBackground(color) // +1 because we're not working with an array anymore
      return
    }
  }
}

/**
 * Function to set up Events on Google Calendars for dates to fill an employee's Reinforcement Report.
 * This function will set up three Events.
 * @param {string} empName The name of the employee.
 * @param {string} position The position the employee will need a report for.
 * @param {number} startDate The number of days in the future for the first report to be filled out.
 *        Each event will be seperated by double the wait time of the previous report (with the exception of the first one).
 */
function createEvents(empName, position, startDate)
{
  // For loop to setup multiple Calendar Events
  for(var i=0; i<=2; i++)
  {
    const start = getRelativeDate(startDate * Math.pow(2,i), 0)
    const end = getRelativeDate((startDate * Math.pow(2,i)) + 4, 0)

    // Event Object needed to Set up an Event
    let event = {
      summary: 'Reinforcement Report ' + (i+1) + ': ' + empName + ' - ' + position,
      description: 'Complete the Reinforcement Report, as part of the training process',
      start: {
        dateTime: start.toISOString()
      },
      end: {
        dateTime: end.toISOString()
      },
    };
    try {
      // call method to insert/create new event in provided calandar
      event = Calendar.Events.insert(event, calendarID);
      console.log('Event ID: ' + event.id);
    } catch (err) {
      console.log('Failed with error %s', err.message);
    }
  }
}

/**
 * Helper function to get a new Date object relative to the current date.
 * @param {number} daysOffset The number of days in the future for the new date.
 * @param {number} hour The hour of the day for the new date, in the time zone of the script.
 * 
 * @return {Date} The new date.
 */
function getRelativeDate(daysOffset, hour) {
  const date = new Date();
  date.setDate(date.getDate() + daysOffset);
  date.setHours(hour);
  date.setMinutes(0);
  date.setSeconds(0);
  date.setMilliseconds(0);
  return date;
}

/**
 * This function is responsible for handling any newly added initial training or retraining reports. THis function is designed to work with sheets 
 * named "Initial_Training" and "Retraining". This function will also check off requests for training that have been requested and are stored in the
 * spreadsheets named "Requests" and "Retraining_Requests" AND make updates to the "Skill_Chart" sheet.
 * 
 * @param {Spreadsheet} ss The Spreadsheet that stores the results from Google Forms.
 * @param {Range} range The entire range of the active sheet minus the header row.
 * 
 * Warning: It's column sensitive, thus any future adjustments to the position of columns in the sheets needs to be handled with care and attention. 
 *      Moving columns may break the code for Preliminary Training Types.
 *      - Column A: Stores Timestamps for the Google Form submission date/time
 *      - Column D: Stores Names of Trainees/Pupils
 *      - Column E: Stores the Emails of Trainees/Pupils (Should be hidden on the Google Spreadsheet view)
 *      - Column F: Stores the Preferred Language; used for translation purposes
 *      - Column G: Stores Position Names
 */
function initialHandler(ss, range)
{
  const ws = ss.getActiveSheet()

  // Move the newest entry to the top row
    sortByOneColumn(range, 1)

    // Reference the last filled row (the newest entry/top entry)
    var employeeName = JSON.stringify(ws.getRange('D2').getValues()).slice(3,-3);
    var trainedPosition = JSON.stringify(ws.getRange('G2').getValues()).slice(3,-3);

    Logger.log(employeeName)
    Logger.log(trainedPosition)

    // Set up Reminders after Initial Training Form has been completed
    if(ws.getName() == 'Initial_Training')
    {
      //Update Skill_Board sheet
      updateSkillBoard(employeeName, trainedPosition, ss.getSheetByName('Skill_Board'), "blue")
      proof_of_fulfilledRequest(employeeName, trainedPosition, ss.getSheetByName("Requests"))
      //createEvents(employeeName, trainedPosition, 14)
    }
    else // Set up Reminders after Retraining Form has been completed
    {
      //Update Skill_Board sheet
      updateSkillBoard(employeeName, trainedPosition, ss.getSheetByName('Skill_Board'), "#9900FF")
      proof_of_fulfilledRequest(employeeName, trainedPosition, ss.getSheetByName("Retraining_Requests"))
      //createEvents(employeeName, trainedPosition, 7)
    }

    // Email Details
    var subject = "Chic-Fil-A Initial Training Report: " + trainedPosition
    var headerIndex = ws.getDataRange().getValues()[0].slice(1,-1)
    var newestEntry = ws.getDataRange().getValues()[1].slice(1,-1)
    var email = JSON.stringify(ws.getRange('E2').getValues()).slice(3,-3);
    var body = emailBodyFormat_HTML(headerIndex, newestEntry, trainedPosition)

    // Translate Results if "Espanol" is saved as a result
    for(var i = 1; i < newestEntry.length; i++)
    {
      if(newestEntry[i] == 'Espanol')
        body =  LanguageApp.translate(body, 'en', 'es');    // Translate to Spanish
        break
    }

    emailResults_HTML(email, subject, body)
}

/**
 * This function is responsibile for sorting newly added training requests and retraining requests made through Google Forms; those requests
 * should be stored in Sheets named "Requests" and "Retraining_Requests". 
 * 
 * @param {Spreadsheet} ss The Spreadsheet that stores the results from Google Forms.
 * @param {Range} range The entire range of the active sheet minus the header row.
 * 
 * Warning: It's column sensitive, thus any future adjustments to the position of columns in the sheets needs to be handled with care and attention. 
 *      Moving columns may break the code. 
 *      - Column B: Store Employee Names
 *      - Column D: Will have Checkboxes inserted into the column for every new entry added
 */
function requestHandler(ss, range)
{
  const ws = ss.getActiveSheet()

    // When the new entry is added, we will add an unmarked checkbox by getting the range of the current sheet and inserting checkmark boxes
    var lastCell = ws.getRange(ws.getLastRow(), 4) // Last Row, Column D
    lastCell.insertCheckboxes()

    // Sorting Algorithm
    sortByTwoColumns(range, 4, 2)
}

/**
 * Sorts newly added Reinforcement Reports made through Google Forms. This function is also resposnible for grading new reports, emailing their 
 * respective trainees, and updating the Skill Board in accordance to the grade they recieve for their report. This function only works for Sheets 
 * named after the respective Chic-Fil-A positions: "Breading", "Buns", "Dish", "Fries", "Prep", "Rotations", Screens", & "Set-Ups"
 *   
 * @param {Spreadsheet} ss The Spreadsheet that stores the results from Google Forms.
 * @param {Range} range The entire range of the active sheet minus the header row.
 * @param {Object[][]} newestEntry An array of data that will be read; this array should come from a row orginating from a spreadsheet.
 * 
 * Warning: It's column sensitive, thus any future adjustments to the position of columns in the sheets needs to be handled with care and attention. 
 *      Moving columns may break the code. 
 */
function reinforcementRepotHandler(ss, range, newestEntry)
{
  ws = ss.getActiveSheet()

    // Begin Collecting E-mail Report Details and Credentials (text-based)
    var subject = "Chic-Fil-A Reinforcement Report: " + ws.getName()
    var headerIndex = ws.getDataRange().getValues()[0].slice(2, -1)
    var email = newestEntry[5]
    var grade = gradeReport(newestEntry, ws)                    //Grade Report and Save it on the spreadsheet
    newestEntry = ws.getDataRange().getValues()[ws.getDataRange().getValues().length-1].slice(2, -1)   // Rerun in order to get the grade results
    // var body = emailBodyFormat(headerIndex, newestEntry)    // Plain Text Version
    var body = emailBodyFormat_HTML(headerIndex, newestEntry, ws.getName()) // HTML Version

    // Set Grade color on the Skill Chart
    var color;
    if(grade >= 86)
      color = "green"
    else if(70 <= grade <= 85)
      color = "#00FF00"
    else if(51 <= grade <= 69)
      color = "orange"
    else if(25 <= grade <= 50)
      color = "#FFFF00"
    else
      color = "red"
    updateSkillBoard(newestEntry[1], ws.getName(), ss.getSheetByName('Skill_Board'), color)

    // Translate Results if "Espanol" is saved as a result
    for(var i = 1; i < newestEntry.length; i++)
    {
      if(newestEntry[i] == 'Espanol')
        body =  LanguageApp.translate(body, 'en', 'es');    // Translate to Spanish
        break
    }

    // Send results via email
    // emailResults(email, subject, body)      // Plain Text
    emailResults_HTML(email, subject, body) // HTML Version

    // Sorting Algorithm
    sortByTwoColumns(range, 4, 1)
}

