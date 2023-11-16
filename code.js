// Google Calendar Id, needed to create events for a specific calendar
const calendarID = ''
// Link to Google Colab
const colab_URL = ''

/**
 * A function to automically sort the rows in a spreadsheet. It's column sensitive, thus any future adjustments to the position
 *  of columns in the sheets needs to be handled with care and attention. Moving columns won't break the code, but it will distort 
 *  the order of rows as specific rows are used to sort an entire column. 
 * Additionally, this function also acts as a main function to set up Events on Google Calenders (for Reinforcement Report due dates),
 *  send emails, and general updates to sister/related spreadsheets (if one sheet get's updated, update another sheet)
 */
function autoSort() {

  // Get current active sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const ws = ss.getActiveSheet()
  const range = ws.getDataRange().offset(1, 0, ws.getLastRow() - 1)
  
  // If the active sheet is a for Reinforcement Reports; enter
  if(!(ws.getName() == 'Requests' || ws.getName() == 'Retraining_Requests' || ws.getName() == 'Skill_Board' || ws.getName() == 'Initial_Training' || ws.getName() == 'Retraining'))
  {
    // E-mail Report Details (text-based)
    var subject = "Chic-Fil-A Reinforcement Report: " + ws.getName()
    var newestEntry = ws.getDataRange().getValues()[ws.getDataRange().getValues().length-1]
    var headerIndex = ws.getDataRange().getValues()[0]
    var email = newestEntry[5]
    var body = emailBodyFormat(headerIndex, newestEntry)

    // Translation Option
    var languagePref = newestEntry[6];
    if(languagePref == 'Espanol')
    {
      body =  LanguageApp.translate(body, 'en', 'es');    // Translate to Spanish
    }

    //Grade Report and Save it on the spreadsheet
    var grade = gradeReport(ws)
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
    updateSkillBoard(newestEntry[2], ws.getName(), ss.getSheetByName('Skill_Board'), color)

    // Send results via email
    emailResults(email, subject, body)

    // Sorting Algorithm
    sortByTwoColumns(range, 4, 1)
  }
  else if(ws.getName() == "Requests" || ws.getName() =="Retraining_Requests") //else if it's for a request; enter
  {
    // When the new entry is added, we will add an unmarked checkbox by getting the range of the current sheet and inserting checkmark boxes
    var lastCell = ws.getRange(ws.getLastRow(), 4) // Last Row, Column D
    lastCell.insertCheckboxes()

    // Sorting Algorithm
    sortByTwoColumns(range, 4, 3)
  }
  else if(ws.getName() == 'Initial_Training' || ws.getName() == 'Retraining') //else if the ws is Initial_Training or Retraining; enter
  {
    // Move the newest entry to the top row
    sortByOneColumn(range, 1)

    // Reference the last filled row (the newest entry/top entry)
    var employeeName = JSON.stringify(ws.getRange('B2').getValues()).slice(3,-3);
    var trainedPosition = JSON.stringify(ws.getRange('D2').getValues()).slice(3,-3);

    // Set up Reminders after Initial Training Form has been completed
    if(ws.getName() == 'Initial_Training')
    {
      //Update Skill_Board sheet
      updateSkillBoard(employeeName, trainedPosition, ss.getSheetByName('Skill_Board'), "blue")
      proof_of_fulfilledRequest(employeeName, trainedPosition, ss.getSheetByName("Requests"))
      createEvents(employeeName, trainedPosition, 14)
    }
    else // Set up Reminders after Retraining Form has been completed
    {
      //Update Skill_Board sheet
      updateSkillBoard(employeeName, trainedPosition, ss.getSheetByName('Skill_Board'), "#9900FF")
      proof_of_fulfilledRequest(employeeName, trainedPosition, ss.getSheetByName("Retraining_Requests"))
      createEvents(employeeName, trainedPosition, 7)
    }
  }
}


/**
 * Function is broken
 */
function gradeReport(ws)
{
  var newestEntry = ws.getDataRange().getValues()[ws.getDataRange().getValues().length-1]
  var score = 0
  var questions = 0

  //Loop through the row: archaic grading criteria (warning  this  can  easily be abused if a text based question is given "Yes" as a response)
  for(var i = 0; i < newestEntry.length; i++)
  {
    if(newestEntry[i] === "Yes")
    {
      score = score +  5
      questions = questions + 1
    }
    else if(1 <= parseInt(newestEntry[i]) && 5 >= parseInt(newestEntry[i]))
    {
      score = score + parseInt(newestEntry[i])
      questions = questions + 1
    }
  }

  //Update Grade 
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
 * @param {} range The range of cells you wish to sort. 
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
 * A function that sends an email.
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
    Logger.log("Invalid address")
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
 * Helper Function to check off Requests and Retraining Requests that have been fullfilled. 
 * Only works for sheets where column D is being changed. Only meant to be used when column D is for boolean values.
 * @param {string} empName The name of the employee.
 * @param {string} position The position the employee was trained on.
 * @param {Sheet} target_sheet The sheet that will have the fullfilled checkbox ticked off; The Sheet that needs to be updated
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
    //Logger.log(cells[i][3])
    //Logger.log(position)
    if(cells[i][1] == empName && cells[i][2] == position && cells[i][3] == false) // if the employee name and position match the ones in a row, enter
    {
      //Logger.log(cells[i])
      target_sheet.getRange(i+2, 4).setValue(true)
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
 * Helper function for updateSkillBoard(); in charge of actually changing the color
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
    if(cells[0][j] === position) // look through header until position is found
    {
      target_sheet.getRange(target_row+1, j+1).setBackground(color) // +1 because we're not working with an array anymore
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
 * @param {number} hour The hour of the day for the new date, in the time zone
 *     of the script.
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

/*
function test()
{
  createEvents("Itsuki", "Rotations")
}*/

/*
function onOpen(){
  var ui = SpreadsheetApp.getUi()
  var menu = ui.createMenu("Custom Menu")
    menu.addItem('Calculate Overall Score', 'calcOverall')
    menu.addItem('Delete Data', 'delete')
    menu.addSeparator()
    menu.addItem('Link to Google Colab', 'linkTo')
}

function calcOverall(){
  var ui = SpreadsheetApp.getUi()
  var myName = ui.prompt("Type in an Employee Name (Last_Name, First_Name)")
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  
  //Loop through every sheet
  for(var i = 0; i < sheets.length; i++)
  {
    var sheet = sheets[i]
    Logger.log("Domain Expansion")

  }

}

function linkTo(){

}
;
*/
