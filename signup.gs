// @Author: Jeffrey Shen (jeffreyshen13@gmail.com)
// Move Form to new account Instructions
// 1. Update all Links for the form and sheets
// 2. Run the findFormQuestionId Function and get the GHIN_QUESTION_ID
// 3. Update GHIN_QUESTION_ID and validate other variables 


// UPDATE FORM GHIN VALIDATION 
// 1. Run the updateFormValidation function
// 2. Deploy and run every 60 minutes

// UPDATE NEXT TOURNAMENT BASED ON RESPONSES
// 1. Run the processResponses function
// 2. Deploy and run every 1 minute

// TODO: 
// ADD PARAMETERS TO A GOOGLE SHEET SHEET TO EASILY UPDATE
// ADD A CACHE VALUE FOR THE ROW ON THE RESPONSES FORM 

// --------------------------------------------------------------------
//             ACCOUNT SPECIFIC SETTINGS 
// --------------------------------------------------------------------
// GOOGLE FORM 
const SIGN_UP_FORM = "https://docs.google.com/forms/d/<XXXXXXXXXXXXXXXXXXX>/edit";
// GOOGLE SHEETS
const NEXT_TOURN_LIST = "https://docs.google.com/spreadsheets/d/<XXXXXXXXXXXXXXXXXXX>/edit"
const MEMBER_LIST = "https://docs.google.com/spreadsheets/d/<XXXXXXXXXXXXXXXXXXX>/edit"
const SIGN_UP_FORM_RESPONSES = "https://docs.google.com/spreadsheets/d/<XXXXXXXXXXXXXXXXXXX>/edit"
const GHIN_QUESTION_ID = "<XXXXXXXXXXXXXXXXXXX>"


// --------------------------------------------------------------------
//             MEMBER LIST SPECIFIC DATA 
// --------------------------------------------------------------------
const NEXT_TOURN_STARTING_ROW = 3
const FORM_RESPONSES_STARTING_ROW = 2
const MEMBER_LIST_COLS = {"name": "A", "phone": "B", "email": "C", "ghin": "D"}
const RESPONSES_SHEET_COLS = {"timestamp": "A", "ghin": "B", "ride": "C", "timing":"D", "added": "E"}
const NEXT_TOURN_COLS = { "timestamp": "A", "name": "B", "ghin": "C", "handicap": "D", "ride": "E", "timing": "F", "position": "G", "email": "H"}
const MEMBER_GHIN_START = 4
const MEMBER_LAST_COL = "D"
const MEMBER_GHIN_END = 1000
const SIGN_UP_RESPONSES_VALID_COL = "A"
var MEMBERS_DICT = {}
// Javascript Object with 
// { ghin (Number) : [name (String, unformatted), phone (String), email (String), Ghin(Number)]}

// --------------------------------------------------------------------
//          GENERAL SETTINGS
// --------------------------------------------------------------------
const send_emails = (getIndicatorVal(SIGN_UP_FORM_RESPONSES, "1", "I") == "Y")
const RUNNING_CELL_COL = "J"
const RUNNING_CELL_ROW = "1"
const RUNNING_INDICATOR = "RUNNING"
const EMAIL_SUBJECT_LINE = "Tournament Sign Up Confirmation";
const INDICATOR_YES = "Y"
const INDICATOR_NO = ""
const TIMEOUT_LIMIT = 6.5;

const userObjNames = {
  "firstName": "FirstName",
  "lastName": "LastName",
  "ghin": "GHIN",
  "handicap": "Handicap",
  "email": "Email",
  "position": "Position",
  "nameFormat": "Name",
  "timestamp": "Time",
  "ride": "Ride",
  "timing": "Timing"
}


// --------------------------
// --------------------------


// ------------------------------------------
// CODE TO UPDATE FORM VALIDATION
// ------------------------------------------

function getRowRange(row, startCol, endCol) {
  return String(startCol) + row  + ":" + String(endCol) + row;
}

function getColRange(col, startRow, endRow) {
  return String(col) + String(startRow) + ":" +  String(col) + String(endRow);
}

function getMemberGhinRange() {
  return MEMBER_LIST_COLS["ghin"] + String(MEMBER_GHIN_START) + ":" + MEMBER_LIST_COLS["ghin"] + String(MEMBER_GHIN_END)
}

function getSheetRangeRaw(sheetUrl, sheetRange) {
  var sheet = SpreadsheetApp.openByUrl(sheetUrl).getActiveSheet();
  var range = sheet.getRange(sheetRange);
  var range_vals = range.getValues();
  return range_vals
}

function getSheetRangeValues(sheetUrl, sheetRange) {
  var sheet = SpreadsheetApp.openByUrl(sheetUrl).getActiveSheet();
  var range = sheet.getRange(sheetRange);
  var range_vals = range.getValues().flat();
  return range_vals
}

function createRegexFromList(vals) {
  // Given a list of values (numbers), return a regex form of the numbers ORed together in group
  // Input (vals): [1, ,02]
  // Output: "(1|02)"
  vals = vals.filter(x => x)
  var regexString = "("
  for (let i = 0; i < vals.length; i++) {
    regexString += vals[i].toString().trim()
    if (i != (vals.length - 1)) {
      regexString += "|"
    }
  }
  regexString += ")"
  console.log("REGEX VALIDATION STRING: " + regexString)
  return regexString
}

function textValidationCol(sheetRange) {
  // Given a sheet Range, Create text validation pattern for the range of numbers 
  var textValidation = FormApp.createTextValidation()
  .setHelpText('Invalid GHIN. Please make sure you are a member and typed in the correct GHIN.').requireTextMatchesPattern(createRegexFromList(getSheetRangeValues(MEMBER_LIST, sheetRange))).build();
  return textValidation
}

function updateFormValidation() {
  // Given Sign Up Form, set the validation regex for GHIN question to the numbers in the range
  var form = FormApp.openByUrl(SIGN_UP_FORM)
  var ghinQuestion = form.getItemById(GHIN_QUESTION_ID).asTextItem();
  ghinQuestion.setValidation(textValidationCol(getMemberGhinRange()))
}
// ------------------------------------------
// CODE TO UPDATE NEXT TOURNAMENT
// ------------------------------------------

function getNextRowtoUpdate(sheetLink, startingRow, indicatorCol, validCol, searchEmpty = false) {
  // Get row number of the next row to update based on 
  // Input: SheetLink - Google sheet link to look through 
  //        startingRow - row to start search at 
  //        indicatorCol - column to indicate whether row has been updated (ex: Added to next tourn sheet)
  //        validCol - column to indicate whether row is valid (not empty)
  //        searchEmpty - when true, only search for next empty row, not using indicator
  // Return index of next valid row or -1
  var responses_sheet = SpreadsheetApp.openByUrl(sheetLink).getActiveSheet();
  var not_found = true;
  while (not_found) {
    var indicator_cell = responses_sheet.getRange(indicatorCol + String(startingRow));
    var indicator_val = indicator_cell.getValue();
    var valid_row = responses_sheet.getRange(validCol + String(startingRow)).getValue() != "";
    // console.log(`Searching Next Row ${startingRow}: ${valid_row} row with indicator val ${indicator_val} `)
    if (searchEmpty) {
      if (!valid_row) {
        return startingRow;
      }
    } else {
      if (indicator_val != INDICATOR_YES && valid_row) {
        not_found = false;
        return startingRow;
      } else if (!valid_row) {
        return -1;
      }
    }
    
    startingRow++;
  }
  return -1;
}

function setIndicatorVal(sheetLink, row, col, val) {
  var sheet = SpreadsheetApp.openByUrl(sheetLink).getActiveSheet();
  sheet.getRange(String(col) + String(row)).setValues([[val]]);
}

function getIndicatorVal(sheetLink, row, col) {
  var sheet = SpreadsheetApp.openByUrl(sheetLink).getActiveSheet();
  return sheet.getRange(String(col) + String(row)).getValue();
}

function parseName(name) {
  // Given name in String "LAST, FIRST"
  // Output: ["First", "Last"]
  // "First Last"
  [last, first] = name.split(',')
  first = first || "";
  last = last || "";
  last = last.trim();
  first = first.trim();
  last = last.charAt(0).toUpperCase() + last.slice(1).toLowerCase();
  first = first.charAt(0).toUpperCase() + first.slice(1).toLowerCase();
  return [first, last]
}

function validateNotDuplicate(nextTournRow, ghin) {
  var ghinsRange = getColRange(NEXT_TOURN_COLS["ghin"], NEXT_TOURN_STARTING_ROW, nextTournRow)
  var prevGhins = getSheetRangeValues(NEXT_TOURN_LIST, ghinsRange)
  return !prevGhins.includes(Number(ghin))
}

function getDataFromMemberList(ghin) {
  // Return [<first_name>, <last_name>, <email>] for specified ghin member 
  var row = -1;
  console.log("Searching for GHIN: " + ghin)
  ghin_vals = getSheetRangeValues(MEMBER_LIST, getMemberGhinRange());
  for (var i = 0; i < ghin_vals.length; i++) {
    if (ghin_vals[i] == ghin) {
      row = MEMBER_GHIN_START + i;
      break;
    }
  }
  if (row == -1) {
    console.error("GHIN " + ghin + " was not found.")
    throw Error("GHIN " + ghin + " was not found.")
    return; 
  }
  var member_sheet = SpreadsheetApp.openByUrl(MEMBER_LIST).getActiveSheet();
  var [firstName, lastName] = parseName(member_sheet.getRange(MEMBER_LIST_COLS["name"] + String(row)).getValue());
  var email = member_sheet.getRange(MEMBER_LIST_COLS["email"] + String(row)).getValue();

  return [firstName, lastName, email]
}

function storeMemberListData() {
  var members_range = "A" + String(MEMBER_GHIN_START) + ":" + MEMBER_LAST_COL + String(MEMBER_GHIN_END)
  
  var members_range_vals = getSheetRangeRaw(MEMBER_LIST, members_range)
  members_range_vals = members_range_vals.filter(x => x[0])
  for (let i = 0; i < members_range_vals.length; i++) {
    let [name, phone, email, ghin] = members_range_vals[i];
    MEMBERS_DICT[Number(ghin)] = [String(name), String(phone), String(email), Number(ghin)]
  }
  return MEMBERS_DICT
}

function getDataFromMemberListCache(ghin) {
  if (!(Number(ghin) in MEMBERS_DICT)) {
    console.error("GHIN " + ghin + " was not found.")
    throw Error("GHIN " + ghin + " was not found.")
  }
  let [name, phone, email, _] = MEMBERS_DICT[Number(ghin)]
  let [firstName, lastName] = parseName(name);
  return [firstName, lastName, email]
}

function getDataFromFormResponse(row) {
  // Get [<timestamp>, <ghin>, <handicap>, <ride>, <timing>, <added>]
  var responsesSheet = SpreadsheetApp.openByUrl(SIGN_UP_FORM_RESPONSES).getActiveSheet();
  var ghin = responsesSheet.getRange(RESPONSES_SHEET_COLS["ghin"] + String(row)).getValue();
  var timestamp = responsesSheet.getRange(RESPONSES_SHEET_COLS["timestamp"] + String(row)).getValue();
  var ride = responsesSheet.getRange(RESPONSES_SHEET_COLS["ride"] + String(row)).getValue();
  var timing = responsesSheet.getRange(RESPONSES_SHEET_COLS["timing"] + String(row)).getValue();
  var handicap = "";
  var added = responsesSheet.getRange(RESPONSES_SHEET_COLS["added"] + String(row)).getValue();
  return [timestamp, ghin, handicap, ride, timing, added]
}

function writeToNextTourn(row, data) {
  // [timestamp,	Name,	GHIN,	Handicap,	Ride,	timing,	Position,	Email Sent] write to 
  var nextTournSheet = SpreadsheetApp.openByUrl(NEXT_TOURN_LIST).getActiveSheet();
  nextTournSheet.getRange(getRowRange(row, NEXT_TOURN_COLS["timestamp"], NEXT_TOURN_COLS["email"])).setValues([data]);
}


function formatRideWalk(ride) {
  switch (ride) {
    case "Ride":
      return "R"
    case "Walk":
      return "W"
    default: 
      return ""
  }

}

// Run function 
// Every X (5) minutes, look through responses and update rows that haven't been sent

//
// For each new row that hasn't been sent, 
  // catch all errors and write an error if row errored 
  // Update the Next Tournament sheet with the next empty row to this current row
    // Fill in the Name,	GHIN,	Handicap	Ride or Walk	Early or Late	Position
      // Get Name from members db using GHIN
      // Get GHIN from form response
      // get handicap from ???
      // Get RIde or walk from form response
      // get early or late from form response
    // Update the current position on the sheet 
    // Send out email with the current position as well as any addition information 

function processResponseRow(row, nextTournRow = -1) {
  var [timestamp, ghin, handicap, ride, timing, added] = getDataFromFormResponse(row);
  var [firstName, lastName, email] = getDataFromMemberListCache(ghin);
  var nameFormat = firstName.toUpperCase() + " " + lastName.toUpperCase()
  var ride = formatRideWalk(ride)
  console.log(`Retrieved Data from row: ${timestamp}, ${ghin}, ${ride}, ${timing}, ${nameFormat}, ${email}`)
  if (nextTournRow == -1) {
    nextTournRow = getNextRowtoUpdate(NEXT_TOURN_LIST, NEXT_TOURN_STARTING_ROW, NEXT_TOURN_COLS["timestamp"], NEXT_TOURN_COLS["timestamp"], true)
  }
  position = (nextTournRow - NEXT_TOURN_STARTING_ROW + 1).toString()
  email_sent = INDICATOR_NO
  user_obj = {
    "FirstName": firstName,
    "LastName": lastName,
    "GHIN": ghin.toString(),
    "Handicap": handicap.toString(),
    "Email": email,
    "Position": position,
    "Name": nameFormat,
    "Time": timestamp.toString(),
    "Ride": ride,
    "Timing": timing
  }
  if (!validateNotDuplicate(nextTournRow, ghin)) {
    console.warn("Duplicate Submission for GHIN: " + ghin.toString())
    return ["WARNING: Duplicate Submission", nextTournRow]
  }
  // TODO: Send emails if not filled in already 
  if (send_emails) {
    email_sent = sendEmailRow(user_obj)
  }
  var data = [timestamp, nameFormat, ghin, handicap, ride, timing, position, email_sent]
  writeToNextTourn(nextTournRow, data);
  console.log(`Wrote to Next Tourn Row ${nextTournRow.toString()}: ${data[0].toString()}, ${data[1]}, ${data[2].toString()}, ${data[3].toString()}, ${data[4]}, ${data[5]}, ${data[6].toString()}`)
  return [INDICATOR_YES, nextTournRow]
}

function ableToRun() {
  var runCell = getIndicatorVal(SIGN_UP_FORM_RESPONSES, RUNNING_CELL_ROW, RUNNING_CELL_COL);
  if (runCell == "") {
    return true
  } else {
    runCell = new Date(runCell);
    if (Date.now() - runCell.getTime() > TIMEOUT_LIMIT * 1000 * 60) {
      return true
    }
  }
  return false;
}

function processResponses() {
  if (ableToRun()) {
    setIndicatorVal(SIGN_UP_FORM_RESPONSES, RUNNING_CELL_ROW, RUNNING_CELL_COL, Date.now())
  } else {
    console.log("Another Instance Currently Running...Stopping.")
    return;
  }
  try {
    // Cache the Member list in memory 
    storeMemberListData();
    var row = getNextRowtoUpdate(SIGN_UP_FORM_RESPONSES, FORM_RESPONSES_STARTING_ROW, RESPONSES_SHEET_COLS["added"], RESPONSES_SHEET_COLS["timestamp"]);
    var nextTournRow = -1;
    while (row != -1) {
      console.log(`Processing Response Row ${row}:`)
      try {
        [message, nextTournRow] = processResponseRow(row, nextTournRow);
        if (message == INDICATOR_YES) {
          nextTournRow += 1;
        }
        setIndicatorVal(SIGN_UP_FORM_RESPONSES, row, RESPONSES_SHEET_COLS["added"], message);
      } catch (e) {
        console.log("Caught Error while Processing Reponse Row " + row)
        console.error(e);
        setIndicatorVal(SIGN_UP_FORM_RESPONSES, row, RESPONSES_SHEET_COLS["added"], "ERROR: " + e.message)
      }
      
      row = getNextRowtoUpdate(SIGN_UP_FORM_RESPONSES, row + 1, RESPONSES_SHEET_COLS["added"], RESPONSES_SHEET_COLS["timestamp"]);
    }
    setIndicatorVal(SIGN_UP_FORM_RESPONSES, RUNNING_CELL_ROW, RUNNING_CELL_COL, "")
  } catch(e) {
    setIndicatorVal(SIGN_UP_FORM_RESPONSES, RUNNING_CELL_ROW, RUNNING_CELL_COL, "")
    console.error("Error: " + e.message)
    throw e;
  }
  
}


// On every new tournament 
// reset the data in the sheet 


// ---------------------=====================================================


function sendEmailRow(userObj) {
  // Given a userObj with user data,
  //  including: "Email", and any template data 
  // Return the time email was sent or any error message 
  console.log("Sending Email to " + userObj["Email"] + " for position " + userObj["Position"])
  const emailTemplate = getGmailTemplateFromDrafts_(EMAIL_SUBJECT_LINE);;
  try {
    const msgObj = fillInTemplateFromObject_(emailTemplate.message, userObj);
    GmailApp.sendEmail(userObj["Email"], msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html,
          // bcc: 'a.bcc@email.com',
          // cc: 'a.cc@email.com',
          // from: 'an.alias@email.com',
          // name: 'name of the sender',
          // replyTo: 'a.reply@email.com',
          // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages
        });
    return new Date();
  } catch(e) {
    console.error("ERROR: Could not send email due to - " + e.message)
    return "ERROR: " + e.message
  }
}

/**
 * Get a Gmail draft message by matching the subject line.
 * @param {string} subject_line to search for draft message
 * @return {object} containing the subject, plain and html message body and attachments
*/
function getGmailTemplateFromDrafts_(subject_line){
  try {
    // get drafts
    const drafts = GmailApp.getDrafts();
    // filter the drafts that match subject line
    const draft = drafts.filter(subjectFilter_(subject_line))[0];
    // get the message object
    const msg = draft.getMessage();

    // Handles inline images and attachments so they can be included in the merge
    // Based on https://stackoverflow.com/a/65813881/1027723
    // Gets all attachments and inline image attachments
    const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true,includeAttachments:false});
    const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
    const htmlBody = msg.getBody(); 

    // Creates an inline image object with the image name as key 
    // (can't rely on image index as array based on insert order)
    const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj) ,{});

    //Regexp searches for all img string positions with cid
    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
    const matches = [...htmlBody.matchAll(imgexp)];

    //Initiates the allInlineImages object
    const inlineImagesObj = {};
    // built an inlineImagesObj from inline image matches
    matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

    return {message: {subject: subject_line, text: msg.getPlainBody(), html:htmlBody}, 
            attachments: attachments, inlineImages: inlineImagesObj };
  } catch(e) {
    throw new Error("ERROR: Can't find Gmail draft");
  }

  /**
   * Filter draft objects with the matching subject linemessage by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} GmailDraft object
  */
  function subjectFilter_(subject_line){
    return function(element) {
      if (element.getMessage().getSubject() === subject_line) {
        return element;
      }
    }
  }
}

/**
 * Fill template string with data object
 * @see https://stackoverflow.com/a/378000/1027723
 * @param {string} template string containing {{}} markers which are replaced with data
 * @param {object} data object used to replace {{}} markers
 * @return {object} message replaced with data
*/
function fillInTemplateFromObject_(template, data) {
  // We have two templates one for plain text and the html body
  // Stringifing the object means we can do a global replace
  let template_string = JSON.stringify(template);
  // Token replacement
  template_string = template_string.replace(/{{[^{}]+}}/g, key => {
    return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
  });
  return  JSON.parse(template_string);
}

/**
 * Escape cell data to make JSON safe
 * @see https://stackoverflow.com/a/9204218/1027723
 * @param {string} str to escape JSON special characters from
 * @return {string} escaped string
*/
function escapeData_(str) {
  return str
    .replace(/[\\]/g, '\\\\')
    .replace(/[\"]/g, '\\\"')
    .replace(/[\/]/g, '\\/')
    .replace(/[\b]/g, '\\b')
    .replace(/[\f]/g, '\\f')
    .replace(/[\n]/g, '\\n')
    .replace(/[\r]/g, '\\r')
    .replace(/[\t]/g, '\\t');
};
 


// -------- TESTING FUNCTIONS --------------------------------
function findFormQuestionId() {
  var form = FormApp.openByUrl(SIGN_UP_FORM)

  var formQuestions = form.getItems()
  const types = formQuestions.map((item) => item.getType().name());
  for (let i = 0; i < formQuestions.length; i++) {
    console.log(formQuestions[i].getId())
  }
  console.log(types)
}


