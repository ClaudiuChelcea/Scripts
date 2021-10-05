// Excel fields columns counting from 1
var email_position = 31;
var firstname_position = 2;
var yesno_position = 36;
var sheet_name="Sheet1";

// Email
var EMAIL_SUBJECT = "Primii pași către BOS! 💚";
// Please modify the body yourself in HTML code, LINE 170


// Main function
function sendEmail() {

  // ACTIVATE THE "Sheet1" SHEET AND MAKE IT THE PRINCIPAL ONE
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).activate();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // GET THE INDEX OF THE LAST ROW
  var last_row = spreadsheet.getLastRow();

  // Check if we can still send emails today
  if(check_emails_status() == 1) {
    return;
  }

  // SEND AN EMAIL TO EVERY PARTICIPANT
  for(var participant_row = 2; participant_row <= last_row; ++participant_row) {

    // Check if we didn't already send an email to that person
    if(check_already_sent(spreadsheet, participant_row) == 1) {
      continue;
    }

    // Check duplicate
    if(participant_row < last_row && get_email(spreadsheet,participant_row).localeCompare(get_email(spreadsheet,participant_row+1)) == 0)
      continue;
    
    // Check if we have emails left to send in the google's limits
    if(MailApp.getRemainingDailyQuota() <= 0) {
      log_errors();
      break;
    }

    // Send the emails
    send_email(spreadsheet, participant_row);

    // Mark the person as someone who has received the email to not receive another one
    spreadsheet.getRange(participant_row, yesno_position).setValue("YES");

    // Mark duplicates as sent as well
    var nr_of_duplicates = count_duplicate(spreadsheet, participant_row, last_row);
    for(var i=1; i<=nr_of_duplicates;i++) {
        var get_value = spreadsheet.getRange(participant_row,yesno_position).getValue();
        spreadsheet.getRange(participant_row - i, yesno_position).setValue(get_value);
      }

  } // end loop
}


// Check how many emails we can send today
// The limit for Google Scripts is 100 emails / 24h!
function check_emails_left() {
  var EMAILS_ABLE_TO_SEND_TODAY_LEFT = MailApp.getRemainingDailyQuota();
  Logger.log("Emails left for today: " + EMAILS_ABLE_TO_SEND_TODAY_LEFT.toString());
  return EMAILS_ABLE_TO_SEND_TODAY_LEFT;
}


// Check if we can still send emails today
function check_emails_status() {
  var EMAILS_ABLE_TO_SEND_TODAY_LEFT = check_emails_left();
  if(EMAILS_ABLE_TO_SEND_TODAY_LEFT == 0) {
    Logger.log("Warning. No more emails to send today! Limit reached!");
    Browser.msgBox("Warning. No more emails to send today! Limit reached!");
    return 1;
  }
  else {
    return 0;
  }
}


// Get person's email
function get_email(spreadsheet, participant_row) {
  //                               ROW        , COLUMN
  return spreadsheet.getRange(participant_row, email_position).getValue();
}


// Get person's name
function get_name(spreadsheet, participant_row) {
  //                                   FIRST NAME                                       
  return spreadsheet.getRange(participant_row, firstname_position).getValue();
}


// Dont send the same email to the same person
function check_already_sent(spreadsheet, participant_row) {
  var bool_alreadySentEmail = spreadsheet.getRange(participant_row, yesno_position).getValue();
  if(bool_alreadySentEmail.toString().toUpperCase() == "YES") {
    Logger.log("Skipped " + spreadsheet.getRange(participant_row, email_position).getValue());
    return 1.
  }
  else {
    Logger.log("Sending to " + spreadsheet.getRange(participant_row, email_position).getValue());
    return 0;
  }
}


// Create email subject
function get_subject() {
  return EMAIL_SUBJECT;
}


// Display error if we can no longer send emails today
function log_errors() {
  Logger.log("Warning. No more emails to send today! Limit reached! The person who is yet to receive the email is " + NAME + "!");
  Browser.msgBox("Warning. No more emails to send today! Limit reached! The person who is yet to receive the email is " + NAME + "!");
}


// Send the email to each person
function send_email(spreadsheet, participant_row, img1, img2, pdf1) {
  MailApp.sendEmail({
      to: get_email(spreadsheet, participant_row),
      subject: get_subject(),
      body: "",
      htmlBody: get_body(get_name(spreadsheet, participant_row)),
    });
}


// Check for duplicate entry in the excel sheet
function count_duplicate(spreadsheet, participant_row, last_row) {
  var duplicates = 0;
  var counter = 1;
  for(var i = participant_row; i >=0 ; i--)
    if(get_email(spreadsheet, participant_row).localeCompare(get_email(spreadsheet, participant_row - counter)) == 0) {
      duplicates++;
      counter++;
    }
    else
      break;
  
  return duplicates;
}


// Create email body
function get_body(NAME) {
  return "Bună, " + NAME + "! 👋<br><br>Venim cu veșți bune! Ne place de ține și te vrem în BOS! 😎<br><br>Pentru asta, oferă-ne în reply la acest email numărul tău de telefon pentru a te putea felicita verbal și a-ti comunica etapele următoare! 🥳<br><br><b>No strangers here! Only friends you've never met! 💚<\/b>";
}
