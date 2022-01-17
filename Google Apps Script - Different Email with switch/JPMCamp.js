// Excel fields columns counting from 1
var email_position = 3;
var firstname_position = 2;
var secondname_position = 1;
var team_position = 4;
var yesno_position = 5;
var sheet_name="Sheet1";

// Email
var EMAIL_SUBJECT = "Proiecte BOS! 💚";

// Main function
function sendEmail() {

  // ACTIVATE THE "Sheet1" SHEET AND MAKE IT THE PRINCIPAL ONE
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).activate();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // GET THE INDEX OF THE LAST ROW
  var last_row = spreadsheet.getLastRow();
  var count_end = 0;

  // Check if we can still send emails today
  if(check_emails_status() == 1) {
    return;
  }

  // SEND AN EMAIL TO EVERY PARTICIPANT
  for(var participant_row = 2; participant_row <= last_row; ++participant_row)
  {
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
  return spreadsheet.getRange(participant_row, firstname_position).getValue() + " " + spreadsheet.getRange(participant_row, secondname_position).getValue();
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


// Send the email to each person
function send_email(spreadsheet, participant_row) {
  var template_default = get_body_template();
  var template_jpm = get_body_template_jpm();
  template_default.replace("[[NAME]]", get_name(spreadsheet, participant_row));
  template_jpm.replace("[[NAME]]", get_name(spreadsheet, participant_row));


  switch(spreadsheet.getRange(participant_row, team_position).getValue().toString()) {
    case "PROUD":
      template_default.replace("[[NUMEPROIECT]]", "PRoud");
      deliver_email(spreadsheet, participant_row, template_default);
      break;
    case "Upgrade":
      template_default.replace("[[NUMEPROIECT]]", "UPgrade");
      deliver_email(spreadsheet, participant_row, template_default);
      break;
    case "IT":
      template_default.replace("[[NUMEPROIECT]]", "IT is Business");
      deliver_email(spreadsheet, participant_row, template_default);
      break;
    case "EVENTUM":
      template_default.replace("[[NUMEPROIECT]]", "EVENTUM");
      deliver_email(spreadsheet, participant_row, template_default);
      break;
    case "IT PM":
      template_default.replace("[[NUMEPROIECT]]", "IT is Business");
      deliver_email(spreadsheet, participant_row, template_default);
      break;
    case "IT JPM":
      template_jpm.replace("[[NUMEPROIECT]]", "IT is Business");
      deliver_email(spreadsheet, participant_row, template_jpm);
      break;
    case "UPGRADE PM":
      template_default.replace("[[NUMEPROIECT]]", "UPgrade");
      deliver_email(spreadsheet, participant_row, template_default);
      break;
    case "UPGRADE JPM":
      template_jpm.replace("[[NUMEPROIECT]]", "UPgrade");
      deliver_email(spreadsheet, participant_row, template_jpm);
      break;
    case "PROUD PM":
      template_default.replace("[[NUMEPROIECT]]", "PRoud");
      deliver_email(spreadsheet, participant_row, template_default);
      break;
    case "PROUD JPM":
      template_jpm.replace("[[NUMEPROIECT]]", "PRoud");
      deliver_email(spreadsheet, participant_row, template_jpm);
      break;
    case "EVENTUM PM":
      template_default.replace("[[NUMEPROIECT]]", "EVENTUM");
      deliver_email(spreadsheet, participant_row, template_default); 
      break;
    case "EVENTUM JPM":
      template_jpm.replace("[[NUMEPROIECT]]", "EVENTUM");
      deliver_email(spreadsheet, participant_row, template_jpm); 
      break;
    default:
      Logger.log("Couldn't identify body for " + get_name(spreadsheet, participant_row));
      return;
  }
}

function deliver_email(spreadsheet, participant_row, body) { 
    MailApp.sendEmail({
      to: get_email(spreadsheet, participant_row),
      subject: EMAIL_SUBJECT,
      body: "",
      htmlBody: body,
    });
}

// Create email body
function get_body_template() {
  return "Bună, [[NAME]]! 👋<br><br>După cum știi, proiectele sunt o parte foarte importantă a organizației noastre, astfel că în fiecare an ne bucurăm să avem bosuleți cu inițiativă și idei complexe! În urma repartizării pe echipe și am decis că cel mai mult te-ai potrivi în echipa proiectului [[NUMEPROIECT]]! 🥳<br><br>De-abia așteptăm să vedem conceptul acestui an și subiectele pe care le veți aborda pentru proiectul [[NUMEPROIECT]]! <br><br><b>No strangers here! Only friends you've never met! 💚<\/b>";
}

// Create email body
function get_body_template_jpm() {
  return "Bună, [[NAME]]! 👋<br><br>După cum știi, proiectele sunt o parte foarte importantă a organizației noastre, astfel că în fiecare an ne bucurăm să avem bosuleți cu inițiativă și idei complexe! Ne-am bucurat foarte mult să te vedem implicându-te în JPM Camp, ai avut atât inițiativă, cât și idei bune. 🥳<br><br>Ți-am admirat curajul și dorința de pune în evidență perspectivă, dar și faptul că știai când să-ți spui punctul vedere, sau când trebuia să-i lași pe alții să o facă, așa că felicitări, ocupi postul de JPM în cadrul [[NUMEPROIECT]]!<br><br><b>No strangers here! Only friends you've never met! 💚<\/b>";
}
