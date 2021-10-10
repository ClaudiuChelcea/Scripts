// Excel fields columns counting from 1
var email_position = 19;
var firstname_position = 1;
var yesno_position = 20;
var sheet_name="Raspunsuri Board";
var status = 12;
var programat_cand_cell = 14;

// Email
var EMAIL_SUBJECT_REMINDER = "Reminder interviu de grup BOS România 💚";
// Please modify the body yourself in HTML code, LINE 170


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
    if(spreadsheet.getRange(participant_row, status).getValue().toString().localeCompare("APROBAT") == 0)
      send_reminder(spreadsheet, participant_row);
    else if(spreadsheet.getRange(participant_row, status).getValue().toString().localeCompare("REFUZAT") == 0)
    {
      spreadsheet.getRange(participant_row, yesno_position).setValue("YES");
      continue;
    }
    else
    {
      Logger.log("Status not finished for " + get_name(spreadsheet, participant_row) + "!");
      count_end = count_end + 1;
      if(count_end >= 3)
      {
        Logger.log("Task finished! Add more entries to the excel sheet or finish the aproval / denial of the current ones before continuing!");
        break;
      }         
      continue;
    }

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

// Send the reminder
function send_reminder(spreadsheet, participant_row)
{
  MailApp.sendEmail({
      to: get_email(spreadsheet, participant_row),
      subject: EMAIL_SUBJECT_REMINDER,
      body: "",
      htmlBody: get_reminder_body(get_name(spreadsheet, participant_row), get_programare(spreadsheet, participant_row)),
    });
}

// Get person's name
function get_programare(spreadsheet, participant_row) {
  //                                   FIRST NAME                                       
  return spreadsheet.getRange(participant_row, programat_cand_cell).getValue();
}

// Create email body
function get_reminder_body(NAME, PROGRAMARE) {
  return "Bună, " + NAME + "! 👋<br><br>Ne bucurăm pentru interesul tău pentru BOS și garantăm că o să ai o studenție de succes alături de noi, de comunitatea noastră și de lucrurile pe care o să le înveți aici! 🤩<br><br>Încă o dată, felicitări pentru trecerea în etapa următoare și anume interviurile de grup! 🥳<br><br>Cum ți-a fost povestit și la telefon, interviul de grup este doar un joculeț în care să ne putem cunoaște mai bine, așa că stai fără grijă!<br><br>Revenim către tine cu link-ul de ZOOM: <br><br>Te rugăm să fii prezent/ă pe PC / laptop în intervalul " + PROGRAMARE + " ,interval stabilit telefonic cu tine.<br>Dacă au apărut oricare schimbări de program și ai nevoie de o reprogramare, nu ezita să dai un reply la acest email.<br><br>Noi deabia așteptăm să te cunoaștem!<br><br><b>No strangers here! Only friends you've never met! 💚<\/b>";
}
