// Global variables
// For anything with _ID, go to drive, "Get link" and get the link between /d/.../view?. We need what`s in the the '...'.
var PDF_FOLDER_ID = "18ZzQfaU0YxLZ5XrxQZOGYQ_GWx8TyNxs"; // Folder to save PDFs, mandatory - they will be automatically sent, but we need to save them
var TMP_FOLDER_ID = "14MW3KBShpu3ufGm8GIkRqnG5HeZ1NPj6"; // Folder to save TMP files, mandatory - they will be automatically sent, but we need to save them
var TEMPLATE_DOCS_FILE_ID = "1UbEriXT6176xUKClGSdU6uG9W3zV7jl9"; // Docs file to create custom invoice from
var sheet_name = "Test"; // The sheet name from the opened excel

// Excel column number for these items
var FIRST_NAME = 2;
var SECOND_NAME = 3;
var email_position = 7;
var PARTICIPANT_ADDRESS = 9;
var PARTICIPANT_CITY = 10;
var PARTICIPANT_POSTAL_CODE = 11;
var PARTICIPANT_COUNTRY = 12;
var PARTICIPANT_ID = 29;
var yesno_position = 30;
var CURRENT_DATE = 31;

// Images or logos
var IMAGE_1_ID = "1TWsaebwhSj3rL6CwxgTPfvt1DTpVUlyN";
var IMAGE_2_ID = "1k2LrD3heWdqNe1Yzo7ls12suVDjKmjVL";

// Email content
var EMAIL_SUBJECT = "Invoice Email";


/*
  * Attention! *
  Keep the excel file clean! Have no extra rows! No side comments (no extra text in any cell, you can have comments of course)
  You still have to modify the get_body(..) function!
*/

// Main function
function main()
{
  // Open spreadsheet
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).activate();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // GET THE INDEX OF THE LAST ROW
  var last_row = spreadsheet.getLastRow();

  // Get images
  var img1 =  DriveApp.getFileById(IMAGE_1_ID);
  var img2 = DriveApp.getFileById(IMAGE_2_ID);

  // Generate PDF for every participant
  for(var participant_row = 2; participant_row <= last_row; ++participant_row) {
    // Check if we didn't already send an email to that person
    if(check_already_sent(spreadsheet, participant_row) == 1) {
        continue;
    }

    // Check duplicate
    if(participant_row < last_row && get_email(spreadsheet, participant_row).localeCompare(get_email(spreadsheet,participant_row+1)) == 0)
      continue;

    // Generate PDF
    generatePDF_sendEmail(spreadsheet, participant_row, img1, img2);

    // Mark the person as someone who has received the email to not receive another one
    spreadsheet.getRange(participant_row, yesno_position).setValue("YES");

    // Mark duplicates as sent as well
    var nr_of_duplicates = count_duplicate(spreadsheet, participant_row);
    for(var i=1; i<=nr_of_duplicates;i++) {
        var get_value = spreadsheet.getRange(participant_row,yesno_position).getValue();
        spreadsheet.getRange(participant_row - i, yesno_position).setValue(get_value);
    }
  }
}


// Get person's email
function get_email(spreadsheet, participant_row) {
  //                               ROW        , COLUMN
  return spreadsheet.getRange(participant_row, email_position).getValue();
}

// Create email subject
function get_subject() {
  return EMAIL_SUBJECT;
}


// Dont send the same email to the same person
function check_already_sent(spreadsheet, participant_row)
{
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


// Return person's name
function get_name(spreadsheet, participant_row) {
  return spreadsheet.getRange(participant_row, FIRST_NAME).getValue().toString().trim() + " " + spreadsheet.getRange(participant_row, SECOND_NAME).getValue().toString().trim();
}


function convertDocx(id, file_name) {
  // Using the normal drive service to get the blob (binary data)
  const docx = DriveApp.getFileById(id);
  const blob = docx.getBlob();

  // Creating a new file
  const newDoc = Drive.newFile();

  // Setting the title
  newDoc.title = file_name;

  // Converting the docx file to GDoc
  const newGDoc = Drive.Files.insert(newDoc, blob, {convert:true});

  // Return the new id
  return newGDoc.id;
}


// Create email body (HMTL)
function get_body(NAME) {
  return "Dear " + NAME +",<br><br>Enclosed you will find the payment request. Please pay the participation fee within 7 working days in order to confirm your participation.<br> Your participation can only be fixed against payment.<br><br>Once your transfer has been done, please send us a justification of the same (bank document scanned, print-screen...) to our email<br>INSERT_COMPANY_EMAIL@gmail.com with the following subject: \"Participation fee Event_NAME - " +  NAME + "\".<br><br>We will process all requests for collective invoices in the upcoming days<br>An event guide will be sent to you prior to the event to guarantee a pleasant participation.<br><br>For further information please visit our website https::randomWEBSITE.com.<br><br>If you have any questions, please do not hesitate to contact us via email at INSERT_COMPANY_EMAIL@gmail.com.<br><br>Looking forward to meeting you very soon,<br><br>Kind regards,<br>Company Name <br><br><img src=\"cid:sampleImage\" width=125 height=125><img src=\"cid:sampleImage2\" width=125 height=125>";
}


// Send the email to each person
function send_email(spreadsheet, participant_row, logo_BOS, logo_JEE_summer, invoice_pdf) {
  MailApp.sendEmail({
      to: get_email(spreadsheet, participant_row),
      subject: get_subject(),
      body: "",
      htmlBody: get_body(get_name(spreadsheet, participant_row)),
      inlineImages: {sampleImage: logo_BOS.getBlob(), sampleImage2: logo_JEE_summer.getBlob()},
      attachments:[invoice_pdf.getAs(MimeType.PDF)]
    });
}


// Generate custom PDF for every person
function generatePDF_sendEmail(spreadsheet, participant_row, img1, img2)
{
  // Get files
  var pdfFolder = DriveApp.getFolderById(PDF_FOLDER_ID);
  var tmpFolder = DriveApp.getFolderById(TMP_FOLDER_ID);
  var templateDocs = DriveApp.getFileById(TEMPLATE_DOCS_FILE_ID);

  // Make temporary files in the tmpFolder
  var newTMP_File = templateDocs.makeCopy(tmpFolder);
  var fn_value = spreadsheet.getRange(participant_row,FIRST_NAME).getValue().toString().trim();
  var sn_value = spreadsheet.getRange(participant_row, SECOND_NAME).getValue().toString().trim();
  var file_name = "JEE Summer Conference 2021 Invoice for " + fn_value + " " + sn_value;
  newTMP_File.setName(file_name);

  // Get body
  var opened_Docs = DocumentApp.openById(convertDocx(newTMP_File.getId(), file_name));
  tmpFolder.removeFile(newTMP_File);
  var body = opened_Docs.getBody();

  // Execute replacements
  var var_name = spreadsheet.getRange(participant_row,FIRST_NAME).getValue().toString().trim();
  var var_lname = spreadsheet.getRange(participant_row, SECOND_NAME).getValue().toString().trim();
  var var_id = spreadsheet.getRange(participant_row, PARTICIPANT_ID).getValue().toString().trim();
  var var_address =  spreadsheet.getRange(participant_row, PARTICIPANT_ADDRESS).getValue().toString().trim();
  var var_city = spreadsheet.getRange(participant_row, PARTICIPANT_CITY).getValue().toString().trim();
  var var_postal = spreadsheet.getRange(participant_row, PARTICIPANT_POSTAL_CODE).getValue().toString().trim();
  var var_country = spreadsheet.getRange(participant_row, PARTICIPANT_COUNTRY).getValue().toString().trim();
  var var_date = spreadsheet.getRange(participant_row, CURRENT_DATE).getValue().toString().trim();
  body.replaceText("{{First Name}}",var_name);
  body.replaceText("{{Last Name}}",var_lname);
  body.replaceText("{{ID}}",var_id);
  body.replaceText("{{Address}}",var_address);
  body.replaceText("{{City}}",var_city);
  body.replaceText("{{Postal Code}}",var_postal);
  body.replaceText("{{Country}}",var_country);
  body.replaceText("{{Date}}", var_date);

  // Save and close
  opened_Docs.saveAndClose();

  // Generate the pdf associated with the file
  const pdfGenerated = opened_Docs.getAs(MimeType.PDF);
  var my_generated_pdf = pdfFolder.createFile(pdfGenerated).setName(file_name + ".pdf");

  // Send email with this PDF
  send_email(spreadsheet, participant_row, img1, img2, my_generated_pdf);
}


// Check for duplicate entry in the excel sheet
function count_duplicate(spreadsheet, participant_row)
{
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

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Create PDFs in drive').addItem('Create PDFs in drive','main').addToUi();
}
