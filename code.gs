/*
* ADD-ON-READY SCRIPT (v9.1 - Style Cleanup)
*
* This script uses an HTML sidebar for all actions and settings.
*
* CHANGE: The 'createMissingDocs' function now applies
* simplified styles: bold questions, plain answers, and no italics.
* Fixed typos from previous version.
*/


// --- GLOBAL CONSTANTS ---
const STUDENT_PERMISSION = "reader"; // Use 'reader' or 'writer', 'commenter' is invalid for Drive API v2
const DOC_STATUS_HEADER = "Doc URL";
const SHARE_STATUS_HEADER = "Share Status";
const EMAIL_COLUMN_TITLE = "Email Address"; // Hard-coded email column


// --- 1. THE MAIN MENU ---


function onOpen() {
 const ui = SpreadsheetApp.getUi();
 // We only create one menu item, which opens our main sidebar
 ui.createMenu("Grader Menu")
   .addItem("Open Grading Panel", "showSidebar")
   .addToUi();
}


/**
* Shows the 'Sidebar.html' file as a sidebar.
*/
function showSidebar() {
 const html = HtmlService.createHtmlOutputFromFile('Sidebar')
     .setTitle("Grading Panel");
 SpreadsheetApp.getUi().showSidebar(html);
}


// --- 2. SETTINGS & SETUP ---


/**
* Saves the settings object from the sidebar
* into the script's hidden PropertiesService.
*/
function saveSettings(settings) {
 try {
   // We only save the properties we receive from the sidebar
   PropertiesService.getUserProperties().setProperties(settings);
 } catch (e) {
   Logger.log("Error saving settings: " + e.message);
   throw new Error("Could not save settings.");
 }
}


/**
* Loads the settings from PropertiesService
* for the sidebar to display.
*/
function loadSettings() {
 return PropertiesService.getUserProperties().getProperties();
}


/**
* Reads settings from PropertiesService
*/
function getSettings() {
 const userProps = PropertiesService.getUserProperties();
 const settings = userProps.getProperties();


 // Check if settings are missing
 if (!settings.folderName) {
   // Simplified error string
   throw new Error("Settings not found. Open panel and save.");
 }
 return settings;
}


/**
* Returns a string instead of ui.alert()
*/
function setupSheet() {
 const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
 if (!sheet) {
   return "Error: Form Responses 1 sheet not found.";
 }
 const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
 let updated = false;


 if (headers.indexOf(DOC_STATUS_HEADER) === -1) {
   const col = sheet.getLastColumn() + 1;
   sheet.getRange(1, col).setValue(DOC_STATUS_HEADER);
   sheet.hideColumns(col);
   updated = true;
 }


 if (headers.indexOf(SHARE_STATUS_HEADER) === -1) {
   const col = sheet.getLastColumn() + 1;
   sheet.getRange(1, col).setValue(SHARE_STATUS_HEADER);
   sheet.hideColumns(col);
   updated = true;
 }
 if (updated) {
   return "Success! Helper columns are set up.";
 } else {
   return "Your sheet is already set up.";
 }
}


// --- 3. CORE SCRIPT: CREATE DOCS ---


function createMissingDocs() {
 let settings;
 try {
   settings = getSettings();
 } catch (e) {
   return e.message; // Show the "Please save settings" error
 }
 const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
 const spreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName();
 const data = sheet.getDataRange().getValues();
 const headers = data[0];


 const indices = {
   email: headers.indexOf(EMAIL_COLUMN_TITLE), // Hard-coded
   docStatus: headers.indexOf(DOC_STATUS_HEADER),
   shareStatus: headers.indexOf(SHARE_STATUS_HEADER)
 };


 if (indices.email === -1) {
   return "Error: Email column not found.";
 }
 if (indices.docStatus === -1 || indices.shareStatus === -1) {
   return "Error: Helper columns not found. Run setup.";
 }
 let folders = DriveApp.getFoldersByName(settings.folderName);
 let folder;
 if (folders.hasNext()) {
   folder = folders.next();
 } else {
   folder = DriveApp.createFolder(settings.folderName);
 }
 const teachers = settings.teacherEmails ? settings.teacherEmails.split(',').map(email => email.trim()) : [];
 let docsCreated = 0;


 // ************************************************
 // *** AESTHETIC CHANGES START HERE ***
 // Style for the main document body
 const bodyStyle = {};
 bodyStyle[DocumentApp.Attribute.FONT_FAMILY] = DocumentApp.FontFamily.ROBOTO;
 bodyStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
 bodyStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = "#333333"; // Dark gray
 bodyStyle[DocumentApp.Attribute.LINE_SPACING] = 1.5;
 bodyStyle[DocumentApp.Attribute.BOLD] = false;


 // Style for the "Email: ..." line
 const emailStyle = {};
 emailStyle[DocumentApp.Attribute.FONT_FAMILY] = DocumentApp.FontFamily.ROBOTO;
 emailStyle[DocumentApp.Attribute.FONT_SIZE] = 9;
 emailStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = "#777777"; // Lighter gray
 emailStyle[DocumentApp.Attribute.LINE_SPACING] = 1.5;


 // Style for the Question titles (replaces HEADING_2)
 const questionTitleStyle = {};
 questionTitleStyle[DocumentApp.Attribute.FONT_FAMILY] = DocumentApp.FontFamily.ROBOTO;
 questionTitleStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
 questionTitleStyle[DocumentApp.Attribute.BOLD] = true;
 questionTitleStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = "#000000"; // Black
 questionTitleStyle[DocumentApp.Attribute.LINE_SPACING] = 1.5;
 // *** AESTHETIC CHANGES END HERE ***
 // ************************************************


 for (let i = 1; i < data.length; i++) {
   const row = data[i];
   const docStatus = row[indices.docStatus];


   if (docStatus === "") {
     const studentEmail = row[indices.email];
     if (!studentEmail) {
       continue;
     }


     const emailPrefix = studentEmail.split('@')[0];
     // --- NEW NAMING SCHEME ---
     const docName = emailPrefix + " - " + spreadsheetName;


     const doc = DocumentApp.create(docName);
     const body = doc.getBody();


     // --- APPLY GLOBAL STYLE ---
     body.setAttributes(bodyStyle);


     body.appendParagraph(emailPrefix).setHeading(DocumentApp.ParagraphHeading.TITLE);


     // --- APPLY EMAIL STYLE ---
     body.appendParagraph("Email: " + studentEmail).setAttributes(emailStyle);


     body.appendHorizontalRule();
     body.appendParagraph(""); // <-- blank line


     headers.forEach((title, colIndex) => {
       if (title === DOC_STATUS_HEADER || title === SHARE_STATUS_HEADER || colIndex === indices.email) {
         return;
       }
       const answer = row[colIndex];
       if (answer && String(answer).trim() !== "") {
         const cleanTitle = String(title || "");
         if (cleanTitle.trim() !== "") {
           // --- APPLY QUESTION TITLE STYLE ---
           body.appendParagraph(cleanTitle)
               .setAttributes(questionTitleStyle);
         }
         // --- EXPLICITLY apply bodyStyle to the answer to ensure it's not bold ---
         body.appendParagraph(String(answer))
             .setAttributes(bodyStyle); // Apply base style
         body.appendParagraph(""); // Add a blank line
       }
     });


     doc.saveAndClose();


     const docFile = DriveApp.getFileById(doc.getId());
     docFile.moveTo(folder);
     const docId = docFile.getId();


     if (teachers.length > 0 && teachers[0] !== "") {
       teachers.forEach(email => {
         try {
           const permission = {
             'role': 'writer', // Use 'writer' for edit, Drive API v2 does not support 'commenter'
             'type': 'user',
             'value': email
           };
           Drive.Permissions.insert(permission, docId, {
             'sendNotificationEmails': false
           });
         } catch (e) {
           Logger.log("[Row " + (i + 1) + "] ERROR adding teacher " + email + ": " + e.message);
         }
       });
     }


     sheet.getRange(i + 1, indices.docStatus + 1).setValue(doc.getUrl());
     docsCreated++;
   }
 }
 return "Complete. Created " + docsCreated + " new docs.";
}


// --- 4. CORE SCRIPT: SHARE DOCS ---


function shareAllDocsWithStudents() {
 let settings;
 try {
   settings = getSettings();
 } catch (e) {
   return e.message; // Show the "Please save settings" error
 }


 const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
 const data = sheet.getDataRange().getValues();
 const headers = data[0];


 const indices = {
   email: headers.indexOf(EMAIL_COLUMN_TITLE), // Hard-coded
   docStatus: headers.indexOf(DOC_STATUS_HEADER),
   shareStatus: headers.indexOf(SHARE_STATUS_HEADER)
 };
 if (indices.email === -1) {
   return "Error: Email column not found.";
 }
 if (indices.docStatus === -1 || indices.shareStatus === -1) {
   return "Error: Helper columns not found. Run setup.";
 }
 let sharedCount = 0;
 // --- UPDATED Regex to handle both /d/ and ?id= URL formats ---
 const docIdRegex = /d\/([a-zA-Z0-9-_]{20,})|id=([a-zA-Z0-9-_]{20,})/;


 for (let i = 1; i < data.length; i++) {
   const row = data[i];
   const studentEmail = row[indices.email];
   const docUrl = row[indices.docStatus];
   const shareStatus = row[indices.shareStatus];


   if (docUrl && !shareStatus) {
     try {
       // --- IMPROVEMENT 1: Get ID from URL directly ---
       const match = docUrl.match(docIdRegex);


       // --- UPDATED ID extraction ---
       // Check match[1] (for /d/ format) or match[2] (for ?id= format)
       const docId = match ? (match[1] || match[2]) : null;


       if (!docId) {
         throw new Error("Could not parse Doc ID from URL. URL was: " + docUrl);
       }


       // --- IMPROVEMENT 2: Use global constant ---
       const permission = {
         'role': STUDENT_PERMISSION, // Use 'reader', 'writer', but not 'commenter'
         'type': 'user',
         'value': studentEmail
       };


       Drive.Permissions.insert(permission, docId, {
         'sendNotificationEmails': true
       });


       sheet.getRange(i + 1, indices.shareStatus + 1).setValue("Shared");
       sharedCount++;


     } catch (e)
     {
       // --- IMPROVED LOGGING ---
       Logger.log("[Row " + (i + 1) + "] ERROR sharing with student: " + e.message);
       sheet.getRange(i + 1, indices.shareStatus + 1).setValue("Error: " + e.message);
     }
   }
 }
 return "Complete. Shared " + sharedCount + " new docs.";
}



