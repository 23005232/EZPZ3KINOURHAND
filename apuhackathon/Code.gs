function onOpen() {
  DocumentApp.getUi().createMenu('Custom Menu')
      .addItem('Show Sidebar', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  initializeStatus(); // Initialize status data
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Name Detector and Email Sender');
  DocumentApp.getUi().showSidebar(html);
}

function getNamesAndEmails() {
  const sheetId = '1QoK_7nrCDdckVH5fxyQpDFaYJHPwWoqA8xMABD54-WM';
  Logger.log('Opening spreadsheet with ID: ' + sheetId);
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Sheet1');
  const data = sheet.getDataRange().getValues();
  Logger.log('Data from spreadsheet: ' + JSON.stringify(data));

  const nameEmailMap = {};
  for (let i = 1; i < data.length; i++) { 
    const name = data[i][0].trim();
    const email = data[i][2];
    nameEmailMap[name] = email;
  }
  Logger.log('Name and Email map: ' + JSON.stringify(nameEmailMap));

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody().getText();
  Logger.log('Document body text: ' + body);
  const names = body.split('\n').filter(line => line.trim().length > 0);
  Logger.log('Names extracted from document: ' + JSON.stringify(names));

  const namesAndEmails = names.map(name => {
    const email = nameEmailMap[name];
    if (email) {
      Logger.log('Matching name: ' + name + ', Email: ' + email);
      return { name: name, email: email };
    }
    Logger.log('No email found for name: ' + name);
    return null;
  }).filter(item => item !== null);

  Logger.log('Names and Emails to return: ' + JSON.stringify(namesAndEmails));
  return namesAndEmails;
}

function sendEmails() {
  const namesAndEmails = getNamesAndEmails();
  if (namesAndEmails.length === 0) {
    Logger.log('No names and emails found.');
    return;
  }

  // Initialize with the document you want to send first
  PropertiesService.getScriptProperties().setProperty('namesAndEmails', JSON.stringify(namesAndEmails));
  PropertiesService.getScriptProperties().setProperty('currentIndex', 0);
  PropertiesService.getScriptProperties().setProperty('currentDocId', DocumentApp.getActiveDocument().getId()); // Save initial document ID

  // Start the process of sending emails
  sendDocumentToNextPerson();

  // Create a time-based trigger to check for signed documents
  createTrigger();
}


function sendDocumentToNextPerson() {
  const namesAndEmails = JSON.parse(PropertiesService.getScriptProperties().getProperty('namesAndEmails'));
  let currentIndex = parseInt(PropertiesService.getScriptProperties().getProperty('currentIndex'), 10);
  const currentDocId = PropertiesService.getScriptProperties().getProperty('currentDocId');

  if (currentIndex >= namesAndEmails.length) {
    Logger.log('All documents have been processed.');
    return;
  }

  const entry = namesAndEmails[currentIndex];
  sendEmailToRecipient(entry.email, entry.name, currentDocId);
}

function sendEmailToRecipient(email, name, docId) {
  const pdf = convertDocToPdf(docId);

  const subject = 'Document for Signature';
  const message = `Hi ${name},\n\nPlease sign the attached document and reply to this email with the signed document attached.\n\nThank you.`;

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: message,
    attachments: [pdf]
  });

  Logger.log(`Email sent to ${email} with document ID: ${docId}`);

  // Update the status to 'Sent'
  updateEmailStatus(email, 'Sent');
}

function updateEmailStatus(name, status) {
  const sheetId = '1QoK_7nrCDdckVH5fxyQpDFaYJHPwWoqA8xMABD54-WM';
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Sheet1');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      sheet.getRange(i + 1, 4).setValue(status); // Assuming status is in the 4th column
      break;
    }
  }
}
function convertDocToPdf(docId) {
  return DriveApp.getFileById(docId).getAs('application/pdf');
}

function processSignedDocument() {
  const namesAndEmails = JSON.parse(PropertiesService.getScriptProperties().getProperty('namesAndEmails'));
  let currentIndex = parseInt(PropertiesService.getScriptProperties().getProperty('currentIndex'), 10);

  const threads = GmailApp.search('is:unread subject:"Document for Signature"');
  threads.forEach(thread => {
    const messages = thread.getMessages();
    const lastMessage = messages[messages.length - 1];
    const attachments = lastMessage.getAttachments();

    if (attachments.length > 0) {
      const signedPdf = attachments[0];
      const file = DriveApp.createFile(signedPdf);
      Logger.log(`Received signed document from ${lastMessage.getFrom()}`);

      // Update the status to 'Replied'
      updateEmailStatus(lastMessage.getFrom(), 'Replied');

      // Process the next recipient
      currentIndex++;
      PropertiesService.getScriptProperties().setProperty('currentIndex', currentIndex);

      if (currentIndex < namesAndEmails.length) {
        const nextEntry = namesAndEmails[currentIndex];
        sendEmailToRecipient(nextEntry.email, nextEntry.name, file.getId());
        PropertiesService.getScriptProperties().setProperty('currentDocId', file.getId());
      } else {
        Logger.log('All recipients have been emailed.');
        deleteTrigger();
      }

      thread.markRead();
    }
  });
}


function createTrigger() {
  // Delete any existing triggers for the processSignedDocument function
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processSignedDocument') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  // Create a new trigger that runs every minute
  ScriptApp.newTrigger('processSignedDocument')
    .timeBased()
    .everyMinutes(1)
    .create();
}

function deleteTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processSignedDocument') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}
function resetNamesAndEmails() {
  PropertiesService.getScriptProperties().deleteProperty('namesAndEmails');
  PropertiesService.getScriptProperties().deleteProperty('currentIndex');
  PropertiesService.getScriptProperties().deleteProperty('currentDocId');
}
function initializeStatus() {
  const namesAndEmails = getNamesAndEmails();
  const statusMap = {};

  namesAndEmails.forEach(entry => {
    statusMap[entry.email] = 'Pending'; // Default status
  });

  PropertiesService.getScriptProperties().setProperty('emailStatuses', JSON.stringify(statusMap));
}

function updateEmailStatus(email, status) {
  const statusMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('emailStatuses') || '{}');
  statusMap[email] = status;
  PropertiesService.getScriptProperties().setProperty('emailStatuses', JSON.stringify(statusMap));
}

function getEmailStatuses() {
  return JSON.parse(PropertiesService.getScriptProperties().getProperty('emailStatuses') || '{}');
}

function getNamesAndEmailsWithStatus() {
  const namesAndEmails = getNamesAndEmails();
  const statuses = getEmailStatuses();

  return namesAndEmails.map(entry => ({
    ...entry,
    status: statuses[entry.email] || 'Pending'
  }));
}

