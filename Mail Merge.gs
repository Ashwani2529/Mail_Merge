// Google Apps Script for mail merge
function mailMerge() {
  // Get DOCS_FILE_ID and SHEETS_FILE_ID dynamically
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
      .addItem('Send Emails', 'sendEmails')
      .addToUi();
  let DOCS_FILE_ID = Browser.inputBox("Mail Merge", 
                                      "Type or copy/paste the Docs URL:",
                                      Browser.Buttons.OK_CANCEL);
  let SHEETS_FILE_ID = Browser.inputBox("Mail Merge", 
                                      "Type or copy/paste the Sheet URL:",
                                      Browser.Buttons.OK_CANCEL);
                          
  DOCS_FILE_ID=extractDocumentIdFromUrl(DOCS_FILE_ID);
  SHEETS_FILE_ID=extractDocumentIdFromUrl(SHEETS_FILE_ID);
function extractDocumentIdFromUrl(url) {
  const regex = /\/d\/([a-zA-Z0-9-_]+)\//;
  const match = url.match(regex);
  if (match && match[1]) {
    return match[1];
  } else {
    throw new Error("Could not extract document ID from URL");
  }
}

  // Constants and settings
  const COLUMNS = ["subject","to_name", "to_title", "to_company", "to_address", "to_email","body","my_name","my_address","my_email","my_phone","filename"]; // Column names to be used

  // Fill-in your data to merge into document template variables
  let merge = {
    // sender data
    "my_name": null,
    "my_address":null,
    "my_email":null,
    "my_phone": null,
    // recipient data (supplied by 'sheets' data source)
    "to_name": null,
    "to_title": null,
    "to_company": null,
    "to_address": null,
    "subject":null,
    "to_email": null,
    "date": Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy MMMM dd"),
    "body": null,
  };

  // Main function to get data and process each form letter
  const data = getDataFromSheets(SHEETS_FILE_ID, COLUMNS);
  // let subjectLine = ""; // Change this to your desired subject line

  for (let i = 0; i < data.length; i++) {
  let row = data[i];
  for (let j = 0; j < COLUMNS.length; j++) {
    merge[COLUMNS[j]] = row[j];
  }
  let docId = mergeTemplate(DOCS_FILE_ID, merge);
  let pdf = convertToPdf(docId);
  let emailBody = merge["body"]; // Get the email body from the data
  draftEmail(merge["to_name"], merge["to_email"], pdf, merge["subject"], emailBody);
}


  function getDataFromSheets(sheetId, columns) {
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("Sheet1");
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];
  const data = values.slice(1); // Exclude header row

  const columnIndices = columns.map(column => headers.indexOf(column));

  return data.map(row => columnIndices.map(index => row[index]));
}


  function mergeTemplate(tmplId, mergeData) {
    const copyId = _copyTemplate(tmplId);
    const doc = DocumentApp.openById(copyId);
    const body = doc.getBody();
    for (let key in mergeData) {
      let value = mergeData[key];
      body.replaceText(`{{${key.toUpperCase()}}}`, value);
    }
    doc.saveAndClose();
    return copyId;
  }

  function _copyTemplate(tmplId) {
    const templateFile = DriveApp.getFileById(tmplId);
    const copy = templateFile.makeCopy(`${merge["filename"]}`);
    return copy.getId();
  }

  function convertToPdf(docId) {
    const doc = DriveApp.getFileById(docId);
    const pdf = doc.getAs('application/pdf');
    return pdf;
  }

  function draftEmail(toName, toEmail, pdf, subjectLine,body) {
    let subject = subjectLine;
    // subjectLine=subject;
    // const body = `Hi ${toName},\n\nPlease find your personalized letter attached.\n\nBest regards,\nAshwani Singh`;
     let draft = GmailApp.createDraft(toEmail, subject, body, {
    attachments: [pdf]
  });
    Logger.log(`Drafted email for ${toName}`);
    sendEmails(subjectLine);
  }
}

const RECIPIENT_COL = "to_email";
const EMAIL_SENT_COL = "Email Sent";

function sendEmails(subjectLine, sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()) {
  let emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  if (!emailTemplate) {
    Logger.log(`No draft found with subject line: ${subjectLine}`);
    return;
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();
  const heads = data.shift();
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  const out = [];

  obj.forEach(function(row, rowIdx) {
    if (row[EMAIL_SENT_COL] == '') {
      try {
        let msgObj = fillInTemplateFromObject_(emailTemplate.message, row);
        GmailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html,
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages
        });
        out.push([new Date()]);
      } catch (e) {
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });
  sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);

  function getGmailTemplateFromDrafts_(subject_line) {
    try {
      const drafts = GmailApp.getDrafts();
      let draft = drafts.filter(subjectFilter_(subject_line))[0];
      if (!draft) {
        Logger.log(`No draft found with subject line: ${subject_line}`);
        return null;
      }
      const msg = draft.getMessage();
      const allInlineImages = draft.getMessage().getAttachments({ includeInlineImages: true, includeAttachments: false });
      const attachments = draft.getMessage().getAttachments({ includeInlineImages: false });
      const htmlBody = msg.getBody();
      const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj), {});

      const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
      const matches = [...htmlBody.matchAll(imgexp)];

      const inlineImagesObj = {};
      matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

      return {
        message: { subject: subject_line, text: msg.getPlainBody(), html: htmlBody },
        attachments: attachments,
        inlineImages: inlineImagesObj
      };
    } catch (e) {
      Logger.log(`Error in getGmailTemplateFromDrafts_: ${e.message}`);
      return null;
    }
  }

  function subjectFilter_(subject_line) {
    return function(element) {
      return element.getMessage().getSubject() === subject_line;
    };
  }

  function fillInTemplateFromObject_(template, data) {
    let template_string = JSON.stringify(template);
    template_string = template_string.replace(/{{[^{}]+}}/g, key => {
      return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
    });
    return JSON.parse(template_string);
  }

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
  }
}
