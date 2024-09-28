function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Auto Send").addItem("Send Emails", "sendEmails").addToUi();
}

function sendEmails() {
  const ss = SpreadsheetApp.getActive();
  const setting = ss.getSheetByName("setting");

  const listName = setting.getRange(2, 1).getValue();
  const title = setting.getRange(2, 2).getValue();
  const documentUrl = setting.getRange(2, 3).getValue();

  let file = getAttachments(setting);

  const from_email = setting.getRange(2, 7).getValue();

  let aliases = GmailApp.getAliases();
  if (!aliases.includes(from_email)) {
    Logger.log("Invalid from email address:" + from_email);
    return;
  }

  const list = ss.getSheetByName(listName);

  const body = DocumentApp.openByUrl(documentUrl).getBody().getText();

  const lastRow = list.getLastRow();

  for (let i = 2; i <= lastRow; i++) {
    const address = list.getRange(i, 1).getValue();
    const name = list.getRange(i, 2).getValue();

    try {
      sendEmail(address, name, title, body, file, from_email);
      Logger.log("Email sent to " + address);
    } catch (error) {
      Logger.log(
        "Failed to send email to " + address + " Error:" + error.message,
      );
    }
  }
}

function getAttachments(setting) {
  let file = [];
  for (let i = 4; i <= 6; i++) {
    if (setting.getRange(2, i).getValue() != "") {
      let fileId = setting
        .getRange(2, i)
        .getValue()
        .match(/[-\w]{25,}/);
      try {
        let blob = DriveApp.getFileById(fileId).getBlob();
        file.push(blob);
      } catch (error) {
        Logger.log("File not found:" + fileId + " Error:" + error.message);
      }
    }
  }
  return file;
}

function sendEmail(address, name, title, body, file, from_email) {
  const draft = makeDraft(address, name, title, body, file, from_email);
  const draftId = draft.getId();
  GmailApp.getDraft(draftId).send();
}

function makeDraft(address, name, title, body, file, from_email) {
  const emailBody = body.replace("{{name}}", name);
  const draft = GmailApp.createDraft(address, title, emailBody, {
    attachments: file,
    from: from_email,
  });
  return draft;
}
