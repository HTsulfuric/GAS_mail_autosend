function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Auto Send").addItem("Send Emails", "sendEmails").addToUi();
}

function sendEmails() {
  const ss = SpreadsheetApp.getActive();

  const setting = ss.getSheetByName("setting");
  const list = ss.getSheetByName(setting.getRange(2, 1).getValue());

  const lastRow = list.getLastRow();

  const title = setting.getRange(2, 2).getValue();

  const documentUrl = setting.getRange(2, 3).getValue();

  const openDoc = DocumentApp.openByUrl(documentUrl);
  const body = openDoc.getBody().getText();

  let file = [];
  for (let i = 4; i <= 6; i++) {
    if (setting.getRange(2, i).getValue() != "") {
      let fileId = setting.getRange(2, i).getValue().split("/")[5];
      let blob = DriveApp.getFileById(fileId).getBlob();
      file.push(blob);
    }
  }

  for (let i = 2; i <= lastRow; i++) {
    let address = list.getRange(i, 1).getValue();
    let name = list.getRange(i, 2).getValue();

    let emailBody = body.replace("{{name}}", name);

    GmailApp.sendEmail(address, title, emailBody, {
      attachments: file,
    });
  }
}
