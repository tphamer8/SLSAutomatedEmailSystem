var lock = LockService.getScriptLock();

function onEdit(e) {
  if (!lock.tryLock(5000)) {
    return;
  }

  try {
    if (e.source.getRange(e.range.getA1Notation()).getValue() !== true) {
      return;
    }

    let sheet = e.source.getSheetByName("New Hires"); // New Hires Tab

    // Extract information about the edited cell
    let range = e.range;
    let row = range.getRow();
    let col = range.getColumn();
    let cellValue = range.getValue();

    // Check if the checkbox was clicked
    if (col == 13 && cellValue === true) {
      // Get relevant info from row
      let name = sheet.getRange(row, 1).getValue();
      let email = sheet.getRange(row, 7).getValue();
      let link = sheet.getRange(row, 8).getValue();

      // Format email body
      let emailBody = `
        <p>Hi ${name},</p>
        <p>We have created a profile page for you on the SLS Website: <a href="${link}">${link}</a></p>
        <p>Please complete your profile by adding a short bio and photo*.</p>
        <ol>
          <li>Log into the website.</li>
          <li>Go to your profile page.</li>
          <li>Click on the edit icon located under your name.</li>
          <li>Fill out the form with the information you would like displayed.</li>
        </ol>
        <p><i><b>*Note: The photo should be high resolution, square in format, at least 800 x 800 pixels, and not too closely cropped.</b></i></p>
        <p>Please see our <a href="https://law.stanford.edu/webteam/updating-your-profile/">tutorial</a> for more information on how to log into the website and edit your profile. If there are changes that you would like to make that are not available on the form, please contact us.</p>
        <p>Thanks!,<br>SLS WebTeam</p>
      `;

      // Send the email
      MailApp.sendEmail({
        to: email,
        subject: 'Your Profile on the SLS Website',
        htmlBody: emailBody,
      });

      // Adds timestamp
      let timestamp = new Date();
      sheet.getRange(row, 14).setValue(timestamp);
    }
  } finally {
    lock.releaseLock();
  }
}

function createOnEditTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var found = false;

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == 'onEdit') {
      found = true;
      break;
    }
  }

  if (!found) {
    ScriptApp.newTrigger('onEdit')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();
  }
}

// onEdit trigger
createOnEditTrigger();
