function sendReminderEmail(ownerName, leadName, alertType, dueDate) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamsSheet = ss.getSheetByName(SHEET_TEAM);
  const teamData = teamsSheet.getDataRange().getValues();

  let ownerEmail = "";


  for (let i = 1; i < teamData.length; i++) {
    if (!teamData[i]) continue;

    if (teamData[i][COL_TEAM_NAME] === ownerName) {

      ownerEmail = teamData[i][COL_TEAM_EMAIL];
      break;

    }
  }

  if (!ownerEmail) {

    Logger.log("ERROR: Could not find email for owner: " + ownerName);
    return;

  }

  const subject = `
                    ${alertType} CRM Reminder: Follow Up with ${leadName}
  `;
  const body = `
          Hello ${ownerName}, \n\nThis is an automated reminder. \n\nYou have a follow-Up with ${leadName} that is ${alertType}.\nDue Date: ${dueDate}\n\nPlease update the CRM.\n\nBest Regards,\nFlowCRM Bot                                                                                                                                                  
  `;


  GmailApp.sendEmail(ownerEmail, subject, body);
  Logger.log(`Email sent successfully to ${ownerEmail}`);

}