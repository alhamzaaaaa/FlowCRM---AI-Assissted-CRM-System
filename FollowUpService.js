function processDailyFollowUps() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const followUpsSheet = ss.getSheetByName(SHEET_FOLLOWUPS);
  const leadsSheet = ss.getSheetByName(SHEET_LEADS);
  const leadData = leadsSheet.getDataRange().getValues();

  const fuData = followUpsSheet.getDataRange().getValues();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = 1; i < fuData.length; i++) {

    if (!fuData[i]) continue;

    const status = fuData[i][COL_FU_STATUS];

    if (status !== "PENDING") continue;

    const leadId = fuData[i][COL_FU_LEAD_ID];
    const rawDueDate = fuData[i][COL_FU_DUE_DATE];
    const owner = fuData[i][COL_FU_OWNER];


    if (!rawDueDate) continue;

    const dueDate = new Date(rawDueDate);
    dueDate.setHours(0, 0, 0, 0);

    let alertType = "DUE TODAY";

    if (dueDate.getTime() === today.getTime()) {
      alertType = "";
    } else if (dueDate.getTime() < today.getTime()) {
      alertType = "OVERDUE";
    } else {
      continue;
    }

    let leadName = "Unknown Lead";

    for (let j = 1; j < leadData.length; j++) {

      if (!leadData[j]) continue;

      if (leadData[j][COL_LEAD_ID] === leadId) {
        leadName = leadData[j][COL_NAME];
        break;
      }
    }

    Logger.log(`Processing: ${alertType} for ${leadName} assigned to ${owner}`);
    sendReminderEmail(owner, leadName, alertType, rawDueDate);

  }
}


































