function assignLead(leadId) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leadsSheet = ss.getSheetByName(SHEET_LEADS);
  const teamsSheet = ss.getSheetByName(SHEET_TEAM);

  const leadData = leadsSheet.getDataRange().getValues();
  let leadRowIndex = -1;
  let requiredService = "";

  for (let i = 1; i < leadData.length; i++) {

    if (leadData[i][COL_LEAD_ID] === leadId) {

      leadRowIndex = i + 1;
      requiredService = leadData[i][COL_SERVICE];
      break;

    }
  }

  if (leadRowIndex === -1) {
    Logger.log("Lead Not Found: " + leadId);
    return "UNASSIGNED.";
  }

  const teamData = teamsSheet.getDataRange().getValues();

  let eligibleSalesPeople = [];

  for (let i = 1; i < teamData.length; i++) {

    if (!teamData[i] || !teamData[i][COL_TEAM_NAME]) {
      continue;
    }

    const name = teamData[i][COL_TEAM_NAME];
    const specialization = teamData[i][COL_TEAM_SPECIALIZATION];
    const status = teamData[i][COL_TEAM_STATUS];
    const activeLeads = teamData[i][COL_TEAM_ACTIVE_LEADS];

    if (status === "ACTIVE" && specialization === requiredService) {
      eligibleSalesPeople.push({ name: name, activeLeads: activeLeads });

    }
  }

  if (eligibleSalesPeople.length === 0) {
    Logger.log("No Eligible Salesperson for service: " + requiredService);
    return "UNASSIGNED";
  }

  eligibleSalesPeople.sort((a, b) => a.activeLeads - b.activeLeads);

  const bestMatch = eligibleSalesPeople[0].name;

  leadsSheet.getRange(leadRowIndex, COL_OWNER + 1).setValue(bestMatch);

  Logger.log("Lead " + leadId + " assigned to " + bestMatch);

  return bestMatch;
}

  const date1 = new Date("2026-10-24T10:00:00");
  const date2 = new Date("2026-10-24T17:00:00");

function isSameDay(date1, date2) {

  const string1 = date1.toDateString();
  const string2 = date2.toDateString();

  return string1 === string2;
}





































