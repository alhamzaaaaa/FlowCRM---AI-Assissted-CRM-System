
function createLead(leadData) {


  const existingLeadId = findLeadByEmail(leadData.email);

  if (existingLeadId) {
    Logger.log("Duplicate detected! Lead already exists.: " + existingLeadId);
    return existingLeadId;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LEADS);

  const leadId = generateLeadId();
  const timestamp = new Date();

  const newRow = [];
  newRow[COL_LEAD_ID] = leadId;
  newRow[COL_CREATED_AT] = timestamp;
  newRow[COL_NAME] = leadData.name || "";
  newRow[COL_COMPANY] = leadData.company || "";
  newRow[COL_EMAIL] = leadData.email || "";
  newRow[COL_PHONE] = leadData.phone || "";
  newRow[COL_SERVICE] = leadData.service || "";
  newRow[COL_BUDGET] = leadData.budget || "";
  newRow[COL_SOURCE] = leadData.source || "";
  newRow[COL_STAGE] = "NEW";
  newRow[COL_OWNER] = ""; // 
  newRow[COL_NEXT_FOLLOWUP] = "";
  newRow[COL_LAST_ACTIVITY] = timestamp;
  newRow[COL_NOTES] = leadData.notes || "";

  sheet.appendRow(newRow);

  return leadId;

}


function findLeadByEmail(email) {

  if (!email) return null;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LEADS);

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    const rowEmail = data[i][COL_EMAIL].toString().toLowerCase().trim();
    const searchEmail = email.toString().toLowerCase().trim();

    if (rowEmail === searchEmail) {

      return data[i][COL_LEAD_ID];

    }
  }

  return null;

}


function testCreateLead() {
  const mockFormData = {
    name: "Rahul Sharma",
    company: "ABC Technologies",
    email: "rahul@example.com",
    phone: "9876543210",
    service: "Web Development",
    budget: 50000,
    source: "Website",
    notes: "Need a business website"
  };
  
  const newLeadId = createLead(mockFormData);
  Logger.log("Successfully created lead: " + newLeadId);
}



























































































































































