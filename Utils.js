function generateLeadId() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LEADS);

  const year = new Date().getFullYear();

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return `LEAD-${year}-0001`;
  }

  const lastRow = data[data.length - 1];
  const lastId = lastRow[COL_LEAD_ID];

  if (!lastId || !lastId.toString().startsWith(`LEAD-${year}-`)) {

    return `LEAD-${year}-0001`;

  }

  const parts = lastId.toString().split("-");

  const lastNumber = parseInt(parts[2], 10);
  const nextNumber = lastNumber + 1;

  const paddedNumber = nextNumber.toString().padStart(5, "0");

  return `LEAD-${year}-${paddedNumber}`;


}


function testIdGeneration() {
  const newId = generateLeadId();
  Logger.log("Generated ID: " + newId);
}