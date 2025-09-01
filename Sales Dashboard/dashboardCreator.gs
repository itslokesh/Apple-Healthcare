function createNewDashboardBatchOptimized() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Sales Invoice responses");
  const targetSheetName = "New Dashboard";
  let targetSheet = ss.getSheetByName(targetSheetName);

  // === CONFIGURE START / END ROWS HERE ===
  const START_ROW = 15003; // first row of source data to process
  const END_ROW = 21929; // last row of source data to process

  // Headers
  const headers = [
    "Bill No",
    "Customer Name",
    "Picked By",
    "Bill Picked Timestamp",
    "Packed By",
    "Packing Timestamp",
    "E-way Bill By",
    "E-way Bill Timestamp",
    "Shipping By",
    "Shipping Timestamp",
    "Courier",
    "No of Boxes",
    "Weight (kg)",
    "AWB Number",
    "AWB Timestamp",
    "Invoice State",
    "Day",
    "Month",
    "Invalid Data"
  ];

  // Create sheet if not exists
  if (!targetSheet) {
    targetSheet = ss.insertSheet(targetSheetName);
    targetSheet.appendRow(headers);
  }

  // Read source data
  const sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length < START_ROW) return;

  // Read existing dashboard data into memory
  const lastRow = targetSheet.getLastRow();
  let dashboardData = lastRow > 1 ? targetSheet.getRange(2, 1, lastRow - 1, headers.length).getValues() : [];
  let dashboardMap = {}; // BillNo -> index in dashboardData

  dashboardData.forEach((row, idx) => {
    const bn = row[0];
    if (bn) dashboardMap[String(bn).trim()] = idx;
  });

  let invalidRows = [];

  // Process selected batch
  for (let i = START_ROW - 1; i <= Math.min(END_ROW - 1, sourceData.length - 1); i++) {
    const row = sourceData[i];
    const billNoRaw = row[3];
    const processType = row[4]?.toString().trim().toLowerCase();
    const timestamp = row[1] || "";

    if (!billNoRaw || String(billNoRaw).trim() === "") {
      invalidRows.push(row);
      continue;
    }

    const billNo = String(billNoRaw).trim();

    let entry;
    if (dashboardMap.hasOwnProperty(billNo)) {
      entry = dashboardData[dashboardMap[billNo]];
    } else {
      entry = Array(headers.length).fill("");
      entry[0] = billNo; // Bill No
      dashboardData.push(entry);
      dashboardMap[billNo] = dashboardData.length - 1;
    }

    switch (processType) {
      case "bill picked":
        entry[1] = row[6] || entry[1]; // Customer
        entry[2] = row[5] || entry[2]; // Picked By
        entry[3] = timestamp || entry[3]; // Bill Picked TS
        break;
      case "packing":
        entry[4] = row[9] || entry[4]; // Packed By
        entry[5] = timestamp || entry[5]; // Packing TS
        entry[11] = row[7] || entry[11]; // No of Boxes
        entry[12] = row[8] || entry[12]; // Weight
        break;
      case "eway bill":
        entry[6] = row[12] || entry[6]; // E-way Bill By
        entry[7] = timestamp || entry[7]; // E-way TS
        break;
      case "shipping":
        entry[8] = row[15] || entry[8]; // Shipping By
        entry[9] = timestamp || entry[9]; // Shipping TS
        entry[10] = row[14] || entry[10]; // Courier
        break;
      case "awb number":
        entry[13] = row[17] || entry[13]; // AWB Number
        entry[14] = timestamp || entry[14]; // AWB TS
        break;
    }

    // Update Invoice State
    if (!entry[2]) entry[15] = "Pending Pick";
    else if (!entry[4]) entry[15] = "Pending Pack";
    else if (!entry[8]) entry[15] = "Pending Ship";
    else entry[15] = "Shipped";

    // Update Day / Month from Bill Picked Timestamp
    if (entry[3]) {
      const dt = new Date(entry[3]);
      entry[16] = dt.getDate();
      entry[17] = dt.getMonth() + 1;
    }
  }

  // Write dashboardData back
  targetSheet.getRange(2, 1, dashboardData.length, headers.length).setValues(dashboardData);

  // Append invalid rows at the end
  if (invalidRows.length > 0) {
    const invalidDataArr = invalidRows.map(() => {
      const arr = Array(headers.length).fill("");
      arr[headers.length - 1] = "Bill Number Missing";
      return arr;
    });
    const startRowIndex = targetSheet.getLastRow() + 1;
    targetSheet.getRange(startRowIndex, 1, invalidDataArr.length, headers.length).setValues(invalidDataArr);
    targetSheet.getRange(startRowIndex, 1, invalidDataArr.length, headers.length).setBackground("red");
  }
}
