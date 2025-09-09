function createNewDashboardBatchOptimized() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Sales Invoice responses");
  const targetSheetName = "New Dashboard";
  let targetSheet = ss.getSheetByName(targetSheetName);

  // === CONFIGURE START / END ROWS HERE ===
  const START_ROW = 2; // first row of source data to process
  const END_ROW = 8000; // last row of source data to process

  // Headers
  const headers = [
    "Bill No",           // index 0
    "Manual Bill Entry", // index 1 - NEW COLUMN
    "Customer Name",     // index 2
    "Picked By",        // index 3
    "Bill Picked Timestamp", // index 4
    "Packed By",        // index 5
    "Packing Timestamp", // index 6
    "E-way Bill By",    // index 7
    "E-way Bill Timestamp", // index 8
    "Shipping By",       // index 9
    "Shipping Timestamp", // index 10
    "Courier",          // index 11
    "No of Boxes",      // index 12
    "Weight (kg)",      // index 13
    "AWB Number",       // index 14
    "AWB Timestamp",    // index 15
    "Invoice State",    // index 16
    "Day",              // index 17
    "Month",            // index 18
    "Invalid Data"      // index 19
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
        entry[2] = row[6] || entry[2]; // Customer Name (index 2)
        entry[3] = row[5] || entry[3]; // Picked By (index 3)
        entry[4] = timestamp || entry[4]; // Bill Picked TS (index 4)
        break;
      case "packing":
        entry[5] = row[9] || entry[5]; // Packed By (index 5)
        entry[6] = timestamp || entry[6]; // Packing TS (index 6)
        entry[12] = row[7] || entry[12]; // No of Boxes (index 12)
        entry[13] = row[8] || entry[13]; // Weight (index 13)
        break;
      case "eway bill":
        entry[7] = row[12] || entry[7]; // E-way Bill By (index 7)
        entry[8] = timestamp || entry[8]; // E-way TS (index 8)
        break;
      case "shipping":
        entry[9] = row[15] || entry[9]; // Shipping By (index 9)
        entry[10] = timestamp || entry[10]; // Shipping TS (index 10)
        entry[11] = row[14] || entry[11]; // Courier (index 11)
        break;
      case "awb number":
        entry[14] = row[17] || entry[14]; // AWB Number (index 14)
        entry[15] = timestamp || entry[15]; // AWB TS (index 15)
        break;
    }

    // Update Invoice State (index 16)
    if (!entry[3]) entry[16] = "Pending Pick";
    else if (!entry[5]) entry[16] = "Pending Pack";
    else if (!entry[9]) entry[16] = "Pending Ship";
    else entry[16] = "Shipped";

    // Update Day / Month from Bill Picked Timestamp (indices 17, 18)
    if (entry[4]) {
      const dt = new Date(entry[4]);
      entry[17] = dt.getDate();
      entry[18] = dt.getMonth() + 1;
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
