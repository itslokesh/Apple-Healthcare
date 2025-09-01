function createNewDashboardBatchOptimized() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Sales Invoice responses");
  const targetSheetName = "New Dashboard";
  let targetSheet = ss.getSheetByName(targetSheetName);
  if (!sourceSheet) throw new Error('Source sheet "Sales Invoice responses" not found');

  // Headers for New Dashboard
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
  } else if (targetSheet.getLastRow() === 0) {
    targetSheet.appendRow(headers);
  }

  // --- Incremental window (no daily guard) ---
  const props = PropertiesService.getScriptProperties();
  const LAST_ROW_KEY = `${ss.getId()}:${sourceSheet.getSheetId()}:salesResp:lastProcessedRow`;

  const sourceLastRow = sourceSheet.getLastRow();
  const sourceLastCol = sourceSheet.getLastColumn();

  // Start after the last processed row; never before row 2 (skip header)
  const lastProcessedRow = Number(props.getProperty(LAST_ROW_KEY)) || 1;
  const startRow = Math.max(lastProcessedRow + 1, 2);
  const endRow = sourceLastRow;

  if (startRow > endRow) {
    // Nothing new
    return;
  }

  // Read only the newly added rows
  const sourceData = sourceSheet.getRange(startRow, 1, endRow - startRow + 1, sourceLastCol).getValues();

  // Read existing dashboard data into memory
  const lastRow = targetSheet.getLastRow();
  const existingRowsCount = Math.max(0, lastRow - 1);
  let dashboardData = existingRowsCount > 0
    ? targetSheet.getRange(2, 1, existingRowsCount, headers.length).getValues()
    : [];
  const dashboardMap = {}; // BillNo -> index in dashboardData

  dashboardData.forEach((row, idx) => {
    const bn = row[0];
    if (bn) dashboardMap[String(bn).trim()] = idx;
  });

  const invalidRows = [];

  // Process appended/new rows only
  // Note: indices below follow your original mapping from the source sheet
  sourceData.forEach(row => {
    const billNoRaw = row[3];
    const processType = row[4]?.toString().trim().toLowerCase();
    const timestamp = row[1] || "";

    if (!billNoRaw || String(billNoRaw).trim() === "") {
      invalidRows.push(row);
      return;
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
  });

  // Write dashboardData back (overwrite body from row 2)
  if (dashboardData.length > 0) {
    targetSheet.getRange(2, 1, dashboardData.length, headers.length).setValues(dashboardData);
  }

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

  // --- Persist checkpoint (only last processed source row) ---
  props.setProperty(LAST_ROW_KEY, String(endRow));
}