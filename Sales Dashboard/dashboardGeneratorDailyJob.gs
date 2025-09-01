function createNewDashboardBatchOptimized() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Sales Invoice responses");
  const targetSheetName = "New Dashboard";
  const configSheetName = "Dashboard Config";
  let targetSheet = ss.getSheetByName(targetSheetName);
  let configSheet = ss.getSheetByName(configSheetName);

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

  // Create or get config sheet to track last processed row
  if (!configSheet) {
    configSheet = ss.insertSheet(configSheetName);
    configSheet.getRange("A1:B2").setValues([
      ["Last Processed Row", "Value"],
      ["Sales Invoice Responses", 1]
    ]);

  }

  // Get last processed row number
  const lastProcessedRow = configSheet.getRange("B2").getValue() || 1;

  // Create target sheet if not exists
  if (!targetSheet) {
    targetSheet = ss.insertSheet(targetSheetName);
    targetSheet.appendRow(headers);
  }

  // Read source data
  const sourceData = sourceSheet.getDataRange().getValues();
  const totalSourceRows = sourceData.length;

  // Determine start and end rows for processing
  const START_ROW = Math.max(2, lastProcessedRow + 1); // Start from row after last processed
  const END_ROW = totalSourceRows;

  if (START_ROW > END_ROW) {
    console.log("No new data to process. All rows up to " + END_ROW + " have been processed.");
    return;
  }



  // Read existing dashboard data into memory
  const lastRow = targetSheet.getLastRow();
  let dashboardData = lastRow > 1 ? targetSheet.getRange(2, 1, lastRow - 1, headers.length).getValues() : [];
  let dashboardMap = {}; // BillNo -> index in dashboardData

  dashboardData.forEach((row, idx) => {
    const bn = row[0];
    if (bn) dashboardMap[String(bn).trim()] = idx;
  });

  let invalidRows = [];
  let processedCount = 0;
  let newEntriesCount = 0;
  let updatedEntriesCount = 0;

  // Process only new rows
  for (let i = START_ROW - 1; i <= Math.min(END_ROW - 1, sourceData.length - 1); i++) {
    const row = sourceData[i];
    const billNoRaw = row[3];
    const processType = row[4]?.toString().trim().toLowerCase();
    const timestamp = row[1] || "";

    if (!billNoRaw || String(billNoRaw).trim() === "") {
      invalidRows.push(row);
      continue;
    }

    processedCount++;
    const billNo = String(billNoRaw).trim();

    let entry;
    let isNewEntry = false;
    
    if (dashboardMap.hasOwnProperty(billNo)) {
      entry = dashboardData[dashboardMap[billNo]];
      updatedEntriesCount++;
    } else {
      entry = Array(headers.length).fill("");
      entry[0] = billNo; // Bill No
      dashboardData.push(entry);
      dashboardMap[billNo] = dashboardData.length - 1;
      newEntriesCount++;
      isNewEntry = true;
    }

    // Process the data based on process type
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

  // Update the last processed row in config sheet
  configSheet.getRange("B2").setValue(END_ROW);
  
  // Log summary
  console.log("=== PROCESSING SUMMARY ===");
  console.log("Rows processed: " + processedCount);
  console.log("New entries added: " + newEntriesCount);
  console.log("Existing entries updated: " + updatedEntriesCount);
  console.log("Invalid rows: " + invalidRows.length);
  console.log("Last processed row updated to: " + END_ROW);
  console.log("Next run will start from row: " + (END_ROW + 1));
}

// Function to reset the processing counter (use this if you want to reprocess all data)
function resetProcessingCounter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Dashboard Config");
  
  if (configSheet) {
    configSheet.getRange("B2").setValue(1);
    console.log("Processing counter reset to row 1. Next run will process all data.");
  } else {
    console.log("Config sheet not found. Run createNewDashboardBatchOptimized() first.");
  }
}

// Function to show current processing status
function showProcessingStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Dashboard Config");
  const sourceSheet = ss.getSheetByName("Sales Invoice responses");
  
  if (!configSheet || !sourceSheet) {
    console.log("Required sheets not found.");
    return;
  }
  
  const lastProcessedRow = configSheet.getRange("B2").getValue() || 1;
  const totalSourceRows = sourceSheet.getDataRange().getValues().length;
  const pendingRows = Math.max(0, totalSourceRows - lastProcessedRow);
  
  console.log("=== PROCESSING STATUS ===");
  console.log("Last processed row: " + lastProcessedRow);
  console.log("Total source rows: " + totalSourceRows);
  console.log("Pending rows to process: " + pendingRows);
  console.log("Next run will process rows: " + (lastProcessedRow + 1) + " to " + totalSourceRows);
}
