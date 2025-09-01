function generateDailyReportBeautified() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName("New Dashboard1");
  const configSheetName = "Automations Config";
  
  if (!dashboardSheet) {
    console.log("Sheet 'New Dashboard1' not found!");
    return;
  }

  // Get config sheet to track last processed row
  let configSheet = ss.getSheetByName(configSheetName);
  if (!configSheet) {
    console.log("Config sheet 'Automations Config' not found!");
    return;
  }

  // Check if Daily Report tracking row exists, if not add it
  const configData = configSheet.getDataRange().getValues();
  let dailyReportRowIndex = -1;
  
  for (let i = 0; i < configData.length; i++) {
    if (configData[i][0] === "Daily Report - New Dashboard1") {
      dailyReportRowIndex = i + 1; // +1 because sheet rows are 1-indexed
      break;
    }
  }
  
  if (dailyReportRowIndex === -1) {
    // Add new row for Daily Report tracking
    const newRowIndex = configSheet.getLastRow() + 1;
    configSheet.getRange(newRowIndex, 1, 1, 2).setValues([
      ["Daily Report - New Dashboard1", 1]
    ]);
    dailyReportRowIndex = newRowIndex;
  }

  // Get last processed row number for Daily Report
  const lastProcessedRow = configSheet.getRange(dailyReportRowIndex, 2).getValue() || 1;
  
  // --- Column indices dynamically ---
  const data = dashboardSheet.getDataRange().getValues();
  const totalSourceRows = data.length;
  
  // Determine start and end rows for processing
  const START_ROW = Math.max(2, lastProcessedRow + 1); // Start from row after last processed
  const END_ROW = totalSourceRows;

  if (START_ROW > END_ROW) {
    console.log("No new data to process. All rows up to " + END_ROW + " have been processed.");
    return;
  }

  
  const headersRow = data[0];
  const dayCol = headersRow.indexOf("Day");
  const monthCol = headersRow.indexOf("Month");
  const pickedByCol = headersRow.indexOf("Picked By");
  const packedByCol = headersRow.indexOf("Packed By");
  const shippingByCol = headersRow.indexOf("Shipping By");
  const invoiceStateCol = headersRow.indexOf("Invoice State");
  const customerCol = headersRow.indexOf("Customer Name");
  const billNoCol = headersRow.indexOf("Bill No");
  const manualBillCol = headersRow.indexOf("Manual Bill Entry");

  if ([dayCol, monthCol, pickedByCol, packedByCol, shippingByCol, invoiceStateCol, customerCol, billNoCol].some(c => c === -1)) {
    throw new Error("One or more required columns not found in New Dashboard1. Verify headers.");
  }

  // --- Daily Report Sheet ---
  let dailySheet = ss.getSheetByName("Daily Report 2025-1");
  if (!dailySheet) dailySheet = ss.insertSheet("Daily Report 2025-1");
  const headers = ["Date", "Employee", "Picked Count", "Packed Count", "Shipped Count", "Pending Pick", "Pending Pack", "Customer Pending", "Daily Summary"];
  if (dailySheet.getLastRow() === 0) dailySheet.appendRow(headers);
  
  // Clear any existing merged ranges to prevent conflicts
  try {
    dailySheet.getRange(1, 1, dailySheet.getMaxRows(), dailySheet.getMaxColumns()).breakApart();
  } catch (e) {}

  // Process only new rows
  const rows = data.slice(START_ROW - 1, Math.min(END_ROW, data.length));
  const dailyMetrics = {}; // dayKey -> metrics

  // --- Aggregate metrics per day with bill numbers ---
  rows.forEach(row => {
    const day = row[dayCol];
    const month = row[monthCol];
    const dayKey = (day !== "" && month !== "") ? `${day}-${month}` : null;
    const pickedBy = row[pickedByCol];
    const packedBy = row[packedByCol];
    const shippingBy = row[shippingByCol];
    const invoiceState = row[invoiceStateCol];
    const customer = row[customerCol];
    const billNo = row[billNoCol];
    const manualBill = manualBillCol >= 0 ? row[manualBillCol] : "";
    const billToken = (billNo && manualBill) ? `${billNo} / ${manualBill}` : (billNo || manualBill || "");

    if (!dayKey || billToken === "") return; // skip invalid rows

    if (!dailyMetrics[dayKey]) dailyMetrics[dayKey] = {
      picked: {}, packed: {}, shipped: {},
      pendingPick: [], pendingPack: {}, rowsForDay: [],
      customerPending: {}
    };

    const dm = dailyMetrics[dayKey];
    dm.rowsForDay.push(row);

    // Track picked bills with employee
    if (pickedBy) {
      if (!dm.picked[pickedBy]) dm.picked[pickedBy] = [];
      dm.picked[pickedBy].push(billToken);
    }

    // Track packed bills with employee
    if (packedBy) {
      if (!dm.packed[packedBy]) dm.packed[packedBy] = [];
      dm.packed[packedBy].push(billToken);
    }

    // Track shipped bills with employee
    if (shippingBy) {
      if (!dm.shipped[shippingBy]) dm.shipped[shippingBy] = [];
      dm.shipped[shippingBy].push(billToken);
    }

    // Track pending pick bills
    if (!pickedBy) dm.pendingPick.push(billToken);

    // Track pending pack bills with employee
    if (!packedBy && pickedBy) {
      if (!dm.pendingPack[pickedBy]) dm.pendingPack[pickedBy] = [];
      dm.pendingPack[pickedBy].push(billToken);
    }

    if (invoiceState !== "Shipped") {
      if (!dm.customerPending[customer]) dm.customerPending[customer] = [];
      dm.customerPending[customer].push(billToken);
    }
  });

  // --- Prepare batch write with bill numbers ---
  const batchValues = [];
  const mergeInfo = [];
  Object.keys(dailyMetrics).sort((a,b)=>{
    const [d1,m1] = a.split("-").map(Number);
    const [d2,m2] = b.split("-").map(Number);
    return m1===m2?d1-d2:m1-m2;
  }).forEach(dayKey => {
    const dm = dailyMetrics[dayKey];
    const employees = new Set([...Object.keys(dm.picked), ...Object.keys(dm.packed), ...Object.keys(dm.shipped), ...Object.keys(dm.pendingPack)]);
    const startRowIndex = batchValues.length;

    employees.forEach(emp => {
      // Format counts with bill numbers in brackets
      const pickedBills = dm.picked[emp] || [];
      const packedBills = dm.packed[emp] || [];
      const shippedBills = dm.shipped[emp] || [];
      const pendingPackBills = dm.pendingPack[emp] || [];

      const pickedCount = pickedBills.length > 0 ? `${pickedBills.length} (${pickedBills.join(', ')})` : "0";
      const packedCount = packedBills.length > 0 ? `${packedBills.length} (${packedBills.join(', ')})` : "0";
      const shippedCount = shippedBills.length > 0 ? `${shippedBills.length} (${shippedBills.join(', ')})` : "0";
      const pendingPackCount = pendingPackBills.length > 0 ? `${pendingPackBills.length} (${pendingPackBills.join(', ')})` : "0";

      batchValues.push([
        dayKey,
        emp,
        pickedCount,
        packedCount,
        shippedCount,
        dm.pendingPick.length > 0 ? `${dm.pendingPick.length} (${dm.pendingPick.join(', ')})` : "0",
        pendingPackCount,
        JSON.stringify(dm.customerPending),
        "" // Daily summary placeholder
      ]);
    });

    const endRowIndex = batchValues.length - 1;
    mergeInfo.push({dayKey, startRowIndex, endRowIndex, dm});
  });

  // Find last used row in employee metrics area (column A) to append after existing employee metrics
  const colAValues = dailySheet.getRange(1, 1, dailySheet.getMaxRows(), 1).getValues();
  let lastEmployeeRow = 1; // header row default
  for (let i = colAValues.length - 1; i >= 1; i--) { // start from bottom, skip row 0 header
    if (colAValues[i][0] !== "") { lastEmployeeRow = i + 1; break; }
  }
  const startWriteRow = Math.max(2, lastEmployeeRow + 1);
  dailySheet.getRange(startWriteRow, 1, batchValues.length, headers.length).setValues(batchValues);

  // --- Merge Date, Pending Ship (H), Customer Pending (I) wrap, Daily Summary & add borders ---
  mergeInfo.forEach(info => {
    const startRow = startWriteRow + info.startRowIndex;
    const endRow = startWriteRow + info.endRowIndex;

    // Merge Date (A) - unmerge first if needed
    try {
      dailySheet.getRange(startRow, 1, endRow - startRow + 1).merge();
      dailySheet.getRange(startRow, 1).setVerticalAlignment("middle");
    } catch (e) {
      console.log("Warning: Could not merge Date column for rows " + startRow + " to " + endRow + ": " + e.message);
    }

    // Merge Pending Ship (H) - unmerge first if needed
    try {
      dailySheet.getRange(startRow, 8, endRow - startRow + 1).merge();
      dailySheet.getRange(startRow, 8).setVerticalAlignment("top").setWrap(true);
    } catch (e) {
      console.log("Warning: Could not merge Pending Ship column for rows " + startRow + " to " + endRow + ": " + e.message);
    }

    // Customer Pending (I) - wrap text - unmerge first if needed
    try {
      dailySheet.getRange(startRow, 9, endRow - startRow + 1).merge();
      dailySheet.getRange(startRow, 9).setVerticalAlignment("top").setWrap(true);
    } catch (e) {
      console.log("Warning: Could not merge Customer Pending column for rows " + startRow + " to " + endRow + ": " + e.message);
    }

    // --- Calculate summary for this day with bill numbers ---
    const totalPicked = Object.values(info.dm.picked).reduce((a,b)=>a+b.length,0);
    const totalPacked = Object.values(info.dm.packed).reduce((a,b)=>a+b.length,0);
    const totalShipped = Object.values(info.dm.shipped).reduce((a,b)=>a+b.length,0);
    const pendingPackSum = Object.values(info.dm.pendingPack).reduce((a,b)=>a+b.length,0);
    const pendingShipCount = Object.values(info.dm.customerPending).reduce((sum, arr) => sum + arr.length, 0);

    // Get all bill numbers for summary
    const allPickedBills = Object.values(info.dm.picked).flat();
    const allPackedBills = Object.values(info.dm.packed).flat();
    const allShippedBills = Object.values(info.dm.shipped).flat();
    const allPendingPickBills = info.dm.pendingPick;
    const allPendingPackBills = Object.values(info.dm.pendingPack).flat();
    const allPendingShipBills = Object.values(info.dm.customerPending).flat();

    const summaryText = `Total Picked: ${totalPicked} (${allPickedBills.join(', ')})\n` +
                        `Total Packed: ${totalPacked} (${allPackedBills.join(', ')})\n` +
                        `Total Shipped: ${totalShipped} (${allShippedBills.join(', ')})\n` +
                        `Pending Pick: ${info.dm.pendingPick.length} (${allPendingPickBills.join(', ')})\n` +
                        `Pending Pack: ${pendingPackSum} (${allPendingPackBills.join(', ')})\n` +
                        `Pending Ship: ${pendingShipCount} (${allPendingShipBills.join(', ')})`;

    dailySheet.getRange(startRow, 9).setValue(summaryText);

    // Outline border for all columns of the day
    dailySheet.getRange(startRow, 1, endRow - startRow + 1, headers.length).setBorder(true, true, true, true, false, false);
    dailySheet.getRange(1, 1, 1, headers.length).setBackground("#f1f3f4");

  });

  // Write headers
  // --- Daily Customer Metrics (L1 onward) ---
  const customerMetricsHeaders = ["Date", "Customer", "Total Orders", "Picked", "Packed", "Shipped", "Pending"];
  const metricsStartCol = 12; // L
  const metricsWidth = customerMetricsHeaders.length;

  // Only clear and recreate headers if this is the first run (sheet is empty)
  if (dailySheet.getLastRow() === 1) {
    // Clear previous customer metrics table
    dailySheet.getRange(1, metricsStartCol, dailySheet.getMaxRows(), metricsWidth).clearContent();
    // Headers
    dailySheet.getRange(1, metricsStartCol, 1, metricsWidth).setValues([customerMetricsHeaders]).setFontWeight("bold");
  } else {
    // Check if headers exist, if not add them
    const existingHeaders = dailySheet.getRange(1, metricsStartCol, 1, metricsWidth).getValues()[0];
    if (existingHeaders[0] !== "Date" || existingHeaders[1] !== "Customer") {
      dailySheet.getRange(1, metricsStartCol, 1, metricsWidth).setValues([customerMetricsHeaders]).setFontWeight("bold");
    }
  }

  // Build per-day, per-customer metrics with bill numbers
  const customerMetricsRows = [];
  
  Object.keys(dailyMetrics).sort((a,b)=>{
    const [d1,m1] = a.split("-").map(Number);
    const [d2,m2] = b.split("-").map(Number);
    return m1===m2?d1-d2:m1-m2;
  }).forEach(dayKey => {
    const dm = dailyMetrics[dayKey];
    const perCustomer = {};

    dm.rowsForDay.forEach(r => {
      const custName = r[customerCol] || "Unknown";
      const billNo = r[billNoCol];
      const manualBill = manualBillCol >= 0 ? r[manualBillCol] : "";
      const billToken = (billNo && manualBill) ? `${billNo} / ${manualBill}` : (billNo || manualBill || "");
      const isPicked = !!r[pickedByCol];
      const isPacked = !!r[packedByCol];
      const isShipped = r[invoiceStateCol] === "Shipped";

      if (!perCustomer[custName]) {
        perCustomer[custName] = { 
          total: 0, picked: 0, packed: 0, shipped: 0,
          totalBills: [], pickedBills: [], packedBills: [], shippedBills: []
        };
      }
      const cs = perCustomer[custName];
      cs.total += 1;
      cs.totalBills.push(billToken);
      if (isPicked) {
        cs.picked += 1;
        cs.pickedBills.push(billToken);
      }
      if (isPacked) {
        cs.packed += 1;
        cs.packedBills.push(billToken);
      }
      if (isShipped) {
        cs.shipped += 1;
        cs.shippedBills.push(billToken);
      }
    });

    Object.keys(perCustomer).sort().forEach(custName => {
      const cs = perCustomer[custName];
      const pending = cs.total - cs.shipped; // pending = not shipped
      const pendingBills = cs.totalBills.filter(bill => !cs.shippedBills.includes(bill));
      
      // Format counts with bill numbers in brackets
      const totalCount = cs.total > 0 ? `${cs.total} (${cs.totalBills.join(', ')})` : "0";
      const pickedCount = cs.picked > 0 ? `${cs.picked} (${cs.pickedBills.join(', ')})` : "0";
      const packedCount = cs.packed > 0 ? `${cs.packed} (${cs.packedBills.join(', ')})` : "0";
      const shippedCount = cs.shipped > 0 ? `${cs.shipped} (${cs.shippedBills.join(', ')})` : "0";
      const pendingCount = pending > 0 ? `${pending} (${pendingBills.join(', ')})` : "0";
      
             customerMetricsRows.push([dayKey, custName, totalCount, pickedCount, packedCount, shippedCount, pendingCount]);
     });
   });
   


  // Track where customer metrics are written
  let customerMetricsWriteStart = null;
  let customerMetricsWriteEnd = null;

  if (customerMetricsRows.length > 0) {
    // Find last non-empty row in column L (metrics area) to append after existing customer metrics
    const lastRowToCheck = Math.max(2, dailySheet.getMaxRows());
    const colLValues = dailySheet.getRange(2, metricsStartCol, lastRowToCheck - 1, 1).getValues();
    let lastNonEmpty = 1; // header row default
    for (let i = colLValues.length - 1; i >= 0; i--) {
      if (colLValues[i][0] !== "") { lastNonEmpty = i + 2; break; }
    }
    const startRow = Math.max(2, lastNonEmpty + 1);
    dailySheet.getRange(startRow, metricsStartCol, customerMetricsRows.length, metricsWidth).setValues(customerMetricsRows);
    customerMetricsWriteStart = startRow;
    customerMetricsWriteEnd = startRow + customerMetricsRows.length - 1;
  }

  // Optional formatting for readability
  dailySheet.setColumnWidths(metricsStartCol, metricsWidth, 140);
  dailySheet.getRange(1, metricsStartCol, 1, metricsWidth).setBackground("#f1f3f4");
   // Merge same-date blocks in L and outline their rows across L:R
  if (customerMetricsRows.length > 0 && customerMetricsWriteStart !== null && customerMetricsWriteEnd !== null) {
    // Use the exact rows where customer metrics were written
    const actualStartRow = customerMetricsWriteStart;
    const actualLastRow = customerMetricsWriteEnd;
    const lCol = 12; // 12 (L)
    const width = 7;     // L:R

    let blockStart = actualStartRow;
    let prevDate = customerMetricsRows[0][0];

    for (let i = 1; i < customerMetricsRows.length; i++) {
      const currDate = customerMetricsRows[i][0];
      if (currDate !== prevDate) {
        const blockEnd = actualStartRow + i - 1;
        try {
          dailySheet.getRange(blockStart, lCol, blockEnd - blockStart + 1, 1)
            .merge()
            .setVerticalAlignment("middle");
        } catch (e) {
          console.log("Warning: Could not merge Date column for customer metrics rows " + blockStart + " to " + blockEnd + ": " + e.message);
        }
        dailySheet.getRange(blockStart, lCol, blockEnd - blockStart + 1, width)
          .setBorder(true, true, true, true, false, false);
        blockStart = blockEnd + 1;
        prevDate = currDate;
      }
    }
    // finalize last block
    try {
      dailySheet.getRange(blockStart, lCol, actualLastRow - blockStart + 1, 1)
        .merge()
        .setVerticalAlignment("middle");
    } catch (e) {
      console.log("Warning: Could not merge Date column for customer metrics rows " + blockStart + " to " + actualLastRow + ": " + e.message);
    }
    dailySheet.getRange(blockStart, lCol, actualLastRow - blockStart + 1, width)
      .setBorder(true, true, true, true, false, false);
  }

  // Update the last processed row in config sheet
  configSheet.getRange(dailyReportRowIndex, 2).setValue(END_ROW);
  
  // Log summary
  console.log("=== PROCESSING SUMMARY ===");
  console.log("Rows processed: " + (END_ROW - START_ROW + 1));
  console.log("New entries added to daily report");
  console.log("Last processed row updated to: " + END_ROW);
  console.log("Next run will start from row: " + (END_ROW + 1));
}

// Function to reset the processing counter (use this if you want to reprocess all data)
function resetDailyReportProcessingCounter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Automations Config");
  
  if (!configSheet) {
    console.log("Config sheet 'Automations Config' not found!");
    return;
  }
  
  // Find the Daily Report tracking row
  const configData = configSheet.getDataRange().getValues();
  let dailyReportRowIndex = -1;
  
  for (let i = 0; i < configData.length; i++) {
    if (configData[i][0] === "Daily Report - New Dashboard1") {
      dailyReportRowIndex = i + 1; // +1 because sheet rows are 1-indexed
      break;
    }
  }
  
  if (dailyReportRowIndex === -1) {
    console.log("Daily Report tracking row not found in Automations Config. Run generateDailyReportBeautified() first.");
    return;
  }
  
  configSheet.getRange(dailyReportRowIndex, 2).setValue(1);
  console.log("Daily Report processing counter reset to row 1. Next run will process all data.");
}

// Function to show current processing status
function showDailyReportProcessingStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Automations Config");
  const dashboardSheet = ss.getSheetByName("New Dashboard1");
  
  if (!configSheet || !dashboardSheet) {
    console.log("Required sheets not found.");
    return;
  }
  
  // Find the Daily Report tracking row
  const configData = configSheet.getDataRange().getValues();
  let dailyReportRowIndex = -1;
  
  for (let i = 0; i < configData.length; i++) {
    if (configData[i][0] === "Daily Report - New Dashboard1") {
      dailyReportRowIndex = i + 1; // +1 because sheet rows are 1-indexed
      break;
    }
  }
  
  if (dailyReportRowIndex === -1) {
    console.log("Daily Report tracking row not found in Automations Config. Run generateDailyReportBeautified() first.");
    return;
  }
  
  const lastProcessedRow = configSheet.getRange(dailyReportRowIndex, 2).getValue() || 1;
  const totalSourceRows = dashboardSheet.getDataRange().getValues().length;
  const pendingRows = Math.max(0, totalSourceRows - lastProcessedRow);
  
  console.log("=== DAILY REPORT PROCESSING STATUS ===");
  console.log("Last processed row: " + lastProcessedRow);
  console.log("Total source rows: " + totalSourceRows);
  console.log("Pending rows to process: " + pendingRows);
  console.log("Next run will process rows: " + (lastProcessedRow + 1) + " to " + totalSourceRows);
}
