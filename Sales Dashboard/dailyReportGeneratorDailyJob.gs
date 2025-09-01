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
  
  // Update Pending Orders Escalation sheet
  updatePendingOrdersEscalation(rows);
  
  // Log summary
  console.log("=== PROCESSING SUMMARY ===");
  console.log("Rows processed: " + (END_ROW - START_ROW + 1));
  console.log("New entries added to daily report");
  console.log("Pending Orders Escalation sheet updated");
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

// Function to update Pending Orders Escalation sheet
function updatePendingOrdersEscalation(processedRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let escalationSheet = ss.getSheetByName("Pending Orders Escalation");
  
  // Create escalation sheet if it doesn't exist
  if (!escalationSheet) {
    escalationSheet = ss.insertSheet("Pending Orders Escalation");
    
    // Define escalation headers (same as New Dashboard1 except last 3 columns)
    const escalationHeaders = [
      "Bill No",
      "Manual Bill Entry", 
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
      "Days Pending",
      "Escalation Level"
    ];
    
    // Set headers
    escalationSheet.getRange(1, 1, 1, escalationHeaders.length)
      .setValues([escalationHeaders])
      .setFontWeight("bold")
      .setBackground("#f1f3f4");
    
    // Auto-resize columns
    escalationSheet.autoResizeColumns(1, escalationHeaders.length);
    
    // Set column widths for timestamp columns
    escalationSheet.setColumnWidth(5, 180);  // Bill Picked Timestamp
    escalationSheet.setColumnWidth(7, 180);  // Packing Timestamp
    escalationSheet.setColumnWidth(9, 180);  // E-way Bill Timestamp
    escalationSheet.setColumnWidth(11, 180); // Shipping Timestamp
    escalationSheet.setColumnWidth(16, 180); // AWB Timestamp
    
    // Freeze header row
    escalationSheet.setFrozenRows(1);
  }
  
  // Get current data from escalation sheet
  const escalationData = escalationSheet.getDataRange().getValues();
  const escalationHeaders = escalationData[0];
  const existingRows = escalationData.slice(1); // Skip header
  
  // Find required column indices
  const billNoCol = escalationHeaders.indexOf("Bill No");
  const invoiceStateCol = escalationHeaders.indexOf("Invoice State");
  const daysPendingCol = escalationHeaders.indexOf("Days Pending");
  const escalationLevelCol = escalationHeaders.indexOf("Escalation Level");
  
  // Get column indices from New Dashboard1 for comparison
  const dashboardSheet = ss.getSheetByName("New Dashboard1");
  const dashboardData = dashboardSheet.getDataRange().getValues();
  const dashboardHeaders = dashboardData[0];
  const dashboardInvoiceStateCol = dashboardHeaders.indexOf("Invoice State");
  const dashboardBillPickedTimestampCol = dashboardHeaders.indexOf("Bill Picked Timestamp");
  const dashboardPackingTimestampCol = dashboardHeaders.indexOf("Packing Timestamp");
  const dashboardEwayBillTimestampCol = dashboardHeaders.indexOf("E-way Bill Timestamp");
  const dashboardShippingTimestampCol = dashboardHeaders.indexOf("Shipping Timestamp");
  
  const currentDate = new Date();
  const updatedRows = [];
  const billsToRemove = new Set();
  
  // Process each row from the escalation sheet
  existingRows.forEach((row, index) => {
    const billNo = row[billNoCol];
    
    // Find corresponding row in New Dashboard1
    const dashboardRow = dashboardData.find(dRow => dRow[dashboardHeaders.indexOf("Bill No")] === billNo);
    
    if (dashboardRow) {
      const invoiceState = dashboardRow[dashboardInvoiceStateCol];
      
      // Condition 1: If bill is shipped, mark for removal
      if (invoiceState === "Shipped") {
        billsToRemove.add(billNo);
        return; // Skip this row
      }
      
      // Calculate days pending based on the most recent activity
      let daysPending = 0;
      let escalationLevel = "";
      
      const billPickedTimestamp = dashboardRow[dashboardBillPickedTimestampCol];
      const packingTimestamp = dashboardRow[dashboardPackingTimestampCol];
      const ewayBillTimestamp = dashboardRow[dashboardEwayBillTimestampCol];
      const shippingTimestamp = dashboardRow[dashboardShippingTimestampCol];
      
      if (billPickedTimestamp) {
        try {
          const pickedDate = new Date(billPickedTimestamp);
          if (!isNaN(pickedDate.getTime())) {
            const timeDiff = currentDate.getTime() - pickedDate.getTime();
            daysPending = Math.ceil(timeDiff / (1000 * 3600 * 24));
          }
        } catch (e) {
          console.log(`Error parsing picked timestamp for bill ${billNo}: ${e.message}`);
        }
      } else if (packingTimestamp) {
        try {
          const packedDate = new Date(packingTimestamp);
          if (!isNaN(packedDate.getTime())) {
            const timeDiff = currentDate.getTime() - packedDate.getTime();
            daysPending = Math.ceil(timeDiff / (1000 * 3600 * 24));
          }
        } catch (e) {
          console.log(`Error parsing packing timestamp for bill ${billNo}: ${e.message}`);
        }
      } else if (ewayBillTimestamp) {
        try {
          const ewayDate = new Date(ewayBillTimestamp);
          if (!isNaN(ewayDate.getTime())) {
            const timeDiff = currentDate.getTime() - ewayDate.getTime();
            daysPending = Math.ceil(timeDiff / (1000 * 3600 * 24));
          }
        } catch (e) {
          console.log(`Error parsing e-way bill timestamp for bill ${billNo}: ${e.message}`);
        }
      } else if (shippingTimestamp) {
        try {
          const shippingDate = new Date(shippingTimestamp);
          if (!isNaN(shippingDate.getTime())) {
            const timeDiff = currentDate.getTime() - shippingDate.getTime();
            daysPending = Math.ceil(timeDiff / (1000 * 3600 * 24));
          }
        } catch (e) {
          console.log(`Error parsing shipping timestamp for bill ${billNo}: ${e.message}`);
        }
      }
      
      // Determine escalation level based on days pending
      if (daysPending >= 1 && daysPending < 2) {
        escalationLevel = "Yellow";
      } else if (daysPending >= 2 && daysPending <= 4) {
        escalationLevel = "Orange";
      } else if (daysPending > 4) {
        escalationLevel = "Red";
      } else {
        escalationLevel = "New";
      }
      
      // Update the row with new values
      const updatedRow = [...row];
      updatedRow[daysPendingCol] = daysPending;
      updatedRow[escalationLevelCol] = escalationLevel;
      
      updatedRows.push({
        rowData: updatedRow,
        daysPending: daysPending,
        escalationLevel: escalationLevel,
        originalIndex: index
      });
    }
  });
  
  // Add new bills that are not in the escalation sheet
  processedRows.forEach(row => {
    const billNo = row[dashboardHeaders.indexOf("Bill No")];
    const invoiceState = row[dashboardInvoiceStateCol];
    
    // Only process bills that are not shipped and not already in escalation sheet
    if (invoiceState !== "Shipped" && !existingRows.some(existingRow => existingRow[billNoCol] === billNo)) {
      // Calculate days pending for new bills
      let daysPending = 0;
      let escalationLevel = "";
      
      const billPickedTimestamp = row[dashboardBillPickedTimestampCol];
      const packingTimestamp = row[dashboardPackingTimestampCol];
      const ewayBillTimestamp = row[dashboardEwayBillTimestampCol];
      const shippingTimestamp = row[dashboardShippingTimestampCol];
      
      if (billPickedTimestamp) {
        try {
          const pickedDate = new Date(billPickedTimestamp);
          if (!isNaN(pickedDate.getTime())) {
            const timeDiff = currentDate.getTime() - pickedDate.getTime();
            daysPending = Math.ceil(timeDiff / (1000 * 3600 * 24));
          }
        } catch (e) {
          console.log(`Error parsing picked timestamp for new bill ${billNo}: ${e.message}`);
        }
      } else if (packingTimestamp) {
        try {
          const packedDate = new Date(packingTimestamp);
          if (!isNaN(packedDate.getTime())) {
            const timeDiff = currentDate.getTime() - packedDate.getTime();
            daysPending = Math.ceil(timeDiff / (1000 * 3600 * 24));
          }
        } catch (e) {
          console.log(`Error parsing packing timestamp for new bill ${billNo}: ${e.message}`);
        }
      } else if (ewayBillTimestamp) {
        try {
          const ewayDate = new Date(ewayBillTimestamp);
          if (!isNaN(ewayDate.getTime())) {
            const timeDiff = currentDate.getTime() - ewayDate.getTime();
            daysPending = Math.ceil(timeDiff / (1000 * 3600 * 24));
          }
        } catch (e) {
          console.log(`Error parsing e-way bill timestamp for new bill ${billNo}: ${e.message}`);
        }
      } else if (shippingTimestamp) {
        try {
          const shippingDate = new Date(shippingTimestamp);
          if (!isNaN(shippingDate.getTime())) {
            const timeDiff = currentDate.getTime() - shippingDate.getTime();
            daysPending = Math.ceil(timeDiff / (1000 * 3600 * 24));
          }
        } catch (e) {
          console.log(`Error parsing shipping timestamp for new bill ${billNo}: ${e.message}`);
        }
      }
      
      // Determine escalation level for new bills
      if (daysPending >= 1 && daysPending < 2) {
        escalationLevel = "Yellow";
      } else if (daysPending >= 2 && daysPending <= 4) {
        escalationLevel = "Orange";
      } else if (daysPending > 4) {
        escalationLevel = "Red";
      } else {
        escalationLevel = "New";
      }
      
      // Create new row data (excluding Invoice State, Day, Month)
      const newRowData = [
        row[dashboardHeaders.indexOf("Bill No")],
        row[dashboardHeaders.indexOf("Manual Bill Entry")] || "",
        row[dashboardHeaders.indexOf("Customer Name")],
        row[dashboardHeaders.indexOf("Picked By")],
        row[dashboardHeaders.indexOf("Bill Picked Timestamp")],
        row[dashboardHeaders.indexOf("Packed By")],
        row[dashboardHeaders.indexOf("Packing Timestamp")],
        row[dashboardHeaders.indexOf("E-way Bill By")],
        row[dashboardHeaders.indexOf("E-way Bill Timestamp")],
        row[dashboardHeaders.indexOf("Shipping By")],
        row[dashboardHeaders.indexOf("Shipping Timestamp")],
        row[dashboardHeaders.indexOf("Courier")],
        row[dashboardHeaders.indexOf("No of Boxes")],
        row[dashboardHeaders.indexOf("Weight (kg)")],
        row[dashboardHeaders.indexOf("AWB Number")],
        row[dashboardHeaders.indexOf("AWB Timestamp")],
        daysPending,
        escalationLevel
      ];
      
      updatedRows.push({
        rowData: newRowData,
        daysPending: daysPending,
        escalationLevel: escalationLevel,
        isNew: true
      });
    }
  });
  
  // Sort by escalation priority (Red > Orange > Yellow > New) and then by days pending
  updatedRows.sort((a, b) => {
    const priorityOrder = { "Red": 4, "Orange": 3, "Yellow": 2, "New": 1 };
    const priorityDiff = priorityOrder[b.escalationLevel] - priorityOrder[a.escalationLevel];
    
    if (priorityDiff !== 0) {
      return priorityDiff;
    }
    
    // If same priority, sort by days pending (descending)
    return b.daysPending - a.daysPending;
  });
  
  // Clear existing data (except headers)
  if (escalationSheet.getLastRow() > 1) {
    escalationSheet.getRange(2, 1, escalationSheet.getLastRow() - 1, escalationHeaders.length).clearContent();
  }
  
  // Write updated data
  if (updatedRows.length > 0) {
    const sheetData = updatedRows.map(item => item.rowData);
    escalationSheet.getRange(2, 1, sheetData.length, escalationHeaders.length).setValues(sheetData);
    
    // Apply color coding and group rows by escalation level
    applyEscalationColorCoding(escalationSheet, updatedRows);
    
    // Add borders
    escalationSheet.getRange(1, 1, updatedRows.length + 1, escalationHeaders.length)
      .setBorder(true, true, true, true, true, true);
  }
  
  console.log("=== PENDING ORDERS ESCALATION UPDATED ===");
  console.log(`Total pending bills: ${updatedRows.length}`);
  console.log(`🔴 Red level (Critical) more than 4 days: ${updatedRows.filter(b => b.escalationLevel === "Red").length}`);
  console.log(`🟠 Orange level (High) 2 to 4 days: ${updatedRows.filter(b => b.escalationLevel === "Orange").length}`);
  console.log(`🟡 Yellow level (Medium) 1 to 2 days: ${updatedRows.filter(b => b.escalationLevel === "Yellow").length}`);
  console.log(`🟢 New level (Low) less than 1 day: ${updatedRows.filter(b => b.escalationLevel === "New").length}`);
}

// Function to apply color coding and group rows by escalation level
function applyEscalationColorCoding(sheet, updatedRows) {
  if (updatedRows.length === 0) return;
  
  // Clear existing conditional formatting
  sheet.clearConditionalFormatRules();
  
  // Group rows by escalation level and apply colors
  let currentRow = 2; // Start after header
  
  // Red level (Critical) - Red background, white text
  const redRows = updatedRows.filter(item => item.escalationLevel === "Red");
  if (redRows.length > 0) {
    const redRange = sheet.getRange(currentRow, 1, redRows.length, 18);
    redRange.setBackground("#ea4335").setFontColor("#ffffff");
    currentRow += redRows.length;
  }
  
  // Orange level (High) - Orange background, black text
  const orangeRows = updatedRows.filter(item => item.escalationLevel === "Orange");
  if (orangeRows.length > 0) {
    const orangeRange = sheet.getRange(currentRow, 1, orangeRows.length, 18);
    orangeRange.setBackground("#ff9800").setFontColor("#000000");
    currentRow += orangeRows.length;
  }
  
  // Yellow level (Medium) - Yellow background, black text
  const yellowRows = updatedRows.filter(item => item.escalationLevel === "Yellow");
  if (yellowRows.length > 0) {
    const yellowRange = sheet.getRange(currentRow, 1, yellowRows.length, 18);
    yellowRange.setBackground("#ffeb3b").setFontColor("#000000");
    currentRow += yellowRows.length;
  }
  
  // New level (Low) - Light green background, black text
  const newRows = updatedRows.filter(item => item.escalationLevel === "New");
  if (newRows.length > 0) {
    const newRange = sheet.getRange(currentRow, 1, newRows.length, 18);
    newRange.setBackground("#c8e6c9").setFontColor("#000000");
  }
}
