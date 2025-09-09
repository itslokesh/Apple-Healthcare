function monthlyReportGenerator() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName("Automation Dashboard");
  const configSheetName = "Automations Config";
  
  if (!dashboardSheet) {
    console.log("Sheet 'Automation Dashboard' not found!");
    return;
  }

  // Get config sheet to track last processed row
  let configSheet = ss.getSheetByName(configSheetName);
  if (!configSheet) {
    console.log("Config sheet 'Automations Config' not found!");
    return;
  }

  // Check if Monthly Report tracking row exists, if not add it
  const configData = configSheet.getDataRange().getValues();
  let monthlyReportRowIndex = -1;
  
  for (let i = 0; i < configData.length; i++) {
    if (configData[i][0] === "Monthly Report") {
      monthlyReportRowIndex = i + 1; // +1 because sheet rows are 1-indexed
      break;
    }
  }
  
  if (monthlyReportRowIndex === -1) {
    // Add new row for Monthly Report tracking
    const newRowIndex = configSheet.getLastRow() + 1;
    configSheet.getRange(newRowIndex, 1, 1, 2).setValues([
      ["Monthly Report", 1]
    ]);
    monthlyReportRowIndex = newRowIndex;
  }

  // Get last processed row number for Monthly Report
  const lastProcessedRow = configSheet.getRange(monthlyReportRowIndex, 2).getValue() || 1;
  
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

  // --- Monthly Report Sheet ---
  let monthlySheet = ss.getSheetByName("Monthly Report 2025");
  if (!monthlySheet) monthlySheet = ss.insertSheet("Monthly Report 2025");

  // Clear any existing merged ranges to prevent conflicts
  try {
    monthlySheet.getRange(1, 1, monthlySheet.getMaxRows(), monthlySheet.getMaxColumns()).breakApart();
  } catch (e) {
    console.log("Break apart failed: " + e.message);
  }

  const headers = ["Month", "Employee", "Picked Count", "Packed Count", "Shipped Count", "Pending Pick", "Pending Pack", "Customer Pending", "Monthly Summary"];

  // Process only new rows
  const rows = data.slice(START_ROW - 1, Math.min(END_ROW, data.length));
  const monthlyMetrics = {}; // monthKey -> metrics

  // --- Aggregate metrics per month with bill numbers ---
  rows.forEach(row => {
    const month = row[monthCol];
    const monthKey = month !== "" ? String(month) : null;
    const pickedBy = row[pickedByCol];
    const packedBy = row[packedByCol];
    const shippingBy = row[shippingByCol];
    const invoiceState = row[invoiceStateCol];
    const customer = row[customerCol];
    const billNo = row[billNoCol];
    const manualBill = manualBillCol >= 0 ? row[manualBillCol] : "";
    const billToken = (billNo && manualBill) ? `${billNo}/${manualBill}` : (billNo || manualBill || "");

    if (!monthKey || billToken === "") return;

    if (!monthlyMetrics[monthKey]) monthlyMetrics[monthKey] = {
      picked: {}, packed: {}, shipped: {},
      pendingPick: [], pendingPack: {}, rowsForMonth: [],
      customerPending: {}
    };

    const mm = monthlyMetrics[monthKey];
    mm.rowsForMonth.push(row);

    // Track picked bills with employee
    if (pickedBy) {
      if (!mm.picked[pickedBy]) mm.picked[pickedBy] = [];
      mm.picked[pickedBy].push(billToken);
    }

    // Track packed bills with employee
    if (packedBy) {
      if (!mm.packed[packedBy]) mm.packed[packedBy] = [];
      mm.packed[packedBy].push(billToken);
    }

    // Track shipped bills with employee
    if (shippingBy) {
      if (!mm.shipped[shippingBy]) mm.shipped[shippingBy] = [];
      mm.shipped[shippingBy].push(billToken);
    }

    // Track pending pick bills
    if (!pickedBy) mm.pendingPick.push(billToken);

    // Track pending pack bills with employee
    if (!packedBy && pickedBy) {
      if (!mm.pendingPack[pickedBy]) mm.pendingPack[pickedBy] = [];
      mm.pendingPack[pickedBy].push(billToken);
    }

    if (invoiceState !== "Shipped") {
      if (!mm.customerPending[customer]) mm.customerPending[customer] = [];
      mm.customerPending[customer].push(billToken);
    }
  });

  // --- Prepare batch write with bill numbers ---
  const batchValues = [];
  const mergeInfo = [];
  Object.keys(monthlyMetrics).sort((a,b)=>Number(a)-Number(b)).forEach(monthKey => {
    const mm = monthlyMetrics[monthKey];
    const employees = new Set([...Object.keys(mm.picked), ...Object.keys(mm.packed), ...Object.keys(mm.shipped), ...Object.keys(mm.pendingPack)]);
    const startRowIndex = batchValues.length;

    employees.forEach(emp => {
      // Format counts with bill numbers in brackets
      const pickedBills = mm.picked[emp] || [];
      const packedBills = mm.packed[emp] || [];
      const shippedBills = mm.shipped[emp] || [];
      const pendingPackBills = mm.pendingPack[emp] || [];

      const pickedCount = pickedBills.length > 0 ? `${pickedBills.length} (${pickedBills.join(', ')})` : "0";
      const packedCount = packedBills.length > 0 ? `${packedBills.length} (${packedBills.join(', ')})` : "0";
      const shippedCount = shippedBills.length > 0 ? `${shippedBills.length} (${shippedBills.join(', ')})` : "0";
      const pendingPackCount = pendingPackBills.length > 0 ? `${pendingPackBills.length} (${pendingPackBills.join(', ')})` : "0";

      batchValues.push([
        monthKey,
        emp,
        pickedCount,
        packedCount,
        shippedCount,
        mm.pendingPick.length > 0 ? `${mm.pendingPick.length} (${mm.pendingPick.join(', ')})` : "0",
        pendingPackCount,
        JSON.stringify(mm.customerPending),
        "" // Monthly summary placeholder
      ]);
    });

    const endRowIndex = batchValues.length - 1;
    mergeInfo.push({monthKey, startRowIndex, endRowIndex, mm});
  });

  // --- Write main table at A20 ---
  const TABLE_START_ROW = 20;

  // Find last used row in employee metrics area (column A) to append after existing employee metrics
  // We need to find the actual end of data, not just the last non-empty cell
  let lastEmployeeRow = TABLE_START_ROW; // start from header row
  
  // Check if there's existing data in the table
  if (monthlySheet.getLastRow() > TABLE_START_ROW) {
    // Get all data from the table area
    const tableData = monthlySheet.getRange(TABLE_START_ROW + 1, 1, monthlySheet.getLastRow() - TABLE_START_ROW, 9).getValues();
    
    // Find the last row that has actual data (not just empty strings)
    for (let i = tableData.length - 1; i >= 0; i--) {
      const row = tableData[i];
      // Check if this row has meaningful data (Month, Employee, or any other column)
      if (row[0] !== "" || row[1] !== "" || row[2] !== "" || row[3] !== "" || row[4] !== "" || row[5] !== "" || row[6] !== "") {
        lastEmployeeRow = TABLE_START_ROW + 1 + i; // Convert back to sheet row number
        break;
      }
    }
  }
  
  const startWriteRow = lastEmployeeRow + 1;

  // Debug logging for employee metrics appending
  console.log("=== EMPLOYEE METRICS APPENDING DEBUG ===");
  console.log("TABLE_START_ROW:", TABLE_START_ROW);
  console.log("Last employee row found:", lastEmployeeRow);
  console.log("Will append starting at row:", startWriteRow);
  console.log("Total rows to append:", batchValues.length);

  // Headers at A20:I20 (only if sheet is empty)
  if (monthlySheet.getLastRow() === 0) {
    monthlySheet.getRange(TABLE_START_ROW, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight("bold")
      .setBackground("#f1f3f4");
  }

  // Data starting at calculated startWriteRow
  if (batchValues.length > 0) {
    monthlySheet.getRange(startWriteRow, 1, batchValues.length, headers.length).setValues(batchValues);
  }

  // --- Merge Month (A), Pending Ship (H), Customer Pending (I) wrap, Monthly Summary & add borders ---
  mergeInfo.forEach(info => {
    const startRow = startWriteRow + info.startRowIndex;
    const endRow = startWriteRow + info.endRowIndex;

    // Merge Month (A)
    try {
      monthlySheet.getRange(startRow, 1, endRow - startRow + 1).merge();
      monthlySheet.getRange(startRow, 1).setVerticalAlignment("middle");
    } catch (e) {
      console.log("Merge failed for Month column: " + e.message);
    }

    // Merge Pending Ship (H)
    try {
      monthlySheet.getRange(startRow, 8, endRow - startRow + 1).merge();
      monthlySheet.getRange(startRow, 8).setVerticalAlignment("top").setWrap(true);
    } catch (e) {
      console.log("Merge failed for Pending Ship column: " + e.message);
    }

    // Customer Pending (I) - wrap text
    try {
      monthlySheet.getRange(startRow, 9, endRow - startRow + 1).merge();
      monthlySheet.getRange(startRow, 9).setVerticalAlignment("top").setWrap(true);
    } catch (e) {
      console.log("Merge failed for Customer Pending column: " + e.message);
    }

    // --- Monthly summary for this month with bill numbers ---
    const totalPicked = Object.values(info.mm.picked).reduce((a,b)=>a+b.length,0);
    const totalPacked = Object.values(info.mm.packed).reduce((a,b)=>a+b.length,0);
    const totalShipped = Object.values(info.mm.shipped).reduce((a,b)=>a+b.length,0);
    const pendingPackSum = Object.values(info.mm.pendingPack).reduce((a,b)=>a+b.length,0);
    const pendingShipCount = Object.values(info.mm.customerPending).reduce((sum, arr) => sum + arr.length, 0);

    // Get all bill numbers for summary
    const allPickedBills = Object.values(info.mm.picked).flat();
    const allPackedBills = Object.values(info.mm.packed).flat();
    const allShippedBills = Object.values(info.mm.shipped).flat();
    const allPendingPickBills = info.mm.pendingPick;
    const allPendingPackBills = Object.values(info.mm.pendingPack).flat();
    const allPendingShipBills = Object.values(info.mm.customerPending).flat();

    const summaryText = `Total Picked: ${totalPicked} (${allPickedBills.join(', ')})\n` +
                        `Total Packed: ${totalPacked} (${allPackedBills.join(', ')})\n` +
                        `Total Shipped: ${totalShipped} (${allShippedBills.join(', ')})\n` +
                        `Pending Pick: ${info.mm.pendingPick.length} (${allPendingPickBills.join(', ')})\n` +
                        `Pending Pack: ${pendingPackSum} (${allPendingPackBills.join(', ')})\n` +
                        `Pending Ship: ${pendingShipCount} (${allPendingShipBills.join(', ')})`;

    monthlySheet.getRange(startRow, 9).setValue(summaryText);

    // Outline border for all columns of the month (no inner gridlines)
    try {
      monthlySheet.getRange(startRow, 1, endRow - startRow + 1, headers.length)
        .setBorder(true, true, true, true, false, false);
    } catch (e) {
      console.log("Border setting failed: " + e.message);
    }
  });

  // Column widths for merged/wrapped columns
  monthlySheet.setColumnWidth(8, 420);  // H - Pending Ship
  monthlySheet.setColumnWidth(9, 520);  // I - Customer Pending and Summary

  // --- Monthly Customer Metrics (L:R) starting at row 20 ---
  const CM_HEADERS = ["Month", "Customer", "Total Orders", "Picked", "Packed", "Shipped", "Pending"];
  const CM_START_COL = 12; // L
  const CM_WIDTH = CM_HEADERS.length;
  const CM_START_ROW = 20; // L20 header, L21 data

  // Only clear and recreate headers if this is the first run (sheet is empty)
  if (monthlySheet.getLastRow() === 0) {
    // Clear previous metrics (L:R)
    monthlySheet.getRange(1, CM_START_COL, monthlySheet.getMaxRows(), CM_WIDTH).clearContent();
    // Headers at L20:R20
    monthlySheet.getRange(CM_START_ROW, CM_START_COL, 1, CM_WIDTH)
      .setValues([CM_HEADERS])
      .setFontWeight("bold")
      .setBackground("#f1f3f4");
  } else {
    // Check if headers exist, if not add them
    const existingHeaders = monthlySheet.getRange(CM_START_ROW, CM_START_COL, 1, CM_WIDTH).getValues()[0];
    if (existingHeaders[0] !== "Month" || existingHeaders[1] !== "Customer") {
      monthlySheet.getRange(CM_START_ROW, CM_START_COL, 1, CM_WIDTH).setValues([CM_HEADERS]).setFontWeight("bold");
    }
  }

  // Build per-month, per-customer metrics with bill numbers
  const customerMetricsRows = [];
  Object.keys(monthlyMetrics).sort((a,b)=>Number(a)-Number(b)).forEach(monthKey => {
    const mm = monthlyMetrics[monthKey];
    const perCustomer = {};

    mm.rowsForMonth.forEach(r => {
      const custName = r[customerCol] || "Unknown";
      const billNo = r[billNoCol];
      const manualBill = manualBillCol >= 0 ? r[manualBillCol] : "";
      const billToken = (billNo && manualBill) ? `${billNo}/${manualBill}` : (billNo || manualBill || "");
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
      
      const totalCount = cs.total > 0 ? `${cs.total} (${cs.totalBills.join(', ')})` : "0";
      const pickedCount = cs.picked > 0 ? `${cs.picked} (${cs.pickedBills.join(', ')})` : "0";
      const packedCount = cs.packed > 0 ? `${cs.packed} (${cs.packedBills.join(', ')})` : "0";
      const shippedCount = cs.shipped > 0 ? `${cs.shipped} (${cs.shippedBills.join(', ')})` : "0";
      const pendingCount = pending > 0 ? `${pending} (${pendingBills.join(', ')})` : "0";
      
      customerMetricsRows.push([monthKey, custName, totalCount, pickedCount, packedCount, shippedCount, pendingCount]);
    });
  });

  // Track where customer metrics are written
  let customerMetricsWriteStart = null;
  let customerMetricsWriteEnd = null;

  if (customerMetricsRows.length > 0) {
    // Find last non-empty row in column L (metrics area) to append after existing customer metrics
    // We need to find the actual end of customer metrics data, not just the last non-empty cell
    let lastCustomerRow = CM_START_ROW; // start from header row
    
    // Check if there's existing customer metrics data
    if (monthlySheet.getLastRow() > CM_START_ROW) {
      // Get all data from the customer metrics area (columns L:R)
      const customerData = monthlySheet.getRange(CM_START_ROW + 1, CM_START_COL, monthlySheet.getLastRow() - CM_START_ROW, CM_WIDTH).getValues();
      
      // Find the last row that has actual customer metrics data
      for (let i = customerData.length - 1; i >= 0; i--) {
        const row = customerData[i];
        // Check if this row has meaningful customer data (Month, Customer, or any other column)
        if (row[0] !== "" || row[1] !== "" || row[2] !== "" || row[3] !== "" || row[4] !== "" || row[5] !== "" || row[6] !== "") {
          lastCustomerRow = CM_START_ROW + 1 + i; // Convert back to sheet row number
          break;
        }
      }
    }
    
    const startRow = lastCustomerRow + 1;
    
    // Debug logging for customer metrics appending
    console.log("=== CUSTOMER METRICS APPENDING DEBUG ===");
    console.log("CM_START_ROW:", CM_START_ROW);
    console.log("Last customer row found:", lastCustomerRow);
    console.log("Will append starting at row:", startRow);
    console.log("Total customer rows to append:", customerMetricsRows.length);
    
    monthlySheet.getRange(startRow, CM_START_COL, customerMetricsRows.length, CM_WIDTH).setValues(customerMetricsRows);
    customerMetricsWriteStart = startRow;
    customerMetricsWriteEnd = startRow + customerMetricsRows.length - 1;

    // Merge same-month blocks in L and outline across L:R
    const lCol = CM_START_COL; // 12
    const width = CM_WIDTH;    // L:R

    let blockStart = startRow;
    let prevMonth = customerMetricsRows[0][0];

    for (let i = 1; i < customerMetricsRows.length; i++) {
      const currMonth = customerMetricsRows[i][0];
      if (currMonth !== prevMonth) {
        const blockEnd = startRow + i - 1;
        try {
          monthlySheet.getRange(blockStart, lCol, blockEnd - blockStart + 1, 1).merge().setVerticalAlignment("middle");
          monthlySheet.getRange(blockStart, lCol, blockEnd - blockStart + 1, width).setBorder(true, true, true, true, false, false);
        } catch (e) {
          console.log("Customer metrics merge failed: " + e.message);
        }
        blockStart = blockEnd + 1;
        prevMonth = currMonth;
      }
    }
    try {
      monthlySheet.getRange(blockStart, lCol, customerMetricsWriteEnd - blockStart + 1, 1).merge().setVerticalAlignment("middle");
      monthlySheet.getRange(blockStart, lCol, customerMetricsWriteEnd - blockStart + 1, width).setBorder(true, true, true, true, false, false);
    } catch (e) {
      console.log("Customer metrics final merge failed: " + e.message);
    }
  }

  // Optional formatting for readability
  monthlySheet.setColumnWidths(CM_START_COL, CM_WIDTH, 140);

  // --- Monthly Summary chart source table (T1 onward) ---
  const SUMMARY_START_COL = 20; // T
  const SUMMARY_START_ROW = 1;
  const chartHeaders = ["Month", "Picked", "Packed", "Shipped", "Pending Pick", "Pending Pack", "Pending Ship"];

  // Generate chart data using actual calculated totals, not parsed strings
  const chartMonthlyMetrics = {};
  
  // First, add current job's data with actual calculated totals
  Object.keys(monthlyMetrics).forEach(monthKey => {
    const mm = monthlyMetrics[monthKey];
    
    // Calculate actual totals directly from the data structures
    const totalPicked = Object.values(mm.picked).reduce((a,b)=>a+b.length,0);
    const totalPacked = Object.values(mm.packed).reduce((a,b)=>a+b.length,0);
    const totalShipped = Object.values(mm.shipped).reduce((a,b)=>a+b.length,0);
    const totalPendingPick = mm.pendingPick.length;
    const totalPendingPack = Object.values(mm.pendingPack).reduce((a,b)=>a+b.length,0);
    const totalPendingShip = Object.values(mm.customerPending).reduce((sum, arr) => sum + arr.length, 0);
    
    // Debug logging for current job totals
    console.log(`=== MONTH ${monthKey} TOTALS ===`);
    console.log(`Picked: ${totalPicked} (from ${Object.keys(mm.picked).length} employees)`);
    console.log(`Packed: ${totalPacked} (from ${Object.keys(mm.packed).length} employees)`);
    console.log(`Shipped: ${totalShipped} (from ${Object.keys(mm.shipped).length} employees)`);
    console.log(`Pending Pick: ${totalPendingPick}`);
    console.log(`Pending Pack: ${totalPendingPack} (from ${Object.keys(mm.pendingPack).length} employees)`);
    console.log(`Pending Ship: ${totalPendingShip} (from ${Object.keys(mm.customerPending).length} customers)`);
    
    chartMonthlyMetrics[monthKey] = {
      picked: totalPicked,
      packed: totalPacked,
      shipped: totalShipped,
      pendingPick: totalPendingPick,
      pendingPack: totalPendingPack,
      pendingShip: totalPendingShip
    };
  });
  
  // Now read existing chart data from the helper table (T:Z) if it exists
  // AND calculate summaries for historical months that don't have them
  
  // Check if there's existing chart data
  if (monthlySheet.getLastRow() >= SUMMARY_START_ROW + 1) {
    const existingChartData = monthlySheet.getRange(SUMMARY_START_ROW + 1, SUMMARY_START_COL, monthlySheet.getLastRow() - SUMMARY_START_ROW, chartHeaders.length).getValues();
    
    console.log("=== EXISTING CHART DATA FOUND ===");
    console.log("Existing chart data rows:", existingChartData.length);
    
    // Process existing chart data to get historical totals
    existingChartData.forEach(row => {
      if (row[0] && row[0] !== "") { // Month column
        const monthKey = String(row[0]);
        
        // If this month already exists in our current data, add to it
        if (chartMonthlyMetrics[monthKey]) {
          const oldPicked = chartMonthlyMetrics[monthKey].picked;
          const oldPacked = chartMonthlyMetrics[monthKey].packed;
          const oldShipped = chartMonthlyMetrics[monthKey].shipped;
          const oldPendingPick = chartMonthlyMetrics[monthKey].pendingPick;
          const oldPendingPack = chartMonthlyMetrics[monthKey].pendingPack;
          const oldPendingShip = chartMonthlyMetrics[monthKey].pendingShip;
          
          chartMonthlyMetrics[monthKey].picked += (parseInt(row[1]) || 0);
          chartMonthlyMetrics[monthKey].packed += (parseInt(row[2]) || 0);
          chartMonthlyMetrics[monthKey].shipped += (parseInt(row[3]) || 0);
          chartMonthlyMetrics[monthKey].pendingPick += (parseInt(row[4]) || 0);
          chartMonthlyMetrics[monthKey].pendingPack += (parseInt(row[5]) || 0);
          chartMonthlyMetrics[monthKey].pendingShip += (parseInt(row[6]) || 0);
          
          console.log(`Month ${monthKey}: Added historical data - Picked: ${oldPicked} + ${parseInt(row[1]) || 0} = ${chartMonthlyMetrics[monthKey].picked}`);
        } else {
          // If this month doesn't exist in current data, add it as historical
          chartMonthlyMetrics[monthKey] = {
            picked: parseInt(row[1]) || 0,
            packed: parseInt(row[2]) || 0,
            shipped: parseInt(row[3]) || 0,
            pendingPick: parseInt(row[4]) || 0,
            pendingPack: parseInt(row[5]) || 0,
            pendingShip: parseInt(row[6]) || 0
          };
          console.log(`Month ${monthKey}: Added as historical data - Picked: ${chartMonthlyMetrics[monthKey].picked}, Packed: ${chartMonthlyMetrics[monthKey].packed}`);
        }
      }
    });
    
         // AUTO-CALCULATE SUMMARIES for historical months that don't have them
     console.log("=== AUTO-CALCULATING HISTORICAL SUMMARIES ===");
     
     // Get all months from the monthly report sheet (A:I and L:R areas)
     const employeeDataRange = monthlySheet.getRange(TABLE_START_ROW + 1, 1, monthlySheet.getLastRow() - TABLE_START_ROW, 9).getValues();
     const customerDataRange = monthlySheet.getRange(CM_START_ROW + 1, CM_START_COL, monthlySheet.getLastRow() - CM_START_ROW, CM_WIDTH).getValues();
     
     // Find months that exist in the report but not in our chart data
     const reportMonths = new Set();
     
     // Collect months from employee metrics (A:I)
     employeeDataRange.forEach(row => {
       if (row[0] && row[0] !== "" && row[0] !== "Month") {
         reportMonths.add(String(row[0]));
       }
     });
     
     // Collect months from customer metrics (L:R)
     customerDataRange.forEach(row => {
       if (row[0] && row[0] !== "" && row[0] !== "Month") {
         reportMonths.add(String(row[0]));
       }
     });
     
     // For each month in the report, calculate summary if not in chart
     reportMonths.forEach(monthKey => {
       if (!chartMonthlyMetrics[monthKey]) {
         console.log(`Calculating summary for historical month: ${monthKey}`);
         
         // For historical months, read the actual calculated summary from column I (Monthly Summary)
         // Find the first row of this month in employee metrics to get the summary
         let monthSummaryRow = null;
         for (let i = 0; i < employeeDataRange.length; i++) {
           if (String(employeeDataRange[i][0]) === monthKey) {
             monthSummaryRow = employeeDataRange[i];
             break;
           }
         }
         
         if (monthSummaryRow && monthSummaryRow[8]) { // Column I (index 8) contains Monthly Summary
           const summaryText = String(monthSummaryRow[8]);
           console.log(`Reading summary from column I for month ${monthKey}: ${summaryText}`);
           
           // Parse the summary text to extract totals
           // Format: "Total Picked: X (BILL001, BILL002)\nTotal Packed: Y (BILL003, BILL004)\n..."
           const pickedMatch = summaryText.match(/Total Picked: (\d+)/);
           const packedMatch = summaryText.match(/Total Packed: (\d+)/);
           const shippedMatch = summaryText.match(/Total Shipped: (\d+)/);
           const pendingPickMatch = summaryText.match(/Pending Pick: (\d+)/);
           const pendingPackMatch = summaryText.match(/Pending Pack: (\d+)/);
           const pendingShipMatch = summaryText.match(/Pending Ship: (\d+)/);
           
           const monthPicked = pickedMatch ? parseInt(pickedMatch[1]) : 0;
           const monthPacked = packedMatch ? parseInt(packedMatch[1]) : 0;
           const monthShipped = shippedMatch ? parseInt(shippedMatch[1]) : 0;
           const monthPendingPick = pendingPickMatch ? parseInt(pendingPickMatch[1]) : 0;
           const monthPendingPack = pendingPackMatch ? parseInt(pendingPackMatch[1]) : 0;
           const monthPendingShip = pendingShipMatch ? parseInt(pendingShipMatch[1]) : 0;
           
           // Add calculated summary to chart data
           chartMonthlyMetrics[monthKey] = {
             picked: monthPicked,
             packed: monthPacked,
             shipped: monthShipped,
             pendingPick: monthPendingPick,
             pendingPack: monthPendingPack,
             pendingShip: monthPendingShip
           };
           
           console.log(`Month ${monthKey} summary read from column I: Picked: ${monthPicked}, Packed: ${monthPacked}, Shipped: ${monthShipped}, Pending Pick: ${monthPendingPick}, Pending Pack: ${monthPendingPack}, Pending Ship: ${monthPendingShip}`);
         } else {
           console.log(`No summary found in column I for month ${monthKey}, skipping`);
         }
       }
     });
  } else {
    console.log("=== NO EXISTING CHART DATA FOUND ===");
  }

  // Safety check: ensure all month keys have valid data structure
  Object.keys(chartMonthlyMetrics).forEach(monthKey => {
    const monthData = chartMonthlyMetrics[monthKey];
    if (!monthData || typeof monthData !== 'object') {
      console.log(`Warning: Invalid month data for ${monthKey}, creating default structure`);
      chartMonthlyMetrics[monthKey] = {
        picked: 0,
        packed: 0,
        shipped: 0,
        pendingPick: 0,
        pendingPack: 0,
        pendingShip: 0
      };
    } else {
      // Ensure all required properties exist
      const requiredProps = ['picked', 'packed', 'shipped', 'pendingPick', 'pendingPack', 'pendingShip'];
      requiredProps.forEach(prop => {
        if (typeof monthData[prop] !== 'number' || isNaN(monthData[prop])) {
          monthData[prop] = 0;
        }
      });
    }
  });

  const monthsSorted = Object.keys(chartMonthlyMetrics)
    .map(m => {
      const num = Number(m);
      return isNaN(num) ? 0 : num; // Convert invalid numbers to 0
    })
    .filter(m => m > 0) // Only keep valid month numbers
    .sort((a,b) => a-b);
  const monthNames = ["","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

  // Debug logging to understand what data is being processed
  console.log("=== CHART DATA DEBUG ===");
  console.log("Chart Monthly Metrics:", JSON.stringify(chartMonthlyMetrics, null, 2));
  console.log("Months Sorted:", monthsSorted);

  const chartRows = monthsSorted.map(m => {
    const monthKey = String(m);
    const mm = chartMonthlyMetrics[monthKey];
    
    // Safety check: ensure the month data exists
    if (!mm) {
      console.log(`Warning: No data found for month ${monthKey}, creating default entry`);
      chartMonthlyMetrics[monthKey] = {
        picked: 0,
        packed: 0,
        shipped: 0,
        pendingPick: 0,
        pendingPack: 0,
        pendingShip: 0
      };
    }
    
    const monthLabel = monthNames[m] || String(m);
    const finalMm = chartMonthlyMetrics[monthKey]; // Get the final data (either existing or default)
    const row = [monthLabel, finalMm.picked, finalMm.packed, finalMm.shipped, finalMm.pendingPick, finalMm.pendingPack, finalMm.pendingShip];
    console.log(`Chart Row for ${monthLabel}:`, row);
    return row;
  });

  // Final chart data summary
  console.log("=== FINAL CHART DATA SUMMARY ===");
  console.log("Total months in chart:", chartRows.length);
  chartRows.forEach((row, index) => {
    console.log(`Row ${index + 1}: ${row[0]} - Picked: ${row[1]}, Packed: ${row[2]}, Shipped: ${row[3]}, Pending Pick: ${row[4]}, Pending Pack: ${row[5]}, Pending Ship: ${row[6]}`);
  });

  // Smart chart data management: Only add new months, preserve existing ones
  if (chartRows.length > 0) {
    // Check if headers exist, if not add them
    const existingHeaders = monthlySheet.getRange(SUMMARY_START_ROW, SUMMARY_START_COL, 1, chartHeaders.length).getValues()[0];
    if (existingHeaders[0] !== "Month" || existingHeaders[1] !== "Picked") {
      monthlySheet.getRange(SUMMARY_START_ROW, SUMMARY_START_COL, 1, chartHeaders.length)
        .setValues([chartHeaders])
        .setFontWeight("bold")
        .setBackground("#f1f3f4");
    }
    
    // Find the last row with chart data to append after
    let lastChartRow = SUMMARY_START_ROW; // Start from header row
    if (monthlySheet.getLastRow() > SUMMARY_START_ROW) {
      const chartDataRange = monthlySheet.getRange(SUMMARY_START_ROW + 1, SUMMARY_START_COL, monthlySheet.getLastRow() - SUMMARY_START_ROW, chartHeaders.length).getValues();
      
      // Find the last row that has actual chart data
      for (let i = chartDataRange.length - 1; i >= 0; i--) {
        const row = chartDataRange[i];
        if (row[0] !== "" && row[0] !== "Month") { // Has month data (not header)
          lastChartRow = SUMMARY_START_ROW + 1 + i;
          break;
        }
      }
    }
    
    // Only add months that don't already exist in the chart
    const existingMonths = new Set();
    if (monthlySheet.getLastRow() > SUMMARY_START_ROW) {
      const existingChartData = monthlySheet.getRange(SUMMARY_START_ROW + 1, SUMMARY_START_COL, monthlySheet.getLastRow() - SUMMARY_START_ROW, 1).getValues();
      existingChartData.forEach(row => {
        if (row[0] && row[0] !== "" && row[0] !== "Month") {
          existingMonths.add(String(row[0]));
        }
      });
    }
    
    // Filter to only new months
    const newMonthRows = chartRows.filter(row => !existingMonths.has(String(row[0])));
    
    if (newMonthRows.length > 0) {
      console.log("=== ADDING NEW MONTHS TO CHART ===");
      console.log("New months to add:", newMonthRows.map(row => row[0]));
      console.log("Appending after row:", lastChartRow);
      
      // Append only new months
      monthlySheet.getRange(lastChartRow + 1, SUMMARY_START_COL, newMonthRows.length, chartHeaders.length)
        .setValues(newMonthRows);
    } else {
      console.log("=== NO NEW MONTHS TO ADD TO CHART ===");
      console.log("All months already exist in chart data");
    }
  }

  // Remove existing charts to avoid duplicates
  monthlySheet.getCharts().forEach(ch => monthlySheet.removeChart(ch));

  // Build and insert chart above the table (at A1)
  // Calculate the actual data range for the chart (including existing + new data)
  const actualChartDataRange = monthlySheet.getRange(SUMMARY_START_ROW + 1, SUMMARY_START_COL, monthlySheet.getLastRow() - SUMMARY_START_ROW, chartHeaders.length).getValues();
  const actualDataRows = actualChartDataRange.filter(row => row[0] !== "" && row[0] !== "Month").length;
  
  if (actualDataRows > 0) {
    const builder = monthlySheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .setPosition(1, 1, 0, 0)
      .setOption('title', 'Monthly Summary')
      .setOption('legend', { position: 'top', textStyle: { color: '#000' } })
      .setOption('hAxis', { title: 'Month' })
      .setOption('vAxis', { title: 'Count' })
      .setOption('width', 1200)
      .setOption('height', 360)
      .setNumHeaders(1); // first row are headers

    // Domain (Month) - use actual data range
    builder.addRange(monthlySheet.getRange(SUMMARY_START_ROW, SUMMARY_START_COL, actualDataRows + 1, 1));

    // Series: Picked..Pending Ship - use actual data range
    for (let c = 1; c < chartHeaders.length; c++) {
      builder.addRange(monthlySheet.getRange(SUMMARY_START_ROW, SUMMARY_START_COL + c, actualDataRows + 1, 1));
    }

    monthlySheet.insertChart(builder.build());
    
    console.log("=== CHART CREATED ===");
    console.log("Chart data range: T1 to T" + (actualDataRows + 1));
    console.log("Total months in chart:", actualDataRows);
  }

  // Update the last processed row in config sheet
  configSheet.getRange(monthlyReportRowIndex, 2).setValue(END_ROW);
  
  // Log summary
  console.log("=== MONTHLY REPORT PROCESSING SUMMARY ===");
  console.log("Rows processed: " + (END_ROW - START_ROW + 1));
  console.log("New entries added to employee metrics: " + batchValues.length);
  console.log("New entries added to customer metrics: " + customerMetricsRows.length);
  console.log("Last processed row updated to: " + END_ROW);
  console.log("Next run will start from row: " + (END_ROW + 1));
}

// Function to reset the processing counter (use this if you want to reprocess all data)
function resetMonthlyReportProcessingCounter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Automations Config");
  
  if (!configSheet) {
    console.log("Config sheet 'Automations Config' not found!");
    return;
  }
  
  // Find the Monthly Report tracking row
  const configData = configSheet.getDataRange().getValues();
  let monthlyReportRowIndex = -1;
  
  for (let i = 0; i < configData.length; i++) {
    if (configData[i][0] === "Monthly Report - New Dashboard1") {
      monthlyReportRowIndex = i + 1; // +1 because sheet rows are 1-indexed
      break;
    }
  }
  
  if (monthlyReportRowIndex === -1) {
    console.log("Monthly Report tracking row not found in Automations Config. Run generateMonthlyReportBeautified() first.");
    return;
  }
  
  configSheet.getRange(monthlyReportRowIndex, 2).setValue(1);
  console.log("Monthly Report processing counter reset to row 1. Next run will process all data.");
}

// Function to show current processing status
function showMonthlyReportProcessingStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Automations Config");
  
  if (!configSheet) {
    console.log("Config sheet 'Automations Config' not found!");
    return;
  }
  
  // Find the Monthly Report tracking row
  const configData = configSheet.getDataRange().getValues();
  let monthlyReportRowIndex = -1;
  
  for (let i = 0; i < configData.length; i++) {
    if (configData[i][0] === "Monthly Report - New Dashboard1") {
      monthlyReportRowIndex = i + 1; // +1 because sheet rows are 1-indexed
      break;
    }
  }
  
  if (monthlyReportRowIndex === -1) {
    console.log("Monthly Report tracking row not found in Automations Config. Run generateMonthlyReportBeautified() first.");
    return;
  }
  
  const lastProcessedRow = configSheet.getRange(monthlyReportRowIndex, 2).getValue() || 1;
  const dashboardSheet = ss.getSheetByName("New Dashboard1");
  
  if (dashboardSheet) {
    const totalRows = dashboardSheet.getLastRow();
    console.log("=== MONTHLY REPORT PROCESSING STATUS ===");
    console.log("Last processed row: " + lastProcessedRow);
    console.log("Total rows in New Dashboard1: " + totalRows);
    console.log("Remaining rows to process: " + Math.max(0, totalRows - lastProcessedRow));
    console.log("Next run will start from row: " + (lastProcessedRow + 1));
  } else {
    console.log("Sheet 'New Dashboard1' not found!");
  }
}
