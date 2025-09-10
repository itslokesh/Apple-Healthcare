function hourlyReportGenerator() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName("Automation Dashboard");
  const configSheetName = "Automations Config";
  
  // Get current execution time in human readable format
  const executionTime = new Date();
  const timeString = executionTime.toLocaleString('en-US', {
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: true,
    timeZone: 'Asia/Kolkata'
  });
  
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

  // Check if Daily Report tracking row exists, if not add it
  const configData = configSheet.getDataRange().getValues();
  let dailyReportRowIndex = -1;
  
  for (let i = 0; i < configData.length; i++) {
    if (configData[i][0] === "Hourly Report") {
      dailyReportRowIndex = i + 1; // +1 because sheet rows are 1-indexed
      break;
    }
  }
  
  if (dailyReportRowIndex === -1) {
    // Add new row for Daily Report tracking
    const newRowIndex = configSheet.getLastRow() + 1;
    configSheet.getRange(newRowIndex, 1, 1, 2).setValues([
      ["Hourly Report", 1]
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
  let dailySheet = ss.getSheetByName("Hourly Report 2025");
  if (!dailySheet) dailySheet = ss.insertSheet("Hourly Report 2025");
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
        `${dayKey} (${timeString})`,
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
   // Scan multiple columns to find the actual last row with data, accounting for merged cells
   const colAValues = dailySheet.getRange(1, 1, dailySheet.getMaxRows(), 1).getValues();
   const colBValues = dailySheet.getRange(1, 2, dailySheet.getMaxRows(), 1).getValues();
   const colCValues = dailySheet.getRange(1, 3, dailySheet.getMaxRows(), 1).getValues();
   
   let lastEmployeeRow = 1; // header row default
   for (let i = colAValues.length - 1; i >= 1; i--) { // start from bottom, skip row 0 header
     // Check if any of the key columns have meaningful data (not just merged cell content)
     if (colAValues[i][0] !== "" || colBValues[i][0] !== "" || colCValues[i][0] !== "") {
       lastEmployeeRow = i + 1; 
       break; 
     }
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
      
             customerMetricsRows.push([`${dayKey} (${timeString})`, custName, totalCount, pickedCount, packedCount, shippedCount, pendingCount]);
     });
   });
   


  // Track where customer metrics are written
  let customerMetricsWriteStart = null;
  let customerMetricsWriteEnd = null;

     if (customerMetricsRows.length > 0) {
     // Find last non-empty row in customer metrics area (columns L-R) to append after existing customer metrics
     // Scan multiple columns to find the actual last row with data, accounting for merged cells
     const lastRowToCheck = Math.max(2, dailySheet.getMaxRows());
     const colLValues = dailySheet.getRange(2, metricsStartCol, lastRowToCheck - 1, 1).getValues();
     const colMValues = dailySheet.getRange(2, metricsStartCol + 1, lastRowToCheck - 1, 1).getValues();
     const colNValues = dailySheet.getRange(2, metricsStartCol + 2, lastRowToCheck - 1, 1).getValues();
     
     let lastNonEmpty = 1; // header row default
     for (let i = colLValues.length - 1; i >= 0; i--) {
       // Check if any of the key customer metrics columns have meaningful data
       if (colLValues[i][0] !== "" || colMValues[i][0] !== "" || colNValues[i][0] !== "") {
         lastNonEmpty = i + 2; 
         break; 
       }
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
  
  // Update Pending Orders Escalation sheet INLINE
  console.log("=== UPDATING PENDING ORDERS ESCALATION INLINE ===");
  console.log("Rows to process:", rows.length);
  console.log("First row sample:", rows[0]);
  
  let updatedRows = [];
  try {
    console.log("=== ESCALATION LOGIC STARTED ===");
    
    // Get or create escalation sheet
    let escalationSheet = ss.getSheetByName("Pending Orders Escalation");
    if (!escalationSheet) {
      console.log("Creating new Pending Orders Escalation sheet");
      escalationSheet = ss.insertSheet("Pending Orders Escalation");
      
      // Define escalation headers
      const escalationHeaders = [
        "Bill No", "Manual Bill Entry", "Customer Name", "Picked By", "Bill Picked Timestamp",
        "Packed By", "Packing Timestamp", "E-way Bill By", "E-way Bill Timestamp",
        "Shipping By", "Shipping Timestamp", "Courier", "No of Boxes", "Weight (kg)",
        "AWB Number", "AWB Timestamp", "Days Pending", "Escalation Level"
      ];
      
      // Set headers
      escalationSheet.getRange(1, 1, 1, escalationHeaders.length)
        .setValues([escalationHeaders])
        .setFontWeight("bold")
        .setBackground("#f1f3f4");
      
      // Auto-resize columns
      escalationSheet.autoResizeColumns(1, escalationHeaders.length);
      escalationSheet.setFrozenRows(1);
      console.log("New escalation sheet created with headers");
    } else {
      console.log("Using existing Pending Orders Escalation sheet");
    }
    
         // Get existing escalation data
     const escalationData = escalationSheet.getDataRange().getValues();
     const escalationHeaders = escalationData[0];
     const existingRows = escalationData.slice(1);
     
     console.log("Escalation sheet has", existingRows.length, "existing rows");
     
     // Define column indices for escalation logic
     const billPickedTimestampCol = headersRow.indexOf("Bill Picked Timestamp");
     const packingTimestampCol = headersRow.indexOf("Packing Timestamp");
     const ewayBillTimestampCol = headersRow.indexOf("E-way Bill Timestamp");
     const shippingTimestampCol = headersRow.indexOf("Shipping Timestamp");
     
     console.log("Escalation column indices:", {
       billPicked: billPickedTimestampCol,
       packing: packingTimestampCol,
       ewayBill: ewayBillTimestampCol,
       shipping: shippingTimestampCol
     });
     
     const currentDate = new Date();
     const updatedRows = [];
     
     // Process new rows from dashboard
     console.log("Processing", rows.length, "new rows from dashboard");
     
     rows.forEach((row, index) => {
       const billNo = row[billNoCol];
       const invoiceState = row[invoiceStateCol];
       
       
       // Only add bills that are NOT shipped and NOT already in escalation sheet
       if (invoiceState !== "Shipped" && !existingRows.some(existingRow => existingRow[0] === billNo)) {         
         
         // Calculate days pending based on most recent activity
         let daysPending = 0;
         let escalationLevel = "New";
         
         const billPickedTimestamp = row[billPickedTimestampCol];
         const packingTimestamp = row[packingTimestampCol];
         const ewayBillTimestamp = row[ewayBillTimestampCol];
         const shippingTimestamp = row[shippingTimestampCol];
        
        // Use the most recent timestamp to calculate days pending
        let mostRecentTimestamp = null;
        
        if (billPickedTimestamp) {
          try {
            const date = new Date(billPickedTimestamp);
            if (!isNaN(date.getTime())) {
              mostRecentTimestamp = date;
            }
          } catch (e) {
            console.log(`Error parsing picked timestamp for bill ${billNo}: ${e.message}`);
          }
        }
        
        if (packingTimestamp) {
          try {
            const date = new Date(packingTimestamp);
            if (!isNaN(date.getTime()) && (!mostRecentTimestamp || date > mostRecentTimestamp)) {
              mostRecentTimestamp = date;
            }
          } catch (e) {
            console.log(`Error parsing packing timestamp for bill ${billNo}: ${e.message}`);
          }
        }
        
        if (ewayBillTimestamp) {
          try {
            const date = new Date(ewayBillTimestamp);
            if (!isNaN(date.getTime()) && (!mostRecentTimestamp || date > mostRecentTimestamp)) {
              mostRecentTimestamp = date;
            }
          } catch (e) {
            console.log(`Error parsing e-way bill timestamp for bill ${billNo}: ${e.message}`);
          }
        }
        
        if (shippingTimestamp) {
          try {
            const date = new Date(shippingTimestamp);
            if (!isNaN(date.getTime()) && (!mostRecentTimestamp || date > mostRecentTimestamp)) {
              mostRecentTimestamp = date;
            }
          } catch (e) {
            console.log(`Error parsing shipping timestamp for bill ${billNo}: ${e.message}`);
          }
        }
        
        // Calculate days pending
        if (mostRecentTimestamp) {
          const timeDiff = currentDate.getTime() - mostRecentTimestamp.getTime();
          daysPending = Math.ceil(timeDiff / (1000 * 3600 * 24));
          
          // Determine escalation level
          if (daysPending >= 1 && daysPending < 2) {
            escalationLevel = "Yellow";
          } else if (daysPending >= 2 && daysPending <= 4) {
            escalationLevel = "Orange";
          } else if (daysPending > 4) {
            escalationLevel = "Red";
          } else {
            escalationLevel = "New";
          }
        }
                
                 // Create new row data (excluding Invoice State, Day, Month)
         const newRowData = [
           row[billNoCol],
           row[manualBillCol] || "",
           row[customerCol],
           row[pickedByCol],
           row[billPickedTimestampCol],
           row[packedByCol],
           row[packingTimestampCol],
           row[headersRow.indexOf("E-way Bill By")],
           row[ewayBillTimestampCol],
           row[shippingByCol],
           row[shippingTimestampCol],
           row[headersRow.indexOf("Courier")],
           row[headersRow.indexOf("No of Boxes")],
           row[headersRow.indexOf("Weight (kg)")],
           row[headersRow.indexOf("AWB Number")],
           row[headersRow.indexOf("AWB Timestamp")],
           daysPending,
           escalationLevel
         ];
        
        updatedRows.push({
          rowData: newRowData,
          daysPending: daysPending,
          escalationLevel: escalationLevel,
          isNew: true
        });
      } else {
        console.log(`Skipping bill ${billNo}: ${invoiceState === "Shipped" ? "Already shipped" : "Already in escalation sheet"}`);
      }
    });
    
    console.log("Total rows to add:", updatedRows.length);
    
    // Add new rows to escalation sheet
    if (updatedRows.length > 0) {
      const sheetData = updatedRows.map(item => item.rowData);
      console.log("Writing", sheetData.length, "rows to escalation sheet");
      
      // Append new rows after existing data
      const startRow = Math.max(2, escalationSheet.getLastRow() + 1);
      escalationSheet.getRange(startRow, 1, sheetData.length, escalationHeaders.length).setValues(sheetData);
      
      // Apply color coding inline
      console.log("Applying color coding starting from row:", startRow);
      
      let currentRow = startRow;
      
      // Red level (Critical) - Red background, white text
      const redRows = updatedRows.filter(item => item.escalationLevel === "Red");
      if (redRows.length > 0) {
        const redRange = escalationSheet.getRange(currentRow, 1, redRows.length, 18);
        redRange.setBackground("#ea4335").setFontColor("#ffffff");
        console.log(`Applied red color to ${redRows.length} rows starting from row ${currentRow}`);
        currentRow += redRows.length;
      }
      
      // Orange level (High) - Orange background, black text
      const orangeRows = updatedRows.filter(item => item.escalationLevel === "Orange");
      if (orangeRows.length > 0) {
        const orangeRange = escalationSheet.getRange(currentRow, 1, orangeRows.length, 18);
        orangeRange.setBackground("#ff9800").setFontColor("#000000");
        console.log(`Applied orange color to ${orangeRows.length} rows starting from row ${currentRow}`);
        currentRow += orangeRows.length;
      }
      
      // Yellow level (Medium) - Yellow background, black text
      const yellowRows = updatedRows.filter(item => item.escalationLevel === "Yellow");
      if (yellowRows.length > 0) {
        const yellowRange = escalationSheet.getRange(currentRow, 1, yellowRows.length, 18);
        yellowRange.setBackground("#ffeb3b").setFontColor("#000000");
        console.log(`Applied yellow color to ${yellowRows.length} rows starting from row ${currentRow}`);
        currentRow += yellowRows.length;
      }
      
      // New level (Low) - Light green background, black text
      const newRows = updatedRows.filter(item => item.escalationLevel === "New");
      if (newRows.length > 0) {
        const newRange = escalationSheet.getRange(currentRow, 1, newRows.length, 18);
        newRange.setBackground("#c8e6c9").setFontColor("#000000");
        console.log(`Applied green color to ${newRows.length} rows starting from row ${currentRow}`);
      }
      
      console.log("=== ESCALATION SHEET UPDATED SUCCESSFULLY ===");
      console.log(`Total pending bills: ${updatedRows.length}`);
      console.log(`🔴 Red level (Critical) more than 4 days: ${updatedRows.filter(b => b.escalationLevel === "Red").length}`);
      console.log(`🟠 Orange level (High) 2 to 4 days: ${updatedRows.filter(b => b.escalationLevel === "Orange").length}`);
      console.log(`🟡 Yellow level (Medium) 1 to 2 days: ${updatedRows.filter(b => b.escalationLevel === "Yellow").length}`);
      console.log(`🟢 New level (Low) less than 1 day: ${updatedRows.filter(b => b.escalationLevel === "New").length}`);
    } else {
      console.log("No new pending bills to add");
    }
    
    console.log("=== ESCALATION LOGIC COMPLETED SUCCESSFULLY ===");
    
  } catch (error) {
    console.log("=== ERROR IN ESCALATION LOGIC ===");
    console.log("Error:", error.message);
    console.log("Stack:", error.stack);
  }
  
  // Build a full snapshot of pending escalations from the sheet for the summary
  let allEscalationRows = [];
  try {
    const escSheet = ss.getSheetByName("Pending Orders Escalation");
    if (escSheet) {
      const escData = escSheet.getDataRange().getValues();
      if (escData.length > 1) {
        // Map rows to objects with at least billNo and escalationLevel
        allEscalationRows = escData.slice(1).map(r => ({ billNo: r[0], escalationLevel: r[17], rowData: r }));
      }
    }
  } catch (e) {
    console.log("Warning: Could not read full escalation snapshot:", e.message);
  }
  
  // Generate WhatsApp Summary Text using full escalation snapshot
  console.log("=== GENERATING WHATSAPP SUMMARY ===");
  const whatsappSummary = generateWhatsAppSummary(dailyMetrics, allEscalationRows, timeString);
  console.log("WhatsApp Summary Generated:");
  console.log(whatsappSummary);
  
  // Send WhatsApp summary via Twilio (uses Script Properties for config)
  try {
    const sentCount = sendWhatsAppSummaryViaTwilio(whatsappSummary);
    console.log(`WhatsApp summary sent to ${sentCount} recipient(s).`);
  } catch (e) {
    console.log("Error sending WhatsApp summary:", e && e.message ? e.message : e);
  }
  
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
  console.log("=== ESCALATION FUNCTION STARTED ===");
  console.log("Received rows:", processedRows.length);
  
  // Simple test to see if function is working
  console.log("Function is executing...");
  
  // Test basic operations
  console.log("Testing basic operations...");
  const testArray = [1, 2, 3];
  console.log("Array test:", testArray.length);
  
  try {
    console.log("Entered try block...");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log("Got spreadsheet object");
    
    // Get or create escalation sheet
    let escalationSheet = ss.getSheetByName("Pending Orders Escalation");
    if (!escalationSheet) {
      console.log("Creating new Pending Orders Escalation sheet");
      escalationSheet = ss.insertSheet("Pending Orders Escalation");
      
      // Define escalation headers
      const escalationHeaders = [
        "Bill No", "Manual Bill Entry", "Customer Name", "Picked By", "Bill Picked Timestamp",
        "Packed By", "Packing Timestamp", "E-way Bill By", "E-way Bill Timestamp",
        "Shipping By", "Shipping Timestamp", "Courier", "No of Boxes", "Weight (kg)",
        "AWB Number", "AWB Timestamp", "Days Pending", "Escalation Level"
      ];
      
      // Set headers
      escalationSheet.getRange(1, 1, 1, escalationHeaders.length)
        .setValues([escalationHeaders])
        .setFontWeight("bold")
        .setBackground("#f1f3f4");
      
      // Auto-resize columns
      escalationSheet.autoResizeColumns(1, escalationHeaders.length);
      escalationSheet.setFrozenRows(1);
      console.log("New escalation sheet created with headers");
    } else {
      console.log("Using existing Pending Orders Escalation sheet");
    }
    
    // Get dashboard data for comparison
    const dashboardSheet = ss.getSheetByName("New Dashboard1");
    const dashboardData = dashboardSheet.getDataRange().getValues();
    const dashboardHeaders = dashboardData[0];
    
    // Find column indices
    const invoiceStateCol = dashboardHeaders.indexOf("Invoice State");
    const billPickedTimestampCol = dashboardHeaders.indexOf("Bill Picked Timestamp");
    const packingTimestampCol = dashboardHeaders.indexOf("Packing Timestamp");
    const ewayBillTimestampCol = dashboardHeaders.indexOf("E-way Bill Timestamp");
    const shippingTimestampCol = dashboardHeaders.indexOf("Shipping Timestamp");
    
    console.log("Dashboard columns found:", {
      invoiceState: invoiceStateCol,
      billPicked: billPickedTimestampCol,
      packing: packingTimestampCol,
      ewayBill: ewayBillTimestampCol,
      shipping: shippingTimestampCol
    });
    
    // Get existing escalation data
    const escalationData = escalationSheet.getDataRange().getValues();
    const escalationHeaders = escalationData[0];
    const existingRows = escalationData.slice(1);
    
    console.log("Escalation sheet has", existingRows.length, "existing rows");
    
    const currentDate = new Date();
    updatedRows = [];
    
    // Process new rows from dashboard
    console.log("Processing", processedRows.length, "new rows from dashboard");
    
    processedRows.forEach((row, index) => {
      const billNo = row[dashboardHeaders.indexOf("Bill No")];
      const invoiceState = row[invoiceStateCol];
      
      console.log(`Row ${index + 1}: Bill ${billNo}, State: ${invoiceState}`);
      
      // Only add bills that are NOT shipped and NOT already in escalation sheet
      if (invoiceState !== "Shipped" && !existingRows.some(existingRow => existingRow[0] === billNo)) {
        console.log(`Adding new pending bill: ${billNo}`);
        
        // Calculate days pending based on most recent activity
        let daysPending = 0;
        let escalationLevel = "New";
        
        const billPickedTimestamp = row[billPickedTimestampCol];
        const packingTimestamp = row[packingTimestampCol];
        const ewayBillTimestamp = row[ewayBillTimestampCol];
        const shippingTimestamp = row[shippingTimestampCol];
        
        // Use the most recent timestamp to calculate days pending
        let mostRecentTimestamp = null;
        
        if (billPickedTimestamp) {
          try {
            const date = new Date(billPickedTimestamp);
            if (!isNaN(date.getTime())) {
              mostRecentTimestamp = date;
            }
          } catch (e) {
            console.log(`Error parsing picked timestamp for bill ${billNo}: ${e.message}`);
          }
        }
        
        if (packingTimestamp) {
          try {
            const date = new Date(packingTimestamp);
            if (!isNaN(date.getTime()) && (!mostRecentTimestamp || date > mostRecentTimestamp)) {
              mostRecentTimestamp = date;
            }
          } catch (e) {
            console.log(`Error parsing packing timestamp for bill ${billNo}: ${e.message}`);
          }
        }
        
        if (ewayBillTimestamp) {
          try {
            const date = new Date(ewayBillTimestamp);
            if (!isNaN(date.getTime()) && (!mostRecentTimestamp || date > mostRecentTimestamp)) {
              mostRecentTimestamp = date;
            }
          } catch (e) {
            console.log(`Error parsing e-way bill timestamp for bill ${billNo}: ${e.message}`);
          }
        }
        
        if (shippingTimestamp) {
          try {
            const date = new Date(shippingTimestamp);
            if (!isNaN(date.getTime()) && (!mostRecentTimestamp || date > mostRecentTimestamp)) {
              mostRecentTimestamp = date;
            }
          } catch (e) {
            console.log(`Error parsing shipping timestamp for bill ${billNo}: ${e.message}`);
          }
        }
        
        // Calculate days pending
        if (mostRecentTimestamp) {
          const timeDiff = currentDate.getTime() - mostRecentTimestamp.getTime();
          daysPending = Math.ceil(timeDiff / (1000 * 3600 * 24));
          
          // Determine escalation level
          if (daysPending >= 1 && daysPending < 2) {
            escalationLevel = "Yellow";
          } else if (daysPending >= 2 && daysPending <= 4) {
            escalationLevel = "Orange";
          } else if (daysPending > 4) {
            escalationLevel = "Red";
          } else {
            escalationLevel = "New";
          }
        }
        
        console.log(`Bill ${billNo}: ${daysPending} days pending, Level: ${escalationLevel}`);
        
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
      } else {
        console.log(`Skipping bill ${billNo}: ${invoiceState === "Shipped" ? "Already shipped" : "Already in escalation sheet"}`);
      }
    });
    
    console.log("Total rows to add:", updatedRows.length);
    
    // Add new rows to escalation sheet
    if (updatedRows.length > 0) {
      const sheetData = updatedRows.map(item => item.rowData);
      console.log("Writing", sheetData.length, "rows to escalation sheet");
      
      // Append new rows after existing data
      const startRow = Math.max(2, escalationSheet.getLastRow() + 1);
      escalationSheet.getRange(startRow, 1, sheetData.length, escalationHeaders.length).setValues(sheetData);
      
      // Apply color coding
      applyEscalationColorCoding(escalationSheet, updatedRows, startRow);
      
      console.log("=== ESCALATION SHEET UPDATED SUCCESSFULLY ===");
      console.log(`Total pending bills: ${updatedRows.length}`);
      console.log(`🔴 Red level (Critical) more than 4 days: ${updatedRows.filter(b => b.escalationLevel === "Red").length}`);
      console.log(`🟠 Orange level (High) 2 to 4 days: ${updatedRows.filter(b => b.escalationLevel === "Orange").length}`);
      console.log(`🟡 Yellow level (Medium) 1 to 2 days: ${updatedRows.filter(b => b.escalationLevel === "Yellow").length}`);
      console.log(`🟢 New level (Low) less than 1 day: ${updatedRows.filter(b => b.escalationLevel === "New").length}`);
    } else {
      console.log("No new pending bills to add");
    }
    
  } catch (error) {
    console.log("ERROR in escalation function:", error.message);
    console.log("Stack:", error.stack);
    throw error;
  }
  
  console.log("=== FUNCTION COMPLETING SUCCESSFULLY ===");
}

// Function to apply color coding and group rows by escalation level
function applyEscalationColorCoding(sheet, updatedRows, startRow) {
  if (updatedRows.length === 0) return;
  
  console.log("Applying color coding starting from row:", startRow);
  
  // Group rows by escalation level and apply colors
  let currentRow = startRow;
  
  // Red level (Critical) - Red background, white text
  const redRows = updatedRows.filter(item => item.escalationLevel === "Red");
  if (redRows.length > 0) {
    const redRange = sheet.getRange(currentRow, 1, redRows.length, 18);
    redRange.setBackground("#ea4335").setFontColor("#ffffff");
    console.log(`Applied red color to ${redRows.length} rows starting from row ${currentRow}`);
    currentRow += redRows.length;
  }
  
  // Orange level (High) - Orange background, black text
  const orangeRows = updatedRows.filter(item => item.escalationLevel === "Orange");
  if (orangeRows.length > 0) {
    const orangeRange = sheet.getRange(currentRow, 1, orangeRows.length, 18);
    orangeRange.setBackground("#ff9800").setFontColor("#000000");
    console.log(`Applied orange color to ${orangeRows.length} rows starting from row ${currentRow}`);
    currentRow += orangeRows.length;
  }
  
  // Yellow level (Medium) - Yellow background, black text
  const yellowRows = updatedRows.filter(item => item.escalationLevel === "Yellow");
  if (yellowRows.length > 0) {
    const yellowRange = sheet.getRange(currentRow, 1, yellowRows.length, 18);
    yellowRange.setBackground("#ffeb3b").setFontColor("#000000");
    console.log(`Applied yellow color to ${yellowRows.length} rows starting from row ${currentRow}`);
    currentRow += yellowRows.length;
  }
  
  // New level (Low) - Light green background, black text
  const newRows = updatedRows.filter(item => item.escalationLevel === "New");
  if (newRows.length > 0) {
    const newRange = sheet.getRange(currentRow, 1, newRows.length, 18);
    newRange.setBackground("#c8e6c9").setFontColor("#000000");
    console.log(`Applied green color to ${newRows.length} rows starting from row ${currentRow}`);
  }
  
  console.log("Color coding completed");
}

// Function to generate WhatsApp summary text
function generateWhatsAppSummary(dailyMetrics, escalationRows, timeString) {
  const currentDate = new Date().toLocaleDateString('en-IN', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric'
  });
  
  // Header optimized for WhatsApp readability
  let summary = `*HOURLY REPORT*\n`;
  summary += `Date: ${currentDate}\n`;
  summary += `Time: ${timeString}\n`;
  summary += `${'-'.repeat(7)}\n\n`;
  
  // Employee Performance Summary
  summary += `*EMPLOYEE SUMMARY*\n`;
  summary += `${'-'.repeat(7)}\n`;
  
  const employeeStats = {};
  let totalPicked = 0, totalPacked = 0, totalShipped = 0;
  
  Object.keys(dailyMetrics).forEach(dayKey => {
    const dm = dailyMetrics[dayKey];
    
    // Aggregate employee stats
    Object.keys(dm.picked).forEach(emp => {
      if (!employeeStats[emp]) employeeStats[emp] = { picked: 0, packed: 0, shipped: 0 };
      employeeStats[emp].picked += dm.picked[emp].length;
      totalPicked += dm.picked[emp].length;
    });
    
    Object.keys(dm.packed).forEach(emp => {
      if (!employeeStats[emp]) employeeStats[emp] = { picked: 0, packed: 0, shipped: 0 };
      employeeStats[emp].packed += dm.packed[emp].length;
      totalPacked += dm.packed[emp].length;
    });
    
    Object.keys(dm.shipped).forEach(emp => {
      if (!employeeStats[emp]) employeeStats[emp] = { picked: 0, packed: 0, shipped: 0 };
      employeeStats[emp].shipped += dm.shipped[emp].length;
      totalShipped += dm.shipped[emp].length;
    });
  });
  
  // Top performers
  const sortedEmployees = Object.entries(employeeStats)
    .sort((a, b) => (b[1].picked + b[1].packed + b[1].shipped) - (a[1].picked + a[1].packed + a[1].shipped))
    .slice(0, 5);
  
  if (sortedEmployees.length > 0) {
    sortedEmployees.forEach(([emp, stats], index) => {
      summary += `${index + 1}. ${emp}: (P:${stats.picked} | Pk:${stats.packed} | S:${stats.shipped})\n`;
    });
  }
  
  summary += `\n*Overall Totals*\n`;
  summary += `• Picked: ${totalPicked} orders\n`;
  summary += `• Packed: ${totalPacked} orders\n`;
  summary += `• Shipped: ${totalShipped} orders\n\n`;
  
  // Customer Summary
  summary += `*CUSTOMER SUMMARY*\n`;
  summary += `${'-'.repeat(7)}\n`;
  
  const customerStats = {};
  let totalCustomerOrders = 0;
  
  Object.keys(dailyMetrics).forEach(dayKey => {
    const dm = dailyMetrics[dayKey];
    dm.rowsForDay.forEach(row => {
      const customerCol = 2; // Assuming customer name is at index 2
      const customer = row[customerCol] || "Unknown";
      const invoiceStateCol = 16; // Assuming invoice state is at index 16
      const isShipped = row[invoiceStateCol] === "Shipped";
      
      if (!customerStats[customer]) {
        customerStats[customer] = { total: 0, shipped: 0, pending: 0 };
      }
      customerStats[customer].total += 1;
      totalCustomerOrders += 1;
      
      if (isShipped) {
        customerStats[customer].shipped += 1;
      } else {
        customerStats[customer].pending += 1;
      }
    });
  });
  
  // Top customers by order volume
  const sortedCustomers = Object.entries(customerStats)
    .sort((a, b) => b[1].total - a[1].total)
    .slice(0, 5);
  
  if (sortedCustomers.length > 0) {
    sortedCustomers.forEach(([customer, stats], index) => {
      const completionRate = stats.total > 0 ? Math.round((stats.shipped / stats.total) * 100) : 0;
      summary += `${index + 1}. ${customer}: ${stats.total} orders (${completionRate}% shipped)\n`;
    });
  }
  
  summary += `\n*Customer Totals*\n`;
  summary += `• Total Orders: ${totalCustomerOrders}\n`;
  
  // Pending Escalations Summary
  summary += `*PENDING ESCALATIONS*\n`;
  summary += `${'-'.repeat(7)}\n`;
  
  if (escalationRows && escalationRows.length > 0) {
    const redItems = escalationRows.filter(r => r.escalationLevel === "Red");
    const redCount = redItems.length;
    // Sort oldest first using Days Pending (index 16 in escalation sheet rowData)
    redItems.sort((a, b) => {
      const da = Array.isArray(a.rowData) ? (parseFloat(a.rowData[16]) || 0) : (a.daysPending || 0);
      const db = Array.isArray(b.rowData) ? (parseFloat(b.rowData[16]) || 0) : (b.daysPending || 0);
      return db - da; // higher days pending first => oldest first
    });
    const redBills = redItems.map(r => Array.isArray(r.rowData) ? r.rowData[0] : (r.billNo || r[0] || "")).filter(Boolean);
    const orangeCount = escalationRows.filter(r => r.escalationLevel === "Orange").length;
    const yellowCount = escalationRows.filter(r => r.escalationLevel === "Yellow").length;
    const newCount = escalationRows.filter(r => r.escalationLevel === "New").length;
    
    summary += `🔴 *CRITICAL (>4 days):* ${redCount} orders\n`;
    summary += `🟠 *HIGH (2-4 days):* ${orangeCount} orders\n`;
    summary += `🟡 *MEDIUM (1-2 days):* ${yellowCount} orders\n`;
    summary += `🟢 *NEW (<1 day):* ${newCount} orders\n`;
    summary += `\n📊 *Total Pending:* ${escalationRows.length} orders\n`;
    
    if (redCount > 0) {
      summary += `\n*URGENT ACTION REQUIRED*\n`;
      summary += `${redCount} orders are critically overdue: ${redBills.join(', ')}\n`;
    }
  } else {
    summary += `✅ No new pending escalations\n`;
  }
  
  return summary;
}
