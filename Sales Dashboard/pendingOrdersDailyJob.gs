function updatePendingOrdersEscalation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const escalationSheet = ss.getSheetByName("Pending Orders Escalation");
  
  if (!escalationSheet) {
    console.log("Sheet 'Pending Orders Escalation' not found! Run generatePendingOrdersEscalation() first.");
    return;
  }

  // Get all data from the escalation sheet
  const data = escalationSheet.getDataRange().getValues();
  if (data.length <= 1) {
    console.log("No data found in Pending Orders Escalation sheet.");
    return;
  }

  const headers = data[0];
  const rows = data.slice(1); // Skip header row
  
  // Find required column indices
  const billPickedTimestampCol = headers.indexOf("Bill Picked Timestamp");
  const escalationLevelCol = headers.indexOf("Escalation Level");
  const daysPendingCol = headers.indexOf("Days Pending");
  
  if (billPickedTimestampCol === -1 || escalationLevelCol === -1 || daysPendingCol === -1) {
    throw new Error("Required columns not found. Verify headers: Bill Picked Timestamp, Escalation Level, Days Pending");
  }

  const currentDate = new Date();
  const updatedRows = [];
  
  // Process each row to recalculate days pending and escalation level
  rows.forEach((row, index) => {
    const billPickedTimestamp = row[billPickedTimestampCol];
    let daysPending = 0;
    let escalationLevel = "";
    
    // Calculate days pending based on bill picked timestamp
    if (billPickedTimestamp) {
      try {
        const pickedDate = new Date(billPickedTimestamp);
        if (!isNaN(pickedDate.getTime())) {
          const timeDiff = currentDate.getTime() - pickedDate.getTime();
          daysPending = Math.ceil(timeDiff / (1000 * 3600 * 24));
        }
      } catch (e) {
        console.log(`Error parsing timestamp for row ${index + 2}: ${e.message}`);
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
  
  // Prepare sorted data for sheet
  const sortedData = updatedRows.map(item => item.rowData);
  
  // Write updated data back to sheet (starting from row 2, after headers)
  escalationSheet.getRange(2, 1, sortedData.length, headers.length).setValues(sortedData);
  
  // Apply conditional formatting based on new escalation levels
  applyEscalationFormatting(escalationSheet, sortedData.length);
  
  // Log summary with time range descriptions
  const summary = {
    total: sortedData.length,
    red: 0,
    orange: 0,
    yellow: 0,
    new: 0
  };
  
  updatedRows.forEach(item => {
    if (item.escalationLevel === "Red") summary.red++;
    else if (item.escalationLevel === "Orange") summary.orange++;
    else if (item.escalationLevel === "Yellow") summary.yellow++;
    else if (item.escalationLevel === "New") summary.new++;
  });
  
  console.log("=== PENDING ORDERS ESCALATION UPDATED ===");
  console.log(`Total pending bills: ${summary.total}`);
  console.log(`🔴 Red level (Critical) more than 4 days: ${summary.red}`);
  console.log(`🟠 Orange level (High) 2 to 4 days: ${summary.orange}`);
  console.log(`🟡 Yellow level (Medium) 1 to 2 days: ${summary.yellow}`);
  console.log(`🟢 New level (Low) less than 1 day: ${summary.new}`);
  
  if (summary.red > 0) {
    console.log(`⚠️  URGENT: ${summary.red} bills require immediate attention!`);
  }
  if (summary.orange > 0) {
    console.log(`⚠️  HIGH PRIORITY: ${summary.orange} bills need attention within 2-4 days`);
  }
}

function applyEscalationFormatting(sheet, dataRows) {
  if (dataRows === 0) return;
  
  // Clear existing conditional formatting rules from the sheet
  sheet.clearConditionalFormatRules();
  
  // Red level formatting (Red background, white text)
  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$R2="Red"')
    .setBackground("#ea4335")
    .setFontColor("#ffffff")
    .setRanges([sheet.getRange(2, 1, dataRows, 18)])
    .build();
  
  // Orange level formatting (Orange background, black text)
  const orangeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$R2="Orange"')
    .setBackground("#ff9800")
    .setFontColor("#000000")
    .setRanges([sheet.getRange(2, 1, dataRows, 18)])
    .build();
  
  // Yellow level formatting (Yellow background, black text)
  const yellowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$R2="Yellow"')
    .setBackground("#ffeb3b")
    .setFontColor("#000000")
    .setRanges([sheet.getRange(2, 1, dataRows, 18)])
    .build();
  
  // New level formatting (Light green background, black text)
  const newRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$R2="New"')
    .setBackground("#c8e6c9")
    .setFontColor("#000000")
    .setRanges([sheet.getRange(2, 1, dataRows, 18)])
    .build();
  
  // Apply all rules
  sheet.setConditionalFormatRules([redRule, orangeRule, yellowRule, newRule]);
}

// Function to refresh escalation data from New Dashboard1 and update the escalation sheet
function refreshEscalationFromDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName("New Dashboard1");
  
  if (!dashboardSheet) {
    console.log("Sheet 'New Dashboard1' not found!");
    return;
  }

  // Get all data from New Dashboard1
  const data = dashboardSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1); // Skip header row
  
  // Find required column indices
  const billNoCol = headers.indexOf("Bill No");
  const manualBillCol = headers.indexOf("Manual Bill Entry");
  const customerCol = headers.indexOf("Customer Name");
  const pickedByCol = headers.indexOf("Picked By");
  const billPickedTimestampCol = headers.indexOf("Bill Picked Timestamp");
  const packedByCol = headers.indexOf("Packed By");
  const packingTimestampCol = headers.indexOf("Packing Timestamp");
  const ewayBillByCol = headers.indexOf("E-way Bill By");
  const ewayBillTimestampCol = headers.indexOf("E-way Bill Timestamp");
  const shippingByCol = headers.indexOf("Shipping By");
  const shippingTimestampCol = headers.indexOf("Shipping Timestamp");
  const courierCol = headers.indexOf("Courier");
  const noOfBoxesCol = headers.indexOf("No of Boxes");
  const weightCol = headers.indexOf("Weight (kg)");
  const awbNumberCol = headers.indexOf("AWB Number");
  const awbTimestampCol = headers.indexOf("AWB Timestamp");
  const invoiceStateCol = headers.indexOf("Invoice State");
  
  // Validate required columns exist
  const requiredCols = [billNoCol, customerCol, pickedByCol, billPickedTimestampCol, packedByCol, 
                       packingTimestampCol, ewayBillByCol, ewayBillTimestampCol, shippingByCol, 
                       shippingTimestampCol, courierCol, noOfBoxesCol, weightCol, awbNumberCol, 
                       awbTimestampCol, invoiceStateCol];
  
  if (requiredCols.some(col => col === -1)) {
    throw new Error("One or more required columns not found in New Dashboard1. Verify headers.");
  }

  // Create or get the Pending Orders Escalation sheet
  let escalationSheet = ss.getSheetByName("Pending Orders Escalation");
  if (escalationSheet) {
    escalationSheet.clear();
  } else {
    escalationSheet = ss.insertSheet("Pending Orders Escalation");
  }

  // Define escalation headers
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

  // Process rows to find pending bills
  const pendingBills = [];
  const currentDate = new Date();
  
  rows.forEach(row => {
    const invoiceState = row[invoiceStateCol];
    
    // Only process bills that are not shipped
    if (invoiceState !== "Shipped") {
      const billNo = row[billNoCol];
      const manualBill = manualBillCol >= 0 ? row[manualBillCol] : "";
      const customerName = row[customerCol];
      const pickedBy = row[pickedByCol];
      const billPickedTimestamp = row[billPickedTimestampCol];
      const packedBy = row[packedByCol];
      const packingTimestamp = row[packingTimestampCol];
      const ewayBillBy = row[ewayBillByCol];
      const ewayBillTimestamp = row[ewayBillTimestampCol];
      const shippingBy = row[shippingByCol];
      const shippingTimestamp = row[shippingTimestampCol];
      const courier = row[courierCol];
      const noOfBoxes = row[noOfBoxesCol];
      const weight = row[weightCol];
      const awbNumber = row[awbNumberCol];
      const awbTimestamp = row[awbTimestampCol];
      
      // Calculate days pending based on the most recent activity
      let daysPending = 0;
      let escalationLevel = "";
      
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
      
      // Add to pending bills array
      pendingBills.push({
        billNo: billNo,
        manualBill: manualBill,
        customerName: customerName,
        pickedBy: pickedBy,
        billPickedTimestamp: billPickedTimestamp,
        packedBy: packedBy,
        packingTimestamp: packingTimestamp,
        ewayBillBy: ewayBillBy,
        ewayBillTimestamp: ewayBillTimestamp,
        shippingBy: shippingBy,
        shippingTimestamp: shippingTimestamp,
        courier: courier,
        noOfBoxes: noOfBoxes,
        weight: weight,
        awbNumber: awbNumber,
        awbTimestamp: awbTimestamp,
        daysPending: daysPending,
        escalationLevel: escalationLevel
      });
    }
  });
  
  // Sort pending bills by escalation priority (Red > Orange > Yellow > New) and then by days pending
  pendingBills.sort((a, b) => {
    const priorityOrder = { "Red": 4, "Orange": 3, "Yellow": 2, "New": 1 };
    const priorityDiff = priorityOrder[b.escalationLevel] - priorityOrder[a.escalationLevel];
    
    if (priorityDiff !== 0) {
      return priorityDiff;
    }
    
    // If same priority, sort by days pending (descending)
    return b.daysPending - a.daysPending;
  });
  
  // Prepare data for sheet
  const sheetData = pendingBills.map(bill => [
    bill.billNo,
    bill.manualBill,
    bill.customerName,
    bill.pickedBy,
    bill.billPickedTimestamp,
    bill.packedBy,
    bill.packingTimestamp,
    bill.ewayBillBy,
    bill.ewayBillTimestamp,
    bill.shippingBy,
    bill.shippingTimestamp,
    bill.courier,
    bill.noOfBoxes,
    bill.weight,
    bill.awbNumber,
    bill.awbTimestamp,
    bill.daysPending,
    bill.escalationLevel
  ]);
  
  // Write data to sheet
  if (sheetData.length > 0) {
    escalationSheet.getRange(2, 1, sheetData.length, escalationHeaders.length).setValues(sheetData);
    
    // Apply conditional formatting based on escalation level
    applyEscalationFormatting(escalationSheet, sheetData.length);
    
    // Auto-resize columns for better readability
    escalationSheet.autoResizeColumns(1, escalationHeaders.length);
    
    // Set column widths for timestamp columns
    escalationSheet.setColumnWidth(5, 180);  // Bill Picked Timestamp
    escalationSheet.setColumnWidth(7, 180);  // Packing Timestamp
    escalationSheet.setColumnWidth(9, 180);  // E-way Bill Timestamp
    escalationSheet.setColumnWidth(11, 180); // Shipping Timestamp
    escalationSheet.setColumnWidth(16, 180); // AWB Timestamp
    
    // Freeze header row
    escalationSheet.setFrozenRows(1);
    
    // Add borders to the data
    escalationSheet.getRange(1, 1, sheetData.length + 1, escalationHeaders.length)
      .setBorder(true, true, true, true, true, true);
    
    console.log(`=== PENDING ORDERS ESCALATION REFRESHED ===`);
    console.log(`Total pending bills: ${sheetData.length}`);
    console.log(`Red level: ${pendingBills.filter(b => b.escalationLevel === "Red").length}`);
    console.log(`Orange level: ${pendingBills.filter(b => b.escalationLevel === "Orange").length}`);
    console.log(`Yellow level: ${pendingBills.filter(b => b.escalationLevel === "Yellow").length}`);
    console.log(`New level: ${pendingBills.filter(b => b.escalationLevel === "New").length}`);
  } else {
    console.log("No pending bills found!");
  }
}

// Function to show current escalation status
function showEscalationStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const escalationSheet = ss.getSheetByName("Pending Orders Escalation");
  
  if (!escalationSheet) {
    console.log("Pending Orders Escalation sheet not found. Run refreshEscalationFromDashboard() first.");
    return;
  }
  
  const data = escalationSheet.getDataRange().getValues();
  if (data.length <= 1) {
    console.log("No pending orders data found.");
    return;
  }
  
  const rows = data.slice(1); // Skip header
  
  const summary = {
    total: rows.length,
    red: 0,
    orange: 0,
    yellow: 0,
    new: 0
  };
  
  rows.forEach(row => {
    const level = row[17]; // Escalation Level column
    if (level === "Red") summary.red++;
    else if (level === "Orange") summary.orange++;
    else if (level === "Yellow") summary.yellow++;
    else if (level === "New") summary.new++;
  });
  
  console.log("=== CURRENT ESCALATION STATUS ===");
  console.log(`Total Pending Bills: ${summary.total}`);
  console.log(`🔴 Red Level (Critical): ${summary.red}`);
  console.log(`🟠 Orange Level (High): ${summary.orange}`);
  console.log(`🟡 Yellow Level (Medium): ${summary.yellow}`);
  console.log(`🟢 New Level (Low): ${summary.new}`);
  
  if (summary.red > 0) {
    console.log(`⚠️  URGENT: ${summary.red} bills require immediate attention!`);
  }
  if (summary.orange > 0) {
    console.log(`⚠️  HIGH PRIORITY: ${summary.orange} bills need attention within 2-4 days`);
  }
}
