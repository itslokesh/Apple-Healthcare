function generateDailyReportBeautified() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName("New Dashboard");
  if (!dashboardSheet) return;

  // --- CHUNK SETTINGS ---
  const START_ROW = 2; // adjust per batch
  const END_ROW = 7210; // adjust per batch

  const data = dashboardSheet.getDataRange().getValues();
  if (data.length < START_ROW) return;

  // --- Column indices dynamically ---
  const headersRow = data[0];
  const dayCol = headersRow.indexOf("Day - Bill picked");
  const monthCol = headersRow.indexOf("Month - Bill picked");
  const pickedByCol = headersRow.indexOf("Picked by");
  const packedByCol = headersRow.indexOf("Packed by");
  const shippingByCol = headersRow.indexOf("Shipping By");
  const invoiceStateCol = headersRow.indexOf("Invoice state");
  const customerCol = headersRow.indexOf("Customer name");
  const billNoCol = headersRow.indexOf("Bill No");

  if ([dayCol, monthCol, pickedByCol, packedByCol, shippingByCol, invoiceStateCol, customerCol, billNoCol].some(c => c === -1)) {
    throw new Error("One or more required columns not found in New Dashboard.");
  }

  // --- Daily Report Sheet ---
  let dailySheet = ss.getSheetByName("Daily Report 2025");
  if (!dailySheet) dailySheet = ss.insertSheet("Daily Report 2025");
  const headers = ["Date", "Employee", "Picked Count", "Packed Count", "Shipped Count", "Pending Pick", "Pending Pack", "Customer Pending", "Daily Summary"];
  if (dailySheet.getLastRow() === 0) dailySheet.appendRow(headers);

  const rows = data.slice(START_ROW - 1, Math.min(END_ROW, data.length));
  const dailyMetrics = {}; // dayKey -> metrics

  // --- Aggregate metrics per day ---
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

    if (!dayKey || !billNo) return; // skip invalid rows

    if (!dailyMetrics[dayKey]) dailyMetrics[dayKey] = {
      picked: {}, packed: {}, shipped: {},
      pendingPick: 0, pendingPack: {}, rowsForDay: [],
      customerPending: {}
    };

    const dm = dailyMetrics[dayKey];
    dm.rowsForDay.push(row);

    if (pickedBy) dm.picked[pickedBy] = (dm.picked[pickedBy] || 0) + 1;
    if (packedBy) dm.packed[packedBy] = (dm.packed[packedBy] || 0) + 1;
    if (shippingBy) dm.shipped[shippingBy] = (dm.shipped[shippingBy] || 0) + 1;

    // Fix pending pick: only if bill is not picked yet
    if (!pickedBy) dm.pendingPick++;

    if (!packedBy && pickedBy) dm.pendingPack[pickedBy] = (dm.pendingPack[pickedBy] || 0) + 1;

    if (invoiceState !== "Shipped") {
      if (!dm.customerPending[customer]) dm.customerPending[customer] = [];
      dm.customerPending[customer].push(billNo);
    }
  });

  // --- Prepare batch write ---
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
      batchValues.push([
        dayKey,
        emp,
        dm.picked[emp] || 0,
        dm.packed[emp] || 0,
        dm.shipped[emp] || 0,
        dm.pendingPick,
        dm.pendingPack[emp] || 0,
        JSON.stringify(dm.customerPending),
        "" // Daily summary placeholder
      ]);
    });

    const endRowIndex = batchValues.length - 1;
    mergeInfo.push({dayKey, startRowIndex, endRowIndex, dm});
  });

  const startWriteRow = dailySheet.getLastRow() + 1;
  dailySheet.getRange(startWriteRow, 1, batchValues.length, headers.length).setValues(batchValues);

  // --- Merge Date, Pending Ship (H), Customer Pending (I) wrap, Daily Summary & add borders ---
  mergeInfo.forEach(info => {
    const startRow = startWriteRow + info.startRowIndex;
    const endRow = startWriteRow + info.endRowIndex;

    // Merge Date (A)
    dailySheet.getRange(startRow, 1, endRow - startRow + 1).merge();
    dailySheet.getRange(startRow, 1).setVerticalAlignment("middle");

    // Merge Pending Ship (H)
    dailySheet.getRange(startRow, 8, endRow - startRow + 1).merge();
    dailySheet.getRange(startRow, 8).setVerticalAlignment("top").setWrap(true);

    // Customer Pending (I) - wrap text
    dailySheet.getRange(startRow, 9, endRow - startRow + 1).merge();
    dailySheet.getRange(startRow, 9).setVerticalAlignment("top").setWrap(true);

    // --- Calculate summary for this day ---
    const totalPicked = Object.values(info.dm.picked).reduce((a,b)=>a+b,0);
    const totalPacked = Object.values(info.dm.packed).reduce((a,b)=>a+b,0);
    const totalShipped = Object.values(info.dm.shipped).reduce((a,b)=>a+b,0);
    const pendingPackSum = Object.values(info.dm.pendingPack).reduce((a,b)=>a+b,0);
    const pendingShipCount = Object.values(info.dm.customerPending).reduce((sum, arr) => sum + arr.length, 0);

    const summaryText = `Total Picked: ${totalPicked}\n` +
                        `Total Packed: ${totalPacked}\n` +
                        `Total Shipped: ${totalShipped}\n` +
                        `Pending Pick: ${info.dm.pendingPick}\n` +
                        `Pending Pack: ${pendingPackSum}\n` +
                        `Pending Ship: ${pendingShipCount}`;

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

  // Clear previous customer metrics table
  dailySheet.getRange(1, metricsStartCol, dailySheet.getMaxRows(), metricsWidth).clearContent();

  // Headers
  dailySheet.getRange(1, metricsStartCol, 1, metricsWidth).setValues([customerMetricsHeaders]).setFontWeight("bold");

  // Build per-day, per-customer metrics
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
      const isPicked = !!r[pickedByCol];
      const isPacked = !!r[packedByCol];
      const isShipped = r[invoiceStateCol] === "Shipped";

      if (!perCustomer[custName]) {
        perCustomer[custName] = { total: 0, picked: 0, packed: 0, shipped: 0 };
      }
      const cs = perCustomer[custName];
      cs.total += 1;
      if (isPicked) cs.picked += 1;
      if (isPacked) cs.packed += 1;
      if (isShipped) cs.shipped += 1;
    });

    Object.keys(perCustomer).sort().forEach(custName => {
      const cs = perCustomer[custName];
      const pending = cs.total - cs.shipped; // pending = not shipped
      customerMetricsRows.push([dayKey, custName, cs.total, cs.picked, cs.packed, cs.shipped, pending]);
    });
  });

  if (customerMetricsRows.length > 0) {
    dailySheet.getRange(2, metricsStartCol, customerMetricsRows.length, metricsWidth).setValues(customerMetricsRows);
  }

  // Optional formatting for readability
  dailySheet.setColumnWidths(metricsStartCol, metricsWidth, 140);
  dailySheet.getRange(1, metricsStartCol, 1, metricsWidth).setBackground("#f1f3f4");
   // Merge same-date blocks in L and outline their rows across L:R
  if (customerMetricsRows.length > 0) {
    const startRow = 2; // data starts at L2
    const lastRow = startRow + customerMetricsRows.length - 1;
    const lCol = 12; // 12 (L)
    const width = 7;     // L:R

    let blockStart = startRow;
    let prevDate = customerMetricsRows[0][0];

    for (let i = 1; i < customerMetricsRows.length; i++) {
      const currDate = customerMetricsRows[i][0];
      if (currDate !== prevDate) {
        const blockEnd = startRow + i - 1;
        dailySheet.getRange(blockStart, lCol, blockEnd - blockStart + 1, 1)
          .merge()
          .setVerticalAlignment("middle");
        dailySheet.getRange(blockStart, lCol, blockEnd - blockStart + 1, width)
          .setBorder(true, true, true, true, false, false);
        blockStart = blockEnd + 1;
        prevDate = currDate;
      }
    }
    // finalize last block
    dailySheet.getRange(blockStart, lCol, lastRow - blockStart + 1, 1)
      .merge()
      .setVerticalAlignment("middle");
    dailySheet.getRange(blockStart, lCol, lastRow - blockStart + 1, width)
      .setBorder(true, true, true, true, false, false);
  }
}