function generateMonthlyReportBeautified() {
  // Config (optionally set SPREADSHEET_ID or DASHBOARD_SHEET_GID)
  const SPREADSHEET_ID = ''; // e.g., '1AbC...'; leave empty to use active
  const DASHBOARD_SHEET_NAME = 'New Dashboard1';
  const DASHBOARD_SHEET_GID = 569724509; // optional fallback

  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  let dashboardSheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!dashboardSheet && DASHBOARD_SHEET_GID) {
    dashboardSheet = ss.getSheets().find(sh => sh.getSheetId() === Number(DASHBOARD_SHEET_GID));
  }
  if (!dashboardSheet) {
    throw new Error(`Sheet not found: name="${DASHBOARD_SHEET_NAME}" gid="${DASHBOARD_SHEET_GID || 'n/a'}"`);
  }

  // --- CHUNK SETTINGS ---
  const START_ROW = 2;
  const END_ROW = 7210;

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

  // --- Monthly Report Sheet ---
  let monthlySheet = ss.getSheetByName("Monthly Report 2025");
  if (!monthlySheet) monthlySheet = ss.insertSheet("Monthly Report 2025");

  const headers = ["Month", "Employee", "Picked Count", "Packed Count", "Shipped Count", "Pending Pick", "Pending Pack", "Customer Pending", "Monthly Summary"];

  const rows = data.slice(START_ROW - 1, Math.min(END_ROW, data.length));
  const monthlyMetrics = {}; // monthKey -> metrics

  // --- Aggregate metrics per month ---
  rows.forEach(row => {
    const month = row[monthCol];
    const monthKey = month !== "" ? String(month) : null;
    const pickedBy = row[pickedByCol];
    const packedBy = row[packedByCol];
    const shippingBy = row[shippingByCol];
    const invoiceState = row[invoiceStateCol];
    const customer = row[customerCol];
    const billNo = row[billNoCol];

    if (!monthKey || !billNo) return;

    if (!monthlyMetrics[monthKey]) monthlyMetrics[monthKey] = {
      picked: {}, packed: {}, shipped: {},
      pendingPick: 0, pendingPack: {}, rowsForMonth: [],
      customerPending: {}
    };

    const mm = monthlyMetrics[monthKey];
    mm.rowsForMonth.push(row);

    if (pickedBy) mm.picked[pickedBy] = (mm.picked[pickedBy] || 0) + 1;
    if (packedBy) mm.packed[packedBy] = (mm.packed[packedBy] || 0) + 1;
    if (shippingBy) mm.shipped[shippingBy] = (mm.shipped[shippingBy] || 0) + 1;

    if (!pickedBy) mm.pendingPick++;
    if (!packedBy && pickedBy) mm.pendingPack[pickedBy] = (mm.pendingPack[pickedBy] || 0) + 1;

    if (invoiceState !== "Shipped") {
      if (!mm.customerPending[customer]) mm.customerPending[customer] = [];
      mm.customerPending[customer].push(billNo);
    }
  });

  // --- Prepare batch write ---
  const batchValues = [];
  const mergeInfo = [];
  Object.keys(monthlyMetrics).sort((a,b)=>Number(a)-Number(b)).forEach(monthKey => {
    const mm = monthlyMetrics[monthKey];
    const employees = new Set([...Object.keys(mm.picked), ...Object.keys(mm.packed), ...Object.keys(mm.shipped), ...Object.keys(mm.pendingPack)]);
    const startRowIndex = batchValues.length;

    employees.forEach(emp => {
      batchValues.push([
        monthKey,
        emp,
        mm.picked[emp] || 0,
        mm.packed[emp] || 0,
        mm.shipped[emp] || 0,
        mm.pendingPick,
        mm.pendingPack[emp] || 0,
        JSON.stringify(mm.customerPending),
        "" // Monthly summary placeholder
      ]);
    });

    const endRowIndex = batchValues.length - 1;
    mergeInfo.push({monthKey, startRowIndex, endRowIndex, mm});
  });

  // --- Write main table at A20 ---
  const TABLE_START_ROW = 20;

  // Clear old table (A:I)
  monthlySheet.getRange(1, 1, monthlySheet.getMaxRows(), headers.length).clearContent();

  // Headers at A20:I20
  monthlySheet.getRange(TABLE_START_ROW, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#f1f3f4");

  // Data starting at A21
  const startWriteRow = TABLE_START_ROW + 1;
  if (batchValues.length > 0) {
    monthlySheet.getRange(startWriteRow, 1, batchValues.length, headers.length).setValues(batchValues);
  }

  // --- Merge Month (A), Pending Ship (H), Customer Pending (I) wrap, Monthly Summary & add borders ---
  mergeInfo.forEach(info => {
    const startRow = startWriteRow + info.startRowIndex;
    const endRow = startWriteRow + info.endRowIndex;

    // Merge Month (A)
    monthlySheet.getRange(startRow, 1, endRow - startRow + 1).merge();
    monthlySheet.getRange(startRow, 1).setVerticalAlignment("middle");

    // Merge Pending Ship (H)
    monthlySheet.getRange(startRow, 8, endRow - startRow + 1).merge();
    monthlySheet.getRange(startRow, 8).setVerticalAlignment("top").setWrap(true);

    // Customer Pending (I) - wrap text
    monthlySheet.getRange(startRow, 9, endRow - startRow + 1).merge();
    monthlySheet.getRange(startRow, 9).setVerticalAlignment("top").setWrap(true);

    // --- Monthly summary for this month ---
    const totalPicked = Object.values(info.mm.picked).reduce((a,b)=>a+b,0);
    const totalPacked = Object.values(info.mm.packed).reduce((a,b)=>a+b,0);
    const totalShipped = Object.values(info.mm.shipped).reduce((a,b)=>a+b,0);
    const pendingPackSum = Object.values(info.mm.pendingPack).reduce((a,b)=>a+b,0);
    const pendingShipCount = Object.values(info.mm.customerPending).reduce((sum, arr) => sum + arr.length, 0);

    const summaryText = `Total Picked: ${totalPicked}\n` +
                        `Total Packed: ${totalPacked}\n` +
                        `Total Shipped: ${totalShipped}\n` +
                        `Pending Pick: ${info.mm.pendingPick}\n` +
                        `Pending Pack: ${pendingPackSum}\n` +
                        `Pending Ship: ${pendingShipCount}`;

    monthlySheet.getRange(startRow, 9).setValue(summaryText);

    // Outline border for all columns of the month (no inner gridlines)
    monthlySheet.getRange(startRow, 1, endRow - startRow + 1, headers.length)
      .setBorder(true, true, true, true, false, false);
  });

  // Column widths for merged/wrapped columns
  monthlySheet.setColumnWidth(8, 420);  // H - Pending Ship
  monthlySheet.setColumnWidth(9, 520);  // I - Customer Pending and Summary

  // --- Monthly Customer Metrics (L:R) starting at row 20 ---
  const CM_HEADERS = ["Month", "Customer", "Total Orders", "Picked", "Packed", "Shipped", "Pending"];
  const CM_START_COL = 12; // L
  const CM_WIDTH = CM_HEADERS.length;
  const CM_START_ROW = 20; // L20 header, L21 data

  // Clear previous metrics (L:R)
  monthlySheet.getRange(1, CM_START_COL, monthlySheet.getMaxRows(), CM_WIDTH).clearContent();

  // Headers at L20:R20
  monthlySheet.getRange(CM_START_ROW, CM_START_COL, 1, CM_WIDTH)
    .setValues([CM_HEADERS])
    .setFontWeight("bold")
    .setBackground("#f1f3f4");

  // Build per-month, per-customer metrics
  const customerMetricsRows = [];
  Object.keys(monthlyMetrics).sort((a,b)=>Number(a)-Number(b)).forEach(monthKey => {
    const mm = monthlyMetrics[monthKey];
    const perCustomer = {};

    mm.rowsForMonth.forEach(r => {
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
      customerMetricsRows.push([monthKey, custName, cs.total, cs.picked, cs.packed, cs.shipped, pending]);
    });
  });

  // Data at L21
  if (customerMetricsRows.length > 0) {
    monthlySheet.getRange(CM_START_ROW + 1, CM_START_COL, customerMetricsRows.length, CM_WIDTH).setValues(customerMetricsRows);

    // Merge same-month blocks in L and outline across L:R
    const startRow = CM_START_ROW + 1; // L21
    const lastRow = startRow + customerMetricsRows.length - 1;
    const lCol = CM_START_COL; // 12
    const width = CM_WIDTH;    // L:R

    let blockStart = startRow;
    let prevMonth = customerMetricsRows[0][0];

    for (let i = 1; i < customerMetricsRows.length; i++) {
      const currMonth = customerMetricsRows[i][0];
      if (currMonth !== prevMonth) {
        const blockEnd = startRow + i - 1;
        monthlySheet.getRange(blockStart, lCol, blockEnd - blockStart + 1, 1).merge().setVerticalAlignment("middle");
        monthlySheet.getRange(blockStart, lCol, blockEnd - blockStart + 1, width).setBorder(true, true, true, true, false, false);
        blockStart = blockEnd + 1;
        prevMonth = currMonth;
      }
    }
    monthlySheet.getRange(blockStart, lCol, lastRow - blockStart + 1, 1).merge().setVerticalAlignment("middle");
    monthlySheet.getRange(blockStart, lCol, lastRow - blockStart + 1, width).setBorder(true, true, true, true, false, false);
  }

  // Optional formatting for readability
  monthlySheet.setColumnWidths(CM_START_COL, CM_WIDTH, 140);

  // --- Monthly Summary chart source table (T1 onward) ---
  const SUMMARY_START_COL = 20; // T
  const SUMMARY_START_ROW = 1;
  const chartHeaders = ["Month", "Picked", "Packed", "Shipped", "Pending Pick", "Pending Pack", "Pending Ship"];

  const monthsSorted = Object.keys(monthlyMetrics).map(Number).sort((a,b)=>a-b);
  const monthNames = ["","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

  const chartRows = monthsSorted.map(m => {
    const mm = monthlyMetrics[String(m)];
    const totalPicked = Object.values(mm.picked).reduce((a,b)=>a+b,0);
    const totalPacked = Object.values(mm.packed).reduce((a,b)=>a+b,0);
    const totalShipped = Object.values(mm.shipped).reduce((a,b)=>a+b,0);
    const pendingPackSum = Object.values(mm.pendingPack).reduce((a,b)=>a+b,0);
    const pendingShipCount = Object.values(mm.customerPending).reduce((sum, arr) => sum + arr.length, 0);
    const monthLabel = monthNames[m] || String(m);
    return [monthLabel, totalPicked, totalPacked, totalShipped, mm.pendingPick, pendingPackSum, pendingShipCount];
  });

  // Clear old helper table (T:Z)
  monthlySheet.getRange(SUMMARY_START_ROW, SUMMARY_START_COL, monthlySheet.getMaxRows(), chartHeaders.length).clearContent();
  if (chartRows.length > 0) {
    monthlySheet.getRange(SUMMARY_START_ROW, SUMMARY_START_COL, 1, chartHeaders.length)
      .setValues([chartHeaders])
      .setFontWeight("bold")
      .setBackground("#f1f3f4");
    monthlySheet.getRange(SUMMARY_START_ROW + 1, SUMMARY_START_COL, chartRows.length, chartHeaders.length)
      .setValues(chartRows);
  }

  // Remove existing charts to avoid duplicates
  monthlySheet.getCharts().forEach(ch => monthlySheet.removeChart(ch));

  // Build and insert chart above the table (at A1)
  if (chartRows.length > 0) {
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

    // Domain (Month)
    builder.addRange(monthlySheet.getRange(SUMMARY_START_ROW, SUMMARY_START_COL, chartRows.length + 1, 1));

    // Series: Picked..Pending Ship
    for (let c = 1; c < chartHeaders.length; c++) {
      builder.addRange(monthlySheet.getRange(SUMMARY_START_ROW, SUMMARY_START_COL + c, chartRows.length + 1, 1));
    }

    monthlySheet.insertChart(builder.build());
  }
}