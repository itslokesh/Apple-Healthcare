function getTwilioConfig_() {
    const props = PropertiesService.getScriptProperties();
    const AccountSid = props.getProperty('TWILIO_ACCOUNT_SID');
    const AuthToken = props.getProperty('TWILIO_AUTH_TOKEN');
    const From = props.getProperty('TWILIO_WHATSAPP_FROM'); // e.g., 'whatsapp:+14155238886'
    const RecipientsRaw = props.getProperty('WHATSAPP_RECIPIENTS'); // "+91..., +91..."
    if (!AccountSid || !AuthToken || !From) {
        throw new Error('Missing Twilio configuration. Please set TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN, TWILIO_WHATSAPP_FROM in Script Properties.');
    }
    const recipients = (RecipientsRaw || '').split(',').map(s => s.trim()).filter(Boolean);
    return { AccountSid, AuthToken, From, recipients };
}

function normalizeWhatsAppNumber_(num) {
    if (!num) return null;
    let n = num.toString().trim();
    if (!n.startsWith('whatsapp:')) {
        if (!n.startsWith('+')) {
        n = '+' + n;
        }
        n = 'whatsapp:' + n;
    }
    return n;
}

function sendWhatsAppMessageViaTwilio_(to, body, cfg) {
    const url = 'https://api.twilio.com/2010-04-01/Accounts/' + cfg.AccountSid + '/Messages.json';
    const from = cfg.From.startsWith('whatsapp:') ? cfg.From : ('whatsapp:' + cfg.From);

    // Helper to split long text into fixed-size chunks (except the last)
    function fixedChunks_(text, maxLen) {
        const t = (text || '');
        const chunks = [];
        if (maxLen <= 0) return [t];
        for (let i = 0; i < t.length; i += maxLen) {
            chunks.push(t.substring(i, Math.min(i + maxLen, t.length)));
        }
        return chunks.length ? chunks : [''];
    }

    // Determine maximum body length per message considering header and newline
    // Header format: (i/N)\n -> length is 3 + digits(i) + digits(N) for the header, plus 1 for newline
    // Use the worst-case for header per batch: digits(i) == digits(N)
    const fullText = body || '';
    // Iteratively compute digits for total parts to get a stable body size
    let digits = 1;
    let bodyMax = 0;
    let totalParts = 0;
    while (true) {
        bodyMax = 1024 - (3 + 2 * digits) - 1; // 1024 - header - newline
        if (bodyMax <= 0) {
            // Fallback safety: no room for body (shouldn't happen)
            bodyMax = 900;
        }
        totalParts = Math.ceil(fullText.length / bodyMax) || 1;
        const neededDigits = String(totalParts).length;
        if (neededDigits === digits) break;
        digits = neededDigits;
    }

    // Now slice into fixed-size chunks so all parts are bodyMax except the last
    const baseParts = fixedChunks_(fullText, bodyMax);
    let lastResponse = null;
    for (let i = 0; i < totalParts; i++) {
        const header = `(${i + 1}/${totalParts})`;
        const bodyPart = baseParts[i] || '';
        const payload = {
            To: normalizeWhatsAppNumber_(to),
            From: from,
            Body: header + "\n" + bodyPart
        };
        const options = {
            method: 'post',
            payload: payload,
            muteHttpExceptions: true,
            headers: {
                Authorization: 'Basic ' + Utilities.base64Encode(cfg.AccountSid + ':' + cfg.AuthToken)
            }
        };
        const res = UrlFetchApp.fetch(url, options);
        const code = res.getResponseCode();
        if (code < 200 || code >= 300) {
            throw new Error('Twilio send failed (' + code + '): ' + res.getContentText());
        }
        lastResponse = JSON.parse(res.getContentText());
    }
    return lastResponse;
}

function sendWhatsAppSummaryViaTwilio(summaryText) {
    const cfg = getTwilioConfig_();
    let sent = 0;
    if (cfg.recipients.length === 0) {
        console.log('No WHATSAPP_RECIPIENTS configured. Skipping send.');
        return sent;
    }
    cfg.recipients.forEach(rec => {
        try {
        sendWhatsAppMessageViaTwilio_(rec, summaryText, cfg);
        sent++;
        } catch (e) {
        console.log('Failed to send to ' + rec + ': ' + e.message);
        }
    });
    return sent;
}

// Function to generate WhatsApp summary text
function generateWhatsAppSummary(dailyMetrics, escalationRows, timeString, summaryType) {
  const currentDate = new Date().toLocaleDateString('en-IN', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric'
  });
  
  let summary = `*${summaryType.toUpperCase()} REPORT*\n`;
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
    .sort((a, b) => (b[1].picked + b[1].packed + b[1].shipped) - (a[1].picked + a[1].packed + a[1].shipped));
  
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
  
  // All customers by order volume
  const sortedCustomers = Object.entries(customerStats).sort((a, b) => b[1].total - a[1].total);
  
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