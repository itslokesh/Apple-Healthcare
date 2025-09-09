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

    // Helper to split long text into <= 1024 char chunks, preferring line boundaries
    function chunkText_(text, maxLen) {
        const chunks = [];
        const lines = (text || '').split('\n');
        let buf = '';
        lines.forEach((line, idx) => {
            const candidate = (buf.length ? buf + '\n' : '') + line;
            if (candidate.length <= maxLen) {
                buf = candidate;
            } else {
                if (buf.length) {
                    chunks.push(buf);
                    buf = '';
                }
                // If a single line itself is longer than maxLen, hard-split it
                let start = 0;
                while (start < line.length) {
                    const end = Math.min(start + maxLen, line.length);
                    chunks.push(line.substring(start, end));
                    start = end;
                }
            }
            // Push last buffer at the end
            if (idx === lines.length - 1 && buf.length) {
                chunks.push(buf);
            }
        });
        return chunks.length ? chunks : [''];
    }

    // Split content to reserve 5 chars for part indicators like (1/9)
    // Assumption per requirements: total parts will not exceed 9
    const baseParts = chunkText_(body, 1019);
    const totalParts = baseParts.length;
    let lastResponse = null;
    for (let i = 0; i < totalParts; i++) {
        const header = `(${i + 1}/${totalParts})`; // part indicator without trailing space
        // Ensure final payload length <= 1024 (reserve 1 char for newline)
        const maxBodyLen = 1024 - header.length - 1;
        const bodyPart = baseParts[i].length > maxBodyLen ? baseParts[i].substring(0, maxBodyLen) : baseParts[i];
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