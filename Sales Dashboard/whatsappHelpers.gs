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
    const payload = {
        To: normalizeWhatsAppNumber_(to),
        From: cfg.From.startsWith('whatsapp:') ? cfg.From : ('whatsapp:' + cfg.From),
        Body: body
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
    return JSON.parse(res.getContentText());
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