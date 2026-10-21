/**
 * OKV TECHNOLOGY - HMS licensing, submissions, customers, messaging and activation codes.
 * Runs entirely on Google Sheets + Google Apps Script. No other server.
 * Setup: see SETUP_GUIDE.md. Run setup() once, then deploy as a Web app.
 */
const CFG = {
  BRAND: 'OKV Technology Consults',
  PRODUCT: 'OKV Hospital Management System',
  OWNER_EMAIL: 'technologyokv@gmail.com',
  WEBSITE: 'www.okvtechnology.com',
  WHATSAPP: '2348104141138',
  PRICE: 120000,
  TRIAL_DAYS: 12,
  TZ: 'Africa/Lagos'
};
const SH = { SUB: 'Submissions', CUS: 'Customers', CODES: 'Codes', MSG: 'Messages', ALERTS: 'Alerts', PAY: 'Payments' };
const HEAD = {
  Submissions: ['ID', 'Time', 'Type', 'Name', 'Organization', 'Email', 'Phone', 'Location', 'Message', 'InstallID', 'PayRef', 'Amount', 'PaidDate', 'ReceiptURL', 'Status', 'Notes'],
  Customers: ['ID', 'Created', 'Name', 'Organization', 'Email', 'Phone', 'Location', 'Source', 'Status', 'InstallID', 'TrialStart', 'ActivationCode', 'CodeSentAt', 'Notes'],
  Codes: ['Serial', 'Code', 'Status', 'AssignedTo', 'SentAt', 'SentBy'],
  Messages: ['Time', 'Channel', 'To', 'Subject', 'Body', 'Status'],
  Alerts: ['ID', 'Time', 'Type', 'Text', 'SubID', 'Read'],
  Payments: ['Reference', 'Time', 'Email', 'Name', 'Organization', 'Phone', 'Amount', 'Currency', 'Channel', 'Status', 'Code', 'InstallID', 'Note']
};
const PUBLIC_GET_ACTIONS = ['trial', 'contact', 'config', 'paystack_init', 'install_request'];

/* ------------------------------------------------------------ basics */
function props_() { return PropertiesService.getScriptProperties(); }
function ss_() { const id = props_().getProperty('SHEET_ID'); return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(n) { const s = ss_().getSheetByName(n); if (!s) throw new Error('Run setup() first'); return s; }
function ts_() { return Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd HH:mm'); }
function clean_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max || 200); }
function validEmail_(e) { return /^[^@\s]{1,64}@[^@\s]+\.[^@\s]{2,}$/.test(e); }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function fail_(m) { return { ok: false, error: m }; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function money_(n) { return 'NGN ' + Number(n || 0).toLocaleString('en-US'); }
function rate_(key, max, secs) {
  const c = CacheService.getScriptCache(), k = 'rl_' + key.slice(0, 200), n = Number(c.get(k) || 0);
  if (n >= max) return false;
  c.put(k, String(n + 1), secs); return true;
}
function nextId_(prefix) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { const p = props_(), k = 'CNT_' + prefix, n = Number(p.getProperty(k) || 0) + 1; p.setProperty(k, String(n)); return prefix + '-' + ('000000' + n).slice(-6); }
  finally { lock.releaseLock(); }
}
function readAll_(name) {
  const v = sheet_(name).getDataRange().getValues(), h = v[0], out = [];
  for (let i = 1; i < v.length; i++) { const o = { _row: i + 1 }; h.forEach((k, j) => o[k] = v[i][j] instanceof Date ? Utilities.formatDate(v[i][j], CFG.TZ, 'yyyy-MM-dd HH:mm') : v[i][j]); out.push(o); }
  return out;
}
function setCell_(name, row, col, val) { sheet_(name).getRange(row, HEAD[name].indexOf(col) + 1).setValue(val); }

/* ------------------------------------------------------------ activation codes (must match okv_hms.py) */
function hmacHex_(msg, key) {
  return Utilities.computeHmacSha256Signature(msg, key).map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function codeFor_(serial, secret) {
  let s = String(serial); while (s.length < 4) s = '0' + s;                     // 4 digits up to 9999, then 5 or 6 digits (the HMS accepts 4 to 6)
  const mac = hmacHex_('HMS-' + s, secret).toUpperCase().slice(0, 10);
  return 'HMS-' + s + '-' + mac.slice(0, 5) + '-' + mac.slice(5);
}
function generateCodes_(n) {
  const sh = sheet_(SH.CODES), secret = props_().getProperty('LICENSE_SECRET'), last = sh.getLastRow();
  const serial = last > 1 ? Number(sh.getRange(last, 1).getValue()) : 0;
  if (serial + n > 999999) throw new Error('Serial numbers go up to 999999 only');
  const rows = []; for (let i = 1; i <= n; i++) rows.push([serial + i, codeFor_(serial + i, secret), 'Available', '', '', '']);
  sh.getRange(last + 1, 1, rows.length, 6).setValues(rows);
  return rows.length;
}

/** RUN THIS ONCE from the editor (Run > setup). */
function setup() {
  const p = props_();
  // Works from a standalone script (no sheet needed): creates the spreadsheet once and remembers it.
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    const saved = p.getProperty('SHEET_ID');
    ss = saved ? SpreadsheetApp.openById(saved) : SpreadsheetApp.create('OKV HMS Licensing');
  }
  p.setProperty('SHEET_ID', ss.getId());
  Logger.log('Your data spreadsheet (open it to see the sheets): ' + ss.getUrl());
  Object.keys(HEAD).forEach(n => {
    const s = ss.getSheetByName(n) || ss.insertSheet(n);
    if (s.getLastRow() === 0) {
      s.getRange(1, 1, 1, HEAD[n].length).setValues([HEAD[n]]).setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
      s.setFrozenRows(1);
    }
    if (n !== 'Codes') s.getRange(2, 1, s.getMaxRows() - 1, HEAD[n].length).setNumberFormat('@'); // plain text: user input can never run as a formula
  });
  const def = ss.getSheetByName('Sheet1'); if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
  if (!p.getProperty('LICENSE_SECRET')) p.setProperty('LICENSE_SECRET', Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''));
  if (!p.getProperty('ADMIN_PASSWORD')) {
    const pw = Utilities.getUuid().replace(/-/g, '').slice(0, 10); p.setProperty('ADMIN_PASSWORD', pw);
    Logger.log('YOUR ADMIN DASHBOARD PASSWORD (change it in Project Settings > Script properties): ' + pw);
  }
  ['DOWNLOAD_URL', 'DOWNLOAD_URL_MAC_ARM', 'DOWNLOAD_URL_MAC_INTEL', 'PRICING_URL'].forEach(k => { if (p.getProperty(k) == null) p.setProperty(k, ''); });
  const defaults = { PRICE: String(CFG.PRICE), TRIALS_ENABLED: '1', PURCHASES_ENABLED: '1', MANUAL_ENABLED: '1', PAYSTACK_ENABLED: '0', PAYSTACK_AUTO_CODE: '1', ALERT_EMAIL_ENABLED: '1',
    PAY_BANK: 'YOUR BANK NAME', PAY_ACCOUNT_NAME: CFG.BRAND, PAY_ACCOUNT_NO: '0000000000', PAY_INSTRUCTIONS: 'Pay by bank transfer, then upload your receipt.',
    INSTALL_TEXT: 'We install and set up the system for you, add your staff, wards and prices, connect your phones and tablets, and train your team. Contact us for a quote.' };
  Object.keys(defaults).forEach(k => { if (p.getProperty(k) == null) p.setProperty(k, defaults[k]); });
  if (sheet_(SH.CODES).getLastRow() <= 1) generateCodes_(1200);
  Logger.log('Setup finished. LICENSE_SECRET (put the same value in okv_hms.py): ' + p.getProperty('LICENSE_SECRET'));
}

/* ------------------------------------------------------------ web app entry points */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.page === 'admin') return ContentService.createTextOutput('The admin dashboard is on your website (admin.html). Open it there.');
  if (p.action === 'pay_return') return payReturnPage_(p);
  if (p.action) {
    let out;
    try { out = PUBLIC_GET_ACTIONS.indexOf(p.action) >= 0 ? handlePublic_(p) : fail_('Unsupported request'); }
    catch (err) { console.error(err); out = fail_('Server error. Please try again.'); }
    const cb = /^[A-Za-z_$][\w$]{0,60}$/.test(p.callback || '') ? p.callback : null;
    return cb ? ContentService.createTextOutput(cb + '(' + JSON.stringify(out) + ');').setMimeType(ContentService.MimeType.JAVASCRIPT) : json_(out);
  }
  return ContentService.createTextOutput('OKV licensing service is running.');
}
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body && body.event && body.data && body.data.reference && !body.action) {   // Paystack webhook: ignored unless Paystack itself confirms the payment
      if (body.event === 'charge.success') paystackVerify_(body.data.reference);
      return json_({ ok: true });
    }
    if (body && body.action === 'admin_call') return json_(adminDispatch_(body));
    return json_(handlePublic_(body));
  }
  catch (err) { console.error(err); return json_(fail_('Could not process the request.')); }
}

/* ------------------------------------------------------------ admin dashboard API (used by admin.html on GitHub Pages)
   Only the functions listed here can be called from the web. Each one still checks the admin session token itself. */
const ADMIN_FNS_ = ['adminLogin', 'adminLogout', 'adminAlerts', 'adminApprove', 'adminApproveMany', 'adminCodes', 'adminConfig',
  'adminCustomers', 'adminDeleteCustomer', 'adminGenerateCodes', 'adminIssueCode', 'adminMarkAlertsRead', 'adminMessages',
  'adminPayments', 'adminRecordPayment', 'adminResetCode', 'adminSaveConfig', 'adminSaveCustomer', 'adminSaveSettings',
  'adminSendCode', 'adminSendMessage', 'adminSettings', 'adminSubmissions', 'adminSummary', 'adminUpdateSubmission', 'adminTestIntegration'];
const ADMIN_MAP_ = {
  adminLogin: adminLogin,
  adminLogout: adminLogout,
  adminAlerts: adminAlerts,
  adminApprove: adminApprove,
  adminApproveMany: adminApproveMany,
  adminCodes: adminCodes,
  adminConfig: adminConfig,
  adminCustomers: adminCustomers,
  adminDeleteCustomer: adminDeleteCustomer,
  adminGenerateCodes: adminGenerateCodes,
  adminIssueCode: adminIssueCode,
  adminMarkAlertsRead: adminMarkAlertsRead,
  adminMessages: adminMessages,
  adminPayments: adminPayments,
  adminRecordPayment: adminRecordPayment,
  adminResetCode: adminResetCode,
  adminSaveConfig: adminSaveConfig,
  adminSaveCustomer: adminSaveCustomer,
  adminSaveSettings: adminSaveSettings,
  adminSendCode: adminSendCode,
  adminSendMessage: adminSendMessage,
  adminSettings: adminSettings,
  adminSubmissions: adminSubmissions,
  adminSummary: adminSummary,
  adminUpdateSubmission: adminUpdateSubmission,
  adminTestIntegration: adminTestIntegration
};
function adminDispatch_(d) {
  const fn = String(d.fn || '');
  if (ADMIN_FNS_.indexOf(fn) < 0) return fail_('Unknown admin request.');
  const args = Array.isArray(d.args) ? d.args : [];
  try {
    const f = ADMIN_MAP_[fn];
    if (typeof f !== 'function') return fail_('Unknown admin request.');
    return { ok: true, result: f.apply(null, args) };
  } catch (err) {
    return fail_(String((err && err.message) || err).replace(/^Error: /, ''));
  }
}

/* ------------------------------------------------------------ public submissions */
function handlePublic_(d) {
  if (clean_(d.hp, 50)) return { ok: true };                                  // honeypot: bots fill hidden fields
  const action = clean_(d.action, 30);
  if (action === 'config') return getConfig_();
  const f = {
    name: clean_(d.name, 100), org: clean_(d.hospital || d.org, 120), email: clean_(d.email, 120).toLowerCase(),
    phone: clean_(d.phone, 30), city: clean_(d.city, 80), message: clean_(d.message, 1000), installId: clean_(d.install_id, 20)
  };
  if (!f.name) return fail_('Please enter your name.');
  if (!validEmail_(f.email)) return fail_('Please enter a valid email address.');
  if (!rate_('em_' + f.email, 8, 3600)) return fail_('Too many submissions from this email. Please try again later.');
  switch (action) {
    case 'trial': {
      if (props_().getProperty('TRIALS_ENABLED') === '0') return fail_('Trial downloads are paused right now. Please message us on WhatsApp +' + CFG.WHATSAPP + '.');
      if (!f.org && f.hospital) f.org = f.hospital;          // the website form sends the name as 'hospital'
      if (!f.org || !f.phone) return fail_('Please enter your hospital/business name and phone number.');
      const dl = downloads_();
      const linksReady = !!(dl.windows || dl.macArm || dl.macIntel);   // false until you add the download links in Settings
      f.message = (f.message ? f.message + ' | ' : '') + 'Type: ' + clean_(d.kind, 40);
      const id = submit_('Trial Download', f, {});
      const links = [['Windows 7, 8, 10 and 11 (32-bit and 64-bit)', dl.windows], ['Mac with Apple chip (M1, M2, M3, M4)', dl.macArm], ['Mac with Intel chip', dl.macIntel]].filter(x => x[1]);
      safeMail_(f.email, 'Your OKV HMS free trial download', wrap_('Your ' + CFG.TRIAL_DAYS + '-day free trial',
        '<p>Hello ' + esc_(f.name) + ',</p><p>Thank you for trying the ' + esc_(CFG.PRODUCT) + '.</p>' +
        links.map(x => '<p><a href="' + esc_(x[1]) + '" style="background:#0f766e;color:#fff;padding:12px 20px;border-radius:8px;text-decoration:none;font-weight:bold;display:inline-block">Download for ' + esc_(x[0]) + '</a></p>').join('') +
        '<ol><li><b>Windows:</b> extract the ZIP and run <b>OKV_HMS_Setup.exe</b>.<br><b>Mac:</b> open the DMG and drag OKV HMS into Applications.</li><li>Open OKV HMS and sign in with <b>admin / admin123</b>.</li><li>Nothing else to install.</li></ol>' +
        '<p>Your free trial lasts <b>' + CFG.TRIAL_DAYS + ' days</b>. To keep using the system afterwards, click <b>Buy licence</b> inside the Admin dashboard or visit our pricing page (one-time ' + money_(Number(props_().getProperty('PRICE') || CFG.PRICE)) + ').</p><p>Questions? Reply to this email or WhatsApp +' + CFG.WHATSAPP + '.</p>'),
        links.map(x => x[0] + ': ' + x[1]).join('\n') + '\nTrial: ' + CFG.TRIAL_DAYS + ' days. First login admin / admin123.');
      return { ok: true, id: id, pending: !linksReady, downloadUrl: dl.windows, macArmUrl: dl.macArm, macIntelUrl: dl.macIntel };
    }
    case 'trial_register': {
      const id = submit_('Trial Registered', f, {});
      safeMail_(f.email, 'We received your trial registration', ackHtml_(f, 'Thank you for registering your free trial. Our team will contact you if you need help.'), 'Thank you for registering.');
      return { ok: true, id: id };
    }
    case 'contact': case 'support': {
      if (!f.message) return fail_('Please type your message.');
      const id = submit_(action === 'support' ? 'Support' : 'Enquiry', f, {});
      safeMail_(f.email, 'We received your message', ackHtml_(f, 'We received your message and will reply shortly.'), 'We received your message.');
      return { ok: true, id: id };
    }
    case 'paystack_init': return paystackInit_(f);
    case 'install_request': {
      if (!f.org) return fail_('Please enter your hospital or business name.');
      const id = submit_('Installation Request', f, {});
      safeMail_(f.email, 'We received your installation request', ackHtml_(f, 'Thank you. Our installation team will contact you shortly to agree a date and quote.'), 'We received your installation request.');
      return { ok: true, id: id };
    }
    case 'activation_request': case 'purchase': {
      if (props_().getProperty('MANUAL_ENABLED') === '0') return fail_('Bank transfer payments are not open right now. Please pay online with Paystack.');
      if (props_().getProperty('PURCHASES_ENABLED') === '0') return fail_('Purchases are paused right now. Please message us on WhatsApp +' + CFG.WHATSAPP + '.');
      const amount = Number(d.amount), ref = clean_(d.payment_ref, 80), paidDate = clean_(d.paid_date, 10);
      if (!ref) return fail_('Please enter the payment reference.');
      if (!(amount > 0)) return fail_('Please enter the amount paid.');
      if (!d.receipt) return fail_('Please attach your payment receipt.');
      const sid = nextId_('S');
      const receiptUrl = saveReceipt_(String(d.receipt), clean_(d.receipt_mime, 40), clean_(d.receipt_name, 80), sid);
      const id = submit_(action === 'purchase' ? 'Purchase Request' : 'Activation Request', f, { payRef: ref, amount: String(amount), paidDate: paidDate, receiptUrl: receiptUrl }, sid);
      safeMail_(f.email, 'We received your activation request', wrap_('Activation request received',
        '<p>Hello ' + esc_(f.name) + ',</p><p>We received your payment details (reference <b>' + esc_(ref) + '</b>, ' + money_(amount) + ').</p>' +
        '<p>Your <b>activation code will be sent to this email address within 24 hours</b>, usually much sooner. Keep this email.</p>' +
        (f.installId ? '<p>Installation ID: <b>' + esc_(f.installId) + '</b></p>' : '') + '<p>Need help? WhatsApp +' + CFG.WHATSAPP + '.</p>'),
        'We received your activation request. Your code will be emailed within 24 hours.');
      return { ok: true, id: id };
    }
    default: return fail_('Unknown request.');
  }
}
function downloads_() { const p = props_(); return { windows: p.getProperty('DOWNLOAD_URL') || '', macArm: p.getProperty('DOWNLOAD_URL_MAC_ARM') || '', macIntel: p.getProperty('DOWNLOAD_URL_MAC_INTEL') || '' }; }
function ackHtml_(f, line) { return wrap_('Thank you', '<p>Hello ' + esc_(f.name) + ',</p><p>' + esc_(line) + '</p><p>WhatsApp +' + CFG.WHATSAPP + ' | ' + esc_(CFG.OWNER_EMAIL) + '</p>'); }

function submit_(type, f, x, presetId) {
  const id = presetId || nextId_('S');
  sheet_(SH.SUB).appendRow([id, ts_(), type, f.name, f.org, f.email, f.phone, f.city, f.message, f.installId, x.payRef || '', x.amount || '', x.paidDate || '', x.receiptUrl || '', 'New', '']);
  upsertCustomer_(f, type);
  alert_(type, id, f, x);
  return id;
}
function upsertCustomer_(f, type) {
  const sh = sheet_(SH.CUS), all = readAll_(SH.CUS), cur = all.filter(c => String(c.Email).toLowerCase() === f.email)[0];
  const pending = /Request|Paystack/.test(type) && type !== 'Installation Request';
  if (!cur) {
    const status = pending ? 'Pending activation' : (type === 'Trial Download' || type === 'Trial Registered') ? 'Trial' : 'Lead';
    sh.appendRow([nextId_('C'), ts_(), f.name, f.org, f.email, f.phone, f.city, type, status, f.installId, (status === 'Trial' ? ts_().slice(0, 10) : ''), '', '', '']);
    return;
  }
  if (f.name) setCell_(SH.CUS, cur._row, 'Name', f.name);
  if (f.org) setCell_(SH.CUS, cur._row, 'Organization', f.org);
  if (f.phone) setCell_(SH.CUS, cur._row, 'Phone', f.phone);
  if (f.city) setCell_(SH.CUS, cur._row, 'Location', f.city);
  if (f.installId) setCell_(SH.CUS, cur._row, 'InstallID', f.installId);
  if (pending && cur.Status !== 'Licensed') setCell_(SH.CUS, cur._row, 'Status', 'Pending activation');
  else if ((type === 'Trial Download' || type === 'Trial Registered') && (cur.Status === 'Lead' || !cur.Status)) { setCell_(SH.CUS, cur._row, 'Status', 'Trial'); setCell_(SH.CUS, cur._row, 'TrialStart', ts_().slice(0, 10)); }
}
function alert_(type, subId, f, x) {
  const text = type + ': ' + f.name + (f.org ? ' (' + f.org + ')' : '') + ' <' + f.email + '>' + (x && x.amount ? ' paid ' + money_(x.amount) : '');
  try { sheet_(SH.ALERTS).appendRow([nextId_('A'), ts_(), type, text, subId, 'No']); } catch (e) { console.error(e); }
  const rows = [['Type', type], ['Name', f.name], ['Organization', f.org], ['Email', f.email], ['Phone', f.phone], ['Location', f.city], ['Message', f.message], ['Install ID', f.installId], ['Payment ref', x && x.payRef], ['Amount', x && x.amount ? money_(x.amount) : ''], ['Receipt', x && x.receiptUrl]]
    .filter(r => r[1]).map(r => '<tr><td style="padding:4px 10px;color:#64748b">' + r[0] + '</td><td style="padding:4px 10px"><b>' + esc_(r[1]) + '</b></td></tr>').join('');
  const pr = props_();
  if (pr.getProperty('ALERT_EMAIL_ENABLED') !== '0') safeMail_(pr.getProperty('OWNER_EMAIL') || CFG.OWNER_EMAIL, '[OKV HMS] New: ' + type + ' - ' + f.name, wrap_('New submission alert', '<table>' + rows + '</table><p>Open the admin dashboard to respond.</p>'), text);
  notifyExtra_('[OKV HMS] ' + type, text, { event: 'submission', id: subId, type: type, name: f.name, organization: f.org, email: f.email, phone: f.phone, message: f.message, amount: x && x.amount || '', installId: f.installId });
}
function notifyExtra_(title, text, payload) {
  const p = props_();
  if (p.getProperty('TG_ENABLED') === '1' && p.getProperty('TG_TOKEN') && p.getProperty('TG_CHAT')) {
    try { UrlFetchApp.fetch('https://api.telegram.org/bot' + p.getProperty('TG_TOKEN') + '/sendMessage', { method: 'post', muteHttpExceptions: true, payload: { chat_id: p.getProperty('TG_CHAT'), text: title + '\n' + text } }); } catch (e) { console.error(e); }
  }
  const url = p.getProperty('WEBHOOK_URL') || '';
  if (p.getProperty('WEBHOOK_ENABLED') === '1' && /^https:\/\//.test(url)) {
    try { UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { 'X-OKV-Secret': p.getProperty('WEBHOOK_SECRET') || '' }, payload: JSON.stringify(payload) }); } catch (e) { console.error(e); }
  }
}
function saveReceipt_(b64, mime, name, id) {
  const bytes = Utilities.base64Decode(b64);
  if (bytes.length > 3.3 * 1024 * 1024) throw new Error('Receipt is too large');
  const b0 = bytes[0] & 255, b1 = bytes[1] & 255, b2 = bytes[2] & 255, b3 = bytes[3] & 255;
  const type = (b0 === 0x25 && b1 === 0x50 && b2 === 0x44 && b3 === 0x46) ? 'application/pdf' : (b0 === 0xff && b1 === 0xd8 && b2 === 0xff) ? 'image/jpeg' : (b0 === 0x89 && b1 === 0x50 && b2 === 0x4e && b3 === 0x47) ? 'image/png' : '';
  if (!type) throw new Error('Receipt must be a JPG, PNG or PDF file');
  const safe = name.replace(/[^A-Za-z0-9._-]/g, '_').slice(0, 60) || 'receipt';
  return receiptFolder_().createFile(Utilities.newBlob(bytes, type, id + '_' + safe)).getUrl();
}
function receiptFolder_() {
  const p = props_(), id = p.getProperty('RECEIPT_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { } }
  const f = DriveApp.createFolder('OKV HMS Payment Receipts'); p.setProperty('RECEIPT_FOLDER_ID', f.getId()); return f;
}

/* ------------------------------------------------------------ email */
function wrap_(title, inner) {
  return '<div style="font-family:Arial,sans-serif;max-width:560px;margin:auto;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden"><div style="background:#0f766e;color:#fff;padding:16px 20px;font-size:18px;font-weight:bold">' + esc_(CFG.PRODUCT) + '</div><div style="padding:20px;color:#0f172a;font-size:15px;line-height:1.5"><h2 style="margin:0 0 10px;font-size:18px">' + esc_(title) + '</h2>' + inner + '</div><div style="background:#f1f5f9;padding:12px 20px;font-size:12px;color:#64748b">' + esc_(CFG.BRAND) + ' | ' + esc_(CFG.WEBSITE) + ' | ' + esc_(CFG.OWNER_EMAIL) + ' | WhatsApp +' + CFG.WHATSAPP + '</div></div>';
}
function mail_(to, subject, html, text) { MailApp.sendEmail({ to: to, subject: subject, htmlBody: html, body: text || '', name: CFG.BRAND, replyTo: CFG.OWNER_EMAIL }); }
function safeMail_(to, subject, html, text) { try { mail_(to, subject, html, text); logMsg_('email (auto)', to, subject, text, 'Sent'); } catch (e) { console.error(e); logMsg_('email (auto)', to, subject, text, 'FAILED: ' + e.message); } }
function logMsg_(channel, to, subject, body, status) { try { sheet_(SH.MSG).appendRow([ts_(), channel, to, subject || '', String(body || '').slice(0, 500), status]); } catch (e) { console.error(e); } }

function sendCode_(p, who) {
  const email = clean_(p.email, 120).toLowerCase(), name = clean_(p.name, 100) || 'Customer';
  if (!validEmail_(email)) throw new Error('A valid customer email is needed');
  if (MailApp.getRemainingDailyQuota() < 1) throw new Error('Daily email limit reached. Try again tomorrow.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let row = 0, serial = 0, code = '';
  try {
    const sh = sheet_(SH.CODES), vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      if (vals[i][2] === 'Available' && (!p.serial || Number(vals[i][0]) === Number(p.serial))) { row = i + 1; serial = vals[i][0]; code = vals[i][1]; break; }
    }
    if (!row) throw new Error(p.serial ? 'That code is not available' : 'No activation codes left. Generate more codes first.');
    sh.getRange(row, 3, 1, 4).setValues([['Sent', email, ts_(), who]]);
  } finally { lock.releaseLock(); }
  const subject = 'Your OKV HMS activation code';
  const html = wrap_('Your activation code', '<p>Hello ' + esc_(name) + ',</p><p>Thank you for your payment. Here is your activation code:</p>' +
    '<p style="font-size:24px;letter-spacing:2px;font-weight:bold;background:#f0fdfa;border:2px dashed #0f766e;border-radius:10px;padding:14px;text-align:center">' + esc_(code) + '</p>' +
    '<ol><li>Sign in to your HMS as <b>Administrator</b>.</li><li>Open <b>Licence</b> in the left menu.</li><li>Paste the code in <b>Enter activation code</b> and click <b>Activate</b>.</li></ol>' +
    '<p>Your data stays exactly as it is. Please keep this email safe and do not share the code.</p>' + (function () { const dl = downloads_(), l = [['Windows', dl.windows], ['Mac (Apple chip)', dl.macArm], ['Mac (Intel)', dl.macIntel]].filter(x => x[1]); return l.length ? '<p>Need the setup again? ' + l.map(x => '<a href="' + esc_(x[1]) + '">' + esc_(x[0]) + '</a>').join(' | ') + '</p>' : ''; })() + '<p>Need help? WhatsApp +' + CFG.WHATSAPP + '.</p>');
  try { mail_(email, subject, html, 'Your OKV HMS activation code: ' + code + '\nSign in as Administrator > Licence > Enter activation code > Activate.'); }
  catch (err) { sheet_(SH.CODES).getRange(row, 3, 1, 4).setValues([['Available', '', '', '']]); throw new Error('Email could not be sent: ' + err.message); }
  const cus = readAll_(SH.CUS).filter(c => String(c.Email).toLowerCase() === email)[0];
  if (cus) { setCell_(SH.CUS, cus._row, 'Status', 'Licensed'); setCell_(SH.CUS, cus._row, 'ActivationCode', code); setCell_(SH.CUS, cus._row, 'CodeSentAt', ts_()); }
  else sheet_(SH.CUS).appendRow([nextId_('C'), ts_(), name, '', email, '', '', 'Manual', 'Licensed', '', '', code, ts_(), '']);
  readAll_(SH.SUB).filter(s => String(s.Email).toLowerCase() === email && /Activation Request|Purchase Request|Paystack/.test(s.Type) && s.Status !== 'Done')
    .forEach(s => { setCell_(SH.SUB, s._row, 'Status', 'Done'); setCell_(SH.SUB, s._row, 'Notes', 'Code ' + code + ' sent ' + ts_()); });
  logMsg_('email (code)', email, subject, 'Activation code ' + code, 'Sent');
  return { code: code, serial: serial, quota: MailApp.getRemainingDailyQuota() };
}


/* ------------------------------------------------------------ remote configuration (controlled from the dashboard) */
function getConfig_() {
  const p = props_(), g = k => p.getProperty(k) || '';
  return { ok: true, price: Number(g('PRICE') || CFG.PRICE), currency: 'NGN', trial_days: CFG.TRIAL_DAYS,
    pay: { bank: g('PAY_BANK'), account_name: g('PAY_ACCOUNT_NAME'), account_no: g('PAY_ACCOUNT_NO'), instructions: g('PAY_INSTRUCTIONS') },
    paystack: g('PAYSTACK_ENABLED') === '1' && !!g('PAYSTACK_SECRET') && g('PURCHASES_ENABLED') !== '0',
    manual_enabled: g('MANUAL_ENABLED') !== '0' && g('PURCHASES_ENABLED') !== '0', pricing_url: g('PRICING_URL'),
    announcement: g('ANNOUNCEMENT'), trials_enabled: g('TRIALS_ENABLED') !== '0', purchases_enabled: g('PURCHASES_ENABLED') !== '0',
    install: { text: g('INSTALL_TEXT'), whatsapp: CFG.WHATSAPP } };
}
function normPhone_(p) { let d = String(p || '').replace(/\D/g, ''); if (d.indexOf('00') === 0) d = d.slice(2); if (d.indexOf('0') === 0) d = '234' + d.slice(1); return d; }
function sendSms_(phone, text) {
  const p = props_(); if (p.getProperty('TERMII_ENABLED') !== '1' || !p.getProperty('TERMII_KEY')) throw new Error('SMS integration is switched off');
  const to = normPhone_(phone); if (to.length < 12) throw new Error('Invalid phone number');
  const r = UrlFetchApp.fetch('https://api.ng.termii.com/api/sms/send', { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ to: to, from: p.getProperty('TERMII_SENDER') || 'OKV', sms: text, type: 'plain', channel: 'generic', api_key: p.getProperty('TERMII_KEY') }) });
  if (r.getResponseCode() >= 300) throw new Error('SMS provider rejected the message');
  return true;
}

/* ------------------------------------------------------------ Paystack (card / transfer / USSD; the code is issued only after Paystack confirms the payment) */
function psFetch_(path, method, payload) {
  const secret = props_().getProperty('PAYSTACK_SECRET'); if (!secret) throw new Error('Paystack is not set up');
  const opt = { method: method, muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + secret } };
  if (payload) { opt.contentType = 'application/json'; opt.payload = JSON.stringify(payload); }
  const r = UrlFetchApp.fetch('https://api.paystack.co' + path, opt);
  let body = {}; try { body = JSON.parse(r.getContentText()); } catch (e) { }
  return { code: r.getResponseCode(), body: body };
}
function paystackInit_(f) {
  const p = props_();
  if (p.getProperty('PAYSTACK_ENABLED') !== '1' || !p.getProperty('PAYSTACK_SECRET')) return fail_('Online payment is not available right now. Please pay by bank transfer.');
  if (p.getProperty('PURCHASES_ENABLED') === '0') return fail_('Purchases are paused right now. Please message us on WhatsApp +' + CFG.WHATSAPP + '.');
  if (!f.org) return fail_('Please enter your hospital or business name.');
  const price = Number(p.getProperty('PRICE') || CFG.PRICE);                      // the amount always comes from here, never from the browser
  const ref = 'okv-' + Utilities.getUuid().replace(/-/g, '').slice(0, 20).toLowerCase();
  const r = psFetch_('/transaction/initialize', 'post', { email: f.email, amount: Math.round(price * 100), currency: 'NGN', reference: ref,
    callback_url: ScriptApp.getService().getUrl() + '?action=pay_return', metadata: { name: f.name, hospital: f.org, phone: f.phone, install_id: f.installId, product: CFG.PRODUCT } });
  const url = r.body && r.body.data && r.body.data.authorization_url;
  if (r.code !== 200 || !r.body.status || !url) return fail_('Could not start the payment. Please try again or pay by bank transfer.');
  sheet_(SH.PAY).appendRow([ref, ts_(), f.email, f.name, f.org, f.phone, price, 'NGN', 'Paystack', 'Initiated', '', f.installId, '']);
  return { ok: true, url: url, reference: ref };
}
function paystackVerify_(refIn) {
  const ref = String(refIn || '').toLowerCase();
  if (!/^okv-[a-f0-9]{20}$/.test(ref)) return fail_('Unknown payment reference.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let row;
  try {
    row = readAll_(SH.PAY).filter(r => String(r.Reference).toLowerCase() === ref)[0];
    if (!row) return fail_('Unknown payment reference.');                         // only payments that this system started are ever processed
    if (row.Status === 'Paid') return { ok: true, paid: true, code: row.Code || '', email: row.Email, already: true };
    if (row.Status === 'Processing') return { ok: false, pending: true, error: 'Your payment is being processed. Please refresh this page in a few seconds.' };
    setCell_(SH.PAY, row._row, 'Status', 'Processing');
  } finally { lock.releaseLock(); }
  try {
    const r = psFetch_('/transaction/verify/' + encodeURIComponent(ref), 'get'), d = r.body && r.body.data;
    if (r.code !== 200 || !r.body.status || !d) { setCell_(SH.PAY, row._row, 'Status', 'Initiated'); return fail_('We could not confirm the payment yet. If you were charged, contact us with reference ' + ref + '.'); }
    if (d.status !== 'success') { setCell_(SH.PAY, row._row, 'Status', (d.status === 'failed' || d.status === 'abandoned') ? 'Failed' : 'Initiated'); return fail_('The payment was not completed (status: ' + d.status + ').'); }
    if (Number(d.amount) !== Math.round(Number(row.Amount) * 100) || String(d.currency).toUpperCase() !== 'NGN' || String(d.reference).toLowerCase() !== ref) {
      setCell_(SH.PAY, row._row, 'Status', 'Mismatch'); setCell_(SH.PAY, row._row, 'Note', 'Paystack amount/currency did not match');
      notifyOwner_('Payment mismatch', 'Reference ' + ref + ' for ' + row.Email + ' did not match the expected amount. Check Paystack.');
      return fail_('The amount paid does not match the price. Please contact us with reference ' + ref + '.');
    }
    setCell_(SH.PAY, row._row, 'Status', 'Paid');
    const f = { name: row.Name, org: row.Organization, email: String(row.Email).toLowerCase(), phone: row.Phone, city: '', message: 'Paystack ref ' + ref, installId: row.InstallID };
    const sid = submit_('Paystack Payment', f, { payRef: ref, amount: String(row.Amount), paidDate: ts_().slice(0, 10), receiptUrl: '' });
    let code = '';
    if (props_().getProperty('PAYSTACK_AUTO_CODE') !== '0') {
      try { code = sendCode_({ email: f.email, name: f.name }, 'paystack').code; setCell_(SH.PAY, row._row, 'Code', code); }
      catch (err) { console.error(err); setCell_(SH.PAY, row._row, 'Note', 'Paid. Code email failed: ' + err.message); notifyOwner_('Send code manually', 'Paystack payment ' + ref + ' for ' + f.email + ' is paid but the code email failed: ' + err.message); }
    } else setCell_(SH.PAY, row._row, 'Note', 'Paid. Waiting for you to send the code');
    return { ok: true, paid: true, code: code, email: f.email, id: sid };
  } catch (err) {
    console.error(err); try { setCell_(SH.PAY, row._row, 'Status', 'Initiated'); } catch (e) { }
    return fail_('Something went wrong while confirming the payment. If you were charged, contact us with reference ' + ref + '.');
  }
}
function notifyOwner_(subject, text) {
  try { mail_(props_().getProperty('OWNER_EMAIL') || CFG.OWNER_EMAIL, '[OKV HMS] ' + subject, wrap_(subject, '<p>' + esc_(text) + '</p>'), text); } catch (e) { console.error(e); }
  notifyExtra_('[OKV HMS] ' + subject, text, { event: 'notice', subject: subject, text: text });
}
function payReturnPage_(p) {
  const ref = clean_(p.reference || p.trxref, 60);
  let r; try { r = paystackVerify_(ref); } catch (e) { console.error(e); r = fail_('Something went wrong. Please contact support with your reference ' + ref + '.'); }
  const body = r.ok && r.paid
    ? '<h2 style="color:#0f766e">Payment received. Thank you!</h2>' + (r.code ? '<p>Your activation code:</p><p style="font-size:24px;letter-spacing:2px;font-weight:bold;background:#f0fdfa;border:2px dashed #0f766e;border-radius:10px;padding:14px;text-align:center">' + esc_(r.code) + '</p><p>We also emailed it to <b>' + esc_(r.email) + '</b>.</p><ol><li>Sign in to your HMS as Administrator.</li><li>Open <b>Licence</b>.</li><li>Paste the code and click <b>Activate</b>.</li></ol>' : '<p>Your payment is confirmed. Your activation code will be emailed to <b>' + esc_(r.email) + '</b> shortly.</p>')
    : '<h2 style="color:#b91c1c">' + (r.pending ? 'Please wait a moment' : 'We could not complete this') + '</h2><p>' + esc_(r.error || 'Please contact support.') + '</p><p>Reference: ' + esc_(ref) + '</p>';
  return HtmlService.createHtmlOutput('<div style="font-family:Arial,sans-serif;max-width:520px;margin:30px auto;padding:20px;border:1px solid #e2e8f0;border-radius:12px">' + body + '<p style="color:#64748b;font-size:13px">' + esc_(CFG.BRAND) + ' | WhatsApp +' + CFG.WHATSAPP + ' | ' + esc_(CFG.OWNER_EMAIL) + '</p></div>').setTitle('OKV HMS payment').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ------------------------------------------------------------ admin: payments, configuration, integrations, reset codes */
function mask_(v) { return v ? '\u2022\u2022\u2022\u2022' + String(v).slice(-4) : ''; }
function adminPayments(token) { adm_(token); return readAll_(SH.PAY).reverse().slice(0, 300); }
function adminRecordPayment(token, o) {
  adm_(token);
  const email = clean_(o.email, 120).toLowerCase(), name = clean_(o.name, 100), amount = Number(o.amount);
  if (!validEmail_(email)) throw new Error('A valid email is required'); if (!name) throw new Error('Name is required'); if (!(amount > 0)) throw new Error('Enter the amount');
  const ref = 'manual-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);
  sheet_(SH.PAY).appendRow([ref, ts_(), email, name, clean_(o.org, 120), clean_(o.phone, 30), amount, 'NGN', 'Manual', 'Paid', '', '', clean_(o.note, 200)]);
  const cus = readAll_(SH.CUS).filter(c => String(c.Email).toLowerCase() === email)[0];
  if (!cus) sheet_(SH.CUS).appendRow([nextId_('C'), ts_(), name, clean_(o.org, 120), email, clean_(o.phone, 30), '', 'Manual', 'Pending activation', '', '', '', '', '']);
  let code = '';
  if (o.sendCode) { code = sendCode_({ email: email, name: name }, 'dashboard').code; const row = readAll_(SH.PAY).filter(r => r.Reference === ref)[0]; setCell_(SH.PAY, row._row, 'Code', code); }
  return { reference: ref, code: code };
}
function adminConfig(token) {
  adm_(token); const p = props_(), g = k => p.getProperty(k) || '';
  return { price: Number(g('PRICE') || CFG.PRICE), pay: { bank: g('PAY_BANK'), account_name: g('PAY_ACCOUNT_NAME'), account_no: g('PAY_ACCOUNT_NO'), instructions: g('PAY_INSTRUCTIONS') },
    announcement: g('ANNOUNCEMENT'), installText: g('INSTALL_TEXT'), trialsEnabled: g('TRIALS_ENABLED') !== '0', purchasesEnabled: g('PURCHASES_ENABLED') !== '0', manualEnabled: g('MANUAL_ENABLED') !== '0', pricingUrl: g('PRICING_URL'),
    paystack: { enabled: g('PAYSTACK_ENABLED') === '1', publicKey: g('PAYSTACK_PUBLIC'), secretMask: mask_(g('PAYSTACK_SECRET')), autoCode: g('PAYSTACK_AUTO_CODE') !== '0' },
    termii: { enabled: g('TERMII_ENABLED') === '1', keyMask: mask_(g('TERMII_KEY')), sender: g('TERMII_SENDER') },
    telegram: { enabled: g('TG_ENABLED') === '1', tokenMask: mask_(g('TG_TOKEN')), chat: g('TG_CHAT') },
    webhook: { enabled: g('WEBHOOK_ENABLED') === '1', url: g('WEBHOOK_URL'), secretMask: mask_(g('WEBHOOK_SECRET')) },
    alerts: { emailEnabled: g('ALERT_EMAIL_ENABLED') !== '0', ownerEmail: g('OWNER_EMAIL') || CFG.OWNER_EMAIL },
    webAppUrl: ScriptApp.getService().getUrl() };
}
function adminSaveConfig(token, c) {
  adm_(token); const p = props_(), setIf = (k, v) => { if (v !== undefined && v !== null) p.setProperty(k, String(v)); };
  const secret = (k, v) => { if (v && String(v).indexOf('\u2022') !== 0) p.setProperty(k, String(v).trim()); };
  const flag = v => v ? '1' : '0';
  if (c.price !== undefined) { const n = Number(c.price); if (!(n >= 100 && n <= 100000000)) throw new Error('Price must be a sensible amount in naira'); p.setProperty('PRICE', String(Math.round(n))); }
  if (c.pay) { setIf('PAY_BANK', clean_(c.pay.bank, 80)); setIf('PAY_ACCOUNT_NAME', clean_(c.pay.account_name, 80)); setIf('PAY_ACCOUNT_NO', clean_(c.pay.account_no, 30)); setIf('PAY_INSTRUCTIONS', clean_(c.pay.instructions, 300)); }
  if (c.announcement !== undefined) setIf('ANNOUNCEMENT', clean_(c.announcement, 300));
  if (c.installText !== undefined) setIf('INSTALL_TEXT', clean_(c.installText, 500));
  if (c.trialsEnabled !== undefined) p.setProperty('TRIALS_ENABLED', flag(c.trialsEnabled));
  if (c.purchasesEnabled !== undefined) p.setProperty('PURCHASES_ENABLED', flag(c.purchasesEnabled));
  if (c.manualEnabled !== undefined) p.setProperty('MANUAL_ENABLED', flag(c.manualEnabled));
  if (c.pricingUrl !== undefined) { const u = clean_(c.pricingUrl, 300); if (u && !/^https:\/\//.test(u)) throw new Error('Pricing page URL must start with https://'); p.setProperty('PRICING_URL', u); }
  if (c.paystack) {
    if (c.paystack.enabled && !(c.paystack.secret && String(c.paystack.secret).indexOf('\u2022') !== 0) && !p.getProperty('PAYSTACK_SECRET')) throw new Error('Enter your Paystack secret key before switching Paystack on');
    p.setProperty('PAYSTACK_ENABLED', flag(c.paystack.enabled)); setIf('PAYSTACK_PUBLIC', clean_(c.paystack.publicKey, 120)); secret('PAYSTACK_SECRET', c.paystack.secret);
    if (c.paystack.autoCode !== undefined) p.setProperty('PAYSTACK_AUTO_CODE', flag(c.paystack.autoCode));
    const sk = p.getProperty('PAYSTACK_SECRET') || ''; if (sk && !/^sk_(live|test)_/.test(sk)) { p.deleteProperty('PAYSTACK_SECRET'); p.setProperty('PAYSTACK_ENABLED', '0'); throw new Error('That does not look like a Paystack secret key (it starts with sk_live_ or sk_test_)'); }
  }
  if (c.termii) { p.setProperty('TERMII_ENABLED', flag(c.termii.enabled)); secret('TERMII_KEY', c.termii.key); setIf('TERMII_SENDER', clean_(c.termii.sender, 11)); }
  if (c.telegram) { p.setProperty('TG_ENABLED', flag(c.telegram.enabled)); secret('TG_TOKEN', c.telegram.token); setIf('TG_CHAT', clean_(c.telegram.chat, 40)); }
  if (c.webhook) {
    const u = clean_(c.webhook.url, 400); if (u && !/^https:\/\//.test(u)) throw new Error('Webhook URL must start with https://');
    p.setProperty('WEBHOOK_ENABLED', flag(c.webhook.enabled)); setIf('WEBHOOK_URL', u); secret('WEBHOOK_SECRET', c.webhook.secret);
  }
  if (c.alerts) { p.setProperty('ALERT_EMAIL_ENABLED', flag(c.alerts.emailEnabled)); if (c.alerts.ownerEmail) { const e = clean_(c.alerts.ownerEmail, 120).toLowerCase(); if (!validEmail_(e)) throw new Error('Alert email is not valid'); p.setProperty('OWNER_EMAIL', e); } }
  return true;
}
function adminTestIntegration(token, which) {
  adm_(token); const p = props_();
  if (which === 'paystack') { const r = psFetch_('/balance', 'get'); if (r.code !== 200 || !r.body.status) throw new Error('Paystack rejected the key (' + r.code + ')'); return 'Paystack key works.'; }
  if (which === 'telegram') { const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + p.getProperty('TG_TOKEN') + '/sendMessage', { method: 'post', muteHttpExceptions: true, payload: { chat_id: p.getProperty('TG_CHAT'), text: 'OKV HMS test alert: Telegram is connected.' } }); if (r.getResponseCode() !== 200) throw new Error('Telegram rejected it (' + r.getResponseCode() + ')'); return 'Test message sent to Telegram.'; }
  if (which === 'termii') { const r = UrlFetchApp.fetch('https://api.ng.termii.com/api/get-balance?api_key=' + encodeURIComponent(p.getProperty('TERMII_KEY') || ''), { muteHttpExceptions: true }); if (r.getResponseCode() !== 200) throw new Error('Termii rejected the key (' + r.getResponseCode() + ')'); return 'Termii key works.'; }
  if (which === 'webhook') { const u = p.getProperty('WEBHOOK_URL') || ''; if (!/^https:\/\//.test(u)) throw new Error('Save a webhook URL first'); const r = UrlFetchApp.fetch(u, { method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { 'X-OKV-Secret': p.getProperty('WEBHOOK_SECRET') || '' }, payload: JSON.stringify({ event: 'test', time: ts_() }) }); return 'Webhook answered with HTTP ' + r.getResponseCode() + '.'; }
  throw new Error('Unknown integration');
}
function resetCodeFor_(installId, day, secret) {
  const mac = hmacHex_('RESET|' + installId + '|' + day, secret).toUpperCase().slice(0, 10);
  return 'RST-' + mac.slice(0, 5) + '-' + mac.slice(5);
}
function adminResetCode(token, installId) {
  adm_(token); installId = clean_(installId, 20).toUpperCase();
  if (!/^HMS-[0-9A-F]{8}$/.test(installId)) throw new Error('Installation ID looks like HMS-1A2B3C4D');
  const code = resetCodeFor_(installId, Utilities.formatDate(new Date(), CFG.TZ, 'yyyyMMdd'), props_().getProperty('LICENSE_SECRET'));
  logMsg_('reset code', installId, '', 'Admin password reset code issued', 'Issued');
  return code;
}

/* ------------------------------------------------------------ admin dashboard API (every call needs a session token) */
function safeEq_(a, b) { if (a.length !== b.length) return false; let r = 0; for (let i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i); return r === 0; }
function adm_(token) { if (!token || !CacheService.getScriptCache().get('tok_' + token)) throw new Error('Session expired. Please sign in again.'); }
function adminLogin(password) {
  const c = CacheService.getScriptCache(), fails = Number(c.get('adm_fail') || 0), real = props_().getProperty('ADMIN_PASSWORD') || '';
  if (fails >= 5) throw new Error('Too many wrong attempts. Try again in 15 minutes.');
  if (!real || !safeEq_(String(password || ''), real)) { c.put('adm_fail', String(fails + 1), 900); throw new Error('Wrong password'); }
  c.remove('adm_fail');
  const t = Utilities.getUuid() + Utilities.getUuid(); c.put('tok_' + t, '1', 21600); return t;
}
function adminLogout(token) { CacheService.getScriptCache().remove('tok_' + token); return true; }
function adminSummary(token) {
  adm_(token);
  const sub = readAll_(SH.SUB), cus = readAll_(SH.CUS), codes = readAll_(SH.CODES), al = readAll_(SH.ALERTS);
  const by = (arr, k) => arr.reduce((o, r) => { o[r[k]] = (o[r[k]] || 0) + 1; return o; }, {});
  return { subNew: sub.filter(s => s.Status === 'New').length, subTotal: sub.length, subByType: by(sub, 'Type'), cusByStatus: by(cus, 'Status'), cusTotal: cus.length,
    codesAvailable: codes.filter(c => c.Status === 'Available').length, codesSent: codes.filter(c => c.Status === 'Sent').length, codesTotal: codes.length,
    unread: al.filter(a => a.Read !== 'Yes').length, quota: MailApp.getRemainingDailyQuota(), pendingActivations: sub.filter(s => /Request|Paystack/.test(s.Type) && s.Status !== 'Done' && s.Type !== 'Installation Request').length,
    installPending: sub.filter(s => s.Type === 'Installation Request' && s.Status !== 'Done').length,
    revenue: readAll_(SH.PAY).filter(r => r.Status === 'Paid').reduce((a, r) => a + Number(r.Amount || 0), 0), paidCount: readAll_(SH.PAY).filter(r => r.Status === 'Paid').length };
}
function adminAlerts(token) {
  adm_(token); const al = readAll_(SH.ALERTS).reverse(), subs = {};
  readAll_(SH.SUB).forEach(x => { subs[x.ID] = x; });
  return { unread: al.filter(a => a.Read !== 'Yes').length, items: al.slice(0, 60).map(a => { const sb = subs[a.SubID]; return { id: a.ID, time: a.Time, type: a.Type, text: a.Text, read: a.Read === 'Yes', subId: a.SubID, canApprove: !!sb && sb.Status !== 'Done' && /Activation Request|Purchase Request|Paystack Payment/.test(sb.Type) }; }) };
}
function adminMarkAlertsRead(token) {
  adm_(token); const sh = sheet_(SH.ALERTS), n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 6, n, 1).setValues(Array.from({ length: n }, () => ['Yes']));
  return true;
}
function adminSubmissions(token, type, status, q) {
  adm_(token); q = String(q || '').toLowerCase();
  return readAll_(SH.SUB).reverse().filter(s => (!type || s.Type === type) && (!status || s.Status === status) &&
    (!q || [s.Name, s.Organization, s.Email, s.Phone, s.PayRef, s.ID].join(' ').toLowerCase().indexOf(q) >= 0)).slice(0, 300);
}
function adminUpdateSubmission(token, id, status, notes) {
  adm_(token); const s = readAll_(SH.SUB).filter(x => x.ID === id)[0]; if (!s) throw new Error('Submission not found');
  if (['New', 'In progress', 'Done'].indexOf(status) < 0) throw new Error('Unknown status');
  setCell_(SH.SUB, s._row, 'Status', status); setCell_(SH.SUB, s._row, 'Notes', clean_(notes, 500)); return true;
}
function adminCustomers(token, q) {
  adm_(token); q = String(q || '').toLowerCase();
  return readAll_(SH.CUS).reverse().filter(c => !q || [c.Name, c.Organization, c.Email, c.Phone, c.Status].join(' ').toLowerCase().indexOf(q) >= 0).slice(0, 500);
}
function adminSaveCustomer(token, c) {
  adm_(token);
  const email = clean_(c.Email, 120).toLowerCase(), name = clean_(c.Name, 100);
  if (!name) throw new Error('Name is required'); if (!validEmail_(email)) throw new Error('A valid email is required');
  if (['Lead', 'Trial', 'Pending activation', 'Licensed'].indexOf(c.Status) < 0) throw new Error('Unknown status');
  const vals = [clean_(c.Name, 100), clean_(c.Organization, 120), email, clean_(c.Phone, 30), clean_(c.Location, 80)];
  if (c.ID) {
    const cur = readAll_(SH.CUS).filter(x => x.ID === c.ID)[0]; if (!cur) throw new Error('Customer not found');
    ['Name', 'Organization', 'Email', 'Phone', 'Location'].forEach((k, i) => setCell_(SH.CUS, cur._row, k, vals[i]));
    setCell_(SH.CUS, cur._row, 'Status', c.Status); setCell_(SH.CUS, cur._row, 'Notes', clean_(c.Notes, 500));
  } else {
    if (readAll_(SH.CUS).some(x => String(x.Email).toLowerCase() === email)) throw new Error('A customer with that email already exists');
    sheet_(SH.CUS).appendRow([nextId_('C'), ts_(), vals[0], vals[1], vals[2], vals[3], vals[4], 'Manual', c.Status, '', '', '', '', clean_(c.Notes, 500)]);
  }
  return true;
}
function adminDeleteCustomer(token, id) {
  adm_(token); const cur = readAll_(SH.CUS).filter(x => x.ID === id)[0]; if (!cur) throw new Error('Customer not found');
  sheet_(SH.CUS).deleteRow(cur._row); return true;
}
function adminCodes(token, status, page) {
  adm_(token); const all = readAll_(SH.CODES), f = status ? all.filter(c => c.Status === status) : all, size = 100, pg = Math.max(0, Number(page) || 0);
  return { total: f.length, page: pg, pages: Math.ceil(f.length / size), available: all.filter(c => c.Status === 'Available').length, rows: f.slice(pg * size, pg * size + size) };
}
function approveSub_(subId, who) {
  const sub = readAll_(SH.SUB).filter(x => x.ID === subId)[0];
  if (!sub) throw new Error('Submission ' + subId + ' not found');
  if (!/Activation Request|Purchase Request|Paystack Payment/.test(sub.Type)) throw new Error('Only activation or payment requests can be approved');
  if (sub.Status === 'Done') throw new Error('Already approved or closed');
  const email = String(sub.Email).toLowerCase(), res = sendCode_({ email: email, name: sub.Name }, who);   // emails the code at once; marks the request Done and the customer Licensed
  const ref = 'manual-' + String(sub.ID).toLowerCase();
  if (sub.Type !== 'Paystack Payment' && !readAll_(SH.PAY).some(r => r.Reference === ref))
    sheet_(SH.PAY).appendRow([ref, ts_(), email, sub.Name, sub.Organization, sub.Phone, Number(sub.Amount) || 0, 'NGN', 'Manual', 'Paid', res.code, sub.InstallID, 'Approved from ' + sub.ID]);
  return { code: res.code, email: email, name: sub.Name, quota: res.quota };
}
function adminApprove(token, subId) { adm_(token); return approveSub_(clean_(subId, 20), 'approved (dashboard)'); }
function adminApproveMany(token, ids) {
  adm_(token); const ok = [], errors = [];
  (ids || []).slice(0, 50).forEach(id => { try { ok.push(approveSub_(clean_(id, 20), 'approved (dashboard)')); } catch (e) { errors.push(id + ': ' + e.message); } });
  return { approved: ok.length, errors: errors, quota: MailApp.getRemainingDailyQuota() };
}
function adminSendCode(token, p) { adm_(token); return sendCode_(p || {}, 'dashboard'); }
function adminIssueCode(token, label) {
  adm_(token); label = clean_(label, 100) || 'issued by hand';
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet_(SH.CODES), vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      if (vals[i][2] === 'Available') { sh.getRange(i + 1, 3, 1, 4).setValues([['Sent', label, ts_(), 'dashboard (by hand)']]); logMsg_('code issued', label, '', 'Activation code ' + vals[i][1] + ' issued by hand', 'Issued'); return { code: vals[i][1], serial: vals[i][0] }; }
    }
  } finally { lock.releaseLock(); }
  throw new Error('No activation codes left. Generate more codes first.');
}
function adminGenerateCodes(token, n) {
  adm_(token); n = Math.floor(Number(n)); if (!(n >= 1 && n <= 2000)) throw new Error('Enter a number from 1 to 2000');
  return generateCodes_(n);
}
function adminSendMessage(token, p) {
  adm_(token);
  if (p.channel === 'sms') return adminSendSms_(p);
  const subject = clean_(p.subject, 150), body = clean_(p.body, 3000);
  if (!subject || !body) throw new Error('Subject and message are required');
  let list = [];
  if (p.segment === 'selected') list = (p.recipients || []).map(r => ({ Name: clean_(r.name, 100), Email: clean_(r.email, 120).toLowerCase() }));
  else {
    const map = { all: null, trial: 'Trial', pending: 'Pending activation', licensed: 'Licensed', lead: 'Lead' };
    if (!(p.segment in map)) throw new Error('Unknown group');
    list = readAll_(SH.CUS).filter(c => !map[p.segment] || c.Status === map[p.segment]).map(c => ({ Name: c.Name, Email: String(c.Email).toLowerCase() }));
  }
  list = list.filter(r => validEmail_(r.Email)); if (!list.length) throw new Error('No recipients with a valid email');
  const quota = MailApp.getRemainingDailyQuota();
  if (list.length > quota) throw new Error('Your daily email limit is ' + quota + ' but you chose ' + list.length + ' recipients. Choose fewer.');
  let sent = 0, failed = 0;
  list.forEach(r => {
    const text = body.replace(/\{name\}/g, r.Name || 'Customer');
    try { mail_(r.Email, subject, wrap_(subject, '<p>' + esc_(text).replace(/\n/g, '<br>') + '</p>'), text); sent++; logMsg_('email', r.Email, subject, text, 'Sent'); }
    catch (e) { failed++; logMsg_('email', r.Email, subject, text, 'FAILED: ' + e.message); }
  });
  return { sent: sent, failed: failed, quota: MailApp.getRemainingDailyQuota() };
}
function adminSendSms_(p) {
  const body = clean_(p.body, 300); if (!body) throw new Error('Type the message');
  const map = { all: null, trial: 'Trial', pending: 'Pending activation', licensed: 'Licensed', lead: 'Lead' };
  let list = p.segment === 'selected' ? (p.recipients || []).map(r => ({ Name: clean_(r.name, 100), Phone: clean_(r.phone, 30) })) :
    readAll_(SH.CUS).filter(c => !map[p.segment] || c.Status === map[p.segment]);
  list = list.filter(r => normPhone_(r.Phone).length >= 12); if (!list.length) throw new Error('No customers with a valid phone number');
  let sent = 0, failed = 0;
  list.forEach(r => { try { sendSms_(r.Phone, body.replace(/\{name\}/g, r.Name || 'Customer')); sent++; logMsg_('sms (auto)', normPhone_(r.Phone), '', body, 'Sent'); } catch (e) { failed++; logMsg_('sms (auto)', r.Phone, '', body, 'FAILED: ' + e.message); } });
  return { sent: sent, failed: failed };
}
function adminLogMessage(token, p) { adm_(token); logMsg_(clean_(p.channel, 20), clean_(p.to, 60), '', clean_(p.body, 500), 'Opened on admin device'); return true; }
function adminMessages(token) { adm_(token); return readAll_(SH.MSG).reverse().slice(0, 200); }
function adminSettings(token) {
  adm_(token); const p = props_();
  return { licenseSecret: p.getProperty('LICENSE_SECRET'), downloadUrl: p.getProperty('DOWNLOAD_URL') || '', macArmUrl: p.getProperty('DOWNLOAD_URL_MAC_ARM') || '', macIntelUrl: p.getProperty('DOWNLOAD_URL_MAC_INTEL') || '', webAppUrl: ScriptApp.getService().getUrl(), ownerEmail: CFG.OWNER_EMAIL, quota: MailApp.getRemainingDailyQuota(), spreadsheetUrl: ss_().getUrl() };
}
function adminSaveSettings(token, s) {
  adm_(token); const p = props_();
  [['downloadUrl', 'DOWNLOAD_URL'], ['macArmUrl', 'DOWNLOAD_URL_MAC_ARM'], ['macIntelUrl', 'DOWNLOAD_URL_MAC_INTEL']].forEach(pair => {
    if (s[pair[0]] != null) { const u = clean_(s[pair[0]], 500); if (u && !/^https:\/\//.test(u)) throw new Error('Download links must start with https://'); p.setProperty(pair[1], u); }
  });
  if (s.newPassword) { const pw = String(s.newPassword); if (pw.length < 8) throw new Error('Password must be at least 8 characters'); p.setProperty('ADMIN_PASSWORD', pw); }
  return true;
}
