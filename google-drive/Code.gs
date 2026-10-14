/**
 * Engler questionnaire → Google Drive uploader and email notifier.
 *
 * Receives files attached on the questionnaire (client-form.html) and saves them to
 * a Drive folder, one subfolder per submission, emails NOTIFY_EMAIL when
 * a new submission is saved, and emails the client a confirmation with their
 * order number, and starts Stripe checkouts for the payment page. Deploy as a web app:
 *   Execute as: Me    Who has access: Anyone
 * then put the /exec URL into DRIVE_UPLOAD_URL in client-form.html.
 *
 * Run setup() once from the editor first: it asks for Drive permission and
 * creates the uploads folder (its link is printed in the execution log).
 */

const ROOT_FOLDER_NAME = 'Engler Questionnaire Uploads';
const MAX_BYTES = 20 * 1024 * 1024;

// Lets the portal (portal.html) show image previews. Files are only viewable by someone
// who has the link, and links are only listed on the password-protected
// responses page. Set to false to keep files private to your Google account
// (previews then only load while you're signed in to Google in that browser).
const SHARE_WITH_LINK = true;

// Where new-submission emails go, and the page linked from them.
const NOTIFY_EMAIL = 'englerwebsites@gmail.com';
const RESPONSES_URL = 'https://englerwebsites.com/portal';

// Confirmation emails to clients. The recipient comes from the public form, so
// the text is fixed, each submission gets at most one, and there's a daily cap.
const CONFIRMATION_SUBJECT = 'Questionnaire confirmation';
const MAX_CONFIRMATIONS_PER_DAY = 50;

// Stripe checkout for englerwebsites.com/payment. Put a Stripe key that can write Checkout
// Sessions in Project Settings > Script properties as STRIPE_SECRET_KEY. The amount always
// comes from the order's published total (orderTotals in Firestore), never from the browser.
const PAYMENT_PAGE_URL = 'https://englerwebsites.com/payment';
const FIRESTORE_DOCS = 'https://firestore.googleapis.com/v1/projects/engler-questionnaire/databases/(default)/documents';
const FIREBASE_WEB_KEY = 'AIzaSyCrmEk6X74MWQH6olX2N5Q5Zd7NJGY_SDk';
const ORDER_NUMBER_PATTERN = /^ENG-[0-9A-HJKMNP-TV-Z]{4}-[0-9A-HJKMNP-TV-Z]{4}$/;

const ALLOWED_TYPES = /^(image\/[\w.+-]+|application\/(pdf|postscript|illustrator|octet-stream))$/;

function setup() {
  Logger.log('Uploads folder: ' + getRootFolder_().getUrl());
}

// Run once from the editor after adding email notifications: it asks for
// permission to send email and sends a test message to NOTIFY_EMAIL.
function testNotification() {
  sendNotification_({
    submissionId: 'testNotification01',
    company: 'TEST — notification check',
    replyTo: '',
    appraisal: { pages: 5, pagesSource: 'estimated page count', pricePerPage: 100, pagesCost: 500,
                 features: ['Custom inquiry forms', 'Blog'], pricePerFeature: 50, featuresCost: 100, total: 600 },
    summary: 'This is a test of the Engler questionnaire email notification.'
  });
  Logger.log('Test email sent to ' + NOTIFY_EMAIL);
}

// Run from the editor to preview the client confirmation email (sent to NOTIFY_EMAIL).
function testConfirmation() {
  MailApp.sendEmail(confirmationEmail_(NOTIFY_EMAIL, orderNumber_('testConfirmation01')));
  Logger.log('Test confirmation sent to ' + NOTIFY_EMAIL);
}

function doGet() {
  return ContentService.createTextOutput('Engler uploader is running.');
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.action === 'checkout') {
      return json_(createCheckout_(String(req.orderNo || '')));
    }
    if (req.action === 'notify') {
      sendNotification_(req);
      try { sendConfirmation_(req); } catch (err) { console.warn('Confirmation email not sent: ' + err); }
      return json_({ ok: true });
    }

    const submissionId = String(req.submissionId || '');
    if (!/^[A-Za-z0-9]{10,40}$/.test(submissionId)) throw new Error('Invalid submission id');

    const type = String(req.type || 'application/octet-stream');
    if (!ALLOWED_TYPES.test(type)) throw new Error('File type not allowed');

    const bytes = Utilities.base64Decode(String(req.data || ''));
    if (!bytes.length) throw new Error('Empty file');
    if (bytes.length > MAX_BYTES) throw new Error('File larger than 20 MB');

    const name = String(req.name || 'file').replace(/[\\/:*?"<>|]+/g, '_').slice(0, 150);
    const company = String(req.company || 'Unknown company').replace(/[\\/:*?"<>|]+/g, '_').slice(0, 100);

    const folder = getSubmissionFolder_(submissionId, company);
    const file = folder.createFile(Utilities.newBlob(bytes, type, name));
    if (SHARE_WITH_LINK) file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return json_({ ok: true, id: file.getId(), url: file.getUrl() });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

// Emails a new submission's answers. The text comes from the public form, so
// it's sent as plain text only.
function sendNotification_(req) {
  const submissionId = String(req.submissionId || '');
  if (!/^[A-Za-z0-9]{10,40}$/.test(submissionId)) throw new Error('Invalid submission id');
  const company = String(req.company || 'Unknown company').replace(/[\r\n]+/g, ' ').slice(0, 100);
  const summary = String(req.summary || '').slice(0, 50000);
  const replyTo = String(req.replyTo || '').trim();
  const ap = req.appraisal || null;

  const num = v => Math.max(0, Math.round(Number(v) || 0));
  const money = v => '$' + num(v).toLocaleString('en-US');
  const subject = 'New questionnaire response: ' + company + (ap ? ' (est. ' + money(ap.total) + ')' : '');

  let body = 'A new questionnaire response was submitted.\n\n' +
    'Business: ' + company + '\n';
  if (ap) {
    const features = Array.isArray(ap.features) ? ap.features.map(String).join(', ') : '';
    body += 'Estimated appraisal: ' + money(ap.total) + '\n' +
      '  • Pages: ' + num(ap.pages) + ' × ' + money(ap.pricePerPage) + ' = ' + money(ap.pagesCost) +
        (ap.pagesSource ? ' (' + String(ap.pagesSource).slice(0, 60) + ')' : '') + '\n' +
      '  • Interactive features: ' + num(Array.isArray(ap.features) ? ap.features.length : 0) + ' × ' + money(ap.pricePerFeature) +
        ' = ' + money(ap.featuresCost) + (features ? ' (' + features.slice(0, 300) + ')' : '') + '\n';
    if (ap.other) body += '  • Not priced, other features described: ' + String(ap.other).slice(0, 300) + '\n';
  }
  body += '\nView it (with any uploaded files):\n' + RESPONSES_URL + '\n\n' +
    '────────────────────────────────────────\n\n' + summary;

  const options = { name: 'Engler Questionnaire' };
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(replyTo)) options.replyTo = replyTo;

  MailApp.sendEmail(NOTIFY_EMAIL, subject, body, options);
}

function getRootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ROOT_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* deleted; recreate below */ }
  }
  const folder = DriveApp.createFolder(ROOT_FOLDER_NAME);
  props.setProperty('ROOT_FOLDER_ID', folder.getId());
  return folder;
}

// Files from one submission upload in parallel, so lock while finding or
// creating its folder to avoid making duplicates.
function getSubmissionFolder_(submissionId, company) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const root = getRootFolder_();
    const name = company + ' — ' + submissionId;
    const existing = root.getFoldersByName(name);
    return existing.hasNext() ? existing.next() : root.createFolder(name);
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Emails the client that their questionnaire arrived, with its order number.
function sendConfirmation_(req) {
  const submissionId = String(req.submissionId || '');
  if (!/^[A-Za-z0-9]{10,40}$/.test(submissionId)) throw new Error('Invalid submission id');
  const to = String(req.replyTo || '').trim();
  if (!/^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/.test(to) || to.length > 200) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const sentKey = 'confirmed_' + submissionId;
    if (props.getProperty(sentKey)) return;
    const today = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
    const countKey = 'confirmations_' + today;
    const count = Number(props.getProperty(countKey) || 0);
    if (count >= MAX_CONFIRMATIONS_PER_DAY) throw new Error('Daily confirmation limit reached');

    MailApp.sendEmail(confirmationEmail_(to, orderNumber_(submissionId)));
    props.setProperty(sentKey, today);
    props.setProperty(countKey, String(count + 1));
  } finally {
    lock.releaseLock();
  }
}

function confirmationEmail_(to, orderNo) {
  const body =
    'Thank you for filling out the website project questionnaire! Your order code is ' + orderNo +
    ', so please hold on to that. Our team will be in touch with you shortly!\n\n' +
    '— The Engler team\nenglerwebsites.com · (530) 691-2191';
  const htmlBody =
    '<div style="font-family:Helvetica,Arial,sans-serif;color:#16233B;max-width:520px;line-height:1.6">' +
      '<h2 style="margin:0 0 12px;font-size:20px">Questionnaire confirmation</h2>' +
      '<p style="margin:0 0 16px">Thank you for filling out the website project questionnaire!</p>' +
      '<p style="margin:0 0 6px">Your order code is</p>' +
      '<p style="margin:0 0 16px;font-family:Menlo,Consolas,monospace;font-size:20px;font-weight:bold;letter-spacing:1px">' + orderNo + '</p>' +
      '<p style="margin:0 0 16px">Please hold on to that. Our team will be in touch with you shortly!</p>' +
      '<p style="margin:24px 0 0;color:#5F6776;font-size:14px">— The Engler team<br>' +
        '<a href="https://englerwebsites.com" style="color:#2F6BFF">englerwebsites.com</a> · (530) 691-2191</p>' +
    '</div>';
  return { to: to, subject: CONFIRMATION_SUBJECT, body: body, htmlBody: htmlBody, name: 'Engler', replyTo: NOTIFY_EMAIL };
}

// Same order numbers as order-number.js on the website (ENG-XXXX-XXXX), worked out from
// the submission ID so the email can't be made to show any other text. Apps Script doesn't
// support BigInt, so the 64-bit hash is done on [high, low] pairs of 32-bit numbers.
function orderNumber_(id) {
  const ALPHABET = '0123456789ABCDEFGHJKMNPQRSTVWXYZ';
  let h = [0xcbf29ce4, 0x84222325];
  for (const ch of String(id)) {
    h = [h[0], (h[1] ^ ch.charCodeAt(0)) >>> 0];
    h = mul64_(h, [0x00000100, 0x000001b3]);
  }
  h = xor64_(h, shr64_(h, 33)); h = mul64_(h, [0xff51afd7, 0xed558ccd]);
  h = xor64_(h, shr64_(h, 33)); h = mul64_(h, [0xc4ceb9fe, 0x1a85ec53]);
  h = xor64_(h, shr64_(h, 33));
  let code = '';
  for (let i = 0; i < 8; i++) {
    code += ALPHABET[h[1] & 31];
    h = shr64_(h, 5);
  }
  return 'ENG-' + code.slice(0, 4) + '-' + code.slice(4);
}

function xor64_(a, b) { return [(a[0] ^ b[0]) >>> 0, (a[1] ^ b[1]) >>> 0]; }

function shr64_(a, n) {
  if (n >= 32) return [0, a[0] >>> (n - 32)];
  return [a[0] >>> n, ((a[1] >>> n) | (a[0] << (32 - n))) >>> 0];
}

// (a × b) mod 2^64, using 16-bit pieces so no step loses precision.
function mul64_(a, b) {
  const x = [a[1] & 0xffff, a[1] >>> 16, a[0] & 0xffff, a[0] >>> 16];
  const y = [b[1] & 0xffff, b[1] >>> 16, b[0] & 0xffff, b[0] >>> 16];
  const r = [0, 0, 0, 0];
  let carry = 0;
  for (let k = 0; k < 4; k++) {
    let t = carry;
    for (let i = 0; i <= k; i++) t += x[i] * y[k - i];
    r[k] = t % 65536;
    carry = Math.floor(t / 65536);
  }
  return [(r[3] * 65536 + r[2]) >>> 0, (r[1] * 65536 + r[0]) >>> 0];
}

// Starts a Stripe Checkout for an order's published total and returns its URL.
function createCheckout_(orderNo) {
  if (!ORDER_NUMBER_PATTERN.test(orderNo)) return { ok: false, error: 'Invalid order number' };
  const key = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');
  if (!key) return { ok: false, error: 'Online payment is not set up yet' };

  const res = UrlFetchApp.fetch(FIRESTORE_DOCS + '/orderTotals/' + encodeURIComponent(orderNo) + '?key=' + FIREBASE_WEB_KEY,
    { muteHttpExceptions: true });
  if (res.getResponseCode() === 404 || res.getResponseCode() === 403) return { ok: false, error: 'Order not found' };
  if (res.getResponseCode() !== 200) throw new Error('Order lookup failed: HTTP ' + res.getResponseCode());
  const fields = JSON.parse(res.getContentText()).fields || {};
  const total = Number(fields.total && (fields.total.doubleValue != null ? fields.total.doubleValue : fields.total.integerValue));
  const company = String(fields.company && fields.company.stringValue || '').slice(0, 200);
  const cents = Math.round(total * 100);
  if (!(cents >= 50)) return { ok: false, error: 'This order has no amount to pay online' };

  const back = PAYMENT_PAGE_URL + '?order=' + encodeURIComponent(orderNo);
  const description = 'Engler website, order ' + orderNo + (company ? ' (' + company + ')' : '');
  const stripe = UrlFetchApp.fetch('https://api.stripe.com/v1/checkout/sessions', {
    method: 'post',
    headers: { Authorization: 'Bearer ' + key },
    muteHttpExceptions: true,
    payload: {
      'mode': 'payment',
      'line_items[0][quantity]': '1',
      'line_items[0][price_data][currency]': 'usd',
      'line_items[0][price_data][unit_amount]': String(cents),
      'line_items[0][price_data][product_data][name]': 'Engler website — order ' + orderNo,
      'line_items[0][price_data][product_data][description]': company || 'Website design and build',
      'client_reference_id': orderNo,
      'metadata[order_number]': orderNo,
      'payment_intent_data[description]': description,
      'payment_intent_data[metadata][order_number]': orderNo,
      'success_url': back + '&paid=1',
      'cancel_url': back + '&cancelled=1'
    }
  });
  const session = JSON.parse(stripe.getContentText());
  if (stripe.getResponseCode() !== 200 || !session.url) {
    console.error('Stripe checkout failed: ' + stripe.getContentText());
    return { ok: false, error: "Couldn't start the payment" };
  }
  return { ok: true, url: session.url };
}
