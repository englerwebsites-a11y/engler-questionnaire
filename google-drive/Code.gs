/**
 * Engler questionnaire → Google Drive uploader and email notifier.
 *
 * Receives files attached on the questionnaire (index.html) and saves them to
 * a Drive folder, one subfolder per submission, and emails NOTIFY_EMAIL when
 * a new submission is saved. Deploy as a web app:
 *   Execute as: Me    Who has access: Anyone
 * then put the /exec URL into DRIVE_UPLOAD_URL in index.html.
 *
 * Run setup() once from the editor first: it asks for Drive permission and
 * creates the uploads folder (its link is printed in the execution log).
 */

const ROOT_FOLDER_NAME = 'Engler Questionnaire Uploads';
const MAX_BYTES = 20 * 1024 * 1024;

// Lets responses.html show image previews. Files are only viewable by someone
// who has the link, and links are only listed on the password-protected
// responses page. Set to false to keep files private to your Google account
// (previews then only load while you're signed in to Google in that browser).
const SHARE_WITH_LINK = true;

// Where new-submission emails go, and the page linked from them.
const NOTIFY_EMAIL = 'englerwebsites@gmail.com';
const RESPONSES_URL = 'https://englerwebsites.com/responses.html';

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

function doGet() {
  return ContentService.createTextOutput('Engler uploader is running.');
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.action === 'notify') {
      sendNotification_(req);
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
