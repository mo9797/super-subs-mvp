const SHEET_ID = '1-tema9OUKu0j2wner8aGvy3mUX1qhswhCRNW4jHfgqI';
const TX_SHEET = 'Transactions';
const MEMBERS_SHEET = 'Members';

function doGet(e) {
  const callback = (e && e.parameter && e.parameter.callback) || '';
  try {
    const email = getUserEmail_();
    assertMember_(email);
    const payload = {ok:true, transactions: readTransactions_(), members: readMembers_(), user: email};
    return json_(payload, callback);
  } catch (err) {
    return json_({ok:false, error: safeError_(err)}, callback);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const email = getUserEmail_();
    assertMember_(email);
    if (body.action === 'add') addTransaction_(body, email);
    else if (body.action === 'member') addMember_(body.email, email);
    else throw new Error('Unsupported action');
    return json_({ok:true});
  } catch (err) {
    return json_({ok:false, error:safeError_(err)});
  }
}

function getUserEmail_() {
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email) throw new Error('Google sign-in is required');
  return email;
}

function assertMember_(email) {
  const rows = readMembers_();
  const found = rows.some(function(row) {
    return String(row.email || '').toLowerCase() === email && String(row.status || '').toLowerCase() === 'active';
  });
  if (!found) throw new Error('This Google account is not an active member');
}

function readMembers_() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(MEMBERS_SHEET);
  const values = sheet.getDataRange().getDisplayValues();
  return values.slice(1).filter(function(r){ return r[0]; }).map(function(r){
    return {email:r[0], role:r[1] || 'Member', status:r[2] || 'Active'};
  });
}

function readTransactions_() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TX_SHEET);
  const values = sheet.getDataRange().getDisplayValues();
  return values.slice(1).filter(function(r){ return r[0]; }).map(function(r){
    return {id:r[0], type:r[1], qty:r[2] ? Number(r[2]) : null, price:r[3] ? Number(r[3]) : 0, total:r[4] ? Number(r[4]) : 0, party:r[5] || '', note:r[6] || '', date:r[7] || '', by:r[8] || ''};
  });
}

function addTransaction_(body, email) {
  const type = String(body.type || '');
  if (['purchase','sale','ad'].indexOf(type) === -1) throw new Error('Invalid transaction type');
  const qty = Number(body.qty);
  const price = Number(body.price);
  if (!isFinite(qty) || qty <= 0 || !isFinite(price) || price < 0) throw new Error('Invalid amount');
  const total = type === 'ad' ? price : qty * price;
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TX_SHEET);
  const party = clean_(body.party, 40);
  if ((type === 'purchase' || type === 'sale') && !party) throw new Error('Financial party is required');
  sheet.appendRow([Utilities.getUuid(), type, type === 'ad' ? 1 : qty, price, total, party, clean_(body.note,120), new Date(), email]);
}

function addMember_(rawEmail, requester) {
  const email = String(rawEmail || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Invalid email');
  const members = readMembers_();
  const owner = members.some(function(m){ return m.email === requester && m.role.toLowerCase() === 'owner'; });
  if (!owner) throw new Error('Only the owner can add members');
  if (members.some(function(m){ return m.email === email; })) return;
  SpreadsheetApp.openById(SHEET_ID).getSheetByName(MEMBERS_SHEET).appendRow([email,'Member','Active']);
}

function clean_(value, max) { return String(value || '').replace(/[<>]/g,'').slice(0,max); }
function safeError_(err) { return String(err && err.message || 'Request failed').slice(0,180); }
function json_(payload, callback) {
  const text = JSON.stringify(payload);
  if (callback && /^[A-Za-z_$][\w.$]*$/.test(callback)) return ContentService.createTextOutput(callback+'('+text+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

function setupHeaders_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const tx = ss.getSheetByName(TX_SHEET);
  if (tx.getRange('A1:I1').getDisplayValues()[0].join('|') !== 'ID|Type|Quantity|Unit Price|Total|Financial Party|Note|Occurred At|Created By') tx.getRange('A1:I1').setValues([['ID','Type','Quantity','Unit Price','Total','Financial Party','Note','Occurred At','Created By']]);
}
function testRead_() { Logger.log(readMembers_()); Logger.log(readTransactions_()); }
function testAdd_() { addTransaction_({type:'ad',qty:1,price:0,note:'test'}, getUserEmail_()); }

function onOpen() { SpreadsheetApp.getUi().createMenu('Super Subs').addItem('تهيئة العناوين','setupHeaders_').addItem('اختبار القراءة','testRead_').addToUi(); }

// Deployment settings: Execute as user accessing the web app; access: anyone with Google account.
// Share the spreadsheet with every active member before testing.
// Web app URL is then configured in index.html as SUPER_SUBS_API_URL.
// Never deploy this script anonymously: membership checks rely on Google sign-in.
