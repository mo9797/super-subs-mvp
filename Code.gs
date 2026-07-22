const SHEET_ID = '1-tema9OUKu0j2wner8aGvy3mUX1qhswhCRNW4jHfgqI';
const TX_SHEET = 'Transactions';
const MEMBERS_SHEET = 'Members';
const VALID_PURCHASE_PARTIES = ['حمدي','علي'];
const VALID_SALE_PARTIES = ['محفظة حمدي','محفظة علي'];

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
    else if (body.action === 'update') updateTransaction_(body, email);
    else if (body.action === 'delete') deleteTransaction_(body.id, email);
    else if (body.action === 'member') addMember_(body.email, email);
    else if (body.action === 'list') return json_({ok:true, transactions:readTransactions_(), members:readMembers_(), user:email});
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

function normalizeTransaction_(body) {
  const type = String(body.type || '');
  if (['purchase','sale','ad'].indexOf(type) === -1) throw new Error('Invalid transaction type');
  const qty = Number(body.qty);
  const price = Number(body.price);
  if (!isFinite(qty) || qty <= 0 || !isFinite(price) || price < 0) throw new Error('Invalid amount');
  const party = validateParty_(type, body.party);
  return {type:type, qty:type === 'ad' ? 1 : qty, price:price, total:type === 'ad' ? price : qty * price, party:party, note:clean_(body.note,120)};
}

function validateParty_(type, rawParty) {
  const party = clean_(rawParty, 40);
  if (type === 'purchase' && VALID_PURCHASE_PARTIES.indexOf(party) === -1) throw new Error('Invalid purchase party');
  if (type === 'sale' && VALID_SALE_PARTIES.indexOf(party) === -1) throw new Error('Invalid sale wallet');
  return type === 'ad' ? '' : party;
}

function addTransaction_(body, email) {
  const tx = normalizeTransaction_(body);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    SpreadsheetApp.openById(SHEET_ID).getSheetByName(TX_SHEET).appendRow([Utilities.getUuid(), tx.type, tx.qty, tx.price, tx.total, tx.party, tx.note, new Date(), email]);
  } finally { lock.releaseLock(); }
}

function updateTransaction_(body, email) {
  const id = clean_(body.id, 80);
  if (!id) throw new Error('Transaction ID is required');
  const tx = normalizeTransaction_(body);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TX_SHEET);
    const row = findTransactionRow_(sheet, id);
    const occurredAt = sheet.getRange(row, 8).getValue() || new Date();
    sheet.getRange(row, 1, 1, 9).setValues([[id, tx.type, tx.qty, tx.price, tx.total, tx.party, tx.note, occurredAt, email]]);
  } finally { lock.releaseLock(); }
}

function deleteTransaction_(rawId, email) {
  const id = clean_(rawId, 80);
  if (!id) throw new Error('Transaction ID is required');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TX_SHEET);
    sheet.deleteRow(findTransactionRow_(sheet, id));
  } finally { lock.releaseLock(); }
}

function findTransactionRow_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('Transaction not found');
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  throw new Error('Transaction not found');
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
