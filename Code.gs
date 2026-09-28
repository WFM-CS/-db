// Live Google Sheets backend for Advisor Productivity Dashboard.
// Bind this Apps Script to the SAME Google Spreadsheet that receives your form data.
const SHEET_NAME = 'Data';
const API_SECRET = ''; // Optional shared token. Leave blank unless you want one.

function doGet(e) {
  try {
    if (!checkToken_(e)) return json_({ok:false,error:'Unauthorized'});
    if ((e.parameter.action || 'read').toLowerCase() !== 'read') return json_({ok:false,error:'Unknown action'});
    const sh = getSheet_();
    const range = sh.getDataRange();
    const values = range.getValues();
    if (!values.length) return json_({ok:true,headers:[],rows:[]});
    const headers = values.shift().map(String);
    const rows = values.map(row => {
      const obj = {};
      headers.forEach((h,i) => {
        const v = row[i];
        obj[h] = v instanceof Date
          ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
          : v;
      });
      return obj;
    });
    return json_({ok:true,headers,rows,refreshedAt:new Date().toISOString(),sheet:SHEET_NAME});
  } catch (err) {
    return json_({ok:false,error:String(err)});
  }
}

function doPost(e) {
  try {
    if (!checkToken_(e)) return json_({ok:false,error:'Unauthorized'});
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    if ((body.action || '') !== 'append') return json_({ok:false,error:'Unknown action'});
    const rows = body.rows || [];
    if (!rows.length) return json_({ok:true,added:0});
    const sh = getSheet_();
    const incomingHeaders = Object.keys(rows[0]);
    if (sh.getLastRow() === 0) sh.getRange(1,1,1,incomingHeaders.length).setValues([incomingHeaders]);
    const existingHeaders = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
    const output = rows.map(obj => existingHeaders.map(h => cleanForSheet_(obj[h], h)));
    sh.getRange(sh.getLastRow()+1,1,output.length,existingHeaders.length).setValues(output);
    return json_({ok:true,added:output.length});
  } catch (err) {
    return json_({ok:false,error:String(err)});
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  return sh;
}
function cleanForSheet_(v, header) {
  if (header === 'Ticket Number') return String(v || '').replace(/\D+/g,'');
  return v == null ? '' : v;
}
function checkToken_(e) {
  if (!API_SECRET) return true;
  return ((e && e.parameter && e.parameter.token) || '') === API_SECRET;
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
