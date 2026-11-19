// THEE Onset Business Unit – Google Sheet backend (use a NEW Google Sheet).
function doPost(e){
  var b=JSON.parse(e.postData.contents),ss=SpreadsheetApp.getActiveSpreadsheet();
  if(b.mode==='pull'){var s=ss.getSheetByName('Backup');return out({data:s&&s.getRange('A1').getValue()?JSON.parse(s.getRange('A1').getValue()):null});}
  var d=b.data,wkly={};
  put(ss,'Backup',[[JSON.stringify(d)]]);
  put(ss,'Records',[['Date','Type','Business','Category','Amount','Note']].concat(d.tx.map(function(t){return[t.date,t.type,t.hustle,t.cat,t.amt,t.note]})));
  put(ss,'Hustles',[['Name','Type','Income','Expenses','Profit']].concat([{name:'Main business'}].concat(d.hustles).map(function(h){var i=0,o=0;d.tx.forEach(function(t){if(t.hustle==h.name){t.type=='in'?i+=t.amt:o+=t.amt}});return[h.name,h.kind||'core',i,o,i-o]})));
  d.tx.forEach(function(t){var k=wk(t.date);wkly[k]=wkly[k]||[0,0];t.type=='in'?wkly[k][0]+=t.amt:wkly[k][1]+=t.amt});
  put(ss,'Weekly',[['Week starting','Income','Expenditure','Profit/Loss']].concat(Object.keys(wkly).sort().map(function(k){return[k,wkly[k][0],wkly[k][1],wkly[k][0]-wkly[k][1]]})));
  return out({ok:true});
}
function wk(s){var d=new Date(s+'T00:00:00');d.setDate(d.getDate()-(d.getDay()+6)%7);return Utilities.formatDate(d,'Africa/Nairobi','yyyy-MM-dd');}
function put(ss,n,r){var s=ss.getSheetByName(n)||ss.insertSheet(n);s.clear();s.getRange(1,1,r.length,r[0].length).setValues(r);}
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
