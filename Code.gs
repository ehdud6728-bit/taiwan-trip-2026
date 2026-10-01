const SPREADSHEET_ID = "1ASQQ3T5BMuxPRzBsvU3dE5ud-hJiZRmN2hHct3hiHSk";
const RATE_KRW_PER_TWD = 42.78;

function ss_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function out_(obj, cb) {
  const body = JSON.stringify(obj);
  if (cb && /^[A-Za-z_$][0-9A-Za-z_$.]*$/.test(cb)) {
    return ContentService.createTextOutput(cb+"("+body+");").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}
function fmt_(v) {
  return Object.prototype.toString.call(v)==="[object Date]"
    ? Utilities.formatDate(v,"Asia/Taipei","yyyy-MM-dd HH:mm:ss") : v;
}
function doGet(e) {
  const ss=ss_(), cb=e.parameter.callback;
  const c=ss.getSheetByName("준비물").getDataRange().getValues();
  const t=ss.getSheetByName("열차").getDataRange().getValues();
  const x=ss.getSheetByName("경비").getDataRange().getValues();
  return out_({ok:true,
    checklist:c.slice(1).filter(r=>r[0]!=="").map((r,i)=>({row:i+2,item:fmt_(r[0]),owner:fmt_(r[1]),done:r[2]===true,note:fmt_(r[3])})),
    trains:t.slice(1).filter(r=>r[0]!=="").map((r,i)=>({row:i+2,date:fmt_(r[0]),route:fmt_(r[1]),departure:fmt_(r[2]),arrival:fmt_(r[3]),trainNo:fmt_(r[4]),status:fmt_(r[5]),note:fmt_(r[6])})),
    expenses:x.slice(1).filter(r=>r[0]!=="").map((r,i)=>({row:i+2,time:fmt_(r[0]),date:fmt_(r[1]),cat:fmt_(r[2]),cur:fmt_(r[3]),raw:Number(r[4])||0,pay:fmt_(r[5]),memo:fmt_(r[6]),rate:Number(r[7])||RATE_KRW_PER_TWD,krw:Number(r[8])||0,inputter:fmt_(r[9])})).reverse()
  },cb);
}
function doPost(e) {
  try {
    const p=e.parameter||{}, a=p.action, ss=ss_();
    if(a==="setChecklist") ss.getSheetByName("준비물").getRange(Number(p.row),3).setValue(String(p.done)==="true");
    else if(a==="addChecklist") ss.getSheetByName("준비물").appendRow([p.item||"",p.owner||"",false,p.note||""]);
    else if(a==="deleteChecklist") ss.getSheetByName("준비물").deleteRow(Number(p.row));
    else if(a==="updateTrain") ss.getSheetByName("열차").getRange(Number(p.row),3,1,5).setValues([[p.departure||"",p.arrival||"",p.trainNo||"",p.status||"",p.note||""]]);
    else if(a==="addExpense") {
      const raw=Number(p.amount)||0, cur=p.cur||"TWD", rate=Number(p.rate)||RATE_KRW_PER_TWD;
      const krw=cur==="KRW"?raw:Math.round(raw*rate);
      ss.getSheetByName("경비").appendRow([Utilities.formatDate(new Date(),"Asia/Taipei","yyyy-MM-dd HH:mm:ss"),p.date||"",p.cat||"기타",cur,raw,p.pay||"",p.memo||"",rate,krw,p.inputter||""]);
    }
    else if(a==="deleteExpense") ss.getSheetByName("경비").deleteRow(Number(p.row));
    else return out_({ok:false,error:"unknown action"});
    return out_({ok:true});
  } catch(err) { return out_({ok:false,error:String(err)}); }
}