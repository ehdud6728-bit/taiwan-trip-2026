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
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
  }
  return v;
}
function fmtDate_(v) {
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, "Asia/Taipei", "yyyy-MM-dd");
  }
  return String(v || "");
}
function fmtTime_(v) {
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, "Asia/Taipei", "HH:mm");
  }
  const s = String(v || "").trim();
  const m = s.match(/(\d{1,2}):(\d{2})/);
  return m ? String(m[1]).padStart(2, "0") + ":" + m[2] : s;
}


const SHARED_CACHE_KEY = "TAIWAN_TRIP_SHARED_V1";
const SHARED_CACHE_SECONDS = 120;

function buildSharedPayload_() {
  const ss=ss_();
  const c=ss.getSheetByName("준비물").getDataRange().getValues();
  const t=ss.getSheetByName("열차").getDataRange().getValues();
  const x=ss.getSheetByName("경비").getDataRange().getValues();
  return {
    checklist:c.slice(1).filter(r=>r[0]!=="").map((r,i)=>({
      row:i+2,item:fmt_(r[0]),owner:fmt_(r[1]),done:r[2]===true,note:fmt_(r[3])
    })),
    trains:t.slice(1).filter(r=>r[0]!=="").map((r,i)=>({
      row:i+2,date:fmtDate_(r[0]),route:fmt_(r[1]),departure:fmtTime_(r[2]),
      arrival:fmtTime_(r[3]),trainNo:fmt_(r[4]),status:fmt_(r[5]),note:fmt_(r[6])
    })),
    expenses:x.slice(1).filter(r=>r[0]!=="").map((r,i)=>({
      row:i+2,time:fmt_(r[0]),date:fmt_(r[1]),cat:fmt_(r[2]),cur:fmt_(r[3]),
      raw:Number(r[4])||0,pay:fmt_(r[5]),memo:fmt_(r[6]),
      rate:Number(r[7])||RATE_KRW_PER_TWD,krw:Number(r[8])||0,inputter:fmt_(r[9])
    })).reverse()
  };
}
function sharedPayloadCached_() {
  const cache=CacheService.getScriptCache();
  const raw=cache.get(SHARED_CACHE_KEY);
  if(raw){
    try { return JSON.parse(raw); } catch(e) {}
  }
  const payload=buildSharedPayload_();
  try { cache.put(SHARED_CACHE_KEY,JSON.stringify(payload),SHARED_CACHE_SECONDS); } catch(e) {}
  return payload;
}
function clearSharedCache_() {
  try { CacheService.getScriptCache().remove(SHARED_CACHE_KEY); } catch(e) {}
}


const FOOD_NICKNAMES_PROP = "FOOD_NICKNAMES_V1";
const FOOD_NICKNAME_POOL = [
  "배고픈오리","곰탕재료푸우","망고에진심인곰","우육면수호대","샤오롱바오헌터",
  "버블티중독자","요우티아오요정","마라훠궈탐정","거위고기감별사","새우잡는펭귄",
  "타이루거두더지","칠성탄갈매기","리위탄잉어","화롄먹보","타이베이너구리",
  "자전거탄고양이","딤섬도둑","계란파전마왕","짠두유전도사","망고빙수기사",
  "야시장방랑자","마가오너구리","에그타르트도깨비","볶음밥용사","훠궈잠수부",
  "굴전수집가","땅콩경단요정","포자먹는판다","우육면길잡이","버섯꼬지대장",
  "황금조개탐험대","야시장청소부","밀크티순찰대","샤오롱바오지킴이","망고빙수도둑",
  "딤섬쫓는오리","두유마시는곰","새우튀김사냥꾼","거위뒤쫓는푸우","타이베이먹깨비"
];

function foodNickMap_() {
  try {
    return JSON.parse(PropertiesService.getScriptProperties().getProperty(FOOD_NICKNAMES_PROP) || "{}");
  } catch(e) {
    return {};
  }
}
function saveFoodNickMap_(m) {
  PropertiesService.getScriptProperties().setProperty(FOOD_NICKNAMES_PROP, JSON.stringify(m || {}));
}
function assignFoodNickname_(voterId) {
  voterId=String(voterId||"").trim().slice(0,80);
  if(!voterId) throw new Error("기기 ID가 없습니다.");

  const lock=LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const map=foodNickMap_();
    if(map[voterId]) return String(map[voterId]);

    const used=new Set(Object.values(map).map(String));
    let available=FOOD_NICKNAME_POOL.filter(n=>!used.has(n));

    // Pool exhausted: append a short number while still avoiding duplicates.
    if(!available.length){
      for(let i=1;i<=999;i++){
        const candidate="배고픈여행자"+i;
        if(!used.has(candidate)){ available=[candidate]; break; }
      }
    }

    const idx=Math.floor(Math.random()*available.length);
    const name=available[idx];
    map[voterId]=name;
    saveFoodNickMap_(map);
    return name;
  } finally {
    lock.releaseLock();
  }
}
function rerollFoodNickname_(voterId) {
  voterId=String(voterId||"").trim().slice(0,80);
  if(!voterId) throw new Error("기기 ID가 없습니다.");

  const lock=LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const map=foodNickMap_();
    const old=String(map[voterId]||"");
    const used=new Set(Object.entries(map).filter(([id])=>id!==voterId).map(([,n])=>String(n)));
    let available=FOOD_NICKNAME_POOL.filter(n=>!used.has(n) && n!==old);

    if(!available.length){
      available=FOOD_NICKNAME_POOL.filter(n=>!used.has(n));
    }
    if(!available.length){
      for(let i=1;i<=999;i++){
        const candidate="배고픈여행자"+i;
        if(!used.has(candidate) && candidate!==old){ available=[candidate]; break; }
      }
    }

    const idx=Math.floor(Math.random()*available.length);
    const name=available[idx] || old || "배고픈여행자";
    map[voterId]=name;
    saveFoodNickMap_(map);

    // Existing votes should display the new nickname too.
    renameFoodVoter_(voterId,name);
    return name;
  } finally {
    lock.releaseLock();
  }
}

const FOOD_VOTES_PROP = "FOOD_VOTES_V1";

function foodVotes_() {
  try {
    return JSON.parse(PropertiesService.getScriptProperties().getProperty(FOOD_VOTES_PROP) || "{}");
  } catch(e) {
    return {};
  }
}
function saveFoodVotes_(v) {
  PropertiesService.getScriptProperties().setProperty(FOOD_VOTES_PROP, JSON.stringify(v || {}));
}
function cleanName_(s) {
  return String(s || "").trim().replace(/\s+/g," ").slice(0,20);
}
function foodVotePayload_() {
  const src=foodVotes_(), out={};
  Object.keys(src).forEach(k=>{
    const arr=Array.isArray(src[k])?src[k]:[];
    out[k]=arr.map(v=>({id:String(v.id||""),name:cleanName_(v.name)})).filter(v=>v.id&&v.name);
  });
  return out;
}
function toggleFoodVote_(foodId,voterId,voterName) {
  foodId=String(foodId||"").trim();
  voterId=String(voterId||"").trim().slice(0,80);
  voterName=cleanName_(voterName);
  if(!foodId || !voterId || !voterName) throw new Error("투표 정보가 부족합니다.");
  const lock=LockService.getScriptLock();
  lock.waitLock(5000);
  try{
    const all=foodVotes_(), arr=Array.isArray(all[foodId])?all[foodId]:[];
    const idx=arr.findIndex(v=>String(v.id)===voterId);
    let selected=false;
    if(idx>=0){
      arr.splice(idx,1);
    }else{
      arr.push({id:voterId,name:voterName});
      selected=true;
    }
    all[foodId]=arr;
    saveFoodVotes_(all);
    return {selected:selected,votes:arr.length};
  }finally{
    lock.releaseLock();
  }
}
function renameFoodVoter_(voterId,voterName) {
  voterId=String(voterId||"").trim().slice(0,80);
  voterName=cleanName_(voterName);
  if(!voterId || !voterName) throw new Error("이름 정보가 부족합니다.");
  const lock=LockService.getScriptLock();
  lock.waitLock(5000);
  try{
    const all=foodVotes_();
    Object.keys(all).forEach(k=>{
      const arr=Array.isArray(all[k])?all[k]:[];
      arr.forEach(v=>{ if(String(v.id)===voterId) v.name=voterName; });
      all[k]=arr;
    });
    saveFoodVotes_(all);
  }finally{
    lock.releaseLock();
  }
}

function doGet(e) {
  const cb=e.parameter.callback, action=String(e.parameter.action||"");
  if(action==="foodVotes") return out_({ok:true,foodVotes:foodVotePayload_()},cb);
  if(action==="foodNickname") {
    try { return out_({ok:true,nickname:assignFoodNickname_(e.parameter.voterId)},cb); }
    catch(err){ return out_({ok:false,error:String(err)},cb); }
  }

  const shared=sharedPayloadCached_();
  return out_({
    ok:true,
    checklist:shared.checklist||[],
    trains:shared.trains||[],
    expenses:shared.expenses||[]
  },cb);
}
function doPost(e) {
  try {
    const p=e.parameter||{}, a=p.action;
    if(a==="toggleFoodVote") {
      const result=toggleFoodVote_(p.foodId,p.voterId,p.voterName);
      return out_({ok:true,result:result});
    }
    if(a==="renameFoodVoter") {
      renameFoodVoter_(p.voterId,p.voterName);
      return out_({ok:true});
    }
    if(a==="rerollFoodNickname") {
      const nickname=rerollFoodNickname_(p.voterId);
      return out_({ok:true,nickname:nickname});
    }
    const ss=ss_();
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
    clearSharedCache_();
    return out_({ok:true});
  } catch(err) { return out_({ok:false,error:String(err)}); }
}