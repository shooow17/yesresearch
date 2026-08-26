/**
 * 代課系統 — GAS Web App API
 * ★ 部署在 Apps Script ★
 *
 * 審核流程：待對方確認 → 待組長審核 → 待主任核定 → 已核准
 *
 * _系統設定：
 *   DASHBOARD_EMAILS       課表概況權限
 *   EDIT_START / EDIT_END   導師拖曳調課開放區間
 *   SWAP_REVIEWER_EMAILS    教學組長 email（第二關）
 *   SWAP_APPROVER_EMAILS    教務主任 email（第三關）
 */

/* ---- 路由 ---- */
function doGet(e) {
  var action = (e.parameter.action || "").trim(), result;
  try {
    switch (action) {
      case "verify":       result = verifyLogin(e.parameter.email); break;
      case "schedule":     result = getScheduleForTeacher(e.parameter.tid, e.parameter.mode, e.parameter.cid||""); break;
      case "dashboard":    result = getDashboardTeachers(); break;
      case "swapTeachers": result = getSwapTeachers(); break;
      case "swapList":     result = getSwapList(e.parameter.email); break;
      case "masterData":   result = getMasterData(); break;
      default:             result = { error: "未知的 action: " + action };
    }
  } catch (err) { result = { error: err.toString() }; }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) {
  var result;
  try {
    var data = JSON.parse(e.postData.contents);
    switch (data.action) {
      case "swap":        result = swapHomeroomLessons(data.teacherId,data.classId,data.dayA,data.periodA,data.dayB,data.periodB); break;
      case "batchUpdate": result = batchUpdateLessons(data.teacherId,data.classId,data.changes); break;
      case "submitSwap":  result = submitSwapRequest(data); break;
      case "reviewSwap":  result = reviewSwapRequest(data); break;
      case "dashboardUpdate": result = dashboardBatchUpdate(data); break;
      case "crossSwapUpdate": result = crossBatchUpdate(data); break;
      case "crossBatchUpdate": result = crossBatchUpdate(data); break;
      default:            result = { error: "未知的 action" };
    }
  } catch (err) { result = { error: err.toString() }; }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

/* ================================================================
   工具
   ================================================================ */
function S(v){return v==null?"":String(v).trim();}
function ix(h,name,fb){for(var i=0;i<h.length;i++){if(S(h[i])===name)return i;}return fb;}
function checkSettingEmail(email,ss,key){
  var st=ss.getSheetByName("_系統設定"); if(!st)return false;
  var d=st.getDataRange().getValues();
  for(var i=1;i<d.length;i++){
    if(S(d[i][0])===key){var list=S(d[i][1]).toLowerCase().split(",");for(var j=0;j<list.length;j++){if(list[j].trim()===email)return true;}}
  } return false;
}
function buildAllMaps(ss){
  var teachers={},classes={},subjects={},classNameToId={};
  var cs=ss.getSheetByName("班級名單");
  if(cs&&cs.getLastRow()>1){var cd=cs.getDataRange().getValues(),ch=cd[0],cidx=ix(ch,"班級代碼",0),cnidx=ix(ch,"班級名稱",1);
    for(var i=1;i<cd.length;i++){var id=S(cd[i][cidx]),name=S(cd[i][cnidx]);if(id&&name){classes[id]={name:name};classNameToId[name]=id;}}}
  var ts=ss.getSheetByName("教師名單");
  if(ts&&ts.getLastRow()>1){var td=ts.getDataRange().getValues(),th=td[0],tidx=ix(th,"教師代碼",0),nidx=ix(th,"姓名",1),ridx=ix(th,"身份",2),cidx2=ix(th,"任教班級代碼",3),eidx=ix(th,"電子郵件",-1);
    for(var i=1;i<td.length;i++){var id=S(td[i][tidx]),name=S(td[i][nidx]),role=S(td[i][ridx]),clsRaw=S(td[i][cidx2]),email=eidx!==-1?S(td[i][eidx]):"";
      if(!id)continue;var classCode=classes[clsRaw]?clsRaw:(classNameToId[clsRaw]||"");
      teachers[id]={name:name,role:role,classCode:classCode,email:email};if(classCode&&classes[classCode])classes[classCode].homeroomId=id;}}
  var ss2=ss.getSheetByName("科目清單");
  if(ss2&&ss2.getLastRow()>1){var sd=ss2.getDataRange().getValues(),sh=sd[0],sidx=ix(sh,"科目代碼",0),snidx=ix(sh,"科目名稱",1);
    for(var i=1;i<sd.length;i++){var id=S(sd[i][sidx]);if(id)subjects[id]=S(sd[i][snidx]);}}
  return{teachers:teachers,classes:classes,subjects:subjects};
}
function getEditWindow(ss){
  var st=ss.getSheetByName("_系統設定"); if(!st)return{canEdit:true,start:"",end:"",message:""};
  var d=st.getDataRange().getValues(),startVal=null,endVal=null;
  for(var i=1;i<d.length;i++){if(S(d[i][0])==="EDIT_START")startVal=d[i][1];if(S(d[i][0])==="EDIT_END")endVal=d[i][1];}
  if(!startVal||!endVal)return{canEdit:true,start:"",end:"",message:""};
  var sd=startVal instanceof Date?startVal:new Date(S(startVal)),ed=endVal instanceof Date?endVal:new Date(S(endVal));
  if(isNaN(sd.getTime())||isNaN(ed.getTime()))return{canEdit:true,start:"",end:"",message:""};
  ed.setHours(23,59,59,999);var now=new Date(),tz=Session.getScriptTimeZone();var can=now>=sd&&now<=ed;
  var fs=Utilities.formatDate(sd,tz,"yyyy/MM/dd"),fe=Utilities.formatDate(ed,tz,"yyyy/MM/dd");
  return{canEdit:can,start:fs,end:fe,message:can?"開放調課中（"+fs+" ~ "+fe+"）":(now<sd?"調課尚未開放（"+fs+" 起）":"調課期間已結束（截止 "+fe+"）")};
}

/* ================================================================
   1. 登入驗證
   ================================================================ */
function verifyLogin(email){
  if(!email||!email.trim())return{error:"請輸入電子郵件"};
  email=email.trim().toLowerCase();
  var ss=SpreadsheetApp.getActiveSpreadsheet(),maps=buildAllMaps(ss);
  var teacher=null;
  for(var tid in maps.teachers){if((maps.teachers[tid].email||"").toLowerCase()===email){teacher=maps.teachers[tid];teacher.id=tid;break;}}
  if(!teacher)return{error:"找不到此電子郵件對應的教師"};
  var isHR=teacher.role.indexOf("導師")!==-1,hrCid="",hrCname="";
  if(isHR&&teacher.classCode){hrCid=teacher.classCode;hrCname=maps.classes[hrCid]?maps.classes[hrCid].name:hrCid;}
  var ys=ss.getSheetByName("學年度"),si={year:"",name:""};
  if(ys&&ys.getLastRow()>1){var y=ys.getDataRange().getValues();si={year:S(y[1][0]),name:S(y[1][1])};}
  return{success:true,
    teacher:{id:teacher.id,name:teacher.name,role:teacher.role},
    isHomeroom:isHR,isDashboard:checkSettingEmail(email,ss,"DASHBOARD_EMAILS"),
    isSwapReviewer:checkSettingEmail(email,ss,"SWAP_REVIEWER_EMAILS"),
    isSwapApprover:checkSettingEmail(email,ss,"SWAP_APPROVER_EMAILS"),
    homeroomClassId:hrCid,homeroomClassName:hrCname,
    schoolInfo:si,editWindow:getEditWindow(ss)};
}

/* ================================================================
   2-4. 課表查詢 / Dashboard / 導師拖曳（不變）
   ================================================================ */
function getScheduleForTeacher(teacherId,mode,classIdOverride){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),maps=buildAllMaps(ss);
  var db=ss.getSheetByName("課表資料庫"); if(!db||db.getLastRow()<=1)return{timetable:[]};
  var dd=db.getDataRange().getValues(),h=dd[0],ci=ix(h,"班級代碼",0),di=ix(h,"星期",1),pi=ix(h,"節次",2),si=ix(h,"科目代碼",3),ti=ix(h,"教師代碼",4);
  var tgt=classIdOverride||"";if(mode==="homeroom"&&!tgt){var t=maps.teachers[teacherId];if(t&&t.classCode)tgt=t.classCode;}
  var tt=[];
  for(var i=1;i<dd.length;i++){var r=dd[i],e={classId:S(r[ci]),day:S(r[di]),period:S(r[pi]),subId:S(r[si]),teaId:S(r[ti])};
    if(!e.classId||!e.day||!e.period)continue;
    e.className=maps.classes[e.classId]?maps.classes[e.classId].name:e.classId;e.subName=maps.subjects[e.subId]||e.subId;
    e.teaName=maps.teachers[e.teaId]?maps.teachers[e.teaId].name:e.teaId;e.teaRole=maps.teachers[e.teaId]?maps.teachers[e.teaId].role:"";
    if(mode==="homeroom"){if(e.classId===tgt)tt.push(e);}else{if(e.teaId===teacherId)tt.push(e);}}
  return{timetable:tt,homeroomClassId:tgt};
}
function getDashboardTeachers(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),maps=buildAllMaps(ss),db=ss.getSheetByName("課表資料庫"),cm={};
  if(db&&db.getLastRow()>1){var dd=db.getDataRange().getValues(),ti=ix(dd[0],"教師代碼",4);for(var i=1;i<dd.length;i++){var tid=S(dd[i][ti]);if(tid)cm[tid]=(cm[tid]||0)+1;}}
  var ro={"導師":1,"科任":2,"行政":3,"特教":4,"教支":5,"外籍":6},res=[];
  for(var tid in maps.teachers){var t=maps.teachers[tid],hc="";if(t.classCode&&maps.classes[t.classCode])hc=maps.classes[t.classCode].name;
    res.push({id:tid,name:t.name,role:t.role,className:hc,total:cm[tid]||0});}
  res.sort(function(a,b){var ra=ro[a.role]||99,rb=ro[b.role]||99;if(ra!==rb)return ra-rb;return a.name.localeCompare(b.name,'zh-Hant');});
  return{teachers:res};
}
function swapHomeroomLessons(teacherId,classId,dayA,periodA,dayB,periodB){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),ew=getEditWindow(ss);if(!ew.canEdit)return{error:ew.message};
  var maps=buildAllMaps(ss),tea=maps.teachers[teacherId];if(!tea||tea.classCode!==classId)return{error:"您不是此班級的導師"};
  var db=ss.getSheetByName("課表資料庫"),dd=db.getDataRange().getValues(),h=dd[0];
  var ci=ix(h,"班級代碼",0),di=ix(h,"星期",1),pi=ix(h,"節次",2),si=ix(h,"科目代碼",3),ti=ix(h,"教師代碼",4);
  var rA=-1,rB=-1;
  for(var i=1;i<dd.length;i++){if(S(dd[i][ci])===classId&&S(dd[i][ti])===teacherId){
    if(S(dd[i][di])===dayA&&S(dd[i][pi])===periodA)rA=i+1;if(S(dd[i][di])===dayB&&S(dd[i][pi])===periodB)rB=i+1;}}
  if(rA===-1||rB===-1)return{error:"找不到要交換的課程"};
  var sA=db.getRange(rA,si+1).getValue(),sB=db.getRange(rB,si+1).getValue();
  db.getRange(rA,si+1).setValue(sB);db.getRange(rB,si+1).setValue(sA);return{success:true};
}
function batchUpdateLessons(teacherId,classId,changes){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),ew=getEditWindow(ss);if(!ew.canEdit)return{error:ew.message};
  var maps=buildAllMaps(ss),tea=maps.teachers[teacherId];if(!tea||tea.classCode!==classId)return{error:"您不是此班級的導師"};
  if(!changes||changes.length===0)return{error:"沒有需要變更的項目"};
  var db=ss.getSheetByName("課表資料庫"),dd=db.getDataRange().getValues(),h=dd[0];
  var ci=ix(h,"班級代碼",0),di=ix(h,"星期",1),pi=ix(h,"節次",2),si=ix(h,"科目代碼",3),ti=ix(h,"教師代碼",4);
  for(var c=0;c<changes.length;c++){var ch=changes[c];for(var i=1;i<dd.length;i++){
    if(S(dd[i][ci])===classId&&S(dd[i][ti])===teacherId&&S(dd[i][di])===ch.day&&S(dd[i][pi])===ch.period){db.getRange(i+1,si+1).setValue(ch.subId);break;}}}
  return{success:true};
}

/* ================================================================
   5. 調課申請系統
   ================================================================ */

/** 教師清單（供選擇代課老師） */
function getSwapTeachers(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),maps=buildAllMaps(ss),res=[];
  for(var tid in maps.teachers){var t=maps.teachers[tid],hc="";if(t.classCode&&maps.classes[t.classCode])hc=maps.classes[t.classCode].name;
    res.push({id:tid,name:t.name,role:t.role,className:hc});}
  res.sort(function(a,b){return a.name.localeCompare(b.name,'zh-Hant');});return{teachers:res};
}

/** 確保「調課紀錄」工作表存在
 *  欄位：0序號 1申請時間 2申請人代碼 3申請人姓名
 *  4自己日期 5自己星期 6自己節次 7自己科目 8自己班級
 *  9對方代碼 10對方姓名 11對方日期 12對方星期 13對方節次 14對方科目 15對方班級
 *  16調課原因 17已告知對方 18狀態
 *  19對方確認時間 20組長審核人 21組長審核時間 22主任審核人 23主任審核時間
 */
function ensureSwapSheet(ss){
  var sheet=ss.getSheetByName("調課紀錄"); if(sheet)return sheet;
  sheet=ss.insertSheet("調課紀錄");
  var headers=["序號","申請時間","申請人代碼","申請人姓名",
    "自己日期","自己星期","自己節次","自己科目","自己班級",
    "對方代碼","對方姓名","對方日期","對方星期","對方節次","對方科目","對方班級",
    "調課原因","已告知對方","狀態",
    "對方確認時間","組長審核人","組長審核時間","主任審核人","主任審核時間"];
  sheet.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight("bold").setBackground("#1e3a5f").setFontColor("#fff");
  sheet.setFrozenRows(1);return sheet;
}

/** 送出調課申請 → 初始狀態「待對方確認」 */
function submitSwapRequest(data){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),maps=buildAllMaps(ss);
  var teaA=maps.teachers[data.tidA]; if(!teaA)return{error:"找不到申請人"};
  var teaB=maps.teachers[data.tidB]; if(!teaB)return{error:"找不到代課老師"};
  if(data.tidA===data.tidB)return{error:"不能跟自己調課"};

  // 衝堂檢查
  var db=ss.getSheetByName("課表資料庫");
  var warnings=[];
  if(db&&db.getLastRow()>1){
    var dd=db.getDataRange().getValues(),h=dd[0],di=ix(h,"星期",1),pi=ix(h,"節次",2),ti=ix(h,"教師代碼",4);
    for(var i=1;i<dd.length;i++){
      if(S(dd[i][ti])===data.tidB&&S(dd[i][di])===data.weekdayA&&S(dd[i][pi])===data.periodA)
        {warnings.push(teaB.name+" 在"+data.weekdayA+"第"+data.periodA+"節已有課程");break;}
    }
    if(data.periodB&&data.weekdayB){
      for(var i=1;i<dd.length;i++){
        if(S(dd[i][ti])===data.tidA&&S(dd[i][di])===data.weekdayB&&S(dd[i][pi])===data.periodB){
          warnings.push(teaA.name+" 在"+data.weekdayB+"第"+data.periodB+"節已有課程");break;}
      }
    }
  }

  var sheet=ensureSwapSheet(ss);
  var seqNo="SW"+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyyMMddHHmmss");
  var now=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy/MM/dd HH:mm");
  var row=[seqNo,now,data.tidA,teaA.name,
    data.dateA||"",data.weekdayA,data.periodA,data.subNameA||"",data.classNameA||"",
    data.tidB,teaB.name,data.dateB||"",data.weekdayB||"",data.periodB||"",data.subNameB||"",data.classNameB||"",
    data.reason||"",data.notified?"是":"否","待對方確認",
    "","","","",""];
  sheet.getRange(sheet.getLastRow()+1,1,1,row.length).setValues([row]);
  return{success:true,seqNo:seqNo,warnings:warnings};
}

/** 取得調課紀錄 */
function getSwapList(email){
  if(!email)return{error:"缺少 email"};
  email=email.trim().toLowerCase();
  var ss=SpreadsheetApp.getActiveSpreadsheet(),maps=buildAllMaps(ss);
  var myTid="";
  for(var tid in maps.teachers){if((maps.teachers[tid].email||"").toLowerCase()===email){myTid=tid;break;}}
  var isRev=checkSettingEmail(email,ss,"SWAP_REVIEWER_EMAILS");
  var isApp=checkSettingEmail(email,ss,"SWAP_APPROVER_EMAILS");

  var sheet=ss.getSheetByName("調課紀錄");
  if(!sheet||sheet.getLastRow()<=1)return{records:[],isReviewer:isRev,isApprover:isApp,myTid:myTid};
  var data=sheet.getDataRange().getValues(),records=[];
  for(var i=1;i<data.length;i++){
    var r=data[i];
    var rec={row:i+1,seqNo:S(r[0]),time:S(r[1]),
      tidA:S(r[2]),nameA:S(r[3]),
      dateA:S(r[4]),weekdayA:S(r[5]),periodA:S(r[6]),subA:S(r[7]),classA:S(r[8]),
      tidB:S(r[9]),nameB:S(r[10]),
      dateB:S(r[11]),weekdayB:S(r[12]),periodB:S(r[13]),subB:S(r[14]),classB:S(r[15]),
      reason:S(r[16]),notified:S(r[17]),status:S(r[18]),
      partnerTime:S(r[19]),revName:S(r[20]),revTime:S(r[21]),appName:S(r[22]),appTime:S(r[23])};
    // 可見：自己申請的 / 自己是代課老師 / 審核者
    if(rec.tidA===myTid||rec.tidB===myTid||isRev||isApp) records.push(rec);
  }
  return{records:records,isReviewer:isRev,isApprover:isApp,myTid:myTid};
}

/** 審核/確認調課申請
 *  流程：待對方確認 → 待組長審核 → 待主任核定 → 已核准
 *  任何階段可駁回 */
function reviewSwapRequest(data){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var email=(data.reviewerEmail||"").trim().toLowerCase();
  var maps=buildAllMaps(ss);
  var revName=email,revTid="";
  for(var tid in maps.teachers){if((maps.teachers[tid].email||"").toLowerCase()===email){revName=maps.teachers[tid].name;revTid=tid;break;}}
  var isRev=checkSettingEmail(email,ss,"SWAP_REVIEWER_EMAILS");
  var isApp=checkSettingEmail(email,ss,"SWAP_APPROVER_EMAILS");

  var sheet=ss.getSheetByName("調課紀錄"); if(!sheet)return{error:"找不到調課紀錄"};
  var row=data.row,rd=sheet.getRange(row,1,1,24).getValues()[0];
  var st=S(rd[18]),tidB=S(rd[9]);
  var now=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy/MM/dd HH:mm");

  // ---- 駁回（任何有權限的階段） ----
  if(data.decision==="reject"){
    sheet.getRange(row,19).setValue("已駁回");
    if(st==="待對方確認"&&revTid===tidB) sheet.getRange(row,20).setValue(now);
    else if(st==="待組長審核"&&isRev){sheet.getRange(row,21).setValue(revName);sheet.getRange(row,22).setValue(now);}
    else if(st==="待主任核定"&&isApp){sheet.getRange(row,23).setValue(revName);sheet.getRange(row,24).setValue(now);}
    else return{error:"您沒有權限駁回此階段的申請"};
    return{success:true,newStatus:"已駁回"};
  }

  // ---- 核准 ----
  if(data.decision==="approve"){
    // 第一關：代課老師確認
    if(st==="待對方確認"&&revTid===tidB){
      sheet.getRange(row,19).setValue("待組長審核");
      sheet.getRange(row,20).setValue(now);
      return{success:true,newStatus:"待組長審核"};
    }
    // 第二關：教學組長
    if(st==="待組長審核"&&isRev){
      sheet.getRange(row,19).setValue("待主任核定");
      sheet.getRange(row,21).setValue(revName);sheet.getRange(row,22).setValue(now);
      return{success:true,newStatus:"待主任核定"};
    }
    // 第三關：教務主任
    if(st==="待主任核定"&&isApp){
      sheet.getRange(row,19).setValue("已核准");
      sheet.getRange(row,23).setValue(revName);sheet.getRange(row,24).setValue(now);
      return{success:true,newStatus:"已核准"};
    }
    return{error:"您沒有權限核准此階段的申請"};
  }
  return{error:"未知的審核動作"};
}

/* ================================================================
   6. Dashboard 課表調整（教學組/主任用）
   ================================================================ */
function dashboardBatchUpdate(data){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var email=(data.email||"").trim().toLowerCase();
  if(!checkSettingEmail(email,ss,"DASHBOARD_EMAILS")) return{error:"您沒有管理權限"};
  var classId=data.classId; if(!classId) return{error:"缺少班級代碼"};
  var changes=data.changes; if(!changes||changes.length===0) return{error:"沒有需要變更的項目"};

  var db=ss.getSheetByName("課表資料庫"); if(!db||db.getLastRow()<=1) return{error:"找不到課表資料庫"};
  var dd=db.getDataRange().getValues(),h=dd[0];
  var ci=ix(h,"班級代碼",0),di=ix(h,"星期",1),pi=ix(h,"節次",2),si=ix(h,"科目代碼",3),ti=ix(h,"教師代碼",4);

  for(var c=0;c<changes.length;c++){
    var ch=changes[c];
    for(var i=1;i<dd.length;i++){
      if(S(dd[i][ci])===classId&&S(dd[i][di])===ch.day&&S(dd[i][pi])===ch.period){
        db.getRange(i+1,si+1).setValue(ch.subId);
        db.getRange(i+1,ti+1).setValue(ch.teaId);
        break;
      }
    }
  }
  return{success:true};
}

/* ================================================================
   7. 基礎資料（班級清單、科目清單）
   ================================================================ */
function getMasterData(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),maps=buildAllMaps(ss);
  var classes=[],subjects=[];
  for(var id in maps.classes) classes.push({id:id,name:maps.classes[id].name});
  for(var id in maps.subjects) subjects.push({id:id,name:maps.subjects[id]});
  classes.sort(function(a,b){
    var ga=parseInt(a.name)||99,gb=parseInt(b.name)||99;if(ga!==gb)return ga-gb;
    var na=parseInt(a.name.replace(/.*年/,''))||0,nb=parseInt(b.name.replace(/.*年/,''))||0;return na-nb;
  });
  subjects.sort(function(a,b){return a.name.localeCompare(b.name,'zh-Hant');});
  return{classes:classes,subjects:subjects};
}

/* ================================================================
   8. 跨師調課（交換＋新增＋刪除）
   changes 陣列，每筆有 type: "swap"|"add"|"delete"
   ================================================================ */
function crossBatchUpdate(data){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var email=(data.email||"").trim().toLowerCase();
  if(!checkSettingEmail(email,ss,"DASHBOARD_EMAILS")) return{error:"您沒有管理權限"};
  var changes=data.changes; if(!changes||changes.length===0) return{error:"沒有需要變更的項目"};

  var db=ss.getSheetByName("課表資料庫"); if(!db||db.getLastRow()<=1) return{error:"找不到課表資料庫"};
  var dd=db.getDataRange().getValues(),h=dd[0];
  var ci=ix(h,"班級代碼",0),di=ix(h,"星期",1),pi=ix(h,"節次",2),si=ix(h,"科目代碼",3),ti=ix(h,"教師代碼",4);
  var colCount=h.length;

  var swapped=0,added=0,deleted=0;
  // 先處理刪除（從後往前避免行號偏移）
  var delRows=[];
  for(var c=0;c<changes.length;c++){
    var ch=changes[c];
    if(ch.type!=='delete') continue;
    for(var i=1;i<dd.length;i++){
      if(S(dd[i][ci])===ch.classId&&S(dd[i][di])===ch.day&&S(dd[i][pi])===ch.period&&S(dd[i][ti])===ch.teaId&&S(dd[i][si])===ch.subId){
        delRows.push(i+1);break;
      }
    }
  }
  delRows.sort(function(a,b){return b-a;}); // 從後往前刪
  for(var d2=0;d2<delRows.length;d2++){db.deleteRow(delRows[d2]);deleted++;}

  // 重新讀取（刪除後行號變了）
  if(deleted>0) dd=db.getDataRange().getValues();

  // 處理交換
  for(var c=0;c<changes.length;c++){
    var ch=changes[c];
    if(ch.type!=='swap') continue;
    // A 邊：classIdA+dayA+periodA+oldTeaIdA → 改成 oldTeaIdB
    for(var i=1;i<dd.length;i++){
      if(S(dd[i][ci])===ch.classIdA&&S(dd[i][di])===ch.dayA&&S(dd[i][pi])===ch.periodA&&S(dd[i][ti])===ch.oldTeaIdA){
        db.getRange(i+1,ti+1).setValue(ch.oldTeaIdB);dd[i][ti]=ch.oldTeaIdB;swapped++;break;
      }
    }
    // B 邊：classIdB+dayB+periodB+oldTeaIdB → 改成 oldTeaIdA
    for(var i=1;i<dd.length;i++){
      if(S(dd[i][ci])===ch.classIdB&&S(dd[i][di])===ch.dayB&&S(dd[i][pi])===ch.periodB&&S(dd[i][ti])===ch.oldTeaIdB){
        db.getRange(i+1,ti+1).setValue(ch.oldTeaIdA);dd[i][ti]=ch.oldTeaIdA;swapped++;break;
      }
    }
  }

  // 處理新增
  for(var c=0;c<changes.length;c++){
    var ch=changes[c];
    if(ch.type!=='add') continue;
    var newRow=[];
    for(var j=0;j<colCount;j++) newRow.push("");
    newRow[ci]=ch.classId;newRow[di]=ch.day;newRow[pi]=ch.period;newRow[si]=ch.subId;newRow[ti]=ch.teaId;
    db.appendRow(newRow);added++;
  }

  return{success:true,swapped:swapped,added:added,deleted:deleted};
}