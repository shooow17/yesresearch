/**
 * 代課系統 — GAS API（供 GitHub Pages 前端呼叫）
 *
 * 部署設定：
 *   執行身分：我（開發者）
 *   誰可以存取：所有人（含匿名）← 這樣 GitHub Pages 才能跨域呼叫
 *
 * 部署完成後，把 Web App URL 貼到前端 index.html 的 API_URL 變數
 */

function doGet(e) {
  var action = (e.parameter.action || "").trim();
  var result;
  try {
    switch (action) {
      case "verify":
        result = verifyLogin(e.parameter.email);
        break;
      case "schedule":
        result = getScheduleForTeacher(e.parameter.tid, e.parameter.mode, e.parameter.cid || "");
        break;
      case "dashboard":
        result = getDashboardTeachers();
        break;
      default:
        result = { error: "未知的 action: " + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var result;
  try {
    var data = JSON.parse(e.postData.contents);
    switch (data.action) {
      case "swap":
        result = swapHomeroomLessons(data.teacherId, data.classId, data.dayA, data.periodA, data.dayB, data.periodB);
        break;
      default:
        result = { error: "未知的 action" };
    }
  } catch (err) {
    result = { error: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

/* ================================================================
   1. 登入驗證
   ================================================================ */
function verifyLogin(email) {
  if (!email || !email.trim()) return { error: "請輸入電子郵件" };
  email = email.trim().toLowerCase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var maps = buildAllMaps(ss);

  var teacher = null;
  for (var tid in maps.teachers) {
    if ((maps.teachers[tid].email || "").toLowerCase() === email) {
      teacher = maps.teachers[tid]; teacher.id = tid; break;
    }
  }
  if (!teacher) return { error: "找不到此電子郵件對應的教師" };

  var isHR = teacher.role.indexOf("導師") !== -1;
  var hrClassId = "", hrClassName = "";
  if (isHR && teacher.classCode) {
    hrClassId = teacher.classCode;
    hrClassName = maps.classes[hrClassId] ? maps.classes[hrClassId].name : hrClassId;
  }

  var yearSheet = ss.getSheetByName("學年度");
  var schoolInfo = { year: "", name: "" };
  if (yearSheet && yearSheet.getLastRow() > 1) { var y = yearSheet.getDataRange().getValues(); schoolInfo = { year: S(y[1][0]), name: S(y[1][1]) }; }

  return {
    success: true,
    teacher: { id: teacher.id, name: teacher.name, role: teacher.role },
    isHomeroom: isHR,
    isDashboard: checkDashboardAccess(email, ss),
    homeroomClassId: hrClassId,
    homeroomClassName: hrClassName,
    schoolInfo: schoolInfo
  };
}

function checkDashboardAccess(email, ss) {
  var st = ss.getSheetByName("_系統設定");
  if (!st) return false;
  var d = st.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (d[i][0] === "DASHBOARD_EMAILS") {
      var list = S(d[i][1]).toLowerCase().split(",");
      for (var j = 0; j < list.length; j++) { if (list[j].trim() === email) return true; }
    }
  }
  return false;
}

/* ================================================================
   2. 取得課表
   ================================================================ */
function getScheduleForTeacher(teacherId, mode, classIdOverride) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var maps = buildAllMaps(ss);
  var dbSheet = ss.getSheetByName("課表資料庫");
  if (!dbSheet || dbSheet.getLastRow() <= 1) return { timetable: [] };

  var dbData = dbSheet.getDataRange().getValues();
  var h = dbData[0];
  var ci=ix(h,"班級代碼",0),di=ix(h,"星期",1),pi=ix(h,"節次",2),si=ix(h,"科目代碼",3),ti=ix(h,"教師代碼",4),ri=ix(h,"指定教室",6);

  var targetClassId = classIdOverride || "";
  if (mode === "homeroom" && !targetClassId) {
    var tea = maps.teachers[teacherId];
    if (tea && tea.classCode) targetClassId = tea.classCode;
  }

  var timetable = [];
  for (var i = 1; i < dbData.length; i++) {
    var row = dbData[i];
    var e = { classId:S(row[ci]), day:S(row[di]), period:S(row[pi]), subId:S(row[si]), teaId:S(row[ti]), roomId:row.length>ri?S(row[ri]):"" };
    if (!e.classId||!e.day||!e.period) continue;
    e.className = maps.classes[e.classId]?maps.classes[e.classId].name:e.classId;
    e.subName = maps.subjects[e.subId]||e.subId;
    e.teaName = maps.teachers[e.teaId]?maps.teachers[e.teaId].name:e.teaId;
    e.teaRole = maps.teachers[e.teaId]?maps.teachers[e.teaId].role:"";
    if (mode==="homeroom") { if(e.classId===targetClassId) timetable.push(e); }
    else { if(e.teaId===teacherId) timetable.push(e); }
  }
  return { timetable:timetable, homeroomClassId:targetClassId };
}

/* ================================================================
   3. Dashboard
   ================================================================ */
function getDashboardTeachers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var maps = buildAllMaps(ss);
  var dbSheet = ss.getSheetByName("課表資料庫");
  var countMap = {};
  if (dbSheet && dbSheet.getLastRow() > 1) {
    var dbData = dbSheet.getDataRange().getValues(); var ti=ix(dbData[0],"教師代碼",4);
    for(var i=1;i<dbData.length;i++){var tid=S(dbData[i][ti]);if(tid)countMap[tid]=(countMap[tid]||0)+1;}
  }
  var roleOrder={"導師":1,"科任":2,"行政":3,"特教":4,"教支":5,"外籍":6};
  var result=[];
  for(var tid in maps.teachers){
    var t=maps.teachers[tid];var hrClass="";
    if(t.classCode&&maps.classes[t.classCode])hrClass=maps.classes[t.classCode].name;
    result.push({id:tid,name:t.name,role:t.role,className:hrClass,total:countMap[tid]||0});
  }
  result.sort(function(a,b){var ra=roleOrder[a.role]||99,rb=roleOrder[b.role]||99;if(ra!==rb)return ra-rb;return a.name.localeCompare(b.name,'zh-Hant');});
  return { teachers:result };
}

/* ================================================================
   4. 導師調課
   ================================================================ */
function swapHomeroomLessons(teacherId, classId, dayA, periodA, dayB, periodB) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var maps = buildAllMaps(ss);
  var tea = maps.teachers[teacherId];
  if (!tea||tea.classCode!==classId) return { error:"您不是此班級的導師" };

  var dbSheet = ss.getSheetByName("課表資料庫");
  var dbData = dbSheet.getDataRange().getValues(); var h=dbData[0];
  var ci=ix(h,"班級代碼",0),di=ix(h,"星期",1),pi=ix(h,"節次",2),si=ix(h,"科目代碼",3),ti=ix(h,"教師代碼",4),ri=ix(h,"指定教室",6);
  var rowA=-1,rowB=-1;
  for(var i=1;i<dbData.length;i++){
    if(S(dbData[i][ci])===classId&&S(dbData[i][ti])===teacherId){
      if(S(dbData[i][di])===dayA&&S(dbData[i][pi])===periodA)rowA=i+1;
      if(S(dbData[i][di])===dayB&&S(dbData[i][pi])===periodB)rowB=i+1;
    }
  }
  if(rowA===-1||rowB===-1) return { error:"找不到要交換的課程" };

  var subA=dbSheet.getRange(rowA,si+1).getValue(),roomA=dbSheet.getRange(rowA,ri+1).getValue();
  var subB=dbSheet.getRange(rowB,si+1).getValue(),roomB=dbSheet.getRange(rowB,ri+1).getValue();
  dbSheet.getRange(rowA,si+1).setValue(subB);dbSheet.getRange(rowA,ri+1).setValue(roomB);
  dbSheet.getRange(rowB,si+1).setValue(subA);dbSheet.getRange(rowB,ri+1).setValue(roomA);

  // 回寫原始排課系統
  try {
    var stSheet=ss.getSheetByName("_系統設定");
    if(stSheet){
      var stD=stSheet.getDataRange().getValues();var srcId=null;
      for(var i=1;i<stD.length;i++){if(stD[i][0]==="SOURCE_SS_ID"){srcId=S(stD[i][1]);break;}}
      if(srcId){
        var srcSS=SpreadsheetApp.openById(srcId);var srcDB=srcSS.getSheetByName("課表資料庫");
        if(srcDB){
          var srcData=srcDB.getDataRange().getValues();var sh=srcData[0];
          var sci=ix(sh,"班級代碼",0),sdi=ix(sh,"星期",1),spi=ix(sh,"節次",2),ssi=ix(sh,"科目代碼",3),sti=ix(sh,"教師代碼",4),sri=ix(sh,"指定教室",6);
          for(var i=1;i<srcData.length;i++){
            if(S(srcData[i][sci])===classId&&S(srcData[i][sti])===teacherId){
              if(S(srcData[i][sdi])===dayA&&S(srcData[i][spi])===periodA){srcDB.getRange(i+1,ssi+1).setValue(subB);srcDB.getRange(i+1,sri+1).setValue(roomB);}
              if(S(srcData[i][sdi])===dayB&&S(srcData[i][spi])===periodB){srcDB.getRange(i+1,ssi+1).setValue(subA);srcDB.getRange(i+1,sri+1).setValue(roomA);}
            }
          }
        }
      }
    }
  } catch(e){Logger.log("回寫失敗："+e.message);}
  return { success:true };
}

/* ================================================================
   工具
   ================================================================ */
function S(v){return v==null?"":String(v).trim();}
function ix(h,name,fb){for(var i=0;i<h.length;i++){if(S(h[i])===name)return i;}return fb;}

function buildAllMaps(ss){
  var teachers={},classes={},subjects={},classNameToId={};
  var cs=ss.getSheetByName("班級名單");
  if(cs&&cs.getLastRow()>1){
    var cd=cs.getDataRange().getValues();var ch=cd[0];var cidx=ix(ch,"班級代碼",0),cnidx=ix(ch,"班級名稱",1);
    for(var i=1;i<cd.length;i++){var id=S(cd[i][cidx]),name=S(cd[i][cnidx]);if(id&&name){classes[id]={name:name};classNameToId[name]=id;}}
  }
  var ts=ss.getSheetByName("教師名單");
  if(ts&&ts.getLastRow()>1){
    var td=ts.getDataRange().getValues();var th=td[0];
    var tidx=ix(th,"教師代碼",0),nidx=ix(th,"姓名",1),ridx=ix(th,"身份",2),cidx2=ix(th,"任教班級代碼",3),eidx=ix(th,"電子郵件",-1);
    for(var i=1;i<td.length;i++){
      var id=S(td[i][tidx]),name=S(td[i][nidx]),role=S(td[i][ridx]);
      var clsRaw=S(td[i][cidx2]);var email=eidx!==-1?S(td[i][eidx]):"";
      if(!id)continue;
      var classCode=classes[clsRaw]?clsRaw:(classNameToId[clsRaw]||"");
      teachers[id]={name:name,role:role,classCode:classCode,email:email};
      if(classCode&&classes[classCode])classes[classCode].homeroomId=id;
    }
  }
  var ss2=ss.getSheetByName("科目清單");
  if(ss2&&ss2.getLastRow()>1){
    var sd=ss2.getDataRange().getValues();var sh=sd[0];var sidx=ix(sh,"科目代碼",0),snidx=ix(sh,"科目名稱",1);
    for(var i=1;i<sd.length;i++){var id=S(sd[i][sidx]);if(id)subjects[id]=S(sd[i][snidx]);}
  }
  return {teachers:teachers,classes:classes,subjects:subjects};
}
