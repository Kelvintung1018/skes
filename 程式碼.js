/* Code.gs - 教甄委員管理系統 (完整整合版) */

// ★★★ 請替換為您的 Web App 網址 (部署後取得) ★★★
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwPaKHUAIE2yUB8kEijIKFYHcEaRl2ts0X2a9efGyZSAo2IsomLWJXkYRvPU_2GIM_B/exec";
const SHEET_NAME = "Candidates";
const SETTINGS_SHEET_NAME = "Settings";

// ==========================================
// 1. API 路由設定 (doGet / doPost)
// ==========================================

// 處理 GET 請求 (前端讀取資料)
function doGet(e) {
  var action = e.parameter.action;
  var data;

  try {
    if (action === 'getAdminData') {
      data = getAdminData();
    } else if (action === 'getSettingsData') {
      data = getSettingsData();
    } else if (action === 'getCandidateInfo') {
      data = getCandidateInfo(e.parameter.uid);
    } else if (action === 'getSmsConfig') {
      data = getSmsConfig();
    } else {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: '無效的 GET 請求' })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: data }))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

// 處理 POST 請求 (前端寫入/修改/刪除/寄信/發簡訊)
function doPost(e) {
  if (!e.postData || !e.postData.contents) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: '缺少 Payload' })).setMimeType(ContentService.MimeType.JSON);
  }

  var requestBody = JSON.parse(e.postData.contents);
  var action = requestBody.action;
  var payload = requestBody.data;
  var responseData = {};

  try {
    if (action === 'adminAddCandidate') { responseData = adminAddCandidate(payload); }
    else if (action === 'adminEditCandidate') { responseData = adminEditCandidate(payload); }
    else if (action === 'adminBatchImport') { responseData = adminBatchImport(payload); }
    else if (action === 'adminDeleteCandidates') { responseData = adminDeleteCandidates(payload); }
    else if (action === 'adminResendEmails') { responseData = adminResendEmails(payload); }
    else if (action === 'saveSmsConfig') { responseData = saveSmsConfig(payload.account, payload.password); }
    else if (action === 'adminSendSMS') { responseData = adminSendSMS(payload.ids, payload.template); }
    else if (action === 'saveSettingsData') { responseData = saveSettingsData(payload); }
    else if (action === 'submitForm') { responseData = submitForm(payload); }
    else {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: '無效的 POST 請求' })).setMimeType(ContentService.MimeType.JSON);
    }

    var finalResponse = responseData.status || responseData.result ? responseData : { status: 'success', data: responseData };
    return ContentService.createTextOutput(JSON.stringify(finalResponse))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 2. 系統設定與範本存取 (Settings)
// ==========================================

function getSettingsData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) return {};
  
  const data = sheet.getDataRange().getValues();
  let settings = {};
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    const val = data[i][1];
    if (key) settings[key] = val;
  }
  return settings;
}

function saveSettingsData(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const updates = {
    'TEMPLATE_EMAIL_INVITE': formObj.tpl_email_invite,
    'TEMPLATE_SMS': formObj.tpl_sms,
    'TEMPLATE_EMAIL_CONFIRM_YES': formObj.tpl_confirm_yes,
    'TEMPLATE_EMAIL_CONFIRM_NO': formObj.tpl_confirm_no
  };
  
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    if (updates.hasOwnProperty(key)) {
      sheet.getRange(i + 1, 2).setValue(updates[key]);
      delete updates[key];
    }
  }
  
  const lastRow = sheet.getLastRow();
  let newRow = lastRow + 1;
  for (const [key, value] of Object.entries(updates)) {
    sheet.getRange(newRow, 1).setValue(key);
    sheet.getRange(newRow, 2).setValue(value);
    newRow++;
  }
  return { status: 'success' };
}

function getSettingValue(key) {
  const settings = getSettingsData();
  return settings[key] || "";
}

function saveSmsConfig(a, p) { 
  const props = PropertiesService.getScriptProperties(); 
  props.setProperty('SMS_ACCOUNT', a); 
  props.setProperty('SMS_PASSWORD', p); 
  return {status:'success'};
}

function getSmsConfig() { 
  const props = PropertiesService.getScriptProperties(); 
  return { 
    account: props.getProperty('SMS_ACCOUNT')||"", 
    hasPassword: !!props.getProperty('SMS_PASSWORD') 
  };
}

// ==========================================
// 3. 資料讀取 API (Dashboard)
// ==========================================

function getAdminData(timestamp) {
  SpreadsheetApp.flush();
  const currentSS = SpreadsheetApp.getActiveSpreadsheet();
  const freshSS = SpreadsheetApp.openById(currentSS.getId());
  const sheet = freshSS.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  const smsTemplate = getSettingValue('TEMPLATE_SMS') || "{{姓名}} 老師您好，誠邀您擔任教甄委員。回覆連結： {{連結}}";
  const timeZone = Session.getScriptTimeZone();
  const fmtTime = (d) => {
    if (!d || !(d instanceof Date)) return "";
    try { return Utilities.formatDate(d, timeZone, "MM/dd HH:mm:ss"); } catch (e) { return ""; }
  };
  
  let stats = { total: 0, sent: 0, read: 0, agree: 0, reject: 0, pending: 0 };
  let list = [];
  
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    stats.total++;
    const batch = data[i][4] || "1"; 
    const status = data[i][5];
    const willingness = data[i][6];
    const diet = data[i][7] || ""; 
    const note = data[i][8];
    const readTime = data[i][10];
    const phone = data[i][11];
    const unit = data[i][12] || ""; 
    const title = data[i][13] || "";
    const smsStatus = data[i][14];

    if (status === "已發送") stats.sent++;
    if (readTime) stats.read++;
    if (willingness === "yes") stats.agree++;
    else if (willingness === "no") stats.reject++;
    else stats.pending++;

    list.push({
      uid: data[i][0],
      name: data[i][1],
      subject: data[i][2],
      email: data[i][3],
      batch: batch, 
      phone: phone, 
      unit: unit,
      title: title,
      status: status,
      smsStatus: smsStatus ? fmtTime(smsStatus) : "",
      willingness: willingness,
      diet: diet,
      note: note,
      readTime: fmtTime(readTime),
      replyTime: fmtTime(data[i][9])
    });
  }
  return { stats: stats, list: list, smsTemplate: smsTemplate, appUrl: WEB_APP_URL };
}

// ==========================================
// 4. CRUD (新增/修改/刪除/匯入)
// ==========================================

function adminAddCandidate(f){
  const lock=LockService.getScriptLock();try{lock.waitLock(10000);}catch(e){return{status:'error'};}
  const s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const nr=s.getLastRow()+1;
  s.getRange(nr,12).setNumberFormat("@");
  
  const d=[Utilities.getUuid(),f.name,f.subject,f.email, f.batch||"1", "未發送","","","","","",String(f.phone||"").trim(),f.unit||"",f.title||"", "", f.diet||""];
  
  s.getRange(nr,1,1,15).setValues([d]);
  SpreadsheetApp.flush();lock.releaseLock();
  return{status:'success'};
}

function adminEditCandidate(f){
  const lock=LockService.getScriptLock();try{lock.waitLock(10000);}catch(e){return{status:'error'};}
  const s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data=s.getDataRange().getValues();
  let ok=false;
  
  for(let i=1;i<data.length;i++){
    if(String(data[i][0])==String(f.uid)){
      const r=i+1;
      s.getRange(r,2).setValue(f.name);
      s.getRange(r,3).setValue(f.subject);
      s.getRange(r,4).setValue(f.email);
      s.getRange(r,5).setValue(f.batch||"1");
      
      s.getRange(r,12).setNumberFormat("@").setValue(String(f.phone));
      s.getRange(r,13).setValue(f.unit||"");
      s.getRange(r,14).setValue(f.title||"");
      s.getRange(r,8).setValue(f.diet||"");
      
      s.getRange(r, 7).setValue(f.willingness);
      if (f.willingness) {
        s.getRange(r, 6).setValue("已回覆");
        s.getRange(r, 10).setValue(new Date());
      } else {
        s.getRange(r, 10).clearContent();
        const currentStatus = s.getRange(r, 6).getValue();
        if (currentStatus === "已回覆") {
          s.getRange(r, 6).setValue("已發送");
        }
      }
      
      s.getRange(r,9).setValue(f.note||"");
      SpreadsheetApp.flush();Utilities.sleep(500);ok=true;break;
    }
  }
  lock.releaseLock();
  return ok?{status:'success'}:{status:'error'};
}

function adminBatchImport(L) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const n = [];
  if (!L || L.length === 0) return { status: 'error' };
  
  for (let i = 0; i < L.length; i++) {
    let m = L[i];
    if (m.name) {
      n.push([
        Utilities.getUuid(),   // 1. UUID
        m.name,                // 2. 姓名
        m.subject,             // 3. 科別
        m.email,               // 4. Email
        m.batch || "1",        // 5. 梯次
        "未發送",              // 6. 狀態
        "",                    // 7. 意願
        "",                    // 8. 飲食
        "",                    // 9. 備註
        "",                    // 10. 回覆時間
        "",                    // 11. 已讀時間
        String(m.phone || ""), // 12. 電話
        m.unit || "",          // 13. 單位
        m.title || "",         // 14. 職稱
        ""                     // 15. 簡訊狀態
      ]);
    }
  }
  
  if (n.length > 0) {
    const r = s.getLastRow() + 1;
    s.getRange(r, 12, n.length, 1).setNumberFormat("@"); 
    s.getRange(r, 1, n.length, 15).setValues(n);
    SpreadsheetApp.flush();
    return { status: 'success', count: n.length };
  }
  return { status: 'error' };
}

function adminDeleteCandidates(ids){
  const s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const d=s.getDataRange().getValues();
  for(let i=d.length-1;i>=1;i--){
    if(ids.includes(d[i][0]))s.deleteRow(i+1);
  }
  SpreadsheetApp.flush();
  return{status:'success'};
}

function getCandidateInfo(u){
  const s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const d=s.getDataRange().getValues();
  const t=String(u).trim();
  for(let i=1;i<d.length;i++){
    if(String(d[i][0]).trim()==t){
      if(d[i][6]!=="")return{status:"responded",name:d[i][1]};
      return{status:"success",name:d[i][1],subject:d[i][2],uid:t};
    }
  }
  return{status:"not_found"};
}

// ==========================================
// 5. 溝通功能 (Email & SMS)
// ==========================================

function adminResendEmails(ids) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  let count = 0;
  let templateStr = getSettingValue('TEMPLATE_EMAIL_INVITE');
  
  if (!templateStr) {
    templateStr = '<p><strong>{{name}}</strong> {{title}} 您好：<br>誠摯邀請您擔任本年度教師甄試委員。<br><a href="{{link}}">點此回覆意願</a></p>';
  }

  for (let i = 1; i < data.length; i++) {
    if (ids.includes(data[i][0]) && data[i][3]) { 
      try {
        const uuid = data[i][0];
        const name = data[i][1];
        const email = data[i][3];
        const title = data[i][13];
        const longLink = WEB_APP_URL + "?uid=" + uuid;
        const shortLink = getShortUrl(longLink); 

        let body = templateStr
          .replace(/{{name}}/g, name)
          .replace(/{{link}}/g, shortLink)
          .replace(/{{title}}/g, title)
          .replace(/<\?= name \?>/g, name)
          .replace(/<\?= link \?>/g, shortLink);

        const tLink = WEB_APP_URL + "?track=true&uid=" + uuid;
        const img = `<img src="${tLink}" width="1" height="1" style="display:none;"/>`;
        
        GmailApp.sendEmail(email, "【誠摯邀請】國小教師甄試委員聘任徵詢 - " + name + title, "", {htmlBody: body + img, name: "教甄委員會"});
        sheet.getRange(i + 1, 6).setValue("已發送"); 
        count++;
      } catch (e) { console.error(e); }
    }
  }
  SpreadsheetApp.flush();
  return { status: 'success', count: count };
}

function adminSendSMS(ids, template) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  const props = PropertiesService.getScriptProperties();
  const SMS_USER = props.getProperty('SMS_ACCOUNT');
  const SMS_PASSWORD = props.getProperty('SMS_PASSWORD');
  
  if (!SMS_USER || !SMS_PASSWORD) {
    return { status: 'error', message: '尚未設定簡訊帳號密碼' };
  }

  let logs = [];
  let successCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (!ids.includes(data[i][0])) continue;
    const name     = data[i][1];
    const phoneRaw = data[i][11];
    const title    = data[i][13];
    
    if (!phoneRaw || String(phoneRaw).trim() === "") {
      logs.push({ name: name, phone: "無", msg: "無電話號碼，略過" });
      continue;
    }

    // 格式化電話號碼 (移除空白與橫線)
    const phone = String(phoneRaw).replace(/[-\s]/g, "");
    const longLink = WEB_APP_URL + "?uid=" + data[i][0];
    const link = getShortUrl(longLink);

    // 替換簡訊內容變數
    const msgContent = template
      .replace(/{{姓名}}/g, name)
      .replace(/{{連結}}/g, link)
      .replace(/{{職稱}}/g, title);
      
    // 準備傳送給 API 的參數 (不用 JSON.stringify，維持物件即可，GAS 會自動轉成 Form Data)
const payload = {
      "UID": SMS_USER,
      "PWD": SMS_PASSWORD,
      "SB": "",
      "MSG": msgContent,
      "DEST": phone,
      "ST": "",
      "RETRYTIME": "1440"
    };
    
    const options = {
      "method": "post",
      // 改為規格書要求的 x-www-form-urlencoded (或者可以直接省略，GAS 預設就是這個)
      "contentType": "application/x-www-form-urlencoded", 
      "payload": payload,
      "muteHttpExceptions": true
      // 刪除所有偽裝的 headers 區塊
    };
    
    try {
      // ★ 修正點：替換為 EVERY8D 官方專用的 API 網域 (api.every8d.com)
      const response = UrlFetchApp.fetch("https://new.e8d.tw/API21/HTTP/sendSMS.ashx", options);
      const resultText = response.getContentText().trim();
      
      // 解析 EVERY8D 回傳的 CSV 格式 (例如: 100.0,1,1,0,3298...)
      const parts = resultText.split(",");
      const firstValue = parseFloat(parts[0]);
      
      if (!isNaN(firstValue) && firstValue < 0) {
        // 如果第一個值是負數，代表帳密錯誤或格式錯誤
        logs.push({ name: name, phone: phone, msg: "失敗 (代碼:" + parts[0] + " / " + parts.slice(1).join(",") + ")" });
      } else {
        const unsend = parseInt(parts[3]);
        const credit = parseFloat(parts[0]);
        if (unsend > 0 && credit === 0) {
          logs.push({ name: name, phone: phone, msg: "警告：點數不足 (BatchID:" + parts[4] + ")" });
        } else {
          logs.push({ name: name, phone: phone, msg: "成功 (餘額:" + parts[0] + " 扣點:" + parts[2] + ")" });
          successCount++;
          // 寫入簡訊發送時間到試算表
          sheet.getRange(i + 1, 15).setValue(new Date());
        }
      }
    } catch (e) {
      logs.push({ name: name, phone: phone, msg: "連線錯誤：" + e.message });
    }
  }

  return { status: 'success', count: successCount, logs: logs };
}

function getShortUrl(l){
  try{
    const r=UrlFetchApp.fetch("https://is.gd/create.php?format=simple&url="+encodeURIComponent(l),{muteHttpExceptions:true});
    if(r.getResponseCode()===200&&r.getContentText().indexOf("http")===0) return r.getContentText();
  }catch(e){}
  return l;
}

// ==========================================
// 6. 表單提交與確認信 (Frontend Submission)
// ==========================================

function submitForm(f) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const uid = String(f.uid).trim();
  let r = -1;
  let targetRowIndex = -1;
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() == uid) {
      r = i + 1;
      targetRowIndex = i; break;
    }
  }

  if (r == -1) return { result: "error" };

  sheet.getRange(r, 6).setValue("已回覆");
  sheet.getRange(r, 7).setValue(f.willingness);

  if (f.willingness === "yes") {
    sheet.getRange(r, 8).setValue(f.diet);
    sheet.getRange(r, 9).setValue(f.memo);
  } else {
    sheet.getRange(r, 9).setValue(f.rejectReason);
  }
  sheet.getRange(r, 10).setValue(new Date());
  
  try {
    const name = data[targetRowIndex][1];
    const email = data[targetRowIndex][3];
    if (email && email.includes("@")) {
      const subject = "【教甄委員會】已收到您的意願回覆 - 確認通知";
      let body = "";
      if (f.willingness === "yes") {
        let tpl = getSettingValue('TEMPLATE_EMAIL_CONFIRM_YES');
        if (!tpl) tpl = "<p>{{name}}  {{title}} 您好：<br>已收到您<strong>願意擔任</strong>的回覆。<br>飲食：{{diet}}<br>備註：{{memo}}</p>";
        body = tpl.replace(/{{name}}/g, name).replace(/{{diet}}/g, f.diet || '未填寫').replace(/{{memo}}/g, f.memo || '無');
      } else {
        let tpl = getSettingValue('TEMPLATE_EMAIL_CONFIRM_NO');
        if (!tpl) tpl = "<p>{{name}}  {{title}} 您好：<br>已收到您<strong>無法擔任</strong>的回覆。感謝您的支持。</p>";
        body = tpl.replace(/{{name}}/g, name);
      }
      GmailApp.sendEmail(email, subject, "", { htmlBody: body, name: "教甄委員會" });
    }
  } catch (e) { console.error("確認信失敗: " + e); }

  return { result: "success", name: data[targetRowIndex][1] };
}

function recordEmailOpen(u){
  const l=LockService.getScriptLock();
  if(l.tryLock(10000)){
    try{
      const s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
      const d=s.getDataRange().getValues();
      const tu=String(u).trim();
      const n=new Date();
      for(let i=1;i<d.length;i++){
        if(String(d[i][0]).trim()===tu){
          if(d[i][10]===""||d[i][10]===null){ s.getRange(i+1,11).setValue(n); }
          break;
        }
      }
    }catch(e){}finally{l.releaseLock();}
  }
}

// ==========================================
// 7. 系統初始化與選單 (Init)
// ==========================================

function initCandidatesSheet() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let s = ss.getSheetByName(SHEET_NAME);
  if (!s) s = ss.insertSheet(SHEET_NAME);
  const h = [["uuid","姓名","科別","Email","梯次","狀態","意願","飲食","備註","填寫時間","已讀時間","電話","單位","職稱","簡訊狀態"]];
  if(s.getLastRow() === 0) {
      s.getRange(1, 1, 1, 15).setValues(h).setFontWeight("bold").setBackground("#4c4c4c").setFontColor("white");
      s.setFrozenRows(1);
      s.getRange("L:L").setNumberFormat("@");
      s.getRange("J:K").setNumberFormat("yyyy/MM/dd HH:mm:ss");
      s.getRange("O:O").setNumberFormat("MM/dd HH:mm");
  }

  let s2 = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!s2) s2 = ss.insertSheet(SETTINGS_SHEET_NAME);
  if (s2.getLastRow() === 0) {
      s2.getRange(1,1,1,3).setValues([["Key", "Value", "Description"]]).setFontWeight("bold");
      const defaults = [
          ["TEMPLATE_EMAIL_INVITE", '<div style="font-family: Microsoft JhengHei, sans-serif; padding: 20px; border: 1px solid #ddd; border-radius: 8px;"> <div style="background-color: #4285f4; color: white; padding: 15px; text-align: center; border-radius: 8px 8px 0 0;"> <h2 style="margin:0;">國小教師甄試委員邀請</h2> </div> <div style="padding: 20px;"> <p><strong>{{name}}</strong> 委員 您好：</p> <p>承蒙您長期對教育界的貢獻，本年度教師甄試委員會，誠摯邀請您擔任甄試委員。</p> <p>敬請撥冗點擊下方連結，回覆您的聘任意願。</p> <div style="text-align: center; margin: 30px 0;"> <a href="{{link}}" style="background-color: #d93025; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold;">點此回覆意願</a> </div> <p style="font-size: 12px; color: #666;">若按鈕無法點擊：<br><a href="{{link}}">{{link}}</a></p> </div> </div>', "邀請信範本"],
          ["TEMPLATE_SMS", "{{姓名}} 老師您好，誠邀您擔任教甄委員。回覆連結： {{連結}}", "簡訊範本"],
          ["TEMPLATE_EMAIL_CONFIRM_YES", '<div style="font-family: Microsoft JhengHei, sans-serif;"> <p><strong>{{name}}</strong> {{title}} 您好：</p> <p>感謝您的回覆，我們已收到您<strong>「願意擔任」</strong>本年度教師甄試委員的通知。</p> <div style="background: #f5f5f5; padding: 15px; border-radius: 5px;"> <ul> <li>飲食：{{diet}}</li> <li>備註：{{memo}}</li> </ul> </div> <p>後續相關時程將盡快與您聯繫。</p> </div>', "確認信-同意"],
          ["TEMPLATE_EMAIL_CONFIRM_NO", '<div style="font-family: Microsoft JhengHei, sans-serif;"> <p><strong>{{name}}</strong> {{title}} 您好：</p> <p>我們已收到您<strong>「無法擔任」</strong>的回覆。</p> <p>感謝您撥冗回覆，希望未來還有機會能邀請您。</p> </div>', "確認信-婉拒"]
      ];
      s2.getRange(2,1,defaults.length,3).setValues(defaults);
      s2.setColumnWidth(2, 400);
  }
  SpreadsheetApp.getUi().alert('資料庫與範本設定初始化完成');
}

function resetDefaultTemplates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Settings");
  if (!sheet) return;

  const tplInvite = `
<div style="font-family: 'Microsoft JhengHei', sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
  <div style="background-color: #4285f4; padding: 20px; text-align: center;">
    <h2 style="color: #ffffff; margin: 0; font-size: 20px;">國小教師甄試委員聘任徵詢</h2>
  </div>
  <div style="padding: 30px; background-color: #ffffff;">
    <p style="font-size: 16px; color: #333; line-height: 1.6;"><strong>{{name}}  {{title}} </strong> 您好：</p>
    <p style="font-size: 15px; color: #555; line-height: 1.6;">
      承蒙 您長期對教育界的專業貢獻與支持，本年度教師甄試委員會正積極籌備中。<br><br>
      誠摯邀請 您擔任本屆甄試委員，藉助 您的專業素養為教育界舉才。敬請撥冗點擊下方按鈕，回覆您的聘任意願及飲食習慣。
    </p>
    <div style="text-align: center; margin: 35px 0;">
      <a href="{{link}}" target="_blank" style="background-color: #d93025; color: #ffffff; font-size: 16px; font-weight: bold; text-decoration: none; padding: 12px 30px; border-radius: 4px; display: inline-block; box-shadow: 0 2px 5px rgba(0,0,0,0.2);">
        點此回覆意願
      </a>
    </div>
    <p style="font-size: 13px; color: #888; border-top: 1px solid #eee; padding-top: 20px; line-height: 1.5;">
      ※ 若上方按鈕無法點擊，請複製以下連結至瀏覽器開啟：<br>
      <a href="{{link}}" target="_blank" style="color: #4285f4; word-break: break-all;">{{link}}</a>
    </p>
  </div>
  <div style="background-color: #f9f9f9; padding: 15px; text-align: center; font-size: 12px; color: #999;">
    教甄委員會 敬上<br>
    (此為系統自動發送信件，請勿直接回覆)
  </div>
</div>`;

  const tplSms = `【教甄邀請】{{姓名}} 老師您好，誠邀您擔任本屆甄試委員。詳情與回覆請點： {{連結}}`;
  
  const tplConfirmYes = `
<div style="font-family: 'Microsoft JhengHei', sans-serif; max-width: 600px; margin: 0 auto; border-left: 4px solid #0f9d58; background-color: #f9f9f9; padding: 20px;">
  <h3 style="color: #0f9d58; margin-top: 0;"> 已收到您的同意回覆</h3>
  <p style="font-size: 15px; color: #333;"><strong>{{name}} {{title}}</strong> 您好：</p>
  <p style="font-size: 15px; color: #555;">感謝您的慨允！我們已收到您<strong>「願意擔任」</strong>本年度教師甄試委員的通知。</p>
  <div style="background-color: #fff; border: 1px solid #ddd; padding: 15px; border-radius: 5px; margin: 20px 0;">
    <strong style="color: #333;">您的登記資料：</strong>
    <ul style="color: #555; font-size: 14px; margin-bottom: 0;">
      <li><strong>飲食習慣：</strong>{{diet}}</li>
      <li><strong>備註事項：</strong>{{memo}}</li>
    </ul>
  </div>
  <p style="font-size: 14px; color: #666;">後續相關聘書發放及詳細甄試時程，我們將會盡快與您聯繫。</p>
</div>`;

  const tplConfirmNo = `
<div style="font-family: 'Microsoft JhengHei', sans-serif; max-width: 600px; margin: 0 auto; border-left: 4px solid #757575; background-color: #f9f9f9; padding: 20px;">
  <h3 style="color: #555; margin-top: 0;">已收到您的回覆</h3>
  <p style="font-size: 15px; color: #333;"><strong>{{name}} {{title}}</strong> 您好：</p>
  <p style="font-size: 15px; color: #555;">我們已收到您<strong>「無法擔任」</strong>本年度教師甄試委員的回覆。</p>
  <p style="font-size: 15px; color: #555;">感謝您撥冗回覆，也謝謝您長期對教育界的關心與支持，期盼未來還有機會能邀請您參與。</p>
</div>`;

  const updates = {
    'TEMPLATE_EMAIL_INVITE': tplInvite,
    'TEMPLATE_SMS': tplSms,
    'TEMPLATE_EMAIL_CONFIRM_YES': tplConfirmYes,
    'TEMPLATE_EMAIL_CONFIRM_NO': tplConfirmNo
  };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    if (updates.hasOwnProperty(key)) {
      sheet.getRange(i + 1, 2).setValue(updates[key]);
    }
  }
}

function onOpen(){
  SpreadsheetApp.getUi().createMenu(' 甄試系統管理')
    .addItem(' 初始化/修復 資料庫','initCandidatesSheet')
    .addItem(' 強制更新預設信件範本','resetDefaultTemplates')
    .addToUi();
}