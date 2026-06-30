// ==========================================
// 1) ตั้งค่าระบบ
// ==========================================

const CHANNEL_ACCESS_TOKEN = "dP8y6cOaNAH3uJaJB+Iqu9D/R5ZS4GstC9BTODQ3tBTmBfs5bbHlHZSESDLufCkXYSL7wGfIPgIhacgxpfMe33NuuVaCkBlxUH7dsJcv7h79gYXPvOXyKk6wkf19YuZnTuqfTO9loJS2GaQlYdd3awdB04t89/1O/w1cDnyilFU=";
const TARGET_PUSH_ID       = "Cef301f4f6f5df2aa74812effaaaf3917";

const TELEGRAM_BOT_TOKEN = "8826936665:AAGIXFboMD6xzhLO2ArOtJJLPOtrc_1za9c";
const TELEGRAM_CHAT_ID   = "-1003646625540";

const SPREADSHEET_ID  = "1AytCW84h5vWSMgBHANIfGrrJli6Y1F1io5oKLcYMRi8";
const FORM_SHEET_NAME = "Form_Responses";
const PHOTO_FOLDER_ID = "1l2huO5opSjl1U_TNmIEVC7fqbRD8_ivU";

const MECHANIC_NAMES = ["สุรชัย", "ศักดิ์สิทธิ์", "ศราวุติ", "กมลภพ", "ชยุตม์"];

const GAS_DEPLOY_URL = "https://script.google.com/macros/s/AKfycbwEyqD4iNnDnchjPRvpbMg4ghfN_iePrZVK3zsEgxJ1YaLUUGDnG_q69tL36Usi4Hrf9Q/exec";

// Column mapping:
// col A-K  (0-10)  = ข้อมูล form
// col L    (11)    = jobId
// col M    (12)    = status
// col N    (13)    = acceptTime
// col O    (14)    = receiverUserId
// col P    (15)    = receiverName
// col Q    (16)    = expectedTime
// col R    (17)    = finishTime
// col S    (18)    = ratingSpeed
// col T    (19)    = ratingQuality
// col U    (20)    = ratingComment
// col V    (21)    = remark (หมายเหตุ รองาน)

// ==========================================
// 2) Webhook / API
// ==========================================

function doGet(e) {
  const type = (e && e.parameter && e.parameter.type) || "";

  if (type === "get_mechanics") {
    return jsonResponse({ status: "success", mechanics: MECHANIC_NAMES });
  }

  if (type === "get_maintenance") {
    const sheet  = getFormSheet();
    const values = sheet.getDataRange().getDisplayValues();
    const jobs   = [];
    for (let i = 1; i < values.length; i++) {
      if (values[i][11]) jobs.push(rowToJob(values[i]));
    }
    jobs.reverse();
    return jsonResponse({ status: "success", jobs: jobs });
  }

  if (type === "rate") {
    const jobId = e.parameter.jobId || "";
    return HtmlService.createHtmlOutput(getRatingHtml(jobId))
      .setTitle("ให้คะแนนบริการซ่อมบำรุง")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return jsonResponse({ status: "error", message: "Invalid type" });
}

function onFormSubmit(e) {
  try {
    const sheet = e.range.getSheet();
    const row   = e.range.getRow();
    const v     = e.values;
    const jobId = "JOB-" + row;

    sheet.getRange(row, 12).setValue(jobId);
    sheet.getRange(row, 13).setValue("รอรับงาน");
    for (let col = 14; col <= 22; col++) sheet.getRange(row, col).clearContent();

    const imageUrl = moveUploadedFile(v[10] || "");
    const job = {
      time: v[0]||"", department: v[1]||"", reporter: v[2]||"",
      location: v[3]||"-", machineNo: v[4]||"-", topic: v[5]||"",
      detail: v[6]||v[7]||v[8]||v[9]||"-",
      imageUrl: imageUrl, jobId: jobId, status: "รอรับงาน",
      receiverName: "-", expectedTime: "-", remark: ""
    };

    sendRepairCard(job);
    sendTelegramRepairCard(TELEGRAM_CHAT_ID, job);
  } catch (err) {
    console.log("onFormSubmit error: " + err);
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) return ContentService.createTextOutput("OK");
    const data = JSON.parse(e.postData.contents);

    const updateId = data.update_id;
    if (updateId) {
      const cache = CacheService.getScriptCache();
      if (cache.get(updateId.toString())) return ContentService.createTextOutput("OK");
      cache.put(updateId.toString(), "1", 300);
    }

    if (data.callback_query) {
      handleTelegramCallback(data.callback_query);
      return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
    }

    if (data.events) {
      data.events.forEach(event => {
        if (event.type === "postback") handlePostback(event);
        if (event.type === "message" && event.message.type === "text") handleTextCommand(event);
      });
    }

    if (data.message && data.message.text) {
      handleTelegramTextCommand(data.message);
      return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
    }

    if (data.type === "accept_job")    return jsonResponse({ status: acceptJob(data.jobId, data.userId || "APP", data.receiverName || "ไม่ระบุชื่อ", data.expectedTime || "-") });
    if (data.type === "finish_job")    return jsonResponse({ status: finishJob(data.jobId, false) });
    if (data.type === "cancel_job")    return jsonResponse({ status: cancelJob(data.jobId, "APP") });
    if (data.type === "submit_rating") return jsonResponse({ status: submitRating(data.jobId, data.ratingSpeed, data.ratingQuality, data.comment) });

    return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    console.log("Error in doPost: " + err.message);
    return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
  }
}

// ==========================================
// 3) ระบบให้คะแนน
// ==========================================

function checkJobExists(jobId) {
  const sheet = getFormSheet();
  const cell  = sheet.createTextFinder(jobId).matchEntireCell(true).findNext();
  if (!cell) return { found: false };
  const row    = cell.getRow();
  const status = sheet.getRange(row, 13).getDisplayValue();
  const rated  = !!(sheet.getRange(row, 19).getValue() || sheet.getRange(row, 20).getValue());
  return { found: true, status: status, alreadyRated: rated };
}

function handleRatingSubmit(jobId, ratingSpeed, ratingQuality, comment) {
  return submitRating(jobId, ratingSpeed, ratingQuality, comment);
}

function submitRating(jobId, ratingSpeed, ratingQuality, comment) {
  const sheet = getFormSheet();
  const cell  = sheet.createTextFinder(jobId).matchEntireCell(true).findNext();
  if (!cell) return "notfound";
  const row = cell.getRow();
  if (sheet.getRange(row, 19).getValue() || sheet.getRange(row, 20).getValue()) return "already_rated";
  // ✅ setNumberFormat ป้องกัน Sheet แปลงตัวเลขเป็นวันที่
  sheet.getRange(row, 19).setNumberFormat("0").setValue(Number(ratingSpeed)   || 0);
  sheet.getRange(row, 20).setNumberFormat("0").setValue(Number(ratingQuality) || 0);
  sheet.getRange(row, 21).setValue(String(comment || ""));
  SpreadsheetApp.flush();
  return "success";
}

function sendRatingQR(jobId, topic, receiverName) {
  const pageUrl = GAS_DEPLOY_URL + "?type=rate&jobId=" + encodeURIComponent(jobId);
  const qrUrl   = "https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=" + encodeURIComponent(pageUrl);

  // ✅ LINE — QR อยู่ใน hero ของ flex (กดขยายได้ + กดเปิดลิงก์ได้)
  pushLine({
    to: TARGET_PUSH_ID,
    messages: [{
      type: "flex",
      altText: "ให้คะแนนงาน " + jobId,
      contents: {
        type: "bubble",
        hero: {
          type: "image",
          url: qrUrl,
          size: "full",
          aspectRatio: "1:1",
          aspectMode: "fit",
          action: { type: "uri", uri: pageUrl }
        },
        body: {
          type: "box", layout: "vertical", spacing: "md", contents: [
            { type: "text", text: "⭐ ให้คะแนนบริการซ่อมบำรุง", weight: "bold", size: "lg", wrap: true },
            { type: "text", text: "งาน: " + jobId, size: "sm", color: "#666666" },
            { type: "text", text: "หัวข้อ: " + topic, size: "sm", color: "#666666", wrap: true },
            { type: "text", text: "ช่าง: " + receiverName, size: "sm", color: "#666666" },
            { type: "text", text: "📱 สแกน QR ด้านบน หรือกดปุ่มด้านล่าง", size: "xs", color: "#999999", wrap: true, margin: "md" }
          ]
        },
        footer: {
          type: "box", layout: "vertical", contents: [{
            type: "button", style: "primary", color: "#1565C0",
            action: { type: "uri", label: "⭐ กดให้คะแนนที่นี่", uri: pageUrl }
          }]
        }
      }
    }]
  });

  // Telegram
  UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendPhoto", {
    method: "post", contentType: "application/json",
    payload: JSON.stringify({
      chat_id: TELEGRAM_CHAT_ID, photo: qrUrl,
      caption: "⭐ <b>ให้คะแนนบริการซ่อมบำรุง</b>\nงาน: <b>" + jobId + "</b> | ช่าง: " + receiverName + "\n📱 สแกน QR หรือกดปุ่มด้านล่าง",
      parse_mode: "HTML",
      reply_markup: { inline_keyboard: [[{ text: "⭐ กดให้คะแนนที่นี่", url: pageUrl }]] }
    }),
    muteHttpExceptions: true
  });
}

function getRatingHtml(jobId) {
  const prefill = jobId ? jobId : "";
  return `<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
  <meta name="theme-color" content="#1565C0">
  <title>ให้คะแนนบริการซ่อมบำรุง</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}
    html,body{height:100%;overflow-x:hidden}
    body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:linear-gradient(160deg,#1565C0 0%,#1976D2 40%,#e8f0fe 40%);min-height:100vh;display:flex;align-items:flex-start;justify-content:center;padding:24px 16px 40px}
    .card{background:#fff;border-radius:28px;padding:28px 20px 32px;width:100%;max-width:420px;box-shadow:0 12px 48px rgba(0,0,0,.18)}
    .logo{text-align:center;margin-bottom:16px}
    .logo span{display:inline-flex;width:64px;height:64px;background:#e8f5e9;border-radius:50%;align-items:center;justify-content:center;font-size:32px}
    h1{text-align:center;font-size:20px;font-weight:700;color:#1a1a1a;margin-bottom:4px}
    .sub{text-align:center;color:#666;font-size:14px;margin-bottom:18px;line-height:1.5}
    .job-section{margin-bottom:20px}
    .job-label{font-size:13px;font-weight:700;color:#444;margin-bottom:6px}
    .job-row{display:flex;gap:8px}
    #jobInput{flex:1;border:2px solid #e0e0e0;border-radius:14px;padding:12px;font-size:18px;font-weight:700;text-align:center;letter-spacing:1px;color:#1565C0;outline:none;font-family:inherit;text-transform:uppercase}
    #jobInput:focus{border-color:#1565C0}
    #jobInput.ok{background:#E3F2FD;border-color:#4CAF50}
    #jobInput.err{border-color:#EF5350}
    #confirmBtn{padding:0 18px;border:none;border-radius:14px;background:#1565C0;color:#fff;font-weight:700;font-size:14px;cursor:pointer;white-space:nowrap}
    #confirmBtn:disabled{background:#BDBDBD}
    .job-msg{font-size:12px;margin-top:6px;min-height:18px}
    .job-msg.ok{color:#4CAF50}.job-msg.err{color:#EF5350}
    .rate-area{opacity:0.3;pointer-events:none;transition:opacity .3s}
    .rate-area.active{opacity:1;pointer-events:auto}
    .rating-block{background:#F8F9FA;border-radius:20px;padding:16px;margin-bottom:14px}
    .rating-title{font-size:13px;font-weight:700;color:#444;margin-bottom:10px;text-align:center}
    .stars{display:flex;justify-content:center;gap:4px;margin-bottom:4px}
    .star{flex:1;max-width:52px;height:52px;display:flex;align-items:center;justify-content:center;font-size:38px;cursor:pointer;color:#E0E0E0;transition:transform .12s,color .12s;border-radius:10px;touch-action:manipulation;-webkit-user-select:none;user-select:none}
    .star.on{color:#FFC107}.star.pop{transform:scale(1.35)}
    .mood{text-align:center;font-weight:700;font-size:14px;min-height:20px;transition:color .2s}
    .divider{height:1px;background:#eee;margin:4px 0 14px}
    .cmt-label{font-size:13px;font-weight:600;color:#444;margin-bottom:8px}
    textarea{width:100%;border:1.5px solid #e0e0e0;border-radius:16px;padding:14px;font-size:16px;line-height:1.5;resize:none;outline:none;font-family:inherit;color:#333}
    textarea:focus{border-color:#1565C0}
    textarea::placeholder{color:#BDBDBD}
    .btn{display:block;width:100%;padding:16px;border:none;border-radius:16px;font-size:16px;font-weight:700;color:#fff;background:#4CAF50;cursor:pointer;margin-top:16px;touch-action:manipulation}
    .btn:active{transform:scale(.98)}.btn:disabled{background:#E0E0E0;color:#BDBDBD;cursor:default}
    #doneCard,#alreadyCard{display:none;text-align:center;padding:12px 0}
    .done-icon{font-size:72px;margin-bottom:16px}
    #doneCard h2{font-size:22px;font-weight:700;color:#2E7D32;margin-bottom:8px}
    #alreadyCard h2{font-size:22px;font-weight:700;color:#1565C0;margin-bottom:8px}
    #doneCard p,#alreadyCard p{color:#777;font-size:15px;line-height:1.6}
    .spinner{display:inline-block;width:18px;height:18px;border:3px solid #fff;border-top-color:transparent;border-radius:50%;animation:spin .6s linear infinite;vertical-align:middle}
    @keyframes spin{to{transform:rotate(360deg)}}
  </style>
</head>
<body>
<div class="card">
  <div id="formCard">
    <div class="logo"><span>🔧</span></div>
    <h1>ให้คะแนนบริการซ่อมบำรุง</h1>
    <p class="sub">สแกน QR หรือกรอกเลขงานด้านล่าง</p>
    <div class="job-section">
      <p class="job-label">เลขงาน (Job ID)</p>
      <div class="job-row">
        <input type="text" id="jobInput" placeholder="เช่น JOB-23" value="${prefill}">
        <button id="confirmBtn" onclick="confirmJob()">ยืนยัน</button>
      </div>
      <p class="job-msg" id="jobMsg"></p>
    </div>
    <div class="rate-area" id="rateArea">
      <div class="rating-block">
        <p class="rating-title">⚡ ความรวดเร็วในการแก้ปัญหา</p>
        <div class="stars" id="starsSpeed">
          <div class="star" data-g="speed" data-v="1">★</div>
          <div class="star" data-g="speed" data-v="2">★</div>
          <div class="star" data-g="speed" data-v="3">★</div>
          <div class="star" data-g="speed" data-v="4">★</div>
          <div class="star" data-g="speed" data-v="5">★</div>
        </div>
        <div class="mood" id="moodSpeed"></div>
      </div>
      <div class="rating-block">
        <p class="rating-title">⭐ คุณภาพการซ่อม / แก้ไข</p>
        <div class="stars" id="starsQuality">
          <div class="star" data-g="quality" data-v="1">★</div>
          <div class="star" data-g="quality" data-v="2">★</div>
          <div class="star" data-g="quality" data-v="3">★</div>
          <div class="star" data-g="quality" data-v="4">★</div>
          <div class="star" data-g="quality" data-v="5">★</div>
        </div>
        <div class="mood" id="moodQuality"></div>
      </div>
      <div class="divider"></div>
      <p class="cmt-label">ความคิดเห็นเพิ่มเติม (ไม่บังคับ)</p>
      <textarea id="cmt" rows="3" placeholder="บอกเราว่าควรปรับปรุงอะไร..."></textarea>
      <button class="btn" id="submitBtn" disabled onclick="doSubmit()">ส่งคะแนน</button>
    </div>
  </div>
  <div id="doneCard">
    <div class="done-icon">🎉</div>
    <h2>ขอบคุณมากครับ!</h2>
    <p>คะแนนของคุณถูกบันทึกแล้ว<br>ทีมช่างจะนำไปปรับปรุงการบริการ</p>
  </div>
  <div id="alreadyCard">
    <div class="done-icon">✅</div>
    <h2>ให้คะแนนไปแล้ว</h2>
    <p id="alreadyMsg">งานนี้ได้รับการให้คะแนนเรียบร้อยแล้วครับ</p>
  </div>
</div>
<script>
  const MOODS  = ['','😞 ต้องปรับปรุง','😐 พอใช้ได้','🙂 ดี','😊 ดีมาก','🤩 ยอดเยี่ยม!'];
  const COLORS = ['','#EF5350','#FF9800','#66BB6A','#26A69A','#1565C0'];
  let selSpeed=0, selQuality=0, confirmedJobId='';

  function showCard(id) {
    ['formCard','doneCard','alreadyCard'].forEach(function(x){
      document.getElementById(x).style.display = x===id?'block':'none';
    });
  }

  function confirmJob() {
    var input=document.getElementById('jobInput'), msg=document.getElementById('jobMsg'), btn=document.getElementById('confirmBtn');
    var val=input.value.trim().toUpperCase();
    if (!val) { msg.textContent='กรุณากรอกเลขงาน'; msg.className='job-msg err'; return; }
    if (!val.startsWith('JOB-')) val='JOB-'+val.replace(/[^0-9]/g,'');
    input.value=val; btn.disabled=true; btn.textContent='...';
    msg.textContent='กำลังตรวจสอบ...'; msg.className='job-msg';
    google.script.run
      .withSuccessHandler(function(r) {
        if (r && r.found) {
          if (r.alreadyRated) { showCard('alreadyCard'); document.getElementById('alreadyMsg').textContent=val+' ได้รับการให้คะแนนเรียบร้อยแล้วครับ'; return; }
          confirmedJobId=val; input.classList.add('ok'); input.classList.remove('err'); input.readOnly=true;
          msg.textContent='✅ พบงาน (สถานะ: '+r.status+')'; msg.className='job-msg ok';
          document.getElementById('rateArea').classList.add('active'); btn.textContent='✓';
        } else {
          input.classList.add('err'); input.classList.remove('ok');
          msg.textContent='❌ ไม่พบเลขงานนี้ในระบบ'; msg.className='job-msg err';
          btn.disabled=false; btn.textContent='ยืนยัน';
        }
      })
      .withFailureHandler(function() {
        msg.textContent='⚠️ เกิดข้อผิดพลาด ลองใหม่'; msg.className='job-msg err';
        btn.disabled=false; btn.textContent='ยืนยัน';
      })
      .checkJobExists(val);
  }

  window.onload=function(){ if(document.getElementById('jobInput').value) confirmJob(); };

  function setupStars(group, moodId, onSelect) {
    document.querySelectorAll('.star[data-g="'+group+'"]').forEach(function(s) {
      s.addEventListener('click', function() {
        var v=+this.dataset.v; onSelect(v);
        document.querySelectorAll('.star[data-g="'+group+'"]').forEach(function(st,i){ st.classList.toggle('on',i<v); });
        var lbl=document.getElementById(moodId); lbl.textContent=MOODS[v]; lbl.style.color=COLORS[v];
        this.classList.add('pop'); var self=this; setTimeout(function(){ self.classList.remove('pop'); },150);
        checkReady();
      });
    });
  }
  setupStars('speed',   'moodSpeed',   function(v){ selSpeed=v; });
  setupStars('quality', 'moodQuality', function(v){ selQuality=v; });

  function checkReady() { document.getElementById('submitBtn').disabled=!(selSpeed>0&&selQuality>0&&confirmedJobId); }

  function doSubmit() {
    var cmt=document.getElementById('cmt').value.trim(), btn=document.getElementById('submitBtn');
    btn.disabled=true; btn.innerHTML='<span class="spinner"></span> กำลังส่ง...';
    google.script.run
      .withSuccessHandler(function(result) {
        if (result==='success') showCard('doneCard');
        else if (result==='already_rated') showCard('alreadyCard');
        else { btn.disabled=false; btn.textContent='ส่งคะแนน'; alert('เกิดข้อผิดพลาด: '+result); }
      })
      .withFailureHandler(function() { btn.disabled=false; btn.textContent='ส่งคะแนน'; alert('เกิดข้อผิดพลาด กรุณาลองใหม่'); })
      .handleRatingSubmit(confirmedJobId, selSpeed, selQuality, cmt);
  }
</script>
</body>
</html>`;
}

// ==========================================
// 4) Telegram Callback
// ==========================================

function handleTelegramCallback(cb) {
  const data   = cb.data;
  const chatId = cb.message.chat.id;
  const msgId  = cb.message.message_id;
  const cbId   = cb.id;
  const apiUrl = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN;

  const getParam = key => { for (let p of data.split("&")) { if (p.startsWith(key+"=")) return decodeURIComponent(p.split("=")[1]); } return ""; };
  const action = getParam("a");
  const jNum   = getParam("j");
  const jobId  = "JOB-" + jNum;

  let reqs = [];
  const q = (ep, payload) => reqs.push({ url: apiUrl+ep, method:"post", contentType:"application/json", payload: JSON.stringify(payload), muteHttpExceptions: true });

  if (action === "acc") {
    const keys = MECHANIC_NAMES.map((name, idx) => ([{ text: "👨‍🔧 "+name, callback_data: "a=sn&j="+jNum+"&n="+idx }]));
    q("/answerCallbackQuery", { callback_query_id: cbId, text: "โปรดเลือกชื่อช่างผู้รับงาน" });
    q("/editMessageReplyMarkup", { chat_id: chatId, message_id: msgId, reply_markup: { inline_keyboard: keys } });
    UrlFetchApp.fetchAll(reqs);
  }
  else if (action === "sn") {
    const nIdx        = parseInt(getParam("n"));
    const minutesList = [15, 30, 60, 120, 180, 240, 360, 480, 1440];
    let keys = [], row = [];
    minutesList.forEach((m, i) => {
      row.push({ text: formatMinuteLabel(m), callback_data: "a=ct&j="+jNum+"&n="+nIdx+"&m="+m });
      if (row.length===3 || i===minutesList.length-1) { keys.push(row); row=[]; }
    });
    q("/answerCallbackQuery", { callback_query_id: cbId, text: "เลือกเวลาคาดการณ์ปิดงาน" });
    q("/editMessageReplyMarkup", { chat_id: chatId, message_id: msgId, reply_markup: { inline_keyboard: keys } });
    UrlFetchApp.fetchAll(reqs);
  }
  else if (action === "ct") {
    q("/answerCallbackQuery", { callback_query_id: cbId });
    UrlFetchApp.fetchAll(reqs); reqs = [];
    const nIdx     = parseInt(getParam("n"));
    const techName = MECHANIC_NAMES[nIdx];
    const minutes  = parseInt(getParam("m"));
    const deadline = new Date();
    deadline.setMinutes(deadline.getMinutes() + minutes);
    const formattedDeadline = Utilities.formatDate(deadline, "GMT+7", "dd/MM/yyyy HH:mm");
    const result   = acceptJob(jobId, cb.from.id.toString(), techName, formattedDeadline);
    if (result === "success") {
      const keys = [[
        { text: "🏁 เสร็จงาน",   callback_data: "a=fin&j="+jNum },
        { text: "⏸️ หมายเหตุ",  callback_data: "a=note&j="+jNum },
        { text: "❌ ยกเลิก",     callback_data: "a=can&j="+jNum }
      ]];
      q("/editMessageCaption", { chat_id: chatId, message_id: msgId, caption: createTelegramCaption(getJobDataById(jobId)), parse_mode: "HTML", reply_markup: { inline_keyboard: keys } });
    } else if (result === "already") {
      q("/sendMessage", { chat_id: chatId, text: "⚠️ ช้าไปครับ งาน "+jobId+" มีคนรับไปแล้ว" });
      q("/editMessageReplyMarkup", { chat_id: chatId, message_id: msgId, reply_markup: { inline_keyboard: [] } });
    }
    UrlFetchApp.fetchAll(reqs);
  }
  else if (action === "fin") {
    const result = finishJob(jobId, true);
    if (result === "success") {
      q("/answerCallbackQuery", { callback_query_id: cbId, text: "ปิดงานเรียบร้อย! กำลังส่ง QR ให้คะแนน..." });
      q("/editMessageCaption", { chat_id: chatId, message_id: msgId, caption: createTelegramCaption(getJobDataById(jobId)), parse_mode: "HTML", reply_markup: { inline_keyboard: [] } });
    } else {
      q("/answerCallbackQuery", { callback_query_id: cbId, text: "งานนี้ถูกปิดไปแล้ว", show_alert: true });
    }
    UrlFetchApp.fetchAll(reqs);
  }
  else if (action === "can") {
    const result = cancelJob(jobId, cb.from.id.toString());
    if (result === "success") {
      const keys = [[{ text: "✅ กดรับงานนี้", callback_data: "a=acc&j="+jNum }]];
      q("/answerCallbackQuery", { callback_query_id: cbId, text: "ยกเลิกรับงานเรียบร้อย" });
      q("/editMessageCaption", { chat_id: chatId, message_id: msgId, caption: createTelegramCaption(getJobDataById(jobId)), parse_mode: "HTML", reply_markup: { inline_keyboard: keys } });
      q("/sendMessage", { chat_id: chatId, text: "↩️ ยกเลิกรับงาน <b>"+jobId+"</b> เรียบร้อย (ส่งกลับเข้าคิว)", parse_mode: "HTML" });
    } else if (result === "notowner") {
      q("/answerCallbackQuery", { callback_query_id: cbId, text: "⚠️ คุณไม่ใช่ช่างผู้กดรับงานนี้!", show_alert: true });
    }
    UrlFetchApp.fetchAll(reqs);
  }
  else if (action === "note") {
    CacheService.getScriptCache().put("note_tg_" + chatId, jobId, 600);
    q("/answerCallbackQuery", { callback_query_id: cbId });
    q("/sendMessage", {
      chat_id: chatId, parse_mode: "HTML",
      text: "📝 โปรดพิมพ์ <b>หมายเหตุ</b> สำหรับงาน <b>"+jobId+"</b>\n\nตัวอย่าง: รอสั่งอะไหล่, รออนุมัติงบ, รอช่างไฟ"
    });
    UrlFetchApp.fetchAll(reqs);
  }
  else if (action === "resume") {
    const result = resumeJob(jobId);
    if (result === "success") {
      const keys = [[
        { text: "🏁 เสร็จงาน",   callback_data: "a=fin&j="+jNum },
        { text: "⏸️ หมายเหตุ",  callback_data: "a=note&j="+jNum },
        { text: "❌ ยกเลิก",     callback_data: "a=can&j="+jNum }
      ]];
      q("/answerCallbackQuery", { callback_query_id: cbId, text: "✅ กลับมาดำเนินการแล้ว" });
      q("/editMessageCaption", { chat_id: chatId, message_id: msgId, caption: createTelegramCaption(getJobDataById(jobId)), parse_mode: "HTML", reply_markup: { inline_keyboard: keys } });
      q("/sendMessage", { chat_id: chatId, text: "🔄 งาน <b>"+jobId+"</b> กลับมาดำเนินการแล้วครับ", parse_mode: "HTML" });
    } else {
      q("/answerCallbackQuery", { callback_query_id: cbId, text: "⚠️ ไม่สามารถดำเนินการได้", show_alert: true });
    }
    UrlFetchApp.fetchAll(reqs);
  }
}

// ==========================================
// 5) Telegram Text Commands
// ==========================================

function handleTelegramTextCommand(msg) {
  const text   = msg.text.trim();
  const chatId = msg.chat.id;
  const tgUrl  = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendMessage";

  const pendingJobId = CacheService.getScriptCache().get("note_tg_" + chatId);
  if (pendingJobId && !text.startsWith("/")) {
    CacheService.getScriptCache().remove("note_tg_" + chatId);
    const result = setJobRemark(pendingJobId, text);
    if (result === "success") {
      UrlFetchApp.fetch(tgUrl, { method:"post", contentType:"application/json",
        payload: JSON.stringify({ chat_id: chatId, parse_mode: "HTML",
          text: "⏸️ บันทึกหมายเหตุงาน <b>"+pendingJobId+"</b> เรียบร้อย\n📝 หมายเหตุ: " + text
        })
      });
    } else {
      UrlFetchApp.fetch(tgUrl, { method:"post", contentType:"application/json",
        payload: JSON.stringify({ chat_id: chatId, text: "⚠️ ไม่สามารถบันทึกได้ สถานะงานอาจเปลี่ยนไปแล้ว" })
      });
    }
    return;
  }

  if (text.startsWith("/show ")) {
    const targetId = text.replace("/show ", "").trim().toUpperCase();
    const job = getJobDataById(targetId);
    if (job) sendTelegramRepairCard(chatId, job);
    else UrlFetchApp.fetch(tgUrl, { method:"post", contentType:"application/json", payload: JSON.stringify({ chat_id: chatId, text: "❌ ไม่พบเลขงาน " + targetId }) });
    return;
  }

  if (text === "/showall" || text === "/รอรับ" || text === "/กำลังทำ" || text === "/รองาน") {
    const sheet  = getFormSheet();
    const values = sheet.getDataRange().getValues();
    let count = 0;

    if (text === "/รองาน") {
      let sent = 0;
      for (let i = values.length-1; i >= 1 && sent < 10; i--) {
        if (values[i][12] !== "กำลังดำเนินการ") continue;
        const job  = rowToJob(values[i]);
        const jNum = job.jobId.replace("JOB-","");
        UrlFetchApp.fetch("https://api.telegram.org/bot"+TELEGRAM_BOT_TOKEN+"/sendPhoto", {
          method:"post", contentType:"application/json",
          payload: JSON.stringify({
            chat_id: chatId,
            photo: job.imageUrl||"https://placehold.co/800x520/png?text=No+Image",
            caption: createTelegramCaption(job), parse_mode: "HTML",
            reply_markup: { inline_keyboard: [[
              { text: "⏸️ หมายเหตุ/ติดขัด", callback_data: "a=note&j="+jNum },
              { text: "🏁 เสร็จงาน",         callback_data: "a=fin&j="+jNum }
            ]]}
          }),
          muteHttpExceptions: true
        });
        sent++;
      }
      if (sent === 0) UrlFetchApp.fetch(tgUrl, { method:"post", contentType:"application/json", payload: JSON.stringify({ chat_id: chatId, text: "✅ ไม่มีงานที่กำลังดำเนินการอยู่ครับ" }) });
      return;
    }

    const header =
      text === "/showall"  ? "📋 <b>รายการงานค้างทั้งหมด:</b>\n━━━━━━━━━━━━━━━━━━\n"
    : text === "/รอรับ"   ? "⏳ <b>รายการงานรอช่างรับ:</b>\n━━━━━━━━━━━━━━━━━━\n"
    :                        "🛠️ <b>รายการงานกำลังดำเนินการ:</b>\n━━━━━━━━━━━━━━━━━━\n";

    let summaryText = header;
    for (let i = values.length-1; i >= 1; i--) {
      const s = values[i][12];
      const match =
        (text === "/showall"  && (s==="รอรับงาน"||s==="กำลังดำเนินการ"||s==="รองาน"))
     || (text === "/รอรับ"   && s==="รอรับงาน")
     || (text === "/กำลังทำ" && (s==="กำลังดำเนินการ"||s==="รองาน"));
      if (!match) continue;
      count++;
      const remark = values[i][21];
      summaryText += `🔹 <b>${values[i][11]}</b> [${s}]\n• แผนก: ${values[i][1]} | ปัญหา: ${values[i][5]}\n`;
      if (values[i][15]) summaryText += `• ช่าง: ${values[i][15]} (คาดเสร็จ: ${values[i][16]})\n`;
      if (remark)        summaryText += `• ⚠️ หมายเหตุ: ${remark}\n`;
      summaryText += "━━━━━━━━━━━━━━━━━━\n";
      if (count >= 10) break;
    }

    const finalMsg = count === 0 ? "✅ ไม่มีงานในหมวดหมู่นี้ครับ" : summaryText;
    UrlFetchApp.fetch(tgUrl, { method:"post", contentType:"application/json", payload: JSON.stringify({ chat_id: chatId, text: finalMsg, parse_mode: "HTML" }) });
  }
}

// ==========================================
// 6) LINE Bot
// ==========================================

function handlePostback(event) {
  const data   = event.postback.data;
  const userId = event.source.userId || "";

  if (data.indexOf("action=note_start") !== -1) {
    const jobId = data.split("jobId=")[1];
    CacheService.getScriptCache().put("note_line_" + userId, jobId, 600);
    replyLine(event.replyToken, "📝 โปรดพิมพ์หมายเหตุสำหรับงาน " + jobId + "\n(เช่น รอสั่งอะไหล่, รออนุมัติงบ, รอช่างไฟ)");
    return;
  }
  if (data.indexOf("action=resume") !== -1) {
    const jobId  = data.split("jobId=")[1];
    const result = resumeJob(jobId);
    if (result === "success") replyLine(event.replyToken, "🔄 งาน " + jobId + " กลับมาดำเนินการแล้วครับ");
    else replyLine(event.replyToken, "⚠️ ไม่สามารถดำเนินการได้");
    return;
  }
  if (data.indexOf("action=accept") !== -1) {
    askMechanicName(event.replyToken, data.split("jobId=")[1]);
    return;
  }
  if (data.indexOf("action=select_name") !== -1) {
    const parts = data.split("&");
    askTimeQuickReply(event.replyToken, parts[1].split("=")[1], parts[2].split("=")[1]);
    return;
  }
  if (data.indexOf("action=confirm_time") !== -1) {
    let jobId="", techName="", minutes=30;
    data.split("&").forEach(p => {
      if (p.startsWith("jobId=")) jobId    = p.split("=")[1];
      if (p.startsWith("name="))  techName = p.split("=")[1];
      if (p.startsWith("min="))   minutes  = parseInt(p.split("=")[1], 10);
    });
    const deadline = new Date();
    deadline.setMinutes(deadline.getMinutes() + minutes);
    const formattedDeadline = Utilities.formatDate(deadline, "GMT+7", "dd/MM/yyyy HH:mm");
    const result = acceptJob(jobId, userId, techName, formattedDeadline);
    if (result === "success") replyAcceptedCard(event.replyToken, jobId, techName, formattedDeadline);
    else if (result === "already") replyLine(event.replyToken, "⚠️ ช้าไปครับ งาน "+jobId+" มีคนรับไปแล้ว");
    return;
  }
  if (data.indexOf("action=finish") !== -1) {
    const result = finishJob(data.split("jobId=")[1], true);
    if (result === "success")               replyLine(event.replyToken, "🎉 ปิดงานเรียบร้อย กำลังส่ง QR ให้คะแนน...");
    else if (result === "already_finished") replyLine(event.replyToken, "✅ งานนี้ปิดไปแล้วครับ");
    return;
  }
  if (data.indexOf("action=cancel") !== -1) {
    const result = cancelJob(data.split("jobId=")[1], userId);
    if (result === "success")       replyLine(event.replyToken, "↩️ ยกเลิกรับงานเรียบร้อย");
    else if (result === "notowner") replyLine(event.replyToken, "⚠️ ต้องเป็นคนที่รับงานนี้เท่านั้น");
  }
}

function handleTextCommand(event) {
  const text   = event.message.text.trim();
  const userId = event.source.userId || event.source.groupId || "";

  const pendingJobId = CacheService.getScriptCache().get("note_line_" + userId);
  if (pendingJobId && !text.startsWith("/")) {
    CacheService.getScriptCache().remove("note_line_" + userId);
    const result = setJobRemark(pendingJobId, text);
    if (result === "success") replyLine(event.replyToken, "⏸️ บันทึกหมายเหตุงาน " + pendingJobId + " เรียบร้อย\n📝 " + text);
    else replyLine(event.replyToken, "⚠️ ไม่สามารถบันทึกได้ สถานะงานอาจเปลี่ยนไปแล้ว");
    return;
  }

  if (text.startsWith("/show "))   return replyJobCard(event.replyToken, text.replace("/show ", "").trim());
  if (text === "/showall")         return replyAllPendingJobCards(event.replyToken);
  if (text === "/รอรับ")           return replyWaitingJobs(event.replyToken);
  if (text === "/กำลังทำ")         return replyAcceptedJobs(event.replyToken);
  if (text === "/รองาน")           return replyNoteJobs(event.replyToken);
  if (text === "/id")              return replyLine(event.replyToken, "groupId/userId: " + (event.source.groupId || event.source.userId || "ไม่พบ"));
  if (text === "/คำสั่ง" || text === "/help") {
    replyLine(event.replyToken,
      "คำสั่งที่ใช้ได้\n" +
      "/show JOB-6 = แสดงการ์ดงาน\n" +
      "/showall = แสดงงานค้างทั้งหมด\n" +
      "/รอรับ = งานรอช่างรับ\n" +
      "/กำลังทำ = งานกำลังดำเนินการ\n" +
      "/รองาน = เพิ่มหมายเหตุ/ติดขัด"
    );
  }
}

function askMechanicName(replyToken, jobId) {
  replyFlex({ replyToken, messages: [{ type:"text", text:"👨‍🔧 ใครรับงาน?", quickReply: { items: MECHANIC_NAMES.map(name => ({ type:"action", action:{ type:"postback", label:name, data:"action=select_name&jobId="+jobId+"&name="+name } })) } }] });
}

function askTimeQuickReply(replyToken, jobId, techName) {
  const minutesList = [15, 30, 60, 120, 180, 240, 300, 360, 480, 1440, 4320, 10080];
  replyFlex({ replyToken, messages: [{ type:"text", text:"⏱️ คาดว่าจะเสร็จ?", quickReply: { items: minutesList.map(m => ({ type:"action", action:{ type:"postback", label:formatMinuteLabel(m), data:"action=confirm_time&jobId="+jobId+"&name="+techName+"&min="+m } })) } }] });
}

function sendRepairCard(job) {
  pushLine({ to: TARGET_PUSH_ID, messages: [{ type:"flex", altText:"มีแจ้งซ่อมใหม่: "+job.topic, contents: createLineBubble(job) }] });
}

function createLineBubble(job) {
  const contents = [
    { type:"text", text:"🔧 แจ้งซ่อม", weight:"bold", size:"xl", wrap:true },
    { type:"text", text:"เลขงาน: "+job.jobId, size:"sm", color:"#666666", margin:"sm" },
    { type:"text", text:"สถานะ: "+job.status, size:"sm", color:"#666666" },
    { type:"separator", margin:"md" },
    { type:"text", text:"เวลา: "+job.time, wrap:true, margin:"md" },
    { type:"text", text:"แผนก: "+job.department, wrap:true },
    { type:"text", text:"ผู้แจ้ง: "+job.reporter, wrap:true },
    { type:"text", text:"หัวข้อ: "+job.topic, wrap:true },
    { type:"text", text:"รายละเอียด: "+job.detail, wrap:true }
  ];
  if (job.receiverName && job.receiverName !== "-") {
    contents.push({ type:"text", text:"👨‍🔧 ผู้รับงาน: "+job.receiverName, wrap:true, color:"#008000", weight:"bold", margin:"md" });
  }
  if (job.expectedTime && job.expectedTime !== "-") {
    contents.push({ type:"text", text:"⏳ คาดเสร็จ: "+job.expectedTime, wrap:true, weight:"bold", color:"#FF9800" });
  }
  if (job.status === "รองาน" && job.remark) {
    contents.push({ type:"text", text:"⚠️ หมายเหตุ: "+job.remark, wrap:true, color:"#E65100", weight:"bold", margin:"md" });
  }

  const footerButtons = [];
  if (job.status === "รอรับงาน") {
    footerButtons.push({ type:"button", style:"primary", color:"#06C755", action:{ type:"postback", label:"รับงาน", data:"action=accept&jobId="+job.jobId } });
  } else if (job.status === "กำลังดำเนินการ") {
    footerButtons.push({ type:"button", style:"primary", color:"#4285F4", action:{ type:"postback", label:"🏁 เสร็จงาน", data:"action=finish&jobId="+job.jobId } });
    footerButtons.push({ type:"button", style:"primary", color:"#FF9800", margin:"sm", action:{ type:"postback", label:"⏸️ หมายเหตุ", data:"action=note_start&jobId="+job.jobId } });
    footerButtons.push({ type:"button", style:"primary", color:"#D32F2F", margin:"sm", action:{ type:"postback", label:"❌ ยกเลิก", data:"action=cancel&jobId="+job.jobId } });
  } else if (job.status === "รองาน") {
    footerButtons.push({ type:"button", style:"primary", color:"#7B1FA2", action:{ type:"postback", label:"🔄 กลับมาทำต่อ", data:"action=resume&jobId="+job.jobId } });
    footerButtons.push({ type:"button", style:"primary", color:"#4285F4", margin:"sm", action:{ type:"postback", label:"🏁 เสร็จงาน", data:"action=finish&jobId="+job.jobId } });
  }

  return {
    type:"bubble",
    hero:{ type:"image", url:job.imageUrl||"https://placehold.co/800x520/png?text=No+Image", size:"full", aspectRatio:"20:13", aspectMode:"cover" },
    body:{ type:"box", layout:"vertical", spacing:"sm", contents:contents },
    footer: footerButtons.length>0 ? { type:"box", layout:"vertical", contents:footerButtons } : undefined
  };
}

function createLineBubbleWithNote(job) {
  return {
    type:"bubble",
    body:{ type:"box", layout:"vertical", spacing:"sm", contents:[
      { type:"text", text:"⏸️ เพิ่มหมายเหตุ", weight:"bold", size:"lg", wrap:true },
      { type:"text", text:"เลขงาน: "+job.jobId, size:"sm", color:"#666666" },
      { type:"separator", margin:"md" },
      { type:"text", text:"หัวข้อ: "+job.topic, wrap:true, margin:"md" },
      { type:"text", text:"ช่าง: "+job.receiverName, wrap:true, color:"#008000", weight:"bold" }
    ]},
    footer:{ type:"box", layout:"vertical", spacing:"sm", contents:[
      { type:"button", style:"primary", color:"#FF9800", action:{ type:"postback", label:"📝 หมายเหตุ/ติดขัด", data:"action=note_start&jobId="+job.jobId } },
      { type:"button", style:"primary", color:"#4285F4", margin:"sm", action:{ type:"postback", label:"🏁 เสร็จงาน", data:"action=finish&jobId="+job.jobId } }
    ]}
  };
}

function replyAcceptedCard(replyToken, jobId, displayName, expectedTime) {
  const contents = [
    { type:"text", text:"✅ กำลังดำเนินการ", weight:"bold", size:"xl", wrap:true },
    { type:"text", text:"เลขงาน: "+jobId, wrap:true },
    { type:"text", text:"👨‍🔧 ผู้รับงาน: "+displayName, wrap:true, color:"#008000", weight:"bold" }
  ];
  if (expectedTime && expectedTime !== "-") contents.push({ type:"text", text:"⏳ คาดเสร็จ: "+expectedTime, wrap:true, color:"#FF9800", weight:"bold" });
  replyFlex({ replyToken, messages: [{ type:"flex", altText:displayName+" รับงาน "+jobId+" แล้ว", contents:{
    type:"bubble",
    body:{ type:"box", layout:"vertical", spacing:"md", contents:contents },
    footer:{ type:"box", layout:"vertical", spacing:"sm", contents:[
      { type:"button", style:"primary", color:"#4285F4", action:{ type:"postback", label:"🏁 เสร็จงาน", data:"action=finish&jobId="+jobId } },
      { type:"button", style:"primary", color:"#FF9800", action:{ type:"postback", label:"⏸️ หมายเหตุ/ติดขัด", data:"action=note_start&jobId="+jobId } },
      { type:"button", style:"primary", color:"#D32F2F", action:{ type:"postback", label:"❌ ยกเลิก", data:"action=cancel&jobId="+jobId } }
    ]}
  } }] });
}

function replyJobCard(replyToken, jobId) {
  const sheet=getFormSheet(), values=sheet.getDataRange().getValues();
  for (let i=1; i<values.length; i++) {
    if (values[i][11]===jobId) { replyFlex({ replyToken, messages:[{ type:"flex", altText:"ข้อมูลงาน "+jobId, contents:createLineBubble(rowToJob(values[i])) }] }); return; }
  }
  replyLine(replyToken, "❌ ไม่พบเลขงาน "+jobId);
}

function replyAllPendingJobCards(replyToken) {
  const sheet=getFormSheet(), values=sheet.getDataRange().getValues(), bubbles=[];
  for (let i=1; i<values.length; i++) {
    const s=values[i][12];
    if (s==="รอรับงาน"||s==="กำลังดำเนินการ"||s==="รองาน") bubbles.push(createLineBubble(rowToJob(values[i])));
    if (bubbles.length>=12) break;
  }
  if (bubbles.length===0) { replyLine(replyToken,"✅ ไม่มีงานค้างในระบบ"); return; }
  replyFlex({ replyToken, messages:[{ type:"flex", altText:"รายการงานค้าง", contents:{ type:"carousel", contents:bubbles } }] });
}

function replyWaitingJobs(replyToken) {
  const sheet=getFormSheet(), values=sheet.getDataRange().getValues(), bubbles=[];
  for (let i=1; i<values.length; i++) {
    if (values[i][12]==="รอรับงาน") bubbles.push(createLineBubble(rowToJob(values[i])));
    if (bubbles.length>=12) break;
  }
  if (bubbles.length===0) { replyLine(replyToken,"✅ ไม่มีงานรอรับครับ"); return; }
  replyFlex({ replyToken, messages:[{ type:"flex", altText:"รายการงานรอรับ", contents:{ type:"carousel", contents:bubbles } }] });
}

function replyAcceptedJobs(replyToken) {
  const sheet=getFormSheet(), values=sheet.getDataRange().getValues(), bubbles=[];
  for (let i=1; i<values.length; i++) {
    const s=values[i][12];
    if (s==="กำลังดำเนินการ"||s==="รองาน") bubbles.push(createLineBubble(rowToJob(values[i])));
    if (bubbles.length>=12) break;
  }
  if (bubbles.length===0) { replyLine(replyToken,"✅ ไม่มีงานกำลังดำเนินการอยู่ครับ"); return; }
  replyFlex({ replyToken, messages:[{ type:"flex", altText:"รายการงานกำลังดำเนินการ", contents:{ type:"carousel", contents:bubbles } }] });
}

function replyNoteJobs(replyToken) {
  const sheet=getFormSheet(), values=sheet.getDataRange().getValues(), bubbles=[];
  for (let i=1; i<values.length; i++) {
    if (values[i][12]==="กำลังดำเนินการ") bubbles.push(createLineBubbleWithNote(rowToJob(values[i])));
    if (bubbles.length>=12) break;
  }
  if (bubbles.length===0) { replyLine(replyToken,"✅ ไม่มีงานที่กำลังดำเนินการอยู่ครับ"); return; }
  replyFlex({ replyToken, messages:[{ type:"flex", altText:"เพิ่มหมายเหตุ/รองาน", contents:{ type:"carousel", contents:bubbles } }] });
}

// ==========================================
// 7) ฐานข้อมูล
// ==========================================

function acceptJob(jobId, userId, displayName, expectedTime) {
  const sheet=getFormSheet();
  const cell=sheet.createTextFinder(jobId).matchEntireCell(true).findNext();
  if (!cell) return "notfound";
  const row=cell.getRow();
  if (sheet.getRange(row,13).getValue() !== "รอรับงาน") return "already";
  if (!expectedTime||expectedTime==="-") {
    const d=new Date(); d.setHours(d.getHours()+2);
    expectedTime=Utilities.formatDate(d,"GMT+7","dd/MM/yyyy HH:mm");
  }
  sheet.getRange(row,13).setValue("กำลังดำเนินการ");
  sheet.getRange(row,14).setValue(new Date());
  sheet.getRange(row,15).setValue(userId);
  sheet.getRange(row,16).setValue(displayName);
  sheet.getRange(row,17).setValue(expectedTime);
  for (let c=18;c<=22;c++) sheet.getRange(row,c).clearContent();
  SpreadsheetApp.flush();
  return "success";
}

function finishJob(jobId, sendQR) {
  const sheet=getFormSheet();
  const cell=sheet.createTextFinder(jobId).matchEntireCell(true).findNext();
  if (!cell) return "notfound";
  const row=cell.getRow();
  const currentStatus = sheet.getRange(row,13).getValue();
  if (currentStatus==="เสร็จงาน") return "already_finished";
  if (currentStatus!=="กำลังดำเนินการ" && currentStatus!=="รองาน") return "invalid_status";
  sheet.getRange(row,13).setValue("เสร็จงาน");
  sheet.getRange(row,18).setValue(new Date());
  SpreadsheetApp.flush();
  if (sendQR) {
    const topic        = sheet.getRange(row,6).getValue()  || jobId;
    const receiverName = sheet.getRange(row,16).getValue() || "ช่าง";
    sendRatingQR(jobId, topic, receiverName);
  }
  return "success";
}

function cancelJob(jobId, userId) {
  const sheet=getFormSheet();
  const cell=sheet.createTextFinder(jobId).matchEntireCell(true).findNext();
  if (!cell) return "notfound";
  const row=cell.getRow();
  const ownerId=sheet.getRange(row,15).getValue().toString();
  if (userId!=="APP" && ownerId!==userId) return "notowner";
  sheet.getRange(row,13).setValue("รอรับงาน");
  for (let c=14;c<=22;c++) sheet.getRange(row,c).clearContent();
  SpreadsheetApp.flush();
  return "success";
}

function setJobRemark(jobId, remark) {
  const sheet=getFormSheet();
  const cell=sheet.createTextFinder(jobId).matchEntireCell(true).findNext();
  if (!cell) return "notfound";
  const row=cell.getRow();
  if (sheet.getRange(row,13).getValue() !== "กำลังดำเนินการ") return "invalid_status";
  sheet.getRange(row,13).setValue("รองาน");
  sheet.getRange(row,22).setValue(remark);
  SpreadsheetApp.flush();
  return "success";
}

function resumeJob(jobId) {
  const sheet=getFormSheet();
  const cell=sheet.createTextFinder(jobId).matchEntireCell(true).findNext();
  if (!cell) return "notfound";
  const row=cell.getRow();
  if (sheet.getRange(row,13).getValue() !== "รองาน") return "invalid_status";
  sheet.getRange(row,13).setValue("กำลังดำเนินการ");
  sheet.getRange(row,22).clearContent();
  SpreadsheetApp.flush();
  return "success";
}

function cancelJobFromApp(jobId) { return cancelJob(jobId, "APP"); }

function getJobDataById(jobId) {
  const sheet=getFormSheet();
  const cell=sheet.createTextFinder(jobId).matchEntireCell(true).findNext();
  if (!cell) return null;
  return rowToJob(sheet.getRange(cell.getRow(),1,1,22).getDisplayValues()[0]);
}

// ==========================================
// 8) สร้างการ์ด Telegram
// ==========================================

function createTelegramCaption(job) {
  if (!job) return "ไม่พบข้อมูลงาน";
  let c = "🔧 <b>สถานะระบบคิวงานซ่อมบำรุง</b>\n━━━━━━━━━━━━━━━━━━\n";
  c += "<b>เลขงาน:</b> "+job.jobId+"\n";
  c += "<b>สถานะ:</b> "
    + (job.status==="รอรับงาน"       ? "⏳ รอช่างรับงาน"
    :  job.status==="กำลังดำเนินการ" ? "🛠️ กำลังดำเนินการ"
    :  job.status==="รองาน"          ? "⏸️ รองาน (ติดขัด)"
    :  job.status==="เสร็จงาน"       ? "🎉 เสร็จงานเรียบร้อย"
    :  job.status) + "\n";
  c += "━━━━━━━━━━━━━━━━━━\n";
  c += "<b>เวลาแจ้ง:</b> "+job.time+"\n<b>แผนก:</b> "+job.department+"\n";
  c += "<b>สถานที่:</b> "+job.location+"\n<b>หัวข้อ:</b> "+job.topic+"\n";
  c += "<b>รายละเอียด:</b> "+job.detail+"\n";
  if (job.receiverName&&job.receiverName!=="-") {
    c += "━━━━━━━━━━━━━━━━━━\n<b>👨‍🔧 ช่าง:</b> "+job.receiverName+"\n<b>⏳ คาดเสร็จ:</b> "+job.expectedTime+"\n";
  }
  if (job.remark) {
    c += "━━━━━━━━━━━━━━━━━━\n⚠️ <b>หมายเหตุ:</b> "+job.remark+"\n";
  }
  if (job.ratingSpeed||job.ratingQuality) {
    c += "━━━━━━━━━━━━━━━━━━\n";
    if (job.ratingSpeed)   c += "<b>⚡ ความเร็ว:</b> "+job.ratingSpeed+"/5\n";
    if (job.ratingQuality) c += "<b>⭐ คุณภาพ:</b> "+job.ratingQuality+"/5\n";
    if (job.comment)       c += "<b>💬 ความคิดเห็น:</b> "+job.comment+"\n";
  }
  return c;
}

function sendTelegramRepairCard(chatId, job) {
  const jNum = job.jobId.replace("JOB-","");
  let keyboard = [];
  if (job.status==="รอรับงาน") {
    keyboard = [[{ text:"✅ กดรับงานนี้", callback_data:"a=acc&j="+jNum }]];
  } else if (job.status==="กำลังดำเนินการ") {
    keyboard = [[
      { text:"🏁 เสร็จงาน",   callback_data:"a=fin&j="+jNum },
      { text:"⏸️ หมายเหตุ",  callback_data:"a=note&j="+jNum },
      { text:"❌ ยกเลิก",     callback_data:"a=can&j="+jNum }
    ]];
  } else if (job.status==="รองาน") {
    keyboard = [[
      { text:"🔄 กลับมาทำต่อ", callback_data:"a=resume&j="+jNum },
      { text:"🏁 เสร็จงาน",   callback_data:"a=fin&j="+jNum }
    ]];
  }
  UrlFetchApp.fetch("https://api.telegram.org/bot"+TELEGRAM_BOT_TOKEN+"/sendPhoto", {
    method:"post", contentType:"application/json",
    payload: JSON.stringify({
      chat_id: chatId,
      photo: job.imageUrl||"https://placehold.co/800x520/png?text=No+Image",
      caption: createTelegramCaption(job), parse_mode:"HTML",
      reply_markup: keyboard.length>0?{ inline_keyboard:keyboard }:undefined
    }),
    muteHttpExceptions: true
  });
}

// ==========================================
// 9) Utilities
// ==========================================

function rowToJob(row) {
  return {
    time: row[0]||"", department: row[1]||"", reporter: row[2]||"",
    location: row[3]||"-", machineNo: row[4]||"-", topic: row[5]||"",
    detail: row[6]||row[7]||row[8]||row[9]||"-",
    imageUrl: getDriveThumbnailUrl(row[10]||""),
    jobId: row[11]||"", status: row[12]||"",
    acceptTime: row[13]||"", receiverUserId: row[14]||"",
    receiverName: row[15]||"-", expectedTime: row[16]||"-",
    finishTime: row[17]||"",
    ratingSpeed:   row[18]||"",
    ratingQuality: row[19]||"",
    comment:       row[20]||"",
    remark:        row[21]||""
  };
}

function formatMinuteLabel(m) {
  if (m<60)     return m+" นาที";
  if (m<1440)   return (m/60)+" ชม.";
  if (m===1440) return "1 วัน";
  if (m===4320) return "3 วัน";
  return "7 วัน";
}

function getFormSheet() { return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(FORM_SHEET_NAME); }

function moveUploadedFile(fileUrl) {
  try {
    if (!fileUrl) return "";
    const fileId=extractDriveFileId(fileUrl); if (!fileId) return "";
    try { DriveApp.getFileById(fileId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
    return "https://drive.google.com/uc?export=view&id="+fileId;
  } catch(e) { return ""; }
}

function getDriveThumbnailUrl(fileUrl) {
  try { const id=extractDriveFileId(fileUrl); return id?"https://drive.google.com/uc?export=view&id="+id:""; } catch(e) { return ""; }
}

function extractDriveFileId(url) {
  let m=String(url||"").trim().match(/[?&]id=([^&]+)/); if(m) return m[1];
  m=String(url||"").trim().match(/\/d\/([^/]+)/);        if(m) return m[1];
  m=String(url||"").trim().match(/[-\w]{25,}/);          if(m) return m[0];
  return "";
}

function replyLine(replyToken, message) { replyFlex({ replyToken, messages:[{ type:"text", text:message }] }); }
function replyFlex(payload) { UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply",{ method:"post", headers:{ "Content-Type":"application/json","Authorization":"Bearer "+CHANNEL_ACCESS_TOKEN }, payload:JSON.stringify(payload), muteHttpExceptions:true }); }
function pushLine(payload)  { UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", { method:"post", headers:{ "Content-Type":"application/json","Authorization":"Bearer "+CHANNEL_ACCESS_TOKEN }, payload:JSON.stringify(payload), muteHttpExceptions:true }); }
function jsonResponse(obj)  { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
