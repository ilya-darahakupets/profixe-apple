/* ============================================================
   Profixi.by — приём заказов
   Google Apps Script → веб-приложение (doPost/doGet)

   Что делает:
   - doPost (заказ): сохраняет заказ в Google-таблицу (лист "orders")
             и отправляет уведомление в Telegram-бота;
   - doPost {action:"status"}: меняет статус заказа;
   - doPost {action:"clear"}: очищает все заказы;
   - doGet ?action=orders&callback=cb: отдаёт заказы как JSONP
             (чтобы страница #orders могла их показать).

   Настройка:
   1. Создайте Google-таблицу (sheets.new), скопируйте её ID из URL.
   2. Откройте https://script.google.com → Новый проект.
   3. Вставьте этот код, впишите SHEET_ID, TG_TOKEN и TG_CHAT.
   4. Развернуть → Новое развёртывание → Веб-приложение:
      - Исполнять от имени: Я
      - Кто имеет доступ: Все
   5. Скопируйте URL (.../exec) и вставьте в index.html → API_URL.
   ============================================================ */

const TG_TOKEN = "9205110270:AAEICi-GhWD-nyUHoykuuV0o1JkHfR8rZWI"; // токен от @BotFather
const TG_CHAT  = "259371486";    // ваш chat id (узнать у @userinfobot)
const SHEET_ID = "1eRdOZIUlvmp93XsmEHmIQM6ZGCCjHq9aS7b-kYGfVlA"; // ID Google-таблицы из URL (между /d/ и /edit)

const HEAD = ["id","date","name","phone","items","total","status"];
const STATUSES = ["Новый","В процессе","Выполнен","Отменен"];

function ok(){return ContentService.createTextOutput("ok")}

function doPost(e){
  try{
    const data = JSON.parse(e.postData.contents);
    const sh = getSheet();
    if(data.action==="clear"){
      sh.clearContents();
      sh.appendRow(HEAD);
      return ok();
    }
    if(data.action==="status"){
      setStatus(sh, data.id, data.status);
      return ok();
    }
    const o = data;
    sh.appendRow([o.id, o.date, o.name, o.phone, JSON.stringify(o.items||[]), o.total, o.status||"Новый"]);
    notifyTG(o);
  }catch(err){ /* отвечаем ok в любом случае */ }
  return ok();
}

function doGet(e){
  const cb = (e.parameter.callback || "cb").replace(/[^a-zA-Z0-9_$]/g,"");
  if(e.parameter.action==="logs"){
    let logs=[];
    try{
      const sh=getSS().getSheetByName("tglog");
      if(sh)logs=sh.getDataRange().getValues().slice(1);
    }catch(err){}
    return ContentService.createTextOutput(cb+"("+JSON.stringify(logs)+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  let orders = [];
  try{
    const sh = getSS().getSheetByName("orders");
    if(sh){
      orders = sh.getDataRange().getValues().slice(1).map(r=>({
        id:r[0], date:r[1], name:r[2], phone:r[3],
        items:safeJSON(r[4]), total:r[5], status:r[6]||"Новый"
      })).filter(o=>o.id!=null&&o.id!=="");
    }
  }catch(err){}
  return ContentService
    .createTextOutput(cb+"("+JSON.stringify(orders)+")")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function safeJSON(s){try{return JSON.parse(s||"[]")}catch(err){return[]}}

function setStatus(sh,id,status){
  if(STATUSES.indexOf(status)===-1)return;
  const vals=sh.getDataRange().getValues();
  for(let i=1;i<vals.length;i++){
    if(String(vals[i][0])===String(id)){sh.getRange(i+1,7).setValue(status);return}
  }
}

function getSS(){
  return (SHEET_ID && SHEET_ID.indexOf("ВСТАВЬТЕ")!==0)
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(){
  const ss = getSS();
  let sh = ss.getSheetByName("orders");
  if(!sh){
    sh = ss.insertSheet("orders");
    sh.appendRow(HEAD);
  }else if(sh.getRange(1,7).getValue()!=="status"){
    sh.getRange(1,7).setValue("status");
  }
  return sh;
}

function rubShort(n){return Math.round(n||0).toLocaleString("ru-RU")+" BYN"}

function notifyTG(o){
  try{
    if(!TG_TOKEN || TG_TOKEN.indexOf("ВСТАВЬТЕ")===0){logTG("no token");return}
    const chat = resolveChatId();
    if(!chat){logTG("no chat_id for order #"+o.id);return}
    const lines = [
      "🛒 Новый заказ #"+o.id, "",
      "👤 "+o.name,
      "📞 "+o.phone,
      "🕒 "+o.date, ""
    ];
    (o.items||[]).forEach(it=>{
      lines.push("• "+it.name+(it.opts?" ("+it.opts+")":"")+(it.price!=null?" — "+rubShort(it.price):""));
    });
    lines.push("", "💰 Итого: "+rubShort(o.total));
    const res = UrlFetchApp.fetch("https://api.telegram.org/bot"+TG_TOKEN+"/sendMessage", {
      method:"post",
      contentType:"application/json",
      payload:JSON.stringify({chat_id:chat, text:lines.join("\n")})
    });
    logTG("order #"+o.id+" chat="+chat+" → "+res.getResponseCode()+" "+res.getContentText().slice(0,300));
  }catch(err){logTG("ERR order #"+o.id+": "+err)}
}

function logTG(msg){
  try{
    const ss=getSS();
    let sh=ss.getSheetByName("tglog");
    if(!sh)sh=ss.insertSheet("tglog");
    sh.appendRow([new Date().toISOString(),msg]);
  }catch(e){}
}

/* chat_id: берём из TG_CHAT, а если там ID бота или пусто —
   находим первый чат, писавший боту, и запоминаем в свойствах скрипта */
function resolveChatId(){
  const botId = TG_TOKEN.split(":")[0];
  if(TG_CHAT && String(TG_CHAT)!==botId) return TG_CHAT;
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty("TG_CHAT_ID");
  if(saved) return saved;
  try{
    const res = UrlFetchApp.fetch("https://api.telegram.org/bot"+TG_TOKEN+"/getUpdates",{muteHttpExceptions:true});
    const data = JSON.parse(res.getContentText());
    const upd = (data.result||[]).find(u=>u.message&&u.message.chat);
    if(upd){
      const id = String(upd.message.chat.id);
      props.setProperty("TG_CHAT_ID", id);
      return id;
    }
  }catch(err){}
  return null;
}
