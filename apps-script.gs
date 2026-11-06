/* ============================================================
   Profixe.by — приём заказов
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
const TG_CHAT  = "1008169639";    // ваш chat id (узнать у @userinfobot)
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
  if(!TG_TOKEN || TG_TOKEN.indexOf("ВСТАВЬТЕ")===0) return;
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
  UrlFetchApp.fetch("https://api.telegram.org/bot"+TG_TOKEN+"/sendMessage", {
    method:"post",
    contentType:"application/json",
    payload:JSON.stringify({chat_id:TG_CHAT, text:lines.join("\n")})
  });
}
