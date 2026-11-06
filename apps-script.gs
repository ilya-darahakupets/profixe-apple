/* ============================================================
   Profixe.by — приём заказов
   Google Apps Script → веб-приложение (doPost/doGet)

   Что делает:
   - doPost: сохраняет заказ в Google-таблицу (лист "orders")
             и отправляет уведомление в Telegram-бота;
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
const TG_CHAT  = "9205110270";    // ваш chat id (узнать у @userinfobot)
const SHEET_ID = "1eRdOZIUlvmp93XsmEHmIQM6ZGCCjHq9aS7b-kYGfVlA"; // ID Google-таблицы из URL (между /d/ и /edit)

function doPost(e){
  try{
    const o = JSON.parse(e.postData.contents);
    const sh = getSheet();
    sh.appendRow([o.id, o.date, o.name, o.phone, JSON.stringify(o.items||[]), o.total]);
    notifyTG(o);
  }catch(err){ /* отвечаем ok в любом случае */ }
  return ContentService.createTextOutput("ok");
}

function doGet(e){
  const cb = (e.parameter.callback || "cb").replace(/[^a-zA-Z0-9_$]/g,"");
  let orders = [];
  try{
    const sh = getSS().getSheetByName("orders");
    if(sh){
      orders = sh.getDataRange().getValues().slice(1).map(r=>({
        id:r[0], date:r[1], name:r[2], phone:r[3],
        items:JSON.parse(r[5]||"[]"), total:r[6]
      }));
    }
  }catch(err){}
  return ContentService
    .createTextOutput(cb+"("+JSON.stringify(orders)+")")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
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
    sh.appendRow(["id","date","name","phone","items","total"]);
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
