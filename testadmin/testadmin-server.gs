// Сервер мониторинга тестирования — обновление 09.11.2026
// ВАЖНО: укажите ID таблицы, В КОТОРУЮ ПОСТУПАЮТ НОВЫЕ ТЕСТЫ.
// ID — часть URL между /d/ и /edit. Если скрипт привязан к нужной таблице,
// можно оставить пустую строку, но явный ID надёжнее.
var SPREADSHEET_ID = "";
var SHEET_NAME = "Ответы";
var ADMIN_KEY = "00000"; // Замените на свой надёжный пароль
var TIME_ZONE = "Asia/Almaty";
var PAGE_SIZE_MAX = 1000;

function getTargetSheet() {
  var ss = SPREADSHEET_ID.trim()
    ? SpreadsheetApp.openById(SPREADSHEET_ID.trim())
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Таблица не найдена. Заполните SPREADSHEET_ID.");
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Лист "' + SHEET_NAME + '" не найден в таблице "' + ss.getName() + '".');
  if (!sheet.getLastRow()) {
    sheet.appendRow(["Дата и время","Район / город","Тип комиссии","Номер комиссии","ФИО","Верно","Неверно","Ошибочные вопросы"]);
    sheet.getRange(1,1,1,8).setFontWeight("bold");
  }
  return sheet;
}
function dateText(value) {
  return value instanceof Date && !isNaN(value.getTime())
    ? Utilities.formatDate(value, TIME_ZONE, "dd.MM.yyyy HH:mm:ss")
    : String(value == null ? "" : value).trim();
}
function cleanRow(r) {
  var kind = String(r[2] || "").trim(), num = String(r[3] || "").trim();
  return {
    ts: dateText(r[0]),
    district: String(r[1] || ""),
    commission: (kind ? kind + " " : "") + (num ? "№" + num : ""),
    fio: String(r[4] || ""),
    ok: Number(r[5]) || 0,
    bad: Number(r[6]) || 0,
    wrong: String(r[7] || "—")
  };
}
function doGet(e) {
  var p = (e && e.parameter) || {}, action = p.action || "", callback = p.callback || "";
  try {
    var sheet = getTargetSheet();
    if (action === "check") {
      // Сохраняем прежний формат проверки, используемый страницей тестирования.
      var parts = String(p.key || "").split("|");
      var ts = parts[0] || "", fio = parts.slice(1).join("|");
      var last = sheet.getLastRow(), found = false;
      if (last > 1) {
        var data = sheet.getRange(2,1,last-1,5).getValues();
        for (var i=0; i<data.length; i++) {
          if (dateText(data[i][0]) === ts.trim() && String(data[i][4]).trim() === fio.trim()) {found=true; break;}
        }
      }
      return jsonpResponse({ok:true, found:found}, callback);
    }
    if (String(p.key || "") !== ADMIN_KEY)
      return jsonpResponse({ok:false,error:"forbidden"},callback);

    if (action === "health") {
      var ss = sheet.getParent(), lastRow = sheet.getLastRow();
      var latest = "", sample = [];
      if (lastRow > 1) {
        var from = Math.max(2,lastRow-99);
        var tail = sheet.getRange(from,1,lastRow-from+1,5).getValues();
        for (var x=tail.length-1;x>=0;x--) {
          if (tail[x][4]) {latest=dateText(tail[x][0]);break;}
        }
      }
      return jsonpResponse({ok:true,spreadsheetName:ss.getName(),spreadsheetId:ss.getId(),sheetName:sheet.getName(),lastRow:lastRow,latestRowDate:latest,checkedAt:Utilities.formatDate(new Date(),TIME_ZONE,"dd.MM.yyyy HH:mm:ss")},callback);
    }
    if (action === "list") {
      var total = Math.max(0,sheet.getLastRow()-1);
      var offset = Math.max(0,Number(p.offset)||0);
      var paged = p.limit !== undefined && String(p.limit).trim() !== "";
      var limit = paged ? Math.min(PAGE_SIZE_MAX,Math.max(1,Number(p.limit)||500)) : total;
      var count = Math.max(0,Math.min(limit,total-offset));
      var values = count ? sheet.getRange(offset+2,1,count,8).getValues() : [];
      var rows = [];
      for (var i=0;i<values.length;i++) if (values[i][4]) rows.push(cleanRow(values[i]));
      return jsonpResponse({ok:true,rows:rows,offset:offset,limit:limit,total:total,hasMore:offset+count<total},callback);
    }
    if (action === "delete") {
      var targetTs=String(p.ts||"").trim(), targetFio=String(p.fio||"").trim();
      var lock=LockService.getScriptLock();lock.waitLock(10000);
      try {
        var n=sheet.getLastRow(), data=n>1?sheet.getRange(2,1,n-1,5).getValues():[];
        for (var i=data.length-1;i>=0;i--) {
          if (dateText(data[i][0])===targetTs && String(data[i][4]||"").trim()===targetFio) {
            sheet.deleteRow(i+2);
            return jsonpResponse({ok:true},callback);
          }
        }
      } finally {lock.releaseLock();}
      return jsonpResponse({ok:false,error:"Запись не найдена в таблице"},callback);
    }
    return jsonpResponse({ok:false,error:"unknown action"},callback);
  } catch(err) {
    return jsonpResponse({ok:false,error:String(err.message||err)},callback);
  }
}
function doPost(e) {
  try {
    var item=JSON.parse((e.postData&&e.postData.contents)||"{}");
    var lock=LockService.getScriptLock();lock.waitLock(20000);
    try {
      getTargetSheet().appendRow([
        Utilities.formatDate(new Date(),TIME_ZONE,"dd.MM.yyyy HH:mm:ss"),
        item.district||"",item.kind||"",item.num||"",item.fio||"",
        Number(item.ok)||0,Number(item.bad)||0,item.wrong||"—"
      ]);
    } finally {lock.releaseLock();}
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err.message||err)})).setMimeType(ContentService.MimeType.JSON);
  }
}
function jsonpResponse(data, callback) {
  var out=JSON.stringify(data);
  if (callback) {
    if (!/^[A-Za-z_$][A-Za-z0-9_$]*$/.test(callback))
      return ContentService.createTextOutput(JSON.stringify({ok:false,error:"invalid callback"})).setMimeType(ContentService.MimeType.JSON);
    return ContentService.createTextOutput(callback+"("+out+");").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}
