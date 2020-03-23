var fileFolderName = "Прайс";
var filePageUrl = "https://obxt.ru/";

var ss = SpreadsheetApp.getActive(); //ss - стандартное название для этой среды
var adminEmail = "admin@obxt.ru";
var adminName = "Администратор";

var errors = "";
var colPosition = {};

function log(event){
  // в консоль
  Logger.log(event);
  
  // в таблицу, чтобы видел пользователь
  var sheet = ss.getSheetByName("Логи");
  var lastRow = sheet.getDataRange().getLastRow() + 1; //нумерация в таблице и в коде отличается
  sheet.getRange(lastRow, 1, 1, 2).setValues([[event, new Date()]]).setHorizontalAlignment("left").setWrap(true);  
}

function errorEmail (settings){  
  log(errors);
  //когда совсем не получилось, письмо админу
  var errorNotification = getNotificationMessage();  
  GmailApp.sendEmail(settings.robotEmail, 
                     errorNotification.error.header,
                     errorNotification.error.body,
                     { name: settings.robotName, replyTo: adminEmail}
                    );
}

//Пункт меню вверху таблицы
function menu() 
{
  var entries = [ {name: "Отправить рассылку", functionName: "main"}];
  ss.addMenu("РАССЫЛКА", entries);
}

function onOpen (e){
  menu();
}

//TODO проверка уникальности списка email
Array.prototype.unique = function() 
{
  var n = {},r=[];
  for(var i = 0; i < this.length; i++) 
  {
    if (!n[this[i]]) 
    {
      n[this[i]] = true; 
      r.push(this[i]); 
    }
  }
  return r;
}
