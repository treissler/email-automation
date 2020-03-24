/* Email-automation: логи, меню и тп */

function log(eventDescription){
  // в консоль
  Logger.log(eventDescription);
  
  // в таблицу
  var sheet = ss.getSheetByName("Логи");
  var lastRow = sheet.getDataRange().getLastRow() + 1;
  sheet.getRange(lastRow, 1, 1, 2).setValues([[eventDescription, new Date()]]).setHorizontalAlignment("left").setWrap(true);  
}

function errorEmail (settings){  
  var errorNotification = getNotificationMessage(); 
  
  GmailApp.sendEmail(settings.robotEmail, 
                     errorNotification.error.header,
                     errorNotification.error.body,
                     { name: settings.robotName, replyTo: adminEmail});
  
  log("Рассылка не сделана, причины: " + errors + " Администратор получил уведомление.");
}

//Пункт меню вверху таблицы
function menu() 
{
  var entries = [ {name: "Отправить", functionName: "main"}];
  ss.addMenu("РАССЫЛКА", entries);
}

function onOpen (e){
  menu();
}
