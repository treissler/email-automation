/* Функция работает раз в день по триггеру. Рассылается прайс из папки Прайс.*/
var err = "";
var days = ["вс", "пн","вт","ср","чт", "пт", "сб"];

function getPriceFile2 () {
  //получить имя ссылки
  var response = UrlFetchApp.fetch("http://obxt.ru/").getContentText(); 
  response = response.match(/media\/price\/[^"]*/); log(response);
  
  //проверка корректности
  if((!response)||(response[0].search("obht") < 0)){ 
    err = err.concat("Ошибка получения прайса"); 
    log(err); 
    return 0; 
  }
  
  response = UrlFetchApp.fetch("https://obxt.ru/".concat(response)).getBlob();
  return response;
}

function sendSelectedEmails() {
  //номер листа = день недели

  var td = new Date();
  td = td.getDay();
  
  //в пт, сб, вс не рассылаем
  if ((td == 0) || (td > 5)) { return 0; } //TODO в пятницу и пн не рассылаем
  
  //Где расположена база клиентов
  var base_sheet_name = days[td];
  
  //Настройки
  var settings = SpreadsheetApp.getActive().getSheetByName("Настройки");
  var adminMail = settings.getRange("B1").getValue();
  var adminName = settings.getRange("B2").getValue();
  var subject = settings.getRange("B3").getValue();
  
  //Получение файла
  var file = getPriceFile2(); //TODO предупредить если файл старый > 6 дней
  
  //Получение адресов клиентов
  var sheet = SpreadsheetApp.getActive().getSheetByName(base_sheet_name);
  var activeRange = sheet.getDataRange();
  var data = activeRange.getValues();
  var res = [];
  var count = 0;
  
  for (var i=2; i< data.length; i++) {
    var row = data[i];
    
    //Получение адреса
    var emailClient = row[0];
    
    //Получение сообщения
    var nameClient = row[2];
    var message_type = row[1];
    var message = getMessage(nameClient, message_type);
    
    //Отправка сообщения
    if ((emailClient)&&(file)&&(message)&&(subject)&&(adminName)&&(adminMail)) {
      if (i == 2) {
        //Уведомление о начале рассылки
        
        GmailApp.sendEmail(adminMail, "Начата рассылка gmail", "Сегодня рассылается лист " + td + " \nКаждый день рассылку с прайсом получает небольшая часть базы клиентов, база разбита по дням недели. Рассылки нет в пн, пт, сб, вс. Прайс скачивается с сайта obxt. Подробнее в файле: https://docs.google.com/spreadsheets/d/1NkEf2cdD5vhYZyokRKQUSqYZC2j3xKvtH3jfyRe-a_w/edit?usp=sharing , файл доступен после авторизации в почте gmail. Данные для авторизации можно запросить по почте", 
                           { name: adminName,
                           replyTo: adminMail,
                           attachments: [file], });
      }
      
      GmailApp.sendEmail(emailClient, subject, message, {
        name: adminName,
        replyTo: adminMail,
        attachments: [file],
      });
      row[4] = new Date();
      count++;
      
    } else { 
      row[4] = "ОШИБКА";
      err = err.concat ("Что-то не так: ").concat("emailClient ").concat(emailClient).concat(" file ").concat(file).concat(" message ").concat(message).concat(" subject ").concat(subject).concat(" adminName ").concat(adminName).concat(" adminMail ").concat(adminMail);
    }
    res.push([row[4]]);
  }
  sheet.getRange(3, 5, data.length - 2, 1).setValues(res);
  
  GmailApp.sendEmail(adminMail, "Закончена рассылка gmail", "Успешно выслано " + count + " писем из " + (data.length - 2 ) + " Ошибки: " + err, 
                     { name: adminName,
                       replyTo: adminMail });

}

function log(t){
  return Logger.log(t); 
}

function getMessage (nameClient, message_type) {  
  //Настройки
  var settings = SpreadsheetApp.getActive().getSheetByName("Настройки");
  var data = settings.getDataRange().getValues();
  var message_foot = data[8][1];
  var message = data[4][1];
  if (!nameClient) { 
    nameClient =  data[3][1]; 
  }
  
  for (var i = 3; i<8; i++) {
    if (message_type == data[i][0]) {
      message = data[i][1];
    }
  }
  return nameClient + message + message_foot;
}

function  getPriceFile () {
  var files = DriveApp.getFoldersByName("Прайс").next().getFiles();
  
  if(files.hasNext()) { 
    return files.next();
  } else { 
    Logger.log("no file"); 
    return false; 
  };
}
/* Меню ****************************************/
function menu() 
{
  var ss = SpreadsheetApp.getActive();
  var entries = [ {name: "Отправить рассылку", functionName: "sendSelectedEmails"}];
  ss.addMenu("БАЗА ХОЗТОРГА", entries);
  Logger.log("OK");
}

function onOpen (e){
  menu();
}

function isUnique() {
  
  var res = [];
  for (var i = 0; i<days.length; i++) {
    var ss;
    try { ss = SpreadsheetApp.getActive().getSheetByName(days[i]);} catch (err){}
    if (ss) {  
      var data = ss.getRange("A3:A").getValues();
      for (var i = 0; i< data.length; i++) {
        if (data[i][0]) {
        res.push(data[i][0]);
        }
      }
    }
  }
  var l1 = res.length;
  res = res.unique();
  var l2 = res.length;
  log(l1 + " "+ l2) ;
}

///////////////////////////////////////////////////////////
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
