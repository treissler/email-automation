/* Работает 1 раз в день по внешнему триггеру на main(). Рассылается файл прайс-листа с сайта. */

function downloadFile () {
  try{
    log("Загрузка страницы " + filePageUrl);
    var response = UrlFetchApp.fetch(filePageUrl).getContentText(); //null или string[]
  } catch (e) {
    errors += "Ошибка загрузки страницы."; 
    return null;
  }
  
  var fileName = response.match(/media\/price\/[^"]*/); //null или string[]  
  
  if((!fileName)||(fileName[0].search("obht") < 0)){ 
    errors += "\nФайл не найден на странице. Ожидается имя вида /media/price/obht_price_2020-12-31(1).xls"; 
    return null; 
  }
  
  var file = UrlFetchApp.fetch(filePageUrl + fileName).getBlob();
  errors += "\nФайл успешно загружен.";
  
  return file;
}

//TODO если сервер с файлом не ответил
function  getFile () {
  //сначала новые, по умолчанию
  var files = DriveApp.getFoldersByName(fileFolderName).next().getFiles();
  
  if(files.hasNext()) { 
    return files.next();
  } else { 
    log("Ошибка: в папке на Диске не найден файл для рассылки."); 
    return null; 
  };
}

function setColumns(tableHeader){
  var colName = {
    "Email(*)": "email",
    "Шаблон письма(*)": "template", 
    "Обращение": "recipientName", 
    "Название компании": "company",
    "Дата отправки": "dateSent",
    "Комментарии для менеджера" : "comment"
  };
  
  for (var i=0; i<tableHeader.length; i++){
    if (!colName[tableHeader[i]]){ 
      errors += "\nПроверьте, что все колонки есть на листе: " + Object.keys(colName);
      return false;
    }
    colPosition[colName[tableHeader[i]]] = i;
  }
  return true;
}

function getUserSettings () {  
  var userSettings = ss.getSheetByName("Настройки").getDataRange().getValues();
  
  var fieldsName = {
    "Email компании (replyTo)": "robotEmail",
    "Название почтового робота": "robotName",
    "Заголовок письма": "subject", 
    "Обращение по умолчанию": "defaultName", 
    "Текст1": "template1",
    "Текст2": "template2",
    "Подпись к письму": "footer"
  };  
  
  var fields = {};
  for (var i=0; i<userSettings.length; i++) {
    var name = userSettings[i][0];
    var val = userSettings[i][1];
    fields[fieldsName[name]] = val;             
  }
  return fields;
}

function getMessageBody (recipientName, messageType, settings) {
  if (!recipientName) { 
    recipientName = settings.defaultName; 
  }
  return recipientName + settings["template" + messageType] + settings.footer;
}

function getNotificationMessage(todaySheetName, sentCount, recipientsCount) {
  var messages = {
    mailing_start: {
      header: "Начата рассылка прайс-листа",
      body: "Сегодня получатели рассылки находятся на листе " 
      + todaySheetName
      + ".\nКаждый день рассылку с прайс-листом получает небольшая группа клиентов, база распределена по дням недели."
      + "\nРассылки нет в пт, сб, вс."
      + "\nПрайс-лист скачивается с сайта obxt.ru."
      + "\nВопросы администратору рассылки: " 
      + adminEmail, 
    },
    mailing_end:  {
      header: "Закончена рассылка прайс-листа",
      body:"Успешно выслано " + sentCount + " писем из " + recipientsCount + "\n\nЗамечания: " + errors      
    },
    error:  {
      header: "Неудачная попытка рассылки прайс-листа",
      body:"Рассылка не начата. Ошибки: " + errors      
    }
  }
  return messages;
}

function main() {
  
  var settings = getUserSettings();
  var file = downloadFile();
  
  if (!file) {
    errors += "\nНе удалось получить файл";
    errorEmail(settings);
    return 0;
  }
  
  var days = [{name: "вс", active: false}, 
              {name: "пн", active: true},
              {name: "вт", active: true},
              {name: "ср", active: true},
              {name: "чт", active: true}, 
              {name: "пт", active: false}, 
              {name: "сб", active: false}
             ];
  var today = (new Date()).getDay(); // 0 = вс, 6 = сб
  
  var startNotification = getNotificationMessage(days[today].name); //можно передавать не все параметры
  
  if (!days[today].active) {
    errors += "\nНельзя рассылать в этот день: " + days[today].name;
    errorEmail (settings);
    return 0;
  }
  
  GmailApp.sendEmail(settings.robotEmail, 
                     startNotification.mailing_start.header,
                     startNotification.mailing_start.body,
                     { name: settings.robotName, replyTo: adminEmail, attachments: [file]});
  
  var todaySheetName = days[today].name;
  var recipientsData = ss.getSheetByName(todaySheetName).getDataRange().getValues();
  var tableHeader = recipientsData[1];
  var haveAllColumns = setColumns(tableHeader);
  
  if (!haveAllColumns) {
    errorEmail(settings);
    return 0;
  }
  
  var subject = settings.subject;
  if (subject.length < 10){
    errors += "Тема письма слишком короткая.";
    errorEmail(settings);
    return 0;
  }
  
  //isEmailsUnique();
  
  var dateSent = [];
  var sentCount = 0;
  
  for (var i=2; i<recipientsData.length; i++) {
    var row = recipientsData[i];
    
    var email = row[colPosition.email];
    var template = row[colPosition.template];
    var messageType = template.replace("Текст","");
    var name = row[colPosition.recipientName];
    //var company = row[colPosition.company];
    var message = getMessageBody(name, messageType, settings);
    
    if ((!email)||(!message)) {
      dateSent.push(["Не отправлено, " + new Date()]);
      errors += "\nНе получилось отправить адресату из строки " + i + " на листе " + days[today].name;
      continue;
    }
      
    GmailApp.sendEmail(email, subject, message, {
      name: settings.robotName,
      replyTo: settings.robotEmail,
      attachments: [file],
    });
    
    dateSent.push([new Date()]);
    sentCount++;
  }
  //Фиксация результата
  ss.getSheetByName(todaySheetName).getRange(3, colPosition.dateSent + 1, recipientsData.length - 2, 1).setValues(dateSent);
  
  var successNotification = getNotificationMessage(days[today].name, sentCount, (recipientsData.length - 2 ));
  
  GmailApp.sendEmail(settings.robotEmail, 
                     successNotification.mailing_end.header,
                     successNotification.mailing_end.body,
                     { name: settings.robotName, replyTo: adminEmail});
}
