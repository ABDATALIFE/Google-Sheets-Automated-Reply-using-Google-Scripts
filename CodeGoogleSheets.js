Skip to content
Search or jump to…
Pull requests

 
function emailAlert() {
  // today's date information
  var today = new Date();
  var todayMonth = today.getMonth() + 1;
  var todayDay = today.getDate();
  var todayYear = today.getFullYear();

  // 8 days from now
  var eightDaysFromToday = new Date();
  eightDaysFromToday.setDate(eightDaysFromToday.getDate() + 8);
  var eightDaysMonth = eightDaysFromToday.getMonth() + 1;
  var eightDaysDay = eightDaysFromToday.getDate();
  var eightDaysyear = eightDaysFromToday.getFullYear();



  


  // getting data from spreadsheet
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 5; // First row of data to process
  var numRows = 180; // Number of rows to process
  

  var dataRange = sheet.getRange(startRow, 1, numRows, 999);
  var data = dataRange.getValues();

  //looping through all of the rows
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];

    var daysLeft = row[11];


  

    var expireDateFormat = Utilities.formatDate(
      new Date(row[10]),
      'ET',
      'MM/dd/yyyy'
    );

    var expireDateFormat2 = Utilities.formatDate(
      new Date(row[10]),
     'America/New_York', 'MMMM dd, yyyy'
    );

    
    
    var ccvariable = 'abasit30495@gmail.com'
   
   
    //var tovariable = 'info@santor.com'
    //var ccvariable = 'info@axiagps.com'

    var newdate = frenchDate(new Date(row[10]));

    var sublevel = row[5].split(",")[0];
    var sublevel2 = sublevel.split(":")[1]
    var sublevel3 = sublevel2+':'+sublevel.split(":")[2]

    // email information
    var subject = 'AXIA SERVICE PLAN GPS';
    var message =
      '**English Below**'+
      '\n' + 
      '\n' +
      'Bonjour' + ' ' + row[0]+','+
      '\n' +
      '\n'+
      "Votre plan de service GPS pour l'unité:"+' '+row[7]+ ' '+ 'expire le' + ' '+ newdate + '.'+
      '\n'+
      '\n'+ 
      'Voici le lien pour le renouvellement :'+ 
      '\n'+
      '\n'+
      row[5].split(":")[0] +
      '\n' +
      '\n'+
      '3 mois:' + '  ' + sublevel3 + 
      '\n' +
      '\n'+
      '1 an:' + '  ' + row[5].split(",")[1] + 
      '\n' + 
      '\n'+ 
      "⚠️ Pour que votre appareil fonctionne correctement, vous devez obtenir un plan de service ,sinon, votre GPS sera désactivé et des frais de 35 $ + taxes peuvent s'appliquer pour réactiver le service suspendu"+
      '\n'+ 
      '\n' + 
      '\n'+
      "Si vous avez des questions, n'hésitez à nous contacter!"+
      '\n' + 
      '\n'+
      'merci'+
      '\n'+
      '\n'+
      'Équipe AXIA' +
      '\n' + 
      '\n' +
      '**************************************' + 
      '\n' +
      '\n' + 
      'Hello' + ' ' + row[0] +','+
      '\n' +
      '\n'+
      'Your GPS service plan for unit:' + ' '+row[7]+ ' '+ 'expires on' + ' '+ expireDateFormat2 +'.' +
      '\n'+
      '\n'+
      'Here is the link for the renewal:'+ 
      '\n'+
      '\n'+
      row[5].split(":")[0] +
      '\n'+
      '\n'+
      '3 months:' + ' ' + sublevel3 + 
      '\n'+
      '\n'+
      '1 year:' + ' ' + row[5].split(",")[1] + 
      '\n'+ 
      '\n'+
      '\n'+
      "⚠️ For your device to work properly, you need to get a service plan; otherwise, your GPS will be disabled and a fee of $35 + tax may apply to reactivate suspended service." + 
      '\n' + 
      '\n'+
      '\n'+
      "If you have any questions, do not hesitate to contact us!"+
      '\n' +
      '\n' +
      'Thank you'+
      '\n'+
      'Axia Team'
      ;

    //expiration date information
    var expireDateMonth = new Date(row[10]).getMonth()+1;
    var expireDateDay = new Date(row[10]).getDate();
    var expireDateYear = new Date(row[10]).getFullYear();
    var cell = sheet.getRange(i+5,13);
    //checking for today
    if (
      expireDateMonth === todayMonth &&
      expireDateDay === todayDay &&
      expireDateYear === todayYear &&
      row[6] != ""
    ) {
    
      MailApp.sendEmail(row[6], subject, message);
      Logger.log('todayyyy!');
    }

    //checking for 8 Days from now
    Logger.log('8 Days, expire ' + ' '+eightDaysMonth + ' '+eightDaysyear);
    if (
      daysLeft > 0 &&
      daysLeft <= 8  &&
      row[6] != "" 
    ) {
      
      cell.setValue("The Email is Sent");
      MailApp.sendEmail(row[6], subject, message,{cc:ccvariable});
      Logger.log(newdate);
      
      
    }
    else{
      cell.setValue("");
    }


  }
}

function frenchDate(date) {
  var month = ['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'];
  var day = ['dimanche','lundi','mardi','mercredi','jeudi','vendredi','samedi'];
  var m = month[date.getMonth()];
  var d = day[date.getDay()];
  var dateStringFr = d+' '+date.getDate()+' '+m+' '+date.getFullYear();
  return dateStringFr
}

