var sheet_id = "[....]; //replace with your sheet-ID (see readme)

function onFormSubmit(e) {
  var ss = SpreadsheetApp.openById(sheet_id);
  var formsheet = e.range.getSheet();
  var arrName = formsheet.getName();
  var submittedData = e.range.getValues();
  var configSheet = ss.getSheetByName("Config");
  var configData = configSheet.getDataRange().getValues();
  var answerRow = e.range.getRow()-1;
  console.log("Arrangement is called "+ arrName)
  console.log("Respondant is the "+ answerRow + "th respondant")


  var arrRow = -1;
  for (var i = 0; i < configData.length; i++) {
    if (configData[i][0] == arrName) {
      arrRow = i;
      console.log("Arrangement on row "+ arrRow)
      break;
    }
  }

    console.log("There are "+configData[arrRow][4]+" spots")

if (answerRow<=configData[arrRow][4]){ //Respondant got a spot
      console.log("Respondant got a spot")
      var name = submittedData[0][2];
      var email = submittedData[0][1];
      console.log("Name: "+name);
      console.log("Email: "+email);
      
      if (configData[arrRow][3]!=""){//There is the possibility to add alcohol
        // Generate HTML content from template file
      var template = HtmlService.createTemplateFromFile('email_plats_med_alkohol');
      }
      else if(configData[arrRow][1]==""){//The arrangement is free
      var template = HtmlService.createTemplateFromFile('email_plats_gratis');
      }
      else{//There is no possibility to add alcohol
        var template = HtmlService.createTemplateFromFile('email_plats');
      }
      template.name = name;
      template.kostnad = configData[arrRow][1];
      template.kostnad_mb = configData[arrRow][2]; ;
      template.kostnad_alkohol = configData[arrRow][3];
      template.infning = configData[arrRow][5];
      template.arr = arrName;

      var emailHtml = template.evaluate().getContent();

      var fromName = "[....]"; //Replace with sender name
      var emailFrom = "[....]"; //Replace with sender email
      var emailTo = email;

      // sends email as you
      MailApp.sendEmail({
        to: emailTo,
        subject: arrName,
        htmlBody: emailHtml,
        from: emailFrom,
        name: fromName,
      });
      formsheet.getRange(1 + answerRow, 12).setValue('The respondant have been given a spot and have been emailed');
}
else{
  var reservNumber = answerRow-configData[arrRow][4];
  console.log("The respondant did not get a spot")
      var name = submittedData[0][2]; 
      var email = submittedData[0][1];
      console.log("Name: "+name);
      console.log("Email: "+email);

      var template = HtmlService.createTemplateFromFile('email_reserv');

      template.name = name;
      template.arr = arrName;

      var emailHtml = template.evaluate().getContent();

       var fromName = "[....]"; //Replace with sender name
      var emailFrom = "[....]"; //Replace with sender email
      var emailTo = email;


      // sends email as you
      MailApp.sendEmail({
        to: emailTo,
        subject: arrName,
        htmlBody: emailHtml,
        from: emailFrom,
        name: fromName,
      });
      formsheet.getRange(1 + answerRow, 12).setValue("Respondant is respondant number "+reservNumber+" and has recieved an email");
}
}

function prepareForSale() {
  console.log("Preparing for ticket sale")
  var ss = SpreadsheetApp.openById(sheet_id);
  var arrSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var avgiftSheet = ss.getSheetByName("Mottagningsavgift");
  var avgiftData = avgiftSheet.getRange('B:B').getValues();
  var avgiftDataFlat = avgiftData.map(function(row) {return row[0];});
  var lastRow = arrSheet.getLastRow();
  var arrData = arrSheet.getRange(2,2, lastRow - 1, 2).getValues();
  console.log(arrData);

  for (var i = 0; i < arrData.length; i++) {
    if (avgiftDataFlat.indexOf(arrData[i][0].toString().toLowerCase()) != -1) {
      console.log(arrData[i][0])
      arrSheet.getRange(2+i, 11).setBackground("orange");
      }
  }

  arrSheet.getRange("K1").setValue("Paid");
  arrSheet.getRange("L1").setValue('=COUNTIF(K2:K; "Ja")')
  arrSheet.hideColumns(1, 1);
  arrSheet.hideColumns(4, 7);


}

function afterSale(){
  var arrSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  arrSheet.showColumns(1, arrSheet.getLastColumn());
}

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  menuEntries.push({ name: "Prepare sheet for sale", functionName: "prepareForSale" });
  menuEntries.push({ name: "Reset sheet after sale", functionName: "afterSale" }); 

  sheet.addMenu("Infochef", menuEntries);
}
  
