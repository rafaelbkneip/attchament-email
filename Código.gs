function myFunction() {
  function myFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  var professores = spreadsheet.getSheetByName('Name').getRange("AB:AB").removeDuplicates();
  var lista_professores = spreadsheet.getSheetByName('Name').getRange("AB:AB").getValues();
  var range = spreadsheet.getSheetByName('Name').getLastRow();

  Logger.log(range)

  var i = 0

  for (i; i <= range; i++){
    Logger.log(lista_professores[i])
    nome = lista_professores[i].toString()
    variavel = spreadsheet.getSheetByName('Página').getRange('B1').setValue(nome)
    SpreadsheetApp.flush()
    Utilities.sleep(10000)
    planilha = spreadsheet.getSheetByName('Página');
    url = spreadsheet.getUrl().replace("edit", "export");

    Logger.log(url)

    email = spreadsheet.getSheetByName('Página').getRange('B2').getValue();

    Logger.log(email)
    
    lr = spreadsheet.getSheetByName('Página').getRange('G1').getValue();

    ultima_columa = Number(lr) + 1
    
    const url2 = url + "?format=pdf&portrait=true&size=A4&gridlines=false&gid=1738435925&top_margin=0.5&bottom_margin=0.25&left_margin=0.25&right_margin=0.25&right_margin=0.25&r1=0&c1=0&r2=" + ultima_columa + "&c2=9"; 
    Logger.log(url2)

    Utilities.sleep(10000)

    var response = UrlFetchApp.fetch(url2, {
      muteHttpExceptions: true,
      headers: {
        Authorization: 'Bearer ' +  ScriptApp.getOAuthToken(),
      },
    });
 
    Utilities.sleep(10000)

    var blob = response.getBlob();
    Logger.log("Content type: " + blob.getContentType());
    Logger.log("File size in MB: " + blob.getBytes().length / 1000000);
    var file = DriveApp.createFile(blob);
    Logger.log(file.getUrl()); 

    var link = file.getId()

    DriveApp.getFolderById(link).setName(lista_professores[i])
    
    
    var emailTemp = HtmlService.createTemplateFromFile("template3");
    var name = "Central de Serviços Pedagógicos"
    var subject = "Pedidos de questão - 2º semestre de 2023"
    
    var htmlMessage = emailTemp.evaluate().getContent()
    var file = DriveApp.getFileById(link)
  

      GmailApp.sendEmail(
        email,
        subject,
        " ",
        {name: name, htmlBody: htmlMessage, attachments: [file]}
      ) 
  }
}
}
