/*
Forma incluir nombre, apellido, direccion de correo, mensaje personalizado
Envia correo HTML
*/
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var range= sheet.getDataRange();
  var values = range.getValues();
  var headers = values.shift();
  var lastColumn = range.getLastColumn();
  var lastRow = range.getLastRow();
var searchRange = sheet.getRange(1,1, lastRow-1, lastColumn-1);
function getParticipantData() {
  for ( i = 0; i < values.length; i++){
    var name  = values[i][0];
    var surname  = values[i][1];
    var email  = values[i][2];
    var sex  = values[i][3];
    //Logger.log(name+","+email+","+sex);
  };
  return values;
}

function main() {
 for ( i = 0; i < values.length; i++){
    var nombre  = values[i][0];
    var apellido  = values[i][1];
    var correo  = values[i][2];
    var sexo  = values[i][3];
    }
  if(sexo=="M")
  { saludo = "o";
  } else {
    saludo = "a";
  };
  var mainEmail = "\""+nombre+" "+apellido+"\""+" <"+correo+">";
  var subject = "Invitaci?n a las V Jornadas de Investigación de UNIEDPA  - 18 y 19 de Octubre de 2018";
  htmlBody = buildMessage(saludo,nombre);
  textBody = buildTextMessage(saludo,nombre);
  Logger.log(textBody);
  sendEmail(mainEmail,subject,saludo, htmlBody, textBody);
};

function sendEmail(mainEmail,subject,saludo, htmlBody, textBody) {
//Logger.log(values[0][1]);
 //var blob = Utilities.newBlob('Insert any HTML content here', 'text/html', 'my_document.html');
 MailApp.sendEmail(mainEmail, subject,textBody, {
    htmlBody: htmlBody,
    name: 'Comisión de Investigación y Postgrado de UNIEDPA'
 });
}

function buildMessage(saludo,nombre) {
  var plantillaId = "10l0CheLgyTEcnrRMcAx834wmrhKqa0j60cgM6l_ZRd0";
  //Esta plantilla esta en carpeta Jornadas
  var plantillaDoc = DocumentApp.openById(plantillaId);
  var plantillaBodyTxt = plantillaDoc.getBody().getText();
  plantillaDoc.saveAndClose();
  var MailMessageFileName = Math.random().toString(36);
  var MailMessageFile = DocumentApp.create(MailMessageFileName);
  var MailMessageId = MailMessageFile.getId();
  var MailMessageDoc = DocumentApp.openById(MailMessageId);
  var MailMessageBody = MailMessageDoc.getBody();
  var MailMessageBody = MailMessageBody.setText(plantillaBodyTxt);
  var MailMessageBody = MailMessageBody.replaceText("{{saludo}}", saludo);
  var MailMessageBody = MailMessageBody.replaceText("{{nombre}}", nombre);
  //Logger.log(MailMessageBody.getText());
  //return MailMessageBody.getText();
  MailMessageDoc.saveAndClose();
  MailMessageFile=DriveApp.getFileById(MailMessageId);
  MailMessageFile.setTrashed(true);
  Drive.Files.emptyTrash();
  return MailMessageBody.getText();
}

function buildTextMessage(saludo,nombre) {
  var plantillaId = "1Iz3rqvgwle2PowITQp-rRUKeVxdpn9Du3jBMnqHix3E";
  var plantillaDoc = DocumentApp.openById(plantillaId);
  var plantillaBodyTxt = plantillaDoc.getBody().getText();
  plantillaDoc.saveAndClose();
  var MailMessageFileName = Math.random().toString(36);
  var MailMessageFile = DocumentApp.create(MailMessageFileName);
  var MailMessageId = MailMessageFile.getId();
  var MailMessageDoc = DocumentApp.openById(MailMessageId);
  var MailMessageBody = MailMessageDoc.getBody();
  var MailMessageBody = MailMessageBody.setText(plantillaBodyTxt);
  var MailMessageBody = MailMessageBody.replaceText("{{saludo}}", saludo);
  var MailMessageBody = MailMessageBody.replaceText("{{nombre}}", nombre);
  //Logger.log(MailMessageBody.getText());
  //return MailMessageBody.getText();
  MailMessageDoc.saveAndClose();
  MailMessageFile=DriveApp.getFileById(MailMessageId);
  MailMessageFile.setTrashed(true);
  Drive.Files.emptyTrash();
  return MailMessageBody.getText();
}

function createHTMLDraftInGmail(asunto, remitente) {
  subject = "Invitación a las V Jornadas de Investigación de UNIEDPA - Octubre 2018";
  var htmlBody=buildMessage(values);
  var remitente = "Comisión de Investigación y Postgrado de UNIEDPA <investigacion@uniedpa.net>";
  var forScope = GmailApp.getInboxUnreadCount(); // needed for auth scope
  var raw = 'From:'+remitente+'\r\n' +
    'To: '+values[0][10]+'\r\n' +
            'Subject:'+asunto+'\r\n' +
            'Content-Type: text/html; charset=UTF-8\r\n' + 
            '\r\n' + htmlBody;

  var draftBody = Utilities.base64Encode(raw, Utilities.Charset.UTF_8).replace(/\//g,'_').replace(/\+/g,'-');

  var params = {
    method      : "post",
    contentType : "application/json",
    headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions:true,
    payload:JSON.stringify({
      "message": {
        "raw": draftBody
      }
    })
  };

  var resp = UrlFetchApp.fetch("https://www.googleapis.com/gmail/v1/users/me/drafts", params);
  Logger.log(resp.getContentText());
}
