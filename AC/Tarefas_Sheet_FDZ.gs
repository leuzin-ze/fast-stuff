function insertLine() {
  const activeSheet = SpreadsheetApp.getActiveSheet();
  let lastRow = activeSheet.getLastRow();
  let lastColumn = activeSheet.getLastColumn();
  let lastItem = activeSheet.getRange(lastRow, 1, 1, lastColumn);
  activeSheet.insertRowAfter(lastRow);
  let newRow = activeSheet.getRange(lastRow+1, 1, 1, lastColumn);
  lastItem.copyTo(newRow, {contentsOnly:false});
  newRow.clearContent();
  activeSheet.getRange(lastRow+1, 1).setValue('**Choose**');
  activeSheet.getRange(lastRow+1, 4).setValue('**Choose**');
  activeSheet.getRange(lastRow+1, 5).setValue('**Choose**');
  activeSheet.getRange(lastRow+1, 10).setValue('NOT_SENT');
  activeSheet.getRange(lastRow+1, 11).setValue('NOT_SENT');
  Logger.log(lastItem.getValues());
  let formulas = lastItem.getFormulasR1C1();
  Logger.log(formulas);
  newRow.setFormulasR1C1(formulas);
}

function analyzeMail(){
  const activeSheet = SpreadsheetApp.getActiveSheet();
  var sheetData = activeSheet.getDataRange().getValues();
  const ui = SpreadsheetApp.getUi();

  var orderToSend = {};
  var paymentToSend = {};

  sheetData.forEach((row, index) => {
    if(index < 4) return;
    var orderEmailOption = row[9];
    var paymentEmailOption = row[10];
    switch(orderEmailOption){
      case 'SEND_EMAIL':
        if(Object.keys(orderToSend).includes(row[3])){
          orderToSend[row[3]].push((index+1));
        }else{
          orderToSend[row[3]] = [(index+1)];
        }
        break;
      default:
        break;
    }
    switch(paymentEmailOption){
      case 'SEND_EMAIL':
        if(Object.keys(paymentToSend).includes(row[3])){
          paymentToSend[row[3]].push((index+1));
        }else{
          paymentToSend[row[3]] = [(index+1)];
        }
        break;
      default:
        break;
    }
  });

  Logger.log(paymentToSend);
  Logger.log(orderToSend);

  var collabOrder = Object.keys(orderToSend);
  var collabOrderString = collabOrder.join(', \n');

  var collabPayment = Object.keys(paymentToSend);
  var collabPaymentString = collabPayment.join(', \n');
  
  if(collabOrder.length >= 1){
    var userResOrder = ui.alert('Deseja mandar um email de solicitação de pedido para: \n' + collabOrderString + ' ?', ui.ButtonSet.YES_NO);
    if(userResOrder == ui.Button.YES){
      sendOrderMail(orderToSend);
    }
  }
  if(collabPayment.length >= 1){
    var userResPayment = ui.alert('Deseja mandar um email de solicitação de pagamento para: \n' + collabPaymentString + ' ?', ui.ButtonSet.YES_NO);
    if(userResPayment == ui.Button.YES){
      sendPaymentMail(paymentToSend);
    }
  }
}

function sendOrderMail(orderDict){
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const collabSheet = SpreadsheetApp.openById('1_LvxngJSs-pNB0nfW14t3AwPiHAktY49OQpQTTDMBy0').getSheetByName('Worksheet');
  var collabData = collabSheet.getDataRange().getValues();
  var collabList = Object.keys(orderDict);
  for(var i = 0; i < collabList.length; i++){
    Logger.log(collabList[i]);
    var collabMail = ''
    collabData.forEach((row, index) => {
    if(index == 0) return;
    if(row[0] == collabList[i]){
      collabMail = row[1];
      Logger.log(collabMail);
    }
  });
    var rowsTask = orderDict[collabList[i]];
    var stringsToSend = []
    for(var k = 0; k < rowsTask.length; k++){
      var dataD = getDate().DataD;
      var task = activeSheet.getRange(rowsTask[k], 1, 1, 9).getValues();
      var stringTask = 
      `<br>
Data: ${dataD} <br>
Task: ${task[0][0]} <br>
Quantidade: ${task[0][1]} <br>
Descrição: ${task[0][2]} <br>
Gera Nota: ${task[0][4]} <br> <br>
<div style="font-weight:600;">
Valor Total: R$ ${task[0][8]},00
</div> <br>`
      Logger.log(stringTask);
      var dataF = getDate().DateF
      var userMail = Session.getActiveUser().getEmail();
      activeSheet.getRange(rowsTask[k], 12).setValue(dataF + '\n' + userMail);
      activeSheet.getRange(rowsTask[k], 10).setValue('ALREADY_SENT');
      stringsToSend.push(stringTask);
    }
    var mailData = {'name': collabList[i], 'tasks': stringsToSend.join('\n \n')};
    MailApp.sendEmail({
      to: collabMail,
      subject: "Solicitação de pedido para colaboração em JOB",
      htmlBody: `<style>
    @import url('https://fonts.googleapis.com/css2?family=Work+Sans:wght@400;700&display=swap');
  </style>
    <h3 style="font-family: 'Work Sans';
    font-style: normal;
    font-weight: 500;
    font-size: 18px;
    line-height: 18px;
    color: #000000;">
      Prezado, ${mailData.name} !
    </h3>
    <p style="font-style: normal;
    font-weight: 400;
    font-size: 15px;
    line-height: 15px;
    color: #000000;">
      Segue a proposta de tarefa para a Fastdezine:
      <br>
      ${mailData.tasks}
      <br>
      Por favor confirme o recebimento e aceite da tarefa.
      <br>
      Atenciosamente,<br>
      Time Fastdezine.
    </p>    `
    });
  }
}
function sendPaymentMail(paymentDict){
  const activeSheet = SpreadsheetApp.getActiveSheet();
  var collabList = Object.keys(paymentDict);
  for(var i = 0; i < collabList.length; i++){
    Logger.log(collabList[i]);
    var rowsTask = paymentDict[collabList[i]];
    var stringsToSend = [];
    for(var k = 0; k < rowsTask.length; k++){
      var dataD = getDate().DataD
      var task = activeSheet.getRange(rowsTask[k], 1, 1, 9).getValues();
      var stringTask = 
      `<br>
Data: ${dataD} <br>
Task: ${task[0][0]} <br>
Quantidade: ${task[0][1]} <br>
Descrição: ${task[0][2]} <br>
Gera Nota: ${task[0][4]} <br> <br>
<div style="font-weight:600;">
Valor Total: R$ ${task[0][8]},00
</div> <br>`
      Logger.log(stringTask);
      var dataF = getDate().DateF
      var userMail = Session.getActiveUser().getEmail();
      activeSheet.getRange(rowsTask[k], 13).setValue(dataF + '\n' + userMail);
      activeSheet.getRange(rowsTask[k], 11).setValue('ALREADY_SENT');
      stringsToSend.push(stringTask);
    }
    var mailData = {'name': collabList[i], 'tasks': stringsToSend.join('\n \n')};
    var recipientEmail = 'accounting@fastdezine.com';
    MailApp.sendEmail({
      to: recipientEmail,
      subject: "Solicitação de pagamento para colaborador em job da Fastdezine",
      htmlBody: `<style>
    @import url('https://fonts.googleapis.com/css2?family=Work+Sans:wght@400;700&display=swap');
  </style>
    <h3 style="font-family: 'Work Sans';
    font-style: normal;
    font-weight: 500;
    font-size: 18px;
    line-height: 18px;
    color: #000000;">
      Prezado Accounting,
    </h3>
    <p style="font-style: normal;
    font-weight: 400;
    font-size: 15px;
    line-height: 15px;
    color: #000000;">
      O ${mailData.name} finalizou a entrega da tarefa listada abaixo com sucesso para a Fastdezine:
      <br>
      ${mailData.tasks}
      <br>
      Por favor agendar seu pagamento.
      <br>
      Atenciosamente,<br>
      Time Fastdezine.
    </p>    `
    });
  }
}

function getDate(){
  const date = new Date;
  var year = date.getUTCFullYear();
  var dia = parseInt(date.getUTCDate());
  var mes = parseInt(date.getUTCMonth()) + 1;
  var hora = parseInt(date.getUTCHours()) - 3;
  var minuto = date.getUTCMinutes();
  var segundo = date.getUTCSeconds();

  if (String(dia).length === 1){
    dia = '0' + dia;
  }
  if (String(mes).length === 1){
    mes = '0' + mes;
  }
  if (String(minuto).length === 1){
    minuto = '0' + minuto;
  }
  if (String(segundo).length === 1){
    segundo = '0' + segundo;
  }
  if (String(hora).length === 1){
    hora = '0' + hora;
  }
  var dataF = year + '-' + mes + '-' + dia + ' / ' + hora + ':' + minuto + ':' + segundo;
  var dataD = dia + '-' + mes + '-' + year;
  let mesExtenso = '';

  switch(mes){
    case '01':
      mesExtenso = 'Jan';
      break;
    case '02':
      mesExtenso = 'Feb';
      break;
    case '03':
      mesExtenso = 'Mar';
      break;
    case '04':
      mesExtenso = 'Apr';
      break;
    case '05':
      mesExtenso = 'May';
      break;
    case '06':
      mesExtenso = 'Jun';
      break;
    case '07':
      mesExtenso = 'Jul';
      break;
    case '08':
      mesExtenso = 'Aug';
      break;
    case '09':
      mesExtenso = 'Sep';
      break;
    case '10':
      mesExtenso = 'Oct';
      break;
    case '11':
      mesExtenso = 'Nov';
      break;
    case '12':
      mesExtenso = 'Dec';
      break;
  }
  return {'DateF': dataF, 'DateExtenso': dia + '-' + mesExtenso + '-' + year, 'Year': year, 'DataD': dataD};
}
