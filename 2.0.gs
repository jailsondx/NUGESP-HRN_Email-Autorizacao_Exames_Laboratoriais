/* 
  
  ANOTAÇÕES
  
  NAVEGUE PELA PLANILHA SE GUIANDO COMO UMA MATRIZ:...
  VAR.FUNC(PARAMENTROS)[LINHA][COLUNA];
  
  ASPAS DUPLAS PARA TEXTO.
  SIMBOLO DE SOMA(+) PARA CONTATENAR.
  
*/



//Função que preenche a coluna Status
function DefineStatus(row){
  var status = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); //Ativa a Planilha/aba atual e faz a contagem matriz iniciar em 1 no lugar de 0
  var celula = status.getRange(row+1,18); //Define o intervalo STATUS
  celula.setValue("ENVIADO"); //Seta valor ENVIADO no intervalo definido
}


//Função de Alerta
function Alert(cont){
  var interface = SpreadsheetApp.getUi(); //variavel de interface
  
  if (cont > 0){
    interface.alert("SUCESSO! ENVIADO(S) " + cont + " NOVO(S) EMAIL(S)"); //modulo de alerta
  } else {
    interface.alert("NÃO HÁ NOVOS ATESTADOS CADASTRADOS"); //modulo de alerta
  }
}


//Função que preenche o email
function Email_Body(row, planilha){

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var nomeDAaba = ss.getName();
  var nomeABA = nomeDAaba;
  
  var assunto = "EXAME AUTORIZADO";
  var email = "EXAME AUTORIZADO\n \nSetor: " + nomeABA + "\nProntuario: " + planilha.getValues()[row][1] + "\nPaciente: " + planilha.getValues()[row][2] + "\nExame: " + planilha.getValues()[row][3];
  var email_address = "laboratorioprotocolosepse@gmail.com,hrnterceirizado@gmail.com,carlos.cac@isgh.org.br";
  
  //Função do GOOGLE API para o envio de email
  MailApp.sendEmail(email_address,assunto,email,{noReply:false});
}


//Função Principal
function Enviar_Email(){
  
  //Define a planilha ativa
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var planilha = sheet.getDataRange();
  
  //contador de execuções para função Alert();
  var cont = 0;
  var row = 6;
  
  for (row; planilha.getValues()[row][12] != ""; row++){
    if ((planilha.getValues()[row][17] == "") && (planilha.getValues()[row][12] == "AUTORIZADO")){
      Email_Body(row, planilha);
      DefineStatus(row);
      cont++;
    }
 }

  Alert(cont);
  
  SpreadsheetApp.flush();//Garante a execução do código ignorando possivel cache

}//Fecha Função Enviar_Email
