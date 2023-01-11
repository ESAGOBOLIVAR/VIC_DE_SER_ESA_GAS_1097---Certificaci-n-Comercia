const postInformationToSheet = () => {
  const currentlyDate = Utilities.formatDate(new Date(), "GMT-05:00", "dd/MM/yyyy:HH:mm");
  const dataSheet = [getResponses()];
  const dataSheetMain = getDataSheet(idSheetMain,nameSheetMain)
  const sheetMain = SpreadsheetApp.openById(idSheetMain);
  let sheetName = sheetMain.getSheetByName(nameSheetMain);
  let lastRow = 1

  dataSheetMain.forEach(( iteratorData  ) => {
    if(iteratorData[0].length > 0){
      lastRow= lastRow +1;
    }
  })
  const secuence = generarFormato("CERT_COM_2023-", "", lastRow-14, 4);
  console.log(secuence,lastRow);
  dataSheet[0].unshift(secuence);
  const urlFile = `https://drive.google.com/file/d/${dataSheet[0][8]}/view?usp=sharing`
  if(dataSheet[0].length > 9){
      const newArray = dataSheet[0].slice(8,dataSheet[0].length).length;
      for(let i = 0; i < newArray;i++){
        dataSheet[0][i + 8] = `https://drive.google.com/file/d/${dataSheet[0][i + 8]}/view?usp=sharing`;
      }
  }
  dataSheet[0][8] = urlFile;
  sheetName.getRange(lastRow + 1,1,dataSheet.length,dataSheet[0].length).setValues(dataSheet);
  sheetName.getRange(`S${lastRow + 1}:S${lastRow + 1}`).setValue(currentlyDate);
  buildEmail(dataSheet);

};


const generarFormato = (textoInicial, textoFinal, identificador, cantidadCeros) => {
  return textoInicial + String(identificador).padStart(cantidadCeros, 0) + textoFinal;
}

const getResponses = () => {  

  let valuesResponses = [];
  const form = FormApp.openById(idForms);
  const formResponses = form.getResponses().pop().getItemResponses();
  const email = form.getResponses().pop().getRespondentEmail();     

  formResponses.forEach((itemResponse,index)=>{
    if(index == 6){
      valuesResponses.push(
        ...itemResponse.getResponse()
      );
    }else{
      valuesResponses.push(
        itemResponse.getResponse()
      );
    }
      
  });
  valuesResponses.unshift(email);
  return valuesResponses;
}

const buildEmail = (responses) => {
  let dataEmail = getBodyEmail();
  let subject = dataEmail.informationEmail[0];
  let body = dataEmail.informationEmail[1];
  
  body = body.replaceAll("[No.Cert]",responses[0][0])
         .replaceAll("[Razón social del prestador de bienes y/o servicios]",responses[0][2])
         .replaceAll("[Número de identificación tributaria]",responses[0][3])
         .replaceAll("[Compañia del Grupo Bolívar]",responses[0][4])
         .replaceAll("[Descripción suministro del bien y/o servicio]",responses[0][5])
         .replaceAll("[orden es de pedido]",responses[0][6])
         .replaceAll("[urlImage]","https://drive.google.com/uc?export=view&id=1JRzAfDYeaAvpnJwkNw5qfmYA5xqChL7O");

  sendEmail(responses[0][1],subject,body);
  sendEmail("compras@segurosbolivar.com",subject,body);
}

const serial_maker = () => {

      var prefix = '';
      var seq = 0;
      return {
          set_prefix: function (p) {
              prefix = String(p);
          },
          set_seq: function (s) {
              seq = s;
          },
          gensym: function ( ) {
              var result = prefix + seq;
              seq += 1;
              return result;
          }
      };

}
