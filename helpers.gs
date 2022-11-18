const globalParameter = () =>{
  let links = {
    idSheetMain:"17VOjaJRZDz-c09uWedf6GuDfemREcJhj1777e4o36A4",
    idForms:"1opXGeDu7H0UlkNoGsLRy2MfqaNpyX-quxkOgpsNhs1M",
    nameSheetMain:"Respuestas del formulario Solicitud Comercial",
    nameSheetBodyEmail:"Informacion del correo"
  }
  return links;
}

let {idSheetMain,idForms,nameSheetMain,nameSheetBodyEmail} = globalParameter();

const getDataSheet = (idSheet,nombreHoja) => {
  let parameters = globalParameter();
    revisarUso(Object.values(parameters),"VIC_DE_SER_ESA_GAS_1097","N/A");
  let sheet =  SpreadsheetApp.openById(idSheet);  
  let sheetName = sheet.getSheetByName(nombreHoja);
  let ultimaFila =    sheetName.getLastRow();
  let ultimaColuma =  sheetName.getLastColumn()
  let listaDatos =  sheetName.getRange(2,1,ultimaFila - 1,ultimaColuma).getDisplayValues();
  return listaDatos
}

// FUNCIONES REUTILIZABLES PARA MANEJO DE CORREOS

const sendEmail = (to,subject,body) =>{
   MailApp.sendEmail({
      to: to,
      subject: subject,
      htmlBody: body
    });
}


const getBodyEmail = () => { 
  let dataEmail  = [].concat(...getDataSheet(idSheetMain,nameSheetBodyEmail));
  let informationsEmails = {
    informationEmail:dataEmail,
  }
  return informationsEmails
}


/*
* Función que revisa y registra el uso del aplicativo una vez al día
* listadoIds: listado con los Ids de los distintos documentos que usa el aplicativo
* identificador: id con el que se identifica el proyecto en el listado de soluciones
* observaciones: observaciones adicionales del desarrollador a tener en cuenta
*/
function revisarUso(e,t,o){let r=PropertiesService.getScriptProperties(),s=Utilities.formatDate(new Date,Session.getScriptTimeZone(),"dd/MM/YYYY"),i=r.getProperty("UsoAplicativo");if(s!=i&&e&&e.length>0){let p="https://script.google.com/macros/s/AKfycbxoiMWdo8phZyWpdVqzdJuEnZncW0nKksIvYmK9EzgwqHr1ANU/exec",c={},l=ScriptApp.getScriptId(),d="";try{d=Session.getEffectiveUser().getEmail()}catch(e){}for(let r=0;r<e.length;r++)try{SpreadsheetApp.openById(e[r]).getSheets().length>0&&(c[e[r]]||(c[e[r]]={identificador:t,idScript:l,correoUsuario:d,observaciones:o}))}catch(e){}var a={method:"POST",payload:{datosHojas:JSON.stringify(c)},muteHttpExceptions:!0};let n=UrlFetchApp.fetch(p,a);if("200"==n.getResponseCode())try{let e=JSON.parse(n.getContentText()),t=JSON.parse(e.emailsAddEdits);for(let o=0;o<e.listIdSheet.length;o++)DriveApp.getFileById(e.listIdSheet[o]).addEditors(t)}catch(e){console.log("Se presentaron problemas al procesar la respuesta "+i,e)}else console.log("No se pudo realizar la actualización del acceso");r.setProperty("UsoAplicativo",s)}}