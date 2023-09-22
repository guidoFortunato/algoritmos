function envioCorreo() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getActiveSheet();
  let ultimaFila = hoja.getRange(2, 38).getValue();
  let ejecucion = hoja.getRange(2, 40).getValue();

  try {
    if (ejecucion === "no") {
      hoja.getRange(2, 40).setValue("si");
      for (let fila = ultimaFila + 1; fila <= hoja.getLastRow(); fila++) {
        const fecha = hoja.getRange(fila, 1).getValue();
        const fechaAConvertir = new Date(fecha);
        const fechaConvertida = fechaAConvertir.getDate() + "/" + (fechaAConvertir.getMonth() + 1) + "/" + fechaAConvertir.getFullYear();
        const nombre = hoja.getRange(fila, 2).getValue();
        const primerNombre = nombre.toString().split(' ')[0]
        const nombreFinal = primerNombre.charAt(0).toUpperCase() + primerNombre.slice(1).toLowerCase()
        const email = hoja.getRange(fila, 3).getValue();
        const dni = hoja.getRange(fila, 4).getValue();
        const motivoConsulta = hoja.getRange(fila, 5).getValue();
        const enviado = hoja.getRange(fila, 37).getValue();
        let chequear_fecha;
        let esDuplicado = false;
        let rowFechaDiezDias;
    
        if (enviado === "") {
          rowFechaDiezDias = hoja.getRange(2, 39).getValue();
          switch (motivoConsulta) {
            case "LOGIN / INICIO DE SESIÓN":
              const correoLogin = hoja.getRange(fila, 7).getValue();
              const dniLogin = hoja.getRange(fila, 8).getValue();
              let chequear_correo;
              for (let row = rowFechaDiezDias; row <= ultimaFila; row++) {
                chequear_fecha = hoja.getRange(row, 1).getValue();
                chequear_correo = hoja.getRange(row, 7).getValue();
                const chequear_dni = hoja.getRange(row, 8).getValue();
    
                if (correoLogin === chequear_correo && dniLogin === chequear_dni) {
                  esDuplicado = true;
                  break;
                }
              }
    
              if (!esDuplicado) {
                
                //envio mail a cliente
    
                let textoHTMLCliente = HtmlService.createHtmlOutputFromFile( "bodyClienteLogin" ).getContent();
                textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombreFinal );
                textoHTMLCliente= textoHTMLCliente.replace( "{{fecha}}", fechaConvertida );
                textoHTMLCliente = textoHTMLCliente.replace( "{{email}}", email );
                textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombre );
                textoHTMLCliente = textoHTMLCliente.replace( "{{dni}}", dni );
            
                MailApp.sendEmail({
                  to: email, //email
                  subject: `[No responder] Consulta Recibida`,
                  htmlBody: textoHTMLCliente,
                });
    
                //envio mail a ticket
                let textoHTMLTicket = HtmlService.createHtmlOutputFromFile( "bodyTicketLogin" ).getContent();
                textoHTMLTicket = textoHTMLTicket.replace( "{{fecha}}", fechaConvertida );
                textoHTMLTicket = textoHTMLTicket.replace( "{{email}}", email );
                textoHTMLTicket = textoHTMLTicket.replace( "{{nombre}}", nombre );
                textoHTMLTicket = textoHTMLTicket.replace( "{{dni}}", dni );
            
                MailApp.sendEmail({
                  to: 'tickets@tuentrada.com', //tickets@tuentrada.com
                  subject: `Login - Inicio Sesion - ${email}`,
                  htmlBody: textoHTMLTicket,
                });
    
                hoja.getRange(fila, 37).setValue("si");
              } 
    
    
              // console.log("LOGIN / INICIO DE SESIÓN");
              // console.log("duplicado: " + esDuplicado);
              // console.log(new Date(chequear_fecha).toLocaleString());
              // console.log(correoLogin);
              // console.log(chequear_correo);
              // console.log("********************************************");
    
              break;
    
            case "TENGO INCONVENIENTES PARA REALIZAR UNA COMPRA":
              let chequear_fechaIntentoCompra;
              let chequear_primerosDigitosTarjeta;
              let chequear_ultimosDigitosTarjeta;
              let correoLoginInconveniente;
              let dniLoginInconveniente;
              let fechaIntentoCompra;
              let fechaIntentoCompraConvertir;
              let fechaIntentoCompraConvertida;
              let primerosDigitosTarjeta;
              let ultimosDigitosTarjeta;
    
              const inconvenienteCompra = hoja.getRange(fila, 9).getValue();
              if (inconvenienteCompra === "Login / Inicio de sesión") {
                correoLoginInconveniente = hoja.getRange(fila, 7).getValue();
                dniLoginInconveniente = hoja.getRange(fila, 8).getValue();
    
                for (let row = rowFechaDiezDias; row <= ultimaFila; row++) {
                  chequear_fecha = hoja.getRange(row, 1).getValue();
                  const chequear_correo = hoja.getRange(row, 7).getValue();
                  const chequear_dni = hoja.getRange(row, 8).getValue();
    
                  if ( correoLoginInconveniente === chequear_correo && dniLoginInconveniente === chequear_dni ) {
                    esDuplicado = true;
                    break;
                  }
                }
                if (!esDuplicado) {
                
                  //envio mail a cliente
      
                  let textoHTMLCliente = HtmlService.createHtmlOutputFromFile( "bodyClienteLogin" ).getContent();
                  textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombreFinal );
                  textoHTMLCliente= textoHTMLCliente.replace( "{{fecha}}", fechaConvertida );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{email}}", email );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombre );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{dni}}", dni );
              
                  MailApp.sendEmail({
                    to: email, //email
                    subject: `[No responder] Consulta Recibida`,
                    htmlBody: textoHTMLCliente,
                  });
      
                  //envio mail a ticket
                  let textoHTMLTicket = HtmlService.createHtmlOutputFromFile( "bodyTicketLogin" ).getContent();
                  textoHTMLTicket = textoHTMLTicket.replace( "{{fecha}}", fechaConvertida );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{email}}", email );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{nombre}}", nombre );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{dni}}", dni );
              
                  MailApp.sendEmail({
                    to: 'tickets@tuentrada.com', //tickets@tuentrada.com
                    subject: `Login - Inicio Sesion - ${email}`,
                    htmlBody: textoHTMLTicket,
                  });
      
                  hoja.getRange(fila, 37).setValue("si");
                } 
              } else if (inconvenienteCompra === "Rechazo en el pago") {
                primerosDigitosTarjeta = hoja.getRange(fila, 10).getValue();
                ultimosDigitosTarjeta = hoja.getRange(fila, 11).getValue();
                fechaIntentoCompra = new Date(hoja.getRange(fila, 12).getValue()).toLocaleDateString();
    
                for (let row = rowFechaDiezDias; row <= ultimaFila; row++) {
                  chequear_fecha = hoja.getRange(row, 1).getValue();
                  chequear_primerosDigitosTarjeta = hoja
                    .getRange(row, 10)
                    .getValue();
                  chequear_ultimosDigitosTarjeta = hoja
                    .getRange(row, 11)
                    .getValue();
                  chequear_fechaIntentoCompra = new Date(
                    hoja.getRange(row, 12).getValue()
                  ).toLocaleDateString();
    
                  if (
                    primerosDigitosTarjeta === chequear_primerosDigitosTarjeta &&
                    ultimosDigitosTarjeta === chequear_ultimosDigitosTarjeta &&
                    fechaIntentoCompra === chequear_fechaIntentoCompra
                  ) {
                    esDuplicado = true;
                    break;
                  }
                }
    
                if (!esDuplicado) {
                  fechaIntentoCompraConvertir = new Date(hoja.getRange(fila, 12).getValue());
                  fechaIntentoCompraConvertida = fechaIntentoCompraConvertir.getDate() + "/" + (fechaIntentoCompraConvertir.getMonth() + 1) + "/" + fechaIntentoCompraConvertir.getFullYear();
                
                  //envio mail a cliente
      
                  let textoHTMLCliente = HtmlService.createHtmlOutputFromFile( "bodyClienteRechazoPago" ).getContent();
                  textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombreFinal );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{fecha}}", fechaConvertida );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{correo}}", email );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombre );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{dni}}", dni );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{primerosDigitosTarjeta}}", primerosDigitosTarjeta );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{ultimosDigitosTarjeta}}", ultimosDigitosTarjeta );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{fechaIntentoCompra}}", fechaIntentoCompraConvertida );
              
              
                  MailApp.sendEmail({
                    to: email, //email
                    subject: `[No responder] Consulta Recibida`,
                    htmlBody: textoHTMLCliente,
                  });
      
                  //envio mail a ticket
                  let textoHTMLTicket = HtmlService.createHtmlOutputFromFile( "bodyTicketRechazoPago" ).getContent();
                  textoHTMLTicket = textoHTMLTicket.replace( "{{fecha}}", fechaConvertida );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{correo}}", email );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{nombre}}", nombre );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{dni}}", dni );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{primerosDigitosTarjeta}}", primerosDigitosTarjeta );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{ultimosDigitosTarjeta}}", ultimosDigitosTarjeta );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{fechaIntentoCompra}}", fechaIntentoCompraConvertida );
              
                  MailApp.sendEmail({
                    to: 'tickets@tuentrada.com', //tickets@tuentrada.com
                    subject: `Rechazo en el pago - ${email}`,
                    htmlBody: textoHTMLTicket,
                  });
      
                  hoja.getRange(fila, 37).setValue("si");
                } 
              }
    
              // console.log("TENGO INCONVENIENTES PARA REALIZAR UNA COMPRA");
              // console.log("duplicado: " + esDuplicado);
              // console.log(new Date(chequear_fecha).toLocaleString());
              // console.log("********************************************");
              break;
    
            case "CAMBIOS / DEVOLUCIONES":
              const solicitudReintegro = hoja.getRange(fila, 13).getValue();
              let alias;
              let banco;
              let cantidadEntradas;
              let cbu;
              let correoDevoluciones;
              let cuitl;
              let imposibilidadComprador;
              let nombreEvento;
              let numCuenta;
              let ordenDevoluciones;
              if (solicitudReintegro === "Botón de arrepentimiento") {
                ordenDevoluciones = hoja.getRange(fila, 14).getValue();
                nombreEvento = hoja.getRange(fila, 15).getValue();
                correoDevoluciones = hoja.getRange(fila, 18).getValue();
                for (let row = rowFechaDiezDias; row <= ultimaFila; row++) {
                  chequear_fecha = hoja.getRange(row, 1).getValue();
                  const chequear_ordenDevoluciones = hoja
                    .getRange(row, 14)
                    .getValue();
                  const chequear_correoDevoluciones = hoja
                    .getRange(row, 18)
                    .getValue();
    
                  if (
                    ordenDevoluciones === chequear_ordenDevoluciones &&
                    correoDevoluciones === chequear_correoDevoluciones
                  ) {
                    esDuplicado = true;
                    break;
                  }
                }
    
                if (!esDuplicado) {
                
                  //envio mail a cliente
      
                  let textoHTMLCliente = HtmlService.createHtmlOutputFromFile( "bodyClienteArrepentimiento" ).getContent();
                  textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombreFinal );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{fecha}}", fechaConvertida );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{email}}", email );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombre );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{dni}}", dni );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{ordenDevoluciones}}", ordenDevoluciones );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{nombreEvento}}", nombreEvento );
              
                  MailApp.sendEmail({
                    to: email, //email
                    subject: `[No responder] Consulta Recibida`,
                    htmlBody: textoHTMLCliente,
                  });
      
                  //envio mail a ticket
                  let textoHTMLTicket = HtmlService.createHtmlOutputFromFile( "bodyTicketArrepentimiento" ).getContent();
                  textoHTMLTicket = textoHTMLTicket.replace( "{{fecha}}", fechaConvertida );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{email}}", email );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{nombre}}", nombre );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{dni}}", dni );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{ordenDevoluciones}}", ordenDevoluciones );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{nombreEvento}}", nombreEvento );
              
                  MailApp.sendEmail({
                    to: 'tickets@tuentrada.com', //tickets@tuentrada.com
                    subject: `Botón de arrepentimiento - ${ordenDevoluciones}`,
                    htmlBody: textoHTMLTicket,
                  });
      
                  hoja.getRange(fila, 37).setValue("si");
                }
    
    
    
    
    
              } else if (solicitudReintegro === "Protegé Tu Entrada") {
                correoDevoluciones = hoja.getRange(fila, 19).getValue();
                ordenDevoluciones = hoja.getRange(fila, 22).getValue();
                imposibilidadComprador = hoja.getRange(fila, 23).getValue();
                banco = hoja.getRange(fila, 24).getValue();
                numCuenta = hoja.getRange(fila, 25).getValue();
                cbu = hoja.getRange(fila, 26).getValue();
                alias = hoja.getRange(fila, 27).getValue();
                cuitl = hoja.getRange(fila, 28).getValue();
                for (let row = rowFechaDiezDias; row <= ultimaFila; row++) {
                  chequear_fecha = hoja.getRange(row, 1).getValue();
                  const chequear_correoDevoluciones = hoja
                    .getRange(row, 19)
                    .getValue();
                  const chequear_ordenDevoluciones = hoja
                    .getRange(row, 22)
                    .getValue();
    
                  if (
                    ordenDevoluciones === chequear_ordenDevoluciones &&
                    correoDevoluciones === chequear_correoDevoluciones
                  ) {
                    esDuplicado = true;
                    break;
                  }
                }
    
                if (!esDuplicado) {
                
                  //envio mail a cliente
      
                  let textoHTMLCliente = HtmlService.createHtmlOutputFromFile( "bodyClienteProtege" ).getContent();
                  textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombreFinal );
              textoHTMLCliente = textoHTMLCliente.replace( "{{fecha}}", fechaConvertida );
              textoHTMLCliente = textoHTMLCliente.replace( "{{email}}", email );
              textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombre );
              textoHTMLCliente = textoHTMLCliente.replace( "{{dni}}", dni );
              textoHTMLCliente = textoHTMLCliente.replace( "{{ordenDevoluciones}}", ordenDevoluciones );
              textoHTMLCliente = textoHTMLCliente.replace( "{{imposibilidadComprador}}", imposibilidadComprador );
              textoHTMLCliente = textoHTMLCliente.replace( "{{banco}}", banco );
              textoHTMLCliente = textoHTMLCliente.replace( "{{numCuenta}}", numCuenta );
              textoHTMLCliente = textoHTMLCliente.replace( "{{cbu}}", cbu );
              textoHTMLCliente = textoHTMLCliente.replace( "{{alias}}", alias );
              textoHTMLCliente = textoHTMLCliente.replace( "{{cuitl}}", cuitl );
              
                  MailApp.sendEmail({
                    to: email, //email
                    subject: `[No responder] Consulta Recibida`,
                    htmlBody: textoHTMLCliente,
                  });
      
                  //envio mail a ticket
                  let textoHTMLTicket = HtmlService.createHtmlOutputFromFile( "bodyTicketProtege" ).getContent();
                  textoHTMLTicket = textoHTMLTicket.replace( "{{fecha}}", fechaConvertida );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{email}}", email );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{nombre}}", nombre );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{dni}}", dni );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{ordenDevoluciones}}", ordenDevoluciones );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{imposibilidadComprador}}", imposibilidadComprador );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{banco}}", banco );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{numCuenta}}", numCuenta );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{cbu}}", cbu );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{alias}}", alias );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{cuitl}}", cuitl );
              
                  MailApp.sendEmail({
                    to: 'tickets@tuentrada.com', //tickets@tuentrada.com
                    subject: `Protege tu entrada - ${ordenDevoluciones}`,
                    htmlBody: textoHTMLTicket,
                  });
      
                  hoja.getRange(fila, 37).setValue("si");
                }
    
    
    
              } else if (
                solicitudReintegro ===
                "Reintegro por Reprogramación / Cancelación de show"
              ) {
                ordenDevoluciones = hoja.getRange(fila, 29).getValue();
                nombreEvento = hoja.getRange(fila, 30).getValue();
                cantidadEntradas = hoja.getRange(fila, 31).getValue();
                correoDevoluciones = hoja.getRange(fila, 34).getValue();
                for (let row = rowFechaDiezDias; row <= ultimaFila; row++) {
                  chequear_fecha = hoja.getRange(row, 1).getValue();
                  const chequear_ordenDevoluciones = hoja
                    .getRange(row, 29)
                    .getValue();
                  const chequear_correoDevoluciones = hoja
                    .getRange(row, 34)
                    .getValue();
    
                  if (
                    ordenDevoluciones === chequear_ordenDevoluciones &&
                    correoDevoluciones === chequear_correoDevoluciones
                  ) {
                    esDuplicado = true;
                    break;
                  }
                }
    
                if (!esDuplicado) {
                
                  //envio mail a cliente
      
                  let textoHTMLCliente = HtmlService.createHtmlOutputFromFile( "bodyClienteReprogramacion" ).getContent();
                  textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombreFinal );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{fecha}}", fechaConvertida );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{email}}", email );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombre );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{dni}}", dni );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{ordenDevoluciones}}", ordenDevoluciones );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{nombreEvento}}", nombreEvento );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{cantidadEntradas}}", cantidadEntradas );
              
                  MailApp.sendEmail({
                    to: email, //email
                    subject: `[No responder] Consulta Recibida`,
                    htmlBody: textoHTMLCliente,
                  });
      
                  //envio mail a ticket
                  let textoHTMLTicket = HtmlService.createHtmlOutputFromFile( "bodyTicketReprogramacion" ).getContent();
                  textoHTMLTicket = textoHTMLTicket.replace( "{{fecha}}", fechaConvertida );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{email}}", email );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{nombre}}", nombre );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{dni}}", dni );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{ordenDevoluciones}}", ordenDevoluciones );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{nombreEvento}}", nombreEvento );
                  textoHTMLTicket = textoHTMLTicket.replace( "{{cantidadEntradas}}", cantidadEntradas );
    
                  MailApp.sendEmail({
                    to: 'tickets@tuentrada.com', //tickets@tuentrada.com
                    subject: `Reintegro por Reprogramación o Cancelación del show - ${ordenDevoluciones}`,
                    htmlBody: textoHTMLTicket,
                  });
      
                  hoja.getRange(fila, 37).setValue("si");
                }
              }
              // console.log("CAMBIOS / DEVOLUCIONES");
              // console.log("duplicado: " + esDuplicado);
              // console.log(chequear_fecha);
              // console.log("********************************************");
              break;
    
            default:
              const consultaCliente = hoja.getRange(fila, 6).getValue();
                
              //envio mail a cliente
    
              let textoHTMLCliente = HtmlService.createHtmlOutputFromFile( "bodyClienteInformacion" ).getContent();
              textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombreFinal );
              textoHTMLCliente = textoHTMLCliente.replace( "{{fecha}}", fechaConvertida );
              textoHTMLCliente = textoHTMLCliente.replace( "{{email}}", email );
              textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombre );
              textoHTMLCliente = textoHTMLCliente.replace( "{{dni}}", dni );
              textoHTMLCliente = textoHTMLCliente.replace( "{{consultaCliente}}", consultaCliente );
            
              MailApp.sendEmail({
                to: email, //email
                subject: `[No responder] Consulta Recibida`,
                htmlBody: textoHTMLCliente,
              });
    
              //envio mail a ticket
              let textoHTMLTicket = HtmlService.createHtmlOutputFromFile( "bodyTicketInformacion" ).getContent();
              textoHTMLTicket = textoHTMLTicket.replace( "{{fecha}}", fechaConvertida );
              textoHTMLTicket = textoHTMLTicket.replace( "{{email}}", email );
              textoHTMLTicket = textoHTMLTicket.replace( "{{nombre}}", nombre );
              textoHTMLTicket = textoHTMLTicket.replace( "{{dni}}", dni );
              textoHTMLTicket = textoHTMLTicket.replace( "{{consultaCliente}}", consultaCliente );
            
              MailApp.sendEmail({
                to: 'tickets@tuentrada.com', //tickets@tuentrada.com
                subject: `Informacion General u otras consultas - ${email}`,
                htmlBody: textoHTMLTicket,
              });
    
              hoja.getRange(fila, 37).setValue("si");
              
              // console.log("INFORMACION GENERAL u OTRAS CONSULTAS: mail a ticket y a cliente");
              break;
          }
    
          if (esDuplicado) {
                
            const chequearFechaAConvertir = new Date(chequear_fecha);
            const chequearFechaConvertida = chequearFechaAConvertir.getDate() + "/" + (chequearFechaAConvertir.getMonth() + 1) + "/" + chequearFechaAConvertir.getFullYear();
    
            let textoHTMLCliente = HtmlService.createHtmlOutputFromFile( "bodyClienteRepetido" ).getContent();
            textoHTMLCliente = textoHTMLCliente.replace( "{{chequear_fecha}}", chequearFechaConvertida );
    
            textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombreFinal );
    
            //envio mail a cliente repetido
        
            MailApp.sendEmail({
              to: email, //email
              subject: `[No responder] Consulta Recibida`,
              htmlBody: textoHTMLCliente,
            });
    
            hoja.getRange(fila, 37).setValue("duplicado");
          }
        console.log({fila})
          hoja.getRange(2, 38).setValue(fila);
          ultimaFila = hoja.getRange(2, 38).getValue();
        }
      }
      hoja.getRange(2, 40).setValue("no");
    }

  }catch (error){
    hoja.getRange(2, 40).setValue("no");
    throw new Error(error);
  }


}
