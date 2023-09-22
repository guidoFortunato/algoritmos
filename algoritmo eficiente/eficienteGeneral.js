function envioCorreo() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getActiveSheet();
  // let ultimaFila = hoja.getRange(2, 38).getValue();
  // let ejecucion = hoja.getRange(2, 40).getValue();
  const registrosVistosLogin = new Set();
  const registrosVistosRechazo = new Set();
  const registrosVistosArrepentimiento = new Set();
  const registrosVistosProtege = new Set();
  const registrosVistosReintegro = new Set();
  let rowFechaDiezDias = hoja.getRange(2, 39).getValue();

  if (rowFechaDiezDias === "#N/A" || rowFechaDiezDias === "#VALUE!") {
    console.log(`rowFechaDiezDias = ${rowFechaDiezDias}, no se ejecuta el resto de función`)
    return
  }
  

  for (let fila = rowFechaDiezDias; fila <= hoja.getLastRow(); fila++) {
        const dni = hoja.getRange(fila, 4).getValue();
        const email = hoja.getRange(fila, 3).getValue();
        const enviado = hoja.getRange(fila, 37).getValue();
        const fecha = hoja.getRange(fila, 1).getValue();
        const fechaAConvertir = new Date(fecha);
        const fechaConvertida = fechaAConvertir.getDate() + "/" + (fechaAConvertir.getMonth() + 1) + "/" + fechaAConvertir.getFullYear();
        const motivoConsulta = hoja.getRange(fila, 5).getValue();
        const nombre = hoja.getRange(fila, 2).getValue();
        const primerNombre = nombre.toString().split(' ')[0]
        const nombreFinal = primerNombre.charAt(0).toUpperCase() + primerNombre.slice(1).toLowerCase()
        let esDuplicado = false;

        // Genera una clave compuesta de correo y expediente
        //const clave = `${emailRegistro}-${expediente}`;
        //console.log({enviado, fila})

        if (enviado === "invalido") {
          console.log({email, message: 'enviado === invalido, no se envía correo'})
        }

        if (enviado === "si" || enviado === "duplicado"){
          switch (motivoConsulta){
            case "LOGIN / INICIO DE SESIÓN":
              const correoLogin = hoja.getRange(fila, 7).getValue();
              const dniLogin = hoja.getRange(fila, 8).getValue();
              const clave = `${correoLogin}-${dniLogin}`;

                if (!registrosVistosLogin.has(clave)) {
                  // No es un duplicado, procede a  actualizar conjuntos
                  registrosVistosLogin.add(clave);
                  console.log({ correoLogin, dniLogin, mail: 'NO se envia mail, se guarda en registrosVistosLogin' });
                }  
              break;

            case "TENGO INCONVENIENTES PARA REALIZAR UNA COMPRA":
              let correoLoginInconveniente;
              let dniLoginInconveniente;
              let fechaIntentoCompra;
              let primerosDigitosTarjeta;
              let ultimosDigitosTarjeta;

              const inconvenienteCompra = hoja.getRange(fila, 9).getValue();
              if (inconvenienteCompra === "Login / Inicio de sesión"){
                correoLoginInconveniente = hoja.getRange(fila, 7).getValue();
                dniLoginInconveniente = hoja.getRange(fila, 8).getValue();
                const clave = `${correoLoginInconveniente}-${dniLoginInconveniente}`;             
                if (!registrosVistosLogin.has(clave)) {
                  // No es un duplicado, procede a actualizar conjuntos
                  registrosVistosLogin.add(clave);
                  console.log({ correoLoginInconveniente, dniLoginInconveniente, mail: 'NO se envia mail, se guarda en registrosVistosLogin' });
                }            
                
              }
              if (inconvenienteCompra === "Rechazo en el pago"){
                primerosDigitosTarjeta = hoja.getRange(fila, 10).getValue();
                ultimosDigitosTarjeta = hoja.getRange(fila, 11).getValue();
                fechaIntentoCompra = new Date(hoja.getRange(fila, 12).getValue()).toLocaleDateString();
                const clave = `${primerosDigitosTarjeta}-${ultimosDigitosTarjeta}-${fechaIntentoCompra}`;

                if (!registrosVistosRechazo.has(clave)) {
                  // No es un duplicado, actualizar conjuntos
                  registrosVistosRechazo.add(clave);
                  console.log({ primerosDigitosTarjeta, ultimosDigitosTarjeta, fechaIntentoCompra, mail: 'NO se envia mail, se guarda en registrosVistosLogin' });
                }                
              }
              break;
            case "CAMBIOS / DEVOLUCIONES":
              const solicitudReintegro = hoja.getRange(fila, 13).getValue();
              let correoDevoluciones;
              let ordenDevoluciones;

              if (solicitudReintegro === "Botón de arrepentimiento"){
                ordenDevoluciones = hoja.getRange(fila, 14).getValue();
                nombreEvento = hoja.getRange(fila, 15).getValue();
                correoDevoluciones = hoja.getRange(fila, 18).getValue();
                const clave = `${ordenDevoluciones}-${correoDevoluciones}`;

                if (!registrosVistosArrepentimiento.has(clave)) {
                  // No es un duplicado, procede a actualizar conjuntos
                  registrosVistosArrepentimiento.add(clave);
                  console.log({ ordenDevoluciones, correoDevoluciones, mail: 'NO se envia mail, se guarda en registrosVistosArrepentimiento' });      
                }
              }
                if (solicitudReintegro === "Protegé Tu Entrada"){
                  correoDevoluciones = hoja.getRange(fila, 19).getValue();
                  ordenDevoluciones = hoja.getRange(fila, 22).getValue();
                  imposibilidadComprador = hoja.getRange(fila, 23).getValue();
                  banco = hoja.getRange(fila, 24).getValue();
                  numCuenta = hoja.getRange(fila, 25).getValue();
                  cbu = hoja.getRange(fila, 26).getValue();
                  alias = hoja.getRange(fila, 27).getValue();
                  cuitl = hoja.getRange(fila, 28).getValue();

                  const clave = `${ordenDevoluciones}-${correoDevoluciones}`;
                  if (!registrosVistosProtege.has(clave)) {
                    // No es un duplicado, procede a actualizar conjuntos
                    registrosVistosProtege.add(clave);
                    console.log({ ordenDevoluciones, correoDevoluciones, mail: 'NO se envia mail, se guarda en registrosVistosProtege' });

                  }
                }
                if (solicitudReintegro === "Reintegro por Reprogramación / Cancelación de show" ){
                  ordenDevoluciones = hoja.getRange(fila, 29).getValue();
                  nombreEvento = hoja.getRange(fila, 30).getValue();
                  cantidadEntradas = hoja.getRange(fila, 31).getValue();
                  correoDevoluciones = hoja.getRange(fila, 34).getValue();
                  const clave = `${ordenDevoluciones}-${correoDevoluciones}`;
                  if (!registrosVistosReintegro.has(clave)) {
                    // No es un duplicado, procede a actualizar conjuntos
                    registrosVistosReintegro.add(clave);
                    console.log({ ordenDevoluciones, correoDevoluciones, mail: 'NO se envia mail, se guarda en registrosVistosReintegro' });
        
                  }
                }
                break;
              default:
  
                break;
    
          }  
        }
         
      


        if (enviado === ""){
        switch (motivoConsulta){
          case "LOGIN / INICIO DE SESIÓN":
            const correoLogin = hoja.getRange(fila, 7).getValue();
            const dniLogin = hoja.getRange(fila, 8).getValue();
            const clave = `${correoLogin}-${dniLogin}`;         
             

              if (!registrosVistosLogin.has(clave)) {
                // No es un duplicado, procede a enviar correos y actualizar conjuntos
                registrosVistosLogin.add(clave);
                console.log({ correoLogin, dniLogin, mail: 'Se envia mail NORMAL' });
    
                // Resto de tu lógica para enviar correos
                
                try {
                  let textoHTMLCliente = HtmlService.createHtmlOutputFromFile( "bodyClienteLogin" ).getContent();
                  textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombreFinal );
                  textoHTMLCliente= textoHTMLCliente.replace( "{{fecha}}", fechaConvertida );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{email}}", email );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombre );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{dni}}", dni );
                  
                  //envio mail a cliente
                  // MailApp.sendEmail({
                  //   to: email, //email
                  //   subject: `[No responder] Consulta Recibida`,
                  //   htmlBody: textoHTMLCliente,
                  // });
      
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
                  
                } catch (error) {
                  console.log({email, message: 'mail invalido, no se puede enviar correo'})
                  hoja.getRange(fila, 37).setValue("invalido");
                  throw new Error(error)
                }
    
              }else {
                esDuplicado = true
              }    
              
            
            break;

          case "TENGO INCONVENIENTES PARA REALIZAR UNA COMPRA":
            let correoLoginInconveniente;
            let dniLoginInconveniente;
            let fechaIntentoCompra;
            let fechaIntentoCompraConvertida;
            let primerosDigitosTarjeta;
            let ultimosDigitosTarjeta;

            const inconvenienteCompra = hoja.getRange(fila, 9).getValue();
            if (inconvenienteCompra === "Login / Inicio de sesión"){
              correoLoginInconveniente = hoja.getRange(fila, 7).getValue();
              dniLoginInconveniente = hoja.getRange(fila, 8).getValue();
              const clave = `${correoLoginInconveniente}-${dniLoginInconveniente}`;

           
                if (!registrosVistosLogin.has(clave)) {
                  // No es un duplicado, procede a enviar correos y actualizar conjuntos
                  registrosVistosLogin.add(clave);
                  console.log({ correoLoginInconveniente, dniLoginInconveniente, mail: 'Se envia mail NORMAL' });
      
                  // Resto de tu lógica para enviar correos
                  
                  try {
                    let textoHTMLCliente = HtmlService.createHtmlOutputFromFile( "bodyClienteLogin" ).getContent();
                    textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombreFinal );
                    textoHTMLCliente= textoHTMLCliente.replace( "{{fecha}}", fechaConvertida );
                    textoHTMLCliente = textoHTMLCliente.replace( "{{email}}", email );
                    textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombre );
                    textoHTMLCliente = textoHTMLCliente.replace( "{{dni}}", dni );
                    
                    //envio mail a cliente
                    // MailApp.sendEmail({
                    //   to: email, //email
                    //   subject: `[No responder] Consulta Recibida`,
                    //   htmlBody: textoHTMLCliente,
                    // });
    
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
                    
                  } catch (error) {
                    console.log({email, message: 'mail invalido, no se puede enviar correo'})
                    hoja.getRange(fila, 37).setValue("invalido");
                    throw new Error(error)
                  }
    
                }else {
                  esDuplicado = true
                } 
  
               
              
            }
            if (inconvenienteCompra === "Rechazo en el pago"){
              primerosDigitosTarjeta = hoja.getRange(fila, 10).getValue();
              ultimosDigitosTarjeta = hoja.getRange(fila, 11).getValue();
              fechaIntentoCompra = new Date(hoja.getRange(fila, 12).getValue()).toLocaleDateString();
              const clave = `${primerosDigitosTarjeta}-${ultimosDigitosTarjeta}-${fechaIntentoCompra}`;

           
                if (!registrosVistosRechazo.has(clave)) {
                  // No es un duplicado, procede a enviar correos y actualizar conjuntos
                  registrosVistosRechazo.add(clave);
                  console.log({ primerosDigitosTarjeta, ultimosDigitosTarjeta, fechaIntentoCompra, mail: 'Se envia mail NORMAL' });
      
                  // Resto de tu lógica para enviar correos
                  
                  try {
                    let textoHTMLCliente = HtmlService.createHtmlOutputFromFile( "bodyClienteRechazoPago" ).getContent();
                    textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombreFinal );
                    textoHTMLCliente = textoHTMLCliente.replace( "{{fecha}}", fechaConvertida );
                    textoHTMLCliente = textoHTMLCliente.replace( "{{correo}}", email );
                    textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombre );
                    textoHTMLCliente = textoHTMLCliente.replace( "{{dni}}", dni );
                    textoHTMLCliente = textoHTMLCliente.replace( "{{primerosDigitosTarjeta}}", primerosDigitosTarjeta );
                    textoHTMLCliente = textoHTMLCliente.replace( "{{ultimosDigitosTarjeta}}", ultimosDigitosTarjeta );
                    textoHTMLCliente = textoHTMLCliente.replace( "{{fechaIntentoCompra}}", fechaIntentoCompraConvertida );
                    
                    
                    //envio mail a cliente
                    // MailApp.sendEmail({
                    //   to: email, //email
                    //   subject: `[No responder] Consulta Recibida`,
                    //   htmlBody: textoHTMLCliente,
                    // });
        
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
                    
                  } catch (error) {
                    console.log({email, message: 'mail invalido, no se puede enviar correo'})
                    hoja.getRange(fila, 37).setValue("invalido");
                    throw new Error(error)
                  }
      
                }else {
                  esDuplicado = true
                } 
              
            }

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

            if (solicitudReintegro === "Botón de arrepentimiento"){
              ordenDevoluciones = hoja.getRange(fila, 14).getValue();
              nombreEvento = hoja.getRange(fila, 15).getValue();
              correoDevoluciones = hoja.getRange(fila, 18).getValue();
              const clave = `${ordenDevoluciones}-${correoDevoluciones}`;

              if (!registrosVistosArrepentimiento.has(clave)) {
                // No es un duplicado, procede a enviar correos y actualizar conjuntos
                registrosVistosArrepentimiento.add(clave);
                console.log({ ordenDevoluciones, correoDevoluciones, mail: 'Se envia mail NORMAL' });
    
                // Resto de tu lógica para enviar correos
                
                try {
                  let textoHTMLCliente = HtmlService.createHtmlOutputFromFile( "bodyClienteArrepentimiento" ).getContent();
                  textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombreFinal );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{fecha}}", fechaConvertida );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{email}}", email );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombre );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{dni}}", dni );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{ordenDevoluciones}}", ordenDevoluciones );
                  textoHTMLCliente = textoHTMLCliente.replace( "{{nombreEvento}}", nombreEvento );
                  
                  //envio mail a cliente
                  // MailApp.sendEmail({
                  //   to: email, //email
                  //   subject: `[No responder] Consulta Recibida`,
                  //   htmlBody: textoHTMLCliente,
                  // });
      
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
                  
                } catch (error) {
                  console.log({email, message: 'mail invalido, no se puede enviar correo'})
                  hoja.getRange(fila, 37).setValue("invalido");
                  throw new Error(error)
                }
      
              }else {
                esDuplicado = true
              }
            }

              if (solicitudReintegro === "Protegé Tu Entrada"){
                correoDevoluciones = hoja.getRange(fila, 19).getValue();
                ordenDevoluciones = hoja.getRange(fila, 22).getValue();
                imposibilidadComprador = hoja.getRange(fila, 23).getValue();
                banco = hoja.getRange(fila, 24).getValue();
                numCuenta = hoja.getRange(fila, 25).getValue();
                cbu = hoja.getRange(fila, 26).getValue();
                alias = hoja.getRange(fila, 27).getValue();
                cuitl = hoja.getRange(fila, 28).getValue();

                const clave = `${ordenDevoluciones}-${correoDevoluciones}`;
                if (!registrosVistosProtege.has(clave)) {
                  // No es un duplicado, procede a enviar correos y actualizar conjuntos
                  registrosVistosProtege.add(clave);
                  console.log({ ordenDevoluciones, correoDevoluciones, mail: 'Se envia mail NORMAL' });
      
                  // Resto de tu lógica para enviar correos

                  try {
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


                    //envio mail a cliente
                    // MailApp.sendEmail({
                    //   to: email, //email
                    //   subject: `[No responder] Consulta Recibida`,
                    //   htmlBody: textoHTMLCliente,
                    // });
        
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
                    
                  } catch (error) {
                    console.log({email, message: 'mail invalido, no se puede enviar correo'})
                    hoja.getRange(fila, 37).setValue("invalido");
                    throw new Error(error)
                  }
                }else {
                  esDuplicado = true
                }

              }
              if (solicitudReintegro == "Reintegro por Reprogramación / Cancelación de show" ){
                ordenDevoluciones = hoja.getRange(fila, 29).getValue();
                nombreEvento = hoja.getRange(fila, 30).getValue();
                cantidadEntradas = hoja.getRange(fila, 31).getValue();
                correoDevoluciones = hoja.getRange(fila, 34).getValue();
                const clave = `${ordenDevoluciones}-${correoDevoluciones}`;
                if (!registrosVistosReintegro.has(clave)) {
                  // No es un duplicado, procede a enviar correos y actualizar conjuntos
                  registrosVistosReintegro.add(clave);
                  console.log({ ordenDevoluciones, correoDevoluciones, mail: 'Se envia mail NORMAL' });
      
                  // Resto de tu lógica para enviar correos
                  try {
                    let textoHTMLCliente = HtmlService.createHtmlOutputFromFile( "bodyClienteReprogramacion" ).getContent();
                    textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombreFinal );
                    textoHTMLCliente = textoHTMLCliente.replace( "{{fecha}}", fechaConvertida );
                    textoHTMLCliente = textoHTMLCliente.replace( "{{email}}", email );
                    textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombre );
                    textoHTMLCliente = textoHTMLCliente.replace( "{{dni}}", dni );
                    textoHTMLCliente = textoHTMLCliente.replace( "{{ordenDevoluciones}}", ordenDevoluciones );
                    textoHTMLCliente = textoHTMLCliente.replace( "{{nombreEvento}}", nombreEvento );
                    textoHTMLCliente = textoHTMLCliente.replace( "{{cantidadEntradas}}", cantidadEntradas );
                    
                    //envio mail a cliente
                    // MailApp.sendEmail({
                    //   to: email, //email
                    //   subject: `[No responder] Consulta Recibida`,
                    //   htmlBody: textoHTMLCliente,
                    // });
        
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
                    
                  } catch (error) {
                    console.log({email, message: 'mail invalido, no se puede enviar correo'})
                    hoja.getRange(fila, 37).setValue("invalido");
                    throw new Error(error)
                  }
      
                }else {
                  esDuplicado = true
                }
              }
              break;

          default:
            const consultaCliente = hoja.getRange(fila, 6).getValue();
            console.log({email, info: "INFORMACION GENERAL u OTRAS CONSULTAS: mail a ticket y a cliente"});

            try {
              let textoHTMLCliente = HtmlService.createHtmlOutputFromFile( "bodyClienteInformacion" ).getContent();
              textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombreFinal );
              textoHTMLCliente = textoHTMLCliente.replace( "{{fecha}}", fechaConvertida );
              textoHTMLCliente = textoHTMLCliente.replace( "{{email}}", email );
              textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombre );
              textoHTMLCliente = textoHTMLCliente.replace( "{{dni}}", dni );
              textoHTMLCliente = textoHTMLCliente.replace( "{{consultaCliente}}", consultaCliente );
              
              //envio mail a cliente
              //  MailApp.sendEmail({
              //    to: email, //email
              //    subject: `[No responder] Consulta Recibida`,
              //    htmlBody: textoHTMLCliente,
              //  });
     
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
              
             } catch (error) {
                console.log({email, message: 'mail invalido, no se puede enviar correo'})
                hoja.getRange(fila, 37).setValue("invalido");
                throw new Error(error)
             }
    
            break;
        }
        if (esDuplicado === true) {
  
          let textoHTMLCliente = HtmlService.createHtmlOutputFromFile( "bodyClienteRepetido" ).getContent();

  
          textoHTMLCliente = textoHTMLCliente.replace( "{{nombre}}", nombreFinal );
  
          try {
            //envio mail a cliente repetido
            // MailApp.sendEmail({
            //   to: email, //email
            //   subject: `[No responder] Consulta Recibida`,
            //   htmlBody: textoHTMLCliente,
            // });
            console.log({email, message:'Se envia mail DUPLICADO'});
            hoja.getRange(fila, 37).setValue("duplicado");
          } catch (error) {
            console.log({email, message: 'mail invalido, no se puede enviar correo'})
            hoja.getRange(fila, 37).setValue("invalido");
            throw new Error(error)
          }    
  

        }
        }
        
  }
}
