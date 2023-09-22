function envioCorreo() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getActiveSheet();
  //let ultimaFila = hoja.getRange(2, 9).getValue();
  let ejecucion = hoja.getRange(2, 10).getValue();

  try {
    if (ejecucion === "no") {
      hoja.getRange(2, 10).setValue("si");
      const registrosVistos = new Set();

      for (let fila = 2; fila <= hoja.getLastRow(); fila++) {
        let enviado = hoja.getRange(fila, 8).getValue();
        const emailRegistro = hoja.getRange(fila, 5).getValue();
        const expediente = hoja.getRange(fila, 6).getValue();
        // Genera una clave compuesta de correo y expediente
        const clave = `${emailRegistro}-${expediente}`;
        //console.log({enviado, fila})

        if (enviado === "si") {
          if (!registrosVistos.has(clave)) {
            // agrego a registros vistos los no duplicados y los que si son enviados
            registrosVistos.add(clave);
            //console.log({ registrosVistosSi: registrosVistos });
            //console.log('cliente no duplicado')
            // Resto de tu lógica para enviar correos
          }
        }
        if (enviado !== "si") {
          if (!registrosVistos.has(clave)) {
            // No es un duplicado, procede a enviar correos y actualizar conjuntos
            registrosVistos.add(clave);
            //console.log({ registrosVistosNo: registrosVistos });

            // Resto de tu lógica para enviar correos
            hoja.getRange(fila, 8).setValue("si");
          } else {
            // Es un duplicado, puedes manejarlo aquí si es necesario
            //console.log('cliente duplicado')
            hoja.getRange(fila, 8).setValue("duplicado");
          }
          hoja.getRange(2, 9).setValue(fila);
        }
      }
      hoja.getRange(2, 10).setValue("no");
    }
  } catch (error) {
    hoja.getRange(2, 10).setValue("no");
    throw new Error(error);
  }
}
