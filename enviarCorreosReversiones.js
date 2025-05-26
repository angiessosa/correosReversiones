//BORRADOOOOOOOOOOOOOOOOOOR!!!!!!!!!!!!!!

function enviarCorreosReversiones() {
  const ui = SpreadsheetApp.getUi();
  
  //Cuadro de dialogo para ingresar datos personales que iran en la firma del correo
  const respuesta =ui.prompt(
    '✏️ Datos para la firma',
    'Ingresa tu nombre y cargo (ej: "Nombre Completo | Perfil a Cargo"):',
    ui.ButtonSet.OK_CANCEL
  );

  //al seleccionar cancelar, se detiene el código y no hara cambios
  if (respuesta.getSelectedButton() === ui.Button.CANCEL) return;
  const datosUsuario = respuesta.getResponseText();

  //URL pública del logo de bolivar
  const urlLogo = "https://cdn.brandfetch.io/idNYmlVYed/w/400/h/400/theme/dark/icon.jpeg?c=1dxbfHSJFAPEGdCLU4o5B";

  // Plantilla de firma con imagen a la izquierda y texto a la derecha
  const firmaHTML = `
    <div style="font-family: Arial, sans-serif; color: #333; margin-top: 20px; border-top: 1px solid #eee; padding-top: 15px;">
      <table style="width: 100%; border-collapse: collapse;">
        <tr>
          <td style="width: 80px; vertical-align: top; padding-right: 15px;">
            <img src="${urlLogo}" 
                 alt="Logo Seguros Bolívar" 
                 style="width: 150px; height: auto; display: block;">
          </td>
          <td style="vertical-align: top; padding-left: 15px; border-left: 1px solid #eee;">
            <p style="margin: 10px 0 5px 0; font-weight: bold;">${datosUsuario}</p>
            <p style="margin: 0 0 3px 0; font-size: 11px; line-height: 1.4;">
              <stronger>Dirección Nacional Administrativa ARL</stronger><br>
              Av. El Dorado #68 B-65 Piso 7<br>
              Bogotá, Colombia
            </p>
            <p style="margin: 5px 0 0 0; font-size: 11px;">
              <a href="https://www.segurosbolivar.com" 
                 style="color: #0066cc; text-decoration: none;">www.segurosbolivar.com</a>
            </p>
          </td>
        </tr>
      </table>
    </div>
  `;

  // Busca la hoja llamada Correo Reversiones
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Correo Reversiones");
  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0];

   // Índices de columnas
  const iEmpresa = encabezados.indexOf("EMPRESA");
  const iNit = encabezados.indexOf("NIT");
  const iLocalidad = encabezados.indexOf("LOCALIDAD");
  const iCanal = encabezados.indexOf("CANAL");
  const iValorReversion = encabezados.indexOf("VALOR REVERSION");
  const iPeriodoIn = encabezados.indexOf("PERIODO INICIO");
  const iPeriodoFn = encabezados.indexOf("PERIODO FIN");
  const iMotivoAjuste = encabezados.indexOf("MOTIVO AJUSTE");
  const iCorreoEnv = encabezados.indexOf("CORREO ENVIADO");

  //guarda en una lista los correos de destinatarios
  const destinatarios = ["angie.sosa@segurosbolivar.com"];

    //Verifica el envio de correo de una empresa con la condicion "SI"
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const correoYaEnviado = fila[iCorreoEnv].toString().toUpperCase() === "SI";

    //Si es Diferente a "SI", procede con el envio de correo obteniendo los datos que se encuentran en la hoja de excel
    if (!correoYaEnviado) {
      const empresa = fila[iEmpresa];
      const nit = fila[iNit];
      const localidad = fila[iLocalidad];
      const canal = fila[iCanal];
      const valorReversion = formatearMoneda(fila[iValorReversion]);
      const periodoInicio = fila[iPeriodoIn];
      const periodoFin = fila[iPeriodoFn];
      const motivoAjuste = fila[iMotivoAjuste];
      

      //Se define la estructura del correo
      const asunto = `ASUNTO SIN CONFIRMAR!!!!!!!!!!!!!!!!`;
      const mensajeHTML = `Buenos días, <br><br>
De manera atenta nos permitimos informar que el aportante <b>${empresa}</b> con NIT <b>${nit}</b> de la localidad <b>${localidad}</b> y canal <b>${canal}</b> presentará una reversión de <b>${valorReversion}</b> este mes, correspondiente a los periodos que se encuentran en mora desde el periodo <b>${periodoInicio}</b> hasta el periodo <b>${periodoFin}</b>, dado el siguiente concepto. <br><br>
<!-- Tabla motivo del ajuste -->
<table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 14px;">
  <tr>
    <td style="border: 1px solid #000; padding: 8px; font-weight: bold; background-color: #f2f2f2;">Motivo del Ajuste</td>
    <td style="border: 1px solid #000; padding: 8px;">➢ ${motivoAjuste}</td>
  </tr>
</table>
<br><br><br>
Coordialmente,<br><br>
${firmaHTML}`;

      // ENVIAR CORREO
      MailApp.sendEmail({
        to: destinatarios.join(","),
        subject: asunto,
        htmlBody: mensajeHTML
      });

      // MARCA COMO ENVIADO
      hoja.getRange(i + 1, iCorreoEnv + 1).setValue("SI");
    }
  }
  
  ui.alert("✅ Correos enviados", "Los correos se han enviado correctamente con tu firma profesional.", ui.ButtonSet.OK);
}


//Funcion para poner valor en Pesos
function formatearMoneda(valor) {
  const numero = parseFloat(valor);
  if (!isNaN(numero)) {
    return '$' + numero.toLocaleString('es-CO');
  }
  return valor;
}
