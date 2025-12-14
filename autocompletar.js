function onFormSubmit(e) {
  // Obtengo los datos del formulario
  if (!e) {
    throw new Error(`onFormSubmit | No se ingresó ningún dato del formulario`);
  }
  const formulario = e.values; // El array con los valores ingresados
  const equipo = formulario[1]; // Tipo de equipo.
  const hojaUrl = formulario[2]; // Enlace de Google Sheets
  const carpetaUrl = formulario[3]; // Enlace de la carpeta de Google Drive
  if (!equipo || !hojaUrl || !carpetaUrl) {
    throw new Error(`❌ onFormSubmit | No se ingestaron bien los datos del formulario. equipo: ${equipo}; hojaUrl:${hojaUrl}; carpetaUrl:${carpetaUrl}`);
  }
  else if (!hojaUrl.includes("spreadsheets")) {
    throw new Error("❌ onFormSubmit | El enlace ingresado NO es una hoja de Google Sheets");
  }

  // Obtengo el url del doc según el equipo seleccionado.
  const plantillas = {
    "Fotovoltaico | On grid": "https://docs.google.com/document/d/1K9Xmt6oDWxxHyMQOP55mi1DzjHkx-m3IYwQfcEVoDQ8/edit?tab=t.0",
    "Fotovoltaico | Off grid": "https://docs.google.com/document/d/1dRijnCy1bA0wUj0LpM4Qs_HMt_VUhAiFk9tqIsYbVjE/edit?tab=t.0",
    "Fotovoltaico | Híbrido": "https://docs.google.com/document/d/1R0QI6EHsg-mdmgCbhtAB4ZjspeanYgLlobISYD0fto0/edit?tab=t.0",
    "Climatización de piletas | Mantas": "https://docs.google.com/document/d/1YfKoWEQ3m3_5Zve-7ZQbMzAvpkW7I00sM16CgTbK_vk/edit?tab=t.0",
    "Climatización de piletas | Bomba de calor": "https://docs.google.com/document/d/14-X8xEtKTLa8BuB2-DgyPBarH8U2xz7IeDea1inWbUU/edit?tab=t.0",
    "Calefacción de agua | Termotanque Solar": "https://docs.google.com/document/d/1tjZGR_6VZOhCvTiy-5S3ijApSvSRlnByvI-zks76C84/edit?tab=t.0"
  };
  let plantillaUrl = plantillas[equipo];
  if (!plantillaUrl) {
    throw new Error(`❌ onFormSubmit | No existe una plantilla para el equipo "${equipo}".`);
  }

  // Genero el número de presupuesto
  const codigos = {
    "Fotovoltaico | On grid": "FV_On_Grid",
    "Fotovoltaico | Off grid": "FV_Off_grid",
    "Fotovoltaico | Híbrido": "FV_hibrido",
    "Climatización de piletas | Mantas": "Mantas",
    "Climatización de piletas | Bomba de calor": "PIL_BC",
    "Calefacción de agua | Termotanque Solar": "TTQ"
  };
  let codigoEquipo = codigos[equipo];
  if (!codigoEquipo) {
    throw new Error(`❌ onFormSubmit | No existe un código para el equipo "${equipo}".`);
  }
  const hoy = new Date();
  const anio = hoy.getFullYear();
  const mes = (hoy.getMonth() + 1).toString().padStart(2, '0'); // Mes en formato 2 dígitos
  const dia = hoy.getDate().toString().padStart(2, '0'); // Día en formato 2 dígitos
  const numRespuestas = e.range.getSheet().getLastRow();
  const numeroPresupuesto = `${anio}_${mes}${dia}_${numRespuestas}`;
  const nombreArchivo = `Presupuesto_${codigoEquipo}_${numeroPresupuesto}`;

  // LLamo a la función que genera el relleno del presupuesto
  llenarPresupuesto(plantillaUrl, hojaUrl, carpetaUrl, numeroPresupuesto, nombreArchivo);
}


function llenarPresupuesto(plantillaUrl, hojaUrl, carpetaUrl, numeroPresupuesto, nombreArchivo) {
  // Datos ingresados:
  const plantillaId = plantillaUrl.match(/[-\w]{25,}/)[0];
  const hojaId = hojaUrl.match(/[-\w]{25,}/)[0];
  const carpetaId = carpetaUrl.match(/[-\w]{25,}/)[0];
  if (!plantillaId) {
    throw new Error(`❌ llenarPresupuesto | No se encuentra el Id de la plantilla "${plantillaId}".`);
  }
  else if (!hojaId) {
    throw new Error(`❌ llenarPresupuesto | No se encuentra el Id de la hoja "${hojaId}".`);
  }
  else if (!carpetaId) {
    throw new Error(`❌ llenarPresupuesto | No se encuentra el Id de la carpeta "${carpetaId}".`);
  }

  // Abro la hoja del formulario obtengo los datos de la hoja "Autorelleno V2":


  // Abrir la hoja de cálculo y seleccionar la hoja llamada "Autorelleno"
  const libro = SpreadsheetApp.openById(hojaId);
  const hoja = libro.getSheetByName("Autorelleno");
  if (!hoja) {
    throw new Error(`❌ llenarPresupuesto | No se encontró la hoja llamada "Autorelleno". Por favor, verifica el nombre.`);
  }
  const datos = hoja.getDataRange().getValues(); // Obtener los datos de la hoja

  // Crear una copia del documento en la carpeta específica
  const carpeta = DriveApp.getFolderById(carpetaId);
  const archivoCopia = DriveApp.getFileById(plantillaId).makeCopy(nombreArchivo, carpeta);
  const copiaId = archivoCopia.getId();
  const copia = DocumentApp.openById(copiaId);
  Logger.log(`💬 llenarPresupuesto | LIBRO: "${libro.getName()}"; CARPETA: "${carpeta.getName()}"; DOCUMENTO "${DriveApp.getFileById(plantillaId).getName()}"`)
  const body = copia.getBody();
  const header = copia.getHeader();
  const footer = copia.getFooter();

  // Reemplaza específicamente 'n_presupuesto'
  Logger.log(`🔠 llenarPresupuesto | Reemplazando: "n_presupuesto"`);
  const marcador = `{{n_presupuesto}}`;
  const valor = numeroPresupuesto;
  reemplazarTextoEnDocumento(body, header, footer, marcador, valor);

  // Reemplaza marcadores de palabras en el cuerpo principal, encabezados y pies de página
  Logger.log(`🔠 llenarPresupuesto | Reemplazando: Marcadores de texto`);
  for (let i = 1; i < datos.length; i++) {
    const marcador = `{{${datos[i][0]}}}`;
    const valor = datos[i][1];
    reemplazarTextoEnDocumento(body, header, footer, marcador, valor);
  }

  // Manejo de marcadores de tabla
  Logger.log(`📆 llenarPresupuesto | Se llama a "insertarTablas"`);
  insertarTablas(libro, body, datos);

  // Manejo de marcadores de gráficos:
  Logger.log(`📊 llenarPresupuesto | Se llama a "insertarGraficosPorTitulo"`);
  insertarGraficosPorTitulo(libro, body, datos);

  // Guardar y cerrar el documento copiado
  Logger.log(`🛟 llenarPresupuesto | Guardando el documento`);
  copia.saveAndClose();

  // Exportar como PDF
  Logger.log(`📄 llenarPresupuesto | Exportando a PDF`);
  const pdfBlob = DriveApp.getFileById(copiaId).getBlob().getAs('application/pdf');
  carpeta.createFile(pdfBlob.setName(`${nombreArchivo}.pdf`));
  Logger.log(`✅ llenarPresupuesto | Documento y PDF "${nombreArchivo}"creados en la carpeta "${carpeta.getName()}"`);
  
  // TODO: Envío de reporte de errores por email
}

/**
 * Reemplaza texto en el body, header y footer.
 * Loguea SOLO UNA VEZ si no aparece en ninguna parte.
 * @param {Body} body 
 * @param {Header} header 
 * @param {Footer} footer 
 * @param {string} marcador 
 * @param {string} valor 
 */
function reemplazarTextoEnDocumento(body, header, footer, marcador, valor) {
  let encontrado = false;

  // Reemplazo en BODY
  encontrado = reemplazarEnParrafos(body.getParagraphs(), marcador, valor) || encontrado;

  // Reemplazo en HEADER
  if (header) {
    encontrado = reemplazarEnParrafos(header.getParagraphs(), marcador, valor) || encontrado;
  }

  // Reemplazo en FOOTER
  if (footer) {
    encontrado = reemplazarEnParrafos(footer.getParagraphs(), marcador, valor) || encontrado;
  }

  // Si no se encontró en ninguna parte:
  if (!encontrado) {
    Logger.log(`❌ ERROR: NO se encontró el marcador "${marcador}".`);
  }
}


/**
 * Reemplaza texto dentro de una lista de párrafos.
 * Devuelve true si lo encontró al menos una vez.
 */
function reemplazarEnParrafos(parrafos, marcador, valor) {
  let encontrado = false;

  parrafos.forEach(p => {
    if (p.getText().includes(marcador)) {
      p.replaceText(marcador, valor);
      encontrado = true;
    }
  });

  return encontrado;
}

/**
 * Función para ajustar el ancho de las columnas en una tabla.
 * Los anchos se determinan a partir de los valores de la primera fila de la tabla.
 * @param tabla - Tabla a realizar los cambios.
 */
function ajustarAnchoColumnas(tabla) {
  const numColumnas = tabla.getRow(0).getNumCells();

  // Leer los valores de la primera fila para determinar los anchos
  const anchosDeseados = [];
  for (let j = 0; j < numColumnas; j++) {
    const celda = tabla.getRow(0).getCell(j);
    const anchoTexto = parseInt(celda.getText(), 10); // Convertir el texto en número
    anchosDeseados.push(isNaN(anchoTexto) ? 50 : anchoTexto); // Si no es número, usar 50 como ancho predeterminado
    tabla.setColumnWidth(j, anchosDeseados[j] * 2.834645669291);
  }
}

function insertarTablas(libro, body, datos) {
  for (let i = 1; i < datos.length; i++) {
    try {
      const marcadorTabla = datos[i][2]; // Columna C: Marcador de tabla
      const nombreHoja = datos[i][3]; // Columna D: Nombre de la hoja
      const celdaInicial = datos[i][4]; // Columna E: Celda inicial
      const celdaFinal = datos[i][5]; // Columna F: Celda final

      if (!marcadorTabla && !nombreHoja && !celdaInicial && !celdaFinal) {
        continue;
      } else if (!marcadorTabla) {
        throw new Error(`❌ insertarTablas | En tabla ${i}, falta el dato marcadorTabla: "${marcadorTabla}"`);
      } else if (!nombreHoja) {
        throw new Error(`❌ insertarTablas | En tabla ${i}, falta el dato nombreHoja: "${nombreHoja}"`);
      } else if (!celdaInicial) {
        throw new Error(`❌ insertarTablas | En tabla ${i}, falta el dato celdaInicial: "${celdaInicial}"`);
      }else if (!celdaFinal) {
        throw new Error(`❌ insertarTablas | En tabla ${i}, falta el dato celdaFinal: "${celdaFinal}"`);
      } else {
        const hojaTabla = libro.getSheetByName(nombreHoja);
        if (!hojaTabla) {
          throw new Error(`❌ insertarTablas | No se encontró la hoja llamada "${nombreHoja}".`);
        }
        try {
          hojaTabla.getRange(celdaInicial);
          hojaTabla.getRange(celdaFinal);
        } catch (e) {
          throw new Error(`❌ insertarTablas | Rango inválido: ${celdaInicial}:${celdaFinal}`);
        }
        
        const rangoTabla = hojaTabla.getRange(`${celdaInicial}:${celdaFinal}`);
        const valores = rangoTabla.getValues();
        const imagenes = rangoTabla.getRichTextValues();
        const parrafoTabla = body.findText(`{{${marcadorTabla}}}`);
        if (!parrafoTabla) {
          throw new Error(`❌ insertarTablas | NO se encontró la tabla "${marcadorTabla}"`);
        }
        else {
          Logger.log(`⏳ insertarTablas | procesando tabla ${marcadorTabla}`)
          const elemento = parrafoTabla.getElement().getParent();
          const tablaInsertada = body.insertTable(body.getChildIndex(elemento), []);
          // Obtener valores formateados de la tabla
          const valoresFormateados = rangoTabla.getDisplayValues();
          // Llenar la tabla con datos y emular el formato
          valores.forEach((fila, i) => {
            const nuevaFila = tablaInsertada.appendTableRow();
            fila.forEach((celda, j) => {
              const nuevaCelda = nuevaFila.appendTableCell();

              // Verificar si es una imagen
              if (imagenes[i][j].getRuns().some(run => run.getLinkUrl() && run.getLinkUrl().includes("http"))) {
                const url = imagenes[i][j].getLinkUrl();
                const blob = UrlFetchApp.fetch(url).getBlob();

                // Crear un párrafo en la celda si no existe
                let parrafo;
                if (nuevaCelda.getNumChildren() === 0) {
                  parrafo = nuevaCelda.appendParagraph('');
                } else {
                  parrafo = nuevaCelda.getChild(0).asParagraph();
                }
                const imagen = parrafo.appendInlineImage(blob); // Insertar la imagen
                // Cambiar tamaño de la imagen
                const anchoDeseado = 80;
                const anchoActual = imagen.getWidth();
                const altoActual = imagen.getHeight();
                const nuevoAlto = (altoActual / anchoActual) * anchoDeseado;
                imagen.setWidth(anchoDeseado);
                imagen.setHeight(nuevoAlto);
              } else {
                // Insertar texto
                const textoFormateado = valoresFormateados[i][j].trim(); // Elimina saltos al inicio y final
                if (textoFormateado !== "") { // Solo insertar si el texto no está vacío
                  const parrafoTexto = nuevaCelda.getChild(0).asParagraph();
                  parrafoTexto.setText(textoFormateado); // Establece el texto directamente
                  // Aplicar formato básico al texto:
                  const esNegrita = rangoTabla.getCell(i + 1, j + 1).getFontWeight() === "bold"; // Negrita
                  parrafoTexto.editAsText().setBold(esNegrita);
                  const colorTexto = rangoTabla.getCell(i + 1, j + 1).getFontColor(); // Color del texto
                  parrafoTexto.editAsText().setForegroundColor(colorTexto);
                  parrafoTexto.editAsText().setBackgroundColor(null); // Eliminar resaltado del texto
                  // Ajustar alineación
                  const alineacion = rangoTabla.getCell(i + 1, j + 1).getHorizontalAlignment();
                  if (alineacion === "center") {
                    parrafoTexto.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
                  } else if (alineacion === "right") {
                    parrafoTexto.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
                  } else {
                    parrafoTexto.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
                  }
                }
              }
              // Aplicar el color de fondo de la celda después
              const colorFondo = rangoTabla.getCell(i + 1, j + 1).getBackground();
              if (colorFondo !== "#ffffff") {
                nuevaCelda.setBackgroundColor(colorFondo);
              }

            });
          });
          // Inserto bordes de 1 punto para la tabla.
          tablaInsertada.setBorderWidth(1);
          // Modifico el ancho de las columnas según la primer fila.
          ajustarAnchoColumnas(tablaInsertada);
          //  Elimino la primer fila que contiene el dato del ancho de las columnas.
          tablaInsertada.removeRow(0)
          // Eliminar marcador original
          elemento.removeFromParent();
          Logger.log(`✅ insertarTablas | ÉXITO → marcador: ${marcadorTabla}; nombreHoja: ${nombreHoja}; rango: ${celdaInicial},${celdaFinal}`);
        }
      }
    } catch (e) {
      Logger.log("❌ insertarTablas | ERROR FATAL ↓");
      Logger.log(e.stack || e.toString());
    }
  }
}

/**
 * Inserta gráficos desde Google Sheets en el documento basándose en el título y marcador especificado.
 * @param {SpreadsheetApp.Spreadsheet} libro - Objeto del libro de Google Sheets.
 * @param {DocumentApp.Body} body - Cuerpo del documento de Google Docs.
 * @param {Array} datos - Matriz con los datos de la hoja "Autorelleno".
 */
function insertarGraficosPorTitulo(libro, body, datos) {
  try {
    for (let i = 1; i < datos.length; i++) {
      const marcadorGrafico = datos[i][6];  // Columna G: Marcador del gráfico.
      const nombreHoja = datos[i][7];  // Columna H: Nombre de la hoja.
      const tituloGrafico = datos[i][8];  // Columna I: Título del gráfico.
      if (!marcadorGrafico || !nombreHoja || !tituloGrafico) {
        continue;
      }
      else {
        const hojaGraficos = libro.getSheetByName(nombreHoja);
        const graficos = hojaGraficos.getCharts(); // Obtener todos los gráficos de la hoja
        const grafico = graficos.find(chart => chart.getOptions().get('title') === tituloGrafico);
        if (!grafico) {
          throw new Error(`insertarGraficosPorTitulo | No se encontró el gráfico con el título ${tituloGrafico}`);
        }
        else {
          const imagenGrafico = grafico.getAs('image/png'); // Convertir gráfico a imagen
          const marcador = `{{${marcadorGrafico}}}`;
          const textoEncontrado = body.findText(marcador);
          if (!textoEncontrado) {
            throw new Error(`insertarGraficosPorTitulo | No se encontró el marcador de gráfico: ${marcadorGrafico}`);
          }
          else {
            const elemento = textoEncontrado.getElement().getParent();
            body.insertImage(body.getChildIndex(elemento), imagenGrafico); // Insertar imagen
            elemento.removeFromParent(); // Eliminar marcador original
          }
        }
        Logger.log(`✅ insertarGraficosPorTitulo | ÉXITO en gráfico ${i}. Datos → marcadorGrafico: ${marcadorGrafico}; nombreHoja: ${nombreHoja}; tituloGrafico: ${tituloGrafico}`);
      }
    }
  }
  catch (e) {
    Logger.log("❌ insertarGraficosPorTitulo | ERROR FATAL ↓");
    Logger.log(e.stack || e.toString());
  }
}