/**
 * @OnlyCurrentDoc
 * Este script procesa la hoja 'Principal', crea nuevas hojas basadas en categorías
 * y genera un documento de Google Docs para cada nueva hoja creada.
 */

// --- CONFIGURACIÓN ---
const PRINCIPAL_SHEET_NAME = 'Principal'; // Nombre de la hoja de origen
const PUBLICABLE_COL_NAME = 'Publicable'; // Nombre columna para filtrar publicables
const RESPUESTA_COL_NAME = 'Respuesta'; // Nombre columna Respuesta original
const RESPUESTA_ACT_COL_NAME = 'Respuesta Actualizada'; // Nombre columna Respuesta actualizada
const ETIQUETA_ORIG_COL_NAME = 'Etiqueta Original'; // Nombre columna Etiqueta original
const ETIQUETA_PROP_COL_NAME = 'Etiqueta Propuesta'; // Nombre columna Etiqueta propuesta

const TAG_SEPARATOR = ';'; // Separador de etiquetas
const EXCLUDED_TAG = 'Sedes'; // Etiqueta específica a excluir COMPLETAMENTE

// Define las reglas de agrupación y los nombres de las hojas de destino
const GROUPING_RULES = {
  'FAQ_Ingresantes': ["Ingreso", "Ingresantes", "Art. 7", "Inscripción", "Beca", "CIVU"],
  'FAQ_Alumnos_Examen': ["Alumno", "Examen"],
  'FAQ_Tramites_Equivalencias_SAG': ["Tramites", "Aranceles", "equivalencias", "SAG"],
  'FAQ_Preguntas_frecuentes_generales': ["Oferta Académica", "Postgrados", "Cursos", "FAQ", "Varios"]
};
// --- FIN CONFIGURACIÓN ---

/**
 * Función principal que orquesta el proceso.
 */
function processAndCreateSheetsAndDocs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const principalSheet = ss.getSheetByName(PRINCIPAL_SHEET_NAME);

  if (!principalSheet) {
    SpreadsheetApp.getUi().alert(`Error: No se encontró la hoja llamada "${PRINCIPAL_SHEET_NAME}".`);
    return;
  }

  const principalData = principalSheet.getDataRange().getValues();
  if (principalData.length < 2) {
    SpreadsheetApp.getUi().alert(`La hoja "${PRINCIPAL_SHEET_NAME}" está vacía o solo tiene encabezados.`);
    return;
  }

  const headers = principalData[0];
  const colIndices = findColumnHeaders(headers);

  // Validar que las columnas necesarias existen
  const requiredCols = [
    PUBLICABLE_COL_NAME, RESPUESTA_COL_NAME, RESPUESTA_ACT_COL_NAME,
    ETIQUETA_ORIG_COL_NAME, ETIQUETA_PROP_COL_NAME
  ];
  for (const colName of requiredCols) {
    if (colIndices[colName] === undefined) {
      SpreadsheetApp.getUi().alert(`Error: No se encontró la columna "${colName}" en la hoja "${PRINCIPAL_SHEET_NAME}".`);
      return;
    }
  }

  // 1. Procesar datos (Filtrar, Transformar, Categorizar)
  const categorizedData = processData(principalData, colIndices);

  // 2. Crear/Limpiar Hojas y Escribir Datos
  const createdSheetNames = createOrClearSheets(ss, categorizedData, headers);

  // 3. Crear Documentos de Google
  if (createdSheetNames.length > 0) {
     createDocsFromSheets(ss, createdSheetNames);
     SpreadsheetApp.getUi().alert(`Proceso completado. Se crearon/actualizaron ${createdSheetNames.length} hojas y sus documentos correspondientes.`);
  } else {
     SpreadsheetApp.getUi().alert('Proceso completado. No se generaron hojas nuevas (quizás no había datos que cumplieran los criterios).');
  }
}

/**
 * Encuentra los índices de las columnas basado en los encabezados.
 * @param {string[]} headers Array de encabezados.
 * @return {object} Un objeto mapeando nombre de columna a índice (basado en 0).
 */
function findColumnHeaders(headers) {
  const indices = {};
  headers.forEach((header, index) => {
    indices[header.trim()] = index;
  });
  return indices;
}

/**
 * Procesa los datos de la hoja principal: filtra, transforma y categoriza.
 * @param {Array<Array<string>>} data Datos completos de la hoja principal (incluye encabezados).
 * @param {object} colIndices Objeto con los índices de las columnas.
 * @return {object} Objeto donde las claves son nombres de hojas de destino y los valores son arrays de filas de datos.
 */
function processData(data, colIndices) {
  const outputData = {};
  const processedRowIndicesBySheet = {}; // Para evitar duplicados por hoja: { sheetName: Set(rowIndex) }

  // Inicializar estructuras de datos
  for (const sheetName in GROUPING_RULES) {
    outputData[sheetName] = [];
    processedRowIndicesBySheet[sheetName] = new Set();
  }

  const headers = data[0]; // Guardamos encabezados para después

  // Iterar sobre las filas de datos (empezar desde 1 para saltar encabezados)
  for (let i = 1; i < data.length; i++) {
    const originalRow = data[i];
    const rowData = [...originalRow]; // Crear una copia para no modificar la original en memoria

    // --- Filtrado Inicial ---
    const publicableValue = rowData[colIndices[PUBLICABLE_COL_NAME]].toString().trim().toUpperCase();
    if (publicableValue === "NO") {
      continue; // Saltar esta fila
    }

    // --- Transformación de Datos ---
    // Respuesta
    const respuestaActualizada = rowData[colIndices[RESPUESTA_ACT_COL_NAME]].toString().trim();
    if (respuestaActualizada !== "") {
      rowData[colIndices[RESPUESTA_COL_NAME]] = respuestaActualizada;
    }
    // Etiqueta
    const etiquetaPropuesta = rowData[colIndices[ETIQUETA_PROP_COL_NAME]].toString().trim();
    if (etiquetaPropuesta !== "") {
      rowData[colIndices[ETIQUETA_ORIG_COL_NAME]] = etiquetaPropuesta;
    }

    // --- Exclusión de Etiqueta Específica ---
    const finalEtiquetaString = rowData[colIndices[ETIQUETA_ORIG_COL_NAME]].toString().trim();
    if (finalEtiquetaString.toLowerCase() === EXCLUDED_TAG.toLowerCase()) {
        continue; // Saltar esta fila si la etiqueta final es exactamente la excluida
    }


    // --- Parseo y Categorización por Etiquetas ---
    const tags = finalEtiquetaString
      .split(TAG_SEPARATOR)
      .map(tag => tag.trim().toLowerCase()) // Convertir a minúsculas para comparación insensible
      .filter(tag => tag !== ""); // Eliminar etiquetas vacías

    if (tags.length === 0) {
        //Logger.log(`Fila ${i+1} sin etiquetas válidas: ${finalEtiquetaString}`);
        continue; // Si no hay etiquetas válidas después de parsear, saltar
    }

    // Revisar cada regla de agrupación
    for (const sheetName in GROUPING_RULES) {
      const keywords = GROUPING_RULES[sheetName].map(kw => kw.toLowerCase()); // Keywords en minúsculas

      let matchFound = false;
      for (const tag of tags) {
        // Si alguna etiqueta de la fila contiene alguna palabra clave de la regla
        if (keywords.some(keyword => tag.includes(keyword))) {
          matchFound = true;
          break; // Encontramos una coincidencia para esta regla, no necesitamos seguir buscando en las tags de esta fila
        }
      }

      // Si hubo coincidencia y la fila AÚN NO ha sido agregada a ESTA hoja específica
      if (matchFound && !processedRowIndicesBySheet[sheetName].has(i)) {
        outputData[sheetName].push(rowData); // Agregar la fila procesada
        processedRowIndicesBySheet[sheetName].add(i); // Marcar esta fila como agregada para esta hoja
      }
    }
  }
  return outputData;
}

/**
 * Crea nuevas hojas o limpia las existentes y escribe los datos procesados.
 * @param {Spreadsheet} ss El objeto Spreadsheet activo.
 * @param {object} categorizedData Datos procesados y agrupados por nombre de hoja.
 * @param {string[]} headers Los encabezados de las columnas.
 * @return {string[]} Array con los nombres de las hojas que fueron creadas o actualizadas.
 */
function createOrClearSheets(ss, categorizedData, headers) {
  const updatedSheetNames = [];
  for (const sheetName in categorizedData) {
    const dataRows = categorizedData[sheetName];

    if (dataRows.length > 0) { // Solo procesar si hay datos para esta categoría
      let targetSheet = ss.getSheetByName(sheetName);
      if (targetSheet) {
        // Si la hoja existe, limpiarla (excepto quizás la primera fila si quieres mantener formato)
        targetSheet.clearContents();
        Logger.log(`Hoja "${sheetName}" limpiada.`);
      } else {
        // Si no existe, crearla
        targetSheet = ss.insertSheet(sheetName);
        Logger.log(`Hoja "${sheetName}" creada.`);
      }

      // Escribir encabezados y datos
      targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold'); // Encabezados en negrita
      targetSheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);

      // Auto-ajustar columnas (opcional)
      targetSheet.autoResizeColumns(1, headers.length);

      updatedSheetNames.push(sheetName);
    } else {
        // Opcional: Si quieres eliminar hojas que quedaron vacías
        let existingSheet = ss.getSheetByName(sheetName);
        if (existingSheet) {
            // ss.deleteSheet(existingSheet);
            // Logger.log(`Hoja "${sheetName}" eliminada por estar vacía.`);
             existingSheet.clearContents(); // O simplemente limpiarla
             Logger.log(`Hoja "${sheetName}" limpiada por estar vacía.`);
        }
    }
  }
  return updatedSheetNames;
}

/**
 * Crea un Google Doc para cada hoja especificada, conteniendo sus datos.
 * @param {Spreadsheet} ss El objeto Spreadsheet activo.
 * @param {string[]} sheetNames Nombres de las hojas de las que crear documentos.
 */
function createDocsFromSheets(ss, sheetNames) {
  const ui = SpreadsheetApp.getUi();
  let folder = null;

  // Preguntar al usuario si desea guardar los documentos en una carpeta específica
  const response = ui.prompt('Creación de Documentos', 'Ingresa el ID de la carpeta de Google Drive donde quieres guardar los documentos (deja vacío para guardar en la raíz de "Mi unidad"):', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const folderId = response.getResponseText().trim();
    if (folderId) {
      try {
        folder = DriveApp.getFolderById(folderId);
        Logger.log(`Documentos se guardarán en la carpeta: ${folder.getName()}`);
      } catch (e) {
        Logger.log(`Error al acceder a la carpeta con ID "${folderId}". Se guardará en la raíz. Error: ${e}`);
        ui.alert(`No se pudo encontrar o acceder a la carpeta con ID "${folderId}". Los documentos se guardarán en la raíz de "Mi unidad".`);
        folder = DriveApp.getRootFolder(); // Guardar en la raíz como fallback
      }
    } else {
        folder = DriveApp.getRootFolder(); // Guardar en la raíz si no se especifica ID
         Logger.log(`Documentos se guardarán en la raíz de "Mi unidad".`);
    }
  } else {
      Logger.log("Creación de documentos cancelada por el usuario.");
      ui.alert("Creación de documentos cancelada.");
      return; // El usuario canceló
  }


  sheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) { // <= 1 porque la fila 1 son encabezados
      Logger.log(`Omitiendo la creación del documento para la hoja "${sheetName}" porque está vacía o no existe.`);
      return; // Saltar si la hoja no existe o está vacía (solo encabezados)
    }

    const data = sheet.getDataRange().getValues();
    const docName = `Reporte - ${sheetName}`; // Nombre del documento

    try {
      // Crear el documento en la carpeta seleccionada (o raíz)
      const doc = DocumentApp.create(docName);
      const docFile = DriveApp.getFileById(doc.getId()); //Obtener el archivo para moverlo
      if (folder) {
          docFile.moveTo(folder); // Mover a la carpeta destino
      }

      const body = doc.getBody();

      // Añadir título al documento
      body.appendParagraph(sheetName).setHeading(DocumentApp.ParagraphHeading.HEADING1);
      body.appendParagraph(""); // Espacio

      // Añadir los datos como una tabla
      // Puedes personalizar qué columnas incluir o cómo formatear aquí
      // Por ahora, incluimos todas las columnas como tabla simple
      const table = body.appendTable(data);

      // Formatear encabezado de la tabla (opcional)
      table.getRow(0).editAsText().setBold(true);

      doc.saveAndClose();
      Logger.log(`Documento "${docName}" creado (ID: ${doc.getId()}) en la carpeta "${folder ? folder.getName() : 'Raíz'}".`);

    } catch (e) {
        Logger.log(`Error al crear el documento para la hoja "${sheetName}": ${e}`);
        // Considera notificar al usuario si falla la creación de un doc específico
        // ui.alert(`Error al crear el documento para "${sheetName}". Detalles en los registros.`);
    }
  });
}


// --- Funciones de Utilidad ---
// (findColumnHeaders ya está arriba)
