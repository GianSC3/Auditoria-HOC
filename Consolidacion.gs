//Esta funcion solo se ejecuta una vez. Si la hoja "Consolidada" ya existe, no se debe ejecutar.
function createConsolidatedSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawSheet = ss.getSheetByName("Respuestas de formulario 1"); // Nombre de la hoja de respuestas crudas
  var consolidatedSheet = ss.getSheetByName("Consolidada");

  var preguntasMaestrasSheet = ss.getSheetByName("Preguntas maestras"); // Hoja con las preguntas maestras
  var preguntasMaestras = preguntasMaestrasSheet.getRange(1, 1, preguntasMaestrasSheet.getLastRow(), 1).getValues().flat();

  if (!consolidatedSheet) {
    consolidatedSheet = ss.insertSheet("Consolidada"); // Crear una nueva hoja "Consolidada" vacía en lugar de copiar la existente
    var rawData = rawSheet.getDataRange().getValues();
    var headers = rawData[0];

    var consolidatedData = [];
    var questionMap = {}; // Mapeo para almacenar las columnas consolidadas

    // Crear encabezados consolidados a partir de la columna G
    for (var i = 6; i < headers.length; i++) {
      var rawQuestion = headers[i].replace(/^\d+\.\s*/, '').trim(); // Ignorar el índice (números, punto y espacio)
      if (rawQuestion === "Subir fotos de desvíos / evidencias") continue; // Excluir la pregunta específica
      var preguntaMaestra = preguntasMaestras.find(p => p.includes(rawQuestion));
      if (preguntaMaestra) {
        rawQuestion = preguntaMaestra;
      }
      if (!questionMap[rawQuestion]) {
        questionMap[rawQuestion] = [];
      }
      questionMap[rawQuestion].push(i);
    }

    consolidatedData.push(headers.slice(0, 6).concat(Object.keys(questionMap))); // Mantener las primeras 6 columnas y agregar encabezados consolidados sin índices

    // Consolidar las respuestas
    for (var row = 1; row < rawData.length; row++) {
      var consolidatedRow = rawData[row].slice(0, 6); // Mantener las primeras 6 columnas
      for (var question in questionMap) {
        var values = questionMap[question].map(colIndex => rawData[row][colIndex]);
        var nonEmptyValue = values.find(value => value !== "");
        consolidatedRow.push(nonEmptyValue || "");
      }
      consolidatedData.push(consolidatedRow);
    }

    // Guardar los datos consolidados en la nueva hoja
    consolidatedSheet.getRange(1, 1, consolidatedData.length, consolidatedData[0].length).setValues(consolidatedData);
  }
}

//Esta función actualiza automáticamente la hoja "Consolidada" con las nuevas auditorías.
function updateConsolidatedSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rawSheet = ss.getSheetByName("Respuestas de formulario 1"); // Nombre de la hoja de respuestas crudas
    var consolidatedSheet = ss.getSheetByName("Consolidada");

    var preguntasMaestrasSheet = ss.getSheetByName("Preguntas maestras"); // Hoja con las preguntas maestras
    var preguntasMaestras = preguntasMaestrasSheet.getRange(1, 1, preguntasMaestrasSheet.getLastRow(), 1).getValues().flat();

    if (consolidatedSheet) {
      // GUARDAMOS LAS FÓRMULAS DE DI, DJ, DK, DL ANTES DE HACER CAMBIOS
      var columnaDI = consolidatedSheet.getRange("DI1").getColumn(); // Obtener el número exacto de columna
      var columnaDJ = columnaDI + 1;  // Columna DJ
      var columnaDK = columnaDI + 2;  // Columna DK
      var columnaDL = columnaDI + 3;  // Columna DL
      var ultimaFila = consolidatedSheet.getLastRow();
      
      // Inicializar los arrays para guardar las fórmulas
      var formulasDI = [];
      var formulasDJ = [];
      var formulasDK = [];
      var formulasDL = [];
      
      if (ultimaFila > 1) {
        // Guardar todas las fórmulas ANTES de la actualización
        formulasDI = consolidatedSheet.getRange(2, columnaDI, ultimaFila - 1, 1).getFormulas();
        formulasDJ = consolidatedSheet.getRange(2, columnaDJ, ultimaFila - 1, 1).getFormulas();
        formulasDK = consolidatedSheet.getRange(2, columnaDK, ultimaFila - 1, 1).getFormulas();
        formulasDL = consolidatedSheet.getRange(2, columnaDL, ultimaFila - 1, 1).getFormulas();
      }
      
      var headers = consolidatedSheet.getRange(1, 1, 1, consolidatedSheet.getLastColumn()).getValues()[0]; // Mantener los encabezados actuales
      var rawData = rawSheet.getDataRange().getValues();

      var consolidatedData = [];
      var questionMap = {}; // Mapeo para almacenar las columnas consolidadas

      // Crear mapeo de preguntas consolidadas a partir de la columna G
      for (var i = 6; i < rawData[0].length; i++) {
        var rawQuestion = rawData[0][i].replace(/^\d+\.\s*/, '').trim(); // Ignorar el índice (números, punto y espacio)
        if (rawQuestion === "Subir fotos de desvíos / evidencias") continue; // Excluir la pregunta específica
        var preguntaMaestra = preguntasMaestras.find(p => p.includes(rawQuestion));
        if (preguntaMaestra) {
          rawQuestion = preguntaMaestra;
        }
        if (!questionMap[rawQuestion]) {
          questionMap[rawQuestion] = [];
        }
        questionMap[rawQuestion].push(i);
      }

      // Consolidar las respuestas
      for (var row = 1; row < rawData.length; row++) {
        var consolidatedRow = rawData[row].slice(0, 6); // Mantener las primeras 6 columnas
        for (var question in questionMap) {
          var values = questionMap[question].map(colIndex => rawData[row][colIndex]);
          var nonEmptyValue = values.find(value => value !== "");
          consolidatedRow.push(nonEmptyValue || "");
        }
        consolidatedData.push(consolidatedRow);
      }

      // Obtener el formato existente de las celdas (alineación y ajuste de texto)
      var rangeToUpdate = consolidatedSheet.getRange(2, 1, consolidatedData.length, headers.length);
      var existingAlignments = rangeToUpdate.getHorizontalAlignments();
      var existingVerticalAlignments = rangeToUpdate.getVerticalAlignments();
      var existingWraps = rangeToUpdate.getWrapStrategies();

      // Actualizar los datos consolidados en la hoja sin afectar los encabezados
      consolidatedSheet.getRange(2, 1, consolidatedData.length, consolidatedData[0].length).setValues(consolidatedData);

      // Restaurar el formato de alineación y ajuste de texto
      rangeToUpdate.setHorizontalAlignments(existingAlignments);
      rangeToUpdate.setVerticalAlignments(existingVerticalAlignments);
      rangeToUpdate.setWrapStrategies(existingWraps);

      // FUNCIÓN AUXILIAR PARA RESTAURAR Y REPLICAR FÓRMULAS
      function restaurarYReplicarFormulas(formulas, columna) {
        if (formulas.length > 0) {
          // Restaurar fórmulas en filas existentes
          var filasActuales = Math.min(formulas.length, consolidatedData.length);
          consolidatedSheet.getRange(2, columna, filasActuales, 1).setFormulas(formulas.slice(0, filasActuales));
          
          // Replicar la fórmula para nuevas filas
          if (consolidatedData.length > formulas.length && formulas.length > 0) {
            var formula = formulas[0][0]; // Tomar la primera fórmula como modelo
            if (formula) {
              for (var i = formulas.length; i < consolidatedData.length; i++) {
                var rowNum = i + 2; // Ajuste para empezar en la fila 2
                var newFormula = formula.replace(/2/g, rowNum); // Reemplazar todos los "2" con el número de fila actual
                consolidatedSheet.getRange(rowNum, columna).setFormula(newFormula);
              }
            }
          }
        }
      }
      
      // RESTAURAR Y REPLICAR LAS FÓRMULAS DE CADA COLUMNA
      restaurarYReplicarFormulas(formulasDI, columnaDI);  // Columna DI
      restaurarYReplicarFormulas(formulasDJ, columnaDJ);  // Columna DJ 
      restaurarYReplicarFormulas(formulasDK, columnaDK);  // Columna DK
      restaurarYReplicarFormulas(formulasDL, columnaDL);  // Columna DL
      
      Logger.log("Actualización completada preservando columnas DI, DJ, DK y DL");
    }
  } catch (error) {
    Logger.log("Error en updateConsolidatedSheet: " + error.toString());
  }
}

function onOpen(e) {
  updateConsolidatedSheet();
}


//Esta función comprueba si es necesario actualizar los permisos de ejecución del Script. De ser necesario, se reporta al email especificado en la última línea de código. 
function checkPermissions() {
  try {
    // Intenta acceder a un recurso que requiere permisos
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rawSheet = ss.getSheetByName("Respuestas de formulario 1");
    var data = rawSheet.getDataRange().getValues();
  } catch (error) {
    MailApp.sendEmail("gilserra@ccu.com.ar", "Error de permisos en Google Sheets", "Hubo un problema con los permisos del script: " + error.message);
  }
}
function createTimeTrigger() {
  ScriptApp.newTrigger('checkPermissions')
    .timeBased()
    .everyDays(1) // Configura para que se ejecute diariamente
    .create();
}

