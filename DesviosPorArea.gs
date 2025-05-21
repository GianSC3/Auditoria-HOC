/**
 * Sistema de distribución de desvíos desde Consolidada a hojas de área
 * Detecta "No Cumple 100%" y los distribuye preservando información existente
 * 
 * Diseñado para usar directamente con cualquier tipo de activador:
 * - Al abrir
 * - Al editar
 * - Basado en tiempo
 */

/**
 * Función principal para distribuir desvíos a las hojas de área
 * Diseñada para usarse directamente como activador
 * @param {Object} e - Evento (opcional, proporcionado por activadores)
 */
function distribuirDesviosAreas(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var consolidadaSheet = ss.getSheetByName('Consolidada');
    
    // Verificar si la hoja existe
    if (!consolidadaSheet) {
      Logger.log("Error: No se encontró la hoja 'Consolidada'");
      return;
    }
    
    // Lista de hojas de áreas
    var hojasAreas = [
      "Almacén",  
      "Áreas Comunes",  
      "Elaboración 1 - Planta de gas",  
      "Elaboración 2 - Bodega",  
      "Elaboración 3 - Filtro",  
      "Elaboración 4 - Cocina",  
      "Elaboración 5 - Molino",  
      "Laboratorio",  
      "L2",  
      "L3",  
      "L4",  
      "L5",  
      "Mantenimiento",  
      "Movimiento Interno",  
      "Planta de Agua",  
      "Sala de Máquinas"
    ];
    
    // Obtener datos de la hoja Consolidada
    var lastRow = consolidadaSheet.getLastRow();
    var lastColumn = consolidadaSheet.getLastColumn();
    
    // Si no hay datos, salir
    if (lastRow <= 1) {
      Logger.log("No hay datos en la hoja Consolidada");
      return;
    }
    
    // Obtener encabezados (nombres de los puntos inconforme)
    var encabezados = consolidadaSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    
    // Obtener todos los datos
    var datosConsolidada = consolidadaSheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
    
    // Contador para estadísticas
    var estadisticas = {
      totalDesvios: 0,
      desviosNuevos: 0,
      desviosPorArea: {}
    };
    
    // Inicializar contador por área
    for (var h = 0; h < hojasAreas.length; h++) {
      estadisticas.desviosPorArea[hojasAreas[h]] = 0;
    }
    
    // Procesar los desvíos para cada área
    for (var h = 0; h < hojasAreas.length; h++) {
      var nombreArea = hojasAreas[h];
      var hojaArea = ss.getSheetByName(nombreArea);
      
      // Verificar si la hoja de área existe
      if (!hojaArea) {
        Logger.log("Advertencia: No se encontró la hoja '" + nombreArea + "'");
        continue;
      }
      
      // Obtener mapa de desvíos existentes en esta área
      var desviosExistentes = obtenerDesviosExistentes(hojaArea);
      
      // Extraer desvíos para esta área desde Consolidada
      var nuevosDesvios = extraerDesviosArea(datosConsolidada, encabezados, nombreArea);
      
      // Filtrar solo los desvíos que no existen
      var desviosAgregar = [];
      for (var i = 0; i < nuevosDesvios.length; i++) {
        var clave = generarClaveDesvio(nuevosDesvios[i][0], nuevosDesvios[i][1]);
        if (!desviosExistentes[clave]) {
          // Añadir el desvío con columnas vacías para D, F y G
          // La columna E tendrá la fórmula que se aplicará después
          desviosAgregar.push([
            nuevosDesvios[i][0], // Fecha
            nuevosDesvios[i][1], // Punto inconforme
            nuevosDesvios[i][2], // Área
            false,               // Resuelta (inicialmente falso)
            "",                  // No resuelta (se aplicará fórmula después)
            "",                  // Fecha de cierre (vacío)
            ""                   // Comentario (vacío)
          ]);
          
          estadisticas.desviosNuevos++;
          estadisticas.desviosPorArea[nombreArea]++;
        }
        estadisticas.totalDesvios++;
      }
      
      // Agregar los nuevos desvíos a la hoja de área
      if (desviosAgregar.length > 0) {
        // Obtener última fila con datos
        var ultimaFila = Math.max(1, hojaArea.getLastRow());
        
        // Si la hoja está vacía, crear encabezados
        if (ultimaFila === 1 && hojaArea.getRange("A1").getValue() === "") {
          var encabezadosArea = ["Fecha", "Punto inconforme", "Área", "Resuelta", "No resuelta", "Fecha de cierre", "Comentario"];
          hojaArea.getRange(1, 1, 1, encabezadosArea.length).setValues([encabezadosArea]);
          
          // Aplicar formato a los encabezados
          var rangoEncabezados = hojaArea.getRange(1, 1, 1, encabezadosArea.length);
          rangoEncabezados.setBackground("#f3f3f3");
          rangoEncabezados.setFontWeight("bold");
        }
        
        // Calcular fila donde agregar (después de la última fila con datos)
        var filaAgregacion = ultimaFila + 1;
        if (ultimaFila === 1 && hojaArea.getRange("A1").getValue() !== "") {
          filaAgregacion = 2; // La primera fila tiene encabezados
        } else if (ultimaFila === 1 && hojaArea.getRange("A1").getValue() === "") {
          filaAgregacion = 2; // Agregamos después de encabezados recién creados
        }
        
        // Agregar los desvíos
        hojaArea.getRange(filaAgregacion, 1, desviosAgregar.length, desviosAgregar[0].length).setValues(desviosAgregar);
        
        // Aplicar fórmulas en la columna E (No resuelta)
        var formulas = [];
        for (var i = 0; i < desviosAgregar.length; i++) {
          formulas.push(['=IF(COUNTA(B' + (filaAgregacion + i) + ')=0,"",IF(D' + (filaAgregacion + i) + '=FALSE,"Si",""))']);
        }
        hojaArea.getRange(filaAgregacion, 5, desviosAgregar.length, 1).setFormulas(formulas);
        
        Logger.log("Agregados " + desviosAgregar.length + " nuevos desvíos a la hoja '" + nombreArea + "'");
      } else {
        Logger.log("No hay nuevos desvíos para agregar a la hoja '" + nombreArea + "'");
      }
    }
    
    // Registrar estadísticas
    Logger.log("Resumen de distribución de desvíos:");
    Logger.log("- Total de desvíos procesados: " + estadisticas.totalDesvios);
    Logger.log("- Nuevos desvíos agregados: " + estadisticas.desviosNuevos);
    for (var area in estadisticas.desviosPorArea) {
      if (estadisticas.desviosPorArea[area] > 0) {
        Logger.log("  - " + area + ": " + estadisticas.desviosPorArea[area] + " nuevos desvíos");
      }
    }
    
    return estadisticas;
  } catch (error) {
    Logger.log("Error en distribuirDesviosAreas: " + error.toString());
    return null;
  }
}

/**
 * Obtiene un mapa de los desvíos existentes en una hoja de área
 * @param {Sheet} hojaArea - La hoja de área a analizar
 * @return {Object} Mapa de desvíos existentes
 */
function obtenerDesviosExistentes(hojaArea) {
  var mapa = {};
  var ultimaFila = hojaArea.getLastRow();
  
  // Si la hoja está vacía o solo tiene encabezados, devolver mapa vacío
  if (ultimaFila <= 1) return mapa;
  
  // Obtener datos de las columnas A y B (Fecha y Punto inconforme)
  var datos = hojaArea.getRange(2, 1, ultimaFila - 1, 2).getValues();
  
  // Crear mapa para búsqueda rápida
  for (var i = 0; i < datos.length; i++) {
    // Solo procesar filas que tengan fecha y punto
    if (datos[i][0] && datos[i][1]) {
      var clave = generarClaveDesvio(datos[i][0], datos[i][1]);
      mapa[clave] = true;
    }
  }
  
  return mapa;
}

/**
 * Extrae los desvíos para un área específica de la hoja Consolidada
 * @param {Array} datosConsolidada - Datos de la hoja Consolidada
 * @param {Array} encabezados - Encabezados de las columnas
 * @param {string} area - Nombre del área a filtrar
 * @return {Array} Lista de desvíos para el área
 */
function extraerDesviosArea(datosConsolidada, encabezados, area) {
  var desvios = [];
  
  // Para cada fila en la hoja Consolidada
  for (var i = 0; i < datosConsolidada.length; i++) {
    // Verificar si la fila corresponde al área buscada
    if (datosConsolidada[i][2] === area) { // La columna C (índice 2) contiene el área
      
      // Obtener la fecha (columna DL) - asumiendo que es la última columna
      var fecha = datosConsolidada[i][encabezados.length - 1];
      
      // Recorrer las columnas desde G hasta DH (índices 6 a 110)
      for (var j = 6; j < 111; j++) {
        // Verificar si es un desvío ("No Cumple 100%")
        if (datosConsolidada[i][j] === "No Cumple 100%") {
          // Obtener el nombre del punto inconforme (encabezado de la columna)
          var puntoInconforme = encabezados[j];
          
          // Agregar a la lista de desvíos
          desvios.push([fecha, puntoInconforme, area]);
        }
      }
    }
  }
  
  return desvios;
}

/**
 * Genera una clave única para un desvío basada en fecha y punto
 * @param {Date|string} fecha - Fecha del desvío
 * @param {string} punto - Punto inconforme
 * @return {string} Clave única
 */
function generarClaveDesvio(fecha, punto) {
  var fechaStr = formatearFecha(fecha);
  return fechaStr + "||" + punto;
}

/**
 * Formatea una fecha para comparación consistente
 * @param {Date|string} date - La fecha a formatear
 * @return {string} Fecha formateada como YYYY-MM-DD
 */
function formatearFecha(date) {
  if (!date) return "";
  
  // Si ya es una cadena, intentar convertir a fecha
  if (typeof date === 'string') {
    return date; // Mantener como está para comparación
  }
  
  // Si es un objeto Date, formatear
  if (date instanceof Date) {
    try {
      return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } catch (e) {
      return date.toString();
    }
  }
  
  // En cualquier otro caso, convertir a string
  return date.toString();
}

/**
 * Función para crear menú personalizado
 * Puede llamarse desde tu función onOpen existente
 */
function crearMenuDesvios() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Gestión de Desvíos')
      .addItem('Distribuir Desvíos a Áreas', 'distribuirDesviosAreas')
      .addToUi();
}

/**
 * Función para configurar activador basado en tiempo
 * Útil para ejecuciones programadas
 */
function crearActivadorTiempo() {
  // Eliminar activadores existentes para evitar duplicados
  var activadores = ScriptApp.getProjectTriggers();
  for (var i = 0; i < activadores.length; i++) {
    if (activadores[i].getHandlerFunction() === 'distribuirDesviosAreas') {
      ScriptApp.deleteTrigger(activadores[i]);
    }
  }
  
  // Crear nuevo activador (diariamente a las 8am)
  ScriptApp.newTrigger('distribuirDesviosAreas')
      .timeBased()
      .atHour(8)
      .everyDays(1)
      .create();
  
  Logger.log("Activador programado creado exitosamente");
}

/**
 * Función para forzar la redistribución de todos los desvíos
 * Útil después de borrar manualmente datos de las hojas
 */
function forzarDistribucionDesvios() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var consolidadaSheet = ss.getSheetByName('Consolidada');
    
    // Verificar si la hoja existe
    if (!consolidadaSheet) {
      Logger.log("Error: No se encontró la hoja 'Consolidada'");
      return;
    }
    
    // Lista de hojas de áreas
    var hojasAreas = [
      "Almacén",  
      "Áreas Comunes",  
      "Elaboración 1 - Planta de gas",  
      "Elaboración 2 - Bodega",  
      "Elaboración 3 - Filtro",  
      "Elaboración 4 - Cocina",  
      "Elaboración 5 - Molino",  
      "Laboratorio",  
      "L2",  
      "L3",  
      "L4",  
      "L5",  
      "Mantenimiento",  
      "Movimiento Interno",  
      "Planta de Agua",  
      "Sala de Máquinas"
    ];
    
    // Obtener datos de la hoja Consolidada
    var lastRow = consolidadaSheet.getLastRow();
    var lastColumn = consolidadaSheet.getLastColumn();
    
    // Si no hay datos, salir
    if (lastRow <= 1) {
      Logger.log("No hay datos en la hoja Consolidada");
      return;
    }
    
    // Obtener encabezados (nombres de los puntos inconforme)
    var encabezados = consolidadaSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    
    // Obtener todos los datos
    var datosConsolidada = consolidadaSheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
    
    // Contador para estadísticas
    var estadisticas = {
      totalDesvios: 0,
      desviosPorArea: {}
    };
    
    // Inicializar contador por área
    for (var h = 0; h < hojasAreas.length; h++) {
      estadisticas.desviosPorArea[hojasAreas[h]] = 0;
    }
    
    // Procesar los desvíos para cada área
    for (var h = 0; h < hojasAreas.length; h++) {
      var nombreArea = hojasAreas[h];
      var hojaArea = ss.getSheetByName(nombreArea);
      
      // Verificar si la hoja de área existe
      if (!hojaArea) {
        Logger.log("Advertencia: No se encontró la hoja '" + nombreArea + "'");
        continue;
      }
      
      // DIFERENCIA CON LA VERSIÓN ORIGINAL: Aquí no obtenemos los desvíos existentes
      // No usaremos obtenerDesviosExistentes(hojaArea)
      
      // Extraer desvíos para esta área desde Consolidada
      var desvios = extraerDesviosArea(datosConsolidada, encabezados, nombreArea);
      
      // Prepararemos TODOS los desvíos para agregarlos
      var desviosAgregar = [];
      for (var i = 0; i < desvios.length; i++) {
        // Añadir el desvío con columnas vacías para D, F y G
        desviosAgregar.push([
          desvios[i][0], // Fecha
          desvios[i][1], // Punto inconforme
          desvios[i][2], // Área
          false,         // Resuelta (inicialmente falso)
          "",            // No resuelta (se aplicará fórmula después)
          "",            // Fecha de cierre (vacío)
          ""             // Comentario (vacío)
        ]);
        
        estadisticas.desviosPorArea[nombreArea]++;
        estadisticas.totalDesvios++;
      }
      
      // Configurar la hoja con los nuevos desvíos
      if (desviosAgregar.length > 0) {
        // Si la hoja está vacía, crear encabezados
        var encabezadosArea = ["Fecha", "Punto inconforme", "Área", "Resuelta", "No resuelta", "Fecha de cierre", "Comentario"];
        
        // Limpiar la hoja PERO preservar la primera fila (encabezados)
        if (hojaArea.getLastRow() > 1) {
          hojaArea.getRange(2, 1, hojaArea.getLastRow() - 1, 7).clearContent();
        }
        
        // Asegurar que existen los encabezados
        if (hojaArea.getLastRow() == 0 || hojaArea.getRange("A1").getValue() === "") {
          hojaArea.getRange(1, 1, 1, encabezadosArea.length).setValues([encabezadosArea]);
          
          // Aplicar formato a los encabezados
          var rangoEncabezados = hojaArea.getRange(1, 1, 1, encabezadosArea.length);
          rangoEncabezados.setBackground("#f3f3f3");
          rangoEncabezados.setFontWeight("bold");
        }
        
        // Agregar los desvíos empezando en la fila 2 (después de encabezados)
        hojaArea.getRange(2, 1, desviosAgregar.length, desviosAgregar[0].length).setValues(desviosAgregar);
        
        // Aplicar fórmulas en la columna E (No resuelta)
        var formulas = [];
        for (var i = 0; i < desviosAgregar.length; i++) {
          formulas.push(['=IF(COUNTA(B' + (2 + i) + ')=0,"",IF(D' + (2 + i) + '=FALSE,"Si",""))']);
        }
        hojaArea.getRange(2, 5, desviosAgregar.length, 1).setFormulas(formulas);
        
        Logger.log("Se redistribuyeron " + desviosAgregar.length + " desvíos a la hoja '" + nombreArea + "'");
      } else {
        Logger.log("No hay desvíos para el área '" + nombreArea + "'");
      }
    }
    
    // Registrar estadísticas
    Logger.log("Resumen de redistribución forzada:");
    Logger.log("- Total de desvíos redistribuidos: " + estadisticas.totalDesvios);
    for (var area in estadisticas.desviosPorArea) {
      if (estadisticas.desviosPorArea[area] > 0) {
        Logger.log("  - " + area + ": " + estadisticas.desviosPorArea[area] + " desvíos");
      }
    }
    
    // Mostrar mensaje de confirmación
    SpreadsheetApp.getUi().alert("Se han redistribuido " + estadisticas.totalDesvios + " desvíos a sus hojas correspondientes.");
    
    return estadisticas;
  } catch (error) {
    Logger.log("Error en forzarDistribucionDesvios: " + error.toString());
    SpreadsheetApp.getUi().alert("Error al redistribuir desvíos: " + error.toString());
    return null;
  }
}

function crearMenuDesvios() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Gestión de Desvíos')
      .addItem('Distribuir Desvíos a Áreas', 'distribuirDesviosAreas')
      .addItem('REINICIAR - Forzar Redistribución', 'forzarDistribucionDesvios')
      .addToUi();
}
