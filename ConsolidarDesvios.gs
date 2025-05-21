/**
 * Sistema de consolidación y sincronización de desvíos en la hoja "Inconformidades"
 * Lee los desvíos de todas las hojas de área y los mantiene actualizados en la hoja central
 * 
 * Diseñado para usarse con cualquier tipo de activador:
 * - Al abrir la planilla
 * - Al editar cualquier hoja de área
 * - Basado en tiempo
 * 
 * VERSIÓN OPTIMIZADA: Actualiza específicamente los desvíos modificados
 */

/**
 * Función principal para consolidar desvíos en la hoja "Inconformidades"
 * @param {Object} e - Evento (opcional, proporcionado por activadores)
 */
function consolidarInconformidades(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaInconf = ss.getSheetByName('Inconformidades');
    
    // Verificar si la hoja existe
    if (!hojaInconf) {
      Logger.log("Error: No se encontró la hoja 'Inconformidades'");
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
    
    // Asegurar que la hoja Inconformidades tiene los encabezados correctos
    asegurarEncabezadosInconformidades(hojaInconf);
    
    // Obtener mapa de desvíos actuales en Inconformidades para búsqueda rápida
    var mapaInconformidades = obtenerMapaInconformidades(hojaInconf);
    
    // Recopilar todos los desvíos de todas las hojas de área
    var todosLosDesvios = [];
    var contadorNuevos = 0;
    var contadorActualizados = 0;
    
    // Recorrer cada hoja de área
    for (var h = 0; h < hojasAreas.length; h++) {
      var nombreArea = hojasAreas[h];
      var hojaArea = ss.getSheetByName(nombreArea);
      
      if (!hojaArea) {
        Logger.log("Advertencia: No se encontró la hoja '" + nombreArea + "'");
        continue;
      }
      
      var ultimaFila = hojaArea.getLastRow();
      
      // Si la hoja solo tiene encabezados o está vacía, continuar con la siguiente
      if (ultimaFila <= 1) continue;
      
      // Obtener todos los desvíos de esta hoja
      var datosArea = hojaArea.getRange(2, 1, ultimaFila - 1, 7).getValues();
      
      // Procesar cada desvío
      for (var i = 0; i < datosArea.length; i++) {
        // Verificar que la fila tenga datos (fecha y punto inconforme)
        if (!datosArea[i][0] || !datosArea[i][1]) continue;
        
        // Generar clave única para este desvío
        var clave = generarClaveDesvio(datosArea[i][0], datosArea[i][1], datosArea[i][2]);
        
        // Verificar si ya existe en Inconformidades
        if (clave in mapaInconformidades) {
          // Ya existe, verificar si hay cambios
          var infoExistente = mapaInconformidades[clave];
          var hayDiferencias = false;
          
          // Comparar las columnas D, E, F, G (Resuelta, No resuelta, Fecha de cierre, Comentario)
          if (datosArea[i][3] !== infoExistente.resuelta ||
              datosArea[i][4] !== infoExistente.noResuelta ||
              formatearFecha(datosArea[i][5]) !== formatearFecha(infoExistente.fechaCierre) ||
              datosArea[i][6] !== infoExistente.comentario) {
            
            // Actualizar el registro en Inconformidades
            hojaInconf.getRange(infoExistente.fila, 4, 1, 4).setValues([[
              datosArea[i][3],  // D - Resuelta
              datosArea[i][4],  // E - No resuelta
              datosArea[i][5],  // F - Fecha de cierre
              datosArea[i][6]   // G - Comentario
            ]]);
            
            contadorActualizados++;
          }
        } else {
          // No existe, agregar a la lista de nuevos desvíos
          todosLosDesvios.push(datosArea[i]);
          contadorNuevos++;
        }
      }
    }
    
    // Agregar nuevos desvíos a Inconformidades
    if (todosLosDesvios.length > 0) {
      var ultimaFilaInconf = Math.max(1, hojaInconf.getLastRow());
      var filaInsercion = ultimaFilaInconf;
      
      // Si solo hay una fila con encabezados, comenzar en la fila 2
      if (ultimaFilaInconf === 1) {
        filaInsercion = 2;
      } else {
        filaInsercion = ultimaFilaInconf + 1;
      }
      
      // Agregar todos los nuevos desvíos
      hojaInconf.getRange(filaInsercion, 1, todosLosDesvios.length, 7).setValues(todosLosDesvios);
      
      // Aplicar formato a la columna D (checkbox)
      var rangoCheckbox = hojaInconf.getRange(filaInsercion, 4, todosLosDesvios.length, 1);
      rangoCheckbox.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
      
      // Aplicar formato a la columna F (fechas)
      var rangoFechas = hojaInconf.getRange(filaInsercion, 6, todosLosDesvios.length, 1);
      rangoFechas.setNumberFormat("dd/mm/yyyy");
    }
    
    // Registrar resultado
    var mensaje = "Consolidación de desvíos completada:\n";
    mensaje += "- Se agregaron " + contadorNuevos + " nuevos desvíos\n";
    mensaje += "- Se actualizaron " + contadorActualizados + " desvíos existentes";
    
    Logger.log(mensaje);
    return {
      nuevos: contadorNuevos,
      actualizados: contadorActualizados
    };
  } catch (error) {
    Logger.log("Error en consolidarInconformidades: " + error.toString());
    return null;
  }
}

/**
 * Asegura que la hoja Inconformidades tenga los encabezados correctos
 * @param {Sheet} hoja - La hoja de Inconformidades
 */
function asegurarEncabezadosInconformidades(hoja) {
  // Definir encabezados esperados
  var encabezados = ["Fecha", "Punto inconforme", "Área", "Resuelta", "No resuelta", "Fecha de cierre", "Comentario"];
  
  // Si la hoja está vacía o los encabezados no son correctos
  if (hoja.getLastRow() === 0 || hoja.getRange("A1").getValue() === "") {
    hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
    
    // Aplicar formato a los encabezados
    var rangoEncabezados = hoja.getRange(1, 1, 1, encabezados.length);
    rangoEncabezados.setBackground("#f3f3f3");
    rangoEncabezados.setFontWeight("bold");
  }
}

/**
 * Obtiene un mapa de los desvíos existentes en la hoja Inconformidades
 * @param {Sheet} hoja - La hoja de Inconformidades
 * @return {Object} Mapa de desvíos con su información
 */
function obtenerMapaInconformidades(hoja) {
  var mapa = {};
  var ultimaFila = hoja.getLastRow();
  
  // Si solo hay encabezados o está vacía, devolver mapa vacío
  if (ultimaFila <= 1) return mapa;
  
  // Obtener todos los datos
  var datos = hoja.getRange(2, 1, ultimaFila - 1, 7).getValues();
  
  // Crear mapa para búsqueda rápida
  for (var i = 0; i < datos.length; i++) {
    // Solo procesar filas que tengan fecha, punto y área
    if (datos[i][0] && datos[i][1] && datos[i][2]) {
      var clave = generarClaveDesvio(datos[i][0], datos[i][1], datos[i][2]);
      
      mapa[clave] = {
        fila: i + 2, // Ajustar para índice base 1 y encabezados
        resuelta: datos[i][3],
        noResuelta: datos[i][4],
        fechaCierre: datos[i][5],
        comentario: datos[i][6]
      };
    }
  }
  
  return mapa;
}

/**
 * Genera una clave única para un desvío
 * @param {Date|string} fecha - Fecha del desvío
 * @param {string} punto - Punto inconforme
 * @param {string} area - Área del desvío
 * @return {string} Clave única
 */
function generarClaveDesvio(fecha, punto, area) {
  var fechaStr = formatearFecha(fecha);
  return fechaStr + "||" + punto + "||" + area;
}

/**
 * Formatea una fecha para comparación consistente
 * @param {Date|string} date - La fecha a formatear
 * @return {string} Fecha formateada como YYYY-MM-DD
 */
function formatearFecha(date) {
  if (!date) return "";
  
  // Si ya es una cadena, devolverla como está
  if (typeof date === 'string') {
    return date;
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
 * Actualiza la hoja Inconformidades basándose en cambios en una hoja de área específica
 * OPTIMIZADA: Solo actualiza el desvío específico que cambió
 * @param {Object} e - Evento de edición
 */
function actualizarDesdeCambioArea(e) {
  // Verificar si el evento tiene información necesaria
  if (!e || !e.source || !e.range) return;
  
  // Obtener hoja activa
  var hoja = e.range.getSheet();
  var nombreHoja = hoja.getName();
  
  // Lista de hojas de áreas para verificar
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
  
  // Verificar si la edición fue en una hoja de área
  if (hojasAreas.indexOf(nombreHoja) === -1) return;
  
  // Verificar si la edición fue en columnas D, F o G (Resuelta, Fecha de cierre, Comentario)
  var columna = e.range.getColumn();
  if (columna !== 4 && columna !== 6 && columna !== 7) return;
  
  // OPTIMIZACIÓN: Solo actualizar el desvío específico que cambió
  var fila = e.range.getRow();
  if (fila <= 1) return; // Ignorar cambios en encabezados
  
  try {
    var hojaArea = e.range.getSheet();
    var datosDesvio = hojaArea.getRange(fila, 1, 1, 7).getValues()[0];
    
    // Verificar que sea un desvío válido (tiene fecha y punto)
    if (!datosDesvio[0] || !datosDesvio[1]) return;
    
    // Actualizar solo este desvío específico en Inconformidades
    actualizarDesvioEspecifico(datosDesvio);
    
    Logger.log("Actualizado desvío específico desde fila " + fila + " en " + nombreHoja);
  } catch (error) {
    Logger.log("Error al actualizar desvío específico: " + error.toString());
    // Si hay algún error, ejecutar la consolidación completa como respaldo
    consolidarInconformidades();
  }
}

/**
 * Actualiza un desvío específico en la hoja Inconformidades
 * Parte de la optimización para actualizar solo lo que cambió
 * @param {Array} datosDesvio - Datos del desvío a actualizar
 */
function actualizarDesvioEspecifico(datosDesvio) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaInconf = ss.getSheetByName('Inconformidades');
  
  if (!hojaInconf) return;
  
  // Buscar el desvío en Inconformidades
  var ultimaFila = hojaInconf.getLastRow();
  if (ultimaFila <= 1) return;
  
  var clave = generarClaveDesvio(datosDesvio[0], datosDesvio[1], datosDesvio[2]);
  var datos = hojaInconf.getRange(2, 1, ultimaFila - 1, 7).getValues();
  
  for (var i = 0; i < datos.length; i++) {
    if (datos[i][0] && datos[i][1] && datos[i][2]) {
      var claveActual = generarClaveDesvio(datos[i][0], datos[i][1], datos[i][2]);
      
      if (claveActual === clave) {
        // Encontrado - actualizar solo las columnas D, E, F, G
        hojaInconf.getRange(i + 2, 4, 1, 4).setValues([[
          datosDesvio[3],  // D - Resuelta
          datosDesvio[4],  // E - No resuelta
          datosDesvio[5],  // F - Fecha de cierre
          datosDesvio[6]   // G - Comentario
        ]]);
        
        Logger.log("Desvío actualizado en fila " + (i + 2) + " de Inconformidades");
        return;
      }
    }
  }
  
  // Si no se encontró, agregarlo al final
  hojaInconf.appendRow(datosDesvio);
  var nuevaFila = hojaInconf.getLastRow();
  
  // Aplicar formato a la columna D (checkbox)
  hojaInconf.getRange(nuevaFila, 4).setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().build()
  );
  
  // Aplicar formato a la columna F (fechas)
  hojaInconf.getRange(nuevaFila, 6).setNumberFormat("dd/mm/yyyy");
  
  Logger.log("Desvío agregado en fila " + nuevaFila + " de Inconformidades");
}

/**
 * Función para agregar al menú existente
 * Puede llamarse desde tu función onOpen actual
 */
function agregarMenuConsolidacion() {
  var ui = SpreadsheetApp.getUi();
  // Verificar si ya existe el menú
  try {
    ui.createMenu('Gestión de Desvíos')
        .addItem('Distribuir Desvíos a Áreas', 'distribuirDesviosAreas')
        .addItem('Consolidar en Inconformidades', 'consolidarInconformidades')
        .addItem('REINICIAR - Forzar Redistribución', 'forzarDistribucionDesvios')
        .addToUi();
  } catch (e) {
    // Si ya existe, añadir solo el ítem de consolidación
    try {
      var menu = ui.getMenu('Gestión de Desvíos');
      menu.addItem('Consolidar en Inconformidades', 'consolidarInconformidades');
    } catch (e2) {
      // Si no se puede modificar el menú existente, crear uno nuevo
      ui.createMenu('Consolidación')
        .addItem('Consolidar Desvíos', 'consolidarInconformidades')
        .addToUi();
    }
  }
}
