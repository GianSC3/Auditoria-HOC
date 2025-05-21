/**
 * Sistema completo de gestión de desvíos
 * Incluye procesamiento de evidencias, distribución a hojas por área y consolidación
 */

//=====================================================================
// PARTE 1: PROCESAMIENTO DE EVIDENCIAS DE DESVÍOS
//=====================================================================

/**
 * Procesa evidencias de desvíos desde "Respuestas de formulario 1" a "EvidenciasDesvios"
 * Versión mejorada: Evita duplicados, formato de fecha sin ceros iniciales, centrado
 * También detecta y crea registros específicos para auditorías sin desvíos reales
 */
function procesarEvidenciasDesvios(mostrarAlertas) {
  // Si no se especifica, no mostrar alertas por defecto (para activadores automáticos)
  mostrarAlertas = (mostrarAlertas === undefined) ? false : mostrarAlertas;
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Obtener o crear hoja EvidenciasDesvios
    var hojaDesvios = ss.getSheetByName("EvidenciasDesvios");
    if (!hojaDesvios) {
      hojaDesvios = ss.insertSheet("EvidenciasDesvios");
      configurarEncabezados(hojaDesvios);
    }
    
    // Obtener hoja de respuestas
    var hojaRespuestas = ss.getSheetByName("Respuestas de formulario 1");
    if (!hojaRespuestas) {
      throw new Error("No se encontró la hoja 'Respuestas de formulario 1'");
    }
    
    // Obtener datos y encabezados
    var datosRespuestas = hojaRespuestas.getDataRange().getValues();
    var encabezados = datosRespuestas[0];
    
    // Identificar índices de columnas por encabezados
    var idxMes = encontrarIndiceEncabezado(encabezados, "Mes correspondiente");
    var idxAnio = encontrarIndiceEncabezado(encabezados, "Año");
    var idxArea = encontrarIndiceEncabezado(encabezados, "Área auditada");
    var idxMarcaTemporal = encontrarIndiceEncabezado(encabezados, "Marca temporal");
    
    // Encontrar índices de observaciones y evidencias
    var indicesObservaciones = [];
    var indicesEvidencias = [];
    
    for (var i = 1; i <= 10; i++) {
      var idxObs = encontrarIndiceEncabezado(encabezados, "Observaciones / Sugerencias de la Evidencia " + i);
      var idxEvid = encontrarIndiceEncabezado(encabezados, "Evidencia " + i);
      
      if (idxObs >= 0 && idxEvid >= 0) {
        indicesObservaciones.push(idxObs);
        indicesEvidencias.push(idxEvid);
      }
    }
    
    // Verificar que se encontraron las columnas necesarias
    if (idxMes < 0 || idxAnio < 0 || idxArea < 0) {
      throw new Error("No se encontraron todas las columnas necesarias. " +
                     "Asegúrese de tener: 'Mes correspondiente', 'Año' y 'Área auditada'");
    }
    
    // Obtener evidencias existentes para evitar duplicados
    var desviosExistentes = {};
    if (hojaDesvios.getLastRow() > 1) {
      var datosExistentes = hojaDesvios.getRange(2, 1, hojaDesvios.getLastRow() - 1, 5).getValues();
      
      // Crear un mapa para verificación rápida de duplicados
      for (var i = 0; i < datosExistentes.length; i++) {
        var fecha = datosExistentes[i][0];
        var area = datosExistentes[i][1];
        var desvio = datosExistentes[i][2];
        
        // Crear clave única: fecha+area+desvio
        var clave = crearClaveEvidencia(fecha, area, desvio);
        desviosExistentes[clave] = true;
      }
    }
    
    // Procesar datos y generar nuevas filas
    var nuevosDesvios = [];
    var totalProcesados = 0;
    var registrosOmitidos = 0;
    var duplicadosEvitados = 0;
    var sinDesviosReales = 0;  // Contador para auditorías sin desvíos reales
    
    // Fecha límite: Mayo 2025
    var mesFiltro = 5;
    var anioFiltro = 2025;
    
    // Para cada respuesta del formulario
    for (var i = 1; i < datosRespuestas.length; i++) {
      var fila = datosRespuestas[i];
      
      // Extraer información de fecha
      var mes = fila[idxMes];
      var anio = fila[idxAnio];
      
      // FILTRO TEMPORAL: Verificar si la fecha es igual o posterior a mayo 2025
      if ((anio > anioFiltro) || (anio == anioFiltro && mes >= mesFiltro)) {
        // Crear fecha (primer día del mes)
        var fecha = new Date(anio, mes - 1, 1); // Los meses en JavaScript son 0-11
        
        // Extraer área
        var area = fila[idxArea];
        
        // NUEVO: Verificar si la primera observación indica "sin desvíos"
        var primeraObservacion = fila[indicesObservaciones[0]];
        var esAuditoriaCompleta = false;
        
        // Comprobar si es una respuesta de "sin desvíos"
        if (primeraObservacion) {
          var observacionLimpia = primeraObservacion.trim();
          if (observacionLimpia === " " || 
              observacionLimpia === "-" || 
              observacionLimpia === "_" || 
              observacionLimpia === "." ||
              observacionLimpia === '"' ||
              observacionLimpia === "'") {
            // Esta es una auditoría 100% completa, sin desvíos reales
            esAuditoriaCompleta = true;
            sinDesviosReales++;
            
            // NUEVO: Agregar un registro explícito que indique que no hay desvíos
            var claveDesvioEspecial = crearClaveEvidencia(fecha, area, "No se presentan desvíos");
            if (!desviosExistentes[claveDesvioEspecial]) {
              nuevosDesvios.push([
                fecha,                    // Fecha
                area,                     // Área
                "No se presentan desvíos", // Desvío (observación especial)
                ""                        // Link (vacío)
              ]);
              
              totalProcesados++;
              desviosExistentes[claveDesvioEspecial] = true;
            }
            
            // Ya agregamos el registro especial, continuamos con la siguiente respuesta
            continue;
          }
        }
        
        // Procesar cada par de observación/evidencia solo si no es una auditoría 100% completa
        if (!esAuditoriaCompleta) {
          for (var j = 0; j < indicesObservaciones.length; j++) {
            var observacion = fila[indicesObservaciones[j]];
            var evidencia = fila[indicesEvidencias[j]];
            
            // MODIFICADO: Filtrar observaciones vacías o caracteres especiales
            if (observacion && observacion.trim() !== "" &&
                observacion.trim() !== " " && 
                observacion.trim() !== "-" && 
                observacion.trim() !== "_" && 
                observacion.trim() !== "." &&
                observacion.trim() !== '"' &&
                observacion.trim() !== "'") {
              
              // Verificar si este desvío ya existe para evitar duplicados
              var claveDesvio = crearClaveEvidencia(fecha, area, observacion);
              
              if (!desviosExistentes[claveDesvio]) {
                // Es un desvío nuevo, agregarlo
                nuevosDesvios.push([
                  fecha,              // Fecha
                  area,               // Área
                  observacion,        // Desvío (observación)
                  evidencia || ""     // Link (evidencia)
                ]);
                
                totalProcesados++;
                
                // Marcar como procesado para evitar duplicados en esta misma ejecución
                desviosExistentes[claveDesvio] = true;
              } else {
                // Ya existe este desvío, omitirlo
                duplicadosEvitados++;
              }
            }
          }
        }
      } else {
        // Contar registros omitidos por el filtro temporal
        registrosOmitidos++;
      }
    }
    
    // Agregar los nuevos desvíos a la hoja
    if (nuevosDesvios.length > 0) {
      // Determinar la última fila con datos
      var ultimaFila = Math.max(1, hojaDesvios.getLastRow());
      var filaInsercion = ultimaFila === 1 ? 2 : ultimaFila + 1; // Saltar encabezados
      
      // Insertar datos
      var rangoInsercion = hojaDesvios.getRange(filaInsercion, 1, nuevosDesvios.length, nuevosDesvios[0].length);
      rangoInsercion.setValues(nuevosDesvios);
      
      // Centrar todo el contenido
      rangoInsercion.setHorizontalAlignment("center");
      rangoInsercion.setVerticalAlignment("middle");
      
      // Generar hipervínculos en la columna E
      var rango = hojaDesvios.getRange(filaInsercion, 5, nuevosDesvios.length, 1);
      var formulas = [];
      
      for (var i = 0; i < nuevosDesvios.length; i++) {
        var fila = filaInsercion + i;
        // Crear fórmula HIPERVINCULO solo si hay link
        if (nuevosDesvios[i][3]) {
          formulas.push(['=HYPERLINK(D' + fila + ',C' + fila + ')']);
        } else {
          formulas.push(['']);
        }
      }
      
      rango.setFormulas(formulas);
      
      // Formatear columna de fecha (sin ceros iniciales)
      hojaDesvios.getRange(filaInsercion, 1, nuevosDesvios.length, 1).setNumberFormat('d/m/yyyy');
    }
    
    // Registrar resultado incluyendo auditorías sin desvíos
    Logger.log("Procesamiento completado: " + totalProcesados + " desvíos procesados. " + 
              registrosOmitidos + " registros omitidos por fecha. " +
              duplicadosEvitados + " duplicados evitados. " +
              sinDesviosReales + " auditorías sin desvíos reales.");
    
    // Mostrar mensaje de éxito solo si se solicita explícitamente
    if (mostrarAlertas && totalProcesados > 0) {
      SpreadsheetApp.getUi().alert("Procesamiento completado: " + totalProcesados + " desvíos procesados.\n" +
                                  "Incluyendo " + sinDesviosReales + " auditorías sin desvíos reales.");
    }
    
    return totalProcesados;
    
  } catch (error) {
    Logger.log("Error en procesarEvidenciasDesvios: " + error.toString());
    
    // Solo mostrar alerta de error si se solicita
    if (mostrarAlertas) {
      SpreadsheetApp.getUi().alert("Error: " + error.message);
    }
    
    return 0;
  }
}

/**
 * Crea una clave única para una evidencia de desvío
 */
function crearClaveEvidencia(fecha, area, desvio) {
  // Normalizar fecha a string YYYY-MM-DD
  var fechaStr = "";
  if (fecha instanceof Date) {
    fechaStr = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd");
  } else {
    fechaStr = String(fecha);
  }
  
  // Normalizar área y desvío
  area = String(area).trim();
  desvio = String(desvio).trim();
  
  // Crear clave combinando los tres valores
  return fechaStr + "|" + area + "|" + desvio;
}

/**
 * Configura los encabezados de la hoja EvidenciasDesvios
 */
function configurarEncabezados(hoja) {
  // Definir encabezados
  var encabezados = ["Fecha", "Área", "Desvío", "Link", "Hipervínculo"];
  
  // Aplicar encabezados
  hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  
  // Formatear encabezados
  var rangoEncabezados = hoja.getRange(1, 1, 1, encabezados.length);
  rangoEncabezados.setBackground("#f3f3f3");
  rangoEncabezados.setFontWeight("bold");
  rangoEncabezados.setHorizontalAlignment("center");
  hoja.setFrozenRows(1);

  // Ajustar anchos de columna
  hoja.setColumnWidth(1, 100);  // Fecha
  hoja.setColumnWidth(2, 150);  // Área
  hoja.setColumnWidth(3, 350);  // Desvío
  hoja.setColumnWidth(4, 150);  // Link
  hoja.setColumnWidth(5, 350);  // Hipervínculo
}

/**
 * Busca un encabezado en un array y devuelve su índice
 */
function encontrarIndiceEncabezado(encabezados, nombre) {
  for (var i = 0; i < encabezados.length; i++) {
    if (encabezados[i] === nombre) {
      return i;
    }
  }
  return -1;
}

//=====================================================================
// PARTE 2: DISTRIBUCIÓN DE DESVÍOS A HOJAS POR SECTOR
//=====================================================================

/**
 * Distribuye los desvíos desde "EvidenciasDesvios" a hojas por sector
 */
function distribuirDesviosAreas(mostrarAlertas) {
  // Si no se especifica, no mostrar alertas por defecto
  mostrarAlertas = (mostrarAlertas === undefined) ? false : mostrarAlertas;
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Obtener hoja de desvíos
    var hojaDesvios = ss.getSheetByName("EvidenciasDesvios");
    if (!hojaDesvios || hojaDesvios.getLastRow() <= 1) {
      throw new Error("No hay datos en la hoja 'EvidenciasDesvios' para distribuir");
    }
    
    // Obtener datos de la hoja intermedia (omitiendo encabezados)
    var datosDesvios = hojaDesvios.getRange(2, 1, hojaDesvios.getLastRow() - 1, 5).getValues();
    
    // Comprobar si hay datos
    if (datosDesvios.length === 0) {
      throw new Error("No hay desvíos para distribuir");
    }
    
    // Agrupar desvíos por área y mes/año
    var desviosPorAreaYMes = agruparDesviosPorAreaYMes(datosDesvios);
    
    // Distribuir desvíos a cada área (versión corregida)
    var estadisticas = distribuirDesviosAHojas(ss, desviosPorAreaYMes);
    
    // Asegurar que todos los desvíos estén consolidados
    consolidarTodosDesvios(false); // No mostrar alertas en la consolidación automática
    
    return { exito: true, estadisticas: estadisticas };
    
  } catch (error) {
    Logger.log("Error en distribuirDesviosAreas: " + error.toString());
    
    if (mostrarAlertas) {
      SpreadsheetApp.getUi().alert("Error: " + error.message);
    }
    
    return { exito: false, error: error.message };
  }
}

/**
 * Agrupa los desvíos por área y mes/año
 */
function agruparDesviosPorAreaYMes(datosDesvios) {
  var desviosPorAreaYMes = {};
  
  datosDesvios.forEach(function(fila) {
    var fecha = fila[0];
    var area = fila[1];
    var desvio = fila[2];
    var link = fila[3];
    var hipervinculo = fila[4];
    
    // Solo procesar si tiene área y desvío
    if (area && desvio) {
      // Crear clave para agrupar por área
      if (!desviosPorAreaYMes[area]) {
        desviosPorAreaYMes[area] = {};
      }
      
      // Crear clave para agrupar por mes/año
      var mes = fecha.getMonth() + 1; // Mes en JavaScript es 0-11
      var anio = fecha.getFullYear();
      var claveYearMonth = anio + "-" + (mes < 10 ? "0" + mes : mes);
      
      if (!desviosPorAreaYMes[area][claveYearMonth]) {
        desviosPorAreaYMes[area][claveYearMonth] = [];
      }
      
      // Agregar desvío al grupo correspondiente
      desviosPorAreaYMes[area][claveYearMonth].push({
        fecha: fecha,
        desvio: desvio,
        link: link,
        hipervinculo: hipervinculo
      });
    }
  });
  
  return desviosPorAreaYMes;
}

/**
 * Distribuye los desvíos a las hojas correspondientes, filtrando por fecha de última auditoría
 */
function distribuirDesviosAHojas(ss, desviosPorAreaYMes) {
  var estadisticas = {
    areasActualizadas: 0,
    desviosDistribuidos: 0,
    desviosEliminados: 0,
    hojaCreadas: 0,
    desviosAntiguosLimpiados: 0  // Nuevo contador para desvíos limpiados
  };
  
  // NUEVO: Obtener fechas de últimas auditorías por área
  var fechasUltimasAuditorias = obtenerFechasUltimasAuditorias();
  
  // Para cada área con desvíos
  for (var area in desviosPorAreaYMes) {
    // Obtener o crear hoja para el área
    var hojaArea = ss.getSheetByName(area);
    if (!hojaArea) {
      hojaArea = ss.insertSheet(area);
      configurarHojaArea(hojaArea);
      estadisticas.hojaCreadas++;
    }
    
    // Primero, consolidar los desvíos existentes antes de cualquier modificación
    consolidarDesviosDesdeHojaArea(hojaArea);
    
    // NUEVO: Limpiar desvíos antiguos en "Inconformidades" cuando hay una nueva auditoría
    var fechaUltimaAuditoria = fechasUltimasAuditorias[area];
    if (fechaUltimaAuditoria) {
      var desviosLimpiados = limpiarDesviosAntiguos(area, fechaUltimaAuditoria);
      estadisticas.desviosAntiguosLimpiados += desviosLimpiados;
    }
    
    // MODIFICADO: Filtrar solo los desvíos de la última auditoría para esta área
    var desviosFiltrados = [];
    
    // Si tenemos fecha de última auditoría
    if (fechaUltimaAuditoria) {
      // Obtener todos los desvíos de todas las fechas para esta área
      var todosDesvios = [];
      for (var mesAnio in desviosPorAreaYMes[area]) {
        todosDesvios = todosDesvios.concat(desviosPorAreaYMes[area][mesAnio]);
      }
      
      // NUEVO: Filtrar desvíos por fecha (mismo día de la última auditoría)
      desviosFiltrados = todosDesvios.filter(function(desvio) {
        // Para cada desvío, verificamos si es del mismo día de la última auditoría
        return esMismaFechaAuditoria(desvio.fecha, fechaUltimaAuditoria);
      });
      
      Logger.log("Área: " + area + " - Encontrados " + desviosFiltrados.length + 
               " desvíos de la última auditoría (" + 
               Utilities.formatDate(fechaUltimaAuditoria, Session.getScriptTimeZone(), "yyyy-MM-dd") + ")");
    }
    
    // Obtener datos existentes
    var datosExistentes = [];
    if (hojaArea.getLastRow() > 1) {
      datosExistentes = hojaArea.getRange(2, 1, hojaArea.getLastRow() - 1, hojaArea.getLastColumn()).getValues();
      estadisticas.desviosEliminados += datosExistentes.length;
    }
    
    // Guardar datos manuales (columnas 4, 5, 6, 7) para cada punto inconforme
    var datosManuales = {};
    datosExistentes.forEach(function(fila) {
      var clave = formatearClave(fila[0], fila[1]); // Fecha + Punto inconforme
      if (clave) {
        datosManuales[clave] = {
          resuelta: fila[3],       // Resuelta
          noResuelta: fila[4],     // No resuelta
          fechaCierre: fila[5],    // Fecha de cierre
          comentario: fila[6]      // Comentario
        };
      }
    });
    
    // Limpiar la hoja (pero preservamos información manual)
    if (hojaArea.getLastRow() > 1) {
      hojaArea.getRange(2, 1, hojaArea.getLastRow() - 1, hojaArea.getLastColumn()).clear();
    }
    
    // Distribuir SOLO los desvíos de la última auditoría conservando datos manuales
    if (desviosFiltrados && desviosFiltrados.length > 0) {
      distribuirDesviosAHojaAreaConservandoDatos(hojaArea, desviosFiltrados, area, datosManuales);
      estadisticas.desviosDistribuidos += desviosFiltrados.length;
      estadisticas.areasActualizadas++;
    }
    // AGREGAR ESTE NUEVO BLOQUE DE CÓDIGO:
    else if (fechaUltimaAuditoria) {
      // Si hay fecha de última auditoría pero no hay desvíos filtrados,
      // significa que la última auditoría no tiene desvíos, por lo que la hoja ya está limpia
      Logger.log("Área: " + area + " - Última auditoría (" + 
                Utilities.formatDate(fechaUltimaAuditoria, Session.getScriptTimeZone(), "yyyy-MM-dd") + 
                ") sin desvíos. Hoja limpiada.");
      estadisticas.areasActualizadas++;
    }
  }
  
  return estadisticas;
}

/**
 * Encuentra el mes/año más reciente en un conjunto de desvíos
 */
function encontrarMesAnioMasReciente(desviosPorMes) {
  var mesesAnios = Object.keys(desviosPorMes);
  if (mesesAnios.length === 0) return null;
  
  // Ordenar por mes/año de forma descendente
  mesesAnios.sort(function(a, b) {
    return b.localeCompare(a); // Comparación lexicográfica inversa
  });
  
  return mesesAnios[0]; // Retornar el más reciente
}

/**
 * Identifica las fechas de la auditoría más reciente para cada área
 * basándose en las respuestas del formulario
 */
function obtenerFechasUltimasAuditorias() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaRespuestas = ss.getSheetByName("Respuestas de formulario 1");
  
  if (!hojaRespuestas) {
    return {};
  }
  
  // Obtener datos de todas las respuestas
  var datosRespuestas = hojaRespuestas.getDataRange().getValues();
  var encabezados = datosRespuestas[0];
  
  // Identificar índices de columnas necesarias
  var idxMarcaTemporal = encontrarIndiceEncabezado(encabezados, "Marca temporal");
  var idxArea = encontrarIndiceEncabezado(encabezados, "Área auditada");
  
  // Si no encontramos las columnas necesarias, retornar objeto vacío
  if (idxMarcaTemporal < 0 || idxArea < 0) {
    return {};
  }
  
  // Mapa para guardar la fecha más reciente de auditoría por área
  var fechasUltimasAuditorias = {};
  
  // Procesar cada respuesta (omitiendo la fila de encabezados)
  for (var i = 1; i < datosRespuestas.length; i++) {
    var fila = datosRespuestas[i];
    var marcaTemporal = fila[idxMarcaTemporal]; // Fecha de la auditoría
    var area = fila[idxArea]; // Área auditada
    
    // Solo procesar si tenemos tanto fecha como área
    if (marcaTemporal && area) {
      // Convertir string a Date si es necesario
      if (typeof marcaTemporal === 'string') {
        marcaTemporal = new Date(marcaTemporal);
      }
      
      // Actualizar la fecha más reciente para esta área
      if (!fechasUltimasAuditorias[area] || marcaTemporal > fechasUltimasAuditorias[area]) {
        fechasUltimasAuditorias[area] = marcaTemporal;
      }
    }
  }
  
  return fechasUltimasAuditorias;
}

/**
 * Helper: Determina si dos fechas pertenecen a la misma auditoría
 * (mismo día, ignorando hora exacta)
 */
function esMismaFechaAuditoria(fecha1, fecha2) {
  // Si alguna fecha es null o undefined
  if (!fecha1 || !fecha2) return false;
  
  // Asegurar que ambas son objetos Date
  var d1 = fecha1 instanceof Date ? fecha1 : new Date(fecha1);
  var d2 = fecha2 instanceof Date ? fecha2 : new Date(fecha2);
  
  // Comparar año, mes y día (ignorar hora)
  return d1.getFullYear() === d2.getFullYear() && 
         d1.getMonth() === d2.getMonth() && 
         d1.getDate() === d2.getDate();
}

/**
 * Configura una nueva hoja de área con los encabezados correctos
 */
function configurarHojaArea(hoja) {
  // Definir encabezados según especificación exacta
  var encabezados = ["Fecha", "Punto inconforme", "Área", "Resuelta", "No resuelta", "Fecha de cierre", "Comentario"];
  
  // Aplicar encabezados
  hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  
  // Formatear encabezados
  var rangoEncabezados = hoja.getRange(1, 1, 1, encabezados.length);
  rangoEncabezados.setBackground("#f3f3f3");
  rangoEncabezados.setFontWeight("bold");
  rangoEncabezados.setHorizontalAlignment("center");
  hoja.setFrozenRows(1);
  
  // Ajustar anchos de columna
  hoja.setColumnWidth(1, 100);   // Fecha
  hoja.setColumnWidth(2, 350);   // Punto inconforme
  hoja.setColumnWidth(3, 150);   // Área
  hoja.setColumnWidth(4, 100);   // Resuelta
  hoja.setColumnWidth(5, 100);   // No resuelta
  hoja.setColumnWidth(6, 100);   // Fecha de cierre
  hoja.setColumnWidth(7, 250);   // Comentario
  
  // Configurar validación para checkbox en "Resuelta"
  var rangoResuelta = hoja.getRange("D2:D1000");
  var validacionCheckbox = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();
  rangoResuelta.setDataValidation(validacionCheckbox);
}

/**
 * Función auxiliar para formatear una clave única para cada desvío
 */
function formatearClave(fecha, texto) {
  if (!fecha || !texto) return null;
  
  // Si fecha es un objeto Date, convertirlo a string en formato YYYY-MM-DD
  var fechaStr = "";
  if (fecha instanceof Date) {
    fechaStr = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd");
  } else {
    fechaStr = fecha.toString();
  }
  
  // Crear una versión simplificada del texto (primeros 50 caracteres)
  var textoSimplificado = texto.toString().substring(0, 50).trim();
  
  return fechaStr + "|" + textoSimplificado;
}

function distribuirDesviosAHojaAreaConservandoDatos(hojaArea, desvios, nombreArea, datosManuales) {
  // Preparar datos para insertar según estructura exacta
  var datosParaInsertar = [];
  var formulasHipervinculo = [];
  
  // Preparar los datos y las fórmulas por separado
  for (var i = 0; i < desvios.length; i++) {
    var desvio = desvios[i];
    
    // Verificar si ya existe información manual para este desvío
    var clave = formatearClave(desvio.fecha, desvio.desvio);
    var datosConservados = datosManuales[clave] || null;
    
    // Datos regulares para todas las columnas
    datosParaInsertar.push([
      desvio.fecha,                                   // Fecha
      "",                                             // Punto inconforme (para la fórmula)
      nombreArea,                                     // Área
      datosConservados ? datosConservados.resuelta : false, // Resuelta (conservado si existe)
      "",                                             // No resuelta (para la fórmula)
      datosConservados ? datosConservados.fechaCierre : "", // Fecha cierre (conservado)
      datosConservados ? datosConservados.comentario : ""   // Comentario (conservado)
    ]);
    
    // Preparar la fórmula de hipervínculo
    var formulaHipervinculo = "";
    if (desvio.hipervinculo && typeof desvio.hipervinculo === 'string' && 
        desvio.hipervinculo.toUpperCase().startsWith('=HYPERLINK')) {
      formulaHipervinculo = desvio.hipervinculo;
    } else if (desvio.link && desvio.link.trim() !== "") {
      var url = desvio.link.replace(/"/g, '""');
      var texto = desvio.desvio.replace(/"/g, '""');
      formulaHipervinculo = '=HYPERLINK("' + url + '","' + texto + '")';
    } else {
      formulaHipervinculo = desvio.desvio || "";
    }
    
    formulasHipervinculo.push([formulaHipervinculo]);
  }
  
  // Si no hay datos para insertar, terminar
  if (datosParaInsertar.length === 0) return;
  
  // Insertar datos en la hoja
  var filaInicio = Math.max(2, hojaArea.getLastRow() + 1);
  var rangoInsercion = hojaArea.getRange(filaInicio, 1, datosParaInsertar.length, datosParaInsertar[0].length);
  rangoInsercion.setValues(datosParaInsertar);
  
  // Centrar verticalmente todas las celdas
  rangoInsercion.setVerticalAlignment("middle");
  
  // Centrar horizontalmente columnas A-F
  hojaArea.getRange(filaInicio, 1, datosParaInsertar.length, 6).setHorizontalAlignment("center");
  
  // Alinear a la izquierda la columna Comentario (G)
  hojaArea.getRange(filaInicio, 7, datosParaInsertar.length, 1).setHorizontalAlignment("left");
  
  // Formatear fecha (sin ceros iniciales)
  hojaArea.getRange(filaInicio, 1, datosParaInsertar.length, 1).setNumberFormat("d/m/yyyy");
  hojaArea.getRange(filaInicio, 6, datosParaInsertar.length, 1).setNumberFormat("d/m/yyyy");
  
  // Aplicar fórmulas de hipervínculo en la columna "Punto inconforme"
  hojaArea.getRange(filaInicio, 2, formulasHipervinculo.length, 1)
    .setFormulas(formulasHipervinculo);
  
  // Configurar fórmula para "No resuelta" basada en "Resuelta" PARA TODOS LOS REGISTROS
  for (var i = 0; i < datosParaInsertar.length; i++) {
    var fila = filaInicio + i;
    hojaArea.getRange(fila, 5).setFormula('=IF(D' + fila + '=FALSE,"Si","")');
  }
}

/**
 * Consolida los desvíos de una hoja antes de limpiarla
 */
function consolidarDesviosDesdeHojaArea(hojaArea) {
  try {
    // Si la hoja está vacía, no hay nada que consolidar
    if (hojaArea.getLastRow() <= 1) return 0;
    
    var nombreArea = hojaArea.getName();
    var datos = hojaArea.getRange(2, 1, hojaArea.getLastRow() - 1, hojaArea.getLastColumn()).getValues();
    
    // Guardar en Inconformidades (no mostrar alertas)
    return consolidarDesviosEnInconformidades(datos, nombreArea);
  } catch (error) {
    Logger.log("Error consolidando hoja " + hojaArea.getName() + ": " + error.toString());
    return 0;
  }
}


//=====================================================================
// PARTE 3: CONSOLIDACIÓN EN HOJA "INCONFORMIDADES"
//=====================================================================
/**
 * Limpia los "Si" de la columna "No resuelta" para desvíos antiguos de un área
 * cuando hay una nueva auditoría
 */
function limpiarDesviosAntiguos(area, fechaUltimaAuditoria) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaInconformidades = ss.getSheetByName("Inconformidades");
    
    if (!hojaInconformidades || hojaInconformidades.getLastRow() <= 1) {
      return 0; // No hay datos para procesar
    }
    
    // Obtener todos los datos de Inconformidades
    var datosInconformidades = hojaInconformidades.getDataRange().getValues();
    var encabezados = datosInconformidades[0];
    
    // Identificar índices de columnas necesarias
    var idxFecha = 0;        // Columna A - Fecha
    var idxArea = 2;         // Columna C - Área
    var idxNoResuelta = 4;   // Columna E - No resuelta
    
    // Contador de desvíos limpiados
    var desviosLimpiados = 0;
    
    // Buscar desvíos antiguos del área específica
    for (var i = 1; i < datosInconformidades.length; i++) {
      var fila = datosInconformidades[i];
      var fechaDesvio = fila[idxFecha];
      var areaDesvio = fila[idxArea];
      var valorNoResuelta = fila[idxNoResuelta];
      
      // Verificar si es un desvío del área buscada
      if (areaDesvio === area) {
        // Verificar si es un desvío antiguo (anterior a la última auditoría)
        if (fechaDesvio instanceof Date && fechaUltimaAuditoria instanceof Date &&
            !esMismaFechaAuditoria(fechaDesvio, fechaUltimaAuditoria) && 
            fechaDesvio < fechaUltimaAuditoria) {
          
          // Si tiene un valor "Si" en la columna "No resuelta", limpiarlo
          if (valorNoResuelta === "Si") {
            // Establecer valor vacío en la columna E
            hojaInconformidades.getRange(i + 1, idxNoResuelta + 1).setValue("");
            desviosLimpiados++;
          }
        }
      }
    }
    
    return desviosLimpiados;
    
  } catch (error) {
    Logger.log("Error al limpiar desvíos antiguos: " + error.toString());
    return 0;
  }
}


/**
 * Consolida todos los desvíos de las hojas de área en "Inconformidades"
 */
function consolidarTodosDesvios(mostrarAlertas) {
  // Si no se especifica, no mostrar alertas por defecto
  mostrarAlertas = (mostrarAlertas === undefined) ? false : mostrarAlertas;
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojas = ss.getSheets();
    var hojaInconformidades = ss.getSheetByName("Inconformidades");
    
    // Crear hoja Inconformidades si no existe
    if (!hojaInconformidades) {
      hojaInconformidades = ss.insertSheet("Inconformidades");
      configurarHojaInconformidades(hojaInconformidades);
    }
    
    var estadisticas = {
      hojasExaminadas: 0,
      desviosConsolidados: 0
    };
    
    // Recorrer todas las hojas buscando las que contienen desvíos
    hojas.forEach(function(hoja) {
      var nombreHoja = hoja.getName();
      
      // Excluir hojas del sistema y la propia hoja de inconformidades
      if (nombreHoja !== "Inconformidades" && 
          nombreHoja !== "EvidenciasDesvios" &&
          nombreHoja !== "Preguntas maestras" &&
          nombreHoja !== "Respuestas de formulario 1" &&
          nombreHoja !== "Consolidada" &&
          !nombreHoja.startsWith("Hoja")) {
        
        // Verificar si la hoja tiene la estructura de desvíos
        // Buscamos encabezados como "Punto inconforme" y "Resuelta"
        var primerFila = [];
        if (hoja.getLastRow() > 0) {
          primerFila = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
        }
        
        if (primerFila.includes("Punto inconforme") || primerFila.includes("Resuelta")) {
          estadisticas.hojasExaminadas++;
          
          // Obtener datos si hay contenido
          if (hoja.getLastRow() > 1) {
            var datosDesvios = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues();
            
            // Consolidar estos desvíos
            var desviosConsolidados = consolidarDesviosEnInconformidades(datosDesvios, nombreHoja);
            estadisticas.desviosConsolidados += desviosConsolidados;
          }
        }
      }
    });
    
    return estadisticas;
    
  } catch (error) {
    Logger.log("Error en consolidarTodosDesvios: " + error.toString());
    
    if (mostrarAlertas) {
      SpreadsheetApp.getUi().alert("Error al consolidar: " + error.message);
    }
    
    return null;
  }
}

/**
 * Configura la hoja de Inconformidades
 */
function configurarHojaInconformidades(hoja) {
  // Definir encabezados con la misma estructura que las hojas de área
  var encabezados = ["Fecha", "Punto inconforme", "Área", "Resuelta", "No resuelta", "Fecha de cierre", "Comentario"];
  
  // Aplicar encabezados
  hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  
  // Formatear encabezados
  var rangoEncabezados = hoja.getRange(1, 1, 1, encabezados.length);
  rangoEncabezados.setBackground("#f3f3f3");
  rangoEncabezados.setFontWeight("bold");
  rangoEncabezados.setHorizontalAlignment("center");
  hoja.setFrozenRows(1);

  // OPCIONAL: Establecer la alineación de la columna Comentario para futuras entradas
  hoja.getRange("G2:G1000").setHorizontalAlignment("left");

  // Ajustar anchos de columna
  hoja.setColumnWidth(1, 100);   // Fecha
  hoja.setColumnWidth(2, 350);   // Punto inconforme
  hoja.setColumnWidth(3, 150);   // Área
  hoja.setColumnWidth(4, 100);   // Resuelta
  hoja.setColumnWidth(5, 100);   // No resuelta
  hoja.setColumnWidth(6, 100);   // Fecha de cierre
  hoja.setColumnWidth(7, 250);   // Comentario
}

/**
 * Consolida desvíos en la hoja "Inconformidades", evitando duplicados
 * Con centrado de texto y formato de fecha sin ceros iniciales
 * Actualizada para mantener fórmulas dinámicas en la columna E
 */
function consolidarDesviosEnInconformidades(datos, nombreArea) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaInconformidades = ss.getSheetByName("Inconformidades");
    
    // Si no existe la hoja Inconformidades, crearla
    if (!hojaInconformidades) {
      hojaInconformidades = ss.insertSheet("Inconformidades");
      configurarHojaInconformidades(hojaInconformidades);
    }
    
    // Obtener datos existentes de Inconformidades
    var datosExistentes = [];
    var mapaRegistrosExistentes = {};
    
    if (hojaInconformidades.getLastRow() > 1) {
      datosExistentes = hojaInconformidades.getRange(2, 1, hojaInconformidades.getLastRow() - 1, 
                                                    hojaInconformidades.getLastColumn()).getValues();
      
      // Crear un mapa de registros existentes para búsqueda rápida
      datosExistentes.forEach(function(fila, indice) {
        // Generar clave única: Fecha + Área + Punto inconforme
        var fecha = fila[0];
        var area = fila[2];
        var puntoInconforme = fila[1];
        
        if (fecha && area && puntoInconforme) {
          var clave = generarClaveUnica(fecha, area, puntoInconforme);
          mapaRegistrosExistentes[clave] = {
            fila: indice + 2, // +2 porque empezamos en la fila 2 y el índice es 0-based
            datos: fila
          };
        }
      });
    }
    
    // Procesar datos para consolidar
    var nuevosRegistros = []; // Para registros que no existen
    var actualizaciones = []; // Para registros que necesitan actualización
    
    datos.forEach(function(fila) {
      // Solo consolidar desvíos con datos
      if (fila[0] && fila[1]) { // Si hay fecha y punto inconforme
        var fecha = fila[0];
        var puntoInconforme = fila[1];
        var area = fila[2] || nombreArea;
        var resuelta = fila[3] || false;
        var noResuelta = ""; // MODIFICADO: dejamos vacío para aplicar fórmula después
        var fechaCierre = fila[5] || "";
        var comentario = fila[6] || "";
        
        // Generar clave única para buscar duplicados
        var clave = generarClaveUnica(fecha, area, puntoInconforme);
        
        // Verificar si ya existe este registro
        if (mapaRegistrosExistentes[clave]) {
          // El registro existe, verificar si hay cambios
          var registroExistente = mapaRegistrosExistentes[clave].datos;
          var filaExistente = mapaRegistrosExistentes[clave].fila;
          
          // Comparar valores importantes (columnas 4, 6, 7)
          var cambios = false;
          
          if (registroExistente[3] != resuelta || 
              // Ya no comparamos la columna noResuelta (índice 4) porque será calculada
              !sonFechasIguales(registroExistente[5], fechaCierre) ||
              registroExistente[6] != comentario) {
            cambios = true;
          }
          
          // Si hay cambios, actualizar el registro existente
          if (cambios) {
            actualizaciones.push({
              fila: filaExistente,
              datos: [fecha, puntoInconforme, area, resuelta, noResuelta, fechaCierre, comentario]
            });
          }
          // Si no hay cambios, no hacemos nada (evitamos duplicados)
          
        } else {
          // Es un registro nuevo, agregarlo a la lista de nuevos
          nuevosRegistros.push([
            fecha, puntoInconforme, area, resuelta, noResuelta, fechaCierre, comentario
          ]);
        }
      }
    });
    
    // Actualizar registros existentes que tienen cambios
    actualizaciones.forEach(function(actualizacion) {
      var rangoActualizacion = hojaInconformidades.getRange(actualizacion.fila, 1, 1, actualizacion.datos.length);
      rangoActualizacion.setValues([actualizacion.datos]);
      
      // NUEVO: Establecer fórmula condicional para la columna E (No resuelta)
      hojaInconformidades.getRange(actualizacion.fila, 5).setFormula('=IF(D' + actualizacion.fila + '=FALSE,"Si","")');
      
      // Centrar columnas A-F
      hojaInconformidades.getRange(actualizacion.fila, 1, 1, 6).setHorizontalAlignment("center");
      // Alinear a la izquierda la columna Comentario (G)
      hojaInconformidades.getRange(actualizacion.fila, 7, 1, 1).setHorizontalAlignment("left");
      // Centrar verticalmente todas las columnas
      rangoActualizacion.setVerticalAlignment("middle");
    });
    
    // Insertar nuevos registros al final
    if (nuevosRegistros.length > 0) {
      var ultimaFila = Math.max(2, hojaInconformidades.getLastRow() + 1);
      
      var rangoNuevos = hojaInconformidades.getRange(ultimaFila, 1, nuevosRegistros.length, nuevosRegistros[0].length);
      rangoNuevos.setValues(nuevosRegistros);
      
      // NUEVO: Establecer fórmulas condicionales para todos los registros nuevos
      for (var i = 0; i < nuevosRegistros.length; i++) {
        var fila = ultimaFila + i;
        hojaInconformidades.getRange(fila, 5).setFormula('=IF(D' + fila + '=FALSE,"Si","")');
      }
      
      // Aplicar alineaciones diferenciadas
      // Centrar verticalmente todo
      rangoNuevos.setVerticalAlignment("middle");
      // Centrar horizontalmente columnas A-F
      hojaInconformidades.getRange(ultimaFila, 1, nuevosRegistros.length, 6).setHorizontalAlignment("center");
      // Alinear a la izquierda la columna Comentario (G)
      hojaInconformidades.getRange(ultimaFila, 7, nuevosRegistros.length, 1).setHorizontalAlignment("left");
      
      // Formatear fechas (sin ceros iniciales)
      hojaInconformidades.getRange(ultimaFila, 1, nuevosRegistros.length, 1).setNumberFormat("d/m/yyyy");
      hojaInconformidades.getRange(ultimaFila, 6, nuevosRegistros.length, 1).setNumberFormat("d/m/yyyy");
      
      // Configurar checkbox en resuelta
      var rangoResuelta = hojaInconformidades.getRange(ultimaFila, 4, nuevosRegistros.length, 1);
      var validacionCheckbox = SpreadsheetApp.newDataValidation()
        .requireCheckbox()
        .build();
      rangoResuelta.setDataValidation(validacionCheckbox);
    }
    
    Logger.log("Consolidación: " + nuevosRegistros.length + " nuevos registros y " + 
              actualizaciones.length + " actualizaciones del área " + nombreArea);
    
    return nuevosRegistros.length + actualizaciones.length;
    
  } catch (error) {
    Logger.log("Error consolidando desvíos: " + error.toString());
    return 0;
  }
}

/**
 * Genera una clave única para identificar un registro
 */
function generarClaveUnica(fecha, area, puntoInconforme) {
  var fechaStr = "";
  if (fecha instanceof Date) {
    fechaStr = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd");
  } else {
    fechaStr = fecha.toString().trim();
  }
  
  area = (area || "").toString().trim();
  puntoInconforme = (puntoInconforme || "").toString().trim().substring(0, 100);
  
  return fechaStr + "|" + area + "|" + puntoInconforme;
}

/**
 * Compara si dos fechas son iguales
 */
function sonFechasIguales(fecha1, fecha2) {
  // Si ambas están vacías, son iguales
  if ((!fecha1 || fecha1 === "") && (!fecha2 || fecha2 === "")) {
    return true;
  }
  
  // Si solo una está vacía, son diferentes
  if ((!fecha1 || fecha1 === "") || (!fecha2 || fecha2 === "")) {
    return false;
  }
  
  // Convertir a objetos Date si no lo son
  var d1 = fecha1 instanceof Date ? fecha1 : new Date(fecha1);
  var d2 = fecha2 instanceof Date ? fecha2 : new Date(fecha2);
  
  // Comparar año, mes y día
  return d1.getFullYear() === d2.getFullYear() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getDate() === d2.getDate();
}

//=====================================================================
// PARTE 4: CONFIGURACIÓN DE ACTIVADORES Y MENÚ
//=====================================================================

/**
 * Configura un activador para procesar los desvíos automáticamente
 */
function configurarActivadorAutomatico(mostrarAlertas) {
  // Si no se especifica, asumir que queremos mostrar alertas si se ejecuta manualmente
  mostrarAlertas = (mostrarAlertas === undefined) ? true : mostrarAlertas;
  
  // Eliminar activadores existentes para evitar duplicados
  var activadores = ScriptApp.getProjectTriggers();
  for (var i = 0; i < activadores.length; i++) {
    if (activadores[i].getHandlerFunction() === 'procesarEvidenciasDesvios') {
      ScriptApp.deleteTrigger(activadores[i]);
    }
  }
  
  // Crear nuevo activador (cada hora)
  ScriptApp.newTrigger('procesarEvidenciasDesvios')
      .timeBased()
      .everyHours(1)
      .create();
      
  if (mostrarAlertas) {
    SpreadsheetApp.getUi().alert("Activador configurado. Los desvíos se procesarán automáticamente cada hora.");
  }
}

/**
 * Configura un activador para consolidación automática periódica
 */
function configurarActivadorConsolidacionAutomatica(mostrarAlertas) {
  // Si no se especifica, asumir que queremos mostrar alertas si se ejecuta manualmente
  mostrarAlertas = (mostrarAlertas === undefined) ? true : mostrarAlertas;
  
  // Eliminar activadores existentes para evitar duplicados
  var activadores = ScriptApp.getProjectTriggers();
  for (var i = 0; i < activadores.length; i++) {
    if (activadores[i].getHandlerFunction() === 'consolidarTodosDesvios') {
      ScriptApp.deleteTrigger(activadores[i]);
    }
  }
  
  // Crear nuevo activador (diario)
  ScriptApp.newTrigger('consolidarTodosDesvios')
      .timeBased()
      .everyDays(1)
      .atHour(2) // 2 AM
      .create();
      
  if (mostrarAlertas) {
    SpreadsheetApp.getUi().alert("Activador configurado. Los desvíos se consolidarán automáticamente una vez al día a las 2 AM.");
  }
}

/**
 * Agrega una opción de menú para acceder a todas las funcionalidades
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    ui.createMenu('Gestión de Desvíos')
      .addItem('Procesar Evidencias de Desvíos', 'procesarEvidenciasDesviosConAlertas')
      .addSeparator()
      .addItem('Distribuir Desvíos a Áreas', 'distribuirDesviosAreasConAlertas')
      .addSeparator()
      .addItem('Consolidar Cambios en Inconformidades', 'consolidarTodosDesviosConAlertas')
      .addSeparator()
      .addItem('Configurar Procesamiento Automático', 'configurarActivadorAutomatico')
      .addItem('Configurar Consolidación Automática', 'configurarActivadorConsolidacionAutomatica')
      .addToUi();
  } catch (e) {
    Logger.log("Error al crear menú: " + e.toString());
  }
}

/**
 * Funciones wrapper para ejecución manual con alertas
 * Estas permiten mantener la interfaz de usuario informativa mientras
 * las ejecuciones automáticas funcionan silenciosamente.
 */
function procesarEvidenciasDesviosConAlertas() {
  return procesarEvidenciasDesvios(true);
}

function distribuirDesviosAreasConAlertas() {
  return distribuirDesviosAreas(true);
}

function consolidarTodosDesviosConAlertas() {
  return consolidarTodosDesvios(true);
}

/**
 * Procesa y distribuye desvíos en una sola operación
 * Útil para simplificar flujos de trabajo manuales
 */
function procesarYDistribuirDesvios() {
  // Procesar evidencias sin mostrar alerta
  var desviosProcesados = procesarEvidenciasDesvios(false);
  
  // Distribuir a hojas por área solo si se procesaron desvíos
  if (desviosProcesados > 0) {
    distribuirDesviosAreas(false);
  }
  
}

/**
 * Crea un activador diario completo que procesa, distribuye y consolida
 */
function configurarActivadorCompletoAutomatico(mostrarAlertas) {
  // Si no se especifica, asumir que queremos mostrar alertas si se ejecuta manualmente
  mostrarAlertas = (mostrarAlertas === undefined) ? true : mostrarAlertas;
  
  // Eliminar activadores existentes para evitar duplicados
  var activadores = ScriptApp.getProjectTriggers();
  for (var i = 0; i < activadores.length; i++) {
    if (activadores[i].getHandlerFunction() === 'procesarYDistribuirDesvios') {
      ScriptApp.deleteTrigger(activadores[i]);
    }
  }
  
  // Crear nuevo activador (diario a las 6 AM)
  ScriptApp.newTrigger('procesarYDistribuirDesvios')
      .timeBased()
      .everyDays(1)
      .atHour(6) // 6 AM
      .create();
      
  if (mostrarAlertas) {
    SpreadsheetApp.getUi().alert("Activador completo configurado. El sistema procesará, distribuirá y consolidará desvíos automáticamente una vez al día a las 6 AM.");
  }
}
