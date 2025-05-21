function createSheets() {
  var sheetNames = [
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

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheetByName("Plantilla");

  if (!templateSheet) {
    Logger.log("No se encontró la hoja Plantilla");
    return;
  }

  for (var i = 0; i < sheetNames.length; i++) {
    var sheetName = sheetNames[i];
    var newSheet = ss.getSheetByName(sheetName);

    // Elimina la hoja si ya existe
    if (newSheet) {
      ss.deleteSheet(newSheet);
    }

    // Duplica la hoja Plantilla
    newSheet = templateSheet.copyTo(ss).setName(sheetName);

    // Agrega la fórmula de filtro en la celda A2 (omitimos la columna D)
    newSheet.getRange("A2").setFormula(`=IFERROR(FILTER(Inconformidades!A2:C, Inconformidades!C2:C = "${sheetName}"), "")`);
    
    // Agrega la fórmula de filtro en la celda E2 para los comentarios
    newSheet.getRange("E2").setFormula(`=IFERROR(FILTER(Inconformidades!E2:E, Inconformidades!C2:C = "${sheetName}"), "")`);
  }
}





function deleteSheets() {
  var sheetNames = [
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

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  for (var i = 0; i < sheetNames.length; i++) {
    var sheetName = sheetNames[i];
    var sheet = ss.getSheetByName(sheetName);

    if (sheet) {
      ss.deleteSheet(sheet);
    } else {
      Logger.log("No se encontró la hoja: " + sheetName);
    }
  }
}
