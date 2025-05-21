function fillLookerAuxTable() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Presupuesto'); // Nombre de la hoja con los datos originales
  var lookerSheet = ss.getSheetByName('PresupuestoAux'); // Nombre de la hoja auxiliar para Looker

  // Obtener los datos de la primera tabla
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  // Limpiar la tabla auxiliar para Looker, si ya tiene datos
  if (lookerSheet.getLastRow() > 1) {
    lookerSheet.getRange(2, 1, lookerSheet.getLastRow() - 1, 4).clearContent();
    lookerSheet.getRange(2, 6, lookerSheet.getLastRow() - 1, 4).clearContent();
  }

  var areas = data.slice(3); // Asume que las áreas están desde la fila 4 en adelante
  var months = [3, 6, 9, 12];
  var rowIndex = 2; // Comienza en la fila 2 de la tabla auxiliar para Looker
  var plantaRowIndex = 2; // Comienza en la fila 2 de la tabla de Presupuesto Planta

  // Recorrer los años y meses para completar las tablas auxiliares
  for (var i = 1; i < data[0].length; i += 4) {
    var year = data[0][i]; // Obtener el año de la celda combinada
    for (var j = 0; j < months.length; j++) {
      var month = months[j];
      for (var k = 0; k < areas.length; k++) {
        var area = areas[k][0]; // Obtener el nombre del área
        var target = areas[k][i + j]; // Obtener el presupuesto correspondiente

        if (area === "PRESUPUESTO PLANTA") {
          // Agregar a la tabla de Presupuesto Planta
          if (target !== "" && target !== null) {
            lookerSheet.getRange(plantaRowIndex, 6).setValue(area);
            lookerSheet.getRange(plantaRowIndex, 7).setValue(target);
            lookerSheet.getRange(plantaRowIndex, 8).setValue(month);
            lookerSheet.getRange(plantaRowIndex, 9).setValue(new Date(year, month - 1, 1));
            plantaRowIndex++;
          }
        } else {
          // Agregar a la tabla auxiliar principal
          if (target !== "" && target !== null) {
            lookerSheet.getRange(rowIndex, 1).setValue(area);
            lookerSheet.getRange(rowIndex, 2).setValue(target);
            lookerSheet.getRange(rowIndex, 3).setValue(month);
            lookerSheet.getRange(rowIndex, 4).setValue(new Date(year, month - 1, 1));
            rowIndex++;
          }
        }
      }
    }
  }
}
