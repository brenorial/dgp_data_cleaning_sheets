function processarPlanilha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var baseSheet = ss.getSheetByName('BASE');
  var data = baseSheet.getDataRange().getValues();
  
  var novaLinha;

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().startsWith("Vaga:")) {
      var nomeVaga = data[i][0].toString().split(":")[1].trim();

      var novaAba = ss.getSheetByName(nomeVaga);
      if (!novaAba) {
        novaAba = ss.insertSheet(nomeVaga);
      } else {
        novaAba.clear();
      }

      novaLinha = 1;

      for (var j = i + 1; j < data.length; j++) {
        if (!data[j][0]) {
          continue;
        }

        if (data[j][0].toString().startsWith("Vaga:")) {
          break;
        }

        novaAba.getRange(novaLinha, 1, 1, data[j].length).setValues([data[j]]);
        novaLinha++;
      }
    }
  }
}
