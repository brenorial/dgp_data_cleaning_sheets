function ordenarAbasDO() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abas = ss.getSheets().filter(sheet => sheet.getName().startsWith("DO"));
  
  abas.forEach(sheet => {
    var data = sheet.getDataRange().getValues();
    if (data.length < 1) return; // Verifica se há dados suficientes
    
    var header = data[2]; // Cabeçalho na linha 3
    var pontuacaoIndex = header.indexOf("PONTUAÇÃO");
    var dataNascimentoIndex = header.indexOf("DATA DE NASCIMENTO");
    
    if (pontuacaoIndex === -1 || dataNascimentoIndex === -1) {
      Logger.log(`Colunas necessárias não encontradas em ${sheet.getName()}`);
      return;
    }
    
    var dados = data.slice(3); // Exclui cabeçalhos
    
    dados.sort((a, b) => {
      var pontuacaoA = parseFloat(a[pontuacaoIndex]) || 0;
      var pontuacaoB = parseFloat(b[pontuacaoIndex]) || 0;
      
      if (pontuacaoB !== pontuacaoA) {
        return pontuacaoB - pontuacaoA; // Ordena por PONTUAÇÃO (decrescente)
      }
      
      var dataA = new Date(a[dataNascimentoIndex]);
      var dataB = new Date(b[dataNascimentoIndex]);
      
      return dataA - dataB; // Ordena por DATA DE NASCIMENTO (crescente)
    });
    
    if (dados.length > 0) {
      sheet.getRange(4, 1, dados.length, dados[0].length).setValues(dados);
    }
  });
}
