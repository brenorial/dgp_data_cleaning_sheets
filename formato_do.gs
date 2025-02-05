function copiarDadosPorArea() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var origem = ss.getSheetByName("teste");
  if (!origem) {
    Logger.log("Aba 'teste' não encontrada.");
    return;
  }
  
  var colunas = ["CARGO", "NOME COMPLETO:", "PPI", "PCD", "PONTUAÇÃO", "DATA DE NASCIMENTO"];
  var cabecalho = origem.getDataRange().getValues()[0];
  var indices = [
    cabecalho.indexOf("CARGO"),
    cabecalho.indexOf("NOME  COMPLETO:"),
    cabecalho.indexOf("[COTA] DESEJA CONCORRER ÀS VAGAS DESTINADAS A CANDIDATOS NEGROS E INDÍGENAS, CONFORME O ITEM 4 DA RESERVA DAS VAGAS PARA NEGROS E INDÍGENAS DO EDITAL?"),
    cabecalho.indexOf("[COTA] DESEJA CONCORRER ÀS VAGAS DESTINADAS ÀS PESSOAS COM DEFICIÊNCIA (PCD), CONFORME O ITEM 5 DO EDITAL?"),
    cabecalho.indexOf("TOTAL_PONTUAÇÃO"),
    cabecalho.indexOf("DATA DE NASCIMENTO:")
  ].filter(i => i !== -1);
  
  if (indices.length !== colunas.length) {
    Logger.log("Algumas colunas não foram encontradas.");
    return;
  }
  
  var dados = origem.getDataRange().getValues();
  var areas = {};
  
  for (var i = 1; i < dados.length; i++) {
    var cargo = dados[i][indices[0]];
    var areaIndex = cabecalho.indexOf("ÁREA");
    var area = areaIndex !== -1 ? dados[i][areaIndex] : "";
    var nomeCompleto = dados[i][indices[1]].replace(/[^A-Z\s]/gi, "").trim().toUpperCase();
    var cotaNegros = dados[i][indices[2]] === "SIM" ? "*" : "";
    var cotaPCD = dados[i][indices[3]] === "SIM" ? "**" : "";
    var totalPontuacao = dados[i][indices[4]];
    var dataNascimento = dados[i][indices[5]];
    
    if (!areas[cargo]) areas[cargo] = [];
    areas[cargo].push([cargo, nomeCompleto, cotaNegros, cotaPCD, totalPontuacao, dataNascimento]);
  }
  
  for (var cargo in areas) {
    areas[cargo].sort((a, b) => a[1].localeCompare(b[1]));
    var abaNome = "DO - " + cargo;
    var destino = ss.getSheetByName(abaNome) || ss.insertSheet(abaNome);
    destino.clear();
    
    destino.getRange("A1:E1").merge().setValue("CARGO: " + cargo).setHorizontalAlignment("center").setBorder(true, true, true, true, true, true).setFontWeight("bold");
    destino.getRange("A2:E2").merge().setValue("Área de Atuação - " + (dados.find(row => row[cabecalho.indexOf("CARGO")] === cargo) ? dados.find(row => row[cabecalho.indexOf("CARGO")] === cargo)[areaIndex] : "Não informada")).setHorizontalAlignment("center").setBorder(true, true, true, true, true, true).setFontWeight("bold");
    
    var headerRange = destino.getRange(3, 1, 1, colunas.length);
    headerRange.setValues([colunas]);
    headerRange.setHorizontalAlignment("center").setBorder(true, true, true, true, true, true).setFontWeight("bold");
    
    var dataRange = destino.getRange(4, 1, areas[cargo].length, colunas.length);
    dataRange.setValues(areas[cargo]);
    dataRange.setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
    
    destino.hideColumns(colunas.length); 
  }
}
