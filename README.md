
# Explicação do Código `copiarDadosPorArea`

Este script do Google Apps Script copia e organiza os dados de uma planilha chamada `"teste"`, criando novas abas para cada tipo de cargo, onde os dados são filtrados, processados e formatados.

---

## 📌 **Passo a Passo do Código**

### 1️⃣ **Obter a Planilha e a Aba de Origem**
```javascript
var ss = SpreadsheetApp.getActiveSpreadsheet();
var origem = ss.getSheetByName("teste");
```
- Obtém a planilha ativa.
- Tenta acessar a aba chamada `"teste"`.
- Se a aba não for encontrada, registra um erro no Logger e interrompe a execução.

### 2️⃣ **Definir as Colunas a Serem Copiadas**
```javascript
var colunas = ["CARGO", "NOME COMPLETO", "PPI", "PCD", "TOTAL PONTUAÇÃO"];
```
- Lista os nomes das colunas que serão extraídas da planilha original.

### 3️⃣ **Identificar os Índices das Colunas**
```javascript
var cabecalho = origem.getDataRange().getValues()[0];
var indices = [
  cabecalho.indexOf("CARGO"),
  cabecalho.indexOf("NOME COMPLETO:"),
  cabecalho.indexOf("[COTA] DESEJA CONCORRER ÀS VAGAS DESTINADAS A CANDIDATOS NEGROS E INDÍGENAS, CONFORME O ITEM 4 DA RESERVA DAS VAGAS PARA NEGROS E INDÍGENAS DO EDITAL?"),
  cabecalho.indexOf("[COTA] DESEJA CONCORRER ÀS VAGAS DESTINADAS ÀS PESSOAS COM DEFICIÊNCIA(PCD), CONFORME O ITEM 5 DO EDITAL?"),
  cabecalho.indexOf("TOTAL_PONTUAÇÃO")
].filter(i => i !== -1);
```
- Localiza os índices das colunas dentro do cabeçalho.
- Filtra para garantir que apenas índices válidos sejam utilizados.

```javascript
var areaIndex = cabecalho.indexOf("ÁREA");
if (indices.length !== colunas.length || areaIndex === -1) {
  Logger.log("Algumas colunas não foram encontradas.");
  return;
}
```
- Adiciona a coluna `"ÁREA"` e verifica se todas as colunas necessárias foram encontradas.

### 4️⃣ **Processar os Dados**
```javascript
var dados = origem.getDataRange().getValues();
var areas = {};
```
- Obtém todos os dados da aba `"teste"` e armazena em `dados`.
- Cria um objeto `areas` para organizar os dados por cargo.

```javascript
for (var i = 1; i < dados.length; i++) {
  var cargo = dados[i][indices[0]];
  var area = dados[i][areaIndex];
  var nomeCompleto = dados[i][indices[1]].replace(/[^A-Z\s]/gi, "").trim().toUpperCase();
  var cotaNegros = dados[i][indices[2]] === "SIM" ? "*" : "";
  var cotaPCD = dados[i][indices[3]] === "SIM" ? "**" : "";
  var totalPontuacao = dados[i][indices[4]];
```
- Remove números e caracteres especiais do nome e o transforma em maiúsculas.
- Adiciona `*` para candidatos PPI e `**` para PCD.

```javascript
if (!areas[cargo]) areas[cargo] = [];
areas[cargo].push([cargo, nomeCompleto, cotaNegros, cotaPCD, totalPontuacao]);
```
- Armazena os dados no objeto `areas`, organizando-os por cargo.

### 5️⃣ **Criar e Preencher as Novas Abas**
```javascript
for (var cargo in areas) {
  areas[cargo].sort((a, b) => a[1].localeCompare(b[1])); // Ordenar por nome completo
  var abaNome = "DO - " + cargo;
  var destino = ss.getSheetByName(abaNome) || ss.insertSheet(abaNome);
  destino.clear();
```
- Para cada cargo encontrado, cria uma nova aba (ou reutiliza uma existente).
- Ordena os dados pelo nome completo.

```javascript
destino.getRange("A1:E1").merge().setValue("CARGO: " + cargo)
  .setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
destino.getRange("A2:E2").merge().setValue("Área de Atuação - " + (dados.find(row => row[cabecalho.indexOf("CARGO")] === cargo) ? dados.find(row => row[cabecalho.indexOf("CARGO")] === cargo)[areaIndex] : "Não informada"))
  .setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
```
- Mescla as células `A1:E1` e `A2:E2` para exibir o nome do cargo e a área de atuação.
- Centraliza e aplica bordas.

```javascript
var headerRange = destino.getRange(3, 1, 1, colunas.length);
headerRange.setValues([colunas]);
headerRange.setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
```
- Adiciona o cabeçalho e aplica formatação.

```javascript
var dataRange = destino.getRange(4, 1, areas[cargo].length, colunas.length);
dataRange.setValues(areas[cargo]);
dataRange.setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
```
- Insere os dados e aplica a formatação de alinhamento centralizado e bordas.

---

## 🔍 **Resumo das Melhorias**
✔ Nomes formatados para maiúsculas, removendo números e caracteres especiais.  
✔ Ordenação dos nomes em cada aba.  
✔ Mesclagem e centralização de títulos nas novas abas.  
✔ Aplicação de bordas em todas as células.  
