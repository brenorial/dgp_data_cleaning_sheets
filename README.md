## Explicação do Código `processarPlanilha()`

Esse código é uma função que processa dados de uma planilha no Google Sheets e cria novas abas para cada vaga encontrada, copiando as linhas correspondentes para essas abas. Abaixo está o detalhamento de cada parte do código:

### 1. Acessando a Planilha Ativa e a Aba "BASE"
```javascript
var ss = SpreadsheetApp.getActiveSpreadsheet();
var baseSheet = ss.getSheetByName('BASE');
var data = baseSheet.getDataRange().getValues();
```
- **`ss`**: Obtém a planilha ativa no Google Sheets.
- **`baseSheet`**: Acessa a aba chamada "BASE" da planilha ativa.
- **`data`**: Obtém todos os dados da aba "BASE", incluindo cabeçalhos e valores, e armazena em um array de duas dimensões.

### 2. Definindo a Variável `novaLinha`
```javascript
var novaLinha;
```
- **`novaLinha`**: A variável será utilizada para armazenar a linha em que as novas informações serão inseridas nas abas criadas para cada "vaga".

### 3. Iterando sobre os Dados da Aba "BASE"
```javascript
for (var i = 0; i < data.length; i++) {
```
- Este laço percorre todas as linhas da aba "BASE". O objetivo é verificar se a célula da coluna 1 (primeira coluna) contém um texto que começa com "Vaga:".

### 4. Verificando se a Linha Contém uma Vaga
```javascript
if (data[i][0] && data[i][0].toString().startsWith("Vaga:")) {
  var nomeVaga = data[i][0].toString().split(":")[1].trim();
```
- Verifica se a primeira célula da linha contém um texto que começa com "Vaga:".
- **`nomeVaga`**: Extrai o nome da vaga a partir do texto que segue após "Vaga:". Isso é feito dividindo a string e pegando a parte após os dois pontos.

### 5. Criando ou Limpar uma Aba para a Vaga
```javascript
var novaAba = ss.getSheetByName(nomeVaga);
if (!novaAba) {
  novaAba = ss.insertSheet(nomeVaga);
} else {
  novaAba.clear();
}
```
- **`novaAba`**: Tenta obter uma aba com o nome da vaga.
- Se a aba não existir, ela é criada usando **`insertSheet`**.
- Se a aba já existir, ela é limpa com **`clear`** para garantir que ela começará com dados novos.

### 6. Processando as Linhas Correspondentes à Vaga
```javascript
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
```
- **`novaLinha`**: Inicializa a linha de inserção como 1 (primeira linha da nova aba).
- **`for (var j = i + 1; j < data.length; j++)`**: Itera pelas linhas seguintes à vaga encontrada, copiando os dados até encontrar outra vaga ou uma linha vazia.
- **`if (!data[j][0])`**: Se a célula da coluna 1 estiver vazia, a linha é ignorada.
- **`if (data[j][0].toString().startsWith("Vaga:"))`**: Se a célula da coluna 1 começar com "Vaga:", a iteração é interrompida (indica o início de uma nova vaga).
- **`novaAba.getRange(novaLinha, 1, 1, data[j].length).setValues([data[j]])`**: Copia os dados da linha para a nova aba na linha atual indicada por **`novaLinha`**.
- **`novaLinha++`**: Avança a linha para a próxima linha da aba.

### 7. Finalizando o Processo
- O processo continua até todas as linhas da aba "BASE" serem verificadas e processadas. Cada vaga encontrada terá sua própria aba criada (ou limpa) e as linhas correspondentes a essa vaga serão copiadas para a nova aba.

---

### Resumo do Funcionamento
Essa função automatiza a organização dos dados da planilha, criando uma nova aba para cada "vaga" encontrada e copiando as linhas seguintes (até encontrar uma nova vaga ou uma linha vazia) para essas abas. Isso ajuda a segmentar os dados de acordo com as vagas e facilita a análise separada por área ou tipo de vaga.

