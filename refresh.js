function UpdateImportJSON() {
  try {
    const url = "https://query1.finance.yahoo.com/v8/finance/chart/MXN=X?period1=1598922000&period2=1598922000&interval=1d"
    const data = ImportJSON(url)

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const startRow = 1; // Inicia na primeira linha
    const startColumn = 1; // Inicia na primeira coluna
      
    // Escreve a matriz na planilha
    sheet.getRange(startRow, startColumn, data.length, data[0].length).setValues(data);

    // Calcula a última linha de dados
    const lastRow = startRow + data.length - 1;
    
    // Adiciona um carimbo de data/hora na primeira linha após a última linha de dados
    const timestamp = new Date();

    // Calcula a posição onde será realizado o registro da última atualização
    const lastLineRow = lastRow + 2
    const lastLineColumn = startColumn + 1
  
    sheet.getRange(lastLineRow, startColumn).setValue('Última atualização');
    sheet.getRange(lastLineRow, lastLineColumn).setValue(timestamp);
  } catch(error) {
    Logger.log("Erro ao atualizar os dados: " + error.message);
  }
}
