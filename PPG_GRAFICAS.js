// IDs fornecidos
const MAIN_SPREADSHEET_ID = "1j1RCyvofFLoSFc3eTdV8O3nf78Bxu4nIiXtWRjJnBKg";
const FOLDER_ROOT_ID = "0AFFS2ipqlJogUk9PVA";

// ----------------------------------------------------------------------
// NOVA FUNÇÃO: Retorna a lista de gráficas da coluna U para montar o Modal
// ----------------------------------------------------------------------
function getListaGraficasConsolidacao() {
  const mainSs = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const sheetApoio = mainSs.getSheetByName("apoio");
  
  if (!sheetApoio) return [];
  
  // Lê da U2 até a última linha com dados
  const maxRow = sheetApoio.getLastRow();
  if (maxRow < 2) return [];
  
  const graficasValores = sheetApoio.getRange("U2:U" + maxRow).getValues();
  
  // Mapeia, remove vazios e tira os espaços das pontas
  let graficasBrutas = graficasValores.map(r => r[0] ? r[0].toString().trim() : "").filter(String);
  
  // Retorna apenas valores únicos para não duplicar checkbox no modal
  return [...new Set(graficasBrutas)];
}

// ----------------------------------------------------------------------
// FUNÇÃO PRINCIPAL DE CONSOLIDAÇÃO (Híbrida: Roda via Portal ou Planilha)
// ----------------------------------------------------------------------
function consolidarArquivos(inputDoPortal) {
  const mainSs = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  
  let filtroInput = "";
  
  // 1. CHECAGEM DE ONDE VEM O COMANDO (Portal vs Planilha)
  if (typeof inputDoPortal === 'string') {
    // Veio do botão do Portal Web App (já vem tratado pela função JavaScript)
    filtroInput = inputDoPortal.trim().toLowerCase();
  } else {
    // Veio do botão desenhado na Planilha (tenta usar o pop-up nativo)
    try {
      let ui = SpreadsheetApp.getUi();
      let resposta = ui.prompt(
        "Atualização de Gráficas",
        "Digite o nome da gráfica que deseja atualizar (ex: FormaCerta, LogPrint).\nPara atualizar várias, separe por vírgula.\nDeixe em branco para atualizar TODAS as gráficas:",
        ui.ButtonSet.OK_CANCEL
      );

      if (resposta.getSelectedButton() !== ui.Button.OK) {
        return "Execução cancelada pelo usuário."; 
      }
      filtroInput = resposta.getResponseText().trim().toLowerCase();
    } catch (e) {
      Logger.log("Execução automatizada ou sem interface de usuário.");
    }
  }

  let startTime = new Date().getTime();
  const MAX_EXECUTION_TIME = 270000; // 4.5 minutos (270.000 ms)

  Logger.log("Iniciando execução...");
 
  const sheetApoio = mainSs.getSheetByName("apoio");
  const sheetArvore = mainSs.getSheetByName("arvore");
  const sheetEspec = mainSs.getSheetByName("Espec");
  const sheetTiragem = mainSs.getSheetByName("Tiragem");
 
  let sheetLog = mainSs.getSheetByName("log");
  if (!sheetLog) {
    sheetLog = mainSs.insertSheet("log");
    sheetLog.appendRow(["Nome da Gráfica", "Nome do Arquivo", "ID do Arquivo", "Data de Modificação"]);
  }
 
  const graficasValores = sheetApoio.getRange("U2:U" + sheetApoio.getMaxRows()).getValues();
  let graficas = graficasValores.map(r => r[0]).filter(String);

  // 2. APLICA O FILTRO DIGITADO PELO USUÁRIO (se existir)
  if (filtroInput !== "") {
    let digitadas = filtroInput.split(",").map(g => g.trim());
    graficas = graficas.filter(g => digitadas.includes(g.toLowerCase()));
    
    if (graficas.length === 0) {
      return "Nenhuma gráfica selecionada encontrada na base de dados.";
    }
  }
 
  const logData = sheetLog.getDataRange().getValues();
  let logMap = {}; 
  for (let i = 1; i < logData.length; i++) {
    let fileId = String(logData[i][2]).trim(); 
    let dataMod = logData[i][3]; 
    logMap[fileId] = { rowIndex: i + 1, dataModificacao: new Date(dataMod).getTime() };
  }
 
  const parentFolder = DriveApp.getFolderById(FOLDER_ROOT_ID);
 
  for (let i = 0; i < graficas.length; i++) {
    if (new Date().getTime() - startTime > MAX_EXECUTION_TIME) {
      return "O tempo limite do Google (5 min) foi atingido. As gráficas pendentes foram pausadas. Rode o script novamente para continuar de onde parou."; 
    }

    let graficaNameReal = String(graficas[i]).trim();
    Logger.log(">> BUSCANDO PASTA: " + graficaNameReal);
    
    let pastasGrafica = parentFolder.searchFolders("title = '" + graficaNameReal + "'");
    if (!pastasGrafica.hasNext()) continue;
    let pastaGrafica = pastasGrafica.next();
 
    let pastasCiclo = pastaGrafica.searchFolders("title contains 'Ciclo'");
    if (!pastasCiclo.hasNext()) continue;
    let pastaCiclo = pastasCiclo.next();
 
    let subpastas = pastaCiclo.getFolders();
    let mapArquivosUnicos = {}; 
 
    while (subpastas.hasNext()) {
      let subpasta = subpastas.next();
      let nomeSubpasta = subpasta.getName().toLowerCase();

      // PROTEÇÃO DE VERSÕES ANTERIORES
      if (nomeSubpasta.includes("versões") || nomeSubpasta.includes("versoes") || nomeSubpasta.includes("anteriores")) {
         continue; 
      }

      let arquivos = subpasta.getFiles();
      while (arquivos.hasNext()) {
        let arquivo = arquivos.next();
        let nomeArquivo = arquivo.getName();

        if (nomeArquivo.startsWith("~$")) continue;

        let dataMod = arquivo.getLastUpdated().getTime();
        // Fica com a versão mais recente caso haja duplicidade de nome
        if (!mapArquivosUnicos[nomeArquivo] || dataMod > mapArquivosUnicos[nomeArquivo].dataModificacao) {
            mapArquivosUnicos[nomeArquivo] = { file: arquivo, dataModificacao: dataMod };
        }
      }
    }
 
    let arquivosParaProcessar = Object.values(mapArquivosUnicos).map(obj => obj.file);
    if (arquivosParaProcessar.length === 0) continue;
 
    let precisaAtualizarGrafica = false;
    for (let f = 0; f < arquivosParaProcessar.length; f++) {
      let arq = arquivosParaProcessar[f];
      let ultimaMod = arq.getLastUpdated().getTime();
      if (!logMap[arq.getId()] || logMap[arq.getId()].dataModificacao !== ultimaMod) {
        precisaAtualizarGrafica = true;
        break;
      }
    }
 
    if (!precisaAtualizarGrafica) continue;
 
    Logger.log(" - ATUALIZANDO GRÁFICA (" + arquivosParaProcessar.length + " arquivos únicos encontrados)");
    apagarDadosAntigos(sheetArvore, 29, graficaNameReal); 
    apagarDadosAntigos(sheetEspec, 80, graficaNameReal); 
    apagarDadosAntigos(sheetTiragem, null, graficaNameReal); 

    for (let f = 0; f < arquivosParaProcessar.length; f++) {
      if (new Date().getTime() - startTime > MAX_EXECUTION_TIME) {
        return "O tempo limite (5 min) foi atingido durante o processamento. Rode o script novamente para finalizar a atualização."; 
      }

      let arquivo = arquivosParaProcessar[f];
 
      try {
        let sourceSs = SpreadsheetApp.openById(arquivo.getId());
        let sourceEsp = sourceSs.getSheetByName("Esp. Técnica (legado)");
        let marcaValue = "";
        if (sourceEsp && sourceEsp.getLastRow() >= 2) {
          marcaValue = sourceEsp.getRange("F2").getValue();
        }
 
        copiarAbaMapeada(sourceSs.getSheetByName("Árvore (legado)"), sheetArvore, graficaNameReal, 29);
        copiarAbaMapeada(sourceSs.getSheetByName("Esp. Técnica (legado)"), sheetEspec, graficaNameReal, 80);
        copiarAbaMapeada(sourceSs.getSheetByName("Tiragem (legado)"), sheetTiragem, graficaNameReal, null, { colIndex: 20, value: marcaValue });
 
        if (logMap[arquivo.getId()]) {
          sheetLog.getRange(logMap[arquivo.getId()].rowIndex, 1, 1, 4).setValues([[graficaNameReal, arquivo.getName(), arquivo.getId(), arquivo.getLastUpdated()]]);
        } else {
          sheetLog.appendRow([graficaNameReal, arquivo.getName(), arquivo.getId(), arquivo.getLastUpdated()]);
          logMap[arquivo.getId()] = { rowIndex: sheetLog.getLastRow(), dataModificacao: arquivo.getLastUpdated().getTime() };
        }
      } catch (e) {
        Logger.log(" - ERRO: " + e.message);
      }
    }
  }
  
  return "Processo finalizado com sucesso!";
}

function copiarAbaMapeada(sourceSheet, destSheet, graficaName, colIndexGrafica, injectData) {
 if (!sourceSheet || !destSheet) return;
 const sourceValues = sourceSheet.getDataRange().getValues();
 if (sourceValues.length < 2) return; 

const sourceHeaders = sourceValues[0];
 const destHeaders = destSheet.getRange(1, 1, 1, destSheet.getMaxColumns()).getValues()[0];
 
let map = {};
 for (let s = 0; s < sourceHeaders.length; s++) {
 let sTitle = String(sourceHeaders[s]).trim().toLowerCase();
 if (!sTitle) continue;
 for (let d = 0; d < destHeaders.length; d++) {
 if (String(destHeaders[d]).trim().toLowerCase() === sTitle) {
 map[s] = d;
 break;
 }
 }
 }
 
let idxGrafica = -1;
 if (colIndexGrafica) {
 idxGrafica = colIndexGrafica - 1; 
} else {
 idxGrafica = destHeaders.findIndex(h => String(h).trim().toLowerCase() === "gráfica" || String(h).trim().toLowerCase() === "grafica");
 }
 
let maxColsNecessarias = destHeaders.length;
 if (idxGrafica >= maxColsNecessarias) {
 maxColsNecessarias = idxGrafica + 1;
 }
 if (injectData && injectData.colIndex > maxColsNecessarias) {
 maxColsNecessarias = injectData.colIndex;
 }
 
let newRows = [];
 for (let i = 1; i < sourceValues.length; i++) {
 let row = new Array(maxColsNecessarias).fill("");
 let sourceRow = sourceValues[i];
 
for (let s in map) {
 row[map[s]] = sourceRow[s];
 }
 
if (idxGrafica >= 0) {
 row[idxGrafica] = graficaName;
 }
 
if (injectData) {
 row[injectData.colIndex - 1] = injectData.value;
 }
 
newRows.push(row);
 }
 
if (newRows.length > 0) {
 if (destSheet.getMaxColumns() < maxColsNecessarias) {
 destSheet.insertColumnsAfter(destSheet.getMaxColumns(), maxColsNecessarias - destSheet.getMaxColumns());
 }
 destSheet.getRange(destSheet.getLastRow() + 1, 1, newRows.length, maxColsNecessarias).setValues(newRows);
 }
}

function apagarDadosAntigos(sheet, colIndexGrafica, graficaName) {
 if (!sheet) return;
 if (!colIndexGrafica) {
 const headers = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues()[0];
 let idx = headers.findIndex(h => String(h).trim().toLowerCase() === "gráfica" || String(h).trim().toLowerCase() === "grafica");
 if (idx >= 0) {
 colIndexGrafica = idx + 1; 
} else {
 return; 
}
 }

 const lastRow = sheet.getLastRow();
 if (lastRow < 2) return;
 const colData = sheet.getRange(2, colIndexGrafica, lastRow - 1, 1).getValues();
 
for (let i = colData.length - 1; i >= 0; i--) {
 if (String(colData[i][0]).trim().toLowerCase() === String(graficaName).trim().toLowerCase()) {
 sheet.deleteRow(i + 2); 
}
 }
}