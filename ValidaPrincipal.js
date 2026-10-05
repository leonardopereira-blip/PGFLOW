// =========================================================================
// ValidaPrincipal.gs — FLUXO "VALIDAÇÃO DE CUSTOS PRINCIPAL" + DASHBOARD PCP
// UI: Index.html (view-custos-principal)
// Entrada: SKUs/gráficas selecionados na planilha DE PARA
// Saída:   e-mail [ESTEIRA_CALCULADORA_PRINCIPAL] + job na Queue_Calculadora
// Depende de: ValidaComum.gs (constantes/helpers) e WebApp.gs (gravarNaFila)
// NÃO chamar nada daqui a partir da Validação Reentradas.
// =========================================================================

function checarStatusEmailPrincipal(hashLote) {
 try {
 var emailNaCaixa = false;
 if (hashLote) {
 var query = 'in:inbox subject:"' + hashLote + '"';
 var threads = GmailApp.search(query, 0, 1);
 emailNaCaixa = (threads.length > 0);
 }
 return { emailNaCaixa: emailNaCaixa };
 } catch (e) {
 return { erro: e.message };
 }
}

function obterDadosDashboardPCP() {
 try {
 var idPlanilha = "1lXG3CJkBSvuIO-Yv1jS5jVoYD5wVUAnZO1z4iph_YZ8";
 var ss = SpreadsheetApp.openById(idPlanilha);
 var aba = ss.getSheetByName("Consolidado") || ss.getSheets()[0];
 var dados = aba.getDataRange().getValues();
 
if (dados.length < 2) return { erro: "A planilha selecionada está vazia ou sem dados estruturados." };
 
var head = dados[0].map(function(h) { return h.toString().toLowerCase(); });
 var idxTiragem = 14; 
 var idxGrafica = 15; 
 var idxStatus = 21; 
 var idxDias = -1; 
 var idxChamado = 0; 
 
for(var h = 0; h < head.length; h++) {
 if(head[h].indexOf("dias de atraso") !== -1 || head[h].indexOf("atraso") !== -1) {
 idxDias = h;
 break;
 }
 }
 
var atrasados = [];
 var emDia = [];
 var finalizados = [];
 
for (var i = 1; i < dados.length; i++) {
 var row = dados[i];
 if (row.length <= idxStatus) continue; 
 
var statusOriginal = row[idxStatus] ? row[idxStatus].toString().trim() : "";
 if (!statusOriginal) continue;
 
var stLower = statusOriginal.toLowerCase();
 var item = {
 linha: i + 1,
 idItem: row[idxChamado] ? row[idxChamado].toString() : ("L." + (i+1)),
 tiragem: parseFloat(row[idxTiragem]) || 0,
 grafica: row[idxGrafica] ? row[idxGrafica].toString() : "-",
 diasAtraso: idxDias !== -1 ? (parseInt(row[idxDias]) || 0) : 0,
 status: statusOriginal
 };
 
if (stLower.indexOf("atrasado") !== -1 || stLower.indexOf("atraso") !== -1) {
 atrasados.push(item);
 } else if (stLower.indexOf("finalizado") !== -1 || stLower.indexOf("conclu") !== -1) {
 finalizados.push(item);
 } else {
 emDia.push(item);
 }
 }
 
return {
 sucesso: true,
 atrasados: atrasados,
 emDia: emDia,
 finalizados: finalizados
 };
 } catch (e) {
 return { erro: "Erro ao extrair dashboard: " + e.message };
 }
}

function obterDadosApoioValidaPrin() {
 try {
 var ID_PLANILHA_DEPARA = "1j1RCyvofFLoSFc3eTdV8O3nf78Bxu4nIiXtWRjJnBKg";
 var ss = SpreadsheetApp.openById(ID_PLANILHA_DEPARA);
 var aba = ss.getSheetByName("Apoio");
 if (!aba) throw new Error("Aba 'Apoio' não encontrada na planilha DE PARA.");
 var dados = aba.getDataRange().getValues();
 var registros = [];
 for (var i = 1; i < dados.length; i++) {
 var g = dados[i][0] ? dados[i][0].toString().trim() : "";
 var m = dados[i][1] ? dados[i][1].toString().trim() : "";
 var s = dados[i][2] ? dados[i][2].toString().trim() : "";
 var k = dados[i][3] ? dados[i][3].toString().trim() : "";
 var t = dados[i][4] ? dados[i][4].toString().trim() : "";
 if (g !== "" || m !== "" || s !== "" || k !== "") {
 registros.push({ grafica: g, marca: m, serie: s, sku: k, tecnologia: t });
 }
 }
 return registros;
 } catch (e) {
 return { erro: e.message };
 }
}

function processarValidaPrincipal(payload) {
 try {
 var ID_PLANILHA = "1j1RCyvofFLoSFc3eTdV8O3nf78Bxu4nIiXtWRjJnBKg";
 var ss = SpreadsheetApp.openById(ID_PLANILHA);
 var skusFiltro = payload.skus || [];
 if (skusFiltro.length === 0) {
 return { erro: "Nenhum SKU selecionado para processamento." };
 }
 
// 1. MAPEAR CADA SKU
 var abaApoio = ss.getSheetByName("Apoio").getDataRange().getValues();
 var mapSkuData = {};
 var graficasFiltro = payload.graficas || []; 

 for(var i=1; i<abaApoio.length; i++){
 var k = abaApoio[i][3] ? abaApoio[i][3].toString().trim().toLowerCase() : "";
 var g = abaApoio[i][0] ? abaApoio[i][0].toString().trim() : "DESCONHECIDA";
 
 if(k){
 if (graficasFiltro.length > 0 && graficasFiltro.indexOf(g) !== -1) {
 mapSkuData[k] = {
 grafica: g,
 marca: abaApoio[i][1] ? abaApoio[i][1].toString().trim() : "",
 serie: abaApoio[i][2] ? abaApoio[i][2].toString().trim() : "",
 tecnologia: abaApoio[i][4] ? abaApoio[i][4].toString().trim() : "PLANA"
 };
 } 
 else if (!mapSkuData[k]) {
 mapSkuData[k] = {
 grafica: g,
 marca: abaApoio[i][1] ? abaApoio[i][1].toString().trim() : "",
 serie: abaApoio[i][2] ? abaApoio[i][2].toString().trim() : "",
 tecnologia: abaApoio[i][4] ? abaApoio[i][4].toString().trim() : "PLANA"
 };
 }
 }
 }
 
// 🔧 CORREÇÃO 2: Garante que as abas leiam a coluna de Gráfica corretamente
 function filtrarEExtrair(nomesAbaPossiveis, maxCols) {
 var aba = null;
 for (var n = 0; n < nomesAbaPossiveis.length; n++) {
 aba = ss.getSheetByName(nomesAbaPossiveis[n]);
 if (aba) break;
 }
 
 if (!aba) return [];
 var dados = aba.getDataRange().getValues();
 if (dados.length <= 1) return [];
 
 var headers = dados[0];
 var idx = { sku: -1, grafica: -1, kit: -1 };
 
 for (var j = 0; j < headers.length; j++) {
 var h = headers[j].toString().trim().toLowerCase();
 // Acha a coluna SKU
 if (idx.sku === -1 && (h === "sku" || h.indexOf("código sku") !== -1 || h.indexOf("codigo sku") !== -1 || (h.indexOf("código") !== -1 && h.indexOf("sku") !== -1))) {
 idx.sku = j;
 }
 // Acha a coluna Gráfica
 if (idx.grafica === -1 && (h === "grafica" || h === "gráfica" || h === "grafico" || h === "gráfico")) {
 idx.grafica = j;
 }
 // Acha a coluna do Kit (NOVO)
 if (idx.kit === -1 && (h.indexOf("código kit") !== -1 || h.indexOf("codigo kit") !== -1 || h === "kit")) {
 idx.kit = j;
 }
 }

 // Pré-mapeamento: rastrear quais códigos de Kit pertencem aos SKUs selecionados
 var kitsValidos = {};
 if (idx.kit !== -1 && idx.sku !== -1) {
 for (var r = 1; r < dados.length; r++) {
 var rSku = dados[r][idx.sku] ? dados[r][idx.sku].toString().trim().toLowerCase() : "";
 var rKit = dados[r][idx.kit] ? dados[r][idx.kit].toString().trim().toLowerCase() : "";
 if (rSku !== "" && rKit !== "") {
 var matchS = skusFiltro.some(function(s) { return s.toString().trim().toLowerCase() === rSku; });
 if (matchS) kitsValidos[rKit] = true;
 }
 }
 }

 var result = [];
 
 for (var r = 1; r < dados.length; r++) {
 var rowSku = idx.sku !== -1 && dados[r][idx.sku] ? dados[r][idx.sku].toString().trim().toLowerCase() : "";
 var rowGrafica = idx.grafica !== -1 && dados[r][idx.grafica] ? dados[r][idx.grafica].toString().trim().toLowerCase() : "";
 var rowKit = idx.kit !== -1 && dados[r][idx.kit] ? dados[r][idx.kit].toString().trim().toLowerCase() : "";
 
 // 1. Checa se o SKU foi selecionado diretamente
 var matchSku = skusFiltro.some(function(s) { 
 return s.toString().trim().toLowerCase() === rowSku; 
 });

 // 2. Lógica de KIT (Inteligente): Checa se é a linha em branco do Kit que o usuário escolheu
 var isHeaderKit = false;
 if (rowSku === "" && rowKit !== "" && kitsValidos[rowKit]) {
 isHeaderKit = true;
 }
 
 // Checa a Gráfica
 var matchGrafica = true; 
 if (graficasFiltro.length > 0 && idx.grafica !== -1 && rowGrafica !== "") {
 matchGrafica = graficasFiltro.some(function(g) {
 return rowGrafica.indexOf(g.toString().trim().toLowerCase()) !== -1 || g.toString().trim().toLowerCase().indexOf(rowGrafica) !== -1;
 });
 }

 // Se for o SKU escolhido OU o cabeçalho do Kit, e passar na gráfica, inclui!
 if ((matchSku || isHeaderKit) && matchGrafica) {
 var rowData = dados[r].slice(0, maxCols);
 while (rowData.length < maxCols) rowData.push(""); 
 result.push(rowData);
 }
 }
 return result;
 }
 
var matTiragem = filtrarEExtrair(["Tiragem", "PPG - Tiragem"], 19);
 var matArvore = filtrarEExtrair(["Árvore de produto", "Arvore de produto", "Arvore", "PPG - Arvore de Produto"], 28); 
var matEspec = filtrarEExtrair(["Especificação Técnica", "Especificacao Tecnica", "Espec", "PPG - Especificações"], 79); 

if (matEspec.length === 0) {
 return { erro: "Nenhum dado encontrado para as diretrizes selecionadas. Verifique se os SKUs passados existem nas abas fonte." };
 }
 
 function colarDestino(nomeAbaDestino, matriz, maxCol) {
 var aba = ss.getSheetByName(nomeAbaDestino);
 if (!aba) return;
 var lastRow = aba.getLastRow();
 if (lastRow > 1) {
 aba.getRange(2, 1, aba.getMaxRows() - 1, maxCol).clearContent();
 }
 if (matriz.length > 0) {
 aba.getRange(2, 1, matriz.length, matriz[0].length).setValues(matriz);
 }
 }
 
colarDestino("PPG - Tiragem", matTiragem, 19);
 colarDestino("PPG - Arvore de Produto", matArvore, 28);
 colarDestino("PPG - Especificações", matEspec, 79);
 
var numLinhas = matEspec.length;
 var abaNc = ss.getSheetByName("NC - PPG Especificações");
 var limNec = 6 + numLinhas - 1;
 if (abaNc.getMaxRows() < limNec) abaNc.insertRowsAfter(abaNc.getMaxRows(), (limNec - abaNc.getMaxRows()) + 5);
 var maxColNc = abaNc.getLastColumn();
 var linhaTipos = abaNc.getRange(5, 1, 1, maxColNc).getValues()[0];
 var linhaInstrucoes = abaNc.getRange(4, 1, 1, maxColNc).getValues()[0];
 for (var col = 1; col <= maxColNc; col++) {
 var tipo = linhaTipos[col - 1] ? linhaTipos[col - 1].toString().trim() : "";
 if (tipo.toLowerCase() === "manual") {
 var instCol = linhaInstrucoes[col - 1] ? linhaInstrucoes[col - 1].toString().toLowerCase() : "";
 var ehNumerico = (instCol.indexOf("cm") !== -1 || instCol.indexOf("área") !== -1 ||
 instCol.indexOf("area") !== -1 || instCol.indexOf("mm") !== -1);
 var padrao = ehNumerico ? "0" : "Não";
 abaNc.getRange(6, col, numLinhas, 1).setValue(padrao);
 } else {
 var c0 = abaNc.getRange(6, col);
 if (c0.getFormula() !== "" || c0.getValue() !== "") {
 c0.copyTo(abaNc.getRange(6, col, numLinhas), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
 }
 }
 }
 
SpreadsheetApp.flush();
 Utilities.sleep(1500);
 var matrizBruta = abaNc.getRange(6, 3, numLinhas, 79).getValues();
 
var lotesPorGrafica = {};
 for (var mIdx = 0; mIdx < matrizBruta.length; mIdx++) {
 var rowCopy = matrizBruta[mIdx].slice();
 var skuClean = rowCopy[0] ? rowCopy[0].toString().trim().toLowerCase() : "";
 
if (skuClean !== "" && skuClean !== "-") {
 var info = mapSkuData[skuClean] || { grafica: "DESCONHECIDA", marca: "", serie: "", tecnologia: "PLANA" };
 var graf = info.grafica || "DESCONHECIDA";
 
if (!lotesPorGrafica[graf]) {
 lotesPorGrafica[graf] = {
 nomeGrafica: graf,
 marcas: {},
 series: {},
 dadosSkus: [],
 dadosCalcGeral: [],
 tiragemTotal: 0
 };
 }
 if(info.marca) lotesPorGrafica[graf].marcas[info.marca] = true;
 if(info.serie) lotesPorGrafica[graf].series[info.serie] = true;
 
lotesPorGrafica[graf].dadosSkus.push(rowCopy.map(serializarValor));
 }
 }
 
var abaCalcGeral = ss.getSheetByName("NC - CALC Geral");
 if (abaCalcGeral) {
 var lastRowG = abaCalcGeral.getLastRow();
 if (lastRowG < 10) lastRowG = 100;
 
var rangeGeral = abaCalcGeral.getRange(1, 1, lastRowG, 21).getValues();
 
for (var idxG = 3; idxG < rangeGeral.length; idxG++) {
 var skuVal = rangeGeral[idxG][11] ? rangeGeral[idxG][11].toString().trim().toLowerCase() : "";
 if (skuVal.indexOf("-") === 0) skuVal = skuVal.substring(1).trim();
 
if (skuVal !== "" && skuVal !== "-" && skuVal.indexOf("total") === -1 && skuVal.indexOf("sku") === -1) {
 
var info = mapSkuData[skuVal];
 if (info && lotesPorGrafica[info.grafica]) {
 var graf = info.grafica;
 lotesPorGrafica[graf].dadosCalcGeral.push([
 skuVal.toUpperCase(), 
 rangeGeral[idxG][15] !== undefined ? rangeGeral[idxG][15] : "",
 rangeGeral[idxG][16] !== undefined ? rangeGeral[idxG][16] : "", 
 rangeGeral[idxG][17] !== undefined ? rangeGeral[idxG][17] : "", 
 rangeGeral[idxG][18] !== undefined ? rangeGeral[idxG][18] : "", 
 info.tecnologia || "PLANA", 
 rangeGeral[idxG][20] !== undefined ? rangeGeral[idxG][20] : "" 
 ]);
 var s = parseFloat(rangeGeral[idxG][18]);
 if (!isNaN(s)) lotesPorGrafica[graf].tiragemTotal += s;
 }
 }
 }
 }
 
// PEGA O HASH DO PAYLOAD SE EXISTIR, SENÃO GERA UM NOVO (Isso garante não perder o tracking!)
 var hashLote = payload.hashLote || new Date().getTime().toString();
 var emailUser = Session.getActiveUser().getEmail();
 var qtdLotes = 0;
 var idsJobsCriados = []; 

// Mesma pasta usada pelas Reentradas (definida em ValidaComum.gs).
 var pastaDest = DriveApp.getFolderById(ID_PASTA_DEPARA_GERADO);
 var resumoLote = [];
 var abasPermitidas = ["PPG - Especificações", "PPG - Arvore de Produto", "PPG - Tiragem", "NC - CALC Geral", "NC - PPG Especificações", "TAB - Ajustes de Nomes"];
 
 function formatSummary(seriesKeys) {
 var anos = [], series = [], inf = [];
 seriesKeys.forEach(function(s) {
 var up = s.toUpperCase();
 var numMatch = up.match(/\d+/);
 var num = numMatch ? numMatch[0] : "";
 if (!num) return;
 if (up.indexOf("INF") !== -1) inf.push(num);
 else if (up.indexOf("ANO") !== -1) anos.push(num);
 else if (up.indexOf("SERIE") !== -1 || up.indexOf("SÉRIE") !== -1) series.push(num);
 });
 function agrupar(arr, sufixo) {
 if (arr.length === 0) return "";
 var unique = arr.filter(function(v, i, a) { return a.indexOf(v) === i; }).sort();
 if (unique.length === 1) return unique[0] + (sufixo ? " " + sufixo : "");
 var last = unique.pop();
 return unique.join(" ") + " E " + last + (sufixo ? " " + sufixo : "");
 }
 var partes = [];
 var strAno = agrupar(anos, "ANO");
 if(strAno) partes.push(strAno);
 var strSerie = agrupar(series, "SÉRIE");
 if(strSerie) partes.push(strSerie);
 var strInf = agrupar(inf, "");
 if(strInf) partes.push("INF " + strInf);
 return partes.join(", ");
 }

 for (var graf in lotesPorGrafica) {
 var lote = lotesPorGrafica[graf];
 
 var marcasArr = Object.keys(lote.marcas);
 var seriesArr = Object.keys(lote.series);
 var resumoAnos = formatSummary(seriesArr);
 var nomeArquivo = "PLANILHA DE PARA PPG " + graf + " V127 " + marcasArr.join(" E ") + (resumoAnos ? " " + resumoAnos : "");
 
 // Sem restrição de sobrescrita: manda o homônimo antigo para a lixeira antes de copiar.
 removerHomonimos(pastaDest, nomeArquivo);
 var arquivoCopia = DriveApp.getFileById(ID_PLANILHA).makeCopy(nomeArquivo, pastaDest);
 var ssCopia = SpreadsheetApp.openById(arquivoCopia.getId());
 var sheetsCopia = ssCopia.getSheets();
 for (var i = 0; i < sheetsCopia.length; i++) {
 if (abasPermitidas.indexOf(sheetsCopia[i].getName()) === -1) {
 ssCopia.deleteSheet(sheetsCopia[i]);
 }
 }
 
 var nomeArquivoCalculadora = "CALCULADORA FORNECEDOR " + graf + " - LIVROS PRINCIPAIS COL 2027 V127 " + marcasArr.join(" E ") + (resumoAnos ? " " + resumoAnos : "") + ".xlsx";
 
 var metadadosPacote = {
 nomeArquivoFinal: nomeArquivoCalculadora,
 idChamado: "PRINCIPAL-" + hashLote.slice(-6) + "-" + qtdLotes,
 nomeGrafica: graf, 
nomeModeloCalculadora: "CALCULADORA FORNECEDOR - LIVROS PRINCIPAIS COL 2027.xlsx",
 usuarioLogado: emailUser,
 dadosSkus: lote.dadosSkus,
 dadosCalcGeral: lote.dadosCalcGeral,
 tiragemTotal: lote.tiragemTotal,
 configuracoesCalc: payload.camposCalc,
 timestamp: new Date().toISOString()
 };
 
var idJob = gravarNaFila(metadadosPacote); 
idsJobsCriados.push(idJob);
 
var assunto = "[ESTEIRA_CALCULADORA_PRINCIPAL] Lote:" + hashLote + " | Gráfica: " + graf;
 var anexoJson = Utilities.newBlob(JSON.stringify(metadadosPacote), "application/json", "pacote_principal_" + graf + ".json");
 var corpo = "Pacote gerado pela Validação Principal.\n\nGráfica: " + graf + "\nMarcas envolvidas: " + Object.keys(lote.marcas).join(", ") + "\n\nLink do arquivo De-Para processado:\n" + arquivoCopia.getUrl();
 
MailApp.sendEmail(EMAIL_FILA_GMAIL, assunto, corpo, { attachments: [anexoJson] });
 resumoLote.push({
 grafica: graf,
 chamado: "",
 marcas: marcasArr.join(", "),
 series: resumoAnos || seriesArr.join(", "),
 skus: lote.dadosSkus.length,
 tiragem: lote.tiragemTotal,
 arquivo: nomeArquivoCalculadora
 });
 qtdLotes++;
 }

return { sucesso: true, skus: skusFiltro.length, lotes: qtdLotes, hashLote: hashLote, jobs: idsJobsCriados, resumo: resumoLote };
} catch (e) {
 return { erro: "Erro na extração principal: " + e.message };
 }
}
