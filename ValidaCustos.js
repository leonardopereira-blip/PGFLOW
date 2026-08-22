// =========================================================================
// ValidaCustos.gs — VERSÃO OTIMIZADA (ALTA PERFORMANCE E CASCATA INDIVIDUAL)
// =========================================================================
var ID_PASTA_CALCULADORA = "1stauNMVLbyFBK409mFKJ62VseIGqIRyT";
var ID_PASTA_DEPARA_PREENCHIDO = "1M5AC0N6-J4Mk6qnJhhyuSxc_yPww1eBJ";
var NOME_PADRAO_DEPARA = "PLANILHA DE PARA PPG";
var ID_PLANILHA_LOG = "1pR9a5-oogD7ZYHxImsE9ABNpHeVyCJaTcIab7_U_bI0";
var EMAIL_FILA_GMAIL = "leonardo.pereira@arcoeducacao.com.br";
var OPCOES_MANUAL = {
 "TECNOLOGIA": ["DIGITAL", "PLANA", "ROTATIVA"],
 "FORNECEDOR": ["ARCO", "GRÁFICA"],
 "PAPEL": ["ARCO", "GRÁFICA"],
 "ACABAMENTO": ["ARCO", "GRÁFICA"],
 "RELEVO": ["Não", "Sim"],
 "UV LOCALIZADO": ["Não", "Sim"],
 "HOT STAMPING": ["Não", "Sim"],
 "ORELHA": ["Não", "Sim"],
 "CORTE / VINCO": ["Não", "Sim"],
 "SERRILHA": ["Não", "Sim"],
 "SHRINK": ["Não", "Sim"],
 "MALETA": ["Não", "Sim"],
 "PERSONALIZADA": ["Não", "Sim"]
};

function resolverOpcoes(titulo, instrucao) {
 var inst = (instrucao || "").toString().toLowerCase();
 if (inst.indexOf("cm") !== -1 || inst.indexOf("área") !== -1 ||
 inst.indexOf("area") !== -1 || inst.indexOf("mm") !== -1) {
 return null;
 }
 if (inst.indexOf("sim") !== -1 && (inst.indexOf("não") !== -1 || inst.indexOf("nao") !== -1)) {
 return ["Não", "Sim"];
 }
 var up = (titulo || "").toString().toUpperCase();
 for (var chave in OPCOES_MANUAL) {
 if (up.indexOf(chave) !== -1) return OPCOES_MANUAL[chave];
 }
 return null;
}

function serializarValor(v) {
 if (v instanceof Date) return Utilities.formatDate(v, 'America/Sao_Paulo', 'dd/MM/yyyy');
 return v;
}

function obterPastasGraficas() {
 try {
 var ss = SpreadsheetApp.openById(ID_PLANILHA_LOG);
 var aba = ss.getSheetByName("Graficas_Lista");
 if (!aba) return [];
 var dados = aba.getDataRange().getValues();
 var lista = [];
 for (var i = 1; i < dados.length; i++) {
 var nome = dados[i][0] ? dados[i][0].toString().trim() : "";
 var id = dados[i][1] ? dados[i][1].toString().trim() : "";
 if (nome !== "") lista.push({ nome: nome, id: id !== "" ? id : nome });
 }
 return lista;
 } catch (e) { return []; }
}

function obterChamadosLogadosMaster() {
 var ss = SpreadsheetApp.openById(ID_PLANILHA_LOG);
 var abaConv = ss.getSheetByName("PPG_Convertida");
 if (!abaConv) return [];
 var valores = abaConv.getRange(2, 1, abaConv.getLastRow() - 1, 1).getValues();
 var log = [];
 valores.forEach(function(l) { if (l && l[0]) log.push(l[0].toString().trim()); });
 return log;
}

function validarFase1Multiplos(idPastaOuNome, chamadosStr) {
 try {
 var chamadosLog = obterChamadosLogadosMaster();
 var digitados = chamadosStr.split(',').map(function(c) { return c.trim(); })
 .filter(function(c) { return c !== ""; });
 var validos = [], invalidos = [];
 digitados.forEach(function(c) {
 (chamadosLog.indexOf(c) !== -1 ? validos : invalidos).push(c);
 });
 if (validos.length === 0) return { erro: "Nenhum chamado encontrado no log." };
 var pastaGrafica;
 if (idPastaOuNome && idPastaOuNome.length > 20 && idPastaOuNome.indexOf(" ") === -1) {
 pastaGrafica = DriveApp.getFolderById(idPastaOuNome);
 } else {
 var pastaMae = DriveApp.getFolderById(ID_PASTA_CALCULADORA);
 var subpastas = pastaMae.getFoldersByName(idPastaOuNome);
 if (!subpastas.hasNext()) return { erro: "Pasta da gráfica não encontrada." };
 pastaGrafica = subpastas.next();
 }
 var arquivos = pastaGrafica.getFilesByName(NOME_PADRAO_DEPARA);
 if (!arquivos.hasNext()) return { erro: "'" + NOME_PADRAO_DEPARA + "' não encontrado." };
 var ssTemplate = SpreadsheetApp.openById(arquivos.next().getId());
 var abaNc = ssTemplate.getSheetByName("NC - PPG Especificações");
 if (!abaNc) return { erro: "Aba 'NC - PPG Especificações' não existe no modelo." };
 var maxCol = abaNc.getLastColumn();
 var linhaTipos = abaNc.getRange(5, 1, 1, maxCol).getValues()[0];
 var linhaTitulos = abaNc.getRange(3, 1, 1, maxCol).getValues()[0];
 var linhaInstrucoes = abaNc.getRange(4, 1, 1, maxCol).getValues()[0];
 var colunasManuais = [];
 
for (var col = 1; col <= maxCol; col++) {
 var tipo = linhaTipos[col - 1] ? linhaTipos[col - 1].toString().trim() : "";
 if (tipo.toLowerCase() === "manual") {
 var titulo = linhaTitulos[col - 1] ? linhaTitulos[col - 1].toString().trim() : "Coluna " + col;
 var instrucao = linhaInstrucoes[col - 1] ? linhaInstrucoes[col - 1].toString().trim() : "Preenchimento Requerido";
 colunasManuais.push({
 index: col.toString(),
 titulo: titulo,
 instrucao: instrucao,
 opcoes: resolverOpcoes(titulo, instrucao)
 });
 }
 }
 return { bundle: true, sucesso: true, validos: validos, invalidos: invalidos, colunasManuais: colunasManuais };
 } catch (e) {
 return { erro: "Erro na Validação: " + e.message };
 }
}

function finalizarProcessamentoLoteDepara(idPastaOuNome, chamadosValidos, opcaoVersao, versoesManuais, respostasManuais, fornecimentoUI) {
 try {
 var pastaDestino = DriveApp.getFolderById(ID_PASTA_DEPARA_PREENCHIDO);
 var pastaGrafica;
 if (idPastaOuNome.length > 20 && idPastaOuNome.indexOf(" ") === -1) {
 pastaGrafica = DriveApp.getFolderById(idPastaOuNome);
 } else {
 pastaGrafica = DriveApp.getFolderById(ID_PASTA_CALCULADORA).getFoldersByName(idPastaOuNome).next();
 }
 var nomeGrafica = pastaGrafica.getName();
 var arquivoOrigDepara = pastaGrafica.getFilesByName(NOME_PADRAO_DEPARA).next();
 var relatorio = [];
 var ssOrigem = SpreadsheetApp.openById(ID_PLANILHA_LOG);
 var emailUsuario = Session.getActiveUser().getEmail();
 var hashLoteAntiThreading = new Date().getTime().toString();
 var idsJobsCriados = [];
 var abaEspecOrigem = ssOrigem.getSheetByName("Especificação Técnica");
 var dadosEspecCache = abaEspecOrigem ? abaEspecOrigem.getDataRange().getValues() : [];
 
var abaArvoreOrigem = ssOrigem.getSheetByName("Árvore de produto");
 var dadosArvoreCache = abaArvoreOrigem ? abaArvoreOrigem.getDataRange().getValues() : [];
 
var abaTiragemOrigem = ssOrigem.getSheetByName("Tiragem");
 var dadosTiragemCache = abaTiragemOrigem ? abaTiragemOrigem.getDataRange().getValues() : [];
 
chamadosValidos.forEach(function(idChamado) {
 try {
 var arquivoTrabalho = arquivoOrigDepara.makeCopy("TEMP_WORKER_" + idChamado, pastaGrafica);
 var ssTemp = SpreadsheetApp.openById(arquivoTrabalho.getId());
 ssTemp.setSpreadsheetLocale('pt_BR');
 ssTemp.setSpreadsheetTimeZone('America/Sao_Paulo');
 var abaTiragemTemp = ssTemp.getSheetByName("PPG - Tiragem");
 var formulaT_Original = abaTiragemTemp ? abaTiragemTemp.getRange("T2").getFormula() : "";
 var numEspec = processarMapeamentoEColagem(dadosEspecCache, "PPG - Especificações", idChamado, "Versão da PPG", opcaoVersao, versoesManuais, ssTemp);
 var numArvore = processarMapeamentoEColagem(dadosArvoreCache, "PPG - Arvore de Produto", idChamado, "Versão da PPG", opcaoVersao, versoesManuais, ssTemp);
 var numTiragem = processarMapeamentoEColagem(dadosTiragemCache, "PPG - Tiragem", idChamado, "Versão da PPG", opcaoVersao, versoesManuais, ssTemp);
 var numLinhas = Math.max(numEspec, numArvore, numTiragem);
 
if (numLinhas === 0) {
 arquivoTrabalho.setTrashed(true);
 relatorio.push(idChamado + " ⚠️ Vazio (Ignorado)");
 return;
 }
 if (abaTiragemTemp && formulaT_Original !== "") {
 abaTiragemTemp.getRange(2, 20, numLinhas, 1).setFormula(formulaT_Original);
 }
 var abaNc = ssTemp.getSheetByName("NC - PPG Especificações");
 var limNec = 6 + numLinhas - 1;
 if (abaNc.getMaxRows() < limNec) abaNc.insertRowsAfter(abaNc.getMaxRows(), (limNec - abaNc.getMaxRows()) + 5);
 var maxCol = abaNc.getLastColumn();
 var linhaTipos = abaNc.getRange(5, 1, 1, maxCol).getValues()[0];
 var linhaInstrucoesFill = abaNc.getRange(4, 1, 1, maxCol).getValues()[0];
 
for (var col = 1; col <= maxCol; col++) {
 var tipo = linhaTipos[col - 1] ? linhaTipos[col - 1].toString().trim() : "";
 if (tipo.toLowerCase() === "manual") {
 var valorResp = "";
 respostasManuais.forEach(function(rm) {
 if (rm.index.toString() === col.toString()) valorResp = rm.valor.trim();
 });
 var instCol = linhaInstrucoesFill[col - 1] ? linhaInstrucoesFill[col - 1].toString().toLowerCase() : "";
 var ehNumerico = (instCol.indexOf("cm") !== -1 || instCol.indexOf("área") !== -1 ||
 instCol.indexOf("area") !== -1 || instCol.indexOf("mm") !== -1);
 var padrao = ehNumerico ? "0" : "Não";
 abaNc.getRange(6, col, numLinhas, 1).setValue(valorResp === "" ? padrao : valorResp);
 } else {
 var c0 = abaNc.getRange(6, col);
 if (c0.getFormula() !== "" || c0.getValue() !== "") {
 c0.copyTo(abaNc.getRange(6, col, numLinhas), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
 }
 }
 }
 SpreadsheetApp.flush();
 Utilities.sleep(1500);
 arquivoTrabalho.moveTo(pastaDestino);
 arquivoTrabalho.setName("DEPARA_CHAMADO_" + idChamado);
 
var matrizBruta = abaNc.getRange(6, 3, numLinhas, 79).getValues();
 var matrizFiltrada = [];
 for (var mIdx = 0; mIdx < matrizBruta.length; mIdx++) {
 var rowCopy = matrizBruta[mIdx].slice();
 for (var cIdx = 0; cIdx < rowCopy.length; cIdx++) {
 if (typeof rowCopy[cIdx] === 'string') {
 rowCopy[cIdx] = rowCopy[cIdx].replace(/\s+/g, ' ').trim();
 }
 }
 var skuClean = rowCopy[0] ? rowCopy[0].toString() : "";
 if (skuClean !== "" && skuClean !== "-") {
 matrizFiltrada.push(rowCopy);
 }
 }
 if (matrizFiltrada.length === 0) {
 arquivoTrabalho.setTrashed(true);
 relatorio.push(idChamado + " ⚠️ Ignorado (Sem SKUs válidos)");
 return;
 }
 var matrizSerializada = matrizFiltrada.map(function(linha) {
 return linha.map(serializarValor);
 });
 
var abaCalcGeral = ssTemp.getSheetByName("NC - CALC Geral");
 var dadosCalcGeralMatrix = [];
 if (abaCalcGeral) {
 var lastRowG = abaCalcGeral.getLastRow();
 if (lastRowG < 10) lastRowG = 100;
 
var rangeGeral = abaCalcGeral.getRange(1, 1, lastRowG, 21).getValues();
 for (var idxG = 3; idxG < rangeGeral.length; idxG++) {
 var skuVal = rangeGeral[idxG][11] ? rangeGeral[idxG][11].toString().trim() : "";
 if (skuVal.indexOf("-") === 0) skuVal = skuVal.substring(1).trim();
 
if (skuVal !== "" && skuVal !== "-" && skuVal.toLowerCase().indexOf("total") === -1 && skuVal.toLowerCase().indexOf("sku") === -1) {
 dadosCalcGeralMatrix.push([
 skuVal,
 rangeGeral[idxG][15] !== undefined ? rangeGeral[idxG][15] : "",
 rangeGeral[idxG][16] !== undefined ? rangeGeral[idxG][16] : "",
 rangeGeral[idxG][17] !== undefined ? rangeGeral[idxG][17] : "",
 rangeGeral[idxG][18] !== undefined ? rangeGeral[idxG][18] : "",
 "N/A",
 rangeGeral[idxG][20] !== undefined ? rangeGeral[idxG][20] : ""
 ]);
 }
 }
 }
 
var tiragemTotal = 0;
 for (var ti = 0; ti < dadosCalcGeralMatrix.length; ti++) {
 var s = parseFloat(dadosCalcGeralMatrix[ti][4]);
 if (!isNaN(s)) tiragemTotal += s;
 }
 
var fUI = fornecimentoUI || {};
 var fornecimentoPacote = {
 tecnologia: (fUI.tecnologia || "DIGITAL").toString().toUpperCase(),
 fornecedor: fUI.fornecedor || "GRÁFICA",
 elementos: fUI.elementos || {}
 };
 
var metadadosPacote = {
 idChamado: idChamado.toString(),
 nomeGrafica: nomeGrafica,
 nomeModeloCalculadora: "CALCULADORA FORNECEDOR " + nomeGrafica.toUpperCase() + " - REENTRADA 2026 - NOVA VERSÃO.xlsx",
 usuarioLogado: emailUsuario,
 dadosSkus: matrizSerializada,
 dadosCalcGeral: dadosCalcGeralMatrix,
 tiragemTotal: tiragemTotal,
 fornecimento: fornecimentoPacote,
 timestamp: new Date().toISOString()
 };
 
var idJobCriado = gravarNaFila(metadadosPacote);
 idsJobsCriados.push(idJobCriado);
 var assunto = "[ESTEIRA_CALCULADORA] Chamado " + idChamado + " " + nomeGrafica + " | Lote:" + hashLoteAntiThreading;
 var anexoJson = Utilities.newBlob(JSON.stringify(metadadosPacote), "application/json", "pacote_" + idChamado + ".json");
 MailApp.sendEmail(EMAIL_FILA_GMAIL, assunto, "Pacote de dados em anexo (JSON).", { attachments: [anexoJson] });
 relatorio.push(idChamado + " ✅ Despachado");
 } catch (errLoop) {
 relatorio.push(idChamado + " ❌ Erro: " + errLoop.message);
 }
 });
 return { mensagem: "Lote processado!", hashLote: hashLoteAntiThreading, jobs: idsJobsCriados };
 } catch (e) {
 return { erro: "Erro Crítico: " + e.message };
 }
}

function processarMapeamentoEColagem(dadosOrigem, nomeAbaDestino, idChamado, nomeColVersao, opcaoVersao, versoesManuais, ssTemp) {
 var abaDestino = ssTemp.getSheetByName(nomeAbaDestino);
 if (!dadosOrigem || dadosOrigem.length <= 1 || !abaDestino) return 0;
 
var maxColDest = abaDestino.getLastColumn();
 if (maxColDest === 0) return 0;
 
var headersOrigem = dadosOrigem[0].map(function(h) { return h ? h.toString().trim() : ""; });
 var headersDestino = abaDestino.getRange(1, 1, 1, maxColDest).getValues()[0].map(function(h) { return h ? h.toString().trim() : ""; });
 var idxVersao = headersOrigem.indexOf(nomeColVersao);
 if (idxVersao === -1) return 0;
 
var versoesAlvo = [];
 if (opcaoVersao === "ultima") {
 var maxV = -1;
 for (var i = 1; i < dadosOrigem.length; i++) {
 if (dadosOrigem[i][0] && dadosOrigem[i][0].toString().trim() === idChamado) {
 var v = parseFloat(dadosOrigem[i][idxVersao]);
 if (!isNaN(v) && v > maxV) maxV = v;
 }
 }
 if (maxV !== -1) versoesAlvo.push(maxV.toString());
 } else {
 if (versoesManuais) versoesAlvo = versoesManuais.split(',').map(function(s) { return s.trim(); });
 }
 
var mapa = [];
 for (var c = 0; c < headersDestino.length; c++) {
 mapa.push(headersDestino[c] !== "" ? headersOrigem.indexOf(headersDestino[c]) : -1);
 }
 var matrizFinal = [];
 for (var i = 1; i < dadosOrigem.length; i++) {
 if (dadosOrigem[i][0] && dadosOrigem[i][0].toString().trim() === idChamado) {
 var vLinha = dadosOrigem[i][idxVersao] ? dadosOrigem[i][idxVersao].toString().trim() : "";
 if (versoesAlvo.indexOf(vLinha) !== -1) {
 var novaLinha = [];
 for (var m = 0; m < mapa.length; m++) novaLinha.push(mapa[m] !== -1 ? dadosOrigem[i][mapa[m]] : "");
 matrizFinal.push(novaLinha);
 }
 }
 }
 if (matrizFinal.length > 0) {
 var lastRow = abaDestino.getLastRow();
 if (lastRow > 1) {
 for (var m = 0; m < mapa.length; m++) {
 if (mapa[m] !== -1) abaDestino.getRange(2, m + 1, lastRow - 1, 1).clearContent();
 }
 }
 abaDestino.getRange(2, 1, matrizFinal.length, matrizFinal[0].length).setValues(matrizFinal);
 }
 return matrizFinal.length;
}

function checarStatusProcessamento(jobs, hashLote) {
 try {
 var emailNaCaixa = false;
 if (hashLote) {
 var threads = GmailApp.search('in:inbox subject:"Lote:' + hashLote + '"', 0, 1);
 emailNaCaixa = (threads.length > 0);
 }
 var queueStatus = "PENDING";
 if (jobs && jobs.length > 0) {
 var ss = SpreadsheetApp.openById(ID_PLANILHA_LOG);
 var aba = ss.getSheetByName("Queue_Calculadora");
 if (aba) {
 var dados = aba.getDataRange().getValues();
 for (var i = dados.length - 1; i > 0; i--) {
 if (jobs.indexOf(dados[i][0].toString()) !== -1) {
 var st = dados[i][4].toString();
 if (st.indexOf("ERRO") !== -1) return { emailNaCaixa: emailNaCaixa, queueStatus: st };
 if (st === "DONE") queueStatus = "DONE";
 break;
 }
 }
 }
 }
 return { emailNaCaixa: emailNaCaixa, queueStatus: queueStatus };
 } catch (e) {
 return { erro: e.message };
 }
}

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

// =========================================================================
// FUNÇÕES DA VALIDAÇÃO PRINCIPAL E DASHBOARD
// =========================================================================

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
 var idx = { sku: -1, grafica: -1 };
 
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
 }

 var result = [];
 
 for (var r = 1; r < dados.length; r++) {
 var rowSku = idx.sku !== -1 && dados[r][idx.sku] ? dados[r][idx.sku].toString().trim().toLowerCase() : "";
 var rowGrafica = idx.grafica !== -1 && dados[r][idx.grafica] ? dados[r][idx.grafica].toString().trim().toLowerCase() : "";
 
 // Checa o SKU
 var matchSku = skusFiltro.some(function(s) { 
 return s.toString().trim().toLowerCase() === rowSku; 
 });
 
 // Checa a Gráfica (Filtro para evitar puxar FormaCerta quando for LogPrint, por ex)
 var matchGrafica = true; 
 if (graficasFiltro.length > 0 && idx.grafica !== -1 && rowGrafica !== "") {
 matchGrafica = graficasFiltro.some(function(g) {
 return rowGrafica.indexOf(g.toString().trim().toLowerCase()) !== -1 || g.toString().trim().toLowerCase().indexOf(rowGrafica) !== -1;
 });
 }

 if (matchSku && matchGrafica) {
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

for (var graf in lotesPorGrafica) {
 var lote = lotesPorGrafica[graf];
 
var metadadosPacote = {
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
 var corpo = "Pacote gerado pela Validação Principal.\n\nGráfica: " + graf + "\nMarcas envolvidas: " + Object.keys(lote.marcas).join(", ");
 
MailApp.sendEmail(EMAIL_FILA_GMAIL, assunto, corpo, { attachments: [anexoJson] });
 qtdLotes++;
 }
 
return { sucesso: true, skus: skusFiltro.length, lotes: qtdLotes, hashLote: hashLote, jobs: idsJobsCriados }; 
} catch (e) {
 return { erro: "Erro na extração principal: " + e.message };
 }
}

function testarPermissaoEmail() {
 var emailUser = Session.getActiveUser().getEmail();
 MailApp.sendEmail(emailUser, "Teste de Permissão de E-mail", "Essa mensagem confirma que o escopo de envio de e-mail está autorizado no seu projeto Google Apps Script.");
 Logger.log("E-mail enviado para " + emailUser);
}