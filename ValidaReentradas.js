// =========================================================================
// ValidaReentradas.gs — FLUXO "VALIDAÇÃO DE CUSTOS REENTRADAS"
// UI: ValidaInterface.html
// Entrada: chamados PPG + pasta da gráfica -> cópia TEMP_WORKER do De-Para
// Saída:   e-mail [ESTEIRA_CALCULADORA] + job na Queue_Calculadora
// Depende de: ValidaComum.gs (constantes/helpers) e WebApp.gs (gravarNaFila)
// NÃO chamar nada daqui a partir da Validação Principal.
// =========================================================================

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

function finalizarProcessamentoLoteDepara(idPastaOuNome, chamadosValidos, opcaoVersao, versoesManuais, respostasManuais, fornecimentoUI, hashLoteUI) {
 try {
 
 var pastaGrafica;
 if (idPastaOuNome.length > 20 && idPastaOuNome.indexOf(" ") === -1) {
 pastaGrafica = DriveApp.getFolderById(idPastaOuNome);
 } else {
 pastaGrafica = DriveApp.getFolderById(ID_PASTA_CALCULADORA).getFoldersByName(idPastaOuNome).next();
 }
 var nomeGrafica = pastaGrafica.getName();
 
 // Salva o De-Para na MESMA pasta usada pela Validação Principal, sobrescrevendo homônimos.
 var pastaDestino = DriveApp.getFolderById(ID_PASTA_DEPARA_GERADO);
 var arquivoOrigDepara = pastaGrafica.getFilesByName(NOME_PADRAO_DEPARA).next();
 var relatorio = [];
 var resumoLote = [];
 var ssOrigem = SpreadsheetApp.openById(ID_PLANILHA_LOG);
 var emailUsuario = Session.getActiveUser().getEmail();
 // Usa o hash gerado no frontend. Assim, se a conexão cair (HTTP 0), a tela ainda
 // sabe qual lote procurar no Gmail. Só gera um novo se a UI não mandar nada.
 var hashLoteAntiThreading = hashLoteUI ? hashLoteUI.toString() : new Date().getTime().toString();
 var idsJobsCriados = [];
 var abaEspecOrigem = ssOrigem.getSheetByName("Especificação Técnica");
 var dadosEspecCache = abaEspecOrigem ? abaEspecOrigem.getDataRange().getValues() : [];
 
var abaArvoreOrigem = ssOrigem.getSheetByName("Árvore de produto");
 var dadosArvoreCache = abaArvoreOrigem ? abaArvoreOrigem.getDataRange().getValues() : [];
 
var abaTiragemOrigem = ssOrigem.getSheetByName("Tiragem");
 var dadosTiragemCache = abaTiragemOrigem ? abaTiragemOrigem.getDataRange().getValues() : [];

 // Índice chamado -> marcas/séries, para montar o resumo do que foi gerado.
 var infoPorChamado = {};
 if (dadosEspecCache.length > 1) {
 var hdrEspec = dadosEspecCache[0];
 var iMarca = hdrEspec.indexOf("2. INF_Marca");
 var iSerie = hdrEspec.indexOf("2. INF_Série");
 for (var iE = 1; iE < dadosEspecCache.length; iE++) {
 var chE = dadosEspecCache[iE][0] ? dadosEspecCache[iE][0].toString().trim() : "";
 if (!chE) continue;
 if (!infoPorChamado[chE]) infoPorChamado[chE] = { marcas: {}, series: {} };
 if (iMarca !== -1 && dadosEspecCache[iE][iMarca]) {
 infoPorChamado[chE].marcas[dadosEspecCache[iE][iMarca].toString().trim()] = true;
 }
 if (iSerie !== -1 && dadosEspecCache[iE][iSerie]) {
 infoPorChamado[chE].series[dadosEspecCache[iE][iSerie].toString().trim()] = true;
 }
 }
 }

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
 var infoCh = infoPorChamado[idChamado] || { marcas: {}, series: {} };
 var marcasCh = Object.keys(infoCh.marcas);
 var seriesCh = Object.keys(infoCh.series);
 var nomeDepara = "PLANILHA DE PARA PPG " + nomeGrafica + " - CHAMADO " + idChamado +
 (marcasCh.length ? " " + marcasCh.join(" E ") : "");
 salvarSobrescrevendo(arquivoTrabalho, nomeDepara, pastaDestino);

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
 
var nomeArquivoCalculadora = "CALCULADORA FORNECEDOR " + nomeGrafica.toUpperCase() +
 " - REENTRADA 2026 - CHAMADO " + idChamado +
 (marcasCh.length ? " " + marcasCh.join(" E ") : "") + ".xlsx";

var metadadosPacote = {
 nomeArquivoFinal: nomeArquivoCalculadora,
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
 resumoLote.push({
 grafica: nomeGrafica,
 chamado: idChamado.toString(),
 marcas: marcasCh.join(", "),
 series: seriesCh.join(", "),
 skus: matrizFiltrada.length,
 tiragem: tiragemTotal,
 arquivo: nomeArquivoCalculadora
 });
 } catch (errLoop) {
 relatorio.push(idChamado + " ❌ Erro: " + errLoop.message);
 }
 });
 return { mensagem: "Lote processado!", hashLote: hashLoteAntiThreading, jobs: idsJobsCriados, resumo: resumoLote };
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

function buscarResumoChamados(chamadosStr) {
  try {
    var ids = (chamadosStr || "").toString().split(',').map(function(c) { return c.trim(); }).filter(function(c) { return c !== ""; });
    if (ids.length === 0) return [];

    var ss = SpreadsheetApp.openById(ID_PLANILHA_LOG);
    var abaEspec = ss.getSheetByName("Especificação Técnica");
    var abaTiragem = ss.getSheetByName("Tiragem");

    var resultados = [];

    ids.forEach(function(idChamado) {
      var skusUnicos = {};
      var marcasSet = {};
      var seriesSet = {};
      var papeisCapaSet = {};
      var papeisMioloSet = {};
      var cdsSet = {};
      var tiragemTotal = 0;

      if (abaEspec) {
        var dadosEspec = abaEspec.getDataRange().getValues();
        if (dadosEspec.length > 1) {
          var headersEspec = dadosEspec[0];
          var idxSkuEspec = headersEspec.indexOf("1. ID_Código SKU");
          var idxMarcaEspec = headersEspec.indexOf("2. INF_Marca");
          var idxSerieEspec = headersEspec.indexOf("2. INF_Série");
          var idxCapaEspec = headersEspec.indexOf("4. CAP_Papel da capa");
          var idxMioloEspec = headersEspec.indexOf("5. MIO_PRIN_Papel");

          for (var i = 1; i < dadosEspec.length; i++) {
            var rowId = dadosEspec[i][0] ? dadosEspec[i][0].toString().trim() : "";
            if (rowId === idChamado) {
              if (idxSkuEspec !== -1 && dadosEspec[i][idxSkuEspec]) {
                var sku = dadosEspec[i][idxSkuEspec].toString().trim();
                if (sku) skusUnicos[sku] = true;
              }
              if (idxMarcaEspec !== -1 && dadosEspec[i][idxMarcaEspec]) {
                marcasSet[dadosEspec[i][idxMarcaEspec].toString().trim()] = true;
              }
              if (idxSerieEspec !== -1 && dadosEspec[i][idxSerieEspec]) {
                seriesSet[dadosEspec[i][idxSerieEspec].toString().trim()] = true;
              }
              if (idxCapaEspec !== -1 && dadosEspec[i][idxCapaEspec]) {
                papeisCapaSet[dadosEspec[i][idxCapaEspec].toString().trim()] = true;
              }
              if (idxMioloEspec !== -1 && dadosEspec[i][idxMioloEspec]) {
                papeisMioloSet[dadosEspec[i][idxMioloEspec].toString().trim()] = true;
              }
            }
          }
        }
      }

      if (abaTiragem) {
        var dadosTiragem = abaTiragem.getDataRange().getValues();
        if (dadosTiragem.length > 1) {
          var headersTiragem = dadosTiragem[0];
          var idxCdTiragem = headersTiragem.indexOf("15. DES_CD envio");
          var idxTiragemVal = -1;
          for (var h = 0; h < headersTiragem.length; h++) {
            if (headersTiragem[h].toString().toLowerCase().indexOf("tir_sku") !== -1) {
              idxTiragemVal = h;
              break;
            }
          }

          for (var j = 1; j < dadosTiragem.length; j++) {
            var rowIdT = dadosTiragem[j][0] ? dadosTiragem[j][0].toString().trim() : "";
            if (rowIdT === idChamado) {
              if (idxCdTiragem !== -1 && dadosTiragem[j][idxCdTiragem]) {
                cdsSet[dadosTiragem[j][idxCdTiragem].toString().trim()] = true;
              }
              if (idxTiragemVal !== -1) {
                var valT = parseFloat(dadosTiragem[j][idxTiragemVal]);
                if (!isNaN(valT)) tiragemTotal += valT;
              }
            }
          }
        }
      }

      resultados.push({
        chamado: idChamado,
        totalSkus: Object.keys(skusUnicos).length || 1,
        tiragemTotal: tiragemTotal,
        marcas: Object.keys(marcasSet).join(", "),
        series: Object.keys(seriesSet).join(", "),
        papeisCapa: Object.keys(papeisCapaSet).join(", "),
        papeisMiolo: Object.keys(papeisMioloSet).join(", "),
        cdsDestino: Object.keys(cdsSet).join(", ")
      });
    });

    return resultados;
  } catch (e) {
    return [];
  }
}
