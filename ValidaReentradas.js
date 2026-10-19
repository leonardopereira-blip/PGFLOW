// Entradas antigas de custos de reentradas desativadas.
// Helpers preservados; Validação Principal permanece separada.

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
  throw new Error('O fluxo antigo de custos de reentradas foi desativado. Atualize a página para usar a API do SPG.');
}

function validarFase1Multiplos() {
  throw new Error('O fluxo antigo de custos de reentradas foi desativado. Atualize a página para usar a API do SPG.');
}

function finalizarProcessamentoLoteDepara() {
  throw new Error('O fluxo antigo de custos de reentradas foi desativado. Atualize a página para usar a API do SPG.');
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

function checarStatusProcessamento() {
  throw new Error('O fluxo antigo de custos de reentradas foi desativado. Atualize a página para usar a API do SPG.');
}

function buscarResumoChamados() {
  throw new Error('O fluxo antigo de custos de reentradas foi desativado. Atualize a página para usar a API do SPG.');
}
