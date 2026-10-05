// =========================================================================
// ValidaComum.gs — BASE COMPARTILHADA
// Constantes e helpers usados PELOS DOIS fluxos (Principal e Reentradas).
// Nada aqui pode conter regra de negócio específica de um dos dois.
// =========================================================================

var ID_PASTA_CALCULADORA = "1stauNMVLbyFBK409mFKJ62VseIGqIRyT";
var ID_PASTA_DEPARA_PREENCHIDO = "1M5AC0N6-J4Mk6qnJhhyuSxc_yPww1eBJ";
var NOME_PADRAO_DEPARA = "PLANILHA DE PARA PPG";
// ID_PLANILHA_LOG é declarado UMA ÚNICA VEZ em WebApp.gs — não redeclarar aqui.
var EMAIL_FILA_GMAIL = "leonardo.pereira@arcoeducacao.com.br";

// Pasta única onde os DOIS fluxos (Principal e Reentradas) salvam o De-Para gerado.
var ID_PASTA_DEPARA_GERADO = "1DKie1AHE7EFHa6roJY2XxE2ZlFhC_lCy";

/**
 * Manda para a lixeira qualquer arquivo homônimo já existente na pasta.
 * Sem restrição de sobrescrita: a versão nova sempre vence.
 */
function removerHomonimos(pastaDestino, nomeArquivo, idIgnorar) {
 var existentes = pastaDestino.getFilesByName(nomeArquivo);
 while (existentes.hasNext()) {
 var antigo = existentes.next();
 if (!idIgnorar || antigo.getId() !== idIgnorar) antigo.setTrashed(true);
 }
}

/** Renomeia e move um arquivo para a pasta destino, sobrescrevendo homônimos. */
function salvarSobrescrevendo(arquivo, nomeFinal, pastaDestino) {
 removerHomonimos(pastaDestino, nomeFinal, arquivo.getId());
 arquivo.setName(nomeFinal);
 arquivo.moveTo(pastaDestino);
 return arquivo;
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


function testarPermissaoEmail() {
 var emailUser = Session.getActiveUser().getEmail();
 MailApp.sendEmail(emailUser, "Teste de Permissão de E-mail", "Essa mensagem confirma que o escopo de envio de e-mail está autorizado no seu projeto Google Apps Script.");
 Logger.log("E-mail enviado para " + emailUser);
}
