// =========================================================================
// 1_WebApp.gs — CONTROLADOR PRINCIPAL DO WEB APP + API DE FILA
// =========================================================================
// ⚠️  ATENÇÃO: ID_PLANILHA_LOG é a única declaração global deste ID em todo
//     o projeto. Os demais arquivos .gs referenciam esta variável diretamente.
// =========================================================================

var ID_PLANILHA_LOG = "1pR9a5-oogD7ZYHxImsE9ABNpHeVyCJaTcIab7_U_bI0";

// ── ROTEADOR PRINCIPAL ──────────────────────────────────────────────────────
// Serve o Web App normalmente OU responde chamadas da API de fila
// (consumida pelo Office Script via Power Automate agendado).
// Publicar com: "Qualquer pessoa na organização" pode acessar.
function doGet(e) {
  // Sem parâmetro "action" → serve a interface HTML normalmente
  if (!e || !e.parameter || !e.parameter.action) {
    return HtmlService.createTemplateFromFile('Index')
        .evaluate()
        .setTitle('Portal de Produção e Custos - ARCO')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Com parâmetro "action" → responde como API JSON para o Office Script
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  switch (e.parameter.action) {

    case "getPending":
      // Office Script chama: ?action=getPending
      // Retorna array de jobs com status PENDING
      output.setContent(JSON.stringify(buscarJobsPendentes()));
      return output;

    case "markDone":
      // Office Script chama: ?action=markDone&id=UUID_DO_JOB
      if (!e.parameter.id) {
        output.setContent(JSON.stringify({ erro: "Parâmetro 'id' ausente." }));
        return output;
      }
      marcarStatusJob(e.parameter.id, "DONE", "");
      output.setContent(JSON.stringify({ ok: true }));
      return output;

    case "markError":
      // Office Script chama: ?action=markError&id=UUID&msg=MENSAGEM
      if (!e.parameter.id) {
        output.setContent(JSON.stringify({ erro: "Parâmetro 'id' ausente." }));
        return output;
      }
      var msgErro = e.parameter.msg ? decodeURIComponent(e.parameter.msg) : "Erro desconhecido";
      marcarStatusJob(e.parameter.id, "ERRO", msgErro);
      output.setContent(JSON.stringify({ ok: true }));
      return output;

    default:
      output.setContent(JSON.stringify({ erro: "Ação inválida: " + e.parameter.action }));
      return output;
  }
}

// Helper para injetar as sub-telas HTML dentro do arquivo Index mestre
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ── FUNÇÕES DE GERENCIAMENTO DA FILA (Queue_Calculadora) ───────────────────

/**
 * Grava um novo job na aba Queue_Calculadora da planilha de log.
 * Chamada pelo ValidaCustos.gs ao finalizar o processamento de um chamado.
 * @param {Object} metadadosPacote - { idChamado, nomeGrafica, nomeModeloCalculadora, dadosSkus }
 * @returns {string} UUID do job gravado
 */
function gravarNaFila(metadadosPacote) {
  var ss = SpreadsheetApp.openById(ID_PLANILHA_LOG);
  var aba = ss.getSheetByName("Queue_Calculadora");

  // Cria a aba e o cabeçalho na primeira execução
  if (!aba) {
    aba = ss.insertSheet("Queue_Calculadora");
    var cabecalho = ["id_job", "idChamado", "nomeGrafica", "nomeModeloCalculadora",
                     "status", "timestamp_criacao", "timestamp_atualizacao", "dados_json"];
    aba.appendRow(cabecalho);
    aba.getRange(1, 1, 1, cabecalho.length).setFontWeight("bold").setBackground("#f3f3f3");
    aba.setFrozenRows(1);
  }

  var idJob = Utilities.getUuid();

  aba.appendRow([
    idJob,
    metadadosPacote.idChamado.toString(),
    metadadosPacote.nomeGrafica.toString(),
    metadadosPacote.nomeModeloCalculadora.toString(),
    "PENDING",
    new Date(),
    "",                                       // timestamp_atualizacao — preenchido ao marcar DONE/ERRO
    JSON.stringify(metadadosPacote.dadosSkus) // matriz serializada (~col H)
  ]);

  Logger.log("Job gravado na fila: " + idJob + " | Chamado: " + metadadosPacote.idChamado);
  return idJob;
}

/**
 * Retorna todos os jobs com status PENDING para o Office Script consumir.
 * @returns {Array<Object>} Lista de jobs pendentes
 */
function buscarJobsPendentes() {
  var ss = SpreadsheetApp.openById(ID_PLANILHA_LOG);
  var aba = ss.getSheetByName("Queue_Calculadora");
  if (!aba || aba.getLastRow() <= 1) return [];

  var dados = aba.getDataRange().getValues();
  var pendentes = [];

  for (var i = 1; i < dados.length; i++) {
    if (dados[i][4].toString() === "PENDING") {
      pendentes.push({
        idJob:              dados[i][0].toString(),
        idChamado:          dados[i][1].toString(),
        nomeGrafica:        dados[i][2].toString(),
        nomeModelo:         dados[i][3].toString(),
        dadosSkusJson:      dados[i][7].toString()  // coluna H (índice 7)
      });
    }
  }
  return pendentes;
}

/**
 * Atualiza o status de um job na fila.
 * @param {string} idJob - UUID do job
 * @param {string} novoStatus - "DONE" | "ERRO"
 * @param {string} mensagem - Detalhes do erro (vazio se DONE)
 */
function marcarStatusJob(idJob, novoStatus, mensagem) {
  var ss = SpreadsheetApp.openById(ID_PLANILHA_LOG);
  var aba = ss.getSheetByName("Queue_Calculadora");
  if (!aba) return;

  var dados = aba.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (dados[i][0].toString() === idJob) {
      var statusFinal = novoStatus === "ERRO" && mensagem
                        ? "ERRO: " + mensagem
                        : novoStatus;
      aba.getRange(i + 1, 5).setValue(statusFinal);       // coluna E (status)
      aba.getRange(i + 1, 7).setValue(new Date());        // coluna G (timestamp_atualizacao)
      return;
    }
  }
  Logger.log("AVISO: Job não encontrado para marcar status: " + idJob);
}
