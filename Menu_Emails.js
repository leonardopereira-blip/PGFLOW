// =========================================================================
// 3_MenuEmail.gs — MENU DA PLANILHA + ENVIO DE E-MAIL PARA GRÁFICA
// =========================================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('✉️ Envio de Formalização')
    .addItem('Enviar E-mail para Gráfica', 'abrirTelaEmail')
    .addToUi();
  
  ui.createMenu('⚙️ Automações')
    .addItem('Validar Custo (Converter Planilhas)', 'abrirMenuConversao')
    .addToUi(); 
}

function abrirTelaEmail() {
  var html = HtmlService.createHtmlOutputFromFile('TelaEmail')
      .setWidth(550)
      .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Envio de Arquivo para Production');
}

function abrirMenuConversao() {
  var html = HtmlService.createHtmlOutputFromFile('ValidaInterface')
      .setWidth(600)
      .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Validar Custos (De-Para)');
}

function obterDadosGraficas() {
  var aba    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Emails_Envio");
  var dados  = aba.getDataRange().getValues();
  var graficas = {};
  
  for (var i = 1; i < dados.length; i++) {
    var nome  = dados[i][0].toString().trim();
    var email = dados[i][1].toString().trim();
    if (nome !== "" && email !== "") {
      if (!graficas[nome]) graficas[nome] = [];
      graficas[nome].push(email);
    }
  }
  return graficas;
}

// Busca a Data, blinda contra Fuso Horário, subtrai 2 dias, trata finais de semana/sextas e captura a coluna F
function obterDadosChamado(idChamado, opcaoVersao, numVersao) {
  idChamado = idChamado.toString().trim();
  var ssLog = SpreadsheetApp.getActiveSpreadsheet();
  
  var chamadoEncontrado = false;
  var dataFormatada = "DD/MM/YYYY";
  var valorColunaF = "";
  
  try {
    // Busca na aba "Solicita_Prod" da própria planilha ativa onde o script está a correr
    var abaFonte = ssLog.getSheetByName("Solicita_Prod");
  
    if(abaFonte) {
      // Pegamos os valores EXATAMENTE como aparecem no ecrã (display values) para neutralizar bugs de fuso horário
      var dadosFonteDisplay = abaFonte.getDataRange().getDisplayValues();
      var dadosFonteValores = abaFonte.getDataRange().getValues();
  
      // Lê de baixo para cima para obter o envio mais recente do mesmo chamado
      for (var i = dadosFonteDisplay.length - 1; i >= 1; i--) {
        if (dadosFonteValores[i][0].toString().trim() === idChamado) {
          chamadoEncontrado = true;
          valorColunaF = dadosFonteDisplay[i][5].toString().trim(); // Captura a Coluna F (Índice 5)
          var dataStr = dadosFonteDisplay[i][15].toString().trim(); // Coluna P (Índice 15) como texto
  
          // Verifica se conseguiu extrair a data no formato DD/MM/AAAA
          if (dataStr && dataStr.indexOf("/") !== -1 && dataStr.indexOf("#REF") === -1) {
            var partes = dataStr.split('/');
            if (partes.length >= 3) {
              var dia = parseInt(partes[0], 10);
              var mes = parseInt(partes[1], 10) - 1; // Em Javascript, os meses começam do 0 (Jan = 0)
              var ano = parseInt(partes[2].substring(0,4), 10); // Garante a captura dos 4 dígitos do ano
  
              // Cria a data fixada ao MEIO-DIA para impedir deslocamentos por fuso horário
              var dataCalculada = new Date(ano, mes, dia, 12, 0, 0);
              dataCalculada.setDate(dataCalculada.getDate() - 2); // Subtrai exatamente 2 dias
  
              // --- REGRA ROBUSTA DE COLETA (Não permite Sexta, Sábado nem Domingo) ---
              var diaSemana = dataCalculada.getDay();
              if (diaSemana === 5) {         // Se cair numa Sexta-feira, recua 1 dia para Quinta
                dataCalculada.setDate(dataCalculada.getDate() - 1);
              } else if (diaSemana === 6) {  // Se cair num Sábado, recua 2 dias para Quinta
                dataCalculada.setDate(dataCalculada.getDate() - 2);
              } else if (diaSemana === 0) {  // Se cair num Domingo, recua 3 dias para Quinta
                dataCalculada.setDate(dataCalculada.getDate() - 3);
              }
              // ------------------------------------------------------------------------
  
              var diaFinal = dataCalculada.getDate().toString().padStart(2, '0');
              var mesFinal = (dataCalculada.getMonth() + 1).toString().padStart(2, '0');
              var anoFinal = dataCalculada.getFullYear();
  
              dataFormatada = diaFinal + "/" + mesFinal + "/" + anoFinal;
            } else {
              dataFormatada = dataStr;
            }
          } else {
            dataFormatada = dataStr; // Fallback: Se não for data com barra, devolve o que estiver lá
          }
          break;
        }
      }
    } else {
      return { erro: "Aba 'Solicita_Prod' não encontrada na planilha atual." };
    }
  } catch(e) {
    return { erro: "ERRO AO LER DADOS DO CHAMADO: " + e.message };
  }
  
  if (!chamadoEncontrado) {
    return { erro: "Chamado " + idChamado + " não encontrado na aba Solicita_Prod. Verifique se o ID está correto." };
  }
  
  var abaUtil  = ssLog.getSheetByName("PPG_Utilizada");
  var dadosUtil = abaUtil.getDataRange().getValues();
  var excelsEncontrados = [];
  
  for (var i = 1; i < dadosUtil.length; i++) {
    if (dadosUtil[i][0].toString().trim() === idChamado) {
      excelsEncontrados.push(dadosUtil[i][4]);  // coluna E — ID da planilha copiada
    }
  }
  
  if (excelsEncontrados.length === 0) {
    return { erro: "Arquivo Excel não encontrado para este chamado. O robô já converteu esse arquivo?" };
  }
  
  var idExcel = null;
  if (opcaoVersao === 'manual') {
    var index = parseInt(numVersao) - 1;
    if (index >= 0 && index < excelsEncontrados.length) {
      idExcel = excelsEncontrados[index];
    } else {
      return { erro: "Versão " + numVersao + " não encontrada. Este chamado possui " + excelsEncontrados.length + " versão(ões)." };
    }
  } else {
    idExcel = excelsEncontrados[excelsEncontrados.length - 1];
  }
  
  if (!idExcel || idExcel.toString().length < 15) {
    return { erro: "O ID do arquivo Excel encontrado na planilha parece ser inválido." };
  }
  
  var arqDrive;
  try {
    arqDrive = DriveApp.getFileById(idExcel.toString().trim());
  } catch(e) {
    return { erro: "ERRO NO EXCEL: " + e.message + " | O usuário não tem permissão para acessar o ID gerado: " + idExcel };
  }
  
  return {
    sucesso:    true,
    dataColeta: dataFormatada,
    idExcel:    idExcel,
    urlExcel:   arqDrive.getUrl(),
    valorF:     valorColunaF
  };
}

// Dispara o e-mail renomeando o anexo utilizando o ALIAS validado
function dispararEmailFinal(dados) {
  try {
    var arquivoExcel = DriveApp.getFileById(dados.idArquivo);
    var assunto      = dados.assunto;
  
    var nomePadronizado = "Chamado " + dados.idChamado + ".xlsx";
    var anexoBlob = arquivoExcel.getBlob().setName(nomePadronizado);
  
    // Disparo oficial em Lote: O robô vai rodar este loop e enviar um email NOVO para cada caixinha que estava na tela
    for (var i = 0; i < dados.arrayDestinos.length; i++) {
      var destinatarioAtual = dados.arrayDestinos[i]; // Emails da caixa da vez
      
      GmailApp.sendEmail(destinatarioAtual, assunto, dados.corpo, {
        attachments: [anexoBlob],
        from: "joao.alberti@arcoeducacao.com.br", // Remetente oficial (Alias)
        name: "Joao Victor Caporal Alberti"       // Nome que aparecerá para os destinatários
      });
    }
  
    return "Enviado com sucesso!";
  } catch (e) {
    return "ERRO AO ACESSAR/ENVIAR ARQUIVO: " + e.message + " | ID do arquivo bloqueado: " + (dados.idArquivo || "N/A");
  }
}