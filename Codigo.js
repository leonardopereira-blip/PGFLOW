function verificarECopiarNovosLinks() {
  // ================= CONFIGURAÇÕES =================
  var ID_PLANILHA_FONTE = "1_wdlykJhe5poscKEde1vVCGmk-zrEiILwJTE_R-0HCM"; 
  var NOME_ABA_FONTE = "Alocação";
  var ID_PASTA_DESTINO = "1M_GnLz8eI1a6NuMdUq7rp3OIdHYaQpXO"; 
  var ID_SUBPASTA_SHEETS = "16FQCAZOflX0Ge1a24gk46SzyE8ZCR5xA"; 
  var LINHA_INICIAL_FONTE = 1420; 
  
  var ID_PLANILHA_LOG = "1pR9a5-oogD7ZYHxImsE9ABNpHeVyCJaTcIab7_U_bI0";
  var NOME_ABA_LOG = "PPG_Utilizada";
  var NOME_ABA_LOG_CONV = "PPG_Convertida"; 
  
  var ABA_CONSOL_1 = "Especificação Técnica";
  var ABA_CONSOL_2 = "Árvore de produto";
  var ABA_CONSOL_3 = "Tiragem";
  
  // CABEÇALHOS DESEJADOS
  var CABECALHOS_ESPEC = ["1. ID_Código SKU", "1. ID_Descrição", "1. ID_ISBN", "1. ID_Código SKU anterior", "2. INF_Marca", "2. INF_Grupo da marca", "2. INF_Segmento", "2. INF_Série", "2. INF_Volume", "2. INF_Envio", "2. INF_Frequência", "2. INF_Usuário", "2. INF_Assunto/ Disciplina/ Área", "2. INF_Tipo de material", "2. INF_Classificação do produto", "2. INF_Cliente personalizado", "3. DIM_B (Aberto)(mm)", "3. DIM_A (Aberto)(mm)", "3. DIM_BxA (Aberto)(mm)", "3. DIM_B (Fechado)(mm)", "3. DIM_A (Fechado)(mm)", "3. DIM_E (mm)", "3. DIM_BXA (Fechado)", "3. DIM_Dimensão acabada (BxAxE)", "3. DIM_Peso (Kg)", "3. DIM_Quantidade total de páginas", "4. CAP_Tipo de capa", "4.Cap_orelha", "4. CAP_Papel da capa", "4. CAP_Cor capa", "4. CAP_Papel revestimento", "4. CAP_Papel forro", "4. CAP_Acabamento da capa", "4. CAP_Acabamento interno capa", "4. CAP_Aproveitamento chapa capa (Cor)", "4. CAP_Corte e vinco", "5. MIO_PRIN_Quantidade de páginas", "5. MIO_PRIN_Papel", "5. MIO_PRIN_Cor", "5. MIO_PRIN_Aproveitamento chapa miolo (Cor)", "5. MIO_PRIN_Corte e vinco", "5. MIO_PRIN_Serrilha", "5. MIO_PRIN_Obs miolo", "6. MIO_PROF_Quantidade de páginas", "6. MIO_PROF_Papel", "6. MIO_PROF_Cor", "6. MIO_PROF_Aproveitamento chapa miolo (Cor)", "6. MIO_PROF_Corte e vinco", "6. MIO_PROF_Serrilha", "6. MIO_PROF_Obs miolo", "7. MIO_ENCA_Quantidade de páginas", "7. MIO_ENCA_Papel", "7. MIO_ENCA_Cor", "7. MIO_ENCA_Aproveitamento chapa encarte (Cor)", "7. MIO_ENCA_Corte e vinco", "7. MIO_ENCA_Serrilha", "7. MIO_ENCA_Obs encarte", "8. MIO_ADES_Quantidade de páginas", "8. MIO_ADES_Papel", "8. MIO_ADES_Cor", "8. MIO_ADES_Aproveitamento chapa adesivo (Cor)", "8. MIO_ADES_Corte e vinco", "8. MIO_ADES_Obs adesivo", "9. ENC_Tipo acabamento", "9. ENC_Posição", "9.ENC_Bitola do espiral", "9. ENC_Cor do espiral", "9. ENC_Obs de encadernação", "10. OBS_PROD_Observação para produção gráfica", "10. Orientação de montagem do livro", "11. REF_Referência troca de chapa", "11. REF_Tipo de personalização", "11. REF_Quantidade de personalizados", "11. REF_Referência personalização", "12. VER_Data verificação editorial", "12. VER_Responsável verificação editorial", "12. VER_Data verificação engenharia", "12. VER_Responsável verificação engenharia"];
  var CABECALHOS_ARVORE = ["Grupo", "1. ID_Código Kit", "1. ID_Descrição Kit", "1. ID_Código SKU", "1. ID_ISBN", "1. ID_Descrição", "2. INF_Marca", "2. INF_Grupo da marca", "2. INF_Segmento", "2. INF_Série", "2. INF_Volume", "2. INF_Envio", "2. INF_Frequência", "2. INF_Usuário", "2. INF_Assunto/ Disciplina/ Área", "2. INF_Tipo de material", "2. INF_Classificação do produto", "2. INF_Cliente personalizado", "3. DIM_E (mm)", "3. DIM_Peso (Kg)", "13. MAN_Total de itens colecionados", "13. MAN_Embalagem", "13. MAN_Observação de embalagem", "13. MAN_Espessura do kit (mm)", "13. MAN_Peso do kit (kg)", "13. MAN Código Caixa NOVA", "13. MAN_Dimensões caixa parda NOVA", "13. MAN_Quantidade de kits por caixa parda NOVA", "Serie", "Envio V1 ?", "Envio V2 ?", "Envio V3 ?", "Envio V4 ?", "Conta Estoque"];
  var CABECALHOS_TIRAGEM = ["1. ID_Código Kit", "1. ID_Descrição Kit", "1. ID_Código SKU", "1. ID_ISBN", "1. ID_Descrição", "2. INF_Segmento", "2. INF_Série", "2. INF_Volume", "2. INF_Usuário", "11. REF_Referência troca de chapa", "11. REF_Referência personalização", "14. EST_Posição de estoque", "15. DES_CD envio", "16. TIR_SKU", "16. xx", "16. xxx", "16. xxxx", "16. TIR_Kit", "16. TIR_Quantidade de caixas pardas", "13. MAN Código Caixa", "17. VER_Data de alteração"];
  
  var planilhaControle = SpreadsheetApp.openById(ID_PLANILHA_LOG);
  var abaControle = planilhaControle.getSheetByName(NOME_ABA_LOG);
  var abaConvertida = planilhaControle.getSheetByName(NOME_ABA_LOG_CONV);

  function getLinhaReal(aba) {
    var last = aba.getLastRow();
    if (last === 0) return 0;
    var vals = aba.getRange(1, 1, last, 1).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      if (vals[i][0] !== "") return i + 1;
    }
    return 0;
  }

  function prepararAbaConsolidacao(nome, listaCabecalhos) {
    var aba = planilhaControle.getSheetByName(nome);
    if (!aba) {
      aba = planilhaControle.insertSheet(nome);
      var head = ["ID Solicitacao", "Link PPG Original"].concat(listaCabecalhos).concat(["Posição do Link"]);
      aba.appendRow(head);
      aba.getRange(1, 1, 1, head.length).setFontWeight("bold").setBackground("#f3f3f3");
    }
    return aba;
  }

  var abaEspec = prepararAbaConsolidacao(ABA_CONSOL_1, CABECALHOS_ESPEC);
  var abaArvore = prepararAbaConsolidacao(ABA_CONSOL_2, CABECALHOS_ARVORE);
  var abaTiragem = prepararAbaConsolidacao(ABA_CONSOL_3, CABECALHOS_TIRAGEM);

  function extrairLinksDaAba(aba) {
    var mapa = {};
    var lr = aba.getLastRow();
    if (lr > 1) {
      var vals = aba.getRange(2, 2, lr - 1, 1).getValues();
      vals.forEach(function(r) { if (r[0]) mapa[r[0].toString().trim()] = true; });
    }
    return mapa;
  }
  
  var linksNaEspec = extrairLinksDaAba(abaEspec);
  var linksNaArvore = extrairLinksDaAba(abaArvore);
  var linksNaTiragem = extrairLinksDaAba(abaTiragem);

  var ultimaLinhaLog = getLinhaReal(abaControle);
  var ultimaLinhaConv = getLinhaReal(abaConvertida);

  if (ultimaLinhaLog === 0) {
    abaControle.getRange(1, 1, 1, 5).setValues([["ID (Coluna A)", "PPG utilizada (Último Link)", "Última atualização", "Status", "ID Planilha Copiada (Excel)"]]);
    abaControle.getRange("A1:E1").setFontWeight("bold").setBackground("#f3f3f3");
    ultimaLinhaLog = 1;
  }
  if (ultimaLinhaConv === 0) {
    abaConvertida.getRange(1, 1, 1, 5).setValues([["ID (Coluna A)", "PPG utilizada (Último Link)", "Última atualização", "Status", "ID Planilha Convertida (Sheets)"]]);
    abaConvertida.getRange("A1:E1").setFontWeight("bold").setBackground("#f3f3f3");
    ultimaLinhaConv = 1;
  }

  var mapExcel = {};
  if (ultimaLinhaLog > 1) {
    var dExcel = abaControle.getRange(2, 2, ultimaLinhaLog - 1, 3).getValues();
    dExcel.forEach(function(r, i) { if (r[0]) mapExcel[r[0].toString().trim()] = { linha: i + 2, status: r[2].toString() }; });
  }

  var mapSheets = {};
  if (ultimaLinhaConv > 1) {
    var dSheets = abaConvertida.getRange(2, 2, ultimaLinhaConv - 1, 4).getValues();
    dSheets.forEach(function(r, i) { if (r[0]) mapSheets[r[0].toString().trim()] = { linha: i + 2, status: r[2].toString(), idPlanilha: r[3].toString() }; });
  }

  var planilhaFonte = SpreadsheetApp.openById(ID_PLANILHA_FONTE);
  var abaFonte = planilhaFonte.getSheetByName(NOME_ABA_FONTE);
  var ultimaLinhaFonte = abaFonte.getLastRow();
  if (ultimaLinhaFonte < LINHA_INICIAL_FONTE) return;
  var dadosFonte = abaFonte.getRange(LINHA_INICIAL_FONTE, 1, (ultimaLinhaFonte - LINHA_INICIAL_FONTE + 1), 17).getValues();

  var pastaDestino = DriveApp.getFolderById(ID_PASTA_DESTINO);
  var subpastaSheets = DriveApp.getFolderById(ID_SUBPASTA_SHEETS);
  var acoesRealizadas = 0;

  function encontrarCabecalhoReal(abaOrigem, listaCabecalhosDesejados) {
    var maxR = Math.min(abaOrigem.getLastRow(), 30);
    var maxC = Math.min(abaOrigem.getLastColumn(), 50);
    if (maxR === 0 || maxC === 0) return null;
    var dadosTop = abaOrigem.getRange(1, 1, maxR, maxC).getValues();
    var melhorLinha = -1;
    var maxMatches = 0;
    for (var r = 0; r < dadosTop.length; r++) {
      var matches = 0;
      for (var c = 0; c < dadosTop[r].length; c++) {
        var val = dadosTop[r][c].toString().trim();
        if (val !== "" && listaCabecalhosDesejados.indexOf(val) !== -1) {
          matches++;
        }
      }
      if (matches > maxMatches) {
        maxMatches = matches;
        melhorLinha = r + 1;
      }
    }
    if (maxMatches >= 2) return melhorLinha;
    return null;
  }

  function consolidarDinamico(ssOrigem, nomeAbaOrigem, listaCabecalhosDesejados, idSolicitacao, linkOriginal, abaDestinoMestre, nomeAncora, posicaoLink) {
    var abaOrigem = ssOrigem.getSheetByName(nomeAbaOrigem);
    if (!abaOrigem) return 0;

    var linhaCab = encontrarCabecalhoReal(abaOrigem, listaCabecalhosDesejados);
    if (!linhaCab) return 0;

    var headFonte = abaOrigem.getRange(linhaCab, 1, 1, abaOrigem.getLastColumn()).getValues()[0];
    var mapaIndices = {};
    headFonte.forEach(function(h, idx) { 
      var nomeCab = h.toString().trim();
      if (nomeCab !== "" && mapaIndices[nomeCab] === undefined) { 
        mapaIndices[nomeCab] = idx; 
      }
    });
    var idxAncora = mapaIndices[nomeAncora];
    var dadosFonte = abaOrigem.getRange(linhaCab + 1, 1, abaOrigem.getLastRow() - linhaCab, abaOrigem.getLastColumn()).getValues();
    var matrizFinal = [];
    for (var r = 0; r < dadosFonte.length; r++) {
      if (idxAncora !== undefined && (dadosFonte[r][idxAncora] === "" || dadosFonte[r][idxAncora] === null)) continue;
      var linhaTemDado = false;
      var novaLinha = [idSolicitacao, linkOriginal];
      
      listaCabecalhosDesejados.forEach(function(cab) {
        var idxNaFonte = mapaIndices[cab];
        if (idxNaFonte !== undefined) {
          var valor = dadosFonte[r][idxNaFonte];
          novaLinha.push(valor);
          if (valor !== "") linhaTemDado = true;
        } else {
          novaLinha.push("");
        
        }
      });
      
      novaLinha.push(posicaoLink);
      
      if (linhaTemDado) matrizFinal.push(novaLinha);
    }

    if (matrizFinal.length > 0) {
      var proximaLinha = getLinhaReal(abaDestinoMestre) + 1;
      abaDestinoMestre.getRange(proximaLinha, 1, matrizFinal.length, matrizFinal[0].length).setValues(matrizFinal);
      return matrizFinal.length;
    }
    return 0;
  }

  // --- LOOP PRINCIPAL ---
  for (var i = 0; i < dadosFonte.length; i++) {
    var idOriginal = dadosFonte[i][0];
    var idOriginalStr = idOriginal.toString().trim();
    var celulaQ = dadosFonte[i][16].toString(); 
    if (!celulaQ) continue;
    var partes = celulaQ.split(',');
    
    for (var k = 0; k < partes.length; k++) {
      var ultimoLink = partes[k].trim();
      if (!ultimoLink) continue;
      
      var posicaoLink = k + 1;
      var idArquivoOrigem = extrairIdDoLink(ultimoLink);
      if (!idArquivoOrigem) continue;

      // OTIMIZAÇÃO RISCO 0: Consolida buscas duplicadas de arquivo. Original buscava o arquivo novamente na Etapa 2.
      // Buscamos apenas UMA vez e reutilizamos.
      var arqDrive = null;
      var arquivoDuplicadoOriginalOriginalmenteBuscadoAquiEmbaixoBuscadoAcima = null; // Marcador apenas para controle visual do que mudou

      // ETAPA 1: EXCEL
      var infoEx = mapExcel[ultimoLink] || { status: "" };
      if (!infoEx.status || infoEx.status.indexOf("Erro") !== -1) {
        try {
          if (!arqDrive) arqDrive = DriveApp.getFileById(idArquivoOrigem); // Otimização
          var arq = arqDrive;
          var exis = pastaDestino.getFilesByName(arq.getName());
          while (exis.hasNext()) exis.next().setTrashed(true);
          var copia = arq.makeCopy(arq.getName(), pastaDestino);
          var linEx = infoEx.linha || ++ultimaLinhaLog;
          abaControle.getRange(linEx, 1, 1, 5).setValues([[idOriginal, ultimoLink, new Date(), "Copiado em Excel", copia.getId()]]);
          mapExcel[ultimoLink] = { linha: linEx, status: "Copiado em Excel" };
          acoesRealizadas++;
        } catch (e) {
          var linEx = infoEx.linha || ++ultimaLinhaLog;
          abaControle.getRange(linEx, 1, 1, 5).setValues([[idOriginal, ultimoLink, new Date(), "Erro Excel: " + e.message, "---"]]);
          mapExcel[ultimoLink] = { linha: linEx, status: "Erro Excel" };
        }
      }

      // ETAPA 2: CONVERSÃO
      var infoSh = mapSheets[ultimoLink] || { status: "", idPlanilha: "" };
      var idSheetsProcessar = infoSh.idPlanilha;
      if (!infoSh.status || (infoSh.status.indexOf("Erro") !== -1 && infoSh.status.indexOf("Consolidação") === -1)) {
        try {
          // OTIMIZAÇÃO RISCO 0: Reutiliza arqDrive buscado acima. O original buscava o arquivo novamente aqui.
          if (!arqDrive) arqDrive = DriveApp.getFileById(idArquivoOrigem); // Otimização
          var arq = arqDrive;
          var nomeSh = arq.getName().replace(/\.xlsx$|\.xls$/gi, "");
          var exisSh = subpastaSheets.getFilesByName(nomeSh);
          while (exisSh.hasNext()) exisSh.next().setTrashed(true);
          var config = { name: nomeSh, parents: [ID_SUBPASTA_SHEETS], mimeType: MimeType.GOOGLE_SHEETS };
          var arqSh = null;
          for (var t = 1; t <= 3; t++) {
            try { arqSh = Drive.Files.copy(config, idArquivoOrigem); break; } 
            catch (e) { Utilities.sleep(2000 * t); }
          }
          if (arqSh) {
            idSheetsProcessar = arqSh.id;
            try {
              var ssNovo = SpreadsheetApp.openById(idSheetsProcessar);
              ssNovo.setSpreadsheetLocale('pt_BR');
              ssNovo.setSpreadsheetTimeZone('America/Sao_Paulo');
            } catch(eLoc) {}

            var linSh = infoSh.linha || ++ultimaLinhaConv;
            abaConvertida.getRange(linSh, 1, 1, 5).setValues([[idOriginal, ultimoLink, new Date(), "Convertido para Sheets", idSheetsProcessar]]);
            infoSh = { linha: linSh, status: "Convertido para Sheets", idPlanilha: idSheetsProcessar };
            mapSheets[ultimoLink] = infoSh;
            acoesRealizadas++;
          } else { throw new Error("Falha após 3 tentativas na API"); }
        } catch (e) {
          var linSh = infoSh.linha || ++ultimaLinhaConv;
          abaConvertida.getRange(linSh, 1, 1, 5).setValues([[idOriginal, ultimoLink, new Date(), "Erro Sheets: " + e.message, "---"]]);
          infoSh = { linha: linSh, status: "Erro Sheets", idPlanilha: "" };
          mapSheets[ultimoLink] = infoSh;
        }
      }

      // ETAPA 3: CONSOLIDAÇÃO
      if (idSheetsProcessar && infoSh.status.indexOf("Erro Sheets") === -1) {
        
        var statusLog = infoSh.status;
        var estaConsolidadoNoLog = (statusLog.indexOf("Consolidado") !== -1);
        var arquivoEraVazio = (statusLog.indexOf("Vazio") !== -1);
        
        var faltaEspec = !linksNaEspec[ultimoLink];
        var faltaArvore = !linksNaArvore[ultimoLink];
        var faltaTiragem = !linksNaTiragem[ultimoLink];
        
        var precisaConsolidarAlguma = false;
        
        if (!estaConsolidadoNoLog || statusLog.indexOf("Erro Consolidação") !== -1) {
          precisaConsolidarAlguma = true;
        } else if (estaConsolidadoNoLog && !arquivoEraVazio && (faltaEspec || faltaArvore || faltaTiragem)) {
          precisaConsolidarAlguma = true;
        }

        if (precisaConsolidarAlguma) {
          try {
            var ss = SpreadsheetApp.openById(idSheetsProcessar);
            var linesAdicionadas = 0;
            
            if (faltaEspec) linesAdicionadas += consolidarDinamico(ss, "Especificação Técnica", CABECALHOS_ESPEC, idOriginal, ultimoLink, abaEspec, "1. ID_Código SKU", posicaoLink);
            if (faltaArvore) linesAdicionadas += consolidarDinamico(ss, "Árvore de produto", CABECALHOS_ARVORE, idOriginal, ultimoLink, abaArvore, "1. ID_Código SKU", posicaoLink);
            if (faltaTiragem) linesAdicionadas += consolidarDinamico(ss, "Tiragem", CABECALHOS_TIRAGEM, idOriginal, ultimoLink, abaTiragem, "1. ID_Código Kit", posicaoLink);

            var novoStatus = "Convertido e Consolidado";
            if (estaConsolidadoNoLog && arquivoEraVazio && linesAdicionadas === 0) {
                novoStatus = "Convertido e Consolidado (Vazio)";
            } else if (!estaConsolidadoNoLog && linesAdicionadas === 0) {
                novoStatus = "Convertido e Consolidado (Vazio)";
            }

            abaConvertida.getRange(infoSh.linha, 3, 1, 2).setValues([[new Date(), novoStatus]]);
            infoSh.status = novoStatus;
            mapSheets[ultimoLink] = infoSh;
            acoesRealizadas++;
            
          } catch (e) {
            abaConvertida.getRange(infoSh.linha, 3, 1, 2).setValues([[new Date(), "Erro Consolidação: " + e.message]]);
          }
        }
      }
    }
  }
  
  // FINAL CORRIGIDO: SEM ABRIR JANELAS, APENAS REGISTRA NO LOG INTERNO
  if (acoesRealizadas > 0) {
    console.log("Processamento concluído. Verifique as abas.");
  } else {
    console.log("Nenhuma atualização ou pendência encontrada.");
  }
}

function extrairIdDoLink(url) {
  var match = url.match(/id=([-\w]{25,})/);
  return match ? match[1] : (url.match(/\/d\/([-\w]{25,})/) ? url.match(/\/d\/([-\w]{25,})/)[1] : null);
}